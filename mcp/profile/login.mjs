#!/usr/bin/env node
// Sign in (or out) for profile-mcp. Requests only User.Read.
//   node --env-file-if-exists=.env mcp/profile/login.mjs    sign in (browser, auth code + PKCE)
//   node --env-file-if-exists=.env mcp/profile/login.mjs --logout   remove cached accounts
//
// The sign-in callback listens on PROFILE_REDIRECT_URI (default http://localhost:7071/auth/callback), which must match a
// redirect URI on the app registration exactly. This runs the auth code flow itself rather than MSAL Node's
// acquireTokenInteractive, which picks its own port and only works as a public client (no PROFILE_CLIENT_SECRET).
// An https redirect URI also works: it uses the SPFx/Heft dev certificate in ~/.rushstack.
import { spawn } from "node:child_process";
import http from "node:http";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import { readFile, rm } from "node:fs/promises";
import { CryptoProvider } from "@azure/msal-node";
import { getPca, SCOPES, CACHE_PATH } from "./auth.mjs";

const REDIRECT_URI = process.env.PROFILE_REDIRECT_URI || "http://localhost:7071/auth/callback";
const TIMEOUT_MS = 5 * 60_000;

function openBrowser(url) {
  const [cmd, argv] = process.platform === "win32" ? ["rundll32", ["url.dll,FileProtocolHandler", url]]
    : process.platform === "darwin" ? ["open", [url]] : ["xdg-open", [url]];
  spawn(cmd, argv, { stdio: "ignore", detached: true }).unref();
  console.log("Opened your browser to sign in. If nothing appeared, open this URL:\n" + url + "\n");
}

const page = (title, body) => `<!doctype html><meta charset="utf-8"><title>${title}</title>
<body style="font:16px system-ui;display:grid;place-items:center;height:90vh;margin:0"><div style="text-align:center"><h1>${title}</h1><p>${body}</p></div>`;
const escapeHtml = (s) => String(s).replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);

async function devCertificate() {
  const dir = path.join(os.homedir(), ".rushstack");
  try {
    return { cert: await readFile(path.join(dir, "rushstack-serve.pem")), key: await readFile(path.join(dir, "rushstack-serve.key")) };
  } catch {
    throw new Error(`No dev certificate in ${dir}. Run \`npx heft trust-dev-cert\` (or \`npm start\` once) to create and trust it, then retry.`);
  }
}

/**
 * Listens on the redirect URI and opens the browser at getAuthUrl() once listening. When Entra sends the user
 * back, redeems the code with redeem(code) before answering the browser, so its page shows the real outcome.
 * Resolves with redeem's result.
 */
async function signInInBrowser(redirect, expectedState, getAuthUrl, redeem) {
  const options = redirect.protocol === "https:" ? await devCertificate() : {};
  const port = Number(redirect.port) || (redirect.protocol === "https:" ? 443 : 80);
  const callbackPath = redirect.pathname || "/";
  return new Promise((resolve, reject) => {
    let timer;
    const finish = (fn, value) => { clearTimeout(timer); server.close(); fn(value); };
    const server = (redirect.protocol === "https:" ? https : http).createServer(options, async (req, res) => {
      const url = new URL(req.url, redirect);
      const params = url.searchParams;
      if (url.pathname !== callbackPath || (!params.has("code") && !params.has("error"))) { res.writeHead(404); return res.end(); }
      res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
      if (params.get("state") !== expectedState) {
        res.end(page("Sign-in failed", "The response didn't match this sign-in. Return to the terminal and retry."));
        return finish(reject, new Error("State mismatch in the sign-in response."));
      }
      if (params.has("error")) {
        res.end(page("Sign-in failed", escapeHtml(params.get("error_description") || params.get("error"))));
        return finish(reject, new Error(params.get("error_description") || params.get("error")));
      }
      try {
        const result = await redeem(params.get("code"));
        res.end(page("Signed in", "You can close this tab and return to the terminal."));
        finish(resolve, result);
      } catch (e) {
        res.end(page("Sign-in failed", "Return to the terminal for details."));
        finish(reject, e);
      }
    });
    server.once("error", (e) => {
      clearTimeout(timer);
      reject(e.code === "EADDRINUSE"
        ? new Error(`Port ${port} is in use, so the sign-in callback (${REDIRECT_URI}) can't listen on it. Stop whatever is using it (for example a running dev server) and retry.`)
        : e);
    });
    timer = setTimeout(() => finish(reject, new Error("Timed out waiting for the browser sign-in.")), TIMEOUT_MS);
    server.listen(port, redirect.hostname === "localhost" ? "localhost" : redirect.hostname, async () => {
      try { openBrowser(await getAuthUrl()); } catch (e) { finish(reject, e); }
    });
  });
}

try {
  const pca = await getPca();

  if (process.argv.includes("--logout")) {
    for (const a of await pca.getTokenCache().getAllAccounts()) await pca.getTokenCache().removeAccount(a);
    await rm(CACHE_PATH, { force: true });
    console.log("Signed out. Token cache removed.");
    process.exit(0);
  }

  const redirect = new URL(REDIRECT_URI);
  const crypto = new CryptoProvider();
  const { verifier, challenge } = await crypto.generatePkceCodes();
  const state = crypto.createNewGuid();
  const r = await signInInBrowser(redirect, state,
    () => pca.getAuthCodeUrl({ scopes: SCOPES, redirectUri: REDIRECT_URI, codeChallenge: challenge, codeChallengeMethod: "S256", state, prompt: "select_account" }),
    (code) => pca.acquireTokenByCode({ code, scopes: SCOPES, redirectUri: REDIRECT_URI, codeVerifier: verifier }));
  console.log(`Signed in as ${r.account.name} <${r.account.username}>`);
  console.log(`Granted: ${r.scopes.join(", ")}`);
  process.exit(0);
} catch (e) {
  const msg = e.errorMessage || e.message || String(e);
  console.error("Sign-in failed:", msg);
  if (/AADSTS7000218/.test(msg)) {
    console.error("\nThis redirect URI is registered under the Web platform, which needs a client secret to redeem the sign-in code. Set PROFILE_CLIENT_SECRET in .env to a secret from the app registration (Certificates & secrets), then retry.");
  }
  process.exit(1);
}
