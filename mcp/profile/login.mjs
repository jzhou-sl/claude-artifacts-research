#!/usr/bin/env node
// Sign in (or out) for profile-mcp. Requests only User.Read.
//   node --env-file-if-exists=.env mcp/profile/login.mjs    sign in (browser, auth code + PKCE)
//   node --env-file-if-exists=.env mcp/profile/login.mjs --logout   remove cached accounts
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import { getPca, SCOPES, CACHE_PATH } from "./auth.mjs";

function openBrowser(url) {
  const [cmd, argv] = process.platform === "win32" ? ["rundll32", ["url.dll,FileProtocolHandler", url]]
    : process.platform === "darwin" ? ["open", [url]] : ["xdg-open", [url]];
  spawn(cmd, argv, { stdio: "ignore", detached: true }).unref();
  console.log("Opened your browser to sign in. If nothing appeared, open this URL:\n" + url + "\n");
  return Promise.resolve();
}

const page = (title, body) => `<!doctype html><meta charset="utf-8"><title>${title}</title>
<body style="font:16px system-ui;display:grid;place-items:center;height:90vh;margin:0"><div style="text-align:center"><h1>${title}</h1><p>${body}</p></div>`;

try {
  const pca = await getPca();

  if (process.argv.includes("--logout")) {
    for (const a of await pca.getTokenCache().getAllAccounts()) await pca.getTokenCache().removeAccount(a);
    await rm(CACHE_PATH, { force: true });
    console.log("Signed out. Token cache removed.");
    process.exit(0);
  }

  const r = await pca.acquireTokenInteractive({
    scopes: SCOPES,
    openBrowser,
    prompt: "select_account",
    successTemplate: page("Signed in", "You can close this tab and return to the terminal."),
    errorTemplate: page("Sign-in failed", "Return to the terminal for details."),
  });
  console.log(`Signed in as ${r.account.name} <${r.account.username}>`);
  console.log(`Granted: ${r.scopes.join(", ")}`);
  process.exit(0);
} catch (e) {
  console.error("Sign-in failed:", e.errorMessage || e.message || e);
  process.exit(1);
}
