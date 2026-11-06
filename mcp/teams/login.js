#!/usr/bin/env node
// Sign in (or out) for teams-mcp.
//   node mcp/teams/login.js              sign in with the base scopes (browser, auth code + PKCE)
//   node mcp/teams/login.js --channels   also try ChannelMessage.Read.All (needs admin consent in most tenants)
//   node mcp/teams/login.js --logout     remove cached accounts
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import { getPca, BASE_SCOPES, CHANNEL_READ, CACHE_PATH } from "./auth.js";

const args = new Set(process.argv.slice(2));

function openBrowser(url) {
  const [cmd, argv] = process.platform === "win32" ? ["rundll32", ["url.dll,FileProtocolHandler", url]]
    : process.platform === "darwin" ? ["open", [url]] : ["xdg-open", [url]];
  spawn(cmd, argv, { stdio: "ignore", detached: true }).unref();
  console.log("Opened your browser to sign in. If nothing appeared, open this URL:\n" + url + "\n");
  return Promise.resolve();
}

const page = (title, body) => `<!doctype html><meta charset="utf-8"><title>${title}</title>
<body style="font:16px system-ui;display:grid;place-items:center;height:90vh;margin:0"><div style="text-align:center"><h1>${title}</h1><p>${body}</p></div>`;

async function interactive(pca, scopes) {
  return pca.acquireTokenInteractive({
    scopes,
    openBrowser,
    prompt: "select_account",
    successTemplate: page("Signed in", "You can close this tab and return to the terminal."),
    errorTemplate: page("Sign-in failed", "Return to the terminal for details."),
  });
}

try {
  const pca = await getPca();

  if (args.has("--logout")) {
    for (const a of await pca.getTokenCache().getAllAccounts()) await pca.getTokenCache().removeAccount(a);
    await rm(CACHE_PATH, { force: true });
    console.log("Signed out. Token cache removed.");
    process.exit(0);
  }

  const r = await interactive(pca, BASE_SCOPES);
  console.log(`Signed in as ${r.account.name} <${r.account.username}>`);
  console.log(`Granted: ${r.scopes.join(", ")}`);

  if (args.has("--channels")) {
    try {
      await interactive(pca, [...BASE_SCOPES, CHANNEL_READ]);
      console.log(`${CHANNEL_READ} granted: channel messages will load.`);
    } catch (e) {
      console.warn(`\nCouldn't get ${CHANNEL_READ}: ${e.errorMessage || e.message}`);
      console.warn("This usually needs an admin to grant consent for the app. Chats and sending still work.");
    }
  } else {
    console.log(`\nChannel messages need ${CHANNEL_READ} (admin consent). If it's already granted for the app, it will work automatically;`);
    console.log("otherwise run `node mcp/teams/login.js --channels` to request it.");
  }
  process.exit(0);
} catch (e) {
  console.error("Sign-in failed:", e.errorMessage || e.message || e);
  process.exit(1);
}
