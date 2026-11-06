#!/usr/bin/env node
// `npm run serve:profile`: serves web-parts/profile-card.html locally against real Microsoft Graph.
//   1. Checks for a cached User.Read token (mcp/profile/auth.mjs) and runs mcp/profile/login.mjs if
//      there isn't one, or it has expired.
//   2. Starts dev-server.mjs on port 7071 (PROFILE_PORT overrides it) with profile-card.html at /, so the page's Profile calls go to mcp/profile.
// Needs PROFILE_TENANT_ID and PROFILE_CLIENT_ID in .env (see README, "Profile server").
import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { getToken } from "../mcp/profile/auth.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const run = (script, env = {}) => new Promise((resolve) => {
  const child = spawn(process.execPath, [path.join(ROOT, script)], { cwd: ROOT, stdio: "inherit", env: { ...process.env, ...env } });
  child.on("exit", (code, signal) => resolve(signal ? 1 : code ?? 1));
});

try {
  await getToken();
  console.log("[serve:profile] Signed in (cached User.Read token).");
} catch (e) {
  if (e?.code !== "needs_reauth") {
    console.error(`[serve:profile] ${e?.message || e}`);
    process.exit(1);
  }
  console.log("[serve:profile] Not signed in. Starting sign-in…");
  const code = await run("mcp/profile/login.mjs");
  if (code !== 0) process.exit(code);
}

const port = process.env.PROFILE_PORT || "7071";
console.log(`[serve:profile] Open http://127.0.0.1:${port}/`);
process.exit(await run("dev-server.mjs", { PORT: port, DEFAULT_PAGE: "profile-card.html", MCP_CONFIG: "mcp.config.json" }));
