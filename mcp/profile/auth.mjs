// MSAL public-client auth for profile-mcp: delegated User.Read only, cached on disk with DPAPI (Windows),
// Keychain (macOS) or libsecret (Linux) via @azure/msal-node-extensions. Separate cache from teams-mcp.
import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-node";
import { PersistenceCreator, PersistenceCachePlugin, DataProtectionScope } from "@azure/msal-node-extensions";
import { readFileSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { AuthError } from "../teams-core/errors.js";

/** The only permission this server uses. No admin consent needed. */
export const SCOPES = ["User.Read"];

export { AuthError };

/* ---------- config: env first, then the profile entry in mcp.config.json ---------- */
function loadConfig() {
  let tenantId = process.env.PROFILE_TENANT_ID, clientId = process.env.PROFILE_CLIENT_ID;
  if (!tenantId || !clientId) {
    try {
      const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
      const cfg = JSON.parse(readFileSync(path.join(root, process.env.MCP_CONFIG || "mcp.config.json"), "utf8"));
      const entry = Object.values(cfg).find((e) => (e.args || []).some((a) => /mcp[\\/]profile[\\/]server\.mjs$/.test(a)));
      tenantId ||= entry?.env?.PROFILE_TENANT_ID; clientId ||= entry?.env?.PROFILE_CLIENT_ID;
    } catch { /* fall through to validation */ }
  }
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!guid.test(clientId || "")) throw new AuthError("not_configured", "PROFILE_CLIENT_ID is missing or not a GUID. Set it in .env or mcp.config.json (env) to your app registration's Application (client) ID.");
  if (!guid.test(tenantId || "") && !/^[\w.-]+\.[a-z]{2,}$/i.test(tenantId || "")) throw new AuthError("not_configured", "PROFILE_TENANT_ID is missing. Set it in .env or mcp.config.json (env) to your Directory (tenant) ID or domain.");
  return { tenantId, clientId };
}

const CACHE_DIR = path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), ".cache"), "profile-mcp");
export const CACHE_PATH = path.join(CACHE_DIR, "msal-cache.bin");

let pcaPromise = null;
export function getPca() {
  pcaPromise ||= (async () => {
    const { tenantId, clientId } = loadConfig();
    await mkdir(CACHE_DIR, { recursive: true });
    const persistence = await PersistenceCreator.createPersistence({
      cachePath: CACHE_PATH,
      dataProtectionScope: DataProtectionScope.CurrentUser,
      serviceName: "profile-mcp",
      accountName: "msal-cache",
      usePlaintextFileOnLinux: false,
    });
    return new PublicClientApplication({
      auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}` },
      cache: { cachePlugin: new PersistenceCachePlugin(persistence) },
    });
  })();
  pcaPromise.catch(() => { pcaPromise = null; });
  return pcaPromise;
}

/** Silent User.Read token. Never prompts: run `node --env-file-if-exists=.env mcp/profile/login.mjs` to sign in. */
export async function getToken() {
  const pca = await getPca();
  const [account] = await pca.getTokenCache().getAllAccounts();
  if (!account) throw new AuthError("needs_reauth", "Not signed in to Microsoft Graph. Run `node --env-file-if-exists=.env mcp/profile/login.mjs`, then retry.");
  try {
    return (await pca.acquireTokenSilent({ account, scopes: SCOPES })).accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError || /interaction_required|invalid_grant|consent_required/i.test(e?.errorCode || "")) {
      throw new AuthError("needs_reauth", "Your Microsoft sign-in has expired or needs consent. Run `node --env-file-if-exists=.env mcp/profile/login.mjs`, then retry.");
    }
    throw e;
  }
}
