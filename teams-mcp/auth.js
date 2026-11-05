// MSAL public-client auth for teams-mcp: delegated tokens, cached on disk with DPAPI (Windows),
// Keychain (macOS) or libsecret (Linux) via @azure/msal-node-extensions.
import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-node";
import { PersistenceCreator, PersistenceCachePlugin, DataProtectionScope } from "@azure/msal-node-extensions";
import { readFileSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { AuthError, CHANNEL_READ } from "../teams-core/errors.js";

/** Scopes every user can consent to. Requested at login. */
export const BASE_SCOPES = ["User.Read", "People.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All", "Chat.ReadWrite", "ChannelMessage.Send"];

export { CHANNEL_READ, AuthError };

/* ---------- config: env first, then the teams-mcp entry in mcp.config.json ---------- */
function loadConfig() {
  let tenantId = process.env.GRAPH_TENANT_ID, clientId = process.env.GRAPH_CLIENT_ID;
  if (!tenantId || !clientId) {
    try {
      const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
      const cfg = JSON.parse(readFileSync(path.join(root, process.env.MCP_CONFIG || "mcp.config.json"), "utf8"));
      const entry = Object.values(cfg).find((e) => (e.args || []).some((a) => /teams-mcp[\\/]server\.js$/.test(a)));
      tenantId ||= entry?.env?.GRAPH_TENANT_ID; clientId ||= entry?.env?.GRAPH_CLIENT_ID;
    } catch { /* fall through to validation */ }
  }
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!guid.test(clientId || "")) throw new AuthError("not_configured", "GRAPH_CLIENT_ID is missing or not a GUID. Set it in mcp.config.json (env) to your app registration's Application (client) ID.");
  if (!guid.test(tenantId || "") && !/^[\w.-]+\.[a-z]{2,}$/i.test(tenantId || "")) throw new AuthError("not_configured", "GRAPH_TENANT_ID is missing. Set it in mcp.config.json (env) to your Directory (tenant) ID or domain.");
  return { tenantId, clientId };
}

const CACHE_DIR = path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), ".cache"), "teams-mcp");
export const CACHE_PATH = path.join(CACHE_DIR, "msal-cache.bin");

let pcaPromise = null;
export function getPca() {
  pcaPromise ||= (async () => {
    const { tenantId, clientId } = loadConfig();
    await mkdir(CACHE_DIR, { recursive: true });
    const persistence = await PersistenceCreator.createPersistence({
      cachePath: CACHE_PATH,
      dataProtectionScope: DataProtectionScope.CurrentUser,
      serviceName: "teams-mcp",
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

export async function getAccount() {
  const pca = await getPca();
  const accounts = await pca.getTokenCache().getAllAccounts();
  return accounts[0] || null;
}

/** Silent token for the given scopes. Never prompts: run `npm run login` to sign in. */
export async function getToken(scopes) {
  const pca = await getPca();
  const account = await getAccount();
  if (!account) throw new AuthError("needs_reauth", "Not signed in to Microsoft Graph. Run `npm run login`, then reload the page.");
  try {
    const r = await pca.acquireTokenSilent({ account, scopes });
    return r.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError || /interaction_required|invalid_grant|consent_required/i.test(e?.errorCode || "")) {
      if (scopes.includes(CHANNEL_READ)) {
        throw new AuthError("consent_missing",
          `Reading channel messages needs the ${CHANNEL_READ} permission, which requires admin consent in most tenants. Ask an admin to grant it to the app, or run \`npm run login -- --channels\` if you're allowed to consent.`, CHANNEL_READ);
      }
      throw new AuthError("needs_reauth", "Your Microsoft sign-in has expired or needs consent. Run `npm run login`, then reload the page.");
    }
    throw e;
  }
}
