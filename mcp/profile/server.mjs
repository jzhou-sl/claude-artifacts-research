#!/usr/bin/env node
// profile-mcp: a least-privilege MCP server (stdio) over Microsoft Graph using only the delegated
// User.Read permission. One tool, get_me, with the same result shape as teams-mcp's get_me.
// Sign in first with `node --env-file-if-exists=.env mcp/profile/login.mjs`.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { getToken, SCOPES, AuthError } from "./auth.mjs";
import { createGraph, GraphError } from "../teams-core/graph.js";

const log = (...a) => console.error("[profile-mcp]", ...a); // stdout is the MCP channel

const SELECT = "id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,preferredLanguage";

const fail = (code, message, extra = {}) => ({ isError: true, content: [{ type: "text", text: message }], structuredContent: { code, message, ...extra } });

// Same error mapping as mcp/teams-core/tools.js, so resultError() and the dev-server bridge treat both servers alike.
function toFailure(e) {
  if (e instanceof AuthError) return fail(e.code === "needs_reauth" ? "needs_reauth" : "tool_error", e.message);
  if (e instanceof GraphError) {
    if (e.status === 401) return fail("needs_reauth", "Microsoft Graph rejected the sign-in. Run `node --env-file-if-exists=.env mcp/profile/login.mjs`, then retry.");
    if (e.status === 403) return fail("tool_error", `Permission denied by Microsoft Graph (${e.code}): ${e.message} Check the app registration has delegated User.Read.`);
    if ([429, 503, 504].includes(e.status)) return fail("server_unavailable", `Microsoft Graph is busy (${e.status}). ${e.message}`, { retryable: true, retryAfterMs: e.retryAfterMs });
    if (e.status >= 500) return fail("upstream_error", `Microsoft Graph error ${e.status}: ${e.message}`, { retryable: true, retryAfterMs: 1500 });
    return fail("tool_error", `${e.message} (${e.code})`);
  }
  log("unexpected error", e);
  return fail("upstream_error", String(e?.message || e), { retryable: true, retryAfterMs: 1500 });
}

const { graph } = createGraph(() => getToken());
const server = new McpServer({ name: "profile-mcp", version: "0.1.0" });

server.registerTool("get_me", {
  description: "Get the signed-in user's Microsoft 365 profile (name, email, job title, department, office, phones).",
  inputSchema: {},
  annotations: { readOnlyHint: true },
}, async () => {
  try {
    const { "@odata.context": _, ...me } = await graph("GET", "/me", { scopes: SCOPES, query: { $select: SELECT } });
    return { content: [{ type: "text", text: JSON.stringify({ ...me, mail: me.mail || me.userPrincipalName }) }] };
  } catch (e) {
    return toFailure(e);
  }
});

await server.connect(new StdioServerTransport());
log("ready");
