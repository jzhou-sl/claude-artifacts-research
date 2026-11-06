#!/usr/bin/env node
// profile-mcp: a least-privilege MCP server (stdio) over Microsoft Graph using only the delegated
// User.Read permission. One tool, get_me, with the same result shape as teams-mcp's get_me.
// The handler lives in tool.mjs, shared with the SharePoint web part.
// Sign in first with `node --env-file-if-exists=.env mcp/profile/login.mjs`.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { getToken } from "./auth.mjs";
import { createGraph } from "../teams-core/graph.js";
import { createProfileTools } from "./tool.mjs";

const log = (...a) => console.error("[profile-mcp]", ...a); // stdout is the MCP channel

const tools = createProfileTools({ ...createGraph(() => getToken()), log });
const server = new McpServer({ name: "profile-mcp", version: "0.1.0" });

for (const [name, t] of Object.entries(tools)) {
  server.registerTool(name, { description: t.description, inputSchema: {}, annotations: t.annotations }, () => t.handler());
}

await server.connect(new StdioServerTransport());
log("ready");
