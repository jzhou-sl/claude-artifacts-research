#!/usr/bin/env node
// teams-mcp: an MCP server (stdio) wrapping Microsoft Graph with delegated permissions.
// Exposes the tool names and result shapes teams-slack-style.html expects (same as mock-mcp/).
// The handlers live in ../teams-core/tools.js (shared with the SharePoint web part); this file adds
// input schemas and MSAL Node auth. Sign in first with `npm run login`.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { getToken } from "./auth.js";
import { createGraph } from "../teams-core/graph.js";
import { createTools } from "../teams-core/tools.js";

const log = (...a) => console.error("[teams-mcp]", ...a); // stdout is the MCP channel

const SCHEMAS = {
  get_me: {},
  search_people: { query: z.string(), limit: z.number().int().positive().max(50).optional() },
  teams_list_teams: {},
  teams_list_channels: { teamId: z.string() },
  teams_list_channel_messages: { teamId: z.string(), channelId: z.string(), parentMessageId: z.string().optional(),
    limit: z.number().int().positive().optional(), cursor: z.string().optional() },
  teams_send_channel_message: { teamId: z.string(), channelId: z.string(), body: z.string().min(1), bodyType: z.enum(["text", "html"]).optional() },
  teams_reply_channel_message: { teamId: z.string(), channelId: z.string(), messageId: z.string(), body: z.string().min(1), bodyType: z.enum(["text", "html"]).optional() },
  teams_list_chats: { limit: z.number().int().positive().optional(), cursor: z.string().optional() },
  teams_send_chat_message: { chatId: z.string(), body: z.string().min(1), bodyType: z.enum(["text", "html"]).optional() },
  teams_create_chat: { members: z.array(z.string()).min(1), chatType: z.enum(["oneOnOne", "group"]).optional(), topic: z.string().optional() },
  chat_message_search: { query: z.string(), limit: z.number().int().positive().optional(), offset: z.number().int().nonnegative().optional(),
    afterDateTime: z.string().optional(), sender: z.string().optional(), recipient: z.string().optional() },
  read_resource: { uri: z.string() },
};

const tools = createTools({ ...createGraph(getToken), log });
const server = new McpServer({ name: "teams-mcp", version: "0.2.0" });
for (const [name, t] of Object.entries(tools)) {
  server.registerTool(name, { description: t.description, inputSchema: SCHEMAS[name] || {}, annotations: t.annotations }, t.handler);
}

await server.connect(new StdioServerTransport());
log("ready");
