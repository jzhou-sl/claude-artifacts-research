#!/usr/bin/env node
// Mock "Microsoft 365" MCP server (stdio). Serves canned Teams data with in-memory state.
//
// Env:
//   MOCK_CHATTER=0          disable the background "someone posts a message" timer
//   MOCK_CHATTER_MS=45000   interval for background messages
//   MOCK_FAIL=tool_a,tool_b make these tools return isError (for testing error handling)

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { people, teams, channels, me, byId, seed, makeChannelMessage, makeChatMessage } from "./fixtures.js";

const state = seed();
const FAIL = new Set((process.env.MOCK_FAIL || "").split(",").map((s) => s.trim()).filter(Boolean));
const log = (...a) => console.error("[mock-m365]", ...a); // stderr only: stdout is the MCP channel

/* ---------- helpers ---------- */
const pub = (o) => Object.fromEntries(Object.entries(o).filter(([k]) => !k.startsWith("_")));
const newestFirst = (a, b) => (a.createdDateTime < b.createdDateTime ? 1 : -1);
const text = (...values) => ({ content: [{ type: "text", text: values.map((v) => JSON.stringify(v, null, 1)).join("\n") }] });
const fail = (message) => ({ isError: true, content: [{ type: "text", text: message }] });
const page = (rows, offset, limit) => {
  const slice = rows.slice(offset, offset + limit);
  const next = offset + slice.length;
  return { slice, more: next < rows.length, next };
};
const toOffset = (cursor) => { const n = parseInt(cursor ?? "0", 10); return Number.isFinite(n) && n >= 0 ? n : 0; };
const findTeamOfChannel = (channelId) => Object.keys(channels).find((t) => channels[t].some((c) => c.id === channelId));
const chatMembers = (chat) => chat.memberIds.map(byId).map((p) => ({ displayName: p.displayName, email: p.mail, userId: p.id }));
const lastActivity = (chat) => state.chatMessages.filter((m) => m.chatId === chat.id).reduce((mx, m) => (m.createdDateTime > mx ? m.createdDateTime : mx), "");

const server = new McpServer({ name: "mock-microsoft-365", version: "0.1.0" });

function tool(name, description, inputSchema, handler) {
  server.registerTool(name, { description, inputSchema }, async (args) => {
    if (FAIL.has(name)) return fail(`Mock failure for ${name} (MOCK_FAIL).`);
    try { return await handler(args); }
    catch (e) { log(name, "threw", e); return fail(String(e && e.message || e)); }
  });
}

/* ---------- profile & people ---------- */
tool("get_me", "Get the signed-in user's profile.", {}, async () => text(pub(me)));

tool("search_people", "Search people in the organisation by name or email.",
  { query: z.string(), limit: z.number().int().positive().optional() },
  async ({ query, limit = 15 }) => {
    const q = query.toLowerCase();
    const rows = people.filter((p) => p.displayName.toLowerCase().includes(q) || p.mail.includes(q)).slice(0, limit);
    return text(rows);
  });

/* ---------- teams & channels ---------- */
tool("teams_list_teams", "List teams the user has joined.", {}, async () => text(teams));

tool("teams_list_channels", "List channels in a team.", { teamId: z.string() }, async ({ teamId }) => {
  const list = channels[teamId];
  if (!list) return fail(`Team ${teamId} not found.`);
  return text(list.map((c) => ({ ...c, webUrl: `https://teams.example/l/channel/${c.id}` })));
});

tool("teams_list_channel_messages", "List messages in a channel, or replies to a message when parentMessageId is set. Newest first.",
  { teamId: z.string(), channelId: z.string(), parentMessageId: z.string().optional(),
    limit: z.number().int().positive().optional(), cursor: z.string().optional() },
  async ({ teamId, channelId, parentMessageId, limit = 20, cursor }) => {
    if (!channels[teamId]?.some((c) => c.id === channelId)) return fail(`Channel ${channelId} not found in team ${teamId}.`);
    const rows = state.channelMessages
      .filter((m) => m.channelId === channelId && (parentMessageId ? m.replyToId === parentMessageId : !m.replyToId))
      .sort(newestFirst);
    const { slice, more, next } = page(rows, toOffset(cursor), limit);
    return text(slice.map(pub), { moreResults: more, nextCursor: more ? String(next) : null });
  });

tool("teams_send_channel_message", "Post a new message to a channel.",
  { teamId: z.string(), channelId: z.string(), body: z.string(), bodyType: z.enum(["text", "html"]).optional() },
  async ({ teamId, channelId, body }) => {
    if (!channels[teamId]?.some((c) => c.id === channelId)) return fail(`Channel ${channelId} not found in team ${teamId}.`);
    const m = makeChannelMessage({ teamId, channelId, authorId: me.id, text: body, at: new Date().toISOString() });
    state.channelMessages.push(m);
    return text({ id: m.id, createdDateTime: m.createdDateTime, webUrl: m.webUrl });
  });

tool("teams_reply_channel_message", "Reply to a channel message.",
  { teamId: z.string(), channelId: z.string(), messageId: z.string(), body: z.string(), bodyType: z.enum(["text", "html"]).optional() },
  async ({ teamId, channelId, messageId, body }) => {
    const root = state.channelMessages.find((m) => m.id === messageId && m.channelId === channelId);
    if (!root) return fail(`Message ${messageId} not found.`);
    const m = makeChannelMessage({ teamId, channelId, authorId: me.id, text: body, at: new Date().toISOString(), parentId: messageId });
    state.channelMessages.push(m);
    return text({ id: m.id, createdDateTime: m.createdDateTime });
  });

/* ---------- chats ---------- */
tool("teams_list_chats", "List the user's chats, most recently active first.",
  { limit: z.number().int().positive().optional(), cursor: z.string().optional() },
  async ({ limit = 25, cursor }) => {
    const rows = state.chats
      .map((c) => ({ id: c.id, chatType: c.chatType, topic: c.topic, memberCount: c.memberIds.length, members: chatMembers(c), lastActivity: lastActivity(c) }))
      .sort((a, b) => (a.lastActivity < b.lastActivity ? 1 : -1));
    const { slice, more, next } = page(rows, toOffset(cursor), limit);
    return text(slice, { moreResults: more, nextCursor: more ? String(next) : null });
  });

tool("teams_send_chat_message", "Send a message to a chat.",
  { chatId: z.string(), body: z.string(), bodyType: z.enum(["text", "html"]).optional() },
  async ({ chatId, body }) => {
    if (!state.chats.some((c) => c.id === chatId)) return fail(`Chat ${chatId} not found.`);
    const m = makeChatMessage({ chatId, authorId: me.id, text: body, at: new Date().toISOString() });
    state.chatMessages.push(m);
    return text({ id: m.id, chatId, createdDateTime: m.createdDateTime });
  });

tool("teams_create_chat", "Create a chat (or return the existing 1:1 chat) with the given members (UPNs).",
  { members: z.array(z.string()).min(1), chatType: z.enum(["oneOnOne", "group"]).optional(), topic: z.string().optional() },
  async ({ members, chatType = "oneOnOne", topic }) => {
    const ids = members.map((upn) => people.find((p) => p.userPrincipalName.toLowerCase() === upn.toLowerCase())?.id);
    if (ids.some((x) => !x)) return fail(`Unknown member(s): ${members.join(", ")}`);
    const memberIds = [me.id, ...ids.filter((x) => x !== me.id)];
    if (chatType === "oneOnOne") {
      const existing = state.chats.find((c) => c.chatType === "oneOnOne" && c.memberIds.length === 2 && c.memberIds.includes(memberIds[1]));
      if (existing) return text({ id: existing.id });
    }
    const id = chatType === "oneOnOne" ? `19:${memberIds.join("_")}@unq.gbl.spaces` : `19:${Date.now().toString(36)}@thread.v2`;
    state.chats.push({ id, chatType, topic: topic ?? null, memberIds });
    return text({ id });
  });

tool("chat_message_search", "Search chat messages. query '*' matches everything.",
  { query: z.string(), limit: z.number().int().positive().optional(), offset: z.number().int().nonnegative().optional(),
    afterDateTime: z.string().optional(), sender: z.string().optional(), recipient: z.string().optional() },
  async ({ query, limit = 25, offset = 0, afterDateTime, sender, recipient }) => {
    const q = query.trim().toLowerCase();
    const after = afterDateTime ? new Date(afterDateTime).toISOString() : null;
    const rows = state.chatMessages.filter((m) => {
      if (q && q !== "*" && !m.summary.toLowerCase().includes(q)) return false;
      if (after && m.createdDateTime < after) return false;
      if (sender && m.from.email.toLowerCase() !== sender.toLowerCase()) return false;
      if (recipient) {
        const chat = state.chats.find((c) => c.id === m.chatId);
        if (!chat || !chatMembers(chat).some((p) => p.email.toLowerCase() === recipient.toLowerCase())) return false;
      }
      return true;
    }).sort(newestFirst);
    const { slice, more, next } = page(rows, offset, limit);
    return text(slice, { moreResults: more, nextOffset: next });
  });

/* ---------- resources (full bodies + reactions) ---------- */
tool("read_resource", "Read a Teams resource by URI (full message body and reactions).", { uri: z.string() }, async ({ uri }) => {
  const m = state.channelMessages.find((x) => x.uri === uri);
  if (!m) return fail(`Resource not found: ${uri}`);
  return text({ id: m.id, body: m._full, reactions: m._reactions });
});

/* ---------- background chatter so live polling has something to show ---------- */
const LINES = ["Quick sync after lunch?", "Pushed a fix, please take a look.", "👍", "Does anyone have the link to the doc?",
  "Running 5 min late.", "Nice!", "I'll pick this up tomorrow.", "Deploy went out fine."];
if (process.env.MOCK_CHATTER !== "0") {
  const ms = Number(process.env.MOCK_CHATTER_MS) || 45_000;
  setInterval(() => {
    const line = LINES[Math.floor(Math.random() * LINES.length)];
    const at = new Date().toISOString();
    if (Math.random() < 0.5) {
      const all = Object.values(channels).flat();
      const ch = all[Math.floor(Math.random() * all.length)];
      const others = people.filter((p) => p.id !== me.id);
      const who = others[Math.floor(Math.random() * others.length)];
      state.channelMessages.push(makeChannelMessage({ teamId: findTeamOfChannel(ch.id), channelId: ch.id, authorId: who.id, text: line, at }));
      log(`chatter: ${who.displayName} in #${ch.displayName}`);
    } else {
      const chat = state.chats[Math.floor(Math.random() * state.chats.length)];
      const others = chat.memberIds.filter((id) => id !== me.id);
      const who = others[Math.floor(Math.random() * others.length)];
      state.chatMessages.push(makeChatMessage({ chatId: chat.id, authorId: who, text: line, at }));
      log(`chatter: ${byId(who).displayName} in chat ${chat.topic || chat.id}`);
    }
  }, ms).unref();
}

await server.connect(new StdioServerTransport());
log("ready", FAIL.size ? `(failing: ${[...FAIL].join(", ")})` : "");
