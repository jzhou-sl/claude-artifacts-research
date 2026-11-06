// Teams tool handlers over Microsoft Graph, environment-neutral (Node MCP server and SharePoint web part).
// createTools({ graph, graphPages, messages? }) -> { [name]: { description, annotations, handler(args) -> CallToolResult } }
// Input validation lives with the caller (zod in mcp/teams/server.js); handlers apply the same defaults.
import { AuthError, CHANNEL_READ } from "./errors.js";
import { GraphError, seg } from "./graph.js";
import * as map from "./map.js";

/* ---------- scopes per area ---------- */
export const SCOPES = {
  me: ["User.Read"],
  people: ["People.Read"],
  teams: ["Team.ReadBasic.All"],
  channels: ["Channel.ReadBasic.All"],
  channelRead: [CHANNEL_READ],
  channelSend: ["ChannelMessage.Send"],
  chats: ["Chat.ReadWrite"],
};

/** Host-specific wording for errors the user has to act on. */
export const NODE_MESSAGES = {
  graph401: "Microsoft Graph rejected the sign-in. Run `node mcp/teams/login.js`, then reload.",
  graph403Hint: "Check the app registration's API permissions and admin consent.",
};

/* ---------- result helpers ---------- */
const text = (...values) => ({ content: [{ type: "text", text: values.map((v) => JSON.stringify(v)).join("\n") }] });
const fail = (code, message, extra = {}) => ({
  isError: true,
  content: [{ type: "text", text: message }],
  structuredContent: { code, message, ...extra },
});

function ttlCache(ttlMs) {
  const m = new Map();
  return {
    async get(key, load) {
      const hit = m.get(key);
      if (hit && Date.now() - hit.at < ttlMs) return hit.value;
      if (m.size > 500) for (const [k, v] of m) if (Date.now() - v.at >= ttlMs) m.delete(k);
      const value = load();          // store the promise so concurrent callers share one request
      m.set(key, { at: Date.now(), value });
      try { return await value; } catch (e) { m.delete(key); throw e; }
    },
    deletePrefix(prefix) { for (const k of m.keys()) if (k.startsWith(prefix)) m.delete(k); },
    clear() { m.clear(); },
  };
}

async function pool(items, n, fn) {
  const out = new Array(items.length); let i = 0;
  await Promise.all(Array.from({ length: Math.min(n, items.length) }, async () => { while (i < items.length) { const k = i++; out[k] = await fn(items[k]); } }));
  return out;
}

export function createTools({ graph, graphPages, messages = NODE_MESSAGES, log = (...a) => console.error("[teams]", ...a) }) {
  const S = SCOPES;

  function toFailure(e) {
    if (e instanceof AuthError) {
      if (e.code === "needs_reauth") return fail("needs_reauth", e.message);
      return fail("tool_error", e.message); // not_configured / consent_missing: page shows the message verbatim
    }
    if (e instanceof GraphError) {
      if (e.status === 401) return fail("needs_reauth", messages.graph401);
      if (e.status === 403) return fail("tool_error", `Permission denied by Microsoft Graph (${e.code}): ${e.message} ${messages.graph403Hint}`);
      if ([429, 503, 504].includes(e.status)) return fail("server_unavailable", `Microsoft Graph is busy (${e.status}). ${e.message}`, { retryable: true, retryAfterMs: e.retryAfterMs });
      if (e.status >= 500) return fail("upstream_error", `Microsoft Graph error ${e.status}: ${e.message}`, { retryable: true, retryAfterMs: 1500 });
      return fail("tool_error", `${e.message} (${e.code})`);
    }
    log("unexpected error", e);
    return fail("upstream_error", String(e?.message || e), { retryable: true, retryAfterMs: 1500 });
  }

  const tools = {};
  function tool(name, description, handler, annotations = {}) {
    tools[name] = { description, annotations, handler: async (args = {}) => { try { return await handler(args); } catch (e) { return toFailure(e); } } };
  }

  /* ---------- caches ---------- */
  const meCache = ttlCache(10 * 60_000);
  const chatsCache = ttlCache(60_000);
  const chatMsgCache = ttlCache(3_000);

  // read_resource serves channel messages already seen in list calls (LRU by uri)
  const seen = new Map();
  function remember(uri, raw) { seen.delete(uri); seen.set(uri, raw); if (seen.size > 2000) seen.delete(seen.keys().next().value); }

  const getMe = () => meCache.get("me", () => graph("GET", "/me", { scopes: S.me, query: { $select: "id,displayName,mail,userPrincipalName,jobTitle" } }));

  /* ---------- profile & people ---------- */
  tool("get_me", "Get the signed-in user's profile.", async () => {
    const me = await getMe();
    return text({ ...me, mail: me.mail || me.userPrincipalName });
  }, { readOnlyHint: true });

  tool("search_people", "Search people relevant to the user by name or email.", async ({ query, limit = 15 }) => {
    const data = await graph("GET", "/me/people", { scopes: S.people, query: { $search: `"${String(query).replace(/"/g, "")}"`, $top: limit } });
    return text((data?.value || []).map(map.person));
  }, { readOnlyHint: true });

  /* ---------- teams & channels ---------- */
  tool("teams_list_teams", "List teams the user has joined.", async () => {
    const rows = await graphPages("/me/joinedTeams", { scopes: S.teams, query: { $select: "id,displayName,description" } });
    return text(rows);
  }, { readOnlyHint: true });

  tool("teams_list_channels", "List channels in a team.", async ({ teamId }) => {
    const rows = await graphPages(`/teams/${seg(teamId)}/channels`, { scopes: S.channels, query: { $select: "id,displayName,description,webUrl,membershipType" } });
    return text(rows);
  }, { readOnlyHint: true });

  tool("teams_list_channel_messages",
    `List messages in a channel (newest first), or replies to a message when parentMessageId is set. Needs ${CHANNEL_READ}.`,
    async ({ teamId, channelId, parentMessageId, limit = 20, cursor }) => {
      const base = `/teams/${seg(teamId)}/channels/${seg(channelId)}/messages` + (parentMessageId ? `/${seg(parentMessageId)}/replies` : "");
      const data = cursor
        ? await graph("GET", cursor, { scopes: S.channelRead })
        : await graph("GET", base, { scopes: S.channelRead, query: { $top: Math.min(limit, 50) } });
      const rows = (data?.value || [])
        .map((m) => { const r = map.channelMsg(m, teamId, channelId); remember(r.uri, m); return r; })
        .sort((a, b) => (a.createdDateTime < b.createdDateTime ? 1 : -1));
      const next = data?.["@odata.nextLink"] || null;
      return text(rows, { moreResults: !!next, nextCursor: next });
    }, { readOnlyHint: true });

  tool("teams_send_channel_message", "Post a new message to a channel as the signed-in user.",
    async ({ teamId, channelId, body, bodyType = "text" }) => {
      const m = await graph("POST", `/teams/${seg(teamId)}/channels/${seg(channelId)}/messages`,
        { scopes: S.channelSend, body: { body: { contentType: bodyType, content: body } } });
      return text({ id: m.id, createdDateTime: m.createdDateTime, webUrl: m.webUrl });
    }, { destructiveHint: false, openWorldHint: true });

  tool("teams_reply_channel_message", "Reply to a channel message as the signed-in user.",
    async ({ teamId, channelId, messageId, body, bodyType = "text" }) => {
      const m = await graph("POST", `/teams/${seg(teamId)}/channels/${seg(channelId)}/messages/${seg(messageId)}/replies`,
        { scopes: S.channelSend, body: { body: { contentType: bodyType, content: body } } });
      return text({ id: m.id, createdDateTime: m.createdDateTime, webUrl: m.webUrl });
    }, { destructiveHint: false, openWorldHint: true });

  /* ---------- chats ---------- */
  async function fetchChatsPage(urlOrNull, top) {
    if (urlOrNull) return graph("GET", urlOrNull, { scopes: S.chats });
    try {
      return await graph("GET", "/me/chats", { scopes: S.chats,
        query: { $expand: "members,lastMessagePreview", $orderby: "lastMessagePreview/createdDateTime desc", $top: top } });
    } catch (e) {
      if (!(e instanceof GraphError) || e.status !== 400) throw e;
      // some tenants reject combining expands with $orderby: fall back and sort locally
      const d = await graph("GET", "/me/chats", { scopes: S.chats, query: { $expand: "members", $top: top } });
      d.value = (d.value || []).sort((a, b) => ((a.lastUpdatedDateTime || "") < (b.lastUpdatedDateTime || "") ? 1 : -1));
      return d;
    }
  }

  /** Recent chats (up to 200), cached for 60 s. Used by chat_message_search. */
  const recentChats = () => chatsCache.get("recent", async () => {
    const out = []; let next = null;
    for (let i = 0; i < 4; i++) {
      const d = await fetchChatsPage(next, 50);
      out.push(...(d?.value || []).map(map.chat));
      next = d?.["@odata.nextLink"]; if (!next) break;
    }
    return out;
  });

  tool("teams_list_chats", "List the user's chats, most recently active first.", async ({ limit = 25, cursor }) => {
    const d = await fetchChatsPage(cursor || null, Math.min(limit, 50));
    const next = d?.["@odata.nextLink"] || null;
    return text((d?.value || []).map(map.chat), { moreResults: !!next, nextCursor: next });
  }, { readOnlyHint: true });

  tool("teams_send_chat_message", "Send a message to a chat as the signed-in user.",
    async ({ chatId, body, bodyType = "text" }) => {
      const m = await graph("POST", `/chats/${seg(chatId)}/messages`, { scopes: S.chats, body: { body: { contentType: bodyType, content: body } } });
      chatMsgCache.deletePrefix(chatId + "|");
      return text({ id: m.id, chatId, createdDateTime: m.createdDateTime });
    }, { destructiveHint: false, openWorldHint: true });

  tool("teams_create_chat", "Create a chat with the given members (UPNs or user ids). For oneOnOne, returns the existing chat if there is one.",
    async ({ members, chatType = "oneOnOne", topic }) => {
      const me = await getMe();
      const others = members.filter((m) => m.toLowerCase() !== String(me.userPrincipalName).toLowerCase() && m !== me.id);
      if (chatType === "oneOnOne" && others.length !== 1) return fail("bad_request", "A oneOnOne chat needs exactly one other member.");
      const bind = (who) => ({ "@odata.type": "#microsoft.graph.aadUserConversationMember", roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${String(who).replace(/'/g, "''")}')` });
      const c = await graph("POST", "/chats", { scopes: S.chats,
        body: { chatType, ...(chatType === "group" && topic ? { topic } : {}), members: [bind(me.id), ...others.map(bind)] } });
      chatsCache.clear();
      return text({ id: c.id, chatType: c.chatType, webUrl: c.webUrl });
    }, { destructiveHint: false, openWorldHint: true });

  /** Messages in one chat newer than `after` (ISO), newest first; cached for 3 s. */
  function chatMessagesSince(chatId, after) {
    return chatMsgCache.get(chatId + "|" + after, () => graphPages(`/chats/${seg(chatId)}/messages`,
      { scopes: S.chats, query: { $top: 50, $orderby: "createdDateTime desc" } },
      { max: 4, keepGoing: (items) => items.length > 0 && items[items.length - 1].createdDateTime >= after }));
  }

  tool("chat_message_search",
    "Search the user's recent chat messages. query '*' matches everything. Filters: sender/recipient (email), afterDateTime. Covers the ~200 most recent chats.",
    async ({ query, limit = 25, offset = 0, afterDateTime, sender, recipient }) => {
      const me = await getMe();
      const myMail = String(me.mail || me.userPrincipalName).toLowerCase();
      const after = new Date(afterDateTime || Date.now() - 30 * 864e5).toISOString();
      const lc = (s) => String(s || "").toLowerCase();
      const has = (c, email) => c.members.some((m) => lc(m.email) === lc(email));

      let candidates = (await recentChats()).filter((c) => !c.lastActivity || c.lastActivity >= after);
      if (sender && lc(sender) !== myMail) candidates = candidates.filter((c) => has(c, sender));
      if (recipient) candidates = candidates.filter((c) => has(c, recipient));
      candidates = candidates.slice(0, 12); // most recently active first

      const q = String(query).trim().toLowerCase();
      const perChat = await pool(candidates, 4, async (c) => {
        const msgs = await chatMessagesSince(c.id, after);
        return msgs
          .filter((m) => m.messageType === "message" && !m.deletedDateTime && m.createdDateTime >= after)
          .map((m) => map.chatHit(m, c.id, c.members))
          .filter((h) => (!sender || lc(h.from.email) === lc(sender)) && (!q || q === "*" || h.summary.toLowerCase().includes(q)));
      });
      const rows = perChat.flat().sort((a, b) => (a.createdDateTime < b.createdDateTime ? 1 : -1));
      const slice = rows.slice(offset, offset + limit);
      const next = offset + slice.length;
      return text(slice, { moreResults: next < rows.length, nextOffset: next });
    }, { readOnlyHint: true });

  /* ---------- resources: full channel message body + reactions ---------- */
  tool("read_resource", "Read a Teams channel message by the uri returned from teams_list_channel_messages (full body and reactions).",
    async ({ uri }) => {
      const ref = map.parseChannelUri(uri);
      if (!ref) return fail("bad_request", `Unsupported resource uri: ${uri}`);
      let m = seen.get(uri);
      if (!m) {
        const p = `/teams/${seg(ref.teamId)}/channels/${seg(ref.channelId)}/messages/` + (ref.parentId ? `${seg(ref.parentId)}/replies/${seg(ref.id)}` : seg(ref.id));
        m = await graph("GET", p, { scopes: S.channelRead });
        remember(uri, m);
      }
      return text({ id: m.id, body: { contentType: m.body?.contentType || "text", content: m.body?.content || "" }, reactions: map.reactions(m) });
    }, { readOnlyHint: true });

  return tools;
}
