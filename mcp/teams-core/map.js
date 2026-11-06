// Graph objects -> the shapes teams-slack-style.html reads (same as mcp/mock/fixtures.js). Pure: runs in Node and the browser.

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };
export function htmlToText(html) {
  return String(html || "")
    .replace(/<(br|\/p|\/div|\/li)[^>]*>/gi, "\n")
    .replace(/<attachment[^>]*>.*?<\/attachment>/gi, "")
    .replace(/<[^>]+>/g, "")
    .replace(/&(#x?[0-9a-f]+|\w+);/gi, (m, e) => {
      if (e[0] === "#") { const n = e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10); return Number.isFinite(n) ? String.fromCodePoint(n) : m; }
      return ENTITIES[e.toLowerCase()] ?? m;
    })
    .replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

export function channelUri(teamId, channelId, id, parentId) {
  const base = `teams:///teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/`;
  return base + (parentId ? `${encodeURIComponent(parentId)}/replies/${encodeURIComponent(id)}` : encodeURIComponent(id));
}
export function parseChannelUri(uri) {
  const m = /^teams:\/\/\/teams\/([^/]+)\/channels\/([^/]+)\/messages\/([^/]+)(?:\/replies\/([^/]+))?$/.exec(uri || "");
  if (!m) return null;
  const [teamId, channelId, a, b] = m.slice(1).map((x) => (x == null ? x : decodeURIComponent(x)));
  return b ? { teamId, channelId, parentId: a, id: b } : { teamId, channelId, id: a };
}

function sender(m) {
  if (m.messageType && m.messageType !== "message") return null; // system events: page hides from-less messages
  const u = m.from?.user, app = m.from?.application;
  if (u) return { displayName: u.displayName || "Unknown", userId: u.id };
  if (app) return { displayName: app.displayName || "App", userId: app.id };
  return null;
}

export function channelMsg(m, teamId, channelId) {
  const parentId = m.replyToId || null;
  return {
    id: m.id,
    replyToId: parentId,
    createdDateTime: m.createdDateTime,
    lastModifiedDateTime: m.lastModifiedDateTime || m.createdDateTime,
    deletedDateTime: m.deletedDateTime || null,
    subject: m.subject || null,
    from: sender(m),
    bodyContentType: m.body?.contentType === "html" ? "html" : "text",
    bodyPreview: m.body?.content || "",
    bodyTruncated: false,
    attachments: (m.attachments || []).filter((a) => a.name).map((a) => ({ name: a.name, contentUrl: a.contentUrl || null })),
    webUrl: m.webUrl || null,
    uri: channelUri(teamId, channelId, m.id, parentId),
  };
}

export function reactions(m) {
  return (m.reactions || []).map((r) => ({ reactionType: r.reactionType, displayName: r.displayName, user: r.user }));
}

export function chat(c) {
  const members = (c.members || []).map((x) => ({ displayName: x.displayName || x.email || "Unknown", email: x.email || null, userId: x.userId || null }));
  return {
    id: c.id,
    chatType: c.chatType,
    topic: c.topic || null,
    memberCount: members.length,
    members,
    lastActivity: c.lastMessagePreview?.createdDateTime || c.lastUpdatedDateTime || null,
    webUrl: c.webUrl || null,
  };
}

/** A chat message as a search hit (normHit in the page). `members` is the chat's mapped member list. */
export function chatHit(m, chatId, members) {
  const uid = m.from?.user?.id;
  const member = uid && members.find((x) => x.userId === uid);
  return {
    id: m.id,
    chatId,
    createdDateTime: m.createdDateTime,
    lastModifiedDateTime: m.lastModifiedDateTime || m.createdDateTime,
    subject: m.subject || null,
    from: { displayName: m.from?.user?.displayName || m.from?.application?.displayName || member?.displayName || "Unknown", email: member?.email || null },
    summary: m.body?.contentType === "html" ? htmlToText(m.body.content) : String(m.body?.content || ""),
    webUrl: `https://teams.microsoft.com/l/message/${encodeURIComponent(chatId)}/${encodeURIComponent(m.id)}?context=${encodeURIComponent('{"contextType":"chat"}')}`,
  };
}

export function person(p) {
  const mail = p.scoredEmailAddresses?.[0]?.address || p.userPrincipalName || null;
  return {
    id: p.id,
    displayName: p.displayName,
    jobTitle: p.jobTitle || null,
    mail,
    userPrincipalName: p.userPrincipalName || mail,
    personType: p.personType ? `${p.personType.class}/${p.personType.subclass}` : null,
  };
}
