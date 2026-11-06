// Seed data for the mock Microsoft 365 MCP server.
// Field names follow what teams-slack-style.html reads (normChannel, normHit, chatTitle, people filter).

const MIN = 60_000;
const ago = (mins) => new Date(Date.now() - mins * MIN).toISOString();

export const DOMAIN = "contoso.example";

export const people = [
  { id: "u-me", displayName: "Alex Morgan", jobTitle: "Software Engineer" },
  { id: "u-priya", displayName: "Priya Shah", jobTitle: "Engineering Manager" },
  { id: "u-tom", displayName: "Tom Becker", jobTitle: "Product Designer" },
  { id: "u-lena", displayName: "Lena Okafor", jobTitle: "QA Lead" },
  { id: "u-sam", displayName: "Sam Rivera", jobTitle: "Data Scientist" },
  { id: "u-jo", displayName: "Jo Kim", jobTitle: "Product Manager" },
  { id: "u-dev", displayName: "Dev Patel", jobTitle: "Platform Engineer" },
  { id: "u-maria", displayName: "Maria Rossi", jobTitle: "Sales Director" },
].map((p) => {
  const upn = p.displayName.toLowerCase().replace(/\s+/g, ".") + "@" + DOMAIN;
  return { ...p, mail: upn, userPrincipalName: upn, personType: "Person/OrganizationUser" };
});

export const byId = (id) => people.find((p) => p.id === id);
export const me = people[0];

export const teams = [
  { id: "t-eng", displayName: "Engineering", description: "Everything engineering" },
  { id: "t-prod", displayName: "Product", description: "Roadmap and discovery" },
  { id: "t-social", displayName: "Social Club", description: "Lunches, games, and the occasional cake" },
];

export const channels = {
  "t-eng": [
    { id: "c-eng-general", displayName: "General", description: "Team-wide announcements" },
    { id: "c-eng-builds", displayName: "Builds", description: "CI and release chatter" },
    { id: "c-eng-incidents", displayName: "Incidents", description: "Live incidents and post-mortems" },
  ],
  "t-prod": [
    { id: "c-prod-general", displayName: "General", description: "Product team" },
    { id: "c-prod-research", displayName: "User Research", description: "Interview notes and insights" },
  ],
  "t-social": [
    { id: "c-soc-general", displayName: "General", description: "Say hi" },
  ],
};

// [channelId, minutesAgo, authorId, text(html allowed), replies?: [[minutesAgo, authorId, text]]]
const seedChannel = [
  ["c-eng-general", 60 * 26, "u-priya", "Welcome to the new Engineering team space! Please keep announcements here.", [
    [60 * 25, "u-dev", "Thanks Priya 🙌"],
    [60 * 25 - 10, "u-me", "Great, will move the old wiki links over."],
  ]],
  ["c-eng-general", 60 * 5, "u-dev", "Reminder: <b>platform upgrade</b> is scheduled for Thursday 18:00 UTC.", []],
  ["c-eng-general", 90, "u-lena", "QA sign-off for 4.2 is done. Release notes are <a href=\"https://example.com/notes\">here</a>.", [
    [80, "u-priya", "Brilliant, thank you!"],
  ]],
  ["c-eng-general", 12, "u-sam", "Anyone know why the staging dashboard is slow today?", []],
  ["c-eng-builds", 60 * 30, "u-dev", "Nightly build #1412 failed: flaky test in <code>auth.spec.ts</code>.", [
    [60 * 29, "u-me", "Looking into it."],
    [60 * 28, "u-me", "Fixed by retrying the token refresh mock. PR is up."],
  ]],
  ["c-eng-builds", 60 * 3, "u-dev", "Build #1420 is green ✅", []],
  ["c-eng-builds", 40, "u-me", "Bumped Node to 24 on the CI runners.", []],
  ["c-eng-incidents", 60 * 48, "u-priya", "<b>INC-221</b>: elevated API latency in EU region. Resolved after cache node restart. Post-mortem to follow.", [
    [60 * 47, "u-dev", "Root cause was a memory leak in the session cache."],
  ]],
  ["c-prod-general", 60 * 20, "u-jo", "Q1 roadmap draft is ready for comments.", [
    [60 * 19, "u-tom", "Left some notes on the onboarding section."],
    [60 * 18, "u-maria", "Sales would love the export feature earlier if possible."],
  ]],
  ["c-prod-general", 55, "u-tom", "New design system tokens are live in Figma.", []],
  ["c-prod-research", 60 * 72, "u-jo", "Summary of last week's five customer interviews: users want <i>faster search</i> and <i>offline mode</i>.", []],
  ["c-soc-general", 60 * 6, "u-maria", "Cake in the kitchen for Lena's birthday 🎂", [
    [60 * 6 - 5, "u-lena", "Thank you all!"],
    [60 * 6 - 3, "u-sam", "Happy birthday!"],
  ]],
  ["c-soc-general", 30, "u-sam", "Board games Friday? Thinking 17:30.", []],
];

export const chats = [
  { id: "19:u-me_u-priya@unq.gbl.spaces", chatType: "oneOnOne", topic: null, memberIds: ["u-me", "u-priya"] },
  { id: "19:u-me_u-tom@unq.gbl.spaces", chatType: "oneOnOne", topic: null, memberIds: ["u-me", "u-tom"] },
  { id: "19:u-me_u-lena@unq.gbl.spaces", chatType: "oneOnOne", topic: null, memberIds: ["u-me", "u-lena"] },
  { id: "19:launch-crew@thread.v2", chatType: "group", topic: "Launch crew", memberIds: ["u-me", "u-jo", "u-dev", "u-maria"] },
  { id: "19:lunch@thread.v2", chatType: "group", topic: null, memberIds: ["u-me", "u-sam", "u-tom"] },
];

// [chatId, minutesAgo, authorId, text]
const seedChat = [
  ["19:u-me_u-priya@unq.gbl.spaces", 60 * 24, "u-priya", "Can we move our 1:1 to Wednesday?"],
  ["19:u-me_u-priya@unq.gbl.spaces", 60 * 24 - 5, "u-me", "Sure, Wednesday 14:00 works."],
  ["19:u-me_u-priya@unq.gbl.spaces", 25, "u-priya", "Also, nice work on the build fix yesterday."],
  ["19:u-me_u-tom@unq.gbl.spaces", 60 * 8, "u-tom", "Got a minute to review the new settings page mock?"],
  ["19:u-me_u-tom@unq.gbl.spaces", 60 * 8 - 20, "u-me", "Yes! Looks clean. Maybe bigger tap targets on mobile."],
  ["19:u-me_u-lena@unq.gbl.spaces", 60 * 50, "u-lena", "Found a regression in the export flow, ticket QA-88."],
  ["19:u-me_u-lena@unq.gbl.spaces", 60 * 49, "u-me", "On it, thanks."],
  ["19:launch-crew@thread.v2", 60 * 3, "u-jo", "Launch checklist: docs ✅ pricing ✅ blog post ⏳"],
  ["19:launch-crew@thread.v2", 60 * 3 - 10, "u-maria", "Customers are asking about the date, can we confirm Tuesday?"],
  ["19:launch-crew@thread.v2", 60 * 2, "u-dev", "Infra is ready for Tuesday."],
  ["19:launch-crew@thread.v2", 15, "u-me", "Blog post draft is in review."],
  ["19:lunch@thread.v2", 60 * 2, "u-sam", "Ramen or tacos?"],
  ["19:lunch@thread.v2", 60 * 2 - 3, "u-tom", "Tacos 🌮"],
];

let seq = 0;
export const nextId = (prefix) => `${prefix}-${Date.now().toString(36)}-${(seq++).toString(36)}`;

export function makeChannelMessage({ teamId, channelId, authorId, text, at, parentId = null }) {
  const author = byId(authorId);
  const id = nextId("m");
  const isHtml = /<[a-z][\s\S]*>/i.test(text);
  const preview = text.length > 160 ? text.slice(0, 160) : text;
  return {
    id,
    replyToId: parentId,
    teamId,
    channelId,
    createdDateTime: at,
    lastModifiedDateTime: at,
    deletedDateTime: null,
    subject: null,
    from: { displayName: author.displayName, userId: author.id },
    bodyContentType: isHtml ? "html" : "text",
    bodyPreview: preview,
    bodyTruncated: preview.length < text.length,
    attachments: [],
    webUrl: `https://teams.example/l/message/${channelId}/${id}`,
    uri: `teams:///teams/${teamId}/channels/${channelId}/messages/${id}`,
    _full: { contentType: isHtml ? "html" : "text", content: text },
    _reactions: [],
  };
}

export function makeChatMessage({ chatId, authorId, text, at }) {
  const author = byId(authorId);
  const id = nextId("cm");
  return {
    id,
    chatId,
    createdDateTime: at,
    lastModifiedDateTime: at,
    subject: null,
    from: { displayName: author.displayName, email: author.mail },
    summary: text,
    webUrl: `https://teams.example/l/chat/${encodeURIComponent(chatId)}/${id}`,
  };
}

export function seed() {
  const channelMessages = [];
  for (const [channelId, mins, authorId, text, replies] of seedChannel) {
    const teamId = Object.keys(channels).find((t) => channels[t].some((c) => c.id === channelId));
    const root = makeChannelMessage({ teamId, channelId, authorId, text, at: ago(mins) });
    channelMessages.push(root);
    for (const [rm, ra, rt] of replies) {
      channelMessages.push(makeChannelMessage({ teamId, channelId, authorId: ra, text: rt, at: ago(rm), parentId: root.id }));
    }
  }
  // a couple of reactions so read_resource has something to show
  const lena = channelMessages.find((m) => m.from.userId === "u-lena" && m.channelId === "c-eng-general");
  if (lena) lena._reactions.push(
    { reactionType: "like", user: { user: { id: "u-priya", displayName: "Priya Shah" } } },
    { reactionType: "heart", user: { user: { id: "u-me", displayName: "Alex Morgan" } } });

  const chatMessages = seedChat.map(([chatId, mins, authorId, text]) => makeChatMessage({ chatId, authorId, text, at: ago(mins) }));
  return { channelMessages, chatMessages, chats: chats.map((c) => ({ ...c })) };
}
