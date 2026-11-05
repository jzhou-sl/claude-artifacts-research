// SharePoint wording for errors whose default text points at claude.ai or `npm run login`.
// Shared by the web part and the local harness.

export const ADMIN_HINT = "A SharePoint admin needs to approve this app’s Microsoft Graph permissions (SharePoint admin center → Advanced → API access).";

const UNREACHABLE = "This page couldn’t reach Microsoft 365. Reload the page, and contact the intranet team if it keeps happening.";

/** window.claude.help: overrides for the page's explain(), keyed by error code. */
export const HELP = {
  needs_reauth: "Your Microsoft 365 sign-in has lapsed. Reload this page to sign in again.",
  server_not_connected: UNREACHABLE,
  not_granted: UNREACHABLE,
  capability_disabled: UNREACHABLE,
};

/** teams-core createTools({ messages }). */
export const MESSAGES = { graph401: HELP.needs_reauth, graph403Hint: ADMIN_HINT };
