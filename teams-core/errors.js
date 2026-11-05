// Shared error types for the Teams tools (no Node or browser dependencies).

/** Needs admin consent in most tenants. Tools degrade without it. */
export const CHANNEL_READ = "ChannelMessage.Read.All";

/** Token problems. code: "needs_reauth" | "consent_missing" | "not_configured". */
export class AuthError extends Error {
  // the default makes `scope` optional when this file is compiled as TypeScript (SharePoint build)
  constructor(code, message, scope = undefined) { super(message); this.code = code; this.scope = scope; }
}
