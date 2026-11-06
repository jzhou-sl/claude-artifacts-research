// profile-mcp's tools over Microsoft Graph, environment-neutral (Node MCP server and SharePoint web part).
// createProfileTools({ graph, messages? }) -> { get_me, get_my_photo: { description, annotations, handler() -> CallToolResult } }
// get_me has the same result shape as teams-mcp's get_me. get_my_photo returns an MCP image block
// ({ type: "image", data: base64, mimeType }), or only a text block when the user has no photo.
// Failures carry a window.claude error code in structuredContent.
import { AuthError } from "../teams-core/errors.js";
import { GraphError } from "../teams-core/graph.js";

export const SCOPES = ["User.Read"];

const SELECT = "id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,preferredLanguage";

/** Host-specific wording for errors the user has to act on. */
export const NODE_MESSAGES = {
  graph401: "Microsoft Graph rejected the sign-in. Run `node --env-file-if-exists=.env mcp/profile/login.mjs`, then retry.",
  graph403Hint: "Check the app registration has delegated User.Read.",
};

/** Base64 without Buffer, so it also runs in the browser. Chunked to stay under argument-count limits. */
function toBase64(bytes) {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

const fail = (code, message, extra = {}) => ({ isError: true, content: [{ type: "text", text: message }], structuredContent: { code, message, ...extra } });

export function createProfileTools({ graph, messages = NODE_MESSAGES, log = (...a) => console.error("[profile]", ...a) }) {
  // Same error mapping as mcp/teams-core/tools.js, so resultError() and the dev-server bridge treat both servers alike.
  function toFailure(e) {
    if (e instanceof AuthError) return fail(e.code === "needs_reauth" ? "needs_reauth" : "tool_error", e.message);
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

  return {
    get_me: {
      description: "Get the signed-in user's Microsoft 365 profile (name, email, job title, department, office, phones).",
      annotations: { readOnlyHint: true },
      async handler() {
        try {
          const { "@odata.context": _, ...me } = await graph("GET", "/me", { scopes: SCOPES, query: { $select: SELECT } });
          return { content: [{ type: "text", text: JSON.stringify({ ...me, mail: me.mail || me.userPrincipalName }) }] };
        } catch (e) {
          return toFailure(e);
        }
      },
    },
    get_my_photo: {
      description: "Get the signed-in user's Microsoft 365 profile photo as an image. Returns only a text block if no photo is set.",
      annotations: { readOnlyHint: true },
      async handler() {
        try {
          const { contentType, bytes } = await graph("GET", "/me/photo/$value", { scopes: SCOPES, binary: true });
          return { content: [{ type: "image", data: toBase64(bytes), mimeType: contentType.split(";")[0].trim() }] };
        } catch (e) {
          // Graph answers 404 (ImageNotFound / ErrorItemNotFound) when the user has no photo: not a failure.
          if (e instanceof GraphError && e.status === 404) return { content: [{ type: "text", text: "No profile photo is set." }] };
          return toFailure(e);
        }
      },
    },
  };
}
