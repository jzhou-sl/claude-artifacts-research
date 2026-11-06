// Minimal Microsoft Graph client: bearer auth, JSON, one retry on throttling, safe nextLink handling.
// Environment-neutral: the caller supplies getToken(scopes) (MSAL Node locally, SPFx in SharePoint).

export const GRAPH = "https://graph.microsoft.com/v1.0";

export class GraphError extends Error {
  constructor(status, code, message, retryAfterMs) { super(message); this.status = status; this.code = code; this.retryAfterMs = retryAfterMs; }
}

/** Encode one path segment. Keeps ':' and '@' readable (Teams ids like 19:abc@thread.v2). */
export const seg = (s) => encodeURIComponent(String(s)).replace(/%3A/gi, ":").replace(/%40/g, "@");

/** Accept a nextLink back from the client only if it points at Graph v1.0. */
export function checkNextLink(url) {
  if (typeof url !== "string" || !url.startsWith(GRAPH + "/")) throw new GraphError(400, "bad_cursor", "Invalid cursor.");
  return url;
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * const { graph, graphPages } = createGraph(getToken)
 * graph("GET", "/me", { scopes, query: { $select: "id" } })
 * graph("GET", nextLink, { scopes })            // absolute Graph URL
 * graph("POST", "/chats/x/messages", { scopes, body })
 * graph("GET", "/me/photo/$value", { scopes, binary: true })   // -> { contentType, bytes: Uint8Array }
 */
export function createGraph(getToken) {
  async function graph(method, pathOrUrl, { scopes, query, body, headers, binary = false } = {}) {
    let url = pathOrUrl.startsWith("https://") ? checkNextLink(pathOrUrl) : GRAPH + pathOrUrl;
    if (query) {
      const qs = Object.entries(query).filter(([, v]) => v != null && v !== "").map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join("&");
      if (qs) url += (url.includes("?") ? "&" : "?") + qs;
    }
    for (let attempt = 0; ; attempt++) {
      const token = await getToken(scopes);
      const res = await fetch(url, {
        method,
        headers: { authorization: `Bearer ${token}`, accept: binary ? "*/*" : "application/json", ...(body ? { "content-type": "application/json" } : {}), ...headers },
        body: body ? JSON.stringify(body) : undefined,
      });
      if (res.status === 204) return null;
      if (binary && res.ok) {
        return { contentType: res.headers.get("content-type") || "application/octet-stream", bytes: new Uint8Array(await res.arrayBuffer()) };
      }
      const text = await res.text();
      let data = null;
      try { data = text ? JSON.parse(text) : null; } catch { data = { raw: text }; }
      if (res.ok) return data;

      const retryAfterMs = (Number(res.headers.get("retry-after")) || 2) * 1000;
      if ([429, 503, 504].includes(res.status) && attempt === 0 && retryAfterMs <= 10_000) { await sleep(retryAfterMs); continue; }
      const err = data?.error || {};
      throw new GraphError(res.status, err.code || `http_${res.status}`, err.message || `Graph returned ${res.status}`, retryAfterMs);
    }
  }

  /** Follow nextLinks until `max` pages or `keepGoing(items)` returns false. */
  async function graphPages(path, opts, { max = 4, keepGoing = () => true } = {}) {
    const out = [];
    let next = path, first = true;
    for (let i = 0; next && i < max; i++) {
      const data = await graph("GET", next, first ? opts : { scopes: opts.scopes });
      first = false;
      const items = data?.value || [];
      out.push(...items);
      next = data?.["@odata.nextLink"] || null;
      if (!keepGoing(items)) break;
    }
    return out;
  }

  return { graph, graphPages };
}
