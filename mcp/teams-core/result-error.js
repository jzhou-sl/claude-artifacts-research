// Turns an isError CallToolResult into the window.claude error shape { code, message, retryable, retryAfterMs }.
// Shared by dev-server.js (MCP bridge) and the SharePoint web part (in-browser transport).

export function resultError(result) {
  const msg = (result.content || []).filter((b) => b.type === "text").map((b) => b.text).join(" ").slice(0, 500);
  // servers may report a window.claude error code directly: structuredContent { code, message, retryable, retryAfterMs }
  const sc = result.structuredContent;
  if (sc && typeof sc.code === "string") {
    return { code: sc.code, message: String(sc.message || msg), retryable: !!sc.retryable, ...(sc.retryAfterMs ? { retryAfterMs: Number(sc.retryAfterMs) } : {}) };
  }
  // the MCP SDK reports invalid arguments / unknown tools as isError results carrying the JSON-RPC code
  if (/^MCP error -32602\b/.test(msg)) return { code: "bad_request", message: msg, retryable: false };
  return { code: "tool_error", message: msg || "Tool reported an error", retryable: false };
}
