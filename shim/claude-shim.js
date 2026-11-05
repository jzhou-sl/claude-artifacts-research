// window.claude shim for running artifact pages locally.
// Implements the "mcp" capability on top of the dev server's /__mcp/call bridge:
//   callTool(server, tool, input, opts)       -> Promise<CallToolResult & { cache: { storedAt } }>
//   watchTool(server, tool, input, cb, opts)  -> unsubscribe()
//   invalidate(server, tool)                  -> Promise<void>
// Errors are rejected/emitted as plain objects: { code, message, retryable, retryAfterMs }.
//
// As a plain <script> (dev-server.js) it installs itself on window, calling the dev server.
// The SharePoint web part instead calls installClaudeShim(iframeWindow, transport, { help }) with an
// in-browser transport(server, tool, input) -> Promise<CallToolResult> that throws the same error shape.
function installClaudeShim(window, transport, options) {
  "use strict";
  var document = window.document;
  // timers belong to the target window, so they stop when an embedding iframe is removed
  var setTimeout = function (f, ms) { return window.setTimeout(f, ms); };
  var clearTimeout = function (t) { window.clearTimeout(t); };
  if (window.claude && !window.claude.__shim) return; // a real host is present: leave it alone
  options = options || {};

  var ENDPOINT = "/__mcp/call";
  var CACHE_TTL_MS = 30000;        // default reads reuse results this fresh
  var WATCH_INTERVAL_MS = 10000;   // watchTool poll interval while the tab is visible
  var WATCH_MAX_BACKOFF_MS = 60000;
  var WRITE_TOOL = /(^|_)(send|reply|create|update|delete|post|set|move|rename|upload|forward|respond|trash|modify)(_|$)/;

  /* ---------- utilities ---------- */
  function stable(v) {
    if (v === null || typeof v !== "object") return JSON.stringify(v);
    if (Array.isArray(v)) return "[" + v.map(stable).join(",") + "]";
    return "{" + Object.keys(v).sort().map(function (k) { return JSON.stringify(k) + ":" + stable(v[k]); }).join(",") + "}";
  }
  function keyOf(server, tool, input) { return server + "\u0000" + tool + "\u0000" + stable(input || {}); }
  function prefixOf(server, tool) { return server + "\u0000" + tool + "\u0000"; }
  function err(code, message, extra) {
    var e = { code: code, message: message || code, retryable: false };
    if (extra) for (var k in extra) e[k] = extra[k];
    return e;
  }
  function stamp(result, storedAt, hit) {
    var out = {}; for (var k in result) out[k] = result[k];
    out.cache = { storedAt: storedAt, hit: !!hit };
    return out;
  }

  /* ---------- transport ---------- */
  async function fetchCall(server, tool, input) {
    var res;
    try {
      res = await fetch(ENDPOINT, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ server: server, tool: tool, input: input || {} }),
      });
    } catch (e) {
      throw err("server_unavailable", "Dev server unreachable: " + (e && e.message || e), { retryable: true, retryAfterMs: 1000 });
    }
    var body = null;
    try { body = await res.json(); } catch (e) { /* non-JSON */ }
    if (body && body.ok) return body.result;
    if (body && body.error) throw body.error;
    if (res.status >= 500) throw err("server_unavailable", "Dev server error " + res.status, { retryable: true, retryAfterMs: 1000 });
    throw err("upstream_error", "Unexpected response " + res.status);
  }
  async function rawCall(server, tool, input) {
    if (!transport) return fetchCall(server, tool, input);
    try { return await transport(server, tool, input || {}); }
    catch (e) { throw e && typeof e.code === "string" ? e : err("upstream_error", String(e && e.message || e), { retryable: true, retryAfterMs: 800 }); }
  }

  /* ---------- cache + in-flight de-duplication ---------- */
  var cache = new Map();    // key -> { result, storedAt(ms) }
  var inflight = new Map(); // key -> Promise<{ result, storedAt }>

  function fetchFresh(server, tool, input, key, store) {
    if (inflight.has(key)) return inflight.get(key);
    var p = rawCall(server, tool, input).then(function (result) {
      var entry = { result: result, storedAt: Date.now() };
      if (store) cache.set(key, entry);
      return entry;
    });
    inflight.set(key, p);
    p.then(function () { inflight.delete(key); }, function () { inflight.delete(key); });
    return p;
  }

  async function callTool(server, tool, input, opts) {
    if (typeof server !== "string" || typeof tool !== "string") throw err("bad_request", "server and tool must be strings");
    opts = opts || {};
    var isWrite = WRITE_TOOL.test(tool);
    var key = keyOf(server, tool, input);
    var mode = isWrite || opts.cache === false ? "bypass" : (opts.cache && opts.cache.refresh) ? "refresh" : "default";

    if (mode === "default") {
      var hit = cache.get(key);
      if (hit && Date.now() - hit.storedAt < CACHE_TTL_MS) return stamp(hit.result, new Date(hit.storedAt).toISOString(), true);
    }
    var entry = mode === "bypass"
      ? { result: await rawCall(server, tool, input), storedAt: Date.now() }
      : await fetchFresh(server, tool, input, key, true);
    if (mode === "refresh") notifyWatchers(key, entry);
    return stamp(entry.result, new Date(entry.storedAt).toISOString(), false);
  }

  /* ---------- watchers (polling; MCP has no push for tool results) ---------- */
  var watchers = new Set();

  function notifyWatchers(key, entry) {
    watchers.forEach(function (w) { if (w.key === key && !w.dead) w.emitData(entry); });
  }

  function watchTool(server, tool, input, cb, opts) {
    if (typeof cb !== "function") throw err("bad_request", "watchTool needs a callback");
    opts = opts || {};
    var interval = opts.intervalMs || WATCH_INTERVAL_MS;
    var w = { key: keyOf(server, tool, input), prefix: prefixOf(server, tool), dead: false, timer: null, delay: interval, busy: false };

    w.emitData = function (entry) {
      if (w.dead) return;
      try { cb({ type: "data", result: stamp(entry.result, new Date(entry.storedAt).toISOString(), false) }); } catch (e) { console.error(e); }
    };
    function schedule(ms) { clearTimeout(w.timer); if (!w.dead) w.timer = setTimeout(tick, ms); }
    async function tick() {
      if (w.dead) return;
      if (document.hidden || w.busy) { schedule(interval); return; }
      w.busy = true;
      try {
        var entry = await fetchFresh(server, tool, input, w.key, true);
        w.delay = interval;
        w.emitData(entry);
      } catch (e) {
        w.delay = Math.min(w.delay * 2, WATCH_MAX_BACKOFF_MS);
        if (!w.dead) try { cb({ type: "error", error: e }); } catch (x) { console.error(x); }
      }
      w.busy = false;
      schedule(w.delay);
    }
    w.kick = function () { schedule(0); };
    watchers.add(w);

    // first emission: serve a fresh-enough cached value immediately, then poll
    var hit = cache.get(w.key);
    if (hit && Date.now() - hit.storedAt < CACHE_TTL_MS) { setTimeout(function () { w.emitData(hit); }, 0); schedule(interval); }
    else schedule(0);

    return function unsubscribe() { w.dead = true; clearTimeout(w.timer); watchers.delete(w); };
  }

  document.addEventListener("visibilitychange", function () {
    if (!document.hidden) watchers.forEach(function (w) { w.kick(); });
  });

  async function invalidate(server, tool) {
    var prefix = tool == null ? server + "\u0000" : prefixOf(server, tool);
    Array.from(cache.keys()).forEach(function (k) { if (k.indexOf(prefix) === 0) cache.delete(k); });
    watchers.forEach(function (w) { if (w.key.indexOf(prefix) === 0) w.kick(); });
  }

  var mcp = Object.freeze({ callTool: callTool, watchTool: watchTool, invalidate: invalidate });

  window.claude = Object.freeze({
    __shim: true,
    help: options.help || null, // optional host-specific wording for explain() in the page
    use: async function (name) {
      if (name === "mcp") return mcp;
      throw err("capability_disabled", "The local shim only provides the \"mcp\" capability (asked for \"" + name + "\").");
    },
  });
  console.info("[claude-shim] window.claude installed (mcp via " + (transport ? "in-page transport" : ENDPOINT) + ")");
}
if (typeof __CLAUDE_SHIM_NO_AUTO__ === "undefined") installClaudeShim(window, null);
