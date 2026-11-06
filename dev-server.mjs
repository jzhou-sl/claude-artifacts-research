#!/usr/bin/env node
// Local dev server: serves artifact HTML pages with a window.claude shim injected,
// and bridges the shim's tool calls to real MCP servers listed in mcp.config.json.
//
//   GET  /                     -> hello-world.html, or DEFAULT_PAGE (shim injected)
//   GET  /<page>.html          -> any .html in web-parts/ (shim injected)
//   GET  /__claude-shim.js     -> the shim
//   POST /__mcp/call           -> { server, tool, input } -> { ok, result } | { ok:false, error }
//   GET  /__mcp/tools?server=  -> list tools (debugging)
//
// Env: PORT (default 5173), DEFAULT_PAGE (default hello-world.html)

import http from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { resultError } from "./mcp/teams-core/result-error.js";

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT) || 5173;
const HOST = "127.0.0.1";
const DEFAULT_PAGE = process.env.DEFAULT_PAGE || "hello-world.html";
const SHIM_PATH = path.join(ROOT, "claude-shim.js");
const PAGES_DIR = path.join(ROOT, "web-parts");
const CONFIG_FILE = process.env.MCP_CONFIG || "mcp.config.json";
const config = JSON.parse(await readFile(path.join(ROOT, CONFIG_FILE), "utf8"));
const ALLOWED_ORIGINS = new Set([`http://127.0.0.1:${PORT}`, `http://localhost:${PORT}`]);

/* ---------- MCP clients (lazy, reconnect after failure) ---------- */
const clients = new Map(); // server name -> Promise<Client>

// Config env values override the environment, except "<…>" placeholders (as in mcp.config.json), so
// values from .env (node --env-file=.env dev-server.mjs) still apply.
const configEnv = (env = {}) => Object.fromEntries(Object.entries(env).filter(([, v]) => !/^<.*>$/.test(String(v))));

function connect(name) {
  const cfg = config[name];
  const transport = cfg.url
    ? new StreamableHTTPClientTransport(new URL(cfg.url), { requestInit: { headers: cfg.headers || {} } })
    : new StdioClientTransport({ command: cfg.command, args: cfg.args || [], env: { ...process.env, ...configEnv(cfg.env) }, cwd: cfg.cwd || ROOT, stderr: "inherit" });
  const client = new Client({ name: "artifact-dev-bridge", version: "0.1.0" });
  client.onclose = () => { console.warn(`[bridge] "${name}" disconnected`); clients.delete(name); };
  client.onerror = (e) => console.warn(`[bridge] "${name}" error:`, e?.message || e);
  const p = client.connect(transport).then(() => { console.log(`[bridge] connected to "${name}"`); return client; });
  p.catch(() => clients.delete(name));
  return p;
}
function getClient(name) {
  if (!clients.has(name)) clients.set(name, connect(name));
  return clients.get(name);
}

/* ---------- error mapping to the window.claude error shape ---------- */
class BridgeError extends Error {
  constructor(code, message, extra = {}) { super(message); this.code = code; Object.assign(this, extra); }
}
function toErrorShape(e) {
  if (e instanceof BridgeError) return { code: e.code, message: e.message, retryable: !!e.retryable, ...(e.retryAfterMs ? { retryAfterMs: e.retryAfterMs } : {}) };
  const msg = String(e?.message || e);
  // JSON-RPC errors from the MCP server
  if (typeof e?.code === "number") {
    if (e.code === -32602) return { code: "bad_request", message: msg, retryable: false };
    if (e.code === -32001) return { code: "server_unavailable", message: msg, retryable: true, retryAfterMs: 1000 }; // request timeout
    return { code: "upstream_error", message: msg, retryable: false };
  }
  // transport-level failures (process died, connection refused, closed)
  if (/closed|ECONNREFUSED|ENOENT|spawn|not connected|EPIPE/i.test(msg)) return { code: "server_unavailable", message: msg, retryable: true, retryAfterMs: 800 };
  return { code: "upstream_error", message: msg, retryable: true, retryAfterMs: 800 };
}

async function callTool({ server, tool, input }) {
  if (typeof server !== "string" || typeof tool !== "string") throw new BridgeError("bad_request", "server and tool must be strings");
  if (!config[server]) throw new BridgeError("server_not_connected", `No MCP server named "${server}" in mcp.config.json`);
  let client;
  try { client = await getClient(server); }
  catch (e) { throw new BridgeError("server_unavailable", `Couldn't start "${server}": ${e.message}`, { retryable: true, retryAfterMs: 1500 }); }
  const result = await client.callTool({ name: tool, arguments: input && typeof input === "object" ? input : {} });
  if (result.isError) {
    const { code, message, ...extra } = resultError(result);
    throw new BridgeError(code, message, extra);
  }
  return result;
}

/* ---------- HTTP ---------- */
const send = (res, status, type, body) => { res.writeHead(status, { "content-type": type, "cache-control": "no-store" }); res.end(body); };
const json = (res, status, obj) => send(res, status, "application/json; charset=utf-8", JSON.stringify(obj));
async function readBody(req, max = 1_000_000) {
  let size = 0; const chunks = [];
  for await (const c of req) { size += c.length; if (size > max) throw new BridgeError("bad_request", "Body too large"); chunks.push(c); }
  return Buffer.concat(chunks).toString("utf8");
}
function injectShim(html) {
  const tag = `<script src="/__claude-shim.js"></script>`;
  return /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, (m) => m + "\n" + tag) : tag + "\n" + html;
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  try {
    if (req.method === "POST" && url.pathname === "/__mcp/call") {
      // tools can send real messages: refuse calls from any other origin (other local sites, file:// pages)
      const origin = req.headers.origin;
      if (origin && !ALLOWED_ORIGINS.has(origin)) return json(res, 403, { ok: false, error: { code: "blocked_by_policy", message: `Origin ${origin} not allowed` } });
      // only accept JSON calls (forces a CORS preflight for cross-origin requests)
      if (!/^application\/json/.test(req.headers["content-type"] || "")) return json(res, 415, { ok: false, error: { code: "bad_request", message: "JSON only" } });
      let payload;
      try { payload = JSON.parse(await readBody(req)); } catch { return json(res, 400, { ok: false, error: { code: "bad_request", message: "Invalid JSON" } }); }
      try {
        const result = await callTool(payload);
        return json(res, 200, { ok: true, result });
      } catch (e) {
        const error = toErrorShape(e);
        console.warn(`[bridge] ${payload.server} / ${payload.tool} -> ${error.code}: ${error.message}`);
        return json(res, 200, { ok: false, error });
      }
    }
    if (req.method === "GET" && url.pathname === "/__mcp/tools") {
      const name = url.searchParams.get("server") || Object.keys(config)[0];
      if (!config[name]) return json(res, 404, { ok: false, error: { code: "server_not_connected" } });
      const client = await getClient(name);
      return json(res, 200, { ok: true, server: name, tools: (await client.listTools()).tools.map((t) => ({ name: t.name, description: t.description })) });
    }
    if (req.method === "GET" && url.pathname === "/__claude-shim.js") {
      return send(res, 200, "text/javascript; charset=utf-8", await readFile(SHIM_PATH));
    }
    if (req.method === "GET" && (url.pathname === "/" || /^\/[\w.-]+\.html$/.test(url.pathname))) {
      const file = url.pathname === "/" ? DEFAULT_PAGE : url.pathname.slice(1);
      let html;
      try { html = await readFile(path.join(PAGES_DIR, path.basename(file)), "utf8"); } catch { return send(res, 404, "text/plain", "Not found"); }
      return send(res, 200, "text/html; charset=utf-8", injectShim(html));
    }
    if (url.pathname === "/favicon.ico") { res.writeHead(204); return res.end(); }
    send(res, 404, "text/plain", "Not found");
  } catch (e) {
    console.error("[dev-server]", e);
    json(res, 500, { ok: false, error: toErrorShape(e) });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Dev server: http://${HOST}:${PORT}/  (${CONFIG_FILE}: ${Object.keys(config).join(", ")})`);
});

const shutdown = async () => {
  for (const p of clients.values()) { try { (await p).close(); } catch {} }
  process.exit(0);
};
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
