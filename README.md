# Local `window.claude` for artifact pages

Run Claude artifact pages (here, `teams-slack-style.html`) outside claude.ai. A dev server injects a shim that implements `window.claude.use("mcp")` and bridges tool calls to MCP servers you configure, by default a mock Microsoft 365 server.

```
browser page ──► shim/claude-shim.js ──POST /__mcp/call──► dev-server.js ──MCP──► mock-mcp/server.js
```

## Quick start

```sh
npm install
npm run dev        # http://127.0.0.1:5173/
```

Open the page and click **Connect Microsoft 365**. Any other `.html` file in this folder is served at `/<name>.html` with the shim injected.

## Files

| Path | Purpose |
|---|---|
| `shim/claude-shim.js` | Browser shim: `callTool`, `watchTool` (polling), `invalidate`, plus a 30 s read cache |
| `dev-server.js` | Injects the shim and bridges calls to MCP clients over stdio or HTTP |
| `mcp.config.json` | Maps the server names pages use to MCP servers |
| `mock-mcp/` | Mock Microsoft 365 MCP server with in-memory state, seeded from `fixtures.js` |

## Using a real MCP server

Edit `mcp.config.json`; the key must match the name the page asks for:

```json
{
  "Microsoft 365": { "command": "npx", "args": ["some-mcp-server"], "env": {} },
  "Other": { "url": "https://example.com/mcp", "headers": { "Authorization": "Bearer …" } }
}
```

## Environment variables

| Variable | Default | Effect |
|---|---|---|
| `PORT` | `5173` | Dev server port (binds to 127.0.0.1 only) |
| `DEFAULT_PAGE` | `teams-slack-style.html` | Page served at `/` |
| `MOCK_CHATTER_MS` | `45000` | How often fake colleagues post (`MOCK_CHATTER=0` turns it off) |
| `MOCK_FAIL` | none | Comma-separated tools that return `isError`, for testing error handling |

## Error mapping

| Situation | `error.code` |
|---|---|
| Tool returned `isError` | `tool_error` |
| Bad arguments or unknown tool | `bad_request` |
| Server name missing from config | `server_not_connected` |
| Dev server unreachable, or MCP server failed to start | `server_unavailable` (retryable) |
| Other JSON-RPC or transport failure | `upstream_error` |

Debug helper: `GET /__mcp/tools?server=Microsoft%20365` lists that server's tools.
