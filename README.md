# Local `window.claude` for artifact pages

Run Claude artifact pages (here, `teams-slack-style.html`) outside claude.ai. A dev server injects a shim that implements `window.claude.use("mcp")` and bridges tool calls to MCP servers you configure. Two servers are available: **real Teams through Microsoft Graph** (`teams-mcp/`) and a **mock** (`mock-mcp/`).

```
browser page ──► shim/claude-shim.js ──POST /__mcp/call──► dev-server.js ──MCP (stdio)──► teams-mcp/ or mock-mcp/
```

## Quick start

```sh
npm install
npm run dev:mock   # mock data, no sign-in needed → http://127.0.0.1:5173/
npm run dev        # real Teams (set up first, see below)
```

Open the page and click **Connect Microsoft 365**. Any other `.html` file in this folder is served at `/<name>.html` with the shim injected.

## Real Teams setup (`teams-mcp`)

> ⚠️ In real mode, sending from the page posts **real Teams messages as you**.

1. **App registration** (Entra ID → App registrations → your app):
   - Authentication: platform *Mobile and desktop applications*, redirect URI `http://localhost`.
   - API permissions (Delegated): `User.Read`, `People.Read`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Chat.ReadWrite`, `ChannelMessage.Send`, and optionally `ChannelMessage.Read.All`.
2. Put the app's tenant ID and client ID in `mcp.config.json` under `env`.
3. Run `npm run login`. A browser window opens for you to sign in. The token cache is encrypted with DPAPI and stored at `%LOCALAPPDATA%\teams-mcp\msal-cache.bin`. To sign out, run `npm run logout`.
4. Run `npm run dev`.

**Channel messages** need `ChannelMessage.Read.All`, which usually requires admin consent. Without it, channels show a "needs admin consent" message while teams, chats and sending still work. Once an admin grants it, it works without signing in again. If you're allowed to consent yourself, run `npm run login -- --channels`.

**Limitations:**
- Chat history and search read recent chats directly. The Search API needs admin consent and is indexed with a delay, so it isn't used. As a result, free-text search covers only the ~200 most recently active chats.
- When you're not signed in, the page shows its built-in "sign-in has lapsed… claude.ai Settings" text. Run `npm run login` instead.

## Files

| Path | Purpose |
|---|---|
| `shim/claude-shim.js` | Browser shim: `callTool`, `watchTool` (polling), `invalidate`, plus a 30 s read cache |
| `dev-server.js` | Injects the shim, bridges calls to MCP servers, and rejects calls from other origins |
| `mcp.config.json` / `mcp.config.mock.json` | Server name → MCP server, for real and mock mode (selected with `MCP_CONFIG`) |
| `teams-mcp/` | Real server: `server.js` (tools), `graph.js` (Graph client), `map.js` (Graph → page shapes), `auth.js` / `login.js` (MSAL) |
| `mock-mcp/` | Mock server with in-memory state, seeded from `fixtures.js` |

Config entries can also point at HTTP MCP servers: `{ "url": "https://…/mcp", "headers": { … } }`.

## Environment variables

| Variable | Default | Effect |
|---|---|---|
| `PORT` | `5173` | Dev server port (binds to 127.0.0.1 only) |
| `MCP_CONFIG` | `mcp.config.json` | Which config file to load |
| `DEFAULT_PAGE` | `teams-slack-style.html` | Page served at `/` |
| `MOCK_CHATTER_MS` | `45000` | Mock: how often fake colleagues post (`MOCK_CHATTER=0` turns it off) |
| `MOCK_FAIL` | none | Mock: comma-separated tools that return `isError` |

## Error mapping

A server can set the code directly by returning `isError` with `structuredContent: { code, message, retryable, retryAfterMs }`; `teams-mcp` does this. Otherwise:

| Situation | `error.code` |
|---|---|
| Tool returned `isError` | `tool_error` |
| Bad arguments or unknown tool | `bad_request` |
| Server name missing from config | `server_not_connected` |
| Dev server unreachable, MCP server failed to start, or Graph throttling | `server_unavailable` (retryable) |
| Not signed in, or token expired | `needs_reauth` |
| Other failure | `upstream_error` |

Debug helper: `GET /__mcp/tools?server=Microsoft%20365` lists that server's tools.
