# Local `window.claude` for artifact pages

Run Claude artifact pages (here, `web-parts/teams-slack-style.html`) outside claude.ai. A dev server injects a shim that implements `window.claude.use("mcp")` and bridges tool calls to MCP servers you configure. Two servers are available: **real Teams through Microsoft Graph** (`mcp/teams/`) and a **mock** (`mcp/mock/`).

```
browser page ──► claude-shim.js ──POST /__mcp/call──► dev-server.mjs ──MCP (stdio)──► mcp/teams/ or mcp/mock/
```

## Quick start

```sh
npm install
node --env-file=.env.mock dev-server.mjs   # mock data, no sign-in needed
node dev-server.mjs                        # real Teams (set up first, see below)
```

`/` serves `web-parts/hello-world.html` (change it with `DEFAULT_PAGE`). For Teams, open <http://127.0.0.1:5173/teams-slack-style.html> and click **Connect Microsoft 365**. For the profile card, open <http://127.0.0.1:5173/profile-card.html>; it uses the `Profile` server (mock mode serves Alex Morgan from `mcp/mock/`). Any other `.html` file in `web-parts/` is served at `/<name>.html` with the shim injected.

## Real Teams setup (`mcp/teams`)

> ⚠️ In real mode, sending from the page posts **real Teams messages as you**.

1. **App registration** (Entra ID → App registrations → your app):
   - Authentication: platform *Mobile and desktop applications*, redirect URI `http://localhost`.
   - API permissions (Delegated): `User.Read`, `People.Read`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Chat.ReadWrite`, `ChannelMessage.Send`, and optionally `ChannelMessage.Read.All`.
2. Put the app's tenant ID and client ID in `mcp.config.json` under `env`.
3. Run `node mcp/teams/login.js`. A browser window opens for you to sign in. The token cache is encrypted with DPAPI and stored at `%LOCALAPPDATA%\teams-mcp\msal-cache.bin`. To sign out, run `node mcp/teams/login.js --logout`.
4. Run `node dev-server.mjs`.

**Channel messages** need `ChannelMessage.Read.All`, which usually requires admin consent. Without it, channels show a "needs admin consent" message while teams, chats and sending still work. Once an admin grants it, it works without signing in again. If you're allowed to consent yourself, run `node mcp/teams/login.js --channels`.

**Limitations:**
- Chat history and search read recent chats directly. The Search API needs admin consent and is indexed with a delay, so it isn't used. As a result, free-text search covers only the ~200 most recently active chats.
- When you're not signed in, the page shows its built-in "sign-in has lapsed… claude.ai Settings" text. Run `node mcp/teams/login.js` instead.

## Files

| Path | Purpose |
|---|---|
| `claude-shim.js` | Browser shim: `callTool`, `watchTool` (polling), `invalidate`, plus a 30 s read cache |
| `dev-server.mjs` | Injects the shim, bridges calls to MCP servers, and rejects calls from other origins |
| `mcp.config.json` / `mcp.config.mock.json` | Server name → MCP server, for real and mock mode (selected with `MCP_CONFIG`) |
| `mcp/teams-core/` | Environment-neutral tool logic used by `mcp/teams` (and kept browser-safe for a later in-browser host): `tools.js` (handlers), `graph.js` (Graph client), `map.js` (Graph → page shapes), `result-error.js` (error mapping) |
| `mcp/profile/` | `User.Read`-only MCP server: `server.mjs` (registration), `tool.mjs` (the `get_me` and `get_my_photo` handlers, browser-safe and also bundled into SharePoint web parts), `auth.mjs` / `login.mjs` (MSAL Node) |
| `mcp/teams/` | Real MCP server: `server.js` (input schemas + registration), `auth.js` / `login.js` (MSAL Node) |
| `web-parts/` | The HTML pages. `hello-world.html` and `profile-card.html` are deployed to SharePoint as the "Hello world" and "My profile" web parts |
| `package.json`, `config/`, `src/`, `scripts/`, `host/`, `pages.json` | SPFx project (the repo root): one web part per page in `pages.json` (see below and [deploy.md](deploy.md)). `src/PageWebPart.ts` is the shared web part (`src/McpPageWebPart.ts` for pages that use MCP), `host/mount.js` the mount code, `scripts/` the generator (`prepare.mjs`, `codegen.mjs`, `pages.mjs`, `add-page.mjs`). `src/webparts/` and `src/generated/` are generated |
| `mcp/mock/` | Mock server with in-memory state, seeded from `fixtures.js` |

Config entries can also point at HTTP MCP servers: `{ "url": "https://…/mcp", "headers": { … } }`.

## SharePoint (intranet hosting)

The repo is also an SPFx project that packages HTML pages as web parts for SharePoint Online: **one web part per page listed in `pages.json`**, all in a single `.sppkg`. Most need no MCP and no Graph. A page can opt in to an MCP server that the web part provides in the browser by listing it in its `pages.json` entry, for example `"mcp": ["Profile"]`. The web part then installs the `window.claude` shim in the page with an in-browser transport that runs the server's tool handlers (here `mcp/profile/tool.mjs`) against Microsoft Graph with the visitor's SharePoint sign-in. So the same page works locally through `dev-server.mjs` and on SharePoint. `Profile` is the only server available this way, and it's why the package requests delegated `User.Read` (approved once by an admin, see [deploy.md](deploy.md) §5). The Teams page above isn't included; it runs locally only.

```
SharePoint page ─► page web part ─► iframe (page markup + CSS)
                                       └─ page script, run against the iframe's window
                                            └─ window.claude.use("mcp") ─► in-browser tool handlers ─► Graph   ("mcp" pages only)
```

Each HTML page stays the single source. Before every build and start, `scripts/prepare.mjs` generates a web part per `pages.json` entry (`src/webparts/<key>/`, a thin subclass of `src/PageWebPart.ts`), the page modules (`src/generated/`) and the bundle list in `config/config.json`. A page's inline script is bundled rather than run inline, so it works under SharePoint's CSP.

```sh
node scripts/add-page.mjs web-parts/my-page.html "My page"   # register a page (path relative to the repo root)
npm start                                                        # local code on the debug page (deploy.md §3)
npm run build                                                    # → sharepoint/solution/intranet-html-pages.sppkg
```

**See [deploy.md](deploy.md)** for page rules, testing on a real SharePoint page, building, the App Catalog and putting the web parts on pages.

## Profile server, `User.Read` only (`mcp/profile`)

This is a least-privilege MCP server with two tools: `get_me` returns the signed-in user's profile, with the same result shape as the Teams server's `get_me`, and `get_my_photo` returns their profile photo (`GET /me/photo/$value`) as an MCP `image` block, or just a text block when no photo is set. It needs only the delegated `User.Read` permission, so no admin consent is required.

1. **App registration**: use a separate app with just one permission.
   - Authentication: platform *Mobile and desktop applications*, redirect URI `http://localhost`.
   - API permissions: Microsoft Graph → Delegated → `User.Read`.
2. Put the tenant ID and client ID in `.env` as `PROFILE_TENANT_ID` and `PROFILE_CLIENT_ID`, or under the `Profile` entry's `env` in `mcp.config.json`.
3. Run `node --env-file-if-exists=.env mcp/profile/login.mjs`. Its token cache is separate from the Teams one, at `%LOCALAPPDATA%\profile-mcp\msal-cache.bin`. To sign out, run `node --env-file-if-exists=.env mcp/profile/login.mjs --logout`.
4. To run it on its own, use `node --env-file-if-exists=.env mcp/profile/server.mjs` (stdio). `node dev-server.mjs` also exposes it to pages as the `Profile` server.

`web-parts/profile-card.html` (the "My profile" web part) calls it. On SharePoint the web part runs the same `get_me` handler in the browser instead, so it needs neither this app registration nor a sign-in step there.

## Environment variables

| Variable | Default | Effect |
|---|---|---|
| `PORT` | `5173` | Dev server port (binds to 127.0.0.1 only) |
| `MCP_CONFIG` | `mcp.config.json` | Which config file to load |
| `DEFAULT_PAGE` | `hello-world.html` | Page in `web-parts/` served at `/` |
| `MOCK_CHATTER_MS` | `45000` | Mock: how often fake colleagues post (`MOCK_CHATTER=0` turns it off) |
| `MOCK_FAIL` | none | Mock: comma-separated tools that return `isError` |
| `SPFX_SERVE_TENANT_DOMAIN` | none | SharePoint: the tenant that `npm start` opens. Set it in `.env`, which is git-ignored; copy it from `.env.example` (see [deploy.md](deploy.md) §3) |

## Error mapping

A server can set the code directly by returning `isError` with `structuredContent: { code, message, retryable, retryAfterMs }`; `mcp/teams` does this. Otherwise:

| Situation | `error.code` |
|---|---|
| Tool returned `isError` | `tool_error` |
| Bad arguments or unknown tool | `bad_request` |
| Server name missing from config | `server_not_connected` |
| Dev server unreachable, MCP server failed to start, or Graph throttling | `server_unavailable` (retryable) |
| Not signed in, or token expired | `needs_reauth` |
| Other failure | `upstream_error` |

Debug helper: `GET /__mcp/tools?server=Microsoft%20365` lists that server's tools.
