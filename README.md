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
| `teams-core/` | Environment-neutral tool logic shared by `teams-mcp` and the SharePoint web part: `tools.js` (handlers), `graph.js` (Graph client), `map.js` (Graph → page shapes), `result-error.js` (error mapping) |
| `teams-mcp/` | Real MCP server: `server.js` (input schemas + registration), `auth.js` / `login.js` (MSAL Node) |
| `sharepoint/` | SPFx full-page app that hosts the page on SharePoint Online (see below). `host/` holds the mount code shared by the web part and the local harness (`harness/`, served at `/sharepoint`); `scripts/codegen.mjs` turns the page and shim into modules |
| `mock-mcp/` | Mock server with in-memory state, seeded from `fixtures.js` |

Config entries can also point at HTTP MCP servers: `{ "url": "https://…/mcp", "headers": { … } }`.

## SharePoint (intranet hosting)

`sharepoint/` packages the page as an SPFx web part for a **full-page app** on SharePoint Online. There is no Node backend: the web part runs the `teams-core` tool handlers in the browser and calls Microsoft Graph with the visitor's SharePoint sign-in.

```
SharePoint page ─► SPFx web part ─► iframe (page markup + CSS)
                        │              └─ page script + shim, run against the iframe's window
                        └─ teams-core tools ─► Graph (token from SPFx AadTokenProvider)
```

`teams-slack-style.html` is still the single source. `npm run build` first runs `scripts/prepare.mjs`, which generates `src/generated/` from the page, `shim/` and `teams-core/`. The page's inline script is bundled rather than run inline, so it works under SharePoint's CSP.

**Test locally (no SPFx install or tenant needed):** run `npm run dev:mock` or `npm run dev` and open <http://127.0.0.1:5173/sharepoint>. The harness mounts the page with the web part's own code (`sharepoint/host/mount.js`), the generated page and shim modules, and `teams-core`, under a SharePoint-like CSP with no inline script. Any CSP violation is shown in the header bar. The header lets you switch between:
- **In-browser Graph (SharePoint path):** `teams-core` runs in the browser and calls Graph directly. Only the token source differs from SharePoint: it comes from your `npm run login` via `POST /__sp/token`, which accepts local origins only, instead of SPFx's `AadTokenProvider`. Available with `npm run dev`.
- **Dev server bridge:** calls go through `/__mcp/call` to the configured MCP server. Use this with `npm run dev:mock` for mock data.

The header also has a **dark SharePoint theme** toggle (simulates a theme with `isInverted`) and a **Remount** button (simulates `onDispose` followed by `render`). Edits to `teams-slack-style.html`, `shim/`, `teams-core/` or `sharepoint/host/` show up on reload.

What the harness can't cover: SPFx's token provider and the admin-approved permission set, SharePoint's real page chrome and CSP, and the `.sppkg` build. Check those in the hosted workbench.

**Build** (Node 22/24, about 1.5 GB of free disk for `node_modules`):

```sh
cd sharepoint
npm install
npm run build        # → sharepoint/sharepoint/solution/teams-slack-sharepoint.sppkg
npm start            # dev: serves to https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx
```

**Deploy** (needs a SharePoint admin):
1. Upload the `.sppkg` to the **tenant App Catalog**, and enable it for all sites or for the intranet site only.
2. **SharePoint admin center → Advanced → API access**: approve the Microsoft Graph requests (`User.Read`, `People.Read`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Chat.ReadWrite`, `ChannelMessage.Send`, `ChannelMessage.Read.All`).
   ⚠️ These are granted to the tenant-wide *SharePoint Online Client Extensibility* principal, so **every** SPFx solution in the tenant can use them. Visitors post **real Teams messages as themselves**.
3. Create the page on the intranet site, for example with PnP PowerShell:
   ```powershell
   Connect-PnPOnline https://<tenant>.sharepoint.com/sites/<intranet> -Interactive
   Add-PnPPage -Name "Teams" -LayoutType SingleWebPartAppPage
   Add-PnPPageWebPart -Page "Teams" -Component "Teams, Slack-style"
   Set-PnPPage -Identity "Teams" -Publish
   ```
   Then add the page to the site navigation.

Until the API access is approved, the page shows a "SharePoint admin needs to approve…" message instead of loading. If only `ChannelMessage.Read.All` is missing, channels report that permission, while chats and sending still work.

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
