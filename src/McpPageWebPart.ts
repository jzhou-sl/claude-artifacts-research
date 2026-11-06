// Base class for page web parts that call MCP servers (pages.json "mcp": [...]). scripts/prepare.mjs makes those
// pages' web parts extend this class instead of PageWebPart.
//
// It installs the window.claude shim in the page's iframe with an in-browser transport: each supported server's
// tool handlers run here against Microsoft Graph, with the visitor's SharePoint sign-in, and return the same
// results and errors as the local MCP servers behind dev-server.mjs. The tokens carry whatever Graph permissions
// an admin approved for the package (config/package-solution.json webApiPermissionRequests).
import type { AadTokenProvider } from '@microsoft/sp-http';

import { PageWebPart } from './PageWebPart';
import { installClaudeShim } from './generated/claudeShim';
import { createGraph } from './generated/core/teams-core/graph';
import { AuthError } from './generated/core/teams-core/errors';
import { resultError } from './generated/core/teams-core/result-error';
import { createProfileTools } from './generated/core/profile/tool';

interface IToolResult { isError?: boolean; content?: unknown[]; structuredContent?: unknown }
interface ITool { handler(input: object): Promise<IToolResult> }
type Graph = ReturnType<typeof createGraph>;
type Tools = Record<string, ITool>;

const ADMIN_HINT = 'A SharePoint admin needs to approve this app’s Microsoft Graph permissions (SharePoint admin center → Advanced → API access).';
const UNREACHABLE = 'This page couldn’t reach Microsoft 365. Reload the page, and contact the intranet team if it keeps happening.';

/** window.claude.help: SharePoint wording for the page's error text, keyed by error code. */
const HELP: Record<string, string> = {
  needs_reauth: 'Your Microsoft 365 sign-in has lapsed. Reload this page to sign in again.',
  server_not_connected: UNREACHABLE,
  not_granted: UNREACHABLE,
  capability_disabled: UNREACHABLE,
};
const MESSAGES = { graph401: HELP.needs_reauth, graph403Hint: ADMIN_HINT };

const log = (...a: unknown[]): void => console.warn('[mcp-web-part]', ...a);

/** Server name (as pages call it, and as in mcp.config.json) -> its tools over Graph. Keep in sync with scripts/pages.mjs MCP_SERVERS. */
const SERVERS: Record<string, (graph: Graph) => Tools> = {
  Profile: (graph) => createProfileTools({ ...graph, messages: MESSAGES, log }) as Tools,
};

/** The window.claude error shape { code, message, retryable, retryAfterMs }, as an Error. */
function mcpError(shape: { code: string; message: string; retryable?: boolean; retryAfterMs?: number }): Error {
  return Object.assign(new Error(shape.message), { retryable: false, ...shape });
}

export abstract class McpPageWebPart extends PageWebPart {
  private _tokenProvider: Promise<AadTokenProvider> | undefined;

  /** The MCP servers the page calls. */
  protected abstract get mcpServers(): string[];

  protected onInit(): Promise<void> {
    this._tokenProvider = this.context.aadTokenProviderFactory.getTokenProvider();
    return super.onInit();
  }

  protected beforeRun(win: Window): void {
    installClaudeShim(win, this._transport(), { help: HELP });
  }

  private _transport(): (server: string, tool: string, input: object) => Promise<IToolResult> {
    const graph = createGraph(() => this._getToken());
    const servers: Record<string, Tools> = {};
    for (const name of this.mcpServers) if (SERVERS[name]) servers[name] = SERVERS[name](graph);

    return async (server, tool, input) => {
      const tools = servers[server];
      if (!tools) throw mcpError({ code: 'server_not_connected', message: `No server named "${server}" in this web part.` });
      const t = tools[tool];
      if (!t) throw mcpError({ code: 'bad_request', message: `Unknown tool ${tool}` });
      const result = await t.handler(input || {});
      if (result.isError) throw mcpError(resultError(result));
      return result;
    };
  }

  /** SPFx tokens carry every Graph permission the admin approved, so the requested scopes are ignored. */
  private async _getToken(): Promise<string> {
    try {
      const provider = await this._tokenProvider!;
      return await provider.getToken('https://graph.microsoft.com');
    } catch (e) {
      const msg = String((e as Error)?.message || e);
      if (/AADSTS(65001|650057|500011)|consent|not.*approved|permission/i.test(msg)) {
        throw new AuthError('consent_missing', `SharePoint couldn’t get a Microsoft Graph token for this page. ${ADMIN_HINT} (${msg.slice(0, 160)})`);
      }
      throw new AuthError('needs_reauth', msg);
    }
  }
}
