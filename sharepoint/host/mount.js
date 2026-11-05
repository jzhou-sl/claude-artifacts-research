// Mounts teams-slack-style.html into a container the way the SharePoint web part does. Shared by the SPFx
// web part (TeamsSlackWebPart.ts) and the local harness (harness/harness.js), so local tests run this exact code.
//
//   const app = mountTeamsPage(container, {
//     page: { markup, runPage },          // from codegen pageModule()
//     installClaudeShim,                  // from codegen shimModule()
//     core: { createGraph, createTools, resultError, AuthError },   // teams-core
//     getToken: async (scopes) => token,  // Graph access token; omit with useDevServer
//     useDevServer: false,                // true: the shim calls the local dev server (/__mcp/call) instead
//     help, messages, dark, log,
//   });
//   app.setDark(true); app.fit(); app.dispose();
//
// The page's markup and CSS go into a same-origin iframe (isolated from the host's styles). Its script and
// the window.claude shim run as bundled code against the iframe's window, never as inline script.

export const SERVER = "Microsoft 365";

export function mountTeamsPage(container, opts) {
  const { page, installClaudeShim, core, getToken, useDevServer = false, help = null, messages, log = console.warn } = opts;
  let dark = !!opts.dark;

  const host = container.ownerDocument.defaultView;
  const frame = container.ownerDocument.createElement("iframe");
  frame.title = "Teams";
  frame.setAttribute("style", "display:block;width:100%;border:0;min-height:480px");
  container.replaceChildren(frame);

  const fit = () => {
    const top = container.getBoundingClientRect().top + host.scrollY;
    frame.style.height = `${Math.max(480, host.innerHeight - top)}px`;
  };
  host.addEventListener("resize", fit);
  fit();

  const win = frame.contentWindow;
  const doc = win.document;
  doc.open();
  doc.write(page.markup); // no <script> elements in markup
  doc.close();

  const applyTheme = () => doc.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
  applyTheme();

  installClaudeShim(win, useDevServer ? null : graphTransport(), { help });
  let storage = null;
  try { storage = win.localStorage; } catch { /* blocked storage: the page copes */ }
  page.runPage(win, doc, storage);

  /** In-browser replacement for dev-server.js + teams-mcp: tool name -> Graph, same results and errors. */
  function graphTransport() {
    const tools = core.createTools({ ...core.createGraph(getToken), messages, log });
    return async (server, tool, input) => {
      if (server !== SERVER) throw { code: "server_not_connected", message: `No server named "${server}"`, retryable: false };
      const t = tools[tool];
      if (!t) throw { code: "bad_request", message: `Unknown tool ${tool}`, retryable: false };
      const result = await t.handler(input || {});
      if (result.isError) throw core.resultError(result);
      return result;
    };
  }

  return {
    frame,
    fit,
    setDark(v) { dark = !!v; applyTheme(); },
    dispose() { host.removeEventListener("resize", fit); frame.remove(); }, // the page's timers and listeners go with the iframe
  };
}
