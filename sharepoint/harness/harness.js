// Local stand-in for the SharePoint page that hosts the web part (served by dev-server.js at /sharepoint).
// Runs the same mount code (host/mount.js), generated page/shim modules and teams-core handlers as the SPFx
// bundle, under a strict CSP with no inline script. Only two pieces differ from SharePoint:
//   - "graph" transport: the Graph token comes from your local `npm run login` (POST /__sp/token)
//     instead of SPFx's AadTokenProvider. Everything after the token is the SharePoint path.
//   - "bridge" transport: calls go through the dev server to the configured MCP server (mock or real).
import { markup, runPage } from "/__sp/page.js";
import { installClaudeShim } from "/__sp/shim.js";
import { mountTeamsPage } from "/sharepoint/host/mount.js";
import { HELP, MESSAGES } from "/sharepoint/host/wording.js";
import { createGraph } from "/teams-core/graph.js";
import { createTools } from "/teams-core/tools.js";
import { AuthError } from "/teams-core/errors.js";
import { resultError } from "/teams-core/result-error.js";

const $ = (s) => document.querySelector(s);
const root = document.documentElement;
const graphAvailable = root.dataset.graph === "1";
const params = new URLSearchParams(location.search);

const modeSel = $("#mode"), darkBox = $("#dark"), status = $("#status");
if (!graphAvailable) {
  modeSel.querySelector('option[value="graph"]').disabled = true;
  modeSel.querySelector('option[value="graph"]').textContent += " (needs real config)";
}
modeSel.value = params.get("mode") === "bridge" || !graphAvailable ? "bridge" : (params.get("mode") || root.dataset.mode);
darkBox.checked = params.get("theme") === "dark";

async function getToken(scopes) {
  let res;
  try {
    res = await fetch("/__sp/token", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ scopes }) });
  } catch (e) {
    throw new AuthError("needs_reauth", "Dev server unreachable for a token: " + (e?.message || e));
  }
  const body = await res.json().catch(() => null);
  if (body?.ok) return body.token;
  throw new AuthError(body?.error?.code || "needs_reauth", body?.error?.message || `Token endpoint returned ${res.status}`);
}

let app = null;
function mount() {
  app?.dispose();
  const mode = modeSel.value;
  root.setAttribute("data-theme", darkBox.checked ? "dark" : "light");
  app = mountTeamsPage($("#canvas"), {
    page: { markup, runPage },
    installClaudeShim,
    core: { createGraph, createTools, resultError, AuthError },
    getToken,
    useDevServer: mode === "bridge",
    help: HELP,
    messages: MESSAGES,
    dark: darkBox.checked,
    log: (...a) => console.warn("[teams-slack]", ...a),
  });
  status.textContent = mode === "graph" ? "Graph called from the browser · token from npm run login" : "Calls bridged through the dev server";
  status.classList.remove("warn");
  const url = new URL(location.href);
  url.searchParams.set("mode", mode); url.searchParams.set("theme", darkBox.checked ? "dark" : "light");
  history.replaceState(null, "", url);
}

modeSel.addEventListener("change", mount);
$("#remount").addEventListener("click", mount);
darkBox.addEventListener("change", () => { root.setAttribute("data-theme", darkBox.checked ? "dark" : "light"); app?.setDark(darkBox.checked); });

// SharePoint enforces a CSP; surface any violation (e.g. something trying to run inline script).
// The iframe inherits this page's CSP, so listen in both documents.
function onViolation(e) {
  status.textContent = `CSP violation: ${e.violatedDirective} (${e.blockedURI || "inline"})`;
  status.classList.add("warn");
  console.error("[harness] CSP violation", e.violatedDirective, e.blockedURI, e.sourceFile, e.lineNumber);
}
document.addEventListener("securitypolicyviolation", onViolation);

mount();
const watchFrame = () => app?.frame.contentDocument?.addEventListener("securitypolicyviolation", onViolation);
watchFrame();
$("#remount").addEventListener("click", watchFrame);
modeSel.addEventListener("change", watchFrame);
