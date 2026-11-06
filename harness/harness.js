// Local stand-in for a SharePoint page hosting the web parts (served by dev-server.js at /sharepoint).
// Runs the same mount code (host/mount.js) and generated page modules as the SPFx bundle, under a strict CSP
// with no inline script. Pick a web part from pages.json in the header; ?page=<key> selects one directly.
import { mountPage } from "/host/mount.js";

const $ = (s) => document.querySelector(s);
const root = document.documentElement;
const params = new URLSearchParams(location.search);
const pageSel = $("#page"), darkBox = $("#dark"), status = $("#status"), crumb = $("#crumb");
darkBox.checked = params.get("theme") === "dark";

// SharePoint enforces a CSP; surface any violation (e.g. something trying to run inline script).
// The iframe inherits this page's CSP, so listen in both documents.
function onViolation(e) {
  status.textContent = `CSP violation: ${e.violatedDirective} (${e.blockedURI || "inline"})`;
  status.classList.add("warn");
  console.error("[harness] CSP violation", e.violatedDirective, e.blockedURI, e.sourceFile, e.lineNumber);
}
document.addEventListener("securitypolicyviolation", onViolation);

function fail(msg) {
  status.textContent = msg;
  status.classList.add("warn");
}

let app = null, pages = [];
async function mount() {
  app?.dispose(); app = null;
  const entry = pages.find((p) => p.key === pageSel.value);
  if (!entry) return;
  root.setAttribute("data-theme", darkBox.checked ? "dark" : "light");
  crumb.textContent = entry.title;
  let page;
  try { page = await import(`/__sp/page.js?key=${encodeURIComponent(entry.key)}&t=${Date.now()}`); }
  catch (e) { fail(`Couldn't load ${entry.file}: ${e.message}`); return; }
  app = mountPage($("#canvas"), { page, dark: darkBox.checked });
  app.frame.contentDocument.addEventListener("securitypolicyviolation", onViolation);
  status.textContent = `Showing ${entry.file}`;
  status.classList.remove("warn");
  const url = new URL(location.href);
  url.searchParams.set("page", entry.key);
  url.searchParams.set("theme", darkBox.checked ? "dark" : "light");
  history.replaceState(null, "", url);
}

pageSel.addEventListener("change", mount);
$("#remount").addEventListener("click", mount);
darkBox.addEventListener("change", () => { root.setAttribute("data-theme", darkBox.checked ? "dark" : "light"); app?.setDark(darkBox.checked); });

const res = await fetch("/__sp/pages").then((r) => r.json()).catch((e) => ({ ok: false, error: { message: String(e) } }));
if (!res.ok) fail(res.error.message);
else {
  pages = res.pages;
  pageSel.replaceChildren(...pages.map((p) => new Option(p.title, p.key)));
  if (pages.some((p) => p.key === params.get("page"))) pageSel.value = params.get("page");
  await mount();
}
