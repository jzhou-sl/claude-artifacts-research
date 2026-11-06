// Mounts an HTML page into a container the way the SharePoint web part does. Shared by the SPFx web part
// (src/PageWebPart.ts) and the local harness (harness/harness.js), so local tests run this exact code.
//
//   const app = mountPage(container, { page: { markup, runPage }, dark });   // page from codegen pageModule()
//   app.setDark(true); app.fit(); app.dispose();
//
// The page's markup and CSS go into a same-origin iframe (isolated from the host's styles). Its script runs
// as bundled code against the iframe's window, never as inline script.

export function mountPage(container, { page, dark = false }) {
  const host = container.ownerDocument.defaultView;
  const frame = container.ownerDocument.createElement("iframe");
  frame.title = "Page";
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

  let storage; // undefined when storage is blocked
  try { storage = win.localStorage; } catch { /* blocked storage: the page copes */ }
  page.runPage(win, doc, storage);

  return {
    frame,
    fit,
    setDark(v) { dark = !!v; applyTheme(); },
    dispose() { host.removeEventListener("resize", fit); frame.remove(); }, // the page's timers and listeners go with the iframe
  };
}
