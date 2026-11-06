// The page registry (pages.json): one entry per HTML page, each built as its own web part.
// Shared by prepare.mjs (SPFx build), add-page.mjs and dev-server.js (local harness).
//
//   { "file": "web-parts/hello-world.html",  // relative to the repo root
//     "id": "7c081f09-…",                    // the web part's permanent GUID: never change it once deployed
//     "key": "helloWorld",                   // camelCase; names the generated folder, class and bundle
//     "title": "Hello world",                // shown in the SharePoint toolbox
//     "description": "…", "icon": "Emoji2" } // icon: a Fluent UI icon name
import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export const SP_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
export const ROOT_DIR = SP_DIR; // the SPFx project is the repo root
export const PAGES_FILE = path.join(SP_DIR, "pages.json");

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const KEY = /^[a-z][A-Za-z0-9]*$/;

/** "my-page.html" -> "myPage" */
export function keyFromFile(file) {
  const words = path.basename(file, path.extname(file)).split(/[^A-Za-z0-9]+/).filter(Boolean);
  const key = words.map((w, i) => (i ? w[0].toUpperCase() + w.slice(1) : w[0].toLowerCase() + w.slice(1))).join("");
  return /^[a-z]/.test(key) ? key : "page" + key.charAt(0).toUpperCase() + key.slice(1);
}
export const className = (key) => key[0].toUpperCase() + key.slice(1) + "WebPart";
export const bundleName = (key) => key.replace(/[A-Z]/g, (c) => "-" + c.toLowerCase()) + "-web-part";

/** Read and validate pages.json; throws an Error listing every problem. */
export async function loadPages() {
  let pages;
  try { pages = JSON.parse(await readFile(PAGES_FILE, "utf8")); }
  catch (e) { throw new Error(`Can't read ${path.relative(process.cwd(), PAGES_FILE)}: ${e.message}`); }
  if (!Array.isArray(pages) || !pages.length) throw new Error("pages.json must be a non-empty array of pages.");

  const problems = [], ids = new Set(), keys = new Set();
  pages.forEach((p, i) => {
    const at = `pages.json[${i}]${p?.file ? ` (${p.file})` : ""}`;
    if (!p || typeof p !== "object") { problems.push(`${at}: not an object`); return; }
    for (const f of ["file", "id", "key", "title"]) if (typeof p[f] !== "string" || !p[f].trim()) problems.push(`${at}: "${f}" is required`);
    if (p.id && !GUID.test(p.id)) problems.push(`${at}: "id" must be a GUID`);
    if (p.key && !KEY.test(p.key)) problems.push(`${at}: "key" must be camelCase letters/digits, like "helloWorld"`);
    if (p.id && ids.has(p.id.toLowerCase())) problems.push(`${at}: duplicate id ${p.id}`);
    if (p.key && keys.has(p.key)) problems.push(`${at}: duplicate key ${p.key}`);
    if (p.file && !existsSync(path.join(ROOT_DIR, p.file))) problems.push(`${at}: file not found (paths are relative to the repo root)`);
    ids.add(String(p.id).toLowerCase()); keys.add(p.key);
  });
  if (problems.length) throw new Error("Invalid pages.json:\n  - " + problems.join("\n  - "));
  return pages.map((p) => ({ description: "", icon: "Page", ...p }));
}
