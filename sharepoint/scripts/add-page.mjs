// Registers an HTML page as a new web part:  npm run add-page -- <file> "<title>" ["<description>"]
//   <file> is relative to the repo root, e.g. my-page.html or pages/news.html
// Checks the page can be packaged, gives it a new permanent GUID and appends it to pages.json.
import { readFile, writeFile } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { pageModule } from "./codegen.mjs";
import { ROOT_DIR, PAGES_FILE, loadPages, keyFromFile } from "./pages.mjs";

const [file, title, description = ""] = process.argv.slice(2);
const fail = (msg) => { console.error(`[add-page] ${msg}`); process.exit(1); };
if (!file || !title) fail('Usage: npm run add-page -- <file relative to repo root> "<title>" ["<description>"]');

const rel = path.relative(ROOT_DIR, path.resolve(ROOT_DIR, file)).split(path.sep).join("/");
if (rel.startsWith("..")) fail(`${file} is outside the repo.`);
let html;
try { html = await readFile(path.join(ROOT_DIR, rel), "utf8"); } catch { fail(`${rel} not found (paths are relative to the repo root).`); }
try { pageModule(html); } catch (e) { fail(`${rel}: ${e.message}`); }

const pages = await loadPages().catch((e) => fail(e.message));
if (pages.some((p) => p.file === rel)) fail(`${rel} is already in pages.json.`);
let key = keyFromFile(rel);
for (let n = 2; pages.some((p) => p.key === key); n++) key = keyFromFile(rel) + n;

const raw = JSON.parse(await readFile(PAGES_FILE, "utf8"));
raw.push({ file: rel, id: randomUUID(), key, title, description, icon: "Page" });
await writeFile(PAGES_FILE, JSON.stringify(raw, null, 2) + "\n");
console.log(`[add-page] added "${title}" (${rel}) as web part ${key}. It is built into the package on the next npm run build.`);
