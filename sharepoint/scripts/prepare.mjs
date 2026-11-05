// Generates src/generated/ for the SPFx build from the files shared with the local dev setup, so the
// HTML page stays the single source of truth:
//   page.ts        <- ../teams-slack-style.html : markup (no scripts) + runPage() (see codegen.mjs)
//   claudeShim.ts  <- ../shim/claude-shim.js     : exports installClaudeShim (no auto-install)
//   mount.ts, wording.ts <- host/*.js            : iframe + shim + transport, SharePoint wording (shared with the local harness)
//   teams-core/    <- ../teams-core/*.js         : Graph client + tool handlers
// The sources are plain JS; they are emitted as .ts with @ts-nocheck so the SPFx TypeScript build compiles
// and bundles them unchanged. Nothing runs as inline script in SharePoint, so the page works under its CSP.
import { readFile, writeFile, mkdir, readdir, rm } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { banner, pageModule, shimModule } from "./codegen.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SP = path.resolve(HERE, "..");
const ROOT = path.resolve(SP, "..");
const OUT = path.join(SP, "src", "generated");
const read = (...p) => readFile(path.join(...p), "utf8");

await rm(OUT, { recursive: true, force: true });
await mkdir(path.join(OUT, "teams-core"), { recursive: true });

await writeFile(path.join(OUT, "page.ts"), pageModule(await read(ROOT, "teams-slack-style.html"), { ts: true }));
await writeFile(path.join(OUT, "claudeShim.ts"), shimModule(await read(ROOT, "shim", "claude-shim.js"), { ts: true }));
for (const f of (await readdir(path.join(SP, "host"))).filter((f) => f.endsWith(".js"))) {
  await writeFile(path.join(OUT, f.replace(/\.js$/, ".ts")), banner(true) + (await read(SP, "host", f)));
}

const coreDir = path.join(ROOT, "teams-core");
for (const f of (await readdir(coreDir)).filter((f) => f.endsWith(".js"))) {
  await writeFile(path.join(OUT, "teams-core", f.replace(/\.js$/, ".ts")), banner(true) + (await read(coreDir, f)));
}

console.log(`[prepare] wrote ${path.relative(process.cwd(), OUT)}`);
