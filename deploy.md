# Deploying HTML pages to SharePoint as web parts

Each page listed in `pages.json` becomes its own web part. All of them ship in one package, `intranet-html-pages.sppkg`, which you deploy to `https://scottlogic.sharepoint.com`. This follows Microsoft's SPFx guides ([connect](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/connect-to-sharepoint), [deploy](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/serve-your-web-part-in-a-sharepoint-page), [debug on pages](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/debug-modern-pages)), using SPFx 1.24 with Heft.

```
HTML page ─► npm run build ─► .sppkg ─► App Catalog ─► site ─► page ─► web part
```

## 1. Prerequisites

- **Node.js 22 or 24**, about **1.5 GB** of free disk, and **Edge** (or Chrome).
- **Upload rights to an App Catalog:** either the tenant one (usually SharePoint admins) or a site collection App Catalog on the intranet site.
- **Edit rights** on the target site.
- The web parts request no API permissions, so no admin approval is needed.
- Run `npm install` in the repo root once.
- **For testing on SharePoint (section 3):** a `.env` file (copied from `.env.example`), and a debug page you can edit (by default `NewcastleSocialCommittee/SitePages/SPFx-dev.aspx`).

## 2. Add a page

A page must be one self-contained HTML file:
- **At most one inline `<script>`, and no `<script src>`.** The script gets bundled, so it runs under SharePoint's content security policy.
- **CSS goes in `<style>`.** `<link>` stylesheets are dropped.
- **Absolute (`https://`) or `data:` URLs only** for images and other files. There's no backend.
- **Dark mode:** `<html>` gets `data-theme="dark"` or `data-theme="light"` from the site theme (see `web-parts/hello-world.html`).

Put the page in `web-parts/` and register it. The path is relative to the repo root:

```sh
node scripts/add-page.mjs web-parts/my-page.html "My page" "Optional description"
```

This appends an entry with a new permanent `id` to `pages.json`. You can edit `title`, `description` and `icon` (a Fluent UI icon name). **Never change `id` once deployed**: SharePoint uses it to find the web part on every page.

## 3. Test on SharePoint with your local code

This runs your local code on a real page, and only you see it. The hosted workbench (`/_layouts/workbench.aspx`) is **retired on 1 December 2026**. Microsoft's replacement is the **SPFx debugging toolbar** on a normal modern page. We use a dedicated debug page, so the intranet's real pages are never touched.

**One-time setup:**

1. Tell the dev server your tenant. In the repo root, copy `.env.example` to `.env`; it's git-ignored and already set to `scottlogic.sharepoint.com`:
   ```powershell
   Copy-Item .env.example .env
   ```
   `npm start` loads `.env` and fills in `{tenantDomain}` in `config/serve.json`. The VS Code **Debug on SharePoint page** launch config can't read `.env`. To use it, also set the variable for your Windows user once with `setx SPFX_SERVE_TENANT_DOMAIN scottlogic.sharepoint.com`, then restart VS Code.
2. Have a debug page. The default is `https://scottlogic.sharepoint.com/sites/NewcastleSocialCommittee/SitePages/SPFx-dev.aspx`. To use your own, create a page on any site where you have edit rights (**+ New → Page**, blank, then **Publish**). Then put its URL in `initialPage` in `config/serve.json` (keep `{tenantDomain}`) and in `.vscode/launch.json`. Creating a site isn't needed: most users can't, and a page is enough.

**Each time:**

1. `npm start`, and wait for the build to finish. It opens `https://scottlogic.sharepoint.com/sites/NewcastleSocialCommittee/SitePages/SPFx-dev.aspx` with the debug parameters already added.
2. Choose **Allow** when Edge asks about local network access, then **Load debug scripts**. The debugging toolbar appears on the page.
3. The first time, **Edit** the page → **+** inside a section → search for the web part's title → **Save as draft**. Don't republish: the page is on a shared site, and other visitors would see a broken web part whenever your dev server isn't running. The draft is only visible to editors, and the web part stays in it from then on.

It rebuilds when you save, so reload the page to see changes. **Stop** on the debug toolbar, or adding `reset=true` to the URL, turns local code off.

**Testing on a different page:** run `npm start -- --nobrowser`, then open any modern page you can edit and add this to the end of its URL (use `&` if the URL already has a `?`):
`?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/build/manifests.js`

## 4. Build the package

1. **Stop `npm start` first.** Otherwise its dev files end up in the package.
2. Bump `version` in `config/package-solution.json` (for example `1.0.0.0` → `1.0.1.0`).
3. `npm run build`.
4. Your package is `sharepoint/solution/intranet-html-pages.sppkg`. It should be about 15–20 KB plus a few KB per page. If it's hundreds of KB, the dev server was still running.

## 5. Upload to the App Catalog

1. Open the catalog:
   - **Tenant:** SharePoint admin center → **More features** → **Apps** → **Open**.
   - **Site collection:** Site contents → **Apps for SharePoint**.
2. **Upload** the `.sppkg`, and confirm **replace** if an older version exists.
3. In the trust dialog, choose **Only enable this app** or **Enable this app and add it to all sites**, then **Enable app**.
4. Check that the **App Package Error Message** column is empty.

## 6. Install on a site

Skip this if you added the app to all sites. Otherwise, on the site: gear icon → **Add an app** → **intranet-html-pages**. Then wait until it shows as installed in **Site contents**.

## 7. Put a web part on a page

- **Normal page:** gear icon → **Add a page**. Add a **Full-width section** (communication sites) or **One column**. Click the **+** *inside* the section (the one on the left edge adds a section), search for the title, then **Publish**.
- **Full-page app** (PnP PowerShell; `-ClientId` is the PnP Entra app registered in our tenant, so ask IT):
  ```powershell
  Connect-PnPOnline https://scottlogic.sharepoint.com/sites/<site> -ClientId <app ID> -Interactive
  Add-PnPPage -Name "Hello" -LayoutType SingleWebPartAppPage
  Add-PnPPageWebPart -Page "Hello" -Component "Hello world"
  Set-PnPPage -Identity "Hello" -Publish
  ```

## 8. Update or remove

- **Update or add pages:** edit the HTML or run `add-page`, bump `version`, rebuild, and re-upload. Existing pages pick up the change.
- **Remove a web part:** delete it from its pages, remove its `pages.json` entry, then rebuild and re-upload.
- **Remove everything:** remove the app from each site (Site contents), then delete the package from the App Catalog.

## 9. Troubleshooting

| Symptom | Fix |
|---|---|
| Local web part missing from the toolbox | Make sure `npm start` is running. Open <https://localhost:4321/temp/build/manifests.js> and accept any certificate warning. Check that the site's **Local network access** permission is **Allow** and that you chose **Load debug scripts**. |
| `npm start` warns that `SPFX_SERVE_TENANT_DOMAIN is not set`, or the browser opens a URL containing `{tenantDomain}` | Create `.env` from `.env.example` (section 3). |
| The debug page returns 404 | The page in `initialPage` (`config/serve.json`) doesn't exist, or its file name differs from its title. Open the page yourself and copy the URL from the address bar into `initialPage` (section 3). |
| Empty canvas, or **+** shows layout options | That's normal. Use the **+** inside a section (section 7). |
| `getPageBasicInfo` / `ComponentPerfLogger` console errors | SharePoint's own telemetry. They're harmless. |
| `Invalid pages.json` or `at most one inline <script>` | Fix what the message lists. |
| Package is hundreds of KB | Stop `npm start` and rebuild. |
| Web part errors on a deployed page | Check the app is installed on that site, and that its `id` hasn't changed. |
