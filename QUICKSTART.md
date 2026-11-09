# My Apps for Microsoft Teams - starter kit

A tiny Teams personal app that works as a doorway: **Teams left rail > My Apps > Open My Apps > https://myapplications.microsoft.com/ opens in the web browser.**

No bot, no API, no database, no backend, no sign-in code, and no Entra app registration.

Design: Shane's Design System, **Black & Purple** theme (Obsidian #111111, Royal Purple #7B1FA2, Silver #ECEFF1, Luminous Purple #B388FF, Deep Black #090909).

The README explains what the kit is, how to get it, and the legal notice.

## What is in this folder

| Path | What it is |
|---|---|
| `web/index.html` | The launcher page Teams shows. |
| `web/app.js` | The logic: finds Teams, then opens My Apps with TeamsJS `app.openLink()`. Falls back to a normal link outside Teams. |
| `web/styles.css` | Black & Purple styling (colors, spacing, focus rings, phone layout). |
| `web/staticwebapp.config.json` | Azure Static Web Apps settings: security headers that let Teams show the page, caching, 404 page. |
| `web/privacy.html`, `terms.html`, `support.html` | Pages the Teams manifest must link to. **Sample wording: replace before production.** |
| `web/404.html`, `web/favicon.svg` | Not-found page and browser tab icon. |
| `appPackage/manifest.json` | Teams app manifest template (schema 1.30) with placeholders. |
| `appPackage/color.png` | 192 x 192 color icon. |
| `appPackage/outline.png` | 32 x 32 white-on-transparent outline icon. |
| `scripts/preview-local.mjs` | Preview the page on your computer with the same headers Azure sends. |
| `scripts/build-teams-package.mjs` | Fills in the placeholders, checks for mistakes, and builds the Teams zip. |
| `scripts/check-website.mjs` | Checks the published site: HTTPS, the headers that let Teams show the page, the TeamsJS tamper seal, files and policy pages. |
| `tests/browser-smoke-test.cjs` | Optional automated browser test (needs Playwright). |

## The short version

1. Preview locally: `node scripts/preview-local.mjs` then open http://localhost:8080
2. Create an Azure Static Web App (plan: Free for the POC, deployment source: Other). Copy its URL and deployment token.
3. Publish the `web` folder (PowerShell):
   ```powershell
   $env:SWA_CLI_DEPLOYMENT_TOKEN = Read-Host "Paste the deployment token"
   npx --yes @azure/static-web-apps-cli deploy ./web --env production
   Remove-Item Env:SWA_CLI_DEPLOYMENT_TOKEN
   ```
4. Check the site: `node scripts/check-website.mjs https://<your-hostname>/`
5. Build the Teams package: `node scripts/build-teams-package.mjs --host <your-hostname> --org "<organization name>"`
6. In Teams: **Apps > Manage your apps > Upload an app > Upload a custom app**, pick `dist/MyAppsTeamsApp-1.0.0.zip`.
7. Admin: Teams admin center **Teams apps > Manage apps > Upload new app**, set availability, then pin it with **Teams apps > Setup policies**.

## Placeholders

| Placeholder | Replace with |
|---|---|
| `REPLACE-WITH-YOUR-HOSTNAME` | Your Static Web App hostname, for example `gentle-sand-0b1c2d30f.5.azurestaticapps.net` (no `https://`, no slash) |
| `REPLACE-WITH-YOUR-ORGANIZATION` | Your organization name as users should see it (32 characters or fewer) |

The build script fills both in for you. If you edit by hand, use Find and Replace in VS Code (Ctrl+H) and replace every match.

## Settings you can change (top of `web/app.js`)

| Setting | Default | Meaning |
|---|---|---|
| `myAppsUrl` | `https://myapplications.microsoft.com/` | Where the button goes. |
| `environmentLabel` | `Test version` | Banner at the top of the page. Set to `""` to hide it. |
| `pageVersion` | `1.0.0` | Shown at the bottom of the page. Raise it with every release. |
| `teamsCheckTimeoutMs` | `4000` | How long the page waits for Teams to answer. |

## Testing Option B (automatic launch)

Add `?launch=auto` to the page address and the page tries to open My Apps as soon as it loads, without a click. To test it inside Teams, build the separate test app:

```
node scripts/build-teams-package.mjs --host <your-hostname> --org "<organization name>" --option-b
```

This makes `dist/MyAppsTeamsApp-OptionB-1.0.0.zip`: an app called "My Apps B" with its own app ID and `?launch=auto` in its `contentUrl`, so you can compare Option A and Option B side by side. Remove it after the POC.

## Updating the TeamsJS library

`web/index.html` loads TeamsJS 2.58.0 from Microsoft's CDN with an `integrity` value (a tamper seal). To change versions, copy both the new address and its matching integrity value from the TeamsJS README on GitHub (OfficeDev/microsoft-teams-library-js, packages/teams-js). Then run `node scripts/check-website.mjs <your site address>`: it downloads the library and confirms the file matches the seal. If the seal does not match, browsers block the library and the page shows "TeamsJS library: not loaded" under Details for testers.

If your network blocks the CDN, host the library yourself: download `MicrosoftTeams.min.js` from the CDN address into the `web` folder, change the tag to `<script src="MicrosoftTeams.min.js" defer></script>`, and publish again. The security header already allows scripts from your own site.

## Version history

| Version | Date | Change |
|---|---|---|
| 1.0.0 | October 28, 2026 | First POC build: launcher page, Teams manifest 1.30, icons, scripts (including the TeamsJS tamper-seal check), tests. |
