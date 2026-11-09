#!/usr/bin/env node
/*
 * Check the published launcher website.
 *
 * Usage (run from the project folder):
 *   node scripts/check-website.mjs https://<your-hostname>/
 *
 * It checks that the page loads over HTTPS, that the security headers that let Teams show the page
 * are present, that the TeamsJS library downloads and matches its tamper seal (integrity value),
 * that the page files load, and that the privacy, terms and support pages exist.
 * It only reads. It changes nothing. Needs Node.js 18 or later.
 */
import { createHash } from "node:crypto";
const input = process.argv[2];
const allowHttp = process.argv.includes("--allow-http"); // for local preview testing only

if (!input) {
  console.error("\nAdd the website address, for example: node scripts/check-website.mjs https://gentle-sand-0b1c2d30f.5.azurestaticapps.net/\n");
  process.exit(1);
}

const base = new URL(input);
let failures = 0;
let warnings = 0;
const pass = (text) => console.log(`PASS  ${text}`);
const warn = (text) => {
  warnings += 1;
  console.log(`WARN  ${text}`);
};
const fail = (text) => {
  failures += 1;
  console.log(`FAIL  ${text}`);
};

async function get(path) {
  const url = new URL(path, base);
  const response = await fetch(url, { redirect: "follow", headers: { "Cache-Control": "no-cache" } });
  const body = await response.text();
  return { url, response, body };
}

try {
  if (base.protocol === "https:") pass("The address uses HTTPS.");
  else if (allowHttp) warn("The address uses HTTP (allowed for a local preview only).");
  else fail("The address must start with https://");

  const home = await get("/");
  if (home.response.status === 200) pass("The launcher page loads (status 200).");
  else fail(`The launcher page returned status ${home.response.status}.`);

  const csp = home.response.headers.get("content-security-policy") || "";
  const scriptTag = home.body.match(/<script[^>]*\ssrc="([^"]*MicrosoftTeams(?:\.min)?\.js)"[^>]*>/i);
  const teamsJsUrl = scriptTag ? new URL(scriptTag[1], base) : null;
  const teamsJsIsExternal = Boolean(teamsJsUrl && teamsJsUrl.origin !== base.origin);

  if (!csp) {
    fail("No Content-Security-Policy header. Check that staticwebapp.config.json was deployed with the page.");
  } else {
    const ancestors = (csp.split(";").find((part) => part.trim().startsWith("frame-ancestors")) || "").trim();
    const needed = ["https://teams.microsoft.com", "https://*.teams.microsoft.com", "https://*.cloud.microsoft"];
    const missing = needed.filter((domain) => !ancestors.includes(domain));
    if (missing.length === 0) pass("Teams is allowed to show the page (frame-ancestors).");
    else fail(`frame-ancestors is missing: ${missing.join(", ")}`);
    if (teamsJsIsExternal) {
      const scriptRule = (csp.split(";").find((part) => part.trim().startsWith("script-src")) || "").trim();
      if (scriptRule.includes(teamsJsUrl.origin)) pass(`The TeamsJS address is allowed (script-src ${teamsJsUrl.origin}).`);
      else fail(`script-src does not allow ${teamsJsUrl.origin}, so the browser would block the Teams library.`);
    }
  }

  const frameOptions = (home.response.headers.get("x-frame-options") || "").toUpperCase();
  if (frameOptions === "DENY" || frameOptions === "SAMEORIGIN") fail(`X-Frame-Options is ${frameOptions}, which stops Teams from showing the page.`);
  else pass("No X-Frame-Options header is blocking Teams.");

  if ((home.response.headers.get("x-content-type-options") || "").toLowerCase() === "nosniff") pass("X-Content-Type-Options is set to nosniff.");
  else warn("X-Content-Type-Options is missing.");

  if (!teamsJsUrl) {
    fail("The page does not load the TeamsJS library (MicrosoftTeams.min.js).");
  } else {
    const integrity = ((scriptTag[0].match(/\sintegrity="([^"]+)"/i) || [])[1] || "").trim().split(/\s+/)[0];
    const versionMatch = teamsJsUrl.pathname.match(/teams-js\/([\d.]+)\//);
    const label = versionMatch ? `TeamsJS ${versionMatch[1]}` : "TeamsJS";
    let library = null;
    try {
      library = await fetch(teamsJsUrl, { headers: { Origin: base.origin } });
    } catch (error) {
      fail(`Could not download ${label} from ${teamsJsUrl}: ${error.message}. Check your internet connection or web filter.`);
    }
    if (library && library.status !== 200) {
      fail(`${label} did not download from ${teamsJsUrl} (status ${library.status}).`);
    } else if (library) {
      const bytes = Buffer.from(await library.arrayBuffer());
      pass(`${label} downloads (${bytes.length} bytes) from ${teamsJsUrl.origin === base.origin ? "your own site" : teamsJsUrl.origin}.`);
      if (integrity) {
        const cut = integrity.indexOf("-");
        const algorithm = integrity.slice(0, cut).toLowerCase();
        const expected = integrity.slice(cut + 1);
        let actual = "";
        try {
          actual = createHash(algorithm).update(bytes).digest("base64");
        } catch (error) {
          fail(`The integrity value uses "${algorithm}", which this checker does not understand.`);
        }
        if (actual && actual === expected) pass(`${label} matches its tamper seal (integrity ${algorithm}).`);
        else if (actual) fail(`${label} does NOT match its tamper seal, so browsers will refuse it. Copy the script tag again from the TeamsJS README. Page expects ${expected}; the file gives ${actual}.`);
        if (teamsJsIsExternal && !(library.headers.get("access-control-allow-origin") || "")) {
          warn(`${teamsJsUrl.origin} did not send a CORS header. With crossorigin="anonymous" some browsers may refuse the file.`);
        }
      } else if (teamsJsIsExternal) {
        fail("TeamsJS comes from another website without a tamper seal (integrity). Add the integrity value from the TeamsJS README.");
      } else {
        pass("TeamsJS is a copy on your own site, so no tamper seal is needed.");
      }
    }
  }

  if (home.body.includes("https://myapplications.microsoft.com/")) pass("The page points to https://myapplications.microsoft.com/");
  else fail("The My Apps address was not found in the page.");

  for (const path of ["/app.js", "/styles.css", "/favicon.svg", "/privacy.html", "/terms.html", "/support.html"]) {
    const result = await get(path);
    if (result.response.status === 200) pass(`${path} loads.`);
    else fail(`${path} returned status ${result.response.status}.`);
  }

  if (base.protocol === "https:") {
    const plain = new URL("/", base);
    plain.protocol = "http:";
    try {
      const redirect = await fetch(plain, { redirect: "manual" });
      const location = redirect.headers.get("location") || "";
      if ([301, 302, 307, 308].includes(redirect.status) && location.startsWith("https://")) pass("Plain http:// requests are sent to https://.");
      else warn(`http:// returned status ${redirect.status} instead of a redirect to https://.`);
    } catch (error) {
      warn("Could not test the http:// to https:// redirect.");
    }
  }
} catch (error) {
  fail(`Could not reach the site: ${error.message}`);
}

console.log(`\nResult: ${failures} failed, ${warnings} warnings.`);
process.exit(failures > 0 ? 1 : 0);
