/*
 * Optional automated browser test for the My Apps launcher page.
 * It simulates the Teams host, so it proves the page logic, NOT real Teams behavior.
 * Real Teams clients still need the manual tests in the playbook.
 *
 * One-time setup (from the project folder):
 *   npm install --no-save playwright
 *   npx playwright install chromium
 * Run:
 *   1. In one terminal:  node scripts/preview-local.mjs
 *   2. In another:       node tests/browser-smoke-test.cjs
 * Screenshots are saved in dist/screenshots.
 */
const { chromium } = require("playwright");
const { mkdirSync } = require("fs");
const path = require("path");

const BASE = process.env.LAUNCHER_URL || "http://localhost:8080/";
const MYAPPS = "https://myapplications.microsoft.com/";
const SHOTS = path.join(__dirname, "..", "dist", "screenshots");
mkdirSync(SHOTS, { recursive: true });
const results = [];

function record(name, ok, detail = "") {
  results.push({ name, ok });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  [" + detail + "]" : ""}`);
}

function teamsMock(mode) {
  window.__openLinkCalls = [];
  window.__notifySuccessCalls = 0;
  window.microsoftTeams = {
    version: "2.58.0 (simulated)",
    app: {
      initialize: () => (mode === "hang" ? new Promise(() => {}) : Promise.resolve()),
      notifySuccess: () => {
        window.__notifySuccessCalls += 1;
        return Promise.resolve();
      },
      getContext: () =>
        Promise.resolve({
          app: { theme: "dark", host: { name: "Teams", clientType: "desktop" } },
          user: { userPrincipalName: "pilot.user@contoso.example" }
        }),
      registerOnThemeChangeHandler: (handler) => {
        window.__themeHandler = handler;
      },
      openLink: (url) => {
        window.__openLinkCalls.push(url);
        return mode === "reject" ? Promise.reject(new Error("simulated failure")) : Promise.resolve();
      }
    }
  };
}

(async () => {
  const browser = await chromium.launch();

  async function open({ mock = null, viewport = { width: 1280, height: 900 }, path = "" } = {}) {
    const context = await browser.newContext({ viewport, deviceScaleFactor: 2 });
    const myAppsRequests = [];
    context.on("request", (request) => {
      if (request.url().startsWith(MYAPPS)) myAppsRequests.push(request.url());
    });
    // Block the real TeamsJS download and My Apps itself: the test simulates Teams and only checks the request.
    await context.route(/^https:\/\/(res\.cdn\.office\.net|myapplications\.microsoft\.com)\//, (route) => route.abort());
    const page = await context.newPage();
    const consoleErrors = [];
    const csp = [];
    page.on("console", (message) => {
      if (message.type() === "error" && !/Failed to load resource|ERR_FAILED/.test(message.text())) consoleErrors.push(message.text());
    });
    page.on("pageerror", (error) => consoleErrors.push(String(error)));
    await page.exposeFunction("__csp", (text) => csp.push(text));
    await page.addInitScript(() => {
      document.addEventListener("securitypolicyviolation", (event) => window.__csp(`${event.violatedDirective} ${event.blockedURI}`));
    });
    if (mock) await page.addInitScript(teamsMock, mock);
    await page.goto(BASE + path);
    return { context, page, consoleErrors, csp, myAppsRequests };
  }

  const text = async (page, selector) => ((await page.locator(selector).textContent()) || "").trim();
  const settle = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  // A. Plain browser (TeamsJS not loaded)
  {
    const t = await open();
    await t.page.waitForFunction(() => document.getElementById("connection-badge").dataset.state === "browser", null, { timeout: 6000 });
    record("Browser: badge shows Web browser", (await text(t.page, "#connection-text")) === "Web browser");
    record("Browser: tester details say TeamsJS not loaded", (await text(t.page, "#diag-teamsjs")) === "not loaded");
    record("Browser: mode banner visible with label", (await text(t.page, "#mode-label")) === "Test version");
    await t.page.screenshot({ path: `${SHOTS}/launcher-browser-desktop.png` });
    const popup = t.context.waitForEvent("page", { timeout: 5000 }).catch(() => null);
    await t.page.click("#open-myapps");
    const newPage = await popup;
    await settle(500);
    record("Browser: click opens a new tab for My Apps", Boolean(newPage) && t.myAppsRequests.length >= 1, `requests=${t.myAppsRequests.length}`);
    record("Browser: status says new browser tab", (await text(t.page, "#status-text")).includes("new browser tab"));
    record("Browser: no CSP violations", t.csp.length === 0, t.csp.join(" | "));
    record("Browser: no script errors", t.consoleErrors.length === 0, t.consoleErrors.join(" | "));
    await t.context.close();
  }

  // B. Inside Teams (simulated host): openLink succeeds
  {
    const t = await open({ mock: "ok" });
    await t.page.waitForFunction(() => document.getElementById("connection-badge").dataset.state === "teams", null, { timeout: 6000 });
    await settle(300);
    record("Teams: badge shows Microsoft Teams", (await text(t.page, "#connection-text")) === "Microsoft Teams");
    record("Teams: notifySuccess called once", (await t.page.evaluate(() => window.__notifySuccessCalls)) === 1);
    record("Teams: tester details show host, client and theme",
      (await text(t.page, "#diag-host")) === "Teams" && (await text(t.page, "#diag-client")) === "desktop" && (await text(t.page, "#diag-theme")) === "dark");
    await t.page.click("#open-myapps");
    await settle(400);
    const calls = await t.page.evaluate(() => window.__openLinkCalls);
    record("Teams: click calls app.openLink with the exact My Apps URL", calls.length === 1 && calls[0] === MYAPPS, JSON.stringify(calls));
    record("Teams: plain link did not also open a tab", t.myAppsRequests.length === 0, `requests=${t.myAppsRequests.length}`);
    record("Teams: success message shown", (await text(t.page, "#status-text")).includes("opening in your web browser"));
    const body = await t.page.evaluate(() => document.body.innerText);
    record("Teams: no user identity shown on the page", !body.includes("pilot.user"));
    await t.page.click("#diagnostics summary");
    await settle(200);
    await t.page.screenshot({ path: `${SHOTS}/launcher-teams-details.png`, fullPage: true });
    record("Teams: no CSP violations", t.csp.length === 0, t.csp.join(" | "));
    record("Teams: no script errors", t.consoleErrors.length === 0, t.consoleErrors.join(" | "));
    await t.context.close();
  }

  // C. Inside Teams: openLink fails, browser fallback is used
  {
    const t = await open({ mock: "reject" });
    await t.page.waitForFunction(() => document.getElementById("connection-badge").dataset.state === "teams", null, { timeout: 6000 });
    await t.page.click("#open-myapps");
    await settle(800);
    const kind = await t.page.locator("#status").getAttribute("data-kind");
    record("Teams failure: falls back to a normal browser window", t.myAppsRequests.length >= 1 && kind === "warning", `kind=${kind}, requests=${t.myAppsRequests.length}`);
    await t.page.screenshot({ path: `${SHOTS}/launcher-teams-fallback.png` });
    await t.context.close();
  }

  // D. Teams never answers: page gives up after the timeout and behaves like a browser
  {
    const t = await open({ mock: "hang" });
    const started = Date.now();
    await t.page.waitForFunction(() => document.getElementById("connection-badge").dataset.state === "browser", null, { timeout: 9000 });
    const waited = Date.now() - started;
    record("Teams silent: switches to browser mode after the timeout", waited >= 3000, `${waited} ms`);
    await t.context.close();
  }

  // E. Option B: automatic launch inside Teams
  {
    const t = await open({ mock: "ok", path: "?launch=auto" });
    await settle(1500);
    const calls = await t.page.evaluate(() => window.__openLinkCalls);
    record("Option B: opens My Apps without a click", calls.length === 1 && calls[0] === MYAPPS, JSON.stringify(calls));
    record("Option B: tester details show automatic mode", (await text(t.page, "#diag-mode")).startsWith("automatic"));
    await t.context.close();
  }

  // F. Keyboard focus is clearly visible
  {
    const t = await open();
    await t.page.waitForFunction(() => document.getElementById("connection-badge").dataset.state === "browser", null, { timeout: 6000 });
    await t.page.keyboard.press("Tab");
    const focused = await t.page.evaluate(() => document.activeElement && document.activeElement.id);
    const outline = await t.page.evaluate(() => getComputedStyle(document.activeElement).outlineStyle);
    record("Keyboard: first Tab lands on Open My Apps", focused === "open-myapps", focused);
    record("Keyboard: focus ring is visible", outline !== "none", outline);
    await t.page.screenshot({ path: `${SHOTS}/launcher-focus.png`, clip: { x: 340, y: 230, width: 600, height: 420 } });
    await t.context.close();
  }

  // G. Phone-sized screen: no sideways scrolling
  {
    const t = await open({ viewport: { width: 390, height: 844 } });
    await t.page.waitForFunction(() => document.getElementById("connection-badge").dataset.state === "browser", null, { timeout: 6000 });
    const overflow = await t.page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    record("Phone: no horizontal overflow", overflow <= 0, `${overflow}px`);
    await t.page.screenshot({ path: `${SHOTS}/launcher-phone.png` });
    await t.context.close();
  }

  // H. Supporting pages
  {
    const t = await open();
    for (const path of ["privacy.html", "terms.html", "support.html"]) {
      const response = await t.page.goto(BASE + path);
      record(`Page ${path} loads`, response.status() === 200 && (await t.page.locator("h1").count()) === 1);
    }
    const missing = await t.page.goto(BASE + "does-not-exist");
    record("Unknown address returns the 404 page", missing.status() === 404 && (await text(t.page, "h1")) === "Page not found");
    await t.page.goto(BASE + "privacy.html");
    await t.page.screenshot({ path: `${SHOTS}/privacy-page.png` });
    record("Supporting pages: no CSP violations", t.csp.length === 0, t.csp.join(" | "));
    await t.context.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed} passed, ${failed} failed`);
  process.exit(failed ? 1 : 0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
