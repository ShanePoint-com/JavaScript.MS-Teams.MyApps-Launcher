/*
 * My Apps launcher for Microsoft Teams
 * Design: Shane's Design System, Black & Purple theme
 *
 * What this file does, in plain words:
 *   1. When the page opens, it checks whether it is running inside Microsoft Teams.
 *   2. Inside Teams, "Open My Apps" asks Teams to open My Apps in the web browser
 *      (the TeamsJS function app.openLink).
 *   3. Outside Teams, "Open My Apps" is a normal link that opens a new browser tab.
 *   4. If anything goes wrong, the page says what happened and offers a direct link.
 *
 * This page stores nothing, signs no one in, reads no user identity and calls no APIs.
 */
(() => {
  "use strict";

  // ===== Settings you can change ==========================================
  const SETTINGS = Object.freeze({
    myAppsUrl: "https://myapplications.microsoft.com/",
    environmentLabel: "Test version", // Banner at the top. Set to "" to hide it.
    pageVersion: "1.0.0",             // Raise this each time you publish a change.
    teamsCheckTimeoutMs: 4000         // How long to wait for Teams to answer.
  });
  // ========================================================================

  const byId = (id) => document.getElementById(id);
  const ui = {
    openButton: byId("open-myapps"),
    directLink: byId("direct-link"),
    hint: byId("open-hint"),
    status: byId("status"),
    statusLabel: byId("status-label"),
    statusText: byId("status-text"),
    modeBanner: byId("mode-banner"),
    modeLabel: byId("mode-label"),
    badge: byId("connection-badge"),
    badgeText: byId("connection-text"),
    pageVersion: byId("page-version"),
    diag: {
      runningIn: byId("diag-running-in"),
      host: byId("diag-host"),
      client: byId("diag-client"),
      theme: byId("diag-theme"),
      mode: byId("diag-mode"),
      teamsjs: byId("diag-teamsjs"),
      version: byId("diag-version")
    }
  };

  const STATUS_LABELS = Object.freeze({
    info: "Note:",
    loading: "Working:",
    success: "Done:",
    warning: "Heads up:",
    error: "Problem:"
  });

  // "?launch=auto" in the address turns on the Option B test (open without a click).
  const launchMode = new URLSearchParams(window.location.search).get("launch") === "auto" ? "auto" : "button";

  let inTeams = false;
  let busy = false;

  function showStatus(kind, message) {
    ui.status.dataset.kind = kind;
    ui.statusLabel.textContent = STATUS_LABELS[kind] || "";
    ui.statusText.textContent = message;
    ui.status.hidden = false;
  }

  function setConnection(state, text) {
    ui.badge.dataset.state = state;
    ui.badgeText.textContent = text;
    ui.diag.runningIn.textContent = text;
  }

  function setBusy(isBusy) {
    busy = isBusy;
    ui.openButton.setAttribute("aria-disabled", String(isBusy));
    ui.openButton.classList.toggle("is-busy", isBusy);
  }

  function withTimeout(promise, ms) {
    let timer;
    const timeout = new Promise((_, reject) => {
      timer = window.setTimeout(() => reject(new Error("timed_out")), ms);
    });
    return Promise.race([promise, timeout]).finally(() => window.clearTimeout(timer));
  }

  function teamsLibraryLoaded() {
    const teams = window.microsoftTeams;
    return Boolean(teams && teams.app && typeof teams.app.initialize === "function");
  }

  // Step 1: try to connect to Teams. Resolves true inside Teams, false anywhere else.
  async function connectToTeams() {
    if (!teamsLibraryLoaded()) {
      return false;
    }
    try {
      await withTimeout(Promise.resolve().then(() => window.microsoftTeams.app.initialize()), SETTINGS.teamsCheckTimeoutMs);
    } catch (error) {
      return false;
    }
    inTeams = true;
    // Tell Teams the page is ready (this hides a loading spinner if Teams shows one).
    try {
      const result = window.microsoftTeams.app.notifySuccess();
      if (result && typeof result.catch === "function") {
        result.catch(() => {});
      }
    } catch (error) {
      // Not important: the page works without it.
    }
    return true;
  }

  function applyTeamsTheme(theme) {
    document.body.dataset.teamsTheme = theme || "default";
    ui.diag.theme.textContent = theme || "unknown";
  }

  // Read harmless details about the Teams client for the "Details for testers" panel.
  // Only host name, client type and theme are read. No user information is used.
  async function readTeamsDetails() {
    try {
      const context = await withTimeout(window.microsoftTeams.app.getContext(), SETTINGS.teamsCheckTimeoutMs);
      const appInfo = (context && context.app) || {};
      const host = appInfo.host || {};
      ui.diag.host.textContent = host.name || "unknown";
      ui.diag.client.textContent = host.clientType || "unknown";
      applyTeamsTheme(appInfo.theme);
      window.microsoftTeams.app.registerOnThemeChangeHandler(applyTeamsTheme);
    } catch (error) {
      ui.diag.client.textContent = "not available";
    }
  }

  // Plain browser fallback. Returns true when a new window or tab opened.
  function openInNewWindow() {
    const newWindow = window.open(SETTINGS.myAppsUrl, "_blank");
    if (newWindow) {
      try {
        newWindow.opener = null;
      } catch (error) {
        // Ignore: some browsers do not allow this, and it is only a safety extra.
      }
      return true;
    }
    return false;
  }

  // Step 2 (inside Teams): ask Teams to open My Apps outside the Teams window.
  async function openMyAppsThroughTeams() {
    if (busy) {
      return;
    }
    setBusy(true);
    showStatus("loading", "Asking Teams to open My Apps in your web browser…");
    try {
      await window.microsoftTeams.app.openLink(SETTINGS.myAppsUrl);
      showStatus("success", "My Apps is opening in your web browser. Look for a new browser window or tab.");
    } catch (error) {
      console.warn("[My Apps launcher] Teams could not open the link.", error);
      if (openInNewWindow()) {
        showStatus("warning", "Teams could not open the link itself, so a regular browser window was used instead.");
      } else {
        showStatus("error", "My Apps did not open automatically. Nothing was changed. Use the \"Open My Apps directly\" link below.");
      }
    } finally {
      setBusy(false);
    }
  }

  function handleOpenClick(event) {
    if (busy) {
      event.preventDefault();
      return;
    }
    if (inTeams) {
      event.preventDefault(); // Teams opens the link instead of the plain web link.
      openMyAppsThroughTeams();
      return;
    }
    // Not in Teams (or Teams has not answered yet): the normal link opens a new tab.
    showStatus("success", "My Apps is opening in a new browser tab.");
  }

  // Option B test: try to open My Apps as soon as the page loads, without a click.
  async function autoLaunch() {
    if (inTeams) {
      await openMyAppsThroughTeams();
      return;
    }
    if (openInNewWindow()) {
      showStatus("success", "My Apps opened in a new browser tab.");
    } else {
      showStatus("warning", "Your browser blocked the automatic launch. Select Open My Apps instead.");
    }
  }

  async function start() {
    ui.pageVersion.textContent = SETTINGS.pageVersion;
    ui.diag.version.textContent = SETTINGS.pageVersion;
    ui.diag.mode.textContent = launchMode === "auto" ? "automatic (Option B)" : "button (Option A)";
    ui.openButton.href = SETTINGS.myAppsUrl;
    ui.directLink.href = SETTINGS.myAppsUrl;

    if (SETTINGS.environmentLabel) {
      ui.modeLabel.textContent = SETTINGS.environmentLabel;
      ui.modeBanner.hidden = false;
    } else {
      ui.modeBanner.hidden = true;
    }

    ui.openButton.addEventListener("click", handleOpenClick);

    const libraryLoaded = teamsLibraryLoaded();
    ui.diag.teamsjs.textContent = libraryLoaded ? (window.microsoftTeams.version || "loaded") : "not loaded";

    setConnection("checking", "Checking for Microsoft Teams…");
    const connected = await connectToTeams();

    if (connected) {
      setConnection("teams", "Microsoft Teams");
      ui.hint.textContent = "Opens in your web browser, outside Teams.";
      readTeamsDetails();
    } else {
      setConnection("browser", "Web browser");
      ui.hint.textContent = "Opens in a new browser tab.";
      ui.diag.host.textContent = "none";
      ui.diag.client.textContent = "web browser";
      ui.diag.theme.textContent = "not in Teams";
    }

    if (launchMode === "auto") {
      await autoLaunch();
    }
  }

  start();
})();
