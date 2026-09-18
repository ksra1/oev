/* global Office */
(function () {
  "use strict";

  Office.onReady(() => {
    applyOfficeTheme();
    renderGroups();
    document.getElementById("backBtn").addEventListener("click", () => respond("cancel"));
    document.getElementById("sendBtn").addEventListener("click", () => respond("send"));
  });

  // See taskpane.js for why this reads Office.context.officeTheme instead of
  // relying on the OS/browser prefers-color-scheme.
  function applyOfficeTheme() {
    const theme = Office.context.officeTheme;
    const bg = theme && theme.bodyBackgroundColor;
    if (!bg) return;
    document.documentElement.setAttribute("data-theme", isDarkColor(bg) ? "dark" : "light");
  }

  function isDarkColor(hex) {
    const match = /^#?([0-9a-f]{6})$/i.exec(hex);
    if (!match) return false;
    const n = parseInt(match[1], 16);
    const r = (n >> 16) & 255;
    const g = (n >> 8) & 255;
    const b = n & 255;
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance < 0.5;
  }

  function getGroupsFromUrl() {
    const params = new URLSearchParams(window.location.search);
    const raw = params.get("data");
    if (!raw) return [];
    try {
      return JSON.parse(decodeURIComponent(raw));
    } catch (e) {
      return [];
    }
  }

  function renderGroups() {
    const container = document.getElementById("groups");
    const groups = getGroupsFromUrl();
    groups.forEach((g) => {
      const row = document.createElement("div");
      row.className = "ov-domain-row";

      const badge = document.createElement("div");
      badge.className = "ov-domain-badge";
      badge.textContent = g.domain === "unknown" ? "?" : g.domain.slice(0, 2);

      const label = document.createElement("div");
      label.textContent = `${g.domain} — ${g.count} recipient${g.count === 1 ? "" : "s"}`;

      row.appendChild(badge);
      row.appendChild(label);
      container.appendChild(row);
    });
  }

  function respond(action) {
    Office.context.ui.messageParent(JSON.stringify({ action }));
  }
})();
