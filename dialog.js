/* global Office */
(function () {
  "use strict";

  Office.onReady(() => {
    renderGroups();
    document.getElementById("backBtn").addEventListener("click", () => respond("cancel"));
    document.getElementById("sendBtn").addEventListener("click", () => respond("send"));
  });

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
