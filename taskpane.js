/* global Office, ValidatorCore */
(function () {
  "use strict";

  const bannerEl = document.getElementById("banner");
  const groupsEl = document.getElementById("groups");
  const refreshBtn = document.getElementById("refreshBtn");

  Office.onReady(() => {
    applyOfficeTheme();
    render();
    attachChangeHandlers();
    refreshBtn.addEventListener("click", render);
  });

  function attachChangeHandlers() {
    const item = Office.context.mailbox.item;
    ["to", "cc", "bcc"].forEach((field) => {
      if (item[field] && item[field].addHandlerAsync) {
        item[field].addHandlerAsync(Office.EventType.RecipientsChanged, render);
      }
    });
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyOfficeTheme);
    }
  }

  // Office.context.officeTheme.isDarkTheme isn't supported in Outlook (per Microsoft's docs),
  // so infer light/dark from the host's own background color instead of the OS/browser
  // prefers-color-scheme, which can disagree with Outlook's actual theme setting.
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

  function getRecipientsAsync() {
    const item = Office.context.mailbox.item;
    return Promise.all([
      getFieldAsync(item, "to"),
      getFieldAsync(item, "cc"),
      getFieldAsync(item, "bcc"),
    ]).then(([to, cc, bcc]) => ValidatorCore.flattenRecipients(to, cc, bcc));
  }

  function getFieldAsync(item, field) {
    return new Promise((resolve) => {
      if (!item[field] || !item[field].getAsync) {
        resolve([]);
        return;
      }
      item[field].getAsync((result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
      });
    });
  }

  function initials(nameOrEmail) {
    const base = (nameOrEmail || "").trim();
    if (!base) return "?";
    const parts = base.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
    return base.slice(0, 2).toUpperCase();
  }

  function render() {
    getRecipientsAsync().then((recipients) => {
      renderBanner(recipients);
      renderGroups(recipients);
    });
  }

  function renderBanner(recipients) {
    bannerEl.innerHTML = "";
    if (recipients.length === 0) {
      bannerEl.appendChild(
        makeBanner("empty", "✉", "Add recipients (To / Cc / Bcc) to see them grouped by company domain.")
      );
      return;
    }
    const domainCount = ValidatorCore.distinctDomainCount(recipients);
    if (domainCount > 1) {
      bannerEl.appendChild(
        makeBanner(
          "warn",
          "⚠",
          `Recipients span ${domainCount} different domains. Double-check before sending.`
        )
      );
    } else {
      bannerEl.appendChild(makeBanner("ok", "✓", "All recipients share one domain."));
    }
  }

  function makeBanner(kind, icon, text) {
    const div = document.createElement("div");
    div.className = `ov-banner ${kind}`;
    const iconSpan = document.createElement("span");
    iconSpan.className = "ov-banner-icon";
    iconSpan.textContent = icon;
    const textSpan = document.createElement("span");
    textSpan.textContent = text;
    div.appendChild(iconSpan);
    div.appendChild(textSpan);
    return div;
  }

  function renderGroups(recipients) {
    groupsEl.innerHTML = "";
    const groups = ValidatorCore.groupByDomain(recipients);
    const isMulti = groups.filter((g) => g.domain !== "unknown").length > 1;

    groups.forEach((group) => {
      const groupEl = document.createElement("div");
      groupEl.className = "ov-group" + (isMulti ? " multi" : "");

      const header = document.createElement("div");
      header.className = "ov-group-header";

      const badge = document.createElement("div");
      badge.className = "ov-domain-badge";
      badge.textContent = group.domain === "unknown" ? "?" : group.domain.slice(0, 2);
      header.appendChild(badge);

      const titleWrap = document.createElement("div");
      titleWrap.className = "ov-group-title";
      const domainName = document.createElement("div");
      domainName.className = "ov-domain-name";
      domainName.textContent = group.domain === "unknown" ? "Unrecognized address" : group.domain;
      const count = document.createElement("div");
      count.className = "ov-domain-count";
      count.textContent = `${group.recipients.length} recipient${group.recipients.length === 1 ? "" : "s"}`;
      titleWrap.appendChild(domainName);
      titleWrap.appendChild(count);
      header.appendChild(titleWrap);

      groupEl.appendChild(header);

      const list = document.createElement("ul");
      list.className = "ov-recipient-list";
      group.recipients.forEach((r) => {
        const li = document.createElement("li");
        li.className = "ov-recipient";

        const avatar = document.createElement("div");
        avatar.className = "ov-recipient-avatar";
        avatar.textContent = initials(r.displayName || r.emailAddress);

        const info = document.createElement("div");
        info.className = "ov-recipient-info";
        const nameEl = document.createElement("div");
        nameEl.className = "ov-recipient-name";
        nameEl.textContent = r.displayName || r.emailAddress;
        const emailEl = document.createElement("div");
        emailEl.className = "ov-recipient-email";
        emailEl.textContent = r.emailAddress;
        info.appendChild(nameEl);
        info.appendChild(emailEl);

        const tag = document.createElement("span");
        tag.className = "ov-field-tag";
        tag.textContent = r.field;

        li.appendChild(avatar);
        li.appendChild(info);
        li.appendChild(tag);
        list.appendChild(li);
      });
      groupEl.appendChild(list);

      groupsEl.appendChild(groupEl);
    });
  }
})();
