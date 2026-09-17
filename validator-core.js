/*
 * Pure, host-agnostic logic shared by taskpane.js and commands.js.
 * No network calls, no Office.js dependency — safe to unit-test standalone.
 */
(function (global) {
  "use strict";

  function getDomain(emailAddress) {
    if (!emailAddress || typeof emailAddress !== "string") return "unknown";
    const at = emailAddress.lastIndexOf("@");
    if (at === -1 || at === emailAddress.length - 1) return "unknown";
    return emailAddress.slice(at + 1).toLowerCase().trim();
  }

  // recipients: [{ displayName, emailAddress, field: "To"|"Cc"|"Bcc" }]
  // returns: [{ domain, recipients: [...] }] sorted by recipient count desc, domain asc
  function groupByDomain(recipients) {
    const map = new Map();
    for (const r of recipients) {
      const domain = getDomain(r.emailAddress);
      if (!map.has(domain)) map.set(domain, []);
      map.get(domain).push(r);
    }
    return Array.from(map.entries())
      .map(([domain, list]) => ({ domain, recipients: list }))
      .sort((a, b) => b.recipients.length - a.recipients.length || a.domain.localeCompare(b.domain));
  }

  function flattenRecipients(toList, ccList, bccList) {
    const tag = (list, field) => (list || []).map((r) => ({
      displayName: r.displayName || r.emailAddress,
      emailAddress: r.emailAddress,
      field,
    }));
    return [...tag(toList, "To"), ...tag(ccList, "Cc"), ...tag(bccList, "Bcc")];
  }

  function distinctDomainCount(recipients) {
    const domains = new Set(recipients.map((r) => getDomain(r.emailAddress)));
    domains.delete("unknown");
    return domains.size || (recipients.length ? 1 : 0);
  }

  const api = { getDomain, groupByDomain, flattenRecipients, distinctDomainCount };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  } else {
    global.ValidatorCore = api;
  }
})(typeof window !== "undefined" ? window : this);
