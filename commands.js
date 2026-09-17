/* global Office */
(function () {
  "use strict";

  // Hardcoded rather than derived from document.location: this file also runs standalone
  // (no DOM) as the classic-Windows JS runtime override, so it can't rely on `document`.
  var BASE_URL = "https://ksra1.github.io/oev/";

  var activeDialog = null;

  function getDomain(emailAddress) {
    if (!emailAddress || typeof emailAddress !== "string") return "unknown";
    var at = emailAddress.lastIndexOf("@");
    if (at === -1 || at === emailAddress.length - 1) return "unknown";
    return emailAddress.slice(at + 1).toLowerCase().trim();
  }

  function flattenRecipients(toList, ccList, bccList) {
    function tag(list, field) {
      return (list || []).map(function (r) {
        return { displayName: r.displayName || r.emailAddress, emailAddress: r.emailAddress, field: field };
      });
    }
    return tag(toList, "To").concat(tag(ccList, "Cc"), tag(bccList, "Bcc"));
  }

  function groupByDomain(recipients) {
    var map = {};
    var order = [];
    recipients.forEach(function (r) {
      var domain = getDomain(r.emailAddress);
      if (!map[domain]) {
        map[domain] = [];
        order.push(domain);
      }
      map[domain].push(r);
    });
    return order
      .map(function (domain) {
        return { domain: domain, recipients: map[domain] };
      })
      .sort(function (a, b) {
        return b.recipients.length - a.recipients.length || a.domain.localeCompare(b.domain);
      });
  }

  function distinctDomainCount(recipients) {
    var domains = {};
    var count = 0;
    recipients.forEach(function (r) {
      var d = getDomain(r.emailAddress);
      if (d !== "unknown" && !domains[d]) {
        domains[d] = true;
        count++;
      }
    });
    if (count === 0 && recipients.length) return 1;
    return count;
  }

  function getFieldAsync(item, field) {
    return new Promise(function (resolve) {
      if (!item[field] || !item[field].getAsync) {
        resolve([]);
        return;
      }
      item[field].getAsync(function (result) {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
      });
    });
  }

  // Runs on the OnMessageSend Smart Alerts event (configured in manifest.xml as a LaunchEvent).
  // Blocks send and asks for confirmation only when recipients span more than one domain.
  function validateOnSend(event) {
    var item = Office.context.mailbox.item;

    Promise.all([getFieldAsync(item, "to"), getFieldAsync(item, "cc"), getFieldAsync(item, "bcc")])
      .then(function (results) {
        var recipients = flattenRecipients(results[0], results[1], results[2]);
        var domainCount = distinctDomainCount(recipients);

        if (domainCount <= 1) {
          event.completed({ allowEvent: true });
          return;
        }

        confirmWithUser(recipients, event);
      })
      .catch(function () {
        // Never block send due to an unexpected read error.
        event.completed({ allowEvent: true });
      });
  }

  function confirmWithUser(recipients, event) {
    var groups = groupByDomain(recipients).map(function (g) {
      return { domain: g.domain, count: g.recipients.length };
    });
    var payload = encodeURIComponent(JSON.stringify(groups));
    var dialogUrl = BASE_URL + "dialog.html?data=" + payload;

    Office.context.ui.displayDialogAsync(dialogUrl, { height: 40, width: 30, displayInIframe: false }, function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        // If the dialog can't open, fail safe by blocking the send rather than losing the warning silently.
        event.completed({ allowEvent: false });
        return;
      }
      activeDialog = asyncResult.value;
      activeDialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        handleDialogMessage(arg, event);
      });
      activeDialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        // Dialog closed without a choice (e.g. user hit the X) -> block the send.
        event.completed({ allowEvent: false });
      });
    });
  }

  function handleDialogMessage(arg, event) {
    var choice;
    try {
      choice = JSON.parse(arg.message).action;
    } catch (e) {
      choice = "cancel";
    }
    if (activeDialog) {
      activeDialog.close();
      activeDialog = null;
    }
    event.completed({ allowEvent: choice === "send" });
  }

  // Maps the manifest's FunctionName="validateOnSend" to this handler.
  Office.actions.associate("validateOnSend", validateOnSend);
})();
