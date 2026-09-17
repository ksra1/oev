/* global Office, ValidatorCore */
(function () {
  "use strict";

  let activeDialog = null;

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

  // Runs on the ItemSend event (configured in manifest.xml as a LaunchEvent).
  // Blocks send and asks for confirmation only when recipients span more than one domain.
  function validateOnSend(event) {
    const item = Office.context.mailbox.item;

    Promise.all([getFieldAsync(item, "to"), getFieldAsync(item, "cc"), getFieldAsync(item, "bcc")])
      .then(([to, cc, bcc]) => {
        const recipients = ValidatorCore.flattenRecipients(to, cc, bcc);
        const domainCount = ValidatorCore.distinctDomainCount(recipients);

        if (domainCount <= 1) {
          event.completed({ allowEvent: true });
          return;
        }

        confirmWithUser(recipients, event);
      })
      .catch(() => {
        // Never block send due to an unexpected read error.
        event.completed({ allowEvent: true });
      });
  }

  function confirmWithUser(recipients, event) {
    const groups = ValidatorCore.groupByDomain(recipients).map((g) => ({
      domain: g.domain,
      count: g.recipients.length,
    }));
    const payload = encodeURIComponent(JSON.stringify(groups));
    const dialogUrl = `${getBaseUrl()}dialog.html?data=${payload}`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 40, width: 30, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          // If the dialog can't open, fail safe by blocking the send rather than losing the warning silently.
          event.completed({ allowEvent: false });
          return;
        }
        activeDialog = asyncResult.value;
        activeDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          handleDialogMessage(arg, event);
        });
        activeDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          // Dialog closed without a choice (e.g. user hit the X) -> block the send.
          event.completed({ allowEvent: false });
        });
      }
    );
  }

  function handleDialogMessage(arg, event) {
    let choice;
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

  function getBaseUrl() {
    const scriptEl = document.querySelector('script[src*="commands.js"]');
    const src = scriptEl ? scriptEl.src : window.location.href;
    return src.slice(0, src.lastIndexOf("/") + 1);
  }

  // Exposed for the manifest's <FunctionName> binding.
  window.validateOnSend = validateOnSend;
  if (Office && Office.actions && Office.actions.associate) {
    Office.actions.associate("validateOnSend", validateOnSend);
  }
})();
