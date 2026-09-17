# Outlook Email Validator

A client-side Outlook add-in that groups the recipients of the message you're
composing by email domain, so a compose window mixing addresses from
`companyx.com` and `companya.com` is obvious before you hit Send.

**Privacy:** everything runs in the browser inside Outlook. No recipient
address, subject, or body is ever sent to any server. This repo only hosts
static HTML/CSS/JS on GitHub Pages — Outlook loads those files directly into
the add-in surfaces described below.

## What's in this repo

| File | Purpose |
|---|---|
| `taskpane.html` / `taskpane.js` / `taskpane.css` | Docked panel opened from the ribbon icon. Shows recipients grouped by domain, live-updates as you edit To/Cc/Bcc. |
| `commands.html` / `commands.js` | UI-less function file. Runs on the `ItemSend` event and blocks sending until you confirm if recipients span more than one domain. |
| `dialog.html` / `dialog.js` / `dialog.css` | The "multiple companies detected" confirmation popup shown by `commands.js` before send. |
| `validator-core.js` | Pure domain-grouping logic shared by the task pane and the send guard. |
| `assets/icon-*.png` | Ribbon icons (16/20/24/32/40/48/64/80 px), generated from a blue circle + envelope + checkmark. |
| `index.html` | Informational landing page only — not used by Outlook. |
| `manifest.xml` | The add-in manifest, pre-filled with this repo's GitHub Pages URLs. |

## GitHub Pages

Serve this repo from the root of `main` (Settings → Pages → Source: `main` /
`/ (root)`). Once enabled, the add-in will be reachable at:

```
https://ksra1.github.io/oev/
```

All URLs below assume that base.

## manifest.xml

Already in this repo, pointing at `https://ksra1.github.io/oev/...` for every
URL. Two things to change before real use:

- `<ProviderName>` currently says `Sravan` — update if you want something else.
- `<Id>` is a random GUID (`ca549072-6cf4-4892-8e6a-9e1932b438b0`) generated
  for this project. Keep it as-is unless you're forking this into a second,
  separate add-in — Outlook uses this ID to tell add-ins apart.

It wires up two extension points:

1. **Task pane command** (`MessageComposeCommandSurface`, requirement set
   `Mailbox 1.5`) — the ribbon button that opens `taskpane.html`.
2. **Send-time guard** (`LaunchEvent` of `Type="ItemSend"`, requirement set
   `Mailbox 1.10`) — binds to `validateOnSend` in `commands.js`, with
   `SendMode="PromptUser"` so `event.completed({allowEvent:false})` actually
   stops the send and our own dialog (`dialog.html`) can take over instead of
   a generic Outlook prompt.

**Known validator false-positive:** running
`npx office-addin-manifest validate manifest.xml` will pass everything except
one line: `Autorun LaunchEvent Type is not valid`. That check calls a
Microsoft-hosted acceptance-test service, not local XML schema validation
(all real schema errors are already fixed — the file validates cleanly
against the XSD). The `ItemSend` LaunchEvent is a newer, less common
extension point, and this hosted service has a known gap in recognizing it;
the manifest structure matches Microsoft's own smart-alerts sample. Trust
sideloading in real Outlook over this one check.

## Sideloading to test

1. Confirm the site is live: `https://ksra1.github.io/oev/taskpane.html`
   should load.
2. In Outlook on the web or new Outlook for Windows: **Get Add-ins** → **My
   add-ins** → **Add a custom add-in** → **Add from file** → pick
   `manifest.xml`.
3. In classic Outlook for Windows/Mac, sideload via a network shared folder
   or the Microsoft 365 admin center, per Microsoft's [sideload docs](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
4. Open a new message, click the **Check Recipients** button in the ribbon,
   and add recipients from two different domains to confirm both the task
   pane grouping and the send-time confirmation dialog show up.

## How the logic works

1. `validator-core.js` exposes `groupByDomain(recipients)`, which buckets
   `{ displayName, emailAddress, field }` objects by the part of the email
   after `@`.
2. `taskpane.js` reads `Office.context.mailbox.item.to/cc/bcc` via
   `getAsync`, re-renders on every `RecipientsChanged` event, and shows one
   card per domain with a red banner if more than one domain is present.
3. `commands.js` runs the same grouping on `ItemSend`. If everything shares
   one domain it allows the send silently. If not, it opens `dialog.html` in
   a modal with a per-domain recipient count and two choices: **Go back &
   review** (blocks the send) or **Send anyway** (allows it). The dialog
   only ever receives domain names and counts, passed via a URL parameter —
   never full addresses or message content.

## Local testing without Outlook

Because there's no build step, you can preview the task pane's rendering
logic (minus real Office.js data) by opening `taskpane.html` directly, or by
sideloading against a real message per Microsoft's [sideload docs](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing),
pointing the manifest's URLs at your GitHub Pages deployment (or `http://localhost`
during development, if your manifest is set up for it).
