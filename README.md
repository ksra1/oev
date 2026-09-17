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
| `commands.html` / `commands.js` | Event-based runtime for the send guard. Runs on the `OnMessageSend` Smart Alerts event and blocks sending until you confirm if recipients span more than one domain. `commands.js` is self-contained (no `ValidatorCore` dependency) because it's also loaded directly as raw JS by classic Outlook for Windows, which has no DOM to load a second `<script>` from. |
| `dialog.html` / `dialog.js` / `dialog.css` | The "multiple companies detected" confirmation popup shown by `commands.js` before send. |
| `validator-core.js` | Pure domain-grouping logic used by the task pane (`taskpane.js`). `commands.js` duplicates the same logic inline rather than depending on this file — see above. |
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

It wires up two extension points, both inside a single nested
`VersionOverrides` `V1_1` block (required because `Runtimes`/`LaunchEvent`
don't exist in `V1_0`):

1. **Task pane command** (`MessageComposeCommandSurface`) — the ribbon
   button that opens `taskpane.html`.
2. **Send-time guard** — a `LaunchEvent` of `Type="OnMessageSend"` (the
   current "Smart Alerts" feature, requirement set `Mailbox 1.12`+). It binds
   to `validateOnSend` in `commands.js` via `Office.actions.associate`, with
   `SendMode="PromptUser"` so `event.completed({allowEvent:false})` actually
   stops the send and our own dialog (`dialog.html`) takes over instead of a
   generic Outlook prompt.

   Note this is deliberately **not** `Type="ItemSend"`. `ItemSend` is an
   older, separate on-send mechanism (`Mailbox 1.8`, registered via a plain
   `<Event>` element, `Block` mode only, add-in-only manifest) that isn't
   valid inside a `LaunchEvent`/`VersionOverridesV1_1` block — using it there
   is what caused Outlook's "the value listed for the Type attribute...is
   invalid" error. `OnMessageSend` is the newer, more flexible replacement
   and the value `office-addin-manifest validate` actually accepts.

   Two runtimes are registered so the guard works everywhere: `commands.html`
   (loads `commands.js` via a `<script>` tag) for the web/Mac/new-Outlook
   browser runtime, and `commands.js` referenced directly as raw JS for the
   classic-Windows JS-only runtime.

Validated with `npx office-addin-manifest validate manifest.xml` — passes
cleanly, including the `OnMessageSend` LaunchEvent type check.

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
3. `commands.js` runs the same grouping on `OnMessageSend`. If everything shares
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
