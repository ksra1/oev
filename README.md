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
| `taskpane.html` / `taskpane.js` / `taskpane.css` | Docked panel opened from the ribbon icon. Shows recipients grouped by domain, live-updates as you edit To/Cc/Bcc. **This is the only surface currently wired into `manifest.xml`** — see "Send-time guard" below for why. |
| `commands.html` / `commands.js` | Event-based runtime for a send-time guard (`OnMessageSend` Smart Alerts) that would block sending until you confirm if recipients span more than one domain. Present in the repo but **not referenced by `manifest.xml`** on this account — see below. |
| `dialog.html` / `dialog.js` / `dialog.css` | The "multiple companies detected" confirmation popup `commands.js` would show before send. Same caveat as above. |
| `validator-core.js` | Pure domain-grouping logic used by the task pane (`taskpane.js`). `commands.js` duplicates the same logic inline rather than depending on this file, since it also has to run with no DOM (see its own comments). |
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

It currently wires up one extension point:

- **Task pane command** (`MessageComposeCommandSurface`, `Mailbox 1.5`) — the
  ribbon button that opens `taskpane.html`. Fully working.

Validated with `npx office-addin-manifest validate manifest.xml` — passes
cleanly.

### Send-time guard: built, but not wired in (Outlook.com limitation)

The repo also contains a working send-time guard — `commands.html` /
`commands.js` / `dialog.*` — meant to run as a `LaunchEvent` of
`Type="OnMessageSend"` (Microsoft's current "Smart Alerts" feature,
`Mailbox 1.12`+) inside a nested `VersionOverridesV1_1` block, binding to
`validateOnSend` via `Office.actions.associate` with `SendMode="PromptUser"`.

It is **not** referenced by `manifest.xml` right now, because on a personal
Outlook.com/`outlook.live.com` mailbox (as opposed to an Exchange
Online-backed Microsoft 365 work/school mailbox), simply having that nested
`VersionOverridesV1_1` `LaunchEvent` block in the manifest causes Outlook to
silently fail to register the **entire** add-in — not just the guard, the
ribbon button disappears too. This was confirmed with four separate isolated
test manifests sideloaded side by side: a bare ribbon-only add-in (worked), a
full copy of this manifest with the guard removed (worked), and two variants
with the guard included — one matching this repo's original resource-id
naming and one with deliberately unique resource ids to rule out an id
collision (both failed identically, icon and all). Microsoft's own support
table for `OnMessageSend`/Smart Alerts lists only Exchange Online, Exchange
Server SE, and on-premises Exchange as supported backends — personal
Outlook.com/Hotmail mailboxes run on separate consumer infrastructure and
aren't covered by that table, which matches the observed behavior.

If you ever install this add-in against an Exchange Online (Microsoft 365
work/school) mailbox, the guard can likely be restored. Take the nested
`VersionOverridesV1_1` block from this file's git history (see the commit
that added `Fix send-guard LaunchEvent: OnMessageSend, not ItemSend`, or the
one just before the guard was removed) and re-add it as a sibling of the
`<Resources>` block inside the outer `VersionOverridesV1_0`, then test the
same way: sideload, open a new message, and confirm the ribbon button still
appears before trusting the guard itself.

## Sideloading to test

1. Confirm the site is live: `https://ksra1.github.io/oev/taskpane.html`
   should load.
2. In Outlook on the web or new Outlook for Windows: **Get Add-ins** → **My
   add-ins** → **Add a custom add-in** → **Add from file** → pick
   `manifest.xml`.
3. In classic Outlook for Windows/Mac, sideload via a network shared folder
   or the Microsoft 365 admin center, per Microsoft's [sideload docs](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
4. Open a new message, click the **Check Recipients** button in the ribbon,
   and add recipients from two different domains to confirm the task pane
   groups them correctly. (The send-time confirmation dialog isn't currently
   wired up — see "Send-time guard" above.)

## How the logic works

1. `validator-core.js` exposes `groupByDomain(recipients)`, which buckets
   `{ displayName, emailAddress, field }` objects by the part of the email
   after `@`.
2. `taskpane.js` reads `Office.context.mailbox.item.to/cc/bcc` via
   `getAsync`, re-renders on every `RecipientsChanged` event, and shows one
   card per domain with a red banner if more than one domain is present.
3. `commands.js` (not currently wired into the manifest — see above) would
   run the same grouping on `OnMessageSend`. If everything shares one domain
   it allows the send silently. If not, it opens `dialog.html` in a modal
   with a per-domain recipient count and two choices: **Go back & review**
   (blocks the send) or **Send anyway** (allows it). The dialog only ever
   receives domain names and counts, passed via a URL parameter — never full
   addresses or message content.

## Local testing without Outlook

Because there's no build step, you can preview the task pane's rendering
logic (minus real Office.js data) by opening `taskpane.html` directly, or by
sideloading against a real message per Microsoft's [sideload docs](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing),
pointing the manifest's URLs at your GitHub Pages deployment (or `http://localhost`
during development, if your manifest is set up for it).
