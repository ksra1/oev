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
| `manifest.xml` | **Not included** — you're adding this by hand. See below for the exact values it needs to reference. |

## GitHub Pages

Serve this repo from the root of `main` (Settings → Pages → Source: `main` /
`/ (root)`). Once enabled, the add-in will be reachable at:

```
https://ksra1.github.io/oev/
```

All URLs below assume that base.

## What your manifest.xml needs to reference

### Task pane command (ribbon icon on the compose window)

- Icons: `assets/icon-16.png`, `assets/icon-32.png`, `assets/icon-80.png`
- Task pane source location: `https://ksra1.github.io/oev/taskpane.html`
- Requirement set: `Mailbox 1.5` is enough for the task pane itself.

Example `VersionOverrides` (XML manifest, `MessageComposeCommandSurface`):

```xml
<bt:Urls>
  <bt:Url id="Commands.Url" DefaultValue="https://ksra1.github.io/oev/commands.html"/>
  <bt:Url id="Taskpane.Url" DefaultValue="https://ksra1.github.io/oev/taskpane.html"/>
</bt:Urls>
<bt:Images>
  <bt:Image id="Icon.16x16" DefaultValue="https://ksra1.github.io/oev/assets/icon-16.png"/>
  <bt:Image id="Icon.32x32" DefaultValue="https://ksra1.github.io/oev/assets/icon-32.png"/>
  <bt:Image id="Icon.80x80" DefaultValue="https://ksra1.github.io/oev/assets/icon-80.png"/>
</bt:Images>

<Control xsi:type="Button" id="ValidatorButton">
  <Label resid="ValidatorButton.Label"/>
  <Supertip>
    <Title resid="ValidatorButton.Label"/>
    <Description resid="ValidatorButton.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="Icon.16x16"/>
    <bt:Image size="32" resid="Icon.32x32"/>
    <bt:Image size="80" resid="Icon.80x80"/>
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="Taskpane.Url"/>
  </Action>
</Control>
```

### Send-time guard (`ItemSend` LaunchEvent)

- Function file: `https://ksra1.github.io/oev/commands.html`
- Function name to bind: `validateOnSend`
- Requirement set: `Mailbox 1.10` (LaunchEvent / `Office.actions.associate`) — if
  your target Outlook clients only support the older `ExecuteFunction`
  pattern, you may need `Mailbox 1.8`/`1.9` and the older event registration
  instead.

Example (inside `VersionOverrides` for the same command surface, alongside
the `Control` above):

```xml
<Runtimes>
  <Runtime resid="Commands.Url" lifetime="short">
    <Override type="javascript" resid="Commands.Url"/>
  </Runtime>
</Runtimes>

<Hosts>
  <Host xsi:type="MailHost">
    <DesktopFormFactor>
      <FunctionFile resid="Commands.Url"/>
      <ExtensionPoint xsi:type="LaunchEvent">
        <LaunchEvents>
          <LaunchEvent Type="ItemSend" FunctionName="validateOnSend" SendMode="PromptUser"/>
        </LaunchEvents>
        <SourceLocation resid="Commands.Url"/>
      </ExtensionPoint>
    </DesktopFormFactor>
  </Host>
</Hosts>
```

`SendMode="PromptUser"` is important — it's what lets `event.completed({allowEvent:false})`
actually stop the send and lets our own dialog (`dialog.html`) take over the
confirmation instead of a generic Outlook prompt.

If you're using the newer unified JSON manifest instead of XML, the
equivalent is a `"launchEvent"` entry under `runtimes[].actions` pointing at
`validateOnSend` with `"onlyOnPrompt": true`.

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
