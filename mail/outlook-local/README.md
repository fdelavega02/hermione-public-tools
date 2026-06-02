# Outlook Local Automation Starter

This is a local-only starter for Francisco's Outlook Web process:

1. Sign in manually in a real local browser window once.
2. Reuse the locally saved browser session to detect inbox emails.
3. Save inbox snapshots and summaries to local files.
4. Give a general reply idea.
5. Generate a local reply draft packet for review and approval.
6. Do not auto-send anything unless a future process explicitly adds that step and the user approves it.

## Security constraints

- Do not request, paste, or store Outlook passwords in chat, terminal prompts, config files, or scripts.
- Use `npm run auth` for a one-time manual browser login on this machine only.
- Browser storage/session state is saved locally in `state/` only.
- This scaffold defaults to writing local files only.
- There is no send script here. Approval stays manual by default.

## Files

- `package.json` - Node project scripts and dependency declaration
- `config.example.json` - local config template
- `auth.mjs` - manual Outlook Web login that saves local browser state
- `sync-inbox.mjs` - inbox snapshot/summarizer that writes local outputs
- `check-alerts.mjs` - recurring alert checker that syncs first, tracks seen inbox items locally, and prints either `NO_REPLY` or a Discord-ready alert summary
- `reply-draft.mjs` - creates a local reply draft packet from an inbox item plus user guidance
- `scripts/setup-task.ps1` - Windows Task Scheduler helper for an every-2-days morning sync

## Setup

1. Use Node.js 18+.
2. Change into this project:

```bash
cd mail/outlook-local
```

3. Install dependencies:

```bash
npm install
```

4. Copy the example config:

```bash
cp config.example.json config.json
```

On Windows PowerShell:

```powershell
Copy-Item .\config.example.json .\config.json
```

5. Edit `config.json` if your Outlook tenant needs different URLs or selector tweaks.
6. If there are routine messages you never want summarized, add subject/sender filters under `ignoreRules`.

## How to adapt this

- Copy `config.example.json` to `config.json`, then keep the real config local.
- Tune selectors and filter rules on your machine after a manual sync.
- Keep private values, sender names, alert terms, local paths, generated files, and screenshots out of git.
- Run auth manually with `npm run auth`; do not paste passwords or session data into chat or config files.
- Treat the example filters as placeholders. Replace them locally with your own private rules.

## Manual login

Run:

```bash
npm run auth
```

That opens a real Chromium window locally. Sign in there, complete MFA there, wait until the mailbox is visible, then return to the terminal and press Enter to save local session state.

If the saved session expires later, rerun `npm run auth`.

## Inbox sync

Run:

```bash
npm run sync
```

Outputs are written to `output/inbox/`:

- `latest.json` - structured inbox rows
- `latest.md` - readable summary
- `latest.html` - captured DOM for selector debugging
- `latest.png` - screenshot for selector debugging
- timestamped history files for each run
- `emails/<reply-id>.json` - per-item local snapshots

`npm run sync` still extracts message-list data only. Full-body fetching is reserved for the alert checker so the inbox detection path stays stable.

If you want to suppress recurring noise, use local ignore rules in `config.json`, for example:

```json
"ignoreRules": {
  "subjectIncludes": ["Routine notice", "Automated reminder"],
  "fromIncludes": ["no-reply@example.com"],
  "previewIncludes": []
}
```

## Recurring new-email alerts

Run:

```bash
npm run check-alerts
```

This script only runs during the configured local alert window. Outside that window it exits with exactly `NO_REPLY` and skips the inbox sync.

When it does run, it syncs the inbox first, then compares the latest filtered inbox snapshot against local alert state in `state/alert-state.json`.

Alert rules:

- only unread emails are considered
- subject, sender, or preview filters can be configured locally
- recurring low-priority messages can be ignored through local config
- public docs intentionally avoid listing private filter terms, names, organizations, or operational details

Behavior:

- Runs with no alertable unread emails: prints exactly `NO_REPLY`
- Runs with one or more alertable unread emails: reuses the saved authenticated Outlook browser session, matches the synced inbox rows against Outlook's internal inbox API, fetches the normalized message body with Outlook's own `FindItem` and `GetItem` service calls, reads the full email content, and prints a short plain-text summary instead of dumping the raw body into the alert
- If that safer Outlook-internal fetch path is unavailable, it still summarizes the inbox preview into a short human line and does not click the reading pane by default
- Matching unread emails continue to alert on later checks until they are no longer unread
- Repeated identical check failures are suppressed after the first failure alert so a broken login or selector does not spam the same error every few minutes

The detection path honors local ignore rules because it reads from the same synced inbox snapshot that `npm run sync` writes after filtering.

## Reply draft process

After a sync, choose a reply ID from `output/inbox/latest.md` or `output/inbox/latest.json`.

Example:

```bash
npm run draft -- --email-id 1-example-subject --idea "Thank them, say I can review this tomorrow morning, and ask whether they want comments inline or in a meeting."
```

Or use a text file for longer guidance:

```bash
npm run draft -- --email-id 1-example-subject --guidance-file ./my-reply-idea.txt
```

Outputs are written to `output/replies/<reply-id>/`:

- `request.json` - structured local request, with `sendApproved: false`
- `assistant-prompt.md` - ready-to-use context for drafting help
- `starter-draft.txt` - a first-pass reply text
- `approval-checklist.md` - manual review checklist

This keeps the approval boundary explicit. The draft is local and editable. Nothing is sent automatically.

## Likely selector tweaks

Outlook Web DOM differs across Microsoft 365 tenants and UI updates. The default selectors are intentionally broad.

You will likely need to tweak:

- `selectors.postLoginReady`
- `selectors.inboxReady`
- `selectors.messageList`
- `selectors.messageRow`
- `selectors.subject`
- `selectors.from`
- `selectors.preview`
- `selectors.time`
- `selectors.unreadIndicator`

Use `output/inbox/latest.html` and `output/inbox/latest.png` after the first sync to inspect the real structure and tighten selectors.

## Scheduling every 2 days in the morning on Windows

After manual auth and a successful manual sync, you can create a Windows scheduled task:

```powershell
.\scripts\setup-task.ps1
```

Default: every 2 days at `07:00`.

Override the run time:

```powershell
.\scripts\setup-task.ps1 -RunAt "08:30"
```

The task only runs `node sync-inbox.mjs` through `wsl.exe` and appends logs to `output/scheduled-run.log`.

If you want scheduler-driven alerts instead of sync-only runs, point the scheduled task at `node check-alerts.mjs` and let the caller decide what to do with stdout.

## Limits and next extensions

- Outlook may expire the session and require a fresh manual login.
- The safer summary path depends on stored browser session data that contains valid Outlook access tokens and on Outlook continuing to support the internal `FindItem` and `GetItem` calls used by Outlook Web.
- The direct fetch path is designed around matching the synced inbox rows to Outlook Web inbox API results; if Microsoft changes that behavior, summaries will fall back to the inbox preview instead of clicking the message open.
- This starter summarizes individual messages, not full conversation threads.
- It does not compose in Outlook's editor or send mail.
- If you later add send automation, keep it gated behind a local approval flag and a separate explicit action.
