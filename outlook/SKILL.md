---
name: outlook
description: Use when the user asks to read, search, send, delete, archive Outlook emails, or read Outlook calendar events and Outlook todos/tasks. Also use for requests similar to /mail, /cal, agenda Outlook, emails Outlook, or tâches Outlook.
---

# Outlook

Use this skill when the user wants to work with Microsoft Outlook from opencode, especially for:

- reading inbox emails;
- searching emails;
- reading a specific email;
- archiving emails;
- deleting emails;
- preparing or sending emails;
- reading Outlook calendar events;
- creating Outlook calendar events;
- deleting Outlook calendar events;
- reading Outlook todos/tasks.

## Core Rules

- Use the scripts bundled in this skill's directory (`.opencode\skills\outlook\`).
- Do not run multiple Outlook COM operations in parallel. Outlook often rejects concurrent COM calls.
- Never send an email unless the user explicitly asks to send and the recipient, subject, and body are clear.
- If the user asks to answer an email but the content is not final, create a draft instead of sending.
- Never permanently delete emails. The delete action only moves messages to Outlook deleted items.
- Before deleting or sending, summarize the target and proceed only when the user intent is explicit.
- For bulk destructive actions, ask one short confirmation question unless the user already provided a precise command and scope.
- Prefer JSON output from scripts, then provide a concise French summary to the user.

## Mail Script

Primary script:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
```

Supported actions:

- `list`: list inbox emails.
- `read`: read a single email by Outlook `EntryID`.
- `search`: search emails by text query.
- `resolve`: search the Outlook Global Address List (GAL) by name or email.
- `archive`: move one email to Archive.
- `delete`: move one email to Deleted Items.
- `draft-new`: create and display a new draft email without sending.
- `draft-reply`: create and display a reply draft.
- `send-new`: send a new email.
- `mark-read`: mark one email as read.
- `count-older-than`: count inbox emails older than N days.
- `list-older-than`: list inbox emails older than N days.

### List Inbox Emails

Use this for `lis mes mails`, `emails récents`, `boîte de réception`, or requests similar to `/mail`.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action list -Limit 20
```

If the user asks for `/mail`-style processing, follow the existing `.opencode\commands\mail.md` behavior: classify as `information`, `meeting`, `todo`, `reply`, or `ignore`; save useful information in `raw/mails/`; add long actions to `todo.md`; archive only `meeting`, `information`, and `ignore`.

### Read One Email

Use this after listing/searching when an `EntryID` is available.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action read -EntryId '<EntryID>'
```

### Search Emails

Use this for `cherche dans mes mails`, `retrouve le mail`, `search Outlook`, or when the user gives a subject, sender, keyword, or date clue.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action search -Query '<texte>' -Limit 20
```

Optional folder scope:

```powershell
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action search -Query '<texte>' -Folder sent -Limit 20
```

Supported `-Folder` values: `inbox`, `sent`, `archive`, `deleted`, `all`.

### Resolve Name / Search Address Book

Use this when the user gives a person's name or email and asks to find their Outlook contact, email address, or phone number.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action resolve -Query 'Nom ou email'
```

Returns matching contacts with name, email (SMTP), department, and job title from the Global Address List.

### Archive Email

Archive only the intended email by `EntryID`.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action archive -EntryId '<EntryID>'
```

### Delete Email

Deletion moves the email to Outlook deleted items. It is not permanent deletion.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action delete -EntryId '<EntryID>'
```

Ask for confirmation before bulk deletion unless the user already provided an explicit, narrow scope.

### Draft New Email

Use this when the user asks to write an email without sending it, or when the content or recipients are not fully confirmed.

Write the body to a temporary UTF-8 text file first, then call `draft-new`. The draft is saved and displayed in Outlook for review before sending.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
$bodyPath = Join-Path (Get-Location) '.opencode\temp\draft-body.txt'
Set-Content -LiteralPath $bodyPath -Encoding UTF8 -Value "Corps du mail"
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action draft-new -To 'destinataire@example.com' -Subject 'Sujet' -BodyPath $bodyPath
```

Optional `-Cc` and `-Bcc` parameters are supported.

Use `resolve` first if the recipient's email address is not known:

```powershell
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action resolve -Query 'Prénom Nom'
```

### Draft Reply

Write the reply body to a temporary UTF-8 text file, then create a displayed Outlook draft.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action draft-reply -EntryId '<EntryID>' -BodyPath '<reply-body.txt>'
```

### Send New Email

Use only when the user explicitly asks to send and all fields are clear.

Write the body to a temporary UTF-8 text file first.

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action send-new -To 'person@example.com' -Cc '' -Subject 'Sujet' -BodyPath '<body.txt>'
```

If the user asks for a draft instead of immediate sending, do not use `send-new`; use Outlook manually only if a dedicated draft action is added later.

## Calendar Script

Primary script:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-cal.ps1'
```

Supported actions:

- `list` (default): list calendar events for a date range.
- `create`: create a new calendar event or meeting.

### List Calendar Events

Read Outlook agenda events with a date range:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-cal.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -StartDate '2026-06-01' -EndDate '2026-06-30'
```

Add `-IncludeEntryId` to retrieve the `entryId` of each event (required for deletion):

```powershell
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -StartDate '2026-07-03' -EndDate '2026-07-03' -IncludeEntryId
```

If the user asks for `/cal` or full calendar export, follow `.opencode\commands\cal.md` and save idempotent Markdown files under `raw/calendrier/`.

### Delete a Calendar Event

Use the mail script with action `delete` and the event `entryId` obtained via `-IncludeEntryId`:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action delete -EntryId '<EntryID>'
```

Deletion moves the event to Outlook deleted items. It is not permanent.

### Create a Calendar Event

Use this when the user asks to add, create, or schedule a calendar event or meeting.

Required parameters: `-Subject`, `-Start`, `-End` (ISO datetime, e.g. `'2026-07-06 09:00'`).

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-cal.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action create `
  -Subject 'Titre de la réunion' `
  -Start '2026-07-06 09:00' `
  -End   '2026-07-06 09:30'
```

Optional parameters:

| Parameter   | Description                                          |
|-------------|------------------------------------------------------|
| `-Location` | Room or Teams link                                   |
| `-Body`     | Description / agenda                                 |
| `-Required` | Comma-separated required attendee emails             |
| `-Optional` | Comma-separated optional attendee emails             |

Example with attendees:

```powershell
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action create `
  -Subject  'Point hebdo' `
  -Start    '2026-07-06 09:00' `
  -End      '2026-07-06 09:30' `
  -Required 'alice@example.com,bob@example.com' `
  -Optional 'charlie@example.com'
```

If the attendee email is unknown, resolve it first with the mail script:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$mailScript = Join-Path $skillDir 'outlook-mail.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $mailScript -Action resolve -Query 'Prénom Nom'
```

Returns JSON with `status`, `entryId`, `subject`, `start`, `end`, `location`, and `isMeeting`.

## Todo/Tasks Script

Primary script:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-todo.ps1'
```

List active Outlook tasks:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-todo.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action list -Status active -Limit 50
```

Search Outlook tasks:

```powershell
$skillDir = 'C:\Users\jfhelie\OneDrive - Sopra Steria\Documents\OpenCodeDesktop\.opencode\skills\outlook'
$script = Join-Path $skillDir 'outlook-todo.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $script -Action search -Query '<texte>' -Limit 50
```

Supported `-Status` values: `all`, `active`, `completed`, `overdue`.

## Response Format

For read/search/list requests, answer in French with:

- number of items read;
- most relevant items first;
- subject/title, date, sender/organizer, and short summary;
- available next actions when useful: read, archive, delete, draft reply, send, mark read.

For sensitive actions, include:

- exact target count;
- action performed;
- errors, if any.

After creating or changing this skill or scripts, the user must quit and restart opencode for skill loading changes to take effect.
