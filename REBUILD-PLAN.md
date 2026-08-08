# HelpfulForms — Clean Rebuild on a New Google Account

Working plan. The old account is frozen as a read-only archive and drained
manually; the new account is built clean rather than transferred.

Status: **planning**. Nothing here is built yet.

---

## Decisions taken

| Decision | Choice |
|---|---|
| Migration | Start clean. Old account frozen as archive, worked down in parallel |
| Status model | Three states per section, one date column per state (see below) |
| UI hosting | Apps Script web app, not GitHub Pages |
| OCR intake | Designed into the schema now, built after migration |
| Siri intake | Designed into the architecture now, built after migration |

### Why clean rather than transfer

Sheets store Drive **URLs**. Transferring preserves file IDs, copying does not —
so a copy breaks every historical receipt link. Starting clean sidesteps the
question: legacy files never move, so legacy links never break. The old page
keeps working against the old account until its backlog is empty.

### Why the UI moves into Apps Script

Removes the two client-side secrets entirely (`SHEETS_API_KEY`,
`DELETE_API_KEY`), gets Google sign-in for free, removes CORS, and collapses two
deploy paths into one. Costs ~1–2s of first paint. For a personal tool, worth it.

---

## Principles

1. **One code path per concern.** Today there are four near-identical
   `loadXSheet` / `renderXTable` / `toggleXStatus` sets client-side and four
   toggle functions server-side. They differ only by configuration, and that
   duplication is why the four sections drifted apart. One generic
   implementation driven by a per-section config object.
2. **Columns resolved by header name, never by index.** The current code is full
   of magic numbers (`rowValues[9]`, `getRange(row, 13)`), and every layout
   difference between sheets became a special case. Look columns up by header
   text once per load; adding or reordering a column then breaks nothing.
3. **Entry creation is a function, not a trigger side effect.** `createEntry()`
   is the single way a row is born. Form submit, Siri and OCR are adapters.
4. **Never report success for work that failed.** Today a toggle returns
   `success: true` even when the file rename and email both failed. Every
   operation returns what actually happened, and the UI shows it.
5. **The client is never trusted.** The server reads status, file URLs and
   identifiers from the sheet, never from the request. (Already true after the
   layer-1 hardening; carry it forward.)
6. **Least privilege per entry point.** The signed-in UI can do everything. The
   Siri endpoint can only create entries.

---

## Schema

Every sheet gets the same spine. Section-specific fields sit between.

### Common columns (identical name and meaning in all four sections)

| Header | Type | Notes |
|---|---|---|
| `Timestamp` | datetime | When the entry was created |
| `Source` | text | `form` / `siri` / `ocr` / `manual` — how it arrived |
| `Status` | text | Current state. Closed vocabulary, see below |
| `Receipt URL` | url | Blank if awaiting a receipt |
| `Receipt State` | text | `attached` / `awaiting` / `none required` |
| `Notes` | text | Free text, never parsed |

`Status` holds only the current state name — never a date, never free text. That
is what fixes the filter dropdowns: three stable options per section instead of
one per claim date.

`Receipt State` is what makes Siri and scan-later work: a row can exist before
its receipt does.

### Status states and their dates

Each state gets **its own date column**. A single shared `Status Date` would
overwrite itself on each transition and lose the history — and for Income all
three dates are distinct business facts worth keeping.

**Work, IVA, Health**

| State | Date column | How the date is set |
|---|---|---|
| `To Do` | — | Creation is `Timestamp` |
| `Claimed` | `Claimed Date` | **Automatic** — the day the state was set |
| `Settled` | `Settled Date` | **Manual** — entered by hand |

**Income**

| State | Date column | How the date is set |
|---|---|---|
| `Invoiced` | `Invoiced Date` | **Manual** |
| `Received` | `Received Date` | **Manual** |
| `Logged` | `Logged Date` | **Automatic** — the day the state was set |

Config per section is therefore a list of states, each flagged `auto` or
`manual`. One generic implementation covers both shapes; adding a fourth state
later is a config edit.

### Status control replaces Done/Undo

Three states cannot be a two-way toggle, so each row gets a **status selector**
rather than Done/Undo buttons. Consequences:

- **"Undo" stops existing as a concept.** Going back is just selecting an
  earlier state, which removes the four inconsistent undo implementations.
- Selecting a state whose date is `manual` **prompts for the date** and will not
  commit until one is given.
- Selecting a state whose date is `auto` fills the date server-side.
- Correcting a mistake is the same action as making the change — no special path.

### Section-specific fields

Carry forward what exists today, confirm exact names during build:

- **Work** — Expense Reason, Expense Date, Amount, Currency, Description
- **IVA** — Número, Data, Emitente NIF, Tipo, Importados, Valor do IVA, Valor Total
- **Health** — Patient, Provider, Treatment Date, Invoice Date, Amount, Details URL
- **Income** — confirm current fields before building

Health's columns M/N (`Original Receipt Filename` / `Original Details Filename`)
are **dropped** — they existed only for the iCloud Shortcut, which is cancelled.

---

## Architecture

### Server (Apps Script)

```
Config
  SECTIONS = { work: {...}, iva: {...}, health: {...}, income: {...} }
    sheet name, section fields, filename pattern, what Done does

Core
  createEntry(section, fields, source)   one way a row is born
  setStatus(section, row, done)          one way status changes
  archiveReason(section, reason)         rows to archive sheet, files to Archived
  resolveColumns(sheet)                  header name -> index, cached per load

Adapters
  onFormSubmit(e)      -> createEntry(...)
  doPost (UI)          -> setStatus / archive / list        (signed in)
  doPost (Siri)        -> createEntry only                  (device key)
  ocrIntake(file)      -> createEntry(...)                  (later)
```

`setStatus` replaces four divergent toggle functions. What each transition does
becomes config, not code:

| | Rename receipt | Email |
|---|---|---|
| Work | prefix on `Claimed` | no |
| IVA | prefix on `Claimed` | **on entry creation**, not on status change |
| Health | prefix on `Claimed`, both files | no |
| Income | n/a | no |

**The IVA claim email moves to `createEntry`** — it fires when the receipt is
uploaded rather than when the status changes. That decouples it from status
entirely, so no transition has a side effect beyond renaming, and re-selecting a
state can never re-send mail. Recipient becomes a Script Property rather than
the hardcoded `jacqueline.eaton@nato.int`.

Moving back to an earlier state reverses the rename. Because that is now the
same code path as moving forward, it cannot drift the way four hand-written
undos did.

### Client

One HTML page served by the Apps Script web app, using `google.script.run`
instead of `fetch` — no API key, no CORS, caller identity known server-side.

One render function driven by section config. Explicit loading, empty and error
states, so a failure stops looking identical to "no data".

### Security

| Entry point | Deployment | Auth | Can do |
|---|---|---|---|
| Web UI | Execute as me, **restricted to my account** | Google sign-in | Everything |
| Siri | Execute as me, anyone with key | Key held on device only | `createEntry` only |

No secret ever reaches a public file. The Siri key lives in the Shortcut and in
Script Properties, nowhere else.

---

## Build order

Each step should leave the system working.

1. **New account + storage confirmed.** Verify quota before anything else.
2. **Spreadsheet** with the four sheets and the standard header spine.
3. **Apps Script project**, bound to the sheet. `resolveColumns`, `SECTIONS`
   config, `createEntry`, `setStatus`.
4. **Four Forms**, fields matching the schema, linked to the sheets.
5. **`installFormTrigger()`**, then submit one test entry per section.
6. **Web UI** in the project. Deploy restricted to your account. Confirm
   sign-in, listing, toggling, archiving.
7. **Script Properties** via `setupScriptProperties()`, verify with
   `checkScriptProperties()`.
8. **Cutover** — see below.
9. **Siri Shortcut** + second narrow deployment.
10. **OCR intake.**

Steps 1–8 restore what you have today, cleanly. 9 and 10 are new capability and
can wait.

## Cutover

1. New system verified with test entries; delete the test rows.
2. Turn **off** "Accepting responses" on all four old forms — prevents
   split-brain where some claims land in each account.
3. New submissions go to the new account from here.
4. Old GitHub Pages page stays live and untouched, pointed at the old account,
   used only to work down the remaining backlog.
5. When the old backlog reaches zero: unpublish the old page, delete the old web
   app deployment, keep the old sheet and Drive as archive.

## What stays frozen

Once cutover happens, **stop changing the old system**. Every change there costs
a push → version → manual UI deploy on the thing you depend on to finish the
backlog. Its known issues are accepted, not fixed:

- Silent load failures, toggles reporting false success
- Filter dropdowns growing one option per claim date
- Health M/N drift

**To stop the iCloud emails without touching code:** delete the `ICLOUD_EMAIL`
Script Property on the old account. `getIcloudEmail()` then throws, the caller
catches and logs, no email is sent, and Health Done still flips status and
renames Drive files. Then delete the iOS automation so the failure notifications
stop.

---

## Open questions

- **Does renaming happen at `Settled` / `Logged` too**, or only at `Claimed`?
  Currently only one transition touches filenames.
- **Is `Invoiced` the state every Income row starts in?** If so its date is
  manual, so the form must ask for it at submission.
- **Should a manual date be allowed to be blank**, i.e. can you mark something
  Settled without knowing the date yet, or is the date required to advance?
- Exact current field names for Income — read from the sheet before building
- IVA email recipient address for the new account
- Whether the four sections should stay four Forms or become one with branching
- Where OCR runs: Drive's built-in OCR vs an API, and how confident it needs to
  be before autofilling rather than prompting
