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
| Intake | **Custom form. No Google Forms** (see below) |
| Supplier registry | Self-populating, learned from entries as they are made |
| OCR intake | Designed into the schema now, built after migration |
| Siri intake | Prompted questions + visual confirmation. Built after migration |

### Why clean rather than transfer

Sheets store Drive **URLs**. Transferring preserves file IDs, copying does not —
so a copy breaks every historical receipt link. Starting clean sidesteps the
question: legacy files never move, so legacy links never break. The old page
keeps working against the old account until its backlog is empty.

### Why there is no Google Form

The deciding argument is the supplier registry. **Google Forms cannot fill one
answer from another** — "type FNAC, get its NIF" is impossible there and trivial
in a form we control. Everything else follows from having made that choice.

Dropping Forms *deletes* work rather than adding it:

- Form dropdown syncing for expense reasons and patients — a real chunk of v1's
  complexity, and of what the management module would have been
- Four forms to create, own, link and migrate to the new account
- The `onFormSubmit` trigger and the "Forms writes the row, we finalise it" split
- `entry.NNN` field IDs, which would have made a Siri→Forms POST brittle

What it costs: file upload has to be built (`<input type="file">` → base64 →
`google.script.run` → Drive, receipts are small), and validation becomes ours —
but `missingFields()` already exists.

Since the UI was already moving into Apps Script, the form is just another view
in the same app, and Siri stops being a special case: form and Siri are two
callers of `createEntry`.

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
| `Date` | date | Date of the transaction itself |
| `Amount` | number | |
| `Currency` | text | |
| `Counterparty` | text | The other party — see labels below |
| `Status` | text | Current state. Closed vocabulary, see below |
| `Receipt URL` | url | Blank if awaiting a receipt |
| `Receipt State` | text | `attached` / `awaiting` / `none required` |
| `Notes` | text | Free text, never parsed. Absorbs v1's Description |

This is a deliberate cut from what v1 collected. Date, Amount, Currency and
Counterparty are shared by all four sections, which is why the generic renderer
can handle nearly every column without knowing which section it is showing.

### Counterparty and category

`Counterparty` is one column with a **per-section display label**, so the sheet
and code stay generic while the UI uses the natural word:

| Section | Label | Example |
|---|---|---|
| Work | Supplier | Uber |
| IVA | Retailer | FNAC |
| Health | Provider | Hospital da Luz |
| Income | Paid by | *the client* |

`category` is an optional extra classifying field, present only where it means
something. Its allowed values are a **managed list** that populates the form
dropdown — the generalisation of v1's add/delete expense reason, which gives
Health add/remove patients for free.

| Section | Category column | Managed list | Required |
|---|---|---|---|
| Work | `Expense Reason` | yes | yes |
| Health | `Patient` | yes | yes |
| Income | `Reason` | no — free text, prefilled | no |
| IVA | — | — | — |

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

| Order | State | Date column |
|---|---|---|
| 1 | `To Do` | — (creation is `Timestamp`) |
| 2 | `Claimed` | `Claimed Date` |
| 3 | `Settled` | `Settled Date` |

**Income**

| Order | State | Date column |
|---|---|---|
| 1 | `Invoiced` | `Invoiced Date` |
| 2 | `Received` | `Received Date` |
| 3 | `Logged` | `Logged Date` |

### Date rules

Every date behaves the same way — there is no auto/manual split:

- Selecting a state **prompts for its date, pre-filled with today**. One tap to
  accept, but it is always seen before it is written.
- Pre-filling rather than writing silently matters because `Invoiced`,
  `Received` and `Settled` are usually **backdated**. A silent "today" would be
  wrong most of the time and never noticed.
- Any date can be **edited later** without changing state.

**The date dialog.** Selecting a state opens a small dialog:

```
        Mark as Settled

  ┌──────────────────────────────┐
  │        Today — 8 Aug         │   ← primary, one tap, done
  └──────────────────────────────┘

  or pick a date

  ┌──────────────┐  ┌────────────┐
  │  08/08/2026  │  │     OK     │
  └──────────────┘  └────────────┘

           Cancel
```

The date field is a native `<input type="date">`, so on iOS it opens the system
date wheel rather than anything custom.

When the target state **already has a date** — which is the case when reverting
— the dialog pre-fills with that existing date and the primary button reads
`Keep 15 Jan` rather than `Today`. That keeps the dialog honest about the
"only fill if blank" rule, so reverting never silently re-stamps.

### Drive layout and filenames

Files follow the status. Each state has a folder, and reaching a state appends
to the filename, so the name carries its own audit trail.

```
<root>/<Section>/Inbox       form uploads land here (To Do)
<root>/<Section>/Claimed
<root>/<Section>/Settled
<root>/<Section>/Archived    archived and soft-deleted
```

Base name: `YYMMDD_Counterparty_Amount_<document>.ext`

```
on upload   250115_HospitalDaLuz_3-45_receipt.pdf
Claimed     250115_HospitalDaLuz_3-45_receipt_Claimed_04-01-2026.pdf
Settled     250115_HospitalDaLuz_3-45_receipt_Claimed_04-01-2026_Settled_20-01-2026.pdf
reverted    250115_HospitalDaLuz_3-45_receipt_Claimed_04-01-2026.pdf
```

The chain is **rebuilt from the row's date columns** on every transition, not
edited in place. Going forward lengthens it, reverting shortens it, and both use
one code path — so there is no separate undo to drift out of step.

### Amounts in filenames

Rendered to two decimal places with the decimal point replaced: `3.4` → `3-40`,
`3.456` → `3-46`, `12` → `12-00`.

Two reasons. A sheet cell holding `3.4500001` would otherwise land in the
filename verbatim. And a name containing `3.45` has two dots, which invites
naive `split('.')` extension parsing to break — in this codebase, or in any
Shortcut, script or OCR step added later. One dot per filename, always the
extension.

Cost: searching Drive means typing `3-45`. Controlled by `DECIMAL_IN_FILENAME`
in `Config.js` — set it to `'.'` to keep the dot.

### Status control replaces Done/Undo

Three states cannot be a two-way toggle, so each row gets a **status selector**
rather than Done/Undo buttons.

- **"Undo" stops existing as a concept.** Going back is just selecting an
  earlier state, which removes the four inconsistent undo implementations.
- Correcting a mistake is the same action as making it — no special path.

### Reverting must not rewrite history

Reverting is expected to be common (mis-taps on a phone), so it has explicit
rules rather than falling out of the implementation:

- Moving to an earlier state **clears the date columns of every state after the
  target**, so the row never claims a date for a state it is no longer in.
- The target state's own date is **only filled if blank**. Reverting Settled →
  Claimed keeps the original `Claimed Date` rather than re-stamping today.
- Any file rename applied by the states being reversed past is undone.

### Reversibility instead of confirmations

Because every state change is freely reversible, state changes get **no
confirmation dialog** — confirmations on frequent actions just train you to tap
through them. Only genuinely irreversible actions (archiving) confirm.

### Section-specific fields — and why they stay

**The forms exist so that a claim can be submitted without reopening the
receipt.** Número and Emitente NIF get retyped into Finanças; a health claim
needs both the treatment date and the invoice date. Capturing them once at entry
is the whole point of the system.

So completeness beats minimalism. Do not trim these back to a tidy shared core —
a dropped field becomes a receipt you have to go and find.

| Section | Extra fields |
|---|---|
| Work | none |
| IVA | `Número`, `Emitente NIF`, `IVA Amount` |
| Health | `Invoice Date`, `Service Type` |
| Income | none — its extra dates are state dates |

**Health `Service Type`** is a deliberately short list: Doctor, Dentist,
Optician, Prescription, Exam/Test. It is **not** the insurer's list, which is
huge and multi-level and is chosen at submission time. This one exists for your
own tracking and is optional.

**Income** is Date, Counterparty, Amount, Currency, optional Reason, plus its
three state dates. `Received Date` and `Logged Date` are ordinary editable
fields shown in the form — settable at entry or later — as well as being filled
by the status control. So an income entry can arrive already `Received`.

`Amount` holds the total in every section; IVA's VAT figure is its own
`IVA Amount` field. `Date` holds the transaction date — for Health that is the
**treatment** date, with the invoice date alongside it.

### Reference values are shown, not stored

`Tipo` and `Importados` are identical on every IVA row, so they were never
really data — they were a reminder. They stop being columns and become a
**reference block displayed in the section**, alongside the NIFs you have to
retype into Finanças and cannot always remember:

| Label | Source |
|---|---|
| JALLC NIF | `REF_JALLC_NIF` |
| My NIF | `REF_MY_NIF` |
| Tipo | `REF_IVA_TIPO` |

Shown with **tap to copy**, since the whole point is transcribing them into
another system.

Values live in Script Properties rather than in `Config.js`. They are identifiers
rather than secrets, but the repository is public and there is no reason to put
a personal NIF into it twice — note that v1's `index.html` already contains
these values publicly, which is worth cleaning up when the old system retires.

### Health has two documents, not one

A health claim needs proof the expense was **necessary** and proof it was
**paid** — usually a prescription or invoice, plus a payment receipt. v1 called
the second file "Details", which hid what it was for.

| Column | Label | Filename suffix |
|---|---|---|
| `Justification URL` | Prescription / Invoice | `_justification` |
| `Receipt URL` | Proof of payment | `_receipt` |

### What v2 does drop

- **Health** — `Original Receipt Filename` / `Original Details Filename` (M/N).
  Existed only for the iCloud Shortcut, which is cancelled.
- **All** — `Description` becomes `Notes`, and is optional.

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

### Supplier registry

A `Suppliers` sheet that **populates itself**. Nothing is entered up front:
every entry teaches it, so it is current by construction rather than by
maintenance.

| Column | Purpose |
|---|---|
| `Name` | Canonical name |
| `Type` | Default service type, e.g. Uber → Taxi, White Clinic → Dentist |
| `NIF` | IVA retailers only |
| `Aliases` | Recurring mishearings, mapped once and fixed forever |
| `Times Used`, `Last Used` | Orders autocomplete by what you actually use |

Serves the form (autocomplete, prefill Type and NIF) and Siri (matching a
misheard name to something real).

**Matching is tiered**, strongest first: exact after normalising (1.00), alias
(0.95), one name containing the other scaled by length ratio (0.75–0.95), then
edit-distance similarity. Accents, case and punctuation are ignored throughout.

**Nothing prefills below 0.85 confidence.** A wrong NIF means a rejected claim,
which is far worse than a blank field — below the bar we keep what was actually
said and let the completion step resolve it. Measured behaviour:

| Heard | Stored | Score | |
|---|---|---|---|
| `wite clinic` | White Clinic | 0.92 | autofills |
| `the white clinic` | White Clinic | 0.90 | autofills |
| `white clinique` | White Clinic | 0.79 | holds |
| `fnak` | FNAC | 0.75 | holds |
| `uber` | Uber Eats | 0.84 | holds |

Short names are inherently brittle — one wrong letter in four tanks the ratio.
That is what `Aliases` is for: correct it once, and it resolves at 0.95 forever.

**What the registry prefills differs per section**, because "usually the same"
is not true everywhere:

| Section | Prefills | Why |
|---|---|---|
| IVA | `Emitente NIF` | A fact about the retailer |
| Health | `Service Type` | White Clinic is usually Dentist |
| Income | `Reason` | Currently fixed per payer |
| Work | nothing | The same supplier serves many trips, so Expense Reason genuinely varies and a default would be wrong more often than right |

Income's `Reason` is therefore free text rather than a managed list: you enter
it once per payer and it fills itself thereafter, which removes the reason it
would otherwise go unfilled.

**Types are cleared on conflict, not overwritten.** A clinic doing both
dentistry and exams has no reliable default, so the second differing type blanks
it. A field that prefills the wrong value is worse than one that prefills
nothing.

### Management module

v1's only management action was add/delete expense reason. v2 gets a proper
management surface, because mistakes are normal and correcting them shouldn't
mean opening the spreadsheet on a phone.

**Edit a row.** Any field, in place, from the table. Editing is not a special
mode — the same validation as `createEntry` applies, so an edited row can never
be less valid than a created one.

**Delete a row — archives it.** Deleting from the main table moves the row to
the section's archive sheet marked `deleted` and moves its documents to the
`Archived` folder. It does **not** remove data. Given how easy it is to mis-tap,
a one-click unrecoverable delete of the wrong row is not worth the convenience,
and the row you meant to remove is junk anyway.

**Hard delete — really deletes, and only from the archive.** The management
module can permanently destroy a row and its files.

The safeguard is structural rather than a scarier dialog: **hard delete only
operates on rows that are already archived.** Live data can never be destroyed
in one action; you archive first, then purge from the archive. Two deliberate
steps, in two different places, with a confirmation on the second.

Drive files are moved to Drive's trash rather than removed outright, which keeps
a 30-day grace period. Storage is only reclaimed when the trash empties — if
space is urgent, permanent removal is a config switch.

| Action | Where | Reversible | Confirms |
|---|---|---|---|
| Change status | table | yes, freely | no |
| Edit a field | table | yes, by editing back | no |
| Delete a row | table | yes — lands in archive | yes |
| Hard delete | management, archived rows only | no (30 days in Drive trash) | yes |
| Archive a category value | management | yes | yes |

**Manage category values.** Add and remove the allowed values of a section's
category field, which updates the linked form's dropdown. Works for Work's
Expense Reason and Health's Patient; hidden for IVA and Income, which have no
category.

### Siri intake

**Decision: prompted questions, then a visual confirmation.** Not one dictated
sentence.

The reason is not recognition accuracy, it is memory: prompts are a **checklist**.
A field the Shortcut insists on asking for cannot be forgotten, whereas a single
sentence quietly omits whatever you did not think to say.

One Shortcut per section, named so the phrase is natural — "Log expense",
"Log health claim". No "which section?" question.

```
You:   Hey Siri, log health claim
Siri:  Who is it for?        -> tap Phoenix from a list
Siri:  Which provider?       -> "White Clinic"
Siri:  How much?             -> "70"

       ┌──────────────────────────────┐
       │  Health claim                │
       │  Phoenix · White Clinic      │
       │  €70.00 · today              │
       │        Save    Cancel        │
       └──────────────────────────────┘
```

**Refinements that follow from choosing prompts:**

- **Do not ask for the date.** Default to today and show it in the confirmation.
  Most entries are same-day, and the confirmation is where an exception gets
  caught. One fewer prompt every time.
- **Category fields are a list, not dictation.** Expense Reason and Patient are
  taps, which cannot be misheard. The list is **fetched from the server**, so
  adding a patient never means editing the Shortcut.
- **Amount uses the Number input type**, so words can never arrive where digits
  belong. Amount errors are the dangerous ones — a supplier typo is obvious in a
  list, "29.80" instead of "298" is not.
- **Currency defaults to EUR** and is not asked.

**Siri captures the core only** — counterparty, amount, category. Número, NIF,
invoice date and documents are left for the completion step. This keeps the
Shortcuts stable: adding a field to a section never requires re-editing them,
because they only ever ask for the same few things.

> Note: an Alert action is right here, where you invoked the Shortcut and are
> looking at the phone. It is wrong in an unattended *automation*, where it
> stalls waiting for a tap — which is why one was removed from the v1 iCloud
> Shortcut.

### Partial entries are the safety net

**An entry does not have to be complete.** Whatever Siri did not capture stays
blank, `Receipt State` becomes `awaiting`, and an email arrives with a
**completion link** that opens the form on that row.

So a mishearing costs one tap on a link rather than a failed capture. It is also
why the registry must hold rather than guess: a blank you will see and fix beats
a confident wrong match that silently corrupts the entry.

The completion link is only possible without Google Forms — Forms cannot reopen
an existing row for editing in any usable way.

**Learning from corrections.** When a completion edit changes the supplier from
what was heard, that spoken form is a candidate `Alias`. Adding it — offered,
not automatic — means the same mishearing resolves next time.

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

- **Keep `3-45` or revert to `3.45` in filenames?** One line in `Config.js`.
- **Is `Invoiced` the state every Income row starts in?** If so its date is
  manual, so the form must ask for it at submission.
- **Income has no files** — so it has no folders and no filename suffixes. Does
  it need a Drive presence at all, or is the sheet row the whole record?
- **Can you advance to a state without its date yet**, or is the date required?
- Exact current field names for Income — read from the sheet before building
- IVA email recipient address for the new account
- Whether the four sections should stay four Forms or become one with branching
- Where OCR runs: Drive's built-in OCR vs an API, and how confident it needs to
  be before autofilling rather than prompting
