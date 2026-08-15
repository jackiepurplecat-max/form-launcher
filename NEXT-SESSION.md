# Start here

Handover note, written 15 Aug 2026. **Read *The plan* below and work down it**;
everything after it is the record of how things got here, and is reference
rather than instruction.

**Operational state only** — the design and the reasons live in
`REBUILD-PLAN.md`, which is the source of truth. This file is disposable:
overwrite it at the end of each session.

## Where things stand

| | |
|---|---|
| Branch | `step-7-web-ui`. **Ahead of `origin`** — push it. No count here: this note is one of the commits, and it cannot include itself |
| Last code commit | the tip — the UI rebuild. `42f0a53` before it is the preview tool |
| Working tree | clean |
| Harness | **727 passing, 0 failed** — unchanged by the UI work, and that is the point |
| Main project | matches `v2/` byte for byte, 13 files |
| Siri project | `v2-siri/`, 2 files — not touched on 15 Aug |
| Deployed | main at **version 28**, cut 15 Aug. Siri at **@2** |
| Shortcuts | all four working. **The IVA one still has no Tipo picker** |

Steps 1–9, 9c and 11 are done. **The web UI was rebuilt around the phone on 15
Aug and is deployed** — see below for what changed and what has still never been
touched by a finger. Cutover (step 10) remains the main thing standing between
this and daily use.

---

# The plan

In order. **Step 1 is blocking and new** — the deployment on 15 Aug turned a
pending item into a broken one.

## 1. Run `bootstrap()`. IVA is broken until you do — 2 minutes

Not optional, and it is more urgent than it was yesterday. `SECTIONS` declares
IVA's `Tipo` column; the IVA sheet does not have it. `columnIndex` throws
`Column "Tipo" not found in IVA` (`v2/Core.js:169`) rather than degrading, so
**`uiListEntries('iva')` fails and the IVA view will not list at all.**

This was harmless while `/exec` served version 27, which predated the change.
**Version 28 includes it**, so the live app is now in that state. The other three
sections are unaffected.

Run `bootstrap()` from the main project's editor. It appends `Tipo` to IVA, is
idempotent, and reports what it added. Then open IVA and confirm the eight codes
appear as a closed list — a closed list rendering as free text is a failure this
project has hit before.

## 2. Look at the new UI on the phone

Deployed and unseen on a real device. Open via the **home screen icon**, which
has its own cookie jar and is why the v2 account wins:

```
https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec?authuser=purplecat.admin@gmail.com
```

Everything below was verified in headless Chrome at 390 and 1100 across all five
views, with no horizontal overflow anywhere. **What headless Chrome cannot
speak for**, and therefore what to actually check:

- **Both date controls.** The status dialog now stacks its input above OK; the
  form's inputs got smaller type and padding. Chrome does not render iOS's
  native date wheel, so this is the fix with the least evidence behind it.
- **The advance/regress pair.** 42px minimum, but feel is not measurable in a
  screenshot.
- **Walk a Health row** To Do → Claimed → Settled and back. The destination
  section should open and scroll to itself while the source stays open.
- **A Health row with one document of two** should read
  `Proof of payment — awaiting` in amber.
- **Momentum scrolling and the accordions** — a long Claimed list on a real
  device.

## 3. Cutover — the four toggles

Unchanged, and still the thing that has been next for four sessions. There is no
code between you and it: `findDebris()` came back clean on 13 Aug, so cutover
step 1 is done.

Signed in as **jackiepurplecat@gmail.com**: `forms.google.com` → each of the four
forms → **Responses** → **Accepting responses** off. Reversible, so the risk is
the gap between the first and the fourth — do them in one sitting.

Then freeze v1, both in the old account:
- Delete the `ICLOUD_EMAIL` Script Property. `getIcloudEmail()` throws, the caller
  catches, no mail is sent, and Health Done still flips status and renames files.
- Delete the iOS automation so the failure notifications stop.

Leave the old GitHub Pages page live and untouched to work the backlog down.

## 4. The IVA Shortcut needs its Tipo picker

`catalog` returns a category for IVA where it returned `null`, so the phone can
offer the closed eight. **The Shortcut has not been updated.**
`SIRI-SHORTCUT-REBUILD.md` shows how health's picker is wired; IVA needs the
same, reading `category.values`.

Until it is done, Siri IVA entries arrive with no Tipo — a deferred field like
any other rather than a failure. **The export needs it**, though — see step 6.

**Export the Shortcut to a file the moment it works.** That rule exists because
three of them were lost on 13 Aug.

## 5. The last unseen surface — NIF warnings on a merge

9c works and has been used, but the NIF handling was tightened afterwards and its
two warnings have never been on screen. Merge two suppliers whose NIFs **differ**;
then merge into one with **no** NIF and confirm it says the core has *inherited*
one. Matching NIFs must say nothing at all.

Note the registry is now reached as **Providers** in the view selector. The sheet,
the server and every identifier still say supplier.

## 6. The IVA export — the substantial piece of new work

Goal: **open DRORIVA pre-filled from the IVA section instead of retyping every
invoice.** The format is fully reverse-engineered in **`v2/IVA-EXPORT-FORMAT.md`** —
read that first, it has five traps that each produce a file the app silently
refuses.

Two decisions to make before writing code, neither of them mechanical:

- **Round-trip proof first.** Regenerate the existing sample from its own values
  and confirm DRORIVA opens it identically — in particular whether omitting the
  61k NUL padding bytes is fine. Cheap, and it de-risks everything after it. Do
  not wire it to real rows before this passes.
- **Which rows, and how does the sheet know?** Presumably complete IVA entries not
  yet submitted, which needs a *submitted* marker the sheet does not have — a new
  column and probably a new state. That is a design decision, not a detail.

  **Note this now interacts with the UI.** A new state is a fifth accordion, and
  the advance/regress buttons are generated from `SECTIONS.states`, so it would
  get its buttons for free — but think about whether "submitted" is a status or a
  flag before adding one.

Then the generator itself. Where it lives is open: an Apps Script function has the
sheet and Drive to hand but holds the whole base64 payload in memory, and one
invoice was 111 KB of PDF for 231 KB of XML. **Size is the thing most likely to
break first.**

## 7. Field validation — step 13

`VALIDATION-PLAN.md`, five rules, none implemented. **Nothing validates values on
any intake path today** — a quoted `"Amount":"abc"` from Siri would be written to
the sheet as text. This is a hole, not a polish item.

AT's own `Decimal_15` is `fractionDigits=2 minInclusive=0` (rules 1 and 2) and
`NifFatura` is `xs:long` (rule 3) — two independent routes to the same rules.

Doing it **after** step 6 is deliberate: the export is what turns a bad value from
an annoyance into a rejected submission, so build the thing that punishes bad data
before hardening against it.

## 8. Not next

- **OCR intake, step 12** — new capability, can wait.
- **The plan's open questions** — `3-45` versus `3.45` in filenames, whether
  `Invoiced` is really every Income row's starting state, whether a state can be
  advanced without its date, and confirming Work's `Type` list against reality.
- **The NIFs in the public repo** — decided 13 Aug to leave. See the section near
  the end; it cannot be closed cleanly until v1 is decommissioned anyway.

---

# The UI rebuild, 15 Aug

The phone is the device this is used on and the page was built desktop-first.
Health rendered fourteen columns in a nowrap table, so moving a status or editing
a row meant scrolling about a thousand pixels right; the header was a
non-wrapping row of six items and Refresh fell off it.

**What the page is now.** Header is the title, your address and Refresh. Below it
a view strip — five pills on a desktop, one `<select>` on a phone — then `+ New`,
which creates into the view you are looking at and is absent in Providers. Each
entry view is four accordions: the three declared statuses, then the archive.
Rows are cards under 600px and the table above it.

**Decisions worth not re-litigating:**

- **Status is a place, not a value you set.** The sections partition by status, so
  moving a row moves it between them. That is why the Status column and the
  Status filter are both gone — each was saying what the heading already said.
- **Advance is named for where it goes, regress for the state it leaves** —
  "Claimed" and "Not Claimed". The negative names the thing on screen you have
  decided is wrong; "To Do" would make you work out what the earlier state was
  called first.
- **The date dialog stays on the way forward.** Claimed, Settled and Received are
  usually backdated, and one tap to accept today is the same tap the dropdown
  cost.
- **The archive is fetched on first open, not with the section.** One round trip
  per view on mobile data beat a count that is always right, so the header reads
  `Archive` rather than `Archive (4)` until you have looked once.
- **"Providers" is a UI relabel only.** The sheet is still Suppliers and so is
  every identifier on the server. Renaming those is a data migration.
- **A status the config does not declare gets its own section**, present only
  when it has rows. Four buckets keyed by status would otherwise drop those rows
  entirely, and a row that exists in the sheet but nowhere on the page is exactly
  the failure the loading/empty/error rule exists to prevent.
- **Cards and table are built by different code**, not one linearised by CSS: the
  column order is the server's and it puts Amount seventh, which makes a headline
  impossible.
- **Empty is now said per section.** Loading and error stay page-level, so a load
  failure still shows no sections at all.

**Two fixes named in use.** A Health row with one of its two documents attached
looked exactly like one with both — `uiRow` drops empty file columns and the
row-level `receiptState` was only ever shown when there were *no* documents. The
page now diffs against `meta.files` and names the gap. And both date controls
overflowed a phone; the dialog's input could not shrink below a native control's
min-content because a flex item defaults to `min-width: auto`.

## `npm run v2:preview` — new, and the reason any of this is trustworthy

`v2/test/preview.js`. Renders the real page against the real server output and
screenshots it, and can click through a sequence:

```
npm run v2:preview                                   # 5 views, 390 and 1100
node v2/test/preview.js --view=health --width=390 --height=844 \
  --click='[data-advance]' --click='#dialogPrimary' --label=moved
```

It reports horizontal overflow per element, which is how IVA's category filter
was caught pushing the page 57px wide. **Two traps baked into it, both found the
hard way and both worth knowing before you change it:**

- **Chrome will not lay out below 500 CSS pixels**, in either headless mode, and
  `--screenshot` then *crops* to the width you asked for. A "390px" shot is a
  500px layout with 110px cut off — which looks exactly like a page overflowing
  its viewport, so the tool would have been inventing the bugs it exists to
  find. `--force-device-scale-factor` does not help; it changes pixel density,
  not the viewport. **The app is rendered in an iframe of the exact width**,
  which is a real viewport that media queries answer to. The dark strip on the
  right of a narrow screenshot is the host page, and marks where the viewport
  ends.
- **`getBoundingClientRect` reports the full width of a node inside a scrolling
  ancestor**, so the desktop table — which is *meant* to be wider than its
  wrapper — buried the one real finding under sixty rows of table cells. The
  audit now skips anything inside an `overflow-x` ancestor.

It cannot reproduce iOS's native date wheel, momentum scrolling or tap-target
feel. Those still need the phone, which is why step 2 exists.

## First thing: establish the baseline

```
npm run v2:test          # expect 727 passing, 0 failed
npm run v2:verify        # expect "Server matches v2/ — 13 files, byte for byte"
npm run v2:siri:verify   # expect "Server matches v2-siri/ — 2 files, byte for byte"
```

If any of those disagree, find out why before changing anything.

## Things that will waste your time if you do not know them

- **Do not put `-X POST` in the curl.** Apps Script answers `/exec` with a 302 to
  `script.googleusercontent.com/macros/echo`, and that endpoint serves **GET
  only**. Plain `curl -L -d …` is right: curl downgrades the redirected request
  to GET by itself. `-X POST` forces the method to stick across the redirect and
  you get **405** with a Drive "Página não encontrada" page — which reads exactly
  like a dead deployment and is not one.
  The tell is in the redirect: `curl -s -D - -o /dev/null -X POST … | grep -i location`.
  A `location` carrying **`&lib=…`** means the library resolved and `doPost`
  already ran, whatever status the next hop returns.
- **Library code reads the LIBRARY's Script Properties — confirmed, not
  assumed.** A live `ping` returned `propertiesVisible` all true and
  `"spreadsheet": "HelpfulForms"`, so `v2/Siri.js` running as a library sees the
  **main** project's properties. **Nothing needs duplicating onto `v2-siri`, and
  nothing should be.**
- **`clasp push` reports success while pushing nothing.** Always verify after.
  It printed `Pushed 13 files.` on 15 Aug and `v2:verify` agreed — but the output
  is not the evidence, the verify is. On a mismatch, `npm run v2:push:force`.
- **`v2-siri/appsscript.json` is generated and gitignored.** It holds the main
  project's script id. `npm run v2:siri:push` regenerates it first. Written with
  **no trailing newline**, for the same one-byte reason as `v2/appsscript.json`.
- **`v2/appsscript.json` has no trailing newline, on purpose.** `wc -c` should be
  **425**, not 426. Do not tidy it.
- **Pushing is not deploying.** A push updates HEAD, which `/dev` serves. `/exec`
  serves a pinned version, so cut a new one — and pass `-i <deploymentId>` or
  clasp creates a *second* deployment on a different URL:
  ```
  cd v2 && clasp --user v2 deploy -i AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo -d "what changed"
  ```
  **The main deployment is version 28 and is current** — cut 15 Aug from the UI
  rebuild. `clasp --user v2 list-deployments` should show exactly two: the
  permanent `@HEAD` and `@28`.
- **Never write the Siri `/exec` URL into a tracked file.** This repo is public,
  and unlike every other URL in this note that one is anonymous: the address *is*
  the reachable surface. It cannot be used to write anything — an unset or wrong
  key returns `Not authorized.` — but a published anonymous endpoint attracts
  scanning and spends execution quota. It happened once already: `@1` went into
  this file, so it was deleted and redeployed as `@2`. Get the URL from
  `clasp --user v2 list-deployments`, keep it in the Shortcut and in `.env`.
- **A clean reload fixes a dead file picker on iOS.** iOS suspends the sandboxed
  iframe, killing the user-activation context the picker needs while the page
  still looks fine. Close the tab and reopen. **Do not write defensive code for
  this** without reproducing it. Nothing is lost — uploads land before anything is
  written.
- **Diagnose access failures from `appsscript.json`, not from the error text.**
  - *"You need access"* = right file, wrong account.
  - *"Cannot open the file"* = no rights to the script itself.
- **`authuser=` does NOT switch accounts.** It *selects* among accounts already
  signed in to that browser. If the v2 account is not signed in there, the
  default answers and the page refuses — no URL can override that.

  **`/u/N/` does not work either.** `script.google.com/u/1/macros/s/<id>/exec`
  returns **404** for every N. Tested 0, 1 and 2 on 13 Aug — do not spend the
  hour again.

  **The symptom is Drive's access error, not ours.** *"Precisa de acesso"* means
  Google refused before the script ran — `access` is `MYSELF`, so a browser
  authenticated as `jackiepurplecat` never reaches `doGet`.
- **Add to Home Screen is the durable phone session**, because iOS gives a
  home-screen web app **its own cookie jar**, separate from Safari's. That is what
  stops the default account winning, and it is stronger than `authuser=` because
  it removes the ambiguity rather than resolving it. Confirmed working 13 Aug.
  The fallback is a Private Browsing tab. iOS 17 Safari Profiles is the other
  clean answer.
- **A refused visitor gets a page that says why** — it names the account you are
  signed in as and offers Switch Google account. **But you will almost never see
  it**: `"access": "MYSELF"` means Google refuses other accounts before `doGet`
  runs and serves *its* page, not ours.
- **The harness still cannot click**, and now there is something that can — see
  `v2:preview` above. Between them the phone, the deployment and the library
  resolution are still outside any automated check.
- **A stray `Folha1` tab** may still be in the spreadsheet. Harmless; delete by
  hand if empty.

## Known and accepted: real NIFs are in the public repo

**`index.html` is tracked, not gitignored, and carries both NIFs in plaintext** —
JALLC's and the personal one, at roughly `index.html:621` and `:623`. `CLAUDE.md`
describes that file as *"(with actual values, gitignored)"*, which is **wrong**:
`git check-ignore` returns nothing and `git ls-files` lists it, so every build has
been published. A real supplier NIF also went into the first version of
`v2/IVA-EXPORT-FORMAT.md`; the working copy is scrubbed to `NNNNNNNNN`, but commit
`c349aa7` still holds it.

**Decided 13 Aug: leave it.** These NIFs appear on invoices and correspondence
anyway, so the exposure was judged low. Recorded rather than fixed so it is a
decision and not an oversight, and so nobody rediscovers it and panics.

**If that judgement ever changes**, in increasing order of cost:

1. `git rm --cached index.html` and add it to `.gitignore` — matches what
   `CLAUDE.md` always claimed, stops future commits, leaves history alone.
2. The above plus `git filter-repo` and a force-push — actually removes them from
   GitHub after GC, at the cost of rewriting every hash and breaking clones.
3. Make the repo private — but v1's launcher page is served from **GitHub Pages**
   and needs it public until the backlog reaches zero.

Note the ordering constraint in 3: this cannot be closed off cleanly until v1 is
decommissioned, which is cutover step 5.

## Settled, so do not re-litigate

- **The Siri logic lives in the MAIN project, not the shim.** The shim is two
  files and holds no logic, no secrets and no configuration. A copy-the-source
  second project would need its own Script Properties and nothing would catch the
  two stores drifting.
- **The library is in development mode**, so the shim runs main's HEAD.
- **`resolve` before `create`.** The confirmation happens before anything is
  written.
- **`create` never re-runs the fuzzy matcher.** Exact match only, to fill blank
  `Type`/NIF.
- **An unset `SIRI_API_KEY` shuts the endpoint.** Not "no key required".
- **A field outside the whitelist is refused, not dropped.** No document column
  is accepted from Siri at all — `extractFileId` would take a Drive id out of any
  string and the script runs as you.
- **A corrected NIF is never backdated.** Pinned by the harness.
- **The completion mail lands in the inbox.** The HTML body with a real `href`
  fixed the junk filing.
- **On a supplier merge the target's spelling survives**, the NIF defaults to the
  **core** entry's, and the registry does not move until every row carries the
  new name.
- **Income's `(none)` reason option was deliberately not built** — there is no
  occasion to log income without knowing what earned it.
- **`Config.js:284` stays `required: false`.** Enforcing it server-side would turn
  existing blank-Reason income rows into INCOMPLETE entries and nag about
  finished work.

## The Shortcuts were destroyed and rebuilt, 13 Aug

Kept because the cause is a standing trap. `Log expense`, `Log receipt` and
`Log income` were lost and unrecoverable, and were rebuilt from the surviving
`Log health claim`.

**Cause: iCloud Drive was switched off on the phone.** That one setting did two
things: exporting a Shortcut to a file failed, because Apple *signs* `.shortcut`
files against a live iCloud session; and Shortcuts sync sat in a state where the
phone's near-empty library reconciled **over** the Mac's four. There was no export
and no Time Machine — `tmutil` reported no destinations and no local snapshots.

**The rules that follow:**

- **Export every Shortcut to a file the moment it works.** Share → Save to Files.
  Not at the end of the session.
- **Those files contain `SIRI_API_KEY` and the `/exec` URL.** iCloud Drive is
  fine; this repo is public and must never hold one.
- The verified template backup is `iCloud Drive/Downloads/Log health claim.shortcut`
  — 26,163 bytes, `AEA1` signed.
- To get one back into Shortcuts on the Mac, **double-click the file.** Importing
  needs no signing, so it works regardless of sync state.

**`v2/SIRI-SHORTCUT-REBUILD.md` is the rebuild record** and the better of the two
Shortcut documents to work from. **It corrects `SIRI-SHORTCUT.md:322`**, which
says to delete steps 3, 4 and 5 for IVA: that predates the medium question, and
step 3 is the `catalog` call feeding `receiptMedium.values`. Keep step 3, delete
only 4 and 5.

**Building Shortcuts leaves debris.** Each failed run writes a blank or
part-filled row and sends a completion mail, and `create` teaches the registry
whatever it was given. Audited 13 Aug and clean, but any future Shortcut work
owes another audit.

**`findDebris()` in `v2/Smoke.js` is the tool.** It **reports and never deletes**,
because a part-filled row awaiting a document is indistinguishable from a real
deferred entry. Two confidence levels: `certain` (no counterparty, or no usable
amount — both intake paths always set both) and `suspect` (complete, Siri-sourced,
awaiting a document, no category). `findDebris('2026-08-13')` narrows to a day.

**An empty report never means "ready for cutover".** A test run that *succeeded*
writes a complete, well-formed row, and nothing about `Bolt, 8 EUR, taxi` says
whether it came from a taxi or from proving a Shortcut works. Read the sheet for
that. The first line of the report distinguishes "scanned forty rows, all fine"
from "scanned nothing" — a distinction it lacked on its first run, when both
printed identically.

**`resetAllData()` is the other tool**, for a clean slate rather than a mixed
sheet. It clears data rows from all four sections and any archive sheets, trashes
the documents those rows point at, empties the registry and the Staging folder.
**Headers, sheets, the Drive tree, the Staging folder itself and Script Properties
all survive**, so `bootstrap()` does not need re-running and Genius Scan keeps
working. It **refuses to run** unless called as
`resetAllData('DELETE ALL TEST DATA')` — there is no confirmation dialog in the
editor, where the last function you picked is one click from running again, so the
safeguard is structural. Documents are trashed **before** their rows are deleted,
because the row is the only record of which file belongs to it.

`smokeCleanup()` is no use for debris — it only matches rows carrying
`SMOKE_MARKER` in Notes, which it wrote itself.
