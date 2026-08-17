# Start here

Handover note, written 15 Aug 2026, updated 18 Aug. **Read *The plan* below and
work down it**; everything after it is the record of how things got here, and is
reference rather than instruction.

**Nothing in `v2/` has changed since version 29.** The 17–18 Aug session went
entirely on the phone's wrong-account problem and ended in documentation only,
because the fix turned out to be a browser setting rather than code. So
everything outstanding on 15 Aug is still outstanding, and the deployed app is
the same one. What *did* change is that the app now opens reliably on the phone,
which is what makes the list in step 1 checkable at last.

**Operational state only** — the design and the reasons live in
`REBUILD-PLAN.md`, which is the source of truth. This file is disposable:
overwrite it at the end of each session.

## Where things stand

| | |
|---|---|
| Branch | `step-7-web-ui`, **pushed to `origin` and in sync** |
| Last commit | `7079831`, **documentation only**. The last *code* commit is still `5a29a30`, the iOS date-input width fix. UI rebuild `33d8622`, preview tool `42f0a53` |
| Working tree | clean |
| Harness | **727 passing, 0 failed** — the 17–18 Aug work added nothing and removed nothing |
| Main project | matches `v2/` byte for byte, 13 files — **re-verified 18 Aug** after the reverted experiment |
| Siri project | `v2-siri/`, 2 files — untouched since 15 Aug |
| Deployed | main at **version 29**, cut 15 Aug. Siri at **@2**. **No deployment was cut on 17–18 Aug** |
| Shortcuts | all four working. **The IVA one still has no Tipo picker** |
| Phone access | **Fixed 17 Aug.** The v2 account must be Safari's *default*, i.e. signed in first. Not a URL problem — see the traps |

Steps 1–9, 9c and 11 are done. **The web UI was rebuilt around the phone on 15
Aug and is deployed** — see below for what changed and what has still never been
touched by a finger. Cutover (step 10) remains the main thing standing between
this and daily use.

**`bootstrap()` was run on 15 Aug**, which unblocked IVA: version 29 declares its
`Tipo` column and the sheet now has it. Before that run the live app threw
`Column "Tipo" not found in IVA` and the section would not list at all. What has
*not* been confirmed is that the eight codes render as a closed list rather than
free text — see step 1.

---

# The plan

In order. Nothing is blocking any more — `bootstrap()` was the last blocker and
it was run on 15 Aug. Step 1 is a check rather than a build.

## 1. Look at the new UI on the phone

Deployed at version 29. **The date controls have now been seen and are right;
everything else on this list has not been.** Open via the **home screen icon**,
which has its own cookie jar and is why the v2 account wins. Hard-refresh it —
it may be holding an older version:

```
https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec
```

The `?authuser=` that used to be on that URL has been dropped: it does nothing on
`/exec` and kept the wrong idea alive. Since 17 Aug the v2 account is Safari's
default, so a plain Safari tab works too — the home screen icon is still the
sturdier route because it does not depend on that ordering holding.

Everything below was verified in headless Chrome at 390 and 1100 across all five
views, with no horizontal overflow anywhere. **What headless Chrome cannot
speak for**, and therefore what to actually check:

- ~~**Both date controls.**~~ **Fixed and confirmed on the phone at version 29.**
  Kept because the diagnosis cost two attempts and the wrong one is plausible:
  font and padding changed nothing, because iOS gives `input[type="date"]` an
  intrinsic width and will not take `width:100%` below it.
  **`-webkit-appearance:none` is what releases that sizing**, and it does *not*
  cost the native wheel — on iOS the picker is bound to the input type, not its
  appearance. An earlier version of `Index.html` avoided the property in a
  comment saying it would, and that belief was what made the first fix a no-op.
  Chrome cannot reproduce the iOS control at all, so nothing local would have
  caught either the bug or the fix.
- **IVA's Tipo.** `bootstrap()` has been run, so the column exists and the
  section lists. Confirm the eight codes appear as a **closed list rather than
  free text** — that is a failure this project has hit before, and it has not
  been looked at since the column was added.
- **The advance/regress pair.** 42px minimum, but feel is not measurable in a
  screenshot.
- **Walk a Health row** To Do → Claimed → Settled and back. The destination
  section should open and scroll to itself while the source stays open.
- **A Health row with one document of two** should read
  `Proof of payment — awaiting` in amber.
- **Momentum scrolling and the accordions** — a long Claimed list on a real
  device.

**Newly outstanding, and the reason step 1 is worth doing now.** The 17 Aug fix
was proved with a *bare* `/exec` link only. The mailed links have never been
opened successfully on the phone, so the thing the whole session was about is
still not verified end to end:

- **A real completion email's "Finish it here" link.** It carries
  `?section=<key>&t=<stamp>`, which the bare test link deliberately did not.
  Confirm it opens the app **on that entry** rather than on the default view.
  `smokeTest()` in the editor sends one; `smokeCleanup()` removes the rows after.
- **Whether the page then works, rather than merely loads.** Everything on it
  goes through `google.script.run`, and each of those 22 functions runs its own
  `requireUiAccess()`. Loading proves `doGet` passed; it does not prove the calls
  do. Save something.
- **The Step 1 staging-folder link** in the same mail. It is a `drive.google.com`
  URL carrying `authuser=`, which is a *different* mechanism from `/exec` and is
  the one place `authuser` genuinely does work — so it can fail independently.
- **That it survives a cold start.** Force-quit Safari, or leave it overnight,
  then tap the mailed link again. The default-session fix is only as good as the
  cookie that carries it.

## 2. Cutover — the four toggles

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

## 3. The IVA Shortcut needs its Tipo picker

`catalog` returns a category for IVA where it returned `null`, so the phone can
offer the closed eight. **The Shortcut has not been updated.**
`SIRI-SHORTCUT-REBUILD.md` shows how health's picker is wired; IVA needs the
same, reading `category.values`.

Until it is done, Siri IVA entries arrive with no Tipo — a deferred field like
any other rather than a failure. **The export needs it**, though — see step 5.

**Export the Shortcut to a file the moment it works.** That rule exists because
three of them were lost on 13 Aug.

## 4. The last unseen surface — NIF warnings on a merge

9c works and has been used, but the NIF handling was tightened afterwards and its
two warnings have never been on screen. Merge two suppliers whose NIFs **differ**;
then merge into one with **no** NIF and confirm it says the core has *inherited*
one. Matching NIFs must say nothing at all.

Note the registry is now reached as **Providers** in the view selector. The sheet,
the server and every identifier still say supplier.

## 5. The IVA export — the substantial piece of new work

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

## 6. Field validation — step 13

`VALIDATION-PLAN.md`, five rules, none implemented. **Nothing validates values on
any intake path today** — a quoted `"Amount":"abc"` from Siri would be written to
the sheet as text. This is a hole, not a polish item.

AT's own `Decimal_15` is `fractionDigits=2 minInclusive=0` (rules 1 and 2) and
`NifFatura` is `xs:long` (rule 3) — two independent routes to the same rules.

Doing it **after** step 5 is deliberate: the export is what turns a bad value from
an annoyance into a rejected submission, so build the thing that punishes bad data
before hardening against it.

## 7. Not next

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
- **A `git push` 403 is the wrong GitHub identity, not a broken remote.** The repo
  is public, so clone and fetch work anonymously and only the push fails — which
  reads like a permissions bug on the repo and is not one. Sorted on 15 Aug: `gh`
  was authenticated as **`pnhknrt7kp`**, which has `pull: true, push: false` on
  `jackiepurplecat-max/form-launcher`.

  Diagnose it with `gh api repos/jackiepurplecat-max/form-launcher --jq
  '.permissions'` rather than from the error text. Both accounts are now stored;
  `jackiepurplecat-max` is active and `gh auth switch` moves between them.

  **`gh auth login` alone was not enough.** Git's helper was `osxkeychain`, which
  still held the old account's token and answered first, so the push kept failing
  with `gh` looking correct. `gh auth setup-git` fixed it by writing a
  github.com-specific helper that takes precedence. The stale token is still in
  the keychain, bypassed rather than removed — so deleting the
  `credential.https://github.com.helper` lines from `~/.gitconfig` would bring the
  403 straight back.
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

  **The account chooser does not work either. Tested 17 Aug, on the phone.**
  `accounts.google.com/AccountChooser?Email=<v2>&continue=<encoded /exec>` was
  built, pushed and opened in Safari: it failed exactly as the bare link does.
  The un-pinned form (`?continue=` alone) *did* show the picker, listing all
  three accounts — **choosing the right one still failed the same way.** So the
  chooser resolves its own selection and `/exec` goes on answering as the
  browser's default regardless. The code was reverted; nothing is left in the
  tree. **The conclusion is stronger than "no URL parameter works": no URL
  works, chooser included, because the session that answers `script.google.com`
  is not something a link gets to choose.** Do not try Method 1, 2 or 3 of any
  suggestion that offers them — all three are now tested and dead.

  **WHAT ACTUALLY FIXED IT, confirmed on the phone 17 Aug: make the v2 account
  Safari's default.** Sign out of every Google account, then sign in as
  `purplecat.admin@gmail.com` **first**, then add the others back. The default
  session is the FIRST one signed in — which is the same reason a Private tab
  always worked, and it is the whole explanation for this bug. The plain
  `/exec` link then opens the app, with no parameter on it and no code change of
  any kind. **This is a browser-session problem, not a link problem**, which is
  why every attempt to solve it in the URL failed.

  **The trap: it comes back.** Nothing pins the ordering. Sign into another
  Google account first on that device — or reset Safari, or get signed out and
  restore in a different order — and the default moves and every link fails
  again, with the same misleading "cannot open the file". If that happens, do
  not debug the deployment or the links: check the account order first. The
  sturdier versions of the same fix are the **home-screen icon** (its own cookie
  jar) and an **iOS 17 Safari Profile** holding only this account.

  **If it ever needs to survive the ordering permanently**, the only real answer
  is to take the Google session out of the path the way the Siri endpoint does —
  `ANYONE_ANONYMOUS` plus a shared key. That is a genuine piece of work, not a
  manifest flip: the 22 `requireUiAccess()` gates read
  `Session.getActiveUser()`, which is blank under anonymous, so the key has to
  reach all of them too. Not worth it while the ordering holds.

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
