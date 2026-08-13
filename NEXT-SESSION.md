# Start here

Handover note, written 12 Aug 2026, **revised 13 Aug** after three of the four
Shortcuts were destroyed and rebuilt the same day. **Operational state only** — the design and
the reasons live in `REBUILD-PLAN.md`, which is the source of truth. Read that
after this. This file is disposable: overwrite it at the end of each session.

## Where things stand

| | |
|---|---|
| Branch | `step-7-web-ui`, pushed to `origin` |
| Last code commit | the tip — `resetAllData()`. `3ece923` before it is `findDebris()`. No hash here on purpose: a commit cannot cite its own |
| Working tree | clean, and pushed to `origin` |
| Harness | **719 passing, 0 failed** (was 710) |
| Main project | matches `v2/` byte for byte, **13 files** — `Siri.js` is new |
| Siri project | **new** — `v2-siri/`, matches byte for byte, 2 files |
| Deployed | main at **version 27**. Siri at **@2**, authorised and answering |
| Shortcuts | **All four working again** — health, work, iva, income. Three were destroyed and rebuilt 13 Aug; see below |

Steps 1–9, 9c and **11 are done**. The Siri endpoint is built, harness-covered,
deployed, configured and exercised live against the real spreadsheet. Three of
the four Shortcuts were destroyed on 13 Aug and rebuilt the same day; **all four
are confirmed working.** Cutover (step 10) is still not started, and is again the
main thing standing between this and daily use.

**Everything code-side is done and checked.** Closed on 13 Aug: the debris audit
(clean), the staging picker, the phone document link, and **a durable phone
session — the home screen icon works**, which was the last thing daily use
needed. **Cutover (step 10) is the next thing and it is operational, not code**:
four toggles in the old account. See "Pick up here", item 7.

**What is genuinely left, in four handfuls:**

1. **Cutover** — turn the old forms off, work the v1 backlog down, decommission.
   No code.
2. **One unseen verification** — the NIF warnings on a supplier merge. Everything
   else has now been on screen at least once.
3. **Field validation, step 13** — `VALIDATION-PLAN.md`. **Nothing validates
   values on any intake path today.** The real hole rather than a polish item.
4. **OCR intake, step 12**, plus the plan's open questions. New capability; can
   wait.

## The Shortcuts were destroyed and rebuilt, 13 Aug

`Log expense`, `Log receipt` and `Log income` were **lost and unrecoverable**,
and were rebuilt from the surviving `Log health claim`. Kept here because the
cause is a standing trap, not a one-off.

**Cause: iCloud Drive was switched off on the phone.** That single setting did
two separate things, and the second is what cost the work:

1. Exporting a Shortcut to a file failed — Apple *signs* `.shortcut` files
   against a live iCloud session, so "signed into iCloud" is not enough. The
   error says you are not signed in, which is misleading.
2. Shortcuts sync sat in a state where the phone's near-empty library
   reconciled **over** the Mac's four. AirDropping health to the phone appears
   to have been the trigger.

There was no export and no Time Machine on that machine — **`tmutil` reports no
destinations and no local snapshots** — so there was nothing to restore from.
Recently Deleted was empty on both devices.

**The rules that follow from this:**

- **Export every Shortcut to a file the moment it works.** Share → Save to
  Files. Not at the end of the session.
- **Those files contain `SIRI_API_KEY` and the `/exec` URL.** iCloud Drive is
  fine; this repo is public and must never hold one.
- The verified template backup is
  `iCloud Drive/Downloads/Log health claim.shortcut` — 26,163 bytes, `AEA1`
  signed, confirmed downloaded on the Mac rather than a placeholder.
- To get one back into Shortcuts on the Mac, **double-click the file.** Importing
  needs no signing, so it works regardless of sync state.

**`v2/SIRI-SHORTCUT-REBUILD.md` is the rebuild record** — three tap-by-tap
walkthroughs as actually performed, replacing the diff table it started as. It is
now the better of the two Shortcut documents to work from; `SIRI-SHORTCUT.md`
keeps the protocol and the six behaviours that cost the original build hours.

**It corrects `SIRI-SHORTCUT.md:322`**, which is wrong and cost time on the
rebuild: it says to delete steps 3, 4 and 5 for IVA, but that predates the medium
question, and step 3 is the `catalog` call feeding `receiptMedium.values`. Keep
step 3, delete only 4 and 5.

**Two decisions taken during the rebuild**, both recorded with reasons in that
file:

- **Income's `(none)` option was deliberately not built** — there is no occasion
  to log income without knowing what earned it, at the moment of logging it. The
  picker guarantees a value anyway.
- **`Config.js:284` stays `required: false`.** Enforcing it server-side would
  turn existing blank-Reason income rows into INCOMPLETE entries and nag about
  finished work — the trap already documented at `Config.js:68`. Capture-time
  enforcement in the Shortcut is where the design puts it.

**The curl test left nothing behind** — its Work row and `ZZ Siri Test` registry
entry were deleted and the removal verified through the endpoint.

**Building the Shortcuts leaves debris — audited 13 Aug and the sheet came back
clean.** Kept because the mechanism is still live: each failed run writes a blank
or part-filled row and sends a completion mail, and `create` teaches the registry
whatever it was given. So any future Shortcut work owes another audit. A stray
blank row is indistinguishable from a real deferred entry — which is the whole
point of deferred entries and the reason this matters.

**Run `findDebris()` from the editor** — added 13 Aug, in `v2/Smoke.js`, and it is
the tool for Cutover step 1. It **reports and never deletes**, because a
part-filled row awaiting a document is indistinguishable from a real deferred
entry and a tool that guessed would destroy real claims. Two confidence levels:

- **`certain`** — no counterparty, or no usable amount. Both intake paths always
  set both, so a row missing either came from a run that failed partway.
- **`suspect`** — complete, but Siri-sourced, awaiting a document, no category.
  Ordinary for a genuine deferred entry, so it is a prompt to look, not a verdict.

`findDebris('2026-08-13')` narrows to the rebuild day. Rows with an unreadable
timestamp are always included rather than filtered out — those are the worst-formed
and likeliest to be debris. Registry entries with `timesUsed <= 1` are reported
and **over-report by design**: a genuine one-off supplier looks identical, and
under-reporting means junk survives cutover.

**Run live on 13 Aug, twice, and the sheet is clean — the debris question is
closed.** The first run returned all zeros and was *not interpretable*: the
report carried findings only, so "scanned forty rows, all fine" and "scanned
nothing" printed identically. `scanned` and a prose `summary` went in for
exactly that reason, and the re-run confirmed it clean for real. The first line
of the report now says which of those two you are looking at, and it is worth
keeping that in mind for any future audit tool here.

**An empty report never means "ready for cutover".** A test run that *succeeded*
writes a complete, well-formed row, and nothing about `Bolt, 8 EUR, taxi` says
whether it came from a taxi or from proving a Shortcut works. Those are silent
here by design — same reason it never deletes. Read the sheet for that; this
tool only finds rows that are malformed.

`smokeCleanup()` is no use here — it only matches rows carrying `SMOKE_MARKER` in
Notes, which it wrote itself. Nothing marks a row abandoned by a half-built
Shortcut.

Then delete by hand, or via `archiveEntry()`. A typo entered through
`+ New reason` is permanent too, because `catalog` reads the column from the
sheet — fix those in the sheet, not in code.

**If the answer turns out to be "all of them", `resetAllData()` is the other
tool** — also in `v2/Smoke.js`, added 13 Aug and harness-covered. `findDebris()`
is for a sheet holding a mix; this one is for a clean slate, and it is the
shorter road into cutover if nothing on the sheet is real yet.

It clears data rows from all four sections and from any archive sheets that
exist, trashes the documents those rows point at, empties the supplier registry,
and empties the Staging folder. **Headers, sheets, the Drive tree, the Staging
folder itself and Script Properties all survive**, so `bootstrap()` does not need
re-running and no id changes — Genius Scan keeps working, because it is pointed
at Staging by id.

It **refuses to run** unless called as `resetAllData('DELETE ALL TEST DATA')`.
There is no confirmation dialog in the editor, where the last function you picked
is one click from running again, so the safeguard is structural instead: the
destructive path can only be reached by typing the string out. Documents are
trashed **before** their rows are deleted, because the row is the only record of
which file belongs to it. It returns what it actually removed — counts per
section, registry, files, staging — plus a `warnings` list for files that were
already gone.

## First thing: establish the baseline

```
npm run v2:test          # expect 719 passing, 0 failed
npm run v2:verify        # expect "Server matches v2/ — 13 files, byte for byte"
npm run v2:siri:verify   # expect "Server matches v2-siri/ — 2 files, byte for byte"
```

If any of those disagree, find out why before changing anything.

## Setup — done

`bootstrap()` re-run, `HEALTH_PATIENTS` set and verified live, the Staging
folder adopted, and the medium question added to the three Shortcuts that have
documents. Kept below because the reasons still apply.

1. **Run `bootstrap()`** from the main project's editor. It appends the new
   `Receipt Medium` header to Work, IVA and Health. Idempotent, so it is safe to
   run and it reports what it added. Until it runs, writing that field fails with
   `Unknown column`.
2. ~~**Set `STAGING_FOLDER_ID`.**~~ **`bootstrap()` now does it** — it adopts an
   existing `Staging` folder under the root by name rather than creating a second
   one beside it, and records the id. Genius Scan is already pointed at that
   folder. An id set by hand always wins, even if it points outside the tree.
3. ~~**Set `HEALTH_PATIENTS`.**~~ **Done and verified live** — `catalog` returns
   `Jackie, Kit, Auryn, Phoenix`, `closed: true`. The sheet now stores full
   names, and the initials already in the Patient column were replaced.
4. ~~**Add the medium question to the three Shortcuts.**~~ **Done, lost, and done
   again.** Health, Work and IVA have it. Income has no documents, so there is
   nothing to ask — and `Receipt Medium` must be *absent* from its `fields`, not
   blank: unknown keys are refused outright rather than dropped, so leaving it in
   fails every income entry. `catalog` returns `receiptMedium.values`, so the list
   is fetched rather than hardcoded, and it goes in `fields` as
   `"Receipt Medium"`.

## Pick up here

**Step 11 is finished** — endpoint and all four Shortcuts, twice over for three of
them. Steps 1–4 below are struck through and kept only because what they proved is
worth knowing and the commands are worth re-running. **Start with the debris
clean-up above, then step 5.**

1. ~~**Run `siriSetup()`.**~~ **Done.** `SPREADSHEET_ID` and `SIRI_API_KEY` are
   both set on the **main** project. The key is in `.env` as `V2_SIRI_API_KEY`
   (gitignored) as well as in Script Properties.

   To **rotate** it, run **`siriRotateKey()`** from the same editor. It replaces
   the key in one go and returns the new one; the old one stops working
   immediately, so update `.env` and every Shortcut. `siriSetup()` deliberately
   will *not* rotate — silently replacing a key would break every Shortcut on the
   phone with nothing to say why — which is why rotating has its own name.
2. ~~**Authorise `v2-siri`.**~~ **Done.** Deployed at **@2**, its four scopes
   granted, and it answers. **Its URL is deliberately not written down here** —
   see the note below on why. Get it with:

   ```
   cd v2-siri && clasp --user v2 list-deployments
   ```

   and build the address as
   `https://script.google.com/macros/s/<the @2 id>/exec`. There should be exactly
   two: the permanent `@HEAD` (`/dev`) and `@2`.

   A keyless probe now returns **`{"ok":false,"error":"Not authorized."}` as
   JSON, HTTP 200** — verified. That is the whole chain working: an anonymous
   request reaches the shim, the shim resolves the library, `siriHandlePost`
   runs, and the key check refuses it. Only the key is missing.

   If it ever goes back to **403** with Drive's *"Acesso negado / Precisa de
   acesso"*, Google is refusing before `doPost` runs: either the scopes were
   revoked, or the deployment's access is not really *Anyone*. Check the latter
   in the editor under Deploy → Manage deployments rather than assume. Redeploy
   with `-i <the @2 id>` to update that URL instead of making a second one.
3. ~~**Prove the endpoint works.**~~ **Done, live, all four actions.**
   - `ping` — `propertiesVisible` all true, `spreadsheet: "HelpfulForms"`.
   - `catalog` — work's `Expense Reason` came back from the sheet, not config.
   - `resolve` — sent the mishearing `"zz siri tst"`, got back
     `"confirm": "ZZ Siri Test", "corrected": true` at **0.92**. That is the
     exact case the resolve-before-create design exists for, confirmed against
     the real registry rather than the harness's.
   - `create` — wrote Work row 5 and returned `complete: false`,
     `outstanding: ["Receipt"]`, `completionEmailed: true`. A file column and a
     `Status` field were both refused live, naming the offending column.

   To re-run any of these without putting the key in your shell history:
   ```
   KEY=$(grep -E '^V2_SIRI_API_KEY=' .env | cut -d= -f2-)
   ID=$(cd v2-siri && clasp --user v2 list-deployments | grep '@2' | awk '{print $2}')
   curl -sL "https://script.google.com/macros/s/$ID/exec" \
     -H 'Content-Type: application/json' \
     -d "{\"key\":\"$KEY\",\"action\":\"catalog\",\"section\":\"health\"}"
   ```
   Note there is **no `-X POST`** — see the curl trap below; `-d` already makes
   it a POST.
4. ~~**Build the Shortcuts.**~~ **All four working — built once, then three of
   them rebuilt on 13 Aug after the iCloud loss.** `v2/SIRI-SHORTCUT-REBUILD.md`
   is the record of that rebuild and the better document to work from: three
   tap-by-tap walkthroughs, the `SIRI-SHORTCUT.md:322` correction, a `fields`
   reference per section, and the two decisions listed above.

   Duplicate from the right one — it is the difference between four edits and
   forty: **work** and **iva** from health, **income** from **work** (it needs
   work's open-list picker, which health does not have).

   Everything section-specific that has to differ: `section` in **all three**
   requests; `Expense Reason` *with* a space for work, `Reason` *without* one for
   income; `Receipt Medium` present for health, work and iva and **absent** for
   income.

   Named to match how the phrase is actually spoken rather than how it is
   spelled — "Log Eva receipt", because that is how IVA is said. Siri matches
   sound, not spelling.

5. ~~**Click the staging picker.**~~ **Done — tested and works.** The last
   never-once-used surface in the web UI, and the one where four of five defects
   in an earlier session were client-side. It picks, files and renames.
6. **Then the older items**, still outstanding from last session and still
   needing a phone:
   - **See the NIF warning on a merge** — 9c works and was used by hand, but the
     NIF handling was tightened afterwards and its two warnings have never been
     on screen. Merge two suppliers whose NIFs **differ**; then merge into one
     with **no** NIF and confirm it says the core has *inherited* one. Matching
     NIFs must say nothing at all.
   - ~~**The document link on the phone.**~~ **Confirmed working, 13 Aug.** The
     `authuser=` parameter `uiFileUrl()` appends is enough; Drive does not need
     the `/u/N/` path form. This was the last thing daily use depended on that
     had never been tried.
   - ~~**A durable phone session.**~~ **Done — the home screen link works.**
     Reached by opening it in a **Private Browsing** tab first, getting it to the
     point where it loaded, and adding to the Home Screen from there. The private
     tab is what supplies a cookie jar holding only the v2 account; the home
     screen icon then keeps its own. **This is now the way in on the phone** —
     before it, the only routes were Siri and a link from a completion mail.
7. **Cutover — step 10, and it is now the next thing.** The plan lists five
   steps; **step 1 is done** — `findDebris()` was re-run after the denominator
   went in and came back clean, so there is no test debris to clear. The rest is
   operational, not code:

   2. **Turn off "Accepting responses" on all four old v1 forms.** In the *old*
      account. This is the actual moment of cutover and the one irreversible-in-
      practice step, because until it happens a claim can land in either system
      and you will not know which — the split-brain the plan names.
   3. New submissions go to the new account from here.
   4. **The old GitHub Pages page stays live and untouched**, pointed at the old
      account, used only to work the remaining backlog down.
   5. **When that backlog hits zero:** unpublish the old page, delete the old web
      app deployment, keep the old sheet and Drive as archive.

   **Everything code-side is ready and checked:** harness green, both projects
   byte for byte, deployment 26 current, the staging picker exercised, the phone
   document link confirmed, and the debris audit clean. Nothing is left to build.

   **Step 2 cannot be done from a terminal, and this is the whole reason cutover
   is still open.** `clasp` moves Apps Script code and nothing else — there is no
   CLI for a Form's *Accepting responses* toggle. Doing it in code would mean
   pushing a `FormApp.setAcceptingResponses(false)` function into v1, which
   "**Do not change v1**" forbids, and only one of the four form ids is even
   recorded (`FORM_ID`, the Work form, and it lives in the old account's Script
   Properties, not in this repo). So it is four toggles in the browser, signed in
   as **jackiepurplecat@gmail.com**: `forms.google.com` → each form → **Responses**
   → switch **Accepting responses** off.

   It is reversible — switch it back on — so the risk is not the toggle, it is
   the gap. Do all four in one sitting.

   **Then stop changing v1** — see "What stays frozen" in the plan. Its known
   bugs are accepted. To stop the iCloud emails without touching its code, delete
   the `ICLOUD_EMAIL` Script Property on the old account, then delete the iOS
   automation so the failure notifications stop.
8. **Validation — `VALIDATION-PLAN.md`**, written this session and not started.
   Five rules noticed in use. The one to know now: **nothing validates values
   anywhere**, on any intake path. A quoted `"Amount":"abc"` would be written to
   the sheet as text; it only failed during the Siri build because unquoted it
   broke the JSON parse first. Not urgent, but it is a hole rather than a polish
   item.

## Things that will waste your time if you do not know them

- **Do not put `-X POST` in the curl.** Apps Script answers `/exec` with a 302 to
  `script.googleusercontent.com/macros/echo`, and that endpoint serves **GET
  only**. Plain `curl -L -d …` is right: curl downgrades the redirected request
  to GET by itself. `-X POST` forces the method to stick across the redirect and
  you get **405** with a Drive "Página não encontrada" page — which reads exactly
  like a dead deployment and is not one. This cost a diagnosis already. Shortcuts
  does the right thing on its own; this is a curl problem only.
  The tell is in the redirect: `curl -s -D - -o /dev/null -X POST … | grep -i location`.
  A `location` carrying **`&lib=…`** means the library resolved and `doPost`
  already ran, whatever status the next hop returns.
- **Library code reads the LIBRARY's Script Properties — confirmed, not
  assumed.** This was the one thing the harness could not reach and the whole
  one-store design rested on it. A live `ping` returned `propertiesVisible` all
  true and `"spreadsheet": "HelpfulForms"`, so `v2/Siri.js` running as a library
  sees the **main** project's properties, and `getSpreadsheet()` resolved through
  the `openById(SPREADSHEET_ID)` fallback exactly as intended. **Nothing needs
  duplicating onto `v2-siri`, and nothing should be** — a second store is the
  drift this design exists to avoid.
- **Never write the Siri `/exec` URL into a tracked file.** This repo is public,
  and unlike every other URL in this note that one is anonymous: the address *is*
  the reachable surface. It cannot be used to write anything — an unset or wrong
  key returns `Not authorized.` — but a published anonymous endpoint attracts
  scanning and spends execution quota on traffic that has nothing to do with you.
  It happened once already: deployment `@1` went into this file and was pushed,
  so it was **deleted and redeployed as `@2`**, and the published address now
  404s. Rotating cost nothing that day because no Shortcut existed yet; once
  four Shortcuts point at it, rotating means editing all four on the phone. Get
  the URL from `clasp --user v2 list-deployments`, keep it in the Shortcut and in
  `.env` if anywhere.
- **`clasp push` reports success while pushing nothing.** Always verify after.
  This bit again while creating `v2-siri`: the first push printed
  `Skipping push.` because the remote manifest was still Google's default.
  `npm run v2:siri:push:force` fixed it. Both projects have their own verify.
- **`v2-siri/appsscript.json` is generated and gitignored.** It holds the main
  project's script id, which is why it is not committed. `npm run v2:siri:push`
  regenerates it first, so it is hard to get wrong — but on a fresh clone it does
  not exist until `v2/.clasp.json` does. Written with **no trailing newline**, for
  the same one-byte reason as `v2/appsscript.json`.
- **`v2/appsscript.json` has no trailing newline, on purpose.** `wc -c` should be
  **425**, not 426. Do not tidy it.
- **Pushing is not deploying.** A push updates HEAD, which `/dev` serves. `/exec`
  serves a pinned version, so cut a new one — and pass `-i <deploymentId>` or
  clasp creates a *second* deployment on a different URL:
  ```
  cd v2 && clasp --user v2 deploy -i AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo -d "what changed"
  ```
  **The main deployment is version 27 and is current** — checked 13 Aug. The only
  `v2/` source to change since is `Smoke.js`, which is editor-only tooling and
  reaches no user-facing surface, so `/exec` and the tree agree on everything the
  web UI runs. Nothing to redeploy before cutover.

  Re-check it the same way rather than trusting this line: `clasp --user v2
  list-deployments` for the pinned version, then `git diff --stat <the commit
  that version was cut from>..HEAD -- v2/`. If only `Smoke.js` and docs move, the
  deployment is still current.
- **A clean reload fixes a dead file picker on iOS.** iOS suspends the sandboxed
  iframe, killing the user-activation context the picker needs while the page
  still looks fine. Close the tab and reopen. **Do not write defensive code for
  this** without reproducing it. Nothing is lost — uploads land before anything is
  written.
- **Diagnose access failures from `appsscript.json`, not from the error text.**
  - *"You need access"* = right file, wrong account.
  - *"Cannot open the file"* = no rights to the script itself.
- **The web app, version 27** — the home page. This is the address to open, to
  bookmark and to add to the home screen. **Always with `authuser`**, or whichever
  Google account the device happens to default to answers and the failure reads
  as a missing file rather than the wrong identity:
  ```
  https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec?authuser=purplecat.admin@gmail.com
  ```
  Safe to write down, unlike the Siri `/exec`: it is **not** anonymous. A signed-out
  request 302s to `accounts.google.com` — verified 13 Aug — and every function the
  page can call checks `Session.getActiveUser()` besides. On the desktop it opens
  as normal; on the iPhone it has been used in a Safari **Private Browsing** tab
  signed in as the v2 account.
- **`authuser=` does NOT switch accounts, and this cost time on 13 Aug.** It
  *selects* among accounts **already signed in to that browser**. If the v2
  account is not signed in there, there is nothing for it to select, the default
  answers, and the page refuses. So the address alone can never override a Safari
  logged in as `jackiepurplecat` — no URL can.

  **`/u/N/` does not work either.** The Drive trick
  `drive.google.com/u/1/file/d/…` has no equivalent for a web app:
  `script.google.com/u/1/macros/s/<id>/exec` returns **404** for every N. Tested
  0, 1 and 2 on 13 Aug — do not spend the hour again.

  **The symptom is Drive's access error, not ours.** *"Precisa de acesso"* /
  *"You need access"* means Google refused before the script ran — `access` is
  `MYSELF`, so a browser authenticated as `jackiepurplecat` never reaches
  `doGet`. Per the rule below: diagnose from `appsscript.json`, not the error
  text. Nothing in the code can catch this, which is why the version 27 denial
  page does not help here.

  **So the browser context itself has to be the v2 account.** That is the whole
  problem, and it is why the Private Browsing tab works — it is a jar with only
  v2 in it. Two durable ways to get the same thing:

  1. **Add the v2 account to normal Safari** — Google account menu → *Add another
     account* → the v2 address. `authuser=` then has something to resolve to.
     Both accounts are needed anyway while the v1 backlog is worked down, so
     signing `jackiepurplecat` out is the wrong move.
  2. **Give it a jar that only knows v2** — Add to Home Screen, or an iOS 17
     Safari Profile. Stronger, because it removes the ambiguity instead of
     resolving it every time.
- **Add to Home Screen is the durable phone session**, and the reason it works is
  that iOS gives a home-screen web app **its own cookie jar**, separate from
  Safari's. That is what stops the default account from winning, and it is a
  stronger fix than `authuser=` because it removes the ambiguity rather than
  resolving it.

  Open the link above in Safari, sign in as the v2 account, confirm the entry list
  actually renders, *then* Share → **Add to Home Screen**. Adding it before
  signing in saves a logged-out page.

  Two things to watch, neither yet seen here:
  - **The separate cookie jar cuts both ways** — the icon may open logged out the
    first time even though Safari is signed in. That is the behaviour working as
    intended: sign in **as the v2 account only** inside it, and that jar then
    never knows any other account, which is the whole point.
  - **Apps Script serves through a redirect to `googleusercontent.com`.** If a
    standalone web app follows that out to Safari, the home-screen session is lost
    and it will keep asking. If that happens, the fallback is the Private Browsing
    tab, which is what has been used until now and works.

  **iOS 17 Safari Profiles** is the other clean answer if the home-screen jar
  misbehaves: a profile gets its own cookies, so one signed in only as v2 gives
  the same isolation without depending on standalone mode.
- **A refused visitor now gets a page that says why** — version 27. It names the
  account you are signed in as and offers **Switch Google account**, returning to
  the pinned `/exec`. It still never names the address that *is* allowed, for the
  reason `requireUiAccess()` does not.

  Before 27 the page was a bare `Not authorized.` with no viewport tag, so on a
  phone it was both uninformative and rendered at desktop width.

  **But you will almost never see it, and it does not fix the wrong-account
  case.** `v2/appsscript.json` sets `"access": "MYSELF"`, so Google refuses any
  other account **before `doGet` runs** and serves *its* page — Drive's
  *"Precisa de acesso"* — not ours. The version 27 page is only reachable by an
  account Google lets execute that then fails the `UI_ALLOWED_EMAILS` check, which
  under `MYSELF` is essentially nobody. It is correct defensive code and the
  harness pins it; it is not the answer to a wrong login. Confirmed on the phone
  13 Aug: the wrong account gets Drive's error, not ours.
- **The harness still cannot click.** It now covers the Siri endpoint end to end
  — the gate, the whitelist, injection, all four sections — but the phone, the
  deployment and the library resolution are all outside it.
- **A stray `Folha1` tab** may still be in the spreadsheet. Harmless; delete by
  hand if empty.

## Settled this session, so do not re-litigate

- **The Siri logic lives in the MAIN project, not the shim.** The shim is two
  files and holds no logic, no secrets and no configuration. A copy-the-source
  second project was considered and rejected: it would need its own Script
  Properties, and nothing would catch the two stores drifting. Reasons in the
  plan under "How step 11 resolved the questions this plan left open".
- **The library is in development mode**, so the shim runs main's HEAD. No
  version to bump, nothing to go stale.
- **`resolve` before `create`.** The confirmation happens before anything is
  written. Canonicalising server-side and reporting it afterwards was the other
  branch the plan offered, and it loses — by the time you read it the row exists.
- **`create` never re-runs the fuzzy matcher.** Exact match only, to fill blank
  `Type`/NIF. Re-matching a name already confirmed on the phone could merge a
  supplier the confirmation had just said was different.
- **An unset `SIRI_API_KEY` shuts the endpoint.** Not "no key required".
- **A field outside the whitelist is refused, not dropped.** No document column
  is accepted from Siri at all — `extractFileId` would take a Drive id out of any
  string and the script runs as you.
- **A corrected NIF is never backdated.** Old records stay as they are. Pinned by
  the harness.
- **The completion mail lands in the inbox.** The HTML body with a real `href`
  fixed the junk filing.
- **On a supplier merge the target's spelling survives**, the NIF defaults to the
  **core** entry's, and the registry does not move until every row carries the
  new name.
