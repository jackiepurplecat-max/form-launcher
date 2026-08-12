# Start here

Handover note, written 12 Aug 2026. **Operational state only** — the design and
the reasons live in `REBUILD-PLAN.md`, which is the source of truth. Read that
after this. This file is disposable: overwrite it at the end of each session.

## Where things stand

| | |
|---|---|
| Branch | `step-7-web-ui`, pushed to `origin` |
| Last code commit | `8e7f27f` — the Siri endpoint. Doc commits follow it |
| Working tree | clean |
| Harness | **623 passing, 0 failed** (was 531) |
| Main project | matches `v2/` byte for byte, **13 files** — `Siri.js` is new |
| Siri project | **new** — `v2-siri/`, matches byte for byte, 2 files |
| Deployed | main is still **version 23**. Siri is at **@2**, authorised, and answering |

Steps 1–9 and 9c are done and verified by hand. **Step 11's server side is
finished: built, pushed, harness-covered, deployed, configured, and every
action exercised live against the real spreadsheet** — including `create`, the
whitelist refusals, and the registry learning what it created. Nothing on the
server is untested. **What is left is the four Shortcuts.** Cutover (step 10) is
still not started.

**A test row may still need deleting.** The live `create` left Work row 5,
`ZZ Siri Test` at €1.23, plus a `ZZ Siri Test` entry in the Suppliers registry —
delete both, or the registry keeps fuzzy-matching it. The Expense Reason
`DELETE ME - siri live test` disappears on its own once the row is gone, because
that list is computed from the sheet.

## First thing: establish the baseline

```
npm run v2:test          # expect 623 passing, 0 failed
npm run v2:verify        # expect "Server matches v2/ — 13 files, byte for byte"
npm run v2:siri:verify   # expect "Server matches v2-siri/ — 2 files, byte for byte"
```

If any of those disagree, find out why before changing anything.

## Pick up here — finishing Siri

Steps 1–3 are done — kept here struck through, because what they proved is worth
knowing and the commands are worth re-running. **Start at step 4.** It needs a
phone; nothing below can be done from the CLI.

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
4. **Build the Shortcuts.** `v2/SIRI-SHORTCUT.md` is the full recipe, including
   the protocol and the per-section field whitelist. Build "Log health claim"
   first and duplicate it.
5. **Then the older items**, still outstanding from last session and still
   needing a phone:
   - **See the NIF warning on a merge** — 9c works and was used by hand, but the
     NIF handling was tightened afterwards and its two warnings have never been
     on screen. Merge two suppliers whose NIFs **differ**; then merge into one
     with **no** NIF and confirm it says the core has *inherited* one. Matching
     NIFs must say nothing at all.
   - **The document link on the phone** — never once confirmed. If Drive wants
     `drive.google.com/u/N/file/d/<id>/view` rather than the `authuser=`
     parameter `uiFileUrl()` appends, that is a one-line change with the harness
     already around it.
   - **A durable phone session** — try **Add to Home Screen** from the working
     private tab; iOS gives home-screen web apps their own cookie jar.
6. **Cutover — step 10.** See the plan.

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
  **The main deployment is still version 23** and does not yet include `Siri.js`.
  That does not matter for Siri — the shim runs the main project's **HEAD**
  through a development-mode library, not the pinned version — but it does mean
  the web UI is running older code than the tree.
- **A clean reload fixes a dead file picker on iOS.** iOS suspends the sandboxed
  iframe, killing the user-activation context the picker needs while the page
  still looks fine. Close the tab and reopen. **Do not write defensive code for
  this** without reproducing it. Nothing is lost — uploads land before anything is
  written.
- **Diagnose access failures from `appsscript.json`, not from the error text.**
  - *"You need access"* = right file, wrong account.
  - *"Cannot open the file"* = no rights to the script itself.
- **The web app, version 23**, on the desktop as normal; on the iPhone a Safari
  **Private Browsing** tab signed in as the v2 account:
  ```
  https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec
  ```
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
