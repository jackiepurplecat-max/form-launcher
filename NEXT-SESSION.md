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
| Harness | **617 passing, 0 failed** (was 531) |
| Main project | matches `v2/` byte for byte, **13 files** — `Siri.js` is new |
| Siri project | **new** — `v2-siri/`, matches byte for byte, 2 files |
| Deployed | main is still **version 23**. Siri is deployed at **@2**, and **returns 403 until it is authorised** — see step 2 |

Steps 1–9 and 9c are done and verified by hand. **Step 11's server side is
built, pushed, covered by the harness and deployed — but not authorised, never
answered a real request, and no Shortcut exists yet.** Cutover (step 10) is
still not started.

## First thing: establish the baseline

```
npm run v2:test          # expect 617 passing, 0 failed
npm run v2:verify        # expect "Server matches v2/ — 13 files, byte for byte"
npm run v2:siri:verify   # expect "Server matches v2-siri/ — 2 files, byte for byte"
```

If any of those disagree, find out why before changing anything.

## Pick up here — finishing Siri

All of it needs hands. Nothing below can be done from the CLI.

1. **Run `siriSetup()`** from the **main** project's editor. It sets
   `SPREADSHEET_ID` off the container and generates `SIRI_API_KEY`, and
   **returns the key once** — copy it somewhere before closing the log. A second
   run will not show it again and will not replace it.
2. **Authorise `v2-siri`.** It is already deployed, at version **@2**, described
   *Siri intake v2*. **Its URL is deliberately not written down here** — see the
   note below on why. Get it with:

   ```
   cd v2-siri && clasp --user v2 list-deployments
   ```

   and build the address as
   `https://script.google.com/macros/s/<the @2 id>/exec`. There should be exactly
   two: the permanent `@HEAD` (`/dev`) and `@2`.

   **It answers 403 to everything, GET and POST alike** — verified with curl
   against the previous deployment, so this is observed and not a guess. The body
   is Drive's *"Acesso negado / Precisa de acesso"*, which is the page you get
   when Google refuses before `doPost` runs.

   The expected cause is that a standalone project deployed execute-as-me is not
   authorised until something in it has run: nobody has granted its four scopes.
   So `npm run v2:siri:open`, run `doGet` once from the editor, accept the
   consent screen, and probe it again with the curl in step 3.

   **If it still 403s after that**, the cause is the other one: the deployment's
   access is not really *Anyone*. `appsscript.json` asks for `ANYONE_ANONYMOUS`
   and the server holds that file byte for byte, but check it in the editor
   under Deploy → Manage deployments rather than assume — the manifest and the
   deployment have disagreed before. Fixing it there and redeploying with
   `-i <the @2 id>` updates that URL rather than making a second one.
3. **Prove the library resolves** before building any Shortcut. This is the one
   thing the harness cannot test — see the trap below.
   ```
   curl -sL -X POST '<the @2 /exec url>' \
     -H 'Content-Type: application/json' \
     -d '{"key":"<the key>","action":"ping"}'
   ```
   A quicker first probe needs no key at all: send `{"action":"ping"}` and expect
   `{"ok":false,"error":"Not authorized."}`. **That JSON is the good outcome** —
   it means the deployment, the shim and the library all work and only the key is
   missing. An HTML page instead means you are still stuck on step 2.
   With the key, expect `"ok": true`, a `spreadsheet` name, and
   `propertiesVisible` **all true**. `-L` matters: Apps Script redirects.
   - `ok:false` with a `SPREADSHEET_ID` message → step 1 did not run, or
     properties resolve to the *shim's* store rather than the library's (below).
   - `Not authorized.` → the key does not match, or step 1 did not run.
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

- **The one unknown the harness cannot reach: whose Script Properties does
  library code see?** `v2/Siri.js` runs as a library when Siri calls it, and
  `getScriptProperties()` is expected to mean the **library's own** project — the
  main one, where every property already lives. If `ping` comes back with
  `propertiesVisible` all false, that expectation is wrong and the properties
  resolve to the *shim's* store instead. It is not fatal: set `SPREADSHEET_ID`,
  `SIRI_API_KEY`, `ROOT_FOLDER_ID` and both recipient addresses on `v2-siri` too.
  But it would mean two stores to keep in step, so write it down here if it
  happens. **Run `ping` before building anything on top of this.**
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
