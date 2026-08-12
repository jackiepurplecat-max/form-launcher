# Start here

Handover note, written 12 Aug 2026. **Operational state only** — the design and
the reasons live in `REBUILD-PLAN.md`, which is the source of truth. Read that
after this. This file is disposable: overwrite it at the end of each session.

## Where things stand

| | |
|---|---|
| Branch | `step-7-web-ui`, pushed to `origin` |
| Working tree | clean |
| Harness | **521 passing, 0 failed** |
| Server | matches `v2/` byte for byte, 12 files |
| Deployed | **version 22**, on the existing deployment id |

**Steps 1–9 of the build order are done and verified by hand.** Step **9c** is
**built but never once used in a browser** — that is the main thing outstanding.

The deferred claim now works end to end, which was the last unexercised path:
an entry made with no receipt holds its claim, the completion mail goes instead,
attaching the receipt through `Edit` sends the claim *then*, and editing the row
again sends **nothing**. The `Claim Emailed` guard holds.

## First thing: establish the baseline

```
npm run v2:test     # expect 521 passing, 0 failed
npm run v2:verify   # expect "Server matches v2/ — 12 files, byte for byte"
```

If either disagrees with those numbers, find out why before changing anything.

The web app, version 22:

```
https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec
```

**On the desktop** it opens as normal. **On the iPhone** it must be a Safari
**Private Browsing** tab signed in as the v2 account — see below.

## Pick up here

Ordered by what unblocks the most.

1. **Hand-test step 9c — supplier editing.** Built this session, 52 new
   assertions around it, never clicked. There is a **Suppliers** button in the
   header. The test worth doing is the real one: find an actual typo in the
   registry, rename it onto the correctly-spelled supplier, and confirm three
   things agree afterwards — the entry rows say the new name, the **Drive
   filenames** were rebuilt, and the two registry rows became one with `Times
   Used` summed. Then check the alias offer appears and that declining it adds
   nothing. Remember the harness cannot click: everything in `Index.html` has to
   be exercised by hand, and four of five defects in one earlier session were
   client-side.
2. **The document link on the phone.** Never once confirmed, because every
   previous attempt was stopped by the account gate that has since been
   understood. Now that a private tab works, tap a document link in the table. If
   Drive wants `drive.google.com/u/N/file/d/<id>/view` rather than the
   `authuser=` parameter `uiFileUrl()` appends, that is a one-line change with the
   harness already around it.
3. **A durable phone session.** The private tab works but Private Browsing drops
   its cookies when the tab closes, so it may mean signing in every time. Try
   **Add to Home Screen** from that session — iOS gives home-screen web apps their
   own cookie jar, which would make it permanent. A second browser app signed in
   only as v2 is the fallback. Opening `webapp.access` to `ANYONE` is the other
   route and **requires guarding the globals first** (see the plan).
4. **Cutover — step 10.** See the plan. Steps 11 (Siri, its own Apps Script
   project) and 12 (OCR intake) are new capability and can wait.

## Things that will waste your time if you do not know them

- **A clean reload fixes a dead file picker on iOS.** Locking the phone part way
  through choosing a file left the `Edit` form's file button doing nothing at all.
  Nothing in the page disables it; iOS had suspended the sandboxed iframe, so the
  user-activation context the picker needs was dead while the page still looked
  fine. Close the tab and reopen. **Do not write defensive code for this** without
  reproducing it first. Nothing was lost — uploads land before anything is
  written, so an abandoned edit changes nothing.
- **Diagnose access failures from `appsscript.json`, not from the error text.**
  `"access": "MYSELF"` means the deployment is visible to exactly one Google
  account. The phone resolves requests as its *default* account, so Google refuses
  before `doGet` runs — and an Apps Script deployment is backed by a Drive file, so
  **Drive** renders the refusal as *"Sorry, unable to open the file at this
  time."* That message mentions no accounts, so it reads as a broken link.
  - *"You need access"* = right file, wrong account.
  - *"Cannot open the file"* = no rights to the script itself.
- **`clasp push` reports success while pushing nothing.** Always
  `npm run v2:verify` after. On a mismatch, `npm run v2:push:force`.
- **`v2/appsscript.json` has no trailing newline, on purpose.** `wc -c` should be
  **425**, not 426. One byte silently refuses every push. Do not tidy it.
- **Pushing is not deploying.** A push updates HEAD, which `/dev` serves. `/exec`
  serves a pinned version, so cut a new one — and pass `-i <deploymentId>` or
  clasp creates a *second* deployment on a different URL:
  ```
  cd v2 && clasp --user v2 deploy -i AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo -d "what changed"
  ```
  `clasp --user v2 list-deployments` should show exactly two: the permanent
  `@HEAD` (`/dev`) endpoint and the one pinned version. More than that means an
  earlier deploy omitted `-i`.
- **`WEB_APP_URL` is set** in Script Properties, and the completion link needs it.
  Unset, the mail silently falls back to the spreadsheet row.
- **The harness now checks the page statically, but it still cannot click.** It
  proves the page's script parses, that every `el('id')` exists in the markup, and
  that every `ui*` name the page calls exists on the server and refuses a
  stranger. None of that is a substitute for tapping the thing.
- **A stray `Folha1` tab** may still be in the spreadsheet — Google's default sheet
  under a Portuguese locale. `bootstrap()` reports it rather than deleting it, in
  case it holds something. Harmless; delete by hand if empty.

## Settled this session, so do not re-litigate

- **The completion mail lands in the inbox.** The HTML body with a real `href`
  fixed the junk filing. No mail rule needed.
- **`Education` and `Boarding Pass`** are on Work's `Type`. The rest of that list
  is still the original proposal, and the `TODO` in `Config.js` says so.
- **On a supplier merge the target's spelling survives**, a conflicting NIF keeps
  the established value, and the registry does not move until every row carries
  the new name. Reasons in the plan under the management module.
- **Correcting a supplier's NIF does not rewrite `Emitente NIF` on past IVA
  entries.** Still genuinely open, but the build takes the option that destroys
  nothing.

## Not started

Cutover (step 10), the Siri endpoint (step 11, its own Apps Script project —
anonymous access is per project), and OCR intake (step 12). There is no `doPost`,
so the only outside surface is the signed-in UI.
