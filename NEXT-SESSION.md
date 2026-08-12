# Start here

Handover note, written 11 Aug 2026. **Operational state only** — the design and
the reasons live in `REBUILD-PLAN.md`, which is the source of truth. Read that
after this. This file is disposable: overwrite it at the end of each session.

## Where things stand

| | |
|---|---|
| Branch | `step-7-web-ui` |
| Last commit | `8a72409` |
| Working tree | clean |
| Harness | **455 passing, 0 failed** |
| Server | matches `v2/` byte for byte |
| Deployed | **version 21**, on the existing deployment id |

**Steps 1–9 of the build order are done and verified by hand.** Not "written and
probably fine" — created, edited, archived, restored, purged, and one real IVA
claim end to end with its receipt, its `Emitente NIF` prefill and its filename
chain. `bootstrap()` has been re-run and all four archive sheets exist.

## First thing: establish the baseline

```
npm run v2:test     # expect 455 passing, 0 failed
npm run v2:verify   # expect "Server matches v2/ — 11 files, byte for byte"
```

If either disagrees with those numbers, find out why before changing anything.

The web app, version 21:

```
https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec
```

**On the desktop** it opens as normal. **On the iPhone** it must be a Safari
**Private Browsing** tab signed in as the v2 account — see below.

## Pick up here

Ordered by what unblocks the most.

1. **The deferred claim — the last path never exercised at all.** Make a Work or
   IVA entry with **no receipt**: the claim must be *held* (no claim mail),
   `Receipt State` = `awaiting`, and a completion email sent instead. Then attach
   the receipt through `Edit` and the claim must go out **now**, with the
   attachment, stamping `Claim Emailed`. Then edit the row again — change the
   Notes — and the claim must **not** send twice. That last step is the one worth
   watching; a double claim is the failure that actually costs something. This is
   the Siri path working before Siri exists, so everything later leans on it.
2. **Does the junk filing stop?** The completion mail went out as plain text with
   a bare 130-character URL and was being filed as junk. Version 21 sends it as
   HTML with a real `href` and a plain-text alternative, which is a less
   spam-shaped message. **Not yet observed** — step 1 will generate one, so just
   look at where it lands. If it is still junk, the fix is a mail rule, not code;
   say so rather than churning the mail body.
3. **The document link on the phone (task #5).** Never once confirmed, because
   every previous attempt was stopped by the account gate that has only just been
   understood. Now that a private tab works, tap a document link in the table. If
   Drive wants `drive.google.com/u/N/file/d/<id>/view` rather than the
   `authuser=` parameter `uiFileUrl()` appends, that is a one-line change with the
   harness already around it.
4. **A durable phone session.** The private tab works but Private Browsing drops
   its cookies when the tab closes, so it may mean signing in every time. Try
   **Add to Home Screen** from that session — iOS gives home-screen web apps their
   own cookie jar, which would make it permanent. A second browser app signed in
   only as v2 is the fallback. Opening `webapp.access` to `ANYONE` is the other
   route and **requires guarding the globals first** (see the plan).
5. **`Education` and `Boarding Pass` on Work's `Type`.** Confirmed with Jax that
   it is `Type`, not `Expense Reason` — `Expense Reason` is open free text and
   needs no change. One line at `Config.js:124`, then push. No `bootstrap()`
   re-run: the column exists, only its option list changes. While in there, the
   whole list still carries a `TODO` saying it was proposed rather than taken from
   real claims.
6. **Supplier editing with the rename propagated — step 9c.** The real build.
   Fully designed in the plan under the management module; the load-bearing
   decision is that the repair re-runs `nameAndFileDocuments()` per row rather
   than pattern-matching the old name inside the existing filename.

## Things that will waste your time if you do not know them

- **Diagnose access failures from `appsscript.json`, not from the error text.**
  `"access": "MYSELF"` means the deployment is visible to exactly one Google
  account. The phone resolves requests as its *default* account, so Google refuses
  before `doGet` runs — and an Apps Script deployment is backed by a Drive file, so
  **Drive** renders the refusal as *"Sorry, unable to open the file at this
  time."* That message mentions no accounts, so it reads as a broken link. Two
  wrong theories were argued from the symptom yesterday, and one line of the
  manifest explained the whole thing.
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
- **`WEB_APP_URL` is set** in Script Properties, and the completion link needs it.
  Unset, the mail silently falls back to the spreadsheet row.
- **Four of yesterday's five defects were client-side**, so 439 passing assertions
  could not see any of them. The harness is necessary and not sufficient: when
  something is in `Index.html`, it has to be clicked.
- **A stray `Folha1` tab** may still be in the spreadsheet — Google's default sheet
  under a Portuguese locale. `bootstrap()` reports it rather than deleting it, in
  case it holds something. Harmless; delete by hand if empty.

## Not started

Siri endpoint (step 11, its own Apps Script project — anonymous access is per
project), OCR intake (step 12), and cutover (step 10). There is no `doPost`, so
the only outside surface is the signed-in UI.
