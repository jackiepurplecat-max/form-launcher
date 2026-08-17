# Two Systems: v1 (Frozen) & v2 (Active)

**Work in v2.** Most of what sits in this repo root is v1, which is frozen —
`Code.js`, `index.html`, `README.md`, `appsscript.json`. Several things that are
true of v1 are **deliberately false** of v2, so do not copy patterns out of the
v1 files: they use column indices, `onFormSubmit` triggers, and headers typed
into sheets, all of which v2 rejects. Do not change v1; its known bugs are
accepted, not fixed.

| | v1 (frozen) | v2 (active) |
|---|---|---|
| Code | `Code.js`, `index.html` | `v2/` |
| Account | Original (default credential) | Separate — `V2_CLASP_ACCOUNT` in `.env` |
| Push | `npm run clasp:push` | `npm run v2:push` (`--user v2` credential) |
| Intake | Google Forms + `onFormSubmit` | `createEntry()` only — no Forms, no triggers |
| UI | GitHub Pages + Sheets API | Apps Script web app (`v2/Web.js` + `v2/Index.html`) |

---

## v2 — Active Development

### Start Here

1. **Read `NEXT-SESSION.md`** — what is deployed, what is verified, what to pick
   up, and the traps. It is disposable; overwrite it at the end of each session.
2. **Read `REBUILD-PLAN.md`** — the source of truth: architecture, decisions and
   the reasons behind them, build order, and a state-of-play section saying
   exactly what is done.
3. **Establish the baseline:** `npm run v2:test` (and `npm run v2:check` for a
   syntax-only pass).

Then, after every change:

```bash
npm run v2:test        # harness — before AND after any change
npm run v2:check       # syntax only
npm run v2:push        # then ALWAYS verify — see below
npm run v2:verify      # pulls and diffs: does the server actually hold this?
npm run v2:push:force  # only when verify reports a mismatch
npm run v2:whoami      # confirm you are on the v2 account
```

`npm run v2:test` runs the real `v2/` source against stand-ins for Apps Script's
services (`v2/test/mocks.js`). Run it **after** changes too, not just before —
it catches things review does not, and several of the bugs it now guards against
were found that way rather than by reading the code. A failure may be a mock gap
rather than a code bug.

### v2 Authentication

`clasp login` writes a single **global** `~/.clasprc.json`, so one default
credential meant whichever account you logged in as last was the one both
projects pushed to. v2 therefore uses a *named* credential:

```bash
npm run v2:login    # once — writes the named `v2` credential
npm run v2:whoami   # confirm which account v2 is talking to
```

Never run the bare `clasp:*` scripts while working in v2 — they use the default
credential and target the frozen v1 project. In particular `npm run clasp:pull`
overwrites the tracked root `Code.js` and `appsscript.json` from the v1 remote.
The v2 equivalents are `v2:pull`, `v2:open`, `v2:status`.

### Critical v2 Rules

**A push that reports success may have done nothing.** When the remote
`appsscript.json` differs from the local one at all, clasp asks before
overwriting it; with no TTY the prompt defaults to no and the push is abandoned.
It then prints either `Skipping push.` — which at least reads as a refusal — or
**`Pushed 9 files.` followed by the full list, having sent none.** The output
cannot be trusted either way. Never take it as evidence: run
`npm run v2:verify`, and on a mismatch `npm run v2:push:force`, which ends the
loop (once the server holds the local manifest byte for byte, plain pushes stop
being refused).

**`v2/appsscript.json` has no trailing final newline, on purpose.** Google stores
it that way, and one byte of difference silently refuses every push. It should be
**425 bytes, not 426** (`wc -c v2/appsscript.json`). Git showing
`\ No newline at end of file` for that file is correct — do not tidy it. This is
the only file in the repo with that exception.

**Pushing is not deploying.** A push updates HEAD, which is what `…/dev` serves;
`…/exec` serves a pinned version until you cut a new one with
`clasp --user v2 deploy -i <deploymentId>`. Omitting `-i` creates a *second*
deployment on a different URL and leaves the old one live — which is how you end
up debugging a page that no longer exists.

**Assume several Google accounts are signed in, on every device including the
phone, and that the v2 account is not the default.** Any Google URL this project
emits — document links, completion links, anything for the phone — must carry
`authuser=<address>`, or the default account answers and **the failure reads as a
missing file rather than the wrong identity**. Note the limit: `authuser=`
only *selects* among accounts already signed in to that browser; it cannot switch
to one that is not. `/u/N/` does not work at all. See `NEXT-SESSION.md` for the
tested details before spending time here.

### True of v2, deliberately false of v1

- **Columns are resolved by header name, never by index.** No magic numbers.
- **Headers are generated from `SECTIONS`** (`v2/Config.js`), so adding a field is
  a config change plus a re-run of `bootstrap()`. **Never type a header into a
  sheet.**
- **Everything section-specific lives in `v2/Config.js`.** An
  `if (section === 'health')` anywhere else is a bug in the wrong place.
- **`createEntry()` is the only way a row is born.** There is no trigger.
- **Never report success for work that failed.** Operations return what actually
  happened; a status change that renames nothing says so.
- **Every value written to a sheet goes through `safeCellValue`.**
- **Claim mail sends once**, when the document is there, and at no other time —
  guarded by a `Claim Emailed` column, not by convention.
- **Every function the web page *can* call checks the caller itself.**
  `google.script.run` reaches any global in the project, so the deployment
  setting is never the only gate. The check reads `Session.getActiveUser()` — the
  visitor — not `getEffectiveUser()`, which under "execute as me" is always you
  and therefore proves nothing.

### Where things are documented

- **`NEXT-SESSION.md`** (repo root) — disposable handover: deployed state,
  verified state, traps.
- **`REBUILD-PLAN.md`** (repo root) — source of truth: architecture, decisions,
  build order, state of play.
- **`v2/`** — the code.
- **`v2/Smoke.js`** — the live counterpart of the harness, run from the Apps
  Script editor; it *is* pushed.
- **`v2/test/`** — local only, never pushed.

Handoffs go in `NEXT-SESSION.md` — overwrite it at the end of each session.

---

## v1 — Reference Only

v1 is frozen and being worked down to zero.

- Google Apps Script–based expense tracking
- Integrates Google Forms, Sheets, automated file management, email notifications
- Workflow: Form submission → Sheet update → File rename → Email with attachment
- UI: GitHub Pages displaying sheet data via Sheets API
- Trigger: `handleTravel()` on form submission (Google Sheet edit)

For v1 setup and configuration see `README.md`. For behaviour, read `Code.js` —
note that it resolves columns by bare index, and the same index means different
things in different sheets (`rowValues[8]` is the email-sent flag, a file URL,
and an amount in three separate handlers), so read the surrounding comments
rather than assuming.

**`index.html` is generated.** `npm run build` injects `.env` values into
`index.template.html` and **overwrites `index.html`**. Confusingly, `index.html`
is the file tracked in git while `index.template.html` is gitignored — so editing
`index.html` directly gets your work destroyed by the next build, and the real
source is invisible to repo-wide greps. Edit the template.

### v1 commands (default credential — v1 only)

```bash
npm install
npm run clasp:login   # default credential — NOT the v2 account
npm run clasp:pull    # overwrites root Code.js from the v1 remote
npm run clasp:open
```

---

## Environment Configuration

Secrets live in `.env` at the repo root (gitignored) — see `README.md` for the
full key table. v2 additionally needs `V2_CLASP_ACCOUNT`, plus the recipient and
Siri keys documented in `REBUILD-PLAN.md`.

Apps Script secrets live in Script Properties, set once via
`setupScriptProperties()` — then clear `SCRIPT_PROPERTY_VALUES` so the secrets are
not left in the script body. Check with `checkScriptProperties()`.
