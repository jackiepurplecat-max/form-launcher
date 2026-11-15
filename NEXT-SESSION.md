# Handover — 15 Nov 2026

Rewrite this from scratch at the end of every session: three parts, 40 lines max. Permanent traps live in `CLAUDE.md`; decisions in `docs/adr/` (older ones in `REBUILD-PLAN.md`).

## What's live

- **Main app: version 41** (commit `4638756`). Siri: @2.
- `/exec` (the app): `https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec`
- `/dev` (HEAD, owner only): `https://script.google.com/macros/s/AKfycbx7eadFkUKax9S5TqYzrI2gRvd3PGhhDPjQf1Tho-0/dev`
- **v1 retired** (`docs/adr/0001-retire-v1.md`); Pages serves only the README and favicon.
- New this session: sort bar on every section (`.scratch/list-sorting/` #01).
  `Last Updated` column stamped by `touchEntry()`; `applyLastUpdatedUpgrade()`
  was run by the owner. Not yet checked by the owner in the app.

Baseline: `npm run v2:test` (947 pass), `v2:verify` (matched).

## What's next

1. **Owner:** check the sort bar on the phone and Mac (HelpfulForms profile);
   then close `.scratch/list-sorting/` #01.
2. **Owner:** Google-side v1 shutdown:
   `.scratch/retire-v1/issues/01-google-side-checklist.md`.
3. Backlog (needs-triage): `.scratch/v1-branding/`, `.scratch/desktop-cards/`,
   `.scratch/field-validation/issues/02-…`, `.scratch/generic-supplier-names/`.
4. Try step 13 on the phone and Mac (bad Amount refused; `FT 1` → `FT1`).

## Traps

- **The phone's home-screen app is `/exec`**: deploy, then swipe it closed.
- **The harness doesn't cover `v2/Index.html`** — except the `sort-compare`
  block, which `test/run.js` extracts and runs. `npm run v2:preview` screenshots.
- **Harness can't see `const`s via `G`**, and sandbox Dates fail `instanceof`
  in run.js — use `Object.prototype.toString`.
- **Provider renames don't stamp Last Updated** on purpose (would lift up to 50 rows).
- **`follows` with `serverOnly: true` never reaches the page.** Keep Invoice Date's rule.
- **Don't turn GitHub Pages off.** It hosts the favicon v2 will use.
- **NIF rule on Providers applies to foreign suppliers too.** Undecided.
- **Test fixtures need valid NIFs.** `123456789` and `500000000` pass.
