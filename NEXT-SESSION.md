# Handover — 15 Nov 2026

Rewrite this from scratch at the end of every session: three parts, 40 lines max. Permanent traps live in `CLAUDE.md`; decisions in `docs/adr/` (older ones in `REBUILD-PLAN.md`).

## What's live

- **Main app: version 44** (commit `ec48784`). Siri: @2.
- `/exec` (the app): `https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec`
- `/dev` (HEAD, owner only): `https://script.google.com/macros/s/AKfycbx7eadFkUKax9S5TqYzrI2gRvd3PGhhDPjQf1Tho-0/dev`
- **v1 retired** (`docs/adr/0001-retire-v1.md`); Pages serves only the README and favicon.
- This session: **favicon** on both v2 pages (`UI_FAVICON_URL`, `v2/Web.js`),
  served from Pages, shrunk to 256px (v42); **purple accent** `#6c3fcf` (v43,
  owner confirmed on phone; `.scratch/v1-branding/` closed); **header logo** (v44,
  not yet seen by the owner).

Baseline: `npm run v2:test` (949 pass), `v2:verify` (matched).

## What's next

1. **Owner:** check the header logo on the phone and in the Chrome profile.
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
- **Don't turn GitHub Pages off** — it hosts favicon + logo (URL in Web.js and Index.html).
- **Home-screen icon / theme-color can't be set**: Google's wrapper page owns
  them. A Shortcuts icon opens plain Safari — back to the account-order trap.
- **NIF rule on Providers applies to foreign suppliers too.** Undecided.
- **Test fixtures need valid NIFs.** `123456789` and `500000000` pass.
