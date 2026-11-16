# Handover — 16 Nov 2026

Rewrite this from scratch at the end of every session: three parts, 40 lines max. Permanent traps live in `CLAUDE.md`; decisions in `docs/adr/` (older ones in `REBUILD-PLAN.md`).

## What's live

- **Main app: version 44** (commit `ec48784`). Siri: @2.
- `/exec` (the app): `https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec`
- `/dev` (HEAD, owner only): `https://script.google.com/macros/s/AKfycbx7eadFkUKax9S5TqYzrI2gRvd3PGhhDPjQf1Tho-0/dev`
- **v1 retired** (`docs/adr/0001-retire-v1.md`); Pages serves only the README and favicon.
- 16 Nov: design only, nothing deployed. **Trips** designed by grilling:
  `docs/adr/0002-…`, glossary Work terms (`d5d4d26`), spec + tickets in
  `.scratch/trip-list/`.

Baseline: `npm run v2:test` (949 pass), `v2:verify` (matched).

## What's next

1. **Build Trips**: `.scratch/trip-list/issues/01`–`08` (ready-for-agent;
   01 and 05 unblocked). Owner first sets Script Property `HOME_AIRPORT`.
2. **Owner:** check the header logo (v44); Google-side v1 shutdown
   (`.scratch/retire-v1/issues/01-…`).
3. needs-info: `.scratch/auto-archive-old-claims/` (which status is
   "submitted"?), `.scratch/siri-shortcuts-performance/`.
4. Backlog: `batch-actions`, `desktop-cards`, `field-validation/02`, `generic-supplier-names`.

## Traps

- **The phone's home-screen app is `/exec`**: deploy, then swipe it closed.
- **The harness doesn't cover `v2/Index.html`** — except the `sort-compare`
  block, which `test/run.js` extracts and runs. `npm run v2:preview` screenshots.
- **Harness can't see `const`s via `G`**, and sandbox Dates fail `instanceof`
  in run.js — use `Object.prototype.toString`.
- **Provider renames don't stamp Last Updated** on purpose (would lift up to 50 rows).
- **`follows` with `serverOnly: true` never reaches the page.** Keep Invoice Date's rule.
- **Don't turn GitHub Pages off** — it hosts favicon + logo (URL in Web.js and Index.html).
- **Home-screen icon / theme-color can't be set** (Google's wrapper owns them).
- **NIF rule on Providers applies to foreign suppliers too.** Undecided.
- **Test fixtures need valid NIFs.** `123456789` and `500000000` pass.
- **Home airport is private** (public repo): Script Property only, never Config.js.
