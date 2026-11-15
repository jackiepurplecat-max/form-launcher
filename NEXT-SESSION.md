# Handover — 15 Nov 2026

Rewrite this from scratch at the end of every session: three parts, 40 lines max.
Permanent traps live in `CLAUDE.md`; decisions in `docs/adr/` (older ones in `REBUILD-PLAN.md`).

## What's live

- **Main app: version 34** (step 13, field validation). Siri: @2.
- `/exec` (the app): `https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec`
- `/dev` (HEAD, owner only): `https://script.google.com/macros/s/AKfycbx7eadFkUKax9S5TqYzrI2gRvd3PGhhDPjQf1Tho-0/dev`
- **v1 retired 15 Nov** (`docs/adr/0001-retire-v1.md`) on branch `retire-v1`,
  committed, **not merged or pushed**. GitHub Pages serves `main`, so the v1
  page stays live until `retire-v1` (and the v2 branches under it) reach `main`.

Baseline: `npm run v2:test` (898 pass), `npm run v2:verify`, `npm run v2:siri:verify`.

## What's next

1. **Owner:** Google-side v1 shutdown —
   `.scratch/retire-v1/issues/01-google-side-checklist.md`.
2. **Merge to `main` and push** (owner to confirm). Then check the Pages root no
   longer shows v1 and `/favicon.png` still loads.
3. Backlog (all needs-triage): `.scratch/v1-branding/` (favicon + purple),
   `.scratch/replace-attached-file/` (make "replace" obvious),
   `.scratch/edit-receipt-state/`, `.scratch/field-validation/issues/02-…`.
4. Try step 13 on the phone and Mac (bad Amount refused; `FT 1` → `FT1`).

## Traps

- **`v2:verify` will report a mismatch** until the next push: the header
  comment in `v2/Config.js` changed (comment only). Expected — push, verify.
- **Don't turn GitHub Pages off** — it hosts the favicon v2 will use.
- **The NIF rule on the Providers screen applies to foreign suppliers too.**
  Undecided; likely fix is to skip it when Country isn't Portugal.
- **Test fixtures need valid NIFs.** `123456789` and `500000000` pass.
- On the Mac, open HelpfulForms links in the HelpfulForms Chrome profile.
