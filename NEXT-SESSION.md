# Handover — 15 Nov 2026

Rewrite this from scratch at the end of every session: three parts, 40 lines max.
Permanent traps live in `CLAUDE.md`; decisions in `docs/adr/` (older ones in `REBUILD-PLAN.md`).

## What's live

- **Main app: version 36** (commit `879a534`). Siri: @2.
- `/exec` (the app): `https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec`
- `/dev` (HEAD, owner only): `https://script.google.com/macros/s/AKfycbx7eadFkUKax9S5TqYzrI2gRvd3PGhhDPjQf1Tho-0/dev`
- **v1 retired** (`docs/adr/0001-retire-v1.md`); Pages serves only the README and favicon.
- Closed this session: `.scratch/invoice-date-follows/` #01. On a new Health
  claim, Treatment date copies Invoice date (one way, page only, via
  `dateFollows` in `Config.js`). Invoice Date's own rule is `serverOnly`: it
  fills a blank at save so Siri claims stay complete. Owner confirmed it works.

Baseline: `npm run v2:test` (901 pass), `v2:verify` (matched), `v2:siri:verify`.

## What's next

1. **Owner:** Google-side v1 shutdown:
   `.scratch/retire-v1/issues/01-google-side-checklist.md`.
2. `.scratch/awaiting-follows-documents/` (needs-triage): both docs picked →
   "Waiting for another document?" = No. Same `applyFollows()` territory.
3. Backlog (needs-triage): `.scratch/v1-branding/` (favicon + purple),
   `.scratch/replace-attached-file/` (make "replace" obvious),
   `.scratch/desktop-cards/`, `.scratch/field-validation/issues/02-…`,
   `.scratch/generic-supplier-names/` ("Farmácia X" corrected to "Farmácia").
4. Try step 13 on the phone and Mac (bad Amount refused; `FT 1` → `FT1`).

## Traps

- **The harness doesn't cover `v2/Index.html`:** confirm page fixes in the app.
- **`follows` with `serverOnly: true` never reaches the page** (`uiFormFields`
  drops it). Don't remove Invoice Date's rule: Siri sends only Date.
- **Don't turn GitHub Pages off.** It hosts the favicon v2 will use.
- **The NIF rule on the Providers screen applies to foreign suppliers too.**
  Undecided; likely fix is to skip it when Country isn't Portugal.
- **Test fixtures need valid NIFs.** `123456789` and `500000000` pass.
- On the Mac, open HelpfulForms links in the HelpfulForms Chrome profile.
