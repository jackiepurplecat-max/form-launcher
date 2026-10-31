# Handover — 31 Oct 2026

Rewrite this from scratch at the end of every session: three parts, 40 lines max.
Permanent traps live in `CLAUDE.md`; decisions in `REBUILD-PLAN.md`.

## What's live

- **Main app: version 34**, deployed 31 Oct — step 13, field validation
  (`VALIDATION-PLAN.md`). Server matched `v2/` after commit 8cf6f69. First real
  run of `npm run v2:deploy`: it worked.
- **Siri: @2**, unchanged — but it runs the main project in development mode,
  so it already has step 13 (bad fields dropped and named in `dropped`).
- `/exec` (the app): `https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec`
- `/dev` (HEAD, owner only): `https://script.google.com/macros/s/AKfycbx7eadFkUKax9S5TqYzrI2gRvd3PGhhDPjQf1Tho-0/dev`

Baseline before changing anything: `npm run v2:test` (898 pass), `npm run v2:verify`,
`npm run v2:siri:verify`.

## What's next

1. **Try step 13 on the phone and Mac**: a bad Amount refused with the form
   kept open; `FT 1` saved as `FT1` with a note; a Siri entry with a bad
   amount saved with the amount blank.
2. **Cutover — on hold.** The owner is checking v1 for claims still in progress.
   Steps are in `REBUILD-PLAN.md` → *Cutover*. Don't start it until they say.
3. `.scratch/field-validation/issues/02-page-validation.md` — the same checks
   on the page while typing (needs-triage).

## Traps

- **The NIF rule on the Providers screen applies to every supplier**, including
  foreign ones. A Spanish tax number there would be refused. Not decided; raise
  it if it happens (likely fix: skip the check when Country isn't Portugal).
- **Test fixtures need valid NIFs now.** `123456789` and `500000000` pass the
  check digit; most round numbers do not.
- On the Mac, open HelpfulForms links in the HelpfulForms Chrome profile (or
  the Dock app). Never add a second Google account to that profile.
