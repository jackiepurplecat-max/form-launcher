# Handover — 30 Oct 2026

Rewrite this from scratch at the end of every session: three parts, 40 lines max.
Permanent traps live in `CLAUDE.md`; decisions in `REBUILD-PLAN.md`.

## What's live

- **Main app: version 33**, deployed 30 Oct (Health in Allianz's two screens;
  Country on Providers). Server matched `v2/` after commit ba735a8.
- **Siri: @2**, unchanged since 15 Aug.
- `/exec` (the app): `https://script.google.com/macros/s/AKfycbxKHouifK8w8hbpMGZ_W0yklTKCdCgp-YHAk9uS7Omji_RH_fa4Za6DGYk1ZjOL5tuo/exec`
- `/dev` (HEAD, owner only): `https://script.google.com/macros/s/AKfycbx7eadFkUKax9S5TqYzrI2gRvd3PGhhDPjQf1Tho-0/dev`

Baseline before changing anything: `npm run v2:test`, `npm run v2:verify`,
`npm run v2:siri:verify`.

## What's next

1. `.scratch/dev-routine/issues/01-upgrade-order.md` — fixed order for upgrades.
2. `.scratch/dev-routine/issues/02-mac-chrome-profile.md` — the owner creates
   the Chrome profile; the ticket covers the rest.
3. **Cutover — on hold.** The owner is checking v1 for claims still in progress
   before deciding whether it's worth a "Retire v1" ticket. Steps are in
   `REBUILD-PLAN.md` → *Cutover*. Don't start it until they say.
4. `.scratch/field-validation/issues/01-intake-validation.md` — step 13.

## Traps

- An upgrade that renames a column breaks the pinned `/exec` until the next
  deploy: run the upgrade and the deploy back to back (ticket 01 fixes this).
