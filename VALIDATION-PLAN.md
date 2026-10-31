# Validation — decided, not built

Rules for field values at intake (build step 13). **None of this is implemented.**
`REBUILD-PLAN.md` stays the source of truth for architecture; this is a single
feature spec that will fold into it once built. Decisions settled 31 Oct 2026.

Written 13 Aug 2026, prompted by a Siri entry: typing letters into the amount
prompt produced `"Amount":abc`, which failed as malformed JSON. That was luck.
**Quoted, `"Amount":"abc"` would have been accepted and written to the sheet**,
because the Siri whitelist checks column *names* and nothing checks values.
The same hole is open on the web form and the edit dialog.

## The rules

| # | Field | Rule | On failure |
|---|---|---|---|
| 1 | `Amount`, `IVA Amount` | A plain number; a comma is read as the decimal point | **Reject** |
| 2 | `Amount`, `IVA Amount` | At most 2 decimal places | **Reject** — never round |
| 3 | `Amount`, `IVA Amount` | Not negative, in every section | **Reject** |
| 4 | IVA section | `IVA Amount` **<** `Amount`, strictly | **Reject** |
| 5 | `Emitente NIF`, supplier-list NIF | Nine digits **and** a valid mod-11 check digit | **Reject** |
| 6 | `Número` | Only `A–Z a–z 0–9 / \ _ . -` | **Normalise** — see below |

**Why each one:**

- **1, comma.** `12,50` is the same number written the Portuguese way, so it is
  accepted and *not* reported — the value has not changed. `abc`, `12abc` and
  `12.50€` are rejected.
- **2, no rounding.** Real invoices never have three decimals, so `12.345` is a
  typo, and rounding it would alter a figure claimed from Finanças. Rejecting
  also means amounts are never normalised, so they never need a change report.
- **3, no negatives.** AT's `Decimal_15` is `minInclusive=0` and the export
  already refuses them. A refund is not recorded as a negative entry.
- **4, strict.** The owner's decision, and it applies **everywhere — the export
  too**, which today rejects only `IVA > total` and so allows equality.
  `ivaLineProblems` in `v2/IvaExport.js` changes to match. Consequence, accepted:
  a €0 invoice with €0 IVA is refused.
- **5, check digit.** Nine digits catches a dropped character; the check digit
  also catches a **transposition**, which is the mistake people actually make.
  Only the IVA section has `registryNifField`, so a foreign provider's tax number
  is never subjected to it.
- **6, Finanças' own rule.** The published `Simplex130.xsd` says only
  `xs:string`, but `DRORIVA.app` validates `numeroFatura` against
  `[/\\a-zA-Z0-9\_\.\-]+` (in `pt.at.simplex130.validator.Quadro02Validator`,
  read from `DRORIVA-UI-1.2.34.0003-app.jar` on 31 Oct 2026). The app refuses
  anything else, so the entry follows that rule **even though the result no
  longer matches the paper exactly**.

## Número normalisation

- **Spaces are removed**: `FT 2024/123` → `FT2024/123`. Not replaced with `-`,
  which would add a character the paper does not have.
- **Accents are dropped**: `Ç` → `C`, `É` → `E`.
- **Anything else outside the set is removed**: `#`, `º`, brackets and so on.
- **If nothing is left, reject.** Someone typed something; turning it into a
  blank would make a mistake look like "not filled in yet".

A normalising rule changes what was submitted, which is not the same act as
refusing it — and this project's position on silent changes is that they say so.
So **every normalisation is reported** to whoever submitted it: in the web form's
confirmation (*"Número saved as FT2024/123 — Finanças doesn't accept spaces"*) and
in Siri's JSON reply. Having the Shortcut *speak* it is a separate, optional
Shortcut edit, not part of step 13.

## Behaviour

- **Blank is always legal.** Validation runs on *supplied, non-empty* values
  only. Partial entries are the safety net — an entry with no amount yet is the
  whole point of the deferred model, and a validator that rejects blanks would
  break Siri, the completion mail, and scan-later in one go. This is the single
  easiest thing to get wrong here.
- **Web form and edit dialog: refuse outright.** The user is there to fix it, and
  what they typed stays on screen.
- **Siri: save the entry, drop the bad field.** Refusing would lose the whole
  spoken entry. Everything valid is saved, each failing field is left blank
  (blank is legal, so it is completed later like any partial entry), and the
  reply names each dropped field and why.
- **Edits check only the fields being changed.** A legacy row with an 8-digit NIF
  must not block an unrelated edit, and nothing rewrites the old value —
  consistent with **history is never rewritten** (see the corrected-NIF decision).
- **Supplier-list values are checked like typed ones.** Siri prefills from the
  registry before `createEntry()` (`v2/Siri.js`), so an old bad NIF there is
  dropped with a note each time; on the web form it is refused on save. Fixing it
  once on the Providers screen — which checks it too — ends it.
- **The export names and skips** an IVA row whose `Número` breaks rule 6, as it
  does for every other problem; it never cleans the value itself. Re-saving the
  Número in that row normalises it. So cleaning only happens where the user sees
  it, and the sheet holds exactly what was sent.

## Where it has to live

**One declaration, in `Config.js`, next to the field.** Everything
section-specific lives there; a rule written into `Form.js` or `Siri.js` is a
rule the other intake paths do not have. `extraFields` entries already carry
`type` and `required`, so a `validate` key belongs beside them.

**Enforced inside `createEntry()`**, which is the only way a row is born — so
the web form and Siri are both covered by construction and neither can skip it.
The **edit path in `Manage.js` needs the same call**: an entry can be created
blank and completed later, so the completion step is where most IVA fields are
actually typed.

**Also `Suppliers.js`** for rule 5 — the registry holds NIFs, they prefill
future entries, and a bad one there propagates.

**The page should apply the same declarations client-side**, for feedback at the
moment of typing rather than after a round trip. Read from the same config, so
the two cannot disagree. The server check is the real one; the client check is
the courtesy.

## Out of scope for step 13

- **Currency.** When it is done it will be a **dropdown (closed list), not free
  text** — so it becomes a `choice` field and the existing closed-list check
  covers it, rather than a new value rule. (Decided 30 Oct 2026.)
- **A report of existing rows that break the rules.** The export already names
  bad IVA rows each run and the supplier list surfaces on use. A ticket of its
  own if it is ever wanted.
- **The export's silent rounding.** `ivaCents` in `v2/IvaExport.js` rounds to
  cents and `parseFloat` reads `12abc` as 12. Intake rules 1–2 stop new bad
  values reaching it; legacy values are not this step's concern.

## Testing

Each rule gets harness coverage in `v2/test/run.js`, and each needs the
**negative** case as well as the positive one — that a bad value is refused, and
that a blank is still allowed through. The blank case is the regression that
would otherwise be found in a waiting room.

Also needed:

- **Rule 6 round-trip:** the normalised value is what gets written *and* what
  gets reported back, or the caller is told something untrue. Plus the
  all-symbols case that normalises to empty and is rejected.
- **Siri partial save:** a bad field is blank in the row, the rest is saved, and
  the reply names the dropped field.
- **Edit of a legacy row:** an unrelated change succeeds despite an old bad value.
- **Export at equality:** a row with `IVA Amount` equal to `Amount` is named and
  left out.
