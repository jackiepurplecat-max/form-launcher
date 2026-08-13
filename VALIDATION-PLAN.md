# Validation — planned, not built

Rules noticed in use, collected here rather than fixed one at a time. **None of
this is implemented.** `REBUILD-PLAN.md` stays the source of truth for
architecture; this is a single feature spec that will fold into it once built.

Written 13 Aug 2026, prompted by a Siri entry: typing letters into the amount
prompt produced `"Amount":abc`, which failed as malformed JSON. That was luck.
**Quoted, `"Amount":"abc"` would have been accepted and written to the sheet**,
because the Siri whitelist checks column *names* and nothing checks values.
The same hole is open on the web form and the edit dialog.

## The rules

| # | Field | Rule | On failure |
|---|---|---|---|
| 1 | `Amount`, `IVA Amount` | Must be a number | **Reject** |
| 2 | `Amount`, `IVA Amount` | Rounded to 2 decimal places | **Normalise** |
| 3 | `Emitente NIF`, supplier NIF | Exactly 9 digits, numeric only | **Reject** |
| 4 | IVA section | `IVA Amount` < `Amount` | **Reject** |
| 5 | `Número` | Alphanumeric and `/` only | **Normalise** — strip the rest |

## Where it has to live

**One declaration, in `Config.js`, next to the field.** Everything
section-specific lives there; a rule written into `Form.js` or `Siri.js` is a
rule the other intake paths do not have. `extraFields` entries already carry
`type` and `required`, so a `validate` key belongs beside them.

**Enforced inside `createEntry()`**, which is the only way a row is born — so
the web form and Siri are both covered by construction and neither can skip it.
The **edit path in `Manage.js` needs the same call**: an entry can be created
blank and completed later, so the completion step is where most IVA fields are
actually typed, and it currently validates nothing.

**Also `Suppliers.js`** for rule 3 — the registry holds NIFs, they prefill
future entries, and a bad one there propagates silently.

**The page should apply the same declarations client-side**, for feedback at the
moment of typing rather than after a round trip. Read from the same config, so
the two cannot disagree. The server check is the real one; the client check is
the courtesy.

## Reject versus normalise — keep them distinct

Rules 2 and 5 **change what was submitted**. That is not the same act as
refusing it, and this project already has a position on silent changes: a
merged NIF defaults but *says so*, precisely so the warning keeps its meaning.

So a normalising rule must **report what it changed**, and the report must reach
wherever the caller can see it — the form's response, and Siri's JSON. A
`Número` quietly stripped of characters the user typed is a value that no longer
matches the paper in their hand.

## Decisions to make before writing any of this

- **Blank must stay legal.** Validation runs on *supplied, non-empty* values
  only. Partial entries are the safety net — an entry with no amount yet is the
  whole point of the deferred model, and a validator that rejects blanks would
  break Siri, the completion mail, and scan-later in one go. This is the single
  easiest thing to get wrong here.
- **Rounding changes a claimed figure.** `12.345 → 12.35` on an IVA entry alters
  what gets submitted to Finanças. Probably right, but decide it deliberately:
  round, or reject anything with more than 2 decimals and make the user retype?
  Rejecting is more honest for the IVA section specifically.
- **Rule 4 is `<`, not `≤`.** `Amount` is Valor Total and `IVA Amount` is the tax
  within it, so equal means the whole invoice was tax. Confirm that is always
  true before pinning it — a zero-rated or fully-exempt edge case would make
  this rule reject a legitimate entry.
- **Should the NIF checksum be validated too?** A Portuguese NIF has a mod-11
  check digit. Nine digits catches a dropped character; the checksum also
  catches a **transposition**, which is the error people actually make and the
  one that produces a plausible-looking wrong number. Worth the extra few lines.
- **What happens to existing rows that already violate a rule?** If validation
  runs on edit, a legacy row with an 8-digit NIF blocks every unrelated change to
  that row until it is fixed. Options: validate only fields being changed, or
  warn rather than reject on edit. Consistent with the standing rule that
  **history is never rewritten** — see the corrected-NIF decision.
- **Currency is unvalidated** and is not on this list. Note it, decide later.

## Testing

Each rule gets harness coverage in `v2/test/run.js`, and each needs the
**negative** case as well as the positive one — that a bad value is refused, and
that a blank is still allowed through. The blank case is the regression that
would otherwise be found in a waiting room.

Rule 5 additionally needs a **round-trip** test: the stripped value is what gets
written *and* what gets reported back, or the caller is told something untrue.
