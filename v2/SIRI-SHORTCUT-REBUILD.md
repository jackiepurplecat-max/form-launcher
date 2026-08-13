# Rebuilding the three lost Shortcuts

Written 13 Aug 2026, after `Log expense`, `Log receipt` and `Log income` were
destroyed by an iCloud reconcile. `Log health claim` survived and is the
template. Rewritten the same day as three tap-by-tap walkthroughs, because the
diff table this file started as was correct and unfollowable.

**Companion to `SIRI-SHORTCUT.md`**, which holds the protocol and — more
usefully — the six Shortcuts behaviours that turn ten minutes into an hour.
Read its *"Read this before building anything"* section first. Everything about
naming results and untyped `Get Dictionary Value` output still applies.

## What went wrong, so it does not again

**iCloud Drive was switched off on the phone.** That one setting did two separate
things, and the second is what cost the work:

1. Exporting a Shortcut to a file failed. Apple **signs** `.shortcut` files
   against a live iCloud session, so being "signed into iCloud" is not enough.
   The error claims you are not signed in, which sends you looking in the wrong
   place.
2. Shortcuts sync sat in a state where the phone's near-empty library reconciled
   **over** the Mac's four. AirDropping health to the phone appears to have been
   the trigger.

No export existed and `tmutil` reported no Time Machine destinations and no local
snapshots, so there was nothing to restore. Recently Deleted was empty on both
devices. Three Shortcuts were simply gone.

**The rules that follow:**

- **Export each Shortcut the moment it works** — Share → Save to Files. Not at
  the end of the session.
- **Those files contain `SIRI_API_KEY` and the `/exec` URL.** iCloud Drive is
  fine. This repo is public and must never hold either.
- Template backup: `iCloud Drive/Downloads/Log health claim.shortcut`.
- To get a Shortcut back into the Mac app, **double-click the file.** Importing
  needs no signing, so it works whatever state sync is in.

## ⚠️ `SIRI-SHORTCUT.md:322` is wrong

It says to delete steps **3, 4 and 5** for `Log receipt`. That predates the
`Receipt Medium` question. IVA *does* ask for medium, and its values come from
`catalog`'s `receiptMedium.values` — step 3 **is** the catalog call.

**Keep step 3. Delete only 4 and 5.** The claimed two-second saving no longer
exists.

## Which one to duplicate from

This matters more than it looks — pick wrong and you build a picker twice, or
delete one you just built.

| Building | Duplicate | Why |
|---|---|---|
| `Log expense` (`work`) | **health** | Swap health's closed picker for the open one |
| `Log receipt` (`iva`) | **health** | No category at all. Health's picker is 2 actions to delete; work's is 10 |
| `Log income` (`income`) | **work** | Needs work's open picker. Building it from health means constructing it again |

**Always duplicate, never edit the original.** Long-press → Duplicate → open the
copy → rename it.

**Before any other edit, change `section` in all three `Get Contents of URL`
actions.** Identify them by their `action` value — `catalog`, `resolve`,
`create` — not by position. Missing one fails silently in the worst way: two
calls answer for health while the row is written to the right sheet.

---

## 1. Log expense — `work`

Duplicate **health**, rename `Log expense`. Five edits and it runs; the picker is
a separate second pass.

### Part 1 — five edits

1. **`section` → `work`** in all three `Get Contents of URL` actions. Expand
   Request Body, tap the value next to `section`.
2. In **`create`** only, expand `fields`. The row keyed `Patient` becomes
   **`Expense Reason`** — *with the space*, it is the real column header. Leave
   the value chip alone.
3. `Ask for Input` → *Which supplier?*
4. `Show Alert` title → *Log expense*.
5. Run it. It should offer your existing expense reasons, ask supplier and
   amount, confirm, save. **Export.**

That is a working `Log expense`. It can only pick reasons that already exist —
Part 2 adds the way to type a new one.

### Part 2 — the `+ New reason` escape

Insert around the existing `Get Dictionary Value` / `Choose from List` near the
top:

```
Get Dictionary Value    category.values      ← already there
Set Variable            Reasons              ← ADD
Text                    + New reason         ← ADD
Set Variable            NewMarker            ← ADD
Add to Variable         Reasons              ← ADD (input is NewMarker)
Choose from List        Reasons              ← already there; input is Reasons
Text                    (just the chosen item)   ← ADD
Set Variable            ReasonChoice         ← ADD
If  ReasonChoice  is  NewMarker              ← ADD
    Ask for Input       New reason?          ← ADD
    Set Variable        Patient              ← ADD (reuse the existing name)
Otherwise
    Set Variable        Patient              ← input is ReasonChoice
End If
```

Two of those look like padding and are not:

- **The `Text` before `ReasonChoice`.** `Choose from List` output has no known
  type, so without it the `If` offers only *has any value*. A `Text` action
  always outputs text, which makes `is` appear.
- **`+ New reason` as a variable, not typed twice.** It is appended to the list
  *and* compared in the `If`. Type it in two places and one day they will not
  match, making "new reason" silently unreachable.

Run, test both paths, **export again**.

### `Type` is not asked

The registry fills it on an exact supplier match, because Uber is always a Taxi.

---

## 2. Log receipt — `iva`

Duplicate **health**, rename `Log receipt`. The easiest of the three — mostly
deletions.

**Order matters.** The `Patient` variable is going away, so remove everything
pointing at it *before* deleting it, or you are left holding orphaned chips.

1. **`section` → `iva`** in all three requests.
2. In **`create`**, delete the **`Patient`** row from `fields`. Leave
   `Counterparty`, `Amount`, `Receipt Medium`.
3. In the middle **`Show Alert`**, delete `Patient · ` from the message, leaving:

   ```
   Confirmed
   €Amount · Current Date
   ```
4. **Now** delete three actions near the top: the `Get Dictionary Value` with
   `category.values`, the *Who is it for?* `Choose from List`, and
   `Set Variable Patient`. Nothing points at them, so nothing breaks.

   *Why:* `catalog` returns `category: null` for IVA. Leave them and you get an
   empty picker.
5. **Keep the medium picker** — the `Get Dictionary Value` with
   `receiptMedium.values` and its `Choose from List`. This is what the wrong
   instruction above would have broken.
6. `Ask for Input` → *Which supplier?*; `Show Alert` title → *Log receipt*.
7. Run — it should ask supplier, amount, medium, and **no category at all**.
   **Export.**

### Every IVA entry arrives incomplete

`complete: false` with outstanding items, and that is correct. **`ok: false` is
failure; `complete: false` is not.** Número, Emitente NIF and Valor do IVA are
deliberately never asked — they are retyped into Finanças from the completion
form later, because standing at a counter is the wrong moment to read numbers off
a receipt.

---

## 3. Log income — `income`

Duplicate **`Log expense`** — it already has the open picker — and rename
`Log income`.

### The one that bites

**`Receipt Medium` must come out completely**, picker and `fields` row. Income
has no documents, so `catalog` returns `receiptMedium: null`. More importantly
`fields` accepts `Counterparty`, `Amount`, `Currency`, `Date`, `Reason` and
**nothing else**, and unknown keys are **refused outright rather than ignored** —
so leaving it in means *every* income entry fails and nothing is written.

It fails loudly. An error alert naming `Receipt Medium` is exactly this.

### The edits

1. **`section` → `income`** in all three requests.
2. In **`create`**, rename the `Expense Reason` key to **`Reason`** — *no space
   this time.* Work's column has one, income's does not.
3. In **`create`**, delete the **`Receipt Medium`** row from `fields`.
4. **Now** delete the medium picker: the `Get Dictionary Value` with
   `receiptMedium.values`, its `Choose from List`, and `Set Variable Medium`.
   Step 3 before step 4, so nothing is orphaned.
5. `Ask for Input` → *Who from?*; `Show Alert` title → *Log income*.
6. Run, test both picker paths, **export**.

### `(none)` was considered and deliberately dropped

An earlier draft added a second `(none)` marker so `Reason` could be left blank,
since it is the one category that is **not required**. It was dropped, on the
reasoning that *there is no occasion to log income without knowing what earned
it, at the moment of logging it.*

Steps 1–6 already guarantee a value: `Choose from List` always yields one
(cancelling stops the Shortcut) and `+ New reason` covers anything not yet in the
list. So requiring something needed no extra work — it needed one branch **not**
built.

**Do not flip `required: true` in `Config.js:284` to enforce this server-side.**
`Config.js:68` documents the trap for a different field: making something
required turns already-finished entries into INCOMPLETE ones and mails completion
requests for work that is done. Existing income rows with a blank Reason would
retroactively start nagging. Enforce it at capture time, in the Shortcut, which
is where the design's own argument puts it — a question that is asked cannot be
forgotten.

Note also `Config.js:287`, `registryTypeField: 'Reason'` — the registry remembers
the reason per payer, so on an exact match it is filled without asking. Blank
income reasons were already rare by design.

**If you ever do want a no-reason option**, prefer a convention over code: add
one deliberate catch-all — `Other` — and reuse it. Zero Shortcut complexity, and
a real category value groups and reports where an empty cell does not.

---

## `fields` reference

Exact keys and types. Anything not listed is **refused**, not dropped.

| Section | Keys | |
|---|---|---|
| work | `Counterparty`, `Amount`, `Expense Reason`, `Receipt Medium` | space in `Expense Reason` |
| health | `Counterparty`, `Amount`, `Patient`, `Receipt Medium` | |
| iva | `Counterparty`, `Amount`, `Receipt Medium` | no category |
| income | `Counterparty`, `Amount`, `Reason` | **no** `Receipt Medium`; no space in `Reason` |

`Currency` and `Date` may also be sent, but default to EUR and today
server-side. Do not send them.

**`Amount` must be Type `Number`.** As Text it is parsed against the
spreadsheet's Portuguese locale, where the decimal separator is a comma — no
error, just a wrong figure. Quick Look the body: you want `"Amount":70`, not
`"Amount":"70"`.

## The variable is called `Patient` in all four

Inherited from health, and it holds a reason in work and income. **Leave it.**
It is cosmetic, and renaming a variable risks orphaning the chips that reference
it. Not worth it for a label.

## Testing each one

1. **The failure path.** Temporarily set `"section": "pets"` in the *create*
   request — expect *"Unknown section: pets…"*. Change it back. Nothing is
   written either way.
2. **`Amount` is a number** — Quick Look, as above.
3. **One real entry**, then check the row landed in the right sheet with the
   right columns filled.
4. **Export to Files.**

~10 seconds of waiting across the three calls is normal — `catalog` ≈ 2s,
`resolve` ≈ 3–4s, `create` ≈ 3–5s. Not frozen.

**Quick Look is the debugger.** Drop one after any action, run, see what it
produced, delete it. Every fault in the original build was found this way in one
or two steps. Pair with **Count → Characters** when a value looks right but
behaves wrong — a key with a trailing space read as 33 characters.

## Clean up after — owed from three builds

- Failed runs write **blank or part-filled rows** and send completion mail. A
  stray blank row is indistinguishable from a legitimate deferred entry. That is
  the design, and the reason this matters.
- `create` **teaches the registry** whatever it was given, so junk suppliers and
  payers from test runs become permanent picker options. Fix them in the sheet,
  not in code — and the same applies to a typo entered through `+ New reason`.
