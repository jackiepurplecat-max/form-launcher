# Rebuilding the three lost Shortcuts

Written 13 Aug 2026, after `Log expense`, `Log receipt` and `Log income` were
destroyed by an iCloud reconcile. `Log health claim` survived and is the
template. Companion to `SIRI-SHORTCUT.md`, which holds the protocol and the
traps — **read its "Read this before building anything" section first**; the six
behaviours listed there are what turn ten minutes into an hour, and they will
happen again.

This file exists because the diff table at the end of `SIRI-SHORTCUT.md`
(*"the other three"*) leaves two substitutions to be done in your head, and
because that table is now **one step out of date** — see the warning below.

## What went wrong, so it does not again

**iCloud Drive was switched off on the phone.** That did two things: exporting a
Shortcut to a file failed (Apple signs `.shortcut` files against a live iCloud
session), and Shortcuts sync sat in a state where the phone's near-empty library
reconciled *over* the Mac's four. There was no Time Machine and no export, so
three were unrecoverable.

So, while building:

- **Export each Shortcut to a file the moment it works**, not at the end.
  Share → Save to Files.
- **Those files contain `SIRI_API_KEY` and the `/exec` URL.** iCloud Drive is
  fine. The git repo is not.
- The verified backup of the template is
  `iCloud Drive/Downloads/Log health claim.shortcut`.

## ⚠️ The doc is out of date about IVA

`SIRI-SHORTCUT.md:322` says of `Log receipt`: *"delete steps 3, 4 and 5
entirely"*, and that dropping the `catalog` call makes it the fastest of the four
by about two seconds.

**That was true before `Receipt Medium` existed. It is now wrong.** IVA *does*
ask `Receipt Medium`, and its values come from `catalog`'s
`receiptMedium.values` — fetched, deliberately not hardcoded. Delete the
`catalog` call and there is no list to choose from.

**For IVA: keep step 3. Delete only steps 4 and 5.** The two-second saving is
gone; take it up with whoever added the medium question.

## The template you are duplicating

`Log health claim` as it stands — 15 steps plus the medium question. Every
sequence below is expressed as a change to this.

| # | Action | Notes |
|---|---|---|
| 1 | **Text** → endpoint URL → **Set Variable** `URL` | |
| 2 | **Text** → key → **Set Variable** `Key` | |
| 3 | **Get Contents of URL** `URL` — POST, `Content-Type: application/json`, JSON body `key`/`section`/`action: catalog` → **Set Variable** `Catalog` | |
| 4 | **Get Dictionary Value** `category.values` from `Catalog` | key you *type*, not a variable |
| 5 | **Choose from List** → *Who is it for?* → **Set Variable** `Patient` | closed list |
| 5a | **Get Dictionary Value** `receiptMedium.values` from `Catalog` → **Choose from List** → **Set Variable** `Medium` | |
| 6 | **Ask for Input** Text → *Which provider?* → **Set Variable** `Provider` | |
| 7 | **Ask for Input** **Number** → *How much?* | |
| 8 | **Get Numbers from Input** → **Set Variable** `Amount` | |
| 9 | **If** `Amount` *does not have any value* → Alert, **Stop This Shortcut** → End If | |
| 10 | **Get Contents of URL** — duplicate of 3; `action: resolve`, add `counterparty: Provider` → **Set Variable** `ResolveResult` | |
| 11 | **Get Dictionary Value** `confirm` from `ResolveResult` → **Set Variable** `Confirmed` | check the source is *resolve* |
| 12 | **Show Alert**, Show Cancel Button **on** | |
| 13 | **Get Contents of URL** — duplicate; `action: create` + `fields` Dictionary → **Set Variable** `CreateResult` | |
| 14 | **Get Dictionary Value** `error` in `CreateResult` → **Set Variable** `ErrorText` | |
| 15 | **If** `ErrorText` *has any value* → Alert `ErrorText`; **Otherwise** → Alert *"Saved…"* → End If | never optional |

**Before editing any duplicate:** change `section` in **all three** Get Contents
of URL actions — steps 3, 10 and 13. Missing one fails confusingly: catalog and
resolve answer for health while create writes to the right sheet.

`Date` and `Currency` default to today and EUR server-side. Do not send them.

---

## 1. Log expense — `work`

The open-list picker replaces the Patient tap. Everything else stands.

**Duplicate → rename `Log expense` → set `section: work` in steps 3, 10, 13.**

**Replace steps 4–5 with:**

```
Get Dictionary Value   category.values      (from Catalog)
Set Variable           Reasons
Text                   + New reason
Set Variable           NewMarker            ← define once; it is compared AND appended
Add to Variable        Reasons              ← input is NewMarker
Choose from List       Reasons                prompt: What for?
Text                   (the chosen item, alone)
Set Variable           ReasonChoice         ← forces the type; without it the If
                                              offers only "has any value"
If  ReasonChoice  is  NewMarker
    Ask for Input      Text                   prompt: New reason?
    Set Variable       Reason
Otherwise
    Set Variable       Reason               ← input is ReasonChoice
End If
```

**Keep 5a** (`Receipt Medium`). **Step 6** prompt → *Which supplier?*
**Step 12** title → *Log expense*, body:

```
Reason · Confirmed
€Amount · Current Date
```

**Step 13 `fields`:**

| Key | Type | Value |
|---|---|---|
| `Counterparty` | Text | `Confirmed` |
| `Amount` | **Number** | `Amount` |
| `Expense Reason` | Text | `Reason` |
| `Receipt Medium` | Text | `Medium` |

`Expense Reason` has a **space** — it is the real column header. `Type` is not
asked: the registry fills it on an exact supplier match, because Uber is always
a Taxi.

---

## 2. Log receipt — `iva`

The shortest of the three, and the only one that loses actions.

**Duplicate → rename `Log receipt` → set `section: iva` in steps 3, 10, 13.**

- **Delete step 4** (`category.values`) and **step 5** (the Patient picker).
  IVA has no category — `catalog` returns `category: null`.
- **Keep step 3.** See the warning above. `5a` needs it.
- **Keep 5a**, and everything from 6 onward.

**Step 6** prompt → *Which supplier?* **Step 12** title → *Log receipt*, body:

```
Confirmed
€Amount · Current Date
```

**Step 13 `fields`:**

| Key | Type | Value |
|---|---|---|
| `Counterparty` | Text | `Confirmed` |
| `Amount` | **Number** | `Amount` |
| `Receipt Medium` | Text | `Medium` |

Número, Emitente NIF and Valor do IVA are **not** asked — they are retyped into
Finanças from the completion form later. **Every IVA entry made this way arrives
incomplete. That is expected, not a fault.**

---

## 3. Log income — `income`

Closest to `work`, but it is the one that **loses the medium question**.

**Duplicate → rename `Log income` → set `section: income` in steps 3, 10, 13.**

**Delete 5a entirely — the picker *and* the `Receipt Medium` row in `fields`.**
Income has no documents, so `catalog` returns `receiptMedium: null` and there is
nothing to ask. This matters more than it looks: `fields` accepts
`Counterparty`, `Amount`, `Currency`, `Date`, `Reason` and **nothing else**, and
anything else is **refused outright rather than dropped**. Leave it in and every
income entry fails. It fails loudly — step 15's alert catches it — but the
message will not obviously point here.

**Replace steps 4–5** with the `work` picker above, with **two** markers instead
of one, because Income's `Reason` is the one category that is *not required*:

```
Get Dictionary Value   category.values      (from Catalog)
Set Variable           Reasons
Text                   + New reason
Set Variable           NewMarker
Text                   (none)
Set Variable           NoneMarker
Add to Variable        Reasons              ← NewMarker
Add to Variable        Reasons              ← NoneMarker
Choose from List       Reasons                prompt: What for?
Text                   (the chosen item, alone)
Set Variable           ReasonChoice
If  ReasonChoice  is  NewMarker
    Ask for Input      Text                   prompt: New reason?
    Set Variable       Reason
Otherwise If  ReasonChoice  is  NoneMarker
    Text               (empty)
    Set Variable       Reason               ← empty string is legitimate here
Otherwise
    Set Variable       Reason               ← ReasonChoice
End If
```

**Step 6** prompt → *Who from?* **Step 12** title → *Log income*, body:

```
Reason · Confirmed
€Amount · Current Date
```

**Step 13 `fields`:**

| Key | Type | Value |
|---|---|---|
| `Counterparty` | Text | `Confirmed` |
| `Amount` | **Number** | `Amount` |
| `Reason` | Text | `Reason` |

`Reason` has **no space** — unlike Work's `Expense Reason`. Sending an empty
string is fine; blank is a legitimate value everywhere in this system.

---

## Test each one before moving on

1. **The failure path.** Temporarily set `"section": "pets"` in the *create*
   request — expect *"Unknown section: pets…"*. Change it back. Nothing is
   written either way, so there is nothing to clean up.
2. **`Amount` really is a number.** Quick Look the create body: you want
   `"Amount":70`, not `"Amount":"70"`. As Text it is parsed against the
   spreadsheet's Portuguese locale, where the decimal separator is a comma — no
   error, just a wrong figure.
3. **One real entry**, then check the row landed in the right sheet with the
   right columns filled.
4. **Export to Files.**

Roughly ten seconds of waiting across the three calls is normal — `catalog` ≈ 2s,
`resolve` ≈ 3–4s, `create` ≈ 3–5s. It is not frozen.

## Clean up after

Building these **leaves debris**, and it is the kind that hides:

- Failed runs write **blank or part-filled rows** and send completion mail. A
  stray blank row is indistinguishable from a legitimate deferred entry — that is
  the whole design, and the reason this matters.
- `create` **teaches the registry** whatever it was given, so junk suppliers from
  test runs become permanent autocomplete options. Fix them in the sheet, not in
  code.

Then update `NEXT-SESSION.md`, which currently claims all four Shortcuts are
built and working.
