# Building the Shortcuts

How to build the four Shortcuts on the phone, and the protocol they speak. The
*reasons* for the design — prompted questions rather than one dictated sentence,
why the confirmation comes before the write — are in `REBUILD-PLAN.md` under
**Siri intake**. This file is the recipe.

One Shortcut per section, named so the phrase is natural. Say the name and Siri
runs it:

| Shortcut name | Section | You say |
|---|---|---|
| Log expense | `work` | "Hey Siri, log expense" |
| Log health claim | `health` | "Hey Siri, log health claim" |
| Log receipt | `iva` | "Hey Siri, log receipt" |
| Log income | `income` | "Hey Siri, log income" |

## Two things to have to hand

- **The endpoint URL** — the `/exec` URL of the `v2-siri` deployment, *not* the
  main web app.
- **The key** — returned once by `siriSetup()`, run from the main project's
  editor. It is in Script Properties as `SIRI_API_KEY` if you need it again.

Everything below posts JSON to that one URL. There are no other endpoints.

## The protocol

Three actions, all `POST`, all with `key` and `section` in the body. Every reply
is **HTTP 200** — `ContentService` cannot set a status code, so *the Shortcut
must read `ok` from the body* rather than assume the request succeeded. That is
the single most important thing on this page.

### 1. `catalog` — what to ask

```json
{ "key": "…", "section": "health", "action": "catalog" }
```

Replies with the labels and the list to tap:

```json
{ "ok": true, "label": "Health Claim", "counterpartyLabel": "Provider",
  "currency": "EUR", "date": "2026-08-12",
  "category": { "header": "Patient", "label": "Patient", "closed": true,
                "required": true, "values": ["Phoenix", "…"] } }
```

`category` is `null` for IVA, which has none. `closed: true` means the values
are the whole list and free text is wrong; `false` means they are suggestions.

This call exists so that **adding a patient never means editing a Shortcut**.

### 2. `resolve` — what was heard

```json
{ "key": "…", "section": "health", "action": "resolve",
  "counterparty": "wite clinic" }
```

```json
{ "ok": true, "heard": "wite clinic", "confirm": "White Clinic",
  "corrected": true, "known": true, "confidence": 0.92 }
```

**Show `confirm`, always.** It is the canonical spelling when the registry is
confident and exactly what was heard when it is not, so it is the right string
either way. `corrected: true` is worth calling out in the alert — that is the
0.92 mishearing being caught before it reaches a filename.

Writes nothing. Nothing exists until step 3.

### 3. `create` — what was confirmed

```json
{ "key": "…", "section": "health", "action": "create",
  "fields": { "Counterparty": "White Clinic", "Amount": 70, "Patient": "Phoenix" } }
```

```json
{ "ok": true, "row": 42, "complete": false,
  "outstanding": ["Invoice date", "Prescription / Invoice", "Proof of payment"],
  "awaitingDocument": true, "missingFields": ["Invoice date"],
  "completionEmailed": true, "counterparty": "White Clinic",
  "amount": 70, "date": "2026-08-12" }
```

`complete: false` **is not a failure.** A partial entry is the point: the row
exists, the completion mail has gone, and its link finishes the entry in one
tap. Report it, do not treat it as an error. `ok: false` is the failure.

`Date` and `Currency` default to today and EUR. Send them only to override.

**`fields` accepts these and nothing else** — anything else is refused outright
rather than dropped, so a Shortcut can never believe it recorded something it
did not:

| Section | Allowed in `fields` |
|---|---|
| work | `Counterparty`, `Amount`, `Currency`, `Date`, `Expense Reason` |
| health | `Counterparty`, `Amount`, `Currency`, `Date`, `Patient` |
| iva | `Counterparty`, `Amount`, `Currency`, `Date` |
| income | `Counterparty`, `Amount`, `Currency`, `Date`, `Reason` |

No document column is accepted from Siri, on purpose — see the header of
`v2/Siri.js`. Photograph the receipt later, through the web form.

## Read this before building anything

Health was built first, by hand, and cost far longer than it should have.
**Every one of the delays came from the same few Shortcuts behaviours**, none of
them to do with this project's server. They are listed here because they will
happen again on the next three.

**1. Name every result the moment it exists.** An action's output is a "magic
variable", and Shortcuts names it after the *kind* of action — `Contents of
URL`, `Dictionary Value`, `Provided Input`. A finished Shortcut here has **three
`Contents of URL` and four `Dictionary Value`, all with the same name**, so when
you drop one into a field Shortcuts *guesses* which you meant. It guesses "most
recent". **Five separate bugs in the first build were wrong guesses**, including
one that silently sent an empty supplier for two runs.

So: after every action whose result is used later, add **Set Variable** and give
it a real name. It is not optional discipline, it is the difference between an
hour and ten minutes.

**2. A variable chip can carry a *property* instead of the value.** Sometimes
Shortcuts attaches a property and shows the *property's* name on the chip —
`File Size`, `Values`. The condition then tests something you never meant. Tap
the chip; the sheet lets you pick the variable and clear the property.

**3. `Get Dictionary Value` output has no known type**, because it could return
anything. So an `If` on it offers only *has any value* / *does not have any
value* — no `is` or `is not`. To compare text, force the type: a **Text** action
containing only that variable, then **Set Variable**. A Text action always
outputs text, so the comparisons appear.

**4. Booleans do not compare to numbers.** `ok` comes back as JSON `true`, not
`1`. `If ok is not 1` fails with *"couldn't convert from Boolean to Dictionary"*.
Test the **presence of `error`** instead — see step 14.

**5. Quick Look is the debugger.** Drop one after any action to see exactly what
it produced, then delete it. Every fault in the first build was found this way in
one or two steps. Pair it with **Count → Characters** when a value looks right
but behaves wrong: a key that had a trailing space read as 33 characters.

**6. It is slow, and that is normal.** `catalog` ≈ 2s, `resolve` ≈ 3–4s,
`create` ≈ 3–5s. **Roughly ten seconds of waiting across the three calls**, with
the worst pause between the amount prompt and the alert. It is not frozen.

## Building "Log health claim"

Build this one, get it working, then duplicate for the other three.

**Setup — the two constants**

1. **Text** → paste the endpoint URL. → **Set Variable** `URL`.
2. **Text** → paste the key. → **Set Variable** `Key`.
   One place to change when the key is rotated. Both Text actions are called
   `Text` in the picker, which is exactly why they need names.

**Ask the questions**

3. **Get Contents of URL** → `URL`. Method `POST`; header `Content-Type` =
   `application/json`; Request Body `JSON` with `key` = `Key` variable,
   `section` = `health`, `action` = `catalog`.
4. **Get Dictionary Value** → `Value` for `category.values`.
   *`category.values` is a **key you type**, not a variable.* If dotted paths do
   not resolve, split into two actions.
5. **Choose from List** → prompt *Who is it for?* → **Set Variable** `Patient`.
   A tap, so it cannot be misheard.
6. **Ask for Input** → Text, *Which provider?* → **Set Variable** `Provider`.
7. **Ask for Input** → **Number**, *How much?*
8. **Get Numbers from Input** → **Set Variable** `Amount`.
9. **If** `Amount` **does not have any value** → **Show Alert** *"Amount must be
   a number"* → **Stop This Shortcut** → **End If**.
   Steps 8–9 exist because letters typed at the prompt otherwise travel all the
   way to the server and come back as a JSON parse error. The Number input type
   gives a numeric keypad on iOS but does not constrain typing on a Mac.

**Confirm before anything is written**

10. **Get Contents of URL** — duplicate the one from step 3 rather than building
    it again; it carries the method, header and body structure. Change `action`
    to `resolve` and add `counterparty` = `Provider`.
11. **Get Dictionary Value** → `confirm` → **Set Variable** `Confirmed`.
    Check the source is the **resolve** response, not the catalog one.
12. **Show Alert**, title *Log health claim*, **Show Cancel Button** on:

    ```
    Patient · Confirmed
    €Amount · Current Date
    ```

    **Use `Confirmed`, never `Provider`** — the corrected name is the entire
    point of step 10. Cancel stops the Shortcut and nothing has been written.

**Save**

13. **Get Contents of URL** — duplicate again; `action` = `create`, and a fourth
    row `fields` with **Type: Dictionary**. *Its expander is the **`>` chevron on
    the left**, not the "0 items" text on the right.* Inside it:

    | Key | Type | Value |
    |---|---|---|
    | `Counterparty` | Text | `Confirmed` |
    | `Amount` | **Number** | `Amount` |
    | `Patient` | Text | `Patient` |

    **`Amount` must be Type `Number`.** As Text it arrives as a string and gets
    parsed against the spreadsheet's Portuguese locale, where the decimal
    separator is a comma. No error, just a wrong figure.

    → **Set Variable** `CreateResult`.
14. **Get Dictionary Value** → `error` in `CreateResult` → **Set Variable**
    `ErrorText`.
15. **If** `ErrorText` **has any value** → **Show Alert** `ErrorText`;
    **Otherwise** → **Show Alert** *"Saved — finish it from the completion
    mail"*; **End If**.

    Testing `error` rather than `ok` avoids the Boolean problem entirely: on
    success the server sends no `error` key, so it is empty.

**Step 15 is not optional.** A refusal returns **HTTP 200** with `ok: false`, so
without it a rejected entry looks exactly like a saved one and you find out from
a gap in the sheet weeks later.

### Test the failure path

With `"section":"pets"` temporarily in the create request you should see
*"Unknown section: pets…"*. Change it back. Nothing is written either way, so
there is nothing to clean up.

### Alternative to the Dictionary builder

If the nested editor will not open, send the body as raw text instead: a **Text**
action holding the whole JSON, and Request Body set to **`File`** pointing at it.

```
{"key":"Key","section":"health","action":"create","fields":{"Counterparty":"Confirmed","Amount":Amount,"Patient":"Patient"}}
```

Quotes around every variable **except `Amount`**, which stays bare so it is a
JSON number. Type the line with placeholder words first and swap each for a
variable afterwards — it is far easier than inserting them as you go. Quick Look
it before running: you want `"Amount":70`, not `"Amount":"70"`.

Step 12 is not optional. Without it a refusal is silent and you find out weeks
later, from a sheet with a gap in it.

> An **Alert** is right here, where you invoked the Shortcut and are looking at
> the screen. It is wrong in an unattended automation, where it stalls waiting
> for a tap — which is why one was removed from the v1 iCloud Shortcut.

### The open-list picker — for Work and Income

Health's Patient list is **closed**: `Choose from List` and nothing else, because
the values are the whole set. Work's `Expense Reason` and Income's `Reason` are
**open** — free text is allowed and the list populates itself from the sheet.

Free text alone drifts. Expense Reason is how a trip's expenses group together
and, unlike suppliers, **nothing fuzzy-matches it** — so "Amsterdam trip" and
"Amsterdam Trip" become two silently separate things. So: offer the list *and* a
way out of it.

Replacing step 5:

```
Get Dictionary Value  category.values
Set Variable          Reasons
Text                  + New reason
Set Variable          NewMarker
Add to Variable       Reasons          ← the marker becomes one of the choices
Choose from List      Reasons            prompt: What for?
Text                  (the chosen item, alone)
Set Variable          ReasonChoice     ← forces the type; see trap 3
If  ReasonChoice  is  NewMarker
    Ask for Input     Text               prompt: New reason?
    Set Variable      Reason
Otherwise
    Set Variable      Reason  ← input set to ReasonChoice
End If
```

Two deliberate details:

- **The marker is a variable, not a typed string.** It is compared in the `If`
  and appended to the list, so writing it twice invites a mismatch that would
  make "new reason" silently unreachable. Define it once.
- **The Text action before `ReasonChoice`** is not decoration. `Choose from List`
  output is untyped, so without it the `If` offers only *has any value*.

A reason entered through `+ New reason` appears in the list next time by itself,
because `catalog` reads the column from the sheet. Nothing to maintain — and by
the same token a **typo becomes a permanent option**, so fix it in the sheet
rather than living with it.

### The other three

**Duplicate the finished health Shortcut** and edit the copy. Change `section`
in **all three** requests — missing one is the obvious mistake, and it fails in a
confusing way: catalog and resolve would answer for the wrong section while
create writes to the right one.

- **Log expense** (`work`) — **built and working.** Step 5 becomes the open-list
  picker above, sent as `Expense Reason` *(note the space — it is the real column
  header)*. `Type` is not asked: the registry fills it on an exact supplier
  match, because Uber is always a Taxi.
- **Log receipt** (`iva`) — **delete steps 3, 4 and 5 entirely.** IVA has no
  category, so there is nothing to fetch and nothing to choose, and dropping the
  `catalog` call makes this the fastest of the four by about two seconds. Número,
  Emitente NIF and Valor do IVA are **not** asked; they are retyped into Finanças
  from the completion form later. Every IVA entry made this way arrives
  incomplete — expected, not a fault.
- **Log income** (`income`) — the same open-list picker, sent as `Reason` (no
  space). Income's `Reason` is the one category that is **not required**, so add
  a second marker — `(none)` — alongside `+ New reason`, and in that branch set
  `Reason` to an empty Text action. Sending an empty string is fine; the field is
  optional and blank is a legitimate value everywhere in this system.

`iva` is the only one that loses actions. `income` is the closest thing to a
straight copy of `work`.

## When it does not work

- **Every reply says `Not authorized.`** — `SIRI_API_KEY` is not set on the
  **main** project, or the Shortcut's key does not match it. An unset key shuts
  the endpoint deliberately: the shim is deployed anonymously, so "no key
  configured" must never mean "no key required".
- **A reply mentions `SPREADSHEET_ID`** — run `siriSetup()` from the main
  project's editor. The shim is standalone and reaches the code as a library, so
  it has no container to resolve and needs the id.
- **Shortcuts shows an HTML page instead of JSON** — the deployment is wrong,
  not the code. Two causes, and Drive's wording does not distinguish them well:
  the project's scopes have never been authorised (open its editor and run
  `doGet` once), or the deployment's access is not really *Anyone*. Probe it
  outside Shortcuts with a keyless `{"action":"ping"}`; `Not authorized.` as
  **JSON** means the plumbing is fine.
- **A change to `v2/Siri.js` has no effect** — `npm run v2:push`, then
  `npm run v2:verify`. A push that reports success may have done nothing. The
  shim runs the main project's HEAD through a development-mode library, so
  there is no version to bump — but there is still a push to actually land.
