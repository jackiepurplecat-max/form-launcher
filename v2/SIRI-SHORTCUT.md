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

## Building "Log health claim"

The others are the same with a different `section` and one fewer or one more
question. Build this one first and duplicate it.

1. **Text** → paste the key. Rename the variable **Key**.
   *Do not type the key into each Get Contents action* — one place to change it.
2. **Text** → paste the endpoint URL. Rename to **URL**.
3. **Get Contents of URL** — URL, Method `POST`, Request Body `JSON`:
   `key` = Key, `section` = `health`, `action` = `catalog`.
4. **Get Dictionary Value** → `category.values` from the result.
   Shortcuts reads dotted keys, so this gets the list in one step.
5. **Choose from List** → that value. Prompt: *Who is it for?*
   This is a tap. A patient cannot be misheard.
6. **Ask for Input** → Text. *Which provider?*
7. **Ask for Input** → **Number**. *How much?*
   Number, not Text. Amount errors are the dangerous ones — a supplier typo is
   obvious in a list, "29.80" instead of "298" is not.
8. **Get Contents of URL** — `key`, `section` = `health`, `action` = `resolve`,
   `counterparty` = the provider from step 6.
9. **Get Dictionary Value** → `confirm`.
10. **Show Alert** (not a notification — you are looking at the phone):

    ```
    Health claim
    <Choose from List result> · <confirm>
    €<amount> · today
    ```

    Buttons **Save** and **Cancel**. Cancel stops the Shortcut here and nothing
    has been written.
11. **Get Contents of URL** — `key`, `section` = `health`, `action` = `create`,
    and a `fields` dictionary: `Counterparty` = the value from step 9 *(the
    resolved one, not what was dictated)*, `Amount` = step 7, `Patient` = step 5.
12. **Get Dictionary Value** → `ok`. **If** it is not `1`, get `error` and
    **Show Alert** with it. Otherwise get `complete`; if that is `0`, show
    *"Saved — completion mail sent"*.

Step 12 is not optional. Without it a refusal is silent and you find out weeks
later, from a sheet with a gap in it.

> An **Alert** is right here, where you invoked the Shortcut and are looking at
> the screen. It is wrong in an unattended automation, where it stalls waiting
> for a tap — which is why one was removed from the v1 iCloud Shortcut.

### The other three

- **Log expense** (`work`) — step 5 becomes **Ask for Input**, Text, *What for?*,
  sent as `Expense Reason`. It is an open list, so offer `category.values` from
  `catalog` as suggestions if you like, but free text must stay allowed. `Type`
  is not asked: the registry fills it, because Uber is always a Taxi.
- **Log receipt** (`iva`) — drop step 5 entirely. Número, Emitente NIF and Valor
  do IVA are **not** asked; they are retyped into Finanças from the completion
  form later. Every IVA entry made this way arrives incomplete, and that is
  expected.
- **Log income** (`income`) — step 5 becomes an optional **Ask for Input** sent
  as `Reason`.

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
