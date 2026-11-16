# 2. Trips replace free-text Expense Reason

Date: 16 Nov 2026 · Status: accepted

Work's Expense Reason was open free text whose suggestions were every value ever
used, and `REBUILD-PLAN.md` expected a managed add/remove list on top of it. We
decided instead that a **Trip** is created on its own Trips tab, which names it
(`YYMMDD City`) and generates its **Planned Lines** from a checklist. The form's
Expense Reason becomes a closed dropdown of open reasons, plus *Add trip…* (a
round trip to the Trips tab and back) and *Other reason…* (free text for
non-trip reasons, refusing the six-digit trip prefix). A reason is offered while
any of its lines is not Settled. That rule is a Config setting, on for Work only.

## Considered options

- **Managed list (the plan's step 9).** Rejected: it would only add and remove
  names. A trip has a predictable set of expenses, and generating them is where
  the time goes. A list would still need a maintenance step for every trip.
- **Keep free text, filter by status.** Rejected: typos and near-duplicate
  names would split one trip across two reasons, and the "drop out when all
  Settled" rule would never fire for the half that holds an empty Planned Line.

## Consequences

- Trips live in a `Trips` sheet tab with headers generated from `Config.js`, like
  every other tab. Planned Lines are born through `createEntry()`, with no
  reminder mail.
- A trip's start date and city are never edited: delete it (only allowed while
  every line is untouched) and create it again. Only the end date is editable.
- Entering an expense with *New* for a trip that has a matching untouched
  Planned Line offers to fill that line instead of adding a duplicate.
  Otherwise the trip could never become all Settled.
- Siri gets a closed list too. Creating trips and filling Planned Lines from
  Siri is deferred (`.scratch/siri-shortcuts-performance/`).
