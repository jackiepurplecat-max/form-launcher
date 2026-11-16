# HelpfulForms

Captures expenses, IVA claims, health claims and income once, at the moment they
happen, so that later claims and declarations can be filled in without going back
to the paperwork.

## Language

### Across all sections

**Counterparty**:
The business or person on the other side of an entry. Each section has its own
word for it: Supplier (work expenses), Retailer (IVA), Provider (health), Paid by
(income). The Providers screen lists every Counterparty, not only health ones.
_Avoid_: Provider or supplier (when meaning any section)

### Health

**Health Claim**:
One treatment invoice to be claimed back from Allianz. It holds exactly what the
Allianz MyHealth app asks for, so the claim can be keyed in without reopening the receipt.
_Avoid_: Health expense, medical receipt

**Patient**:
The family member who received the treatment. A closed list.
_Avoid_: Claimant, member

**Provider**:
The clinic, doctor, pharmacy or lab that issued the invoice: Health's word for
its Counterparty. Not the same as the Providers screen, which covers all sections.
_Avoid_: Supplier, doctor (when meaning the business)

**Treatment Date**:
The day the treatment happened.
_Avoid_: Date (unqualified), appointment date

**Invoice Date**:
The date printed on the invoice, which can differ from the Treatment Date.
_Avoid_: Receipt date, billing date

**Treatment Type**:
Allianz's top-level category for the treatment, in Allianz's own wording.
_Avoid_: Type (unqualified), category

**Treatment Amount**:
What the treatment cost, in the invoice's currency, as Allianz's claim asks for it.
_Avoid_: Amount (unqualified), total, price

**Country of Treatment**:
The country the treatment took place in.
_Avoid_: Location

**Prescription / Invoice**:
The document showing the treatment was needed or what it was. Not every claim has one.
_Avoid_: Details, justification

**Proof of Payment**:
The document showing the treatment was paid for. A combined invoice-receipt counts as this.
_Avoid_: Receipt (when the invoice is meant)

**Awaiting Document**:
A Health Claim that is still waiting for another document to arrive. Set by you,
not inferred from which slots are filled.
_Avoid_: Incomplete, missing file

### IVA

**Número**:
The invoice number as Finanças will accept it, which can differ from the number
printed on the paper: only unaccented letters, digits and `/ \ _ . -`.
_Avoid_: Invoice number (when the printed one is meant), reference

### Work

**Expense Reason**:
What a work expense is claimed under: a trip, an education allowance, or another
claimable occasion. Many expenses share one Expense Reason.
_Avoid_: Trip (when the reason is not a trip), category

**Trip**:
An Expense Reason planned ahead, with a destination, a start date and an end
date. Its name is the start date and city, e.g. `261012 Brussels`. The only kind
of Expense Reason created on its own.
_Avoid_: Journey, travel

**Planned Line**:
A work expense created by a Trip before it happens, from a checklist of what that
kind of trip usually costs. It holds the date, Type and a note, and is filled in
when the receipt arrives. It stays untouched until it has a supplier, a receipt
or an amount other than its starting one.
_Avoid_: Placeholder, template line, draft
