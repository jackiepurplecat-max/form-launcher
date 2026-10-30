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
