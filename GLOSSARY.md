# HelpfulForms

Captures expenses, IVA claims, health claims and income once, at the moment they
happen, so that later claims and declarations can be filled in without going back
to the paperwork.

## Language

### Health

**Health Claim**:
One treatment invoice to be claimed back from Allianz. It holds exactly what the
Allianz MyHealth app asks for, so the claim can be keyed in without reopening the receipt.
_Avoid_: Health expense, medical receipt

**Patient**:
The family member who received the treatment. A closed list.
_Avoid_: Claimant, member

**Provider**:
The clinic, doctor, pharmacy or lab that issued the invoice.
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
