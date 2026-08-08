/**
 * v2 — section configuration.
 *
 * NOT YET DEPLOYED. This is for the new Google account. The old account still
 * runs the frozen Code.js at the repository root and must not be changed.
 *
 * Everything that differs between Work, IVA, Health and Income lives in this
 * file. Core.js is identical for all four sections. If you find yourself
 * writing `if (section === 'health')` in Core.js, it belongs here instead.
 *
 * Columns are referenced by HEADER NAME, never by index. Reordering or
 * inserting a column in a sheet must not require a code change.
 */

/**
 * Columns every section has, with identical headers and meaning.
 *
 * Date / Amount / Currency / Counterparty are the entry itself; the rest is
 * bookkeeping. Because these are shared, the generic renderer handles almost
 * every column without knowing which section it is showing.
 */
const COMMON = {
  timestamp: 'Timestamp',
  source: 'Source',

  date: 'Date',
  amount: 'Amount',
  currency: 'Currency',
  counterparty: 'Counterparty',

  status: 'Status',
  receiptUrl: 'Receipt URL',
  receiptState: 'Receipt State',
  notes: 'Notes'
};

/** Valid values for the "Receipt State" column. */
const RECEIPT_STATE = {
  attached: 'attached',
  awaiting: 'awaiting',
  notRequired: 'none required'
};

/**
 * WHY extraFields EXIST
 *
 * The point of these forms is that a claim can be submitted without reopening
 * the receipt. Fields like Número and Emitente NIF are retyped into Finanças,
 * and Health needs both the treatment date and the invoice date. Capturing
 * them once at entry is the entire value of the system.
 *
 * So completeness beats minimalism here. Do not trim these back to a tidy
 * shared core - a dropped field becomes a receipt you have to go and find.
 */

/**
 * Per-section configuration.
 *
 * states[] is ordered. Index position defines what "earlier" and "later" mean
 * when reverting, so the order here is significant.
 *
 *   name        displayed, and written verbatim into the Status column
 *   dateColumn  header of this state's own date column, or omitted for none
 *   filePrefix  label prefixed to receipt filenames while in this state,
 *               rendered as "<label> (DD-MM-YYYY) ". Omit for no prefix.
 *
 * counterpartyLabel
 *   What the shared Counterparty column is called in the UI. The column header
 *   stays "Counterparty" everywhere so the code stays generic.
 *
 * category
 *   Optional extra classifying field, present only where it means something.
 *   `managed: true` means its allowed values are a list you maintain, which
 *   populates the form dropdown — the generalisation of v1's add/delete
 *   expense reason.
 *
 * extraFields
 *   Section-specific columns beyond the shared core. Declared with type and
 *   label so the form, the table and the edit dialog can all render them
 *   without special-casing. See the note above on why these matter.
 *
 * fileColumns
 *   Each uploaded document, with the filename suffix it contributes.
 *   Health has two because a claim needs proof of need AND proof of payment.
 */
const SECTIONS = {
  work: {
    label: 'Work Expenses',
    sheet: 'Work',
    counterpartyLabel: 'Supplier',
    category: { header: 'Expense Reason', label: 'Expense Reason', managed: true },
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', filePrefix: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date' }
    ],
    fileColumns: [
      { header: COMMON.receiptUrl, label: 'Receipt', suffix: 'receipt' }
    ],
    extraFields: [],
    emailOnCreate: null
  },

  iva: {
    label: 'IVA Claim Receipts',
    sheet: 'IVA',
    counterpartyLabel: 'Retailer',
    category: null,
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', filePrefix: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date' }
    ],
    fileColumns: [
      { header: COMMON.receiptUrl, label: 'Fatura', suffix: 'fatura' }
    ],
    // All required by Finanças at submission time and retyped from here rather
    // than from the receipt. COMMON.amount holds Valor Total; the VAT figure is
    // its own field.
    extraFields: [
      { header: 'Número', label: 'Número', type: 'text', required: true },
      { header: 'Emitente NIF', label: 'Emitente NIF', type: 'text', required: true },
      { header: 'Tipo', label: 'Tipo', type: 'text', required: false },
      { header: 'Importados', label: 'Importados', type: 'boolean', required: false },
      { header: 'IVA Amount', label: 'Valor do IVA', type: 'number', required: true }
    ],
    // Sent when the entry is created (receipt uploaded), NOT on status change,
    // so no transition can re-send it.
    emailOnCreate: {
      recipientProperty: 'IVA_CLAIM_RECIPIENT',
      attachReceipt: true
    }
  },

  health: {
    label: 'Health Claim',
    sheet: 'Health',
    counterpartyLabel: 'Provider',
    category: { header: 'Patient', label: 'Patient', managed: true },
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', filePrefix: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date' }
    ],
    // A health claim needs two documents: proof that the expense was necessary,
    // and proof that it was paid. v1 called the first one "Details", which hid
    // what it was for.
    fileColumns: [
      { header: 'Justification URL', label: 'Prescription / Invoice', suffix: 'justification' },
      { header: COMMON.receiptUrl, label: 'Proof of payment', suffix: 'receipt' }
    ],
    // COMMON.date holds the treatment date - the event itself, consistent with
    // the other sections. The invoice date is also required by the claim.
    extraFields: [
      { header: 'Invoice Date', label: 'Invoice date', type: 'date', required: true }
    ],
    emailOnCreate: null
  },

  income: {
    label: 'Log Income',
    sheet: 'Income',
    counterpartyLabel: 'Paid by',
    category: null,
    states: [
      { name: 'Invoiced', dateColumn: 'Invoiced Date' },
      { name: 'Received', dateColumn: 'Received Date' },
      { name: 'Logged', dateColumn: 'Logged Date' }
    ],
    fileColumns: [],
    extraFields: [],
    emailOnCreate: null
  }
};

/**
 * OPEN QUESTION — does Settled / Logged also rename the file?
 * Currently only Claimed carries a filePrefix. If Settled should rename too,
 * add `filePrefix: 'Settled'` to that state. No code change needed: Core.js
 * strips whatever prefix is present and applies whatever the target state
 * declares.
 */

/** The state a newly created entry starts in: the first in the list. */
function initialState(section) {
  return section.states[0].name;
}

/** Display label for the Counterparty column in a given section. */
function counterpartyLabel(section) {
  return section.counterpartyLabel || 'Counterparty';
}
