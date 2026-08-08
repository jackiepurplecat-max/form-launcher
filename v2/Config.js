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
    fileColumns: [COMMON.receiptUrl],
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
    fileColumns: [COMMON.receiptUrl],
    // OPEN: v1 also collected Número, Emitente NIF, Tipo, Importados,
    // Valor do IVA and Valor Total. Confirm whether the reclaim process still
    // needs any of these before dropping them.
    extraFields: [],
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
    fileColumns: [COMMON.receiptUrl, 'Details URL'],
    extraFields: [],
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
