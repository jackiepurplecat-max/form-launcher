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

/** Column headers that every section shares. */
const COMMON = {
  timestamp: 'Timestamp',
  source: 'Source',
  status: 'Status',
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
 */
const SECTIONS = {
  work: {
    label: 'Work Expenses',
    sheet: 'Work',
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', filePrefix: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date' }
    ],
    fileColumns: ['Receipt URL'],
    emailOnCreate: null
  },

  iva: {
    label: 'IVA Claim Receipts',
    sheet: 'IVA',
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', filePrefix: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date' }
    ],
    fileColumns: ['Receipt URL'],
    // Sent when the entry is created (receipt uploaded), NOT on status change,
    // so no transition can re-send it.
    emailOnCreate: {
      recipientProperty: 'IVA_CLAIM_RECIPIENT',
      subjectColumns: ['Número', 'Data'],
      attachReceipt: true
    }
  },

  health: {
    label: 'Health Claim',
    sheet: 'Health',
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', filePrefix: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date' }
    ],
    fileColumns: ['Receipt URL', 'Details URL'],
    emailOnCreate: null
  },

  income: {
    label: 'Log Income',
    sheet: 'Income',
    states: [
      { name: 'Invoiced', dateColumn: 'Invoiced Date' },
      { name: 'Received', dateColumn: 'Received Date' },
      { name: 'Logged', dateColumn: 'Logged Date' }
    ],
    fileColumns: [],
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
