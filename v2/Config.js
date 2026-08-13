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

/**
 * Column recording when a section's claim was emailed. Present only in sections
 * that have an emailOnCreate.
 *
 * This is what makes "email once, when the document arrives, and at no other
 * time" a property of the data rather than of the call graph. Without it the
 * guarantee rests on nobody ever calling the send path twice, and the completion
 * step - which exists precisely to finish an entry whose receipt came later -
 * has to call it a second time.
 */
const CLAIM_EMAILED_COLUMN = 'Claim Emailed';

/** Valid values for the "Receipt State" column. */
const RECEIPT_STATE = {
  attached: 'attached',
  awaiting: 'awaiting',
  notRequired: 'none required'
};

/**
 * WHERE THE DOCUMENT IS, as opposed to whether it has arrived yet.
 *
 * `Receipt State` says whether a document is attached. This says where to go and
 * find it when it is not: in the mail, or on paper in a bag, or both. The
 * completion mail reads it and tells you where to look.
 *
 * It exists because it is only reliably knowable AT CAPTURE TIME. Standing at
 * the counter you know whether you were handed paper; three days later, reading
 * a reminder, you do not. That is the same argument the prompted-questions
 * design rests on — a question that is asked cannot be forgotten.
 *
 * NOT REQUIRED, deliberately. Making it required would turn every work expense
 * submitted through the web form with its receipt already attached into an
 * INCOMPLETE entry, and mail a completion request for something that is
 * finished. It earns its place by being useful when set, not by being demanded.
 */
const RECEIPT_MEDIUM = {
  electronic: 'Electronic',
  physical: 'Physical',
  both: 'Both'
};

/**
 * One definition, shared by the three sections that have documents. Income has
 * no fileColumns, so it has nothing to find and does not get the field.
 *
 * Shared BY REFERENCE on purpose — that is what makes it one definition rather
 * than three that drift. Nothing may mutate it; the renderers build their own
 * objects from it.
 */
const RECEIPT_MEDIUM_FIELD = {
  header: 'Receipt Medium',
  label: 'Receipt is',
  type: 'choice',
  required: false,
  options: [RECEIPT_MEDIUM.electronic, RECEIPT_MEDIUM.physical, RECEIPT_MEDIUM.both]
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
 *   fileSuffix  label APPENDED to filenames on reaching this state, as
 *               "_<label>_<DD-MM-YYYY>". Suffixes accumulate, so a settled
 *               claim reads ..._Claimed_04-01-2026_Settled_20-01-2026.pdf and
 *               carries its own audit trail. Omit for no suffix.
 *   folder      subfolder under <root>/<Section>/ that files live in while in
 *               this state. Omit to leave files in the section inbox.
 *
 * counterpartyLabel
 *   What the shared Counterparty column is called in the UI. The column header
 *   stays "Counterparty" everywhere so the code stays generic.
 *
 * category
 *   Optional extra classifying field, present only where it means something.
 *   `managed: true` means its allowed values are a maintained list offered in
 *   the form. Since v2 has no Google Form, this list lives in the sheet rather
 *   than in a form question.
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
    // The trip varies, but the KIND of expense does not: Uber is always Taxi
    // whether the trip is Amsterdam or Plymouth. So Type prefills from the
    // registry and Expense Reason never does.
    registryTypeField: 'Type',
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', fileSuffix: 'Claimed', folder: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date', fileSuffix: 'Settled', folder: 'Settled' }
    ],
    fileColumns: [
      { header: COMMON.receiptUrl, label: 'Receipt', suffix: 'receipt' }
    ],
    extraFields: [
      // Education and Boarding Pass are confirmed against real claims and belong
      // HERE rather than on Expense Reason, which is open free text and needs no
      // list at all. The rest of this list is still the original proposal - add to
      // it when a claim turns up that none of these describe.
      {
        header: 'Type', label: 'Type', type: 'choice', required: false,
        options: [
          'Taxi', 'Train', 'Flight', 'Boarding Pass', 'Hotel', 'Meals',
          'Parking', 'Fuel', 'Education', 'Other'
        ]
      },
      RECEIPT_MEDIUM_FIELD
    ],
    // v1 mailed every work expense on form submission, receipt attached. That
    // is the claim being filed, so it carries forward - but tied to creation
    // like IVA's, and only sent once the entry is actually complete.
    emailOnCreate: {
      recipientProperty: 'WORK_CLAIM_RECIPIENT',
      attachReceipt: true
    }
  },

  iva: {
    label: 'IVA Claim Receipts',
    sheet: 'IVA',
    counterpartyLabel: 'Retailer',
    category: null,
    registryTypeField: null,
    registryNifField: 'Emitente NIF',
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', fileSuffix: 'Claimed', folder: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date', fileSuffix: 'Settled', folder: 'Settled' }
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
      { header: 'IVA Amount', label: 'Valor do IVA', type: 'number', required: true },
      RECEIPT_MEDIUM_FIELD
    ],
    // Fixed for every claim, so they are shown for reference rather than asked
    // per row. Values come from Script Properties, not from this file.
    reference: [
      { label: 'JALLC NIF', property: 'REF_JALLC_NIF' },
      { label: 'My NIF', property: 'REF_MY_NIF' },
      { label: 'Tipo', property: 'REF_IVA_TIPO' }
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
    // A closed list, unlike Work's Expense Reason. The patients are the family
    // and the set does not drift, so this is a dropdown rather than free text
    // with suggestions: one misspelling typed once becomes a second patient
    // forever and splits that person's claims across two values, with nothing
    // to warn you. Adding someone is a line here plus a push - deliberately a
    // change to the configuration rather than something a keystroke can do.
    //
    // INITIALS, on purpose. This repository is public, and these are family
    // members attached to health claims. The initials are unambiguous to the
    // one person who uses this and meaningless to anyone else, which is the
    // same reasoning that keeps the NIFs in Script Properties - with the
    // difference that a list this short stays readable in version control.
    category: {
      header: 'Patient', label: 'Patient', managed: true,
      options: ['J', 'K', 'A', 'P']
    },
    // White Clinic is usually Dentist, so remember it
    registryTypeField: 'Type',
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', fileSuffix: 'Claimed', folder: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date', fileSuffix: 'Settled', folder: 'Settled' }
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
      { header: 'Invoice Date', label: 'Invoice date', type: 'date', required: true },
      // Deliberately NOT the insurer's list, which is huge and multi-level.
      // This is a short one for your own tracking; the insurer's value is
      // chosen at submission time, not here.
      {
        header: 'Type', label: 'Type', type: 'choice', required: false,
        options: ['Doctor', 'Dentist', 'Optician', 'Prescription', 'Exam/Test']
      },
      RECEIPT_MEDIUM_FIELD
    ],
    emailOnCreate: null
  },

  income: {
    label: 'Log Income',
    sheet: 'Income',
    counterpartyLabel: 'Paid by',
    category: { header: 'Reason', label: 'Reason', managed: false, required: false },
    // The reason is currently fixed per payer, so let the registry remember it
    // rather than asking every time
    registryTypeField: 'Reason',
    // Income's three dates are business facts rather than bookkeeping: an
    // invoice is usually backdated, and money can arrive before the row is
    // made. So the form asks for them as ordinary fields, as well as the status
    // control filling them. No other section does this - a new entry is in the
    // first state, and setStatus clears the dates of every state after the
    // target, so a Claimed Date typed at creation would be wiped by the first
    // transition. That is also the answer to the plan's open question about
    // Invoiced Date being stamped with today when blank: now it can be entered.
    stateDatesInForm: true,
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
 * Script Property holding the ID of the app's root Drive folder.
 * Layout beneath it, all created on demand:
 *
 *   <root>/<Section>/Inbox      form uploads land here (state: To Do)
 *   <root>/<Section>/Claimed
 *   <root>/<Section>/Settled
 *   <root>/<Section>/Archived   archived and soft-deleted entries
 */
const ROOT_FOLDER_PROPERTY = 'ROOT_FOLDER_ID';

/**
 * Script Property holding the spreadsheet's own ID.
 *
 * Unused by anything running inside the container — getSpreadsheet() prefers
 * getActiveSpreadsheet() and never reaches for this. It exists for the Siri
 * project, which is standalone and reaches this code as a library, and so has
 * no container to resolve. See getSpreadsheet() in Core.js.
 */
const SPREADSHEET_ID_PROPERTY = 'SPREADSHEET_ID';

/**
 * Script Property holding the Drive folder that scans and saved mail
 * attachments land in, before they belong to an entry.
 *
 * Documents are PICKED out of it rather than uploaded again, so choosing one
 * moves it into the HelpfulForms tree and it leaves the folder by itself. That
 * is what stops this becoming a second copy of every receipt against a quota
 * managed by hand, and what keeps the folder's contents meaning "not yet filed".
 */
const STAGING_FOLDER_PROPERTY = 'STAGING_FOLDER_ID';

/** Folder used for entries in a state that declares no folder of its own. */
const INBOX_FOLDER = 'Inbox';

/** Folder that archived and soft-deleted files are moved to. */
const ARCHIVE_FOLDER = 'Archived';

/**
 * How an amount is rendered inside a filename.
 *
 * Always two decimal places, and the decimal point replaced, so a filename
 * contains exactly one dot - the extension. Multiple dots invite naive
 * split('.') parsing to break, in this codebase or any script, Shortcut or
 * OCR step added later.
 *
 * 3.4  -> "3-40"      3.456 -> "3-46"      1234.5 -> "1234-50"
 *
 * Set DECIMAL_IN_FILENAME to '.' to keep the dot instead, at the cost of
 * reintroducing that ambiguity.
 */
const DECIMAL_IN_FILENAME = '-';

/**
 * Currency the form starts with, and the one Siri will not ask about.
 *
 * A default rather than a fixed value: the column still accepts anything, but
 * almost every entry is euros and typing it each time is three taps for nothing.
 */
const DEFAULT_CURRENCY = 'EUR';

function formatAmountForFilename(amount) {
  // A blank amount must produce nothing, not "0-00". Number('') is 0, so
  // without this a partial Siri entry awaiting its amount would be filed under
  // a figure it does not have.
  if (amount === '' || amount === null || amount === undefined) return '';
  const n = Number(amount);
  if (!isFinite(n)) return '';
  return n.toFixed(2).replace('.', DECIMAL_IN_FILENAME);
}

/**
 * Every Script Property v2 reads.
 *
 * Real addresses and identifiers live here rather than in this file, because
 * the repository is public. v1 hardcoded the IVA recipient into Code.js and
 * put both NIFs in index.html; v2 does not repeat that.
 */
const SCRIPT_PROPERTY_INFO = {
  ROOT_FOLDER_ID: {
    required: true, secret: false,
    description: 'Drive folder containing <Section>/{Inbox,Claimed,Settled,Archived}'
  },
  IVA_CLAIM_RECIPIENT: {
    required: true, secret: false,
    description: 'Where an IVA claim is emailed when the entry is created'
  },
  WORK_CLAIM_RECIPIENT: {
    required: true, secret: false,
    description: 'Where a work expense claim is emailed when the entry is created. ' +
      'v1 sent these from sendWorkExpenseEmail() to RECIPIENT_EMAIL'
  },
  COMPLETION_EMAIL_RECIPIENT: {
    required: true, secret: false,
    description: 'Where "more info needed" mail goes when an entry arrives incomplete. ' +
      'Deliberately separate from IVA_CLAIM_RECIPIENT - this is a note to yourself ' +
      'and must never reach whoever processes claims'
  },
  WEB_APP_URL: {
    required: false, secret: false,
    description: 'The /exec URL of the versioned Web app deployment. Used to build the ' +
      'completion link, which opens the form on the entry rather than the sheet row. ' +
      'Set by hand so it names the PINNED deployment: getService().getUrl() returns ' +
      'whichever endpoint the running context belongs to, which for anything run from ' +
      'the editor is /dev. Unset, the completion mail falls back to the spreadsheet row'
  },
  REF_JALLC_NIF: {
    required: false, secret: false,
    description: 'Shown in the IVA section for copying into Finanças'
  },
  REF_MY_NIF: {
    required: false, secret: false,
    description: 'Shown in the IVA section for copying into Finanças'
  },
  REF_IVA_TIPO: {
    required: false, secret: false,
    description: 'Fixed Tipo text shown in the IVA section'
  },
  UI_ALLOWED_EMAILS: {
    required: false, secret: false,
    description: 'Comma-separated addresses allowed to use the web UI. Unset means ' +
      'only the account the script runs as, which is the intent for a personal tool'
  },
  SIRI_API_KEY: {
    required: false, secret: true,
    description: 'Key held only in the Shortcut, for the create-only endpoint'
  },
  STAGING_FOLDER_ID: {
    required: false, secret: false,
    description: 'Drive folder that Genius Scan and saved email attachments write to. ' +
      'The completion mail links to it, and the form lists what is in it so a document ' +
      'can be PICKED rather than uploaded - picking moves the file out, which is what ' +
      'keeps the folder meaningful and avoids a second copy of every receipt'
  },
  SPREADSHEET_ID: {
    required: false, secret: false,
    description: 'This spreadsheet\'s own id. Never read by anything running inside ' +
      'the container - only by the standalone Siri project, which reaches this code ' +
      'as a library and so has no active spreadsheet to resolve. Required as soon as ' +
      'the Siri endpoint exists; harmless before then'
  }
};

/** The state a newly created entry starts in: the first in the list. */
function initialState(section) {
  return section.states[0].name;
}

/**
 * The Receipt Medium field for a section, or null where it has none.
 *
 * Asked rather than inferred from `fileColumns`, so that a section could have
 * documents without being asked where they are — the two are related but not
 * the same question, and the config should decide, not this function.
 */
function sectionReceiptMedium(section) {
  return (section.extraFields || [])
    .filter(field => field.header === RECEIPT_MEDIUM_FIELD.header)[0] || null;
}

/** Display label for the Counterparty column in a given section. */
function counterpartyLabel(section) {
  return section.counterpartyLabel || 'Counterparty';
}
