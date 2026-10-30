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
 * Countries offered for Health's Country of Treatment. Portugal is the default
 * and nearly always right; the rest cover treatment abroad, Other the remainder.
 */
const COUNTRIES = [
  'Portugal', 'Spain', 'France', 'United Kingdom', 'Ireland', 'Netherlands',
  'Germany', 'Italy', 'Switzerland', 'United States', 'Other'
];

/**
 * WHY extraFields EXIST
 *
 * The point of these forms is that a claim can be keyed in without reopening
 * the receipt. So a section collects exactly what its claim form asks for:
 * IVA mirrors what Finanças asks, Health mirrors what the Allianz MyHealth app
 * asks, in Allianz's own wording and order (fieldGroups). Not more - every field
 * is retyped by hand - and not less, or it becomes a receipt you have to go and
 * find. Decided 30 Oct 2026; it replaced an older "collect everything" rule.
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
 *     defaultValue  what a new entry starts with when nothing is supplied
 *     follows       { header } copies that field's value into a blank one;
 *                   { header, values, otherwise } looks it up instead. Applied
 *                   when an entry is created, and live in the form until the
 *                   field is edited by hand.
 *     formerly      { header, values } - the column's old name and old values,
 *                   converted once by upgradeSection().
 *
 * fileColumns
 *   Each uploaded document, with the filename suffix it contributes.
 *
 * Optional, where a section needs them:
 *   dateLabel            what COMMON.date is called in the UI (default 'Date')
 *   fieldGroups          [{ heading?, headers }] - the order the form and the
 *                        card show fields in, and the headings between them
 *   amountLabel          what COMMON.amount is called in the UI (default 'Amount')
 *   cardTitle            headers whose values, joined, title a card in place of
 *                        the counterparty alone
 *   filenameInitialFrom  header whose first letter starts every filename
 *   awaitingDocumentField  Yes/No column deciding completeness when only some
 *                        fileColumns are filled. Without one, all must be.
 *   registryCountryField header the registry remembers per counterparty, as it
 *                        does registryTypeField
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
    // Quadro 02's "Tipo" — `codigoBem` in the declaration file, and a CLOSED list
    // because the tax app will not accept anything else.
    //
    // These eight are not a preference. DRORIVA derives the per-line Tipo from
    // the ENTITY type in Quadro 01, and for `d` — Funcionários de Embaixadas e
    // Organismos Internacionais, which is what REF_IVA_TIPO holds — exactly these
    // are legal. Read out of the app's own EntidadeEnum rather than off the
    // screen, so the list is complete instead of whatever was visible.
    //
    // What is deliberately ABSENT matters as much: food and drink, property
    // works, utilities, catering, accommodation and telephone are all claimable
    // by the embassy itself (type `c`) and NOT by an employee. Filing one of
    // those is a rejected line, so the list refuses it at intake instead.
    //
    // Codes live in the label so the exporter can recover them without a second
    // table to keep in step. See v2/IVA-EXPORT-FORMAT.md.
    category: {
      header: 'Tipo', label: 'Tipo', managed: false, required: false,
      options: [
        '101 — Vestuário e calçado',
        '102 — Electrodomésticos',
        '103 — Móveis',
        '104 — Jóias, bijutarias',
        '106 — Artigos de escritório',
        '107 — Outros bens não especificados',
        '156 — Reparação ou manutenção de veículos automóveis',
        '157 — Outros serviços não especificados'
      ]
    },
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
    // Finanças' rules for what may be attached to a pedido de restituição, from
    // DRORIVA's own "Ajuda — Anexação de ficheiros" panel:
    //
    //   - Apenas são aceites ficheiros anexos do formato PDF
    //   - Não são aceites ficheiros com mais do que 200 Kb
    //
    // Neither is in the XSD and neither is in the jar, so the app does NOT stop
    // you: a 3 MB scan attaches fine, saves fine, and is refused at submission.
    // That makes this an intake rule — the point of catching it is while the
    // scanner is still on the desk, not in the reply from AT.
    //
    // Generic on purpose (`documentRules`, not an IVA branch): any section may
    // declare it, and Form.js enforces whatever is declared. Only IVA has one
    // because only IVA hands its documents to a tax authority.
    documentRules: {
      extensions: ['pdf'],
      maxBytes: 200 * 1024,
      // Named in the message, because "200 Kb" is what the help says and a byte
      // count would send the user looking for a rule in those words.
      sizeLabel: '200 Kb',
      source: 'Finanças only accepts PDF faturas of 200 Kb or less'
    },
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
      // "Tipo de Entidade", not "Tipo": Quadro 01's entity type, which is fixed
      // for every claim. Renamed once the per-line Tipo above became a real
      // column, because two different things called Tipo on one screen — the
      // claimant's category and the purchase's — is a mis-selection waiting to
      // happen. The declaration calls them TipoEntidade and CodigoBem.
      { label: 'Tipo de Entidade', property: 'REF_IVA_TIPO' }
    ],
    // Sent when the entry is created (receipt uploaded), NOT on status change,
    // so no transition can re-send it.
    emailOnCreate: {
      recipientProperty: 'IVA_CLAIM_RECIPIENT',
      attachReceipt: true
    },
    // The declaration file DRORIVA opens — see v2/IvaExport.js and
    // v2/IVA-EXPORT-FORMAT.md. Only IVA has one; nothing else is submitted to a
    // tax authority.
    declaration: {
      // Quadro 01's TipoEntidade: `d`, Funcionários de Embaixadas e Organismos
      // Internacionais. A constant, and NOT read from REF_IVA_TIPO — that
      // property holds display text for the screen and would put a sentence
      // where the app wants one letter. The string `Embaixada` appears nowhere
      // in the app, so this code can only be learned from a saved file.
      entityTypeCode: 'd',
      // Both NIFs are Script Properties because this repo is public.
      entityNifProperty: 'REF_JALLC_NIF',       // NifEntidade — the employing body
      claimantNifProperty: 'REF_MY_NIF',        // NifFuncionario — you
      // Which rows are unsubmitted. Advancing to Claimed is what records a
      // submission, so the status already answers this and no `Submitted`
      // column is needed. Change this one string if that ever stops being true.
      sourceState: 'To Do',
      // Its own folder, not a state's. A declaration is not a receipt and does
      // not belong among the things it declares.
      folder: 'Declarations',
      // Base64 inflates by a third and these are scans; Apps Script holds the
      // whole declaration in memory as one string. Lines past the budget are
      // deferred to the next batch and reported, never dropped.
      maxBase64Bytes: 20 * 1024 * 1024,
      // "Não podem ser anexadas mais do que 50 faturas por pedido" — same help
      // panel as `documentRules` above. AT's own remedy for going over is to
      // split rather than to trim: "admite-se a submissão de tantos pedidos
      // quanto os necessários para completar o mesmo". So this is a BATCH SIZE,
      // handled by the same deferral as the byte budget, and never a refusal.
      maxLinesPerDeclaration: 50
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
    // THE LIST LIVES IN A SCRIPT PROPERTY, not here. These are family members
    // attached to health claims and this repository is public, so the names
    // cannot be committed - the same reasoning that keeps the NIFs and the
    // recipient addresses out of this file.
    //
    // It was initials for exactly that reason, and initials are worse: the
    // sheet is not public, so there is no benefit to storing "P" in it, and a
    // column of single letters is unreadable in the one place the data is
    // actually used. Moving the list to a property lets the real names be both
    // stored and displayed while staying out of version control.
    //
    // Still a CLOSED list. Where the values come from changed; what they mean
    // did not. One misspelling typed once would still become a second patient
    // for ever and split that person's claims.
    category: {
      header: 'Patient', label: 'Patient', managed: true,
      optionsProperty: 'HEALTH_PATIENTS'
    },
    // White Clinic is always Dental Expenses, and always in Portugal
    registryTypeField: 'Treatment Type',
    registryCountryField: 'Country of Treatment',
    // COMMON.date holds the treatment date - the event itself, consistent with
    // the other sections - so here it is labelled as such.
    dateLabel: 'Treatment date',
    // The Allianz app asks in two screens, so the form and the card show the
    // same two blocks, in its order and under its headings, to be read top to
    // bottom while keying a claim in. The last group has no heading: Allianz
    // asks for none of it. Anything not listed follows.
    fieldGroups: [
      {
        heading: 'Invoice information',
        headers: ['Patient', 'Country of Treatment', COMMON.counterparty, COMMON.currency, 'Invoice Date']
      },
      {
        heading: 'Treatment information',
        headers: [COMMON.date, 'Treatment Type', COMMON.amount]
      },
      {
        headers: [
          'Justification URL', COMMON.receiptUrl, 'Awaiting Document',
          RECEIPT_MEDIUM_FIELD.header, COMMON.notes
        ]
      }
    ],
    amountLabel: 'Treatment amount',
    // Who and where, so a claim is found by the patient first. The amount
    // stays on the right as everywhere else.
    cardTitle: ['Patient', COMMON.counterparty],
    // Filenames start with the patient's initial, so the Inbox groups by person
    // when picking a file in the Allianz app. The family's initials are distinct.
    filenameInitialFrom: 'Patient',
    states: [
      { name: 'To Do' },
      { name: 'Claimed', dateColumn: 'Claimed Date', fileSuffix: 'Claimed', folder: 'Claimed' },
      { name: 'Settled', dateColumn: 'Settled Date', fileSuffix: 'Settled', folder: 'Settled' }
    ],
    // Up to two documents: proof of what the treatment was, and proof that it
    // was paid. Not every claim has both - a doctor's receipt usually says what
    // the visit was, and a combined invoice-receipt goes in Proof of payment -
    // so which documents a claim needs is answered by Awaiting Document below,
    // not by counting filled slots.
    fileColumns: [
      { header: 'Justification URL', label: 'Prescription / Invoice', suffix: 'justification' },
      { header: COMMON.receiptUrl, label: 'Proof of payment', suffix: 'receipt' }
    ],
    // With both slots filled nothing more can come; with none, something must.
    // In between, this column decides.
    awaitingDocumentField: 'Awaiting Document',
    extraFields: [
      {
        header: 'Country of Treatment', label: 'Country of treatment', type: 'choice',
        required: false, options: COUNTRIES, defaultValue: 'Portugal'
      },
      // Usually the treatment date, so it starts as a copy of it
      {
        header: 'Invoice Date', label: 'Invoice date', type: 'date', required: true,
        follows: { header: COMMON.date }
      },
      // Allianz's first-level list, word for word. Its second level is chosen
      // in the app at claim time and not recorded here.
      {
        header: 'Treatment Type', label: 'Treatment type', type: 'choice', required: false,
        options: [
          'Doctor Visit', 'Medication and Medical Aids', 'Dental Expenses',
          'Optical', 'Tests/Scans', 'Alternative treatment'
        ],
        // Was "Type", with a short list of our own. upgradeSection() renames the
        // column and converts the old values, here and in the registry.
        formerly: {
          header: 'Type',
          values: {
            'Doctor': 'Doctor Visit', 'Dentist': 'Dental Expenses', 'Optician': 'Optical',
            'Prescription': 'Medication and Medical Aids', 'Exam/Test': 'Tests/Scans'
          }
        }
      },
      // A visit's receipt is usually the whole claim; a prescription or an exam
      // needs the receipt as well, and a dentist may send them separately.
      {
        header: 'Awaiting Document', label: 'Waiting for another document?', type: 'choice',
        required: false, options: ['Yes', 'No'],
        follows: {
          header: 'Treatment Type',
          values: { 'Doctor Visit': 'No', 'Alternative treatment': 'No' },
          otherwise: 'Yes'
        }
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
  HEALTH_PATIENTS: {
    required: false, secret: false,
    description: 'Comma-separated patient names for the Health section, in the order ' +
      'they should be offered. Lives here rather than in Config.js because the ' +
      'repository is public and these are family members attached to health claims. ' +
      'Unset means the dropdown is empty - the list is still CLOSED, so nothing can ' +
      'be entered until it is set'
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
 * The declared values of a section's category, or null when it is an open list.
 *
 * A category is CLOSED when it declares its values and OPEN when it does not —
 * open ones populate themselves from the column. The values may be declared
 * here in `options`, or named by `optionsProperty` for a list that cannot be
 * committed because the repository is public.
 *
 * Both are equally closed. Everything that used to read `category.options`
 * directly must come through here, or a property-backed list silently reads as
 * an open one: free text allowed, no dropdown, and the first typo becomes a
 * permanent second value.
 */
function categoryOptions(section) {
  const category = section.category;
  if (!category) return null;

  if (category.options && category.options.length) return category.options.slice();

  if (category.optionsProperty) {
    const raw = PropertiesService.getScriptProperties()
      .getProperty(category.optionsProperty) || '';
    const values = raw.split(',')
      .map(value => value.trim())
      .filter(value => value !== '');

    // An unset or empty property means "not configured yet", NOT "open list".
    // Returning null there would quietly re-open a closed list.
    return values.length ? values : [];
  }

  return null;
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

/** What the shared Date column is called in this section's UI. */
function dateLabel(section) {
  return section.dateLabel || 'Date';
}

/** What the shared Amount column is called in this section's UI. */
function amountLabel(section) {
  return section.amountLabel || 'Amount';
}

/**
 * Sort a list of { header } objects into the order of the section's
 * fieldGroups, and stamp each with `group`: the heading of the group it is in,
 * or null. Headers not in any group keep their relative order and follow the
 * listed ones. A section without fieldGroups gets its list back unchanged.
 */
function orderFields(section, list) {
  if (!section.fieldGroups) return list;
  const order = [];
  const headingOf = {};
  section.fieldGroups.forEach(group => {
    group.headers.forEach(header => {
      order.push(header);
      headingOf[header] = group.heading || null;
    });
  });
  const rank = item => {
    const at = order.indexOf(item.header);
    return at === -1 ? order.length : at;
  };
  return list
    .map((item, index) => ({ item: item, index: index }))
    .sort((a, b) => rank(a.item) - rank(b.item) || a.index - b.index)
    .map(pair => Object.assign({}, pair.item, { group: headingOf[pair.item.header] || null }));
}

/**
 * The value a `follows` rule gives a field, from the value of the field it
 * follows. '' when there is nothing to follow.
 */
function followedValue(follows, sourceValue) {
  if (!follows.values) return sourceValue;
  const key = (sourceValue === null || sourceValue === undefined) ? '' : String(sourceValue);
  return Object.prototype.hasOwnProperty.call(follows.values, key)
    ? follows.values[key]
    : (follows.otherwise || '');
}
