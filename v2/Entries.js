/**
 * v2 — entry creation.
 *
 * NOT YET DEPLOYED. See Config.js.
 *
 * An entry can be born two ways, and they differ in one respect only:
 *
 *   Google Form  Forms writes the row, then the trigger fires. The row already
 *                exists, so the adapter finalises it.
 *   Siri / OCR   Nothing exists yet, so the row is appended first.
 *
 * Both then run initializeEntry(), which is the single place bookkeeping
 * columns are set, documents are named and filed, and creation mail is sent.
 * Adding a third intake means writing an adapter, not touching this logic.
 */

/* ============================== Filenames ================================= */

/**
 * Reduce free text to something safe and consistent inside a filename:
 * accents flattened, each word capitalised so the result stays readable once
 * the spaces go, then everything non-alphanumeric removed.
 *
 *   "Hospital da Luz"  -> "HospitalDaLuz"
 *   "Farmácia Sá"      -> "FarmaciaSa"
 *   "José & Cia, Lda." -> "JoseCiaLda"
 */
function slugForFilename(text) {
  return (text || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // strip combining accents exposed by NFD
    .split(/\s+/)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join('')
    .replace(/[^A-Za-z0-9]/g, '');
}

/**
 * Base filename for a document: YYMMDD_Counterparty_Amount_<document>
 * The state suffix chain is appended later by applyFileState().
 */
function buildBaseFilename(sheet, cols, row, fileCol) {
  const dateValue = readCell(sheet, cols, row, COMMON.date);
  const stamp = dateValue
    ? Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), 'yyMMdd')
    : 'nodate';

  const who = slugForFilename(readCell(sheet, cols, row, COMMON.counterparty)) || 'unknown';
  const amount = formatAmountForFilename(readCell(sheet, cols, row, COMMON.amount));

  return [stamp, who, amount, fileCol.suffix].filter(part => part !== '').join('_');
}

/* ============================== Validation ================================ */

/**
 * Fields an entry should have. Returns a list of missing labels rather than
 * throwing, so the caller decides whether that is fatal.
 */
function missingFields(section, sheet, cols, row) {
  const missing = [];

  const core = [
    { header: COMMON.date, label: 'Date' },
    { header: COMMON.amount, label: 'Amount' },
    { header: COMMON.counterparty, label: counterpartyLabel(section) }
  ];
  if (section.category) {
    core.push({ header: section.category.header, label: section.category.label });
  }

  core.forEach(field => {
    const value = readCell(sheet, cols, row, field.header);
    if (value === '' || value === null || value === undefined) missing.push(field.label);
  });

  section.extraFields.filter(f => f.required).forEach(field => {
    const value = readCell(sheet, cols, row, field.header);
    if (value === '' || value === null || value === undefined) missing.push(field.label);
  });

  return missing;
}

/** attached / awaiting / none required, from what actually arrived. */
function receiptStateFor(section, sheet, cols, row) {
  if (!section.fileColumns.length) return RECEIPT_STATE.notRequired;

  const present = section.fileColumns.filter(fileCol => {
    const value = readCell(sheet, cols, row, fileCol.header);
    return value !== '' && value !== null && value !== undefined;
  });

  if (!present.length) return RECEIPT_STATE.awaiting;
  return present.length === section.fileColumns.length
    ? RECEIPT_STATE.attached
    : RECEIPT_STATE.awaiting;
}

/* ============================ Creation email ============================== */

/**
 * Mail sent when an entry is created — currently IVA only.
 *
 * Deliberately tied to creation rather than to a status change, so no
 * transition has a side effect beyond moving files, and re-selecting a state
 * can never re-send it.
 */
function sendCreationEmail(section, sheet, cols, row) {
  const spec = section.emailOnCreate;
  if (!spec) return null;

  try {
    const recipient = PropertiesService.getScriptProperties()
      .getProperty(spec.recipientProperty);
    if (!recipient) {
      return { ok: false, error: `${spec.recipientProperty} not set in Script Properties` };
    }

    const who = readCell(sheet, cols, row, COMMON.counterparty);
    const amount = readCell(sheet, cols, row, COMMON.amount);
    const currency = readCell(sheet, cols, row, COMMON.currency);
    const subject = `${section.label}: ${who} ${amount} ${currency}`.trim();

    const lines = [`${section.label} entry created.`, ''];
    [COMMON.date, COMMON.counterparty, COMMON.amount, COMMON.currency].forEach(header => {
      lines.push(`${header}: ${readCell(sheet, cols, row, header)}`);
    });
    section.extraFields.forEach(field => {
      lines.push(`${field.label}: ${readCell(sheet, cols, row, field.header)}`);
    });

    const options = {};
    if (spec.attachReceipt) {
      const attachments = [];
      section.fileColumns.forEach(fileCol => {
        const fileId = extractFileId(readCell(sheet, cols, row, fileCol.header));
        if (fileId) attachments.push(DriveApp.getFileById(fileId).getBlob());
      });
      if (attachments.length) options.attachments = attachments;
    }

    GmailApp.sendEmail(recipient, subject, lines.join('\n'), options);
    return { ok: true, recipient: recipient };

  } catch (error) {
    // Reported, never fatal: the entry exists and must not be lost because
    // mail failed.
    return { ok: false, error: error.toString() };
  }
}

/* ============================ Initialisation ============================== */

/**
 * Fill in bookkeeping, name and file the documents, send any creation mail.
 * Shared by every intake path.
 */
function initializeEntry(section, sheet, row, source) {
  const cols = resolveColumns(sheet);

  if (!readCell(sheet, cols, row, COMMON.timestamp)) {
    writeCell(sheet, cols, row, COMMON.timestamp, new Date());
  }
  writeCell(sheet, cols, row, COMMON.source, source || 'manual');

  // Initial state. Its date is taken from the entry when supplied, because a
  // state like Invoiced is normally backdated; today is only a fallback.
  const first = section.states[0];
  writeCell(sheet, cols, row, COMMON.status, first.name);
  if (first.dateColumn && !readCell(sheet, cols, row, first.dateColumn)) {
    writeCell(sheet, cols, row, first.dateColumn, today());
  }

  writeCell(sheet, cols, row, COMMON.receiptState, receiptStateFor(section, sheet, cols, row));

  // Name each document, then let applyFileState move it to the state folder.
  const renames = [];
  section.fileColumns.forEach(fileCol => {
    const fileId = extractFileId(readCell(sheet, cols, row, fileCol.header));
    if (!fileId) return;
    try {
      const file = DriveApp.getFileById(fileId);
      const ext = splitExtension(file.getName()).ext;
      file.setName(`${buildBaseFilename(sheet, cols, row, fileCol)}${ext}`);
      renames.push({ column: fileCol.header, ok: true, name: file.getName() });
    } catch (error) {
      renames.push({ column: fileCol.header, ok: false, error: error.toString() });
    }
  });

  const files = applyFileState(section, sheet, cols, row, 0);
  const warnings = missingFields(section, sheet, cols, row);
  const email = sendCreationEmail(section, sheet, cols, row);

  if (warnings.length) {
    Logger.log(`${section.sheet} row ${row}: incomplete — missing ${warnings.join(', ')}`);
  }

  return {
    ok: true,
    section: section.sheet,
    row: row,
    state: first.name,
    warnings: warnings,
    renames: renames,
    files: files,
    fileErrors: files.filter(f => !f.ok).concat(renames.filter(r => !r.ok)),
    email: email
  };
}

/* ================================ Intakes ================================= */

/**
 * Append a new entry. Used by Siri, OCR and manual creation.
 *
 * @param {string} sectionKey key into SECTIONS
 * @param {Object} fields     keyed by COLUMN HEADER, matching the sheet
 * @param {string} source     'siri' | 'ocr' | 'manual'
 */
function createEntry(sectionKey, fields, source) {
  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const cols = resolveColumns(sheet);

  const row = sheet.getLastRow() + 1;
  Object.keys(fields || {}).forEach(header => {
    if (!cols[header]) throw new Error(`Unknown column "${header}" for ${sectionKey}`);
    sheet.getRange(row, cols[header]).setValue(fields[header]);
  });

  const result = initializeEntry(section, sheet, row, source || 'manual');

  // Unlike a form submission, a programmatic caller can be told it got it
  // wrong, so incompleteness is surfaced as an error rather than a warning.
  if (result.warnings.length) {
    result.ok = false;
    result.error = `Missing required: ${result.warnings.join(', ')}`;
  }
  return result;
}

/** Find the section whose sheet a form response landed in. */
function sectionForSheet(sheetName) {
  const key = Object.keys(SECTIONS).find(k => SECTIONS[k].sheet === sheetName);
  return key ? { key: key, section: SECTIONS[key] } : null;
}

/**
 * Form submit trigger. Forms has already written the row, so this finalises it
 * rather than creating it.
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  if (row === 1) return;

  const match = sectionForSheet(sheet.getName());
  if (!match) {
    Logger.log(`No section configured for sheet: ${sheet.getName()}`);
    return;
  }

  const result = initializeEntry(match.section, sheet, row, 'form');
  Logger.log(`${sheet.getName()} row ${row}: created via form${
    result.fileErrors.length ? ` with ${result.fileErrors.length} file error(s)` : ''
  }`);
  return result;
}

/** Install the form submit trigger, replacing any existing ones. */
function installFormTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  Logger.log('Form submit trigger installed for onFormSubmit()');
}
