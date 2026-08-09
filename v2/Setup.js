/**
 * v2 — one-time setup and configuration.
 *
 * NOT YET DEPLOYED. See Config.js.
 *
 * Run bootstrap() once against a blank spreadsheet and it builds everything
 * the rest of the code assumes exists: the sheets, their header spine, the
 * registry, and the Drive folder tree.
 *
 * WHY HEADERS ARE GENERATED RATHER THAN TYPED
 *
 * Every read and write in v2 resolves columns by header name, so a header
 * typo is not a cosmetic problem - it is a runtime failure in a function that
 * looks correct. Generating the headers FROM the same config the readers use
 * means the two cannot drift: add a field to SECTIONS, re-run bootstrap, and
 * the column appears.
 *
 * Everything here is idempotent. Re-running bootstrap() is also how you check
 * the setup, because it reports what it found as well as what it changed.
 */

/** Name of the Drive folder created to hold everything, if one is not set. */
const ROOT_FOLDER_NAME = 'HelpfulForms';

/* ============================== Header spine ============================== */

/**
 * The full ordered header list for a section.
 *
 * Order is for human readability only - nothing in the code depends on it,
 * which is the entire point of resolving by name. It runs roughly in the
 * order you would read an entry: when and how it arrived, what it was, how
 * much, where it has got to, and what is attached.
 *
 * Note that Receipt URL is NOT added separately. It is declared in the
 * fileColumns of the sections that have it, and Income has no documents at
 * all, so adding it here would give Income a column nothing ever writes.
 */
function sectionHeaders(section) {
  const headers = [
    COMMON.timestamp,
    COMMON.source,
    COMMON.date,
    COMMON.counterparty
  ];

  if (section.category) headers.push(section.category.header);
  section.extraFields.forEach(field => headers.push(field.header));

  headers.push(COMMON.amount, COMMON.currency, COMMON.status);

  section.states.forEach(state => {
    if (state.dateColumn) headers.push(state.dateColumn);
  });
  section.fileColumns.forEach(fileCol => headers.push(fileCol.header));

  headers.push(COMMON.receiptState);
  // Only where a claim is actually mailed, so the other sections do not carry a
  // column nothing ever writes
  if (section.emailOnCreate) headers.push(CLAIM_EMAILED_COLUMN);
  headers.push(COMMON.notes);

  // A duplicate header is a config error worth catching loudly. resolveColumns
  // builds a name -> index map, so a repeat would silently win and every read
  // of the first column would land on the second one instead.
  const seen = {};
  headers.forEach(header => {
    if (seen[header]) {
      throw new Error(
        `Duplicate header "${header}" configured for sheet ${section.sheet}`
      );
    }
    seen[header] = true;
  });

  return headers;
}

/**
 * Bring one sheet's headers into line with the config.
 *
 * Missing headers are APPENDED, never inserted, and existing ones are never
 * reordered or removed. Columns have data underneath them: reordering would
 * silently reassign it. So this is safe to run against a sheet already in use,
 * and adding a field to SECTIONS later is a one-line change plus a re-run.
 */
function applyHeaders(sheet, headers) {
  const width = sheet.getLastColumn();
  const existing = width
    ? sheet.getRange(1, 1, 1, width).getValues()[0]
        .map(header => (header || '').toString().trim())
    : [];
  const present = existing.filter(Boolean);

  if (!present.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    clearColumnCache();
    return { created: headers.length, added: [], extra: [] };
  }

  const added = headers.filter(header => present.indexOf(header) === -1);
  if (added.length) {
    // Appended past the last USED column, not past the count of non-blank
    // headers. A blank cell anywhere in the header row makes those two numbers
    // differ, and the difference would put the first new header on top of an
    // existing column - and its data.
    sheet.getRange(1, width + 1, 1, added.length).setValues([added]);
    clearColumnCache();
  }

  // Reported, not removed. An unrecognised column is usually something you
  // added deliberately, and deleting a column of real data to satisfy a config
  // file is not a trade this function gets to make.
  const extra = present.filter(header => headers.indexOf(header) === -1);

  return { created: 0, added: added, extra: extra };
}

/* ============================== Drive layout ============================== */

/**
 * Find or create the root folder and record its ID.
 *
 * Uses an already-configured ROOT_FOLDER_ID when there is one, so re-running
 * bootstrap() against a live system never orphans the existing tree.
 */
function ensureRootFolder() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(ROOT_FOLDER_PROPERTY);

  if (existingId) {
    try {
      const folder = DriveApp.getFolderById(existingId);
      return { folder: folder, id: existingId, created: false };
    } catch (error) {
      throw new Error(
        `${ROOT_FOLDER_PROPERTY} is set to "${existingId}" but that folder ` +
        `cannot be opened. Clear the property to create a new one, or fix ` +
        `the ID. (${error})`
      );
    }
  }

  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty(ROOT_FOLDER_PROPERTY, folder.getId());
  return { folder: folder, id: folder.getId(), created: true };
}

/**
 * Every folder a section needs: the inbox, one per state that declares a
 * folder, and the archive.
 *
 * Built by calling Core's own sectionFolder() rather than by reimplementing
 * the path here. That matters - if the two disagreed about where a section's
 * files live, bootstrap would quietly create a second, empty tree alongside
 * the real one and nothing would look wrong until a file went missing.
 */
function ensureSectionFolders(section) {
  // A section with no documents gets no folders. Core only ever creates one
  // lazily, when there is actually a file to put in it, so building them here
  // would leave Income with an empty tree that nothing writes to.
  if (!section.fileColumns.length) return [];

  const names = [INBOX_FOLDER];
  section.states.forEach(state => {
    if (state.folder && names.indexOf(state.folder) === -1) names.push(state.folder);
  });
  names.push(ARCHIVE_FOLDER);

  names.forEach(name => sectionFolder(section, name));
  return names;
}

/* ================================ Bootstrap =============================== */

/**
 * Build everything. Safe to run repeatedly.
 *
 * Run this from the Apps Script editor with the project bound to the target
 * spreadsheet. It will ask for Drive and Spreadsheet authorisation the first
 * time.
 */
function bootstrap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = { spreadsheet: ss.getName(), sections: {}, warnings: [] };

  Object.keys(SECTIONS).forEach(key => {
    const section = SECTIONS[key];
    let sheet = ss.getSheetByName(section.sheet);
    const isNew = !sheet;
    if (isNew) sheet = ss.insertSheet(section.sheet);

    const headers = applyHeaders(sheet, sectionHeaders(section));
    if (headers.extra.length) {
      report.warnings.push(
        `${section.sheet}: unrecognised column(s) left in place - ${headers.extra.join(', ')}`
      );
    }

    report.sections[key] = {
      sheet: section.sheet,
      sheetCreated: isNew,
      headersWritten: headers.created,
      headersAdded: headers.added,
      unrecognisedColumns: headers.extra
    };
  });

  // The registry looks after its own sheet, so creating it is just a matter of
  // asking for it once.
  getOrCreateRegistrySheet();
  report.registry = REGISTRY_SHEET;

  // Drive comes last: it is the only part that can fail on quota or
  // authorisation, and a half-built folder tree is easier to live with than a
  // half-built spreadsheet.
  const root = ensureRootFolder();
  report.rootFolder = {
    id: root.id,
    name: root.folder.getName(),
    created: root.created,
    url: root.folder.getUrl()
  };
  Object.keys(SECTIONS).forEach(key => {
    report.sections[key].folders = ensureSectionFolders(SECTIONS[key]);
  });

  // Google puts a default sheet in every new spreadsheet, and names it in the
  // account's own language - "Sheet1", "Folha1", "Hoja1". Checking for the
  // English name only would silently miss it, so report any tab that is not one
  // this code knows about. Reported rather than removed, in case it is the one
  // holding your notes.
  const known = Object.keys(SECTIONS).map(key => SECTIONS[key].sheet).concat([REGISTRY_SHEET]);
  const leftover = ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => known.indexOf(name) === -1);
  if (leftover.length) {
    report.warnings.push(
      `Unrecognised sheet(s) present - delete by hand if unused: ${leftover.join(', ')}`
    );
  }
  report.unrecognisedSheets = leftover;

  report.propertiesStillNeeded = Object.keys(SCRIPT_PROPERTY_INFO).filter(key => {
    if (!SCRIPT_PROPERTY_INFO[key].required) return false;
    return !PropertiesService.getScriptProperties().getProperty(key);
  });

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/* ============================ Script Properties =========================== */

/**
 * Values to write. Fill in what you want to set, then run
 * setupScriptProperties(). Anything left blank is skipped, so re-running can
 * never overwrite a real value with a placeholder.
 *
 * DO NOT COMMIT REAL VALUES HERE. This repository is public. Fill these in the
 * Apps Script editor, run the function, then blank them again - the values
 * live in Script Properties from that point on.
 *
 * ROOT_FOLDER_ID is deliberately absent: bootstrap() sets it.
 *
 * Where to find the values, none of which live in this repository's history:
 *
 *   IVA_CLAIM_RECIPIENT        .env, as V2_IVA_CLAIM_RECIPIENT (gitignored)
 *   WORK_CLAIM_RECIPIENT       v1 sent these to RECIPIENT_EMAIL in .env
 *   COMPLETION_EMAIL_RECIPIENT .env, as V2_COMPLETION_EMAIL_RECIPIENT
 *   REF_JALLC_NIF         v1's built index.html, in the IVA reference block
 *   REF_MY_NIF            likewise
 *   REF_IVA_TIPO          likewise
 *   SIRI_API_KEY          generate when the Siri endpoint is built; not needed
 *                         to get the system running
 */
const SCRIPT_PROPERTY_VALUES = {
  IVA_CLAIM_RECIPIENT: '',
  WORK_CLAIM_RECIPIENT: '',
  COMPLETION_EMAIL_RECIPIENT: '',
  REF_JALLC_NIF: '',
  REF_MY_NIF: '',
  REF_IVA_TIPO: '',
  SIRI_API_KEY: ''
};

/**
 * Store configuration in Script Properties.
 * Safe to re-run: blank entries are left untouched.
 */
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  const written = [];
  const skipped = [];

  Object.keys(SCRIPT_PROPERTY_VALUES).forEach(key => {
    if (!SCRIPT_PROPERTY_INFO[key]) {
      throw new Error(`"${key}" is not declared in SCRIPT_PROPERTY_INFO`);
    }
    const value = (SCRIPT_PROPERTY_VALUES[key] || '').toString().trim();
    if (!value) {
      skipped.push(key);
      return;
    }
    props.setProperty(key, value);
    written.push(key);
  });

  Logger.log(
    `Script Properties written: ${written.join(', ') || 'none'}\n` +
    `Left untouched: ${skipped.join(', ') || 'none'}`
  );
  return checkScriptProperties();
}

/**
 * Report what is configured, without revealing secrets.
 *
 * Run after setupScriptProperties(), and any time something fails in a way
 * that smells like missing configuration.
 */
function checkScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const report = { set: [], missing: [], missingRequired: [], unknown: [] };

  Object.keys(SCRIPT_PROPERTY_INFO).forEach(key => {
    const info = SCRIPT_PROPERTY_INFO[key];
    const value = all[key];
    if (value) {
      report.set.push({
        key: key,
        value: info.secret ? `set (${value.length} chars)` : value
      });
    } else {
      report.missing.push(key);
      if (info.required) report.missingRequired.push(key);
    }
  });

  report.unknown = Object.keys(all).filter(key => !SCRIPT_PROPERTY_INFO[key]);
  report.ok = report.missingRequired.length === 0;

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
