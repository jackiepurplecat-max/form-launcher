/**
 * v2 — core operations, identical for all four sections.
 *
 * NOT YET DEPLOYED. See Config.js.
 *
 * Everything section-specific comes from SECTIONS. Nothing here branches on
 * which section it is handling.
 */

/* ========================= Section and sheet access ======================= */

function getSection(sectionKey) {
  const section = SECTIONS[sectionKey];
  if (!section) throw new Error(`Unknown section: ${sectionKey}`);
  return section;
}

function getSheet(section) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(section.sheet);
  if (!sheet) throw new Error(`Sheet not found: ${section.sheet}`);
  return sheet;
}

/**
 * Validate a row number supplied by a caller. Never trust it: reject the header
 * row and anything past the end of the data.
 */
function resolveDataRow(sheet, sheetRow) {
  const row = Number(sheetRow);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
    throw new Error(`Invalid row for ${sheet.getName()}: ${sheetRow}`);
  }
  return row;
}

/* ============================ Columns by name ============================= */

const _columnCache = {};

/**
 * Map header text to 1-based column index for a sheet.
 *
 * This is what lets the four sheets have different layouts without the code
 * caring, and it is why v2 has no magic column numbers.
 */
function resolveColumns(sheet) {
  const key = sheet.getName();
  if (_columnCache[key]) return _columnCache[key];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, i) => {
    const name = (header || '').toString().trim();
    if (name) map[name] = i + 1;
  });

  _columnCache[key] = map;
  return map;
}

function columnIndex(cols, sheetName, header) {
  const index = cols[header];
  if (!index) throw new Error(`Column "${header}" not found in ${sheetName}`);
  return index;
}

function readCell(sheet, cols, row, header) {
  return sheet.getRange(row, columnIndex(cols, sheet.getName(), header)).getValue();
}

function writeCell(sheet, cols, row, header, value) {
  sheet.getRange(row, columnIndex(cols, sheet.getName(), header)).setValue(value);
}

/* ================================= States ================================= */

function stateIndex(section, stateName) {
  const name = (stateName || '').toString().trim();
  return section.states.findIndex(s => s.name === name);
}

function requireStateIndex(section, stateName) {
  const index = stateIndex(section, stateName);
  if (index === -1) {
    const valid = section.states.map(s => s.name).join(', ');
    throw new Error(`Unknown state "${stateName}". Expected one of: ${valid}`);
  }
  return index;
}

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ========================= Drive folders and names ======================== */

function extractFileId(fileRef) {
  if (!fileRef) return null;
  const match = fileRef.toString().match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/** Get a subfolder by name, creating it if absent. */
function childFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parent.createFolder(name);
}

/** Resolve <root>/<Section>/<folderName>, creating anything missing. */
function sectionFolder(section, folderName) {
  const rootId = PropertiesService.getScriptProperties().getProperty(ROOT_FOLDER_PROPERTY);
  if (!rootId) throw new Error(`${ROOT_FOLDER_PROPERTY} not set in Script Properties`);
  return childFolder(childFolder(DriveApp.getFolderById(rootId), section.label), folderName);
}

/** Where files for a given state belong. States with no folder use the inbox. */
function folderForState(section, state) {
  return sectionFolder(section, state.folder || INBOX_FOLDER);
}

function splitExtension(filename) {
  const match = filename.match(/^(.*?)(\.[^.\s]+)$/);
  return match ? { base: match[1], ext: match[2] } : { base: filename, ext: '' };
}

/**
 * Regex matching the accumulated state suffix chain at the end of a base name,
 * e.g. "_Claimed_04-01-2026_Settled_20-01-2026".
 */
function suffixChainPattern(section) {
  const labels = section.states
    .filter(s => s.fileSuffix)
    .map(s => s.fileSuffix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  if (!labels.length) return null;
  return new RegExp(`(?:_(?:${labels.join('|')})_\\d{2}-\\d{2}-\\d{4})+$`);
}

function stripSuffixChain(section, base) {
  const pattern = suffixChainPattern(section);
  return pattern ? base.replace(pattern, '') : base;
}

/**
 * Build the suffix chain for every state up to and including the target that
 * has a suffix and a recorded date.
 *
 * Derived from the row's dates rather than by editing the existing string, so
 * reverting simply produces a shorter chain. There is no separate "undo"
 * path that could drift out of step with the forward one.
 */
function buildSuffixChain(section, sheet, cols, row, targetIndex) {
  let chain = '';
  section.states.forEach((state, i) => {
    if (i > targetIndex || !state.fileSuffix || !state.dateColumn) return;
    const value = readCell(sheet, cols, row, state.dateColumn);
    if (!value) return;
    const stamp = Utilities.formatDate(
      new Date(value), Session.getScriptTimeZone(), 'dd-MM-yyyy'
    );
    chain += `_${state.fileSuffix}_${stamp}`;
  });
  return chain;
}

/**
 * Rename and relocate every document to match the target state.
 *
 * The name is rebuilt as <base><chain><ext>, where base is whatever remains
 * after removing any existing chain. Moving forwards lengthens the chain,
 * reverting shortens it, and both use this one path.
 *
 * Returns one result per file rather than throwing: a failure here must be
 * reported, not swallowed, but it must not roll back the status change.
 */
function applyFileState(section, sheet, cols, row, targetIndex) {
  const target = section.states[targetIndex];
  const chain = buildSuffixChain(section, sheet, cols, row, targetIndex);
  const results = [];
  let destination = null;

  section.fileColumns.forEach(fileCol => {
    const header = fileCol.header;
    const fileRef = (readCell(sheet, cols, row, header) || '').toString().trim();
    if (!fileRef) return;

    const fileId = extractFileId(fileRef);
    if (!fileId) {
      results.push({ column: header, ok: false, error: 'Not a Drive file reference' });
      return;
    }

    try {
      const file = DriveApp.getFileById(fileId);
      const { base, ext } = splitExtension(file.getName());
      const newName = `${stripSuffixChain(section, base)}${chain}${ext}`;
      if (newName !== file.getName()) file.setName(newName);

      // Resolved lazily and once: creating folders is slow, and a section with
      // no attached files should not create any.
      if (!destination) destination = folderForState(section, target);
      file.moveTo(destination);

      results.push({ column: header, ok: true, name: newName, folder: destination.getName() });

    } catch (error) {
      results.push({ column: header, ok: false, error: error.toString() });
    }
  });

  return results;
}

/* =============================== setStatus ================================ */

/**
 * Move an entry to a state. This is the only way Status ever changes, and it
 * replaces the four divergent toggle functions of v1.
 *
 * Moving backwards is an ordinary call, not a special "undo" path:
 *
 *   - Dates for states AFTER the target are cleared, so a row never claims a
 *     date for a state it is no longer in.
 *   - The target's own date is filled only if blank, so reverting
 *     Settled -> Claimed keeps the original Claimed Date instead of
 *     re-stamping today. This matters: mis-taps must not rewrite history.
 *   - An explicit date always wins, so any date can be corrected later.
 *
 * @param {string} sectionKey  key into SECTIONS
 * @param {number} sheetRow    1-based sheet row
 * @param {string} newState    target state name
 * @param {string} [dateISO]   yyyy-MM-dd; defaults to today only when blank
 */
function setStatus(sectionKey, sheetRow, newState, dateISO) {
  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const row = resolveDataRow(sheet, sheetRow);
  const cols = resolveColumns(sheet);

  const targetIndex = requireStateIndex(section, newState);
  const target = section.states[targetIndex];
  const previous = (readCell(sheet, cols, row, COMMON.status) || '').toString().trim();

  // Clear dates belonging to states later than the target
  section.states.forEach((state, i) => {
    if (state.dateColumn && i > targetIndex) {
      writeCell(sheet, cols, row, state.dateColumn, '');
    }
  });

  // Fill the target's date: explicit wins, otherwise only if currently blank
  let effectiveDate = null;
  if (target.dateColumn) {
    const existing = readCell(sheet, cols, row, target.dateColumn);
    if (dateISO) {
      effectiveDate = dateISO;
    } else if (!existing) {
      effectiveDate = today();
    } else {
      effectiveDate = existing;
    }
    writeCell(sheet, cols, row, target.dateColumn, effectiveDate);
  }

  writeCell(sheet, cols, row, COMMON.status, target.name);

  // After the dates are written, so the suffix chain reflects them
  const files = applyFileState(section, sheet, cols, row, targetIndex);

  const failed = files.filter(f => !f.ok);
  Logger.log(
    `${section.sheet} row ${row}: "${previous}" -> "${target.name}"` +
    (failed.length ? ` (${failed.length} file error(s))` : '')
  );

  // Report what actually happened. A caller must be able to tell that the
  // status moved but a rename failed - v1 returned success either way.
  return {
    ok: true,
    section: sectionKey,
    row: row,
    previousState: previous,
    state: target.name,
    date: effectiveDate,
    files: files,
    fileErrors: failed
  };
}

/**
 * Correct a date without changing state.
 * Only date columns declared by this section's states may be written.
 */
function setEntryDate(sectionKey, sheetRow, dateColumn, dateISO) {
  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const row = resolveDataRow(sheet, sheetRow);
  const cols = resolveColumns(sheet);

  const allowed = section.states.some(s => s.dateColumn === dateColumn);
  if (!allowed) throw new Error(`"${dateColumn}" is not a date column of ${sectionKey}`);

  writeCell(sheet, cols, row, dateColumn, dateISO || '');
  return { ok: true, section: sectionKey, row: row, column: dateColumn, date: dateISO || '' };
}
