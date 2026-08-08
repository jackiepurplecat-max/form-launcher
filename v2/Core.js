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

/* ============================ Filename prefixes =========================== */

/** Regex matching any prefix any state in this section could have applied. */
function knownPrefixPattern(section) {
  const labels = section.states
    .filter(s => s.filePrefix)
    .map(s => s.filePrefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  if (!labels.length) return null;
  return new RegExp(`^(?:${labels.join('|')}) \\(\\d{2}-\\d{2}-\\d{4}\\) `);
}

function stripKnownPrefix(section, filename) {
  const pattern = knownPrefixPattern(section);
  return pattern ? filename.replace(pattern, '') : filename;
}

function extractFileId(fileRef) {
  if (!fileRef) return null;
  const match = fileRef.toString().match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Bring every receipt filename into line with the target state.
 *
 * Always strips whatever prefix is currently there before applying the target
 * state's prefix, so moving forwards and backwards use the same code path and
 * cannot drift apart.
 *
 * Returns one result per file rather than throwing: a rename failure must be
 * reported, not silently swallowed, but it must not roll back the status.
 */
function applyFilePrefixes(section, sheet, cols, row, targetState, dateForPrefix) {
  const results = [];

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
      const bare = stripKnownPrefix(section, file.getName());

      let newName = bare;
      if (targetState.filePrefix) {
        const stamp = Utilities.formatDate(
          new Date(dateForPrefix), Session.getScriptTimeZone(), 'dd-MM-yyyy'
        );
        newName = `${targetState.filePrefix} (${stamp}) ${bare}`;
      }

      if (newName !== file.getName()) file.setName(newName);
      results.push({ column: header, ok: true, name: newName });

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

  const files = applyFilePrefixes(
    section, sheet, cols, row, target, effectiveDate || today()
  );

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
