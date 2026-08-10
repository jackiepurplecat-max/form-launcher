/**
 * v2 — management: archive, restore, and permanent deletion.
 *
 * WHY DELETE DOES NOT DELETE
 *
 * Deleting from the table moves the row to the section's archive sheet and its
 * documents to the Archived folder. It removes nothing. Mis-tapping on a phone
 * is normal, a one-click unrecoverable delete of the wrong row is not worth the
 * convenience, and the row you meant to remove is junk anyway — so there is
 * nothing to gain by destroying it immediately.
 *
 * Permanent deletion exists, and its safeguard is structural rather than a
 * scarier dialog: hardDeleteEntry() only operates on the ARCHIVE sheet. Live
 * data cannot be destroyed in one action from anywhere, because the function
 * that destroys things cannot see it. Two deliberate steps, in two different
 * places.
 *
 * ORDER OF OPERATIONS
 *
 * Archiving writes the copy BEFORE it removes anything. Every ordering has a
 * failure mode; this one's is a row that exists in both places, which is
 * visible, harmless and repairable. The alternative is losing the row, so the
 * choice is not close. File moves happen in between and are reported rather
 * than rolled back, exactly as setStatus does with renames.
 */

/** Columns the archive adds on top of the section's own. */
const ARCHIVE_COLUMNS = {
  archivedAt: 'Archived',
  reason: 'Archive Reason'
};

/** Why a row is in the archive. Closed vocabulary, like Status. */
const ARCHIVE_REASON = {
  deleted: 'deleted',
  archived: 'archived'
};

/** The archive sheet for a section: "Work" -> "Work Archive". */
function archiveSheetName(section) {
  return `${section.sheet} Archive`;
}

/**
 * The archive's header row: the section's own spine, then the archive columns.
 *
 * Generated from the same sectionHeaders() the live sheet uses, so the two
 * cannot drift — adding a field to SECTIONS widens both.
 */
function archiveHeaders(section) {
  return sectionHeaders(section).concat([
    ARCHIVE_COLUMNS.archivedAt,
    ARCHIVE_COLUMNS.reason
  ]);
}

function getArchiveSheet(section) {
  const name = archiveSheetName(section);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`No sheet named "${name}". Run bootstrap() to create it.`);
  }
  return sheet;
}

/* ================================ Archive ================================= */

/**
 * Move one entry out of the live sheet and into the archive.
 *
 * Returns what actually happened, including any document that could not be
 * moved. A file failure does not abort the archive: the row is the record, and
 * a document left in the wrong folder is recoverable while a half-archived row
 * is not.
 */
function archiveEntry(sectionKey, sheetRow, reason) {
  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const row = resolveDataRow(sheet, sheetRow);
  const cols = resolveColumns(sheet);

  const why = reason === ARCHIVE_REASON.archived
    ? ARCHIVE_REASON.archived
    : ARCHIVE_REASON.deleted;

  const archive = getArchiveSheet(section);
  const archiveCols = resolveColumns(archive);
  const width = archive.getLastColumn();

  // Read the live row once, by header, so the two sheets need not share an
  // order - and so an archive sheet with columns in any arrangement still works.
  const headers = sectionHeaders(section);
  const carried = {};
  headers.forEach(header => {
    if (cols[header]) carried[header] = readCell(sheet, cols, row, header);
  });

  const values = new Array(width).fill('');
  Object.keys(carried).forEach(header => {
    const index = archiveCols[header];
    if (index) values[index - 1] = safeCellValue(carried[header]);
  });
  values[columnIndex(archiveCols, archive.getName(), ARCHIVE_COLUMNS.archivedAt) - 1] = new Date();
  values[columnIndex(archiveCols, archive.getName(), ARCHIVE_COLUMNS.reason) - 1] = why;

  // The copy lands first. Everything after this can fail without losing data.
  const archiveRow = withLock(() => {
    const target = archive.getLastRow() + 1;
    archive.getRange(target, 1, 1, width).setValues([values]);
    SpreadsheetApp.flush();
    return target;
  });

  const fileResults = [];
  section.fileColumns.forEach(fileCol => {
    const fileId = extractFileId(carried[fileCol.header]);
    if (!fileId) return;
    try {
      DriveApp.getFileById(fileId).moveTo(sectionFolder(section, ARCHIVE_FOLDER));
      fileResults.push({ column: fileCol.header, ok: true, folder: ARCHIVE_FOLDER });
    } catch (error) {
      fileResults.push({ column: fileCol.header, ok: false, error: error.toString() });
    }
  });

  sheet.deleteRow(row);
  SpreadsheetApp.flush();

  return {
    ok: true,
    section: sectionKey,
    row: row,
    archiveRow: archiveRow,
    reason: why,
    fileErrors: fileResults.filter(result => !result.ok),
    files: fileResults
  };
}

/**
 * Put an archived entry back in its section.
 *
 * The documents are re-filed by applyFileState rather than by moving them back
 * to wherever they came from, so the folder AND the filename suffix chain are
 * rebuilt from the row's own dates. That is the same code path a status change
 * uses, which is what stops restore drifting away from it.
 */
function restoreEntry(sectionKey, archiveSheetRow) {
  const section = getSection(sectionKey);
  const archive = getArchiveSheet(section);
  const row = resolveDataRow(archive, archiveSheetRow);
  const archiveCols = resolveColumns(archive);

  const sheet = getSheet(section);
  const cols = resolveColumns(sheet);
  const width = sheet.getLastColumn();

  const headers = sectionHeaders(section);
  const values = new Array(width).fill('');
  headers.forEach(header => {
    if (!archiveCols[header] || !cols[header]) return;
    values[cols[header] - 1] = safeCellValue(readCell(archive, archiveCols, row, header));
  });

  const liveRow = withLock(() => {
    const target = sheet.getLastRow() + 1;
    sheet.getRange(target, 1, 1, width).setValues([values]);
    SpreadsheetApp.flush();
    return target;
  });

  // An off-vocabulary status restores to the first state's folder rather than
  // throwing: the row is back either way, and the UI is where a bad status gets
  // noticed and repaired.
  const status = (readCell(sheet, cols, liveRow, COMMON.status) || '').toString().trim();
  const index = Math.max(0, stateIndex(section, status));
  const files = applyFileState(section, sheet, cols, liveRow, index);

  archive.deleteRow(row);
  SpreadsheetApp.flush();

  return {
    ok: true,
    section: sectionKey,
    row: liveRow,
    archiveRow: row,
    fileErrors: files.filter(result => !result.ok),
    files: files
  };
}

/* ============================= Hard delete ================================ */

/**
 * Destroy an archived entry and its documents.
 *
 * ONLY operates on the archive sheet, and that is the safeguard. Live data
 * cannot be destroyed in one action because this function cannot reach it — you
 * archive first, then purge, in two different places. A confirmation dialog
 * would be a weaker version of the same idea, since it is one tap away from
 * being trained out.
 *
 * Files go to Drive's trash rather than being removed outright, which leaves 30
 * days to change your mind. Storage is only reclaimed when the trash empties.
 */
function hardDeleteEntry(sectionKey, archiveSheetRow) {
  const section = getSection(sectionKey);
  const archive = getArchiveSheet(section);
  const row = resolveDataRow(archive, archiveSheetRow);
  const cols = resolveColumns(archive);

  const fileResults = [];
  section.fileColumns.forEach(fileCol => {
    const fileId = extractFileId(readCell(archive, cols, row, fileCol.header));
    if (!fileId) return;
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      fileResults.push({ column: fileCol.header, ok: true });
    } catch (error) {
      fileResults.push({ column: fileCol.header, ok: false, error: error.toString() });
    }
  });

  // The row goes whether or not every file could be trashed. A file that
  // survives is reported and can be removed by hand; a row that survives a
  // "permanently deleted" would be a lie about what happened.
  archive.deleteRow(row);
  SpreadsheetApp.flush();

  return {
    ok: true,
    section: sectionKey,
    archiveRow: row,
    filesTrashed: fileResults.filter(result => result.ok).length,
    fileErrors: fileResults.filter(result => !result.ok)
  };
}

/* ============================== UI wrappers =============================== */

/** Delete from the table: archives, removes nothing. */
function uiArchiveEntry(sectionKey, sheetRow) {
  requireUiAccess();
  return archiveEntry(sectionKey, sheetRow, ARCHIVE_REASON.deleted);
}

function uiRestoreEntry(sectionKey, archiveSheetRow) {
  requireUiAccess();
  const result = restoreEntry(sectionKey, archiveSheetRow);
  result.entry = uiEntry(sectionKey, result.row);
  return result;
}

/** Permanent, and reachable only for a row that is already archived. */
function uiHardDeleteEntry(sectionKey, archiveSheetRow) {
  requireUiAccess();
  return hardDeleteEntry(sectionKey, archiveSheetRow);
}

/**
 * The archive, newest first.
 *
 * Rendered by the same generic table as the live rows, so it returns the same
 * shape — with `archived` and `reason` alongside, and no status options, since
 * an archived row has no transitions to offer.
 */
function uiListArchive(sectionKey) {
  const viewer = requireUiAccess();

  const section = getSection(sectionKey);
  const archive = getArchiveSheet(section);
  const cols = resolveColumns(archive);
  const lastRow = archive.getLastRow();

  let rows = [];
  if (lastRow >= 2) {
    const values = archive.getRange(2, 1, lastRow - 1, archive.getLastColumn()).getValues();
    rows = values
      .map((rowValues, i) => {
        const entry = uiRow(section, archive.getName(), cols, rowValues, i + 2, viewer);
        if (!entry) return null;
        entry.options = [];  // nothing to transition an archived row to
        entry.archivedAt = uiDateISO(rowValues[cols[ARCHIVE_COLUMNS.archivedAt] - 1]);
        entry.reason = (rowValues[cols[ARCHIVE_COLUMNS.reason] - 1] || '').toString();
        return entry;
      })
      .filter(row => row !== null);
    rows.reverse();
  }

  return {
    ok: true,
    section: sectionKey,
    sheet: archive.getName(),
    meta: uiSectionMeta(sectionKey),
    rows: rows
  };
}
