/**
 * v2 — the live smoke test. Build order step 5.
 *
 * The Apps Script editor can only run functions that take no arguments, so
 * "call createEntry() from the editor" needs a wrapper. This is it.
 *
 * It is the live counterpart of v2/test/run.js: the harness proves the logic
 * against stand-ins, this proves it against real Sheets and real Drive, where
 * quotas, authorisation and date coercion actually behave. Run it once, read
 * the log, then delete the rows it made.
 *
 *   smokeTest()      creates one entry per section, cycles its status, checks
 *                    the filenames and folders at each step, reports
 *   smokeCleanup()   deletes the rows and files it created, and nothing else
 *
 * MAIL: it never sends a claim to a real claims address. Every section that mails
 * on creation - Work and IVA - is given no document, so its claim defers by
 * design, and that deferral is itself one of the things being checked. To test a
 * claim mail for real, see the deployment notes: point that section's recipient
 * property at yourself first.
 *
 * It DOES send one "more info needed" mail per mailing section, to
 * COMPLETION_EMAIL_RECIPIENT. That is the point: it proves the address works.
 * Expect two - one naming the Receipt as outstanding, one naming the Fatura.
 */

/** Marker written into Notes, so cleanup can find exactly these rows. */
const SMOKE_MARKER = 'SMOKE TEST - safe to delete';

/** Counterparty used throughout, so the registry entry it teaches is findable. */
const SMOKE_COUNTERPARTY = 'Smoke Test Ltd';

function _smokeFile(name) {
  // A real Drive file, so renaming and moving are genuinely exercised
  return DriveApp.createFile(Utilities.newBlob('smoke test placeholder', 'text/plain', name));
}

/**
 * Create one entry per section, walk it through its states, and check what
 * actually happened to the sheet and to Drive.
 */
function smokeTest() {
  const results = [];
  const checks = [];
  const check = (label, ok, detail) => checks.push({ label: label, ok: !!ok, detail: detail });

  const stamp = today();

  Object.keys(SECTIONS).forEach(key => {
    const section = SECTIONS[key];
    const fields = {};

    fields[COMMON.date] = stamp;
    fields[COMMON.counterparty] = SMOKE_COUNTERPARTY;
    fields[COMMON.amount] = 3.45;
    fields[COMMON.currency] = 'EUR';
    fields[COMMON.notes] = SMOKE_MARKER;

    if (section.category) fields[section.category.header] = 'Smoke';
    section.extraFields.forEach(field => {
      if (!field.required) return;
      fields[field.header] = field.type === 'date' ? stamp
        : field.type === 'number' ? 1
          : 'Smoke';
    });

    // Any section that mails on creation is left WITHOUT its document on
    // purpose. A missing receipt is what makes the claim defer, and deferring is
    // the only thing standing between a smoke test and a junk claim landing in
    // a real work inbox. Driven off emailOnCreate rather than a section name, so
    // giving another section a claim email cannot quietly re-arm this.
    //
    // The cost is that file naming and filing are not exercised for those
    // sections - Health covers that path in full, with two documents, and the
    // path is entirely generic.
    const files = {};
    if (!section.emailOnCreate) {
      section.fileColumns.forEach(fileCol => {
        const file = _smokeFile(`smoke_${key}_${fileCol.suffix}.txt`);
        files[fileCol.header] = file;
        fields[fileCol.header] = file.getId();
      });
    }

    let created;
    try {
      created = createEntry(key, fields, 'manual');
    } catch (error) {
      check(`${key}: createEntry`, false, error.toString());
      return;
    }

    check(`${key}: created`, created.ok === true, created.error || created.warnings);
    check(`${key}: no file errors`, created.fileErrors.length === 0, created.fileErrors);
    check(`${key}: registry learned`, !!created.registry, created.registry);
    if (section.emailOnCreate) {
      check(`${key}: claim mail deferred, not sent`,
        created.email && created.email.deferred === true, created.email);
      check(`${key}: asked for the missing document instead`,
        created.completionRequest && created.completionRequest.ok === true,
        created.completionRequest);
    }

    // Filenames and folders, at creation and then at each state
    Object.keys(files).forEach(header => {
      const file = DriveApp.getFileById(files[header].getId());
      check(`${key}: ${header} named`, file.getName().indexOf('SmokeTestLtd_3-45') !== -1, file.getName());
    });

    section.states.slice(1).forEach(state => {
      const moved = setStatus(key, created.row, state.name);
      check(`${key}: -> ${state.name}`, moved.fileErrors.length === 0, moved.fileErrors);
      if (state.fileSuffix) {
        Object.keys(files).forEach(header => {
          const file = DriveApp.getFileById(files[header].getId());
          check(
            `${key}: ${state.name} suffix on ${header}`,
            file.getName().indexOf(`_${state.fileSuffix}_`) !== -1,
            file.getName()
          );
          check(
            `${key}: ${state.name} folder`,
            file.getParents().next().getName() === (state.folder || INBOX_FOLDER),
            file.getParents().next().getName()
          );
        });
      }
    });

    // Back to the first state: the chain must shorten, not accumulate
    const reverted = setStatus(key, created.row, section.states[0].name);
    check(`${key}: reverted to ${section.states[0].name}`, reverted.ok === true, reverted);
    Object.keys(files).forEach(header => {
      const file = DriveApp.getFileById(files[header].getId());
      const bare = section.states.every(s => !s.fileSuffix || file.getName().indexOf(`_${s.fileSuffix}_`) === -1);
      check(`${key}: ${header} chain stripped on revert`, bare, file.getName());
    });

    results.push({ section: key, row: created.row, fileIds: Object.keys(files).map(h => files[h].getId()) });
  });

  const failed = checks.filter(c => !c.ok);
  const report = {
    ok: failed.length === 0,
    passed: checks.length - failed.length,
    failed: failed,
    entries: results,
    next: failed.length
      ? 'Fix the failures above before building anything on top.'
      : 'All good. Run smokeCleanup() to remove these rows and files.'
  };

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/**
 * Remove what smokeTest() made.
 *
 * Only ever touches rows whose Notes hold SMOKE_MARKER, so it cannot take a
 * real entry with it. Rows are deleted bottom-up, because deleting one shifts
 * every row beneath it.
 */
function smokeCleanup() {
  const removed = { rows: {}, files: 0, registry: null, warnings: [] };

  Object.keys(SECTIONS).forEach(key => {
    const section = SECTIONS[key];
    const sheet = getSheet(section);
    const cols = resolveColumns(sheet);
    if (sheet.getLastRow() < 2) return;

    const notesCol = columnIndex(cols, sheet.getName(), COMMON.notes);
    const notes = sheet.getRange(2, notesCol, sheet.getLastRow() - 1, 1).getValues();

    const targets = [];
    notes.forEach((value, i) => {
      if ((value[0] || '').toString().trim() === SMOKE_MARKER) targets.push(i + 2);
    });

    // Trash the documents before the row that points at them is gone
    targets.forEach(row => {
      section.fileColumns.forEach(fileCol => {
        const fileId = extractFileId(readCell(sheet, cols, row, fileCol.header));
        if (!fileId) return;
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
          removed.files++;
        } catch (error) {
          removed.warnings.push(`${key} row ${row}: ${error}`);
        }
      });
    });

    targets.slice().reverse().forEach(row => sheet.deleteRow(row));
    removed.rows[key] = targets;
  });

  // The registry learns from every entry, including these, so the supplier it
  // taught has to go too or "Smoke Test Ltd" autocompletes forever.
  const registry = getOrCreateRegistrySheet();
  const registryRow = loadRegistry().find(
    entry => normalizeName(entry.name) === normalizeName(SMOKE_COUNTERPARTY)
  );
  if (registryRow) {
    registry.deleteRow(registryRow.row);
    removed.registry = SMOKE_COUNTERPARTY;
  }

  Logger.log(JSON.stringify(removed, null, 2));
  return removed;
}

/* =============================== Full reset =============================== */

/**
 * The exact string resetAllData() demands before it will do anything.
 *
 * A confirmation dialog is not available here - this runs from the editor, where
 * the last function you picked is one click from running again. So the safeguard
 * is structural, in the same spirit as hardDeleteEntry() only being able to see
 * the archive: the destructive path cannot be reached by clicking Run, only by
 * typing this out.
 */
const RESET_CONFIRMATION = 'DELETE ALL TEST DATA';

/**
 * Trash the documents a sheet's rows point at, then delete the rows.
 *
 * Documents first, deliberately: the row is the only record of which file
 * belongs to it, so deleting rows first would leave files that nothing can
 * identify. The reverse leaves rows pointing at trashed files, which is visible
 * and harmless for the two lines it survives. Same argument as archiveEntry's
 * write-before-remove.
 */
function _resetSheetData(sheet, section, report, label) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;

  const cols = resolveColumns(sheet);
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();

  (section.fileColumns || []).forEach(fileCol => {
    if (!cols[fileCol.header]) return;
    const at = cols[fileCol.header] - 1;
    values.forEach((rowValues, i) => {
      const fileId = extractFileId(rowValues[at]);
      if (!fileId) return;
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
        report.filesTrashed++;
      } catch (error) {
        // Already gone is the common case and not worth failing over, but it is
        // still reported: a reset that silently skipped files would be the same
        // class of lie as a status change that renames nothing.
        report.warnings.push(`${label} row ${i + 2}: ${error}`);
      }
    });
  });

  sheet.deleteRows(2, last - 1);
  return last - 1;
}

/**
 * Empty every sheet of data and trash the documents behind it. Headers, sheets
 * and folders all survive, so nothing needs re-bootstrapping afterwards.
 *
 * Written for Cutover step 1 in the case where the answer to "which of these are
 * test rows" turns out to be "all of them". findDebris() is the tool when the
 * sheet holds a mix; this one is for a clean slate.
 *
 * WHAT IT CLEARS
 *
 * Data rows in all four section sheets and any archive sheets that exist, the
 * documents those rows reference, the whole supplier registry, and everything in
 * the Staging folder. The Staging folder itself stays, because Genius Scan is
 * pointed at it by id.
 *
 * WHAT IT DOES NOT TOUCH
 *
 * Headers, sheet structure, the Drive folder tree, and Script Properties. So
 * bootstrap() does not need re-running and no id changes.
 *
 * @param {string} confirmation Must be exactly RESET_CONFIRMATION.
 * @return {Object} What was actually removed, also written to the log.
 */
function resetAllData(confirmation) {
  if (confirmation !== RESET_CONFIRMATION) {
    throw new Error(
      `resetAllData() refused: pass exactly resetAllData('${RESET_CONFIRMATION}')`
    );
  }

  const report = {
    rows: {}, archiveRows: {}, registry: 0,
    filesTrashed: 0, stagingTrashed: 0, warnings: []
  };

  Object.keys(SECTIONS).forEach(key => {
    const section = SECTIONS[key];
    report.rows[key] = _resetSheetData(getSheet(section), section, report, key);

    // getSheetByName rather than getArchiveSheet: the latter creates one, and a
    // reset that left four new empty sheets behind would be absurd.
    const archive = getSpreadsheet().getSheetByName(archiveSheetName(section));
    report.archiveRows[key] = archive
      ? _resetSheetData(archive, section, report, `${key} archive`)
      : 0;
  });

  const registry = getOrCreateRegistrySheet();
  const registryLast = registry.getLastRow();
  if (registryLast >= 2) {
    registry.deleteRows(2, registryLast - 1);
    report.registry = registryLast - 1;
  }

  // Ids collected before anything is trashed: a Drive iterator walking a folder
  // that is being emptied underneath it is not something to rely on.
  const stagingId = uiStagingFolderId();
  if (stagingId) {
    try {
      const ids = [];
      const iterator = DriveApp.getFolderById(stagingId).getFiles();
      while (iterator.hasNext()) ids.push(iterator.next().getId());
      ids.forEach(id => {
        try {
          DriveApp.getFileById(id).setTrashed(true);
          report.stagingTrashed++;
        } catch (error) {
          report.warnings.push(`staging ${id}: ${error}`);
        }
      });
    } catch (error) {
      report.warnings.push(`staging folder: ${error}`);
    }
  }

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/* ============================== Debris audit ============================== */

/**
 * Report rows and registry entries that look like they were left behind by
 * testing. Run from the editor before cutover; see REBUILD-PLAN's Cutover step
 * 1, "delete the test rows".
 *
 * WHY THIS REPORTS AND NEVER DELETES
 *
 * A part-filled row awaiting a document is exactly what a legitimate deferred
 * entry looks like. That is not a flaw in the schema, it is the point of
 * deferred entries - so no test this function could apply separates the two with
 * certainty, and a cleanup that guessed would destroy real claims. It prints;
 * you decide; the sheet or archiveEntry() does the removing.
 *
 * smokeCleanup() can delete safely because it only matches rows carrying
 * SMOKE_MARKER in Notes, which it wrote itself. Nothing marks a row abandoned by
 * a half-built Shortcut, which is why that function is useless here and this one
 * exists.
 *
 * THE SIGNALS, AND HOW MUCH THEY ARE WORTH
 *
 * `certain` - no counterparty, or no usable amount. Both intake paths always set
 * both: Siri asks for them before it will call create, and the form requires
 * them. A row missing either was written by a run that failed partway, so it is
 * debris whatever else is true of it.
 *
 * `suspect` - complete enough to be real, but Siri-sourced and still awaiting a
 * document with no category set. Ordinary for a genuine deferred entry, which is
 * why it is only ever a prompt to look. Narrow it with `sinceIso`.
 *
 * WHAT IT CANNOT SEE, WHICH MATTERS BEFORE CUTOVER
 *
 * A test run that SUCCEEDED writes a complete, well-formed row. Nothing about
 * "Bolt, 8 EUR, taxi" says whether it came from a real taxi or from proving a
 * Shortcut works, so this function is silent about it by design - the same
 * reason it never deletes. An empty report therefore means "no malformed rows",
 * NOT "the sheet is ready for cutover". Read the sheet for that.
 *
 * WHY IT COUNTS WHAT IT SCANNED
 *
 * `scanned` carries the denominator, because an all-zero report is otherwise
 * ambiguous in the worst way: "looked at forty rows and they are all fine" and
 * "looked at nothing" print identically, and the second reads as the first. Per
 * section it gives `rows` present and `considered` after `sinceIso` is applied,
 * so a date filter that excluded everything is visible rather than silent.
 *
 * @param {string} [sinceIso] Only consider rows created on or after this date,
 *   e.g. '2026-08-13'. Rows with an unreadable timestamp are always included -
 *   a filter that silently dropped them would hide the worst-formed rows, which
 *   are the ones most likely to be debris.
 * @return {Object} The report, also written to the log.
 */
function findDebris(sinceIso) {
  let since = null;
  if (sinceIso) {
    since = new Date(sinceIso);
    if (isNaN(since.getTime())) throw new Error(`Unreadable date: ${sinceIso}`);
  }

  const report = {
    since: sinceIso || '(everything)',
    summary: '',
    scanned: {},
    sections: {},
    registry: [],
    totals: { certain: 0, suspect: 0, registry: 0 }
  };

  Object.keys(SECTIONS).forEach(key => {
    const section = SECTIONS[key];
    const sheet = getSheet(section);
    const findings = [];
    report.sections[key] = findings;

    // Recorded before the early return, so a section with no data rows reports
    // a hard zero rather than being absent from the report entirely.
    const counts = { rows: 0, considered: 0 };
    report.scanned[key] = counts;
    if (sheet.getLastRow() < 2) return;

    const cols = resolveColumns(sheet);
    const values = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .getValues();

    // Required columns go through columnIndex so a missing one is a loud error
    // rather than an audit that quietly finds nothing. The optional ones are
    // guarded instead: Income has no documents, so it has no receipt state, and
    // IVA has no category at all.
    const iTime = columnIndex(cols, sheet.getName(), COMMON.timestamp) - 1;
    const iParty = columnIndex(cols, sheet.getName(), COMMON.counterparty) - 1;
    const iAmount = columnIndex(cols, sheet.getName(), COMMON.amount) - 1;
    const iSource = columnIndex(cols, sheet.getName(), COMMON.source) - 1;
    const iState = cols[COMMON.receiptState] ? cols[COMMON.receiptState] - 1 : -1;
    const iCategory = section.category && cols[section.category.header]
      ? cols[section.category.header] - 1
      : -1;

    counts.rows = values.length;

    values.forEach((rowValues, i) => {
      const created = rowValues[iTime];
      const readable = created instanceof Date && !isNaN(created.getTime());
      if (since && readable && created < since) return;
      counts.considered++;

      const party = (rowValues[iParty] || '').toString().trim();
      const amount = Number(rowValues[iAmount]);
      const reasons = [];

      if (!party) reasons.push('no counterparty');
      if (!isFinite(amount) || amount === 0) reasons.push('no usable amount');

      if (!reasons.length) {
        const awaiting = iState >= 0 &&
          (rowValues[iState] || '').toString().trim() === RECEIPT_STATE.awaiting;
        const siri = (rowValues[iSource] || '').toString().trim() === 'siri';
        const noCategory = iCategory >= 0 &&
          !(rowValues[iCategory] || '').toString().trim();
        if (siri && awaiting && noCategory) reasons.push('siri, awaiting, no category');
        else return;
      }

      const certain = reasons[0] !== 'siri, awaiting, no category';
      findings.push({
        row: i + 2,
        confidence: certain ? 'certain' : 'suspect',
        counterparty: party || '(blank)',
        amount: rowValues[iAmount],
        created: readable ? created : '(unreadable)',
        reasons: reasons
      });
      report.totals[certain ? 'certain' : 'suspect']++;
    });
  });

  // A supplier learned from a test run autocompletes forever, so it is as much
  // debris as the row that taught it. timesUsed <= 1 over-reports by design: a
  // genuine one-off supplier looks identical, and under-reporting here means
  // junk survives cutover.
  const entries = loadRegistry();
  const registryCounts = { rows: entries.length, considered: 0 };
  report.scanned.registry = registryCounts;

  // The date guard runs before the timesUsed one so `considered` means the same
  // thing here as it does for a section: survived `sinceIso`. timesUsed is the
  // finding criterion, not a narrowing of scope.
  entries.forEach(entry => {
    const last = entry.lastUsed instanceof Date ? entry.lastUsed : null;
    if (since && last && last < since) return;
    registryCounts.considered++;
    if (entry.timesUsed > 1) return;
    report.registry.push({
      row: entry.row,
      name: entry.name,
      timesUsed: entry.timesUsed,
      lastUsed: entry.lastUsed
    });
  });
  report.totals.registry = report.registry.length;

  const rows = Object.keys(SECTIONS)
    .reduce((sum, key) => sum + report.scanned[key].rows, 0);
  const considered = Object.keys(SECTIONS)
    .reduce((sum, key) => sum + report.scanned[key].considered, 0);
  const found = report.totals.certain + report.totals.suspect + report.totals.registry;

  // Spelled out in prose because the caller is a human reading the editor log,
  // and the distinction this makes is the one the numbers alone lost.
  if (!rows && !registryCounts.rows) {
    report.summary =
      'Nothing to audit: no data rows in any section and an empty registry. ' +
      'This is not a clean bill of health, it is an empty sheet.';
  } else if (!considered && !registryCounts.considered) {
    report.summary =
      `Nothing considered: ${rows} row(s) and ${registryCounts.rows} registry ` +
      `entr(ies) exist, but all fall before ${report.since}.`;
  } else if (!found) {
    report.summary =
      `Scanned ${considered} of ${rows} row(s) and ${registryCounts.considered} ` +
      `of ${registryCounts.rows} registry entr(ies): no malformed rows. ` +
      'Complete rows left by successful test runs are invisible here - read the sheet.';
  } else {
    report.summary =
      `Scanned ${considered} of ${rows} row(s) and ${registryCounts.considered} ` +
      `of ${registryCounts.rows} registry entr(ies): ${report.totals.certain} certain, ` +
      `${report.totals.suspect} suspect, ${report.totals.registry} registry.`;
  }

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
