/**
 * v2 — editing a supplier, and repairing what its name is written into.
 *
 * WHY THIS EXISTS
 *
 * The registry populates itself, which means it also learns your typos. `whitee
 * clinic` was entered once and became a second supplier, splitting that
 * provider's history, with the misspelling baked into a receipt's filename. The
 * form now corrects a confident match before it is saved, but nothing repairs
 * what is already stored — and doing it by hand means editing a sheet cell, a
 * Drive filename and a registry row in three places and getting all three
 * consistent.
 *
 * So a name change here is not an edit to one cell. It is a rename across
 * everything derived from that name.
 *
 * THE DECISIONS, AND THE REASONS
 *
 * - **The repair goes through nameAndFileDocuments(), never through string
 *   surgery on the old filename.** That function rebuilds a document's name from
 *   the row's own values and is already shared by creation, editing and status
 *   changes. So the operation is: write the new name into the affected rows, then
 *   re-run it per row. Pattern-matching the old name inside the existing filename
 *   would be a second naming rule, free to drift from the first — and it would
 *   have to know that `White Clinic` is `WhiteClinic` in a filename.
 *
 * - **The rows are the index, not Drive.** Affected entries are found by their
 *   Counterparty column across all four sections; Drive is never searched by
 *   filename. The row holds the document's URL, so it can say which file to fix
 *   without guessing, and a file whose name was ALREADY wrong is still found.
 *
 * - **The archive sheets are included.** They carry the same spine, their
 *   documents sit in Archived, and restoreEntry rebuilds names from the row — so
 *   a rename that skipped them would sit quietly until a restore resurrected the
 *   old spelling. Their documents are renamed but NOT re-filed; see the
 *   folderName note in applyFileState.
 *
 * - **Merging is the common case, not renaming.** Correcting a typo usually means
 *   the target name already exists. A rename that silently created a THIRD
 *   supplier would be the original bug with more steps.
 *
 * - **On a merge the surviving spelling is the target's, not the one typed.**
 *   You are folding a typo into an established supplier, so the established name
 *   wins and only the typo's rows are touched. Editing the established spelling
 *   is a separate, deliberate act: open that supplier and rename it.
 *
 * - **Offer the old spelling as an alias, do not add it.** A real recurring
 *   mishearing should resolve at 0.95 forever; a one-off typo should not be
 *   taught. Only you know which it was, so the result says what could be added
 *   and addSupplierAlias() is a second, separate call.
 *
 * - **A merge keeps the CORE entry's NIF, and says so loudly when that matters.**
 *   The core is the supplier being merged into — the established record, against
 *   a row that is by assumption a typo. Choosing between the two by hand was
 *   considered and rejected: the default is right almost every time, and a picker
 *   on every merge is a decision you would learn to click through. So it defaults
 *   and warns instead, both before the merge and after, in the two cases where
 *   the core ends up holding a number nobody checked — the NIFs disagreed, or the
 *   core had none and inherited the typo's. See mergeSupplierNif.
 *
 * - **Report per row and per document.** A rename can fail halfway through, so
 *   this returns what actually happened rather than a tick.
 *
 * A CORRECTED NIF IS NEVER BACKDATED
 *
 * Correcting a supplier's NIF does NOT rewrite `Emitente NIF` on past IVA
 * entries, and that is a decision rather than an omission. A submitted claim is a
 * record of WHAT WAS SUBMITTED, so rewriting the figure afterwards makes the row
 * disagree with what Finanças actually received — worse than a row that is merely
 * out of date, because it destroys the evidence of what happened. The registry
 * value changes, so every FUTURE entry prefills correctly, and that is the whole
 * benefit available without touching history.
 *
 * The harness pins it. Do not "improve" this into a repair.
 */

/**
 * Rows repaired in a single run.
 *
 * Apps Script kills an execution at six minutes. Each row costs one cell write
 * plus up to two Drive renames and two moves, so a supplier with hundreds of
 * entries could reach that limit and be discovered as a timeout — which is the
 * one outcome with no report attached.
 *
 * So the work stops at a known point instead. When it does, the REGISTRY IS LEFT
 * UNTOUCHED and the result says how many rows remain: the rows already done no
 * longer carry the old name, the ones left still do, and running the same edit
 * again continues from there. Partial progress is the point, not a failure mode.
 */
const SUPPLIER_REPAIR_ROW_LIMIT = 50;

/* ============================ Finding the rows ============================ */

/**
 * Every sheet a counterparty name can be written into: each section's live sheet
 * and its archive.
 *
 * A missing archive sheet is reported rather than thrown, because the live rows
 * are still worth repairing and bootstrap() is the fix. A missing LIVE sheet is
 * fatal — that is a broken installation, not a gap.
 */
function supplierScanTargets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [];
  const skipped = [];

  Object.keys(SECTIONS).forEach(sectionKey => {
    const section = SECTIONS[sectionKey];
    targets.push({
      sectionKey: sectionKey, section: section,
      sheet: getSheet(section), archived: false
    });

    const archiveName = archiveSheetName(section);
    const archive = ss.getSheetByName(archiveName);
    if (archive) {
      targets.push({
        sectionKey: sectionKey, section: section,
        sheet: archive, archived: true
      });
    } else {
      skipped.push(archiveName);
    }
  });

  return { targets: targets, skipped: skipped };
}

/**
 * Rows in one sheet whose Counterparty is the given name.
 *
 * Compared through normalizeName, so casing, accents and punctuation do not
 * hide a row: correcting `whitee clinic` must also catch `Whitee Clinic`. The
 * column is read in one range rather than cell by cell — this runs over eight
 * sheets and the scan has to be cheap enough to do before asking for
 * confirmation.
 */
function supplierRowsIn(target, normalized) {
  const sheet = target.sheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const cols = resolveColumns(sheet);
  const index = columnIndex(cols, sheet.getName(), COMMON.counterparty);
  const values = sheet.getRange(2, index, lastRow - 1, 1).getValues();

  const rows = [];
  values.forEach((cell, i) => {
    if (normalizeName(cell[0]) === normalized) rows.push(i + 2);
  });
  return rows;
}

/**
 * Every entry row using a supplier, across every section, live and archived.
 *
 * This is what the confirmation step shows: you see the blast radius before
 * anything is written, which is the whole difference between a rename and an
 * accident.
 */
function findSupplierEntries(name) {
  const normalized = normalizeName(name);
  // A blank normalises to '', which would match every row whose Counterparty is
  // empty. That is the one input that must never reach the scan.
  if (!normalized) throw new Error('A supplier name is required');

  const scan = supplierScanTargets();
  const matches = [];
  const bySection = {};

  scan.targets.forEach(target => {
    const rows = supplierRowsIn(target, normalized);
    if (!rows.length) return;

    if (!bySection[target.sectionKey]) {
      bySection[target.sectionKey] = { section: target.sectionKey, live: 0, archived: 0 };
    }
    bySection[target.sectionKey][target.archived ? 'archived' : 'live'] += rows.length;

    rows.forEach(row => matches.push({ target: target, row: row }));
  });

  return {
    name: name,
    matches: matches,
    total: matches.length,
    bySection: Object.keys(bySection).map(key => bySection[key]),
    skippedSheets: scan.skipped
  };
}

/* =============================== The repair =============================== */

/**
 * Write a supplier name into every row that uses it, and rebuild the documents.
 *
 * writeName may equal matchName, and that is not a wasted call: it re-runs the
 * naming rule over rows whose documents are stale for any reason, which is what
 * makes "run it again" a real repair rather than advice. A rename is the same
 * operation with a different name written.
 *
 * A document failure does NOT stop the row: the sheet is the record, and a file
 * left with an old name is visible, reported and fixable by running this again.
 * A failure to write the CELL does stop that row, and is what makes the caller
 * leave the registry alone.
 */
function applySupplierToEntries(matchName, writeName, limit) {
  const found = findSupplierEntries(matchName);
  const cap = limit || SUPPLIER_REPAIR_ROW_LIMIT;
  const doing = found.matches.slice(0, cap);

  const rows = [];
  doing.forEach(match => {
    const target = match.target;
    const sheet = target.sheet;
    const cols = resolveColumns(sheet);

    const record = {
      section: target.sectionKey,
      sheet: sheet.getName(),
      archived: target.archived,
      row: match.row,
      ok: false,
      renames: [],
      files: []
    };

    // The CELL is what "this row is done" means, and it is tracked separately
    // from the documents on purpose. record.ok gates the registry, and the
    // registry may move as soon as every row carries the new name — a document
    // that could not be renamed is reported and repaired by running again, in
    // exactly the way a file failure never rolls back a status change. Reporting
    // the row as failed because of a Drive error would block the registry
    // forever on something re-running cannot fix.
    try {
      writeCell(sheet, cols, match.row, COMMON.counterparty, writeName);
      record.ok = true;
    } catch (error) {
      record.error = error.toString();
      rows.push(record);
      return;
    }

    try {
      // The status decides the suffix chain, exactly as it does on a restore. An
      // off-vocabulary status falls back to the first state rather than throwing:
      // the name is already corrected either way, and a bad status is a separate
      // problem visible in the table.
      const status = (readCell(sheet, cols, match.row, COMMON.status) || '').toString().trim();
      const stateAt = Math.max(0, stateIndex(target.section, status));

      const documents = nameAndFileDocuments(
        target.section, sheet, cols, match.row, stateAt,
        target.archived ? ARCHIVE_FOLDER : null
      );
      record.renames = documents.renames;
      record.files = documents.files;

    } catch (error) {
      // nameAndFileDocuments reports per-file rather than throwing, so reaching
      // here means the ROW could not be read at all - a missing Status column,
      // say. Recorded against the row's documents, not against the row.
      record.documentError = error.toString();
    }

    rows.push(record);
  });

  SpreadsheetApp.flush();

  const rowErrors = rows.filter(record => !record.ok);
  const documentErrors = [];
  rows.forEach(record => {
    if (record.documentError) {
      documentErrors.push({
        sheet: record.sheet, row: record.row, column: '', error: record.documentError
      });
    }
    record.renames.concat(record.files)
      .filter(result => !result.ok)
      .forEach(result => documentErrors.push({
        sheet: record.sheet, row: record.row,
        column: result.column, error: result.error
      }));
  });

  return {
    from: matchName,
    to: writeName,
    rows: rows,
    rowsChanged: rows.filter(record => record.ok).length,
    documentsTouched: rows.reduce((total, record) => total + record.files.length, 0),
    rowErrors: rowErrors,
    documentErrors: documentErrors,
    // Everything the caller needs to decide whether the registry may now change.
    remaining: Math.max(0, found.total - doing.length),
    complete: found.total <= doing.length && rowErrors.length === 0,
    skippedSheets: found.skippedSheets
  };
}

/* ============================== The registry ============================== */

/**
 * Split the aliases field back into a list.
 *
 * Aliases share one comma-separated cell, so a comma inside one would come back
 * as two. Splitting on commas is therefore the definition rather than a parse of
 * it, and addSupplierAlias() rejects a comma for the same reason.
 */
function parseAliasList(value) {
  const seen = {};
  return (value === null || value === undefined ? '' : value)
    .toString()
    .split(',')
    .map(alias => alias.trim())
    .filter(alias => {
      if (!alias) return false;
      const key = normalizeName(alias);
      if (!key || seen[key]) return false;
      seen[key] = true;
      return true;
    });
}

/**
 * Reconcile two stored Types when suppliers merge.
 *
 * Same rule recordSupplier applies when it sees a second, different type: stop
 * guessing rather than guess wrongly. A field that prefills the wrong value is
 * worse than one that prefills nothing.
 */
function mergeSupplierType(a, b) {
  const left = (a || '').toString().trim();
  const right = (b || '').toString().trim();
  if (!left) return { value: right, cleared: false };
  if (!right) return { value: left, cleared: false };
  if (left === right) return { value: left, cleared: false };
  return { value: '', cleared: true };
}

/**
 * Which NIF a merge keeps, and what is worth warning about.
 *
 * Defaults to the CORE entry's — the supplier being merged into — because that is
 * the established record and the other row is, by assumption, a typo. It is never
 * cleared, unlike Type: a supplier really can have two types and cannot have two
 * NIFs.
 *
 * Both ways the core can end up holding a value you did not check are reported,
 * because a wrong NIF is a rejected claim and that makes a silent change to the
 * core the one outcome worth being noisy about:
 *
 *   kept     the two disagreed. The core's survived; the other is named so you
 *            can decide whether the core was the right one to trust.
 *   adopted  the core had NO NIF, so it has just inherited the typo's. If that
 *            number was wrong, the core is now wrong — and nothing said so
 *            before this existed.
 *
 * Shared by updateSupplier and uiSupplierPreview on purpose, so the warning shown
 * BEFORE a merge is produced by the rule that runs DURING it rather than by a
 * second copy free to drift.
 */
function mergeSupplierNif(submitted, target, targetName) {
  const mine = (submitted || '').toString().trim();
  const theirs = (target || '').toString().trim();

  if (!theirs && mine) {
    return { value: mine, kept: null, adopted: { value: mine, into: targetName } };
  }
  if (theirs && mine && mine !== theirs) {
    return {
      value: theirs,
      kept: { kept: theirs, discarded: mine, into: targetName },
      adopted: null
    };
  }
  return { value: theirs || mine, kept: null, adopted: null };
}

/** The later of two Last Used values, either of which may be blank or junk. */
function laterDate(a, b) {
  const left = a ? new Date(a).getTime() : NaN;
  const right = b ? new Date(b).getTime() : NaN;
  if (isNaN(left)) return isNaN(right) ? '' : b;
  if (isNaN(right)) return a;
  return left >= right ? a : b;
}

/**
 * Edit one supplier, and propagate a name change to everything derived from it.
 *
 * ORDER OF OPERATIONS, and why it is this way round: the ENTRIES are repaired
 * first, and the registry is only changed once every one of them succeeded.
 *
 * A merge deletes the source registry row, and once it is gone there is no
 * supplier left to open and no edit left to re-run. So if the repair stops at the
 * row limit, or any row refuses to be written, the registry keeps both entries
 * and the caller is told to run it again — the rows already renamed no longer
 * match, so a second pass picks up exactly where the first stopped. The
 * intermediate state is two registry entries and some rows moved, which is
 * visible and self-correcting. The alternative loses the ability to finish.
 *
 * Document failures do not block the registry: the sheet values are right, and
 * uiRepairSupplierDocuments() re-runs the naming over the surviving name.
 *
 * Type, NIF and Aliases are taken from the payload as given, so a blank CLEARS
 * the field — the same rule uiUpdateEntry follows, and the only way to empty one.
 * The form therefore sends all four fields, never a subset.
 *
 * @param {number} sheetRow row in the Suppliers sheet
 * @param {Object} payload  { name, type, nif, aliases, was }
 * @param {number} [limit]  rows to repair in this run. Defaults to
 *   SUPPLIER_REPAIR_ROW_LIMIT; uiUpdateSupplier never passes it, so the page
 *   cannot choose its own batch size. Exists so the incomplete path is testable
 *   without inventing fifty entries.
 */
function updateSupplier(sheetRow, payload, limit) {
  const submitted = payload || {};
  const newName = (submitted.name || '').toString().trim();
  if (!newName) throw new Error('Name is required');

  const sheet = getOrCreateRegistrySheet();
  const row = resolveDataRow(sheet, sheetRow);
  const registry = loadRegistry();

  const source = registry.filter(entry => entry.row === row)[0];
  if (!source) throw new Error(`Row ${row} of ${REGISTRY_SHEET} is not a supplier`);

  // Rows shift when one is deleted, so the row number alone is not proof that
  // this is the supplier the form loaded. Renaming the wrong one would be
  // silent and would take its documents with it.
  const was = (submitted.was || '').toString().trim();
  if (was && normalizeName(was) !== normalizeName(source.name)) {
    throw new Error(
      `Row ${row} now holds "${source.name}", not "${was}". Reload the list and try again.`
    );
  }

  const type = (submitted.type || '').toString().trim();
  const nif = (submitted.nif || '').toString().trim();
  const aliases = parseAliasList(submitted.aliases);

  // An alias that equals the supplier's own name is noise: findSupplier already
  // matches the name exactly, at a higher confidence than an alias.
  const ownName = normalizeName(newName);

  // A merge target may be matched by NAME or by ALIAS, because recordSupplier
  // resolves both — leaving a supplier whose name collides with another's alias
  // would make the registry ambiguous about which one it means.
  const target = registry.filter(entry =>
    entry.row !== source.row && (
      normalizeName(entry.name) === ownName ||
      entry.aliases.some(alias => normalizeName(alias) === ownName)
    )
  )[0] || null;

  // On a merge the target's spelling survives, so that is the name written into
  // the entry rows — not the string that was typed to find it.
  const finalName = target ? target.name : newName;
  const renaming = finalName !== source.name;

  const repair = renaming ? applySupplierToEntries(source.name, finalName, limit) : null;

  if (repair && !repair.complete) {
    return {
      ok: false,
      incomplete: true,
      error: repair.rowErrors.length
        ? `${repair.rowErrors.length} row(s) could not be renamed; the registry is unchanged.`
        : `Renamed ${repair.rowsChanged} of ${repair.rowsChanged + repair.remaining} entries. ` +
          `Save again to continue — the registry is unchanged until every row is done.`,
      merged: null,
      renamedTo: finalName,
      repair: repair,
      aliasOffer: null
    };
  }

  const registryResult = withLock(() => {
    const cols = resolveColumns(sheet);

    if (!target) {
      writeCell(sheet, cols, source.row, REGISTRY.name, newName);
      writeCell(sheet, cols, source.row, REGISTRY.type, type);
      writeCell(sheet, cols, source.row, REGISTRY.nif, nif);
      writeCell(sheet, cols, source.row, REGISTRY.aliases,
        aliases.filter(alias => normalizeName(alias) !== ownName).join(', '));
      SpreadsheetApp.flush();
      return {
        merged: null, row: source.row, name: newName,
        typeCleared: false, nifKept: null, nifAdopted: null
      };
    }

    // Merge. Times Used sums, Last Used keeps the later, aliases union, Type
    // follows the clear-on-conflict rule, and NIF is never cleared — it is a
    // fact about the supplier, so a conflict keeps the established value and
    // reports the one it displaced rather than silently choosing.
    const mergedType = mergeSupplierType(type, target.type);

    const targetName = normalizeName(target.name);
    const mergedAliases = [];
    const seen = {};
    target.aliases.concat(source.aliases).concat(aliases).forEach(alias => {
      const key = normalizeName(alias);
      if (!key || key === targetName || seen[key]) return;
      seen[key] = true;
      mergedAliases.push(alias);
    });

    const mergedNif = mergeSupplierNif(nif, target.nif, target.name);

    writeCell(sheet, cols, target.row, REGISTRY.type, mergedType.value);
    writeCell(sheet, cols, target.row, REGISTRY.nif, mergedNif.value);
    writeCell(sheet, cols, target.row, REGISTRY.aliases, mergedAliases.join(', '));
    writeCell(sheet, cols, target.row, REGISTRY.timesUsed, target.timesUsed + source.timesUsed);
    writeCell(sheet, cols, target.row, REGISTRY.lastUsed,
      laterDate(target.lastUsed, source.lastUsed));

    // Last, so a failure above leaves both rows and the edit can be re-run.
    sheet.deleteRow(source.row);
    SpreadsheetApp.flush();

    return {
      merged: { into: target.name, timesUsed: target.timesUsed + source.timesUsed },
      // The target's own row shifts up if the deleted source sat above it.
      row: source.row < target.row ? target.row - 1 : target.row,
      name: target.name,
      typeCleared: mergedType.cleared,
      nifKept: mergedNif.kept,
      nifAdopted: mergedNif.adopted
    };
  });

  return {
    ok: true,
    incomplete: false,
    section: null,
    row: registryResult.row,
    name: registryResult.name,
    merged: registryResult.merged,
    renamedFrom: renaming ? source.name : null,
    renamedTo: renaming ? finalName : null,
    typeCleared: registryResult.typeCleared,
    nifKept: registryResult.nifKept,
    nifAdopted: registryResult.nifAdopted,
    repair: repair,
    // Offered, never applied: only you know whether the old spelling was a
    // recurring mishearing worth teaching or a one-off typo worth forgetting.
    aliasOffer: renaming ? { name: registryResult.name, alias: source.name } : null,
    documentErrors: repair ? repair.documentErrors : []
  };
}

/* ============================== UI wrappers =============================== */

/**
 * The registry, most-used first — the same order the autocomplete offers, so the
 * list you manage reads like the list you are offered.
 */
function uiListSuppliers() {
  requireUiAccess();

  const suppliers = loadRegistry()
    .map(entry => ({
      row: entry.row,
      name: entry.name,
      type: entry.type,
      nif: entry.nif,
      aliases: entry.aliases.join(', '),
      timesUsed: entry.timesUsed,
      lastUsed: uiDateISO(entry.lastUsed)
    }))
    .sort((a, b) => b.timesUsed - a.timesUsed || a.name.localeCompare(b.name));

  return { ok: true, suppliers: suppliers };
}

/**
 * What saving this edit would do, before it does it.
 *
 * Called when the name has been changed, so the confirmation can name the
 * damage: how many entries in which sections, whether this is a merge rather
 * than a rename, and what a merge would do to the surviving NIF. Reads only.
 *
 * submittedNif is what the form currently holds, so the NIF verdict is computed
 * from the same values the save would use — through mergeSupplierNif, the same
 * function the merge itself calls.
 */
function uiSupplierPreview(sheetRow, newName, submittedNif) {
  requireUiAccess();

  const sheet = getOrCreateRegistrySheet();
  const row = resolveDataRow(sheet, sheetRow);
  const registry = loadRegistry();

  const source = registry.filter(entry => entry.row === row)[0];
  if (!source) throw new Error(`Row ${row} of ${REGISTRY_SHEET} is not a supplier`);

  const wanted = normalizeName(newName);
  const target = registry.filter(entry =>
    entry.row !== source.row && (
      normalizeName(entry.name) === wanted ||
      entry.aliases.some(alias => normalizeName(alias) === wanted)
    )
  )[0] || null;

  const found = findSupplierEntries(source.name);

  // Undefined means the caller did not say, in which case the stored value is
  // what a save would send - not a blank, which would read as "clear it".
  const nif = submittedNif === undefined || submittedNif === null
    ? source.nif
    : submittedNif;
  const nifOutcome = target ? mergeSupplierNif(nif, target.nif, target.name) : null;

  return {
    ok: true,
    from: source.name,
    to: target ? target.name : (newName || '').toString().trim(),
    merge: target
      ? { name: target.name, timesUsed: target.timesUsed, nif: target.nif, type: target.type }
      : null,
    nifKept: nifOutcome ? nifOutcome.kept : null,
    nifAdopted: nifOutcome ? nifOutcome.adopted : null,
    total: found.total,
    bySection: found.bySection,
    // A truthful warning rather than a refusal: the work will stop at the limit
    // and say so, and saving again continues it.
    willStopAt: found.total > SUPPLIER_REPAIR_ROW_LIMIT ? SUPPLIER_REPAIR_ROW_LIMIT : null,
    skippedSheets: found.skippedSheets
  };
}

function uiUpdateSupplier(sheetRow, payload) {
  requireUiAccess();
  return updateSupplier(sheetRow, payload);
}

/**
 * Rebuild every document name for a supplier, changing no data.
 *
 * This is what makes "run it again" the repair for a document that could not be
 * renamed during a merge — after the merge the old name is gone, so there is
 * nothing left to re-run the rename from. Also fixes a filename left stale by
 * anything else, since names are always rebuilt from the row.
 */
function uiRepairSupplierDocuments(sheetRow) {
  requireUiAccess();

  const sheet = getOrCreateRegistrySheet();
  const row = resolveDataRow(sheet, sheetRow);
  const source = loadRegistry().filter(entry => entry.row === row)[0];
  if (!source) throw new Error(`Row ${row} of ${REGISTRY_SHEET} is not a supplier`);

  return applySupplierToEntries(source.name, source.name);
}

/** Teach the registry a spelling, so the next mishearing resolves to this one. */
function uiAddSupplierAlias(name, alias) {
  requireUiAccess();
  return addSupplierAlias(name, alias);
}
