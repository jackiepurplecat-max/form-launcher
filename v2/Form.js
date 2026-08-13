/**
 * v2 — the custom form: the server side of creating an entry from the web app.
 *
 * This is step 8, and the reason Google Forms was dropped. A Form cannot fill
 * one answer from another, so "type FNAC, get its NIF" is impossible there and
 * is most of what this file exists to do.
 *
 * Everything here is derived from SECTIONS. There is no per-section code: the
 * fields, their order, which are required, which offer a managed list and which
 * accept a document all come from config, so adding a field stays a config
 * change plus a re-run of bootstrap().
 *
 * WHAT THIS FILE DOES NOT TRUST
 *
 * The page is signed in and restricted, and it is still not trusted:
 *
 *   - Headers are whitelisted against the field list this file generates. An
 *     unknown header is refused by name rather than dropped, because a silently
 *     ignored field is data you believed you had entered.
 *   - A value for a FILE column is refused outright. extractFileId will take a
 *     Drive ID out of any string and the script runs as you, so accepting one
 *     would let a caller have a file of yours renamed and moved into
 *     HelpfulForms. Documents arrive as uploads and nowhere else. The same rule
 *     is written down for doPost in REBUILD-PLAN.md; it is not weaker here just
 *     because the caller signed in.
 *   - Every value still goes through createEntry, so safeCellValue still runs.
 *
 * ORDER OF OPERATIONS, AND WHY
 *
 * Uploads happen BEFORE the row is created, and a failed upload aborts the
 * whole thing. You are standing there holding the receipt: telling you the
 * upload failed is better than creating a row without it and mailing you a
 * completion request about a file you already have. If the row then fails to
 * create, the uploads are trashed — otherwise they would sit in Drive with
 * nothing pointing at them, which is exactly the orphan state checkDocuments()
 * exists to find.
 */

/** Refuse anything larger than this. Receipts are small; a video is a mistake. */
const MAX_UPLOAD_BYTES = 10 * 1024 * 1024;

/**
 * Extension to give an upload that arrived without one.
 *
 * iOS shares images with no extension often enough to matter, and a file with
 * no extension is a real problem here rather than a cosmetic one: the rename
 * chain carries the extension over from the original name, so one lost at
 * upload stays lost through every transition.
 */
const UPLOAD_EXTENSIONS = {
  'application/pdf': '.pdf',
  'image/jpeg': '.jpg',
  'image/png': '.png',
  'image/heic': '.heic',
  'image/heif': '.heic',
  'image/webp': '.webp',
  'text/plain': '.txt'
};

/* ============================== Field list ================================ */

/**
 * The form's fields, in the order they are asked for, derived from SECTIONS.
 *
 * `required` here MUST agree with missingFields() in Entries.js — the form
 * marking a field optional that the server then reports missing would produce
 * an entry that reports itself incomplete the moment it is made. The harness
 * asserts the two agree rather than trusting this comment.
 *
 * `role` is the only behavioural hint the client gets:
 *   counterparty  offers registry autocomplete, and prefills from a match
 *   category      offers the values already in use
 */
function uiFormFields(section) {
  const fields = [
    { header: COMMON.date, label: 'Date', type: 'date', required: true, role: null },
    {
      header: COMMON.counterparty, label: counterpartyLabel(section),
      type: 'text', required: true, role: 'counterparty'
    }
  ];

  if (section.category) {
    // A category with a declared option list is a closed choice - Health's
    // Patient. One without is free text with suggestions, because Work's
    // Expense Reason is a new trip most times it is asked for.
    const declared = categoryOptions(section);
    const closed = declared !== null;
    fields.push({
      header: section.category.header,
      label: section.category.label,
      type: closed ? 'choice' : 'text',
      required: section.category.required !== false,
      options: closed ? declared : null,
      role: closed ? null : 'category'
    });
  }

  section.extraFields.forEach(field => {
    fields.push({
      header: field.header,
      label: field.label,
      type: field.type,
      required: !!field.required,
      options: field.options || null,
      role: null
    });
  });

  fields.push({ header: COMMON.amount, label: 'Amount', type: 'number', required: true, role: null });
  fields.push({
    header: COMMON.currency, label: 'Currency', type: 'text',
    required: false, role: null, defaultValue: DEFAULT_CURRENCY
  });

  // Income's three dates are business facts, not just bookkeeping - an invoice
  // is usually backdated, and money can arrive before the row is made. So they
  // are ordinary fields here as well as being filled by the status control.
  // Work, IVA and Health do not get theirs: a new entry is in the first state,
  // and setStatus clears the dates of every state after the target, so a
  // Claimed Date typed at creation would be wiped by the first transition.
  if (section.stateDatesInForm) {
    section.states.forEach(state => {
      if (!state.dateColumn) return;
      fields.push({
        header: state.dateColumn, label: state.dateColumn,
        type: 'date', required: false, role: null
      });
    });
  }

  fields.push({ header: COMMON.notes, label: 'Notes', type: 'text', required: false, role: null });

  section.fileColumns.forEach(fileCol => {
    fields.push({
      header: fileCol.header, label: fileCol.label,
      type: 'file', required: false, role: null
    });
  });

  return fields;
}

/** Headers the form may write. Files are excluded: they arrive as uploads. */
function uiWritableHeaders(section) {
  return uiFormFields(section)
    .filter(field => field.type !== 'file')
    .map(field => field.header);
}

/**
 * Values already in use in a section's category column, most used first.
 *
 * INTERIM, and deliberately so. The plan calls for a managed list, and step 9
 * builds add / remove on top. Until then the list populates itself from what
 * you actually enter, exactly as the supplier registry does — so adding a
 * patient is typing it once, and there is no list to maintain before the form
 * is usable. Free text stays allowed, which is what makes the first use of a
 * new value possible at all.
 */
function uiCategoryValues(sectionKey) {
  requireUiAccess();
  return categoryValues(sectionKey);
}

/*
 * The same list, without the UI gate, for callers that run their own check.
 * The Siri endpoint needs it to offer Patient as a tap rather than dictation,
 * and it authenticates by key rather than by Google sign-in — so the gate has
 * to sit in the ui* wrapper, not in here. Nothing else may call this without
 * checking its caller first.
 */
function categoryValues(sectionKey) {
  const section = getSection(sectionKey);
  if (!section.category) return [];

  // A declared list is the answer, and the only answer. Appending whatever
  // happens to be in the column would quietly re-open a list whose whole
  // purpose is being closed - a misspelling already in the sheet would come
  // back as a suggestion and get picked again.
  const declared = categoryOptions(section);
  if (declared !== null) return declared;

  const sheet = getSheet(section);
  const cols = resolveColumns(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const column = columnIndex(cols, sheet.getName(), section.category.header);
  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();

  const counts = {};
  values.forEach(cell => {
    const value = (cell[0] === null || cell[0] === undefined) ? '' : cell[0].toString().trim();
    if (value) counts[value] = (counts[value] || 0) + 1;
  });

  return Object.keys(counts).sort((a, b) => counts[b] - counts[a] || a.localeCompare(b));
}

/* ============================== The registry ============================== */

/** Autocomplete as you type. Names only — nothing is filled in from this. */
function uiSuggestCounterparty(prefix, limit) {
  requireUiAccess();
  return suggestSuppliers(prefix || '', limit || 8);
}

/**
 * What the registry knows about a name, and what it is confident enough to fill.
 *
 * Returns the match either way. Below the confidence bar `autofill` is false and
 * `prefill` is empty: a wrong NIF means a rejected claim, which is worse than a
 * blank field, so the page shows what it suspects and fills nothing.
 */
function uiLookupCounterparty(sectionKey, name) {
  requireUiAccess();
  getSection(sectionKey);  // reject an unknown section before touching the sheet
  return lookupCounterparty(sectionKey, name || '');
}

/* ================================ Uploads ================================= */

/** A name safe to put in Drive, with its extension kept or supplied. */
function uploadFilename(name, mimeType) {
  const cleaned = (name || '')
    .toString()
    .replace(/[\\/]/g, '_')                // no path separators
    .replace(/[\x00-\x1f\x7f]/g, '')       // no control characters
    .trim() || 'upload';

  if (splitExtension(cleaned).ext) return cleaned;
  return cleaned + (UPLOAD_EXTENSIONS[(mimeType || '').toLowerCase()] || '');
}

/* ============================= Staging folder ============================= */

/**
 * The folder scans and saved mail attachments land in before they belong to an
 * entry.
 *
 * WHY PICKING BEATS UPLOADING. A document that is already in Drive does not need
 * uploading again — `nameAndFileDocuments()` works from a file ID and does not
 * care where the file came from, so choosing one MOVES and renames it into the
 * HelpfulForms tree and it leaves this folder by itself. Uploading a second copy
 * would instead leave the original behind for ever, doubling storage against a
 * quota managed by hand, and leaving a folder that only grows — which destroys
 * the one useful property it has: that what is in it is what is not yet filed.
 */
function uiStagingFolderId() {
  return PropertiesService.getScriptProperties().getProperty(STAGING_FOLDER_PROPERTY);
}

/**
 * What is waiting to be filed. Empty list when no folder is configured, rather
 * than an error: the picker is an addition, and the form must still work
 * without it.
 */
function uiStagingFiles() {
  requireUiAccess();

  const folderId = uiStagingFolderId();
  if (!folderId) return [];

  const files = [];
  const iterator = DriveApp.getFolderById(folderId).getFiles();
  while (iterator.hasNext()) {
    const file = iterator.next();
    files.push({ id: file.getId(), name: file.getName() });
  }

  return files.sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Turn a picked id into a file, refusing anything not actually in the staging
 * folder.
 *
 * The check is the point. `extractFileId` will take a Drive id out of any
 * string, and this script runs as me — so without it a stale or mistyped id
 * would have some unrelated file of mine renamed and moved into HelpfulForms.
 * The web UI is gated to my own account, so this is guarding against an accident
 * rather than an attacker, but it is the same accident either way.
 */
function uiResolveStagingPick(fileId) {
  const wanted = extractFileId((fileId || '').toString());
  if (!wanted) throw new Error('No file was chosen.');

  const folderId = uiStagingFolderId();
  if (!folderId) {
    throw new Error(`${STAGING_FOLDER_PROPERTY} is not set in Script Properties.`);
  }

  const iterator = DriveApp.getFolderById(folderId).getFiles();
  while (iterator.hasNext()) {
    const file = iterator.next();
    if (file.getId() === wanted) return file;
  }

  throw new Error('That file is not in the staging folder. Refresh the list and try again.');
}

/**
 * Documents for one submission, from uploads and picks together.
 *
 * `picked` is tracked per file and it matters: the callers trash everything in
 * this list if the write then fails, which is right for an upload they just
 * created and WRONG for a pick — that is the original, sitting in the staging
 * folder, and trashing it would destroy the only copy.
 */
function uiCollectDocuments(section, uploads, picks) {
  const collected = [];

  // Cleans up after ITSELF. A second upload failing must not strand the first
  // one, which has already been created in the section inbox — an unreferenced
  // file in the tree is the orphan state that cannot be told from a live
  // document by looking. The callers only see all-or-nothing.
  try {
    (uploads || []).forEach(upload => {
      collected.push({ header: upload.header, file: uiStoreUpload(section, upload), picked: false });
    });

    (picks || []).forEach(pick => {
      const header = (pick && pick.header) || '';
      if (!section.fileColumns.some(col => col.header === header)) {
        throw new Error(`Not a document column for ${section.sheet}: "${header}"`);
      }
      collected.push({ header: header, file: uiResolveStagingPick(pick.id), picked: true });
    });
  } catch (error) {
    uiDiscardDocuments(collected);
    throw error;
  }

  return collected;
}

/** Undo stored documents after a failed write. Never touches a picked file. */
function uiDiscardDocuments(collected) {
  collected.forEach(item => {
    if (item.picked) return;
    try { item.file.setTrashed(true); } catch (ignored) { /* best effort */ }
  });
}

/**
 * Put one uploaded document in the section's inbox and return its file.
 *
 * Throws rather than returning a failure, because the caller aborts the whole
 * creation on any upload problem — see the note at the top of this file.
 */
function uiStoreUpload(section, upload) {
  const header = (upload && upload.header) || '';
  const fileCol = section.fileColumns.filter(col => col.header === header)[0];
  if (!fileCol) {
    throw new Error(`"${header}" is not a document for ${section.sheet}`);
  }

  const data = (upload.data || '').toString();
  if (!data) throw new Error(`${fileCol.label}: no file content arrived`);

  // Checked before decoding: base64 is about 4/3 of the bytes it represents, so
  // this refuses an oversized upload without first materialising it in memory.
  const approximateBytes = Math.floor(data.length * 3 / 4);
  if (approximateBytes > MAX_UPLOAD_BYTES) {
    throw new Error(
      `${fileCol.label}: ${Math.round(approximateBytes / 1024 / 1024)} MB is too large ` +
      `(limit ${Math.round(MAX_UPLOAD_BYTES / 1024 / 1024)} MB)`
    );
  }

  const mimeType = (upload.mimeType || 'application/octet-stream').toString();
  const blob = Utilities.newBlob(
    Utilities.base64Decode(data), mimeType, uploadFilename(upload.name, mimeType)
  );

  // Straight into the inbox. initializeEntry renames it from the row's own
  // values and applyFileState files it, so nothing here needs to know the
  // naming rules.
  return sectionFolder(section, INBOX_FOLDER).createFile(blob);
}

/* ============================= Validation ================================= */

/**
 * Check a submitted set of values against a field list, and throw on the first
 * problem with the field named.
 *
 * Shared by creating and editing on purpose. The plan's rule is that editing is
 * not a special mode — an edited row must never be able to be less valid than a
 * created one — and the only way to guarantee that is for both to run the same
 * function rather than two that look alike today.
 */
function validateSubmitted(section, submitted, fields) {
  const writable = fields.filter(f => f.type !== 'file').map(f => f.header);
  const fileHeaders = section.fileColumns.map(col => col.header);

  Object.keys(submitted).forEach(header => {
    if (fileHeaders.indexOf(header) !== -1) {
      throw new Error(`"${header}" is a document and cannot be set directly`);
    }
    if (writable.indexOf(header) === -1) {
      throw new Error(`"${header}" is not a field of ${section.sheet}`);
    }
  });

  // Dates are checked here rather than left to the sheet, so a typo is refused
  // with the field named instead of landing as text in a date column.
  fields.filter(field => field.type === 'date').forEach(field => {
    const value = (submitted[field.header] || '').toString().trim();
    if (value && !isValidDateISO(value)) {
      throw new Error(`${field.label} must be a valid yyyy-MM-dd date`);
    }
  });

  // A closed list is only closed if the server says so. The page renders a
  // dropdown, but google.script.run does not have to go through the page - and
  // the whole point of Patient being a list is that one misspelling would
  // silently become a second patient, splitting that person's claims in two.
  fields
    .filter(field => field.type === 'choice' && field.options && field.options.length)
    .forEach(field => {
      const value = (submitted[field.header] || '').toString().trim();
      if (value && field.options.indexOf(value) === -1) {
        throw new Error(
          `${field.label} must be one of: ${field.options.join(', ')} — got "${value}"`
        );
      }
    });
}

/* ============================== Create ==================================== */

/**
 * Create an entry from the form.
 *
 * Returns what actually happened, including `warnings` when the entry is
 * incomplete — createEntry still writes the row in that case, by design, since
 * a partial entry is the safety net rather than an error. The page says so
 * instead of showing a plain tick.
 */
function uiCreateEntry(sectionKey, payload) {
  requireUiAccess();

  const section = getSection(sectionKey);
  const submitted = (payload && payload.values) || {};
  const uploads = (payload && payload.files) || [];

  validateSubmitted(section, submitted, uiFormFields(section));

  const values = {};
  Object.keys(submitted).forEach(header => {
    const value = submitted[header];
    if (value === null || value === undefined || value.toString().trim() === '') return;
    values[header] = value;
  });

  const documents = uiCollectDocuments(section, uploads, (payload && payload.picked) || []);

  try {
    documents.forEach(item => { values[item.header] = item.file.getId(); });

    const result = createEntry(sectionKey, values, 'form');
    result.entry = uiEntry(sectionKey, result.row);
    return result;

  } catch (error) {
    // Nothing points at these now, and an unreferenced file in the tree is the
    // orphan state that is impossible to tell from a live document by looking.
    // A PICKED file is left alone — it is the original in the staging folder,
    // and trashing it would destroy the only copy of the receipt.
    uiDiscardDocuments(documents);
    throw error;
  }
}
