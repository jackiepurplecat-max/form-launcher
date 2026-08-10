/**
 * v2 — the web app: one page, listing and the status control.
 *
 * This is the FIRST code in v2 with a surface outside the editor. Everything
 * before it could only be reached by someone already in the Apps Script project;
 * from here on there is a URL. So the rules are stricter than elsewhere:
 *
 *   - Every function a browser can call checks the caller itself. The deployment
 *     setting is not the only gate (see WHY, below).
 *   - Nothing section-specific is written here. The table, the status selector
 *     and the date dialog are all driven by SECTIONS, so adding a field is still
 *     a config change plus a re-run of bootstrap().
 *   - The client is told what actually happened. setStatus already reports file
 *     errors alongside ok:true; the wrappers here pass that through, and the page
 *     shows it rather than a flat "Saved".
 *
 * WHY THE AUTH CHECK EXISTS AT ALL
 *
 * The UI deployment is "execute as me, access restricted to myself", so Google
 * already refuses everyone else. But webapp.access is a property of the PROJECT,
 * not of a deployment: the moment the Siri endpoint needs ANYONE_ANONYMOUS, the
 * UI deployment opens with it. Checking here means that change cannot silently
 * hand the whole system to a URL.
 *
 * Reading the caller's address needs the userinfo.email scope, which is why it
 * is pinned in appsscript.json alongside the other three.
 *
 * A CONSEQUENCE WORTH KNOWING: under ANYONE_ANONYMOUS, Google signs nobody in,
 * so Session.getActiveUser() is blank for EVERYONE including me — and this code
 * then denies everyone. That is the correct direction to fail, but it means the
 * Siri endpoint gets its own Apps Script project rather than a second deployment
 * of this one. Recorded in REBUILD-PLAN.md so the decision is not rediscovered.
 */

/**
 * Script Property holding the addresses allowed to use the UI, comma-separated.
 *
 * Optional. Left unset, the only allowed address is the account the script runs
 * as — which is the whole intent for a personal tool, and means there is no
 * configuration step you can forget in a way that opens access rather than
 * closing it.
 */
const UI_ALLOWED_PROPERTY = 'UI_ALLOWED_EMAILS';

/** Name of the HTML file served, without its extension. */
const UI_PAGE = 'Index';

const UI_TITLE = 'HelpfulForms';

/* ============================ Who is calling ============================== */

/**
 * Read an address out of a Session user, treating any failure as "unknown".
 *
 * A missing scope or a revoked authorisation must produce a denial, not an
 * exception that some caller might catch and carry on from.
 */
function uiEmailOf(getUser) {
  try {
    const user = getUser();
    return user ? (user.getEmail() || '').toString().trim().toLowerCase() : '';
  } catch (error) {
    Logger.log(`Could not read the caller's identity: ${error}`);
    return '';
  }
}

/** Addresses permitted to use the UI. */
function uiAllowedEmails() {
  const configured = PropertiesService.getScriptProperties()
    .getProperty(UI_ALLOWED_PROPERTY);

  const listed = (configured || '')
    .split(',')
    .map(email => email.trim().toLowerCase())
    .filter(Boolean);
  if (listed.length) return listed;

  // Nothing configured: only the account this runs as, which is me.
  const owner = uiEmailOf(() => Session.getEffectiveUser());
  return owner ? [owner] : [];
}

/**
 * Decide whether the caller may use the UI, without throwing.
 *
 * Separate from requireUiAccess() so doGet can render a page instead of an
 * Apps Script error screen, and so the harness can assert on the reason.
 */
function uiAccessCheck() {
  const email = uiEmailOf(() => Session.getActiveUser());
  const allowed = uiAllowedEmails();

  if (!email) {
    return { ok: false, email: '', reason: 'no identifiable signed-in user' };
  }
  if (!allowed.length) {
    return {
      ok: false, email: email,
      reason: `no allowed addresses could be determined; set ${UI_ALLOWED_PROPERTY}`
    };
  }
  if (allowed.indexOf(email) === -1) {
    return { ok: false, email: email, reason: 'address is not allowed' };
  }
  return { ok: true, email: email };
}

/**
 * Gate for every function the page can call.
 *
 * google.script.run reaches any global function in the project, so the page
 * being restricted is not on its own a reason to leave these open — each one
 * asks for itself.
 */
function requireUiAccess() {
  const verdict = uiAccessCheck();
  if (!verdict.ok) {
    Logger.log(`UI access denied (${verdict.email || 'anonymous'}): ${verdict.reason}`);
    // Deliberately says nothing about who is allowed.
    throw new Error('Not authorized.');
  }
  return verdict.email;
}

/**
 * Editor-runnable check on the access gate. Zero arguments, because that is all
 * the Apps Script editor can run — the same reason Smoke.js exists.
 *
 * Run this when the page says "Not authorized" and you cannot tell why. It
 * reports whether the caller could be identified at all, which separates the two
 * failures that look identical from the browser: a blank active user (scope not
 * granted, or an anonymous deployment) versus an address that simply is not on
 * the allowed list.
 *
 * Running it also forces the authorisation prompt if one is still outstanding,
 * since Apps Script asks for the manifest's whole scope list at once.
 */
function checkUiAccess() {
  const verdict = uiAccessCheck();
  const report = {
    ok: verdict.ok,
    activeUser: verdict.email || '(blank — nobody identifiable)',
    effectiveUser: uiEmailOf(() => Session.getEffectiveUser()) || '(blank)',
    allowed: uiAllowedEmails(),
    allowedFrom: PropertiesService.getScriptProperties().getProperty(UI_ALLOWED_PROPERTY)
      ? UI_ALLOWED_PROPERTY
      : 'the account the script runs as',
    reason: verdict.reason || 'allowed'
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/* ================================= doGet ================================== */

/**
 * Serve the page.
 *
 * No templating: the page fetches everything through google.script.run, so
 * there is one path by which data reaches the client rather than two. It also
 * means nothing is interpolated into the HTML, so there is nowhere for a sheet
 * value to arrive as markup.
 */
function doGet(e) {
  const verdict = uiAccessCheck();
  if (!verdict.ok) {
    Logger.log(`doGet denied (${verdict.email || 'anonymous'}): ${verdict.reason}`);
    return HtmlService
      .createHtmlOutput('<p style="font:16px system-ui">Not authorized.</p>')
      .setTitle(UI_TITLE);
  }

  return HtmlService.createHtmlOutputFromFile(UI_PAGE)
    .setTitle(UI_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ============================ Section metadata ============================ */

/**
 * Columns the table shows, in reading order, each with the label the UI uses.
 *
 * Derived from SECTIONS, so a field added there appears in the table without
 * anything here changing. Status, the state dates and the documents are NOT in
 * this list: they are rendered as controls rather than as text.
 */
function uiColumns(section) {
  const columns = [
    { header: COMMON.date, label: 'Date', type: 'date' },
    { header: COMMON.counterparty, label: counterpartyLabel(section), type: 'text' }
  ];

  if (section.category) {
    columns.push({
      header: section.category.header,
      label: section.category.label,
      type: 'text'
    });
  }

  section.extraFields.forEach(field => {
    columns.push({ header: field.header, label: field.label, type: field.type });
  });

  columns.push({ header: COMMON.amount, label: 'Amount', type: 'number' });
  columns.push({ header: COMMON.currency, label: 'Currency', type: 'text' });
  columns.push({ header: COMMON.notes, label: 'Notes', type: 'text' });

  return columns;
}

/**
 * The reference block — values that are the same on every row and so were never
 * really data. Blank properties are dropped rather than shown empty.
 */
function uiReference(section) {
  if (!section.reference) return [];
  const props = PropertiesService.getScriptProperties();
  return section.reference
    .map(item => ({ label: item.label, value: (props.getProperty(item.property) || '').trim() }))
    .filter(item => item.value);
}

/** Everything the client needs to render a section, and nothing else. */
function uiSectionMeta(sectionKey) {
  const section = getSection(sectionKey);
  return {
    key: sectionKey,
    label: section.label,
    sheet: section.sheet,
    counterpartyLabel: counterpartyLabel(section),
    category: section.category
      ? { header: section.category.header, label: section.category.label }
      : null,
    columns: uiColumns(section),
    states: section.states.map(state => ({
      name: state.name,
      dateColumn: state.dateColumn || null
    })),
    files: section.fileColumns.map(fileCol => ({
      header: fileCol.header,
      label: fileCol.label
    })),
    emailsClaim: !!section.emailOnCreate,
    reference: uiReference(section)
  };
}

/**
 * One round trip on page load: who you are, what today is, and the four
 * sections' shapes.
 */
function uiBootstrap() {
  const email = requireUiAccess();
  return {
    ok: true,
    user: email,
    today: today(),
    sections: Object.keys(SECTIONS).map(uiSectionMeta)
  };
}

/* ============================ Value formatting ============================ */

/**
 * yyyy-MM-dd for anything a date cell might hold, or '' when it holds no date.
 *
 * A real sheet returns a Date from a date-formatted cell and a string from a
 * text one, and a cell typed by hand can hold neither. All three arrive here.
 */
function uiDateISO(value) {
  if (value === '' || value === null || value === undefined) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime())
      ? ''
      : Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const text = value.toString().trim();
  if (isValidDateISO(text)) return text;

  const parsed = new Date(text);
  return isNaN(parsed.getTime())
    ? ''
    : Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * A cell value the client can use directly: strings and numbers only.
 *
 * Dates become ISO text so nothing downstream has to care whether a value came
 * back as a Date object, and so a row survives the trip through
 * google.script.run unchanged.
 */
function uiCellValue(value) {
  if (value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') return uiDateISO(value);
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  return value.toString();
}

/** A URL that is already pointing at Drive, whatever shape it takes. */
const UI_DRIVE_URL = /^https?:\/\/(?:[\w-]+\.)*drive\.google\.com\//i;

/**
 * A link to a document from whatever the file column holds.
 *
 * The column may hold a full URL or a bare ID depending on how the entry was
 * made, so both are turned into something clickable.
 *
 * WHY THE ACCOUNT HINT. A bare drive.google.com link resolves against whichever
 * Google account the browser has as its default, which is not necessarily the
 * one signed into this page. With two accounts signed in, every document read
 * "You need access" — the file was fine and the link was fine, it was being
 * opened as the wrong person. `authuser` says which account to open it as, and
 * the address used is the one that just passed the access check, so the link is
 * built for whoever is actually looking rather than for a hardcoded account.
 *
 * Drive links are therefore rebuilt from the ID rather than passed through, so
 * the hint lands on the ones stored as full URLs too — which is all of them, in
 * practice, since createEntry writes a URL. Anything that is not a Drive link is
 * left exactly as stored: guessing an ID out of some other service's URL would
 * turn a working link into a broken one.
 */
function uiFileUrl(fileRef, viewerEmail) {
  const text = (fileRef === null || fileRef === undefined) ? '' : fileRef.toString().trim();
  if (!text) return '';

  const isUrl = /^https?:\/\//i.test(text);
  if (isUrl && !UI_DRIVE_URL.test(text)) return text;

  const id = extractFileId(text);
  if (!id) return isUrl ? text : '';

  const url = `https://drive.google.com/file/d/${id}/view`;
  return viewerEmail ? `${url}?authuser=${encodeURIComponent(viewerEmail)}` : url;
}

/* ================================= Rows =================================== */

/**
 * Turn one sheet row into what the table renders.
 *
 * `options` is the point of this function. Each state is returned with its date
 * column and the date the row ALREADY has for it, which is what lets the date
 * dialog offer "Keep 15 Jan" rather than "Today" when reverting — the honest
 * wording for setStatus's "only fill if blank" rule. Computing it here rather
 * than in the page means the harness can test it, and the client stays
 * presentation only.
 *
 * Returns null for a row with nothing in it, which is what a row deleted by
 * hand in the sheet leaves behind.
 *
 * `viewerEmail` only reaches the document links — see uiFileUrl.
 */
function uiRow(section, sheetName, cols, rowValues, rowNumber, viewerEmail) {
  const raw = header => rowValues[columnIndex(cols, sheetName, header) - 1];

  const spine = [COMMON.date, COMMON.counterparty, COMMON.amount, COMMON.status];
  const empty = spine.every(header => {
    const value = raw(header);
    return value === '' || value === null || value === undefined;
  });
  if (empty) return null;

  const status = (raw(COMMON.status) || '').toString().trim();

  const dates = {};
  section.states.forEach(state => {
    if (state.dateColumn) dates[state.dateColumn] = uiDateISO(raw(state.dateColumn));
  });

  const cells = {};
  uiColumns(section).forEach(column => {
    const value = raw(column.header);
    cells[column.header] = column.type === 'date'
      ? (uiDateISO(value) || uiCellValue(value))
      : uiCellValue(value);
  });

  const files = section.fileColumns
    .map(fileCol => ({ label: fileCol.label, url: uiFileUrl(raw(fileCol.header), viewerEmail) }))
    .filter(file => file.url);

  const options = section.states.map(state => {
    const existing = state.dateColumn ? (dates[state.dateColumn] || '') : '';
    return {
      state: state.name,
      dateColumn: state.dateColumn || null,
      existingDate: existing,
      // True when moving here must not re-stamp a date the row already holds
      keepExisting: !!existing
    };
  });

  return {
    row: rowNumber,
    status: status,
    statusIndex: stateIndex(section, status),
    cells: cells,
    dates: dates,
    files: files,
    receiptState: (raw(COMMON.receiptState) || '').toString(),
    claimEmailed: section.emailOnCreate ? !!raw(CLAIM_EMAILED_COLUMN) : null,
    options: options
  };
}

/**
 * Every entry in a section, newest first.
 *
 * Read as ONE range rather than a getValue per cell — a phone on mobile data
 * notices the difference — while still resolving every column by header name.
 */
function uiListEntries(sectionKey) {
  const viewer = requireUiAccess();

  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const cols = resolveColumns(sheet);
  const meta = uiSectionMeta(sectionKey);

  const lastRow = sheet.getLastRow();
  let rows = [];

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    rows = values
      .map((rowValues, i) => uiRow(section, sheet.getName(), cols, rowValues, i + 2, viewer))
      .filter(row => row !== null);
    rows.reverse();
  }

  return {
    ok: true,
    section: sectionKey,
    today: today(),
    meta: meta,
    rows: rows
  };
}

/**
 * One row, re-read from the sheet. Used to refresh a row after a change.
 *
 * Checks the caller even though its callers here have already done so: it is a
 * global that returns row data, and google.script.run can reach it directly
 * without going through uiSetStatus. The repeated check costs a property read.
 */
function uiEntry(sectionKey, sheetRow) {
  const viewer = requireUiAccess();

  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const row = resolveDataRow(sheet, sheetRow);
  const cols = resolveColumns(sheet);
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return uiRow(section, sheet.getName(), cols, values, row, viewer);
}

/* ================================ Actions ================================= */

/**
 * The status control. A single call whichever direction it moves, because
 * setStatus has no separate undo — going back is selecting an earlier state.
 *
 * Returns setStatus's own result, including fileErrors, plus the row as the
 * sheet now holds it. The page renders from that rather than from what it hoped
 * would happen, so a rename failure cannot look like a success.
 */
function uiSetStatus(sectionKey, sheetRow, newState, dateISO) {
  requireUiAccess();

  const result = setStatus(sectionKey, sheetRow, newState, dateISO || null);
  result.date = uiDateISO(result.date);
  result.entry = uiEntry(sectionKey, result.row);
  return result;
}

/**
 * Correct a date without changing state. Blank clears it.
 * setEntryDate refuses any column that is not one of this section's state dates.
 */
function uiSetEntryDate(sectionKey, sheetRow, dateColumn, dateISO) {
  requireUiAccess();

  const result = setEntryDate(sectionKey, sheetRow, dateColumn, dateISO || '');
  result.entry = uiEntry(sectionKey, result.row);
  return result;
}
