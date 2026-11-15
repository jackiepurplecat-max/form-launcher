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

/**
 * Make a value safe to put inside HTML.
 *
 * Used for the mail that is sent as HTML. Its body is built from sheet cells, and
 * a counterparty or note is the same untrusted text that `safeCellValue` exists to
 * neutralise on the way in — a `<` in a supplier name would otherwise silently
 * swallow the rest of the sentence, and that is the harmless version.
 *
 * Also correct for an href: `&` becomes `&amp;`, which is what an attribute
 * carrying a query string is supposed to contain.
 */
function escapeHtml(value) {
  return (value === null || value === undefined ? '' : value)
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * The key a section is held under, for the places that have the object and need
 * the name — a URL, a message. Returns '' rather than throwing: a caller that
 * only wants to build a link should not be the thing that fails.
 */
function sectionKeyOf(section) {
  const keys = Object.keys(SECTIONS);
  for (let i = 0; i < keys.length; i++) {
    if (SECTIONS[keys[i]] === section) return keys[i];
  }
  return '';
}

/**
 * The spreadsheet, whether or not this code is running inside its container.
 *
 * WHY THIS IS NOT JUST getActiveSpreadsheet(). The Siri endpoint has to be a
 * SEPARATE Apps Script project — `webapp.access` is per project, so opening it
 * to ANYONE_ANONYMOUS for Siri would open the UI too, and anonymous access
 * blanks Session.getActiveUser() for everyone including me, so no check inside
 * doGet could compensate. A second project cannot also be bound to this
 * spreadsheet: a container may have exactly one bound script. It therefore
 * reaches this code as a LIBRARY, and a library resolves
 * getActiveSpreadsheet() against its CALLER's container — which for a
 * standalone project is nothing at all.
 *
 * So: prefer the container when there is one, and fall back to opening by id.
 * The order matters. Bound execution — the editor, the web UI, smokeTest —
 * keeps working untouched even if SPREADSHEET_ID is never set, and only the
 * library path pays for the property.
 *
 * ONLY the openById result is cached, because only it is a real fetch and a
 * single createEntry touches the sheet a dozen times. The active spreadsheet is
 * asked for every time: caching it would save nothing and would freeze whatever
 * object the first call happened to see.
 */
let SPREADSHEET_CACHE = null;

function getSpreadsheet() {
  // Bound execution. Wrapped because a standalone caller does not merely get
  // null here — depending on the context it can throw — and that must fall
  // through to the id rather than escape.
  let active = null;
  try {
    active = SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    active = null;
  }
  if (active) return active;

  if (SPREADSHEET_CACHE) return SPREADSHEET_CACHE;

  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
  if (!id) {
    throw new Error(
      'No active spreadsheet, and no SPREADSHEET_ID script property to fall back on. ' +
      'Library callers need it set — see checkScriptProperties().'
    );
  }

  SPREADSHEET_CACHE = SpreadsheetApp.openById(id);
  return SPREADSHEET_CACHE;
}

/** Drop the cached spreadsheet. Tests only; one execution never needs it. */
function clearSpreadsheetCache() {
  SPREADSHEET_CACHE = null;
}

function getSheet(section) {
  const sheet = getSpreadsheet().getSheetByName(section.sheet);
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
 *
 * Cached per execution. Apps Script globals do not survive between runs, so
 * this is only ever a within-run cache - but anything that ADDS a column mid-run
 * must call clearColumnCache(), or subsequent lookups will miss it.
 */
function resolveColumns(sheet) {
  const key = sheet.getName();
  if (_columnCache[key]) return _columnCache[key];

  const width = sheet.getLastColumn();
  if (!width) {
    throw new Error(
      `Sheet "${key}" has no header row. Run bootstrap() to create the columns.`
    );
  }

  const headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  const map = {};
  headers.forEach((header, i) => {
    const name = (header || '').toString().trim();
    if (name) map[name] = i + 1;
  });

  _columnCache[key] = map;
  return map;
}

/** Forget cached header positions. Call after adding or renaming a column. */
function clearColumnCache() {
  Object.keys(_columnCache).forEach(key => delete _columnCache[key]);
}

function columnIndex(cols, sheetName, header) {
  const index = cols[header];
  if (!index) throw new Error(`Column "${header}" not found in ${sheetName}`);
  return index;
}

/**
 * Neutralise a value that Sheets would otherwise treat as a formula.
 *
 * A cell whose text starts with =, +, - or @ becomes a formula, so a supplier
 * name of "=IMPORTXML(...)" would execute on write rather than being stored.
 * Prefixing with an apostrophe forces literal text; Sheets strips it again on
 * read, so nothing downstream sees the difference.
 *
 * Genuine negative numbers are left alone - "-50" is data, not an attack, and
 * escaping it would turn an amount into text.
 *
 * This matters most for the Siri endpoint, which accepts field values from
 * anyone holding the device key. Principle 5: the client is never trusted.
 */
function safeCellValue(value) {
  if (typeof value !== 'string') return value;
  if (!/^[=+\-@\t\r]/.test(value)) return value;
  if (value.trim() !== '' && isFinite(Number(value))) return value;
  return `'${value}`;
}

/* ============================ Field validation ============================ */

/**
 * The value rules a field can declare, by name (VALIDATION-PLAN.md).
 *
 * Each takes a non-blank value and returns `{ value }` when it passes, or
 * `{ problem }` naming what is wrong in words that follow "<label> ".
 */
const FIELD_RULES = {
  money: value => {
    const text = value.toString().trim();
    if (/^-\d+([.,]\d+)?$/.test(text)) return { problem: 'cannot be negative' };
    if (!/^\d+([.,]\d+)?$/.test(text)) return { problem: 'must be a number' };
    // Refused rather than rounded: no invoice has three decimals, so it is a
    // typo, and rounding it would change a figure claimed from Finanças.
    if (/[.,]\d{3,}$/.test(text)) return { problem: 'can have at most 2 decimal places' };
    // A comma is the Portuguese decimal point. Stored as a number, not as the
    // text typed, so the sheet can add it up.
    return { value: Number(text.replace(',', '.')) };
  },

  nif: value => {
    const text = value.toString().trim();
    if (!/^\d{9}$/.test(text)) return { problem: 'must be 9 digits' };
    return isValidNif(text) ? { value: text } : { problem: 'is not a valid NIF (check digit)' };
  },

  // Finanças' own rule, which its published schema does not state: DRORIVA's
  // Quadro02Validator checks numeroFatura against [/\\a-zA-Z0-9\_\.\-]+. So the
  // value is CLEANED to fit, even though it then no longer matches the paper -
  // and the change is reported, because a silent one would leave the user
  // believing the paper's number is what was saved.
  numero: value => {
    const typed = value.toString().trim();
    const cleaned = typed
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')   // Ç -> C, É -> E
      .replace(/[^A-Za-z0-9\/\\_.\-]/g, '');
    // Someone typed something; turning it into a blank would make a mistake
    // look like "not filled in yet".
    if (!cleaned) return { problem: 'has none of the characters Finanças accepts (letters, digits, / \\ _ . -)' };
    if (cleaned === typed) return { value: cleaned };
    return { value: cleaned, explain: 'Finanças accepts only letters, digits and / \\ _ . -' };
  }
};

/**
 * A Portuguese NIF's ninth digit is computed from the first eight (mod 11).
 * Nine digits alone catches a dropped character; this also catches two digits
 * swapped, which is the mistake people make and the one that looks plausible.
 */
function isValidNif(nineDigits) {
  let sum = 0;
  for (let i = 0; i < 8; i++) sum += Number(nineDigits[i]) * (9 - i);
  const remainder = sum % 11;
  const check = remainder < 2 ? 0 : 11 - remainder;
  return Number(nineDigits[8]) === check;
}

/** Every field of a section that declares a rule: [{ header, label, rule }]. */
function fieldRules(section) {
  const rules = Object.keys(COMMON_VALIDATE).map(header => ({
    header: header,
    label: header === COMMON.amount ? amountLabel(section) : header,
    rule: COMMON_VALIDATE[header]
  }));
  (section.extraFields || []).filter(field => field.validate).forEach(field => {
    rules.push({ header: field.header, label: field.label, rule: field.validate });
  });
  return rules;
}

/**
 * Check supplied values against the section's rules.
 *
 * BLANK IS ALWAYS LEGAL: only supplied, non-empty values are checked. A partial
 * entry is the safety net, and a validator that refused blanks would break
 * Siri, the completion mail and scan-later in one go.
 *
 * Returns `{ values, problems, changes }`: `values` is `fields` with each value
 * in its checked form, `problems` lists `{ header, message }` for every value
 * refused, and `changes` lists `{ header, from, to, message }` for every value
 * a rule CLEANED rather than refused. A cleaned value is a different act from a
 * refused one and must reach the caller - see VALIDATION-PLAN.md.
 *
 * `others`, optional, holds values already stored. They are never checked
 * themselves - only used to complete a `lessThan` pair whose other side is in
 * `fields` - so an edit is judged on what it changes.
 */
function checkFieldValues(section, fields, others) {
  const values = Object.assign({}, fields);
  const problems = [];
  const changes = [];

  fieldRules(section).forEach(entry => {
    const value = values[entry.header];
    if (value === null || value === undefined || value.toString().trim() === '') return;
    const outcome = FIELD_RULES[entry.rule](value);
    if (outcome.problem) {
      problems.push({ header: entry.header, message: `${entry.label} ${outcome.problem}` });
    } else {
      values[entry.header] = outcome.value;
      if (outcome.explain) {
        changes.push({
          header: entry.header, from: value, to: outcome.value,
          message: `${entry.label} saved as ${outcome.value} — ${outcome.explain}`
        });
      }
    }
  });

  // Pairs, compared only once both sides are present and each passed its own
  // rule - a refused or missing value has nothing to say about the other.
  (section.lessThan || []).forEach(pair => {
    if (!(pair.header in values) && !(pair.than in values)) return;
    const refused = problems.some(p => p.header === pair.header || p.header === pair.than);
    const side = header => (header in values ? values[header] : (others || {})[header]);
    const small = side(pair.header);
    const large = side(pair.than);
    if (refused || typeof small !== 'number' || typeof large !== 'number') return;
    if (small < large) return;
    problems.push({
      header: pair.header,
      message: `${ruleLabel(section, pair.header)} must be less than ${ruleLabel(section, pair.than)}`
    });
  });

  return { values: values, problems: problems, changes: changes };
}

/** The label a rule's message uses for a header: what the form calls it. */
function ruleLabel(section, header) {
  const entry = fieldRules(section).filter(e => e.header === header)[0];
  return entry ? entry.label : header;
}

function readCell(sheet, cols, row, header) {
  return sheet.getRange(row, columnIndex(cols, sheet.getName(), header)).getValue();
}

function writeCell(sheet, cols, row, header, value) {
  sheet.getRange(row, columnIndex(cols, sheet.getName(), header))
    .setValue(safeCellValue(value));
}

/**
 * Record that something just happened to a row: Last Updated becomes now.
 *
 * Every operation that changes an entry calls this - creating, editing,
 * attaching, a status or state-date change, restoring (archiving stamps the
 * archive copy itself). It is what "Last updated" sorts by, so an operation that
 * forgets it leaves its row sinking down the list as if untouched. The harness
 * checks each one.
 *
 * Called BEFORE the operation's other writes, so a sheet without the column
 * (the upgrade not yet run) refuses the operation outright instead of half
 * doing it and then failing.
 */
function touchEntry(sheet, cols, row) {
  writeCell(sheet, cols, row, COMMON.lastUpdated, new Date());
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

/** True for yyyy-MM-dd naming a date the calendar actually has. */
function isValidDateISO(value) {
  const text = (value === null || value === undefined) ? '' : value.toString().trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return false;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);

  // Rejects 2026-02-31, which Date would roll forward into March
  return date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day;
}

/**
 * A cell as a Date, or null when it holds none. Sheets hands back a Date for a
 * date cell and text for anything typed; yyyy-MM-dd is read as a local date so
 * it lands on the day it names.
 */
function cellAsDate(value) {
  if (value === '' || value === null || value === undefined) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime()) ? null : value;
  }
  const text = value.toString().trim();
  if (isValidDateISO(text)) {
    const parts = text.split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Reject a bad date BEFORE anything is written.
 *
 * Without this a caller could put arbitrary text in a date column, and the
 * failure would surface later inside buildSuffixChain - after the status and
 * dates had already been written, leaving the row half-changed.
 */
function requireDateISO(value, label) {
  if (!isValidDateISO(value)) {
    throw new Error(`${label} must be a valid yyyy-MM-dd date, got "${value}"`);
  }
  return value.toString().trim();
}

/* ================================ Locking ================================= */

/**
 * Depth of the script lock held by THIS execution.
 *
 * Nested withLock() calls are a no-op rather than a second acquisition, because
 * createEntry -> initializeEntry -> learnCounterparty -> recordSupplier would
 * otherwise ask for a lock it already holds.
 */
let _lockDepth = 0;

/**
 * Run fn while holding the script lock.
 *
 * Needed anywhere a row is appended: every append computes its target from
 * getLastRow(), so two callers arriving together - the form and Siri, or two
 * Siri taps - would compute the same row and one would overwrite the other.
 *
 * Scope this as narrowly as the race requires. Slow work (Drive, Gmail) belongs
 * outside it.
 */
function withLock(fn, timeoutMs) {
  if (_lockDepth > 0) return fn();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(timeoutMs || 20000)) {
    throw new Error('Timed out waiting for the script lock; another write is in progress.');
  }

  _lockDepth++;
  try {
    return fn();
  } finally {
    _lockDepth--;
    lock.releaseLock();
  }
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

/**
 * Resolve <root>/<Section>/<folderName>, creating anything missing.
 *
 * Named from section.sheet, not section.label, so the folders match the sheet
 * tabs: Work / IVA / Health / Income. label is a UI string - "Log Income" is a
 * button, not a sensible place to keep files.
 *
 * Folders are matched BY NAME, so renaming one in Drive does not move the
 * files here - it makes the next call create a fresh empty folder alongside.
 * Change this only while the tree is empty.
 */
function sectionFolder(section, folderName) {
  const rootId = PropertiesService.getScriptProperties().getProperty(ROOT_FOLDER_PROPERTY);
  if (!rootId) throw new Error(`${ROOT_FOLDER_PROPERTY} not set in Script Properties`);
  return childFolder(childFolder(DriveApp.getFolderById(rootId), section.sheet), folderName);
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

    // Skipped rather than thrown: a date typed by hand into the sheet must not
    // be able to abort a transition half way through, and the state itself is
    // still recorded in the Status column either way.
    const parsed = new Date(value);
    if (isNaN(parsed.getTime())) {
      Logger.log(
        `${sheet.getName()} row ${row}: "${state.dateColumn}" holds "${value}", ` +
        `which is not a date - omitted from the filename`
      );
      return;
    }

    const stamp = Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd-MM-yyyy');
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
 *
 * folderName overrides where the documents land, and exists for exactly one
 * caller: repairing an ARCHIVED row. Those files were moved to Archived when the
 * row was, and their names still have to be rebuilt when a supplier is renamed —
 * but filing them by status would quietly un-archive them. Naming and filing are
 * one path on purpose, so the override lives here rather than in a second copy
 * of the naming rule. Omit it and nothing changes.
 */
function applyFileState(section, sheet, cols, row, targetIndex, folderName) {
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
      if (!destination) {
        destination = folderName
          ? sectionFolder(section, folderName)
          : folderForState(section, target);
      }
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

  // Validated up front, so a bad date cannot leave the row with its later
  // dates already cleared
  const requested = dateISO ? requireDateISO(dateISO, target.dateColumn || 'Date') : null;

  const previous = (readCell(sheet, cols, row, COMMON.status) || '').toString().trim();

  touchEntry(sheet, cols, row);

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
    if (requested) {
      effectiveDate = requested;
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
 *
 * A date for a state the row has NOT reached is refused. setStatus clears the
 * dates of every state after the target, so such a date is wiped by the next
 * transition — accepting it means writing something that quietly disappears,
 * which is worse than saying no. Selecting the state is how its date gets set,
 * and that is the one path that keeps status and dates agreeing.
 *
 * Two deliberate exceptions:
 *   - Clearing is always allowed. Removing a date can never create the
 *     inconsistency this guards against.
 *   - A row whose Status is not a recognised state is allowed through, because
 *     refusing would leave no way to repair it from the UI - which is the only
 *     place a hand-edited sheet gets noticed.
 */
function setEntryDate(sectionKey, sheetRow, dateColumn, dateISO) {
  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const row = resolveDataRow(sheet, sheetRow);
  const cols = resolveColumns(sheet);

  const targetIndex = section.states.findIndex(s => s.dateColumn === dateColumn);
  if (targetIndex === -1) {
    throw new Error(`"${dateColumn}" is not a date column of ${sectionKey}`);
  }

  // Blank clears the date; anything else must be a real one
  const value = dateISO ? requireDateISO(dateISO, dateColumn) : '';

  if (value) {
    const status = (readCell(sheet, cols, row, COMMON.status) || '').toString().trim();
    const currentIndex = stateIndex(section, status);
    if (currentIndex !== -1 && targetIndex > currentIndex) {
      throw new Error(
        `"${dateColumn}" belongs to ${section.states[targetIndex].name}, which this ` +
        `entry has not reached - it is "${status}". Select that state instead; ` +
        `it asks for the date.`
      );
    }
  }

  touchEntry(sheet, cols, row);
  writeCell(sheet, cols, row, dateColumn, value);
  return { ok: true, section: sectionKey, row: row, column: dateColumn, date: value };
}
