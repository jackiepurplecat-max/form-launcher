/**
 * v2 — the supplier registry.
 *
 * NOT YET DEPLOYED. See Config.js.
 *
 * A self-populating list of counterparties. Nothing is entered up front: every
 * entry you create teaches it, so it is current by construction rather than by
 * maintenance.
 *
 * It serves two callers:
 *   the form   autocomplete, and prefilling Type / NIF once known
 *   Siri       matching a misheard name against something real
 *
 * Only possible because v2 has no Google Form — Forms cannot fill one answer
 * from another, which is exactly what "type FNAC, get its NIF" requires.
 */

const REGISTRY_SHEET = 'Suppliers';

const REGISTRY = {
  name: 'Name',
  type: 'Type',
  nif: 'NIF',
  aliases: 'Aliases',
  timesUsed: 'Times Used',
  lastUsed: 'Last Used'
};

/**
 * Confidence at or above which a match may prefill data on its own.
 *
 * Deliberately strict. A wrong NIF produces a rejected claim, which is far
 * worse than an empty field — below this bar we keep what was actually said
 * and let the completion step sort it out.
 */
const REGISTRY_AUTOFILL_CONFIDENCE = 0.85;

/* ================================ Storage ================================= */

function getOrCreateRegistrySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REGISTRY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRY_SHEET);
    const headers = [
      REGISTRY.name, REGISTRY.type, REGISTRY.nif,
      REGISTRY.aliases, REGISTRY.timesUsed, REGISTRY.lastUsed
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    Logger.log(`Created "${REGISTRY_SHEET}"`);
  }
  return sheet;
}

/**
 * Flatten a name for comparison: accents removed, punctuation dropped, case
 * and spacing ignored. "Farmácia Sá" and "farmacia sa" become the same thing.
 */
function normalizeName(text) {
  return (text || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function loadRegistry() {
  const sheet = getOrCreateRegistrySheet();
  if (sheet.getLastRow() < 2) return [];

  const cols = resolveColumns(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  return values.map((rowValues, i) => {
    const cell = header => rowValues[cols[header] - 1];
    const aliases = (cell(REGISTRY.aliases) || '').toString()
      .split(',').map(a => a.trim()).filter(Boolean);
    return {
      row: i + 2,
      name: (cell(REGISTRY.name) || '').toString(),
      type: (cell(REGISTRY.type) || '').toString(),
      nif: (cell(REGISTRY.nif) || '').toString(),
      aliases: aliases,
      timesUsed: Number(cell(REGISTRY.timesUsed)) || 0
    };
  }).filter(entry => entry.name);
}

/* ================================ Matching ================================ */

/** Edit distance, capped so a long mismatch exits cheaply. */
function editDistance(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;

  let previous = [];
  for (let j = 0; j <= b.length; j++) previous[j] = j;

  for (let i = 1; i <= a.length; i++) {
    const current = [i];
    for (let j = 1; j <= b.length; j++) {
      current[j] = Math.min(
        previous[j] + 1,
        current[j - 1] + 1,
        previous[j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1)
      );
    }
    previous = current;
  }
  return previous[b.length];
}

function similarity(a, b) {
  const longest = Math.max(a.length, b.length);
  if (!longest) return 0;
  return 1 - editDistance(a, b) / longest;
}

/**
 * Best registry match for a name, with a confidence and the reason for it.
 *
 * Tiers, strongest first:
 *   1.00  exact after normalising
 *   0.95  matches a recorded alias — the cure for repeated Siri mishearings
 *   0.80  one name contains the other ("white clinic" vs "the white clinic")
 *   ~     edit-distance similarity, for ordinary mishearings
 *
 * Returns null when nothing is close, rather than offering a poor guess.
 */
function findSupplier(spokenName) {
  const target = normalizeName(spokenName);
  if (!target) return null;

  let best = null;
  const consider = (entry, confidence, reason) => {
    if (!best || confidence > best.confidence) {
      best = { entry: entry, confidence: confidence, reason: reason };
    }
  };

  loadRegistry().forEach(entry => {
    const name = normalizeName(entry.name);
    if (name === target) return consider(entry, 1, 'exact');

    if (entry.aliases.some(alias => normalizeName(alias) === target)) {
      return consider(entry, 0.95, 'alias');
    }
    if (name.indexOf(target) !== -1 || target.indexOf(name) !== -1) {
      // Scaled by how much of the longer string is accounted for, so
      // "white clinic" inside "the white clinic" scores high while "uber"
      // inside "uber eats" does not.
      const ratio = Math.min(name.length, target.length) / Math.max(name.length, target.length);
      return consider(entry, 0.75 + 0.2 * ratio, 'contains');
    }
    const score = similarity(name, target);
    if (score >= 0.6) consider(entry, score, 'similar');
  });

  if (!best) return null;
  return {
    name: best.entry.name,
    type: best.entry.type,
    nif: best.entry.nif,
    confidence: best.confidence,
    reason: best.reason,
    // The caller must not prefill identifiers on a guess
    autofill: best.confidence >= REGISTRY_AUTOFILL_CONFIDENCE
  };
}

/** Registry entries matching a typed prefix, most-used first, for autocomplete. */
function suggestSuppliers(prefix, limit) {
  const target = normalizeName(prefix);
  const matches = loadRegistry().filter(entry => {
    if (!target) return true;
    const name = normalizeName(entry.name);
    return name.indexOf(target) !== -1 ||
      entry.aliases.some(alias => normalizeName(alias).indexOf(target) !== -1);
  });

  matches.sort((a, b) => b.timesUsed - a.timesUsed || a.name.localeCompare(b.name));
  return matches.slice(0, limit || 10).map(entry => ({
    name: entry.name, type: entry.type, nif: entry.nif
  }));
}

/* =============================== Learning ================================= */

/**
 * Record use of a supplier, creating it if new.
 *
 * Type handling is deliberately cautious. Some suppliers offer several services
 * — a clinic doing both dentistry and exams — so once a second, different type
 * is seen the stored default is CLEARED rather than overwritten. A field that
 * prefills the wrong value is worse than one that prefills nothing.
 *
 * NIF is never cleared: it is a fact about the supplier, not about the visit.
 */
function recordSupplier(name, details) {
  const clean = (name || '').toString().trim();
  if (!clean) return null;

  const sheet = getOrCreateRegistrySheet();
  const cols = resolveColumns(sheet);
  const target = normalizeName(clean);

  const existing = loadRegistry().find(entry =>
    normalizeName(entry.name) === target ||
    entry.aliases.some(alias => normalizeName(alias) === target)
  );

  const type = ((details || {}).type || '').toString().trim();
  const nif = ((details || {}).nif || '').toString().trim();

  if (!existing) {
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, cols[REGISTRY.name]).setValue(clean);
    if (type) sheet.getRange(row, cols[REGISTRY.type]).setValue(type);
    if (nif) sheet.getRange(row, cols[REGISTRY.nif]).setValue(nif);
    sheet.getRange(row, cols[REGISTRY.timesUsed]).setValue(1);
    sheet.getRange(row, cols[REGISTRY.lastUsed]).setValue(new Date());
    Logger.log(`Registry: learned "${clean}"`);
    return { created: true, name: clean, row: row };
  }

  sheet.getRange(existing.row, cols[REGISTRY.timesUsed]).setValue(existing.timesUsed + 1);
  sheet.getRange(existing.row, cols[REGISTRY.lastUsed]).setValue(new Date());

  if (nif && !existing.nif) {
    sheet.getRange(existing.row, cols[REGISTRY.nif]).setValue(nif);
  }
  if (type) {
    if (!existing.type) {
      sheet.getRange(existing.row, cols[REGISTRY.type]).setValue(type);
    } else if (existing.type !== type) {
      // Conflicting types seen - stop guessing rather than guess wrongly
      sheet.getRange(existing.row, cols[REGISTRY.type]).setValue('');
      Logger.log(`Registry: "${existing.name}" seen as both ${existing.type} and ${type}; type default cleared`);
    }
  }

  return { created: false, name: existing.name, row: existing.row };
}

/** Attach an alias, so a recurring mishearing resolves next time. */
function addSupplierAlias(name, alias) {
  const sheet = getOrCreateRegistrySheet();
  const cols = resolveColumns(sheet);
  const target = normalizeName(name);

  const entry = loadRegistry().find(e => normalizeName(e.name) === target);
  if (!entry) throw new Error(`No registry entry named "${name}"`);

  const clean = (alias || '').toString().trim();
  if (!clean) throw new Error('Alias is required');
  if (entry.aliases.some(a => normalizeName(a) === normalizeName(clean))) {
    return { ok: false, error: `"${clean}" is already an alias of ${entry.name}` };
  }

  const updated = entry.aliases.concat([clean]);
  sheet.getRange(entry.row, cols[REGISTRY.aliases]).setValue(updated.join(', '));
  return { ok: true, name: entry.name, aliases: updated };
}
