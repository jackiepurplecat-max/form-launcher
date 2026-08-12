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

/** Header spine of the registry sheet, in reading order. */
const REGISTRY_HEADERS = [
  REGISTRY.name, REGISTRY.type, REGISTRY.nif,
  REGISTRY.aliases, REGISTRY.timesUsed, REGISTRY.lastUsed
];

/**
 * Confidence at or above which a match may prefill data on its own.
 *
 * Deliberately strict. A wrong NIF produces a rejected claim, which is far
 * worse than an empty field — below this bar we keep what was actually said
 * and let the completion step sort it out.
 */
const REGISTRY_AUTOFILL_CONFIDENCE = 0.85;

/**
 * Similarity at or above which a name is worth *offering* in the dropdown.
 *
 * Much lower than the autofill bar, and that is the point: suggesting a name
 * fills nothing in. The strict bar exists to stop a wrong NIF reaching a claim,
 * so applying it to a list of candidates would be borrowing a rule from a
 * decision this is not. Matches findSupplier's own floor for its weakest tier.
 */
const REGISTRY_SUGGEST_SIMILARITY = 0.6;

/* ================================ Storage ================================= */

/**
 * The registry sheet, with its headers guaranteed.
 *
 * Headers go through applyHeaders() on every call, not only on creation. A
 * sheet that exists but is missing a column would otherwise read as undefined
 * through loadRegistry() and then fail on the first write to it, in a way that
 * points nowhere near the cause. applyHeaders only ever appends, so this is
 * safe against a registry already in use.
 */
function getOrCreateRegistrySheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(REGISTRY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRY_SHEET);
    Logger.log(`Created "${REGISTRY_SHEET}"`);
  }
  applyHeaders(sheet, REGISTRY_HEADERS);
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
    // columnIndex rather than cols[header]: a missing column would otherwise
    // index by NaN and read as a blank string, silently losing every alias.
    const cell = header => rowValues[columnIndex(cols, REGISTRY_SHEET, header) - 1];
    const aliases = (cell(REGISTRY.aliases) || '').toString()
      .split(',').map(a => a.trim()).filter(Boolean);
    return {
      row: i + 2,
      name: (cell(REGISTRY.name) || '').toString(),
      type: (cell(REGISTRY.type) || '').toString(),
      nif: (cell(REGISTRY.nif) || '').toString(),
      aliases: aliases,
      timesUsed: Number(cell(REGISTRY.timesUsed)) || 0,
      // Raw, not parsed: a merge has to keep the later of two, and the
      // management list shows it. Matching is indifferent to it.
      lastUsed: cell(REGISTRY.lastUsed)
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
 *   1.00        exact after normalising
 *   0.95        matches a recorded alias — the cure for repeated mishearings
 *   0.75–0.95   one name contains the other, scaled by how much of the longer
 *               one is accounted for ("white clinic" in "the white clinic"
 *               scores 0.90; "uber" in "uber eats" only 0.84)
 *   ~           edit-distance similarity, for ordinary mishearings
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

/**
 * Registry entries matching typed text, most-used first, for autocomplete.
 * Matches anywhere in the name or an alias, not just at the start, so "luz"
 * finds "Hospital da Luz". A blank prefix returns the most-used entries.
 *
 * Substring matching cannot see past a typo, and that showed: "white" offered
 * White Clinic while "whitee clinic" offered nothing, because a wrong letter
 * mid-string means no registry name contains the text and the text contains no
 * registry name. findSupplier scores that same string at 0.92 — the matcher was
 * never the problem, the dropdown just never asked it. So substring matches come
 * first, and near misses top the list up to the limit.
 *
 * Length-sensitive similarity is what keeps this from dragging in junk: a single
 * typed letter scores near zero against every name, so the fuzzy half only
 * contributes once enough has been typed to be wrong rather than incomplete.
 */
function suggestSuppliers(prefix, limit) {
  const target = normalizeName(prefix);
  const cap = limit || 10;
  const registry = loadRegistry();

  const matches = registry.filter(entry => {
    if (!target) return true;
    const name = normalizeName(entry.name);
    return name.indexOf(target) !== -1 ||
      entry.aliases.some(alias => normalizeName(alias).indexOf(target) !== -1);
  });

  matches.sort((a, b) => b.timesUsed - a.timesUsed || a.name.localeCompare(b.name));

  if (target && matches.length < cap) {
    const already = {};
    matches.forEach(entry => { already[entry.name] = true; });

    registry
      .filter(entry => !already[entry.name])
      .map(entry => ({ entry: entry, score: similarity(normalizeName(entry.name), target) }))
      .filter(scored => scored.score >= REGISTRY_SUGGEST_SIMILARITY)
      .sort((a, b) => b.score - a.score || b.entry.timesUsed - a.entry.timesUsed)
      .slice(0, cap - matches.length)
      .forEach(scored => matches.push(scored.entry));
  }

  return matches.slice(0, cap).map(entry => ({
    name: entry.name, type: entry.type, nif: entry.nif
  }));
}

/**
 * Look a counterparty up for a given section and say what may be prefilled.
 *
 * This is the half of the registry that Google Forms made impossible: "type
 * FNAC, get its NIF". The form calls it as you type; Siri calls it with
 * whatever it heard.
 *
 * prefill is keyed by COLUMN HEADER so a caller can pass it straight to
 * createEntry() without translating anything. It is empty below the autofill
 * threshold - the match is still returned, so the UI can offer it as a
 * suggestion, but nothing is filled in on a guess.
 *
 * The canonical name is part of the prefill: correcting "wite clinic" to
 * "White Clinic" is the most useful thing a confident match can do, and it
 * stops near-miss spellings accumulating as separate registry entries.
 */
function lookupCounterparty(sectionKey, spokenName) {
  const section = getSection(sectionKey);
  const match = findSupplier(spokenName);
  if (!match) return null;

  const prefill = {};
  if (match.autofill) {
    prefill[COMMON.counterparty] = match.name;
    if (section.registryTypeField && match.type) {
      prefill[section.registryTypeField] = match.type;
    }
    if (section.registryNifField && match.nif) {
      prefill[section.registryNifField] = match.nif;
    }
  }

  return {
    name: match.name,
    confidence: match.confidence,
    reason: match.reason,
    autofill: match.autofill,
    prefill: prefill
  };
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

  const type = ((details || {}).type || '').toString().trim();
  const nif = ((details || {}).nif || '').toString().trim();

  // Locked for the same reason createEntry is: a new supplier is appended at
  // getLastRow() + 1, and every write here goes through writeCell, so a name
  // learned from a Siri field is stored as text rather than as a formula.
  return withLock(() => {
    const sheet = getOrCreateRegistrySheet();
    const cols = resolveColumns(sheet);
    const target = normalizeName(clean);

    const existing = loadRegistry().find(entry =>
      normalizeName(entry.name) === target ||
      entry.aliases.some(alias => normalizeName(alias) === target)
    );

    if (!existing) {
      const row = sheet.getLastRow() + 1;
      writeCell(sheet, cols, row, REGISTRY.name, clean);
      if (type) writeCell(sheet, cols, row, REGISTRY.type, type);
      if (nif) writeCell(sheet, cols, row, REGISTRY.nif, nif);
      writeCell(sheet, cols, row, REGISTRY.timesUsed, 1);
      writeCell(sheet, cols, row, REGISTRY.lastUsed, new Date());
      SpreadsheetApp.flush();
      Logger.log(`Registry: learned "${clean}"`);
      return { created: true, name: clean, row: row };
    }

    writeCell(sheet, cols, existing.row, REGISTRY.timesUsed, existing.timesUsed + 1);
    writeCell(sheet, cols, existing.row, REGISTRY.lastUsed, new Date());

    if (nif && !existing.nif) {
      writeCell(sheet, cols, existing.row, REGISTRY.nif, nif);
    }
    if (type) {
      if (!existing.type) {
        writeCell(sheet, cols, existing.row, REGISTRY.type, type);
      } else if (existing.type !== type) {
        // Conflicting types seen - stop guessing rather than guess wrongly
        writeCell(sheet, cols, existing.row, REGISTRY.type, '');
        Logger.log(`Registry: "${existing.name}" seen as both ${existing.type} and ${type}; type default cleared`);
      }
    }

    return { created: false, name: existing.name, row: existing.row };
  });
}

/** Attach an alias, so a recurring mishearing resolves next time. */
function addSupplierAlias(name, alias) {
  const clean = (alias || '').toString().trim();
  if (!clean) throw new Error('Alias is required');

  // Aliases share one cell, comma-separated, so an alias containing a comma
  // would silently come back as two.
  if (clean.indexOf(',') !== -1) {
    throw new Error(`An alias cannot contain a comma: "${clean}"`);
  }

  return withLock(() => {
    const sheet = getOrCreateRegistrySheet();
    const cols = resolveColumns(sheet);
    const target = normalizeName(name);

    const entry = loadRegistry().find(e => normalizeName(e.name) === target);
    if (!entry) throw new Error(`No registry entry named "${name}"`);

    if (entry.aliases.some(a => normalizeName(a) === normalizeName(clean))) {
      return { ok: false, error: `"${clean}" is already an alias of ${entry.name}` };
    }

    const updated = entry.aliases.concat([clean]);
    writeCell(sheet, cols, entry.row, REGISTRY.aliases, updated.join(', '));
    return { ok: true, name: entry.name, aliases: updated };
  });
}
