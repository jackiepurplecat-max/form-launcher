/**
 * v2 — the Siri intake endpoint.
 *
 * WHERE THIS CODE ACTUALLY RUNS. Not here. This file lives in the main,
 * spreadsheet-bound project, but every call into it arrives from a SECOND Apps
 * Script project — `v2-siri/`, whose entire contents is a manifest and a
 * five-line doPost that delegates to `siriHandlePost`. That project is the one
 * deployed with `access: ANYONE_ANONYMOUS`.
 *
 * WHY TWO PROJECTS. `webapp.access` is a property of the project, not of the
 * deployment. Opening this project to anonymous callers so Siri could reach it
 * would also open the web UI — and worse, anonymous access blanks
 * Session.getActiveUser() for EVERYONE INCLUDING ME, so `uiAccessCheck()` would
 * start refusing me too and no check inside doGet could tell the difference.
 * The isolation is the point; see Security in REBUILD-PLAN.md.
 *
 * WHY THE LOGIC IS NEVERTHELESS HERE. The shim holds no logic, no secrets and
 * no configuration, so there is one copy of the code, one set of Script
 * Properties, and — this is the part that matters day to day — the whole
 * endpoint is exercised by `npm run v2:test` against the real source. Only the
 * five-line delegation is untested, and it has nothing in it to get wrong.
 *
 * THE SHAPE OF A CONVERSATION. Three calls, because the confirmation has to
 * happen before anything is written:
 *
 *   catalog  what to ask -> the category list, the labels, today's date
 *   resolve  what was heard -> the canonical supplier, written nothing
 *   create   what was confirmed -> the row
 *
 * `resolve` exists because nothing canonicalises the counterparty on the
 * server. That is deliberate — two near-identical names can be two real
 * businesses, so the correction is shown and overridable rather than applied
 * silently. The custom form has a natural moment for that; Siri only has one if
 * it is built in, and without it a 0.92 mishearing lands in a filename and
 * appends a second supplier row on a path with nobody watching.
 *
 * EVERY RESPONSE IS HTTP 200. ContentService cannot set a status code, so the
 * outcome lives in the body's `ok` field and a Shortcut must read it rather
 * than trusting that the request succeeded.
 */

/** Script Property holding the shared key. Unset means the endpoint is shut. */
const SIRI_KEY_PROPERTY = 'SIRI_API_KEY';

/** Value written to the Source column by this path. */
const SIRI_SOURCE = 'siri';

/* ============================== The envelope ============================== */

function siriJson(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function siriFail(message) {
  return siriJson({ ok: false, error: message });
}

/**
 * Compare two secrets without returning early on the first differing byte.
 *
 * Timing analysis across Apps Script's own latency is not a realistic attack,
 * so this is about intent rather than a threat being modelled: a key check that
 * visibly does not leak is one nobody has to reason about later.
 */
function siriKeyMatches(supplied, expected) {
  const a = (supplied === null || supplied === undefined) ? '' : supplied.toString();
  const b = (expected === null || expected === undefined) ? '' : expected.toString();
  if (!a || !b) return false;

  let diff = a.length ^ b.length;
  const len = Math.max(a.length, b.length);
  for (let i = 0; i < len; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

/**
 * The gate. Fails closed in every direction: no key configured, no key
 * supplied, or a key that does not match are all the same refusal.
 *
 * An unset SIRI_API_KEY shutting the endpoint is the important one. The shim
 * project is deployed to ANYONE_ANONYMOUS, so "no key configured" must never
 * mean "no key required" — that would leave the endpoint open the moment it is
 * deployed and before it is configured, which is exactly the window in which
 * nobody is looking.
 */
function siriAuthorised(request) {
  const expected = PropertiesService.getScriptProperties().getProperty(SIRI_KEY_PROPERTY);
  if (!expected) {
    Logger.log('Siri: refused — SIRI_API_KEY is not set, so the endpoint is shut.');
    return false;
  }
  return siriKeyMatches(request.key, expected);
}

/* ============================== The actions =============================== */

/**
 * What the Shortcut needs in order to ask its questions.
 *
 * Fetched rather than baked into the Shortcut so that adding a patient, or a
 * new Expense Reason, never means editing four Shortcuts on a phone. That is
 * the whole reason this call exists.
 */
function siriCatalog(request) {
  const section = getSection(request.section);
  const medium = sectionReceiptMedium(section);

  return {
    ok: true,
    section: request.section,
    label: section.label,
    counterpartyLabel: counterpartyLabel(section),
    currency: SIRI_DEFAULT_CURRENCY,
    date: today(),
    // Sent so the Shortcut offers the choices without hardcoding them. Null for
    // Income, which has no documents and so nothing to go looking for.
    receiptMedium: medium ? {
      header: medium.header,
      label: medium.label,
      required: medium.required === true,
      values: medium.options.slice()
    } : null,
    category: section.category ? {
      header: section.category.header,
      label: section.category.label,
      // A declared list is closed: Health's patients are the family and a
      // misspelling typed once would become a second patient forever. An open
      // one is a suggestion list and free text stays allowed.
      closed: !!(section.category.options && section.category.options.length),
      required: section.category.required !== false,
      values: categoryValues(request.section)
    } : null
  };
}

/**
 * Match what was heard against the registry. Writes nothing.
 *
 * `confirm` is the string the Shortcut should put in front of you. When the
 * registry is confident it is the canonical spelling; otherwise it is exactly
 * what was heard, because a poor guess shown as if it were a correction is
 * worse than no correction at all.
 */
function siriResolve(request) {
  getSection(request.section);

  const heard = (request.counterparty === null || request.counterparty === undefined)
    ? '' : request.counterparty.toString().trim();
  if (!heard) return { ok: false, error: 'Nothing was heard for the counterparty.' };

  const match = lookupCounterparty(request.section, heard);

  // No match, or one too weak to fill anything in: keep what was heard. A new
  // supplier is a normal event, not an error - the registry learns it on save.
  if (!match || !match.autofill) {
    return {
      ok: true,
      heard: heard,
      confirm: heard,
      corrected: false,
      known: false,
      confidence: match ? match.confidence : 0
    };
  }

  return {
    ok: true,
    heard: heard,
    confirm: match.name,
    // Only a change of SPELLING is a correction worth showing. An exact hit
    // rewrites nothing, and saying "corrected" of it would train you to ignore
    // the word in the case that matters.
    corrected: normalizeName(match.name) !== normalizeName(heard) || match.name !== heard,
    known: true,
    confidence: match.confidence
  };
}

/**
 * Create the row.
 *
 * THE WHITELIST IS THE SECURITY BOUNDARY OF THIS FILE. `createEntry` accepts
 * any column the sheet has, and this script runs as me. A file column reaching
 * it would be the whole exploit: `extractFileId` pulls a Drive id out of
 * whatever string it is given, so a key holder passing a URL for a file of MINE
 * would have that file renamed and moved into HelpfulForms. Siri sends the core
 * fields and nothing else.
 *
 * An unknown or forbidden field is REFUSED rather than dropped. Silently
 * ignoring it would mean a Shortcut that thinks it is recording something it is
 * not, and that failure is invisible for as long as nobody checks the sheet.
 */
function siriCreate(request) {
  const section = getSection(request.section);

  const allowed = siriAllowedFields(section);
  const supplied = request.fields || {};

  const rejected = Object.keys(supplied).filter(header => allowed.indexOf(header) === -1);
  if (rejected.length) {
    return {
      ok: false,
      error: `Not accepted from Siri: ${rejected.join(', ')}. ` +
        `This endpoint takes ${allowed.join(', ')} only.`
    };
  }

  const fields = {};
  allowed.forEach(header => {
    if (supplied[header] !== undefined) fields[header] = supplied[header];
  });

  // Not asked for, and not a decision to leave to the caller. Both are shown in
  // the confirmation before this call is made.
  if (!fields[COMMON.date]) fields[COMMON.date] = today();
  if (!fields[COMMON.currency]) fields[COMMON.currency] = SIRI_DEFAULT_CURRENCY;

  siriPrefillFromRegistry(section, fields);

  const result = createEntry(request.section, fields, SIRI_SOURCE);

  /*
   * WHAT "COMPLETE" HAS TO MEAN HERE. Not `result.ok`. createEntry reports
   * ok:false only for a missing REQUIRED field, and by that measure a work
   * expense with no receipt is complete — which is exactly the entry that most
   * needs finishing. The honest signal is the completion request: it is raised
   * when a required field is blank OR a document is still awaited, and its
   * `outstanding` list already names the documents in the words the form uses.
   *
   * An incomplete entry is not a failure. A partial entry is the safety net —
   * the row exists, the mail has gone, and the link finishes it in one tap. So
   * this reports ok:true and says what is left.
   */
  const completion = result.completionRequest;

  return {
    ok: true,
    row: result.row,
    complete: !completion,
    outstanding: (completion && completion.outstanding) || [],
    // Distinguished because they are finished differently: a blank field is
    // typed in, a document has to be found and attached.
    awaitingDocument: result.receiptState === RECEIPT_STATE.awaiting,
    missingFields: result.warnings || [],
    // Never assume the reminder went. Unset COMPLETION_EMAIL_RECIPIENT and it
    // silently does not, which is worth knowing on the phone rather than later.
    completionEmailed: !!(completion && completion.ok),
    counterparty: fields[COMMON.counterparty] || '',
    amount: fields[COMMON.amount],
    date: fields[COMMON.date]
  };
}

/** Default currency. Not asked for; shown in the confirmation. */
const SIRI_DEFAULT_CURRENCY = 'EUR';

/**
 * The only columns this endpoint will write.
 *
 * Built from SECTIONS rather than listed, so a renamed category header cannot
 * leave a stale string behind — but deliberately NOT built by subtracting the
 * file columns from the full header list. A whitelist that is "everything
 * except the dangerous ones" grants every column added in future by default,
 * and the next section-specific field would be writable from an anonymous
 * endpoint without anyone deciding it should be.
 */
function siriAllowedFields(section) {
  const allowed = [COMMON.counterparty, COMMON.amount, COMMON.currency, COMMON.date];
  if (section.category) allowed.push(section.category.header);

  // The one deliberate exception to "Siri captures the core only". That rule
  // exists so adding a field never means re-editing four Shortcuts on a phone,
  // and it is a good rule. This field breaks it because it is the only one that
  // is knowable ONLY at capture time: standing at the counter you know whether
  // you were handed paper, and by the time the completion mail arrives you do
  // not. A field that can only be answered now must be asked now.
  const medium = sectionReceiptMedium(section);
  if (medium) allowed.push(medium.header);

  return allowed;
}

/**
 * Fill Type and NIF from the registry, on an EXACT match only.
 *
 * The value of the registry is that Uber is always a Taxi, and Siri has no
 * moment to ask. But the fuzzy matcher must not run here: the counterparty
 * arriving in this call is the one that was confirmed on the phone, and
 * re-matching it could quietly merge a supplier the confirmation had just
 * established was a different business. An exact hit cannot do that — it fills
 * blanks and changes no name.
 */
function siriPrefillFromRegistry(section, fields) {
  const name = (fields[COMMON.counterparty] || '').toString().trim();
  if (!name) return;

  const match = lookupCounterparty(sectionKeyOf(section), name);
  if (!match || !match.autofill) return;
  if (normalizeName(match.name) !== normalizeName(name)) return;

  Object.keys(match.prefill).forEach(header => {
    // Never the counterparty itself: whatever was confirmed on the phone wins,
    // including its spelling.
    if (header === COMMON.counterparty) return;
    if (!fields[header]) fields[header] = match.prefill[header];
  });
}

/**
 * Reachability check, for the one thing the harness cannot prove.
 *
 * This code runs as a LIBRARY when Siri calls it, and two things behave
 * differently there than they do in the editor: `getActiveSpreadsheet()`
 * resolves against the caller's container, which a standalone project does not
 * have, and `getScriptProperties()` may or may not mean this project's
 * properties. Both are load-bearing and neither can be tested in node. So the
 * shim can ask, and get a straight answer, before anything is built on top.
 */
function siriPing() {
  const props = PropertiesService.getScriptProperties();
  const seen = {};
  ['ROOT_FOLDER_ID', 'SPREADSHEET_ID', 'SIRI_API_KEY'].forEach(key => {
    seen[key] = !!props.getProperty(key);
  });

  let spreadsheet = null;
  let error = null;
  try {
    spreadsheet = getSpreadsheet().getName();
  } catch (e) {
    error = e.message;
  }

  return {
    ok: !!spreadsheet,
    propertiesVisible: seen,
    spreadsheet: spreadsheet,
    error: error,
    sections: Object.keys(SECTIONS)
  };
}

/* ================================ Setup =================================== */

/**
 * Configure the two properties the Siri endpoint needs. Run once, from the
 * editor of the MAIN project.
 *
 * Both are set here rather than typed into Project Settings, for the same
 * reason in each case: typing them is the step that goes wrong.
 *
 *   SPREADSHEET_ID  is read off the container rather than copied out of a URL,
 *                   so it cannot be a character short or the id of some other
 *                   sheet. Nothing bound ever reads it — only the library path.
 *
 *   SIRI_API_KEY    is generated rather than invented. A key chosen by hand is
 *                   short and memorable, and this one guards an endpoint that
 *                   anyone on the internet can reach.
 *
 * The key is RETURNED, once, so it can be copied into the Shortcut. It is not
 * shown again: a second run reports that one already exists and leaves it
 * alone, because regenerating it silently would break every Shortcut on the
 * phone. To deliberately replace it, delete the property first.
 */
function siriSetup() {
  const props = PropertiesService.getScriptProperties();
  const report = { spreadsheetId: null, key: null, keyAlreadySet: false };

  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  props.setProperty(SPREADSHEET_ID_PROPERTY, id);
  report.spreadsheetId = id;

  const existing = props.getProperty(SIRI_KEY_PROPERTY);
  if (existing) {
    report.keyAlreadySet = true;
    report.key = '(unchanged — delete the property first to replace it)';
  } else {
    const key = siriGenerateKey();
    props.setProperty(SIRI_KEY_PROPERTY, key);
    report.key = key;
  }

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/**
 * Replace the key deliberately, in one run.
 *
 * `siriSetup()` refuses to touch an existing key, which is right — it is the
 * function you run when setting things up, and silently rotating from there
 * would break every Shortcut on the phone with nothing to say why. But rotating
 * is a real operation, not an accident to be prevented: a key ends up somewhere
 * it should not be and has to be replaced. Making that mean hand-editing Script
 * Properties turns a legitimate thing into fiddling in a UI, which is how it
 * ends up not being done.
 *
 * So it gets its own name rather than its own ceremony. Nobody runs
 * `siriRotateKey` by accident, and the old key stops working the instant it
 * does — which is the whole reason this is not what `siriSetup()` does.
 */
function siriRotateKey() {
  const props = PropertiesService.getScriptProperties();
  const had = !!props.getProperty(SIRI_KEY_PROPERTY);

  const key = siriGenerateKey();
  props.setProperty(SIRI_KEY_PROPERTY, key);

  const report = {
    key: key,
    replacedAnExistingKey: had,
    reminder: 'The old key stopped working just now. Update every Shortcut, and .env.'
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/** 32 hex characters from Apps Script's own UUID source. */
function siriGenerateKey() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 32);
}

/* =============================== The router =============================== */

const SIRI_ACTIONS = {
  catalog: siriCatalog,
  resolve: siriResolve,
  create: siriCreate
};

/**
 * Everything the shim's doPost does, so that the shim itself has nothing in it.
 *
 * `e` is the Apps Script event object. The body is JSON rather than form
 * parameters because Shortcuts sends JSON natively and it keeps `fields` a
 * nested object rather than something to be flattened and parsed back.
 */
function siriHandlePost(e) {
  let request;
  try {
    const body = e && e.postData && e.postData.contents;
    if (!body) return siriFail('Empty request.');
    request = JSON.parse(body);
  } catch (error) {
    return siriFail('Body was not valid JSON.');
  }

  if (!request || typeof request !== 'object') return siriFail('Body was not an object.');

  // Before anything else looks at anything else.
  if (!siriAuthorised(request)) return siriFail('Not authorized.');

  // ping is deliberately inside the gate. It reports which properties are set,
  // which is not something to tell an anonymous caller.
  if (request.action === 'ping') return siriJson(siriPing());

  const handler = SIRI_ACTIONS[request.action];
  if (!handler) {
    return siriFail(`Unknown action: ${request.action}. Expected one of ` +
      `${Object.keys(SIRI_ACTIONS).join(', ')}, ping.`);
  }

  if (!SECTIONS[request.section]) {
    return siriFail(`Unknown section: ${request.section}. Expected one of ` +
      `${Object.keys(SECTIONS).join(', ')}.`);
  }

  try {
    return siriJson(handler(request));
  } catch (error) {
    // The Shortcut has to be told something it can show. An unhandled throw
    // reaches Shortcuts as an HTML error page and reads as a broken network.
    Logger.log(`Siri ${request.action} failed: ${error.stack || error.message}`);
    return siriFail(error.message);
  }
}
