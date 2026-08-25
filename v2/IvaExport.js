/**
 * v2 — the DRORIVA declaration export.
 *
 * Turns complete, unsubmitted IVA rows into the XML file the Portuguese tax
 * app opens, so a claim is reviewed rather than retyped. The format is
 * reverse-engineered in v2/IVA-EXPORT-FORMAT.md — read that before changing
 * anything here; five of its traps are load-time failures with no error message
 * worth reading.
 *
 * The emitter below was proved by round-trip: v2/test/iva-roundtrip.js parses a
 * declaration DRORIVA saved itself, rebuilds it through `ivaDeclarationXml` and
 * compares byte for byte. Keep that true. If you change the whitespace, the tag
 * spacing or the totals, re-run it against a real sample before pushing.
 *
 * WHAT IS DELIBERATELY NOT HERE: this generates a file, it does not submit one
 * and it does not advance any row. Generating is not submitting — the app can
 * still refuse the file — and a status that moved on the strength of a download
 * would be reporting success for work that had not happened.
 */

/* ========================= The emitter (format) ========================== */

const IVA_XML_HEADER =
  "<?xml version='1.0' encoding='ISO-8859-1'?>\n" +
  '\n' +
  '<Simplex130 xmlns="http://www.dgci.gov.pt/2013/Simplex130"' +
  ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
  ' xsi:schemaLocation="http://www.dgci.gov.pt/2013/Simplex130 Simplex130.xsd"' +
  ' versao="1" >\n';

function ivaXmlEscape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * `<Tag >text</Tag>` — the stray space is the app's own writer joining an empty
 * attribute map. Cosmetic almost certainly, matched anyway: it costs nothing and
 * it is what a byte-identical round trip requires.
 */
function ivaTag(indent, name, text) {
  return indent + '<' + name + ' >' + ivaXmlEscape(text) + '</' + name + '>\n';
}

/**
 * Amounts as integer cents. SomaIVA and SomaValorFatura are WRITTEN INTO the
 * file rather than computed on load, so float drift produces a declaration that
 * disagrees with itself — and the app believes the totals.
 */
function ivaCents(value) {
  const n = typeof value === 'number' ? value : parseFloat(String(value).replace(',', '.'));
  if (!isFinite(n)) throw new Error(`Not a number: ${value}`);
  return Math.round(n * 100);
}

function ivaMoney(centsValue) {
  const sign = centsValue < 0 ? '-' : '';
  const abs = Math.abs(centsValue);
  return sign + Math.floor(abs / 100) + '.' + ('0' + (abs % 100)).slice(-2);
}

/**
 * The declaration, from a model. Pure string building — no sheet, no Drive, no
 * Apps Script services — which is what lets the round-trip harness run it
 * outside Apps Script against a file the app itself wrote.
 */
function ivaDeclarationXml(model) {
  let out = IVA_XML_HEADER;
  out += '  <Rosto>\n';
  out += '    <Quadro01>\n';
  out += ivaTag('      ', 'NifEntidade', model.nifEntidade);
  out += ivaTag('      ', 'TipoEntidade', model.tipoEntidade);
  out += ivaTag('      ', 'NifFuncionario', model.nifFuncionario);
  out += '    </Quadro01>\n';
  out += '    <Quadro02>\n';
  out += '      <Quadro02T1 >\n';

  let somaIva = 0;
  let somaTotal = 0;

  for (let i = 0; i < model.lines.length; i++) {
    const line = model.lines[i];
    const iva = ivaCents(line.valorIVA);
    const total = ivaCents(line.valorTotalFatura);
    somaIva += iva;
    somaTotal += total;

    // Numbered from 1, and `numero` is required.
    out += '        <Quadro02T1-Linha numero="' + (i + 1) + '">\n';
    out += ivaTag('          ', 'NumeroFatura', line.numeroFatura);
    out += ivaTag('          ', 'DataFatura', line.dataFatura);
    out += ivaTag('          ', 'NifFatura', line.nifFatura);
    if (line.codigoBem) out += ivaTag('          ', 'CodigoBem', line.codigoBem);
    out += ivaTag('          ', 'ValorIVA', ivaMoney(iva));
    out += ivaTag('          ', 'ValorTotalFatura', ivaMoney(total));

    const f = line.ficheiro;
    // Two spaces after `Ficheiro` — same empty-attribute-map join as above.
    out += '          <Ficheiro  attachmentDate="' + ivaXmlEscape(f.attachmentDate) +
      '" fileName="' + ivaXmlEscape(f.fileName) +
      '" extension="' + ivaXmlEscape(f.extension) + '">' +
      f.base64 + '</Ficheiro>\n';
    out += '        </Quadro02T1-Linha>\n';
  }

  out += '      </Quadro02T1>\n';
  out += ivaTag('      ', 'SomaIVA', ivaMoney(somaIva));
  out += ivaTag('      ', 'SomaValorFatura', ivaMoney(somaTotal));
  out += '    </Quadro02>\n';
  out += '  </Rosto>\n';
  out += '</Simplex130>\n';
  return out;
}

/* ======================== Validation at export time ======================= */

/**
 * AT's own types, applied where they bite rather than at intake.
 *
 * VALIDATION-PLAN.md proposes the same rules for intake and that is step 13.
 * These are here as well, and not instead, because the export is the point at
 * which a bad value stops being an annoyance and becomes a rejected submission —
 * and because existing rows predate any intake rule. A row that fails is NAMED
 * AND LEFT OUT, never silently dropped and never quietly corrected.
 *
 *   Decimal_15  fractionDigits=2 minInclusive=0
 *   NifFatura   xs:long — nine digits, numeric
 */
function ivaLineProblems(line) {
  const problems = [];

  if (!line.numeroFatura) problems.push('no Número');

  if (!isValidDateISO(line.dataFatura)) problems.push('no usable Date');

  if (!/^\d{9}$/.test(String(line.nifFatura || ''))) {
    problems.push(`Emitente NIF is not nine digits (${line.nifFatura || 'blank'})`);
  }

  let iva = null;
  let total = null;
  try {
    iva = ivaCents(line.valorIVA);
  } catch (err) {
    problems.push('Valor do IVA is not a number');
  }
  try {
    total = ivaCents(line.valorTotalFatura);
  } catch (err) {
    problems.push('Amount is not a number');
  }

  if (iva !== null && iva < 0) problems.push('Valor do IVA is negative');
  if (total !== null && total < 0) problems.push('Amount is negative');
  // Not a schema rule, but a transposed pair is the mistake this catches and the
  // app would take it: VAT above the invoice total is arithmetically impossible.
  if (iva !== null && total !== null && iva > total) {
    problems.push('Valor do IVA is more than the Amount');
  }

  if (!line.codigoBem) {
    // `Tipo` is required: false at intake on purpose (Config.js), so this is a
    // real and expected gap on older rows rather than a broken entry.
    problems.push('no Tipo — set one before exporting');
  }

  if (!line.ficheiro) problems.push('no Fatura attached');

  return problems;
}

/**
 * `101 — Vestuário e calçado` → `101`.
 *
 * The code lives in the option label so there is no second table to keep in
 * step with Config.js. Anything that is not a leading run of digits returns
 * null and the row is reported rather than exported.
 */
function ivaCodigoBem(tipo) {
  const m = String(tipo || '').trim().match(/^(\d{3})\b/);
  return m ? m[1] : null;
}

/* ========================= Gathering the rows ============================ */

/** Format for `attachmentDate` — `yyyy-MM-dd'T'HH:mm:ss'Z'`, always UTC. */
function ivaAttachmentStamp(date) {
  const d = date || new Date();
  const pad = n => ('0' + n).slice(-2);
  return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()) +
    'T' + pad(d.getUTCHours()) + ':' + pad(d.getUTCMinutes()) + ':' + pad(d.getUTCSeconds()) + 'Z';
}

function ivaCell(cols, sheetName, rowValues, header) {
  return rowValues[columnIndex(cols, sheetName, header) - 1];
}

function ivaText(value) {
  if (value === null || value === undefined) return '';
  return value.toString().trim();
}

/**
 * Which rows go in the declaration.
 *
 * THE ROWS IN THE FIRST STATE — `To Do` — that are complete and have their
 * Fatura. Submitting is what moves a row on to `Claimed`, exactly as it does in
 * Health, so the sheet already knows what has been submitted and no `Submitted`
 * column is needed. A boolean beside the status saying the same thing would be
 * a second answer to a question the status already answers, and a fifth
 * accordion for a state that is not a place an entry sits in.
 *
 * If that reading of `Claimed` is ever wrong, this is one line: the state name
 * comes from `section.declaration.sourceState`.
 */
function ivaCollectLines(section, sheet, cols, options) {
  const opts = options || {};
  const sheetName = sheet.getName();
  const sourceState = (section.declaration && section.declaration.sourceState) ||
    section.states[0].name;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { lines: [], skipped: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const fileHeader = section.fileColumns[0].header;

  const lines = [];
  const skipped = [];

  values.forEach((rowValues, i) => {
    const rowNumber = i + 2;
    const status = ivaText(ivaCell(cols, sheetName, rowValues, COMMON.status));
    if (!status) return;                       // blank row
    if (status !== sourceState) return;        // already submitted, or settled

    if (opts.rows && opts.rows.indexOf(rowNumber) === -1) return;

    const line = {
      row: rowNumber,
      counterparty: ivaText(ivaCell(cols, sheetName, rowValues, COMMON.counterparty)),
      numeroFatura: ivaText(ivaCell(cols, sheetName, rowValues, 'Número')),
      dataFatura: uiDateISO(ivaCell(cols, sheetName, rowValues, COMMON.date)),
      nifFatura: ivaText(ivaCell(cols, sheetName, rowValues, 'Emitente NIF')),
      codigoBem: ivaCodigoBem(ivaCell(cols, sheetName, rowValues, section.category.header)),
      valorIVA: ivaCell(cols, sheetName, rowValues, 'IVA Amount'),
      valorTotalFatura: ivaCell(cols, sheetName, rowValues, COMMON.amount),
      fileRef: ivaText(ivaCell(cols, sheetName, rowValues, fileHeader))
    };
    line.ficheiro = line.fileRef ? true : null;

    const problems = ivaLineProblems(line);
    if (problems.length) {
      skipped.push({ row: rowNumber, counterparty: line.counterparty, reasons: problems });
      return;
    }
    lines.push(line);
  });

  return { lines: lines, skipped: skipped };
}

/**
 * Pull each Fatura out of Drive and base64 it.
 *
 * SIZE IS THE THING MOST LIKELY TO BREAK. Base64 inflates by a third and these
 * are scans — one real invoice was 111 KB of PDF for 231 KB of XML — and Apps
 * Script holds the whole declaration in memory as one string. So the budget is
 * enforced here, per line, and a line that would take the total past it is
 * DEFERRED AND REPORTED rather than dropped: the caller says how many are left
 * for the next batch, because a truncated declaration that says nothing reads
 * exactly like a complete one.
 *
 * No NUL padding. The app pads its own blobs and a PDF is self-terminating at
 * %%EOF, so the padding is a buffer the writer never truncates — see the round
 * trip in v2/test/iva-roundtrip.js, and the `.nopad` file it produces, which is
 * how that was confirmed against the app rather than assumed.
 */
function ivaAttachDocuments(lines, budgetBytes, stamp) {
  const attached = [];
  const deferred = [];
  let used = 0;

  lines.forEach(line => {
    if (deferred.length) {                      // keep the batch contiguous
      deferred.push({ row: line.row, counterparty: line.counterparty, reasons: ['batch full'] });
      return;
    }

    const fileId = extractFileId(line.fileRef);
    if (!fileId) {
      deferred.push({ row: line.row, counterparty: line.counterparty, reasons: ['Fatura link is not a Drive file'] });
      return;
    }

    let file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (err) {
      deferred.push({ row: line.row, counterparty: line.counterparty, reasons: [`Fatura not readable: ${err.message}`] });
      return;
    }

    const base64 = Utilities.base64Encode(file.getBlob().getBytes());
    if (used + base64.length > budgetBytes && attached.length) {
      deferred.push({ row: line.row, counterparty: line.counterparty, reasons: ['batch full'] });
      return;
    }
    used += base64.length;

    const split = splitExtension(file.getName());
    attached.push({
      row: line.row,
      counterparty: line.counterparty,
      numeroFatura: line.numeroFatura,
      dataFatura: line.dataFatura,
      nifFatura: line.nifFatura,
      codigoBem: line.codigoBem,
      valorIVA: line.valorIVA,
      valorTotalFatura: line.valorTotalFatura,
      ficheiro: {
        attachmentDate: stamp,
        // Without the extension — it is a separate attribute.
        fileName: split.base,
        extension: (split.ext || 'pdf').replace(/^\./, ''),
        base64: base64
      }
    });
  });

  return { attached: attached, deferred: deferred, bytes: used };
}

/* ============================== The export =============================== */

/**
 * Build the declaration for the IVA section and return it as text plus a
 * summary of what went in and what did not.
 *
 * Returns rather than writes, so the harness can assert on the XML and the
 * caller decides whether it becomes a Drive file or a download.
 */
function buildIvaDeclaration(options) {
  const opts = options || {};
  const section = getSection('iva');
  const decl = section.declaration;
  if (!decl) throw new Error('The iva section has no `declaration` config.');

  const sheet = getSheet(section);
  const cols = resolveColumns(sheet);

  const props = PropertiesService.getScriptProperties();
  const nifEntidade = ivaText(props.getProperty(decl.entityNifProperty));
  const nifFuncionario = ivaText(props.getProperty(decl.claimantNifProperty));

  // Both NIFs live in Script Properties and neither is in this repo. Unset, the
  // app would take the file and reject the submission, which is a far worse
  // failure than refusing to build it.
  const missingRefs = [];
  if (!/^\d{9}$/.test(nifEntidade)) missingRefs.push(decl.entityNifProperty);
  if (!/^\d{9}$/.test(nifFuncionario)) missingRefs.push(decl.claimantNifProperty);
  if (missingRefs.length) {
    throw new Error(
      `Cannot build the declaration: ${missingRefs.join(' and ')} must be set to a ` +
      'nine-digit NIF in Script Properties.'
    );
  }

  const collected = ivaCollectLines(section, sheet, cols, opts);
  const stamp = ivaAttachmentStamp(opts.now);
  const budget = opts.maxBytes || decl.maxBase64Bytes;
  const withDocs = ivaAttachDocuments(collected.lines, budget, stamp);

  const summary = {
    ok: withDocs.attached.length > 0,
    lines: withDocs.attached.map(l => ({
      row: l.row,
      counterparty: l.counterparty,
      numeroFatura: l.numeroFatura,
      dataFatura: l.dataFatura,
      valorIVA: ivaMoney(ivaCents(l.valorIVA)),
      valorTotalFatura: ivaMoney(ivaCents(l.valorTotalFatura)),
      codigoBem: l.codigoBem
    })),
    skipped: collected.skipped.concat(withDocs.deferred),
    base64Bytes: withDocs.bytes,
    sourceState: decl.sourceState
  };
  summary.somaIVA = ivaMoney(withDocs.attached.reduce((t, l) => t + ivaCents(l.valorIVA), 0));
  summary.somaValorFatura =
    ivaMoney(withDocs.attached.reduce((t, l) => t + ivaCents(l.valorTotalFatura), 0));

  if (!withDocs.attached.length) {
    summary.xml = null;
    summary.message = collected.skipped.length
      ? `Nothing to export: ${collected.skipped.length} ${decl.sourceState} ` +
        `${collected.skipped.length === 1 ? 'row is' : 'rows are'} not ready.`
      : `Nothing to export: no ${decl.sourceState} rows.`;
    return summary;
  }

  summary.xml = ivaDeclarationXml({
    nifEntidade: nifEntidade,
    tipoEntidade: decl.entityTypeCode,
    nifFuncionario: nifFuncionario,
    lines: withDocs.attached
  });
  summary.message = `${withDocs.attached.length} ` +
    `${withDocs.attached.length === 1 ? 'invoice' : 'invoices'}, ` +
    `${summary.somaValorFatura} total, ${summary.somaIVA} IVA` +
    (summary.skipped.length ? `. ${summary.skipped.length} left out — see below.` : '.');
  return summary;
}

/**
 * Write the declaration into the IVA section's folder and hand back a link.
 *
 * ISO-8859-1, not UTF-8. Supplier names carry Portuguese accents and the
 * declaration says Latin-1 in its own prolog, so the bytes have to agree with
 * it — `newBlob(string)` would write UTF-8 and the app would disagree with the
 * file it was handed.
 */
function ivaWriteDeclaration(summary, folder) {
  const section = getSection('iva');
  const name = `Simplex130 ${today()}.xml`;
  const blob = Utilities.newBlob('', 'text/xml', name)
    .setDataFromString(summary.xml, 'ISO-8859-1');
  const target = folder || sectionFolder(section, section.declaration.folder);
  const file = target.createFile(blob);
  return { id: file.getId(), name: name, url: file.getUrl ? file.getUrl() : null };
}

/**
 * Web-callable: build the declaration, save it, return where it went.
 *
 * Checks the caller itself. Every global the page can reach does — reaching it
 * through google.script.run does not go past the deployment setting.
 */
function uiIvaExport() {
  requireUiAccess();

  const summary = buildIvaDeclaration();
  if (!summary.ok) {
    return {
      ok: false,
      message: summary.message,
      lines: [],
      skipped: summary.skipped
    };
  }

  const written = ivaWriteDeclaration(summary);
  return {
    ok: true,
    message: summary.message,
    file: written,
    lines: summary.lines,
    skipped: summary.skipped,
    somaIVA: summary.somaIVA,
    somaValorFatura: summary.somaValorFatura,
    // Say what the file does NOT do. Nothing has moved to Claimed: the rows move
    // when the submission is accepted, not when the file is generated.
    reminder: `Nothing has moved to Claimed. Advance the rows once Finanças has ` +
      `accepted the declaration.`
  };
}
