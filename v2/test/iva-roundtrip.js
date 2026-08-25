#!/usr/bin/env node
/**
 * Round-trip proof for the DRORIVA declaration format.
 *
 * Step 5 of NEXT-SESSION.md says to prove a generated file loads BEFORE wiring
 * anything to real rows. This does the cheap half of that: it parses a
 * declaration the app itself saved, rebuilds it from the parsed values, and
 * compares byte for byte. If the rebuild is identical, the emitter is right and
 * the only thing left to prove is the one assumption bytes cannot settle —
 * whether DRORIVA minds the missing NUL padding.
 *
 *   node v2/test/iva-roundtrip.js ~/test_iva
 *
 * Writes two files next to the sample:
 *   <sample>.rebuilt   — blob preserved exactly, must be byte-identical
 *   <sample>.nopad     — trailing NULs stripped, OPEN THIS IN DRORIVA
 *
 * Outputs are refused inside the repo: a real sample carries a real supplier
 * NIF and a real invoice PDF, and this repo is public.
 *
 * IT RUNS THE SHIPPED EMITTER, not a copy of it. `ivaDeclarationXml` is loaded
 * out of v2/IvaExport.js into a bare context — the function touches no Apps
 * Script service, which is what makes that possible and is why it was written
 * that way. A local copy would prove a copy, and the copy would drift.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SOURCE = path.resolve(__dirname, '..', 'IvaExport.js');
const ctx = vm.createContext({ console });
vm.runInContext(fs.readFileSync(SOURCE, 'utf8'), ctx, { filename: 'IvaExport.js' });
const buildDeclaration = ctx.ivaDeclarationXml;
if (typeof buildDeclaration !== 'function') {
  throw new Error('ivaDeclarationXml not found in ' + SOURCE);
}

// ---------------------------------------------------------------------------
// Parsing the sample back into a model. Test-only — regex is fine on a file this
// app wrote itself, and pulling in a parser to read seven fields is not.
// ---------------------------------------------------------------------------

function one(xml, name) {
  const m = xml.match(new RegExp('<' + name + '\\s*>([^<]*)</' + name + '>'));
  return m ? m[1] : null;
}

function parseDeclaration(xml) {
  const model = {
    nifEntidade: one(xml, 'NifEntidade'),
    tipoEntidade: one(xml, 'TipoEntidade'),
    nifFuncionario: one(xml, 'NifFuncionario'),
    lines: []
  };

  const lineRe = /<Quadro02T1-Linha numero="(\d+)">([\s\S]*?)<\/Quadro02T1-Linha>/g;
  let m;
  while ((m = lineRe.exec(xml)) !== null) {
    const body = m[2];
    const fich = body.match(/<Ficheiro\s*([^>]*)>([^<]*)<\/Ficheiro>/);
    const attr = (name) => {
      const a = fich[1].match(new RegExp(name + '="([^"]*)"'));
      return a ? a[1] : null;
    };
    model.lines.push({
      numero: m[1],
      numeroFatura: one(body, 'NumeroFatura'),
      dataFatura: one(body, 'DataFatura'),
      nifFatura: one(body, 'NifFatura'),
      codigoBem: one(body, 'CodigoBem'),
      valorIVA: one(body, 'ValorIVA'),
      valorTotalFatura: one(body, 'ValorTotalFatura'),
      ficheiro: {
        attachmentDate: attr('attachmentDate'),
        fileName: attr('fileName'),
        extension: attr('extension'),
        base64: fich[2]
      }
    });
  }
  return model;
}

// ---------------------------------------------------------------------------

function stripNulPadding(base64) {
  const buf = Buffer.from(base64, 'base64');
  let end = buf.length;
  while (end > 0 && buf[end - 1] === 0) end--;
  return { base64: buf.slice(0, end).toString('base64'), stripped: buf.length - end, kept: end };
}

function firstDifference(a, b) {
  const len = Math.min(a.length, b.length);
  for (let i = 0; i < len; i++) {
    if (a[i] !== b[i]) return i;
  }
  return a.length === b.length ? -1 : len;
}

function context(buf, at) {
  const from = Math.max(0, at - 40);
  return JSON.stringify(buf.slice(from, at + 40).toString('latin1'));
}

function main() {
  const samplePath = process.argv[2] || path.join(process.env.HOME, 'test_iva');
  if (!fs.existsSync(samplePath)) {
    console.error('No sample at ' + samplePath);
    console.error('Pass the path to a declaration DRORIVA saved itself.');
    process.exit(1);
  }

  const outDir = path.dirname(path.resolve(samplePath));
  const repoRoot = path.resolve(__dirname, '..', '..');
  if (outDir === repoRoot || outDir.indexOf(repoRoot + path.sep) === 0) {
    console.error('Refusing to write inside the repo: a real sample carries a real');
    console.error('supplier NIF and a real invoice, and this repo is public.');
    console.error('Move the sample outside ' + repoRoot + ' and re-run.');
    process.exit(1);
  }

  const original = fs.readFileSync(samplePath);
  const model = parseDeclaration(original.toString('latin1'));

  console.log('Sample:  ' + samplePath + '  (' + original.length + ' bytes)');
  console.log('Lines:   ' + model.lines.length);
  for (const line of model.lines) {
    const blob = Buffer.from(line.ficheiro.base64, 'base64');
    console.log('  ' + line.numero + '. ' + line.numeroFatura + '  ' + line.dataFatura +
      '  NIF ' + line.nifFatura +
      '  CodigoBem ' + (line.codigoBem || '(absent)') +
      '  IVA ' + line.valorIVA + ' of ' + line.valorTotalFatura +
      '  ' + line.ficheiro.fileName + '.' + line.ficheiro.extension +
      ' (' + blob.length + ' bytes)');
  }

  // 1. Rebuild with the blob untouched. This must be byte-identical, and it is
  //    the only part of the proof that does not need the app.
  const rebuilt = Buffer.from(buildDeclaration(model), 'latin1');
  const rebuiltPath = samplePath + '.rebuilt';
  fs.writeFileSync(rebuiltPath, rebuilt);

  const diffAt = firstDifference(original, rebuilt);
  console.log('');
  if (diffAt === -1) {
    console.log('PASS  rebuild is byte-identical (' + rebuilt.length + ' bytes)');
  } else {
    console.log('FAIL  rebuild differs at byte ' + diffAt);
    console.log('  original: ' + context(original, diffAt));
    console.log('  rebuilt:  ' + context(rebuilt, diffAt));
    console.log('  sizes: ' + original.length + ' vs ' + rebuilt.length);
  }

  // 2. Same declaration, padding removed. Bytes cannot settle whether the app
  //    minds; only opening it can.
  const noPadModel = JSON.parse(JSON.stringify(model));
  let strippedTotal = 0;
  for (const line of noPadModel.lines) {
    const s = stripNulPadding(line.ficheiro.base64);
    line.ficheiro.base64 = s.base64;
    strippedTotal += s.stripped;
    console.log('  padding on line ' + line.numero + ': ' + s.stripped +
      ' NUL bytes after ' + s.kept + ' of payload');
  }
  const noPad = Buffer.from(buildDeclaration(noPadModel), 'latin1');
  const noPadPath = samplePath + '.nopad';
  fs.writeFileSync(noPadPath, noPad);

  console.log('');
  console.log('Wrote ' + rebuiltPath + '  (' + rebuilt.length + ' bytes)');
  console.log('Wrote ' + noPadPath + '  (' + noPad.length + ' bytes, ' +
    strippedTotal + ' NUL bytes dropped)');
  console.log('');
  console.log('Now open the .nopad file in DRORIVA. If every field comes back and');
  console.log('the receipt opens, unpadded blobs are safe and the format is proven.');

  process.exit(diffAt === -1 ? 0 : 1);
}

main();
