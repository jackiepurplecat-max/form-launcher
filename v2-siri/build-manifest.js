/*
 * Generate v2-siri/appsscript.json from the template.
 *
 *   npm run v2:siri:manifest
 *
 * WHY THIS IS GENERATED. The shim reaches the main project's code as a library,
 * so its manifest has to name the main project's script id — and every clasp
 * identifier in this repo is gitignored, because the repo is public. So the
 * committed file is the template with a placeholder, and the real manifest is
 * built from v2/.clasp.json, which is the copy that already exists locally.
 * Nothing new to configure and nothing that can drift from the project clasp
 * actually pushes to.
 *
 * NO TRAILING NEWLINE, on purpose. Google stores appsscript.json without one,
 * and when the remote manifest differs from the local one by so much as a byte
 * clasp abandons the push while still reporting success. v2/appsscript.json is
 * kept at exactly 425 bytes for the same reason. See NEXT-SESSION.md.
 */

const fs = require('fs');
const path = require('path');

const HERE = __dirname;
const mainClasp = path.join(HERE, '..', 'v2', '.clasp.json');

if (!fs.existsSync(mainClasp)) {
  console.error(
    'Cannot find v2/.clasp.json, so there is no main script id to point the library at.\n' +
    'That file is gitignored; on a fresh clone, clone the main project first.'
  );
  process.exit(1);
}

const scriptId = JSON.parse(fs.readFileSync(mainClasp, 'utf8')).scriptId;
if (!scriptId) {
  console.error('v2/.clasp.json has no scriptId.');
  process.exit(1);
}

const template = fs.readFileSync(path.join(HERE, 'appsscript.template.json'), 'utf8');
const out = template.replace('{{MAIN_SCRIPT_ID}}', scriptId).replace(/\n+$/, '');

fs.writeFileSync(path.join(HERE, 'appsscript.json'), out);
console.log(`Wrote v2-siri/appsscript.json — library HF -> ${scriptId} (development mode), ${out.length} bytes, no trailing newline.`);
