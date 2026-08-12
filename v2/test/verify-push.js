/*
 * Does the server actually hold what v2/ holds?
 *
 * WHY THIS EXISTS. `clasp push` is not trustworthy on its own. When the remote
 * appsscript.json differs from the local one, clasp asks before overwriting it;
 * with no TTY that prompt defaults to no, and the push is abandoned - but clasp
 * 3.0.6-alpha still prints "Pushed 9 files." and lists them. A push that did
 * nothing is indistinguishable from one that worked, by output alone.
 *
 * That is not hypothetical: checkDocuments() sat unpushed through a whole
 * session because of it, while the plan told you to go and run it. The trigger
 * was a MISSING TRAILING NEWLINE in the server's copy of the manifest. Nothing
 * semantic - one byte.
 *
 * So the only honest check is to fetch what the server has and compare it. This
 * pulls into a temporary directory - never over v2/ - and diffs every file clasp
 * would push.
 *
 *   npm run v2:verify
 *
 * A mismatch is fixed with `npm run v2:push:force`, which also ends the loop:
 * once the server holds the local manifest byte for byte, plain pushes stop
 * being refused.
 *
 * TWO PROJECTS. Since the Siri endpoint exists there is a second, separate
 * Apps Script project — v2-siri/, the anonymous shim. Pass a directory to check
 * that one instead:
 *
 *   npm run v2:verify        the main project
 *   npm run v2:siri:verify   the shim
 *
 * The shim is two files and changes almost never, which is exactly why it needs
 * checking: nobody would notice it had stopped matching.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const V2 = path.resolve(__dirname, '..', process.argv[2] || '.');
const LABEL = path.basename(V2);

/*
 * The files clasp pushes.
 *
 * Read from .claspignore rather than guessed from the directory listing. Both
 * projects write their ignore file as "everything, then !name per allowed
 * file", so the allowlist is exact — and it has to be: v2-siri/ holds
 * build-manifest.js and appsscript.template.json, which are local tooling.
 * Inferring the list from the extension would report those as missing from the
 * server for ever, and a check that always fails is one that gets ignored.
 */
function allowedByClaspignore(dir) {
  const file = path.join(dir, '.claspignore');
  if (!fs.existsSync(file)) return null;

  const allowed = fs.readFileSync(file, 'utf8')
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.startsWith('!'))
    .map(line => line.slice(1));

  return allowed.length ? allowed.sort() : null;
}

function pushedFiles(dir, allowlist) {
  const present = fs.readdirSync(dir)
    .filter(name => !name.startsWith('.'))
    .filter(name => fs.statSync(path.join(dir, name)).isFile());

  if (allowlist) return present.filter(name => allowlist.indexOf(name) !== -1).sort();

  return present
    .filter(name => ['.js', '.html', '.json'].indexOf(path.extname(name)) !== -1)
    .sort();
}

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'v2-verify-'));
fs.copyFileSync(path.join(V2, '.clasp.json'), path.join(tmp, '.clasp.json'));

try {
  execFileSync('clasp', ['--user', 'v2', 'pull'], { cwd: tmp, stdio: 'pipe' });
} catch (error) {
  console.error('Could not pull from the server:\n' + (error.stderr || error.message));
  process.exit(2);
}

const allowlist = allowedByClaspignore(V2);
const local = pushedFiles(V2, allowlist);
const differences = [];

local.forEach(name => {
  const remote = path.join(tmp, name);
  if (!fs.existsSync(remote)) {
    differences.push(`${name} — not on the server at all`);
    return;
  }
  if (!fs.readFileSync(path.join(V2, name)).equals(fs.readFileSync(remote))) {
    differences.push(`${name} — server copy differs`);
  }
});

// The pulled copy has no .claspignore, so list it by extension: anything the
// server holds that the allowlist does not name is a real difference — a file
// left behind by an earlier push, which is exactly what this should catch.
pushedFiles(tmp, null).forEach(name => {
  if (local.indexOf(name) === -1) differences.push(`${name} — on the server but not in ${LABEL}/`);
});

fs.rmSync(tmp, { recursive: true, force: true });

if (differences.length) {
  console.error(`\nServer does NOT match ${LABEL}/ (${differences.length}):\n`);
  differences.forEach(line => console.error('  ' + line));
  console.error('\nRun the matching push:force, then this again.\n');
  process.exit(1);
}

console.log(`Server matches ${LABEL}/ — ${local.length} files, byte for byte.`);
