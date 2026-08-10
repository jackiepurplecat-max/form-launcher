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
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const V2 = path.join(__dirname, '..');

/* The files clasp pushes: top-level source, no dotfiles, no test directory. */
function pushedFiles(dir) {
  return fs.readdirSync(dir)
    .filter(name => !name.startsWith('.'))
    .filter(name => fs.statSync(path.join(dir, name)).isFile())
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

const local = pushedFiles(V2);
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

pushedFiles(tmp).forEach(name => {
  if (local.indexOf(name) === -1) differences.push(`${name} — on the server but not in v2/`);
});

fs.rmSync(tmp, { recursive: true, force: true });

if (differences.length) {
  console.error(`\nServer does NOT match v2/ (${differences.length}):\n`);
  differences.forEach(line => console.error('  ' + line));
  console.error('\nRun `npm run v2:push:force`, then this again.\n');
  process.exit(1);
}

console.log(`Server matches v2/ — ${local.length} files, byte for byte.`);
