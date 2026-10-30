/*
 * Move the pinned /exec deployment onto what the server holds now.
 *
 *   npm run v2:deploy
 *
 * WHY THIS EXISTS. Deploying by hand meant typing
 * `clasp --user v2 deploy -i <deploymentId>` and looking the ID up. Leaving out
 * `-i` creates a SECOND deployment on a different URL and leaves the old one
 * live, and nothing stopped you deploying code the server did not actually
 * hold. This refuses, before it touches anything, when:
 *
 *   - V2_DEPLOYMENT_ID is missing from .env
 *   - v2/ has uncommitted changes (the order starts with a commit)
 *   - v2:verify finds the server does not match v2/
 *
 * The version is labelled with the latest commit, so list-deployments says
 * which commit is live. Afterwards it lists the deployments and warns unless
 * there are exactly two: @HEAD and the pinned version.
 *
 * Main project only. It never pushes and never forces. The Siri deployment ID
 * must never be put in .env under this key, or anywhere in the repo.
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const V2 = path.join(ROOT, 'v2');

function stop(message) {
  console.error('\n' + message + '\nNothing was deployed.\n');
  process.exit(1);
}

// Read the one key as text. .env is never handed to a shell, and no other value
// in it is printed.
function deploymentId() {
  const file = path.join(ROOT, '.env');
  if (!fs.existsSync(file)) return '';
  const line = fs.readFileSync(file, 'utf8')
    .split('\n')
    .find(entry => entry.trim().startsWith('V2_DEPLOYMENT_ID='));
  if (!line) return '';
  return line.trim().slice('V2_DEPLOYMENT_ID='.length).replace(/^["']|["']$/g, '').trim();
}

const id = deploymentId();
if (!id) {
  stop('V2_DEPLOYMENT_ID is not set in .env. Get it from `clasp --user v2 list-deployments`\n' +
       '(the pinned version, not @HEAD).');
}

const dirty = execFileSync('git', ['status', '--porcelain', '--', 'v2'], { cwd: ROOT, encoding: 'utf8' }).trim();
if (dirty) stop('v2/ has uncommitted changes — commit first:\n' + dirty);

try {
  execFileSync('node', [path.join(__dirname, 'verify-push.js')], { stdio: 'inherit' });
} catch (error) {
  stop('The server does not match v2/. Push (or push:force) and verify again.');
}

const label = execFileSync('git', ['log', '-1', '--format=%h %s'], { cwd: ROOT, encoding: 'utf8' })
  .trim().slice(0, 100);

console.log(`\nDeploying "${label}"…`);
try {
  execFileSync('clasp', ['--user', 'v2', 'deploy', '-i', id, '-d', label], { cwd: V2, stdio: 'inherit' });
} catch (error) {
  console.error('\nThe deploy failed. Check `clasp --user v2 list-deployments`.\n');
  process.exit(1);
}

const listing = execFileSync('clasp', ['--user', 'v2', 'list-deployments'], { cwd: V2, encoding: 'utf8' });
console.log('\n' + listing.trim());

const count = listing.split('\n').filter(line => line.trim().startsWith('- ')).length;
if (count !== 2) {
  console.error(`\nWARNING: ${count} deployments, expected exactly two (@HEAD and the pinned version).\n`);
  process.exit(1);
}

console.log('\nDeployed. Fully close the home-screen app (swipe it away) before checking.\n');
