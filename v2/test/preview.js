/**
 * v2 — render Index.html locally and screenshot it.
 *
 *   node v2/test/preview.js                       # every view, phone and desktop
 *   node v2/test/preview.js --view=health --width=390
 *   node v2/test/preview.js --open                # write the page, skip Chrome
 *
 * WHY THIS EXISTS. The harness cannot click, and it cannot see. It runs the
 * server against mocks and reads Index.html off disk only to prove the file is
 * there - so every line of the page's own CSS and JavaScript is invisible to
 * `npm run v2:test`. Before this, the first time a layout change was seen was on
 * the phone, after a push and a deploy.
 *
 * WHY IT USES THE REAL SERVER. The shapes the page renders - meta.columns,
 * row.options, row.files - come from uiSectionMeta and uiRow, and hand-written
 * fixtures drift from them silently. So this loads the same v2 sources the
 * harness does, against the same mocks, seeds plausible rows through the real
 * createEntry/setStatus, and inlines what uiBootstrap/uiListEntries/
 * uiListArchive/uiListSuppliers actually returned as a stub google.script.run.
 * If the server's shape changes, the preview changes with it.
 *
 * WHAT IT CANNOT TELL YOU. Headless Chrome is not iOS Safari. It will catch
 * overflow, a wrong breakpoint and a broken accordion; it will not reproduce the
 * native date wheel, momentum scrolling, or the tap-target feel. Those still
 * need the phone.
 *
 * Local only. v2/.claspignore allows twelve named files, so nothing under test/
 * is ever pushed.
 */
const fs = require('fs');
const vm = require('vm');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const DIR = path.join(__dirname, '..');
const OUT = path.join(os.tmpdir(), 'hf-preview');
const CHROME = process.env.CHROME_PATH ||
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

/* ------------------------------- arguments -------------------------------- */

const argv = process.argv.slice(2);
const flag = name => {
  const hit = argv.filter(a => a.indexOf('--' + name + '=') === 0)
    .map(a => a.slice(name.length + 3));
  return hit;
};
const VIEWS = flag('view').length ? flag('view') : ['work', 'iva', 'health', 'income', 'providers'];
const WIDTHS = flag('width').length ? flag('width').map(Number) : [390, 1100];
const PAGE_ONLY = argv.indexOf('--open') !== -1;
// --click='[data-toggle="__archive"]' --click='#dialogPrimary', in order.
const CLICKS = flag('click');
const LABEL = flag('label')[0] || '';
/*
 * Frame height. Tall by default so a whole list fits in one image; pass a real
 * handset height when looking at an overlay, because the dialogs are
 * position:fixed and centre themselves in the viewport - in a 2400px frame they
 * land a thousand pixels below anything you can see on a phone.
 */
const HEIGHT = Number(flag('height')[0] || 2400);

/* --------------------------- the real v2 source --------------------------- */

const mocks = require('./mocks.js');
const FILES = ['Config.js', 'Core.js', 'Entries.js', 'Registry.js', 'Setup.js',
  'Smoke.js', 'Web.js', 'Form.js', 'Manage.js', 'Suppliers.js', 'Siri.js'];

const sandbox = Object.assign({ console }, mocks);
sandbox.globalThis = sandbox;
const ctx = vm.createContext(sandbox);
vm.runInContext(FILES.map(f => fs.readFileSync(path.join(DIR, f), 'utf8')).join('\n;\n'),
  ctx, { filename: 'v2-concat.js' });
const G = sandbox;

/* ------------------------------- the fixture ------------------------------ */

/*
 * Stand-in patients, for the same reason the harness uses them: the real list is
 * a Script Property because these are family members and the repository public.
 */
mocks._props.HEALTH_PATIENTS = 'Jackie, Kit, Auryn, Phoenix';
G.bootstrap();

/*
 * Rows chosen to stress the layout rather than to be tidy: a provider name too
 * long for a phone, a note long enough to wrap, one row per status so every
 * accordion has something in it, a Health row with BOTH documents and another
 * with only one - which is the case that currently says nothing at all.
 */
const seed = [];

seed.push(G.createEntry('health', {
  'Date': '2026-07-28', 'Counterparty': 'Clínica Dentária de São João', 'Amount': 240.5,
  'Currency': 'EUR', 'Patient': 'Auryn', 'Type': 'Dentist', 'Invoice Date': '2026-07-28',
  'Notes': 'Two fillings and the follow-up X-ray, claimed under the annual dental limit',
  'Receipt Medium': 'Paper',
  'Justification URL': 'https://drive.google.com/file/d/' + 'j'.repeat(28) + '/view',
  'Receipt URL': 'https://drive.google.com/file/d/' + 'r'.repeat(28) + '/view'
}, 'form'));
G.setStatus('health', seed[0].row, 'Claimed', '2026-08-01');

// One document of two. The gap this preview exists to show.
seed.push(G.createEntry('health', {
  'Date': '2026-08-03', 'Counterparty': 'Farmácia Central', 'Amount': 18.4,
  'Currency': 'EUR', 'Patient': 'Kit', 'Type': 'Prescription', 'Invoice Date': '2026-08-03',
  'Receipt Medium': 'Email',
  'Justification URL': 'https://drive.google.com/file/d/' + 'p'.repeat(28) + '/view'
}, 'siri'));

seed.push(G.createEntry('health', {
  'Date': '2026-08-11', 'Counterparty': 'Dr. Marta Oliveira', 'Amount': 75,
  'Currency': 'EUR', 'Patient': 'Jackie', 'Type': 'Doctor', 'Invoice Date': '2026-08-10',
  'Notes': 'Annual check', 'Receipt Medium': 'Paper'
}, 'form'));
G.setStatus('health', seed[2].row, 'Claimed', '2026-08-12');
G.setStatus('health', seed[2].row, 'Settled', '2026-08-14');

// Archived, so the fourth accordion is not empty.
const gone = G.createEntry('health', {
  'Date': '2026-05-02', 'Counterparty': 'Optica Lisboa', 'Amount': 129,
  'Currency': 'EUR', 'Patient': 'Phoenix', 'Type': 'Optician', 'Invoice Date': '2026-05-02',
  'Receipt Medium': 'Paper'
}, 'form');
G.archiveEntry('health', gone.row, 'deleted');

G.createEntry('work', {
  'Date': '2026-08-09', 'Counterparty': 'Bolt', 'Amount': 8.2, 'Currency': 'EUR',
  'Expense Reason': 'Client visit', 'Notes': 'Taxi to the Lisbon office',
  'Receipt Medium': 'Email'
}, 'siri');
const wk = G.createEntry('work', {
  'Date': '2026-07-14', 'Counterparty': 'Ementa Restaurante', 'Amount': 46.9,
  'Currency': 'EUR', 'Expense Reason': 'Client lunch', 'Receipt Medium': 'Paper',
  'Receipt URL': 'https://drive.google.com/file/d/' + 'w'.repeat(28) + '/view'
}, 'form');
G.setStatus('work', wk.row, 'Claimed', '2026-07-20');

G.createEntry('iva', {
  'Date': '2026-08-05', 'Counterparty': 'Galp', 'Amount': 62.3, 'Currency': 'EUR',
  'Receipt Medium': 'Paper'
}, 'siri');

const inc = G.createEntry('income', {
  'Date': '2026-07-01', 'Counterparty': 'JALLC', 'Amount': 3200, 'Currency': 'EUR',
  'Reason': 'August retainer'
}, 'form');
G.setStatus('income', inc.row, 'Received', '2026-07-31');

/*
 * A row holding a status that is not one of the three. Written straight to the
 * sheet because nothing in the app can produce it - which is the point: this is
 * what a hand edit or a half-finished Siri run leaves behind, and it is the row
 * the four-way partition would otherwise drop on the floor.
 */
(function oddStatus() {
  const sheet = mocks._ss.getSheetByName('Work');
  const cols = G.resolveColumns(sheet);
  const row = G.createEntry('work', {
    'Date': '2026-06-06', 'Counterparty': 'Unknown Vendor', 'Amount': 12,
    'Currency': 'EUR', 'Expense Reason': 'Client visit'
  }, 'form').row;
  sheet.getRange(row, cols['Status']).setValue('Pending?');
})();

/* ------------------------------- the payload ------------------------------ */

const SECTIONS = ['work', 'iva', 'health', 'income'];
const payload = { uiBootstrap: G.uiBootstrap(), entries: {}, archive: {}, suppliers: G.uiListSuppliers() };
SECTIONS.forEach(key => {
  payload.entries[key] = G.uiListEntries(key);
  payload.archive[key] = G.uiListArchive(key);
});

if (argv.indexOf('--dump') !== -1) {
  console.log(JSON.stringify(payload, null, 2));
}

/* --------------------------------- the page ------------------------------- */

/*
 * Only what the page actually calls. Anything not stubbed reports itself through
 * the failure handler rather than hanging, so a missing stub reads as an error on
 * screen instead of a page stuck on "Loading...".
 */
function stub(view) {
  return `
<script>
/*
 * A thrown error in the page's own script leaves a blank white screenshot and
 * nothing to go on, so it is painted onto the page instead. This is the preview
 * harness's equivalent of the loading/empty/error rule the page itself keeps.
 */
window.addEventListener('error', function (e) {
  var box = document.createElement('pre');
  box.style.cssText = 'background:#fdecea;color:#b3261e;padding:1rem;margin:0;' +
    'white-space:pre-wrap;font:12px/1.4 ui-monospace,monospace;position:relative;z-index:99';
  box.textContent = 'PAGE ERROR\\n' + (e.message || e.error) +
    '\\n' + (e.filename || '') + ':' + (e.lineno || '') +
    '\\n\\n' + ((e.error && e.error.stack) || '');
  (document.body || document.documentElement).prepend(box);
});
window.google = { script: {
  url: { getLocation: function (cb) { cb({ parameter: {} }); } },
  run: (function () {
    var DATA = ${JSON.stringify(payload)};
    var h = {};
    function reply(value) { var f = h.ok; setTimeout(function () { f(value); }, 0); }
    function fail(msg) { var f = h.err; setTimeout(function () { f(new Error(msg)); }, 0); }
    var api = {
      withSuccessHandler: function (f) { h.ok = f; return api; },
      withFailureHandler: function (f) { h.err = f; return api; },
      uiBootstrap: function () { reply(DATA.uiBootstrap); },
      uiListEntries: function (k) { reply(DATA.entries[k]); },
      uiListArchive: function (k) { reply(DATA.archive[k]); },
      uiListSuppliers: function () { reply(DATA.suppliers); },
      uiCategoryValues: function () { reply({ ok: true, values: [] }); },
      uiStagingFiles: function () { reply({ ok: true, files: [] }); },

      /*
       * Enough of setStatus to exercise the page: move the row, stamp the date,
       * hand back the entry. Not the server's logic - it does not rename files
       * or clear later dates - but the shape it returns is the server's, so the
       * follow-the-row behaviour can be driven and seen.
       */
      uiSetStatus: function (section, rowNumber, state, dateISO) {
        var data = DATA.entries[section];
        var row = data.rows.filter(function (r) { return r.row === Number(rowNumber); })[0];
        if (!row) return fail('preview: no row ' + rowNumber + ' in ' + section);
        var names = data.meta.states.map(function (s) { return s.name; });
        var target = data.meta.states[names.indexOf(state)];
        row.status = state;
        row.statusIndex = names.indexOf(state);
        if (target && target.dateColumn) row.dates[target.dateColumn] = dateISO || '';
        reply({ ok: true, entry: row, date: dateISO || '', fileErrors: [] });
      },

      uiArchiveEntry: function (section, rowNumber) {
        var data = DATA.entries[section];
        var keep = [];
        data.rows.forEach(function (r) {
          if (r.row === Number(rowNumber)) {
            r.archivedAt = '2026-08-15';
            r.reason = '';
            r.options = [];
            DATA.archive[section].rows.unshift(r);
          } else { keep.push(r); }
        });
        data.rows = keep;
        reply({ ok: true, fileErrors: [] });
      },

      uiRestoreEntry: function (section, rowNumber) {
        var arc = DATA.archive[section];
        var keep = [];
        arc.rows.forEach(function (r) {
          if (r.row === Number(rowNumber)) {
            delete r.archivedAt;
            delete r.reason;
            delete r.archived;
            DATA.entries[section].rows.unshift(r);
          } else { keep.push(r); }
        });
        arc.rows = keep;
        reply({ ok: true, fileErrors: [] });
      }
    };
    return new Proxy(api, { get: function (t, k) {
      if (k in t) return t[k];
      return function () { fail('preview: ' + String(k) + ' is not stubbed'); };
    } });
  })()
} };
window.PREVIEW_VIEW = ${JSON.stringify(view)};
window.PREVIEW_CLICKS = ${JSON.stringify(CLICKS)};
</script>
`;
}

/*
 * Drives the page to the view under test after boot. select() is a global in the
 * page's own script, so this needs no hooks in Index.html - the preview stays
 * something the page knows nothing about.
 */
/*
 * Drives the page to the view under test, then audits it for horizontal
 * overflow.
 *
 * The audit is here because the eye is bad at this and a screenshot is worse: a
 * page 40px too wide and a page 400px too wide look identical once the content
 * is clipped at the frame, and the element actually responsible is usually not
 * the one you can see sticking out. It reports the widest offenders by how far
 * past the viewport they reach, naming each one enough to find it in the CSS.
 *
 * Written into the DOM rather than logged, because headless Chrome's console
 * needs a debugger connection to read and --dump-dom does not.
 */
const DRIVE = `
<script>
(function () {
  var tries = 0;
  var timer = setInterval(function () {
    if (++tries > 60) return clearInterval(timer);
    if (typeof select !== 'function' || !window.app || !app.sections.length) return;
    clearInterval(timer);
    if (PREVIEW_VIEW && PREVIEW_VIEW !== 'default') select(PREVIEW_VIEW);
    setTimeout(clickThrough, 400);
  }, 25);

  /*
   * Clicks, in order, with a gap for each render to settle.
   *
   * This is what makes the interactive paths testable at all: expanding the
   * archive, advancing a status through the date dialog, following the row into
   * its new section. Every one of those is invisible to the harness, and a
   * still screenshot of the page as it first loads never reaches them.
   *
   * A selector that matches nothing is reported rather than skipped - a step
   * that quietly did nothing would make the rest of the sequence a lie.
   */
  function clickThrough() {
    var steps = (PREVIEW_CLICKS || []).slice();
    var missed = [];

    (function next() {
      if (!steps.length) return audit(missed);
      var selector = steps.shift();
      var node = document.querySelector(selector);
      if (node) { node.click(); } else { missed.push(selector); }
      setTimeout(next, 350);
    })();
  }

  function name(node) {
    return node.tagName.toLowerCase() +
      (node.id ? '#' + node.id : '') +
      (node.className && typeof node.className === 'string'
        ? '.' + node.className.trim().split(/\\s+/).join('.') : '');
  }

  /*
   * A node inside a scrolling ancestor is not overflowing anything - the wide
   * desktop table is meant to be wider than its wrapper, and getBoundingClientRect
   * reports its full width whether or not the wrapper clips it. Counting those
   * buried the one real finding under sixty rows of table cells.
   */
  function scrolls(node) {
    for (var p = node.parentElement; p && p !== document.body; p = p.parentElement) {
      var overflow = getComputedStyle(p).overflowX;
      if (overflow === 'auto' || overflow === 'scroll' || overflow === 'hidden') return true;
    }
    return false;
  }

  function audit(missed) {
    var vw = document.documentElement.clientWidth;
    var over = [];
    Array.prototype.forEach.call(document.querySelectorAll('body *'), function (node) {
      // Fixed-position layers are allowed to be their own size, and a hidden
      // node's box is not on screen to overflow anything.
      if (getComputedStyle(node).position === 'fixed') return;
      if (!node.offsetParent && node.offsetWidth === 0) return;
      if (scrolls(node)) return;
      var right = node.getBoundingClientRect().right;
      if (right > vw + 0.5) over.push({ n: name(node), by: Math.round(right - vw) });
    });
    over.sort(function (a, b) { return b.by - a.by; });

    var report = document.createElement('pre');
    report.id = 'overflowReport';
    report.style.display = 'none';
    report.textContent = 'VIEWPORT ' + vw +
      ' DOCUMENT ' + document.documentElement.scrollWidth + '\\n' +
      ((missed && missed.length)
        ? '  CLICK MATCHED NOTHING: ' + missed.join(', ') + '\\n' : '') +
      (over.length
        ? over.slice(0, 14).map(function (o) { return '  +' + o.by + 'px  ' + o.n; }).join('\\n')
        : '  no horizontal overflow');
    document.body.appendChild(report);
  }
})();
</script>
`;

const source = fs.readFileSync(path.join(DIR, 'Index.html'), 'utf8');

function build(view) {
  return source
    // Apps Script adds this server-side via addMetaTag; the local file needs it
    // spelled out or every width renders as a 980px desktop.
    .replace('<head>', '<head>\n  <meta name="viewport" content="width=device-width, initial-scale=1">')
    .replace('<script>', stub(view) + '  <script>')
    .replace('</body>', DRIVE + '</body>');
}

/* -------------------------------- capture -------------------------------- */

/*
 * WHY AN IFRAME. Chrome's --window-size will not go below 500 CSS pixels - ask
 * for 390 and you get a 500px layout, and --screenshot then CROPS the image to
 * 390. The result looks exactly like a page overflowing its viewport, which is
 * the very bug this tool is for, so the tool would have been inventing them.
 * --force-device-scale-factor does not help; it changes pixel density, not the
 * viewport. Neither headless mode lifts the floor.
 *
 * An iframe has its own viewport, and media queries inside it answer to the
 * frame's width rather than the window's. So the app is rendered in a frame of
 * exactly the width under test, inside a host page wide enough for Chrome to
 * accept, and the screenshot is cropped back to the frame. The app genuinely
 * believes it is 390 wide, because it is.
 *
 * Both files are file:// URLs, hence --allow-file-access-from-files: without it
 * Chrome gives each its own opaque origin and the host cannot read the report
 * the frame leaves for it.
 */
function hostPage(pageFile, width) {
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  html, body { margin: 0; background: #202430; }
  iframe { display: block; border: 0; width: ${width}px; height: ${HEIGHT}px; background: #fff; }
</style></head>
<body>
<iframe id="frame" src="${path.basename(pageFile)}"></iframe>
<pre id="overflowReport" style="display:none"></pre>
<script>
(function () {
  var tries = 0;
  var timer = setInterval(function () {
    if (++tries > 120) {
      clearInterval(timer);
      document.getElementById('overflowReport').textContent =
        'no report - the frame never finished (cross-origin? missing --allow-file-access-from-files?)';
      return;
    }
    try {
      var doc = document.getElementById('frame').contentDocument;
      var found = doc && doc.getElementById('overflowReport');
      var error = doc && doc.querySelector('pre');
      if (found) {
        clearInterval(timer);
        document.getElementById('overflowReport').textContent = found.textContent;
      } else if (error && /^PAGE ERROR/.test(error.textContent)) {
        clearInterval(timer);
        document.getElementById('overflowReport').textContent = error.textContent;
      }
    } catch (e) { /* not readable yet */ }
  }, 50);
})();
</script>
</body></html>`;
}

fs.mkdirSync(OUT, { recursive: true });
const written = [];

VIEWS.forEach(view => {
  const file = path.join(OUT, 'page-' + view + '.html');
  fs.writeFileSync(file, build(view));
  written.push(file);
  if (PAGE_ONLY) return;

  WIDTHS.forEach(width => {
    const shot = path.join(OUT, view + (LABEL ? '-' + LABEL : '') + '-' + width + '.png');
    const host = path.join(OUT, 'host-' + view + '-' + width + '.html');
    fs.writeFileSync(host, hostPage(file, width));

    const common = [
      '--headless=new', '--disable-gpu', '--no-sandbox', '--hide-scrollbars',
      '--allow-file-access-from-files',
      // The host only has to clear Chrome's 500px floor; the frame inside it is
      // the width that matters, and the shot is cropped back to it.
      '--window-size=' + Math.max(width, 500) + ',' + HEIGHT,
      '--virtual-time-budget=6000'
    ];

    try {
      execFileSync(CHROME, common.concat(['--screenshot=' + shot, 'file://' + host]),
        { stdio: ['ignore', 'ignore', 'pipe'] });
      written.push(shot);
    } catch (e) {
      console.error('  chrome failed for ' + view + '@' + width + ': ' + e.message);
      return;
    }

    try {
      const dom = execFileSync(CHROME, common.concat(['--dump-dom', 'file://' + host]),
        { stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 64 * 1024 * 1024 }).toString();
      const hit = dom.match(/<pre id="overflowReport"[^>]*>([\s\S]*?)<\/pre>/);
      const body = hit
        ? hit[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim()
        : '(no report — the host could not read the frame)';
      console.log('\n' + view + '@' + width + '\n' +
        body.split('\n').map(l => '  ' + l.trim()).join('\n'));
    } catch (e) {
      console.error('  audit failed for ' + view + '@' + width);
    }
  });
});

console.log('\n' + written.map(f => '  ' + f).join('\n'));
console.log('\n' + OUT);
