/* Minimal Apps Script stand-ins: enough to actually run v2 end to end. */

const Logger = { entries: [], log(m) { this.entries.push(String(m)); } };

const Session = { getScriptTimeZone: () => 'UTC' };

function _pad(n, w) { return String(n).padStart(w, '0'); }

const Utilities = {
  formatDate(date, tz, fmt) {
    const d = new Date(date);
    if (isNaN(d.getTime())) throw new Error('Utilities.formatDate: invalid date');
    const map = {
      'yyyy-MM-dd': `${d.getUTCFullYear()}-${_pad(d.getUTCMonth() + 1, 2)}-${_pad(d.getUTCDate(), 2)}`,
      'dd-MM-yyyy': `${_pad(d.getUTCDate(), 2)}-${_pad(d.getUTCMonth() + 1, 2)}-${d.getUTCFullYear()}`,
      'yyMMdd': `${String(d.getUTCFullYear()).slice(2)}${_pad(d.getUTCMonth() + 1, 2)}${_pad(d.getUTCDate(), 2)}`
    };
    if (!map[fmt]) throw new Error('unhandled format ' + fmt);
    return map[fmt];
  },
  newBlob(content, type, name) { return { content, type, name }; }
};

/* ------------------------------- Sheets ---------------------------------- */

class Range {
  constructor(sheet, row, col, nr, nc) {
    Object.assign(this, { sheet, row, col, nr, nc });
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nr; r++) {
      const line = [];
      for (let c = 0; c < this.nc; c++) line.push(this.sheet._get(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(values) {
    if (values.length !== this.nr || values[0].length !== this.nc) {
      throw new Error(`setValues shape mismatch: got ${values.length}x${values[0].length}, range is ${this.nr}x${this.nc}`);
    }
    values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
    return this;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
}

let _gid = 0;
class Sheet {
  constructor(name) { this.name = name; this.cells = new Map(); this.frozen = 0; this.gid = ++_gid; }
  getName() { return this.name; }
  getSheetId() { return this.gid; }
  _key(r, c) { return r + ':' + c; }
  _get(r, c) { const v = this.cells.get(this._key(r, c)); return v === undefined ? '' : v; }
  _set(r, c, v) {
    if (r < 1 || c < 1) throw new Error(`bad cell reference row=${r} col=${c}`);
    if (v === '' || v === null || v === undefined) this.cells.delete(this._key(r, c));
    else this.cells.set(this._key(r, c), v);
  }
  _extent() {
    let maxR = 0, maxC = 0;
    for (const key of this.cells.keys()) {
      const [r, c] = key.split(':').map(Number);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    }
    return { maxR, maxC };
  }
  getLastRow() { return this._extent().maxR; }
  getLastColumn() { return this._extent().maxC; }
  getRange(row, col, nr, nc) { return new Range(this, row, col, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc); }
  setFrozenRows(n) { this.frozen = n; return this; }
  deleteRow(row) {
    const next = new Map();
    for (const [key, v] of this.cells) {
      const [r, c] = key.split(':').map(Number);
      if (r === row) continue;
      next.set((r > row ? r - 1 : r) + ':' + c, v);
    }
    this.cells = next;
    return this;
  }
}

class Spreadsheet {
  constructor(name) { this.name = name; this.sheets = []; }
  getName() { return this.name; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }
}

const _ss = new Spreadsheet('HelpfulForms v2');
_ss.getUrl = () => 'https://docs.google.com/spreadsheets/d/TESTSHEETID/edit';
const SpreadsheetApp = { getActiveSpreadsheet: () => _ss, flush() { this.flushes = (this.flushes || 0) + 1; } };

/* -------------------------------- Drive ---------------------------------- */

let _driveId = 0;
class DFile {
  constructor(name, parent) { this.id = 'file' + (++_driveId) + 'x'.repeat(22); this.name = name; this.parent = parent; }
  getName() { return this.name; }
  getId() { return this.id; }
  setName(n) { this.name = n; return this; }
  moveTo(folder) { this.parent = folder; return this; }
  getBlob() { return { _blob: this.name }; }
  getParents() { const p = this.parent ? [this.parent] : []; let i = 0; return { hasNext: () => i < p.length, next: () => p[i++] }; }
  setTrashed(v) { this.trashed = !!v; return this; }
}
class DFolder {
  constructor(name) { this.id = 'fold' + (++_driveId) + 'y'.repeat(22); this.name = name; this.children = []; }
  getName() { return this.name; }
  getId() { return this.id; }
  getUrl() { return 'https://drive.test/' + this.id; }
  createFolder(n) { const f = new DFolder(n); this.children.push(f); _folders[f.id] = f; return f; }
  getFoldersByName(n) {
    const hits = this.children.filter(c => c instanceof DFolder && c.name === n);
    let i = 0;
    return { hasNext: () => i < hits.length, next: () => hits[i++] };
  }
  path() { return this.name; }
}
const _folders = {}, _files = {};
const DriveApp = {
  createFolder(n) { const f = new DFolder(n); _folders[f.id] = f; return f; },
  getFolderById(id) { const f = _folders[id]; if (!f) throw new Error('No folder ' + id); return f; },
  getFileById(id) { const f = _files[id]; if (!f) throw new Error('No file ' + id); return f; },
  _addFile(name) { const f = new DFile(name, null); _files[f.id] = f; return f; },
  createFile(blob) { return DriveApp._addFile(blob && blob.name ? blob.name : 'untitled'); }
};

/* ------------------------------ Properties -------------------------------- */

const _props = {};
const PropertiesService = {
  getScriptProperties: () => ({
    getProperty: k => (_props[k] === undefined ? null : _props[k]),
    setProperty: (k, v) => { _props[k] = v; },
    getProperties: () => Object.assign({}, _props)
  })
};

/* -------------------------------- Misc ----------------------------------- */

const MailApp = { sent: [], sendEmail(to, subject, body, opts) { this.sent.push({ to, subject, body, opts }); } };

let _lockHeld = false;
const LockService = {
  getScriptLock: () => ({
    tryLock(ms) { if (_lockHeld) return false; _lockHeld = true; return true; },
    releaseLock() { _lockHeld = false; }
  })
};

module.exports = {
  Logger, Session, Utilities, SpreadsheetApp, DriveApp, PropertiesService,
  MailApp, LockService, _props, _files, _folders, _ss
};
