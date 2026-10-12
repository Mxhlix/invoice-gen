'use strict';
/**
 * GAS(Google Apps Script)グローバルAPIの最小モック（invoice-gen 用）。
 * Code.gs を Node 上でそのまま動かして、自動発行・実行ログ・エラー通知の経路を検証する。
 * Node標準のみ（外部パッケージなし）。
 */

class Range {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col;
    this.numRows = numRows || 1; this.numCols = numCols || 1;
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this.numRows; i++) {
      const rowArr = this.sheet.data[this.row - 1 + i] || [];
      const vals = [];
      for (let j = 0; j < this.numCols; j++) {
        const v = rowArr[this.col - 1 + j];
        vals.push(v === undefined ? '' : v);
      }
      out.push(vals);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(values) {
    for (let i = 0; i < this.numRows; i++) {
      const rowIdx = this.row - 1 + i;
      while (this.sheet.data.length <= rowIdx) this.sheet.data.push([]);
      const rowArr = this.sheet.data[rowIdx];
      for (let j = 0; j < this.numCols; j++) {
        const colIdx = this.col - 1 + j;
        while (rowArr.length <= colIdx) rowArr.push('');
        rowArr[colIdx] = values[i][j];
      }
    }
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  setFontWeight() { return this; }
}

class OverGridImage {
  constructor(script) { this.script = script || null; this.width = 0; this.height = 0; }
  setWidth(w) { this.width = w; return this; }
  setHeight(h) { this.height = h; return this; }
  assignScript(fn) { this.script = fn; return this; }
  getScript() { return this.script; }
}

class Sheet {
  constructor(name) { this.name = name; this.data = []; this.images = []; }
  getName() { return this.name; }
  getRange(row, col, numRows, numCols) { return new Range(this, row, col, numRows, numCols); }
  getDataRange() {
    return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }
  getLastRow() { return this.data.length; }
  getLastColumn() { let m = 0; this.data.forEach(r => { if (r.length > m) m = r.length; }); return m; }
  appendRow(values) { this.data.push(values.slice()); }
  setFrozenRows() { return this; }
  getImages() { return this.images; }
  insertImage(blob, col, row) {
    const img = new OverGridImage(null);
    img.blob = blob; img.col = col; img.row = row;
    this.images.push(img);
    return img;
  }
  // テスト用: ヘッダ行を除いた行配列
  rows() { return this.data.slice(1); }
}

class Spreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  insertSheet(name) { const s = new Sheet(name); this.sheets[name] = s; return s; }
  getSheets() { return Object.keys(this.sheets).map(k => this.sheets[k]); }
}

class DriveFile {
  constructor(name) { this.name = name; }
  getName() { return this.name; }
  getUrl() { return 'https://drive.example.test/file/' + encodeURIComponent(this.name); }
}

class DriveFolder {
  constructor(name) { this.name = name; this.files = []; }
  getName() { return this.name; }
  createFile(blob) {
    const f = new DriveFile(blob && blob.getName ? blob.getName() : 'untitled');
    this.files.push(f);
    return f;
  }
}

/** 固定時刻の Date。引数なしの new Date() だけを nowMs に固定し、他は素の Date と同じ。 */
function makeFixedDate(nowMs) {
  const RealDate = Date;
  class FixedDate extends RealDate {
    constructor(...args) {
      if (args.length === 0) super(nowMs);
      else super(...args);
    }
    static now() { return nowMs; }
  }
  return FixedDate;
}

/** Utilities.formatDate の簡易版（yyyy/MM/dd/HH/mm/ss のみ。タイムゾーンはローカル扱い）。 */
function formatDate(d, tz, fmt) {
  const p = (n, w) => String(n).padStart(w || 2, '0');
  const map = {
    yyyy: String(d.getFullYear()),
    MM: p(d.getMonth() + 1),
    dd: p(d.getDate()),
    HH: p(d.getHours()),
    mm: p(d.getMinutes()),
    ss: p(d.getSeconds())
  };
  return String(fmt).replace(/yyyy|MM|dd|HH|mm|ss/g, t => map[t]);
}

/**
 * サンドボックスを組み立てる。
 * opts:
 *   nowMs           … new Date() が返す固定時刻（省略時は実時刻）
 *   effectiveUserEmail … Session.getEffectiveUser().getEmail() の値
 *   failPdf         … true で PDF 変換(getAs)が例外を投げる（発行失敗の再現）
 */
function buildSandbox(opts) {
  opts = opts || {};
  const ss = new Spreadsheet();
  const folders = {};
  const triggers = [];
  const calls = { mail: [], log: [] };
  const effectiveUserEmail =
    (opts.effectiveUserEmail === undefined) ? 'owner@example.test' : opts.effectiveUserEmail;

  const MimeType = { HTML: 'text/html', PDF: 'application/pdf', PNG: 'image/png' };

  function makeBlob(content, type, name) {
    return {
      getName() { return name; },
      getContentText() { return String(content); },
      getContentType() { return type; },
      setName(n) { name = n; return this; },
      getAs(t) {
        if (opts.failPdf && t === MimeType.PDF) {
          throw new Error('PDF変換に失敗しました（ハーネスの疑似エラー）');
        }
        return makeBlob(content, t, name);
      }
    };
  }

  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ss,
    getUi: () => {
      const menu = {
        addItem() { return menu; },
        addSeparator() { return menu; },
        addToUi() { return menu; }
      };
      return { createMenu: () => menu };
    }
  };

  const Utilities = {
    newBlob: (content, type, name) => makeBlob(content, type, name),
    base64Decode: (s) => new Array(String(s).length).fill(0),
    formatDate
  };

  const DriveApp = {
    getFoldersByName: (name) => {
      const hit = folders[name] ? [folders[name]] : [];
      let i = 0;
      return { hasNext: () => i < hit.length, next: () => hit[i++] };
    },
    createFolder: (name) => { folders[name] = new DriveFolder(name); return folders[name]; }
  };

  const MailApp = {
    sendEmail: (...args) => {
      // Code.gs は {to, subject, body} 形式で呼ぶ。位置引数形式も一応受ける
      const rec = (args.length === 1 && typeof args[0] === 'object')
        ? args[0]
        : { to: args[0], subject: args[1], body: args[2] };
      calls.mail.push(rec);
    }
  };

  const Session = {
    getEffectiveUser: () => ({ getEmail: () => effectiveUserEmail })
  };

  const ScriptApp = {
    getProjectTriggers: () => triggers.map(t => ({
      getHandlerFunction: () => t.handler,
      _raw: t
    })),
    deleteTrigger: (t) => {
      const raw = t && t._raw ? t._raw : t;
      const i = triggers.indexOf(raw);
      if (i >= 0) triggers.splice(i, 1);
    },
    newTrigger: (handler) => ({
      timeBased: () => ({
        everyDays: (days) => ({
          atHour: (hour) => ({
            create: () => { triggers.push({ handler, days, hour }); }
          })
        })
      })
    })
  };

  const Logger = { log: (...args) => calls.log.push(args) };

  const sandbox = {
    SpreadsheetApp, Utilities, DriveApp, MailApp, Session, ScriptApp, Logger, MimeType
  };
  if (opts.nowMs !== undefined) sandbox.Date = makeFixedDate(opts.nowMs);

  return { sandbox, ss, folders, triggers, calls, effectiveUserEmail };
}

module.exports = { buildSandbox, Sheet, Spreadsheet, Range };
