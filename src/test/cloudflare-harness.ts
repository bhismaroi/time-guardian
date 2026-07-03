// Cloudflare test harness — load the static Cloudflare bundle in
// jsdom with a stubbed ExcelJS so we can run the same inputs through
// both the React path (src/lib) and the Cloudflare path
// (cloudflare-pages-attendance/public/browserCompiler.js) and
// compare the resulting cells.
//
// ExcelJS is normally loaded via <script src="...cdn...">. We stub
// it in-memory: a Workbook class that captures every sheet, every
// cell value/numFmt/font/fill/border, every merged range, and
// every conditional-formatting rule. The Cloudflare code calls the
// same shape of API as the real ExcelJS, so the captured cells are
// a faithful record of what the Cloudflare deployment would write.

import * as XLSX from 'xlsx';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

class StubCell {
  constructor() {
    this.value = null;
    this.numFmt = null;
    this.font = null;
    this.fill = null;
    this.border = null;
    this.alignment = null;
  }
}

class StubRow {
  constructor(rowNumber) {
    this.rowNumber = rowNumber;
    this._cells = {};
  }

  getCell(col) {
    if (!this._cells[col]) this._cells[col] = new StubCell();
    return this._cells[col];
  }
}

class StubWorksheet {
  constructor(name) {
    this.name = name;
    this._rows = {};
    this.properties = {};
    this.views = [];
    this.columns = [];
    this._mergedCells = [];
    this._conditionalFormatting = [];
  }

  getRow(rowNumber) {
    if (!this._rows[rowNumber]) this._rows[rowNumber] = new StubRow(rowNumber);
    return this._rows[rowNumber];
  }

  getCell(address) {
    // Parse the leading letters as a column and the trailing digits
    // as a row. A1 -> col A row 1. A1:M1 stays a range — not used
    // for cell reads.
    const match = address.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`StubWorksheet.getCell: bad address ${address}`);
    const colIndex = this._colLetterToIndex(match[1]);
    return this.getRow(Number(match[2])).getCell(colIndex);
  }

  eachRow(callback) {
    for (const rowNumber of Object.keys(this._rows).map(Number).sort((a, b) => a - b)) {
      callback(this.getRow(rowNumber), rowNumber);
    }
  }

  mergeCells(range) {
    this._mergedCells.push(range);
  }

  addConditionalFormatting(spec) {
    this._conditionalFormatting.push(spec);
  }

  _colLetterToIndex(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }
}

class StubWorkbook {
  constructor() {
    this.worksheets = [];
  }

  // Construct from a pre-parsed XLSX representation. The Cloudflare
  // bundle calls `workbook.xlsx.load(buffer)`; we shortcut that by
  // parsing the buffer with the real SheetJS library (a project
  // dependency) and feeding the result into this stub.
  static fromXlsxBuffer(buffer) {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const wb = new StubWorkbook();
    workbook.SheetNames.forEach((name) => {
      const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
        header: 1,
        raw: true,
        defval: null,
      });
      const stubSheet = new StubWorksheet(name);
      sheet.forEach((row, rowIndex) => {
        const rowNumber = rowIndex + 1; // SheetJS is 0-based, ExcelJS is 1-based
        if (rowNumber === 1) {
          // The Cloudflare parser uses eachRow and skips rowNumber === 1
          // for fingerprint files. We still populate it so cell-by-cell
          // access works for the online parser, which uses getCell('A3')
          // and getCell('A6') (not row 1).
        }
        if (!row) return;
        row.forEach((value, colIndex) => {
          const col = colIndex + 1;
          if (value === null || value === undefined) return;
          // The Cloudflare code reads cell.value and handles Date,
          // number, string, and { text, richText, result }. SheetJS
          // returns the value as-is (Date object for date cells,
          // number for numeric, string for text).
          stubSheet.getRow(rowNumber).getCell(col).value = value;
        });
      });
      wb.worksheets.push(stubSheet);
    });
    return wb;
  }

  addWorksheet(name) {
    const sheet = new StubWorksheet(name);
    this.worksheets.push(sheet);
    return sheet;
  }

  getWorksheet(name) {
    return this.worksheets.find((s) => s.name === name) || null;
  }

  // No-op for the stub. The TrackedWorkbook constructor in
  // loadCloudflareBundle() pre-loads the buffer via
  // StubWorkbook.fromXlsxBuffer before the Cloudflare parser calls
  // `await workbook.xlsx.load(buffer)`. The Cloudflare code awaits
  // this method, so we provide a no-op to satisfy the API surface.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  xlsx: any = {
    load: async (_buffer: ArrayBuffer) => {
      // already loaded by TrackedWorkbook constructor
    },
  };
}

let loaded = null;

function loadCloudflareBundle() {
  if (loaded) return loaded;
  if (typeof window === 'undefined') {
    throw new Error('loadCloudflareBundle must run in a browser/jsdom context');
  }

  // Install the ExcelJS stub on the global. The Cloudflare bundle
  // calls `new ExcelJS.Workbook()` twice in sequence (once per
  // source file). Each call must return a parsed workbook instance,
  // so we maintain a queue of buffers that the TrackedWorkbook
  // constructor shifts from. The Cloudflare code may also call
  // `await workbook.xlsx.load(buffer)`; since the buffer was
  // already parsed at construction time, xlsx.load is a no-op.
  const bufferQueue = [];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__enqueueWorkbookBuffer = (buf: ArrayBuffer) => {
    bufferQueue.push(buf);
  };
  const TrackedWorkbook = function TrackedWorkbook() {
    const buf = bufferQueue.shift();
    return buf ? StubWorkbook.fromXlsxBuffer(buf) : new StubWorkbook();
  };
  TrackedWorkbook.prototype = StubWorkbook.prototype;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).ExcelJS = { Workbook: TrackedWorkbook };

  // Read the policy.js and browserCompiler.js bundles from disk
  // and evaluate them in the current window context. The IIFE in
  // browserCompiler.js will set window.AttendanceCompiler.
  const repoRoot = resolve(__dirname, '../..');
  const policySource = readFileSync(resolve(repoRoot, 'cloudflare-pages-attendance/public/policy.js'), 'utf8');
  const compilerSource = readFileSync(resolve(repoRoot, 'cloudflare-pages-attendance/public/browserCompiler.js'), 'utf8');

  // eslint-disable-next-line no-new-func
  const runPolicy = new Function(policySource);
  runPolicy.call(window);
  // eslint-disable-next-line no-new-func
  const runCompiler = new Function(compilerSource);
  runCompiler.call(window);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  loaded = (window as any).AttendanceCompiler;
  if (!loaded || typeof loaded.buildCompiledWorkbookFromFiles !== 'function') {
    throw new Error('Cloudflare bundle did not expose buildCompiledWorkbookFromFiles');
  }
  return loaded;
}

function makeFileLike(buffer, name) {
  // The Cloudflare code calls `await file.arrayBuffer()`. jsdom's
  // built-in File should support this, but in some jsdom versions
  // the third-arg options object is not accepted by the File
  // constructor, so we wrap a plain object that exposes the same
  // surface (name, arrayBuffer) and let the Cloudflare parser
  // consume it.
  const file = new File([buffer], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  if (typeof file.arrayBuffer !== 'function') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (file as any).arrayBuffer = async () => buffer;
  }
  return file;
}

/**
 * Run the Cloudflare bundle on the given fingerprint and online
 * xlsx buffers and return the captured workbook plus the merge
 * summary. The returned workbook has `.worksheets` (array of
 * StubWorksheet) where each worksheet's cells can be inspected
 * directly.
 */
export async function compileWithCloudflare(fingerprintBuffer, onlineBuffer, options: { debug?: boolean } = {}) {
  const AC = loadCloudflareBundle();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__enqueueWorkbookBuffer(fingerprintBuffer);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__enqueueWorkbookBuffer(onlineBuffer);
  const fingerprintFile = makeFileLike(fingerprintBuffer, 'fingerprint.xlsx');
  const onlineFile = makeFileLike(onlineBuffer, 'online.xlsx');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const result = await (AC as any).buildCompiledWorkbookFromFiles(fingerprintFile, onlineFile);
  if (options.debug) {
    // eslint-disable-next-line no-console
    console.log('=== Cloudflare compile result ===');
    // eslint-disable-next-line no-console
    console.log('warnings:', JSON.stringify(result.warnings, null, 2));
    // eslint-disable-next-line no-console
    console.log('summary:', JSON.stringify(result.summary, null, 2));
    // eslint-disable-next-line no-console
    console.log('sheets:', result.workbook.worksheets.map((s: { name: string }) => s.name));
  }
  return result;
}
