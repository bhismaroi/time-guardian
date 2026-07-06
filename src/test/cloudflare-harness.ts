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
  // The real ExcelJS Cell type has a complex union (string | number |
  // Date | { formula, result } | { text, richText } | CellRichTextValue
  // | CellHyperlinkValue | CellErrorValue | null). For the test stub
  // we just need an `any`-shaped field that the test code can read
  // and write freely. Explicit `any` here is the right call — this
  // is a mock, not production code.
  value: any = null;
  numFmt: string | undefined = undefined;
  font: any = null;
  fill: any = null;
  border: any = null;
  alignment: any = null;
}

class StubRow {
  rowNumber: number;
  _cells: Record<number, StubCell>;

  constructor(rowNumber: number) {
    this.rowNumber = rowNumber;
    this._cells = {};
  }

  getCell(col: number): StubCell {
    if (!this._cells[col]) this._cells[col] = new StubCell();
    return this._cells[col];
  }
}

class StubWorksheet {
  name: string;
  _rows: Record<number, StubRow>;
  properties: Record<string, unknown>;
  views: unknown[];
  columns: { width: number }[];
  _mergedCells: string[];
  _conditionalFormatting: unknown[];

  constructor(name: string) {
    this.name = name;
    this._rows = {};
    this.properties = {};
    this.views = [];
    this.columns = [];
    this._mergedCells = [];
    this._conditionalFormatting = [];
  }

  getRow(rowNumber: number): StubRow {
    if (!this._rows[rowNumber]) this._rows[rowNumber] = new StubRow(rowNumber);
    return this._rows[rowNumber];
  }

  getCell(address: string): StubCell {
    // Parse the leading letters as a column and the trailing digits
    // as a row. A1 -> col A row 1. A1:M1 stays a range — not used
    // for cell reads.
    const match = address.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`StubWorksheet.getCell: bad address ${address}`);
    const colIndex = this._colLetterToIndex(match[1]);
    return this.getRow(Number(match[2])).getCell(colIndex);
  }

  eachRow(callback: (row: StubRow, rowNumber: number) => void): void {
    for (const rowNumber of Object.keys(this._rows).map(Number).sort((a, b) => a - b)) {
      callback(this.getRow(rowNumber), rowNumber);
    }
  }

  mergeCells(range: string): void {
    this._mergedCells.push(range);
  }

  addConditionalFormatting(spec: unknown): void {
    this._conditionalFormatting.push(spec);
  }

  _colLetterToIndex(letters: string): number {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }
}

class StubWorkbook {
  worksheets: StubWorksheet[];

  constructor() {
    this.worksheets = [];
  }

  // Construct from a pre-parsed XLSX representation. The Cloudflare
  // bundle calls `workbook.xlsx.load(buffer)`; we shortcut that by
  // parsing the buffer with the real SheetJS library (a project
  // dependency) and feeding the result into this stub.
  static fromXlsxBuffer(buffer: ArrayBuffer): StubWorkbook {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const wb = new StubWorkbook();
    workbook.SheetNames.forEach((name: string) => {
      const sheet = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], {
        header: 1,
        raw: true,
        defval: null,
      });
      const stubSheet = new StubWorksheet(name);
      sheet.forEach((row: unknown[] | null | undefined, rowIndex: number) => {
        const rowNumber = rowIndex + 1; // SheetJS is 0-based, ExcelJS is 1-based
        if (rowNumber === 1) {
          // The Cloudflare parser uses eachRow and skips rowNumber === 1
          // for fingerprint files. We still populate it so cell-by-cell
          // access works for the online parser, which uses getCell('A3')
          // and getCell('A6') (not row 1).
        }
        if (!row) return;
        row.forEach((value: unknown, colIndex: number) => {
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

  addWorksheet(name: string): StubWorksheet {
    const sheet = new StubWorksheet(name);
    this.worksheets.push(sheet);
    return sheet;
  }

  getWorksheet(name: string): StubWorksheet | null {
    return this.worksheets.find((s: StubWorksheet) => s.name === name) || null;
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

// Module-level cache of the loaded AttendanceCompiler. The shape is
// the one browserCompiler.js exposes (a buildCompiledWorkbookFromFiles
// entry plus warnings/summary fields); see the call site at the
// bottom of loadCloudflareBundle for the runtime assignment.
type AttendanceCompiler = {
  buildCompiledWorkbookFromFiles: (fingerprint: File, online: File, opts?: { onProgress?: (msg: string) => void }) => Promise<{
    workbook: StubWorkbook;
    fileName: string;
    warnings: string[];
    summary: {
      employees: number;
      matchedEmployees: number;
      fingerprintOnlyEmployees: number;
      onlineOnlyEmployees: number;
      lowConfidenceMatches: number;
      month: string;
    };
  }>;
};
let loaded: AttendanceCompiler | null = null;

function loadCloudflareBundle(): AttendanceCompiler {
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
  const bufferQueue: ArrayBuffer[] = [];
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

  const runPolicy = new Function(policySource);
  runPolicy.call(window);
  const runCompiler = new Function(compilerSource);
  runCompiler.call(window);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  loaded = (window as unknown as { AttendanceCompiler: AttendanceCompiler }).AttendanceCompiler;
  if (!loaded || typeof loaded.buildCompiledWorkbookFromFiles !== 'function') {
    throw new Error('Cloudflare bundle did not expose buildCompiledWorkbookFromFiles');
  }
  return loaded;
}

function makeFileLike(buffer: ArrayBuffer, name: string): File {
  // The Cloudflare code calls `await file.arrayBuffer()`. jsdom's
  // built-in File should support this, but in some jsdom versions
  // the third-arg options object is not accepted by the File
  // constructor, so we wrap a plain object that exposes the same
  // surface (name, arrayBuffer) and let the Cloudflare parser
  // consume it.
  const file = new File([buffer], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  if (typeof file.arrayBuffer !== 'function') {
    (file as unknown as { arrayBuffer: () => Promise<ArrayBuffer> }).arrayBuffer = async () => buffer;
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
export async function compileWithCloudflare(
  fingerprintBuffer: ArrayBuffer,
  onlineBuffer: ArrayBuffer,
  options: { debug?: boolean } = {}
) {
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
    console.log('=== Cloudflare compile result ===');
    console.log('warnings:', JSON.stringify(result.warnings, null, 2));
    console.log('summary:', JSON.stringify(result.summary, null, 2));
    console.log('sheets:', result.workbook.worksheets.map((s: { name: string }) => s.name));
  }
  return result;
}
