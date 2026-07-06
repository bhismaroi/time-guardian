// Excel file parsing utilities

import * as XLSX from 'xlsx';
import type { RawFingerprintRecord } from './types';
import { MONTH_LOOKUP } from './policy';
import {
  extractNameParts,
  extractTime,
  formatDateIso,
  getEarlierTime,
  getLaterTime,
  normalizeName,
  normalizeWhitespace,
  parseDate,
} from './timeUtils';

type DailyClock = { clockIn: string | null; clockOut: string | null };

/**
 * The reporting period covered by the online workbook. For single-month
 * workbooks, start and end are equal. For multi-month reports that cross a
 * year boundary (e.g. "Dec 1, 2025 - Jan 31, 2026"), the two halves are used
 * to resolve which calendar year a given month label belongs to. Months are
 * 0-based (Jan = 0).
 */
type ReportContext = {
  startMonth: number;
  startYear: number;
  endMonth: number;
  endYear: number;
};

function parseReportContext(data: unknown[][]): ReportContext | null {
  for (let rowIndex = 0; rowIndex < Math.min(data.length, 6); rowIndex++) {
    const row = data[rowIndex];
    if (!row) continue;

    const text = row.map((cell) => String(cell ?? '')).join(' ');
    // Capture groups: 1=startMonth, 2=startYear, 3=endMonth, 4=endYear.
    // The original regex had only three groups and read "endYear" where the
    // end month should have been, so cross-year ranges silently fell through
    // to the single-match path and lost the second year.
    const rangeMatch = text.match(/([A-Za-z]{3,9})\s+\d{1,2},\s*(\d{4})\s*-\s*([A-Za-z]{3,9})\s+\d{1,2},\s*(\d{4})/);
    if (rangeMatch) {
      const startMonth = MONTH_LOOKUP[rangeMatch[1].toLowerCase()];
      const endMonth = MONTH_LOOKUP[rangeMatch[3].toLowerCase()];
      if (startMonth !== undefined && endMonth !== undefined) {
        return {
          startMonth,
          startYear: Number(rangeMatch[2]),
          endMonth,
          endYear: Number(rangeMatch[4]),
        };
      }
    }

    const singleMatch = text.match(/([A-Za-z]{3,9})\s+\d{1,2},\s*(\d{4})/);
    if (singleMatch) {
      const month = MONTH_LOOKUP[singleMatch[1].toLowerCase()];
      if (month !== undefined) {
        const year = Number(singleMatch[2]);
        return { startMonth: month, startYear: year, endMonth: month, endYear: year };
      }
    }
  }

  return null;
}

/**
 * Resolve the calendar year for a month label inside a reporting range.
 * For a single-month range (start === end) this returns the only year. For
 * a cross-year range like "Dec 2025 - Jan 2026", months >= startMonth map
 * to startYear and the rest map to endYear. Without this, a "Dec 1, 2025 -
 * Jan 31, 2026" report would tag January 1 as 2025-01-01, never matching
 * the fingerprint calendar dates.
 */
function resolveReportYear(month: number, context: ReportContext | null): number {
  if (!context) return new Date().getFullYear();
  if (context.startMonth === context.endMonth) return context.startYear;
  return month >= context.startMonth ? context.startYear : context.endYear;
}

function parseDateFromLabel(label: string, context: ReportContext | null): string | null {
  const match = label.match(/(\d{1,2})\s+([A-Za-z]{3,9})/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = MONTH_LOOKUP[match[2].toLowerCase()];
  if (month === undefined) return null;

  const year = resolveReportYear(month, context);
  return formatDateIso(new Date(year, month, day));
}

function toDateKey(value: unknown): { date: Date; dateKey: string } | null {
  if (value === null || value === undefined || value === '') return null;

  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    const date = new Date(parsed.y, parsed.m - 1, parsed.d);
    return { date, dateKey: formatDateIso(date) };
  }

  const stringValue = normalizeWhitespace(String(value));
  if (!stringValue) return null;

  const parsed = parseDate(stringValue, new Date().getFullYear());
  if (!parsed) return null;

  return { date: parsed, dateKey: formatDateIso(parsed) };
}

// Returns the index of the first header that includes any of the given
// priority labels. Used by the fingerprint parser to pick the
// policy-preferred column when several candidates may be present (e.g.
// "actual in" is preferred over "clock in"). Returns -1 if none match.
function pickPriorityIndex(headers: string[], priority: string[]): number {
  for (const label of priority) {
    const index = headers.findIndex((header) => header.includes(label));
    if (index !== -1) return index;
  }
  return -1;
}

function mergeClock(existing: DailyClock | undefined, next: DailyClock): DailyClock {
  if (!existing) return next;

  const mergedIn = getEarlierTime(existing.clockIn, next.clockIn);
  const mergedOut = getLaterTime(existing.clockOut, next.clockOut);

  return { clockIn: mergedIn, clockOut: mergedOut };
}

function getHeaderRow(data: unknown[][], labels: string[]): number {
  for (let rowIndex = 0; rowIndex < Math.min(data.length, 12); rowIndex++) {
    const row = data[rowIndex];
    if (!row) continue;
    const normalized = row.map((cell) => normalizeName(String(cell ?? '')));
    if (labels.every((label) => normalized.some((value) => value.includes(label)))) {
      return rowIndex;
    }
  }
  return -1;
}

function addAliases(
  store: Map<string, Map<string, DailyClock>>,
  aliasSource: string,
  records: Map<string, DailyClock>
): void {
  const normalized = normalizeName(aliasSource);
  if (!normalized) return;

  const parts = extractNameParts(normalized);
  // We deliberately do NOT register single-token aliases (first name or last
  // name alone). Two distinct employees can share a first or last name, and
  // adding single-token aliases was the root cause of `findOnlineMatch`
  // attributing one employee's online data to multiple fingerprint employees.
  // The full-name and reversed full-name aliases below cover the legitimate
  // "Adi Wijaya" vs "Wijaya Adi" name-order flip.
  const aliases = new Set<string>([
    normalized,
    parts.join(' '),
    parts.slice().reverse().join(' '),
  ]);

  for (const alias of aliases) {
    if (!alias) continue;
    store.set(alias, records);
  }
}

function parseOnlineMatrixFormat(
  data: unknown[][],
  reportContext: ReportContext | null
): Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>> {
  const result = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();

  const headerRowIndex = getHeaderRow(data, ['last name', 'first name']);
  const dateRowIndex = headerRowIndex >= 0 ? headerRowIndex + 1 : 4;
  const employeeStartRow = dateRowIndex + 1;
  const dateRow = data[dateRowIndex];
  if (!dateRow) return result;

  const columnToDateKey = new Map<number, string>();

  for (let col = 2; col < dateRow.length; col++) {
    const cell = String(dateRow[col] ?? '').trim();
    const match = cell.match(/(\d{1,2})\s+([A-Za-z]{3,9})/);
    if (!match) continue;

    const day = Number(match[1]);
    const month = MONTH_LOOKUP[match[2].toLowerCase()];
    if (month === undefined) continue;

    // Resolve the year via the report context so cross-year ranges (e.g.
    // "Dec 1, 2025 - Jan 31, 2026") tag each column with the correct year.
    const year = resolveReportYear(month, reportContext);
    columnToDateKey.set(col, formatDateIso(new Date(year, month, day)));
  }

  for (let rowIndex = employeeStartRow; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    if (!row) continue;

    const lastName = String(row[0] ?? '').trim();
    const firstName = String(row[1] ?? '').trim();
    if (!lastName && !firstName) continue;

    const displayName = firstName ? `${firstName} ${lastName}`.trim() : lastName;
    const reverseDisplayName = lastName ? `${lastName} ${firstName}`.trim() : firstName;

    const employeeRecords = new Map<string, { clockIn: string | null; clockOut: string | null }>();

    for (const [col, dateKey] of columnToDateKey) {
      if (col >= row.length) continue;

      const cellValue = String(row[col] ?? '').trim();
      if (!cellValue || /^do$/i.test(cellValue) || /^off$/i.test(cellValue)) {
        continue;
      }

      const match = cellValue.match(/(\d{1,2}:\d{2}|_+)\s*-\s*(\d{1,2}:\d{2}|_+)/);
      if (!match) continue;

      const clockIn = match[1].includes('_') ? null : extractTime(match[1]);
      const clockOut = match[2].includes('_') ? null : extractTime(match[2]);

      if (!clockIn && !clockOut) continue;

      const next = { clockIn, clockOut };
      const existing = employeeRecords.get(dateKey);
      employeeRecords.set(dateKey, mergeClock(existing, next));
    }

    if (employeeRecords.size === 0) continue;

    addAliases(result, displayName, employeeRecords);
    addAliases(result, reverseDisplayName, employeeRecords);
    addAliases(result, firstName, employeeRecords);
    addAliases(result, lastName, employeeRecords);
  }

  return result;
}

function parseOnlineBlockFormat(
  data: unknown[][],
  reportContext: ReportContext | null
): Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>> {
  const result = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();

  for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    if (!row) continue;

    const label = String(row[1] ?? '').trim();
    if (normalizeName(label) !== 'full name') continue;

    const employeeName = String(row[2] ?? '').trim();
    if (!employeeName) continue;

    let headerRowIndex = -1;
    for (let probe = rowIndex + 1; probe < Math.min(rowIndex + 10, data.length); probe++) {
      const probeRow = data[probe];
      if (!probeRow) continue;
      const probeLabel = normalizeName(String(probeRow[1] ?? ''));
      if (probeLabel === 'schedule') {
        headerRowIndex = probe;
        break;
      }
    }

    if (headerRowIndex === -1) continue;

    const employeeRecords = new Map<string, { clockIn: string | null; clockOut: string | null }>();
    const startRow = headerRowIndex + 1;

    for (let dayRow = startRow; dayRow < data.length; dayRow++) {
      const current = data[dayRow];
      if (!current) continue;

      const currentMarker = normalizeName(String(current[1] ?? ''));
      if (currentMarker === 'full name') {
        break;
      }

      const dateLabel = String(current[0] ?? '').trim();
      if (!dateLabel) {
        if (String(current[1] ?? '').trim() === '') {
          continue;
        }
        continue;
      }

      const dateKey = parseDateFromLabel(dateLabel, reportContext);
      if (!dateKey) continue;
      const clockIn = extractTime(String(current[3] ?? ''));
      const clockOut = extractTime(String(current[4] ?? ''));

      if (!clockIn && !clockOut) continue;

      const existing = employeeRecords.get(dateKey);
      employeeRecords.set(dateKey, mergeClock(existing, { clockIn, clockOut }));
    }

    if (employeeRecords.size === 0) continue;

    addAliases(result, employeeName, employeeRecords);
  }

  return result;
}

/**
 * Parse the Fingerprint Excel file.
 */
export function parseFingerprintExcel(file: ArrayBuffer): RawFingerprintRecord[] {
  const workbook = XLSX.read(file, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 }) as unknown[][];

  if (data.length === 0) return [];

  const headerRow = data[0].map((cell) => normalizeName(String(cell ?? '')));
  // For columns with a single candidate label (emp no, name, date, working
  // hours), the previous Math.max(idx, default) idiom was fine — either the
  // header is found or we fall back to a default position.
  //
  // For clockIn/clockOut the previous code did
  //   Math.max(actualInIdx, clockInIdx)
  // which is a position-based choice, not a priority-based one. If the
  // workbook has "Actual In" at column C and "Clock In" at column D, the
  // parser read the "Clock In" column — the opposite of the policy intent.
  // The priority is: 'actual in' (post-shift time) is preferred over
  // 'clock in' (raw punch time). pickPriorityIndex returns the first match
  // in priority order, or -1 if none are present.
  const columnIndex = {
    empNo: Math.max(headerRow.findIndex((header) => header.includes('emp no')), 0),
    name: Math.max(headerRow.findIndex((header) => header === 'name'), 3),
    date: Math.max(headerRow.findIndex((header) => header.includes('date')), 5),
    workingHours: Math.max(headerRow.findIndex((header) => header.includes('working hours')), 6),
    clockIn: pickPriorityIndex(headerRow, ['actual in', 'clock in time', 'clock in']),
    clockOut: pickPriorityIndex(headerRow, ['actual out', 'clock out time', 'clock out']),
  };

  // Surface a clear error if the workbook has no recognisable clock-in/out
  // columns. Pre-fix, both indices resolved to -1, every record carried
  // null clock times, and the workbook was still produced — the user got a
  // zero-everywhere report with no indication of the structural problem.
  if (columnIndex.clockIn < 0 && columnIndex.clockOut < 0) {
    throw new Error(
      'Fingerprint file is missing both "Actual In/Out" and "Clock In/Out" columns. '
      + `Detected headers: [${headerRow.join(', ')}]`
    );
  }

  const records: RawFingerprintRecord[] = [];

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    if (!row) continue;

    const empNo = String(row[columnIndex.empNo] ?? '').trim();
    const name = String(row[columnIndex.name] ?? '').trim();
    const dateValue = row[columnIndex.date];
    const workingHours = String(row[columnIndex.workingHours] ?? '').trim();

    if (!name || !dateValue) continue;

    const parsedDate = toDateKey(dateValue);
    if (!parsedDate) continue;

    // The header-row column detection (columnIndex.clockIn /
    // clockOut) is authoritative — it picked the policy-preferred
    // column once for the whole file. Per-row fallbacks to
    // pickTimeCell() (which scans the row's cells for an "actual in"
    // label) and the row[7] / row[8] magic indices are removed:
    // they added O(rows * candidates * labels) work for each
    // parsed row with no test coverage that justifies it. The
    // previous Phase 1.7 throws if neither clockIn nor clockOut
    // was found, so the case below is "exactly one of them exists"
    // or "neither" (in which case we just produce nulls).
    const rawClockIn = columnIndex.clockIn >= 0 ? String(row[columnIndex.clockIn] ?? '') : '';
    const rawClockOut = columnIndex.clockOut >= 0 ? String(row[columnIndex.clockOut] ?? '') : '';

    const actualIn = extractTime(rawClockIn);
    const actualOut = extractTime(rawClockOut);

    // Same as actualIn/actualOut: pickPriorityIndex already named
    // the preferred column.
    const clockIn = extractTime(rawClockIn);
    const clockOut = extractTime(rawClockOut);

    records.push({
      empNo,
      name,
      date: formatDateIso(parsedDate.date),
      dateKey: parsedDate.dateKey,
      workingHours,
      clockIn,
      clockOut,
      actualIn,
      actualOut,
    });
  }

  return records;
}

/**
 * Parse the Online Excel file.
 */
export function parseOnlineExcel(
  file: ArrayBuffer
): Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>> {
  const workbook = XLSX.read(file, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 }) as unknown[][];

  if (data.length === 0) return new Map();

  const reportContext = parseReportContext(data);

  // Try the block parser first. The block format requires a "Full name" cell
  // in column B *and* a "Schedule" header within 10 rows of it. A single
  // stray "Full name" cell in an otherwise matrix-format workbook satisfies
  // the first condition but never the second, so a block parse on such a
  // file returns an empty map. Falling through to the matrix parser in that
  // case recovers the data; the previous code returned zero records and no
  // error if a matrix file had a stray "Full name" cell.
  const blockResult = parseOnlineBlockFormat(data, reportContext);
  let totalBlockRecords = 0;
  for (const employeeRecords of blockResult.values()) {
    totalBlockRecords += employeeRecords.size;
  }
  if (totalBlockRecords > 0) return blockResult;
  return parseOnlineMatrixFormat(data, reportContext);
}

/**
 * Get unique employees from fingerprint records.
 */
export function getUniqueEmployees(records: RawFingerprintRecord[]): { empNo: string; name: string }[] {
  const seen = new Map<string, { empNo: string; name: string }>();

  for (const record of records) {
    const key = normalizeName(record.name);
    if (!seen.has(key)) {
      seen.set(key, {
        empNo: record.empNo,
        name: record.name,
      });
    }
  }

  return Array.from(seen.values());
}

/**
 * Get all dates from a month. Iterates by day number rather than
 * mutating a loop variable, which would be DST-fragile: in a DST
 * timezone, the spring-forward or fall-back transition can skip a
 * day (advance skips from 23:00 to 01:00) or repeat one (fall-back
 * repeats 01:00-02:00). Building each Date from its (year, month,
 * day) components avoids that pitfall.
 */
export function getMonthDates(year: number, month: number): Date[] {
  const dates: Date[] = [];
  const lastDay = new Date(year, month + 1, 0).getDate();

  for (let day = 1; day <= lastDay; day++) {
    dates.push(new Date(year, month, day));
  }

  return dates;
}
