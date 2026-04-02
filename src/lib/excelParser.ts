// Excel file parsing utilities

import * as XLSX from 'xlsx';
import type { RawFingerprintRecord } from './types';
import {
  extractNameParts,
  extractTime,
  formatDateIso,
  normalizeName,
  normalizeWhitespace,
  parseDate,
  parseTimeToMinutes,
} from './timeUtils';

type DailyClock = { clockIn: string | null; clockOut: string | null };

function getCellValue(row: unknown[], index: number): string {
  return normalizeWhitespace(String(row[index] ?? ''));
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

function pickTimeCell(row: unknown[], headers: string[], candidates: string[]): string | null {
  for (const candidate of candidates) {
    const index = headers.findIndex((header) => header.includes(candidate));
    if (index !== -1) {
      const raw = String(row[index] ?? '');
      const time = extractTime(raw);
      if (time) return time;
    }
  }
  return null;
}

function mergeClock(existing: DailyClock | undefined, next: DailyClock): DailyClock {
  if (!existing) return next;

  const mergedIn = getEarlierTime(existing.clockIn, next.clockIn);
  const mergedOut = getLaterTime(existing.clockOut, next.clockOut);

  return { clockIn: mergedIn, clockOut: mergedOut };
}

function getEarlierTime(time1: string | null, time2: string | null): string | null {
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);

  if (min1 === null && min2 === null) return null;
  if (min1 === null) return time2;
  if (min2 === null) return time1;
  return min1 <= min2 ? time1 : time2;
}

function getLaterTime(time1: string | null, time2: string | null): string | null {
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);

  if (min1 === null && min2 === null) return null;
  if (min1 === null) return time2;
  if (min2 === null) return time1;
  return min1 >= min2 ? time1 : time2;
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
  const aliases = new Set<string>([
    normalized,
    parts.join(' '),
    parts.slice().reverse().join(' '),
    ...parts,
  ]);

  for (const alias of aliases) {
    if (!alias) continue;
    store.set(alias, records);
  }
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
  const columnIndex = {
    empNo: Math.max(headerRow.findIndex((header) => header.includes('emp no')), 0),
    name: Math.max(headerRow.findIndex((header) => header === 'name'), 3),
    date: Math.max(headerRow.findIndex((header) => header.includes('date')), 5),
    workingHours: Math.max(headerRow.findIndex((header) => header.includes('working hours')), 6),
    clockIn: Math.max(
      headerRow.findIndex((header) => header.includes('actual in')),
      headerRow.findIndex((header) => header.includes('clock in'))
    ),
    clockOut: Math.max(
      headerRow.findIndex((header) => header.includes('actual out')),
      headerRow.findIndex((header) => header.includes('clock out'))
    ),
  };

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

    const rawClockIn = columnIndex.clockIn >= 0 ? String(row[columnIndex.clockIn] ?? '') : '';
    const rawClockOut = columnIndex.clockOut >= 0 ? String(row[columnIndex.clockOut] ?? '') : '';
    const fallbackClockIn = pickTimeCell(row, headerRow, ['actual in', 'clock in time', 'clock in']);
    const fallbackClockOut = pickTimeCell(row, headerRow, ['actual out', 'clock out time', 'clock out']);

    const actualIn = extractTime(rawClockIn) ?? fallbackClockIn;
    const actualOut = extractTime(rawClockOut) ?? fallbackClockOut;

    const clockIn = extractTime(rawClockIn) ?? extractTime(String(row[7] ?? '')) ?? actualIn;
    const clockOut = extractTime(rawClockOut) ?? extractTime(String(row[8] ?? '')) ?? actualOut;

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

  const result = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();
  if (data.length === 0) return result;

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
    const month = match[2].toLowerCase();
    const monthIndex: Record<string, number> = {
      jan: 0,
      january: 0,
      feb: 1,
      february: 1,
      mar: 2,
      march: 2,
      apr: 3,
      april: 3,
      may: 4,
      jun: 5,
      june: 5,
      jul: 6,
      july: 6,
      aug: 7,
      august: 7,
      sep: 8,
      sept: 8,
      september: 8,
      oct: 9,
      october: 9,
      nov: 10,
      november: 10,
      dec: 11,
      december: 11,
    };

    const monthIndexValue = monthIndex[month];
    if (monthIndexValue === undefined) continue;

    const date = new Date(2025, monthIndexValue, day);
    columnToDateKey.set(col, formatDateIso(date));
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
 * Get all dates from a month.
 */
export function getMonthDates(year: number, month: number): Date[] {
  const dates: Date[] = [];
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);

  for (let date = new Date(firstDay); date <= lastDay; date.setDate(date.getDate() + 1)) {
    dates.push(new Date(date));
  }

  return dates;
}
