// Time and name helpers for attendance processing

import { MONTH_LOOKUP } from './policy';

export function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

export function normalizeName(value: string | null | undefined): string {
  if (!value) return '';
  return normalizeWhitespace(value)
    .toLowerCase()
    .replace(/[.,/\\'`"-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function extractNameParts(value: string | null | undefined): string[] {
  const normalized = normalizeName(value);
  if (!normalized) return [];
  return normalized
    .split(' ')
    .map((part) => part.trim())
    .filter(Boolean);
}

/**
 * Parse a time string (HH:MM or HH:MM:SS) into its component parts.
 * Returns { hours, minutes, seconds } or null if the string is empty,
 * is a placeholder (e.g. '__', '-'), or has invalid values.
 * Both parseTimeToMinutes and extractTime are thin wrappers over this.
 */
function parseTimeParts(time: string | null | undefined): { hours: number; minutes: number; seconds: number } | null {
  if (!time) return null;

  const cleanTime = time.trim();
  if (!cleanTime || cleanTime === '__' || cleanTime.includes('__') || cleanTime === '-') {
    return null;
  }

  const match = cleanTime.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return null;

  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = match[3] ? Number(match[3]) : 0;

  if (!Number.isFinite(hours) || !Number.isFinite(minutes) || !Number.isFinite(seconds)) {
    return null;
  }
  if (hours > 23 || minutes > 59 || seconds > 59) {
    return null;
  }

  return { hours, minutes, seconds };
}

/**
 * Parse a time string (HH:MM or HH:MM:SS) to minutes from midnight.
 */
export function parseTimeToMinutes(time: string | null | undefined): number | null {
  const parts = parseTimeParts(time);
  if (parts === null) return null;
  return parts.hours * 60 + parts.minutes;
}

/**
 * Convert minutes from midnight to HH:MM format.
 */
export function minutesToTimeString(minutes: number): string {
  if (!Number.isFinite(minutes) || minutes <= 0) return '0:00';
  const hours = Math.floor(minutes / 60);
  const mins = Math.round(minutes % 60);
  return `${hours}:${mins.toString().padStart(2, '0')}`;
}

/**
 * Convert a Date to an Excel serial date number.
 */
export function dateToExcelSerial(date: Date): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  const excelEpoch = Date.UTC(1899, 11, 30);
  return (utc - excelEpoch) / (24 * 60 * 60 * 1000);
}

/**
 * Get day name abbreviation.
 */
export function getDayName(date: Date): string {
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return days[date.getDay()];
}

// The following helpers were removed in Phase 5.5 because they
// became dead once shared/policy.js became the canonical source:
//   - excelTimeToString        (was an unused Excel-formatter)
//   - timeStringToExcelFraction (was an unused Excel-formatter)
//   - getDayOfWeek             (was an unused JS Date wrapper)
//   - getDayNameLong           (was an unused long-form variant)
//   - isFriday                 (now in shared/policy.js as isFriday)
//   - isWeekend                (now in shared/policy.js as isWeekend)

/**
 * Calculate break overlap in minutes.
 */
export function calculateBreakOverlap(
  clockInMinutes: number,
  clockOutMinutes: number,
  breakStartMinutes: number,
  breakEndMinutes: number
): number {
  if (clockOutMinutes <= breakStartMinutes || clockInMinutes >= breakEndMinutes) {
    return 0;
  }

  const overlapStart = Math.max(clockInMinutes, breakStartMinutes);
  const overlapEnd = Math.min(clockOutMinutes, breakEndMinutes);

  return Math.max(0, overlapEnd - overlapStart);
}

/**
 * Whether a numeric date such as "8/1/2026" should be read as
 * day-first (1 August) or month-first (8 January). A single export uses
 * one convention throughout, so this is decided once per date column.
 */
export type NumericDateOrder = 'day-first' | 'month-first';

// Two-digit years are expanded into 2000-2069 (the fingerprint export
// writes "26" for 2026), falling back to 1900-1999 at or above the
// cutoff.
function expandTwoDigitYear(value: number): number {
  if (value >= 100) return value;
  return value < 70 ? 2000 + value : 1900 + value;
}

/**
 * Build a Date only from a real calendar day. Out-of-range values and
 * rollovers such as 31 February (which JS would silently turn into
 * 3 March) return null instead of a wrong date.
 */
function makeDateParts(year: number, month: number, day: number): Date | null {
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }
  return date;
}

/**
 * Infer whether a column of numeric dates ("8/1/2026") is day-first or
 * month-first.
 *
 * A component greater than 12 can only be a day, which pins the order
 * outright. When every value is ambiguous (both components <= 12) the
 * month of a single-month export stays constant across rows while the
 * day varies, so the interpretation whose SECOND component is constant
 * is the month-first one. "8/1, 8/2, ... 8/12" is therefore month-first
 * and "1/8, 2/8, ... 12/8" is day-first.
 */
export function detectNumericDateOrder(
  values: Iterable<string | null | undefined>
): NumericDateOrder {
  const firstComponents = new Set<number>();
  const secondComponents = new Set<number>();
  let mustBeDayFirst = false;
  let mustBeMonthFirst = false;

  for (const value of values) {
    if (!value) continue;

    const match = normalizeWhitespace(String(value)).match(
      /^(\d{1,2})[/\-.](\d{1,2})[/\-.]\d{2,4}$/
    );
    if (!match) continue;

    const first = Number(match[1]);
    const second = Number(match[2]);
    firstComponents.add(first);
    secondComponents.add(second);

    if (first > 12) mustBeDayFirst = true;
    if (second > 12) mustBeMonthFirst = true;
  }

  if (mustBeDayFirst) return 'day-first';
  if (mustBeMonthFirst) return 'month-first';

  return firstComponents.size < secondComponents.size ? 'month-first' : 'day-first';
}

/**
 * Parse a date from multiple workbook formats.
 *
 * Supported shapes:
 *   - Fingerprint export: "01-Sep-26", "1 Sep 2026", "01/Sep/2026"
 *   - ISO:                "2026-09-01"
 *   - Numeric:            "8/1/2026", "1-8-2026", "1/8"
 *   - Month name:         "Sep 1, 2026", "September 1 2026"
 *
 * `numericOrder` only affects the ambiguous numeric forms, where the
 * same string can mean either 1 August or 8 January; callers pass the
 * value that detectNumericDateOrder returned for the whole column.
 */
export function parseDate(
  dateStr: string,
  defaultYear = 2025,
  numericOrder: NumericDateOrder = 'day-first'
): Date | null {
  if (!dateStr) return null;

  const trimmed = normalizeWhitespace(dateStr);
  if (!trimmed) return null;

  // ISO: 2026-09-01.
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const [year, month, day] = trimmed.split('-').map(Number);
    return makeDateParts(year, month, day);
  }

  // Numeric with a year: 8/1/2026, 1-8-26.
  const numeric = trimmed.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$/);
  if (numeric) {
    const first = Number(numeric[1]);
    const second = Number(numeric[2]);
    const year = expandTwoDigitYear(Number(numeric[3]));
    const month = numericOrder === 'month-first' ? first : second;
    const day = numericOrder === 'month-first' ? second : first;
    return makeDateParts(year, month, day);
  }

  // Numeric without a year: 1/8.
  const numericNoYear = trimmed.match(/^(\d{1,2})[/\-.](\d{1,2})$/);
  if (numericNoYear) {
    const first = Number(numericNoYear[1]);
    const second = Number(numericNoYear[2]);
    const month = numericOrder === 'month-first' ? first : second;
    const day = numericOrder === 'month-first' ? second : first;
    return makeDateParts(defaultYear, month, day);
  }

  // Day, month name, year: "01-Sep-26", "1 Sep 2026", "01/Sep/2026".
  const dayMonthNameYear = trimmed.match(
    /^(\d{1,2})[\s\-/.,]+([A-Za-z]{3,9})[\s\-/.,]+(\d{2,4})$/
  );
  if (dayMonthNameYear) {
    const month = MONTH_LOOKUP[dayMonthNameYear[2].toLowerCase()];
    if (month !== undefined) {
      return makeDateParts(
        expandTwoDigitYear(Number(dayMonthNameYear[3])),
        month + 1,
        Number(dayMonthNameYear[1])
      );
    }
  }

  // Month name, day, year: "Sep 1, 2026".
  const monthNameDayYear = trimmed.match(
    /^([A-Za-z]{3,9})[\s\-/.,]+(\d{1,2})[\s\-/.,]+(\d{2,4})$/
  );
  if (monthNameDayYear) {
    const month = MONTH_LOOKUP[monthNameDayYear[1].toLowerCase()];
    if (month !== undefined) {
      return makeDateParts(
        expandTwoDigitYear(Number(monthNameDayYear[3])),
        month + 1,
        Number(monthNameDayYear[2])
      );
    }
  }

  return null;
}

export function formatDateShort(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  return `${day}/${month}`;
}

export function formatDateFull(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

export function formatDateIso(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${year}-${month}-${day}`;
}

/**
 * Extract time from a cell value that may contain extra text. Returns
 * the HH:MM portion (with zero-padded hour) or null if no valid time
 * is found. Delegates to parseTimeParts so the regex and validation
 * are defined in exactly one place.
 */
export function extractTime(timeStr: string | null | undefined): string | null {
  const normalized = timeStr ? normalizeWhitespace(timeStr) : '';
  const parts = parseTimeParts(normalized);
  if (parts === null) return null;
  return `${String(parts.hours).padStart(2, '0')}:${String(parts.minutes).padStart(2, '0')}`;
}

/**
 * Compare two times and return the earlier one.
 */
export function getEarlierTime(time1: string | null, time2: string | null): string | null {
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);

  if (min1 === null && min2 === null) return null;
  if (min1 === null) return time2;
  if (min2 === null) return time1;

  return min1 <= min2 ? time1 : time2;
}

/**
 * Compare two times and return the later one.
 */
export function getLaterTime(time1: string | null, time2: string | null): string | null {
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);

  if (min1 === null && min2 === null) return null;
  if (min1 === null) return time2;
  if (min2 === null) return time1;

  return min1 >= min2 ? time1 : time2;
}
