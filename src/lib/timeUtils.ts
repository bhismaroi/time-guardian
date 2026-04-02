// Time and name helpers for attendance processing

const MONTH_LOOKUP: Record<string, number> = {
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
 * Parse a time string (HH:MM or HH:MM:SS) to minutes from midnight.
 */
export function parseTimeToMinutes(time: string | null | undefined): number | null {
  if (!time) return null;

  const cleanTime = time.trim();
  if (!cleanTime || cleanTime === '__' || cleanTime.includes('__') || cleanTime === '-') {
    return null;
  }

  const match = cleanTime.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return null;

  const hours = Number(match[1]);
  const minutes = Number(match[2]);

  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) {
    return null;
  }

  return hours * 60 + minutes;
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
 * Convert an Excel time fraction into an HH:MM string.
 */
export function excelTimeToString(value: number | null | undefined): string | null {
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return null;
  }

  const totalMinutes = Math.round(value * 24 * 60);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
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
 * Convert HH:MM to an Excel time fraction.
 */
export function timeStringToExcelFraction(time: string | null | undefined): number | null {
  const minutes = parseTimeToMinutes(time);
  if (minutes === null) return null;
  return minutes / (24 * 60);
}

/**
 * Get the day of week (0 = Sunday, 1 = Monday, etc.).
 */
export function getDayOfWeek(date: Date): number {
  return date.getDay();
}

/**
 * Get day name abbreviation.
 */
export function getDayName(date: Date): string {
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return days[date.getDay()];
}

export function getDayNameLong(date: Date): string {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return days[date.getDay()];
}

export function isFriday(date: Date): boolean {
  return date.getDay() === 5;
}

export function isMondayOrThursday(date: Date): boolean {
  const day = date.getDay();
  return day === 1 || day === 4;
}

export function isWeekend(date: Date): boolean {
  const day = date.getDay();
  return day === 0 || day === 6;
}

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
 * Parse a date from multiple workbook formats.
 */
export function parseDate(dateStr: string, defaultYear = 2025): Date | null {
  if (!dateStr) return null;

  const trimmed = normalizeWhitespace(dateStr);

  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const [year, month, day] = trimmed.split('-').map(Number);
    return new Date(year, month - 1, day);
  }

  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed)) {
    const [day, month, year] = trimmed.split('/').map(Number);
    return new Date(year, month - 1, day);
  }

  if (/^\d{1,2}\/\d{1,2}$/.test(trimmed)) {
    const [day, month] = trimmed.split('/').map(Number);
    return new Date(defaultYear, month - 1, day);
  }

  const monthDayYear = trimmed.match(/^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})$/);
  if (monthDayYear) {
    const month = MONTH_LOOKUP[monthDayYear[1].toLowerCase()];
    if (month !== undefined) {
      return new Date(Number(monthDayYear[3]), month, Number(monthDayYear[2]));
    }
  }

  const dayMonthYear = trimmed.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$/);
  if (dayMonthYear) {
    const month = MONTH_LOOKUP[dayMonthYear[2].toLowerCase()];
    if (month !== undefined) {
      return new Date(Number(dayMonthYear[3]), month, Number(dayMonthYear[1]));
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
 * Extract time from a cell value that may contain extra text.
 */
export function extractTime(timeStr: string | null | undefined): string | null {
  if (!timeStr) return null;
  const trimmed = normalizeWhitespace(timeStr);
  if (!trimmed || trimmed === '__' || trimmed.includes('__') || trimmed === '-') return null;

  const timeMatch = trimmed.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!timeMatch) return null;

  return `${timeMatch[1].padStart(2, '0')}:${timeMatch[2]}`;
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
