// Excel file parsing utilities

import * as XLSX from 'xlsx';
import type { RawFingerprintRecord } from './types';
import { extractTime } from './timeUtils';

/**
 * Parse the Fingerprint Excel file
 */
export function parseFingerprintExcel(file: ArrayBuffer): RawFingerprintRecord[] {
  const workbook = XLSX.read(file, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 }) as unknown[][];

  const records: RawFingerprintRecord[] = [];

  // Skip header row (first row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !Array.isArray(row) || row.length < 10) continue;

    const empNo = String(row[0] || '');
    const name = String(row[3] || '').trim();
    const dateValue = row[5];
    const workingHours = String(row[6] || '');
    const actualIn = String(row[9] || '');
    const actualOut = String(row[10] || '');

    if (!name || !dateValue) continue;

    // Parse date - could be a number (Excel date) or string
    let dateStr = '';
    if (typeof dateValue === 'number') {
      const date = XLSX.SSF.parse_date_code(dateValue);
      dateStr = `${date.d.toString().padStart(2, '0')}/${(date.m).toString().padStart(2, '0')}/${date.y}`;
    } else {
      dateStr = String(dateValue);
    }

    records.push({
      empNo,
      name,
      date: dateStr,
      workingHours,
      actualIn: extractTime(actualIn) || '',
      actualOut: extractTime(actualOut) || '',
    });
  }

  return records;
}

/**
 * Parse the Online Excel file
 * Structure:
 * - Row 9 (index 8): Section headers including "Clock-in / Clock-out"
 * - Row 10 (index 9): Date headers like "01 Oct, We", "02 Oct, Th"
 * - Row 11+ (index 10+): Employee data with clock times in format "HH:MM - HH:MM" or "HH:MM - __" or "__ - HH:MM"
 */
export function parseOnlineExcel(file: ArrayBuffer): Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>> {
  const workbook = XLSX.read(file, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 }) as unknown[][];

  // Result: Map<employeeName, Map<dateStr, {clockIn, clockOut}>>
  const result = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();

  // Find the header row with section names (row 9, index 8)
  const sectionHeaderRow = data[8];
  // Find the date row (row 10, index 9)
  const dateRow = data[9];
  
  if (!dateRow || !Array.isArray(dateRow)) {
    console.log('No date row found at index 9');
    return result;
  }
  if (!sectionHeaderRow || !Array.isArray(sectionHeaderRow)) {
    console.log('No section header row found at index 8');
    return result;
  }

  // Find the "Clock-in / Clock-out" section by scanning the section header row
  let clockSectionStart = -1;
  let clockSectionEnd = -1;
  
  for (let col = 0; col < sectionHeaderRow.length; col++) {
    const header = String(sectionHeaderRow[col] || '').toLowerCase();
    if (header.includes('clock-in') && header.includes('clock-out')) {
      clockSectionStart = col;
      console.log(`Found "Clock-in / Clock-out" section at column ${col}`);
      break;
    }
  }

  // If we found the section header, find where it ends (next non-empty header that's not part of dates)
  if (clockSectionStart !== -1) {
    // Count consecutive date columns starting from clockSectionStart
    for (let col = clockSectionStart; col < Math.min(clockSectionStart + 35, dateRow.length); col++) {
      const dateStr = String(dateRow[col] || '');
      // Check if this column has a date pattern like "01 Oct, We"
      if (dateStr.match(/\d{1,2}\s+\w+,\s*\w+/)) {
        clockSectionEnd = col + 1;
      }
    }
    console.log(`Clock section spans columns ${clockSectionStart} to ${clockSectionEnd}`);
  }

  // If we couldn't find the section by header, try to find it by data pattern
  if (clockSectionStart === -1) {
    console.log('Section header not found, scanning for clock patterns in data...');
    // Look at first employee row for clock patterns
    const firstEmployeeRow = data[10];
    if (firstEmployeeRow && Array.isArray(firstEmployeeRow)) {
      for (let col = 0; col < firstEmployeeRow.length; col++) {
        const cellValue = String(firstEmployeeRow[col] || '');
        // Look for patterns like "HH:MM - HH:MM" or "__ - __" or "HH:MM - __"
        if (cellValue.match(/(\d{1,2}:\d{2}|_+)\s*-\s*(\d{1,2}:\d{2}|_+)/)) {
          // Check if the date row at this column has a valid date
          const dateStr = String(dateRow[col] || '');
          if (dateStr.match(/\d{1,2}\s+\w+/)) {
            clockSectionStart = col;
            console.log(`Found clock section by pattern at column ${col}: "${cellValue}"`);
            break;
          }
        }
      }
    }
    
    if (clockSectionStart !== -1) {
      clockSectionEnd = clockSectionStart + 31; // Assume 31 days max
    }
  }

  if (clockSectionStart === -1) {
    console.log('Could not find Clock-in / Clock-out section');
    return result;
  }

  // Build a map of column index to normalized date string
  const columnToDate = new Map<number, string>();
  for (let col = clockSectionStart; col <= clockSectionEnd && col < dateRow.length; col++) {
    const dateStr = String(dateRow[col] || '');
    const dateMatch = dateStr.match(/(\d{1,2})\s+(\w+)/);
    if (dateMatch) {
      const day = parseInt(dateMatch[1], 10);
      const monthStr = dateMatch[2];
      const monthMap: Record<string, number> = {
        'Oct': 10, 'Nov': 11, 'Dec': 12, 'Jan': 1, 'Feb': 2, 'Mar': 3,
        'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9
      };
      const month = monthMap[monthStr] || 10;
      const normalizedDate = `${day.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/2025`;
      columnToDate.set(col, normalizedDate);
    }
  }
  console.log(`Mapped ${columnToDate.size} columns to dates`);

  // Parse employee rows (starting from row 11, index 10)
  for (let rowIdx = 10; rowIdx < data.length; rowIdx++) {
    const row = data[rowIdx];
    if (!row || !Array.isArray(row) || row.length < 3) continue;

    const lastName = String(row[0] || '').trim();
    const firstName = String(row[1] || '').trim();
    
    if (!lastName && !firstName) continue;
    
    const fullName = firstName ? `${firstName} ${lastName}`.trim() : lastName;
    const reverseName = firstName ? `${lastName} ${firstName}`.trim() : lastName;
    
    const employeeRecords = new Map<string, { clockIn: string | null; clockOut: string | null }>();
    
    // Parse each column in the clock section
    for (const [col, normalizedDate] of columnToDate) {
      if (col >= row.length) continue;
      
      const cellValue = String(row[col] || '');
      
      // Skip if it's not a clock pattern (could be "DO" for day off, or empty, or "8h 0m")
      // Valid patterns: "HH:MM - HH:MM", "HH:MM - __", "__ - HH:MM", "__ - __"
      const clockPattern = cellValue.match(/(\d{1,2}:\d{2}|_+)\s*-\s*(\d{1,2}:\d{2}|_+)/);
      
      if (clockPattern) {
        const inPart = clockPattern[1];
        const outPart = clockPattern[2];
        
        const clockIn = inPart.includes('_') ? null : extractTime(inPart);
        const clockOut = outPart.includes('_') ? null : extractTime(outPart);
        
        if (clockIn || clockOut) {
          // Merge with existing record for this date (get earliest in, latest out)
          const existing = employeeRecords.get(normalizedDate);
          if (existing) {
            const mergedIn = getEarlierTimeStr(existing.clockIn, clockIn);
            const mergedOut = getLaterTimeStr(existing.clockOut, clockOut);
            employeeRecords.set(normalizedDate, { clockIn: mergedIn, clockOut: mergedOut });
          } else {
            employeeRecords.set(normalizedDate, { clockIn, clockOut });
          }
        }
      }
    }
    
    if (employeeRecords.size > 0) {
      // Store under multiple name variations for better matching
      result.set(fullName.toLowerCase(), employeeRecords);
      if (reverseName.toLowerCase() !== fullName.toLowerCase()) {
        result.set(reverseName.toLowerCase(), employeeRecords);
      }
      // Also store just first name and just last name for partial matching
      if (firstName && firstName.length > 2) {
        result.set(firstName.toLowerCase(), employeeRecords);
      }
      if (lastName && lastName.length > 2) {
        result.set(lastName.toLowerCase(), employeeRecords);
      }
      console.log(`Online: ${fullName} has ${employeeRecords.size} records. Sample dates: ${Array.from(employeeRecords.keys()).slice(0, 3).join(', ')}`);
    }
  }

  console.log(`Total employees parsed from Online: ${result.size / 4} (with name variations: ${result.size})`);
  return result;
}

/**
 * Helper to get earlier time (for clock-in)
 */
function getEarlierTimeStr(time1: string | null, time2: string | null): string | null {
  if (!time1) return time2;
  if (!time2) return time1;
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);
  if (min1 === null) return time2;
  if (min2 === null) return time1;
  return min1 <= min2 ? time1 : time2;
}

/**
 * Helper to get later time (for clock-out)
 */
function getLaterTimeStr(time1: string | null, time2: string | null): string | null {
  if (!time1) return time2;
  if (!time2) return time1;
  const min1 = parseTimeToMinutes(time1);
  const min2 = parseTimeToMinutes(time2);
  if (min1 === null) return time2;
  if (min2 === null) return time1;
  return min1 >= min2 ? time1 : time2;
}

/**
 * Parse time string to minutes (local helper)
 */
function parseTimeToMinutes(time: string | null): number | null {
  if (!time) return null;
  const parts = time.split(':');
  if (parts.length < 2) return null;
  const hours = parseInt(parts[0], 10);
  const minutes = parseInt(parts[1], 10);
  if (isNaN(hours) || isNaN(minutes)) return null;
  return hours * 60 + minutes;
}

/**
 * Get unique employees from fingerprint records
 */
export function getUniqueEmployees(records: RawFingerprintRecord[]): { empNo: string; name: string }[] {
  const seen = new Map<string, { empNo: string; name: string }>();
  
  for (const record of records) {
    if (!seen.has(record.name.toLowerCase())) {
      seen.set(record.name.toLowerCase(), {
        empNo: record.empNo,
        name: record.name,
      });
    }
  }
  
  return Array.from(seen.values());
}

/**
 * Get all dates from a month
 */
export function getMonthDates(year: number, month: number): Date[] {
  const dates: Date[] = [];
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  
  for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
    dates.push(new Date(d));
  }
  
  return dates;
}
