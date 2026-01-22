// Excel file parsing utilities

import * as XLSX from 'xlsx';
import type { RawFingerprintRecord, RawOnlineRecord } from './types';
import { extractTime, parseDate } from './timeUtils';

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
 * This file has a complex structure with multiple columns per date
 * Row 9 (index 8) contains section headers like "Clock-in / Clock-out"
 * Row 10 (index 9) contains date headers like "01 Oct, We"
 * Row 11+ contains employee data
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
  
  if (!dateRow || !Array.isArray(dateRow)) return result;
  if (!sectionHeaderRow || !Array.isArray(sectionHeaderRow)) return result;

  // Find the start column of "Clock-in / Clock-out" section
  let clockSectionStart = -1;
  let clockSectionEnd = -1;
  
  for (let col = 0; col < sectionHeaderRow.length; col++) {
    const header = String(sectionHeaderRow[col] || '').toLowerCase();
    if (header.includes('clock-in') && header.includes('clock-out')) {
      clockSectionStart = col;
      // Find the end of this section (next non-empty header or section with dates)
      for (let endCol = col + 1; endCol < sectionHeaderRow.length; endCol++) {
        const nextHeader = String(sectionHeaderRow[endCol] || '').trim();
        if (nextHeader && !nextHeader.toLowerCase().includes('clock')) {
          clockSectionEnd = endCol;
          break;
        }
      }
      if (clockSectionEnd === -1) {
        clockSectionEnd = clockSectionStart + 32; // Assume 31 days max
      }
      break;
    }
  }

  // If we couldn't find the section by header, fall back to pattern matching
  if (clockSectionStart === -1) {
    // Look for columns that contain clock patterns in employee rows
    for (let col = 0; col < (data[10]?.length || 0); col++) {
      const cellValue = String(data[10]?.[col] || '');
      if (cellValue.match(/(\d{1,2}:\d{2}|_+)\s*-\s*(\d{1,2}:\d{2}|_+)/)) {
        clockSectionStart = col;
        clockSectionEnd = col + 32;
        break;
      }
    }
  }

  console.log('Clock section found:', clockSectionStart, 'to', clockSectionEnd);

  // Parse employee rows (starting from row 11, index 10)
  for (let rowIdx = 10; rowIdx < data.length; rowIdx++) {
    const row = data[rowIdx];
    if (!row || !Array.isArray(row) || row.length < 3) continue;

    const lastName = String(row[0] || '').trim();
    const firstName = String(row[1] || '').trim();
    
    if (!lastName && !firstName) continue;
    
    const fullName = firstName ? `${firstName} ${lastName}`.trim() : lastName;
    
    const employeeRecords = new Map<string, { clockIn: string | null; clockOut: string | null }>();
    
    // Parse the clock-in/clock-out section columns
    const startCol = clockSectionStart !== -1 ? clockSectionStart : 0;
    const endCol = clockSectionEnd !== -1 ? clockSectionEnd : row.length;
    
    for (let col = startCol; col < Math.min(endCol, row.length); col++) {
      const cellValue = String(row[col] || '');
      
      // Look for clock-in/clock-out pattern: "HH:MM - HH:MM" or "HH:MM - __" or "__ - HH:MM"
      const clockPattern = cellValue.match(/(\d{1,2}:\d{2}|_+)\s*-\s*(\d{1,2}:\d{2}|_+)/);
      
      if (clockPattern) {
        // Get the date from the date row at the same column
        const dateStr = String(dateRow[col] || '');
        const dateMatch = dateStr.match(/(\d{1,2})\s+(\w+),?\s*(\w+)?/);
        
        if (dateMatch) {
          const day = parseInt(dateMatch[1], 10);
          const monthStr = dateMatch[2];
          const monthMap: Record<string, number> = {
            'Oct': 10, 'Nov': 11, 'Dec': 12, 'Jan': 1, 'Feb': 2, 'Mar': 3,
            'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9
          };
          const month = monthMap[monthStr] || 10;
          const normalizedDate = `${day.toString().padStart(2, '0')}/${month.toString().padStart(2, '0')}/2025`;
          
          const clockIn = clockPattern[1].includes('_') ? null : extractTime(clockPattern[1]);
          const clockOut = clockPattern[2].includes('_') ? null : extractTime(clockPattern[2]);
          
          if (clockIn || clockOut) {
            // If we already have a record for this date, merge the times
            const existing = employeeRecords.get(normalizedDate);
            if (existing) {
              // Get earliest clock-in and latest clock-out
              const mergedIn = getEarlierTimeStr(existing.clockIn, clockIn);
              const mergedOut = getLaterTimeStr(existing.clockOut, clockOut);
              employeeRecords.set(normalizedDate, { clockIn: mergedIn, clockOut: mergedOut });
            } else {
              employeeRecords.set(normalizedDate, { clockIn, clockOut });
            }
          }
        }
      }
    }
    
    if (employeeRecords.size > 0) {
      result.set(fullName.toLowerCase(), employeeRecords);
      console.log(`Online: ${fullName} has ${employeeRecords.size} records`);
    }
  }

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
