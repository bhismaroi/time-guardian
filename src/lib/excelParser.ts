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
 */
export function parseOnlineExcel(file: ArrayBuffer): Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>> {
  const workbook = XLSX.read(file, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 }) as unknown[][];

  // Result: Map<employeeName, Map<dateStr, {clockIn, clockOut}>>
  const result = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();

  // Find the header rows that contain date information
  // Row 10 (index 9) typically contains date headers like "01 Oct, We"
  const dateRow = data[9];
  if (!dateRow || !Array.isArray(dateRow)) return result;

  // Find the Clock-in/Clock-out section - typically starts around column index 70+
  // We need to find where the actual clock data is
  
  // Parse employee rows (starting from row 11, index 10)
  for (let rowIdx = 10; rowIdx < data.length; rowIdx++) {
    const row = data[rowIdx];
    if (!row || !Array.isArray(row) || row.length < 3) continue;

    const lastName = String(row[0] || '').trim();
    const firstName = String(row[1] || '').trim();
    
    if (!lastName && !firstName) continue;
    
    const fullName = firstName ? `${firstName} ${lastName}`.trim() : lastName;
    
    // Find clock-in/clock-out data - it's in a specific section of the columns
    // The data includes patterns like "08:12 - __" or "08:00 - 17:00" or "__ - 17:00"
    const employeeRecords = new Map<string, { clockIn: string | null; clockOut: string | null }>();
    
    // The clock-in/clock-out section appears after several other data sections
    // We need to look for the pattern in the row data
    // Typically it's around columns 70-100 area based on the structure
    
    for (let col = 0; col < row.length; col++) {
      const cellValue = String(row[col] || '');
      
      // Look for clock-in/clock-out pattern: "HH:MM - HH:MM" or "HH:MM - __" or "__ - HH:MM"
      const clockPattern = cellValue.match(/(\d{1,2}:\d{2}|__|_\s*_)\s*-\s*(\d{1,2}:\d{2}|__|_\s*_)/);
      
      if (clockPattern && dateRow[col]) {
        const dateStr = String(dateRow[col]);
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
            employeeRecords.set(normalizedDate, { clockIn, clockOut });
          }
        }
      }
    }
    
    if (employeeRecords.size > 0) {
      result.set(fullName.toLowerCase(), employeeRecords);
    }
  }

  return result;
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
