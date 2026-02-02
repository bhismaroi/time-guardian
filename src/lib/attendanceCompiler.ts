// Attendance data compilation logic

import type { CompiledEmployee, MergedAttendanceRecord, RawFingerprintRecord } from './types';
import { 
  getDayName, 
  isWeekend, 
  formatDateFull, 
  getEarlierTime, 
  getLaterTime,
  parseDate 
} from './timeUtils';
import { calculateAttendance, formatCalculationResults } from './attendanceCalculator';
import { getUniqueEmployees, getMonthDates } from './excelParser';

/**
 * Compile attendance data from fingerprint and online sources
 */
export function compileAttendance(
  fingerprintRecords: RawFingerprintRecord[],
  onlineData: Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>
): CompiledEmployee[] {
  const employees = getUniqueEmployees(fingerprintRecords);
  const compiledEmployees: CompiledEmployee[] = [];

  // Generate dates for October 2025 (since that's what the files contain)
  const dates = getMonthDates(2025, 9); // Month is 0-indexed, so 9 = October

  console.log('Compiling attendance for', employees.length, 'employees');
  console.log('Online data keys:', Array.from(onlineData.keys()));

  for (const employee of employees) {
    // Group fingerprint records by date for this employee
    const fingerprintByDate = new Map<string, { in: string | null; out: string | null }>();
    
    for (const record of fingerprintRecords) {
      if (record.name.toLowerCase() === employee.name.toLowerCase()) {
        const dateKey = record.date;
        const existing = fingerprintByDate.get(dateKey);
        
        if (existing) {
          // Get earliest in and latest out
          existing.in = getEarlierTime(existing.in, record.actualIn || null);
          existing.out = getLaterTime(existing.out, record.actualOut || null);
        } else {
          fingerprintByDate.set(dateKey, {
            in: record.actualIn || null,
            out: record.actualOut || null,
          });
        }
      }
    }

    // Try to find online data for this employee using multiple matching strategies
    // The online file uses "FirstName LastName" format, fingerprint uses full name
    let employeeOnlineData = onlineData.get(employee.name.toLowerCase());
    
    // If not found, try various matching strategies
    if (!employeeOnlineData) {
      const nameParts = employee.name.toLowerCase().split(/\s+/).filter(p => p.length > 2);
      
      // Strategy 1: Try each name part as a direct key
      for (const part of nameParts) {
        if (onlineData.has(part)) {
          employeeOnlineData = onlineData.get(part);
          console.log(`Matched "${employee.name}" to online via part "${part}"`);
          break;
        }
      }
      
      // Strategy 2: Try partial matching on the keys
      if (!employeeOnlineData) {
        for (const [onlineName, data] of onlineData) {
          const onlineParts = onlineName.split(/\s+/).filter(p => p.length > 2);
          
          // Check if any significant part matches
          const matchCount = nameParts.filter(part => 
            onlineParts.some(op => op === part || op.includes(part) || part.includes(op))
          ).length;
          
          // Require at least one good match
          if (matchCount >= 1) {
            employeeOnlineData = data;
            console.log(`Matched "${employee.name}" to online "${onlineName}" (${matchCount} parts matched)`);
            break;
          }
        }
      }
    }

    if (employeeOnlineData) {
      console.log(`Employee "${employee.name}" has ${employeeOnlineData.size} online records`);
    } else {
      console.log(`Employee "${employee.name}" has NO online data`);
    }

    // Build records for each date
    const records: MergedAttendanceRecord[] = [];

    for (const date of dates) {
      const dateStr = formatDateFull(date);
      const dayName = getDayName(date);
      
      // Get fingerprint data
      const fingerprint = fingerprintByDate.get(dateStr);
      const fingerprintIn = fingerprint?.in || null;
      const fingerprintOut = fingerprint?.out || null;
      
      // Get online data
      const online = employeeOnlineData?.get(dateStr);
      const onlineIn = online?.clockIn || null;
      const onlineOut = online?.clockOut || null;
      
      // Merge: get earliest clock-in and latest clock-out from both sources
      const actualIn = getEarlierTime(fingerprintIn, onlineIn);
      const actualOut = getLaterTime(fingerprintOut, onlineOut);
      
      // Debug log for dates with data
      if (fingerprintIn || fingerprintOut || onlineIn || onlineOut) {
        console.log(`${employee.name} ${dateStr}: FP(${fingerprintIn}/${fingerprintOut}) + Online(${onlineIn}/${onlineOut}) = (${actualIn}/${actualOut})`);
      }
      
      // Calculate attendance
      const calculation = calculateAttendance(date, actualIn, actualOut);
      const formatted = formatCalculationResults(calculation);
      
      // Determine remarks
      let remarks = '';
      if (isWeekend(date)) {
        remarks = '';
      } else if (!actualIn && !actualOut) {
        remarks = '0';
      }

      records.push({
        date,
        dayOfWeek: dayName,
        fingerprintIn,
        fingerprintOut,
        onlineIn,
        onlineOut,
        actualIn,
        actualOut,
        totalHours: isWeekend(date) ? '' : formatted.totalHours,
        tardiness: formatted.tardiness,
        leaveEarlier: formatted.leaveEarlier,
        overtime: formatted.overtime,
        remarks,
      });
    }

    compiledEmployees.push({
      empNo: employee.empNo,
      name: employee.name,
      nip: `000${employee.empNo}`.slice(-6),
      division: 'MITSUI OSK LINES',
      department: 'IDACT',
      section: 'IDACT',
      records,
    });
  }

  return compiledEmployees;
}
