// Business logic for attendance calculations

import {
  parseTimeToMinutes,
  minutesToTimeString,
  isFriday,
  isWeekend,
  calculateBreakOverlap,
} from './timeUtils';
import {
  BREAK_MON_THU,
  BREAK_FRI,
  STANDARD_CLOCKIN_END,
  FLEXI_1_START,
  FLEXI_1_END,
  FLEXI_2_START,
  FLEXI_2_END,
  LATE_THRESHOLD,
  getBreakWindow,
  getAllowedClockOut as getAllowedClockOutPolicy,
  getOvertimeThreshold as getOvertimeThresholdPolicy,
  determineFlexiType as determineFlexiTypePolicy,
  isFriday as isFridayPolicy,
  isWeekend as isWeekendPolicy,
} from './policy';
import type { AttendanceCalculation, FlexiType } from './types';

// The constants and helpers above live in shared/policy.js. They are
// re-exported from src/lib/policy.ts with TypeScript types. We re-import
// the runtime values here so the calculator has a single source of
// truth for break windows, flexi thresholds, and clock-out rules.

// The local determineFlexiType below delegates to the policy version.
// We keep the local name to preserve the existing function signature
// and JSDoc comment (and so existing tests keep working).
export function determineFlexiType(clockInMinutes: number): FlexiType {
  return determineFlexiTypePolicy(clockInMinutes);
}

/**
 * Get the break duration for a given date
 */
export function getBreakDuration(date: Date): { start: number; end: number; duration: number } {
  const win = getBreakWindow(date);
  return { start: win.start, end: win.end, duration: win.duration };
}

/**
 * Get the allowed clock-out time based on flexi type and day
 */
export function getAllowedClockOut(flexiType: FlexiType, date: Date): number {
  return getAllowedClockOutPolicy(flexiType, date);
}

/**
 * Get the overtime start threshold for a given date
 */
export function getOvertimeThreshold(date: Date): number {
  return getOvertimeThresholdPolicy(date);
}

/**
 * Calculate attendance metrics for a single day
 */
export function calculateAttendance(
  date: Date,
  clockIn: string | null,
  clockOut: string | null
): AttendanceCalculation {
  const result: AttendanceCalculation = {
    totalMinutes: 0,
    breakMinutes: 0,
    workMinutes: 0,
    overtimeMinutes: 0,
    tardinessMinutes: 0,
    leaveEarlierMinutes: 0,
    flexiType: null,
  };

  // If weekend, no calculation needed
  if (isWeekend(date)) {
    return result;
  }

  const clockInMinutes = parseTimeToMinutes(clockIn);
  const clockOutMinutes = parseTimeToMinutes(clockOut);

  // If no clock-in or clock-out, return empty
  if (clockInMinutes === null || clockOutMinutes === null) {
    return result;
  }

  // Calculate raw total time
  result.totalMinutes = Math.max(0, clockOutMinutes - clockInMinutes);

  // Determine flexi type
  const flexiType = determineFlexiType(clockInMinutes);
  result.flexiType = flexiType;

  // Calculate tardiness (only if late)
  if (clockInMinutes > LATE_THRESHOLD) {
    result.tardinessMinutes = clockInMinutes - LATE_THRESHOLD;
  }

  // Get break info for the day
  const breakInfo = getBreakDuration(date);

  // Calculate break overlap
  result.breakMinutes = calculateBreakOverlap(
    clockInMinutes,
    clockOutMinutes,
    breakInfo.start,
    breakInfo.end
  );

  // Calculate work minutes (total - break)
  result.workMinutes = Math.max(0, result.totalMinutes - result.breakMinutes);

  // Calculate leave earlier
  const allowedClockOut = getAllowedClockOut(flexiType, date);
  if (clockOutMinutes < allowedClockOut) {
    result.leaveEarlierMinutes = allowedClockOut - clockOutMinutes;
  }

  // Calculate overtime
  const overtimeThreshold = getOvertimeThreshold(date);
  if (clockOutMinutes > overtimeThreshold) {
    result.overtimeMinutes = clockOutMinutes - overtimeThreshold;
  }

  return result;
}

/**
 * Format calculation result to display strings
 */
export function formatCalculationResults(calc: AttendanceCalculation): {
  totalHours: string;
  tardiness: string | null;
  leaveEarlier: string | null;
  overtime: string | null;
} {
  return {
    totalHours: minutesToTimeString(calc.workMinutes),
    tardiness: calc.tardinessMinutes > 0 ? minutesToTimeString(calc.tardinessMinutes) : null,
    leaveEarlier: calc.leaveEarlierMinutes > 0 ? minutesToTimeString(calc.leaveEarlierMinutes) : null,
    overtime: calc.overtimeMinutes > 0 ? minutesToTimeString(calc.overtimeMinutes) : null,
  };
}
