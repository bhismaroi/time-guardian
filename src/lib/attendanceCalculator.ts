// Business logic for attendance calculations
//
// The policy module (shared/policy.js, re-exported via src/lib/policy.ts)
// is the source of truth for the break window, flexi thresholds, and
// clock-out rules. We re-import the runtime values here and re-export
// the policy helpers under their existing names so callers in this
// codebase (the compiler, tests, etc.) keep working without an
// import rewrite.

import {
  parseTimeToMinutes,
  minutesToTimeString,
  calculateBreakOverlap,
} from './timeUtils';
import {
  LATE_THRESHOLD,
  getBreakWindow,
  getAllowedClockOut,
  getOvertimeThreshold,
  determineFlexiType,
  isWeekend,
} from './policy';
import type { AttendanceCalculation, FlexiType } from './types';

// Re-export the policy helpers under their existing names.
export { determineFlexiType, getAllowedClockOut, getOvertimeThreshold };

// getBreakWindow returns the same { start, end, duration } shape the
// local getBreakDuration used to. Exposed under the old name for
// backward compat with external callers.
export function getBreakDuration(date: Date): { start: number; end: number; duration: number } {
  return getBreakWindow(date);
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
