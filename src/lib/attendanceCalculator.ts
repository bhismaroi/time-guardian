// Business logic for attendance calculations

import {
  parseTimeToMinutes,
  minutesToTimeString,
  isFriday,
  isMondayOrThursday,
  isWeekend,
  calculateBreakOverlap,
} from './timeUtils';
import type { AttendanceCalculation, FlexiType } from './types';

// Break time constants (in minutes from midnight)
const BREAK_MON_THU_START = 12 * 60; // 12:00
const BREAK_MON_THU_END = 12 * 60 + 30; // 12:30
const BREAK_FRI_START = 11 * 60 + 30; // 11:30
const BREAK_FRI_END = 13 * 60; // 13:00

// Flexi time thresholds (in minutes from midnight)
const FLEXI_1_START = 8 * 60; // 08:00
const FLEXI_1_END = 8 * 60 + 15; // 08:15
const FLEXI_2_START = 8 * 60 + 15; // 08:15
const FLEXI_2_END = 8 * 60 + 30; // 08:30
const LATE_THRESHOLD = 8 * 60 + 30; // 08:30

// Allowed clock-out times
const FLEXI_1_CLOCKOUT_MON_THU = 16 * 60 + 45; // 16:45
const FLEXI_1_CLOCKOUT_FRI = 17 * 60 + 15; // 17:15
const FLEXI_2_CLOCKOUT_MON_THU = 17 * 60; // 17:00
const FLEXI_2_CLOCKOUT_FRI = 17 * 60 + 30; // 17:30

// Overtime thresholds
const OVERTIME_START_MON_THU = 17 * 60 + 30; // 17:30
const OVERTIME_START_FRI = 18 * 60; // 18:00

/**
 * Determine the flexi type based on clock-in time
 */
export function determineFlexiType(clockInMinutes: number): FlexiType {
  if (clockInMinutes >= FLEXI_1_START && clockInMinutes < FLEXI_1_END) {
    return 'flexi1';
  }
  if (clockInMinutes >= FLEXI_2_START && clockInMinutes < FLEXI_2_END) {
    return 'flexi2';
  }
  return 'late';
}

/**
 * Get the break duration for a given date
 */
export function getBreakDuration(date: Date): { start: number; end: number; duration: number } {
  if (isFriday(date)) {
    return {
      start: BREAK_FRI_START,
      end: BREAK_FRI_END,
      duration: 90, // 90 minutes
    };
  }
  // Monday to Thursday
  return {
    start: BREAK_MON_THU_START,
    end: BREAK_MON_THU_END,
    duration: 30, // 30 minutes
  };
}

/**
 * Get the allowed clock-out time based on flexi type and day
 */
export function getAllowedClockOut(flexiType: FlexiType, date: Date): number {
  if (isFriday(date)) {
    return flexiType === 'flexi1' ? FLEXI_1_CLOCKOUT_FRI : FLEXI_2_CLOCKOUT_FRI;
  }
  return flexiType === 'flexi1' ? FLEXI_1_CLOCKOUT_MON_THU : FLEXI_2_CLOCKOUT_MON_THU;
}

/**
 * Get the overtime start threshold for a given date
 */
export function getOvertimeThreshold(date: Date): number {
  return isFriday(date) ? OVERTIME_START_FRI : OVERTIME_START_MON_THU;
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
  result.flexiType = determineFlexiType(clockInMinutes);

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
  if (clockInMinutes >= FLEXI_1_START && clockInMinutes < FLEXI_2_END) {
    const allowedClockOut = getAllowedClockOut(result.flexiType === 'late' ? 'flexi2' : result.flexiType, date);
    if (clockOutMinutes < allowedClockOut) {
      result.leaveEarlierMinutes = allowedClockOut - clockOutMinutes;
    }
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
