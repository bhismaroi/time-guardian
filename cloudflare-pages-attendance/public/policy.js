// Attendance policy — single source of truth.
//
// This file is the canonical definition of the office attendance rules.
// It is consumed by:
//   - The TypeScript React app (via src/lib/policy.ts which re-exports
//     these values for the TypeScript calculator and Excel generator).
//   - The Cloudflare Pages static build (via cloudflare-pages-attendance/
//     public/policy.js, a copy of this file that the browser loads as a
//     plain <script> tag before browserCompiler.js).
//
// Both consumers must produce identical formulas for the same input. The
// formula builders below accept a `dayExpr` parameter (the cell name
// convention to use for the day-of-week check) so the React app can pass
// `WEEKDAY(A${row},2)=5` and the Cloudflare app can pass `B${row}="Fri"`,
// and the same JS function emits the right string for each environment.
//
// All time constants are minutes from midnight. Months are 0-based
// (Jan = 0) in MONTH_LOOKUP and `month` indices throughout.

const MONTH_LOOKUP = Object.freeze({
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
});

// Break windows (in minutes from midnight).
const BREAK_MON_THU = Object.freeze({ start: 12 * 60, end: 12 * 60 + 30, duration: 30 });
const BREAK_FRI = Object.freeze({ start: 11 * 60 + 30, end: 13 * 60, duration: 90 });

// Flexi time thresholds (in minutes from midnight). The boundaries are
// 08:00 (last standard clock-in), 08:01-08:15 (flexi 1), 08:16-08:30
// (flexi 2), anything later is "late".
const STANDARD_CLOCKIN_END = 8 * 60;            // 08:00
const FLEXI_1_START = 8 * 60 + 1;               // 08:01
const FLEXI_1_END = 8 * 60 + 15;                // 08:15
const FLEXI_2_START = 8 * 60 + 16;              // 08:16
const FLEXI_2_END = 8 * 60 + 30;                // 08:30
const LATE_THRESHOLD = 8 * 60 + 30;             // 08:30

// Allowed clock-out per flexi type and weekday group (in minutes from
// midnight). Late arrivals still fall back to the flexi-2 expected
// clock-out, not the standard one.
const STANDARD_CLOCKOUT_MON_THU = 16 * 60 + 30; // 16:30
const STANDARD_CLOCKOUT_FRI = 17 * 60;          // 17:00
const FLEXI_1_CLOCKOUT_MON_THU = 16 * 60 + 45;  // 16:45
const FLEXI_1_CLOCKOUT_FRI = 17 * 60 + 15;      // 17:15
const FLEXI_2_CLOCKOUT_MON_THU = 17 * 60;       // 17:00
const FLEXI_2_CLOCKOUT_FRI = 17 * 60 + 30;      // 17:30

// Overtime thresholds (any clock-out past this counts as overtime, in
// minutes from midnight).
const OVERTIME_START_MON_THU = 17 * 60 + 30;    // 17:30
const OVERTIME_START_FRI = 18 * 60;             // 18:00

// ---- Helpers ----

// True if the JS Date is a Friday (0 = Sun, 5 = Fri, 6 = Sat).
function isFriday(date) {
  return date.getDay() === 5;
}

// True if the JS Date is a weekend (Sat or Sun).
function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

// Determine the flexi type based on clock-in minutes from midnight.
function determineFlexiType(clockInMinutes) {
  if (clockInMinutes <= STANDARD_CLOCKIN_END) return 'standard';
  if (clockInMinutes >= FLEXI_1_START && clockInMinutes <= FLEXI_1_END) return 'flexi1';
  if (clockInMinutes >= FLEXI_2_START && clockInMinutes <= FLEXI_2_END) return 'flexi2';
  return 'late';
}

// Get the break window for a given date. Returns { start, end, duration }.
function getBreakWindow(date) {
  return isFriday(date) ? BREAK_FRI : BREAK_MON_THU;
}

// Get the allowed clock-out (in minutes from midnight) for a given flexi
// type and date. Late arrivals fall back to the flexi-2 expected out.
function getAllowedClockOut(flexiType, date) {
  const onFri = isFriday(date);
  if (flexiType === 'standard') return onFri ? STANDARD_CLOCKOUT_FRI : STANDARD_CLOCKOUT_MON_THU;
  if (flexiType === 'flexi1') return onFri ? FLEXI_1_CLOCKOUT_FRI : FLEXI_1_CLOCKOUT_MON_THU;
  // flexi2 OR late -> flexi2 expected out
  return onFri ? FLEXI_2_CLOCKOUT_FRI : FLEXI_2_CLOCKOUT_MON_THU;
}

// Get the overtime start threshold (in minutes from midnight) for a date.
function getOvertimeThreshold(date) {
  return isFriday(date) ? OVERTIME_START_FRI : OVERTIME_START_MON_THU;
}

// ---- Excel formula builders ----
//
// These produce Excel formula strings. Each builder takes (row,
// dayChecks) where dayChecks is { friday, weekday, weekend } — the
// three day-of-week boolean expressions to inline. The two callers
// produce these from their own column conventions:
//
//   React     (date is in column A, mode-2 weekday number):
//     friday  = WEEKDAY(A${row},2)=5
//     weekday = WEEKDAY(A${row},2)<=4
//     weekend = WEEKDAY(A${row},2)>5
//
//   Cloudflare (date is in column A but the day-name lives in column B):
//     friday  = B${row}="Fri"
//     weekday = OR(B${row}="Mon",B${row}="Tue",B${row}="Wed",B${row}="Thu")
//     weekend = OR(B${row}="Sat",B${row}="Sun")
//
// Output strings for the React case are byte-identical to the
// hand-written formulas that excelGenerator.ts used before Phase 2.

// Build the break-deduction formula.
function buildBreakFormula(row, dayChecks) {
  return `IF(${dayChecks.friday},IF(AND(G${row}<TIME(13,0,0),H${row}>TIME(11,30,0)),MIN(H${row},TIME(13,0,0))-MAX(G${row},TIME(11,30,0)),0),IF(${dayChecks.weekday},IF(AND(G${row}<TIME(12,30,0),H${row}>TIME(12,0,0)),MIN(H${row},TIME(12,30,0))-MAX(G${row},TIME(12,0,0)),0),0))`;
}

// Build the total-hours formula.
function buildTotalHoursFormula(row, dayChecks) {
  return `=IF(OR(G${row}="",H${row}="",${dayChecks.weekend}),"",MAX(0,(H${row}-G${row})-(${buildBreakFormula(row, dayChecks)})))`;
}

// Build the tardiness formula.
function buildTardinessFormula(row, dayChecks) {
  return `=IF(OR(G${row}="",${dayChecks.weekend}),"",MAX(0,G${row}-TIME(8,30,0)))`;
}

// Build the leave-earlier formula. Uses the Friday check to pick the
// expected clock-out per flexi type, then MAX(0, expected - actual).
function buildLeaveEarlierFormula(row, dayChecks) {
  const standardClockOut = `IF(${dayChecks.friday},TIME(17,0,0),TIME(16,30,0))`;
  const flexi1ClockOut = `IF(${dayChecks.friday},TIME(17,15,0),TIME(16,45,0))`;
  const flexi2ClockOut = `IF(${dayChecks.friday},TIME(17,30,0),TIME(17,0,0))`;
  const expectedClockOut = `IF(G${row}<=TIME(8,0,0),${standardClockOut},IF(G${row}<=TIME(8,15,0),${flexi1ClockOut},${flexi2ClockOut}))`;
  return `=IF(OR(G${row}="",H${row}="",${dayChecks.weekend}),"",MAX(0,${expectedClockOut}-H${row}))`;
}

// Build the overtime formula.
function buildOvertimeFormula(row, dayChecks) {
  return `=IF(OR(H${row}="",${dayChecks.weekend}),"",MAX(0,H${row}-IF(${dayChecks.friday},TIME(18,0,0),TIME(17,30,0))))`;
}

// Day-of-week checks for the React app (date in column A, mode-2
// weekday number where 1=Mon..7=Sun).
function reactDayChecks(row) {
  return {
    friday: `WEEKDAY(A${row},2)=5`,
    weekday: `WEEKDAY(A${row},2)<=4`,
    weekend: `WEEKDAY(A${row},2)>5`,
  };
}

// Day-of-week checks for the Cloudflare app (day-name in column B).
function cloudflareDayChecks(row) {
  return {
    friday: `B${row}="Fri"`,
    weekday: `OR(B${row}="Mon",B${row}="Tue",B${row}="Wed",B${row}="Thu")`,
    weekend: `OR(B${row}="Sat",B${row}="Sun")`,
  };
}

// User-facing descriptions for the help section in Index.tsx. Keeping
// them here (instead of inline in the React component) means a policy
// change updates the help text in lockstep with the calculator.
const POLICY_DESCRIPTIONS = Object.freeze({
  breakMonThu: 'Monday to Thursday: 30-minute lunch at 12:00 - 12:30.',
  breakFri: 'Friday: 90-minute lunch at 11:30 - 13:00.',
  flexiTime: 'Standard clock-in is up to 08:00. Flexi 1 is 08:01 - 08:15, Flexi 2 is 08:16 - 08:30. After 08:30 is considered late.',
  overtime: 'Overtime starts after 17:30 Mon-Thu and after 18:00 on Friday.',
  dataMerging: 'Earliest clock-in and latest clock-out are merged across the fingerprint and online sources.',
});


// Cloudflare Pages loads this file as a plain <script> (not a module),
// so the public API is attached to a global rather than imported.
// window.AttendancePolicy mirrors the named exports of shared/
// policy.js — keep the two in sync.
if (typeof window !== 'undefined') {
  window.AttendancePolicy = {
    MONTH_LOOKUP,
    BREAK_MON_THU,
    BREAK_FRI,
    STANDARD_CLOCKIN_END,
    FLEXI_1_START,
    FLEXI_1_END,
    FLEXI_2_START,
    FLEXI_2_END,
    LATE_THRESHOLD,
    STANDARD_CLOCKOUT_MON_THU,
    STANDARD_CLOCKOUT_FRI,
    FLEXI_1_CLOCKOUT_MON_THU,
    FLEXI_1_CLOCKOUT_FRI,
    FLEXI_2_CLOCKOUT_MON_THU,
    FLEXI_2_CLOCKOUT_FRI,
    OVERTIME_START_MON_THU,
    OVERTIME_START_FRI,
    isFriday,
    isWeekend,
    determineFlexiType,
    getBreakWindow,
    getAllowedClockOut,
    getOvertimeThreshold,
    buildBreakFormula,
    buildTotalHoursFormula,
    buildTardinessFormula,
    buildLeaveEarlierFormula,
    buildOvertimeFormula,
    POLICY_DESCRIPTIONS,
  };
}
