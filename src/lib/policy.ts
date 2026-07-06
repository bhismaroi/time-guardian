// TypeScript re-export of the canonical policy module.
//
// The actual source of truth is shared/policy.js — a vanilla ES2020
// module that has no dependencies and is consumable by both the React
// app and the Cloudflare Pages static build (which loads it via a
// <script> tag). This file re-exports the values and adds TypeScript
// types so the React app can import the policy as if it were a normal
// .ts module.

import {
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
  reactDayChecks,
  cloudflareDayChecks,
  POLICY_DESCRIPTIONS,
} from '../../shared/policy.js';

export type DayChecks = {
  friday: string;
  weekday: string;
  weekend: string;
};

export type FlexiType = 'standard' | 'flexi1' | 'flexi2' | 'late';

export interface BreakWindow {
  start: number;
  end: number;
  duration: number;
}

export {
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
  reactDayChecks,
  cloudflareDayChecks,
  POLICY_DESCRIPTIONS,
};
