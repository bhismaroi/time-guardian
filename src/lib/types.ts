// Core types for the attendance system
//
// The runtime shapes that ended up being used are MergedAttendanceRecord
// (one row per employee per day in the compiled output) and
// CompiledEmployee (one entry per employee). The other interfaces
// (EmployeeAttendanceRecord, EmployeeData, RawOnlineRecord) were
// vestigial: they were declared early in the project but never
// matched the actual shape returned by the compiler. Removed in
// Phase 5.5.

export interface RawFingerprintRecord {
  empNo: string;
  name: string;
  date: string;
  dateKey: string;
  workingHours: string;
  clockIn: string | null;
  clockOut: string | null;
  actualIn: string | null;
  actualOut: string | null;
}

export interface MergedAttendanceRecord {
  date: Date;
  dayOfWeek: string;
  fingerprintIn: string | null;
  fingerprintOut: string | null;
  onlineIn: string | null;
  onlineOut: string | null;
  actualIn: string | null;
  actualOut: string | null;
  totalHours: string;
  tardiness: string | null;
  leaveEarlier: string | null;
  overtime: string | null;
  remarks: string;
}

export interface CompiledEmployee {
  empNo: string;
  name: string;
  sheetName: string;
  records: MergedAttendanceRecord[];
}

export type FlexiType = 'standard' | 'flexi1' | 'flexi2' | 'late';

export interface AttendanceCalculation {
  totalMinutes: number;
  breakMinutes: number;
  workMinutes: number;
  overtimeMinutes: number;
  tardinessMinutes: number;
  leaveEarlierMinutes: number;
  flexiType: FlexiType | null;
}
