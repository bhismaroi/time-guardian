// Core types for the attendance system

export interface EmployeeAttendanceRecord {
  date: Date;
  dayOfWeek: string;
  actualIn: string | null;
  actualOut: string | null;
  totalHours: string;
  tardiness: string | null;
  leaveEarlier: string | null;
  overtime: string | null;
  remarks: string;
}

export interface EmployeeData {
  empNo: string;
  name: string;
  nip: string;
  records: EmployeeAttendanceRecord[];
}

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

export interface RawOnlineRecord {
  lastName: string;
  firstName: string;
  date: string;
  clockIn: string;
  clockOut: string;
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
  nip: string;
  division: string;
  department: string;
  section: string;
  records: MergedAttendanceRecord[];
}

export type FlexiType = 'flexi1' | 'flexi2' | 'late';

export interface AttendanceCalculation {
  totalMinutes: number;
  breakMinutes: number;
  workMinutes: number;
  overtimeMinutes: number;
  tardinessMinutes: number;
  leaveEarlierMinutes: number;
  flexiType: FlexiType | null;
}
