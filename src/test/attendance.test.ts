import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { calculateAttendance } from '@/lib/attendanceCalculator';
import { compileAttendance } from '@/lib/attendanceCompiler';
import { buildAttendanceWorkbook } from '@/lib/excelGenerator';
import type { RawFingerprintRecord } from '@/lib/types';

describe('attendance calculations', () => {
  it('deducts the correct break and overtime on a Monday', () => {
    const date = new Date(2025, 9, 6);
    const result = calculateAttendance(date, '08:10', '17:40');

    expect(result.breakMinutes).toBe(30);
    expect(result.workMinutes).toBe(540);
    expect(result.overtimeMinutes).toBe(10);
    expect(result.tardinessMinutes).toBe(0);
  });

  it('deducts the Friday lunch break and flags flexi overtime', () => {
    const date = new Date(2025, 9, 3);
    const result = calculateAttendance(date, '08:20', '18:15');

    expect(result.breakMinutes).toBe(90);
    expect(result.workMinutes).toBe(505);
    expect(result.overtimeMinutes).toBe(15);
    expect(result.tardinessMinutes).toBe(0);
  });
});

describe('attendance compilation', () => {
  it('matches names using first and last name overlap and merges earliest/latest times', () => {
    const fingerprintRecords: RawFingerprintRecord[] = [
      {
        empNo: '427',
        name: 'Adi Misykatul Anwar',
        date: '2025-10-01',
        dateKey: '2025-10-01',
        workingHours: 'Office Hour',
        clockIn: '08:15',
        clockOut: '17:00',
        actualIn: '08:15',
        actualOut: '17:00',
      },
    ];

    const onlineData = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();
    onlineData.set(
      'misykatul anwar adi',
      new Map([
        ['2025-10-01', { clockIn: '08:10', clockOut: '18:14' }],
      ])
    );

    const compiled = compileAttendance(fingerprintRecords, onlineData);
    const employee = compiled[0];
    const firstDay = employee.records.find((record) => record.date.getDate() === 1);

    expect(employee.name).toBe('Adi Misykatul Anwar');
    expect(employee.sheetName).toBe('Adi');
    expect(firstDay?.actualIn).toBe('08:10');
    expect(firstDay?.actualOut).toBe('18:14');
  });

  it('writes a workbook with formula cells for calculated columns', async () => {
    const compiled = compileAttendance(
      [
        {
          empNo: '427',
          name: 'Adi Misykatul Anwar',
          date: '2025-10-01',
          dateKey: '2025-10-01',
          workingHours: 'Office Hour',
          clockIn: '08:10',
          clockOut: '17:40',
          actualIn: '08:10',
          actualOut: '17:40',
        },
        {
          empNo: '427',
          name: 'Adi Misykatul Anwar',
          date: '2025-10-02',
          dateKey: '2025-10-02',
          workingHours: 'Office Hour',
          clockIn: '08:10',
          clockOut: '17:40',
          actualIn: '08:10',
          actualOut: '17:40',
        },
      ],
      new Map()
    );

    const workbook = buildAttendanceWorkbook(compiled);
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const parsedWorkbook = XLSX.read(buffer, { type: 'array', cellFormula: true });
    const sheet = parsedWorkbook.Sheets['Adi'];

    expect(parsedWorkbook.SheetNames).toContain('Template');
    expect(parsedWorkbook.SheetNames).toContain('Adi');
    expect(sheet?.['I10']?.f).toContain('H10-G10');
    expect(sheet?.['L10']?.f).toContain('TIME(17,30,0)');
  });
});
