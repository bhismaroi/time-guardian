import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { calculateAttendance } from '@/lib/attendanceCalculator';
import { compileAttendance } from '@/lib/attendanceCompiler';
import { buildAttendanceWorkbook } from '@/lib/excelGenerator';
import { parseOnlineExcel } from '@/lib/excelParser';
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

  it('does not mark leave earlier for clock-ins before 08:00', () => {
    const date = new Date(2025, 9, 6);
    const result = calculateAttendance(date, '07:46', '16:40');

    expect(result.leaveEarlierMinutes).toBe(0);
  });
});

describe('attendance compilation', () => {
  it('parses the new block-style online workbook format', () => {
    const rows: unknown[][] = [];
    rows[1] = ['Report'];
    rows[2] = ['Mar 1, 2026 - Mar 31, 2026'];
    rows[6] = [null, 'Full name', 'Adi Misykatul'];
    rows[7] = [null, 'Code', 'E-427'];
    rows[8] = [null, 'Position', 'Staff'];
    rows[9] = [null, 'Department', 'IDACT'];
    rows[10] = [null, 'Location', 'Kantor Menara Astra'];
    rows[11] = [null, 'Schedule', 'Template', 'Clock-in', 'Clock-out', 'Worked', 'Late', 'Overtime (non approved)', 'Early departure', 'Worked on day off'];
    rows[12] = ['01 Mar, Su', 'DO', null, '-', '-', 0, 0, 0, 0, 0];
    rows[13] = ['02 Mar, Mo', '08:00 - 16:30', null, '08:10', '18:10', 1, 0, 0, 0, 0];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const parsed = parseOnlineExcel(buffer);

    expect(parsed.get('adi misykatul')?.get('2026-03-02')).toEqual({
      clockIn: '08:10',
      clockOut: '18:10',
    });
  });

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
    const sheet = workbook.Sheets['Adi'];

    expect(workbook.SheetNames).toContain('Template');
    expect(workbook.SheetNames).toContain('Adi');
    expect(sheet?.['A6']?.v).toBeTypeOf('number');
    expect(sheet?.['A6']?.z).toBe('dd/mm');
    expect(sheet?.['I6']?.f).toContain('H6-G6');
    expect(sheet?.['I6']?.z).toBe('[h]:mm');
    expect(sheet?.['K6']?.f).toContain('TIME(8,15,0)');
    expect(sheet?.['K6']?.f).toContain('TIME(8,30,0)');
    expect(sheet?.['K6']?.f).toContain('TIME(8,0,0)');
    expect(sheet?.['K6']?.z).toBe('[h]:mm');
    expect(sheet?.['A6']?.v).not.toBe('Divisi : MITSUI OSK LINES');
    expect(sheet?.['A6']?.v).not.toBe('NIP : 000427   Nama : ADI MISYKATUL ANWAR');
  });
});
