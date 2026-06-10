import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { calculateAttendance } from '@/lib/attendanceCalculator';
import { compileAttendance } from '@/lib/attendanceCompiler';
import { buildAttendanceWorkbook } from '@/lib/excelGenerator';
import { parseFingerprintExcel, parseOnlineExcel } from '@/lib/excelParser';
import { extractTime, parseTimeToMinutes } from '@/lib/timeUtils';
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

  it('uses standard hours for clock-ins before 08:00', () => {
    const date = new Date(2025, 9, 6);
    const result = calculateAttendance(date, '07:46', '16:40');

    expect(result.leaveEarlierMinutes).toBe(0);
  });

  it('uses standard hours for an exact 08:00 clock-in', () => {
    const date = new Date(2026, 2, 5);
    const result = calculateAttendance(date, '08:00', '16:34');

    expect(result.flexiType).toBe('standard');
    expect(result.leaveEarlierMinutes).toBe(0);
  });

  it('flags leave earlier against flexi 1 expected clock-out', () => {
    const date = new Date(2026, 2, 11);
    const result = calculateAttendance(date, '08:07', '16:30');

    expect(result.flexiType).toBe('flexi1');
    expect(result.leaveEarlierMinutes).toBe(15);
  });

  it('flags leave earlier before standard clock-out for early clock-ins', () => {
    const date = new Date(2026, 2, 5);
    const result = calculateAttendance(date, '07:46', '16:20');

    expect(result.flexiType).toBe('standard');
    expect(result.leaveEarlierMinutes).toBe(10);
  });

  it('ignores invalid clock times instead of calculating with impossible values', () => {
    expect(parseTimeToMinutes('25:99')).toBeNull();
    expect(extractTime('clocked 24:00')).toBeNull();
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

  it('falls back to matrix format when a matrix workbook has a stray "Full name" cell', () => {
    // Regression: a single "Full name" cell in column B used to flip the
    // auto-detect into block mode, which then returned zero records
    // because the block parser also needs a "Schedule" header within 10
    // rows. The fix tries block first and falls through to matrix when
    // the block parse yields nothing.
    //
    // Matrix format: row 0 is the header (cols 0/1 = "Last name"/"First
    // name"), row 1 is the date row (cells from col 2 onward hold date
    // labels), subsequent rows are employees with time ranges directly
    // under each date column. We embed a stray "Full name" cell at the
    // end of row 2 to trip the old auto-detect.
    const rows: unknown[][] = [];
    rows[0] = ['Last name', 'First name'];
    rows[1] = ['Dates', 'Dates', '02 Mar, Mo', '03 Mar, Tu'];
    rows[2] = ['Smith', 'Adi', '08:00 - 16:30', '08:10 - 18:10', '', '', '', '', '', '', '', 'Full name'];
    rows[3] = ['Doe', 'Budi', '08:00 - 16:30', '08:05 - 18:00', '', '', '', '', '', '', '', ''];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const parsed = parseOnlineExcel(buffer);

    // The matrix parser should have produced records for both employees.
    // Pre-fix, the stray "Full name" cell flipped the parser to block
    // mode and the result was an empty map.
    expect(parsed.size).toBeGreaterThan(0);
    expect(parsed.get('adi smith')).toBeDefined();
    expect(parsed.get('budi doe')).toBeDefined();
  });

  it('tags cross-year date labels with the correct calendar year (block format)', () => {
    // Regression for the bug where parseDateFromLabel used a single
    // reportYear for every date label, so a "Dec 1, 2025 - Jan 31, 2026"
    // report tagged January dates as 2025-01-01 and never matched the
    // fingerprint calendar. The fix introduces a ReportContext with both
    // ends of the range and resolves year per month.
    const rows: unknown[][] = [];
    rows[1] = ['Report'];
    rows[2] = ['Dec 1, 2025 - Jan 31, 2026'];
    rows[6] = [null, 'Full name', 'Adi Misykatul'];
    rows[11] = [null, 'Schedule', 'Template', 'Clock-in', 'Clock-out', 'Worked', 'Late', 'Overtime (non approved)', 'Early departure', 'Worked on day off'];
    rows[12] = ['30 Dec, Tu', '08:00 - 16:30', null, '08:05', '17:00', 1, 0, 0, 0, 0];
    rows[13] = ['02 Jan, Fr', '08:00 - 16:30', null, '08:10', '17:05', 1, 0, 0, 0, 0];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const parsed = parseOnlineExcel(buffer);

    // The December row should land on 2025-12-30 (start year side of the range).
    expect(parsed.get('adi misykatul')?.get('2025-12-30')).toEqual({
      clockIn: '08:05',
      clockOut: '17:00',
    });
    // The January row should land on 2026-01-02 (end year side). Pre-fix,
    // this was keyed as '2025-01-02' and the fingerprint calendar would
    // never find a match.
    expect(parsed.get('adi misykatul')?.get('2026-01-02')).toEqual({
      clockIn: '08:10',
      clockOut: '17:05',
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

    // The online alias is the fingerprint name without the middle name. This
    // is the realistic shape of office data: fingerprint exports carry the
    // full name while online exports sometimes drop a middle name. The
    // match scores 100 (first=first, last=last, shared tokens) and merges
    // the earliest in (08:10 < 08:15) and latest out (18:14 > 17:00).
    const onlineData = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();
    onlineData.set(
      'adi anwar',
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

  it('throws when the fingerprint workbook has no clock-in or clock-out columns', () => {
    // Regression: parseFingerprintExcel used to silently produce a
    // workbook full of null clock times when the header row had no
    // recognisable clock-in/out columns. The user saw a zero-everywhere
    // report with no indication of the structural problem.
    const rows: unknown[][] = [];
    rows[0] = ['Emp No', 'Code', 'Division', 'Name', 'Position', 'Date', 'Working Hours'];
    rows[1] = ['427', 'E-427', 'IDACT', 'Adi Misykatul', 'Staff', '2026-03-05', 'Office Hour'];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    expect(() => parseFingerprintExcel(buffer)).toThrowError(/missing both "Actual In\/Out" and "Clock In\/Out" columns/);
  });

  it('reads the policy-priority clock-in column even when it has the lower column index', () => {
    // Regression: parseFingerprintExcel used Math.max(actualInIdx,
    // clockInIdx) to pick the column. The Math.max call is a
    // position-based choice, not a priority-based one. With "Actual
    // In" at col 4 and "Clock In" at col 5, Math.max(4, 5) = 5 read
    // from the "Clock In" column — the opposite of policy intent.
    // The fix uses pickPriorityIndex, which returns the first match
    // in priority order ('actual in' before 'clock in') regardless
    // of column position.
    // The parser uses hardcoded default column positions when the header
    // labels aren't found: Name -> col 3, Date -> col 5, Working Hours
    // -> col 6. The fixture mirrors that layout, with "Actual In" placed
    // BEFORE "Clock In" (lower column index) to trigger the Math.max bug.
    const rows: unknown[][] = [];
    rows[0] = ['Emp No', 'Code', 'Division', 'Name', 'Position', 'Date', 'Working Hours', 'Actual In', 'Clock In', 'Actual Out', 'Clock Out'];
    rows[1] = ['427', 'E-427', 'IDACT', 'Adi Misykatul', 'Staff', '2026-03-05', 'Office Hour', '08:00', '08:30', '17:00', '17:30'];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    const records = parseFingerprintExcel(buffer);

    expect(records).toHaveLength(1);
    // The policy says "Actual In" (08:00) wins over "Clock In" (08:30).
    // Pre-fix, Math.max(4, 5) = 5 picked "Clock In" and the record
    // carried 08:30.
    expect(records[0].actualIn).toBe('08:00');
    expect(records[0].clockIn).toBe('08:00');
    expect(records[0].actualOut).toBe('17:00');
    expect(records[0].clockOut).toBe('17:00');
  });

  it('does not let two fingerprint employees with a shared first name collide on the same online alias', () => {
    // Regression: `addAliases` used to register single-token aliases (e.g.
    // "adi"), so two distinct fingerprint employees ("Adi Wijaya" and
    // "Adi Saputra") both inherited the online clock-ins of whichever
    // online employee happened to register the "adi" alias first. The fix
    // drops single-token aliases from the store. With no single-token
    // alias available, a 2-token fingerprint name that shares only the
    // first token with a 3-token online name scores 40 (first only) which
    // is below the 60 minimum; the second employee ends up with no online
    // match and empty clock-ins.
    const fingerprintRecords: RawFingerprintRecord[] = [
      {
        empNo: '001',
        name: 'Adi Wijaya',
        date: '2025-10-01',
        dateKey: '2025-10-01',
        workingHours: 'Office Hour',
        clockIn: '08:00',
        clockOut: '17:00',
        actualIn: '08:00',
        actualOut: '17:00',
      },
      {
        empNo: '002',
        name: 'Adi Saputra',
        date: '2025-10-01',
        dateKey: '2025-10-01',
        workingHours: 'Office Hour',
        clockIn: '08:00',
        clockOut: '17:00',
        actualIn: '08:00',
        actualOut: '17:00',
      },
    ];

    const onlineData = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();
    onlineData.set(
      'adi pratama',
      new Map([
        ['2025-10-01', { clockIn: '08:30', clockOut: '17:30' }],
      ])
    );

    const compiled = compileAttendance(fingerprintRecords, onlineData);
    const wijaya = compiled.find((e) => e.name === 'Adi Wijaya');
    const saputra = compiled.find((e) => e.name === 'Adi Saputra');
    const wijayaDay = wijaya?.records.find((r) => r.date.getDate() === 1);
    const saputraDay = saputra?.records.find((r) => r.date.getDate() === 1);

    // Neither employee should inherit the online clock-ins of a different
    // online employee whose only overlap is the first name. The match
    // score (40, first-name only) is below the 60 minimum.
    expect(wijayaDay?.onlineIn).toBeNull();
    expect(saputraDay?.onlineIn).toBeNull();
  });

  it('omits the cached value on no-attendance formula cells so the formula result is shown', () => {
    // Regression: when a row had no actualIn/actualOut (e.g. a weekend
    // day, or a missed punch), the workbook used to ship a cached value
    // of 0 in I/J/K/L because toFormulaFraction(null) was 0. The
    // formula itself returns '' via IF(OR(G="",H="",...),""), but
    // viewers render the cached 0 until the user forces F9 to
    // recalculate. The fix is to emit `{ f, t: 'n', z }` with no `v`
    // when the cached value is null, so viewers display the formula's
    // own "" result.
    const compiled = compileAttendance(
      [
        // Saturday 4 Oct 2025: weekend, no attendance.
        {
          empNo: '427',
          name: 'Adi Misykatul Anwar',
          date: '2025-10-04',
          dateKey: '2025-10-04',
          workingHours: 'Office Hour',
          clockIn: null,
          clockOut: null,
          actualIn: null,
          actualOut: null,
        },
      ],
      new Map()
    );

    const workbook = buildAttendanceWorkbook(compiled);
    const sheet = workbook.Sheets['Adi'];
    // The first data row is at row 6 (row 1: title, row 2: period, row
    // 3: blank, row 4: headers, row 5: subheaders, row 6: first day).
    // The weekend row is the 4th record, so it's at row 6 + 3 = 9.
    const totalHoursCell = sheet?.['I9'];
    const leaveEarlierCell = sheet?.['K9'];

    // The formula is still present, the cell type is numeric, and
    // there is no cached numeric value pre-populated.
    expect(totalHoursCell?.f).toContain('H9-G9');
    expect(totalHoursCell?.t).toBe('n');
    expect(totalHoursCell?.v).toBeUndefined();
    expect(leaveEarlierCell?.f).toBeDefined();
    expect(leaveEarlierCell?.t).toBe('n');
    expect(leaveEarlierCell?.v).toBeUndefined();
  });

  it('throws when no date can be detected in either source file', () => {
    // Regression: getMonthFromDates used to fall back to a stale
    // { year: 2025, month: 9 } literal when no dates parsed, and the
    // workbook was built for that month with every cell empty. The fix
    // throws a descriptive error so the UI surfaces the problem
    // instead of silently producing an empty report.
    const fingerprintRecords: RawFingerprintRecord[] = [
      {
        empNo: '427',
        name: 'Adi Misykatul Anwar',
        // Both date and dateKey are non-parseable strings; the parser
        // accepts any of the formats in timeUtils.parseDate but these
        // don't match any of them.
        date: 'not-a-date',
        dateKey: 'still-not-a-date',
        workingHours: 'Office Hour',
        clockIn: '08:00',
        clockOut: '17:00',
        actualIn: '08:00',
        actualOut: '17:00',
      },
    ];
    const onlineData = new Map<string, Map<string, { clockIn: string | null; clockOut: string | null }>>();

    expect(() => compileAttendance(fingerprintRecords, onlineData)).toThrowError(
      /Could not detect reporting period from uploaded files/
    );
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
    expect(sheet?.['I6']?.f).toContain('WEEKDAY(A6,2)>5');
    expect(sheet?.['I6']?.f).toContain('MIN(H6,TIME(12,30,0))-MAX(G6,TIME(12,0,0))');
    expect(sheet?.['I6']?.z).toBe('[h]:mm');
    expect(sheet?.['K6']?.f).toContain('TIME(8,15,0)');
    expect(sheet?.['K6']?.f).toContain('TIME(8,0,0)');
    expect(sheet?.['K6']?.f).toContain('WEEKDAY(A6,2)>5');
    expect(sheet?.['K6']?.f).toContain('TIME(16,30,0)');
    expect(sheet?.['K6']?.f).toContain('TIME(17,30,0)');
    expect(sheet?.['K6']?.z).toBe('[h]:mm');
    expect(sheet?.['I9']?.f).toContain('WEEKDAY(A9,2)>5');
    expect(sheet?.['K9']?.f).toContain('WEEKDAY(A9,2)>5');
    expect(sheet?.['A6']?.v).not.toBe('Divisi : MITSUI OSK LINES');
    expect(sheet?.['A6']?.v).not.toBe('NIP : 000427   Nama : ADI MISYKATUL ANWAR');
  });

  it('keeps Cloudflare static workbook formulas aligned with source formulas', () => {
    const compiler = readFileSync('cloudflare-pages-attendance/public/browserCompiler.js', 'utf8');
    const policy = readFileSync('cloudflare-pages-attendance/public/policy.js', 'utf8');

    // The Cloudflare file loads the policy as a sibling <script> and
    // delegates every formula builder to it. The day-of-week strings
    // now come from cloudflareDayChecks in shared/policy.js (and its
    // sync'd copy at cloudflare-pages-attendance/public/policy.js), so
    // the test asserts the policy.js content is in place and the
    // browserCompiler.js wires the four cell-level formulas to the
    // policy builders.
    expect(policy).toContain('B${row}="Sat"');
    expect(policy).toContain('B${row}="Sun"');
    expect(policy).toContain('B${row}="Fri"');
    expect(policy).toContain('MIN(H${row},TIME(12,30,0))-MAX(G${row},TIME(12,0,0))');
    expect(policy).toContain('IF(G${row}<=TIME(8,0,0)');

    // The browserCompiler.js script must reference the AttendancePolicy
    // global for each of the four cell-level formula builders.
    expect(compiler).toContain('AttendancePolicy.buildTotalHoursFormula');
    expect(compiler).toContain('AttendancePolicy.buildTardinessFormula');
    expect(compiler).toContain('AttendancePolicy.buildLeaveEarlierFormula');
    expect(compiler).toContain('AttendancePolicy.buildOvertimeFormula');
    expect(compiler).toContain('AttendancePolicy.cloudflareDayChecks');
  });
});
