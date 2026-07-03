import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { calculateAttendance } from '@/lib/attendanceCalculator';
import { compileAttendance } from '@/lib/attendanceCompiler';
import { buildAttendanceWorkbook } from '@/lib/excelGenerator';
import { parseFingerprintExcel, parseOnlineExcel } from '@/lib/excelParser';
import { extractTime, parseTimeToMinutes } from '@/lib/timeUtils';
import { compileWithCloudflare } from './cloudflare-harness';
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

  it('Cloudflare harness loads the static bundle in jsdom and produces a populated workbook', async () => {
    // Smoke test for the Cloudflare test harness. The Cloudflare
    // parser has its own column expectations: fingerprint name is in
    // col 4, date col 6, actual in/out cols 10/11; the online
    // workbook uses a "Full name" marker in column B with day labels
    // in column A and times in columns D/E. We build synthetic
    // workbooks matching that shape and assert the captured stub
    // workbook has the expected sheets, cells, and formula strings.
    // SheetJS's aoa_to_sheet collapses leading null/empty rows, so
    // every "row" in the AOA must contain at least one cell value.
    // We use a single empty-string cell in A1 and A2 to act as a
    // placeholder and preserve the data positions the Cloudflare
    // parser expects.
    const fingerprintRows: unknown[][] = [
      [''],                                                                  // row 1: skipped by the Cloudflare parser
      ['', '', '', 'Adi Misykatul', '', '2026-03-03', '', '', '', '08:00', '17:00'],
      ['', '', '', 'Adi Misykatul', '', '2026-03-04', '', '', '', '08:05', '17:30'],
    ];
    const fingerprintWorkbook = XLSX.utils.book_new();
    const fingerprintWorksheet = XLSX.utils.aoa_to_sheet(fingerprintRows);
    XLSX.utils.book_append_sheet(fingerprintWorkbook, fingerprintWorksheet, 'Sheet1');
    const fingerprintBuffer = XLSX.write(fingerprintWorkbook, { bookType: 'xlsx', type: 'array' });

    // The Cloudflare online parser requires:
    //   - Cell A3 (row 3 col 1) = the report period.
    //   - A "Full name" marker in column B (col 2) of the employee
    //     section.
    //   - A "Schedule" label in column B of the next row.
    //   - Day labels in column A matching /^(\d{2})\s+([A-Za-z]{3}),/.
    //   - Times in columns D (col 4) and E (col 5).
    const onlineRows: unknown[][] = [
      [''],                                                                    // row 1
      [''],                                                                    // row 2
      ['Mar 1, 2026 - Mar 31, 2026'],                                          // row 3: report period
      [''],                                                                    // row 4
      [''],                                                                    // row 5
      [null, 'Full name', 'Adi Misykatul'],                                    // row 6: name marker
      [null, 'Schedule', 'Template', 'Clock-in', 'Clock-out'],                  // row 7: schedule header
      ['03 Mar, Tu', '08:00 - 16:30', null, '08:10', '18:10'],                  // row 8: day
      ['04 Mar, We', '08:00 - 16:30', null, '08:15', '18:15'],                  // row 9: day
    ];
    const onlineWorkbook = XLSX.utils.book_new();
    const onlineWorksheet = XLSX.utils.aoa_to_sheet(onlineRows);
    XLSX.utils.book_append_sheet(onlineWorkbook, onlineWorksheet, 'Sheet1');
    const onlineBuffer = XLSX.write(onlineWorkbook, { bookType: 'xlsx', type: 'array' });

    const result = await compileWithCloudflare(fingerprintBuffer, onlineBuffer);

    // The Cloudflare bundle returns the workbook (with all sheets
    // populated), a fileName, warnings, and a summary.
    expect(result).toBeDefined();
    expect(result.workbook).toBeDefined();
    expect(result.fileName).toMatch(/^Compiled Attendance/);
    expect(result.summary.employees).toBeGreaterThan(0);

    // The captured sheets include the Template sheet and one per
    // employee. The matched employee is "Adi Misykatul" so we expect
    // a sheet named "Adi" (first-name) or "Adi Misykatul" (truncated).
    const sheetNames = result.workbook.worksheets.map((s: { name: string }) => s.name);
    expect(sheetNames).toContain('Template');
    const employeeSheet = result.workbook.worksheets.find(
      (s: { name: string }) => s.name !== 'Template'
    );
    expect(employeeSheet).toBeDefined();

    // The per-employee sheet has the layout from styleSheet:
    // A1 = title, A2 = period, A4..M4 = headers, A6 = "Nama : ..."
    // The data rows start at row 7.
    const titleCell = employeeSheet.getCell('A1');
    expect(titleCell.value).toBe('Laporan Absensi Harian');
    const periodCell = employeeSheet.getCell('A2');
    expect(periodCell.value).toMatch(/^Periode 01\/03\/2026 s\/d 31\/03\/2026$/);
    const nameCell = employeeSheet.getCell('A6');
    expect(nameCell.value).toBe('Nama : Adi Misykatul');

    // The Cloudflare compiler iterates over every day in the detected
    // month and writes one row per day starting at row 7. Day 1 is
    // 2026-03-01 (Sunday), day 2 is 2026-03-02 (Monday), day 3 is
    // 2026-03-03 (Tuesday), and so on. The fingerprint and online
    // data only contain records for March 3 and March 4 — so the
    // attendance columns for rows 7-8 (March 1, March 2) should be
    // null, and the values for row 9 (March 3) should be the merged
    // 08:00 / 18:10 from the test data.
    const firstDataRow = employeeSheet.getRow(7);
    expect(firstDataRow.getCell(2).value).toBe('Sun'); // B7 = day name for March 1
    // Rows 7 and 8 have no attendance. G7 and H7 should be null.
    expect(firstDataRow.getCell(7).value).toBeNull();
    expect(firstDataRow.getCell(8).value).toBeNull();

    const march3Row = employeeSheet.getRow(9);
    expect(march3Row.getCell(2).value).toBe('Tue'); // B9 = day name for March 3
    // G9 is the earliest of fingerprint (08:00) and online (08:10)
    // = 08:00 = 8/24.
    expect(march3Row.getCell(7).value).toBeCloseTo(8 / 24, 6);
    // H9 is the latest of fingerprint (17:00) and online (18:10)
    // = 18:10 = (18*60+10)/(24*60).
    expect(march3Row.getCell(8).value).toBeCloseTo((18 * 60 + 10) / (24 * 60), 6);

    // The formula cells (I-L) are ExcelJS formula objects with the
    // formula produced by the policy builders. The day-name check
    // uses the cloudflareDayChecks (B="Sat", B="Sun", B="Fri")
    // convention.
    const totalHours = march3Row.getCell(9);
    const tardiness = march3Row.getCell(10);
    const leaveEarlier = march3Row.getCell(11);
    const overtime = march3Row.getCell(12);
    expect(typeof totalHours.value).toBe('object');
    expect((totalHours.value as { formula: string }).formula).toContain('OR(B9="Sat",B9="Sun")');
    expect((totalHours.value as { formula: string }).formula).toContain('MAX(0,(H9-G9)');
    expect((tardiness.value as { formula: string }).formula).toContain('OR(B9="Sat",B9="Sun")');
    expect((leaveEarlier.value as { formula: string }).formula).toContain('B9="Fri"');
    expect((overtime.value as { formula: string }).formula).toContain('IF(B9="Fri",TIME(18,0,0),TIME(17,30,0))');
  });

  // Behavioural parity: drive the React path and the Cloudflare
  // path with the same logical attendance data and verify that the
  // resulting sheets agree on the things an HR operator would notice
  // (per-row date, day, merged actual in/out) and on the shape of
  // the formula cells (each formula string contains the right
  // policy constants — exact byte match is not expected because
  // the React path uses WEEKDAY(A,2)=5 and the Cloudflare path
  // uses B="Sat"/"Sun"/"Fri").
  //
  // Note on sheet name divergence: the React path uses the first
  // space-separated token (e.g. "Adi Misykatul Anwar" -> "Adi")
  // while the Cloudflare path uses the full displayName. To make
  // both paths produce the same sheet name, the test uses a
  // single-token employee name ("Adi") — both paths produce a
  // sheet called "Adi".
  it('React path and Cloudflare path produce equivalent output for the same inputs', async () => {
    const records = [
      { isoDate: '2026-03-02', in: '08:05', out: '17:30', weekday: 'Mon' },
      { isoDate: '2026-03-03', in: '08:10', out: '18:00', weekday: 'Tue' },
      { isoDate: '2026-03-04', in: '08:15', out: '17:30', weekday: 'Wed' },
      { isoDate: '2026-03-06', in: '08:00', out: '17:15', weekday: 'Fri' },
    ];
    const employeeName = 'Adi';

    // React-format fingerprint: header row + data rows. The React
    // parser uses Math.max(findIndex('name'), 3) to pick the name
    // column, so a "Name" header at col 1 is overridden to col 3.
    // We pick a layout that matches the parser's expectation: the
    // employee name sits at col 3 (the default fallback position).
    const reactFpRows: unknown[][] = [
      ['Emp No', 'Code', 'Division', 'Name', 'Position', 'Date', 'Working Hours', 'Actual In', 'Clock In', 'Actual Out', 'Clock Out'],
      ...records.map((r) => ['E-427', 'C-427', 'IDACT', employeeName, 'Staff', r.isoDate, 'Office Hour', r.in, r.in, r.out, r.out]),
    ];
    const reactFpWb = XLSX.utils.book_new();
    const reactFpWs = XLSX.utils.aoa_to_sheet(reactFpRows);
    XLSX.utils.book_append_sheet(reactFpWb, reactFpWs, 'Sheet1');
    const reactFpBuffer = XLSX.write(reactFpWb, { bookType: 'xlsx', type: 'array' });

    // Cloudflare-format fingerprint: hardcoded column positions
    // (col 4 = name, col 6 = date, col 10 = actual in, col 11 = actual out).
    // Row 1 is skipped by the Cloudflare parser.
    const cfFpRows: unknown[][] = [
      [''],
      ...records.map((r) => ['', '', '', employeeName, '', r.isoDate, '', '', '', r.in, r.out]),
    ];
    const cfFpWb = XLSX.utils.book_new();
    const cfFpWs = XLSX.utils.aoa_to_sheet(cfFpRows);
    XLSX.utils.book_append_sheet(cfFpWb, cfFpWs, 'Sheet1');
    const cfFpBuffer = XLSX.write(cfFpWb, { bookType: 'xlsx', type: 'array' });

    // Shared online block-format buffer. Both the React and the
    // Cloudflare parsers expect this exact shape (a "Full name"
    // marker in column B, a "Schedule" header, then day rows).
    const onlineRows: unknown[][] = [
      [''],
      [''],
      ['Mar 1, 2026 - Mar 31, 2026'],
      [''],
      [''],
      [null, 'Full name', employeeName],
      [null, 'Schedule', 'Template', 'Clock-in', 'Clock-out'],
      ...records.map((r, i) => {
        const dayNumber = String(i + 2).padStart(2, '0');
        // Map the isoDate month component ("03") to the 3-letter
        // abbreviation the Cloudflare parser's parseOnlineDateLabel
        // regex expects. Only March is exercised by these tests.
        const monthLabel: Record<string, string> = {
          '01': 'Jan', '02': 'Feb', '03': 'Mar',
          '04': 'Apr', '05': 'May', '06': 'Jun',
          '07': 'Jul', '08': 'Aug', '09': 'Sep',
          '10': 'Oct', '11': 'Nov', '12': 'Dec',
        };
        const month = r.isoDate.slice(5, 7);
        const dateLabelMap: Record<string, string> = {
          Mon: 'Mo', Tue: 'Tu', Wed: 'We', Thu: 'Th', Fri: 'Fr',
        };
        return [`${dayNumber} ${monthLabel[month]}, ${dateLabelMap[r.weekday]}`, '08:00 - 16:30', null, r.in, '18:00'];
      }),
    ];
    const onlineWb = XLSX.utils.book_new();
    const onlineWs = XLSX.utils.aoa_to_sheet(onlineRows);
    XLSX.utils.book_append_sheet(onlineWb, onlineWs, 'Sheet1');
    const onlineBuffer = XLSX.write(onlineWb, { bookType: 'xlsx', type: 'array' });

    // React path.
    const reactFpRecords = parseFingerprintExcel(reactFpBuffer);
    const reactOnlineMap = parseOnlineExcel(onlineBuffer);
    const reactCompiled = compileAttendance(reactFpRecords, reactOnlineMap);
    const reactWorkbook = buildAttendanceWorkbook(reactCompiled);
    const reactSheet = reactWorkbook.Sheets['Adi'];

    // Cloudflare path.
    const cfResult = await compileWithCloudflare(cfFpBuffer, onlineBuffer);
    const cfEmployeeSheet = cfResult.workbook.worksheets.find(
      (s: { name: string }) => s.name !== 'Template'
    );

    // Both paths should produce a sheet named after the employee.
    // With a single-token name like "Adi", the React path's
    // first-name sheet-name logic and the Cloudflare path's
    // full-name sheet-name logic both yield "Adi".
    expect(reactWorkbook.SheetNames).toContain(employeeName);
    expect(cfEmployeeSheet?.name).toBe(employeeName);

    // The Cloudflare path's styleSheet writes the employee name in
    // A6 ("Nama : <name>"). The React path's buildSheetData leaves
    // A6 empty (a blank separator row) — so we only assert the
    // Cloudflare cell here. Note this as a known shape difference
    // between the two implementations.
    expect(cfEmployeeSheet?.getCell('A6').value).toBe(`Nama : ${employeeName}`);

    // The React A2 cell carries the period label. The React path
    // uses a double space ("s/d  ") while the Cloudflare path
    // uses a single space ("s/d "). This is a known cosmetic
    // inconsistency between the two implementations.
    expect(reactSheet?.['A2']?.v).toBe('Periode 01/03/2026 s/d  31/03/2026');

    // The total data row count is the same — both paths iterate
    // every day in March (31 days) starting at row 7.
    expect(reactCompiled[0].records.length).toBe(31);

    // Sample the non-formula cells for the day that has attendance
    // (March 4 = row 10 in the merged timeline). We compare
    // data cells (A, B, G, H) but not the formula cells verbatim.
    const targetRow = 10; // March 4 (rows 7-9 are March 1-3, row 10 is March 4)
    const reactTarget = reactSheet?.[`A${targetRow}`];
    const cfTarget = cfEmployeeSheet?.getRow(targetRow);

    // The React cell carries the date as an Excel serial number
    // formatted dd/mm. The Cloudflare cell carries a JS Date. We
    // normalise both to YYYY-MM-DD via the dateKey or via the
    // dateToExcelSerial / Date.UTC conversion.
    // The compiled.records are 0-indexed; row 7 is March 1, row 10 is March 4.
    const compiledRecord = reactCompiled[0].records.find((r) => r.date.getUTCDate() === 4);
    expect(compiledRecord?.date.toISOString()).toContain('2026-03-04');

    // Day name in column B.
    expect(cfTarget?.getCell(2).value).toBe('Wed');
    // The React side stores the day name as a static formula
    // referencing the date; we don't compare formulas here, but we
    // verify the merged-in/merged-out values are equivalent.
    const mergedIn = compiledRecord?.actualIn;   // earliest between fp and online
    const mergedOut = compiledRecord?.actualOut; // latest between fp and online
    expect(mergedIn).toBe('08:00');  // 08:00 < 08:15
    expect(mergedOut).toBe('18:00'); // 18:00 > 17:30

    // The Cloudflare cell stores the same minutes as an Excel time
    // fraction (day.mergedIn / 1440). For March 4 both the
    // fingerprint and the online are '08:15', so cfIn = 495/1440.
    const targetRecord = records.find((r) => r.isoDate === '2026-03-04');
    const expectedInMinutes = Number(targetRecord!.in.slice(0, 2)) * 60 + Number(targetRecord!.in.slice(3, 5));
    const cfIn = cfTarget?.getCell(7).value as number;
    expect(cfIn).toBeCloseTo(expectedInMinutes / 1440, 4);

    // Structural check on the four formula cells in row 10 (March 4,
    // Wednesday). Each cell should reference its row and the
    // policy constant appropriate to that formula. We don't
    // compare exact strings because the two paths use different
    // day-name conventions.
    const cell9 = cfTarget?.getCell(9);
    const cell10 = cfTarget?.getCell(10);
    const cell11 = cfTarget?.getCell(11);
    const cell12 = cfTarget?.getCell(12);

    const formula9 = (cell9?.value as { formula: string }).formula;
    expect(formula9).toContain(`H${targetRow}`);
    expect(formula9).toContain(`G${targetRow}`);
    expect(formula9).toContain('TIME(12,30,0)');  // Mon-Thu break
    expect(formula9).toContain('TIME(11,30,0)');  // Fri break

    const formula10 = (cell10?.value as { formula: string }).formula;
    expect(formula10).toContain(`G${targetRow}`);
    expect(formula10).toContain('TIME(8,30,0)');   // 08:30 late threshold

    const formula11 = (cell11?.value as { formula: string }).formula;
    expect(formula11).toContain(`H${targetRow}`);
    expect(formula11).toContain(`B${targetRow}="Fri"`);
    expect(formula11).toContain('TIME(8,0,0)');    // 08:00 flexi threshold

    const formula12 = (cell12?.value as { formula: string }).formula;
    expect(formula12).toContain(`H${targetRow}`);
    expect(formula12).toContain(`B${targetRow}="Fri"`);
    expect(formula12).toContain('TIME(17,30,0)');  // Mon-Thu overtime
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
