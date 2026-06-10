// Excel report generation

import * as XLSX from 'xlsx';
import type { CompiledEmployee, MergedAttendanceRecord } from './types';
import { dateToExcelSerial, formatDateFull, getDayName, parseTimeToMinutes } from './timeUtils';
import { isFriday, isWeekend } from './policy';
import {
  buildBreakFormula as policyBuildBreakFormula,
  buildTotalHoursFormula as policyBuildTotalHoursFormula,
  buildTardinessFormula as policyBuildTardinessFormula,
  buildLeaveEarlierFormula as policyBuildLeaveEarlierFormula,
  buildOvertimeFormula as policyBuildOvertimeFormula,
  reactDayExpr,
} from './policy';

const COLUMN_WIDTHS = [
  { wch: 10 },
  { wch: 6 },
  { wch: 8 },
  { wch: 18 },
  { wch: 14 },
  { wch: 8 },
  { wch: 12 },
  { wch: 12 },
  { wch: 13 },
  { wch: 11 },
  { wch: 13 },
  { wch: 11 },
  { wch: 10 },
];

const DATE_DISPLAY_FORMAT = 'dd/mm';
const TIME_DISPLAY_FORMAT = '[h]:mm';

function toFormulaFraction(value: string | null | undefined): number | null {
  const minutes = parseTimeToMinutes(value);
  if (minutes === null) return null;
  return minutes / (24 * 60);
}

// Formula builders. Thin wrappers over the canonical policy builders
// (shared/policy.js) that pre-supply the React day-of-week expression
// (WEEKDAY(A${row},2)=5). Pre-Phase 2 each formula was hand-written
// inline here, with the same Mon-Thu / Friday IF(WEEKDAY(...)=5) and
// IF(WEEKDAY(...)<=4) branching repeated. Centralising in policy.js
// means the Cloudflare bundle can produce byte-identical formulas
// from the same source.

function buildBreakFormula(row: number): string {
  return policyBuildBreakFormula(row, reactDayExpr(row));
}

function buildTotalHoursFormula(row: number): string {
  return policyBuildTotalHoursFormula(row, reactDayExpr(row));
}

function buildTardinessFormula(row: number): string {
  return policyBuildTardinessFormula(row, reactDayExpr(row));
}

function buildLeaveEarlierFormula(row: number): string {
  return policyBuildLeaveEarlierFormula(row, reactDayExpr(row));
}

function buildOvertimeFormula(row: number): string {
  return policyBuildOvertimeFormula(row, reactDayExpr(row));
}

function makeFormulaCell(formula: string, cachedValue: number | string | null, numberFormat?: string): XLSX.CellObject {
  if (cachedValue === null || cachedValue === undefined || cachedValue === '') {
    // No cached value: emit the formula with `t: 'n'` and the number format
    // but no `v`. The formula's own evaluation yields '' for rows with no
    // clock-in/out (the IF guard returns ""), and the absent `v` prevents
    // SheetJS from pre-populating the cell with a misleading 0. Some
    // viewers (and the round-trip via XLSX.write) interpret a missing `t`
    // inconsistently; explicitly setting `t: 'n'` makes the cell type
    // stable. Pre-fix, the cached value was 0 on no-attendance rows
    // because toFormulaFraction(null) returned 0, and the workbook
    // displayed 0:00 in the I/J/K/L columns until the user forced F9.
    return numberFormat ? { f: formula, z: numberFormat, t: 'n' } : { f: formula, t: 'n' };
  }

  return typeof cachedValue === 'number'
    ? { f: formula, v: cachedValue, t: 'n', z: numberFormat }
    : { f: formula, v: cachedValue, t: 's', z: numberFormat };
}

function buildRowMetadata(record: MergedAttendanceRecord): { shift: string; officeIn: string; officeOut: string } {
  if (isWeekend(record.date)) {
    return { shift: '', officeIn: '', officeOut: '' };
  }

  if (isFriday(record.date)) {
    return { shift: '08.00 - 17.00', officeIn: 'C 08:00', officeOut: 'C 17:00' };
  }

  return { shift: '08.00 - 16.30', officeIn: 'C 08:00', officeOut: 'C 16:30' };
}

function buildSheetData(
  employee: CompiledEmployee,
  records: MergedAttendanceRecord[],
  periodLabel: string
): XLSX.WorkSheet {
  const wsData: (string | number | null)[][] = [];

  wsData.push(['Laporan Absensi Harian', '', '', '', '', '', '', '', '', '', '', '', '']);
  wsData.push([`Periode ${periodLabel}`, '', '', '', '', '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
  wsData.push(['Date', 'Day', 'Kal', 'Shift', 'Office Hours', '', 'Actual In', 'Actual Out', 'Total Hours', 'Tardiness', 'Leave Earlier', 'Overtime', 'Remarks']);
  wsData.push(['', '', '', '', 'In', 'Out', '', '', '', '', '', '', '']);

  for (const record of records) {
    const rowMeta = buildRowMetadata(record);
    wsData.push([
      dateToExcelSerial(record.date),
      getDayName(record.date),
      'WD',
      rowMeta.shift,
      rowMeta.officeIn,
      rowMeta.officeOut,
      record.actualIn || '',
      record.actualOut || '',
      '',
      '',
      '',
      '',
      record.remarks || '',
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = COLUMN_WIDTHS;
  ws['!merges'] = [
    { s: { c: 0, r: 0 }, e: { c: 12, r: 0 } },
    { s: { c: 0, r: 1 }, e: { c: 12, r: 1 } },
    { s: { c: 0, r: 3 }, e: { c: 0, r: 4 } },
    { s: { c: 1, r: 3 }, e: { c: 1, r: 4 } },
    { s: { c: 2, r: 3 }, e: { c: 2, r: 4 } },
    { s: { c: 3, r: 3 }, e: { c: 3, r: 4 } },
    { s: { c: 4, r: 3 }, e: { c: 5, r: 3 } },
    { s: { c: 6, r: 3 }, e: { c: 6, r: 4 } },
    { s: { c: 7, r: 3 }, e: { c: 7, r: 4 } },
    { s: { c: 8, r: 3 }, e: { c: 8, r: 4 } },
    { s: { c: 9, r: 3 }, e: { c: 9, r: 4 } },
    { s: { c: 10, r: 3 }, e: { c: 10, r: 4 } },
    { s: { c: 11, r: 3 }, e: { c: 11, r: 4 } },
    { s: { c: 12, r: 3 }, e: { c: 12, r: 4 } },
  ];
  const dataStartRow = 6;

  records.forEach((record, index) => {
    const rowNumber = dataStartRow + index;
    const totalHoursCell = `I${rowNumber}`;
    const tardinessCell = `J${rowNumber}`;
    const leaveEarlierCell = `K${rowNumber}`;
    const overtimeCell = `L${rowNumber}`;

    const totalHoursValue = toFormulaFraction(record.totalHours);
    const tardinessValue = toFormulaFraction(record.tardiness);
    const leaveEarlierValue = toFormulaFraction(record.leaveEarlier);
    const overtimeValue = toFormulaFraction(record.overtime);

    ws[totalHoursCell] = makeFormulaCell(buildTotalHoursFormula(rowNumber), totalHoursValue, TIME_DISPLAY_FORMAT);
    ws[tardinessCell] = makeFormulaCell(buildTardinessFormula(rowNumber), tardinessValue, TIME_DISPLAY_FORMAT);
    ws[leaveEarlierCell] = makeFormulaCell(buildLeaveEarlierFormula(rowNumber), leaveEarlierValue, TIME_DISPLAY_FORMAT);
    ws[overtimeCell] = makeFormulaCell(buildOvertimeFormula(rowNumber), overtimeValue, TIME_DISPLAY_FORMAT);

    const dateCell = `A${rowNumber}`;
    const dayCell = `B${rowNumber}`;
    ws[dateCell] = { t: 'n', v: dateToExcelSerial(record.date), z: DATE_DISPLAY_FORMAT };
    ws[dayCell] = { f: `TEXT(A${rowNumber},"ddd")`, v: getDayName(record.date), t: 's' };
  });

  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:M1');
  range.e.r = Math.max(range.e.r, dataStartRow + records.length - 1);
  range.e.c = 12;
  ws['!ref'] = XLSX.utils.encode_range(range);

  return ws;
}

/**
 * Build the compiled attendance workbook.
 */
export function buildAttendanceWorkbook(employees: CompiledEmployee[]): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new();

  if (employees.length === 0) {
    const emptySheet = XLSX.utils.aoa_to_sheet([['No data']]);
    XLSX.utils.book_append_sheet(workbook, emptySheet, 'Template');
    return workbook;
  }

  const firstEmployee = employees[0];
  const firstRecord = firstEmployee.records[0];
  const lastRecord = firstEmployee.records[firstEmployee.records.length - 1];
  const periodLabel = `${formatDateFull(firstRecord.date)} s/d  ${formatDateFull(lastRecord.date)}`;

  const templateSheet = buildSheetData(firstEmployee, firstEmployee.records, periodLabel);
  XLSX.utils.book_append_sheet(workbook, templateSheet, 'Template');

  for (const employee of employees) {
    const sheet = buildSheetData(employee, employee.records, periodLabel);
    XLSX.utils.book_append_sheet(workbook, sheet, employee.sheetName);
  }

  return workbook;
}

/**
 * Generate the compiled attendance Excel file.
 */
export function generateAttendanceExcel(employees: CompiledEmployee[]): Blob {
  const workbook = buildAttendanceWorkbook(employees);
  const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/**
 * Download the Excel file.
 */
export function downloadExcel(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
