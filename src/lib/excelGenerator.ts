// Excel report generation

import * as XLSX from 'xlsx';
import type { CompiledEmployee, MergedAttendanceRecord } from './types';
import { dateToExcelSerial, formatDateFull, getDayName, isFriday, isWeekend, parseTimeToMinutes } from './timeUtils';

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

function toFormulaFraction(value: string | null | undefined): number | null {
  const minutes = parseTimeToMinutes(value);
  if (minutes === null) return null;
  return minutes / (24 * 60);
}

function buildBreakFormula(row: number): string {
  const day = `WEEKDAY(A${row},2)`;
  return `IF(${day}=5,IF(AND(G${row}<TIME(13,0,0),H${row}>TIME(11,30,0)),MIN(H${row},TIME(13,0,0))-MAX(G${row},TIME(11,30,0)),0),IF(${day}<=4,IF(AND(G${row}<TIME(12,30,0),H${row}>TIME(12,0,0)),MIN(H${row},TIME(12,30,0))-MAX(G${row},TIME(12,0,0)),0),0))`;
}

function buildTotalHoursFormula(row: number): string {
  return `=IF(OR(G${row}="",H${row}=""),"",MAX(0,(H${row}-G${row})-(${buildBreakFormula(row)})))`;
}

function buildTardinessFormula(row: number): string {
  return `=IF(G${row}="","",MAX(0,G${row}-TIME(8,30,0)))`;
}

function buildLeaveEarlierFormula(row: number): string {
  return `=IF(OR(G${row}="",H${row}=""),"",IF(G${row}<TIME(8,15,0),IF(WEEKDAY(A${row},2)=5,MAX(0,TIME(17,15,0)-H${row}),MAX(0,TIME(16,45,0)-H${row})),IF(G${row}<TIME(8,30,0),IF(WEEKDAY(A${row},2)=5,MAX(0,TIME(17,30,0)-H${row}),MAX(0,TIME(17,0,0)-H${row})),0)))`;
}

function buildOvertimeFormula(row: number): string {
  return `=IF(H${row}="","",MAX(0,H${row}-IF(WEEKDAY(A${row},2)=5,TIME(18,0,0),TIME(17,30,0))))`;
}

function makeFormulaCell(formula: string, cachedValue: number | string | null): XLSX.CellObject {
  if (cachedValue === null || cachedValue === undefined || cachedValue === '') {
    return { f: formula };
  }

  return typeof cachedValue === 'number'
    ? { f: formula, v: cachedValue, t: 'n' }
    : { f: formula, v: cachedValue, t: 's' };
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
  periodLabel: string,
  includeEmployeeInfo: boolean
): XLSX.WorkSheet {
  const wsData: (string | number | null)[][] = [];

  wsData.push(['Laporan Absensi Harian', '', '', '', '', '', '', '', '', '', '', '', '']);
  wsData.push([`Periode ${periodLabel}`, '', '', '', '', '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
  wsData.push(['Date', 'Day', 'Kal', 'Shift', 'Office Hours', '', 'Actual In', 'Actual Out', 'Total Hours', 'Tardiness', 'Leave Earlier', 'Overtime', 'Remarks']);
  wsData.push(['', '', '', '', 'In', 'Out', '', '', '', '', '', '', '']);

  if (includeEmployeeInfo) {
    wsData.push([`Divisi : ${employee.division}`, '', '', '', '', '', '', '', '', '', '', '', '']);
    wsData.push([`Departemen : ${employee.department}`, '', '', '', '', '', '', '', '', '', '', '', '']);
    wsData.push([`Seksi : ${employee.section}`, '', '', '', '', '', '', '', '', '', '', '', '']);
    wsData.push([`NIP : ${employee.nip}   Nama : ${employee.name}`, '', '', '', '', '', '', '', '', '', '', '', '']);
  } else {
    wsData.push(['Divisi : MITSUI OSK LINES', '', '', '', '', '', '', '', '', '', '', '', '']);
    wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
    wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
    wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
  }

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
    { s: { c: 0, r: 5 }, e: { c: 12, r: 5 } },
    { s: { c: 0, r: 6 }, e: { c: 12, r: 6 } },
    { s: { c: 0, r: 7 }, e: { c: 12, r: 7 } },
    { s: { c: 0, r: 8 }, e: { c: 12, r: 8 } },
  ];
  const dataStartRow = 10;

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

    ws[totalHoursCell] = makeFormulaCell(buildTotalHoursFormula(rowNumber), totalHoursValue);
    ws[tardinessCell] = makeFormulaCell(buildTardinessFormula(rowNumber), tardinessValue);
    ws[leaveEarlierCell] = makeFormulaCell(buildLeaveEarlierFormula(rowNumber), leaveEarlierValue);
    ws[overtimeCell] = makeFormulaCell(buildOvertimeFormula(rowNumber), overtimeValue);

    const dateCell = `A${rowNumber}`;
    const dayCell = `B${rowNumber}`;
    ws[dateCell] = { t: 'n', v: dateToExcelSerial(record.date) };
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

  const templateSheet = buildSheetData(firstEmployee, firstEmployee.records, periodLabel, false);
  XLSX.utils.book_append_sheet(workbook, templateSheet, 'Template');

  for (const employee of employees) {
    const sheet = buildSheetData(employee, employee.records, periodLabel, true);
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
