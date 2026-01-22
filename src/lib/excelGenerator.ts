// Excel report generation

import * as XLSX from 'xlsx';
import type { CompiledEmployee } from './types';
import { formatDateShort } from './timeUtils';

/**
 * Generate the compiled attendance Excel file
 */
export function generateAttendanceExcel(employees: CompiledEmployee[]): Blob {
  const workbook = XLSX.utils.book_new();

  for (const employee of employees) {
    // Create worksheet data
    const wsData: (string | number | null)[][] = [];

    // Header section
    wsData.push(['Laporan Absensi Harian', '', '', '', '', '', '', '', '']);
    wsData.push([`Periode 01/10/2025 s/d 31/10/2025`, '', '', '', '', '', '', '', '']);
    wsData.push(['', '', '', '', '', '', '', '', '']);
    wsData.push(['Date', 'Day', 'Actual In', 'Actual Out', 'Total Hours', 'Tardiness', 'Leave Earlier', 'Overtime', 'Remarks']);
    wsData.push(['', '', '', '', '', '', '', '', '']);
    
    // Employee info
    wsData.push([`Divisi : ${employee.division}`, '', '', '', '', '', '', '', '']);
    wsData.push([`Departemen : ${employee.department}`, '', '', '', '', '', '', '', '']);
    wsData.push([`Seksi : ${employee.section}`, '', '', '', '', '', '', '', '']);
    wsData.push([`NIP : ${employee.nip} Nama : ${employee.name}`, '', '', '', '', '', '', '', '']);

    // Data rows
    for (let i = 0; i < employee.records.length; i++) {
      const record = employee.records[i];
      const rowNum = i + 10; // Starting row number (1-indexed, after headers)
      
      wsData.push([
        formatDateShort(record.date),
        record.dayOfWeek,
        record.actualIn || '',
        record.actualOut || '',
        record.totalHours || '0:00',
        record.tardiness || '',
        record.leaveEarlier || '',
        record.overtime || '',
        record.remarks || '0',
      ]);
    }

    // Create the worksheet
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Set column widths
    ws['!cols'] = [
      { wch: 10 }, // Date
      { wch: 6 },  // Day
      { wch: 12 }, // Actual In
      { wch: 12 }, // Actual Out
      { wch: 12 }, // Total Hours
      { wch: 10 }, // Tardiness
      { wch: 12 }, // Leave Earlier
      { wch: 10 }, // Overtime
      { wch: 15 }, // Remarks
    ];

    // Add formulas for Total Hours calculation
    // Note: XLSX doesn't support complex time calculations easily,
    // so we're using pre-calculated values but structuring for future formula support
    
    // Sanitize sheet name (Excel has a 31 char limit and no special chars)
    const sheetName = employee.name
      .replace(/[\\/*?[\]:]/g, '')
      .substring(0, 31);

    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  }

  // Generate the file
  const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  return new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/**
 * Download the Excel file
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
