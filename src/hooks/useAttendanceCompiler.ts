import { useState, useCallback } from 'react';
import type { CompiledEmployee } from '@/lib/types';
import { parseFingerprintExcel, parseOnlineExcel } from '@/lib/excelParser';
import { compileAttendance } from '@/lib/attendanceCompiler';
import { generateAttendanceExcel, downloadExcel } from '@/lib/excelGenerator';

// First four bytes of any ZIP archive: 0x50 0x4B 0x03 0x04 ("PK\x03\x04").
// An xlsx file is a ZIP container; if the input doesn't have this
// header the upload is not a real Excel file (CSV, plain text, or
// other binary). Sniffing before handing the buffer to SheetJS gives
// a friendlier error than the library's "Invalid signature".
const XLSX_MAGIC = [0x50, 0x4b, 0x03, 0x04];

function assertXlsxBuffer(buffer: ArrayBuffer, label: string): void {
  if (buffer.byteLength < 4) {
    throw new Error(`${label} is empty or unreadable. Please upload a valid .xlsx file.`);
  }
  const view = new Uint8Array(buffer, 0, 4);
  for (let i = 0; i < 4; i++) {
    if (view[i] !== XLSX_MAGIC[i]) {
      throw new Error(
        `${label} is not a valid Excel file (missing ZIP signature). Please upload a .xlsx file exported from your attendance system.`
      );
    }
  }
}

export function useAttendanceCompiler() {
  const [fingerprintFile, setFingerprintFile] = useState<File | null>(null);
  const [onlineFile, setOnlineFile] = useState<File | null>(null);
  const [compiledData, setCompiledData] = useState<CompiledEmployee[]>([]);
  const [isCompiling, setIsCompiling] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const canCompile = fingerprintFile !== null && onlineFile !== null;
  const hasData = compiledData.length > 0;

  const compile = useCallback(async () => {
    if (!fingerprintFile || !onlineFile) {
      setError('Please upload both Fingerprint and Online Excel files');
      return;
    }

    setIsCompiling(true);
    setError(null);

    try {
      // Read files
      const [fingerprintBuffer, onlineBuffer] = await Promise.all([
        fingerprintFile.arrayBuffer(),
        onlineFile.arrayBuffer(),
      ]);

      // ZIP magic bytes: "PK\x03\x04". An xlsx file is a ZIP
      // container; if either input doesn't start with these four
      // bytes, the file is not a real Excel file (CSV, plain text,
      // binary garbage) and SheetJS will throw a confusing "Invalid
      // signature" error. Surface a friendlier message up front.
      assertXlsxBuffer(fingerprintBuffer, 'Fingerprint file');
      assertXlsxBuffer(onlineBuffer, 'Online file');

      // Parse files
      const fingerprintRecords = parseFingerprintExcel(fingerprintBuffer);
      const onlineData = parseOnlineExcel(onlineBuffer);

      // Compile attendance
      const compiled = compileAttendance(fingerprintRecords, onlineData);

      if (compiled.length === 0) {
        setError('No employee data found in the uploaded files');
        return;
      }

      setCompiledData(compiled);
    } catch (err) {
      console.error('Compilation error:', err);
      setError(err instanceof Error ? err.message : 'An error occurred during compilation');
      // Drop any previously compiled data so the UI doesn't show
      // stale employees while the red Alert banner is up. Without
      // this, hitting "Compile" with a malformed file would keep
      // showing the prior good run in the employee tabs while the
      // error message was dismissed — misleading.
      setCompiledData([]);
    } finally {
      setIsCompiling(false);
    }
  }, [fingerprintFile, onlineFile]);

  const downloadReport = useCallback(() => {
    if (compiledData.length === 0) return;

    const blob = generateAttendanceExcel(compiledData);
    const firstDate = compiledData[0]?.records[0]?.date;
    const periodTag = firstDate
      ? `${firstDate.getFullYear()}-${String(firstDate.getMonth() + 1).padStart(2, '0')}`
      : new Date().toISOString().slice(0, 10);
    const filename = `Compiled_Attendance_${periodTag}.xlsx`;
    downloadExcel(blob, filename);
  }, [compiledData]);

  const reset = useCallback(() => {
    setFingerprintFile(null);
    setOnlineFile(null);
    setCompiledData([]);
    setError(null);
  }, []);

  return {
    fingerprintFile,
    setFingerprintFile,
    onlineFile,
    setOnlineFile,
    compiledData,
    isCompiling,
    error,
    canCompile,
    hasData,
    compile,
    downloadReport,
    reset,
  };
}
