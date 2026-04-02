import { useState, useCallback } from 'react';
import type { CompiledEmployee } from '@/lib/types';
import { parseFingerprintExcel, parseOnlineExcel } from '@/lib/excelParser';
import { compileAttendance } from '@/lib/attendanceCompiler';
import { generateAttendanceExcel, downloadExcel } from '@/lib/excelGenerator';

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
