// Attendance data compilation logic

import type { CompiledEmployee, MergedAttendanceRecord, RawFingerprintRecord } from './types';
import { formatDateIso, getDayName, getEarlierTime, getLaterTime, isWeekend, normalizeName, parseDate, extractNameParts } from './timeUtils';
import { calculateAttendance, formatCalculationResults } from './attendanceCalculator';
import { getUniqueEmployees, getMonthDates } from './excelParser';

type OnlineDayRecord = { clockIn: string | null; clockOut: string | null };

function getFingerprintDateKey(record: RawFingerprintRecord): string {
  return record.dateKey || record.date || '';
}

function getEmployeeTokens(name: string): string[] {
  return extractNameParts(name).filter((token) => token.length > 1);
}

function scoreNameMatch(fingerprintName: string, onlineKey: string): number {
  const fingerprintTokens = getEmployeeTokens(fingerprintName);
  const onlineTokens = getEmployeeTokens(onlineKey);

  if (fingerprintTokens.length === 0 || onlineTokens.length === 0) return 0;

  const fingerprintFirst = fingerprintTokens[0];
  const fingerprintLast = fingerprintTokens[fingerprintTokens.length - 1];
  const onlineFirst = onlineTokens[0];
  const onlineLast = onlineTokens[onlineTokens.length - 1];

  if (normalizeName(fingerprintName) === normalizeName(onlineKey)) {
    return 100;
  }

  let score = 0;
  if (fingerprintFirst === onlineFirst) score += 40;
  if (fingerprintLast === onlineLast) score += 40;
  if (fingerprintTokens.some((token) => onlineTokens.includes(token))) score += 20;
  if (fingerprintTokens.some((token) => onlineTokens.some((candidate) => candidate.includes(token) || token.includes(candidate)))) {
    score += 10;
  }

  return score;
}

function findOnlineMatch(
  fingerprintName: string,
  onlineData: Map<string, Map<string, OnlineDayRecord>>
): Map<string, OnlineDayRecord> | undefined {
  const normalizedFingerprint = normalizeName(fingerprintName);
  if (!normalizedFingerprint) return undefined;

  if (onlineData.has(normalizedFingerprint)) {
    return onlineData.get(normalizedFingerprint);
  }

  let bestMatch: { key: string; score: number } | null = null;

  for (const key of onlineData.keys()) {
    const score = scoreNameMatch(fingerprintName, key);
    if (score > 0 && (!bestMatch || score > bestMatch.score)) {
      bestMatch = { key, score };
    }
  }

  if (!bestMatch) return undefined;
  return onlineData.get(bestMatch.key);
}

function buildEmployeeName(employee: { name: string }): string {
  return employee.name.trim();
}

function getMonthFromDates(records: RawFingerprintRecord[], onlineData: Map<string, Map<string, OnlineDayRecord>>): { year: number; month: number } {
  const candidates: Date[] = [];

  for (const record of records.slice(0, 50)) {
    const parsed = parseDate(record.date) || parseDate(record.dateKey);
    if (parsed) candidates.push(parsed);
  }

  for (const employeeRecords of onlineData.values()) {
    for (const dateKey of employeeRecords.keys()) {
      const parsed = parseDate(dateKey);
      if (parsed) candidates.push(parsed);
    }
  }

  const first = candidates[0];
  if (!first) {
    return { year: 2025, month: 9 };
  }

  return { year: first.getFullYear(), month: first.getMonth() };
}

function buildUniqueSheetName(name: string, usedNames: Set<string>): string {
  const base = name
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.replace(/[\\/*?[\]:]/g, ''))
    .join(' ')
    .trim()
    .slice(0, 31) || 'Sheet';

  let candidate = base;
  let suffix = 2;

  while (usedNames.has(candidate)) {
    const trimmedBase = base.slice(0, Math.max(0, 31 - ` (${suffix})`.length)).trim();
    candidate = `${trimmedBase} (${suffix})`;
    suffix += 1;
  }

  usedNames.add(candidate);
  return candidate;
}

function pickSourceTime(record: RawFingerprintRecord, field: 'clockIn' | 'clockOut' | 'actualIn' | 'actualOut'): string | null {
  const primary = record[field];
  if (primary) return primary;

  if (field === 'actualIn') return record.clockIn ?? null;
  if (field === 'actualOut') return record.clockOut ?? null;
  return null;
}

/**
 * Compile attendance data from fingerprint and online sources.
 */
export function compileAttendance(
  fingerprintRecords: RawFingerprintRecord[],
  onlineData: Map<string, Map<string, OnlineDayRecord>>
): CompiledEmployee[] {
  const employees = getUniqueEmployees(fingerprintRecords);
  const compiledEmployees: CompiledEmployee[] = [];
  const usedSheetNames = new Set<string>();
  const { year, month } = getMonthFromDates(fingerprintRecords, onlineData);
  const dates = getMonthDates(year, month);

  for (const employee of employees) {
    const fingerprintByDate = new Map<string, { in: string | null; out: string | null }>();

    for (const record of fingerprintRecords) {
      if (normalizeName(record.name) !== normalizeName(employee.name)) continue;

      const dateKey = getFingerprintDateKey(record);
      if (!dateKey) continue;

      const actualIn = pickSourceTime(record, 'actualIn');
      const actualOut = pickSourceTime(record, 'actualOut');

      const existing = fingerprintByDate.get(dateKey);
      if (existing) {
        existing.in = getEarlierTime(existing.in, actualIn);
        existing.out = getLaterTime(existing.out, actualOut);
      } else {
        fingerprintByDate.set(dateKey, {
          in: actualIn,
          out: actualOut,
        });
      }
    }

    const employeeOnlineData = findOnlineMatch(buildEmployeeName(employee), onlineData);

    const records: MergedAttendanceRecord[] = [];
    for (const date of dates) {
      const dateKey = formatDateIso(date);
      const dayName = getDayName(date);

      const fingerprint = fingerprintByDate.get(dateKey);
      const online = employeeOnlineData?.get(dateKey);

      const fingerprintIn = fingerprint?.in || null;
      const fingerprintOut = fingerprint?.out || null;
      const onlineIn = online?.clockIn || null;
      const onlineOut = online?.clockOut || null;

      const actualIn = getEarlierTime(fingerprintIn, onlineIn);
      const actualOut = getLaterTime(fingerprintOut, onlineOut);

      const calculation = calculateAttendance(date, actualIn, actualOut);
      const formatted = formatCalculationResults(calculation);

      const hasAttendance = Boolean(actualIn || actualOut);
      let remarks = '';
      if (!isWeekend(date) && !hasAttendance) {
        remarks = '0';
      }

      records.push({
        date,
        dayOfWeek: dayName,
        fingerprintIn,
        fingerprintOut,
        onlineIn,
        onlineOut,
        actualIn,
        actualOut,
        totalHours: isWeekend(date) ? '' : formatted.totalHours,
        tardiness: formatted.tardiness,
        leaveEarlier: formatted.leaveEarlier,
        overtime: formatted.overtime,
        remarks,
      });
    }

    const employeeName = employee.name || 'Unknown';
    const sheetName = buildUniqueSheetName(employeeName.split(' ')[0] || employeeName, usedSheetNames);

    compiledEmployees.push({
      empNo: employee.empNo,
      name: employeeName,
      sheetName,
      nip: `000${employee.empNo || '0'}`.slice(-6),
      division: 'MITSUI OSK LINES',
      department: 'IDACT',
      section: 'IDACT',
      records,
    });
  }

  return compiledEmployees;
}
