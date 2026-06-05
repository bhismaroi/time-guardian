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

// Minimum score for a fuzzy name match. The score is the sum of first-name
// and last-name matches (each 0 or 40). To accept a fuzzy match, the
// candidate must match on BOTH the first and the last name (score 80). A
// single shared first or last name alone (40) is not enough: that is the
// failure mode that previously attributed "Adi Saputra"'s online data to
// "Adi Wijaya" (both share first name "adi" — 40 points).
const MIN_FUZZY_MATCH_SCORE = 80;

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
  // The token-shared clause (previously +20 for any shared token) and the
  // substring clause (previously +10 for any substring containment) are
  // removed. Both produced false positives: "Adi Saputra" matching "Adi
  // Pratama" (only first token shared) and "Eric" matching "Erica"
  // (substring). First+last name matching alone is sufficient for the
  // legitimate same-person cases: "Adi Misykatul Anwar" vs "Adi Anwar"
  // scores 80, and the reversed-name case is handled by the explicit
  // reversed-alias check in findOnlineMatch.

  return score;
}

function findOnlineMatch(
  fingerprintName: string,
  onlineData: Map<string, Map<string, OnlineDayRecord>>,
  usedKeys: Set<string>
): Map<string, OnlineDayRecord> | undefined {
  const normalizedFingerprint = normalizeName(fingerprintName);
  if (!normalizedFingerprint) return undefined;

  // 1. Exact normalized full-name match.
  if (onlineData.has(normalizedFingerprint) && !usedKeys.has(normalizedFingerprint)) {
    usedKeys.add(normalizedFingerprint);
    return onlineData.get(normalizedFingerprint);
  }

  // 2. Reversed full-name match ("Adi Wijaya" vs "Wijaya Adi"). The alias
  //    seeder registers reversed full names too, so this is a hash lookup.
  const tokens = getEmployeeTokens(fingerprintName);
  const reversed = tokens.slice().reverse().join(' ');
  if (reversed && reversed !== normalizedFingerprint && onlineData.has(reversed) && !usedKeys.has(reversed)) {
    usedKeys.add(reversed);
    return onlineData.get(reversed);
  }

  // 3. Fuzzy match with a minimum score threshold. We collect every
  //    candidate at or above the threshold so that ties can be detected
  //    and reported as an ambiguous match (returning undefined).
  const candidates: { key: string; score: number }[] = [];
  for (const key of onlineData.keys()) {
    if (usedKeys.has(key)) continue;
    const score = scoreNameMatch(fingerprintName, key);
    if (score >= MIN_FUZZY_MATCH_SCORE) {
      candidates.push({ key, score });
    }
  }

  if (candidates.length === 0) return undefined;

  // Deterministic tie-break: highest score, then alphabetical key.
  candidates.sort((a, b) => (b.score - a.score) || a.key.localeCompare(b.key));
  const topScore = candidates[0].score;
  const tied = candidates.filter((c) => c.score === topScore);
  if (tied.length > 1) {
    // Two or more distinct online employees look equally plausible.
    // Surface the ambiguity rather than silently picking one. The proper
    // warnings channel lands in Phase 2; for now this goes to console.
    console.warn(
      `[attendance] ambiguous online match for "${fingerprintName}": ${tied.length} candidates tied at score ${topScore} (${tied.map((t) => t.key).join(', ')}). Skipping.`
    );
    return undefined;
  }

  const winner = candidates[0];
  usedKeys.add(winner.key);
  return onlineData.get(winner.key);
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
  // Tracks online keys already bound to a fingerprint employee in this
  // compile pass. Defence in depth on top of the no-single-token-alias
  // rule: even if a fingerprint name scores equally well against two
  // online employees, only the first match consumes the record map.
  const usedOnlineKeys = new Set<string>();
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

    const employeeOnlineData = findOnlineMatch(buildEmployeeName(employee), onlineData, usedOnlineKeys);

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
      records,
    });
  }

  return compiledEmployees;
}
