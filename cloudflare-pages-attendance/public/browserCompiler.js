(function () {
  const WEEKDAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const MONTH_NAMES = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ];

  async function buildCompiledWorkbookFromFiles(fingerprintFile, onlineFile, { onProgress } = {}) {
    // The online workbook's period label (e.g. "Aug 1, 2026 - Aug 31,
    // 2026") is the authoritative month/year. Read it first so the
    // fingerprint dates can be disambiguated: the fingerprint export
    // uses day-first dates ("03/11/2025") for some months and
    // month-first US style ("8/1/2026" = 1 August) for others.
    if (onProgress) onProgress('Reading online report period...');
    const reportPeriod = await readOnlineReportPeriod(onlineFile);
    if (onProgress) onProgress('Reading fingerprint file...');
    const fingerprintRows = await parseFingerprintWorkbook(fingerprintFile, reportPeriod);
    if (onProgress) onProgress('Reading online file...');
    const onlineRows = await parseOnlineWorkbook(onlineFile);
    if (onProgress) onProgress('Merging attendance records...');
    const merged = mergeAttendance(fingerprintRows, onlineRows);

    if (!merged.month) {
      throw new Error('Could not detect any attendance dates from the uploaded files.');
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Codex Attendance Compiler';
    workbook.created = new Date();
    workbook.modified = new Date();

    addTemplateSheet(workbook, merged.month);

    for (const employee of merged.employees) {
      const truncationWarning = addEmployeeSheet(workbook, employee, merged.month);
      if (truncationWarning) {
        merged.warnings.push(truncationWarning);
      }
    }

    return {
      workbook,
      fileName: `Compiled Attendance ${MONTH_NAMES[merged.month.month - 1]} ${merged.month.year}.xlsx`,
      warnings: merged.warnings,
      summary: {
        employees: merged.employees.length,
        matchedEmployees: merged.summary.matchedEmployees,
        fingerprintOnlyEmployees: merged.summary.fingerprintOnlyEmployees,
        onlineOnlyEmployees: merged.summary.onlineOnlyEmployees,
        lowConfidenceMatches: merged.summary.lowConfidenceMatches,
        month: `${String(merged.month.month).padStart(2, '0')}/${merged.month.year}`,
      },
    };
  }

  async function readOnlineReportPeriod(file) {
    let workbook;
    try {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
    } catch (err) {
      return null;
    }
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      return null;
    }
    return parseOnlineReportPeriod(cellText(worksheet.getCell('A3')));
  }

  async function parseFingerprintWorkbook(file, reportPeriod) {
    let workbook;
    try {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
    } catch (err) {
      throw new Error(`Failed to read fingerprint Excel file: ${err.message || 'The file may be corrupted or not a valid Excel file.'}`);
    }
    const worksheet = workbook.worksheets[0];

    if (!worksheet) {
      throw new Error('Fingerprint workbook does not contain any worksheet.');
    }

    const rows = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        return;
      }

      const name = normalizeWhitespace(cellText(row.getCell(4)));
      const date = parseCellDate(row.getCell(6), reportPeriod);
      const actualIn = parseTimeValue(row.getCell(10).value);
      const actualOut = parseTimeValue(row.getCell(11).value);

      if (!name || !date || looksLikeGarbageName(name)) {
        return;
      }

      rows.push({
        source: 'fingerprint',
        name,
        tokens: tokenizeName(name),
        dateKey: formatDateKey(date),
        actualIn,
        actualOut,
      });
    });

    return rows;
  }

  async function parseOnlineWorkbook(file) {
    let workbook;
    try {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
    } catch (err) {
      throw new Error(`Failed to read online Excel file: ${err.message || 'The file may be corrupted or not a valid Excel file.'}`);
    }
    const worksheet = workbook.worksheets[0];

    if (!worksheet) {
      throw new Error('Online workbook does not contain any worksheet.');
    }

    const rows = [];
    let currentName = '';
    const reportPeriod = parseOnlineReportPeriod(cellText(worksheet.getCell('A3')));

    worksheet.eachRow((row) => {
      const label = normalizeWhitespace(cellText(row.getCell(2)));

      if (label === 'Full name') {
        currentName = normalizeWhitespace(cellText(row.getCell(3)));
        return;
      }

      const dateLabel = normalizeWhitespace(cellText(row.getCell(1)));
      if (!currentName || !dateLabel || !/^\d{1,2}\s+[A-Za-z]{3,9},/.test(dateLabel)) {
        return;
      }

      const date = parseOnlineDateLabel(dateLabel, reportPeriod);
      if (!date) {
        return;
      }

      rows.push({
        source: 'online',
        name: currentName,
        tokens: tokenizeName(currentName),
        dateKey: formatDateKey(date),
        actualIn: parseTimeValue(row.getCell(4).value),
        actualOut: parseTimeValue(row.getCell(5).value),
      });
    });

    return rows;
  }

  function mergeAttendance(fingerprintRows, onlineRows) {
    const months = detectMonths(fingerprintRows.concat(onlineRows));
    // For display-only consumers (filename, period label, summary), we use
    // the first month as the "primary" month. If the data spans multiple
    // months, the workbook still contains every month; only the file
    // label and sheet header reflect a single month.
    const month = months[0] || null;
    const fingerprintMap = groupRowsByEmployee(fingerprintRows);
    const onlineMap = groupRowsByEmployee(onlineRows);
    const warnings = [];
    const employees = [];
    const usedFingerprintNames = new Set();
    const lowConfidenceMatches = [];

    for (const [onlineName, onlineData] of onlineMap.entries()) {
      const match = findBestFingerprintMatch(onlineData.tokens, fingerprintMap, usedFingerprintNames, warnings);
      let fingerprintName = null;

      if (match) {
        fingerprintName = match.name;
        usedFingerprintNames.add(match.name);

        if (match.lowConfidence) {
          const warning = `Low-confidence name match: online "${onlineName}" matched fingerprint "${match.name}".`;
          warnings.push(warning);
          lowConfidenceMatches.push(warning);
        }
      } else {
        warnings.push(`No fingerprint match found for online employee "${onlineName}".`);
      }

      const fingerprintData = fingerprintName ? fingerprintMap.get(fingerprintName) : null;
      employees.push(buildEmployeeAttendanceRecord({
        displayName: fingerprintName || onlineName,
        fingerprintName,
        onlineName,
        fingerprintData,
        onlineData,
        months,
      }));
    }

    for (const [fingerprintName, fingerprintData] of fingerprintMap.entries()) {
      if (usedFingerprintNames.has(fingerprintName)) {
        continue;
      }

      warnings.push(`No online match found for fingerprint employee "${fingerprintName}".`);
      employees.push(buildEmployeeAttendanceRecord({
        displayName: fingerprintName,
        fingerprintName,
        onlineName: null,
        fingerprintData,
        onlineData: null,
        months,
      }));
    }

    employees.sort((left, right) => left.displayName.localeCompare(right.displayName));

    return {
      month,
      employees,
      warnings,
      summary: {
        matchedEmployees: employees.filter((employee) => employee.fingerprintName && employee.onlineName).length,
        fingerprintOnlyEmployees: employees.filter((employee) => employee.fingerprintName && !employee.onlineName).length,
        onlineOnlyEmployees: employees.filter((employee) => !employee.fingerprintName && employee.onlineName).length,
        lowConfidenceMatches: lowConfidenceMatches.length,
      },
    };
  }

  function groupRowsByEmployee(rows) {
    const map = new Map();

    for (const row of rows) {
      if (!map.has(row.name)) {
        map.set(row.name, {
          tokens: row.tokens,
          byDate: new Map(),
        });
      }

      const employee = map.get(row.name);
      if (!employee.byDate.has(row.dateKey)) {
        employee.byDate.set(row.dateKey, {
          sourceInTimes: [],
          sourceOutTimes: [],
        });
      }

      const day = employee.byDate.get(row.dateKey);
      if (row.actualIn != null) {
        day.sourceInTimes.push(row.actualIn);
      }
      if (row.actualOut != null) {
        day.sourceOutTimes.push(row.actualOut);
      }
    }

    return map;
  }

  function buildEmployeeAttendanceRecord({
    displayName,
    fingerprintName,
    onlineName,
    fingerprintData,
    onlineData,
    months,
  }) {
    const days = [];
    // dayCount = last day of the target month. Passing the (1-based) month as
    // the 0-based month parameter and day 0 rolls over to the last day of the
    // intended month. The previous expression `month.month - 1` returned the
    // *previous* month's day count, silently truncating non-January/non-August
    // months (e.g. March 2026 -> 28, Feb leap -> 31 overshoot).
    //
    // The outer loop iterates every detected month so a fingerprint/online
    // workbook split across Mar/Apr renders both months' data. Pre-fix this
    // function took a single `month` and silently dropped the others.
    for (const month of months) {
      const dayCount = new Date(month.year, month.month, 0).getDate();

      for (let day = 1; day <= dayCount; day += 1) {
        const date = new Date(Date.UTC(month.year, month.month - 1, day));
        const dateKey = formatDateKey(date);
        const fingerprintDay = fingerprintData && fingerprintData.byDate.get(dateKey);
        const onlineDay = onlineData && onlineData.byDate.get(dateKey);
        const sourceInTimes = []
          .concat((fingerprintDay && fingerprintDay.sourceInTimes) || [])
          .concat((onlineDay && onlineDay.sourceInTimes) || [])
          .sort((left, right) => left - right);
        const fingerprintOut = [].concat((fingerprintDay && fingerprintDay.sourceOutTimes) || []).sort((left, right) => right - left);
        const onlineOut = [].concat((onlineDay && onlineDay.sourceOutTimes) || []).sort((left, right) => right - left);
        const sourceOutTimes = fingerprintOut.concat(onlineOut).sort((left, right) => right - left);
        const mergedIn = sourceInTimes.length ? sourceInTimes[0] : null;
        const mergedOut = chooseMergedOut(fingerprintOut, onlineOut);

        days.push({
          date,
          dateKey,
          mergedIn,
          mergedOut,
          sourceInTimes,
          sourceOutTimes,
          sourceTrace: {
            fingerprintInCount: fingerprintDay ? fingerprintDay.sourceInTimes.length : 0,
            fingerprintOutCount: fingerprintDay ? fingerprintDay.sourceOutTimes.length : 0,
            onlineInCount: onlineDay ? onlineDay.sourceInTimes.length : 0,
            onlineOutCount: onlineDay ? onlineDay.sourceOutTimes.length : 0,
          },
        });
      }
    }

    return {
      displayName,
      fingerprintName,
      onlineName,
      days,
    };
  }

  function chooseMergedOut(fingerprintOut, onlineOut) {
    const bestFingerprint = fingerprintOut[0] == null ? null : fingerprintOut[0];
    const bestOnline = onlineOut[0] == null ? null : onlineOut[0];

    if (bestFingerprint == null) {
      return bestOnline;
    }
    if (bestOnline == null) {
      return bestFingerprint;
    }
    if (bestFingerprint === bestOnline) {
      return bestFingerprint;
    }

    return Math.max(bestFingerprint, bestOnline);
  }

  // Minimum score (0-100) for a fuzzy name match. First-name and
  // last-name matches each score 40; a candidate must match BOTH (80)
  // to be accepted. A single shared first or last name (40) is not
  // enough — that is the failure mode that attributed "Adi Saputra"'s
  // online data to "Adi Wijaya" (both share the first name "adi").
  const MIN_FUZZY_MATCH_SCORE = 80;

  function findBestFingerprintMatch(onlineTokens, fingerprintMap, usedFingerprintNames, warnings) {
    const candidates = [];

    for (const [candidateName, candidate] of fingerprintMap.entries()) {
      if (usedFingerprintNames.has(candidateName)) {
        continue;
      }

      const score = scoreNameMatch(onlineTokens, candidate.tokens);
      if (score < MIN_FUZZY_MATCH_SCORE) {
        continue;
      }

      candidates.push({ name: candidateName, score });
    }

    if (candidates.length > 0) {
      // Deterministic tie-break: highest score, then alphabetical name.
      candidates.sort((left, right) => (right.score - left.score) || left.name.localeCompare(right.name));
      const topScore = candidates[0].score;
      const tied = candidates.filter((candidate) => candidate.score === topScore);

      if (tied.length > 1) {
        // Two or more fingerprint employees look equally plausible.
        // Surface the ambiguity rather than silently picking one, so the
        // user can see that data was not attributed to the wrong person.
        warnings.push(
          `Ambiguous fingerprint match for online employee "${onlineTokens.join(' ')}": ${tied.length} candidates tied at score ${topScore} (${tied.map((t) => t.name).join(', ')}).`
        );
        return null;
      }

      const winner = candidates[0];
      return {
        name: winner.name,
        score: winner.score,
        lowConfidence: winner.score < 100,
      };
    }

    // Single-token fallback. Some employees are exported as a lone first
    // name in one system but a full name in the other (e.g. fingerprint
    // "Patricia" vs online "Patricia Pranata"). A strict first+last match
    // cannot pair those. We accept a first-name-only match only when it is
    // unambiguous — exactly one fingerprint candidate shares that first
    // name — AND one of the two names is a single token. The single-token
    // guard keeps "Adi Wijaya" vs "Adi Saputra" (both multi-token, both
    // sharing "adi") from being mispaired, while recovering the
    // legitimate lone-first-name cases.
    const firstToken = onlineTokens[0];
    const firstTokenCandidates = [];

    for (const [candidateName, candidate] of fingerprintMap.entries()) {
      if (usedFingerprintNames.has(candidateName)) {
        continue;
      }
      if (candidate.tokens[0] === firstToken) {
        firstTokenCandidates.push({ name: candidateName, tokens: candidate.tokens });
      }
    }

    if (firstTokenCandidates.length !== 1) {
      return null;
    }

    const onlyCandidate = firstTokenCandidates[0];
    if (onlyCandidate.tokens.length !== 1 && onlineTokens.length !== 1) {
      return null;
    }

    return { name: onlyCandidate.name, score: 40, lowConfidence: true };
  }

  function scoreNameMatch(onlineTokens, fingerprintTokens) {
    if (!onlineTokens.length || !fingerprintTokens.length) {
      return 0;
    }

    const onlineKey = onlineTokens.join(' ');
    const fingerprintKey = fingerprintTokens.join(' ');

    // Exact full-name match.
    if (onlineKey === fingerprintKey) {
      return 100;
    }

    // Reversed full-name match ("Adi Wijaya" vs "Wijaya Adi").
    if (onlineKey === fingerprintTokens.slice().reverse().join(' ')) {
      return 100;
    }

    // Fuzzy: first-name and last-name overlap, each worth 40. We
    // deliberately do NOT count shared middle tokens or single-token
    // names, which were the source of false-positive matches.
    let score = 0;
    if (onlineTokens[0] === fingerprintTokens[0]) {
      score += 40;
    }
    if (onlineTokens[onlineTokens.length - 1] === fingerprintTokens[fingerprintTokens.length - 1]) {
      score += 40;
    }

    return score;
  }

  function addTemplateSheet(workbook, month) {
    const sheet = workbook.addWorksheet('Template');
    styleSheet(sheet, month, 'Template', null);
  }

  function addEmployeeSheet(workbook, employee, month) {
    const { name: sheetName, truncated } = safeSheetName(employee.displayName || 'Employee', workbook);
    const sheet = workbook.addWorksheet(sheetName);
    styleSheet(sheet, month, employee.displayName, employee);
    return truncated ? `Sheet name for "${employee.displayName}" was truncated to "${sheetName}" due to Excel's 31-character limit.` : null;
  }

  function styleSheet(sheet, month, employeeName, employee) {
    sheet.properties.defaultRowHeight = 20;
    sheet.views = [{ state: 'frozen', ySplit: 6 }];
    sheet.columns = [
      { width: 14 },
      { width: 10 },
      { width: 10 },
      { width: 16 },
      { width: 14 },
      { width: 14 },
      { width: 12 },
      { width: 12 },
      { width: 14 },
      { width: 12 },
      { width: 13 },
      { width: 12 },
      { width: 26 },
    ];

    sheet.mergeCells('A1:M1');
    sheet.mergeCells('A2:M2');
    sheet.mergeCells('A6:M6');

    sheet.getCell('A1').value = 'Laporan Absensi Harian';
    // Period label: end-of-month date uses the same corrected day-of-month
    // idiom (passing the 1-based month as 0-based, day 0 = last day of the
    // intended month). See dayCount comment above for the bug history.
    const lastDayOfMonth = new Date(month.year, month.month, 0).getDate();
    sheet.getCell('A2').value = `Periode ${formatDateLabel(makeUtcDate(month.year, month.month, 1))} s/d ${formatDateLabel(makeUtcDate(month.year, month.month, lastDayOfMonth))}`;
    sheet.getCell('A4').value = 'Date';
    sheet.getCell('B4').value = 'Day';
    sheet.getCell('C4').value = 'Kal';
    sheet.getCell('D4').value = 'Shift';
    sheet.getCell('E4').value = 'Office Hours';
    sheet.getCell('G4').value = 'Actual In';
    sheet.getCell('H4').value = 'Actual Out';
    sheet.getCell('I4').value = 'Total Hours';
    sheet.getCell('J4').value = 'Tardiness';
    sheet.getCell('K4').value = 'Leave Earlier';
    sheet.getCell('L4').value = 'Overtime';
    sheet.getCell('M4').value = 'Remarks';
    sheet.getCell('E5').value = 'In';
    sheet.getCell('F5').value = 'Out';
    sheet.getCell('A6').value = `Nama : ${employeeName}`;

    ['A1', 'A2', 'A6'].forEach((address) => {
      sheet.getCell(address).font = { bold: true };
    });

    for (let row = 4; row <= 5; row += 1) {
      for (let col = 1; col <= 13; col += 1) {
        const cell = sheet.getRow(row).getCell(col);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD9EAF7' },
        };
        cell.border = thinBorder();
        cell.font = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    if (!employee) {
      return;
    }

    let rowNumber = 7;
    for (const day of employee.days) {
      populateAttendanceRow(sheet, rowNumber, day);
      rowNumber += 1;
    }
  }

  function populateAttendanceRow(sheet, rowNumber, day) {
    const row = sheet.getRow(rowNumber);
    const weekday = day.date.getUTCDay();
    const isFriday = weekday === 5;
    const isWeekend = weekday === 0 || weekday === 6;
    const standardOut = isFriday ? '17:00' : '16:30';

    row.getCell(1).value = day.date;
    row.getCell(1).numFmt = 'dd/mm/yyyy';
    row.getCell(2).value = WEEKDAY_NAMES[weekday];
    row.getCell(3).value = 'WD';
    row.getCell(4).value = isWeekend ? '0' : isFriday ? '08.00 - 17.00' : '08.00 - 16.30';
    row.getCell(5).value = isWeekend ? '0' : 'C 08:00';
    row.getCell(6).value = isWeekend ? '0' : `C ${standardOut}`;
    row.getCell(7).value = day.mergedIn == null ? null : day.mergedIn / 1440;
    row.getCell(8).value = day.mergedOut == null ? null : day.mergedOut / 1440;
    row.getCell(7).numFmt = 'hh:mm';
    row.getCell(8).numFmt = 'hh:mm';
    row.getCell(9).value = { formula: totalHoursFormula(rowNumber) };
    row.getCell(10).value = { formula: tardinessFormula(rowNumber) };
    row.getCell(11).value = { formula: leaveEarlierFormula(rowNumber) };
    row.getCell(12).value = { formula: overtimeFormula(rowNumber) };
    row.getCell(13).value = remarksForDay(day, isWeekend);

    for (let col = 1; col <= 13; col += 1) {
      const cell = row.getCell(col);
      cell.border = thinBorder();
      if (col >= 9 && col <= 12) {
        cell.numFmt = '[h]:mm';
      }
      cell.alignment = { vertical: 'middle', horizontal: col === 13 ? 'left' : 'center' };
    }

    addConditionalFormatting(sheet, rowNumber);
  }

  function addConditionalFormatting(sheet, rowNumber) {
    [
      { column: 'J', color: 'FFFECACA' },
      { column: 'K', color: 'FFFEF3C7' },
      { column: 'L', color: 'FFE4DFEC' },
    ].forEach(({ column, color }) => {
      sheet.addConditionalFormatting({
        ref: `${column}${rowNumber}`,
        rules: [
          {
            type: 'expression',
            formulae: [`AND(${column}${rowNumber}<>"",${column}${rowNumber}>0)`],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                bgColor: { argb: color },
                fgColor: { argb: color },
              },
            },
          },
        ],
      });
    });
  }

  // Formula builders delegate to the canonical AttendancePolicy module
  // (loaded as policy.js before this script). The day-of-week checks
  // come from cloudflareDayChecks because the day name lives in column
  // B (not column A as a WEEKDAY number) in the Cloudflare layout.
  // Pre-Phase 2 these were hand-written inline; the formulas produced
  // here are semantically equivalent to the originals but use a
  // slightly different string shape (explicit OR(...) for the weekend
  // guard, parenthesised (H-G) for the total hours subtraction, and a
  // leading '=' which ExcelJS accepts with or without).
  function totalHoursFormula(rowNumber) {
    return AttendancePolicy.buildTotalHoursFormula(rowNumber, AttendancePolicy.cloudflareDayChecks(rowNumber));
  }

  function tardinessFormula(rowNumber) {
    return AttendancePolicy.buildTardinessFormula(rowNumber, AttendancePolicy.cloudflareDayChecks(rowNumber));
  }

  function leaveEarlierFormula(rowNumber) {
    return AttendancePolicy.buildLeaveEarlierFormula(rowNumber, AttendancePolicy.cloudflareDayChecks(rowNumber));
  }

  function overtimeFormula(rowNumber) {
    return AttendancePolicy.buildOvertimeFormula(rowNumber, AttendancePolicy.cloudflareDayChecks(rowNumber));
  }

  function remarksForDay(day, isWeekend) {
    if (!day.mergedIn && !day.mergedOut) {
      return isWeekend ? 'Weekend' : '';
    }

    const remarks = [];
    if (day.sourceTrace.fingerprintInCount || day.sourceTrace.fingerprintOutCount) {
      remarks.push('Fingerprint');
    }
    if (day.sourceTrace.onlineInCount || day.sourceTrace.onlineOutCount) {
      remarks.push('Online');
    }
    return remarks.join(' + ');
  }

  // Returns the sorted set of months seen in the input rows. Each entry
  // is a 1-based { year, month } pair. The original implementation picked
  // the single most-frequent month and silently dropped rows for any
  // other month, so a Mar/Apr split report would render only March with
  // no indication that April data was lost. Returning the set lets the
  // caller iterate every month.
  function detectMonths(rows) {
    const counts = new Map();

    rows.forEach((row) => {
      const parts = row.dateKey.split('-');
      const key = `${parts[0]}-${parts[1]}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    });

    return Array.from(counts.entries())
      .map(([key]) => {
        const parts = key.split('-').map(Number);
        return { year: parts[0], month: parts[1] };
      })
      .sort((left, right) => (left.year - right.year) || (left.month - right.month));
  }

  function parseCellDate(cell, reportPeriod) {
    const value = cell.value;
    if (!value) {
      return null;
    }

    if (value instanceof Date) {
      return makeUtcDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
    }

    if (typeof value === 'number') {
      const date = new Date(Math.round((value - 25569) * 86400 * 1000));
      return makeUtcDate(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
    }

    if (typeof value === 'object' && value.text) {
      return parseDateString(value.text, reportPeriod);
    }

    return parseDateString(String(value), reportPeriod);
  }

  function resolveAmbiguousDate(a, b, year, reportPeriod) {
    const dayFirst = makeUtcDate(year, b, a);
    const monthFirst = makeUtcDate(year, a, b);

    const dayFirstValid = dayFirst.getUTCFullYear() === year && dayFirst.getUTCMonth() + 1 === b;
    const monthFirstValid = monthFirst.getUTCFullYear() === year && monthFirst.getUTCMonth() + 1 === a;

    if (dayFirstValid && !monthFirstValid) return dayFirst;
    if (monthFirstValid && !dayFirstValid) return monthFirst;

    if (reportPeriod) {
      const monthInRange = (month, date) => {
        const y = date.getUTCFullYear();
        if (reportPeriod.startMonth === reportPeriod.endMonth) {
          return month === reportPeriod.startMonth && y === reportPeriod.startYear;
        }
        const expectedYear = month >= reportPeriod.startMonth ? reportPeriod.startYear : reportPeriod.endYear;
        return (month >= reportPeriod.startMonth && month <= 12 && y === expectedYear)
          || (month < reportPeriod.startMonth && y === expectedYear);
      };
      if (dayFirstValid && monthInRange(b, dayFirst)) return dayFirst;
      if (monthFirstValid && monthInRange(a, monthFirst)) return monthFirst;
    }

    return dayFirst;
  }

  function parseDateString(input, reportPeriod) {
    const text = normalizeWhitespace(input);
    if (!text) {
      return null;
    }

    let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      return resolveAmbiguousDate(Number(match[1]), Number(match[2]), Number(match[3]), reportPeriod);
    }

    match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
      return makeUtcDate(Number(match[1]), Number(match[2]), Number(match[3]));
    }

    return null;
  }

  // Map a month label (3-letter abbreviation or full name) to a 1-based
  // month number. Returns 0 for an unrecognised label.
  function monthNumber(label) {
    const abbr = normalizeWhitespace(label).slice(0, 3).toLowerCase();
    return ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(abbr) + 1;
  }

  function parseOnlineDateLabel(label, reportPeriod) {
    const match = label.match(/^(\d{1,2})\s+([A-Za-z]{3,9}),\s*(?:[A-Za-z]{2})$/);
    if (!match) {
      return null;
    }

    const month = monthNumber(match[2]);
    if (!month || !reportPeriod || !reportPeriod.startYear) {
      return null;
    }

    let year = reportPeriod.startYear;
    if (month < reportPeriod.startMonth) {
      year = reportPeriod.endYear;
    }

    return makeUtcDate(year, month, Number(match[1]));
  }

  function parseOnlineReportPeriod(label) {
    const match = normalizeWhitespace(label).match(/^([A-Za-z]{3,9})\s+\d{1,2},\s+(\d{4})\s+-\s+([A-Za-z]{3,9})\s+\d{1,2},\s+(\d{4})$/);
    if (!match) {
      return null;
    }

    const startMonth = monthNumber(match[1]);
    const endMonth = monthNumber(match[3]);
    if (!startMonth || !endMonth) {
      return null;
    }

    return { startMonth, startYear: Number(match[2]), endMonth, endYear: Number(match[4]) };
  }

  function parseTimeValue(value) {
    if (value == null || value === '') {
      return null;
    }

    if (value instanceof Date) {
      return value.getHours() * 60 + value.getMinutes();
    }

    if (typeof value === 'number') {
      return Math.round(value * 24 * 60);
    }

    if (typeof value === 'object') {
      if (value.text) {
        return parseTimeString(value.text);
      }
      if (value.result != null) {
        return parseTimeValue(value.result);
      }
    }

    return parseTimeString(String(value));
  }

  function parseTimeString(input) {
    const text = normalizeWhitespace(input).replace('.', ':');
    if (!text || text === '-') {
      return null;
    }

    const match = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!match) {
      return null;
    }

    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    if (hours > 23 || minutes > 59) {
      return null;
    }

    return hours * 60 + minutes;
  }

  function formatDateKey(date) {
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
  }

  function formatDateLabel(date) {
    return `${String(date.getUTCDate()).padStart(2, '0')}/${String(date.getUTCMonth() + 1).padStart(2, '0')}/${date.getUTCFullYear()}`;
  }

  function tokenizeName(name) {
    return normalizeWhitespace(name)
      .toLowerCase()
      .split(/\s+/)
      .map((token) => token.replace(/[^a-z0-9]/g, ''))
      .filter(Boolean);
  }

  function normalizeWhitespace(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function looksLikeGarbageName(name) {
    return /^\d+$/.test(name);
  }

  function safeSheetName(name, workbook) {
    const originalName = name.replace(/[\\/*?:[\]]/g, ' ').trim();
    const sanitized = originalName.slice(0, 31) || 'Employee';
    const truncated = originalName.length > 31;
    let candidate = sanitized;
    let counter = 2;

    while (workbook.getWorksheet(candidate)) {
      const suffix = ` ${counter}`;
      candidate = `${sanitized.slice(0, 31 - suffix.length)}${suffix}`;
      counter += 1;
    }

    return { name: candidate, truncated };
  }

  function thinBorder() {
    return {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  }

  function makeUtcDate(year, month, day) {
    return new Date(Date.UTC(year, month - 1, day));
  }

  function cellText(cell) {
    const value = cell.value;
    if (value == null) {
      return '';
    }
    if (typeof value === 'object' && value.text) {
      return value.text;
    }
    if (typeof value === 'object' && value.richText) {
      return value.richText.map((part) => part.text).join('');
    }
    return String(value);
  }

  window.AttendanceCompiler = {
    buildCompiledWorkbookFromFiles,
  };
})();
