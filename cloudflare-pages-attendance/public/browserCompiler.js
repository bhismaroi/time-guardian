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

  async function buildCompiledWorkbookFromFiles(fingerprintFile, onlineFile) {
    const [fingerprintRows, onlineRows] = await Promise.all([
      parseFingerprintWorkbook(fingerprintFile),
      parseOnlineWorkbook(onlineFile),
    ]);
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
      addEmployeeSheet(workbook, employee, merged.month);
    }

    return {
      workbook,
      fileName: `Compiled Attendance ${MONTH_NAMES[merged.month.month - 1]}${merged.month.year}.xlsx`,
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

  async function parseFingerprintWorkbook(file) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
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
      const date = parseCellDate(row.getCell(6));
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
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
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
      if (!currentName || !dateLabel || !/^\d{2}\s+\w{3},/.test(dateLabel)) {
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
    const month = detectMonth(fingerprintRows.concat(onlineRows));
    const fingerprintMap = groupRowsByEmployee(fingerprintRows);
    const onlineMap = groupRowsByEmployee(onlineRows);
    const warnings = [];
    const employees = [];
    const usedFingerprintNames = new Set();
    const lowConfidenceMatches = [];

    for (const [onlineName, onlineData] of onlineMap.entries()) {
      const match = findBestFingerprintMatch(onlineData.tokens, fingerprintMap, usedFingerprintNames);
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
        month,
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
        month,
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
    month,
  }) {
    const days = [];
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

  function findBestFingerprintMatch(onlineTokens, fingerprintMap, usedFingerprintNames) {
    let bestMatch = null;

    for (const [candidateName, candidate] of fingerprintMap.entries()) {
      if (usedFingerprintNames.has(candidateName)) {
        continue;
      }

      const score = scoreNameMatch(onlineTokens, candidate.tokens);
      if (score <= 0) {
        continue;
      }

      if (!bestMatch || score > bestMatch.score || (score === bestMatch.score && candidateName < bestMatch.name)) {
        bestMatch = {
          name: candidateName,
          score,
          lowConfidence: score < 5,
        };
      }
    }

    return bestMatch;
  }

  function scoreNameMatch(onlineTokens, fingerprintTokens) {
    if (!onlineTokens.length || !fingerprintTokens.length) {
      return 0;
    }

    const onlineSet = new Set(onlineTokens);
    const fingerprintSet = new Set(fingerprintTokens);
    const intersection = Array.from(onlineSet).filter((token) => fingerprintSet.has(token));

    if (!intersection.length) {
      return 0;
    }

    let score = intersection.length * 3;

    if (onlineTokens[0] && fingerprintTokens[0] && onlineTokens[0] === fingerprintTokens[0]) {
      score += 2;
    }
    if (
      onlineTokens[onlineTokens.length - 1] &&
      fingerprintTokens[fingerprintTokens.length - 1] &&
      onlineTokens[onlineTokens.length - 1] === fingerprintTokens[fingerprintTokens.length - 1]
    ) {
      score += 2;
    }

    return score;
  }

  function addTemplateSheet(workbook, month) {
    const sheet = workbook.addWorksheet('Template');
    styleSheet(sheet, month, 'Template', null);
  }

  function addEmployeeSheet(workbook, employee, month) {
    const sheetName = safeSheetName(employee.displayName || 'Employee', workbook);
    const sheet = workbook.addWorksheet(sheetName);
    styleSheet(sheet, month, employee.displayName, employee);
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
    sheet.getCell('A2').value = `Periode ${formatDateLabel(makeUtcDate(month.year, month.month, 1))} s/d ${formatDateLabel(makeUtcDate(month.year, month.month, new Date(month.year, month.month, 0).getDate()))}`;
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

  function totalHoursFormula(rowNumber) {
    return `IF(OR(G${rowNumber}="",H${rowNumber}=""),"",MAX(0,H${rowNumber}-G${rowNumber}-IF(B${rowNumber}="Fri",TIME(1,30,0),IF(OR(B${rowNumber}="Mon",B${rowNumber}="Tue",B${rowNumber}="Wed",B${rowNumber}="Thu"),TIME(0,30,0),0))))`;
  }

  function tardinessFormula(rowNumber) {
    return `IF(G${rowNumber}="","",MAX(0,G${rowNumber}-TIME(8,30,0)))`;
  }

  function leaveEarlierFormula(rowNumber) {
    return `IF(OR(G${rowNumber}="",H${rowNumber}=""),"",MAX(0,IF(B${rowNumber}="Fri",IF(G${rowNumber}<TIME(8,0,0),TIME(17,0,0),IF(G${rowNumber}<=TIME(8,15,0),TIME(17,15,0),IF(G${rowNumber}<=TIME(8,30,0),TIME(17,30,0),TIME(17,0,0)))),IF(OR(B${rowNumber}="Mon",B${rowNumber}="Tue",B${rowNumber}="Wed",B${rowNumber}="Thu"),IF(G${rowNumber}<TIME(8,0,0),TIME(16,30,0),IF(G${rowNumber}<=TIME(8,15,0),TIME(16,45,0),IF(G${rowNumber}<=TIME(8,30,0),TIME(17,0,0),TIME(16,30,0)))),0))-H${rowNumber}))`;
  }

  function overtimeFormula(rowNumber) {
    return `IF(H${rowNumber}="","",MAX(0,H${rowNumber}-IF(B${rowNumber}="Fri",TIME(18,0,0),IF(OR(B${rowNumber}="Mon",B${rowNumber}="Tue",B${rowNumber}="Wed",B${rowNumber}="Thu"),TIME(17,30,0),24))))`;
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

  function detectMonth(rows) {
    const counts = new Map();

    rows.forEach((row) => {
      const parts = row.dateKey.split('-');
      const key = `${parts[0]}-${parts[1]}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    });

    const best = Array.from(counts.entries()).sort((left, right) => right[1] - left[1])[0];
    if (!best) {
      return null;
    }

    const parts = best[0].split('-').map(Number);
    return { year: parts[0], month: parts[1] };
  }

  function parseCellDate(cell) {
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
      return parseDateString(value.text);
    }

    return parseDateString(String(value));
  }

  function parseDateString(input) {
    const text = normalizeWhitespace(input);
    if (!text) {
      return null;
    }

    let match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      return makeUtcDate(Number(match[3]), Number(match[2]), Number(match[1]));
    }

    match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
      return makeUtcDate(Number(match[1]), Number(match[2]), Number(match[3]));
    }

    return null;
  }

  function parseOnlineDateLabel(label, reportPeriod) {
    const match = label.match(/^(\d{2})\s+([A-Za-z]{3}),\s*(?:[A-Za-z]{2})$/);
    if (!match) {
      return null;
    }

    const month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].indexOf(label.substring(3, 6)) + 1;
    if (!month || !reportPeriod || !reportPeriod.year) {
      return null;
    }

    return makeUtcDate(reportPeriod.year, month, Number(match[1]));
  }

  function parseOnlineReportPeriod(label) {
    const match = normalizeWhitespace(label).match(/^([A-Za-z]{3})\s+\d{1,2},\s+(\d{4})\s+-\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4}$/);
    if (!match) {
      return null;
    }

    const month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].indexOf(match[1]) + 1;
    if (!month) {
      return null;
    }

    return { month, year: Number(match[2]) };
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
    const sanitized = name.replace(/[\\/*?:[\]]/g, ' ').trim().slice(0, 31) || 'Employee';
    let candidate = sanitized;
    let counter = 2;

    while (workbook.getWorksheet(candidate)) {
      const suffix = ` ${counter}`;
      candidate = `${sanitized.slice(0, 31 - suffix.length)}${suffix}`;
      counter += 1;
    }

    return candidate;
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
