// Verify the React-side formula output of shared/policy.js against
// the formulas that excelGenerator.ts produced before Phase 2. Each
// builder's actual output must equal the expected string byte-for-byte
// so existing workbooks are not affected by the refactor.

import * as m from '../shared/policy.js';

const row = 6;
const dayChecks = m.reactDayChecks(row);

const expected = {
  break: `IF(${dayChecks.friday},IF(AND(G${row}<TIME(13,0,0),H${row}>TIME(11,30,0)),MIN(H${row},TIME(13,0,0))-MAX(G${row},TIME(11,30,0)),0),IF(${dayChecks.weekday},IF(AND(G${row}<TIME(12,30,0),H${row}>TIME(12,0,0)),MIN(H${row},TIME(12,30,0))-MAX(G${row},TIME(12,0,0)),0),0))`,
  total: `=IF(OR(G${row}="",H${row}="",${dayChecks.weekend}),"",MAX(0,(H${row}-G${row})-(${m.buildBreakFormula(row, dayChecks)})))`,
  tardiness: `=IF(OR(G${row}="",${dayChecks.weekend}),"",MAX(0,G${row}-TIME(8,30,0)))`,
  leaveEarlier: `=IF(OR(G${row}="",H${row}="",${dayChecks.weekend}),"",MAX(0,IF(G${row}<=TIME(8,0,0),IF(${dayChecks.friday},TIME(17,0,0),TIME(16,30,0)),IF(G${row}<=TIME(8,15,0),IF(${dayChecks.friday},TIME(17,15,0),TIME(16,45,0)),IF(${dayChecks.friday},TIME(17,30,0),TIME(17,0,0))))-H${row}))`,
  overtime: `=IF(OR(H${row}="",${dayChecks.weekend}),"",MAX(0,H${row}-IF(${dayChecks.friday},TIME(18,0,0),TIME(17,30,0))))`,
};

const actual = {
  break: m.buildBreakFormula(row, dayChecks),
  total: m.buildTotalHoursFormula(row, dayChecks),
  tardiness: m.buildTardinessFormula(row, dayChecks),
  leaveEarlier: m.buildLeaveEarlierFormula(row, dayChecks),
  overtime: m.buildOvertimeFormula(row, dayChecks),
};

let allMatch = true;
for (const k of Object.keys(expected)) {
  const match = actual[k] === expected[k];
  if (!match) allMatch = false;
  console.log(match ? 'OK  ' : 'FAIL', k);
  if (!match) {
    console.log('  expected:', expected[k]);
    console.log('  actual:  ', actual[k]);
  }
}

// Also smoke-test the Cloudflare dayChecks shape.
const cfChecks = m.cloudflareDayChecks(6);
console.log('cloudflare dayChecks:', JSON.stringify(cfChecks));
process.exit(allMatch ? 0 : 1);
