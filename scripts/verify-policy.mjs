import * as m from '../shared/policy.js';

const row = 6;
const dayExpr = m.reactDayExpr(row);
const weekday = dayExpr.replace('=5', '');

const expected = {
  break: `IF(${dayExpr}=5,IF(AND(G${row}<TIME(13,0,0),H${row}>TIME(11,30,0)),MIN(H${row},TIME(13,0,0))-MAX(G${row},TIME(11,30,0)),0),IF(${dayExpr}<=4,IF(AND(G${row}<TIME(12,30,0),H${row}>TIME(12,0,0)),MIN(H${row},TIME(12,30,0))-MAX(G${row},TIME(12,0,0)),0),0))`,
  total: `=IF(OR(G${row}="",H${row}="",${dayExpr}>5),"",MAX(0,(H${row}-G${row})-(${m.buildBreakFormula(row, dayExpr)})))`,
  tardiness: `=IF(OR(G${row}="",${dayExpr}>5),"",MAX(0,G${row}-TIME(8,30,0)))`,
  leaveEarlier: `=IF(OR(G${row}="",H${row}="",${weekday}>5),"",MAX(0,IF(G${row}<=TIME(8,0,0),IF(${weekday}=5,TIME(17,0,0),TIME(16,30,0)),IF(G${row}<=TIME(8,15,0),IF(${weekday}=5,TIME(17,15,0),TIME(16,45,0)),IF(${weekday}=5,TIME(17,30,0),TIME(17,0,0))))-H${row}))`,
  overtime: `=IF(OR(H${row}="",${dayExpr}>5),"",MAX(0,H${row}-IF(${dayExpr}=5,TIME(18,0,0),TIME(17,30,0))))`,
};

const actual = {
  break: m.buildBreakFormula(row, dayExpr),
  total: m.buildTotalHoursFormula(row, dayExpr),
  tardiness: m.buildTardinessFormula(row, dayExpr),
  leaveEarlier: m.buildLeaveEarlierFormula(row, dayExpr),
  overtime: m.buildOvertimeFormula(row, dayExpr),
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
process.exit(allMatch ? 0 : 1);
