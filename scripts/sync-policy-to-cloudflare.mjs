// Convert shared/policy.js (ES module with `export` statements) into
// a browser script that attaches its public API to
// window.AttendancePolicy, so the Cloudflare Pages static build
// (which loads plain <script> tags) can consume it.
//
// Usage: node scripts/sync-policy-to-cloudflare.mjs
// Run after any change to shared/policy.js. The output is written
// to cloudflare-pages-attendance/public/policy.js.

import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..');
const sourcePath = resolve(repoRoot, 'shared/policy.js');
const destPath = resolve(repoRoot, 'cloudflare-pages-attendance/public/policy.js');

let body = readFileSync(sourcePath, 'utf8');

// Strip the `export ` keyword from declarations and `export { ... }`
// blocks. The shared file uses `export const X`, `export function Y`,
// and a `export { ... }` (none currently) form. We strip the keyword
// and any trailing `export { ... }` block in one pass.
body = body.replace(/^export\s+const\s+/gm, 'const ');
body = body.replace(/^export\s+function\s+/gm, 'function ');
// Strip any `export { ... }` block at the bottom (none in current
// shared/policy.js, but defensive).
body = body.replace(/^export\s*\{[\s\S]*?\};?\s*$/gm, '');

// Append the window-attachment block. Only attach in a browser context
// (so the same file is harmless to load in Node for testing).
const footer = `

// Cloudflare Pages loads this file as a plain <script> (not a module),
// so the public API is attached to a global rather than imported.
// window.AttendancePolicy mirrors the named exports of shared/
// policy.js — keep the two in sync.
if (typeof window !== 'undefined') {
  window.AttendancePolicy = {
    MONTH_LOOKUP,
    BREAK_MON_THU,
    BREAK_FRI,
    STANDARD_CLOCKIN_END,
    FLEXI_1_START,
    FLEXI_1_END,
    FLEXI_2_START,
    FLEXI_2_END,
    LATE_THRESHOLD,
    STANDARD_CLOCKOUT_MON_THU,
    STANDARD_CLOCKOUT_FRI,
    FLEXI_1_CLOCKOUT_MON_THU,
    FLEXI_1_CLOCKOUT_FRI,
    FLEXI_2_CLOCKOUT_MON_THU,
    FLEXI_2_CLOCKOUT_FRI,
    OVERTIME_START_MON_THU,
    OVERTIME_START_FRI,
    isFriday,
    isWeekend,
    determineFlexiType,
    getBreakWindow,
    getAllowedClockOut,
    getOvertimeThreshold,
    buildBreakFormula,
    buildTotalHoursFormula,
    buildTardinessFormula,
    buildLeaveEarlierFormula,
    buildOvertimeFormula,
    POLICY_DESCRIPTIONS,
  };
}
`;

writeFileSync(destPath, body + footer, 'utf8');
console.log(`Wrote ${destPath} (${body.split('\n').length} body lines + ${footer.split('\n').length} footer lines)`);
