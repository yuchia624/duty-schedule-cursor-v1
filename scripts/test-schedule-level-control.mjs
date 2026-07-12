#!/usr/bin/env node
/**
 * Unit tests for schedule-level-control.js
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

const LEVEL_ORDER = ['M', 'm', '2A', '2A*', '1A', '學生', '見習'];

function getScheduleBlockALevelRank(levelRaw) {
  const raw = String(levelRaw || '').trim();
  if (!raw) return 1000;
  const exactIdx = LEVEL_ORDER.indexOf(raw);
  if (exactIdx >= 0) return exactIdx;
  if (/^2A\s*\*$/i.test(raw)) return LEVEL_ORDER.indexOf('2A*');
  const stripped = raw.replace(/[\s*!\-]/g, '');
  if (/^2A$/i.test(stripped) && !raw.includes('*')) return LEVEL_ORDER.indexOf('2A');
  if (/^1A$/i.test(stripped)) return LEVEL_ORDER.indexOf('1A');
  const key = stripped === '见習' ? '見習' : stripped;
  const keyRank = { M: 0, m: 1, '2A': 2, '1A': 4, '學生': 5, '見習': 6 };
  if (Object.prototype.hasOwnProperty.call(keyRank, key)) return keyRank[key];
  return 1000;
}

function normalizeStaffLevelFilterKey(levelRaw) {
  const raw = String(levelRaw || '').trim();
  if (!raw) return '';
  const stripped = raw.replace(/[\s*!\-]/g, '');
  if (stripped === '见習' || stripped === '见习') return '見習';
  if (stripped === '学生') return '學生';
  if (stripped === '見習DIC' || stripped.toUpperCase() === 'TRAINEEDIC') return 'DIC';
  if (stripped === 'M') return 'M';
  if (stripped === 'm') return 'm';
  const upper = stripped.toUpperCase();
  if (upper === '1A' || upper === '2A') return upper;
  return stripped;
}

const code = readFileSync(join(root, 'schedule-level-control.js'), 'utf8');
// eslint-disable-next-line no-eval
eval(code);

const SLC = globalThis.ScheduleLevelControl;

function assert(label, cond) {
  if (!cond) throw new Error(`FAIL: ${label}`);
  console.log(`ok: ${label}`);
}

globalThis.document = {
  getElementById() { return null; },
  querySelectorAll() { return []; }
};

SLC.init({
  onChange() {},
  getPersonLevelRank(person) {
    return getScheduleBlockALevelRank(person?.level);
  },
  getPersonLevelKey(person) {
    return normalizeStaffLevelFilterKey(person?.level);
  }
});

const row = (idx, level) => ({
  originalIndex: idx,
  person: { level },
  isPinDivider: false
});

SLC.applyMenuAction('filter-exact-2A');
assert('exact 2A matches 2A*', SLC.personMatchesLevelFilter({ level: '2A*' }) === true);
assert('exact 2A rejects M', SLC.personMatchesLevelFilter({ level: 'M' }) === false);

const exactRows = SLC.applyToDisplayRows([
  row(0, '2A'),
  row(1, '2A*'),
  row(2, 'M'),
  row(3, '1A')
]);
assert('filter exact 2A rows', exactRows.map(r => r.originalIndex).join(',') === '0,1');

SLC.clearAll({ skipNotify: true });
SLC.applyMenuAction('filter-min-1A');
const min1aRows = SLC.applyToDisplayRows([
  row(0, 'M'),
  row(1, '2A*'),
  row(2, '1A'),
  row(3, '學生')
]);
assert('filter min 1A rows', min1aRows.map(r => r.originalIndex).join(',') === '0,1,2');

SLC.clearAll({ skipNotify: true });
SLC.applyMenuAction('filter-min-2A');
const min2aRows = SLC.applyToDisplayRows([
  row(0, 'M'),
  row(1, '2A'),
  row(2, '2A*'),
  row(3, '1A')
]);
assert('filter min 2A rows', min2aRows.map(r => r.originalIndex).join(',') === '0,1,2');

SLC.clearAll({ skipNotify: true });
SLC.applyMenuAction('filter-min-m');
const minMRows = SLC.applyToDisplayRows([
  row(0, 'M'),
  row(1, 'm'),
  row(2, '2A')
]);
assert('filter min m rows', minMRows.map(r => r.originalIndex).join(',') === '0,1');

SLC.clearAll({ skipNotify: true });
SLC.applyMenuAction('filter-exact-M');
assert('filter exact M', SLC.applyToDisplayRows([
  row(0, 'M'),
  row(1, 'm'),
  row(2, '2A')
]).map(r => r.originalIndex).join(',') === '0');

SLC.clearAll({ skipNotify: true });
SLC.applyMenuAction('filter-exact-student');
assert('filter exact student', SLC.applyToDisplayRows([
  row(0, '學生'),
  row(1, '見習'),
  row(2, '1A')
]).map(r => r.originalIndex).join(',') === '0');

SLC.clearAll({ skipNotify: true });
SLC.applyMenuAction('filter-exact-DIC');
assert('filter exact DIC', SLC.applyToDisplayRows([
  row(0, 'DIC'),
  row(1, '見習DIC'),
  row(2, 'M')
]).map(r => r.originalIndex).join(',') === '0,1');

SLC.clearAll({ skipNotify: true });
SLC.applyMenuAction('sort-desc');
const sortRows = SLC.applyToDisplayRows([
  row(0, '2A'),
  row(1, 'M'),
  row(2, '1A')
]);
assert('sort desc by level', sortRows.map(r => r.originalIndex).join(',') === '1,0,2');

console.log('\nAll schedule-level-control tests passed.');
