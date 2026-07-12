#!/usr/bin/env node
/**
 * Unit tests for duty-stats.js counting rules.
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

const code = readFileSync(join(root, 'duty-stats.js'), 'utf8');
// eslint-disable-next-line no-eval
eval(code);

const DS = globalThis.DutyStats;

function assert(label, cond) {
  if (!cond) throw new Error(`FAIL: ${label}`);
  console.log(`ok: ${label}`);
}

function item(role, extras = {}) {
  return { personIndex: 0, role, ...extras };
}

const schedule = [
  item('RC', { personIndex: 0 }),
  item('RC', { personIndex: 0 }),
  item('BG', { personIndex: 0 }),
  item('接機RC', { personIndex: 0 }),
  item('接機C', { personIndex: 0 }),
  item('PRE', { personIndex: 0 }),
  item('ABG', { personIndex: 0 }),
  item('DIC', { personIndex: 0, dicKind: 'DIC' }),
  item('TC', { personIndex: 0 }),
  item('大件', { personIndex: 0 }),
  item('休', { personIndex: 0 }),
  item('其他', { personIndex: 0, otherKind: 'S' }),
  item('其他', { personIndex: 0, otherKind: 'TA' }),
  item('其他', { personIndex: 0, otherKind: '特4hr' }),
  item('其他', { personIndex: 0, otherKind: '補4hr' }),
  item('其他', { personIndex: 0, otherKind: 'OJT' }),
  item('其他', { personIndex: 0, otherKind: 'B/F' }),
  item('其他', { personIndex: 0, otherKind: '設備' }),
  item('其他', { personIndex: 0, otherKind: '關單' }),
  item('其他', { personIndex: 0, otherKind: '備品' }),
  item('其他', { personIndex: 0, otherKind: '值日生' }),
  item('其他', { personIndex: 0, otherKind: '漏查單' }),
  item('其他', { personIndex: 0, otherKind: '抄Load' }),
  item('PPT', { personIndex: 0 }),
  item('PPT', { personIndex: 0 }),
  item('T', { personIndex: 0, tMode: 'staff', tKind: 'T1' }),
  item('RC', { personIndex: 1 }),
  item('BG', { personIndex: 1 }),
  item('BG', { personIndex: 1 })
];

const stats0 = DS.countPersonStats(schedule, 0);
assert('RC only counts role RC', stats0.rc === 2);
assert('BG only counts role BG', stats0.bg === 1);
assert('PPT counts role PPT', stats0.ppt === 2);
assert('T counts role T', stats0.t === 1);
assert('duty includes 接機RC and 接機C', stats0.duty === 14);
// 14 = prior 11 + 2 PPT + 1 T

const stats1 = DS.countPersonStats(schedule, 1);
assert('second person RC', stats1.rc === 1);
assert('second person BG', stats1.bg === 2);
assert('second person duty', stats1.duty === 3);
assert('second person ppt/t zero', stats1.ppt === 0 && stats1.t === 0);

assert('接機RC not RC count', !DS.isRcCountItem(item('接機RC')));
assert('接機C not BG count', !DS.isBgCountItem(item('接機C')));
assert('接機RC in duty total', DS.isDutyTotalItem(item('接機RC')));
assert('休 excluded', DS.isExcludedFromDutyTotal(item('休')));
assert('S excluded', DS.isExcludedFromDutyTotal(item('其他', { otherKind: 'S' })));
assert('抄Load included', DS.isDutyTotalItem(item('其他', { otherKind: '抄Load' })));

const map = DS.buildStatsMap(schedule, [{ name: 'A' }, { name: 'B' }]);
assert('buildStatsMap size', map.size === 2);
assert('buildStatsMap person 0', map.get(0).duty === 14);
assert('buildStatsMap person 0 ppt', map.get(0).ppt === 2);
assert('buildStatsMap person 1', map.get(1).bg === 2);

globalThis.document = {
  getElementById() { return null; }
};
const levelRank = level => {
  const order = ['M', 'm', '2A', '2A*', '1A', '學生', '見習'];
  const idx = order.indexOf(String(level || '').trim());
  return idx >= 0 ? idx : 1000;
};
DS.init({
  onChange() {},
  getPersonLevelRank(person) {
    return levelRank(person?.level);
  }
});
const row = idx => ({ originalIndex: idx, person: {}, isPinDivider: false });
const statsMap2 = new Map([
  [0, { rc: 2, bg: 1, ppt: 0, t: 1, duty: 4 }],
  [1, { rc: 1, bg: 3, ppt: 2, t: 0, duty: 2 }],
  [2, { rc: 0, bg: 0, ppt: 1, t: 2, duty: 5 }]
]);
DS.setVisible(true, { skipNotify: true, force: true });
DS.clearQuery({ skipNotify: true });
const baseRows = [row(0), row(1), row(2)];

DS.setFilter({ dutyMin: 4 });
assert('filter duty >= 4', DS.applyToDisplayRows(baseRows, statsMap2).map(r => r.originalIndex).join(',') === '0,2');
DS.clearQuery({ skipNotify: true });

DS.setSort('bg', 'desc');
assert('sort bg desc', DS.applyToDisplayRows(baseRows, statsMap2).map(r => r.originalIndex).join(',') === '1,0,2');
DS.setSort('bg', 'asc');
assert('sort bg asc', DS.applyToDisplayRows(baseRows, statsMap2).map(r => r.originalIndex).join(',') === '2,0,1');
DS.clearQuery({ skipNotify: true });

DS.setFilter({ pptMin: 2 });
assert('filter ppt >= 2', DS.applyToDisplayRows(baseRows, statsMap2).map(r => r.originalIndex).join(',') === '1');
DS.clearQuery({ skipNotify: true });

DS.setSort('t', 'desc');
assert('sort t desc', DS.applyToDisplayRows(baseRows, statsMap2).map(r => r.originalIndex).join(',') === '2,0,1');
DS.clearQuery({ skipNotify: true });

DS.setFilter({ rcMin: 1, bgMin: 2 });
assert('combined filter', DS.applyToDisplayRows(baseRows, statsMap2).map(r => r.originalIndex).join(',') === '1');
assert('hasActiveQuery when filter set', DS.hasActiveQuery());

const levelRows = [
  { originalIndex: 0, person: { level: '2A' }, isPinDivider: false },
  { originalIndex: 1, person: { level: 'M' }, isPinDivider: false },
  { originalIndex: 2, person: { level: 'm' }, isPinDivider: false }
];
const levelStats = new Map([
  [0, { rc: 0, bg: 0, ppt: 0, t: 0, duty: 1 }],
  [1, { rc: 0, bg: 0, ppt: 0, t: 0, duty: 1 }],
  [2, { rc: 1, bg: 0, ppt: 0, t: 0, duty: 1 }]
]);
DS.setVisible(true, { skipNotify: true, force: true });
DS.clearQuery({ skipNotify: true });
DS.setSort('rc', 'asc');
assert('secondary level desc after rc asc', DS.applyToDisplayRows(levelRows, levelStats).map(r => r.originalIndex).join(',') === '1,0,2');
DS.toggleLevelSort();
assert('secondary level asc after rc asc', DS.applyToDisplayRows(levelRows, levelStats).map(r => r.originalIndex).join(',') === '0,1,2');

DS.setVisible(false, { skipNotify: true, force: true });
assert('hidden stats skips query', !DS.hasActiveQuery());
assert('hidden stats keeps rows', DS.applyToDisplayRows(baseRows, statsMap2).length === 3);

console.log('\nAll duty-stats tests passed.');
