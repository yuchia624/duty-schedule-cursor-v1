#!/usr/bin/env node
/**
 * Unit tests for schedule-diff-core.js
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

const code = readFileSync(join(root, 'schedule-diff-core.js'), 'utf8');
// eslint-disable-next-line no-eval
eval(code);

const Core = globalThis.ScheduleDiffCore;

function assert(label, cond) {
  if (!cond) throw new Error(`FAIL: ${label}`);
  console.log(`ok: ${label}`);
}

function item(uid, personIndex, start, end, extras = {}) {
  return {
    uid,
    personIndex,
    start,
    end,
    role: 'RC',
    flight: '385A',
    compactLabel: '385A',
    ...extras
  };
}

const helpers = {
  getLabel: (i) => i.compactLabel,
  getPersonName: (idx) => ['Alice', 'Bob', 'Carol'][idx] || `#${idx}`,
  getPersonSeq: (idx) => [12, 5, 8][idx] ?? null
};

const base = [
  item(1, 0, 800, 840),
  item(2, 0, 900, 940),
  item(3, 1, 820, 860, { role: 'BG', flight: '160A', compactLabel: '160A' }),
  item(4, 2, 700, 740, { role: '休', flight: '', compactLabel: '休' })
];

const curr = [
  item(1, 1, 800, 840),
  item(2, 0, 920, 960),
  item(3, 1, 820, 860, { role: 'BG', flight: '160A', compactLabel: '160A' }),
  item(5, 0, 1000, 1040, { role: 'PPT', flight: '67P', compactLabel: '67P' })
];

const peopleKeys = ['alice', 'bob', 'carol'];
const changes = Core.diffSchedule(Core.slimSchedule(base), Core.slimSchedule(curr), {
  getPersonKey: (idx) => peopleKeys[idx] || String(idx)
});

assert('detects moved person', changes.some(c => c.type === 'changed' && c.uid === 1 && c.personChanged));
assert('same uid time change shows 改時', changes.some(c => c.type === 'changed' && c.uid === 2 && c.timeChanged && !c.personChanged));
assert('unchanged block omitted', !changes.some(c => c.uid === 3));
assert('detects added block', changes.some(c => c.type === 'added' && c.uid === 5));
assert('detects removed block', changes.some(c => c.type === 'removed' && c.uid === 4));

const movedPerson = changes.find(c => c.uid === 1);
const summary = Core.formatChangeSummary(movedPerson, helpers);
assert('summary kind is 移人', summary.kindLabel === '移人');
assert('summary shows seq in detail', summary.detail.includes('序12') && summary.detail.includes('序5'));
assert('summary seq column uses curr person', summary.seq === 5);

const affected = Core.collectAffectedPersonIndices(changes);
assert('affected includes moved from/to', affected.has(0) && affected.has(1) && affected.has(2));

const rows = [
  { originalIndex: 0, pinned: false },
  { originalIndex: 1, pinned: false },
  { originalIndex: 2, pinned: false },
  { originalIndex: 3, pinned: false }
];
const sorted = Core.applyDiffRowSort(rows, new Set([2, 0]));
assert('changed rows bubble up', sorted[0].originalIndex === 0 && sorted[1].originalIndex === 2);

const exportRows = Core.buildExportSheetRows(changes.slice(0, 1), helpers, {
  baselineTime: '16:00',
  dateLabel: '2026/07/22',
  exportTime: '17:00'
});
assert('export has header row', exportRows[5][0] === '類型');
assert('export row includes seq', exportRows[6][1] === 5);

const timeOnly = Core.diffSchedule(
  [Core.slimScheduleItem(item(20, 0, 800, 840))],
  [Core.slimScheduleItem(item(20, 0, 920, 960))],
  { getPersonKey: (idx) => String(idx) }
);
assert('explicit same-person time-only shows 改時', timeOnly.length === 1 && timeOnly[0].timeChanged);

const mixedKeyRoundTrip = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(41, 2, 1320, 1380, {
    role: '其他', otherKind: 'TA', compactLabel: 'TA', label: 'TA'
  })), { personKey: '2', dutyKey: '其他:TA' })],
  [Object.assign(Core.slimScheduleItem(item(99, 2, 1320, 1380, {
    role: '其他', otherKind: 'TA', compactLabel: 'TA', label: 'TA'
  })), { personKey: '2', dutyKey: 'TA' })],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey }
);
assert('mixed duty key delete re-add same slot nets zero', mixedKeyRoundTrip.length === 0);

const mixedKeyTimeShift = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(42, 2, 1320, 1380, {
    role: '其他', otherKind: 'TA', compactLabel: 'TA', label: 'TA'
  })), { personKey: '2', dutyKey: '其他:TA' })],
  [Object.assign(Core.slimScheduleItem(item(99, 2, 1330, 1390, {
    role: '其他', otherKind: 'TA', compactLabel: 'TA', label: 'TA'
  })), { personKey: '2', dutyKey: 'TA' })],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey }
);
assert('mixed duty key time shift becomes 改時', mixedKeyTimeShift.length === 1
  && mixedKeyTimeShift[0].type === 'changed'
  && mixedKeyTimeShift[0].timeChanged);

const restTimeShift = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(50, 2, 1320, 1380, {
    role: '休', flight: '', compactLabel: '休', label: '休'
  })), { personKey: '2', dutyKey: '休' })],
  [Object.assign(Core.slimScheduleItem(item(50, 2, 1330, 1390, {
    role: '休', flight: '', compactLabel: '休', label: '休'
  })), { personKey: '2', dutyKey: '休' })],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey }
);
assert('休 same-person time shift hidden', restTimeShift.length === 0);

const preTimeShift = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(51, 2, 800, 860, {
    role: 'PRE', flight: 'PRE1', compactLabel: 'PRE', label: 'PRE'
  })), { personKey: '2', dutyKey: 'PRE:PRE1:PRE' })],
  [Object.assign(Core.slimScheduleItem(item(51, 2, 820, 880, {
    role: 'PRE', flight: 'PRE1', compactLabel: 'PRE', label: 'PRE'
  })), { personKey: '2', dutyKey: 'PRE:PRE1:PRE' })],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey }
);
assert('PRE same-person time shift hidden', preTimeShift.length === 0);

const restDeleteReaddTime = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(52, 2, 1320, 1380, {
    role: '休', flight: '', compactLabel: '休', label: '休'
  })), { personKey: '2', dutyKey: '休' })],
  [Object.assign(Core.slimScheduleItem(item(99, 2, 1330, 1390, {
    role: '休', flight: '', compactLabel: '休', label: '休'
  })), { personKey: '2', dutyKey: '休' })],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey }
);
assert('休 delete re-add time shift hidden', restDeleteReaddTime.length === 0);

const moveWithTime = Core.diffSchedule(
  [Core.slimScheduleItem(item(21, 0, 800, 840))],
  [Core.slimScheduleItem(item(21, 1, 920, 960))],
  { getPersonKey: (idx) => String(idx) }
);
assert('person move with time still reported', moveWithTime.length === 1 && moveWithTime[0].personChanged);

const roundTrip = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(30, 2, 1320, 1380, {
    role: '其他', flight: 'OTHER_TA', compactLabel: 'TA', label: 'TA', otherKind: 'TA'
  })), { personKey: '2', dutyKey: 'TA' })],
  [Object.assign(Core.slimScheduleItem(item(99, 2, 1320, 1380, {
    role: '其他', flight: 'OTHER_TA', compactLabel: 'TA', label: 'TA', otherKind: 'TA'
  })), { personKey: '2', dutyKey: 'TA' })],
  {
    getPersonKey: (idx) => String(idx),
    getDutyKey: (it) => it.dutyKey || 'TA'
  }
);
assert('delete and re-add same slot nets to zero', roundTrip.length === 0);

assert('TA items equivalent via dutyKey', Core.dutyItemsEquivalent(
  { personIndex: 2, start: 1320, end: 1380, personKey: 'huang', dutyKey: 'TA' },
  { personIndex: 2, start: 1320, end: 1380, personKey: 'huang', dutyKey: 'TA' },
  { getPersonKey: () => 'other' }
));

const sameUidTime = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(20, 0, 800, 840)), { personKey: '0', dutyKey: '385A' })],
  [Object.assign(Core.slimScheduleItem(item(20, 0, 920, 960)), { personKey: '0', dutyKey: '385A' })],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey || '385A' }
);
assert('same uid same-person time shows 改時', sameUidTime.length === 1
  && sameUidTime[0].type === 'changed'
  && sameUidTime[0].timeChanged);

assert('mixed duty keys still match', Core.dutyKeysMatch('TA', '其他:TA'));

const timeShift = Core.diffSchedule(
  [Object.assign(Core.slimScheduleItem(item(40, 2, 1320, 1380, {
    role: '其他', otherKind: 'TA', compactLabel: 'TA', label: 'TA'
  })), { personKey: '2', dutyKey: 'TA' })],
  [Object.assign(Core.slimScheduleItem(item(99, 2, 1330, 1390, {
    role: '其他', otherKind: 'TA', compactLabel: 'TA', label: 'TA'
  })), { personKey: '2', dutyKey: 'TA' })],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey }
);
assert('same person duty time shift becomes 改時', timeShift.length === 1
  && timeShift[0].type === 'changed'
  && timeShift[0].timeChanged
  && !timeShift[0].personChanged);
const timeSummary = Core.formatChangeSummary(timeShift[0], helpers);
assert('time shift summary label', timeSummary.kindLabel === '改時');
assert('time shift summary has range', timeSummary.detail.includes('22:00') && timeSummary.detail.includes('22:10'));

const partialRoundTrip = Core.diffSchedule(
  [
    Object.assign(Core.slimScheduleItem(item(31, 0, 800, 840)), { personKey: '0', dutyKey: '385A' }),
    Object.assign(Core.slimScheduleItem(item(32, 1, 900, 940, { role: 'BG', flight: '160A', compactLabel: '160A' })), { personKey: '1', dutyKey: '160A' })
  ],
  [
    Object.assign(Core.slimScheduleItem(item(88, 0, 800, 840)), { personKey: '0', dutyKey: '385A' }),
    Object.assign(Core.slimScheduleItem(item(32, 1, 900, 940, { role: 'BG', flight: '160A', compactLabel: '160A' })), { personKey: '1', dutyKey: '160A' })
  ],
  { getPersonKey: (idx) => String(idx), getDutyKey: (it) => it.dutyKey }
);
assert('only unmatched add/remove remain', partialRoundTrip.length === 0);

const contentChange = Core.diffSchedule(
  [Core.slimScheduleItem(item(10, 0, 800, 840, { coveredFlights: ['A'] }))],
  [Core.slimScheduleItem(item(10, 0, 800, 840, { coveredFlights: ['B'] }))],
  { getPersonKey: (idx) => String(idx) }
);
assert('detects coveredFlights change', contentChange.length === 1 && contentChange[0].contentChanged);

console.log('\nAll schedule-diff-core tests passed.');
