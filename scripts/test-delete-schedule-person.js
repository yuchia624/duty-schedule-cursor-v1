#!/usr/bin/env node
/**
 * E2E-style tests for schedule person delete index safety.
 * Logic mirrors index.html: getPersonDeleteToken, personMatchesDeleteToken,
 * resolvePersonDeleteIndex, applySchedulePersonDelete, display order reindex.
 */
'use strict';

function normalizeText(v) {
  return String(v || '').trim().toLowerCase().replace(/\s+/g, '');
}
function isTempControlRow(p) {
  return !!p?.isTempRow;
}

function tempRowIdentityKeys(person) {
  if (!isTempControlRow(person)) return [];
  const keys = [];
  const code = normalizeText(person.code);
  if (code) keys.push(code);
  const name = String(person.name || '').trim();
  if (name) keys.push(normalizeText(name));
  return keys;
}

function tempRowDisplayLabelTaken(people, label, excludeIndex) {
  const key = normalizeText(String(label || '').trim());
  if (!key) return false;
  return people.some((p, idx) => {
    if (idx === excludeIndex || !isTempControlRow(p)) return false;
    return tempRowIdentityKeys(p).includes(key);
  });
}

function syncTempRowIdentity(people, personIndex, name) {
  const person = people[personIndex];
  if (!person?.isTempRow) return { ok: true };
  const v = String(name ?? '').trim();
  if (v && tempRowDisplayLabelTaken(people, v, personIndex)) {
    return { ok: false, reason: 'duplicate' };
  }
  person.name = v;
  person.personnelId = '';
  return { ok: true };
}

function nextTempControlRowCode(people) {
  let max = 0;
  const usedCodes = new Set();
  people.forEach(p => {
    if (!isTempControlRow(p)) return;
    const code = normalizeText(p.code);
    if (code) usedCodes.add(code);
    const m = String(p.code || p.name || '').match(/^CTR(\d+)$/i);
    if (m) max = Math.max(max, Number(m[1]) || 0);
  });
  let n = max + 1;
  let candidate = `CTR${n}`;
  while (usedCodes.has(normalizeText(candidate)) || tempRowDisplayLabelTaken(people, candidate, -1)) {
    n += 1;
    candidate = `CTR${n}`;
  }
  return candidate;
}

function getPersonDeleteToken(person, resolvePersonnelId, resolveFullName) {
  if (!person) return null;
  if (isTempControlRow(person)) {
    return { kind: 'temp', code: String(person.code || person.name || '').trim() };
  }
  const personnelId = String(person.personnelId || '').trim() || resolvePersonnelId(person);
  if (personnelId) return { kind: 'formal', personnelId };
  return {
    kind: 'formal',
    code: String(person.code || '').trim(),
    name: String(resolveFullName(person) || person.name || person.code || '').trim()
  };
}

function personMatchesDeleteToken(person, token, resolvePersonnelId, resolveFullName) {
  if (!person || !token) return false;
  if (token.kind === 'temp') {
    if (!isTempControlRow(person)) return false;
    return normalizeText(person.code || person.name) === normalizeText(token.code);
  }
  if (token.kind !== 'formal' || isTempControlRow(person)) return false;
  if (token.personnelId) {
    const pid = String(resolvePersonnelId(person) || person.personnelId || '').trim();
    return pid === token.personnelId;
  }
  if (!token.code || !token.name) return false;
  const codeMatch = normalizeText(person.code) === normalizeText(token.code);
  const nameMatch = normalizeText(resolveFullName(person) || person.name || person.code) === normalizeText(token.name);
  return codeMatch && nameMatch;
}

function resolvePersonDeleteIndex(people, token, resolvePersonnelId, resolveFullName) {
  if (!token) return -1;
  return people.findIndex(p => personMatchesDeleteToken(p, token, resolvePersonnelId, resolveFullName));
}

function isValidPeopleDisplayOrder(order, peopleLen) {
  if (!Array.isArray(order) || order.length !== peopleLen) return false;
  const seen = new Set();
  for (const idx of order) {
    if (!Number.isInteger(idx) || idx < 0 || idx >= peopleLen || seen.has(idx)) return false;
    seen.add(idx);
  }
  return seen.size === peopleLen;
}

function isDicSchedulePerson(person) {
  if (!person || isTempControlRow(person)) return false;
  return String(person.level || '').trim() === 'DIC';
}

function getPersonShiftBand(person) {
  if (!person) return '';
  if (isTempControlRow(person)) {
    const shift = String(person.shift || '').trim();
    if (shift) {
      const shiftType = shift.startsWith('I') || shift === 'hM' ? 'early'
        : (shift.startsWith('h') && shift !== 'hM' ? 'late' : '');
      if (shiftType) return shiftType;
    }
    return person.tempRowView === 'late' ? 'late' : 'early';
  }
  const shift = String(person.shift || '').trim();
  return shift.startsWith('I') || shift === 'hM' ? 'early'
    : (shift.startsWith('h') && shift !== 'hM' ? 'late' : '');
}

function insertTempRowIntoDisplayOrder(people, peopleDisplayOrder, personIndex) {
  if (!Array.isArray(peopleDisplayOrder)) peopleDisplayOrder = [];
  const tempPerson = people[personIndex];
  const tempBand = getPersonShiftBand(tempPerson);
  peopleDisplayOrder = peopleDisplayOrder.filter(idx => idx !== personIndex);
  let insertAt = peopleDisplayOrder.length;
  for (let i = 0; i < peopleDisplayOrder.length; i++) {
    const p = people[peopleDisplayOrder[i]];
    if (!p || !isDicSchedulePerson(p)) continue;
    if (tempBand && getPersonShiftBand(p) !== tempBand) continue;
    insertAt = i;
    break;
  }
  if (insertAt === peopleDisplayOrder.length && tempBand) {
    let lastInBand = -1;
    for (let i = 0; i < peopleDisplayOrder.length; i++) {
      const p = people[peopleDisplayOrder[i]];
      if (p && getPersonShiftBand(p) === tempBand) lastInBand = i;
    }
    if (lastInBand >= 0) insertAt = lastInBand + 1;
  }
  while (insertAt < peopleDisplayOrder.length && isTempControlRow(people[peopleDisplayOrder[insertAt]])) {
    const otherBand = getPersonShiftBand(people[peopleDisplayOrder[insertAt]]);
    if (tempBand && otherBand !== tempBand) break;
    insertAt += 1;
  }
  peopleDisplayOrder.splice(insertAt, 0, personIndex);
  return peopleDisplayOrder;
}

function restoreTempRowsDisplayOrderBeforeDic(people, peopleDisplayOrder) {
  people.forEach((p, idx) => {
    if (isTempControlRow(p)) {
      peopleDisplayOrder = insertTempRowIntoDisplayOrder(people, peopleDisplayOrder, idx);
    }
  });
  return peopleDisplayOrder;
}

function reindexPeopleDisplayOrderAfterPersonRemoved(people, peopleDisplayOrder, removedIndex) {
  if (!Array.isArray(peopleDisplayOrder) || !peopleDisplayOrder.length) {
    return people.length ? people.map((_, i) => i) : [];
  }
  let nextOrder = peopleDisplayOrder
    .filter(idx => idx !== removedIndex)
    .map(idx => (idx > removedIndex ? idx - 1 : idx));
  const seen = new Set(nextOrder);
  for (let idx = 0; idx < people.length; idx++) {
    if (!seen.has(idx)) nextOrder.push(idx);
  }
  if (!isValidPeopleDisplayOrder(nextOrder, people.length)) {
    nextOrder = people.map((_, i) => i);
  }
  nextOrder = restoreTempRowsDisplayOrderBeforeDic(people, nextOrder);
  return nextOrder;
}

function buildAllPersonRows(people) {
  return people.map((p, idx) => ({ person: p, originalIndex: idx, matchedRow: false }));
}

function getDisplaySortedRows(people, peopleDisplayOrder) {
  const allRows = buildAllPersonRows(people);
  const rowMap = new Map(allRows.map(row => [row.originalIndex, row]));
  const ordered = [];
  const seen = new Set();
  (peopleDisplayOrder || []).forEach(idx => {
    const row = rowMap.get(idx);
    if (!row || seen.has(idx)) return;
    ordered.push(row);
    seen.add(idx);
  });
  allRows.forEach(row => {
    if (seen.has(row.originalIndex)) return;
    ordered.push(row);
    seen.add(row.originalIndex);
  });
  return ordered;
}

function getTempRowPinnedShiftView(person) {
  if (!isTempControlRow(person)) return '';
  return person.tempRowView === 'late' ? 'late' : 'early';
}

function personMatchesShiftView(person, activeShiftView) {
  if (isTempControlRow(person)) {
    const shift = String(person.shift || '').trim();
    if (shift) {
      const shiftType = shift.startsWith('I') || shift === 'hM' || shift === 'hO' ? 'early'
        : (shift.startsWith('h') && shift !== 'hM' && shift !== 'hO' ? 'late' : '');
      if (shiftType) return shiftType === activeShiftView;
    }
    return getTempRowPinnedShiftView(person) === activeShiftView;
  }
  const shift = String(person.shift || '').trim();
  const shiftType = shift.startsWith('I') || shift === 'hM' ? 'early' : '';
  return shiftType === activeShiftView;
}

function getDisplayRows(people, peopleDisplayOrder, activeShiftView = 'early') {
  return getDisplaySortedRows(people, peopleDisplayOrder).filter(row => personMatchesShiftView(row.person, activeShiftView));
}

function applySchedulePersonDelete(state, token, resolvePersonnelId, resolveFullName) {
  const { people, schedule, peopleDisplayOrder } = state;
  const personIndex = resolvePersonDeleteIndex(people, token, resolvePersonnelId, resolveFullName);
  if (personIndex < 0) return { ok: false, reason: 'not-found' };
  const person = people[personIndex];
  if (!person || !personMatchesDeleteToken(person, token, resolvePersonnelId, resolveFullName)) {
    return { ok: false, reason: 'token-mismatch' };
  }
  const tempCountBefore = people.filter(p => p.isTempRow).length;
  const removedDuties = schedule.filter(s => s.personIndex === personIndex);
  const nextSchedule = schedule
    .filter(s => s.personIndex !== personIndex)
    .map(s => (s.personIndex > personIndex ? { ...s, personIndex: s.personIndex - 1 } : { ...s }));
  const nextPeople = people.slice();
  nextPeople.splice(personIndex, 1);
  const nextOrder = reindexPeopleDisplayOrderAfterPersonRemoved(nextPeople, peopleDisplayOrder, personIndex);
  const tempCountAfter = nextPeople.filter(p => p.isTempRow).length;
  return {
    ok: true,
    deletedIndex: personIndex,
    deletedLabel: person.name || person.code,
    removedDuties,
    people: nextPeople,
    schedule: nextSchedule,
    peopleDisplayOrder: nextOrder,
    tempCountBefore,
    tempCountAfter,
    displayRows: getDisplayRows(nextPeople, nextOrder),
    sortedRows: getDisplaySortedRows(nextPeople, nextOrder)
  };
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

let passed = 0;
let failed = 0;

function test(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
    passed++;
  } catch (e) {
    console.error(`FAIL ${name}: ${e.message}`);
    failed++;
  }
}

const resolvePid = p => String(p.personnelId || '').trim();
const resolveName = p => String(p.name || p.code || '').trim();

test('delete formal A by personnelId only removes A', () => {
  const people = [
    { code: 'A1', name: '楊亞璇', personnelId: 'P001', shift: 'hM' },
    { code: 'CTR1', name: 'CTR1', isTempRow: true, shift: '' },
    { code: 'B1', name: '蕭宇謙', personnelId: 'P002', shift: 'I9' }
  ];
  const schedule = [
    { uid: 1, personIndex: 0, role: 'TC' },
    { uid: 2, personIndex: 2, role: 'TC' }
  ];
  const state = { people, schedule, peopleDisplayOrder: [0, 1, 2] };
  const token = getPersonDeleteToken(people[0], resolvePid, resolveName);
  const result = applySchedulePersonDelete(state, token, resolvePid, resolveName);
  assert(result.ok, 'delete failed');
  assert(result.people.length === 2, 'expected 2 people');
  assert(result.people.every(p => p.personnelId !== 'P001'), 'A should be gone');
  assert(result.people.some(p => p.personnelId === 'P002'), 'B should remain');
  assert(result.people.some(p => p.isTempRow && p.code === 'CTR1'), 'CTR1 should remain');
  assert(result.schedule.length === 1, 'only A duty removed');
  assert(result.schedule[0].personIndex === 1, 'B duty reindexed to 1');
  assert(isValidPeopleDisplayOrder(result.peopleDisplayOrder, 2), 'display order valid');
});

test('delete formal preserves CTR1 in people[] display and duty', () => {
  const people = [
    { code: 'A1', name: '甲', personnelId: 'PA', shift: 'hM', level: 'M' },
    { code: 'CTR1', name: '支援小王', isTempRow: true, level: '支援', shift: 'IB' },
    { code: 'B1', name: '乙', personnelId: 'PB', shift: 'hM', level: 'M' }
  ];
  let peopleDisplayOrder = [0, 1, 2];
  peopleDisplayOrder = insertTempRowIntoDisplayOrder(people, peopleDisplayOrder, 1);
  const schedule = [
    { uid: 10, personIndex: 0, role: 'TC' },
    { uid: 11, personIndex: 1, role: 'TC', start: 100, end: 200 },
    { uid: 12, personIndex: 2, role: 'TC' }
  ];
  const token = { kind: 'formal', personnelId: 'PA' };
  const result = applySchedulePersonDelete(
    { people, schedule, peopleDisplayOrder },
    token,
    resolvePid,
    resolveName
  );
  assert(result.ok, 'delete failed');
  assert(result.tempCountBefore === 1 && result.tempCountAfter === 1, 'temp row count unchanged');
  assert(result.people.some(p => p.isTempRow && p.code === 'CTR1'), 'CTR1 still in people[]');
  assert(
    result.displayRows.some(r => r.person.isTempRow && r.person.code === 'CTR1'),
    'CTR1 still in getDisplayRows()'
  );
  assert(
    result.sortedRows.some(r => r.person.isTempRow && r.person.code === 'CTR1'),
    'CTR1 still in getDisplaySortedRows()'
  );
  const ctrIdx = result.people.findIndex(p => p.isTempRow && p.code === 'CTR1');
  assert(
    result.schedule.some(s => s.uid === 11 && s.personIndex === ctrIdx),
    'CTR1 duty preserved with reindexed personIndex'
  );
  assert(result.people.some(p => p.personnelId === 'PB'), 'formal B remains');
});

test('delete first formal when CTR inserted before DIC keeps CTR visible', () => {
  const people = [
    { code: 'A1', name: '甲', personnelId: 'PA', shift: 'hM' },
    { code: 'CTR1', name: '', isTempRow: true, level: '支援', shift: '' },
    { code: 'D1', name: 'Dic', personnelId: 'PD', shift: 'hM', level: 'DIC' }
  ];
  let peopleDisplayOrder = [0, 1, 2];
  peopleDisplayOrder = insertTempRowIntoDisplayOrder(people, peopleDisplayOrder, 1);
  const result = applySchedulePersonDelete(
    { people, schedule: [{ uid: 1, personIndex: 1, role: 'TC' }], peopleDisplayOrder },
    { kind: 'formal', personnelId: 'PA' },
    resolvePid,
    resolveName
  );
  assert(result.ok, 'delete failed');
  assert(result.people.length === 2, 'CTR + DIC remain');
  assert(result.people[0].isTempRow && result.people[0].code === 'CTR1', 'CTR1 first in people[]');
  assert(result.displayRows.some(r => r.person.code === 'CTR1'), 'CTR1 visible after delete');
  assert(result.schedule.length === 1 && result.schedule[0].personIndex === 0, 'CTR duty reindexed to 0');
});

test('delete temp CTR1 only removes CTR1', () => {
  const people = [
    { code: 'A1', name: '楊亞璇', personnelId: 'P001' },
    { code: 'CTR1', name: 'CTR1', isTempRow: true },
    { code: 'B1', name: '蕭宇謙', personnelId: 'P002' }
  ];
  const state = { people, schedule: [{ uid: 1, personIndex: 1 }], peopleDisplayOrder: [0, 1, 2] };
  const token = getPersonDeleteToken(people[1], resolvePid, resolveName);
  const result = applySchedulePersonDelete(state, token, resolvePid, resolveName);
  assert(result.ok, 'delete failed');
  assert(result.people.length === 2, 'expected 2 people');
  assert(result.people[0].personnelId === 'P001', 'above person remains');
  assert(result.people[1].personnelId === 'P002', 'below person remains');
  assert(result.schedule.length === 0, 'temp duty removed');
});

test('delete formal without personnelId uses code+name', () => {
  const people = [
    { code: 'X9', name: '王小明' },
    { code: 'CTR2', name: 'CTR2', isTempRow: true }
  ];
  const token = { kind: 'formal', code: 'X9', name: '王小明' };
  const result = applySchedulePersonDelete(
    { people, schedule: [], peopleDisplayOrder: [0, 1] },
    token,
    resolvePid,
    resolveName
  );
  assert(result.ok, 'delete failed');
  assert(result.people.length === 1 && result.people[0].code === 'CTR2', 'only temp remains');
});

test('sequential delete two different people does not delete wrong person', () => {
  let people = [
    { code: 'A1', name: '甲', personnelId: 'PA' },
    { code: 'CTR1', name: 'CTR1', isTempRow: true },
    { code: 'B1', name: '乙', personnelId: 'PB' }
  ];
  let schedule = [
    { uid: 1, personIndex: 0 },
    { uid: 2, personIndex: 1 },
    { uid: 3, personIndex: 2 }
  ];
  let peopleDisplayOrder = [0, 1, 2];

  const tokenTemp = getPersonDeleteToken(people[1], resolvePid, resolveName);
  let r1 = applySchedulePersonDelete({ people, schedule, peopleDisplayOrder }, tokenTemp, resolvePid, resolveName);
  assert(r1.ok, 'first delete failed');
  people = r1.people;
  schedule = r1.schedule;
  peopleDisplayOrder = r1.peopleDisplayOrder;

  assert(people.length === 2, 'after first delete 2 remain');
  assert(people[0].personnelId === 'PA', '甲 remains');
  assert(people[1].personnelId === 'PB', '乙 remains');
  assert(schedule.length === 2, 'two duties remain');
  assert(schedule.every(s => s.personIndex === 0 || s.personIndex === 1), 'duty indices valid');

  const tokenB = { kind: 'formal', personnelId: 'PB' };
  let r2 = applySchedulePersonDelete({ people, schedule, peopleDisplayOrder }, tokenB, resolvePid, resolveName);
  assert(r2.ok, 'second delete failed');
  assert(r2.people.length === 1 && r2.people[0].personnelId === 'PA', 'only 甲 remains');
  assert(r2.schedule.length === 1 && r2.schedule[0].personIndex === 0, 'only 甲 duty remains');
});

test('delete person with duty does not remove other duties', () => {
  const people = [
    { code: 'A1', name: '甲', personnelId: 'PA' },
    { code: 'B1', name: '乙', personnelId: 'PB' }
  ];
  const schedule = [
    { uid: 10, personIndex: 0, role: 'TC', start: 100, end: 200 },
    { uid: 11, personIndex: 1, role: 'TC', start: 300, end: 400 }
  ];
  const token = { kind: 'formal', personnelId: 'PA' };
  const result = applySchedulePersonDelete(
    { people, schedule, peopleDisplayOrder: [0, 1] },
    token,
    resolvePid,
    resolveName
  );
  assert(result.ok, 'delete failed');
  assert(result.removedDuties.length === 1 && result.removedDuties[0].uid === 10, 'only A duty removed');
  assert(result.schedule.length === 1 && result.schedule[0].uid === 11, 'B duty remains');
  assert(result.schedule[0].personIndex === 0, 'B duty reindexed to 0');
});

test('token resolve ignores stale array index (index shift simulation)', () => {
  const people = [
    { code: 'A1', name: '甲', personnelId: 'PA' },
    { code: 'CTR1', name: 'CTR1', isTempRow: true }
  ];
  const token = { kind: 'temp', code: 'CTR1' };
  assert(resolvePersonDeleteIndex(people, token, resolvePid, resolveName) === 1, 'CTR1 at index 1');
  const after = applySchedulePersonDelete(
    { people, schedule: [], peopleDisplayOrder: [0, 1] },
    token,
    resolvePid,
    resolveName
  );
  assert(after.ok && after.people.length === 1 && after.people[0].personnelId === 'PA', 'only temp removed');
  assert(resolvePersonDeleteIndex(after.people, token, resolvePid, resolveName) === -1, 'token no longer resolves');
});

test('temp token does not match formal person with same code text', () => {
  const people = [{ code: 'CTR1', name: 'CTR1', personnelId: 'PX' }];
  const token = { kind: 'temp', code: 'CTR1' };
  assert(resolvePersonDeleteIndex(people, token, resolvePid, resolveName) === -1, 'must not match formal');
});

test('late shift temp row inserts before late DIC not early DIC', () => {
  const people = [
    { code: 'E1', name: '早班人', shift: 'hM', level: 'M' },
    { code: 'ED', name: '早DIC', shift: 'hM', level: 'DIC' },
    { code: 'L1', name: '晚班人', shift: 'hN', level: 'M' },
    { code: 'LD', name: '晚DIC', shift: 'hN', level: 'DIC' }
  ];
  const ctrIdx = 4;
  people.push({ code: 'CTR1', name: '', isTempRow: true, shift: 'hN', tempRowView: 'late', level: '支援' });
  let order = [0, 1, 2, 3];
  order = insertTempRowIntoDisplayOrder(people, order, ctrIdx);
  const ctrPos = order.indexOf(ctrIdx);
  const lateDicPos = order.indexOf(3);
  const earlyDicPos = order.indexOf(1);
  assert(ctrPos < lateDicPos, 'late CTR should be before late DIC');
  assert(ctrPos > earlyDicPos, 'late CTR should stay after early DIC block');
});

test('temp row with early shift only shows in early view', () => {
  const people = [
    { code: 'CTR1', name: '', isTempRow: true, level: '支援', shift: 'IB', tempRowView: 'early' }
  ];
  assert(personMatchesShiftView(people[0], 'early'), 'IB temp visible in early');
  assert(!personMatchesShiftView(people[0], 'late'), 'IB temp hidden in late');
});

test('temp row with empty shift only shows in pinned view', () => {
  const earlyOnly = { code: 'CTR1', name: '', isTempRow: true, shift: '', tempRowView: 'early' };
  const lateOnly = { code: 'CTR2', name: '', isTempRow: true, shift: '', tempRowView: 'late' };
  assert(personMatchesShiftView(earlyOnly, 'early'), 'empty shift pinned early shows in early');
  assert(!personMatchesShiftView(earlyOnly, 'late'), 'empty shift pinned early hidden in late');
  assert(personMatchesShiftView(lateOnly, 'late'), 'empty shift pinned late shows in late');
  assert(!personMatchesShiftView(lateOnly, 'early'), 'empty shift pinned late hidden in early');
});

test('temp row keeps CTR code when display name differs', () => {
  const people = [
    { code: 'A1', name: '甲', personnelId: 'PA' },
    { code: 'CTR1', name: '支援小王', isTempRow: true }
  ];
  const tokenFormal = { kind: 'formal', code: 'A1', name: '甲' };
  const tokenTemp = { kind: 'temp', code: 'CTR1' };
  assert(resolvePersonDeleteIndex(people, tokenFormal, resolvePid, resolveName) === 0, 'formal resolves to A');
  assert(resolvePersonDeleteIndex(people, tokenTemp, resolvePid, resolveName) === 1, 'temp resolves by CTR code');
  const result = applySchedulePersonDelete(
    { people, schedule: [], peopleDisplayOrder: [0, 1] },
    tokenFormal,
    resolvePid,
    resolveName
  );
  assert(result.people.length === 1 && result.people[0].code === 'CTR1', 'CTR code preserved after formal delete');
});

test('temp row rejects duplicate display name', () => {
  const people = [
    { code: 'CTR1', name: 'CTR3', isTempRow: true },
    { code: 'CTR2', name: '小王', isTempRow: true }
  ];
  const r1 = syncTempRowIdentity(people, 1, 'CTR3');
  assert(!r1.ok && r1.reason === 'duplicate', 'should reject duplicate name');
  assert(people[1].name === '小王', 'name unchanged on reject');
  const r2 = syncTempRowIdentity(people, 1, '小李');
  assert(r2.ok, 'different name allowed');
  assert(people[1].name === '小李', 'name updated');
});

test('temp row rejects name matching another temp code', () => {
  const people = [
    { code: 'CTR3', name: 'CTR3', isTempRow: true },
    { code: 'CTR4', name: '', isTempRow: true }
  ];
  const r = syncTempRowIdentity(people, 1, 'CTR3');
  assert(!r.ok, 'cannot reuse another temp row code as name');
});

test('nextTempControlRowCode skips taken labels', () => {
  const people = [
    { code: 'CTR1', name: 'CTR1', isTempRow: true },
    { code: 'CTR3', name: 'CTR3', isTempRow: true }
  ];
  assert(nextTempControlRowCode(people) === 'CTR4', 'skips CTR3 when taken');
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
