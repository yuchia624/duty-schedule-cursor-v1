#!/usr/bin/env node
/**
 * Smoke tests for duty-qualification.js (mirrors index.html globals at runtime).
 */
import { createRequire } from 'module';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

const global = globalThis;
global.SCHEDULE_BLOCK_A_LEVEL_ORDER = ['M', 'm', '2A', '2A*', '1A', '學生', '見習'];
global.getScheduleBlockALevelRank = levelRaw => {
  const raw = String(levelRaw || '').trim();
  return global.SCHEDULE_BLOCK_A_LEVEL_ORDER.indexOf(raw);
};
global.normalizeStaffLevelFilterKey = raw => String(raw || '').trim();
global.resolveLevel = person => String(person?.level || '').trim();
global.getDutyDisplayLabel = duty => String(duty?.label || duty?.role || '').trim();
global.isInternCopyDuty = duty => !!(duty?.paletteCopyIntern);
global.flightDefs = [{ flight: 'CX123', type: 'DEP' }];
global.ForeignDutyRules = {
  isForeignFlightDef: def => !!(def && def.flight === 'CX123'),
  isOzFlightDef: () => false
};

const code = readFileSync(join(root, 'duty-qualification.js'), 'utf8');
// eslint-disable-next-line no-eval
eval(code);

function assert(label, cond) {
  if (!cond) throw new Error(`FAIL: ${label}`);
  console.log(`ok: ${label}`);
}

const student = { level: '學生' };
const support = { level: '支援' };
const dic = { level: 'DIC' };
const m = { level: 'M' };
const twoAStar = { level: '2A*' };

assert('student blocked RC', !global.personQualifiesForDuty(student, { role: 'RC', flight: 'BR123' }));
assert('student allowed ABG', global.personQualifiesForDuty(student, { role: 'ABG', flight: 'BR123' }));
assert('support only ABG homeline', global.personQualifiesForDuty(support, { role: 'ABG', flight: 'BR123' }));
assert('support blocked RC', !global.personQualifiesForDuty(support, { role: 'RC', flight: 'BR123' }));
assert('DIC can DIC duty', global.personQualifiesForDuty(dic, { role: 'DIC', dicKind: 'DIC' }));
assert('M qualifies TC', global.personQualifiesForDuty(m, { role: 'TC' }));
assert('2A* blocked foreign BG', !global.personMeetsForeignBgQualification(twoAStar, {
  role: 'BG',
  flight: 'CX123',
  flightType: 'DEP'
}));

const soft = global.evaluatePersonDutyQualification({ level: '見習' }, { role: 'PRE' });
assert('intern PRE soft fail', !soft.ok && soft.kind === 'soft');

console.log('All duty-qualification smoke tests passed.');
