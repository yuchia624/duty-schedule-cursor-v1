/**
 * Browser E2E: delete formal person must not remove CTR temp row.
 *
 * Run with the app open (roll-call mode on), then paste in DevTools console,
 * or evaluate via CDP Runtime.evaluate on index.html.
 */
async function e2eDeleteFormalPreservesCtr() {
  const out = { steps: [], pass: false };
  const log = (step, data) => {
    const entry = { step, ...data };
    out.steps.push(entry);
    console.log('[e2eDeleteFormalPreservesCtr]', step, data);
  };

  if (!rollCallMode) document.getElementById('rollCallToggle')?.click();

  const formalCountBefore = people.filter(p => !p.isTempRow).length;
  if (formalCountBefore < 1) {
    people.unshift(normalizePerson({ code: 'E2E-F1', name: 'E2E正式', personnelId: 'E2E-P1', shift: 'hM', level: 'M' }));
    schedule.forEach(s => { if (s.personIndex >= 0) s.personIndex += 1; });
    peopleDisplayOrder = people.map((_, i) => i);
    renderRows();
  }

  let ctrIdx = people.findIndex(p => p.isTempRow && /^CTR\d+$/i.test(p.code || ''));
  if (ctrIdx < 0) {
    addTempControlRow();
    ctrIdx = people.findIndex(p => p.isTempRow);
  }
  const ctrCode = people[ctrIdx].code;

  const tpl = createCustomTemplate('E2E-CTR-DUTY', '07:00', 60);
  if (tpl) {
    customTemplates.push(tpl);
    rebuildTemplates();
    renderPalette();
  }
  const tplId = templates.find(t => t.label === 'E2E-CTR-DUTY')?.id;
  const timeline = document.querySelector(`.timeline[data-original-index="${ctrIdx}"]`);
  if (tplId && timeline) {
    handleTimelinePaletteDrop({
      preventDefault() {},
      stopPropagation() {},
      dataTransfer: { types: ['text/template'], getData: t => (t === 'text/template' ? tplId : '') },
      clientX: timeline.getBoundingClientRect().left + 100
    }, timeline);
  }

  const dutyUidBefore = schedule.find(s => s.personIndex === ctrIdx)?.uid;
  log('beforeDelete', {
    peopleLength: people.length,
    tempRows: people.filter(p => p.isTempRow).map(p => p.code),
    peopleDisplayOrder: peopleDisplayOrder?.slice(),
    displaySorted: getDisplaySortedRows().map(r => r.person.code),
    displayRows: getDisplayRows().map(r => r.person.code),
    ctrDutyUid: dutyUidBefore
  });

  const formalIdx = people.findIndex(p => !p.isTempRow);
  const token = getPersonDeleteToken(people[formalIdx]);
  log('deleteTarget', { formalIdx, token, code: people[formalIdx].code });

  const origConfirm = window.appConfirm;
  window.appConfirm = async () => true;
  schedulePersonDeleteInFlight = false;
  schedulePersonDeleteMutedUntil = 0;
  try {
    await deleteSchedulePersonRow(token);
    await new Promise(r => setTimeout(r, 500));
  } finally {
    window.appConfirm = origConfirm;
  }

  const ctrAfterIdx = people.findIndex(p => p.isTempRow && p.code === ctrCode);
  const ctrDutyAfter = schedule.filter(s => people[s.personIndex]?.isTempRow);
  const ctrDom = document.querySelector(`.person-cell[data-original-index="${ctrAfterIdx}"]`);

  log('afterDelete', {
    peopleLength: people.length,
    tempRows: people.filter(p => p.isTempRow).map(p => p.code),
    peopleDisplayOrder: peopleDisplayOrder?.slice(),
    displaySorted: getDisplaySortedRows().map(r => r.person.code),
    displayRows: getDisplayRows().map(r => r.person.code),
    ctrInPeople: ctrAfterIdx >= 0,
    ctrInDisplay: getDisplayRows().some(r => r.person.isTempRow && r.person.code === ctrCode),
    ctrDomVisible: !!ctrDom,
    ctrDutyCount: ctrDutyAfter.length,
    ctrDutyUids: ctrDutyAfter.map(d => d.uid)
  });

  out.pass =
    ctrAfterIdx >= 0 &&
    getDisplayRows().some(r => r.person.isTempRow && r.person.code === ctrCode) &&
    !!ctrDom &&
    (!dutyUidBefore || ctrDutyAfter.some(d => d.uid === dutyUidBefore));

  console.log(out.pass ? 'E2E PASS' : 'E2E FAIL', out);
  return out;
}

if (typeof window !== 'undefined') {
  window.e2eDeleteFormalPreservesCtr = e2eDeleteFormalPreservesCtr;
}
