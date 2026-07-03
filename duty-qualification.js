/**
 * 戰力 / 勤務資格規則（自 index.html 抽出）
 * 依賴 index.html 內的全域函式與狀態（執行時解析，非載入時）。
 */
(function (global) {
  'use strict';

  function resolveLevel(person) { return global.resolveLevel(person); }
  function normalizeStaffLevelFilterKey(raw) { return global.normalizeStaffLevelFilterKey(raw); }
  function getDutyDisplayLabel(duty) { return global.getDutyDisplayLabel(duty); }
  function isInternCopyDuty(duty) { return global.isInternCopyDuty(duty); }
  function getScheduleBlockALevelRank(levelRaw) { return global.getScheduleBlockALevelRank(levelRaw); }
  function getScheduleBlockALevelOrder() { return global.SCHEDULE_BLOCK_A_LEVEL_ORDER; }
  function getScheduleOneALevelRank() { return getScheduleBlockALevelOrder().indexOf('1A'); }
  function getFlightDefs() { return global.flightDefs || []; }

  function getInternCopyDutyTierRuleRole(duty) {
    const role = duty?.role || '';
    if (role === 'RC' || role === 'BG' || role === 'TC' || role === 'DIC') return role;
    return '';
  }
  function personMeetsInternCopyDutyQualification(person, duty) {
    const role = getInternCopyDutyTierRuleRole(duty);
    if (!role) return true;
    const tier = normalizeLevelTier(person);
    if (role === 'RC') {
      if (isOzRcDuty(duty)) return tier === 'm';
      return tier === 'M' || tier === 'm' || tier === '2A';
    }
    if (role === 'BG' || role === 'TC') {
      if (role === 'BG') return tier === 'M' || tier === 'm' || tier === '2A' || tier === '1A';
      return tier === 'M' || tier === 'm' || tier === '2A';
    }
    if (role === 'DIC') {
      const levelKey = getPersonLevelFilterKey(person);
      return tier === 'M' || tier === 'm' || levelKey === 'DIC';
    }
    return true;
  }
  function getInternCopyDutyQualificationFailure(person, duty) {
    if (!isInternCopyDuty(duty)) return '';
    const role = getInternCopyDutyTierRuleRole(duty);
    if (!role) return '';
    const tier = normalizeLevelTier(person);
    if (tier === '學生') {
      const levelLabel = String(person?.level || resolveLevel(person) || tier).trim() || tier;
      const dutyLabel = getDutyDisplayLabel(duty) || role;
      return `${levelLabel} 不可排 ${dutyLabel}`;
    }
    if (tier === '見習') {
      const levelLabel = String(person?.level || resolveLevel(person) || tier).trim() || tier;
      const dutyLabel = getDutyDisplayLabel(duty) || role;
      return `${levelLabel} 不可排 ${dutyLabel}`;
    }
    if (role === 'RC' && isOzRcDuty(duty)) return '';
    if (personMeetsInternCopyDutyQualification(person, duty)) return '';
    const levelLabel = String(person?.level || resolveLevel(person) || tier).trim() || tier;
    const dutyLabel = getDutyDisplayLabel(duty) || role;
    if (role === 'RC') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'BG') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'TC') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'DIC') return `${levelLabel} 不可排 ${dutyLabel}`;
    return '';
  }

  function normalizeLevelTier(person) {
    const raw = String(person?.level || resolveLevel(person) || '').trim();
    if (!raw) return '';
    const key = normalizeStaffLevelFilterKey(raw);
    if (key === '學生' || key === '見習') return key;
    if (key === 'M') return 'M';
    if (key === 'm') return 'm';
    if (key === '2A') return '2A';
    if (key === '1A') return '1A';
    if (raw === 'M') return 'M';
    const stripped = raw.replace(/[*!\-]/g, '');
    if (stripped === 'm') return 'm';
    const core = stripped.toUpperCase();
    if (core === '2A') return '2A';
    if (core === '1A') return '1A';
    if (core === 'M') return 'M';
    return core;
  }
  function isOzFlight(flight) {
    return String(flight || '').trim().toUpperCase().startsWith('OZ');
  }
  function dutyInvolvesOzFlight(duty) {
    if (!duty) return false;
    const flightDefs = getFlightDefs();
    const check = flight => {
      const f = String(flight || '').trim();
      if (!f) return false;
      if (isOzFlight(f)) return true;
      const depDef = flightDefs.find(x => x.flight === f && x.type === 'DEP');
      const arrDef = flightDefs.find(x => x.flight === f && x.type === 'ARR');
      return !!(global.ForeignDutyRules?.isOzFlightDef?.(depDef) || global.ForeignDutyRules?.isOzFlightDef?.(arrDef));
    };
    if (duty.connectingDuty) {
      return check(duty.arrFlight) || check(duty.depFlight || duty.flight);
    }
    return check(duty.flight);
  }
  function isOzRcDuty(duty) {
    if (!duty) return false;
    const role = duty.role || '';
    if (role !== 'RC' && role !== '接機RC') return false;
    return dutyInvolvesOzFlight(duty);
  }
  function isInternOzRcDuty(duty) {
    return isInternCopyDuty(duty) && isOzRcDuty(duty);
  }
  function getPersonLevelFilterKey(person) {
    return normalizeStaffLevelFilterKey(String(person?.level || resolveLevel(person) || '').trim());
  }
  function isPersonDicLevel(person) {
    return getPersonLevelFilterKey(person) === 'DIC' || normalizeLevelTier(person) === 'DIC';
  }
  function isPersonSupervisorLevel(person) {
    return getPersonLevelFilterKey(person) === '業督' || normalizeLevelTier(person) === '業督';
  }
  function isPersonDicOrSupervisorLevel(person) {
    return isPersonDicLevel(person) || isPersonSupervisorLevel(person);
  }
  function isPersonSupportLevel(person) {
    return getPersonLevelFilterKey(person) === '支援' || normalizeLevelTier(person) === '支援';
  }
  function isPersonClerkLevel(person) {
    return getPersonLevelFilterKey(person) === '事務員' || normalizeLevelTier(person) === '事務員';
  }
  function isPersonLevelAtOrAboveM(person) {
    const tier = normalizeLevelTier(person);
    return tier === 'M' || tier === 'm';
  }
  function getDicQualificationHardFailure(person, duty) {
    if (!duty || duty.role !== 'DIC') return '';
    const dicKind = String(duty.dicKind || duty.label || duty.compactLabel || '').trim();
    const levelLabel = String(person?.level || resolveLevel(person) || '').trim() || '此人員';
    const dutyLabel = dicKind || getDutyDisplayLabel(duty) || 'DIC';
    const levelKey = getPersonLevelFilterKey(person);
    if (dicKind === '業督') {
      if (levelKey !== '業督') return `${levelLabel} 不可排 ${dutyLabel}`;
      return '';
    }
    if (dicKind === '小天使') {
      if (!isPersonLevelAtOrAboveM(person) && !isPersonDicOrSupervisorLevel(person)) {
        return `${levelLabel} 不可排 ${dutyLabel}`;
      }
      return '';
    }
    if (levelKey !== 'DIC') return `${levelLabel} 不可排 ${dutyLabel}`;
    return '';
  }
  function dutyRequiresPersonQualification(duty) {
    if (!duty) return false;
    const role = duty.role;
    return role && role !== '休' && role !== 'DIC' && role !== '其他' && role !== '自訂';
  }
  function getFlightDefForDuty(duty) {
    if (!duty?.flight) return null;
    const flightType = duty.flightType || 'DEP';
    return getFlightDefs().find(f => f.flight === duty.flight && f.type === flightType) || null;
  }
  function isOzGuidePptDuty(duty) {
    if (duty?.role !== 'PPT') return false;
    return !!global.ForeignDutyRules?.getOzGuideConfig?.(getFlightDefForDuty(duty));
  }
  function dutyInvolvesForeignFlight(duty) {
    if (!duty) return false;
    const flightDefs = getFlightDefs();
    const isForeign = f => !!(f && global.ForeignDutyRules?.isForeignFlightDef?.(f));
    if (duty.connectingDuty) {
      const arrFlight = String(duty.arrFlight || '').trim();
      const depFlight = String(duty.depFlight || duty.flight || '').trim();
      const arrDef = arrFlight ? flightDefs.find(f => f.flight === arrFlight && f.type === 'ARR') : null;
      const depDef = depFlight ? flightDefs.find(f => f.flight === depFlight && f.type === 'DEP') : null;
      return isForeign(arrDef) || isForeign(depDef);
    }
    return isForeign(getFlightDefForDuty(duty));
  }
  function isRcFamilyDuty(duty) {
    const role = duty?.role || '';
    return role === 'RC' || role === '接機RC';
  }
  const INBOUND_RC_BLOCKED_LEVEL_KEYS = new Set(['學生', '見習', '支援', '事務員']);
  function isInboundRcDuty(duty) {
    return (duty?.role || '') === '接機RC';
  }
  function personMeetsInboundRcQualification(person, duty) {
    if (!person || !isInboundRcDuty(duty)) return true;
    const levelKey = getPersonLevelFilterKey(person);
    if (INBOUND_RC_BLOCKED_LEVEL_KEYS.has(levelKey)) return false;
    const tier = normalizeLevelTier(person);
    return tier !== '學生' && tier !== '見習';
  }
  function getInboundRcQualificationHardFailure(person, duty) {
    if (!isInboundRcDuty(duty) || personMeetsInboundRcQualification(person, duty)) return '';
    const levelLabel = String(person?.level || resolveLevel(person) || '').trim() || '此人員';
    const dutyLabel = getDutyDisplayLabel(duty) || '接機RC';
    return `${levelLabel} 不可排 ${dutyLabel}`;
  }
  function isBgFamilyDuty(duty) {
    const role = duty?.role || '';
    return role === 'BG' || role === '接機C';
  }
  function isPersonLevel2AStar(person) {
    const raw = String(person?.level || resolveLevel(person) || '').trim();
    if (!raw || normalizeLevelTier(person) !== '2A') return false;
    return raw.includes('*');
  }
  function isForeignBgDuty(duty) {
    if (!duty || isInternCopyDuty(duty) || !isBgFamilyDuty(duty)) return false;
    return dutyInvolvesForeignFlight(duty);
  }
  function personMeetsForeignBgQualification(person, duty) {
    if (!isForeignBgDuty(duty)) return true;
    if (isPersonLevel2AStar(person)) return false;
    const tier = normalizeLevelTier(person);
    if (tier === '見習' || tier === '學生' || tier === '1A') return false;
    return tier === 'M' || tier === 'm' || tier === '2A';
  }
  function isTcFamilyDuty(duty) {
    return (duty?.role || '') === 'TC' && !isInternCopyDuty(duty);
  }
  function personMeetsTcQualification(person, duty) {
    if (!isTcFamilyDuty(duty)) return true;
    const tier = normalizeLevelTier(person);
    return tier === 'M' || tier === 'm' || tier === '2A';
  }
  function getForeignBgQualificationHardFailure(person, duty) {
    if (!isForeignBgDuty(duty)) return '';
    if (personMeetsForeignBgQualification(person, duty)) return '';
    const levelLabel = String(person?.level || resolveLevel(person) || '').trim() || '此人員';
    const dutyLabel = getDutyDisplayLabel(duty) || duty.role || 'BG';
    return `${levelLabel} 不可排 ${dutyLabel}`;
  }
  function personTierCanPerformDutyTierRules(tier, duty) {
    if (!dutyRequiresPersonQualification(duty)) return true;
    if (isRcFamilyDuty(duty) && !isInternCopyDuty(duty)) {
      if (isInboundRcDuty(duty)) {
        return tier !== '學生' && tier !== '見習' && tier !== '支援' && tier !== '事務員';
      }
      if (isOzRcDuty(duty)) return tier === 'M';
      return tier === 'M' || tier === 'm' || tier === 'DIC' || tier === '業督';
    }
    const role = duty.role;
    const flightType = duty.flightType;
    const flight = duty.flight;
    if (role === 'PRE') {
      return tier === 'M' || tier === 'm' || tier === '2A' || tier === '1A' || tier === 'DIC' || tier === '業督';
    }
    if (role === 'TC' && !isInternCopyDuty(duty)) {
      return tier === 'M' || tier === 'm' || tier === '2A';
    }
    if (tier === 'DIC' || tier === '業督') return true;
    if (tier === 'M') return true;
    if (tier === 'm') {
      if (isOzFlight(flight) && (role === 'RC' || role === '接機RC')) return false;
      return true;
    }
    if (tier === '2A' || tier === '1A') {
      if (tier === '1A' && flightType === 'DEP' && role === 'BG') return false;
      return true;
    }
    return false;
  }
  function getDutyMinimumQualificationRank(duty) {
    const blockOrder = getScheduleBlockALevelOrder();
    if (!dutyRequiresPersonQualification(duty)) return getScheduleOneALevelRank();
    if ((duty?.role || '') === 'TC') return blockOrder.indexOf('2A');
    if (isForeignBgDuty(duty)) return blockOrder.indexOf('2A');
    const order = ['M', 'm', '2A', '1A'];
    for (let i = order.length - 1; i >= 0; i--) {
      if (personTierCanPerformDutyTierRules(order[i], duty)) {
        return getScheduleBlockALevelRank(order[i]);
      }
    }
    return getScheduleOneALevelRank();
  }
  function isDutyAtOrBelow1ALevel(duty) {
    return getDutyMinimumQualificationRank(duty) >= getScheduleOneALevelRank();
  }
  const STUDENT_ALLOWED_FLIGHT_ROLES = new Set(['ABG', '大件', 'BA', 'L1', 'G2', '接機C', 'PPT']);
  const STUDENT_ALLOWED_OTHER_KINDS = new Set([
    '抄Load', '入境通報', '設備', '關單', '備品', '查房員', '遺失物', 'C5R指引', '值日生', 'Q2'
  ]);
  const STUDENT_BLOCKED_OTHER_KINDS = new Set(['UA P/F', '客服']);
  function isStudentAllowedDuty(duty) {
    if (!duty) return false;
    const role = duty.role || '';
    if (role === '其他') {
      const kind = String(duty.otherKind || duty.label || '').trim();
      return STUDENT_ALLOWED_OTHER_KINDS.has(kind);
    }
    if (!STUDENT_ALLOWED_FLIGHT_ROLES.has(role)) return false;
    if (role === '接機C') return (duty.flightType || '') === 'ARR';
    return true;
  }
  function isStudentBlockedOtherDuty(duty) {
    if (!duty || duty.role !== '其他') return false;
    const kind = String(duty.otherKind || duty.label || '').trim();
    return STUDENT_BLOCKED_OTHER_KINDS.has(kind);
  }
  function getStudentQualificationHardFailure(person, duty) {
    const tier = normalizeLevelTier(person);
    if (tier !== '學生') return '';
    const levelLabel = String(person.level || resolveLevel(person) || tier).trim() || tier;
    const role = duty?.role || '';
    const dutyLabel = getDutyDisplayLabel(duty) || role || '此勤務';
    if (dutyInvolvesForeignFlight(duty)) {
      return `${levelLabel} 不可排 ${dutyLabel}`;
    }
    if (isInboundRcDuty(duty)) return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'RC') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'BG') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'CKIN') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'PRE') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'TC') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === 'T' && duty.tMode !== 'tc') return `${levelLabel} 不可排 ${dutyLabel}`;
    if (role === '其他' && isStudentBlockedOtherDuty(duty)) {
      return `${levelLabel} 不可排 ${dutyLabel}`;
    }
    if (dutyRequiresPersonQualification(duty) && !isStudentAllowedDuty(duty)) {
      return `${levelLabel} 不可排 ${dutyLabel}`;
    }
    return '';
  }
  function getSupportClerkQualificationHardFailure(person, duty) {
    if (!person || !duty) return '';
    if (!isPersonSupportLevel(person) && !isPersonClerkLevel(person)) return '';
    if (personMeetsBaseTierQualification(person, duty)) return '';
    const levelLabel = String(person?.level || resolveLevel(person) || '').trim() || '此人員';
    const dutyLabel = getDutyDisplayLabel(duty) || duty.role || '此勤務';
    return `${levelLabel} 不可排 ${dutyLabel}`;
  }
  function personMeetsBaseTierQualification(person, duty) {
    if (!person || !duty) return false;
    if (isPersonClerkLevel(person)) {
      if (duty.role === '休') return true;
      if (!dutyRequiresPersonQualification(duty)) {
        if (duty.role === '其他') {
          const kind = String(duty.otherKind || duty.label || '').trim();
          return kind === '抄Load';
        }
        return false;
      }
      return false;
    }
    if (isPersonSupportLevel(person)) {
      if (!dutyRequiresPersonQualification(duty)) return duty.role === '休';
      return (duty.role || '') === 'ABG' && !dutyInvolvesForeignFlight(duty);
    }
    if (!dutyRequiresPersonQualification(duty)) return true;
    const tier = normalizeLevelTier(person);
    const role = duty.role;
    if (role === 'PRE') {
      if (tier === '學生') return false;
      if (tier === '見習') return true;
      if (isPersonDicOrSupervisorLevel(person)) return true;
      return tier === 'M' || tier === 'm' || tier === '2A' || tier === '1A';
    }
    if (isRcFamilyDuty(duty) && !isInternCopyDuty(duty)) {
      if (isInboundRcDuty(duty)) return personMeetsInboundRcQualification(person, duty);
      if (isOzRcDuty(duty)) return tier === 'M';
      if (tier === 'm' && dutyInvolvesOzFlight(duty)) return false;
      return tier === 'M' || tier === 'm' || isPersonDicOrSupervisorLevel(person);
    }
    if (isForeignBgDuty(duty)) {
      return personMeetsForeignBgQualification(person, duty);
    }
    if (role === 'TC' && !isInternCopyDuty(duty)) {
      return personMeetsTcQualification(person, duty);
    }
    if (isPersonDicOrSupervisorLevel(person)) return true;
    const flightType = duty.flightType;
    const flight = duty.flight;
    if (tier === 'M') return true;
    if (tier === 'm') {
      if (isOzFlight(flight) && role === '接機RC') return false;
      return true;
    }
    if (tier === '2A' || tier === '1A') {
      if (tier === '1A' && flightType === 'DEP' && role === 'BG') return false;
      return true;
    }
    if (tier === '見習') return true;
    if (tier === '學生') {
      if (!dutyRequiresPersonQualification(duty)) {
        return !isStudentBlockedOtherDuty(duty);
      }
      return isStudentAllowedDuty(duty);
    }
    return personTierCanPerformDutyTierRules(tier, duty);
  }
  function getBaseTierQualificationFailureMessage(person, duty) {
    const tier = normalizeLevelTier(person);
    const levelLabel = String(person.level || resolveLevel(person) || tier || '').trim() || tier || '此人員';
    const dutyLabel = getDutyDisplayLabel(duty) || duty.role || '此勤務';
    const inboundRcFail = getInboundRcQualificationHardFailure(person, duty);
    if (inboundRcFail) return inboundRcFail;
    const foreignBgFail = getForeignBgQualificationHardFailure(person, duty);
    if (foreignBgFail) return foreignBgFail;
    return `${levelLabel} 不可排 ${dutyLabel}`;
  }
  function getInternQualificationHardFailure(person, duty) {
    const tier = normalizeLevelTier(person);
    if (tier !== '見習') return '';
    if (!dutyRequiresPersonQualification(duty)) return '';
    if (isDutyAtOrBelow1ALevel(duty)) return '';
    const levelLabel = String(person.level || resolveLevel(person) || tier).trim() || tier;
    const dutyLabel = getDutyDisplayLabel(duty) || duty.role || '此勤務';
    return `${levelLabel} 不可排 ${dutyLabel}`;
  }
  function getInternQualificationSoftFailure(person, duty) {
    const tier = normalizeLevelTier(person);
    if (tier !== '見習') return '';
    const role = duty?.role;
    if (!role || role === '休') return '';
    const levelLabel = String(person.level || resolveLevel(person) || tier).trim() || tier;
    const dutyLabel = getDutyDisplayLabel(duty) || role || '此勤務';
    const message = `${levelLabel} 不可排 ${dutyLabel}？`;
    if (role === 'PRE') return message;
    if (!dutyRequiresPersonQualification(duty)) return message;
    if (!isDutyAtOrBelow1ALevel(duty)) return '';
    return message;
  }
  function evaluatePersonDutyQualification(person, duty) {
    if (!person || !duty) return { ok: false, kind: 'hard', message: '' };
    if (isInternOzRcDuty(duty)) {
      const tier = normalizeLevelTier(person);
      const levelLabel = String(person?.level || resolveLevel(person) || tier).trim() || tier;
      const dutyLabel = getDutyDisplayLabel(duty) || '見習 OZ RC';
      if (tier === 'm') return { ok: true, kind: 'ok', message: '' };
      return { ok: false, kind: 'hard', message: `${levelLabel} 不可排 ${dutyLabel}` };
    }
    const internCopyFail = getInternCopyDutyQualificationFailure(person, duty);
    if (internCopyFail) return { ok: false, kind: 'hard', message: internCopyFail };
    const supportClerkFail = getSupportClerkQualificationHardFailure(person, duty);
    if (supportClerkFail) return { ok: false, kind: 'hard', message: supportClerkFail };
    if (isInternCopyDuty(duty) && getInternCopyDutyTierRuleRole(duty)) {
      const studentFail = getStudentQualificationHardFailure(person, duty);
      if (studentFail) return { ok: false, kind: 'hard', message: studentFail };
      return { ok: true, kind: 'ok', message: '' };
    }
    const dicFail = getDicQualificationHardFailure(person, duty);
    if (dicFail) return { ok: false, kind: 'hard', message: dicFail };
    const foreignBgFail = getForeignBgQualificationHardFailure(person, duty);
    if (foreignBgFail) return { ok: false, kind: 'hard', message: foreignBgFail };
    if (!dutyRequiresPersonQualification(duty)) {
      const studentFail = getStudentQualificationHardFailure(person, duty);
      if (studentFail) return { ok: false, kind: 'hard', message: studentFail };
      const internSoftFail = getInternQualificationSoftFailure(person, duty);
      if (internSoftFail) return { ok: false, kind: 'soft', message: internSoftFail };
      return { ok: true, kind: 'ok', message: '' };
    }
    if (!personMeetsBaseTierQualification(person, duty)) {
      return { ok: false, kind: 'hard', message: getBaseTierQualificationFailureMessage(person, duty) };
    }
    const studentFail = getStudentQualificationHardFailure(person, duty);
    if (studentFail) return { ok: false, kind: 'hard', message: studentFail };
    const internHardFail = getInternQualificationHardFailure(person, duty);
    if (internHardFail) return { ok: false, kind: 'hard', message: internHardFail };
    const internSoftFail = getInternQualificationSoftFailure(person, duty);
    if (internSoftFail) return { ok: false, kind: 'soft', message: internSoftFail };
    return { ok: true, kind: 'ok', message: '' };
  }
  function personQualifiesForDuty(person, duty) {
    return evaluatePersonDutyQualification(person, duty).ok;
  }
  function getPersonDutyQualificationFailureMessage(person, duty) {
    const issue = evaluatePersonDutyQualification(person, duty);
    return issue.ok ? '' : issue.message;
  }

  Object.assign(global, {
    getInternCopyDutyTierRuleRole,
    personMeetsInternCopyDutyQualification,
    getInternCopyDutyQualificationFailure,
    normalizeLevelTier,
    isOzFlight,
    dutyInvolvesOzFlight,
    isOzRcDuty,
    isInternOzRcDuty,
    getPersonLevelFilterKey,
    isPersonDicLevel,
    isPersonSupervisorLevel,
    isPersonDicOrSupervisorLevel,
    isPersonSupportLevel,
    isPersonClerkLevel,
    isPersonLevelAtOrAboveM,
    getDicQualificationHardFailure,
    dutyRequiresPersonQualification,
    getFlightDefForDuty,
    isOzGuidePptDuty,
    dutyInvolvesForeignFlight,
    isRcFamilyDuty,
    isInboundRcDuty,
    personMeetsInboundRcQualification,
    getInboundRcQualificationHardFailure,
    isBgFamilyDuty,
    isPersonLevel2AStar,
    isForeignBgDuty,
    personMeetsForeignBgQualification,
    isTcFamilyDuty,
    personMeetsTcQualification,
    getForeignBgQualificationHardFailure,
    personTierCanPerformDutyTierRules,
    getDutyMinimumQualificationRank,
    isDutyAtOrBelow1ALevel,
    isStudentAllowedDuty,
    isStudentBlockedOtherDuty,
    getStudentQualificationHardFailure,
    getSupportClerkQualificationHardFailure,
    personMeetsBaseTierQualification,
    getBaseTierQualificationFailureMessage,
    getInternQualificationHardFailure,
    getInternQualificationSoftFailure,
    evaluatePersonDutyQualification,
    personQualifiesForDuty,
    getPersonDutyQualificationFailureMessage
  });
})(typeof window !== 'undefined' ? window : globalThis);
