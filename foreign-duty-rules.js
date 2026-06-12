/**
 * 外家航班 duty 規則（本家 = BR/B7，其餘為外家）
 * Phase 1：HX 依機型分人力與 ABG/PPT 時段
 * Phase 2：外家接飛 RC/BG 合併 duty
 */
(function (global) {
  const HOMELINE_PREFIXES = ['BR', 'B7'];
  const CONNECTING_BEFORE_STA_MIN = 15;

  /** HX 機型：A330 家族（匯入簡化為 330；相容舊碼 333、33G） */
  const HX_AC_TYPE_A330 = new Set(['330', '333', '33G']);

  const HX_DEP_DUTY = {
    rcBeforeMin: 70,
    bgBeforeMin: 60,
    abgBeforeMin: 50,
    abgSlots: 2,
    pptBeforeMin: 70,
    pptSlotsA330: 1
  };

  function flightAirline(flight) {
    const m = String(flight || '').trim().toUpperCase().match(/^([A-Z]{2})/);
    return m ? m[1] : '';
  }

  function isHomelineFlight(flight) {
    const s = String(flight || '').trim().toUpperCase();
    return HOMELINE_PREFIXES.some(prefix => s.startsWith(prefix));
  }

  function isForeignFlightDef(f) {
    return !!(f && f.flight && !isHomelineFlight(f.flight));
  }

  function isHxFlightDef(f) {
    return isForeignFlightDef(f) && flightAirline(f.flight) === 'HX';
  }

  function normalizeAcType(acType) {
    return String(acType || '').trim().toUpperCase();
  }

  /** @returns {'A330'|'A321'|null} */
  function getHxAcFamily(f) {
    if (!f || f.type !== 'DEP' || !isHxFlightDef(f)) return null;
    const ac = normalizeAcType(f.acType);
    if (HX_AC_TYPE_A330.has(ac)) return 'A330';
    return 'A321';
  }

  function getHxDepDutyConfig(f) {
    const family = getHxAcFamily(f);
    if (!family) return null;
    return {
      family,
      rcBeforeMin: HX_DEP_DUTY.rcBeforeMin,
      bgBeforeMin: HX_DEP_DUTY.bgBeforeMin,
      abgBeforeMin: HX_DEP_DUTY.abgBeforeMin,
      abgSlots: HX_DEP_DUTY.abgSlots,
      pptBeforeMin: HX_DEP_DUTY.pptBeforeMin,
      pptSlots: family === 'A330' ? HX_DEP_DUTY.pptSlotsA330 : 0
    };
  }

  function getHxPptConfig(f) {
    const hx = getHxDepDutyConfig(f);
    if (!hx || !hx.pptSlots) return null;
    return {
      slots: hx.pptSlots,
      beforeMinutes: hx.pptBeforeMin,
      endAtStd: true
    };
  }

  function getHxAbgSlots(f) {
    return getHxDepDutyConfig(f)?.abgSlots ?? 0;
  }

  /** HX253 專用 check-in 協勤：STD 前 2h40m～STD 前 40m，1 人，分配表不顯示 */
  const HX253_CKIN_BEFORE_MIN = 160;
  const HX253_CKIN_END_BEFORE_STD_MIN = 40;

  function isHx253CkinFlight(f, normalizeFlightNo) {
    if (!f || f.type !== 'DEP' || f.status === 'CANX') return false;
    const norm = typeof normalizeFlightNo === 'function'
      ? normalizeFlightNo
      : (v) => String(v || '').trim().toUpperCase();
    return norm(f.flight) === norm('HX253');
  }

  function getHx253CkinConfig(f, normalizeFlightNo) {
    if (!isHx253CkinFlight(f, normalizeFlightNo)) return null;
    const flight = String(f.flight || '').trim();
    const label = `${flight} CKIN`;
    return {
      label,
      beforeMinutes: HX253_CKIN_BEFORE_MIN,
      endBeforeStdMin: HX253_CKIN_END_BEFORE_STD_MIN
    };
  }

  function roleShort(role) {
    if (role === 'RC') return 'R';
    if (role === 'BG') return 'B';
    return role;
  }

  function compactDutyFlightNo(flight) {
    const s = String(flight || '').trim().toUpperCase();
    if (!s) return '';
    const formatDutyNum = (numPart) => {
      if (!/^\d+$/.test(numPart)) return numPart;
      const trimmed = numPart.replace(/^0+/, '') || '0';
      return trimmed.length === 1 ? trimmed.padStart(2, '0') : trimmed;
    };
    if (s.startsWith('B7')) {
      const numPart = s.slice(2);
      if (/^\d+$/.test(numPart)) return `B7${formatDutyNum(numPart)}`;
    }
    if (s.startsWith('BR')) {
      const numPart = s.slice(2);
      if (/^\d+$/.test(numPart)) return formatDutyNum(numPart);
    }
    const m = s.match(/^([A-Z]{2})(\d+)$/);
    if (m) return `${m[1]}${formatDutyNum(m[2])}`;
    return s;
  }

  /** 接飛合併標籤，例如 HX252/253R */
  function connectingCompactLabel(arrFlight, depFlight, role) {
    const arrPart = compactDutyFlightNo(arrFlight);
    const depPart = compactDutyFlightNo(depFlight);
    const depSuffix = depPart.replace(/^[A-Z]{2}/i, '') || depPart;
    return `${arrPart}/${depSuffix}${roleShort(role)}`;
  }

  function connectingTemplateId(arrFlight, depFlight, role) {
    return `CONN_${arrFlight}_${depFlight}_${role}`;
  }

  function getFixedConnectingMap(normalizeFlightNo) {
    const norm = typeof normalizeFlightNo === 'function'
      ? normalizeFlightNo
      : (v) => String(v || '').trim().toUpperCase();
    const utils = global.FlightAssignmentUtils;
    return utils?.buildFixedConnectingMap?.(norm) || new Map();
  }

  function flightDefKey(f, normalizeFlightNo) {
    const norm = normalizeFlightNo(f?.flight);
    return norm && f?.type ? `${norm}|${f.type}` : '';
  }

  function findFlightDef(flightDefs, flight, type, normalizeFlightNo) {
    const target = normalizeFlightNo(flight);
    return (flightDefs || []).find(f =>
      f?.type === type && normalizeFlightNo(f.flight) === target
    ) || null;
  }

  function getConnectingPair(arrDef, flightDefs, normalizeFlightNo) {
    if (!arrDef || arrDef.type !== 'ARR' || arrDef.status === 'CANX') return null;
    if (!isForeignFlightDef(arrDef)) return null;
    const norm = normalizeFlightNo(arrDef.flight);
    const map = getFixedConnectingMap(normalizeFlightNo);
    let depNorm = arrDef.connectingFlight ? normalizeFlightNo(arrDef.connectingFlight) : '';
    if (!depNorm && map.has(norm)) depNorm = map.get(norm);
    if (!depNorm) return null;
    const depDef = findFlightDef(flightDefs, depNorm, 'DEP', normalizeFlightNo);
    if (!depDef || depDef.status === 'CANX' || !isForeignFlightDef(depDef)) return null;
    return {
      arrFlight: arrDef.flight,
      depFlight: depDef.flight,
      arrDef,
      depDef
    };
  }

  function isConnectingArrFlight(arrDef, flightDefs, normalizeFlightNo) {
    return !!getConnectingPair(arrDef, flightDefs, normalizeFlightNo);
  }

  function isConnectingDepFlight(depDef, flightDefs, normalizeFlightNo) {
    if (!depDef || depDef.type !== 'DEP' || depDef.status === 'CANX') return false;
    if (!isForeignFlightDef(depDef)) return false;
    const depNorm = normalizeFlightNo(depDef.flight);
    return (flightDefs || []).some(arr => {
      if (arr?.type !== 'ARR' || arr.status === 'CANX') return false;
      const pair = getConnectingPair(arr, flightDefs, normalizeFlightNo);
      return !!(pair && normalizeFlightNo(pair.depFlight) === depNorm);
    });
  }

  function buildConnectingTemplates(flightDefs, normalizeFlightNo, getArrStaMin, getDepStdMin) {
    const out = [];
    const seen = new Set();
    (flightDefs || []).forEach(arrDef => {
      const pair = getConnectingPair(arrDef, flightDefs, normalizeFlightNo);
      if (!pair) return;
      const key = `${normalizeFlightNo(pair.arrFlight)}|${normalizeFlightNo(pair.depFlight)}`;
      if (seen.has(key)) return;
      seen.add(key);
      const sta = getArrStaMin(pair.arrDef);
      const std = getDepStdMin(pair.depDef);
      if (sta == null || std == null) return;
      const start = Math.max(0, sta - CONNECTING_BEFORE_STA_MIN);
      const end = std;
      ['RC', 'BG'].forEach(role => {
        out.push({
          id: connectingTemplateId(pair.arrFlight, pair.depFlight, role),
          label: `${pair.arrFlight}/${pair.depFlight} ${role}`,
          compactLabel: connectingCompactLabel(pair.arrFlight, pair.depFlight, role),
          role,
          flight: pair.depFlight,
          arrFlight: pair.arrFlight,
          depFlight: pair.depFlight,
          flightType: 'CONN',
          connectingDuty: true,
          start,
          end
        });
      });
    });
    return out;
  }

  global.ForeignDutyRules = {
    HX_AC_TYPE_A330,
    CONNECTING_BEFORE_STA_MIN,
    isHomelineFlight,
    isForeignFlightDef,
    isHxFlightDef,
    getHxAcFamily,
    getHxDepDutyConfig,
    getHxPptConfig,
    getHxAbgSlots,
    isHx253CkinFlight,
    getHx253CkinConfig,
    connectingCompactLabel,
    connectingTemplateId,
    getConnectingPair,
    isConnectingArrFlight,
    isConnectingDepFlight,
    buildConnectingTemplates
  };
})(typeof window !== 'undefined' ? window : globalThis);
