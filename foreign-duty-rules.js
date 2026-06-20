/**
 * 外家航班 duty 規則（本家 = BR/B7，其餘為外家）
 * Phase 1：HX 依機型分人力與 ABG/PPT 時段
 * Phase 2：外家接飛 RC/BG 合併 duty
 * Phase 3：OZ 依機型分人力、接飛 ABG×2、A380 指引；OZ/BX 接飛 STA 前 20 分
 * Phase 4：BX 出境 ABG×1（STD 前 40 分）；BX794 ABG 含 check-in（STD 前 2h45m，管制表 CKIN+A）
 * Phase 5：NZ 出境 PPT×2（STD 前 70 分）；ABG CKIN×2（STD 前 3h35m，CKIN+A）＋登機門×1（STD 前 50 分）
 * Phase 6：TK 出境 RC×1（STD 前 90 分～STD 前 60 分）
 */
(function (global) {
  const HOMELINE_PREFIXES = ['BR', 'B7'];
  const CONNECTING_BEFORE_STA_MIN = 15;
  const CONNECTING_BEFORE_STA_OZ_BX_MIN = 20;

  /** HX 機型：A330 家族（匯入簡化為 330；相容舊碼 333、33G） */
  const HX_AC_TYPE_A330 = new Set(['330', '333', '33G']);

  /** OZ 機型：A380 */
  const OZ_AC_TYPE_A380 = new Set(['380']);

  const HX_DEP_DUTY = {
    rcBeforeMin: 70,
    bgBeforeMin: 60,
    abgBeforeMin: 50,
    abgSlots: 2,
    pptBeforeMin: 70,
    pptSlotsA330: 1
  };

  const OZ_DEP_DUTY = {
    abgSlots: 2,
    guideBeforeMin: 40
  };

  const BX_DEP_DUTY = {
    abgBeforeMin: 40,
    abgSlots: 1
  };

  /** BX794：ABG 需協助 check-in 後就定位，STD 前 2h45m～STD */
  const BX794_ABG_BEFORE_MIN = 165;

  const NZ_DEP_DUTY = {
    pptSlots: 2,
    pptBeforeMin: 70,
    abgCkinSlots: 2,
    abgCkinBeforeMin: 215,
    abgGateSlots: 1,
    abgGateBeforeMin: 50
  };

  const TK_DEP_DUTY = {
    rcBeforeMin: 90,
    rcEndBeforeStdMin: 60
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

  function isOzFlightDef(f) {
    return isForeignFlightDef(f) && flightAirline(f.flight) === 'OZ';
  }

  function isBxFlightDef(f) {
    return isForeignFlightDef(f) && flightAirline(f.flight) === 'BX';
  }

  function isNzFlightDef(f) {
    return isForeignFlightDef(f) && flightAirline(f.flight) === 'NZ';
  }

  function isTkFlightDef(f) {
    return isForeignFlightDef(f) && flightAirline(f.flight) === 'TK';
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

  /** @returns {'A380'|'OTHER'|null} */
  function getOzAcFamily(f) {
    if (!f || f.type !== 'DEP' || !isOzFlightDef(f)) return null;
    const ac = normalizeAcType(f.acType);
    if (OZ_AC_TYPE_A380.has(ac)) return 'A380';
    return 'OTHER';
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

  function getOzDepDutyConfig(f) {
    const family = getOzAcFamily(f);
    if (!family) return null;
    return {
      family,
      abgSlots: OZ_DEP_DUTY.abgSlots,
      guideBeforeMin: OZ_DEP_DUTY.guideBeforeMin,
      guideSlots: family === 'A380' ? 1 : 0
    };
  }

  /** A380 指引：STD 前 40 分～STD，分配表 PPT 欄顯示「指引:代號」 */
  function getOzGuideConfig(f) {
    const oz = getOzDepDutyConfig(f);
    if (!oz || !oz.guideSlots) return null;
    const flight = String(f.flight || '').trim();
    return {
      slots: oz.guideSlots,
      beforeMinutes: oz.guideBeforeMin,
      endAtStd: true,
      compactLabel: `${compactDutyFlightNo(flight)}指引`,
      assignmentPrefix: '指引:'
    };
  }

  function getOzAbgSlots(f) {
    return getOzDepDutyConfig(f)?.abgSlots ?? 0;
  }

  function isBx794Flight(f, normalizeFlightNo) {
    if (!f || f.type !== 'DEP' || f.status === 'CANX') return false;
    const norm = typeof normalizeFlightNo === 'function'
      ? normalizeFlightNo
      : (v) => String(v || '').trim().toUpperCase();
    return norm(f.flight) === norm('BX794');
  }

  /** BX 出境 ABG×1；BX794 延長至 STD 前 2h45m，管制表顯示「BX794 CKIN+A」 */
  function getBxAbgConfig(f, normalizeFlightNo) {
    if (!f || f.type !== 'DEP' || f.status === 'CANX' || !isBxFlightDef(f)) return null;
    const flight = String(f.flight || '').trim();
    if (isBx794Flight(f, normalizeFlightNo)) {
      return {
        abgSlots: BX_DEP_DUTY.abgSlots,
        beforeMinutes: BX794_ABG_BEFORE_MIN,
        endAtStd: true,
        compactLabel: `${compactDutyFlightNo(flight)} CKIN+A`
      };
    }
    return {
      abgSlots: BX_DEP_DUTY.abgSlots,
      beforeMinutes: BX_DEP_DUTY.abgBeforeMin,
      endAtStd: true
    };
  }

  function getBxAbgSlots(f) {
    if (!f || f.type !== 'DEP' || !isBxFlightDef(f)) return 0;
    return BX_DEP_DUTY.abgSlots;
  }

  function getNzPptConfig(f) {
    if (!f || f.type !== 'DEP' || f.status === 'CANX' || !isNzFlightDef(f)) return null;
    return {
      slots: NZ_DEP_DUTY.pptSlots,
      beforeMinutes: NZ_DEP_DUTY.pptBeforeMin,
      endAtStd: true
    };
  }

  /** NZ 出境 ABG：CKIN×2 與登機門×1 分開 palette chip，分配表皆顯示於 ABG 欄 */
  function getNzAbgConfigs(f) {
    if (!f || f.type !== 'DEP' || f.status === 'CANX' || !isNzFlightDef(f)) return null;
    const flight = String(f.flight || '').trim();
    const compact = compactDutyFlightNo(flight);
    return [
      {
        kind: 'CKIN',
        abgSlots: NZ_DEP_DUTY.abgCkinSlots,
        beforeMinutes: NZ_DEP_DUTY.abgCkinBeforeMin,
        endAtStd: true,
        compactLabel: `${compact} CKIN+A`
      },
      {
        kind: 'GATE',
        abgSlots: NZ_DEP_DUTY.abgGateSlots,
        beforeMinutes: NZ_DEP_DUTY.abgGateBeforeMin,
        endAtStd: true,
        compactLabel: `${compact}A`
      }
    ];
  }

  function getNzAbgSlots(f) {
    if (!f || f.type !== 'DEP' || !isNzFlightDef(f)) return 0;
    return NZ_DEP_DUTY.abgCkinSlots + NZ_DEP_DUTY.abgGateSlots;
  }

  /** TK 出境 RC×1：STD 前 90 分～STD 前 60 分 */
  function getTkRcConfig(f) {
    if (!f || f.type !== 'DEP' || f.status === 'CANX' || !isTkFlightDef(f)) return null;
    return {
      beforeMinutes: TK_DEP_DUTY.rcBeforeMin,
      endBeforeStdMin: TK_DEP_DUTY.rcEndBeforeStdMin
    };
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

  function getConnectingBeforeStaMin(depDef) {
    const airline = flightAirline(depDef?.flight);
    if (airline === 'OZ' || airline === 'BX') return CONNECTING_BEFORE_STA_OZ_BX_MIN;
    return CONNECTING_BEFORE_STA_MIN;
  }

  function roleShort(role) {
    if (role === 'RC') return 'R';
    if (role === 'BG') return 'B';
    if (role === 'ABG') return 'A';
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

  /** 接飛合併標籤；ABG 僅顯示出境班 + A，例如 OZ712A */
  function connectingCompactLabel(arrFlight, depFlight, role) {
    if (role === 'ABG') {
      return `${compactDutyFlightNo(depFlight)}A`;
    }
    const arrPart = compactDutyFlightNo(arrFlight);
    const depPart = compactDutyFlightNo(depFlight);
    const depSuffix = depPart.replace(/^[A-Z]{2}/i, '') || depPart;
    return `${arrPart}/${depSuffix}${roleShort(role)}`;
  }

  function connectingTemplateId(arrFlight, depFlight, role) {
    return `CONN_${arrFlight}_${depFlight}_${role}`;
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

  function getHomelineConnectingPair(arrDef, flightDefs, normalizeFlightNo) {
    if (!arrDef || arrDef.type !== 'ARR' || arrDef.status === 'CANX') return null;
    if (!isHomelineFlight(arrDef.flight)) return null;
    const depNorm = arrDef.connectingFlight ? normalizeFlightNo(arrDef.connectingFlight) : '';
    if (!depNorm) return null;
    const depDef = findFlightDef(flightDefs, depNorm, 'DEP', normalizeFlightNo);
    if (!depDef || depDef.status === 'CANX' || !isHomelineFlight(depDef.flight)) return null;
    return {
      arrFlight: arrDef.flight,
      depFlight: depDef.flight,
      arrDef,
      depDef
    };
  }

  function isHomelineConnectingArrFlight(arrDef, flightDefs, normalizeFlightNo) {
    return !!getHomelineConnectingPair(arrDef, flightDefs, normalizeFlightNo);
  }

  function isHomelineConnectingDepFlight(depDef, flightDefs, normalizeFlightNo) {
    if (!depDef || depDef.type !== 'DEP' || depDef.status === 'CANX') return false;
    if (!isHomelineFlight(depDef.flight)) return false;
    const depNorm = normalizeFlightNo(depDef.flight);
    return (flightDefs || []).some(arr => {
      if (arr?.type !== 'ARR' || arr.status === 'CANX') return false;
      const pair = getHomelineConnectingPair(arr, flightDefs, normalizeFlightNo);
      return !!(pair && normalizeFlightNo(pair.depFlight) === depNorm);
    });
  }

  function buildHomelineConnectingTemplates(flightDefs, normalizeFlightNo, getArrAnchorMin, getDepStdMin, shouldMergePair) {
    const out = [];
    const seen = new Set();
    const mergeCheck = typeof shouldMergePair === 'function' ? shouldMergePair : () => true;
    (flightDefs || []).forEach(arrDef => {
      const pair = getHomelineConnectingPair(arrDef, flightDefs, normalizeFlightNo);
      if (!pair) return;
      const key = `${normalizeFlightNo(pair.arrFlight)}|${normalizeFlightNo(pair.depFlight)}`;
      if (seen.has(key)) return;
      seen.add(key);
      if (!mergeCheck(pair.arrFlight, pair.depFlight)) return;
      const anchor = getArrAnchorMin(pair.arrDef);
      const std = getDepStdMin(pair.depDef);
      if (anchor == null || std == null) return;
      const start = Math.max(0, anchor - CONNECTING_BEFORE_STA_MIN);
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
          homelineConnecting: true,
          start,
          end
        });
      });
    });
    return out;
  }

  function getConnectingPair(arrDef, flightDefs, normalizeFlightNo) {
    if (!arrDef || arrDef.type !== 'ARR' || arrDef.status === 'CANX') return null;
    if (!isForeignFlightDef(arrDef)) return null;
    const depNorm = arrDef.connectingFlight ? normalizeFlightNo(arrDef.connectingFlight) : '';
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

  /** OZ 接飛：ABG 併入接飛 duty，出境班不另建獨立 ABG */
  function isOzConnectingDepFlight(depDef, flightDefs, normalizeFlightNo) {
    return isConnectingDepFlight(depDef, flightDefs, normalizeFlightNo) && isOzFlightDef(depDef);
  }

  function buildConnectingTemplates(flightDefs, normalizeFlightNo, getArrStaMin, getDepStdMin, shouldMergePair) {
    const out = [];
    const seen = new Set();
    const mergeCheck = typeof shouldMergePair === 'function' ? shouldMergePair : () => true;
    (flightDefs || []).forEach(arrDef => {
      const pair = getConnectingPair(arrDef, flightDefs, normalizeFlightNo);
      if (!pair) return;
      const key = `${normalizeFlightNo(pair.arrFlight)}|${normalizeFlightNo(pair.depFlight)}`;
      if (seen.has(key)) return;
      seen.add(key);
      if (!mergeCheck(pair.arrFlight, pair.depFlight)) return;
      const sta = getArrStaMin(pair.arrDef);
      const std = getDepStdMin(pair.depDef);
      if (sta == null || std == null) return;
      const beforeSta = getConnectingBeforeStaMin(pair.depDef);
      const start = Math.max(0, sta - beforeSta);
      const end = std;
      const roles = ['RC', 'BG'];
      if (isOzFlightDef(pair.depDef)) roles.push('ABG');
      roles.forEach(role => {
        const entry = {
          id: connectingTemplateId(pair.arrFlight, pair.depFlight, role),
          label: `${pair.arrFlight}/${pair.depFlight} ${role}`,
          compactLabel: connectingCompactLabel(pair.arrFlight, pair.depFlight, role),
          role,
          flight: pair.depFlight,
          arrFlight: pair.arrFlight,
          depFlight: pair.depFlight,
          flightType: 'CONN',
          connectingDuty: true,
          foreignConnecting: true,
          start,
          end
        };
        if (role === 'ABG') entry.abgSlots = getOzAbgSlots(pair.depDef) || OZ_DEP_DUTY.abgSlots;
        out.push(entry);
      });
    });
    return out;
  }

  global.ForeignDutyRules = {
    HX_AC_TYPE_A330,
    OZ_AC_TYPE_A380,
    CONNECTING_BEFORE_STA_MIN,
    CONNECTING_BEFORE_STA_OZ_BX_MIN,
    isHomelineFlight,
    isForeignFlightDef,
    isHxFlightDef,
    isOzFlightDef,
    isBxFlightDef,
    isNzFlightDef,
    isTkFlightDef,
    getHxAcFamily,
    getOzAcFamily,
    getHxDepDutyConfig,
    getHxPptConfig,
    getHxAbgSlots,
    getOzDepDutyConfig,
    getOzGuideConfig,
    getOzAbgSlots,
    getBxAbgConfig,
    getBxAbgSlots,
    getNzPptConfig,
    getNzAbgConfigs,
    getNzAbgSlots,
    getTkRcConfig,
    getConnectingBeforeStaMin,
    isHx253CkinFlight,
    getHx253CkinConfig,
    connectingCompactLabel,
    connectingTemplateId,
    getConnectingPair,
    getHomelineConnectingPair,
    isHomelineConnectingArrFlight,
    isHomelineConnectingDepFlight,
    buildHomelineConnectingTemplates,
    isConnectingArrFlight,
    isConnectingDepFlight,
    isOzConnectingDepFlight,
    buildConnectingTemplates
  };
})(typeof window !== 'undefined' ? window : globalThis);
