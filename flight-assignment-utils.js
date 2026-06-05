/**
 * 分配表航班欄位：變更紀錄、登機門改色、接飛自動配對
 */
(function (global) {
  const FIXED_CONNECTING_RAW = [
    ['HX252', 'HX253'],
    ['HX254', 'HX255'],
    ['HX260', 'HX261'],
    ['HX282', 'HX283'],
    ['OZ711', 'OZ712'],
    ['OZ713', 'OZ714'],
    ['BX791', 'BX792'],
    ['BX793', 'BX794'],
    ['NZ77', 'NZ78']
  ];

  let flightDefIdSeq = 0;

  function cleanAcNo(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim().toUpperCase();
  }

  function scheduleDayTimeMs(selectedDate, timeStr) {
    if (typeof ForeignScheduleParser !== 'undefined' && typeof ForeignScheduleParser.datetimeOnScheduleDay === 'function') {
      const dt = ForeignScheduleParser.datetimeOnScheduleDay(timeStr, selectedDate);
      return dt ? dt.ms : null;
    }
    const s = String(timeStr || '').trim();
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    if (!m) return null;
    const iso = String(selectedDate || '').trim();
    const dm = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!dm) return null;
    return new Date(Number(dm[1]), Number(dm[2]) - 1, Number(dm[3]), Number(m[1]), Number(m[2]), 0).getTime();
  }

  function turnaroundDiffMs(selectedDate, staTime, stdTime) {
    const staMs = scheduleDayTimeMs(selectedDate, staTime);
    const stdMs = scheduleDayTimeMs(selectedDate, stdTime);
    if (staMs == null || stdMs == null) return null;
    let diff = stdMs - staMs;
    if (diff <= 0) diff += 24 * 60 * 60 * 1000;
    return diff;
  }

  function buildFixedConnectingMap(normalizeFlightNo) {
    const norm = typeof normalizeFlightNo === 'function'
      ? normalizeFlightNo
      : (v) => String(v || '').trim().toUpperCase();
    const map = new Map();
    FIXED_CONNECTING_RAW.forEach(([arr, dep]) => {
      map.set(norm(arr), norm(dep));
    });
    return map;
  }

  function ensureFlightDefIds(flightDefs) {
    if (!Array.isArray(flightDefs)) return;
    flightDefs.forEach(f => {
      if (!f || typeof f !== 'object') return;
      if (!f.id) {
        flightDefIdSeq += 1;
        f.id = `fd_${Date.now().toString(36)}_${flightDefIdSeq.toString(36)}`;
      }
      if (!Array.isArray(f.edits)) f.edits = [];
      if (f.type === 'ARR') {
        if (f.connectingFlight == null) f.connectingFlight = '';
        if (f.connectingSource == null) f.connectingSource = '';
      }
    });
  }

  function getFieldEdits(flightDef, field) {
    return (flightDef?.edits || []).filter(e => e.field === field);
  }

  function getFieldEditCount(flightDef, field) {
    return getFieldEdits(flightDef, field).length;
  }

  /** 改回匯入原值（如 C7→C8→C7）視為第 3 次，顯示綠色 */
  function getFieldEditDisplayCount(flightDef, field) {
    const edits = getFieldEdits(flightDef, field);
    if (!edits.length) return 0;
    let count = edits.length;
    if (edits.length >= 2 && String(edits[edits.length - 1].to) === String(edits[0].from)) {
      count = Math.max(count, 3);
    }
    return count;
  }

  function getFieldEditColorClass(count) {
    if (count <= 0) return '';
    if (count === 1) return 'field-edits-1';
    if (count === 2) return 'field-edits-2';
    if (count === 3) return 'field-edits-3';
    return 'field-edits-4plus';
  }

  function getLastFieldEdit(flightDef, field) {
    const edits = getFieldEdits(flightDef, field);
    return edits.length ? edits[edits.length - 1] : null;
  }

  function recordFieldEdit(flightDef, field, from, to) {
    if (!flightDef) return;
    if (!Array.isArray(flightDef.edits)) flightDef.edits = [];
    const fromVal = String(from ?? '');
    const toVal = String(to ?? '').trim();
    if (fromVal === toVal) return;
    flightDef.edits.push({
      field,
      from: fromVal,
      to: toVal,
      at: new Date().toISOString()
    });
  }

  /** 修正最後一筆改門：更新 to、保留 at；若 to === from 則刪除該筆 */
  function correctLastFieldEdit(flightDef, field, newTo) {
    if (!flightDef) return false;
    const last = getLastFieldEdit(flightDef, field);
    if (!last) return false;
    const newVal = String(newTo ?? '').trim();
    if (String(last.to ?? '').trim() === newVal) return false;
    last.to = newVal;
    if (field === 'depGate') flightDef.depGate = newVal;
    else if (field === 'arrGate') flightDef.arrGate = newVal;
    if (String(last.from) === String(last.to)) {
      const idx = (flightDef.edits || []).indexOf(last);
      if (idx >= 0) flightDef.edits.splice(idx, 1);
    }
    return true;
  }

  /** 依改門鏈節點 index 修正歷史，並同步 depGate/arrGate 為最後節點 */
  function correctFieldEditAtChainIndex(flightDef, field, chainIndex, newValue) {
    if (!flightDef) return false;
    const edits = getFieldEdits(flightDef, field);
    if (!edits.length) return false;
    const chain = buildFieldHistoryChain(flightDef, field);
    if (chainIndex < 0 || chainIndex >= chain.length) return false;

    const newVal = String(newValue ?? '').trim();
    if (String(chain[chainIndex] ?? '').trim() === newVal) return false;

    chain[chainIndex] = newVal;

    const otherEdits = (flightDef.edits || []).filter(e => e.field !== field);
    const newFieldEdits = [];
    for (let i = 0; i < chain.length - 1; i++) {
      const from = String(chain[i] ?? '');
      const to = String(chain[i + 1] ?? '').trim();
      if (from === to) continue;
      const at = edits[i]?.at || new Date().toISOString();
      newFieldEdits.push({ field, from, to, at });
    }

    flightDef.edits = [...otherEdits, ...newFieldEdits];

    const currentGate = String(chain[chain.length - 1] ?? '').trim();
    if (field === 'depGate') flightDef.depGate = currentGate;
    else if (field === 'arrGate') flightDef.arrGate = currentGate;

    return true;
  }

  function buildFieldHistoryChain(flightDef, field) {
    const edits = getFieldEdits(flightDef, field);
    if (!edits.length) return [];
    return [edits[0].from, ...edits.map(e => e.to)];
  }

  function buildFieldHistoryHtml(flightDef, field, escapeHtml) {
    const esc = typeof escapeHtml === 'function' ? escapeHtml : (s) => String(s);
    const chain = buildFieldHistoryChain(flightDef, field);
    if (chain.length <= 1) return '';
    return chain.map((v, i) => {
      const text = esc(v || '—');
      return i < chain.length - 1 ? `<s>${text}</s>` : text;
    }).join(' → ');
  }

  function gatesReadyForPair(arr, dep) {
    return !!(String(arr?.arrGate || '').trim() && String(dep?.depGate || '').trim());
  }

  function computeConnectingFlights(flightDefs, selectedDate, opts = {}) {
    if (!Array.isArray(flightDefs) || !flightDefs.length) return;
    ensureFlightDefIds(flightDefs);
    const normalizeFlightNo = typeof opts.normalizeFlightNo === 'function'
      ? opts.normalizeFlightNo
      : (v) => String(v || '').trim().toUpperCase();
    const preserveManual = opts.preserveManual !== false;
    const fixedMap = buildFixedConnectingMap(normalizeFlightNo);
    const depFlights = flightDefs.filter(f => f.type === 'DEP' && f.status !== 'CANX');
    const arrFlights = flightDefs.filter(f => f.type === 'ARR');

    arrFlights.forEach(arr => {
      if (preserveManual && arr.connectingSource === 'manual') return;

      const arrNorm = normalizeFlightNo(arr.flight);
      if (fixedMap.has(arrNorm)) {
        arr.connectingFlight = fixedMap.get(arrNorm);
        arr.connectingSource = 'foreign-fixed';
        return;
      }

      const acNo = cleanAcNo(arr.acNo);
      if (!acNo) {
        if (arr.connectingSource !== 'manual') {
          arr.connectingFlight = '';
          arr.connectingSource = '';
        }
        return;
      }

      let bestDep = null;
      let bestDiff = Infinity;

      depFlights.forEach(dep => {
        if (cleanAcNo(dep.acNo) !== acNo) return;
        const diffMs = turnaroundDiffMs(selectedDate, arr.baseTime, dep.baseTime);
        if (diffMs == null || diffMs <= 0) return;

        const bothGates = gatesReadyForPair(arr, dep);
        let maxMs;
        if (bothGates) {
          if (String(arr.arrGate).trim() !== String(dep.depGate).trim()) return;
          maxMs = 6 * 60 * 60 * 1000;
        } else {
          maxMs = 3 * 60 * 60 * 1000;
        }
        if (diffMs > maxMs) return;
        if (diffMs < bestDiff) {
          bestDiff = diffMs;
          bestDep = dep;
        }
      });

      if (bestDep) {
        arr.connectingFlight = bestDep.flight;
        arr.connectingSource = 'auto';
      } else if (arr.connectingSource !== 'manual') {
        arr.connectingFlight = '';
        arr.connectingSource = '';
      }
    });
  }

  global.FlightAssignmentUtils = {
    FIXED_CONNECTING_RAW,
    ensureFlightDefIds,
    getFieldEdits,
    getFieldEditCount,
    getFieldEditDisplayCount,
    getFieldEditColorClass,
    getLastFieldEdit,
    recordFieldEdit,
    correctLastFieldEdit,
    correctFieldEditAtChainIndex,
    buildFieldHistoryChain,
    buildFieldHistoryHtml,
    computeConnectingFlights,
    buildFixedConnectingMap,
    turnaroundDiffMs
  };
})(typeof window !== 'undefined' ? window : globalThis);
