/**
 * Flight Import Reconciliation — 純函式核對引擎
 *
 * 職責：比較「匯入前的班表狀態(baseline)」與「匯入後的班表狀態(current)」，
 * 產生一份結構化的變更計畫(ReconciliationPlan)，供 UI 顯示與後續套用。
 *
 * 設計原則：
 *   - 純函式、零副作用：不讀寫全域、不碰 DOM、不改 flightDefs / schedule。
 *   - 所有相依（正規化航班號、2 小時判斷、已排 duty 清單）都由 opts 注入。
 *   - 本版聚焦：新增航班 / 取消航班 / STD·STA 時間異動 / 接飛異動 / 已排 duty 影響。
 *   - 本版「不」處理 manual 接飛的套用邏輯（僅標記，不納入建議動作）。
 *
 * Snapshot 結構（由呼叫端在 index.html 產生）：
 *   {
 *     flights:    [{ flight, type:'DEP'|'ARR', baseMinutes, baseTime, status }],
 *     connecting: [{ arr, dep, source }]   // 每班 ARR 目前接的 DEP（dep 可為空字串）
 *   }
 */
(function (global) {
  const DEFAULT_NORM = (v) => String(v ?? '').trim().toUpperCase();

  function flightKey(flight, type, norm) {
    return `${norm(flight)}|${String(type || '').toUpperCase()}`;
  }

  function indexFlights(flights, norm) {
    const map = new Map();
    (Array.isArray(flights) ? flights : []).forEach((f) => {
      if (!f || !f.flight || !f.type) return;
      map.set(flightKey(f.flight, f.type, norm), f);
    });
    return map;
  }

  function connectingMap(connecting, norm) {
    const map = new Map();
    (Array.isArray(connecting) ? connecting : []).forEach((c) => {
      if (!c || !c.arr) return;
      map.set(norm(c.arr), {
        dep: c.dep ? norm(c.dep) : '',
        source: c.source || ''
      });
    });
    return map;
  }

  // 找出參照到某航班的已排 duty（非接飛）
  function dutiesForFlight(scheduleItems, flight, type, norm) {
    const fn = norm(flight);
    const out = [];
    scheduleItems.forEach((it) => {
      if (!it) return;
      if (it.connectingDuty) {
        if (norm(it.arrFlight) === fn || norm(it.depFlight || it.flight) === fn) {
          out.push(describeDuty(it));
        }
        return;
      }
      if (it.flightType === type && norm(it.flight) === fn) out.push(describeDuty(it));
    });
    return out;
  }

  // 找出某接飛配對(arr[/dep])上的已排合併接飛 duty
  function connectingDutiesFor(scheduleItems, arrNorm, depNorm, norm) {
    const out = [];
    scheduleItems.forEach((it) => {
      if (!it || !it.connectingDuty) return;
      if (norm(it.arrFlight) !== arrNorm) return;
      if (depNorm && norm(it.depFlight || it.flight) !== depNorm) return;
      out.push(describeDuty(it));
    });
    return out;
  }

  function describeDuty(it) {
    return {
      uid: it.uid,
      role: it.role,
      personIndex: it.personIndex,
      connecting: !!it.connectingDuty,
      arrFlight: it.arrFlight || '',
      depFlight: it.depFlight || (it.connectingDuty ? it.flight : '') || '',
      flight: it.flight || '',
      flightType: it.flightType || ''
    };
  }

  /**
   * 產生變更計畫。
   * @param {object} baseline  匯入前 snapshot
   * @param {object} current   匯入後 snapshot
   * @param {object} opts      { normalizeFlightNo, isWithinTwoHours(arr,dep), scheduleItems, source }
   * @returns {object} ReconciliationPlan
   */
  function buildPlan(baseline, current, opts = {}) {
    const norm = typeof opts.normalizeFlightNo === 'function' ? opts.normalizeFlightNo : DEFAULT_NORM;
    const within2h = typeof opts.isWithinTwoHours === 'function' ? opts.isWithinTwoHours : () => true;
    const scheduleItems = Array.isArray(opts.scheduleItems) ? opts.scheduleItems : [];

    const base = baseline || {};
    const cur = current || {};
    const baseFlights = indexFlights(base.flights, norm);
    const curFlights = indexFlights(cur.flights, norm);

    const added = [];
    const timeChanged = [];

    curFlights.forEach((cf, key) => {
      const bf = baseFlights.get(key);
      if (!bf) {
        if (cf.status !== 'CANX') {
          added.push({ flight: cf.flight, type: cf.type, time: cf.baseTime || '' });
        }
        return;
      }
      // STD (DEP) / STA (ARR) 只比對 FIS 基準時間 baseMinutes
      if (Number(bf.baseMinutes) !== Number(cf.baseMinutes)) {
        timeChanged.push({
          flight: cf.flight,
          type: cf.type,
          field: cf.type === 'DEP' ? 'STD' : 'STA',
          from: bf.baseTime || '',
          to: cf.baseTime || '',
          fromMinutes: Number(bf.baseMinutes),
          toMinutes: Number(cf.baseMinutes),
          affectedDuties: dutiesForFlight(scheduleItems, cf.flight, cf.type, norm)
        });
      }
    });

    const cancelled = [];
    baseFlights.forEach((bf, key) => {
      const cf = curFlights.get(key);
      if (!cf) {
        cancelled.push({
          flight: bf.flight,
          type: bf.type,
          reason: 'missing',
          affectedDuties: dutiesForFlight(scheduleItems, bf.flight, bf.type, norm)
        });
      } else if (cf.status === 'CANX' && bf.status !== 'CANX') {
        cancelled.push({
          flight: bf.flight,
          type: bf.type,
          reason: 'CANX',
          affectedDuties: dutiesForFlight(scheduleItems, bf.flight, bf.type, norm)
        });
      }
    });

    const connecting = buildConnectingChanges({
      baseConn: connectingMap(base.connecting, norm),
      curConn: connectingMap(cur.connecting, norm),
      baseFlights,
      curFlights,
      norm,
      within2h,
      scheduleItems
    });

    const hasChanges = !!(added.length || cancelled.length || timeChanged.length || connecting.length);
    return {
      meta: {
        source: opts.source || '',
        hasChanges,
        counts: {
          added: added.length,
          cancelled: cancelled.length,
          timeChanged: timeChanged.length,
          connecting: connecting.length
        }
      },
      added,
      cancelled,
      timeChanged,
      connecting
    };
  }

  function flightLookup(flightsMaps, flight, type, norm) {
    if (!flight) return null;
    const key = flightKey(flight, type, norm);
    for (const map of flightsMaps) {
      const f = map && map.get(key);
      if (f) return f;
    }
    return null;
  }

  function depBaseTime(flightsMaps, dep, norm) {
    const f = flightLookup(flightsMaps, dep, 'DEP', norm);
    return f && f.baseTime ? String(f.baseTime) : '';
  }

  /** 接飛時間 = 出境 STD − 入境 STA（分鐘）；跨午夜時 +1440。 */
  function connectingGapMinutes(arr, dep, arrMaps, depMaps, norm) {
    if (!arr || !dep) return null;
    const arrF = flightLookup(arrMaps, arr, 'ARR', norm);
    const depF = flightLookup(depMaps, dep, 'DEP', norm);
    if (!arrF || !depF) return null;
    const arrMin = Number(arrF.baseMinutes);
    const depMin = Number(depF.baseMinutes);
    if (!Number.isFinite(arrMin) || !Number.isFinite(depMin)) return null;
    let diff = depMin - arrMin;
    if (diff <= 0) diff += 1440;
    return diff;
  }

  function buildConnectingChanges({ baseConn, curConn, baseFlights, curFlights, norm, within2h, scheduleItems }) {
    // 新狀態下每個 DEP 被哪些 ARR 接（找重複）
    const depToArrs = new Map();
    curConn.forEach((info, arr) => {
      if (!info.dep) return;
      if (!depToArrs.has(info.dep)) depToArrs.set(info.dep, []);
      depToArrs.get(info.dep).push(arr);
    });

    const rows = [];
    const seenArr = new Set();
    const arrKeys = new Set([...baseConn.keys(), ...curConn.keys()]);
    const flightMaps = { baseFlights, curFlights };

    arrKeys.forEach((arr) => {
      const oldInfo = baseConn.get(arr) || { dep: '', source: '' };
      const newInfo = curConn.get(arr) || { dep: '', source: '' };
      const oldDep = oldInfo.dep;
      const newDep = newInfo.dep;
      const isManual = newInfo.source === 'manual' || oldInfo.source === 'manual';
      const dup = newDep ? (depToArrs.get(newDep) || []).length > 1 : false;

      // 只納入「有變更」或「新狀態產生重複」
      if (oldDep === newDep && !dup) return;

      rows.push(makeConnectingRow({ arr, oldDep, newDep, dup, isManual, within2h, scheduleItems, norm, flightMaps }));
      seenArr.add(arr);
    });

    // 補：接飛未變、但因別班而形成重複的 ARR
    depToArrs.forEach((arrs, dep) => {
      if (arrs.length <= 1) return;
      arrs.forEach((arr) => {
        if (seenArr.has(arr)) return;
        const oldInfo = baseConn.get(arr) || { dep: '', source: '' };
        const newInfo = curConn.get(arr) || { dep: '', source: '' };
        const isManual = newInfo.source === 'manual' || oldInfo.source === 'manual';
        rows.push(makeConnectingRow({
          arr, oldDep: oldInfo.dep, newDep: dep, dup: true, isManual, within2h, scheduleItems, norm, flightMaps
        }));
        seenArr.add(arr);
      });
    });

    return rows;
  }

  function makeConnectingRow({ arr, oldDep, newDep, dup, isManual, within2h, scheduleItems, norm, flightMaps }) {
    let suggested;
    if (isManual) {
      suggested = 'manual-skip';
    } else if (!newDep) {
      suggested = 'split'; // 接飛已移除 → 拆開已排接飛 duty
    } else if (dup) {
      suggested = 'pick'; // 重複 → 需選擇保留哪班入境
    } else if (!within2h(arr, newDep)) {
      suggested = 'split'; // 超過 2 小時 → 拆開
    } else {
      suggested = 'update'; // 換 DEP 且 <=2h → 更新接飛
    }
    const baseFlights = flightMaps?.baseFlights;
    const curFlights = flightMaps?.curFlights;
    // 原配對：優先用匯入前時間；新配對：優先用 FIS 新狀態（無→有時可直接看到 STD−STA）
    const oldConnMinutes = connectingGapMinutes(
      arr, oldDep, [baseFlights, curFlights], [baseFlights, curFlights], norm
    );
    const newConnMinutes = connectingGapMinutes(
      arr, newDep, [curFlights, baseFlights], [curFlights, baseFlights], norm
    );
    const oldDepTime = depBaseTime([baseFlights, curFlights], oldDep, norm);
    const newDepTime = depBaseTime([curFlights, baseFlights], newDep, norm);
    return {
      arrFlight: arr,
      oldDep,
      newDep,
      oldDepTime,
      newDepTime,
      oldConnMinutes,
      newConnMinutes,
      duplicate: dup,
      manual: isManual,
      suggested,
      affectedDuties: connectingDutiesFor(scheduleItems, arr, oldDep, norm)
    };
  }

  global.FlightReconciliation = {
    buildPlan,
    // 匯出內部工具，方便未來測試/擴充
    _internal: { flightKey, indexFlights, connectingMap }
  };
})(typeof window !== 'undefined' ? window : globalThis);
