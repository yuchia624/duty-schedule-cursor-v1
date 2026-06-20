/**
 * 外家時刻表主檔：carrier + season + 日期區間
 * rows 分 key 存放，僅在 FIS 匯入或航班管理頁按需載入（不影響開頁速度）
 */
(function (global) {
  const REGISTRY_VERSION = '2026-06-18-v1';
  const META_KEY = 'cursor_v1_foreign_schedule_meta_v1';
  const ROWS_KEY_PREFIX = 'cursor_v1_foreign_schedule_rows_v1_';
  const KNOWN_CARRIERS = ['HX', 'OZ', 'NZ', 'BX'];

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function isValidIsoDate(s) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '').trim());
  }

  function compareIsoDate(a, b) {
    return String(a || '').localeCompare(String(b || ''));
  }

  function rangesOverlap(aStart, aEnd, bStart, bEnd) {
    return compareIsoDate(aStart, bEnd) <= 0 && compareIsoDate(bStart, aEnd) <= 0;
  }

  function makeScheduleId(carrier, startDate) {
    const c = clean(carrier).toUpperCase() || 'XX';
    const d = clean(startDate).replace(/-/g, '') || 'nodate';
    return `${c.toLowerCase()}-${d}-${Date.now().toString(36)}`;
  }

  function loadMeta() {
    try {
      const raw = localStorage.getItem(META_KEY);
      if (!raw) return { version: REGISTRY_VERSION, schedules: [] };
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') return { version: REGISTRY_VERSION, schedules: [] };
      if (!Array.isArray(parsed.schedules)) parsed.schedules = [];
      return parsed;
    } catch (_) {
      return { version: REGISTRY_VERSION, schedules: [] };
    }
  }

  function saveMeta(meta) {
    const next = {
      version: REGISTRY_VERSION,
      updatedAt: new Date().toISOString(),
      schedules: Array.isArray(meta?.schedules) ? meta.schedules : [],
      connectingFltMigratedFromFixed: !!meta?.connectingFltMigratedFromFixed
    };
    localStorage.setItem(META_KEY, JSON.stringify(next));
    return next;
  }

  function rowsStorageKey(scheduleId) {
    return `${ROWS_KEY_PREFIX}${scheduleId}`;
  }

  function loadRows(scheduleId) {
    if (!scheduleId) return [];
    try {
      const raw = localStorage.getItem(rowsStorageKey(scheduleId));
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      return [];
    }
  }

  function saveRows(scheduleId, rows) {
    if (!scheduleId) return;
    localStorage.setItem(rowsStorageKey(scheduleId), JSON.stringify(rows || []));
  }

  function deleteRows(scheduleId) {
    if (!scheduleId) return;
    localStorage.removeItem(rowsStorageKey(scheduleId));
  }

  function hasAnySchedules() {
    return loadMeta().schedules.length > 0;
  }

  function listSchedules() {
    return loadMeta().schedules.slice().sort((a, b) => {
      const ca = compareIsoDate(a.carrier, b.carrier);
      if (ca) return ca;
      return compareIsoDate(a.startDate, b.startDate);
    });
  }

  function getSchedule(scheduleId) {
    const meta = loadMeta();
    const entry = meta.schedules.find(s => s.scheduleId === scheduleId);
    if (!entry) return null;
    return { ...entry, rows: loadRows(scheduleId) };
  }

  function validateScheduleEntry(entry, excludeScheduleId) {
    const carrier = clean(entry?.carrier).toUpperCase();
    const startDate = clean(entry?.startDate);
    const endDate = clean(entry?.endDate);
    const seasonLabel = clean(entry?.seasonLabel);
    if (!carrier) return '請選擇航司';
    if (!seasonLabel) return '請填寫季節名稱';
    if (!isValidIsoDate(startDate) || !isValidIsoDate(endDate)) return '請填寫有效日期區間';
    if (compareIsoDate(startDate, endDate) > 0) return '結束日不可早於起始日';
    const meta = loadMeta();
    for (const s of meta.schedules) {
      if (excludeScheduleId && s.scheduleId === excludeScheduleId) continue;
      if (s.carrier !== carrier) continue;
      if (rangesOverlap(startDate, endDate, s.startDate, s.endDate)) {
        return `與「${s.seasonLabel}」（${s.startDate}～${s.endDate}）日期重疊`;
      }
    }
    return null;
  }

  function resolveSchedulesForDate(dateISO, meta) {
    if (!isValidIsoDate(dateISO)) return [];
    const schedules = (meta || loadMeta()).schedules || [];
    return schedules.filter(s =>
      isValidIsoDate(s.startDate)
      && isValidIsoDate(s.endDate)
      && compareIsoDate(s.startDate, dateISO) <= 0
      && compareIsoDate(dateISO, s.endDate) <= 0
    );
  }

  function upsertSchedule(entry, rows) {
    const err = validateScheduleEntry(entry, entry.scheduleId || null);
    if (err) return { ok: false, error: err };
    const meta = loadMeta();
    const now = new Date().toISOString();
    const scheduleId = entry.scheduleId || makeScheduleId(entry.carrier, entry.startDate);
    const patch = {
      scheduleId,
      carrier: clean(entry.carrier).toUpperCase(),
      seasonLabel: clean(entry.seasonLabel),
      seasonKind: clean(entry.seasonKind) || 'custom',
      startDate: clean(entry.startDate),
      endDate: clean(entry.endDate),
      source: clean(entry.source) || '',
      rowCount: Array.isArray(rows) ? rows.length : 0,
      updatedAt: now
    };
    const idx = meta.schedules.findIndex(s => s.scheduleId === scheduleId);
    if (idx >= 0) meta.schedules[idx] = { ...meta.schedules[idx], ...patch };
    else meta.schedules.push(patch);
    saveRows(scheduleId, rows || []);
    saveMeta(meta);
    return { ok: true, schedule: patch };
  }

  function deleteSchedule(scheduleId) {
    const meta = loadMeta();
    const next = meta.schedules.filter(s => s.scheduleId !== scheduleId);
    if (next.length === meta.schedules.length) return { ok: false, error: '找不到時刻表' };
    deleteRows(scheduleId);
    saveMeta({ ...meta, schedules: next });
    return { ok: true };
  }

  function filterRowsForCarrier(rows, carrier) {
    const code = clean(carrier).toUpperCase();
    if (!code || typeof global.ForeignScheduleParser?.inferCarrierFromFlt !== 'function') return rows || [];
    return (rows || []).filter(row => global.ForeignScheduleParser.inferCarrierFromFlt(row.flt) === code);
  }

  function buildCandidatesForDate(dateISO, opts = {}) {
    migrateLegacyConnectingFltFromFixed();
    const meta = loadMeta();
    const active = resolveSchedulesForDate(dateISO, meta);
    if (!active.length) {
      return {
        candidates: [],
        skipped: [],
        stats: { dep: 0, arr: 0, total: 0 },
        applied: [],
        warnings: ['航班管理尚未設定此外家時刻表區間'],
        registryVersion: REGISTRY_VERSION
      };
    }
    const Parser = global.ForeignScheduleParser;
    if (!Parser || typeof Parser.buildCandidatesFromNormalizedRows !== 'function') {
      return { error: '外家 parser 未載入', candidates: [], skipped: [], stats: { dep: 0, arr: 0, total: 0 } };
    }
    let mergedRows = [];
    const applied = [];
    active.forEach(sched => {
      const rows = loadRows(sched.scheduleId);
      mergedRows = mergedRows.concat(rows);
      applied.push({
        scheduleId: sched.scheduleId,
        carrier: sched.carrier,
        seasonLabel: sched.seasonLabel,
        startDate: sched.startDate,
        endDate: sched.endDate,
        rowCount: rows.length
      });
    });
    const result = Parser.buildCandidatesFromNormalizedRows(mergedRows, dateISO, opts);
    result.applied = applied;
    result.registryVersion = REGISTRY_VERSION;
    result.warnings = [];
    const carriersInRegistry = new Set(meta.schedules.map(s => s.carrier));
    KNOWN_CARRIERS.forEach(code => {
      if (!carriersInRegistry.has(code)) return;
      const hit = active.some(s => s.carrier === code);
      if (!hit) result.warnings.push(`${code} 在 ${dateISO} 無適用時刻表區間`);
    });
    return result;
  }

  function previewForDate(dateISO, opts = {}) {
    return buildCandidatesForDate(dateISO, opts);
  }

  function collectRowsForAcTypeSync(dateISO) {
    const active = resolveSchedulesForDate(dateISO, loadMeta());
    let rows = [];
    active.forEach(s => { rows = rows.concat(loadRows(s.scheduleId)); });
    return rows;
  }

  function removeRowAt(scheduleId, rowIndex) {
    const rows = loadRows(scheduleId);
    if (rowIndex < 0 || rowIndex >= rows.length) return { ok: false, error: '找不到航班列' };
    rows.splice(rowIndex, 1);
    if (!rows.length) return deleteSchedule(scheduleId);
    saveRows(scheduleId, rows);
    const meta = loadMeta();
    const entry = meta.schedules.find(s => s.scheduleId === scheduleId);
    if (entry) {
      entry.rowCount = rows.length;
      entry.updatedAt = new Date().toISOString();
      saveMeta(meta);
    }
    return { ok: true };
  }

  function findScheduleForNewRow(carrier, fromDate, tillDate) {
    const code = clean(carrier).toUpperCase();
    return (loadMeta().schedules || []).find(s =>
      s.carrier === code
      && isValidIsoDate(s.startDate)
      && isValidIsoDate(s.endDate)
      && compareIsoDate(s.startDate, fromDate) <= 0
      && compareIsoDate(tillDate, s.endDate) <= 0
    ) || null;
  }

  function isInboundScheduleRow(dep, arr) {
    const d = clean(dep).toUpperCase();
    const a = clean(arr).toUpperCase();
    if (d === 'TPE') return false;
    return a === 'TPE';
  }

  function normalizeConnectingFlt(dep, arr, value) {
    if (!isInboundScheduleRow(dep, arr)) return '';
    return clean(value).toUpperCase();
  }

  /** 一次性：將舊固定配對表寫入 registry connectingFlt（移除固定表前的資料遷移） */
  const LEGACY_FIXED_CONNECTING = [
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

  function normalizeFltNo(v) {
    const s = clean(v).toUpperCase().replace(/\s/g, '');
    const m = s.match(/^([A-Z0-9]{2})(\d+)$/);
    if (!m) return s;
    return m[1] + String(Number(m[2]));
  }

  function legacyFixedConnectingDep(arrFlt) {
    const arrNorm = normalizeFltNo(arrFlt);
    for (const [arr, dep] of LEGACY_FIXED_CONNECTING) {
      if (normalizeFltNo(arr) === arrNorm) return normalizeFltNo(dep);
    }
    return '';
  }

  function migrateLegacyConnectingFltFromFixed() {
    const meta = loadMeta();
    if (meta.connectingFltMigratedFromFixed) return { ok: true, updated: 0 };
    let updated = 0;
    (meta.schedules || []).forEach(sched => {
      const rows = loadRows(sched.scheduleId);
      let changed = false;
      rows.forEach(row => {
        if (!isInboundScheduleRow(row.dep, row.arr)) return;
        if (clean(row.connectingFlt)) return;
        const depFlt = legacyFixedConnectingDep(row.flt);
        if (!depFlt) return;
        row.connectingFlt = depFlt;
        changed = true;
        updated += 1;
      });
      if (changed) saveRows(sched.scheduleId, rows);
    });
    saveMeta({ ...meta, connectingFltMigratedFromFixed: true });
    return { ok: true, updated };
  }

  function addRowAt(patch) {
    const Parser = global.ForeignScheduleParser;
    const infer = Parser && typeof Parser.inferCarrierFromFlt === 'function'
      ? Parser.inferCarrierFromFlt.bind(Parser)
      : () => '';
    const parseDate = Parser && typeof Parser.parseExcelDateValue === 'function'
      ? Parser.parseExcelDateValue.bind(Parser)
      : (v) => clean(v);
    const flt = clean(patch?.flt).toUpperCase();
    const carrier = clean(patch?.carrier || infer(flt)).toUpperCase();
    const fromDate = parseDate(patch?.fromDate);
    const tillDate = parseDate(patch?.tillDate);
    if (!flt) return { ok: false, error: '請填寫 FLT No' };
    if (!carrier) return { ok: false, error: '請選擇或填寫航司' };
    if (!isValidIsoDate(fromDate) || !isValidIsoDate(tillDate)) return { ok: false, error: '請填寫有效日期區間' };
    if (compareIsoDate(fromDate, tillDate) > 0) return { ok: false, error: '結束日不可早於起始日' };

    let scheduleId;
    const existing = findScheduleForNewRow(carrier, fromDate, tillDate);
    if (existing) {
      scheduleId = existing.scheduleId;
    } else {
      const created = upsertSchedule({
        carrier,
        seasonLabel: `${carrier} 手動維護`,
        seasonKind: 'manual',
        startDate: fromDate,
        endDate: tillDate,
        source: 'manual-add'
      }, []);
      if (!created.ok) return created;
      scheduleId = created.schedule.scheduleId;
    }

    const rows = loadRows(scheduleId);
    const next = {
      flt,
      fromDate,
      tillDate,
      dep: clean(patch?.dep).toUpperCase(),
      arr: clean(patch?.arr).toUpperCase(),
      std: patch?.std ?? '',
      sta: patch?.sta ?? '',
      schedule: clean(patch?.schedule),
      acType: clean(patch?.acType),
      connectingFlt: normalizeConnectingFlt(patch?.dep, patch?.arr, patch?.connectingFlt),
      createdAt: new Date().toISOString()
    };
    rows.push(next);
    saveRows(scheduleId, rows);
    const meta = loadMeta();
    const entry = meta.schedules.find(s => s.scheduleId === scheduleId);
    if (entry) {
      entry.rowCount = rows.length;
      entry.updatedAt = new Date().toISOString();
      saveMeta(meta);
    }
    return { ok: true, scheduleId, rowIndex: rows.length - 1, row: next };
  }

  function updateRowAt(scheduleId, rowIndex, patch) {
    const rows = loadRows(scheduleId);
    if (rowIndex < 0 || rowIndex >= rows.length) return { ok: false, error: '找不到航班列' };
    const current = rows[rowIndex];
    const Parser = global.ForeignScheduleParser;
    const parseDate = Parser && typeof Parser.parseExcelDateValue === 'function'
      ? Parser.parseExcelDateValue.bind(Parser)
      : (v) => clean(v);
    const next = {
      ...current,
      flt: clean(patch?.flt ?? current.flt),
      fromDate: parseDate(patch?.fromDate ?? current.fromDate),
      tillDate: parseDate(patch?.tillDate ?? current.tillDate),
      dep: clean(patch?.dep ?? current.dep).toUpperCase(),
      arr: clean(patch?.arr ?? current.arr).toUpperCase(),
      std: patch?.std ?? current.std,
      sta: patch?.sta ?? current.sta,
      schedule: clean(patch?.schedule ?? current.schedule),
      acType: clean(patch?.acType ?? current.acType),
      connectingFlt: normalizeConnectingFlt(
        patch?.dep ?? current.dep,
        patch?.arr ?? current.arr,
        patch?.connectingFlt ?? current.connectingFlt
      ),
      createdAt: current.createdAt || new Date().toISOString()
    };
    if (!next.flt) return { ok: false, error: '請填寫 FLT No' };
    rows[rowIndex] = next;
    saveRows(scheduleId, rows);
    const meta = loadMeta();
    const entry = meta.schedules.find(s => s.scheduleId === scheduleId);
    if (entry) entry.updatedAt = new Date().toISOString();
    saveMeta(meta);
    return { ok: true, row: next };
  }

  function listAllFlightRows() {
    migrateLegacyConnectingFltFromFixed();
    const Parser = global.ForeignScheduleParser;
    const infer = Parser && typeof Parser.inferCarrierFromFlt === 'function'
      ? Parser.inferCarrierFromFlt.bind(Parser)
      : () => '';
    const out = [];
    listSchedules().forEach(meta => {
      loadRows(meta.scheduleId).forEach((row, rowIndex) => {
        out.push({
          scheduleId: meta.scheduleId,
          rowIndex,
          carrier: infer(row.flt) || meta.carrier,
          seasonLabel: meta.seasonLabel,
          flt: row.flt,
          fromDate: row.fromDate || meta.startDate,
          tillDate: row.tillDate || meta.endDate,
          dep: row.dep,
          arr: row.arr,
          std: row.std,
          sta: row.sta,
          schedule: row.schedule,
          acType: row.acType,
          connectingFlt: row.connectingFlt || '',
          createdAt: row.createdAt || meta.updatedAt || ''
        });
      });
    });
    return out;
  }

  function upsertBundles(bundles) {
    const saved = [];
    const errors = [];
    (bundles || []).forEach(bundle => {
      const res = upsertSchedule({
        carrier: bundle.carrier,
        seasonLabel: bundle.seasonLabel,
        seasonKind: bundle.seasonKind || 'import',
        startDate: bundle.startDate,
        endDate: bundle.endDate,
        source: bundle.source || 'excel-import'
      }, bundle.rows || []);
      if (res.ok) saved.push(res.schedule);
      else errors.push(`${bundle.carrier} ${bundle.startDate || '?'}～${bundle.endDate || '?'}：${res.error}`);
    });
    return { ok: !errors.length, saved, errors };
  }

  global.ForeignScheduleRegistry = {
    REGISTRY_VERSION,
    KNOWN_CARRIERS,
    loadMeta,
    listSchedules,
    getSchedule,
    hasAnySchedules,
    validateScheduleEntry,
    resolveSchedulesForDate,
    upsertSchedule,
    deleteSchedule,
    loadRows,
    saveRows,
    filterRowsForCarrier,
    buildCandidatesForDate,
    previewForDate,
    collectRowsForAcTypeSync,
    compareIsoDate,
    isValidIsoDate,
    upsertBundles,
    removeRowAt,
    addRowAt,
    updateRowAt,
    listAllFlightRows,
    migrateLegacyConnectingFltFromFixed
  };
})(typeof window !== 'undefined' ? window : globalThis);
