/**
 * 本家 EDW 訂位主檔：依日期分片存放，排班作業按需查詢
 */
(function (global) {
  const REGISTRY_VERSION = '2026-06-20-v5';
  const DEFAULT_RETENTION_DAYS = 45;
  const MIN_RETENTION_DAYS = 7;
  const META_KEY = 'cursor_v1_homeline_pax_meta_v2';
  const DAY_KEY_PREFIX = 'cursor_v1_homeline_pax_day_v2_';
  const LEGACY_META_KEY = 'cursor_v1_homeline_pax_meta_v1';
  const LEGACY_DAY_KEY_PREFIX = 'cursor_v1_homeline_pax_day_v1_';
  const STORAGE_MIGRATION_KEY = 'cursor_v1_homeline_pax_migrated_v2';

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function isValidIsoDate(s) {
    return /^\d{4}-\d{2}-\d{2}$/.test(clean(s));
  }

  function compareIsoDate(a, b) {
    return String(a || '').localeCompare(String(b || ''));
  }

  function dayStorageKey(dateIso) {
    return `${DAY_KEY_PREFIX}${clean(dateIso)}`;
  }

  function toIntOrNull(v) {
    if (Number.isFinite(v)) return v;
    const n = parseInt(clean(v).replace(/,/g, ''), 10);
    return Number.isFinite(n) ? n : null;
  }

  function normalizeLookupFlight(flight) {
    if (typeof HomelinePaxParser !== 'undefined' && HomelinePaxParser.normalizeLookupFlight) {
      return HomelinePaxParser.normalizeLookupFlight(flight);
    }
    const s = clean(flight).toUpperCase();
    if (!s) return s;
    if (s.startsWith('B7')) {
      const num = s.slice(2);
      if (/^\d+$/.test(num)) return `B7${num.padStart(4, '0')}`;
      return s;
    }
    if (s.startsWith('BR')) {
      const num = s.slice(2);
      if (/^\d+$/.test(num)) return `BR${num.padStart(4, '0')}`;
      return s;
    }
    const m = s.match(/^([A-Z]{2})(\d+)$/);
    if (m) return `${m[1]}${m[2].padStart(4, '0')}`;
    return s;
  }

  function purgeLegacyStorage() {
    try {
      const legacyMetaRaw = localStorage.getItem(LEGACY_META_KEY);
      if (legacyMetaRaw) {
        const legacyMeta = JSON.parse(legacyMetaRaw);
        (legacyMeta?.coveredDates || []).forEach(dateIso => {
          localStorage.removeItem(`${LEGACY_DAY_KEY_PREFIX}${dateIso}`);
        });
        localStorage.removeItem(LEGACY_META_KEY);
      }
      for (let i = localStorage.length - 1; i >= 0; i -= 1) {
        const key = localStorage.key(i);
        if (key && key.startsWith(LEGACY_DAY_KEY_PREFIX)) localStorage.removeItem(key);
      }
    } catch (_) { /* ignore */ }
  }

  function ensureStorageMigration() {
    if (localStorage.getItem(STORAGE_MIGRATION_KEY) === REGISTRY_VERSION) return;
    purgeLegacyStorage();
    localStorage.setItem(STORAGE_MIGRATION_KEY, REGISTRY_VERSION);
  }

  function loadMeta() {
    ensureStorageMigration();
    try {
      const raw = localStorage.getItem(META_KEY);
      if (!raw) {
        return { version: REGISTRY_VERSION, imports: [], coveredDates: [] };
      }
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object') {
        return { version: REGISTRY_VERSION, imports: [], coveredDates: [] };
      }
      if (!Array.isArray(parsed.imports)) parsed.imports = [];
      if (!Array.isArray(parsed.coveredDates)) parsed.coveredDates = [];
      return parsed;
    } catch (_) {
      return { version: REGISTRY_VERSION, imports: [], coveredDates: [] };
    }
  }

  function pruneCoveredDates(dates, maxDays = DEFAULT_RETENTION_DAYS) {
    const sorted = [...new Set((dates || []).filter(isValidIsoDate))].sort();
    const limit = Math.max(MIN_RETENTION_DAYS, maxDays || DEFAULT_RETENTION_DAYS);
    if (sorted.length <= limit) return sorted;
    return sorted.slice(-limit);
  }

  function pruneLocalToRetention(maxDays = DEFAULT_RETENTION_DAYS) {
    const meta = loadMeta();
    const nextDates = pruneCoveredDates(meta.coveredDates, maxDays);
    const keep = new Set(nextDates);
    (meta.coveredDates || []).forEach(dateIso => {
      if (!keep.has(dateIso)) deleteDay(dateIso);
    });
    meta.coveredDates = nextDates;
    return saveMeta(meta);
  }

  function saveMeta(meta) {
    const coveredDates = pruneCoveredDates(meta?.coveredDates, meta?.retentionDays || DEFAULT_RETENTION_DAYS);
    const next = {
      version: REGISTRY_VERSION,
      updatedAt: new Date().toISOString(),
      retentionDays: Number.isFinite(meta?.retentionDays) ? meta.retentionDays : DEFAULT_RETENTION_DAYS,
      imports: Array.isArray(meta?.imports) ? meta.imports.slice(-20) : [],
      coveredDates
    };
    localStorage.setItem(META_KEY, JSON.stringify(next));
    return next;
  }

  function computeDirectPax(pax, transit, group) {
    if (!Number.isFinite(pax)) return null;
    const t = Number.isFinite(transit) ? transit : 0;
    const g = Number.isFinite(group) ? group : 0;
    return Math.max(0, pax - t - g);
  }

  function computeWheelchair(wchc, wchr, wchs) {
    const vals = [wchc, wchr, wchs].filter(Number.isFinite);
    return vals.length ? vals.reduce((a, b) => a + b, 0) : null;
  }

  function computeLoadFactor(pax, capacity) {
    if (!Number.isFinite(pax) || !Number.isFinite(capacity) || capacity <= 0) return null;
    return Math.round((pax / capacity) * 1000) / 1000;
  }

  function enrichDayRow(row) {
    if (!row || typeof row !== 'object') return row;
    const pax = toIntOrNull(row.pax);
    const capacity = toIntOrNull(row.capacity);
    const transit = toIntOrNull(row.transit);
    const group = toIntOrNull(row.group);
    const wchc = toIntOrNull(row.wchc);
    const wchr = toIntOrNull(row.wchr);
    const wchs = toIntOrNull(row.wchs);
    const wheelchairFromParts = computeWheelchair(wchc, wchr, wchs);
    const wheelchair = wheelchairFromParts !== null
      ? wheelchairFromParts
      : toIntOrNull(row.wheelchair);
    const directPax = computeDirectPax(pax, transit, group);
    return {
      ...row,
      pax,
      capacity,
      transit,
      group,
      wchc,
      wchr,
      wchs,
      directPax,
      wheelchair,
      loadFactor: computeLoadFactor(pax, capacity)
    };
  }

  function loadDay(dateIso) {
    if (!isValidIsoDate(dateIso)) return [];
    try {
      const raw = localStorage.getItem(dayStorageKey(dateIso));
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed.map(enrichDayRow);
    } catch (_) {
      return [];
    }
  }

  function saveDay(dateIso, rows) {
    if (!isValidIsoDate(dateIso)) return;
    try {
      localStorage.setItem(dayStorageKey(dateIso), JSON.stringify(rows || []));
    } catch (err) {
      throw new Error(`訂位資料儲存失敗（可能超出瀏覽器容量）：${err?.message || err}`);
    }
  }

  function deleteDay(dateIso) {
    if (!isValidIsoDate(dateIso)) return;
    localStorage.removeItem(dayStorageKey(dateIso));
  }

  function groupRowsByDate(rows) {
    const byDate = {};
    (rows || []).forEach(row => {
      const dateIso = clean(row?.dateIso);
      if (!isValidIsoDate(dateIso)) return;
      if (!byDate[dateIso]) byDate[dateIso] = [];
      byDate[dateIso].push(row);
    });
    return byDate;
  }

  function serializeDayRow(row, sourceImportId) {
    const enriched = enrichDayRow(row);
    return {
      flight: normalizeLookupFlight(enriched.flight),
      carrier: clean(enriched.carrier).toUpperCase(),
      fltNo: clean(enriched.fltNo),
      std: clean(enriched.std),
      dest: clean(enriched.dest).toUpperCase(),
      capacity: enriched.capacity,
      pax: enriched.pax,
      group: enriched.group,
      transit: enriched.transit,
      directPax: enriched.directPax,
      wchc: enriched.wchc,
      wchr: enriched.wchr,
      wchs: enriched.wchs,
      wheelchair: enriched.wheelchair,
      loadFactor: enriched.loadFactor,
      sourceImportId: sourceImportId || enriched.sourceImportId || ''
    };
  }

  function importRows(rows, importMeta = {}) {
    if (!Array.isArray(rows) || !rows.length) {
      return { ok: false, error: '沒有可匯入的訂位資料。' };
    }

    const importId = `edw-${Date.now().toString(36)}`;
    const byDate = groupRowsByDate(rows);
    const importedDates = Object.keys(byDate).sort();
    if (!importedDates.length) {
      return { ok: false, error: '找不到有效日期。' };
    }

    importedDates.forEach(dateIso => {
      const dayRows = byDate[dateIso]
        .map(r => serializeDayRow(r, importId))
        .sort((a, b) => {
          const ta = compareIsoDate(a.std, b.std);
          if (ta) return ta;
          return compareIsoDate(a.flight, b.flight);
        });
      saveDay(dateIso, dayRows);
    });

    const meta = loadMeta();
    const coveredSet = new Set(meta.coveredDates || []);
    importedDates.forEach(d => coveredSet.add(d));

    const record = {
      importId,
      fileName: clean(importMeta.fileName) || 'EDW 訂位報表',
      importedAt: new Date().toISOString(),
      parserVersion: typeof HomelinePaxParser !== 'undefined' ? HomelinePaxParser.PARSER_VERSION : '',
      dateMin: importedDates[0],
      dateMax: importedDates[importedDates.length - 1],
      dateCount: importedDates.length,
      rowCount: rows.length
    };
    meta.imports = [record, ...(meta.imports || [])];
    meta.coveredDates = [...coveredSet].sort();
    meta.retentionDays = DEFAULT_RETENTION_DAYS;
    saveMeta(meta);
    const savedMeta = pruneLocalToRetention(DEFAULT_RETENTION_DAYS);

    return {
      ok: true,
      importId,
      importedDates,
      meta: savedMeta,
      record
    };
  }

  function getCoverage() {
    const meta = loadMeta();
    const dates = meta.coveredDates || [];
    if (!dates.length) {
      return {
        dateCount: 0,
        dateMin: '',
        dateMax: '',
        lastImport: null
      };
    }
    const sorted = dates.slice().sort();
    return {
      dateCount: sorted.length,
      dateMin: sorted[0],
      dateMax: sorted[sorted.length - 1],
      lastImport: meta.imports?.[0] || null
    };
  }

  function lookup(dateIso, flight) {
    if (!isValidIsoDate(dateIso)) return null;
    const key = normalizeLookupFlight(flight);
    if (!key) return null;
    const dayRows = loadDay(dateIso);
    return dayRows.find(r => normalizeLookupFlight(r.flight) === key) || null;
  }

  function filterAndSortHomelineRows(rows, opts = {}) {
    let list = (rows || []).slice();
    const search = clean(opts.search).toUpperCase();
    if (search) {
      list = list.filter(r =>
        normalizeLookupFlight(r.flight).includes(search)
        || clean(r.dest).toUpperCase().includes(search)
      );
    }
    list.sort((a, b) => {
      const da = compareIsoDate(a.dateIso, b.dateIso);
      if (da) return da;
      const ta = compareIsoDate(a.std, b.std);
      if (ta) return ta;
      return compareIsoDate(a.flight, b.flight);
    });
    return list;
  }

  function listForDate(dateIso, opts = {}) {
    if (!isValidIsoDate(dateIso)) return [];
    const rows = loadDay(dateIso).map(r => ({ ...r, dateIso }));
    return filterAndSortHomelineRows(rows, opts);
  }

  function listForDateRange(startIso, endIso, opts = {}) {
    if (!isValidIsoDate(startIso)) return [];
    const end = isValidIsoDate(endIso) ? endIso : startIso;
    const from = compareIsoDate(startIso, end) <= 0 ? startIso : end;
    const to = compareIsoDate(startIso, end) <= 0 ? end : startIso;
    const meta = loadMeta();
    const dates = (meta.coveredDates || []).filter(
      d => compareIsoDate(d, from) >= 0 && compareIsoDate(d, to) <= 0
    );
    const rows = [];
    dates.forEach(dateIso => {
      loadDay(dateIso).forEach(row => rows.push({ ...row, dateIso }));
    });
    return filterAndSortHomelineRows(rows, opts);
  }

  function hasDataForDate(dateIso) {
    return loadDay(dateIso).length > 0;
  }

  function hasAnyData() {
    return (loadMeta().coveredDates || []).length > 0;
  }

  function buildRowsByDate(dates) {
    const rowsByDate = {};
    (dates || []).forEach(dateIso => {
      if (!isValidIsoDate(dateIso)) return;
      const dayRows = loadDay(dateIso).map(r => serializeDayRow(r, r.sourceImportId));
      if (dayRows.length) rowsByDate[dateIso] = dayRows;
    });
    return rowsByDate;
  }

  function exportMasterPayload(opts = {}) {
    const retentionDays = Number.isFinite(opts.retentionDays)
      ? Math.max(MIN_RETENTION_DAYS, opts.retentionDays)
      : DEFAULT_RETENTION_DAYS;
    const meta = loadMeta();
    const coveredDates = pruneCoveredDates(meta.coveredDates, retentionDays);
    const exportMeta = {
      version: REGISTRY_VERSION,
      updatedAt: meta.updatedAt || new Date().toISOString(),
      retentionDays,
      imports: Array.isArray(meta.imports) ? meta.imports.slice(-20) : [],
      coveredDates
    };
    return {
      version: REGISTRY_VERSION,
      retentionDays,
      meta: exportMeta,
      rowsByDate: buildRowsByDate(coveredDates)
    };
  }

  function importMasterPayload(payload) {
    if (!payload || typeof payload !== 'object') return { ok: false, error: '無效資料' };
    const rawMeta = payload.meta;
    if (!rawMeta || !Array.isArray(rawMeta.coveredDates)) return { ok: false, error: '無效主檔' };
    const retentionDays = Number.isFinite(payload.retentionDays)
      ? Math.max(MIN_RETENTION_DAYS, payload.retentionDays)
      : (Number.isFinite(rawMeta.retentionDays) ? rawMeta.retentionDays : DEFAULT_RETENTION_DAYS);
    const coveredDates = pruneCoveredDates(rawMeta.coveredDates, retentionDays);
    const rowsByDate = payload.rowsByDate && typeof payload.rowsByDate === 'object'
      ? payload.rowsByDate
      : {};
    const oldMeta = loadMeta();
    const keep = new Set(coveredDates);
    (oldMeta.coveredDates || []).forEach(dateIso => {
      if (!keep.has(dateIso)) deleteDay(dateIso);
    });
    let rowCount = 0;
    coveredDates.forEach(dateIso => {
      const rows = Array.isArray(rowsByDate[dateIso]) ? rowsByDate[dateIso] : [];
      const nextRows = rows
        .map(r => serializeDayRow(r, r.sourceImportId))
        .sort((a, b) => {
          const ta = compareIsoDate(a.std, b.std);
          if (ta) return ta;
          return compareIsoDate(a.flight, b.flight);
        });
      rowCount += nextRows.length;
      saveDay(dateIso, nextRows);
    });
    saveMeta({
      ...rawMeta,
      retentionDays,
      coveredDates
    });
    return { ok: true, dateCount: coveredDates.length, rowCount, retentionDays };
  }

  function clearAll() {
    const meta = loadMeta();
    (meta.coveredDates || []).forEach(deleteDay);
    localStorage.removeItem(META_KEY);
    return saveMeta({ version: REGISTRY_VERSION, imports: [], coveredDates: [] });
  }

  ensureStorageMigration();

  global.HomelinePaxRegistry = {
    REGISTRY_VERSION,
    DEFAULT_RETENTION_DAYS,
    MIN_RETENTION_DAYS,
    loadMeta,
    getCoverage,
    importRows,
    lookup,
    listForDate,
    listForDateRange,
    hasDataForDate,
    hasAnyData,
    loadDay,
    clearAll,
    pruneLocalToRetention,
    exportMasterPayload,
    importMasterPayload,
    normalizeLookupFlight,
    isValidIsoDate,
    enrichDayRow
  };
})(typeof window !== 'undefined' ? window : globalThis);
