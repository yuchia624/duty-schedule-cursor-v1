/**
 * 航點／航線分類主檔：依目的地 IATA 對應一或多個航線群（本家）
 */
(function (global) {
  const REGISTRY_VERSION = '2026-06-20-v5';
  const DATA_KEY = 'cursor_v1_dest_class_groups_v1';
  const CATALOG_META_KEY = 'cursor_v1_dest_class_catalog_meta_v1';

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function normalizeIata(code) {
    const s = clean(code).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!s) return '';
    if (s.length === 3) return s;
    const m = s.match(/^([A-Z0-9]{3})/);
    return m ? m[1] : '';
  }

  function parseIataList(text) {
    const raw = clean(text);
    if (!raw) return [];
    const parts = raw.split(/[\s,;、，\/]+/).map(normalizeIata).filter(Boolean);
    return [...new Set(parts)];
  }

  function slugFromLabel(label) {
    const base = clean(label)
      .toLowerCase()
      .replace(/\s+/g, '-')
      .replace(/[^\w\u4e00-\u9fff-]/g, '')
      .slice(0, 32);
    return base || `group-${Date.now().toString(36)}`;
  }

  function makeGroupId(label, existingIds) {
    const base = slugFromLabel(label);
    if (!existingIds.has(base)) return base;
    let n = 2;
    while (existingIds.has(`${base}-${n}`)) n += 1;
    return `${base}-${n}`;
  }

  function normalizeAirports(list) {
    if (!Array.isArray(list)) return [];
    const out = [];
    const seen = new Set();
    list.forEach(item => {
      const code = normalizeIata(item);
      if (!code || seen.has(code)) return;
      seen.add(code);
      out.push(code);
    });
    return out.sort();
  }

  function loadGroupsRaw() {
    try {
      const raw = localStorage.getItem(DATA_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      return [];
    }
  }

  function resolveRegionId(entry) {
    const raw = clean(entry?.regionId);
    if (raw && global.DestClassSeed?.getRegion?.(raw)) return raw;
    const inferred = global.DestClassSeed?.resolveRegionId?.(clean(entry?.id));
    if (inferred && global.DestClassSeed?.getRegion?.(inferred)) return inferred;
    return '';
  }

  function normalizeGroup(entry) {
    if (!entry || typeof entry !== 'object') return null;
    const label = clean(entry.label);
    const id = clean(entry.id) || slugFromLabel(label);
    if (!id || !label) return null;
    return {
      id,
      label,
      icon: clean(entry.icon) || '',
      regionId: resolveRegionId(entry),
      airports: normalizeAirports(entry.airports),
      sortOrder: Number.isFinite(entry.sortOrder) ? entry.sortOrder : 0,
      updatedAt: clean(entry.updatedAt) || new Date().toISOString()
    };
  }

  function loadCatalogMeta() {
    try {
      const raw = localStorage.getItem(CATALOG_META_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (_) {
      return {};
    }
  }

  function saveCatalogMeta(patch) {
    const next = { ...loadCatalogMeta(), ...patch, at: new Date().toISOString() };
    localStorage.setItem(CATALOG_META_KEY, JSON.stringify(next));
    return next;
  }

  function listDeletedGroupIds() {
    const meta = loadCatalogMeta();
    return Array.isArray(meta.deletedGroupIds)
      ? meta.deletedGroupIds.map(clean).filter(Boolean)
      : [];
  }

  function markGroupDeleted(id) {
    const key = clean(id);
    if (!key) return;
    const next = new Set(listDeletedGroupIds());
    next.add(key);
    saveCatalogMeta({ deletedGroupIds: [...next] });
  }

  function unmarkGroupDeleted(id) {
    const key = clean(id);
    if (!key) return;
    const next = listDeletedGroupIds().filter(item => item !== key);
    saveCatalogMeta({ deletedGroupIds: next });
  }

  function isGroupDeleted(id) {
    return listDeletedGroupIds().includes(clean(id));
  }

  function hasAnyGroups() {
    return listGroups().length > 0;
  }

  function saveGroups(groups) {
    const next = (groups || [])
      .map(normalizeGroup)
      .filter(Boolean)
      .sort((a, b) => {
        const so = (a.sortOrder || 0) - (b.sortOrder || 0);
        if (so) return so;
        return a.label.localeCompare(b.label, 'zh-Hant');
      });
    localStorage.setItem(DATA_KEY, JSON.stringify(next));
    return next;
  }

  function listGroups() {
    return loadGroupsRaw()
      .map(normalizeGroup)
      .filter(Boolean)
      .sort((a, b) => {
        const so = (a.sortOrder || 0) - (b.sortOrder || 0);
        if (so) return so;
        return a.label.localeCompare(b.label, 'zh-Hant');
      });
  }

  function getGroup(id) {
    const key = clean(id);
    if (!key) return null;
    return listGroups().find(g => g.id === key) || null;
  }

  function validateGroupInput(entry, excludeId) {
    const label = clean(entry?.label);
    if (!label) return '請填寫航線名稱';
    if (label.length > 40) return '航線名稱不可超過 40 字';
    const regionId = resolveRegionId(entry);
    if (!regionId) return '請選擇區域分類';
    const airports = normalizeAirports(entry?.airports);
    if (!airports.length) return '請至少加入一個 IATA 航點';
    const dup = listGroups().find(g =>
      g.label === label && (!excludeId || g.id !== excludeId)
    );
    if (dup) return `已有同名航線群「${label}」`;
    return null;
  }

  function upsertGroup(entry) {
    const label = clean(entry?.label);
    const airports = normalizeAirports(entry?.airports);
    const regionId = resolveRegionId(entry);
    const err = validateGroupInput({ label, airports, regionId }, entry?.id || null);
    if (err) return { ok: false, error: err };
    const groups = listGroups();
    const existingIds = new Set(groups.map(g => g.id));
    const now = new Date().toISOString();
    let id = clean(entry?.id);
    if (id && !getGroup(id)) id = '';
    if (!id) id = makeGroupId(label, existingIds);
    const patch = {
      id,
      label,
      icon: clean(entry?.icon) || getGroup(id)?.icon || '',
      regionId,
      airports,
      sortOrder: Number.isFinite(entry?.sortOrder) ? entry.sortOrder : groups.length,
      updatedAt: now
    };
    const idx = groups.findIndex(g => g.id === id);
    if (idx >= 0) groups[idx] = { ...groups[idx], ...patch };
    else groups.push(patch);
    unmarkGroupDeleted(id);
    saveGroups(groups);
    return { ok: true, group: patch };
  }

  function deleteGroup(id) {
    const key = clean(id);
    if (!key) return { ok: false, error: '找不到航線群' };
    const groups = listGroups();
    const next = groups.filter(g => g.id !== key);
    if (next.length === groups.length) return { ok: false, error: '找不到航線群' };
    markGroupDeleted(key);
    saveGroups(next);
    return { ok: true };
  }

  function lookupGroupsForAirport(code) {
    const iata = normalizeIata(code);
    if (!iata) return [];
    return listGroups().filter(g => g.airports.includes(iata));
  }

  function lookupLabelsForAirport(code) {
    return lookupGroupsForAirport(code).map(g => g.label);
  }

  function buildAirportIndex() {
    const index = new Map();
    listGroups().forEach(group => {
      group.airports.forEach(iata => {
        if (!index.has(iata)) index.set(iata, []);
        index.get(iata).push(group);
      });
    });
    return index;
  }

  function listAllAirports() {
    const set = new Set();
    listGroups().forEach(g => g.airports.forEach(iata => set.add(iata)));
    return [...set].sort();
  }

  function getStats() {
    const groups = listGroups();
    const airports = listAllAirports();
    return {
      groupCount: groups.length,
      airportCount: airports.length,
      version: REGISTRY_VERSION,
      updatedAt: groups.reduce((max, g) => {
        const t = clean(g.updatedAt);
        return t > max ? t : max;
      }, '')
    };
  }

  function collectAirportsFromHomelinePax() {
    const set = new Set();
    if (typeof global.HomelinePaxRegistry === 'undefined') return [];
    const meta = global.HomelinePaxRegistry.loadMeta?.() || { coveredDates: [] };
    (meta.coveredDates || []).forEach(dateIso => {
      const rows = global.HomelinePaxRegistry.loadDay?.(dateIso) || [];
      rows.forEach(row => {
        const code = normalizeIata(row?.dest);
        if (code) set.add(code);
      });
    });
    return [...set].sort();
  }

  function collectAirportsFromForeignSchedule() {
    const set = new Set();
    if (typeof global.ForeignScheduleRegistry === 'undefined') return [];
    (global.ForeignScheduleRegistry.listAllFlightRows?.() || []).forEach(row => {
      [row.dep, row.arr].forEach(ap => {
        const code = normalizeIata(ap);
        if (code && code !== 'TPE') set.add(code);
      });
    });
    return [...set].sort();
  }

  function collectSuggestedAirports() {
    const set = new Set([
      ...collectAirportsFromHomelinePax(),
      ...collectAirportsFromForeignSchedule()
    ]);
    return [...set].sort();
  }

  function listUnclassifiedAirports() {
    const classified = new Set(listAllAirports());
    return collectSuggestedAirports().filter(iata => !classified.has(iata));
  }

  function replaceAllGroups(entries) {
    const now = new Date().toISOString();
    const next = (entries || [])
      .map((entry, idx) => normalizeGroup({
        ...entry,
        sortOrder: Number.isFinite(entry?.sortOrder) ? entry.sortOrder : idx,
        updatedAt: now
      }))
      .filter(Boolean);
    saveGroups(next);
    return { ok: true, count: next.length, groups: next };
  }

  function seedDefaultGroupsIfEmpty() {
    if (listGroups().length) return { ok: true, seeded: false };
    const seed = global.DestClassSeed?.HOMELINE_SUMMER_2026_GROUPS;
    if (!Array.isArray(seed) || !seed.length) return { ok: false, error: '找不到預設種子' };
    return { ...replaceAllGroups(seed), seeded: true };
  }

  function importGroupsFromAirports(airportList) {
    if (typeof global.DestClassSeed?.buildGroupsFromAirports !== 'function') {
      return { ok: false, error: '分類規則未載入' };
    }
    const groups = global.DestClassSeed.buildGroupsFromAirports(airportList);
    if (!groups.length) return { ok: false, error: '找不到可分類的航點' };
    return replaceAllGroups(groups);
  }

  /** 依現有航點補上新版規則中缺少的標準航線群 */
  function ensureCatalogGroups() {
    const airports = listAllAirports();
    if (!airports.length) return { ok: true, added: 0 };
    if (typeof global.DestClassSeed?.buildGroupsFromAirports !== 'function') {
      return { ok: false, error: '分類規則未載入' };
    }
    const built = global.DestClassSeed.buildGroupsFromAirports(airports);
    const existing = listGroups();
    const byId = new Map(existing.map(g => [g.id, g]));
    let added = 0;
    built.forEach(group => {
      if (byId.has(group.id)) return;
      if (isGroupDeleted(group.id)) return;
      byId.set(group.id, { ...group, updatedAt: new Date().toISOString() });
      added += 1;
    });
    if (!added) return { ok: true, added: 0 };
    saveGroups([...byId.values()]);
    return { ok: true, added };
  }

  /** 依夏季班表種子補上缺少的標準航線群（例如新加坡／吉隆坡／印尼／柬埔寨線） */
  function mergeMissingCatalogGroups() {
    const seedGroups = global.DestClassSeed?.HOMELINE_SUMMER_2026_GROUPS;
    if (!Array.isArray(seedGroups) || !seedGroups.length) return { ok: true, added: 0 };
    const byId = new Map(listGroups().map(g => [g.id, g]));
    let added = 0;
    seedGroups.forEach(group => {
      if (byId.has(group.id)) return;
      if (isGroupDeleted(group.id)) return;
      const normalized = normalizeGroup(group);
      if (!normalized) return;
      byId.set(normalized.id, normalized);
      added += 1;
    });
    if (!added) return { ok: true, added: 0 };
    saveGroups([...byId.values()]);
    return { ok: true, added };
  }

  /** 種子版本更新時，自動補上缺少的標準航線群 */
  function syncCatalogUpgrade() {
    const targetVersion = global.DestClassSeed?.SEED_VERSION || '';
    const meta = loadCatalogMeta();
    if (meta.seedVersion === targetVersion) return { ok: true, added: 0 };

    let added = 0;
    const fromAirports = ensureCatalogGroups();
    if (fromAirports.ok) added += fromAirports.added || 0;

    const fromSeed = mergeMissingCatalogGroups();
    if (fromSeed.ok) added += fromSeed.added || 0;

    saveCatalogMeta({ seedVersion: targetVersion });
    return { ok: true, added };
  }

  function exportMasterPayload() {
    return {
      version: REGISTRY_VERSION,
      groups: listGroups(),
      catalogMeta: loadCatalogMeta()
    };
  }

  function importMasterPayload(payload) {
    if (!payload || typeof payload !== 'object') return { ok: false, error: '無效資料' };
    const groups = Array.isArray(payload.groups)
      ? payload.groups.map(normalizeGroup).filter(Boolean)
      : [];
    const catalogMeta = payload.catalogMeta && typeof payload.catalogMeta === 'object'
      ? payload.catalogMeta
      : {};
    saveGroups(groups);
    const nextMeta = { ...catalogMeta };
    if (!nextMeta.at) nextMeta.at = new Date().toISOString();
    localStorage.setItem(CATALOG_META_KEY, JSON.stringify(nextMeta));
    return { ok: true, groupCount: groups.length };
  }

  global.DestClassRegistry = {
    REGISTRY_VERSION,
    normalizeIata,
    normalizeAirports,
    parseIataList,
    listGroups,
    getGroup,
    hasAnyGroups,
    upsertGroup,
    deleteGroup,
    lookupGroupsForAirport,
    lookupLabelsForAirport,
    buildAirportIndex,
    listAllAirports,
    getStats,
    collectSuggestedAirports,
    collectAirportsFromHomelinePax,
    collectAirportsFromForeignSchedule,
    listUnclassifiedAirports,
    replaceAllGroups,
    seedDefaultGroupsIfEmpty,
    importGroupsFromAirports,
    ensureCatalogGroups,
    mergeMissingCatalogGroups,
    syncCatalogUpgrade,
    saveCatalogMeta,
    resolveRegionId,
    exportMasterPayload,
    importMasterPayload
  };
})(typeof window !== 'undefined' ? window : globalThis);
