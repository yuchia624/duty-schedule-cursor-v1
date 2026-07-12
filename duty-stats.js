/**
 * 勤務統計：每人 RC / BG / PPT / T / 勤務總數
 */
(function (global) {
  'use strict';

  const STORAGE_KEY = 'cursor_v1_duty_stats_visible';
  const SORT_FIELDS = new Set(['rc', 'bg', 'ppt', 't', 'duty']);
  const SORT_LABELS = {
    rc: 'RC',
    bg: 'BG',
    ppt: 'PPT',
    t: 'T',
    duty: '勤'
  };

  /** @type {ReadonlySet<string>} */
  const EXCLUDED_OTHER_KINDS = new Set([
    'S',
    'TA',
    '特4hr',
    '補4hr',
    'B/F',
    'OJT',
    '設備',
    '關單',
    '備品',
    '值日生',
    '漏查單'
  ]);

  let visible = false;
  /** @type {(() => void) | null} */
  let onChange = null;
  /** @type {'rc'|'bg'|'ppt'|'t'|'duty'|null} */
  let sortField = null;
  /** @type {'desc'|'asc'} */
  let sortDir = 'desc';
  /** @type {'desc'|'asc'} 次排序：戰力（desc = M 等高戰力先） */
  let levelSortDir = 'desc';
  /** @type {((person: unknown) => number) | null} */
  let getPersonLevelRank = null;
  /** @type {{ rcMin: number|null, bgMin: number|null, pptMin: number|null, tMin: number|null, dutyMin: number|null }} */
  let filter = { rcMin: null, bgMin: null, pptMin: null, tMin: null, dutyMin: null };
  let filterBarBound = false;

  function emptyStats() {
    return { rc: 0, bg: 0, ppt: 0, t: 0, duty: 0 };
  }

  function readStoredVisible() {
    try {
      return localStorage.getItem(STORAGE_KEY) === '1';
    } catch (_) {
      return false;
    }
  }

  function writeStoredVisible(next) {
    try {
      localStorage.setItem(STORAGE_KEY, next ? '1' : '0');
    } catch (_) { /* ignore */ }
  }

  function isOtherKindExcluded(otherKind) {
    return EXCLUDED_OTHER_KINDS.has(String(otherKind || '').trim());
  }

  function isExcludedFromDutyTotal(item) {
    if (!item || typeof item !== 'object') return true;
    if (item.role === '休') return true;
    if (item.role === '其他') return isOtherKindExcluded(item.otherKind);
    return false;
  }

  function isRcCountItem(item) {
    return !!item && item.role === 'RC';
  }

  function isBgCountItem(item) {
    return !!item && item.role === 'BG';
  }

  function isPptCountItem(item) {
    return !!item && item.role === 'PPT';
  }

  function isTCountItem(item) {
    return !!item && item.role === 'T';
  }

  function isDutyTotalItem(item) {
    return !!item && !isExcludedFromDutyTotal(item);
  }

  function countPersonStats(schedule, personIndex) {
    const pid = Number(personIndex);
    if (!Number.isFinite(pid)) return emptyStats();
    const stats = emptyStats();
    (schedule || []).forEach(item => {
      if (item?.personIndex !== pid) return;
      if (isRcCountItem(item)) stats.rc += 1;
      if (isBgCountItem(item)) stats.bg += 1;
      if (isPptCountItem(item)) stats.ppt += 1;
      if (isTCountItem(item)) stats.t += 1;
      if (isDutyTotalItem(item)) stats.duty += 1;
    });
    return stats;
  }

  function buildStatsMap(schedule, people) {
    const map = new Map();
    (people || []).forEach((_, idx) => {
      map.set(idx, countPersonStats(schedule, idx));
    });
    return map;
  }

  function parseMinInput(value) {
    const s = String(value ?? '').trim();
    if (!s) return null;
    const n = Number(s);
    if (!Number.isFinite(n) || n < 0) return null;
    return Math.floor(n);
  }

  function hasActiveFilter(state = filter) {
    return state.rcMin != null
      || state.bgMin != null
      || state.pptMin != null
      || state.tMin != null
      || state.dutyMin != null;
  }

  function hasActiveSort() {
    return sortField != null;
  }

  function hasActiveQuery() {
    return visible && (hasActiveSort() || hasActiveFilter());
  }

  function matchesFilter(stats, state = filter) {
    const s = stats || emptyStats();
    if (state.rcMin != null && s.rc < state.rcMin) return false;
    if (state.bgMin != null && s.bg < state.bgMin) return false;
    if (state.pptMin != null && s.ppt < state.pptMin) return false;
    if (state.tMin != null && s.t < state.tMin) return false;
    if (state.dutyMin != null && s.duty < state.dutyMin) return false;
    return true;
  }

  function compareLevelRank(rowA, rowB) {
    if (typeof global.ScheduleLevelControl?.compareRows === 'function') {
      return global.ScheduleLevelControl.compareRows(rowA, rowB, levelSortDir);
    }
    if (typeof getPersonLevelRank !== 'function') return 0;
    const rankA = getPersonLevelRank(rowA?.person);
    const rankB = getPersonLevelRank(rowB?.person);
    const diff = rankA - rankB;
    if (diff === 0) return 0;
    return levelSortDir === 'asc' ? -diff : diff;
  }

  function applyToDisplayRows(rows, statsMap) {
    if (!visible) return rows;
    let result = (rows || []).filter(row => !row.isPinDivider);
    if (hasActiveFilter()) {
      result = result.filter(row => matchesFilter(statsMap.get(row.originalIndex)));
    }
    if (hasActiveSort() && sortField) {
      const dir = sortDir === 'asc' ? 1 : -1;
      result = [...result].sort((a, b) => {
        const sa = statsMap.get(a.originalIndex) || emptyStats();
        const sb = statsMap.get(b.originalIndex) || emptyStats();
        const diff = (sa[sortField] || 0) - (sb[sortField] || 0);
        if (diff !== 0) return diff * dir;
        const levelDiff = compareLevelRank(a, b);
        if (levelDiff !== 0) return levelDiff;
        return a.originalIndex - b.originalIndex;
      });
    }
    return result;
  }

  function getSort() {
    return { field: sortField, dir: sortDir };
  }

  function getFilter() {
    return { ...filter };
  }

  function setSort(field, dir) {
    sortField = SORT_FIELDS.has(field) ? field : null;
    sortDir = dir === 'asc' ? 'asc' : 'desc';
  }

  function toggleSort(field) {
    if (!SORT_FIELDS.has(field)) return;
    if (sortField !== field) {
      sortField = field;
      sortDir = 'desc';
    } else {
      sortDir = sortDir === 'desc' ? 'asc' : 'desc';
    }
    syncFilterBar();
    if (typeof onChange === 'function') onChange();
  }

  function setFilter(next) {
    filter = {
      rcMin: parseMinInput(next?.rcMin),
      bgMin: parseMinInput(next?.bgMin),
      pptMin: parseMinInput(next?.pptMin),
      tMin: parseMinInput(next?.tMin),
      dutyMin: parseMinInput(next?.dutyMin)
    };
  }

  function toggleLevelSort() {
    if (!hasActiveSort()) return;
    levelSortDir = levelSortDir === 'desc' ? 'asc' : 'desc';
    syncFilterBar();
    if (typeof onChange === 'function') onChange();
  }

  function clearQuery(options = {}) {
    sortField = null;
    sortDir = 'desc';
    levelSortDir = 'desc';
    filter = { rcMin: null, bgMin: null, pptMin: null, tMin: null, dutyMin: null };
    syncFilterBar();
    if (!options.skipNotify && typeof onChange === 'function') onChange();
  }

  function readFilterFromBar() {
    const bar = document.getElementById('dutyStatsFilterBar');
    if (!bar) return;
    setFilter({
      rcMin: bar.querySelector('[data-filter="rcMin"]')?.value,
      bgMin: bar.querySelector('[data-filter="bgMin"]')?.value,
      pptMin: bar.querySelector('[data-filter="pptMin"]')?.value,
      tMin: bar.querySelector('[data-filter="tMin"]')?.value,
      dutyMin: bar.querySelector('[data-filter="dutyMin"]')?.value
    });
  }

  function applyFilterFromInputs() {
    readFilterFromBar();
    syncFilterBar();
    if (typeof onChange === 'function') onChange();
  }

  function isVisible() {
    return visible;
  }

  function setVisible(next, options = {}) {
    const normalized = !!next;
    if (visible === normalized && !options.force) return;
    visible = normalized;
    writeStoredVisible(normalized);
    syncBoardClass();
    syncToolbarBtn();
    syncFilterBar();
    if (!options.skipNotify && typeof onChange === 'function') onChange();
  }

  function toggleVisible() {
    setVisible(!visible);
  }

  function syncBoardClass() {
    const board = document.getElementById('scheduleBoard');
    if (!board) return;
    board.classList.toggle('duty-stats-visible', visible && !board.classList.contains('roll-call-mode'));
  }

  function syncToolbarBtn() {
    const btn = document.getElementById('toggleDutyStatsBtn');
    if (!btn) return;
    btn.classList.toggle('active', visible);
    btn.setAttribute('aria-pressed', visible ? 'true' : 'false');
    const label = visible ? '收合勤務統計' : '勤務統計';
    btn.setAttribute('aria-label', label);
    const tip = btn.querySelector('.toolbar-icon-btn__tip');
    if (tip) tip.textContent = label;
    if (typeof global.setToolbarIconBtnLabel === 'function') {
      global.setToolbarIconBtnLabel(btn, label);
    }
  }

  function renderSortIndicator(field) {
    if (sortField !== field) return '';
    return `<span class="duty-stats-sort-ind" aria-hidden="true">${sortDir === 'asc' ? '↑' : '↓'}</span>`;
  }

  function renderSortHeader(field, label) {
    const active = sortField === field ? ' is-active' : '';
    const ariaSort = sortField === field
      ? (sortDir === 'asc' ? 'ascending' : 'descending')
      : 'none';
    return [
      `<div class="rc-col duty-stats-col duty-stats-col-${field}">`,
      `<button type="button" class="duty-stats-sort-btn${active}" data-sort-field="${field}" aria-label="依${label}排序" aria-sort="${ariaSort}">`,
      `<span class="rc-h">${label}</span>${renderSortIndicator(field)}`,
      '</button>',
      '</div>'
    ].join('');
  }

  function renderBadge(value, kind) {
    const n = Number(value);
    const safe = Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
    const zeroClass = safe === 0 ? ' is-zero' : '';
    return `<span class="duty-stats-badge duty-stats-badge-${kind}${zeroClass}">${safe}</span>`;
  }

  function getLevelHeaderHtml() {
    if (!visible || !hasActiveSort()) {
      return '<span class="rc-h">戰力</span>';
    }
    const hint = levelSortDir === 'asc' ? '戰力次排序（低→高）' : '戰力次排序（高→低）';
    return [
      '<button type="button" class="duty-stats-level-sort-btn is-active" aria-label="', hint, '">',
      '<span class="rc-h">戰力</span>',
      `<span class="duty-stats-sort-ind" aria-hidden="true">${levelSortDir === 'asc' ? '↑' : '↓'}</span>`,
      '</button>'
    ].join('');
  }

  function getHeaderHtml() {
    if (!visible) return '';
    return [
      renderSortHeader('rc', 'RC'),
      renderSortHeader('bg', 'BG'),
      renderSortHeader('ppt', 'PPT'),
      renderSortHeader('t', 'T'),
      renderSortHeader('duty', '勤')
    ].join('');
  }

  function getRowHtml(stats) {
    if (!visible) return '';
    const s = stats || emptyStats();
    return [
      `<div class="rc-col duty-stats-col duty-stats-col-rc">${renderBadge(s.rc, 'rc')}</div>`,
      `<div class="rc-col duty-stats-col duty-stats-col-bg">${renderBadge(s.bg, 'bg')}</div>`,
      `<div class="rc-col duty-stats-col duty-stats-col-ppt">${renderBadge(s.ppt, 'ppt')}</div>`,
      `<div class="rc-col duty-stats-col duty-stats-col-t">${renderBadge(s.t, 't')}</div>`,
      `<div class="rc-col duty-stats-col duty-stats-col-duty">${renderBadge(s.duty, 'duty')}</div>`
    ].join('');
  }

  function syncFilterBar() {
    const bar = document.getElementById('dutyStatsFilterBar');
    if (!bar) return;
    const show = visible && !document.getElementById('scheduleBoard')?.classList.contains('roll-call-mode');
    bar.classList.toggle('hidden', !show);
    bar.querySelector('[data-filter="rcMin"]').value = filter.rcMin != null ? String(filter.rcMin) : '';
    bar.querySelector('[data-filter="bgMin"]').value = filter.bgMin != null ? String(filter.bgMin) : '';
    bar.querySelector('[data-filter="pptMin"]').value = filter.pptMin != null ? String(filter.pptMin) : '';
    bar.querySelector('[data-filter="tMin"]').value = filter.tMin != null ? String(filter.tMin) : '';
    bar.querySelector('[data-filter="dutyMin"]').value = filter.dutyMin != null ? String(filter.dutyMin) : '';
    bar.classList.toggle('has-query', hasActiveQuery());
  }

  function bindFilterBar() {
    if (filterBarBound) return;
    const bar = document.getElementById('dutyStatsFilterBar');
    if (!bar) return;
    filterBarBound = true;
    bar.querySelector('#dutyStatsFilterClear')?.addEventListener('click', () => clearQuery());
    bar.querySelector('#dutyStatsFilterApply')?.addEventListener('click', () => applyFilterFromInputs());
    bar.querySelectorAll('.duty-stats-filter-input').forEach(input => {
      input.addEventListener('keydown', e => {
        if (e.key === 'Enter') {
          e.preventDefault();
          applyFilterFromInputs();
        }
      });
    });
  }

  function bindHeaderSortDelegation() {
    const header = document.getElementById('nameColHeader');
    if (!header || header.dataset.dutyStatsSortBound) return;
    header.dataset.dutyStatsSortBound = '1';
    header.addEventListener('click', e => {
      const btn = e.target.closest('.duty-stats-sort-btn');
      if (!btn) return;
      e.preventDefault();
      toggleSort(btn.dataset.sortField);
    });
  }

  function init(options = {}) {
    onChange = typeof options.onChange === 'function' ? options.onChange : null;
    getPersonLevelRank = typeof options.getPersonLevelRank === 'function'
      ? options.getPersonLevelRank
      : null;
    visible = readStoredVisible();
    bindFilterBar();
    bindHeaderSortDelegation();
    syncBoardClass();
    syncToolbarBtn();
    syncFilterBar();
    document.getElementById('toggleDutyStatsBtn')?.addEventListener('click', toggleVisible);
  }

  function notifyBoardModeChanged() {
    syncBoardClass();
    syncFilterBar();
  }

  function getFilterSummaryText() {
    const parts = [];
    if (filter.rcMin != null) parts.push(`RC ≥ ${filter.rcMin}`);
    if (filter.bgMin != null) parts.push(`BG ≥ ${filter.bgMin}`);
    if (filter.pptMin != null) parts.push(`PPT ≥ ${filter.pptMin}`);
    if (filter.tMin != null) parts.push(`T ≥ ${filter.tMin}`);
    if (filter.dutyMin != null) parts.push(`勤 ≥ ${filter.dutyMin}`);
    if (sortField) {
      const label = SORT_LABELS[sortField] || sortField;
      parts.push(`排序：${label}${sortDir === 'asc' ? '↑' : '↓'}`);
      parts.push(`戰力${levelSortDir === 'asc' ? '↑' : '↓'}`);
    }
    return parts.join(' · ');
  }

  global.DutyStats = {
    EXCLUDED_OTHER_KINDS,
    isOtherKindExcluded,
    isExcludedFromDutyTotal,
    isRcCountItem,
    isBgCountItem,
    isPptCountItem,
    isTCountItem,
    isDutyTotalItem,
    countPersonStats,
    buildStatsMap,
    parseMinInput,
    hasActiveFilter,
    hasActiveSort,
    hasActiveQuery,
    matchesFilter,
    applyToDisplayRows,
    getSort,
    getFilter,
    setSort,
    toggleSort,
    setFilter,
    clearQuery,
    isVisible,
    setVisible,
    toggleVisible,
    syncBoardClass,
    syncToolbarBtn,
    syncFilterBar,
    getHeaderHtml,
    getLevelHeaderHtml,
    getRowHtml,
    getFilterSummaryText,
    toggleLevelSort,
    init,
    notifyBoardModeChanged
  };
})(typeof window !== 'undefined' ? window : globalThis);
