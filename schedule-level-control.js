/**
 * 管制表人員列：戰力排序與篩選（選單 UI）
 */
(function (global) {
  'use strict';

  /** @type {null | 'min-1A' | 'min-2A' | 'exact-2A' | 'min-m' | 'exact-M' | 'exact-student' | 'exact-DIC'} */
  let filterMode = null;
  /** @type {null | 'desc' | 'asc'} */
  let sortDir = null;
  /** @type {(() => void) | null} */
  let onChange = null;
  /** @type {((person: unknown) => number) | null} */
  let getPersonLevelRank = null;
  /** @type {((person: unknown) => string) | null} */
  let getPersonLevelKey = null;

  let menuOpen = false;
  let bound = false;

  const FILTER_LABELS = {
    'min-1A': '1A 以上',
    'min-2A': '2A 以上',
    'exact-2A': '只看 2A',
    'min-m': 'm 以上',
    'exact-M': '只看 M',
    'exact-student': '學生',
    'exact-DIC': 'DIC'
  };

  const MENU_ITEMS = [
    { action: 'all', label: '全部' },
    { type: 'sep' },
    { action: 'sort-desc', label: '排序：高→低' },
    { action: 'sort-asc', label: '排序：低→高' },
    { type: 'sep' },
    { action: 'filter-min-1A', label: '1A 以上' },
    { action: 'filter-min-2A', label: '2A 以上' },
    { action: 'filter-exact-2A', label: '只看 2A' },
    { action: 'filter-min-m', label: 'm 以上' },
    { action: 'filter-exact-M', label: '只看 M' },
    { action: 'filter-exact-student', label: '學生' },
    { action: 'filter-exact-DIC', label: 'DIC' }
  ];

  function usesDutyStatsLevelSecondary() {
    return !!(global.DutyStats?.isVisible?.() && global.DutyStats?.hasActiveSort?.());
  }

  function isRollCallBoard() {
    return !!document.getElementById('scheduleBoard')?.classList.contains('roll-call-mode');
  }

  function getMinRankThreshold(mode) {
    if (!getPersonLevelRank) return null;
    if (mode === 'min-1A') return getPersonLevelRank({ level: '1A' });
    if (mode === 'min-m') return getPersonLevelRank({ level: 'm' });
    if (mode === 'min-2A') return getPersonLevelRank({ level: '2A*' });
    return null;
  }

  function personMatchesLevelFilter(person) {
    if (!filterMode || !getPersonLevelRank) return true;
    const rank = getPersonLevelRank(person);
    const key = typeof getPersonLevelKey === 'function'
      ? getPersonLevelKey(person)
      : '';
    if (filterMode === 'exact-2A') return key === '2A';
    if (filterMode === 'exact-M') return key === 'M';
    if (filterMode === 'exact-student') return key === '學生';
    if (filterMode === 'exact-DIC') return key === 'DIC';
    const threshold = getMinRankThreshold(filterMode);
    if (threshold == null) return true;
    return rank <= threshold;
  }

  function compareRows(rowA, rowB, dir) {
    if (typeof getPersonLevelRank !== 'function') return 0;
    const rankA = getPersonLevelRank(rowA?.person);
    const rankB = getPersonLevelRank(rowB?.person);
    const diff = rankA - rankB;
    if (diff === 0) return 0;
    const normalizedDir = dir === 'asc' ? 'asc' : 'desc';
    return normalizedDir === 'asc' ? -diff : diff;
  }

  function hasActiveFilter() {
    return filterMode != null;
  }

  function hasActiveSort() {
    return sortDir != null && !usesDutyStatsLevelSecondary();
  }

  function hasActiveQuery() {
    if (isRollCallBoard()) return false;
    return hasActiveFilter() || hasActiveSort();
  }

  function applyToDisplayRows(rows) {
    if (!hasActiveQuery()) return rows;
    let result = (rows || []).filter(row => !row.isPinDivider);
    if (hasActiveFilter()) {
      result = result.filter(row => personMatchesLevelFilter(row.person));
    }
    if (hasActiveSort() && sortDir) {
      const dir = sortDir;
      result = [...result].sort((a, b) => {
        const diff = compareRows(a, b, dir);
        if (diff !== 0) return diff;
        return a.originalIndex - b.originalIndex;
      });
    }
    return result;
  }

  function getState() {
    return { filterMode, sortDir, menuOpen };
  }

  function setFilterMode(mode) {
    filterMode = mode || null;
  }

  function setSortDir(dir) {
    sortDir = dir === 'asc' || dir === 'desc' ? dir : null;
  }

  function clearAll(options = {}) {
    filterMode = null;
    sortDir = null;
    closeMenu();
    syncSummaryBar();
    if (!options.skipNotify && typeof onChange === 'function') onChange();
  }

  function applyMenuAction(action) {
    switch (action) {
      case 'all':
        clearAll({ skipNotify: true });
        break;
      case 'sort-desc':
        filterMode = null;
        sortDir = 'desc';
        break;
      case 'sort-asc':
        filterMode = null;
        sortDir = 'asc';
        break;
      case 'filter-exact-2A':
        filterMode = 'exact-2A';
        sortDir = null;
        break;
      case 'filter-min-1A':
        filterMode = 'min-1A';
        sortDir = null;
        break;
      case 'filter-min-2A':
        filterMode = 'min-2A';
        sortDir = null;
        break;
      case 'filter-min-m':
        filterMode = 'min-m';
        sortDir = null;
        break;
      case 'filter-exact-M':
        filterMode = 'exact-M';
        sortDir = null;
        break;
      case 'filter-exact-student':
        filterMode = 'exact-student';
        sortDir = null;
        break;
      case 'filter-exact-DIC':
        filterMode = 'exact-DIC';
        sortDir = null;
        break;
      default:
        return;
    }
    closeMenu();
    syncSummaryBar();
    if (typeof onChange === 'function') onChange();
  }

  function getSummaryText() {
    const parts = [];
    if (filterMode) parts.push(`戰力 ${FILTER_LABELS[filterMode]}`);
    if (hasActiveSort()) parts.push(`排序：${sortDir === 'asc' ? '低→高' : '高→低'}`);
    return parts.join(' · ');
  }

  function syncSummaryBar() {
    const bar = document.getElementById('scheduleLevelFilterBar');
    const label = document.getElementById('scheduleLevelFilterLabel');
    if (!bar || !label) return;
    const show = hasActiveQuery() && !isRollCallBoard();
    bar.classList.toggle('hidden', !show);
    label.textContent = show ? getSummaryText() : '';
  }

  function closeMenu() {
    menuOpen = false;
    document.querySelectorAll('.level-control-menu.is-open').forEach(el => {
      el.classList.remove('is-open');
      el.classList.add('hidden');
    });
    document.querySelectorAll('.level-control-trigger[aria-expanded="true"]').forEach(el => {
      el.setAttribute('aria-expanded', 'false');
    });
  }

  function toggleMenu(trigger, menu) {
    if (!trigger || !menu) return;
    const next = !menuOpen;
    closeMenu();
    if (!next) return;
    menuOpen = true;
    menu.classList.remove('hidden');
    menu.classList.add('is-open');
    trigger.setAttribute('aria-expanded', 'true');
  }

  function renderMenuHtml() {
    return MENU_ITEMS.map(item => {
      if (item.type === 'sep') return '<div class="level-control-menu-sep" role="separator"></div>';
      const active = (item.action === 'sort-desc' && sortDir === 'desc')
        || (item.action === 'sort-asc' && sortDir === 'asc')
        || (item.action === 'filter-min-1A' && filterMode === 'min-1A')
        || (item.action === 'filter-exact-2A' && filterMode === 'exact-2A')
        || (item.action === 'filter-min-2A' && filterMode === 'min-2A')
        || (item.action === 'filter-min-m' && filterMode === 'min-m')
        || (item.action === 'filter-exact-M' && filterMode === 'exact-M')
        || (item.action === 'filter-exact-student' && filterMode === 'exact-student')
        || (item.action === 'filter-exact-DIC' && filterMode === 'exact-DIC')
        || (item.action === 'all' && !filterMode && !sortDir);
      return `<button type="button" class="level-control-menu-item${active ? ' is-active' : ''}" role="menuitem" data-level-action="${item.action}">${item.label}</button>`;
    }).join('');
  }

  function getLevelHeaderHtml() {
    if (isRollCallBoard()) return '<span class="rc-h">戰力</span>';
    if (usesDutyStatsLevelSecondary()) {
      return global.DutyStats.getLevelHeaderHtml();
    }
    const sortMark = sortDir === 'asc' ? '↑' : sortDir === 'desc' ? '↓' : '';
    const filterMark = filterMode ? '●' : '';
    const marks = [sortMark, filterMark].filter(Boolean).join('');
    const markHtml = marks
      ? `<span class="level-control-marks" aria-hidden="true">${marks}</span>`
      : '';
    return [
      '<div class="level-control-head">',
      '<button type="button" class="level-control-trigger" aria-haspopup="menu" aria-expanded="false" aria-label="戰力排序與篩選">',
      '<span class="rc-h">戰力</span>',
      '<span class="level-control-chevron" aria-hidden="true">▾</span>',
      markHtml,
      '</button>',
      `<div class="level-control-menu hidden" role="menu">${renderMenuHtml()}</div>`,
      '</div>'
    ].join('');
  }

  function bindHeaderControls() {
    if (bound) return;
    const header = document.getElementById('nameColHeader');
    if (!header) return;
    bound = true;

    header.addEventListener('click', e => {
      if (usesDutyStatsLevelSecondary()) {
        const levelBtn = e.target.closest('.duty-stats-level-sort-btn');
        if (levelBtn) {
          e.preventDefault();
          global.DutyStats?.toggleLevelSort?.();
        }
        return;
      }
      const menuItem = e.target.closest('[data-level-action]');
      if (menuItem) {
        e.preventDefault();
        e.stopPropagation();
        applyMenuAction(menuItem.dataset.levelAction);
        return;
      }
      const trigger = e.target.closest('.level-control-trigger');
      if (trigger) {
        e.preventDefault();
        e.stopPropagation();
        const head = trigger.closest('.level-control-head');
        const menu = head?.querySelector('.level-control-menu');
        toggleMenu(trigger, menu);
      }
    });

    document.addEventListener('click', e => {
      if (!menuOpen) return;
      if (e.target.closest('.level-control-head')) return;
      closeMenu();
    });

    document.getElementById('clearScheduleLevelFilter')?.addEventListener('click', () => clearAll());
  }

  function init(options = {}) {
    onChange = typeof options.onChange === 'function' ? options.onChange : null;
    getPersonLevelRank = typeof options.getPersonLevelRank === 'function'
      ? options.getPersonLevelRank
      : null;
    getPersonLevelKey = typeof options.getPersonLevelKey === 'function'
      ? options.getPersonLevelKey
      : null;
    bindHeaderControls();
    syncSummaryBar();
  }

  function notifyBoardModeChanged() {
    closeMenu();
    syncSummaryBar();
  }

  global.ScheduleLevelControl = {
    personMatchesLevelFilter,
    compareRows,
    hasActiveFilter,
    hasActiveSort,
    hasActiveQuery,
    applyToDisplayRows,
    getState,
    getSummaryText,
    getLevelHeaderHtml,
    clearAll,
    applyMenuAction,
    syncSummaryBar,
    init,
    notifyBoardModeChanged
  };
})(typeof window !== 'undefined' ? window : globalThis);
