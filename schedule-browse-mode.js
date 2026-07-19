/**
 * 手機班表瀏覽模式：≤768px 一律唯讀；桌機可編輯。
 *
 * 對外 API（掛在 window，供 index.html 呼叫）：
 *   isScheduleBrowseMode()
 *   canEditScheduleBoard()
 *   ensureScheduleEditable(actionHint?)
 *   setScheduleBrowseView('assignment'|'control')
 *   syncScheduleBrowseMode()
 *   initScheduleBrowseMode(deps?)
 *   setScheduleBrowseZoom(next) / adjustScheduleBrowseZoom(delta)
 *
 * init deps（可選）：
 *   openAssignmentPanel(forceOpen: boolean)
 *   clearAssignmentColumnUnlocks()
 *   onSearch(query: string)
 *   onZoomChange(zoom: number)
 */
(function (global) {
  'use strict';

  const SCHEDULE_BROWSE_MQ = global.matchMedia
    ? global.matchMedia('(max-width: 768px)')
    : { matches: false, addEventListener() {}, addListener() {} };

  const ZOOM_MIN = 0.5;
  const ZOOM_MAX = 1.5;
  const ZOOM_STEP = 0.1;

  let scheduleBrowseView = 'assignment';
  let browseZoom = 1;
  let deps = {
    openAssignmentPanel: null,
    clearAssignmentColumnUnlocks: null,
    onSearch: null,
    onZoomChange: null
  };

  function isScheduleBrowseMode() {
    return !!SCHEDULE_BROWSE_MQ.matches;
  }

  function canEditScheduleBoard() {
    return !isScheduleBrowseMode();
  }

  function ensureScheduleEditable(actionHint) {
    if (canEditScheduleBoard()) return true;
    const tip = actionHint ? `（${actionHint}）` : '';
    if (typeof global.appAlert === 'function') {
      global.appAlert(`手機瀏覽模式僅供查看${tip}，請使用電腦編輯班表。`, '唯讀', 'info');
    }
    return false;
  }

  function openAssignmentIfNeeded(forceOpen) {
    if (typeof deps.openAssignmentPanel === 'function') {
      deps.openAssignmentPanel(forceOpen);
      return;
    }
    if (typeof global.toggleAssignmentPanel === 'function') {
      global.toggleAssignmentPanel(forceOpen);
    }
  }

  function clearUnlocksIfNeeded() {
    if (typeof deps.clearAssignmentColumnUnlocks === 'function') {
      deps.clearAssignmentColumnUnlocks();
      return;
    }
    const set = global.assignmentUnlockedColumns;
    if (set && typeof set.clear === 'function' && set.size) {
      set.clear();
      if (typeof global.syncAssignmentColumnEditUi === 'function') {
        global.syncAssignmentColumnEditUi();
      }
    }
  }

  function hideWorkspaceSplitter() {
    const splitter = document.getElementById('workspaceSplitter');
    if (!splitter) return;
    splitter.hidden = true;
    splitter.setAttribute('aria-hidden', 'true');
    splitter.style.cssText = 'display:none!important;height:0!important;margin:0!important;padding:0!important;border:0!important;flex:0 0 0!important;overflow:hidden!important;pointer-events:none!important;visibility:hidden!important;';
  }

  function clampZoom(z) {
    const n = Number(z);
    if (!Number.isFinite(n)) return 1;
    return Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(n * 10) / 10));
  }

  function syncZoomLabel() {
    const label = document.getElementById('scheduleBrowseZoomLabel');
    if (label) label.textContent = `${Math.round(browseZoom * 100)}%`;
    const outBtn = document.getElementById('scheduleBrowseZoomOut');
    const inBtn = document.getElementById('scheduleBrowseZoomIn');
    if (outBtn) outBtn.disabled = browseZoom <= ZOOM_MIN + 1e-6;
    if (inBtn) inBtn.disabled = browseZoom >= ZOOM_MAX - 1e-6;
  }

  function setScheduleBrowseZoom(next) {
    browseZoom = clampZoom(next);
    document.documentElement.style.setProperty('--browse-zoom', String(browseZoom));
    syncZoomLabel();
    if (typeof deps.onZoomChange === 'function') deps.onZoomChange(browseZoom);
  }

  function adjustScheduleBrowseZoom(delta) {
    setScheduleBrowseZoom(browseZoom + Number(delta || 0));
  }

  function syncBrowseSearchFromTimeline() {
    const browseInput = document.getElementById('scheduleBrowseSearch');
    const timelineInput = document.getElementById('timelineFlightSearch');
    if (!browseInput) return;
    const v = timelineInput ? String(timelineInput.value || '') : '';
    if (document.activeElement !== browseInput) browseInput.value = v;
  }

  function setScheduleBrowseView(view) {
    const next = view === 'control' ? 'control' : 'assignment';
    scheduleBrowseView = next;
    document.body.classList.toggle('schedule-browse-view-assignment', next === 'assignment');
    document.body.classList.toggle('schedule-browse-view-control', next === 'control');
    document.querySelectorAll('.schedule-browse-tab').forEach(btn => {
      const active = btn.dataset.browseView === next;
      btn.classList.toggle('is-active', active);
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
    if (isScheduleBrowseMode() && next === 'assignment') {
      openAssignmentIfNeeded(true);
    }
    hideWorkspaceSplitter();
  }

  function syncScheduleBrowseMode() {
    const on = isScheduleBrowseMode();
    const wasOn = document.body.classList.contains('schedule-browse-mode');
    document.body.classList.toggle('schedule-browse-mode', on);
    if (!on) {
      document.body.classList.remove('schedule-browse-view-assignment', 'schedule-browse-view-control');
      document.documentElement.style.removeProperty('--browse-zoom');
      if (wasOn && typeof global.renderRows === 'function') global.renderRows();
      return;
    }
    clearUnlocksIfNeeded();
    setScheduleBrowseZoom(browseZoom || 1);
    setScheduleBrowseView(scheduleBrowseView || 'assignment');
    if (scheduleBrowseView === 'assignment') openAssignmentIfNeeded(true);
    hideWorkspaceSplitter();
    syncBrowseSearchFromTimeline();
    if (typeof global.updateShiftViewButtonLabel === 'function') {
      global.updateShiftViewButtonLabel();
    }
    // 進入瀏覽模式時重排，隱藏釘選置頂（桌機釘選狀態仍保留）
    if (!wasOn && typeof global.renderRows === 'function') global.renderRows();
  }

  function bindBrowseChrome() {
    document.getElementById('scheduleBrowseTabs')?.addEventListener('click', e => {
      const btn = e.target.closest('.schedule-browse-tab');
      if (!btn || !isScheduleBrowseMode()) return;
      setScheduleBrowseView(btn.dataset.browseView);
    });

    const browseInput = document.getElementById('scheduleBrowseSearch');
    if (browseInput && browseInput.dataset.browseBound !== '1') {
      browseInput.dataset.browseBound = '1';
      browseInput.addEventListener('input', () => {
        const q = browseInput.value || '';
        const timelineInput = document.getElementById('timelineFlightSearch');
        if (timelineInput && document.activeElement !== timelineInput) timelineInput.value = q;
        if (typeof deps.onSearch === 'function') deps.onSearch(q);
        else if (typeof global.scheduleTimelineFlightSearch === 'function') {
          global.scheduleTimelineFlightSearch(q);
        }
      });
    }

    document.getElementById('scheduleBrowseZoomOut')?.addEventListener('click', () => {
      if (!isScheduleBrowseMode()) return;
      adjustScheduleBrowseZoom(-ZOOM_STEP);
    });
    document.getElementById('scheduleBrowseZoomIn')?.addEventListener('click', () => {
      if (!isScheduleBrowseMode()) return;
      adjustScheduleBrowseZoom(ZOOM_STEP);
    });
  }

  function initScheduleBrowseMode(initDeps) {
    if (initScheduleBrowseMode._done) return;
    initScheduleBrowseMode._done = true;
    if (initDeps && typeof initDeps === 'object') {
      if (typeof initDeps.openAssignmentPanel === 'function') {
        deps.openAssignmentPanel = initDeps.openAssignmentPanel;
      }
      if (typeof initDeps.clearAssignmentColumnUnlocks === 'function') {
        deps.clearAssignmentColumnUnlocks = initDeps.clearAssignmentColumnUnlocks;
      }
      if (typeof initDeps.onSearch === 'function') deps.onSearch = initDeps.onSearch;
      if (typeof initDeps.onZoomChange === 'function') deps.onZoomChange = initDeps.onZoomChange;
    }
    bindBrowseChrome();
    const apply = () => syncScheduleBrowseMode();
    if (typeof SCHEDULE_BROWSE_MQ.addEventListener === 'function') {
      SCHEDULE_BROWSE_MQ.addEventListener('change', apply);
    } else if (typeof SCHEDULE_BROWSE_MQ.addListener === 'function') {
      SCHEDULE_BROWSE_MQ.addListener(apply);
    }
    apply();
  }

  global.isScheduleBrowseMode = isScheduleBrowseMode;
  global.canEditScheduleBoard = canEditScheduleBoard;
  global.ensureScheduleEditable = ensureScheduleEditable;
  global.setScheduleBrowseView = setScheduleBrowseView;
  global.syncScheduleBrowseMode = syncScheduleBrowseMode;
  global.initScheduleBrowseMode = initScheduleBrowseMode;
  global.setScheduleBrowseZoom = setScheduleBrowseZoom;
  global.adjustScheduleBrowseZoom = adjustScheduleBrowseZoom;
  global.ScheduleBrowseMode = {
    isActive: isScheduleBrowseMode,
    canEdit: canEditScheduleBoard,
    ensureEditable: ensureScheduleEditable,
    setView: setScheduleBrowseView,
    sync: syncScheduleBrowseMode,
    init: initScheduleBrowseMode,
    setZoom: setScheduleBrowseZoom,
    adjustZoom: adjustScheduleBrowseZoom
  };
})(typeof window !== 'undefined' ? window : globalThis);
