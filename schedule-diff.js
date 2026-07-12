/**
 * 班表基準對照：基準版、差異清單、畫面高亮、變動列置頂
 */
(function (global) {
  'use strict';

  const Core = global.ScheduleDiffCore;
  if (!Core) return;

  /** @type {null | { items: ReturnType<typeof Core.slimScheduleItem>[], capturedAt: number }} */
  let baseline = null;
  /** @type {Array<Record<string, unknown>>} */
  let lastChanges = [];
  /** @type {Set<number>} */
  let affectedPersonIndices = new Set();
  let highlightActive = false;
  let diffViewActive = false;
  let uiBound = false;

  /** @type {(() => unknown) | null} */
  let getSnapshot = null;
  /** @type {((item: unknown) => string) | null} */
  let getDutyLabel = null;
  /** @type {((personIndex: number) => string) | null} */
  let getPersonName = null;
  /** @type {((personIndex: number) => string) | null} */
  let getPersonKey = null;
  /** @type {((personIndex: number) => number|string|null) | null} */
  let getPersonSeq = null;
  /** @type {((item: unknown) => { personIndex: number, left: number, width: number } | null) | null} */
  let getBlockLayout = null;
  /** @type {(() => string) | null} */
  let getExportFileDate = null;
  /** @type {(() => string) | null} */
  let getExportDateText = null;
  /** @type {(() => void) | null} */
  let onChange = null;

  function escapeHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function formatBaselineTime(ts) {
    const d = new Date(ts);
    if (Number.isNaN(d.getTime())) return '';
    return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
  }

  function getHelpers() {
    return {
      getLabel: (item) => (getDutyLabel ? getDutyLabel(item) : (item?.compactLabel || item?.role || '?')),
      getPersonName: (idx) => (getPersonName ? getPersonName(idx) : `#${idx}`),
      getPersonSeq: (idx) => (getPersonSeq ? getPersonSeq(idx) : null)
    };
  }

  function buildPersonKeyResolver(people) {
    if (typeof getPersonKey === 'function') {
      return (idx) => getPersonKey(idx);
    }
    return (idx) => {
      const p = Array.isArray(people) ? people[idx] : null;
      return String(p?.personnelId || p?.name || idx);
    };
  }

  function buildDiffItems(schedule, people) {
    const resolver = buildPersonKeyResolver(people);
    const diffOptions = {
      getPersonKey: resolver,
      getDutyKey: (item) => (getDutyLabel ? getDutyLabel(item) : Core.getDefaultDutyKey(item))
    };
    const items = [];
    (schedule || []).forEach(raw => {
      const slim = Core.slimScheduleItem(raw);
      if (!slim) return;
      slim.personKey = resolver(raw.personIndex);
      slim.dutyKey = Core.normalizeDutyKey(diffOptions.getDutyKey(raw));
      items.push(slim);
    });
    return { items, diffOptions };
  }

  function computeChangesFromSnapshot(snapshot) {
    if (!baseline || !snapshot) return [];
    const { items: currItems, diffOptions } = buildDiffItems(snapshot.schedule, snapshot.people);
    return Core.diffSchedule(baseline.items, currItems, diffOptions);
  }

  function computeChanges() {
    if (!getSnapshot) return [];
    return computeChangesFromSnapshot(getSnapshot());
  }

  function updateAffectedSet(changes) {
    affectedPersonIndices = Core.collectAffectedPersonIndices(changes);
  }

  function setBaseline() {
    if (!getSnapshot) return false;
    const snap = getSnapshot();
    const built = buildDiffItems(snap.schedule, snap.people);
    baseline = {
      items: built.items,
      capturedAt: Date.now()
    };
    lastChanges = [];
    affectedPersonIndices = new Set();
    exitDiffView(false);
    syncUi();
    return true;
  }

  function clearBaseline() {
    baseline = null;
    lastChanges = [];
    exitDiffView(true);
    closeModal();
    syncUi();
  }

  function exitDiffView(triggerRender) {
    diffViewActive = false;
    highlightActive = false;
    affectedPersonIndices = new Set();
    clearHighlights();
    clearRemovedGhosts();
    if (triggerRender && onChange) onChange();
  }

  function hasBaseline() {
    return !!baseline;
  }

  function isDiffViewActive() {
    return diffViewActive;
  }

  function isPersonAffected(personIndex) {
    return affectedPersonIndices.has(personIndex);
  }

  function getChanges() {
    return lastChanges.slice();
  }

  function isHighlightActive() {
    return highlightActive;
  }

  function clearHighlights() {
    document.querySelectorAll('.block.schedule-diff-added, .block.schedule-diff-changed').forEach(el => {
      el.classList.remove('schedule-diff-added', 'schedule-diff-changed', 'schedule-diff-flash');
    });
  }

  function clearRemovedGhosts() {
    document.querySelectorAll('.block.schedule-diff-ghost').forEach(el => el.remove());
  }

  function applyHighlights(changes) {
    clearHighlights();
    changes.forEach(change => {
      if (change.type === 'removed') return;
      const uid = change.uid;
      const cls = change.type === 'added' ? 'schedule-diff-added' : 'schedule-diff-changed';
      document.querySelectorAll(`.block[data-uid="${uid}"]`).forEach(el => el.classList.add(cls));
    });
    highlightActive = changes.some(c => c.type === 'added' || c.type === 'changed' || c.type === 'removed');
  }

  function appendDiffGhost(change, item, ghostKind, prefixLabel) {
    const layout = getBlockLayout(item);
    if (!layout) return;
    const timeline = document.querySelector(`.timeline[data-original-index="${layout.personIndex}"]`);
    if (!timeline) return;
    const label = getDutyLabel ? getDutyLabel(item) : (item.compactLabel || item.role || '?');
    const ghost = document.createElement('div');
    ghost.className = `block schedule-diff-ghost schedule-diff-${ghostKind}`;
    ghost.dataset.ghostUid = String(change.uid);
    ghost.style.left = `${layout.left}px`;
    ghost.style.width = `${layout.width}px`;
    ghost.innerHTML = `<div class="label schedule-diff-ghost-label">${prefixLabel} ${escapeHtml(label)}</div>`;
    ghost.addEventListener('click', e => {
      e.stopPropagation();
      focusGhost(change.uid);
    });
    timeline.appendChild(ghost);
  }

  function renderDiffGhosts(changes) {
    clearRemovedGhosts();
    if (!diffViewActive || !getBlockLayout) return;
    changes.forEach(change => {
      if (change.type === 'removed') {
        appendDiffGhost(change, change.item, 'removed', '刪除');
        return;
      }
      if (change.type === 'changed' && change.personChanged) {
        appendDiffGhost(change, change.base, 'moved-away', '移除');
        return;
      }
      if (change.type === 'changed' && change.timeChanged && !change.personChanged) {
        appendDiffGhost(change, change.base, 'time-moved', '改時');
      }
    });
  }

  function applyDiffVisuals(changes) {
    applyHighlights(changes);
    renderDiffGhosts(changes);
  }

  function focusGhost(uid) {
    const el = document.querySelector(`.block.schedule-diff-ghost[data-ghost-uid="${uid}"]`);
    if (!el) return;
    el.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    el.classList.add('schedule-diff-flash');
    window.setTimeout(() => el.classList.remove('schedule-diff-flash'), 1400);
  }

  function focusBlock(uid) {
    const el = document.querySelector(`.block[data-uid="${uid}"]`);
    if (!el) return;
    el.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    el.classList.add('schedule-diff-flash');
    window.setTimeout(() => el.classList.remove('schedule-diff-flash'), 1400);
  }

  function focusChange(change) {
    if (!change) return;
    if (change.type === 'removed') focusGhost(change.uid);
    else focusBlock(change.uid);
  }

  function closeHighlight() {
    exitDiffView(true);
    syncUi();
  }

  function applyToDisplayRows(displayRows) {
    if (!diffViewActive || !affectedPersonIndices.size) return displayRows;
    return Core.applyDiffRowSort(displayRows, affectedPersonIndices);
  }

  function exportDiffExcel() {
    if (!baseline || !lastChanges.length) {
      showBriefToast('沒有可匯出的變更');
      return;
    }
    if (typeof global.XLSX === 'undefined') {
      showBriefToast('Excel 元件未載入');
      return;
    }
    const now = new Date();
    const exportTime = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const rows = Core.buildExportSheetRows(lastChanges, getHelpers(), {
      baselineTime: formatBaselineTime(baseline.capturedAt),
      dateLabel: getExportDateText ? getExportDateText() : '',
      exportTime
    });
    const sheet = global.XLSX.utils.aoa_to_sheet(rows);
    const wb = global.XLSX.utils.book_new();
    global.XLSX.utils.book_append_sheet(wb, sheet, '基準對照');
    const fileDate = getExportFileDate ? getExportFileDate() : 'today';
    global.XLSX.writeFile(wb, `基準對照_${fileDate}.xlsx`);
    showBriefToast('已匯出 Excel');
  }

  function ensureUi() {
    if (document.getElementById('scheduleDiffBaselineBar')) return;

    const filterBars = document.querySelector('.duty-stats-filter-bar');
    const bar = document.createElement('div');
    bar.className = 'schedule-diff-baseline-bar hidden';
    bar.id = 'scheduleDiffBaselineBar';
    bar.setAttribute('aria-live', 'polite');
    bar.innerHTML = [
      '<span id="scheduleDiffBaselineLabel"></span>',
      '<div class="schedule-diff-baseline-bar__actions">',
      '<button type="button" class="schedule-diff-baseline-bar__btn schedule-diff-baseline-bar__btn--primary" id="scheduleDiffCompareBtn">對照基準</button>',
      '<button type="button" class="schedule-diff-baseline-bar__btn" id="scheduleDiffClearBaselineBtn">清除基準</button>',
      '</div>'
    ].join('');
    if (filterBars?.parentNode) {
      filterBars.parentNode.insertBefore(bar, filterBars.nextSibling);
    } else {
      document.body.appendChild(bar);
    }

    const modal = document.createElement('div');
    modal.id = 'scheduleDiffModal';
    modal.className = 'rollcall-import-dialog schedule-diff-modal';
    modal.hidden = true;
    modal.innerHTML = [
      '<div class="rollcall-import-backdrop" data-schedule-diff-dismiss="1"></div>',
      '<div class="rollcall-import-panel" role="dialog" aria-modal="true" aria-labelledby="scheduleDiffModalTitle">',
      '<h3 id="scheduleDiffModalTitle">基準對照</h3>',
      '<p class="schedule-diff-modal__meta" id="scheduleDiffModalMeta"></p>',
      '<p class="schedule-diff-modal__hint hidden" id="scheduleDiffModalHint">有變動的人員列已置頂；原位置會以紫框標示刪除/移除。</p>',
      '<div class="schedule-diff-modal__empty hidden" id="scheduleDiffModalEmpty">與基準版相同，沒有變更。</div>',
      '<ul class="schedule-diff-change-list" id="scheduleDiffChangeList"></ul>',
      '<div class="schedule-diff-modal__foot">',
      '<button type="button" class="schedule-diff-modal__btn" id="scheduleDiffExportBtn" hidden>匯出 Excel</button>',
      '<button type="button" class="schedule-diff-modal__btn" id="scheduleDiffCloseHighlightBtn" hidden>關閉對照</button>',
      '<button type="button" class="schedule-diff-modal__btn schedule-diff-modal__btn--primary" id="scheduleDiffModalCloseBtn">關閉</button>',
      '</div>',
      '</div>'
    ].join('');
    document.body.appendChild(modal);
  }

  const BASELINE_BTN_ICON = [
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">',
    '<path d="m19 21-7-5-7 5V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z"/>',
    '<path d="M12 7v6"/>',
    '<path d="M9 10h6"/>',
    '</svg>'
  ].join('');

  function ensureToolbarButtons() {
    const side = document.querySelector('.toolbar-schedule-side');
    if (!side) return;
    const overtimeBtn = document.getElementById('toggleOvertimeToolbarBtn');
    let btn = document.getElementById('scheduleDiffSetBaselineBtn');
    if (!btn) {
      btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'toolbar-icon-btn';
      btn.id = 'scheduleDiffSetBaselineBtn';
      btn.setAttribute('aria-label', '設為基準');
      btn.setAttribute('aria-pressed', 'false');
      btn.innerHTML = [
        '<span class="toolbar-icon-btn__icon" aria-hidden="true">',
        BASELINE_BTN_ICON,
        '</span>',
        '<span class="toolbar-icon-btn__tip">設為基準</span>'
      ].join('');
      side.appendChild(btn);
    } else {
      const icon = btn.querySelector('.toolbar-icon-btn__icon');
      if (icon) icon.innerHTML = BASELINE_BTN_ICON;
    }
    if (overtimeBtn && btn.previousElementSibling !== overtimeBtn) {
      overtimeBtn.insertAdjacentElement('afterend', btn);
    }
  }

  function syncToolbarBtn() {
    const btn = document.getElementById('scheduleDiffSetBaselineBtn');
    if (!btn) return;
    const active = !!baseline;
    btn.classList.toggle('schedule-diff-baseline-active', active);
    btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    const tip = btn.querySelector('.toolbar-icon-btn__tip');
    if (tip) {
      tip.textContent = active
        ? `基準：${formatBaselineTime(baseline.capturedAt)}`
        : '設為基準';
    }
    btn.title = active
      ? `基準已設定 ${formatBaselineTime(baseline.capturedAt)}（再按可更新）`
      : '將目前班表設為對照基準';
  }

  function syncBaselineBar() {
    const bar = document.getElementById('scheduleDiffBaselineBar');
    const label = document.getElementById('scheduleDiffBaselineLabel');
    const compareBtn = document.getElementById('scheduleDiffCompareBtn');
    if (!bar || !label) return;
    if (!baseline) {
      bar.classList.add('hidden');
      bar.classList.toggle('is-diff-view', false);
      return;
    }
    bar.classList.remove('hidden');
    bar.classList.toggle('is-diff-view', diffViewActive);
    const changeCount = computeChanges().length;
    const viewHint = diffViewActive ? ' · 變動列已置頂' : '';
    label.textContent = `基準 ${formatBaselineTime(baseline.capturedAt)}${changeCount ? ` · ${changeCount} 項變更${viewHint}` : ''}`;
    if (compareBtn) {
      compareBtn.textContent = diffViewActive ? '重新對照' : '對照基準';
    }
  }

  function renderChangeList(changes) {
    const list = document.getElementById('scheduleDiffChangeList');
    const empty = document.getElementById('scheduleDiffModalEmpty');
    const meta = document.getElementById('scheduleDiffModalMeta');
    const hint = document.getElementById('scheduleDiffModalHint');
    const closeHlBtn = document.getElementById('scheduleDiffCloseHighlightBtn');
    const exportBtn = document.getElementById('scheduleDiffExportBtn');
    if (!list || !empty || !meta) return;

    const helpers = getHelpers();
    meta.textContent = baseline
      ? `基準時間 ${formatBaselineTime(baseline.capturedAt)} · 共 ${changes.length} 項`
      : '';

    if (hint) hint.classList.toggle('hidden', !changes.length);

    if (!changes.length) {
      empty.classList.remove('hidden');
      list.innerHTML = '';
      if (closeHlBtn) closeHlBtn.hidden = true;
      if (exportBtn) exportBtn.hidden = true;
      return;
    }

    empty.classList.add('hidden');
    list.innerHTML = changes.map((change, idx) => {
      const summary = Core.formatChangeSummary(change, helpers);
      const kindClass = change.type === 'added'
        ? 'is-added'
        : change.type === 'removed'
          ? 'is-removed'
          : 'is-changed';
      const seqText = summary.seq !== '' && summary.seq != null ? String(summary.seq) : '—';
      return [
        `<li class="schedule-diff-change-item ${kindClass} is-clickable"`,
        ` data-change-index="${idx}" role="button" tabindex="0">`,
        `<span class="schedule-diff-change-kind">${escapeHtml(summary.kindLabel)}</span>`,
        `<span class="schedule-diff-change-seq">${escapeHtml(seqText)}</span>`,
        `<span class="schedule-diff-change-label">${escapeHtml(summary.label)}</span>`,
        `<span class="schedule-diff-change-detail">${escapeHtml(summary.detail)}</span>`,
        '</li>'
      ].join('');
    }).join('');

    if (closeHlBtn) closeHlBtn.hidden = !diffViewActive;
    if (exportBtn) exportBtn.hidden = false;
  }

  function enterDiffView() {
    if (!baseline) return;
    lastChanges = computeChanges();
    updateAffectedSet(lastChanges);
    diffViewActive = true;
    if (onChange) onChange();
    applyDiffVisuals(lastChanges);
    renderChangeList(lastChanges);
    syncUi();

    const modal = document.getElementById('scheduleDiffModal');
    if (modal) modal.hidden = false;
  }

  function openDiffPanel() {
    if (!baseline) return;
    ensureUi();
    enterDiffView();
  }

  function closeModal() {
    const modal = document.getElementById('scheduleDiffModal');
    if (modal) modal.hidden = true;
  }

  function syncUi() {
    syncToolbarBtn();
    syncBaselineBar();
    const board = document.getElementById('scheduleBoard');
    board?.classList.toggle('schedule-diff-view-active', diffViewActive);
  }

  function bindUi() {
    if (uiBound) return;
    ensureUi();
    ensureToolbarButtons();
    uiBound = true;

    document.getElementById('scheduleDiffSetBaselineBtn')?.addEventListener('click', () => {
      setBaseline();
      showBriefToast(baseline ? `已設基準（${formatBaselineTime(baseline.capturedAt)}）` : '設定失敗');
    });

    document.getElementById('scheduleDiffCompareBtn')?.addEventListener('click', openDiffPanel);

    document.getElementById('scheduleDiffClearBaselineBtn')?.addEventListener('click', () => {
      clearBaseline();
      showBriefToast('已清除基準');
    });

    document.getElementById('scheduleDiffModalCloseBtn')?.addEventListener('click', closeModal);

    document.getElementById('scheduleDiffCloseHighlightBtn')?.addEventListener('click', () => {
      closeHighlight();
      renderChangeList(lastChanges);
    });

    document.getElementById('scheduleDiffExportBtn')?.addEventListener('click', exportDiffExcel);

    document.getElementById('scheduleDiffModal')?.addEventListener('click', e => {
      if (e.target.closest('[data-schedule-diff-dismiss="1"]')) {
        closeModal();
        return;
      }
      const row = e.target.closest('[data-change-index]');
      if (row) {
        const change = lastChanges[Number(row.dataset.changeIndex)];
        focusChange(change);
      }
    });

    document.getElementById('scheduleDiffChangeList')?.addEventListener('keydown', e => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const row = e.target.closest('[data-change-index]');
      if (!row) return;
      e.preventDefault();
      const change = lastChanges[Number(row.dataset.changeIndex)];
      focusChange(change);
    });
  }

  function showBriefToast(msg) {
    if (typeof global.showScheduleBriefToast === 'function') {
      global.showScheduleBriefToast(msg, 1200);
      return;
    }
    console.log(msg);
  }

  function debugDump() {
    return {
      hasBaseline: !!baseline,
      diffViewActive,
      capturedAt: baseline ? new Date(baseline.capturedAt).toISOString() : null,
      baselineCount: baseline?.items?.length ?? 0,
      changeCount: computeChanges().length,
      affectedPeople: [...affectedPersonIndices],
      changes: computeChanges().map(c => Core.formatChangeSummary(c, getHelpers()))
    };
  }

  function init(options = {}) {
    getSnapshot = typeof options.getSnapshot === 'function' ? options.getSnapshot : null;
    getDutyLabel = typeof options.getDutyLabel === 'function' ? options.getDutyLabel : null;
    getPersonName = typeof options.getPersonName === 'function' ? options.getPersonName : null;
    getPersonKey = typeof options.getPersonKey === 'function' ? options.getPersonKey : null;
    getPersonSeq = typeof options.getPersonSeq === 'function' ? options.getPersonSeq : null;
    getBlockLayout = typeof options.getBlockLayout === 'function' ? options.getBlockLayout : null;
    getExportFileDate = typeof options.getExportFileDate === 'function' ? options.getExportFileDate : null;
    getExportDateText = typeof options.getExportDateText === 'function' ? options.getExportDateText : null;
    onChange = typeof options.onChange === 'function' ? options.onChange : null;
    bindUi();
    syncUi();
  }

  function notifyRowsRendered() {
    if (!diffViewActive || !lastChanges.length) return;
    applyDiffVisuals(lastChanges);
  }

  function notifyDateChanged() {
    clearBaseline();
  }

  global.ScheduleDiff = {
    init,
    setBaseline,
    clearBaseline,
    hasBaseline,
    openDiffPanel,
    closeModal,
    closeHighlight,
    isHighlightActive,
    isDiffViewActive,
    isPersonAffected,
    getChanges,
    computeChanges,
    applyToDisplayRows,
    exportDiffExcel,
    debugDump,
    notifyRowsRendered,
    notifyDateChanged
  };
})(typeof window !== 'undefined' ? window : globalThis);
