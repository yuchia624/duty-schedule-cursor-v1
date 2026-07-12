/**
 * 班表基準對照：純 diff 邏輯（零 DOM）
 */
(function (global) {
  'use strict';

  function formatMinutes(total) {
    if (total == null || !Number.isFinite(total)) return '—';
    total = ((total % 1440) + 1440) % 1440;
    const h = Math.floor(total / 60);
    const m = total % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }

  function formatTimeRange(start, end) {
    if (start == null || end == null) return '—';
    return `${formatMinutes(start)}–${formatMinutes(end)}`;
  }

  function slimScheduleItem(item) {
    if (!item || item.uid == null) return null;
    return {
      uid: item.uid,
      personIndex: item.personIndex,
      start: item.start,
      end: item.end,
      role: item.role,
      flight: item.flight,
      compactLabel: item.compactLabel,
      label: item.label,
      otherKind: item.otherKind,
      dicKind: item.dicKind,
      coveredFlights: Array.isArray(item.coveredFlights)
        ? item.coveredFlights.slice()
        : null
    };
  }

  function getDefaultDutyKey(item) {
    if (!item) return '';
    const role = String(item.role || '').trim();
    if (role === '休') return '休';
    if (role === '其他') {
      return `其他:${String(item.otherKind || item.compactLabel || item.label || '').trim()}`;
    }
    if (role === 'DIC') {
      return `DIC:${String(item.dicKind || item.compactLabel || item.label || '').trim()}`;
    }
    const flight = String(item.flight || '').trim();
    const label = String(item.compactLabel || item.label || '').trim();
    return `${role}:${flight}:${label}`;
  }

  function resolvePersonKey(item, getPersonKey) {
    if (item?.personKey != null && item.personKey !== '') return String(item.personKey);
    return getPersonKey(item.personIndex);
  }

  function normalizeDutyKey(key) {
    const s = String(key || '').trim();
    if (!s) return '';
    if (s.startsWith('其他:')) return s.slice(3).trim();
    if (s.startsWith('DIC:')) return s.slice(4).trim();
    return s;
  }

  function dutyKeysMatch(a, b) {
    const na = normalizeDutyKey(a);
    const nb = normalizeDutyKey(b);
    return na !== '' && na === nb;
  }

  function resolveDutyKey(item, getDutyKey) {
    if (item?.dutyKey != null && item.dutyKey !== '') return normalizeDutyKey(item.dutyKey);
    if (typeof getDutyKey === 'function') return normalizeDutyKey(getDutyKey(item));
    return normalizeDutyKey(getDefaultDutyKey(item));
  }

  function isTimeDiffExcludedItem(item) {
    const role = String(item?.role || '').trim();
    return role === '休' || role === 'PRE';
  }

  function personsMatch(a, b, getPersonKey) {
    if (a.personIndex === b.personIndex) return true;
    return resolvePersonKey(a, getPersonKey) === resolvePersonKey(b, getPersonKey);
  }

  function dutyItemsEquivalent(a, b, options = {}) {
    const getPersonKey = typeof options.getPersonKey === 'function'
      ? options.getPersonKey
      : (idx) => String(idx);
    if (!personsMatch(a, b, getPersonKey)) return false;
    if (Number(a.start) !== Number(b.start) || Number(a.end) !== Number(b.end)) return false;
    return dutyKeysMatch(resolveDutyKey(a, options.getDutyKey), resolveDutyKey(b, options.getDutyKey));
  }

  function dutyItemsSamePersonAndDuty(a, b, options = {}) {
    const getPersonKey = typeof options.getPersonKey === 'function'
      ? options.getPersonKey
      : (idx) => String(idx);
    if (!personsMatch(a, b, getPersonKey)) return false;
    return dutyKeysMatch(resolveDutyKey(a, options.getDutyKey), resolveDutyKey(b, options.getDutyKey));
  }

  function netOutEquivalentAddedRemoved(changes, options = {}) {
    const removed = [];
    const added = [];
    const other = [];
    changes.forEach(change => {
      if (change.type === 'removed') removed.push(change);
      else if (change.type === 'added') added.push(change);
      else other.push(change);
    });
    if (!removed.length || !added.length) return changes;

    const addedUsed = new Array(added.length).fill(false);
    const promoted = [];
    const keptRemoved = [];

    removed.forEach(rem => {
      let matched = false;
      for (let i = 0; i < added.length; i++) {
        if (addedUsed[i]) continue;
        if (dutyItemsEquivalent(rem.item, added[i].item, options)) {
          addedUsed[i] = true;
          matched = true;
          break;
        }
      }
      if (matched) return;

      for (let i = 0; i < added.length; i++) {
        if (addedUsed[i]) continue;
        if (!dutyItemsSamePersonAndDuty(rem.item, added[i].item, options)) continue;
        const add = added[i];
        addedUsed[i] = true;
        matched = true;
        if (isTimeDiffExcludedItem(rem.item)) break;
        const timeChanged = Number(rem.item.start) !== Number(add.item.start)
          || Number(rem.item.end) !== Number(add.item.end);
        if (timeChanged) {
          promoted.push({
            type: 'changed',
            uid: add.uid,
            base: rem.item,
            curr: add.item,
            personChanged: false,
            timeChanged: true,
            contentChanged: false
          });
        }
        break;
      }
      if (!matched) keptRemoved.push(rem);
    });

    const keptAdded = added.filter((_, i) => !addedUsed[i]);
    const merged = [...other, ...promoted, ...keptRemoved, ...keptAdded];
    merged.sort((a, b) => {
      const order = { changed: 0, added: 1, removed: 2 };
      const ta = order[a.type] ?? 9;
      const tb = order[b.type] ?? 9;
      if (ta !== tb) return ta - tb;
      return (a.uid || 0) - (b.uid || 0);
    });
    return merged;
  }

  /**
   * @param {ReturnType<typeof slimScheduleItem>[]} baseItems
   * @param {ReturnType<typeof slimScheduleItem>[]} currItems
   * @param {{ getPersonKey?: (personIndex: number) => string, getDutyKey?: (item: unknown) => string }} [options]
   */
  function diffSchedule(baseItems, currItems, options = {}) {
    const getPersonKey = typeof options.getPersonKey === 'function'
      ? options.getPersonKey
      : (idx) => String(idx);

    const baseMap = new Map();
    const currMap = new Map();
    baseItems.forEach(item => baseMap.set(item.uid, item));
    currItems.forEach(item => currMap.set(item.uid, item));

    /** @type {Array<Record<string, unknown>>} */
    const changes = [];

    for (const [uid, curr] of currMap) {
      const base = baseMap.get(uid);
      if (!base) {
        changes.push({ type: 'added', uid, item: curr });
        continue;
      }
      const personChanged = resolvePersonKey(base, getPersonKey) !== resolvePersonKey(curr, getPersonKey);
      const timeChanged = Number(base.start) !== Number(curr.start) || Number(base.end) !== Number(curr.end);
      const contentChanged = base.role !== curr.role
        || base.flight !== curr.flight
        || !arraysEqual(base.coveredFlights, curr.coveredFlights)
        || !dutyKeysMatch(resolveDutyKey(base, options.getDutyKey), resolveDutyKey(curr, options.getDutyKey));
      if (!personChanged && timeChanged && !contentChanged) {
        if (!isTimeDiffExcludedItem(base)) {
          changes.push({
            type: 'changed',
            uid,
            base,
            curr,
            personChanged: false,
            timeChanged: true,
            contentChanged: false
          });
        }
        continue;
      }
      if (personChanged || timeChanged || contentChanged) {
        changes.push({
          type: 'changed',
          uid,
          base,
          curr,
          personChanged,
          timeChanged,
          contentChanged
        });
      }
    }

    for (const [uid, base] of baseMap) {
      if (!currMap.has(uid)) {
        changes.push({ type: 'removed', uid, item: base });
      }
    }

    changes.sort((a, b) => {
      const order = { changed: 0, added: 1, removed: 2 };
      const ta = order[a.type] ?? 9;
      const tb = order[b.type] ?? 9;
      if (ta !== tb) return ta - tb;
      return (a.uid || 0) - (b.uid || 0);
    });

    return netOutEquivalentAddedRemoved(changes, options);
  }

  function slimSchedule(schedule) {
    if (!Array.isArray(schedule)) return [];
    return schedule.map(slimScheduleItem).filter(Boolean);
  }

  function arraysEqual(a, b) {
    if (!a && !b) return true;
    if (!a || !b || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
      if (a[i] !== b[i]) return false;
    }
    return true;
  }

  function getChangeKinds(change) {
    if (!change) return [];
    if (change.type === 'added') return ['added'];
    if (change.type === 'removed') return ['removed'];
    const kinds = [];
    if (change.personChanged) kinds.push('person');
    if (change.timeChanged) kinds.push('time');
    if (change.contentChanged) kinds.push('content');
    return kinds.length ? kinds : ['changed'];
  }

  function formatPersonRef(personIndex, helpers) {
    const getPersonName = typeof helpers.getPersonName === 'function'
      ? helpers.getPersonName
      : (idx) => `#${idx}`;
    const name = getPersonName(personIndex);
    const seq = typeof helpers.getPersonSeq === 'function'
      ? helpers.getPersonSeq(personIndex)
      : null;
    if (seq != null && seq !== '') return `序${seq} ${name}`;
    return name;
  }

  function formatChangeSummary(change, helpers = {}) {
    const getLabel = typeof helpers.getLabel === 'function'
      ? helpers.getLabel
      : (item) => item?.compactLabel || item?.flight || item?.role || '?';
    const primarySeq = (() => {
      if (change.type === 'added' || change.type === 'removed') {
        return typeof helpers.getPersonSeq === 'function'
          ? helpers.getPersonSeq(change.item.personIndex)
          : '';
      }
      if (change.type === 'changed') {
        return typeof helpers.getPersonSeq === 'function'
          ? helpers.getPersonSeq(change.curr.personIndex)
          : '';
      }
      return '';
    })();

    if (change.type === 'added') {
      return {
        kindLabel: '新增',
        seq: primarySeq,
        label: getLabel(change.item),
        detail: `${formatPersonRef(change.item.personIndex, helpers)} · ${formatTimeRange(change.item.start, change.item.end)}`
      };
    }
    if (change.type === 'removed') {
      return {
        kindLabel: '刪除',
        seq: primarySeq,
        label: getLabel(change.item),
        detail: `${formatPersonRef(change.item.personIndex, helpers)} · ${formatTimeRange(change.item.start, change.item.end)}`
      };
    }
    if (change.type === 'changed') {
      const parts = [];
      if (change.personChanged) {
        parts.push(`${formatPersonRef(change.base.personIndex, helpers)} → ${formatPersonRef(change.curr.personIndex, helpers)}`);
      } else if (change.timeChanged) {
        parts.push(formatPersonRef(change.curr.personIndex, helpers));
      }
      if (change.timeChanged) {
        parts.push(`${formatTimeRange(change.base.start, change.base.end)} → ${formatTimeRange(change.curr.start, change.curr.end)}`);
      }
      if (change.contentChanged && !change.personChanged && !change.timeChanged) {
        parts.push('內容變更');
      }
      let kindLabel = '變更';
      if (change.personChanged && !change.timeChanged && !change.contentChanged) kindLabel = '移人';
      else if (change.timeChanged && !change.personChanged && !change.contentChanged) kindLabel = '改時';
      else if (change.personChanged && change.timeChanged) kindLabel = '移人+改時';
      return {
        kindLabel,
        seq: primarySeq,
        label: getLabel(change.curr),
        detail: parts.join(' · ') || '—'
      };
    }
    return { kindLabel: '?', seq: '', label: '?', detail: '' };
  }

  function collectAffectedPersonIndices(changes) {
    const set = new Set();
    if (!Array.isArray(changes)) return set;
    changes.forEach(change => {
      if (change.type === 'added' || change.type === 'removed') {
        set.add(change.item.personIndex);
        return;
      }
      if (change.type === 'changed') {
        set.add(change.base.personIndex);
        set.add(change.curr.personIndex);
      }
    });
    return set;
  }

  function applyDiffRowSort(displayRows, affectedSet) {
    if (!Array.isArray(displayRows) || !displayRows.length || !affectedSet?.size) return displayRows;
    const pinned = [];
    const dividers = [];
    const affected = [];
    const rest = [];
    displayRows.forEach(row => {
      if (row.isPinDivider) {
        dividers.push(row);
        return;
      }
      if (row.pinned) {
        pinned.push(row);
        return;
      }
      if (affectedSet.has(row.originalIndex)) affected.push(row);
      else rest.push(row);
    });
    const out = [...pinned];
    if (dividers.length) out.push(dividers[0]);
    out.push(...affected, ...rest);
    return out;
  }

  function exportPersonCells(personIndex, helpers) {
    const seq = typeof helpers.getPersonSeq === 'function'
      ? helpers.getPersonSeq(personIndex)
      : '';
    const name = typeof helpers.getPersonName === 'function'
      ? helpers.getPersonName(personIndex)
      : `#${personIndex}`;
    return [seq === '' || seq == null ? '' : seq, name];
  }

  function buildExportSheetRows(changes, helpers, meta = {}) {
    const getLabel = typeof helpers.getLabel === 'function'
      ? helpers.getLabel
      : (item) => item?.compactLabel || item?.flight || item?.role || '?';
    const rows = [
      ['基準對照'],
      ['基準時間', meta.baselineTime || ''],
      ['班表日期', meta.dateLabel || ''],
      ['匯出時間', meta.exportTime || ''],
      [],
      ['類型', '序', '姓名', '勤務', '基準人', '基準序', '基準時間', '現況人', '現況序', '現況時間', '說明']
    ];
    changes.forEach(change => {
      const summary = formatChangeSummary(change, helpers);
      if (change.type === 'added') {
        const [seq, name] = exportPersonCells(change.item.personIndex, helpers);
        rows.push([
          summary.kindLabel,
          seq,
          name,
          getLabel(change.item),
          '',
          '',
          '',
          name,
          seq,
          formatTimeRange(change.item.start, change.item.end),
          summary.detail
        ]);
        return;
      }
      if (change.type === 'removed') {
        const [seq, name] = exportPersonCells(change.item.personIndex, helpers);
        rows.push([
          summary.kindLabel,
          seq,
          name,
          getLabel(change.item),
          name,
          seq,
          formatTimeRange(change.item.start, change.item.end),
          '',
          '',
          '',
          summary.detail
        ]);
        return;
      }
      if (change.type === 'changed') {
        const [baseSeq, baseName] = exportPersonCells(change.base.personIndex, helpers);
        const [currSeq, currName] = exportPersonCells(change.curr.personIndex, helpers);
        rows.push([
          summary.kindLabel,
          currSeq,
          currName,
          getLabel(change.curr),
          baseName,
          baseSeq,
          formatTimeRange(change.base.start, change.base.end),
          currName,
          currSeq,
          formatTimeRange(change.curr.start, change.curr.end),
          summary.detail
        ]);
      }
    });
    return rows;
  }

  global.ScheduleDiffCore = {
    formatMinutes,
    formatTimeRange,
    slimScheduleItem,
    slimSchedule,
    getDefaultDutyKey,
    normalizeDutyKey,
    dutyKeysMatch,
    isTimeDiffExcludedItem,
    diffSchedule,
    dutyItemsEquivalent,
    dutyItemsSamePersonAndDuty,
    netOutEquivalentAddedRemoved,
    getChangeKinds,
    formatPersonRef,
    formatChangeSummary,
    collectAffectedPersonIndices,
    applyDiffRowSort,
    buildExportSheetRows
  };
})(typeof window !== 'undefined' ? window : globalThis);
