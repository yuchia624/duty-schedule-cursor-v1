/**
 * 外家時刻表 Excel → 當日候選 flightDefs
 * 需搭配 SheetJS（XLSX）；匯出 parseForeignScheduleWorkbook(workbook, selectedDateISO, opts)
 *
 * 支援欄位（新）：FLT No、DEP、ARR、STD、STA、Frequency、Aircraft Type
 * 相容舊版：航班、航線、班期、機型、備註
 */
(function (global) {
  const STATION = 'TPE';
  const DOW_ZH = ['日', '一', '二', '三', '四', '五', '六'];
  const PARSER_VERSION = '2026-05-27-bx792-debug';

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function isBx792(flt) {
    return /BX\s*0*792/i.test(clean(flt));
  }

  function shouldDebugForeignSchedule(opts) {
    if (opts && Object.prototype.hasOwnProperty.call(opts, 'debug')) return !!opts.debug;
    try {
      const v = localStorage.getItem('foreignScheduleDebug');
      if (v === '0') return false;
      if (v === '1') return true;
    } catch (_) {}
    return true;
  }

  function debugLog(opts, ...args) {
    if (!shouldDebugForeignSchedule(opts)) return;
    console.log('[ForeignScheduleParser]', ...args);
  }

  function normalizeSelectedDate(selectedDate) {
    if (!selectedDate) return null;
    if (selectedDate instanceof Date && !Number.isNaN(selectedDate.getTime())) {
      return new Date(selectedDate.getFullYear(), selectedDate.getMonth(), selectedDate.getDate());
    }
    if (typeof selectedDate === 'string') {
      const s = selectedDate.trim();
      const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
      const slash = s.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
      if (slash) return new Date(Number(slash[1]), Number(slash[2]) - 1, Number(slash[3]));
    }
    return null;
  }

  function parseIsoDate(iso) {
    return normalizeSelectedDate(iso);
  }

  function timeToMinutes(timeStr) {
    const [hh, mm] = String(timeStr || '').split(':').map(Number);
    if (Number.isNaN(hh) || Number.isNaN(mm)) return null;
    return hh * 60 + mm;
  }

  function parseForeignTimeValue(value) {
    if (value == null || value === '') return null;
    if (typeof value === 'number' && value >= 0 && value < 1) {
      const total = Math.round(value * 24 * 60);
      const hh = Math.floor(total / 60) % 24;
      const mm = total % 60;
      return {
        time: `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`,
        raw: value
      };
    }
    const s = clean(value).replace(/\+1.*$/i, '');
    if (/[／/]/.test(s)) return null;
    const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!m) return null;
    return {
      time: `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`,
      raw: String(value)
    };
  }

  function formatTimeRaw(value) {
    const p = parseForeignTimeValue(value);
    return p?.time || clean(value);
  }

  function parseRoute(routeStr) {
    const s = clean(routeStr).replace(/\s+/g, '');
    const m = s.match(/^([A-Z]{3})[→\-\>]+([A-Z]{3})$/i);
    if (!m) return null;
    return { dep: m[1].toUpperCase(), arr: m[2].toUpperCase() };
  }

  /** Excel Frequency 數字：週一=1 … 週日=7 */
  function excelWeekday(date) {
    const jsDay = date.getDay();
    return jsDay === 0 ? 7 : jsDay;
  }

  function matchesSchedule(scheduleStr, date) {
    const s = clean(scheduleStr).replace(/\*+$/, '');
    if (!s || s === '每日' || /^daily$/i.test(s)) return true;

    const tokens = s.split(/[,，、\s]+/).map(x => clean(x)).filter(Boolean);
    if (tokens.length && tokens.every(p => /^\d+$/.test(p))) {
      const dow = excelWeekday(date);
      return tokens.some(p => Number(p) === dow);
    }

    const dow = DOW_ZH[date.getDay()];
    if (tokens.length) return tokens.includes(dow);
    return s.includes(dow);
  }

  function resolveRoute(row) {
    const dep = clean(row.dep).toUpperCase();
    const arr = clean(row.arr).toUpperCase();
    if (dep && arr) {
      return { dep, arr, routeStr: `${dep} → ${arr}` };
    }
    const parsed = parseRoute(row.route);
    if (!parsed) return null;
    return {
      dep: parsed.dep,
      arr: parsed.arr,
      routeStr: clean(row.route)
    };
  }

  function defaultForeignAcType(flt) {
    const f = clean(flt).toUpperCase();
    if (f.startsWith('BX')) return '321';
    if (f.startsWith('NZ')) return '787';
    return '';
  }

  function normalizeWorkbookRow(raw, rawIndex) {
    const fltKeys = {
      'FLT No': raw['FLT No'],
      FLT: raw.FLT,
      航班: raw['航班'],
      flt: raw.flt
    };
    const flt = clean(fltKeys['FLT No'] || fltKeys.FLT || fltKeys['航班'] || fltKeys.flt);
    if (!flt || flt === 'FLT No' || flt === 'FLT' || flt === '航班') {
      return { _skip: true, _rawIndex: rawIndex, _raw: raw, _reason: '無效 FLT 或表頭列', _fltKeys: fltKeys };
    }
    const dep = clean(raw.DEP || raw.dep || '');
    const arr = clean(raw.ARR || raw.arr || '');
    const route = clean(raw['航線'] || raw.route || '');
    if (!dep && !arr && !route) {
      return { _skip: true, _rawIndex: rawIndex, _raw: raw, _reason: '缺少 DEP/ARR/航線', _fltKeys: fltKeys };
    }

    return {
      _rawIndex: rawIndex,
      _raw: raw,
      flt,
      dep,
      arr,
      route,
      std: raw.STD ?? raw.std ?? '',
      sta: raw.STA ?? raw.sta ?? '',
      schedule: clean(raw.Frequency || raw['班期'] || raw.schedule || ''),
      acType: clean(raw['Aircraft Type'] || raw['機型'] || raw.acType || ''),
      remark: clean(raw['備註'] || raw.remark || '')
    };
  }

  function resolveDepStd(row) {
    const parsed = parseForeignTimeValue(row.std);
    return parsed?.time || null;
  }

  function datetimeOnScheduleDay(timeStr, selectedDate) {
    const mins = timeToMinutes(timeStr);
    if (mins == null || !timeStr) return null;
    const date = normalizeSelectedDate(selectedDate);
    if (!date) return null;
    const y = date.getFullYear();
    const mo = date.getMonth();
    const da = date.getDate();
    const hh = Math.floor(mins / 60);
    const mm = mins % 60;
    return {
      minutes: mins,
      ms: new Date(y, mo, da, hh, mm, 0).getTime()
    };
  }

  /** 出境 STD：D 日 04:30（含）～ D+1 日 04:30（不含）；入境 STA：D 日 03:30（含）～ D+1 日 03:30（不含） */
  function inOperationalWindow(type, timeStr, selectedDate) {
    const dt = datetimeOnScheduleDay(timeStr, selectedDate);
    if (!dt) return false;
    const date = normalizeSelectedDate(selectedDate);
    if (!date) return false;
    const y = date.getFullYear();
    const mo = date.getMonth();
    const da = date.getDate();
    const startH = type === 'DEP' ? 4 : 3;
    const startM = 30;
    const startMs = new Date(y, mo, da, startH, startM, 0).getTime();
    const endMs = new Date(y, mo, da + 1, startH, startM, 0).getTime();
    return dt.ms >= startMs && dt.ms < endMs;
  }

  function sheetRows(workbook) {
    const name = workbook.SheetNames[0];
    if (!name) return [];
    return global.XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '' });
  }

  function buildForeignFlightDef(row, type, anchorTime, normalizeFlightNo, route) {
    const mins = timeToMinutes(anchorTime);
    if (mins == null) return null;
    const flight = normalizeFlightNo(row.flt);
    const def = {
      flight,
      type,
      baseTime: anchorTime,
      baseMinutes: mins,
      offset: 0,
      extension: 0,
      flt: clean(row.flt).toUpperCase(),
      acNo: '',
      acType: clean(row.acType) || defaultForeignAcType(row.flt),
      dep: route.dep,
      arr: route.arr,
      std: type === 'DEP' ? anchorTime : formatTimeRaw(row.std),
      sta: type === 'ARR' ? anchorTime : formatTimeRaw(row.sta),
      etd: '',
      etaRaw: '',
      depGate: '',
      arrGate: '',
      pax: '',
      serviceType: '',
      status: '',
      eChat: '',
      source: 'foreign',
      route: route.routeStr,
      schedule: clean(row.schedule),
      remark: clean(row.remark)
    };
    if (type === 'DEP') {
      def.changeTime = '';
      def.eta = '';
    } else {
      def.eta = '';
      def.etaManual = false;
      def.changeTime = '';
    }
    return def;
  }

  function extractForeignAcTypeOptions(workbook) {
    const opts = { OZ: [], HX: [], BX: [], NZ: [] };
    sheetRows(workbook).forEach(raw => {
      Object.keys(opts).forEach(code => {
        const v = clean(raw[code]);
        if (!v || opts[code].includes(v)) return;
        opts[code].push(v);
      });
    });
    return opts;
  }

  function parseForeignScheduleWorkbook(workbook, selectedDateIso, opts = {}) {
    const normalizeFlightNo = typeof opts.normalizeFlightNo === 'function'
      ? opts.normalizeFlightNo
      : (v) => clean(v).toUpperCase();

    const selectedDate = parseIsoDate(selectedDateIso);
    if (!selectedDate) {
      return { error: '系統排班日期無效，無法解析外家時刻表。', candidates: [] };
    }

    const rawRows = sheetRows(workbook);
    debugLog(opts, 'VERSION', PARSER_VERSION, 'selectedDate', selectedDateIso, 'excelWeekday', excelWeekday(selectedDate), 'rawRowCount', rawRows.length);

    rawRows.forEach((raw, rawIndex) => {
      if (!shouldDebugForeignSchedule(opts)) return;
      const entry = {
        rawIndex,
        raw,
        fltNo: raw['FLT No'],
        flt: raw.FLT,
        航班: raw['航班'],
        dep: raw.DEP,
        arr: raw.ARR,
        std: raw.STD,
        sta: raw.STA,
        frequency: raw.Frequency,
        acType: raw['Aircraft Type']
      };
      if (isBx792(entry.fltNo || entry.flt || entry.航班)) {
        console.log('[ForeignScheduleParser][BX792][raw-row]', entry);
      } else {
        debugLog(opts, 'raw-row', entry);
      }
    });

    const normalizedRows = rawRows.map((raw, rawIndex) => normalizeWorkbookRow(raw, rawIndex));
    const rows = normalizedRows.filter(row => !row._skip);
    normalizedRows.filter(row => row._skip).forEach(row => {
      const msg = { rawIndex: row._rawIndex, reason: row._reason, fltKeys: row._fltKeys, raw: row._raw };
      if (isBx792(row._raw?.['FLT No'] || row._raw?.FLT || row._raw?.['航班'])) {
        console.warn('[ForeignScheduleParser][BX792][skipped-normalize]', msg);
      } else {
        debugLog(opts, 'skipped-normalize', msg);
      }
    });

    debugLog(opts, 'normalizedRowCount', rows.length);

    const candidates = [];
    const skipped = [];

    rows.forEach((row, rowIndex) => {
      const baseLog = {
        rowIndex,
        rawIndex: row._rawIndex,
        flt: row.flt,
        dep: row.dep,
        arr: row.arr,
        std: row.std,
        sta: row.sta,
        schedule: row.schedule,
        acType: row.acType,
        selectedDate: selectedDateIso,
        excelWeekday: excelWeekday(selectedDate)
      };
      const bx792 = isBx792(row.flt);
      const logBx792 = (...args) => { if (bx792) console.log('[ForeignScheduleParser][BX792]', ...args); };

      const scheduleMatch = matchesSchedule(row.schedule, selectedDate);
      logBx792('schedule-check', { ...baseLog, scheduleMatch });
      debugLog(opts, 'row-check', { ...baseLog, scheduleMatch });

      if (!scheduleMatch) {
        skipped.push({ flt: row.flt, reason: '班期不符', schedule: row.schedule, excelWeekday: excelWeekday(selectedDate) });
        logBx792('SKIP 班期不符', row.schedule, 'weekday', excelWeekday(selectedDate));
        return;
      }

      const route = resolveRoute(row);
      if (!route) {
        skipped.push({ flt: row.flt, reason: '航線格式無法解析' });
        logBx792('SKIP 航線格式無法解析', row);
        return;
      }
      logBx792('route', route);

      let type = '';
      let anchorTime = '';
      if (route.dep === STATION && route.arr !== STATION) {
        type = 'DEP';
        anchorTime = resolveDepStd(row);
      } else if (route.arr === STATION && route.dep !== STATION) {
        type = 'ARR';
        const parsed = parseForeignTimeValue(row.sta);
        anchorTime = parsed?.time || null;
      } else {
        skipped.push({ flt: row.flt, reason: '非 TPE 出入境航線' });
        logBx792('SKIP 非 TPE 出入境航線', route);
        return;
      }

      logBx792('type/time', { type, anchorTime, std: row.std, sta: row.sta });

      if (!anchorTime) {
        skipped.push({ flt: row.flt, reason: '時間無法解析' });
        logBx792('SKIP 時間無法解析');
        return;
      }

      const inWindow = inOperationalWindow(type, anchorTime, selectedDate);
      logBx792('operational-window', { inWindow, type, anchorTime });
      if (!inWindow) {
        skipped.push({ flt: row.flt, reason: '不在當日營運時間窗口', type, anchorTime });
        logBx792('SKIP 不在當日營運時間窗口');
        return;
      }

      const flightDef = buildForeignFlightDef(row, type, anchorTime, normalizeFlightNo, route);
      if (!flightDef) {
        skipped.push({ flt: row.flt, reason: '無法建立航班資料' });
        logBx792('SKIP 無法建立 flightDef');
        return;
      }

      const candidate = {
        id: `fr_${row._rawIndex}_${flightDef.flight}_${type}`,
        flight: flightDef.flight,
        type,
        flt: flightDef.flt,
        dep: route.dep,
        arr: route.arr,
        route: flightDef.route,
        anchorTime,
        anchorLabel: type === 'DEP' ? 'STD' : 'STA',
        flightDef
      };
      candidates.push(candidate);
      logBx792('CANDIDATE ADDED', candidate);
      debugLog(opts, 'candidate-added', candidate);
    });

    candidates.sort((a, b) => {
      if (a.type !== b.type) return a.type === 'DEP' ? -1 : 1;
      return a.flightDef.baseMinutes - b.flightDef.baseMinutes || a.flight.localeCompare(b.flight);
    });

    const bx792Candidates = candidates.filter(c => /BX\s*0*792/i.test(c.flight) || /BX\s*0*792/i.test(c.flt));
    console.log('[ForeignScheduleParser] VERSION', PARSER_VERSION, 'summary', {
      selectedDate: selectedDateIso,
      excelWeekday: excelWeekday(selectedDate),
      candidates: candidates.length,
      dep: candidates.filter(c => c.type === 'DEP').length,
      arr: candidates.filter(c => c.type === 'ARR').length,
      skipped: skipped.length,
      bx792Candidates
    });
    if (bx792Candidates.length) {
      console.log('[ForeignScheduleParser][BX792] final candidates', bx792Candidates);
    } else {
      const bx792Skipped = skipped.filter(s => isBx792(s.flt));
      console.warn('[ForeignScheduleParser][BX792] NOT in candidates. skipped=', bx792Skipped);
    }

    return {
      candidates,
      skipped,
      stats: {
        dep: candidates.filter(c => c.type === 'DEP').length,
        arr: candidates.filter(c => c.type === 'ARR').length,
        total: candidates.length
      },
      parserVersion: PARSER_VERSION
    };
  }

  global.ForeignScheduleParser = {
    PARSER_VERSION,
    parseForeignScheduleWorkbook,
    extractForeignAcTypeOptions,
    matchesSchedule,
    inOperationalWindow,
    datetimeOnScheduleDay,
    normalizeSelectedDate,
    STATION
  };
})(typeof window !== 'undefined' ? window : globalThis);
