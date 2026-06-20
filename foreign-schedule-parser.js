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

  function formatIsoDate(date) {
    const d = normalizeSelectedDate(date);
    if (!d) return '';
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    const da = String(d.getDate()).padStart(2, '0');
    return `${y}-${mo}-${da}`;
  }

  /** Excel From/Till：2026/3/29、ISO、或 Excel 序號 */
  function parseExcelDateValue(value) {
    if (value == null || value === '') return '';
    if (value instanceof Date && !Number.isNaN(value.getTime())) return formatIsoDate(value);
    if (typeof value === 'number' && Number.isFinite(value)) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(epoch.getTime() + Math.round(value) * 86400000);
      return formatIsoDate(d);
    }
    const s = clean(value);
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) return s;
    const slash = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (slash) {
      return `${slash[1]}-${String(Number(slash[2])).padStart(2, '0')}-${String(Number(slash[3])).padStart(2, '0')}`;
    }
    return formatIsoDate(s);
  }

  function compareIsoDate(a, b) {
    return String(a || '').localeCompare(String(b || ''));
  }

  function rowCoversDate(row, selectedDateIso) {
    const fromDate = clean(row.fromDate);
    const tillDate = clean(row.tillDate);
    if (fromDate && compareIsoDate(fromDate, selectedDateIso) > 0) return false;
    if (tillDate && compareIsoDate(selectedDateIso, tillDate) > 0) return false;
    return true;
  }

  function buildAutoSeasonLabel(carrier, startDate, endDate) {
    const c = clean(carrier).toUpperCase();
    const start = parseExcelDateValue(startDate);
    const end = parseExcelDateValue(endDate);
    if (!start || !end) return c ? `${c} 時刻表` : '時刻表';
    const sm = Number(start.slice(5, 7));
    const em = Number(end.slice(5, 7));
    let kind = '時刻表';
    if (sm >= 3 && sm <= 4 && em >= 9 && em <= 10) kind = '夏季';
    else if (sm >= 10 && em <= 3) kind = '冬季';
    const year = start.slice(0, 4);
    return `${c} ${year} ${kind}`;
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
        plusOne: false,
        raw: value
      };
    }
    const rawStr = clean(value);
    const plusOne = /\+1|\(\+1\)/i.test(rawStr);
    const s = rawStr.replace(/\s*\(?\+1\)?/gi, '').trim();
    if (/[／/]/.test(s)) return null;
    const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!m) return null;
    return {
      time: `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`,
      plusOne,
      raw: String(value)
    };
  }

  function formatTimeRaw(value) {
    const p = parseForeignTimeValue(value);
    if (!p?.time) return clean(value);
    return p.plusOne ? `${p.time}+1` : p.time;
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
      fromDate: parseExcelDateValue(raw.From ?? raw.from ?? raw['From'] ?? raw['起始'] ?? ''),
      tillDate: parseExcelDateValue(raw.Till ?? raw.till ?? raw['Till'] ?? raw['結束'] ?? ''),
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

  function datetimeOnScheduleDay(timeInput, selectedDate) {
    let timeStr = timeInput;
    let plusOne = false;
    if (timeInput && typeof timeInput === 'object' && timeInput.time) {
      timeStr = timeInput.time;
      plusOne = !!timeInput.plusOne;
    } else {
      const parsed = parseForeignTimeValue(timeInput);
      if (parsed) {
        timeStr = parsed.time;
        plusOne = !!parsed.plusOne;
      }
    }
    const mins = timeToMinutes(timeStr);
    if (mins == null || !timeStr) return null;
    const date = normalizeSelectedDate(selectedDate);
    if (!date) return null;
    const y = date.getFullYear();
    const mo = date.getMonth();
    const da = date.getDate();
    const hh = Math.floor(mins / 60);
    const mm = mins % 60;
    const dayOffset = plusOne ? 1 : 0;
    return {
      minutes: mins + dayOffset * 24 * 60,
      ms: new Date(y, mo, da + dayOffset, hh, mm, 0).getTime()
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
    let mins = timeToMinutes(anchorTime);
    if (type === 'ARR') {
      const parsedSta = parseForeignTimeValue(row.sta);
      if (parsedSta?.plusOne && mins != null) mins += 24 * 60;
    }
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
      sta: formatTimeRaw(row.sta),
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
      const connectingFlt = clean(row.connectingFlt);
      if (connectingFlt) {
        def.connectingFlight = normalizeFlightNo(connectingFlt);
        def.connectingSource = 'foreign-schedule';
      } else {
        def.connectingFlight = '';
        def.connectingSource = '';
      }
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

  function inferCarrierFromFlt(flt) {
    const s = clean(flt).toUpperCase().replace(/\s/g, '');
    const m = s.match(/^([A-Z0-9]{2})/);
    return m ? m[1] : '';
  }

  function parseWorkbookToNormalizedRows(workbook, opts = {}) {
    const rawRows = sheetRows(workbook);
    debugLog(opts, 'parseWorkbookToNormalizedRows', 'rawRowCount', rawRows.length);
    const normalizedRows = rawRows.map((raw, rawIndex) => normalizeWorkbookRow(raw, rawIndex));
    const rows = normalizedRows.filter(row => !row._skip);
    normalizedRows.filter(row => row._skip).forEach(row => {
      debugLog(opts, 'skipped-normalize', { rawIndex: row._rawIndex, reason: row._reason });
    });
    return rows;
  }

  function extractForeignAcTypeOptionsFromRows(rows) {
    const opts = { OZ: [], HX: [], BX: [], NZ: [] };
    (rows || []).forEach(row => {
      const carrier = inferCarrierFromFlt(row.flt);
      const v = clean(row.acType);
      if (!carrier || !opts[carrier] || !v || opts[carrier].includes(v)) return;
      opts[carrier].push(v);
    });
    return opts;
  }

  function detectScheduleBundlesFromRows(rows) {
    const map = new Map();
    (rows || []).forEach(row => {
      const carrier = inferCarrierFromFlt(row.flt);
      if (!carrier) return;
      const startDate = clean(row.fromDate);
      const endDate = clean(row.tillDate);
      const bundleKey = `${carrier}|${startDate}|${endDate}`;
      if (!map.has(bundleKey)) {
        map.set(bundleKey, {
          bundleKey,
          carrier,
          startDate,
          endDate,
          seasonLabel: buildAutoSeasonLabel(carrier, startDate, endDate),
          rows: []
        });
      }
      map.get(bundleKey).rows.push(row);
    });
    return [...map.values()].sort((a, b) => {
      const ca = compareIsoDate(a.carrier, b.carrier);
      if (ca) return ca;
      return compareIsoDate(a.startDate, b.startDate);
    });
  }

  function buildCandidatesFromNormalizedRows(rows, selectedDateIso, opts = {}) {
    const normalizeFlightNo = typeof opts.normalizeFlightNo === 'function'
      ? opts.normalizeFlightNo
      : (v) => clean(v).toUpperCase();

    const selectedDate = parseIsoDate(selectedDateIso);
    if (!selectedDate) {
      return { error: '系統排班日期無效，無法解析外家時刻表。', candidates: [], skipped: [], stats: { dep: 0, arr: 0, total: 0 } };
    }

    debugLog(opts, 'buildCandidatesFromNormalizedRows', 'selectedDate', selectedDateIso, 'rowCount', (rows || []).length);

    const candidates = [];
    const skipped = [];

    (rows || []).forEach((row, rowIndex) => {
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

      if (!rowCoversDate(row, selectedDateIso)) {
        skipped.push({ flt: row.flt, reason: '不在 From/Till 區間', fromDate: row.fromDate, tillDate: row.tillDate });
        logBx792('SKIP 不在 From/Till 區間', row.fromDate, row.tillDate);
        return;
      }

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
      let windowTimeInput = '';
      if (route.dep === STATION && route.arr !== STATION) {
        type = 'DEP';
        anchorTime = resolveDepStd(row);
        windowTimeInput = anchorTime;
      } else if (route.arr === STATION && route.dep !== STATION) {
        type = 'ARR';
        const parsedSta = parseForeignTimeValue(row.sta);
        anchorTime = parsedSta?.time || null;
        windowTimeInput = parsedSta || anchorTime;
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

      const inWindow = inOperationalWindow(type, windowTimeInput, selectedDate);
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
        tillDate: clean(row.tillDate),
        carrier: inferCarrierFromFlt(row.flt) || '',
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

  function parseForeignScheduleWorkbook(workbook, selectedDateIso, opts = {}) {
    const selectedDate = parseIsoDate(selectedDateIso);
    if (!selectedDate) {
      return { error: '系統排班日期無效，無法解析外家時刻表。', candidates: [] };
    }
    const rows = parseWorkbookToNormalizedRows(workbook, opts);
    return buildCandidatesFromNormalizedRows(rows, selectedDateIso, opts);
  }

  global.ForeignScheduleParser = {
    PARSER_VERSION,
    parseForeignScheduleWorkbook,
    parseWorkbookToNormalizedRows,
    buildCandidatesFromNormalizedRows,
    detectScheduleBundlesFromRows,
    extractForeignAcTypeOptions,
    extractForeignAcTypeOptionsFromRows,
    inferCarrierFromFlt,
    parseExcelDateValue,
    buildAutoSeasonLabel,
    rowCoversDate,
    parseForeignTimeValue,
    formatTimeRaw,
    matchesSchedule,
    inOperationalWindow,
    datetimeOnScheduleDay,
    normalizeSelectedDate,
    compareIsoDate,
    STATION
  };
})(typeof window !== 'undefined' ? window : globalThis);
