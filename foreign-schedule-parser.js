/**
 * 外家時刻表 Excel → 當日候選 flightDefs
 * 需搭配 SheetJS（XLSX）；匯出 parseForeignScheduleWorkbook(workbook, selectedDateISO, opts)
 */
(function (global) {
  const STATION = 'TPE';
  const DOW_ZH = ['日', '一', '二', '三', '四', '五', '六'];

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
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

  function matchesSchedule(scheduleStr, date) {
    const s = clean(scheduleStr).replace(/\*+$/, '');
    if (!s || s === '每日') return true;
    const dow = DOW_ZH[date.getDay()];
    const parts = s.split(/[、,，]/).map(x => clean(x)).filter(Boolean);
    if (parts.length) return parts.includes(dow);
    return s.includes(dow);
  }

  function resolveDepStd(row, date) {
    const flt = clean(row.flt).toUpperCase();
    const stdRaw = clean(row.std);
    if (flt === 'BX792' || /16:35/.test(stdRaw) && /16:40/.test(stdRaw)) {
      const d = date.getDay();
      if ([2, 4, 6].includes(d)) return '16:40';
      return '16:35';
    }
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

  function buildForeignFlightDef(row, type, anchorTime, normalizeFlightNo) {
    const mins = timeToMinutes(anchorTime);
    if (mins == null) return null;
    const flight = normalizeFlightNo(row.flt);
    const route = parseRoute(row.route) || { dep: '', arr: '' };
    const def = {
      flight,
      type,
      baseTime: anchorTime,
      baseMinutes: mins,
      offset: 0,
      extension: 0,
      flt: clean(row.flt).toUpperCase(),
      acNo: '',
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
      route: clean(row.route),
      schedule: clean(row.schedule),
      remark: clean(row.remark)
    };
    if (type === 'DEP') {
      def.changeTime = '';
      def.eta = '';
    } else {
      def.eta = '';
      def.changeTime = '';
    }
    return def;
  }

  function parseForeignScheduleWorkbook(workbook, selectedDateIso, opts = {}) {
    const normalizeFlightNo = typeof opts.normalizeFlightNo === 'function'
      ? opts.normalizeFlightNo
      : (v) => clean(v).toUpperCase();

    const selectedDate = parseIsoDate(selectedDateIso);
    if (!selectedDate) {
      return { error: '系統排班日期無效，無法解析外家時刻表。', candidates: [] };
    }

    const rows = sheetRows(workbook).map(row => ({
      flt: clean(row['航班'] || row.flt),
      route: clean(row['航線'] || row.route),
      std: row['STD'] ?? row.std ?? '',
      sta: row['STA'] ?? row.sta ?? '',
      schedule: clean(row['班期'] || row.schedule),
      remark: clean(row['備註'] || row.remark)
    })).filter(r => r.flt && r.route);

    const candidates = [];
    const skipped = [];

    rows.forEach(row => {
      if (!matchesSchedule(row.schedule, selectedDate)) {
        skipped.push({ flt: row.flt, reason: '班期不符' });
        return;
      }
      const route = parseRoute(row.route);
      if (!route) {
        skipped.push({ flt: row.flt, reason: '航線格式無法解析' });
        return;
      }

      let type = '';
      let anchorTime = '';
      if (route.dep === STATION && route.arr !== STATION) {
        type = 'DEP';
        anchorTime = resolveDepStd(row, selectedDate);
      } else if (route.arr === STATION && route.dep !== STATION) {
        type = 'ARR';
        const parsed = parseForeignTimeValue(row.sta);
        anchorTime = parsed?.time || null;
      } else {
        skipped.push({ flt: row.flt, reason: '非 TPE 出入境航線' });
        return;
      }

      if (!anchorTime) {
        skipped.push({ flt: row.flt, reason: '時間無法解析' });
        return;
      }
      if (!inOperationalWindow(type, anchorTime, selectedDate)) {
        skipped.push({ flt: row.flt, reason: '不在當日營運時間窗口' });
        return;
      }

      const flightDef = buildForeignFlightDef(row, type, anchorTime, normalizeFlightNo);
      if (!flightDef) {
        skipped.push({ flt: row.flt, reason: '無法建立航班資料' });
        return;
      }

      candidates.push({
        id: `${flightDef.flight}_${type}`,
        flight: flightDef.flight,
        type,
        flt: flightDef.flt,
        route: flightDef.route,
        anchorTime,
        anchorLabel: type === 'DEP' ? 'STD' : 'STA',
        flightDef
      });
    });

    candidates.sort((a, b) => {
      if (a.type !== b.type) return a.type === 'DEP' ? -1 : 1;
      return a.flightDef.baseMinutes - b.flightDef.baseMinutes || a.flight.localeCompare(b.flight);
    });

    return {
      candidates,
      skipped,
      stats: {
        dep: candidates.filter(c => c.type === 'DEP').length,
        arr: candidates.filter(c => c.type === 'ARR').length,
        total: candidates.length
      }
    };
  }

  global.ForeignScheduleParser = {
    parseForeignScheduleWorkbook,
    matchesSchedule,
    inOperationalWindow,
    datetimeOnScheduleDay,
    normalizeSelectedDate,
    STATION
  };
})(typeof window !== 'undefined' ? window : globalThis);
