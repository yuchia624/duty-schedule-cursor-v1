/**
 * GLOBAL FIS Excel 解析 → flightDefs
 * 需搭配 SheetJS（XLSX）；匯出 parseFisWorkbook(workbook, selectedDateISO)
 */
(function (global) {
  const STATION = 'TPE';
  const ALLOWED_SERVICE = new Set(['A', 'C', 'G', 'J', 'K', 'P', 'Q', 'R', 'X']);
  const HEADER_ALIASES = {
    flt: ['FLT'],
    acNo: ['A/C No', 'A/C NO', 'AC No'],
    dep: ['DEP'],
    arr: ['ARR'],
    std: ['STD'],
    sta: ['STA'],
    etd: ['ETD'],
    eta: ['ETA'],
    depGate: ['Dep Gate', 'DEP Gate', 'DepGate'],
    arrGate: ['Arr Gate', 'ARR Gate', 'ArrGate'],
    pax: ['PAX'],
    status: ['Status'],
    serviceType: ['Service Type', 'ServiceType'],
    eChat: ['eChat', 'EChat']
  };

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function normHeader(v) {
    return clean(v).replace(/\s+/g, '').toLowerCase();
  }

  function parseIsoDate(iso) {
    const m = String(iso || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }

  function isoFromDate(d) {
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    const da = String(d.getDate()).padStart(2, '0');
    return `${y}-${mo}-${da}`;
  }

  function addDays(date, n) {
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    d.setDate(d.getDate() + n);
    return d;
  }

  function parseFisTimeValue(value) {
    if (value == null || value === '') return null;
    if (typeof value === 'number' && value >= 0 && value < 1) {
      const total = Math.round(value * 24 * 60);
      const hh = Math.floor(total / 60) % 24;
      const mm = total % 60;
      return {
        time: `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`,
        daySuffix: '',
        raw: value
      };
    }
    const s = clean(value);
    const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\((\d+)\)$/);
    if (m) {
      return {
        time: `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`,
        daySuffix: m[3],
        raw: s
      };
    }
    const plain = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (plain) {
      return {
        time: `${String(Number(plain[1])).padStart(2, '0')}:${plain[2]}`,
        daySuffix: '',
        raw: s
      };
    }
    return null;
  }

  function timeToMinutes(timeStr) {
    const [hh, mm] = String(timeStr || '').split(':').map(Number);
    if (Number.isNaN(hh) || Number.isNaN(mm)) return null;
    return hh * 60 + mm;
  }

  function buildSuffixDateMap(suffixCounts, selectedDate) {
    const entries = Object.entries(suffixCounts).filter(([, c]) => c > 0);
    if (!entries.length) return new Map();
    entries.sort((a, b) => b[1] - a[1] || Number(a[0]) - Number(b[0]));
    const anchorSuffix = entries[0][0];
    const anchorNum = Number(anchorSuffix);
    const map = new Map();
    entries.forEach(([suffix]) => {
      const delta = Number(suffix) - anchorNum;
      map.set(suffix, isoFromDate(addDays(selectedDate, delta)));
    });
    return map;
  }

  function datetimeFromFisTime(parsed, suffixMap, selectedDate) {
    if (!parsed || !parsed.time) return null;
    let dateIso = isoFromDate(selectedDate);
    if (parsed.daySuffix && suffixMap.has(parsed.daySuffix)) {
      dateIso = suffixMap.get(parsed.daySuffix);
    } else if (parsed.daySuffix) {
      const delta = Number(parsed.daySuffix) - selectedDate.getDate();
      dateIso = isoFromDate(addDays(selectedDate, delta));
    }
    const mins = timeToMinutes(parsed.time);
    if (mins == null) return null;
    const [y, mo, da] = dateIso.split('-').map(Number);
    return {
      dateIso,
      minutes: mins,
      ms: new Date(y, mo - 1, da, Math.floor(mins / 60), mins % 60, 0).getTime()
    };
  }

  /** 出境 STD：D 日 04:30（含）～ D+1 日 04:30（不含）；入境 STA：D 日 03:30（含）～ D+1 日 03:30（不含） */
  function inOperationalWindow(type, dt, selectedDate) {
    if (!dt) return false;
    const y = selectedDate.getFullYear();
    const mo = selectedDate.getMonth();
    const da = selectedDate.getDate();
    const startH = type === 'DEP' ? 4 : 3;
    const startM = 30;
    const startMs = new Date(y, mo, da, startH, startM, 0).getTime();
    const endMs = new Date(y, mo, da + 1, startH, startM, 0).getTime();
    return dt.ms >= startMs && dt.ms < endMs;
  }

  /** PRE 匯入（額外、與入境無關）：美加歐澳線、STA 日期 = 班表日+1、隔日 00:01（含）～ 09:00（不含） */
  const PREFILING_ORIGIN_AIRPORTS = new Set([
    'LAX', 'SFO', 'SEA', 'ORD', 'JFK', 'IAH', 'DFW', 'IAD',
    'YYZ', 'YVR',
    'VIE', 'MUC', 'CDG', 'MXP', 'AMS', 'LHR',
    'BNE'
  ]);
  const PREFILING_IMPORT_STA_START_MIN = 1;
  const PREFILING_IMPORT_STA_END_MIN = 9 * 60;

  function isPrefilingOriginAirport(dep) {
    const origin = clean(dep).toUpperCase();
    return !!origin && PREFILING_ORIGIN_AIRPORTS.has(origin);
  }

  function inPrefilingImportWindow(dt, selectedDate, dep) {
    if (!dt || !isPrefilingOriginAirport(dep)) return false;
    const nextIso = isoFromDate(addDays(selectedDate, 1));
    if (dt.dateIso !== nextIso) return false;
    return dt.minutes >= PREFILING_IMPORT_STA_START_MIN && dt.minutes < PREFILING_IMPORT_STA_END_MIN;
  }

  function matchesPrefilingImportWindow(parsed, suffixMap, selectedDate, dep) {
    const dt = datetimeFromFisTime(parsed, suffixMap, selectedDate);
    if (!dt) return false;
    return inPrefilingImportWindow(dt, selectedDate, dep);
  }

  function matchesSelectedScheduleDay(type, parsed, suffixMap, selectedDate) {
    const dt = datetimeFromFisTime(parsed, suffixMap, selectedDate);
    if (!dt) return false;
    return inOperationalWindow(type, dt, selectedDate);
  }

  function normalizeGate(v) {
    if (v == null || v === '') return '';
    return clean(v);
  }

  function normalizeStatus(v) {
    const s = clean(v);
    if (!s) return '';
    if (/cancel/i.test(s)) return 'CANX';
    return s;
  }

  function normalizeServiceType(v) {
    const s = clean(v).toUpperCase().replace(/Ｘ/g, 'X');
    return s.length === 1 ? s : s.charAt(0);
  }

  function isAllowedService(v) {
    return ALLOWED_SERVICE.has(normalizeServiceType(v));
  }

  function getCellMatrix(ws) {
    if (!ws || !ws['!ref']) return [];
    return global.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  }

  function findHeaderRow(matrix) {
    for (let r = 0; r < Math.min(matrix.length, 40); r++) {
      const row = matrix[r];
      if (!row || !row.length) continue;
      if (clean(row[0]).toUpperCase() === 'FLT') return r;
    }
    return -1;
  }

  function buildColumnIndex(headerRow) {
    const index = {};
    headerRow.forEach((cell, c) => {
      const n = normHeader(cell);
      if (!n) return;
      Object.keys(HEADER_ALIASES).forEach(key => {
        if (index[key] != null) return;
        const hit = HEADER_ALIASES[key].some(alias => normHeader(alias) === n);
        if (hit) index[key] = c;
      });
    });
    return index;
  }

  function readSheetStats(matrix) {
    let depCount = null;
    let arrCount = null;
    for (let r = 0; r < Math.min(matrix.length, 20); r++) {
      const label = clean(matrix[r]?.[0]).toLowerCase();
      const val = matrix[r + 1]?.[0];
      if (label === 'departure' && val !== '' && !Number.isNaN(Number(val))) depCount = Number(val);
      if (label === 'arrival' && val !== '' && !Number.isNaN(Number(val))) arrCount = Number(val);
    }
    return { depCount, arrCount };
  }

  function pickFormattedCell(ws, matrix, r, colIndex, key) {
    const c = colIndex[key];
    if (c == null) return '';
    const raw = matrix[r]?.[c];
    if (raw != null && raw !== '') {
      // FIS 匯出檔常見 cell.w 與儲存格實值錯位（上一列殘留），字串時間以 raw 為準。
      if (typeof raw !== 'number') return String(raw).trim();
    }
    const addr = global.XLSX.utils.encode_cell({ r, c });
    const cell = ws?.[addr];
    if (cell?.w != null && String(cell.w).trim()) return String(cell.w).trim();
    if (raw == null || raw === '') return '';
    return String(raw).trim();
  }

  function rowToObject(ws, matrix, r, colIndex) {
    const row = matrix[r] || [];
    const pick = key => (colIndex[key] != null ? row[colIndex[key]] : '');
    return {
      flt: clean(pick('flt')),
      acNo: clean(pick('acNo')),
      dep: clean(pick('dep')).toUpperCase(),
      arr: clean(pick('arr')).toUpperCase(),
      std: pickFormattedCell(ws, matrix, r, colIndex, 'std'),
      sta: pickFormattedCell(ws, matrix, r, colIndex, 'sta'),
      etd: pickFormattedCell(ws, matrix, r, colIndex, 'etd'),
      eta: pickFormattedCell(ws, matrix, r, colIndex, 'eta'),
      depGate: pick('depGate'),
      arrGate: pick('arrGate'),
      pax: clean(pick('pax')),
      status: pick('status'),
      serviceType: pick('serviceType'),
      eChat: pick('eChat')
    };
  }

  function analyzeSheet(ws) {
    const matrix = getCellMatrix(ws);
    const headerRowIdx = findHeaderRow(matrix);
    if (headerRowIdx < 0) {
      return { error: '找不到 FLT 表頭列', matrix, headerRowIdx: -1 };
    }
    const colIndex = buildColumnIndex(matrix[headerRowIdx]);
    if (colIndex.flt == null) {
      return { error: '表頭缺少 FLT 欄', matrix, headerRowIdx, colIndex };
    }
    let depTpe = 0;
    let arrTpe = 0;
    const rows = [];
    for (let r = headerRowIdx + 1; r < matrix.length; r++) {
      const obj = rowToObject(ws, matrix, r, colIndex);
      if (!obj.flt) continue;
      rows.push(obj);
      if (obj.dep === STATION) depTpe += 1;
      if (obj.arr === STATION) arrTpe += 1;
    }
    return {
      matrix,
      headerRowIdx,
      colIndex,
      rows,
      depTpe,
      arrTpe,
      stats: readSheetStats(matrix),
      missingCols: Object.keys(HEADER_ALIASES).filter(k => colIndex[k] == null)
    };
  }

  function classifySheets(sheetInfos) {
    if (sheetInfos.length < 2) {
      return { error: 'FIS 檔必須包含兩個工作表（出境／入境）。' };
    }
    const ranked = sheetInfos.map((info, idx) => ({ idx, ...info }));
    ranked.sort((a, b) => (b.depTpe - a.depTpe) || (a.arrTpe - b.arrTpe));
    const depCandidate = ranked[0];
    const arrRanked = [...sheetInfos]
      .map((info, idx) => ({ idx, ...info }))
      .sort((a, b) => (b.arrTpe - a.arrTpe) || (a.depTpe - b.depTpe));
    let arrCandidate = arrRanked[0];
    if (arrCandidate.idx === depCandidate.idx && arrRanked.length > 1) {
      arrCandidate = arrRanked[1];
    }
    if (depCandidate.idx === arrCandidate.idx) {
      const byStats = sheetInfos
        .map((info, idx) => ({
          idx,
          score: (info.stats.depCount || 0) - (info.stats.arrCount || 0)
        }))
        .sort((a, b) => b.score - a.score);
      const depIdx = byStats[0]?.idx;
      const arrIdx = byStats.find(x => x.idx !== depIdx)?.idx;
      if (depIdx != null && arrIdx != null) {
        return { depSheet: sheetInfos[depIdx], arrSheet: sheetInfos[arrIdx], depIdx, arrIdx };
      }
      return { error: '無法判斷出境／入境工作表，請確認 FIS 匯出格式。' };
    }
    return {
      depSheet: sheetInfos[depCandidate.idx],
      arrSheet: sheetInfos[arrCandidate.idx],
      depIdx: depCandidate.idx,
      arrIdx: arrCandidate.idx
    };
  }

  function collectSuffixCounts(depRows, arrRows) {
    const counts = {};
    const bump = suffix => {
      if (!suffix) return;
      counts[suffix] = (counts[suffix] || 0) + 1;
    };
    depRows.forEach(row => {
      const p = parseFisTimeValue(row.std);
      if (p?.daySuffix) bump(p.daySuffix);
    });
    arrRows.forEach(row => {
      const p = parseFisTimeValue(row.sta);
      if (p?.daySuffix) bump(p.daySuffix);
    });
    return counts;
  }

  function buildFlightDef(type, row, normalizeFlightNo) {
    const anchor = type === 'DEP' ? parseFisTimeValue(row.std) : parseFisTimeValue(row.sta);
    const time = anchor?.time || '';
    const mins = timeToMinutes(time);
    if (mins == null || !time) return null;
    const etdParsed = parseFisTimeValue(row.etd);
    const flight = normalizeFlightNo(row.flt);
    const def = {
      flight,
      type,
      baseTime: time,
      baseMinutes: mins,
      offset: 0,
      extension: 0,
      flt: row.flt,
      acNo: row.acNo,
      dep: row.dep,
      arr: row.arr,
      std: clean(row.std),
      sta: clean(row.sta),
      etd: clean(row.etd),
      etaRaw: '',
      depGate: normalizeGate(row.depGate),
      arrGate: normalizeGate(row.arrGate),
      pax: row.pax,
      serviceType: normalizeServiceType(row.serviceType),
      status: normalizeStatus(row.status),
      eChat: row.eChat
    };
    if (type === 'DEP') {
      def.changeTime = etdParsed?.time && etdParsed.time !== time ? etdParsed.time : '';
      def.eta = '';
      def.etaRaw = '';
    } else {
      def.eta = '';
      def.etaRaw = '';
      def.etaManual = false;
      def.changeTime = '';
    }
    return def;
  }

  function parseFisWorkbook(workbook, selectedDateIso, opts = {}) {
    const warnings = [];
    const normalizeFlightNo = typeof opts.normalizeFlightNo === 'function'
      ? opts.normalizeFlightNo
      : (v) => clean(v).toUpperCase();

    if (!workbook || !Array.isArray(workbook.SheetNames) || !workbook.SheetNames.length) {
      return { error: '無法讀取 Excel 工作表。', flightDefs: [], warnings, stats: { dep: 0, arr: 0, total: 0 } };
    }
    if (workbook.SheetNames.length < 2) {
      return { error: 'FIS 檔必須包含兩個工作表（出境／入境）。', flightDefs: [], warnings, stats: { dep: 0, arr: 0, total: 0 } };
    }

    const selectedDate = parseIsoDate(selectedDateIso);
    if (!selectedDate) {
      return { error: '系統排班日期無效，無法匯入 FIS。', flightDefs: [], warnings, stats: { dep: 0, arr: 0, total: 0 } };
    }

    const sheetInfos = workbook.SheetNames.map(name => {
      const ws = workbook.Sheets[name];
      const info = analyzeSheet(ws);
      return { name, ...info };
    });

    const headerErrors = sheetInfos.filter(s => s.error);
    if (headerErrors.length === sheetInfos.length) {
      return {
        error: '找不到 FLT 表頭列，請確認為 GLOBAL FIS 匯出格式。',
        flightDefs: [],
        warnings,
        stats: { dep: 0, arr: 0, total: 0 }
      };
    }
    if (headerErrors.length) {
      warnings.push(`部分工作表缺少標準 FLT 表頭：${headerErrors.map(s => s.name).join('、')}`);
    }

    const validSheets = sheetInfos.filter(s => !s.error && s.rows.length);
    const classified = classifySheets(validSheets.length >= 2 ? validSheets : sheetInfos.filter(s => !s.error));
    if (classified.error) {
      return { error: classified.error, flightDefs: [], warnings, stats: { dep: 0, arr: 0, total: 0 } };
    }

    const { depSheet, arrSheet } = classified;
    [depSheet, arrSheet].forEach(sheet => {
      if (sheet.missingCols?.length) {
        warnings.push(`工作表「${sheet.name}」缺少欄位：${sheet.missingCols.join('、')}（與標準 FIS 格式不同，部分資料可能空白）`);
      }
    });

    const depCandidates = depSheet.rows.filter(row => row.dep === STATION && isAllowedService(row.serviceType));
    const arrCandidates = arrSheet.rows.filter(row => row.arr === STATION && isAllowedService(row.serviceType));

    const suffixMap = buildSuffixDateMap(
      collectSuffixCounts(depCandidates, arrCandidates),
      selectedDate
    );

    const flightDefs = [];
    let prefilingArrAdded = 0;

    depCandidates.forEach(row => {
      const parsed = parseFisTimeValue(row.std);
      if (!matchesSelectedScheduleDay('DEP', parsed, suffixMap, selectedDate)) return;
      const def = buildFlightDef('DEP', row, normalizeFlightNo);
      if (def) flightDefs.push(def);
    });

    arrCandidates.forEach(row => {
      const parsed = parseFisTimeValue(row.sta);
      if (!matchesSelectedScheduleDay('ARR', parsed, suffixMap, selectedDate)) return;
      if (matchesPrefilingImportWindow(parsed, suffixMap, selectedDate, row.dep)) return;
      const def = buildFlightDef('ARR', row, normalizeFlightNo);
      if (!def) return;
      const dt = datetimeFromFisTime(parsed, suffixMap, selectedDate);
      if (dt?.dateIso) def.scheduleStaDateIso = dt.dateIso;
      def.isPrefilingImport = false;
      def.isInboundImport = true;
      flightDefs.push(def);
    });

    arrCandidates.forEach(row => {
      const parsed = parseFisTimeValue(row.sta);
      if (!matchesPrefilingImportWindow(parsed, suffixMap, selectedDate, row.dep)) return;
      const def = buildFlightDef('ARR', row, normalizeFlightNo);
      if (!def) return;
      const dt = datetimeFromFisTime(parsed, suffixMap, selectedDate);
      if (dt?.dateIso) def.scheduleStaDateIso = dt.dateIso;
      def.isPrefilingImport = true;
      def.isInboundImport = false;
      flightDefs.push(def);
      prefilingArrAdded += 1;
    });

    if (prefilingArrAdded) {
      warnings.push(`PRE 額外匯入 ${prefilingArrAdded} 筆隔日美加歐澳航班（STA 日期為班表日+1，00:01～09:00）`);
    }

    flightDefs.sort((a, b) => {
      if (a.type !== b.type) return a.type === 'DEP' ? -1 : 1;
      if (a.type === 'ARR') {
        const ae = timeToMinutes(a.eta) ?? a.baseMinutes;
        const be = timeToMinutes(b.eta) ?? b.baseMinutes;
        return ae - be || a.flight.localeCompare(b.flight);
      }
      return a.baseMinutes - b.baseMinutes || a.flight.localeCompare(b.flight);
    });

    const dep = flightDefs.filter(f => f.type === 'DEP').length;
    const arr = flightDefs.filter(f => f.type === 'ARR').length;

    return {
      flightDefs,
      warnings,
      stats: { dep, arr, total: flightDefs.length },
      meta: {
        depSheet: depSheet.name,
        arrSheet: arrSheet.name,
        selectedDate: selectedDateIso
      }
    };
  }

  global.FisFlightParser = {
    parseFisWorkbook,
    parseFisTimeValue,
    STATION,
    ALLOWED_SERVICE
  };
})(typeof window !== 'undefined' ? window : globalThis);
