/**
 * AirlineData.xls 第一頁（AirlineData）→ 登機門更新
 * 入境機坪 → arrGate、出境機坪 → depGate
 */
(function (global) {
  const SHEET_DATA = 'AirlineData';
  const SHEET_MAP = '工作表1';

  const DEFAULT_AIRLINE_PREFIX = {
    AAR: 'OZ',
    ABG: 'RL',
    ABL: 'BX',
    ABW: 'RU',
    ACA: 'AC',
    ANA: 'NH',
    ANZ: 'NZ',
    CAO: '國貨',
    CQH: '9C',
    CRK: 'HX',
    DKH: 'HO',
    ESR: 'ZE',
    EVA: 'BR',
    HKC: '港貨',
    NCT: 'XW',
    SBI: 'S7',
    SFJ: '7G',
    SLK: 'MI',
    SNJ: '6J',
    TGW: 'TR',
    THA: 'TG',
    THD: 'WE',
    UIA: 'B7',
    VJC: 'VJ'
  };

  const HEADER_ALIASES = {
    airline: ['航空公司'],
    arrFlt: ['入境班次'],
    depFlt: ['出境班次'],
    arrGate: ['入境機坪'],
    depGate: ['出境機坪']
  };

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function normalizeGate(v) {
    if (v == null || v === '') return '';
    const n = Number(v);
    if (Number.isFinite(n) && String(v).trim() !== '' && !String(v).includes(':')) {
      return clean(String(Math.trunc(n)));
    }
    return clean(v);
  }

  function getCellMatrix(ws) {
    if (!ws || !global.XLSX) return [];
    return global.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  }

  function findHeaderRow(matrix) {
    for (let r = 0; r < Math.min(matrix.length, 20); r++) {
      const row = matrix[r] || [];
      const texts = row.map(clean);
      if (texts.includes('航空公司') && (texts.includes('入境機坪') || texts.includes('出境機坪'))) {
        return r;
      }
    }
    return -1;
  }

  function headerIndexMap(headerRow) {
    const map = new Map();
    (headerRow || []).forEach((cell, idx) => {
      const key = clean(cell);
      if (key) map.set(key, idx);
    });
    const out = {};
    Object.keys(HEADER_ALIASES).forEach(field => {
      const aliases = HEADER_ALIASES[field];
      const hit = aliases.find(a => map.has(a));
      out[field] = hit != null ? map.get(hit) : -1;
    });
    return out;
  }

  function readAirlinePrefixMap(workbook) {
    const map = { ...DEFAULT_AIRLINE_PREFIX };
    if (!workbook?.SheetNames?.includes(SHEET_MAP)) return map;
    const ws = workbook.Sheets[SHEET_MAP];
    const matrix = getCellMatrix(ws);
    matrix.forEach(row => {
      const code = clean(row[0]);
      const prefix = clean(row[1]);
      if (code && prefix) map[code.toUpperCase()] = prefix;
    });
    return map;
  }

  function buildFlightNo(prefix, flightNum, normalizeFlightNo) {
    const p = clean(prefix);
    const n = clean(flightNum);
    if (!n) return '';
    const norm = typeof normalizeFlightNo === 'function'
      ? normalizeFlightNo
      : (v) => clean(v).toUpperCase();
    if (/^[A-Z]{2}/i.test(n) || /^[A-Z]\d/i.test(n) || n.startsWith('國') || n.startsWith('港')) {
      return norm(n);
    }
    if (!p) return norm(n);
    return norm(p + n);
  }

  function parseAirlineGateWorkbook(workbook, opts = {}) {
    const normalizeFlightNo = opts.normalizeFlightNo;
    if (!workbook || !Array.isArray(workbook.SheetNames) || !workbook.SheetNames.length) {
      return { error: '無法讀取 Excel 工作表。', updates: [], stats: {} };
    }
    const sheetName = workbook.SheetNames.includes(SHEET_DATA)
      ? SHEET_DATA
      : workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const matrix = getCellMatrix(ws);
    const headerRowIdx = findHeaderRow(matrix);
    if (headerRowIdx < 0) {
      return { error: '找不到 AirlineData 表頭（需含「航空公司」「入境機坪」或「出境機坪」）。', updates: [], stats: {} };
    }
    const cols = headerIndexMap(matrix[headerRowIdx]);
    if (cols.arrGate < 0 && cols.depGate < 0) {
      return { error: '找不到「入境機坪」或「出境機坪」欄位。', updates: [], stats: {} };
    }

    const prefixMap = readAirlinePrefixMap(workbook);
    const updates = [];
    let arrRows = 0;
    let depRows = 0;

    for (let r = headerRowIdx + 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const airline = clean(row[cols.airline]).toUpperCase();
      if (!airline) continue;
      const prefix = prefixMap[airline] || airline;

      const arrFlt = cols.arrFlt >= 0 ? clean(row[cols.arrFlt]) : '';
      const depFlt = cols.depFlt >= 0 ? clean(row[cols.depFlt]) : '';
      const arrGate = cols.arrGate >= 0 ? normalizeGate(row[cols.arrGate]) : '';
      const depGate = cols.depGate >= 0 ? normalizeGate(row[cols.depGate]) : '';

      if (arrFlt && arrGate) {
        const flight = buildFlightNo(prefix, arrFlt, normalizeFlightNo);
        if (flight) {
          updates.push({ flight, type: 'ARR', gate: arrGate, airline });
          arrRows++;
        }
      }
      if (depFlt && depGate) {
        const flight = buildFlightNo(prefix, depFlt, normalizeFlightNo);
        if (flight) {
          updates.push({ flight, type: 'DEP', gate: depGate, airline });
          depRows++;
        }
      }
    }

    if (!updates.length) {
      return { error: '檔案中沒有可用的登機門資料。', updates: [], stats: { arrRows: 0, depRows: 0, total: 0 } };
    }

    return {
      updates,
      stats: { arrRows, depRows, total: updates.length },
      sheetName
    };
  }

  function applyGateUpdates(flightDefs, updates, opts = {}) {
    const normalizeFlightNo = typeof opts.normalizeFlightNo === 'function'
      ? opts.normalizeFlightNo
      : (v) => clean(v).toUpperCase();
    const index = new Map();
    (flightDefs || []).forEach(f => {
      if (!f?.flight || !f?.type) return;
      index.set(`${normalizeFlightNo(f.flight)}|${f.type}`, f);
    });

    const applied = { arr: 0, dep: 0 };
    const unmatched = [];
    const seen = new Set();

    (updates || []).forEach(u => {
      const key = `${normalizeFlightNo(u.flight)}|${u.type}`;
      const target = index.get(key);
      if (!target) {
        if (!seen.has(key)) {
          unmatched.push({ flight: u.flight, type: u.type });
          seen.add(key);
        }
        return;
      }
      if (u.type === 'ARR') {
        target.arrGate = u.gate;
        applied.arr++;
      } else if (u.type === 'DEP') {
        target.depGate = u.gate;
        applied.dep++;
      }
    });

    return {
      applied,
      unmatched,
      totalApplied: applied.arr + applied.dep
    };
  }

  global.AirlineDataGateParser = {
    parseAirlineGateWorkbook,
    applyGateUpdates,
    DEFAULT_AIRLINE_PREFIX
  };
})(typeof window !== 'undefined' ? window : globalThis);
