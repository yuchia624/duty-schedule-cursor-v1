/**
 * 長榮立榮夏季班表 Excel → 航點清單
 * 欄位：DEP(C/J) ARR(D/K)，TPE 出境目的地 + TPE 入境起點
 */
(function (global) {
  const PARSER_VERSION = '2026-summer-homeline-v1';

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function normalizeIata(code) {
    const s = clean(code).toUpperCase().replace(/[^A-Z0-9]/g, '');
    return s.length === 3 ? s : '';
  }

  function sheetToMatrix(sheet) {
    if (!sheet || !sheet['!ref']) return [];
    if (typeof XLSX === 'undefined' || typeof XLSX.utils?.sheet_to_json !== 'function') return [];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  }

  function extractAirportsFromMatrix(matrix) {
    const set = new Set();
    (matrix || []).slice(1).forEach(row => {
      if (!Array.isArray(row)) return;
      const dep = normalizeIata(row[2]);
      const arr = normalizeIata(row[3]);
      const inboundDep = normalizeIata(row[9]);
      const inboundArr = normalizeIata(row[10]);
      if (dep === 'TPE' && arr) set.add(arr);
      if (inboundArr === 'TPE' && inboundDep) set.add(inboundDep);
    });
    return [...set].sort();
  }

  function parseHomelineSummerWorkbook(workbook) {
    if (!workbook || !workbook.SheetNames?.length) {
      return { ok: false, error: '找不到工作表。', airports: [] };
    }
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = sheetToMatrix(sheet);
    if (!matrix.length) {
      return { ok: false, error: '檔案內容為空。', airports: [] };
    }
    const airports = extractAirportsFromMatrix(matrix);
    if (!airports.length) {
      return { ok: false, error: '找不到 TPE 出境／入境航點（需 DEP/ARR 欄）。', airports: [] };
    }
    return { ok: true, airports, rowCount: matrix.length - 1 };
  }

  global.DestClassParser = {
    PARSER_VERSION,
    parseHomelineSummerWorkbook,
    extractAirportsFromMatrix
  };
})(typeof window !== 'undefined' ? window : globalThis);
