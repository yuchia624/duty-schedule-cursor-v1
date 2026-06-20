/**
 * EDW 本家訂位報表解析（Tab 分隔偽 xls / 純文字 TSV）
 * 匯出 HomelinePaxParser.parseText / parseArrayBuffer
 */
(function (global) {
  const PARSER_VERSION = '2026-06-20-v4';
  const HOMELINE_CARRIERS = new Set(['BR', 'B7']);
  const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

  /** TPE EDW 劃位人數預報表固定欄位（0-based） */
  const EDW_COL = {
    date: 0,
    std: 1,
    carrier: 2,
    fltNo: 3,
    dest: 4,
    capacity: 5,
    pax: 6,
    group: 9,
    transit: 11,
    wchc: 13,
    wchr: 14,
    wchs: 15
  };

  function clean(v) {
    return String(v ?? '').replace(/\u3000/g, '').trim();
  }

  function isValidIsoDate(s) {
    return ISO_DATE_RE.test(clean(s));
  }

  function normalizeLookupFlight(flight) {
    const s = clean(flight).toUpperCase();
    if (!s) return s;
    if (s.startsWith('B7')) {
      const num = s.slice(2);
      if (/^\d+$/.test(num)) return `B7${num.padStart(4, '0')}`;
      return s;
    }
    if (s.startsWith('BR')) {
      const num = s.slice(2);
      if (/^\d+$/.test(num)) return `BR${num.padStart(4, '0')}`;
      return s;
    }
    const m = s.match(/^([A-Z]{2})(\d+)$/);
    if (m) return `${m[1]}${m[2].padStart(4, '0')}`;
    return s;
  }

  function normalizeFlightNo(carrier, fltNo) {
    const c = clean(carrier).toUpperCase();
    const digits = clean(fltNo).replace(/^0+/, '') || '0';
    if (!c || !/^\d+$/.test(digits)) return '';
    return normalizeLookupFlight(`${c}${digits}`);
  }

  function normalizeTime(std) {
    const s = clean(std);
    const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!m) return s;
    return `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`;
  }

  function parseIntField(v) {
    const n = parseInt(clean(v).replace(/,/g, ''), 10);
    return Number.isFinite(n) ? n : null;
  }

  function computeLoadFactor(pax, capacity) {
    if (!Number.isFinite(pax) || !Number.isFinite(capacity) || capacity <= 0) return null;
    return Math.round((pax / capacity) * 1000) / 1000;
  }

  function capPaxToCapacity(pax, capacity) {
    if (!Number.isFinite(pax) || !Number.isFinite(capacity)) return pax;
    return Math.min(pax, capacity);
  }

  function sumOptionalInts(values) {
    let sum = 0;
    let hasAny = false;
    (values || []).forEach(v => {
      if (!Number.isFinite(v)) return;
      hasAny = true;
      sum += v;
    });
    return hasAny ? sum : null;
  }

  function computeDirectPax(pax, transit, group) {
    if (!Number.isFinite(pax)) return null;
    const t = Number.isFinite(transit) ? transit : 0;
    const g = Number.isFinite(group) ? group : 0;
    return Math.max(0, pax - t - g);
  }

  function computeWheelchair(wchc, wchr, wchs) {
    return sumOptionalInts([wchc, wchr, wchs]);
  }

  function detectDataStartLine(lines) {
    for (let i = 0; i < Math.min(lines.length, 8); i++) {
      const parts = String(lines[i] || '').split('\t');
      if (parts.some(p => clean(p).toUpperCase() === 'CARRIER')) return i + 1;
    }
    return 2;
  }

  function pickField(parts, index) {
    if (!Number.isInteger(index) || index < 0 || index >= parts.length) return null;
    return parseIntField(parts[index]);
  }

  function decodeArrayBuffer(buffer) {
    if (!buffer) return '';
    const bytes = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : buffer;
    if (bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b) {
      if (typeof XLSX === 'undefined') {
        throw new Error('此檔案為 Excel 二進位格式，需要 SheetJS（XLSX）才能解析。');
      }
      const wb = XLSX.read(buffer, { type: 'array' });
      const sheetName = wb.SheetNames[0];
      if (!sheetName) return '';
      const sheet = wb.Sheets[sheetName];
      return XLSX.utils.sheet_to_csv(sheet, { FS: '\t' });
    }
    const decoderLabels = ['utf-8', 'big5', 'windows-950', 'x-big5'];
    for (const label of decoderLabels) {
      try {
        const text = new TextDecoder(label, { fatal: false }).decode(bytes).replace(/^\uFEFF/, '');
        if (!clean(text)) continue;
        if (text.includes('Carrier') || text.includes('CARRIER') || /^\d{4}-\d{2}-\d{2}\t/m.test(text)) {
          return text;
        }
      } catch (_) { /* try next */ }
    }
    return new TextDecoder('utf-8', { fatal: false }).decode(bytes).replace(/^\uFEFF/, '');
  }

  function parseText(text, opts = {}) {
    const warnings = [];
    const rows = [];
    const carriers = new Set();
    const dates = new Set();
    let skipped = 0;
    let duplicateKeys = 0;
    const seenByDateFlight = new Map();

    const lines = String(text || '').replace(/^\uFEFF/, '').split(/\r?\n/);
    const dataStart = detectDataStartLine(lines);

    lines.forEach((line, idx) => {
      if (!clean(line)) return;
      if (idx < dataStart) return;
      const parts = line.split('\t');
      if (parts.length < 12) {
        skipped += 1;
        return;
      }
      const dateIso = clean(parts[EDW_COL.date]);
      const std = normalizeTime(parts[EDW_COL.std]);
      const carrier = clean(parts[EDW_COL.carrier]).toUpperCase();
      const fltNo = clean(parts[EDW_COL.fltNo]);
      const dest = clean(parts[EDW_COL.dest]).toUpperCase();
      const capacity = pickField(parts, EDW_COL.capacity);
      const rawPax = pickField(parts, EDW_COL.pax);
      const group = pickField(parts, EDW_COL.group);
      const transit = pickField(parts, EDW_COL.transit);
      const wchc = pickField(parts, EDW_COL.wchc);
      const wchr = pickField(parts, EDW_COL.wchr);
      const wchs = pickField(parts, EDW_COL.wchs);
      const wheelchair = computeWheelchair(wchc, wchr, wchs);

      if (!isValidIsoDate(dateIso)) {
        skipped += 1;
        return;
      }
      if (!/^\d+$/.test(fltNo)) {
        skipped += 1;
        return;
      }
      if (!HOMELINE_CARRIERS.has(carrier)) {
        skipped += 1;
        return;
      }
      if (!Number.isFinite(capacity) || !Number.isFinite(rawPax)) {
        warnings.push(`第 ${idx + 1} 行 ${carrier}${fltNo}：艙容或訂位數無效，已略過`);
        skipped += 1;
        return;
      }

      const pax = capPaxToCapacity(rawPax, capacity);
      if (rawPax > capacity) {
        warnings.push(`第 ${idx + 1} 行 ${carrier}${fltNo}：訂位數 ${rawPax} 超過艙容 ${capacity}，已以艙容計`);
      }
      const directPax = computeDirectPax(pax, transit, group);

      const flight = normalizeFlightNo(carrier, fltNo);
      if (!flight) {
        skipped += 1;
        return;
      }

      const key = `${dateIso}|${flight}`;
      if (seenByDateFlight.has(key)) duplicateKeys += 1;
      seenByDateFlight.set(key, true);

      carriers.add(carrier);
      dates.add(dateIso);
      rows.push({
        dateIso,
        std,
        carrier,
        fltNo,
        flight,
        dest,
        capacity,
        pax,
        group: Number.isFinite(group) ? group : null,
        transit: Number.isFinite(transit) ? transit : null,
        directPax,
        wchc: Number.isFinite(wchc) ? wchc : null,
        wchr: Number.isFinite(wchr) ? wchr : null,
        wchs: Number.isFinite(wchs) ? wchs : null,
        wheelchair,
        loadFactor: computeLoadFactor(pax, capacity)
      });
    });

    if (!rows.length) {
      return {
        ok: false,
        error: '找不到有效的 BR/B7 訂位列。請確認檔案格式與 EDW 報表一致。',
        rows: [],
        warnings,
        stats: { rows: 0, dates: 0, skipped, duplicateKeys, carriers: [] }
      };
    }

    const sortedDates = [...dates].sort();
    return {
      ok: true,
      rows,
      warnings,
      stats: {
        rows: rows.length,
        dates: dates.size,
        dateMin: sortedDates[0],
        dateMax: sortedDates[sortedDates.length - 1],
        skipped,
        duplicateKeys,
        carriers: [...carriers].sort()
      }
    };
  }

  function parseArrayBuffer(buffer, opts = {}) {
    try {
      const text = decodeArrayBuffer(buffer);
      if (!clean(text)) {
        return {
          ok: false,
          error: '檔案內容為空或無法讀取。',
          rows: [],
          warnings: [],
          stats: { rows: 0, dates: 0, skipped: 0, duplicateKeys: 0, carriers: [] }
        };
      }
      return parseText(text, opts);
    } catch (err) {
      return {
        ok: false,
        error: err?.message || String(err),
        rows: [],
        warnings: [],
        stats: { rows: 0, dates: 0, skipped: 0, duplicateKeys: 0, carriers: [] }
      };
    }
  }

  global.HomelinePaxParser = {
    PARSER_VERSION,
    EDW_COL,
    HOMELINE_CARRIERS,
    clean,
    isValidIsoDate,
    normalizeLookupFlight,
    normalizeFlightNo,
    computeLoadFactor,
    capPaxToCapacity,
    computeDirectPax,
    computeWheelchair,
    parseText,
    parseArrayBuffer
  };
})(typeof window !== 'undefined' ? window : globalThis);
