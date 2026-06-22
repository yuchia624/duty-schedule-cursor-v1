/**
 * 點名表 Excel 解析（點名表整理匯出工具 v5.2 邏輯）
 * 輸出：代號、姓名、班別；可選隔日班別比對
 */
(function (global) {
  const SKIP_LABELS = new Set([
    '時間', '人數', '人員', '人員清單', '班別', '班別時間', '班別明細',
    '上班人員清單', '加班類型', '加班及休假明細', '課程/會議明細', '課程／會議明細',
    '類型', '未加班', '休假人員', '差假', '點名表', '點名表(Roll Call)',
  ]);

  function cleanName(v) { return String(v ?? '').replace(/\u3000/g, '').trim(); }
  function normName(v) { return cleanName(v).replace(/\s/g, ''); }

  function normalizeTime(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number' && v >= 0 && v < 1) {
      const total = Math.round(v * 24 * 60);
      return `${String(Math.floor(total / 60)).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}`;
    }
    if (v instanceof Date && !isNaN(v)) {
      return `${String(v.getHours()).padStart(2, '0')}:${String(v.getMinutes()).padStart(2, '0')}`;
    }
    const m = cleanName(v).match(/^(\d{1,2}):(\d{2})/);
    if (m) return `${String(+m[1]).padStart(2, '0')}:${m[2]}`;
    return cleanName(v);
  }

  function isValidTime(t) { return /^\d{2}:\d{2}$/.test(t); }

  function isShiftCode(code) {
    const c = cleanName(code);
    if (!c || c.length > 4 || SKIP_LABELS.has(c)) return false;
    return /^[A-Za-z0-9]+$/.test(c);
  }

  function safeSheetName(name) {
    return name.replace(/[\\/?*\[\]:]/g, '-').slice(0, 31) || '點名表日期';
  }

  function getCell(ws, r, c) {
    const cell = ws[global.XLSX.utils.encode_cell({ r, c })];
    if (!cell) return null;
    return cell.w != null ? cell.w : cell.v;
  }

  function findSectionRow(ws, range, label) {
    for (let r = range.s.r; r <= range.e.r; r++) {
      if (normName(getCell(ws, r, 1)) === label) return r;
    }
    return -1;
  }

  function parseDateFromSheet(ws, range) {
    let raw = '';
    for (let r = range.s.r; r <= Math.min(range.s.r + 7, range.e.r); r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = getCell(ws, r, c);
        if (v != null && String(v).includes('日期:')) raw = String(v);
      }
    }
    const m = raw.match(/日期\s*:\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\((.)\)/);
    if (!m) return { display: raw || '未找到日期', sheet: '點名表日期' };
    const [, mm, dd, yyyy, dow] = m;
    const sheet = `${yyyy}-${String(+mm).padStart(2, '0')}-${String(+dd).padStart(2, '0')}(${dow})`;
    return {
      display: `${yyyy}/${String(+mm).padStart(2, '0')}/${String(+dd).padStart(2, '0')}(${dow})`,
      sheet: safeSheetName(sheet),
    };
  }

  function parseRosterFromSheet(ws, range, endRow) {
    const roster = [];
    let currentTime = null;
    const last = endRow >= 0 ? endRow : range.e.r;
    const nameCols = [];
    for (let c = 9; c <= Math.min(range.e.c, 60); c += 3) nameCols.push(c);

    let startR = 4;
    for (let r = range.s.r; r <= Math.min(range.e.r, 30); r++) {
      if (normName(getCell(ws, r, 1)) === '時間') { startR = r + 1; break; }
    }

    for (let r = startR; r < last; r++) {
      const label = normName(getCell(ws, r, 1));
      if (label === '班別明細' || label === '班別') break;
      const rowTime = normalizeTime(getCell(ws, r, 1));
      if (isValidTime(rowTime)) currentTime = rowTime;
      if (!currentTime) continue;
      for (const c of nameCols) {
        const raw = cleanName(getCell(ws, r, c));
        if (!raw || raw === ' ') continue;
        const name = normName(raw);
        if (SKIP_LABELS.has(name)) continue;
        if (!/[\u4e00-\u9fff]/.test(name) || name.length < 2) continue;
        roster.push({ time: currentTime, name });
      }
    }
    return roster;
  }

  function parseShiftDetailFromSheet(ws, range, startRow) {
    const shifts = [];
    if (startRow < 0) return shifts;
    for (let r = startRow + 1; r <= range.e.r; r++) {
      const code = cleanName(getCell(ws, r, 1));
      const timeRange = cleanName(getCell(ws, r, 13));
      if (!code || !timeRange.includes('~')) {
        if (shifts.length && !code && !cleanName(getCell(ws, r, 5))) break;
        continue;
      }
      if (!isShiftCode(code)) continue;
      const people = cleanName(getCell(ws, r, 5));
      const tokens = Array.from(people);
      shifts.push({ code, start: normalizeTime(timeRange.split('~')[0].trim()), tokens });
    }
    return shifts;
  }

  function normalizeChar(ch) {
    try { return ch.normalize('NFKC'); } catch { return ch; }
  }

  function tokenMatchesName(token, name) {
    const t = normName(token);
    if (!t) return false;
    const n = normName(name);
    if (n.includes(t)) return true;
    for (const ch of n) {
      if (ch === t || normalizeChar(ch) === t) return true;
    }
    return false;
  }

  function scoreNameForToken(name, token) {
    if (!tokenMatchesName(token, name)) return -1;
    const n = normName(name);
    let score = 0;
    if (token.length === 1) {
      if (n.endsWith(token)) score += 25;
      else if (n.startsWith(token)) score += 5;
      else score += 12;
    } else {
      if (n.startsWith(token)) score += 20;
      if (n.endsWith(token)) score += 15;
    }
    score += (token.length / n.length) * 10;
    if ([...n].filter((c) => c === token).length === 1) score += 10;
    return score;
  }

  function buildRowsFromShifts(shifts, roster) {
    const byTime = new Map();
    const allPeople = [];

    for (const p of roster) {
      const t = normalizeTime(p.time);
      if (!isValidTime(t)) continue;
      const person = { name: p.name, time: t, assigned: false };
      allPeople.push(person);
      if (!byTime.has(t)) byTime.set(t, []);
      byTime.get(t).push(person);
    }

    function pickPerson(pool, token) {
      let cands = pool.filter((p) => !p.assigned && scoreNameForToken(p.name, token) >= 0);
      if (!cands.length) {
        cands = allPeople.filter((p) => !p.assigned && scoreNameForToken(p.name, token) >= 0);
      }
      if (!cands.length) return '';
      cands.sort((a, b) => {
        const ds = scoreNameForToken(b.name, token) - scoreNameForToken(a.name, token);
        if (ds !== 0) return ds;
        return pool.indexOf(a) - pool.indexOf(b);
      });
      cands[0].assigned = true;
      return cands[0].name;
    }

    const result = [];
    for (const shift of shifts) {
      const pool = byTime.get(normalizeTime(shift.start)) || [];
      const unassigned = pool.filter((p) => !p.assigned);

      if (shift.tokens.length > 0 && shift.tokens.length === unassigned.length) {
        for (let i = 0; i < shift.tokens.length; i++) {
          unassigned[i].assigned = true;
          result.push({ 代號: shift.tokens[i], 姓名: unassigned[i].name, 班別: shift.code });
        }
        continue;
      }

      for (const token of shift.tokens) {
        result.push({ 代號: token, 姓名: pickPerson(pool, token), 班別: shift.code });
      }
    }
    return result;
  }

  function getMainSheet(workbook) {
    const sheetName = workbook.SheetNames.find((n) => String(n).includes('工作')) || workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const range = global.XLSX.utils.decode_range(ws['!ref'] || 'A1:AZ200');
    return { ws, range, sheetName };
  }

  function parseWorkbook(workbook) {
    const { ws, range } = getMainSheet(workbook);
    const dateInfo = parseDateFromSheet(ws, range);
    const shiftSectionRow = findSectionRow(ws, range, '班別明細');
    const roster = parseRosterFromSheet(ws, range, shiftSectionRow >= 0 ? shiftSectionRow : -1);
    const shifts = parseShiftDetailFromSheet(ws, range, shiftSectionRow);
    const result = buildRowsFromShifts(shifts, roster);
    return { dateInfo, result, ws, range };
  }

  function parseLeaveFromSheet(ws, range) {
    const leaveTokens = new Set();
    const addTokensFromText = (txt) => {
      const raw = cleanName(txt);
      if (!raw) return;
      // 先嘗試以常見分隔符切，再退回逐字（符合原點名表代號字）
      const split = raw.split(/[、,，\s/]+/).filter(Boolean);
      if (split.length > 1) {
        split.forEach((t) => leaveTokens.add(normName(t)));
      } else {
        Array.from(raw).forEach((t) => leaveTokens.add(normName(t)));
      }
    };
    for (let r = range.s.r; r <= range.e.r; r++) {
      const label = normName(getCell(ws, r, 19));
      if (label === '休假人員' || label === '差假') {
        const people = cleanName(getCell(ws, r, 23));
        const count = parseInt(String(getCell(ws, r, 31) ?? '').trim(), 10);
        const tokens = Number.isFinite(count) && count > 0 ? Array.from(people).slice(0, count) : Array.from(people);
        tokens.forEach((t) => leaveTokens.add(normName(t)));
      }
      // 兼容欄位位移：只要該列有「休假人員/差假」，就掃描右側所有儲存格
      const rowTexts = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = getCell(ws, r, c);
        if (v != null) rowTexts.push(cleanName(v));
      }
      const rowJoined = rowTexts.join(' ');
      if (/休假人員|差假/.test(rowJoined)) {
        rowTexts.forEach((t) => {
          if (!t || t === '休假人員' || t === '差假') return;
          // 避免把日期/時間/純數字當作代號
          if (/^\d+$/.test(t) || /^\d{1,2}:\d{2}$/.test(t)) return;
          addTokensFromText(t);
        });
      }
    }
    return leaveTokens;
  }

  function toPersonMap(rows) {
    const map = new Map();
    for (const row of rows) {
      if (!row.姓名) continue;
      if (!map.has(row.姓名)) map.set(row.姓名, { 代號: row.代號, 班別: row.班別 });
    }
    return map;
  }

  function isOnLeave(row, leaveTokens) {
    if (!leaveTokens.size) return false;
    if (row.代號 && leaveTokens.has(row.代號)) return true;
    if (!row.姓名) return false;
    for (const token of leaveTokens) {
      if (tokenMatchesName(token, row.姓名)) return true;
    }
    return false;
  }

  function findNextShift(row, nextRows, nextMap, leaveTokens) {
    // 休假人員／差假優先：班別明細代號字常會誤配到其他姓名（如「嬛」→ hs）
    if (isOnLeave(row, leaveTokens)) return '休';
    if (row.姓名 && nextMap.has(row.姓名)) return nextMap.get(row.姓名).班別;
    if (row.代號) {
      const byCode = nextRows.filter((n) => n.代號 === row.代號 && n.班別);
      if (byCode.length === 1) return byCode[0].班別;
      if (row.姓名) {
        const matched = byCode.filter(
          (n) => normName(n.姓名) === normName(row.姓名) || tokenMatchesName(row.代號, n.姓名),
        );
        if (matched.length === 1) return matched[0].班別;
      }
    }
    // 你的流程定義：隔日沒有班別時，視為休假
    return '休';
  }

  function buildNextDayRows(todayRows, nextWorkbook) {
    const { result: nextRows, ws, range } = parseWorkbook(nextWorkbook);
    const leaveTokens = parseLeaveFromSheet(ws, range);
    const nextMap = toPersonMap(nextRows);
    return todayRows.map((row) => ({
      ...row,
      隔日班別: findNextShift(row, nextRows, nextMap, leaveTokens),
    }));
  }

  function extractLeaveTokens(nextWorkbook) {
    const { ws, range } = getMainSheet(nextWorkbook);
    return Array.from(parseLeaveFromSheet(ws, range));
  }

  function tryParseWorkbook(workbook) {
    if (!workbook || !global.XLSX) return null;
    try {
      const parsed = parseWorkbook(workbook);
      if (!parsed.result.length) return null;
      return parsed;
    } catch {
      return null;
    }
  }

  function looksLikeRollCallWorkbook(workbook) {
    if (!workbook || !workbook.SheetNames?.length) return false;
    const parsed = tryParseWorkbook(workbook);
    return !!(parsed && parsed.result.length >= 3);
  }

  global.RollCallParser = {
    parseWorkbook,
    buildNextDayRows,
    extractLeaveTokens,
    tryParseWorkbook,
    looksLikeRollCallWorkbook,
    tokenMatchesName,
    normName,
  };
})(typeof window !== 'undefined' ? window : global);
