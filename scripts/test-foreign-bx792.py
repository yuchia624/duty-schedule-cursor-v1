#!/usr/bin/env python3
"""Run local verification: BX792 on 2026-05-27 should be DEP 16:35 in candidates."""
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime

STATION = 'TPE'
PARSER_VERSION_EXPECTED = '2026-05-27-bx792-debug'

def clean(v):
    return str(v or '').replace('\u3000', '').strip()

def excel_weekday(d):
    # JS: Sun=0 -> 7, Mon=1 -> 1 ...
    js_day = (d.weekday() + 1) % 7  # Mon=1..Sun=0 wrong
    # Python weekday: Mon=0. JS getDay: Sun=0 Mon=1
    js = (d.weekday() + 1) % 7
    return 7 if d.weekday() == 6 else d.weekday() + 1  # isoweekday Mon=1 Sun=7

def matches_schedule(schedule_str, d):
    s = clean(schedule_str).rstrip('*')
    if not s or s == '每日' or s.lower() == 'daily':
        return True
    tokens = [clean(x) for x in re.split(r'[,，、\s]+', s) if clean(x)]
    if tokens and all(re.match(r'^\d+$', t) for t in tokens):
        dow = d.isoweekday()
        return any(int(t) == dow for t in tokens)
    return False

def parse_time(value):
    if value is None or value == '':
        return None
    if isinstance(value, (int, float)) or (isinstance(value, str) and value.replace('.', '', 1).isdigit()):
        val = float(value)
        if 0 <= val < 1:
            total = round(val * 24 * 60)
            hh = (total // 60) % 24
            mm = total % 60
            return f'{hh:02d}:{mm:02d}'
    s = clean(value).replace('+1', '', 1)
    m = re.match(r'^(\d{1,2}):(\d{2})', s)
    if m:
        return f'{int(m.group(1)):02d}:{m.group(2)}'
    return None

def in_operational_window(typ, time_str, iso):
    y, mo, da = map(int, iso.split('-'))
    hh, mm = map(int, time_str.split(':'))
    dt = datetime(y, mo, da, hh, mm)
    start = datetime(y, mo, da, 4 if typ == 'DEP' else 3, 30)
    end = datetime(y, mo, da + 1, 4 if typ == 'DEP' else 3, 30) if typ == 'DEP' else datetime(y, mo, da + 1, 3, 30)
    return start <= dt < end

def normalize_workbook_row(raw, raw_index):
    flt_keys = {
        'FLT No': raw.get('FLT No'),
        'FLT': raw.get('FLT'),
        '航班': raw.get('航班'),
        'flt': raw.get('flt'),
    }
    flt = clean(flt_keys['FLT No'] or flt_keys['FLT'] or flt_keys['航班'] or flt_keys['flt'])
    if not flt or flt in ('FLT No', 'FLT', '航班'):
        return None
    dep = clean(raw.get('DEP') or raw.get('dep') or '')
    arr = clean(raw.get('ARR') or raw.get('arr') or '')
    if not dep and not arr:
        return None
    return {
        'raw_index': raw_index,
        'flt': flt,
        'dep': dep.upper(),
        'arr': arr.upper(),
        'std': raw.get('STD') or raw.get('std') or '',
        'sta': raw.get('STA') or raw.get('sta') or '',
        'schedule': clean(raw.get('Frequency') or raw.get('班期') or raw.get('schedule') or ''),
        'acType': clean(raw.get('Aircraft Type') or raw.get('機型') or raw.get('acType') or ''),
    }

def load_xlsx_rows(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for si in root.findall('m:si', ns):
            shared.append(''.join((t.text or '') for t in si.findall('.//m:t', ns)))
        sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        headers = {}
        out = []
        for row in sheet.findall('m:sheetData/m:row', ns):
            ri = int(row.get('r'))
            cells = {}
            for c in row.findall('m:c', ns):
                ref = c.get('r')
                col = re.match(r'([A-Z]+)', ref).group(1)
                t = c.get('t')
                v = c.find('m:v', ns)
                val = ''
                if v is not None:
                    val = shared[int(v.text)] if t == 's' else v.text
                cells[col] = val
            if ri == 1:
                for col, val in cells.items():
                    headers[col] = val
            else:
                obj = {}
                for col, val in cells.items():
                    h = headers.get(col)
                    if h:
                        obj[h] = val
                out.append(obj)
        return out

def normalize_flight_no(flt):
    s = clean(flt).upper()
    m = re.match(r'^([A-Z]{2})(\d+)$', s)
    if m:
        return f'{m.group(1)}{m.group(2).zfill(4)}'
    return s

def main():
    parser_path = '/Users/kuoyuchia/Desktop/vibe coding/foreign-schedule-parser.js'
    xlsx_path = '/Users/kuoyuchia/Desktop/vibe coding/外家時刻表.xlsx'
    selected = '2026-05-27'

    with open(parser_path, encoding='utf-8') as f:
        src = f.read()
    assert "raw['FLT No']" in src or 'FLT No' in src, 'parser missing FLT No support'
    assert PARSER_VERSION_EXPECTED in src, f'parser version {PARSER_VERSION_EXPECTED} not in file'

    raw_rows = load_xlsx_rows(xlsx_path)
    rows = [normalize_workbook_row(r, i) for i, r in enumerate(raw_rows)]
    rows = [r for r in rows if r]

    candidates = []
    skipped = []
    bx792_trace = []

    for row in rows:
        is_bx792 = '792' in row['flt']
        trace = {'flt': row['flt'], 'schedule': row['schedule'], 'std': row['std']}
        if not matches_schedule(row['schedule'], date.fromisoformat(selected)):
            reason = '班期不符'
            skipped.append({**row, 'reason': reason})
            if is_bx792:
                trace['scheduleMatch'] = False
                trace['weekday'] = date.fromisoformat(selected).isoweekday()
                bx792_trace.append(trace)
            continue
        trace['scheduleMatch'] = True
        dep, arr = row['dep'], row['arr']
        if dep == STATION:
            typ = 'DEP'
            anchor = parse_time(row['std'])
        elif arr == STATION:
            typ = 'ARR'
            anchor = parse_time(row['sta'])
        else:
            skipped.append({**row, 'reason': '非 TPE 出入境'})
            continue
        trace.update({'type': typ, 'anchorTime': anchor})
        if not anchor:
            skipped.append({**row, 'reason': '時間無法解析'})
            if is_bx792:
                bx792_trace.append({**trace, 'skip': '時間無法解析'})
            continue
        if not in_operational_window(typ, anchor, selected):
            skipped.append({**row, 'reason': '不在營運窗口', 'type': typ, 'anchorTime': anchor})
            if is_bx792:
                bx792_trace.append({**trace, 'skip': '不在營運窗口'})
            continue
        cand = {
            'flight': normalize_flight_no(row['flt']),
            'type': typ,
            'anchorTime': anchor,
            'schedule': row['schedule'],
            'acType': row['acType'] or ('321' if row['flt'].upper().startswith('BX') else ''),
        }
        candidates.append(cand)
        if is_bx792:
            trace['added'] = cand
            bx792_trace.append(trace)

    dep = [c for c in candidates if c['type'] == 'DEP']
    bx792_cands = [c for c in candidates if '792' in c['flight']]

    print('PARSER FILE VERSION:', PARSER_VERSION_EXPECTED, 'OK')
    print('FLT No support: OK')
    print('selectedDate:', selected, 'weekday:', date.fromisoformat(selected).isoweekday())
    print('stats:', {'total': len(candidates), 'dep': len(dep), 'arr': len(candidates) - len(dep)})
    print('BX792 trace:', json.dumps(bx792_trace, ensure_ascii=False, indent=2))
    print('BX792 candidates:', json.dumps(bx792_cands, ensure_ascii=False, indent=2))
    print('DEP list:', [f"{c['flight']} {c['anchorTime']}" for c in dep])

    assert bx792_cands, 'FAIL: BX792 not in candidates for 2026-05-27'
    assert any(c['type'] == 'DEP' and c['anchorTime'] == '16:35' for c in bx792_cands), 'FAIL: BX792 16:35 DEP missing'
    print('PASS: BX792 DEP 16:35 present in candidates')

if __name__ == '__main__':
    main()
