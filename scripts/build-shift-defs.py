#!/usr/bin/env python3
"""Generate SHIFT_DEFS block for index.html from 班別時間 Excel."""
import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

DEFAULT_XLSX = Path.home() / 'Desktop/當日人力/新的班別時間.xlsx'


def read_xlsx(path: Path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in root.findall('m:si', ns):
                texts = [t.text or '' for t in si.findall('.//m:t', ns)]
                shared.append(''.join(texts))
        sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        rows = []
        for row in sheet.findall('m:sheetData/m:row', ns):
            vals = []
            for c in row.findall('m:c', ns):
                t = c.get('t')
                v = c.find('m:v', ns)
                if v is None:
                    vals.append('')
                elif t == 's':
                    vals.append(shared[int(v.text)])
                else:
                    vals.append(v.text)
            rows.append(vals)
        return rows


def excel_time(v):
    if not v or str(v).strip() in ('', '-'):
        return ''
    s = str(v).strip()
    if ':' in s:
        return s
    try:
        f = float(v)
        h = int(f * 24)
        m = int(round((f * 24 - h) * 60))
        if h == 24 and m == 0:
            return '24:00'
        return f'{h:02d}:{m:02d}'
    except ValueError:
        return ''


def js_str(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


def build_defs(rows):
    out = []
    for r in rows[1:]:
        if len(r) < 5:
            continue
        code = str(r[1]).strip()
        if not code:
            continue
        start = excel_time(r[2])
        end = excel_time(r[3])
        kind = str(r[4]).strip()
        common = str(r[5]).strip() == 'O' if len(r) > 5 else False
        note = str(r[6]).strip() if len(r) > 6 else ''
        out.append({
            'code': code,
            'start': start,
            'end': end,
            'kind': kind,
            'common': common,
            'note': note,
        })
    return out


def render_js(defs):
    lines = ['const SHIFT_DEFS = {']
    for d in defs:
        c = js_str(d['code'])
        lines.append(
            f'  {c}: {{ start: {js_str(d["start"])}, end: {js_str(d["end"])}, '
            f'kind: {js_str(d["kind"])}, common: {"true" if d["common"] else "false"}, '
            f'note: {js_str(d["note"])} }},'
        )
    lines.append('};')
    leave_req = [d['code'] for d in defs if d['kind'] == '請假']
    lines.append(f'const SHIFT_LEAVE_REQUEST_CODES = {json.dumps(leave_req, ensure_ascii=False)};')
    return '\n'.join(lines)


def patch_index(html_path: Path, block: str):
    text = html_path.read_text(encoding='utf-8')
    start = text.index('const SHIFT_DEFS = {')
    end = text.index('function getShiftDef(key)')
    text = text[:start] + block + '\n' + text[end:]
    html_path.write_text(text, encoding='utf-8')


def main():
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_XLSX
    html = Path(__file__).resolve().parent.parent / 'index.html'
    if not xlsx.exists():
        print(f'Missing {xlsx}', file=sys.stderr)
        sys.exit(1)
    defs = build_defs(read_xlsx(xlsx))
    block = render_js(defs)
    if '--stdout' in sys.argv:
        print(block)
        print(f'// {len(defs)} defs', file=sys.stderr)
        return
    patch_index(html, block)
    print(f'Patched {html} with {len(defs)} shift defs')


if __name__ == '__main__':
    main()
