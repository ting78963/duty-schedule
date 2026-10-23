from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from lxml import etree
import openpyxl, json, io, os, re, copy, zipfile, requests as req, hmac
from datetime import datetime

app = Flask(__name__)

SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')

def sb_headers():
    return {
        'apikey': SUPABASE_KEY,
        'Authorization': f'Bearer {SUPABASE_KEY}',
        'Content-Type': 'application/json',
        'Prefer': 'return=representation'
    }

def sb_get(table, params=None):
    r = req.get(f"{SUPABASE_URL}/rest/v1/{table}", headers=sb_headers(), params=params)
    return r.json()

def sb_post(table, data):
    headers = sb_headers()
    headers['Prefer'] = 'resolution=merge-duplicates,return=representation'
    r = req.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=headers, json=data)
    return r.json()

def sb_patch(table, match_key, match_val, data):
    r = req.patch(f"{SUPABASE_URL}/rest/v1/{table}", headers=sb_headers(),
                  params={match_key: f'eq.{match_val}'}, json=data)
    return r.json()

def sb_delete(table, match_key, match_val):
    r = req.delete(f"{SUPABASE_URL}/rest/v1/{table}", headers=sb_headers(),
                   params={match_key: f'eq.{match_val}'})
    return r.status_code

ASSETS = os.path.join(os.path.dirname(__file__), 'assets')
ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

TEMPLATE_BUCKET = 'duty-templates'
TEMPLATE_NAMES = {
    '平日': 'tmpl_平日.docx', '假日': 'tmpl_假日.docx',
    '平日取酒': 'tmpl_平日取酒.docx', '假日取酒': 'tmpl_假日取酒.docx'
}
DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

def template_bytes(tmpl_path):
    """Use persistent uploaded template first; bundled GitHub template as fallback."""
    name = os.path.basename(tmpl_path)
    if SUPABASE_URL and SUPABASE_KEY:
        url = f"{SUPABASE_URL}/storage/v1/object/authenticated/{TEMPLATE_BUCKET}/{name}"
        response = req.get(url, headers={'apikey': SUPABASE_KEY,
                            'Authorization': f'Bearer {SUPABASE_KEY}'}, timeout=20)
        if response.status_code == 200:
            validate_template_bytes(response.content)
            return response.content
        if response.status_code != 404:
            raise RuntimeError(f'無法讀取已上傳的範本：HTTP {response.status_code}')
    with open(tmpl_path, 'rb') as f:
        return f.read()

def validate_template_bytes(payload):
    if len(payload) > 5 * 1024 * 1024:
        raise ValueError('範本不得超過5MB')
    if not zipfile.is_zipfile(io.BytesIO(payload)):
        raise ValueError('檔案不是有效的Word docx')
    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        members = archive.infolist()
        if len(members) > 3000 or sum(item.file_size for item in members) > 60 * 1024 * 1024:
            raise ValueError('Word檔案內部容量異常，拒絕上傳')
        if 'word/document.xml' not in archive.namelist():
            raise ValueError('Word檔案缺少必要內容')
    doc = Document(io.BytesIO(payload))
    if not doc.tables:
        raise ValueError('Word範本沒有勤務表格')
    table = doc.tables[0]
    if len(table.columns) not in (33, 35, 37):
        raise ValueError('Word範本欄位數不符合勤務表')
    code_col = len(table.columns) - 5
    name_col = len(table.columns) - 1
    codes = {int(row.cells[code_col].text.strip()) for row in table.rows
             if row.cells[code_col].text.strip().isdigit()}
    if not {19, 29}.issubset(codes):
        raise ValueError('Word範本缺少正式代號19或29')
    for code, label in ((19, '休(補)假'), (29, '公假')):
        matching = [row for row in table.rows
                    if row.cells[code_col].text.strip() == str(code)]
        if not matching or label not in matching[0].cells[name_col].text:
            raise ValueError(f'代號{code}同排必須保留「{label}」標題')
    return doc


def get_tmpl(is_holiday, has_sake):
    if is_holiday and has_sake: return os.path.join(ASSETS, 'tmpl_假日取酒.docx')
    if is_holiday:              return os.path.join(ASSETS, 'tmpl_假日.docx')
    if has_sake:                return os.path.join(ASSETS, 'tmpl_平日取酒.docx')
    return os.path.join(ASSETS, 'tmpl_平日.docx')

def get_cols(is_holiday):
    if is_holiday:
        return 28, 29, 30, 31, 32
    return 30, 31, 32, 33, 34

def set_cell(cell, text, size=None, color='000000'):
    """Replace text while preserving the approved Word template's layout and run formatting."""
    text = str(text) if text else ''
    para = cell.paragraphs[0]
    runs = para.runs
    if runs:
        run = runs[0]
        run.text = text
        for extra in runs[1:]:
            extra.text = ''
    else:
        # Blank roster cells have no runs. Copy a formatted run from an
        # adjacent populated cell in the same table column when available.
        from copy import deepcopy
        reference = None
        parent_row = cell._tc.getparent()
        table = parent_row.getparent()
        cells = list(parent_row)
        col_idx = next((i for i, tc in enumerate(cells) if tc is cell._tc), None)
        if col_idx is not None:
            for other_row in table:
                if other_row.tag != f'{{{ns}}}tr':
                    continue
                other_cells = [c for c in other_row if c.tag == f'{{{ns}}}tc']
                if col_idx >= len(other_cells):
                    continue
                other_tc = other_cells[col_idx]
                for rpr in other_tc.iter(f'{{{ns}}}rPr'):
                    if rpr.find(f'{{{ns}}}rFonts') is not None:
                        reference = deepcopy(rpr)
                        break
                if reference is not None:
                    break
        run = para.add_run(text)
        if reference is not None:
            current = run._r.rPr
            if current is not None:
                run._r.remove(current)
            run._r.insert(0, reference)
    rpr = run._r.get_or_add_rPr()
    fonts = rpr.find(f'{{{ns}}}rFonts')
    if fonts is None:
        fonts = etree.SubElement(rpr, f'{{{ns}}}rFonts')
    for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        fonts.set(f'{{{ns}}}{attr}', '標楷體')
    if size is not None:
        for tag in ('sz', 'szCs'):
            el = rpr.find(f'{{{ns}}}{tag}')
            if el is None:
                el = etree.SubElement(rpr, f'{{{ns}}}{tag}')
            el.set(f'{{{ns}}}val', str(size))
    if color != '000000':
        el = rpr.find(f'{{{ns}}}color')
        if el is None:
            el = etree.SubElement(rpr, f'{{{ns}}}color')
        el.set(f'{{{ns}}}val', color)

def set_red_border(row):
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.find(f'{{{ns}}}tcPr')
        if tcPr is None:
            tcPr = etree.SubElement(tc, f'{{{ns}}}tcPr')
        tcBorders = tcPr.find(f'{{{ns}}}tcBorders')
        if tcBorders is None:
            tcBorders = etree.SubElement(tcPr, f'{{{ns}}}tcBorders')
        for side in ['top','bottom','left','right']:
            el = tcBorders.find(f'{{{ns}}}{side}')
            if el is None:
                el = etree.SubElement(tcBorders, f'{{{ns}}}{side}')
            el.set(f'{{{ns}}}val', 'single')
            el.set(f'{{{ns}}}sz', '16')
            el.set(f'{{{ns}}}color', 'FF0000')

def find_section_rows(t, oname_col, label):
    # Some approved templates (e.g. 11/20) repeat 休(補)假 inside the
    # same section. Keep those heading rows untouched and use the blank/name
    # slots on both sides of the repeated heading.
    start = None
    for r_idx in range(len(t.rows)):
        if oname_col >= len(t.rows[r_idx].cells): continue
        if label in t.rows[r_idx].cells[oname_col].text.strip():
            start = r_idx + 1
            break
    if start is None: return []
    rows = []
    for r_idx in range(start, len(t.rows)):
        if oname_col >= len(t.rows[r_idx].cells): break
        txt = t.rows[r_idx].cells[oname_col].text.strip()
        if label in txt:
            continue
        if txt and any(x in txt for x in ['假','附記','巡邏','調適']): break
        rows.append(r_idx)
    return rows

def parse_sake_time(cell_value):
    if not cell_value: return None
    m = re.search(r'(\d{1,2})-(\d{1,2})', str(cell_value))
    if m: return {'start': int(m.group(1)), 'end': int(m.group(2))}
    return None

def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    monthly_data = {}
    staff_from_excel = {}  # name -> True
    sake_by_date = {}

    # ── 外勤工作表 ──────────────────────────────────────
    ws_ext = None
    for name in ['外勤', 'Sheet1']:
        if name in wb.sheetnames:
            ws_ext = wb[name]
            break
    if not ws_ext:
        ws_ext = wb.worksheets[0]

    # 列2：人員姓名
    name_cols = {}
    for c in range(3, 42, 2):
        v = ws_ext.cell(row=2, column=c).value
        if v and isinstance(v, str) and len(v) >= 2 and v not in ['當日休息人數','固定番人數','日期','星期']:
            name_cols[c] = v
            staff_from_excel[v] = True

    # 列3以後：每天
    for row in range(3, ws_ext.max_row+1):
        date_val = ws_ext.cell(row=row, column=1).value
        if not date_val: continue
        if isinstance(date_val, datetime):
            date_obj = date_val
        else:
            try:
                date_obj = datetime.strptime(str(date_val)[:10], '%Y-%m-%d')
            except:
                continue

        roc_year = date_obj.year - 1911
        month = date_obj.month
        day = date_obj.day
        weekday = ['一','二','三','四','五','六','日'][date_obj.weekday()]
        is_holiday = date_obj.weekday() >= 5
        date_str = f'{roc_year}年{month}月{day}日'

        on_duty = []
        off_duty = []

        for col, name in name_cols.items():
            # 番號在姓名左欄，狀態在右欄
            ban_val = ws_ext.cell(row=row, column=col).value
            status_val = ws_ext.cell(row=row, column=col+1).value

            ban = None
            if ban_val is not None:
                ban_str = str(ban_val).strip()
                if ban_str.isdigit():
                    ban = int(ban_str)

            status = ''
            if status_val and str(status_val).strip():
                raw = str(status_val).strip()
                for s in ['喪假','輪','優','休','補','公']:
                    if s in raw:
                        status = s
                        break
                if not status:
                    status = raw[:2]  # 常訓等備考

            person = {'name': name, 'ban': ban, 'status': status}
            off_statuses = ['輪','優','休','補','公','喪假']
            if any(s == status for s in off_statuses):
                off_duty.append(person)
            else:
                on_duty.append(person)

        monthly_data[date_str] = {
            'date_str': date_str,
            'roc_date': date_str,
            'weekday': weekday,
            'is_holiday': is_holiday,
            'has_sake': False,
            'sake_times': [],
            'on_duty': sorted(on_duty, key=lambda x: x['ban'] or 99),
            'off_duty': sorted(off_duty, key=lambda x: x['ban'] or 99),
        }

    # ── 內勤：使用已驗收 Word 的正式代號，Excel 決定當日姓名/狀態 ──
    # 不以 Excel 欄位順序推測代號；沒有代號的新進人員不寫入。
    internal_codes = {
        '林筠臻':23, '黃雅琳':26, '洪君智':27, '楊智為':28,
        '陳淑卿':29, '謝函儒':31, '曾義雄':33, '姚宗志':35,
        '柯育婷':36, '林裕善':38, '王庭偉':39, '林家豪':41,
        '張銥凌':42, '翁水成':43, '洪呈欣':44, '簡竹芳':45,
        '沈宗漢':46, '曾煙煌':47, '陳駿騰':48, '李智偉':50,
        '林新慶':51, '李金孔':52, '董安民':53, '陳昱靜':54,
    }
    if '全隊大表' in wb.sheetnames:
        ws_all = wb['全隊大表']
        cols = [(c, str(ws_all.cell(2, c).value).strip())
                for c in range(49, 75) if ws_all.cell(2, c).value]
        for row in range(3, ws_all.max_row + 1):
            date_val = ws_all.cell(row, 47).value
            if not isinstance(date_val, datetime):
                continue
            key = f'{date_val.year - 1911}年{date_val.month}月{date_val.day}日'
            if key not in monthly_data:
                continue
            day = monthly_data[key]
            for col, name in cols:
                code = internal_codes.get(name)
                if code is None:
                    continue
                raw = str(ws_all.cell(row, col).value or '').strip()
                is_field = name in ('李金孔', '林新慶', '李智偉', '董安民')
                stopped_leave = is_field and '輪' in raw and ('Δ' in raw or '△' in raw or '▲' in raw or '▵' in raw)
                person = {'name': name, 'ban': code,
                          'status': '留職停薪' if name == '林筠臻' else ('輪' if is_field and '輪' in raw and not stopped_leave else raw),
                          'stopped_leave': stopped_leave}
                # 三角形加「輪」表示停休支援取締酒駕：備考顯示停休，不列右側輪休。
                is_leave = not stopped_leave and (raw in ('優', '休', '補', '公') or (is_field and '輪' in raw))
                (day['off_duty'] if is_leave else day['on_duty']).append(person)
                staff_from_excel[name] = True
            day['on_duty'].sort(key=lambda p: p.get('ban') or 999)
            day['off_duty'].sort(key=lambda p: p.get('ban') or 999)

    # ── 全隊大表：取締酒駕 ──────────────────────────────
    if '全隊大表' in wb.sheetnames:
        ws_all = wb['全隊大表']
        for row in range(1, ws_all.max_row+1):
            date_val = ws_all.cell(row=row, column=1).value
            ao_val = ws_all.cell(row=row, column=41).value
            if date_val and ao_val and '取締' in str(ao_val):
                sake_time = parse_sake_time(ao_val)
                if sake_time:
                    if isinstance(date_val, datetime):
                        date_obj = date_val
                    else:
                        try:
                            date_obj = datetime.strptime(str(date_val)[:10], '%Y-%m-%d')
                        except:
                            continue
                    roc_year = date_obj.year - 1911
                    date_str = f'{roc_year}年{date_obj.month}月{date_obj.day}日'
                    sake_by_date[date_str] = sake_time
                    if date_str in monthly_data:
                        monthly_data[date_str]['has_sake'] = True
                        monthly_data[date_str]['sake_times'] = [sake_time]

    return {
        'monthlyData': monthly_data,
        'staffFromExcel': list(staff_from_excel.keys()),
        'sakeByDate': sake_by_date,
        'totalDays': len(monthly_data)
    }

def save_to_supabase(monthly_data):
    """把解析好的大表資料存到 Supabase monthly_duty"""
    records = []
    for date_str, d in monthly_data.items():
        records.append({
            'date_str': date_str,
            'roc_date': d['roc_date'],
            'weekday': d['weekday'],
            'is_holiday': d['is_holiday'],
            'has_sake': d['has_sake'],
            'sake_times': d['sake_times'],
            'on_duty': d['on_duty'],
            'off_duty': d['off_duty'],
        })
    # 批次 upsert（每次50筆）
    for i in range(0, len(records), 50):
        batch = records[i:i+50]
        sb_post('monthly_duty', batch)
    return len(records)

def sync_staff(staff_names):
    """把大表人員同步到 Supabase staff"""
    existing = sb_get('staff', params={'type': 'eq.輪班', 'order': 'code'})
    existing_names = {s['name'] for s in existing if s.get('name')}
    new_names = [n for n in staff_names if n not in existing_names]
    return new_names  # 回傳新人員名單讓前端確認

def generate_word(data):
    on_duty  = data['onDuty']
    off_duty = data['offDuty']
    roc_date = data['rocDate']
    weekday  = data['weekday']
    is_holiday = data['isHoliday']
    has_sake   = data.get('hasSake', False)
    sake_times = data.get('sakeTimes', [])

    tmpl_path = get_tmpl(is_holiday, has_sake)
    # Bundled template is a fallback; the latest upload lives in Supabase Storage.

    # The DUI templates have two extra columns; derive roster positions from
    # the actual Word table rather than assuming every weekday/holiday layout matches.

    lx, bx, gg = [], [], []
    for s in sorted(off_duty, key=lambda x: x.get('ban') or 99):
        st = s.get('status') or ''
        if st == '輪': lx.append(s)
        elif st in ['優','休','補','喪假']: bx.append(s)
        elif st == '公': gg.append(s)

    doc = Document(io.BytesIO(template_bytes(tmpl_path)))
    t = doc.tables[0]
    table_width = len(t.columns)
    if table_width not in (33, 35, 37):
        raise ValueError(f'不支援的勤務表欄數：{table_width}')
    CODE_COL = table_width - 5
    NAME_COL, NOTE_COL = CODE_COL + 1, CODE_COL + 2
    OCODE_COL, ONAME_COL = CODE_COL + 3, CODE_COL + 4

    ban_to_row = {}
    for r_idx in range(4, len(t.rows)):
        if CODE_COL >= len(t.rows[r_idx].cells): continue
        ct = t.rows[r_idx].cells[CODE_COL].text.strip()
        if ct.isdigit() and 1 <= int(ct) <= 54:
            ban_to_row[int(ct)] = r_idx

    for ban, r_idx in ban_to_row.items():
        for ci in [NAME_COL, NOTE_COL]:
            for para in t.rows[r_idx].cells[ci].paragraphs:
                for run in para.runs: run.text = ''

    for s in on_duty:
        ban = s.get('ban')
        if not ban or ban not in ban_to_row: continue
        set_cell(t.rows[ban_to_row[ban]].cells[NAME_COL], s['name'])
        set_cell(t.rows[ban_to_row[ban]].cells[NOTE_COL], '留職停薪' if ban == 23 and s['name'] == '林筠臻' else ('停休' if s.get('stopped_leave') else ''))

    for s in off_duty:
        ban = s.get('ban')
        if not ban or ban not in ban_to_row: continue
        set_cell(t.rows[ban_to_row[ban]].cells[NAME_COL], s['name'])
        set_cell(t.rows[ban_to_row[ban]].cells[NOTE_COL], '留職停薪' if ban == 23 and s['name'] == '林筠臻' else ('停休' if s.get('stopped_leave') else ''))

    # The leave heading belongs on the row of OFFICIAL ROSTER CODE 19,
    # never on the physical nineteenth Word table row.
    from copy import deepcopy
    leave_anchor = ban_to_row.get(19)
    if leave_anchor is None:
        raise ValueError('Word範本找不到人員代號19，無法定位休(補)假')
    heading_rows = [i for i in range(len(t.rows))
                    if '休(補)假' in t.rows[i].cells[ONAME_COL].text]
    if not heading_rows:
        raise ValueError('Word範本缺少休(補)假標題')
    if leave_anchor not in heading_rows:
        source = t.rows[heading_rows[0]].cells[ONAME_COL].paragraphs[0]
        target = t.rows[leave_anchor].cells[ONAME_COL].paragraphs[0]
        if target.text.strip():
            raise ValueError('代號19同排已有文字，無法安全移動休(補)假標題')
        if source._p.pPr is not None:
            if target._p.pPr is not None:
                target._p.remove(target._p.pPr)
            target._p.insert(0, deepcopy(source._p.pPr))
        for run_xml in list(target._p.r_lst):
            target._p.remove(run_xml)
        for run_xml in source._p.r_lst:
            target._p.append(deepcopy(run_xml))
    for r_idx in heading_rows:
        if r_idx != leave_anchor:
            set_cell(t.rows[r_idx].cells[ONAME_COL], '')
            set_cell(t.rows[r_idx].cells[OCODE_COL], '')

    # The leave heading belongs on the row of OFFICIAL ROSTER CODE 29,
    # never on the physical nineteenth Word table row.
    from copy import deepcopy
    public_anchor = ban_to_row.get(29)
    if public_anchor is None:
        raise ValueError('Word範本找不到人員代號29，無法定位公假')
    public_heading_rows = [i for i in range(len(t.rows))
                    if '公假' in t.rows[i].cells[ONAME_COL].text]
    if not public_heading_rows:
        raise ValueError('Word範本缺少公假標題')
    if public_anchor not in public_heading_rows:
        source = t.rows[public_heading_rows[0]].cells[ONAME_COL].paragraphs[0]
        target = t.rows[public_anchor].cells[ONAME_COL].paragraphs[0]
        if target.text.strip():
            raise ValueError('代號29同排已有文字，無法安全移動公假標題')
        if source._p.pPr is not None:
            if target._p.pPr is not None:
                target._p.remove(target._p.pPr)
            target._p.insert(0, deepcopy(source._p.pPr))
        for run_xml in list(target._p.r_lst):
            target._p.remove(run_xml)
        for run_xml in source._p.r_lst:
            target._p.append(deepcopy(run_xml))
    for r_idx in public_heading_rows:
        if r_idx != public_anchor:
            set_cell(t.rows[r_idx].cells[ONAME_COL], '')
            set_cell(t.rows[r_idx].cells[OCODE_COL], '')

    # Roster slots are defined by the two approved fixed headings.
    # Do not let legacy heading positions in old templates reduce capacity.
    wheel_headings = [i for i in range(leave_anchor)
                      if '輪' in t.rows[i].cells[ONAME_COL].text
                      and '休' in t.rows[i].cells[ONAME_COL].text]
    if not wheel_headings:
        raise ValueError('Word範本缺少輪休標題')
    lx_rows = list(range(wheel_headings[0] + 1, leave_anchor))

    # Leave entries are bounded by the fixed public-leave heading at code 29.
    # Do not reuse legacy scattered positions from older Word templates.
    bx_rows = []
    for r_idx in range(leave_anchor + 1, public_anchor):
        txt = t.rows[r_idx].cells[ONAME_COL].text.strip()
        if txt and any(x in txt for x in ('公假', '附記', '巡邏', '調適')):
            break
        if '休(補)假' in txt:
            continue
        bx_rows.append(r_idx)
    gg_rows = [i for i in range(public_anchor + 1, len(t.rows))
               if not any(x in t.rows[i].cells[ONAME_COL].text
                          for x in ('附記', '巡邏', '調適'))]

    for rows in [lx_rows, bx_rows, gg_rows]:
        for r_idx in rows:
            set_cell(t.rows[r_idx].cells[OCODE_COL], '')
            set_cell(t.rows[r_idx].cells[ONAME_COL], '')

    # Merged right-hand code cells can silently overwrite two different staff
    # codes (seen in the 11/14 approved template). Refuse unsafe output.
    for section_name, people, rows in [('輪休', lx, lx_rows), ('休補假', bx, bx_rows), ('公假', gg, gg_rows)]:
        selected = rows[:len(people)]
        for col, label in ((OCODE_COL, '代號'), (ONAME_COL, '姓名')):
            ids = [id(t.rows[i].cells[col]._tc) for i in selected]
            if len(ids) != len(set(ids)):
                raise ValueError(f'{section_name}{label}欄有合併儲存格，無法逐人填入；請修正Word範本')
    for section_name, people, rows in [('輪休', lx, lx_rows), ('休補假', bx, bx_rows), ('公假', gg, gg_rows)]:
        if len(people) > len(rows):
            raise ValueError(f'{section_name}名單有{len(people)}人，但Word只有{len(rows)}個欄位，已停止產生以免漏人')
        for j, s in enumerate(people):
            set_cell(t.rows[rows[j]].cells[OCODE_COL], str(s.get('ban','')))
            set_cell(t.rows[rows[j]].cells[ONAME_COL], s['name'])

    day = roc_date.split('年')[1].split('日')[0] + '日'
    for cell in t.rows[0].cells:
        if '民國' in cell.text and '年' in cell.text:
            for para in cell.paragraphs:
                for run in para.runs:
                    if '年月日【' in run.text:
                        run.text = f'年{day}【'
                    elif run.text == '星期':
                        run.text = f'星期{weekday}'
            break

    if sake_times and has_sake:
        time_to_row = {}
        for r_idx in range(len(t.rows)):
            row = t.rows[r_idx]
            if not row.cells: continue
            cell_txt = row.cells[0].text.strip()
            m = re.match(r'(\d{1,2})[－\-–](\d{1,2})', cell_txt)
            if m:
                h = int(m.group(1))
                time_to_row[h] = r_idx

        for st in sake_times:
            s_h = int(st['start'])
            e_h = int(st['end'])
            hours = []
            h = s_h
            while True:
                hours.append(h % 24)
                h += 1
                if h % 24 == e_h % 24: break
                if len(hours) > 16: break
            for h in hours:
                if h in time_to_row:
                    set_red_border(t.rows[time_to_row[h]])

    # Clear only daily assignment codes after all other Word edits are complete.
    # Unknown template layouts abort instead of risking a damaged duty sheet.
    from duty_blank_safe import blank_duty_document
    blank_duty_document(doc)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ── API ──────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/staff', methods=['GET'])
def get_staff():
    if not SUPABASE_URL: return jsonify([])
    data = sb_get('staff', params={'order': 'code'})
    return jsonify(data)

@app.route('/api/staff', methods=['POST'])
def add_staff():
    return jsonify(sb_post('staff', request.json))

@app.route('/api/staff/<int:code>', methods=['PUT'])
def update_staff(code):
    return jsonify(sb_patch('staff', 'code', code, request.json))

@app.route('/api/staff/<int:code>', methods=['DELETE'])
def delete_staff(code):
    sb_delete('staff', 'code', code)
    return jsonify({'ok': True})

@app.route('/api/parse-excel', methods=['POST'])
def parse_excel_api():
    f = request.files.get('file')
    if not f: return jsonify({'error': '未上傳檔案'}), 400
    try:
        result = parse_excel(f.read())
        # 自動存到 Supabase
        saved = save_to_supabase(result['monthlyData'])
        result['savedDays'] = saved
        # 回傳新人員名單（前端確認後再同步）
        new_staff = sync_staff(result['staffFromExcel'])
        result['newStaff'] = new_staff
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/duty/<date_str>', methods=['GET'])
def get_duty(date_str):
    """從 Supabase 取特定日期的勤務資料"""
    if not SUPABASE_URL: return jsonify(None)
    data = sb_get('monthly_duty', params={'date_str': f'eq.{date_str}'})
    if data and len(data) > 0:
        d = data[0]
        return jsonify({
            'rocDate': d['roc_date'],
            'weekday': d['weekday'],
            'isHoliday': d['is_holiday'],
            'hasSake': d['has_sake'],
            'sakeTimes': d['sake_times'],
            'onDuty': d['on_duty'],
            'offDuty': d['off_duty'],
        })
    return jsonify(None)

@app.route('/api/duty-month/<int:year>/<int:month>', methods=['GET'])
def get_duty_month(year, month):
    """取某月所有勤務資料（給月曆用）"""
    if not SUPABASE_URL: return jsonify([])
    prefix = f'{year}年{month}月'
    data = sb_get('monthly_duty', params={
        'date_str': f'like.{prefix}%',
        'order': 'date_str'
    })
    return jsonify(data)

@app.route('/api/generate', methods=['POST'])
def generate_api():
    data = request.json
    try:
        buf = generate_word(data)
        fname = f'勤務表_{data.get("rocDate","")}.docx'
        return send_file(buf, as_attachment=True, download_name=fname,
                        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-batch', methods=['POST'])
def generate_batch():
    days_data = request.json
    zip_buf = io.BytesIO()
    failures = []
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for d in days_data:
            try:
                buf = generate_word(d)
                fname = f"勤務表_{d.get('rocDate','')}.docx"
                zf.writestr(fname, buf.read())
            except Exception as exc:
                failures.append({'date': d.get('rocDate', ''), 'error': str(exc)})
    if failures:
        return jsonify({'error': '部分日期產生失敗，未提供不完整的壓縮檔', 'failures': failures}), 422
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True, download_name='勤務表批次.zip',
                    mimetype='application/zip')

@app.route('/api/upload-template', methods=['POST'])
def upload_template():
    tmpl_type = request.form.get('type')
    f = request.files.get('file')
    if not f or tmpl_type not in TEMPLATE_NAMES:
        return jsonify({'error': '請選擇正確的範本類型與Word檔案'}), 400
    if not SUPABASE_URL or not SUPABASE_KEY:
        return jsonify({'error': '尚未設定永久儲存服務，未執行上傳'}), 503
    expected_token = os.environ.get('TEMPLATE_UPLOAD_TOKEN', '')
    supplied_token = request.headers.get('X-Template-Upload-Token', '')
    if not expected_token or not hmac.compare_digest(supplied_token, expected_token):
        return jsonify({'error': '範本上傳未授權，請輸入管理密碼'}), 403
    try:
        payload = f.read(5 * 1024 * 1024 + 1)
        validate_template_bytes(payload)
    except Exception as exc:
        return jsonify({'error': str(exc)}), 400
    name = TEMPLATE_NAMES[tmpl_type]
    url = f"{SUPABASE_URL}/storage/v1/object/{TEMPLATE_BUCKET}/{name}"
    headers = {'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}',
               'Content-Type': DOCX_MIME, 'x-upsert': 'true',
               'Cache-Control': 'no-cache'}
    try:
        result = req.post(url, headers=headers, data=payload, timeout=30)
        if result.status_code not in (200, 201):
            return jsonify({'error': f'永久儲存失敗 HTTP {result.status_code}'}), 502
        # Verify the uploaded bytes are actually readable.
        saved = template_bytes(os.path.join(ASSETS, name))
        if saved != payload:
            return jsonify({'error': '儲存後內容驗證失敗'}), 502
    except Exception as exc:
        return jsonify({'error': f'範本儲存或驗證失敗：{exc}'}), 502
    return jsonify({'ok': True, 'type': tmpl_type, 'persistent': True})


if __name__ == '__main__':
    app.run(debug=True)