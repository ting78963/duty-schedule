from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from lxml import etree
import openpyxl, json, io, os, re, copy, zipfile, requests as req, hmac, uuid
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
    params = {'on_conflict': 'date_str'} if table == 'monthly_duty' else None
    r = req.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=headers, params=params, json=data, timeout=30)
    if not r.ok:
        raise RuntimeError(f'Supabase {table} 寫入失敗 HTTP {r.status_code}: {r.text[:300]}')
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
# Supabase Storage object keys use stable ASCII names; user-facing Word filenames remain Chinese.
TEMPLATE_STORAGE_NAMES = {
    'tmpl_平日.docx': 'weekday.docx',
    'tmpl_假日.docx': 'holiday.docx',
    'tmpl_平日取酒.docx': 'weekday_sake.docx',
    'tmpl_假日取酒.docx': 'holiday_sake.docx',
}
DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

def template_bytes(tmpl_path):
    """Read the uploaded private Supabase template first, then use bundled file if no upload exists."""
    name = os.path.basename(tmpl_path)
    storage_name = TEMPLATE_STORAGE_NAMES.get(name, name)
    if SUPABASE_URL and SUPABASE_KEY:
        url = f"{SUPABASE_URL}/storage/v1/object/authenticated/{TEMPLATE_BUCKET}/{storage_name}"
        response = req.get(url + '?v=' + uuid.uuid4().hex, headers={
            'apikey': SUPABASE_KEY,
            'Authorization': f'Bearer {SUPABASE_KEY}',
            'Cache-Control': 'no-cache'
        }, timeout=20)
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
    text = str(text) if text else ''
    for para in cell.paragraphs:
        runs = para.runs
        if runs:
            runs[0].text = text
            for rr in runs[1:]: rr.text = ''
            rpr = runs[0]._r.find(f'{{{ns}}}rPr')
            if rpr is None:
                rpr = etree.SubElement(runs[0]._r, f'{{{ns}}}rPr')
                runs[0]._r.insert(0, rpr)
            fonts = rpr.find(f'{{{ns}}}rFonts')
            if fonts is None: fonts = etree.SubElement(rpr, f'{{{ns}}}rFonts')
            for attr in ['ascii','hAnsi','eastAsia','cs']:
                fonts.set(f'{{{ns}}}{attr}', '標楷體')
            col_el = rpr.find(f'{{{ns}}}color')
            if col_el is None: col_el = etree.SubElement(rpr, f'{{{ns}}}color')
            col_el.set(f'{{{ns}}}val', color)
            if size:
                for tag in ['sz','szCs']:
                    el = rpr.find(f'{{{ns}}}{tag}')
                    if el is None: el = etree.SubElement(rpr, f'{{{ns}}}{tag}')
                    el.set(f'{{{ns}}}val', str(size))
            # When size is unspecified, retain the template's exact run font size.
            # Removing w:sz/w:szCs made narrow remarks (e.g. 內部管理) wrap.
            return
    cell.paragraphs[0].add_run(text)

def set_cell_preserving_format(cell, text):
    """Replace visible text without changing approved Word run/paragraph formatting."""
    text = str(text) if text is not None else ''
    paragraphs = cell.paragraphs
    if not paragraphs:
        return
    first = True
    for para in paragraphs:
        for run in para.runs:
            run.text = text if first else ''
            first = False
    if first:
        paragraphs[0].add_run(text)


def set_red_border_cells(cells):
    for cell in cells:
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

def find_section_rows(t, oname_col, label, row_cells=None):
    # Reuse resolved rows: python-docx repeatedly reconstructs merged-cell grids.
    cells_at = (lambda i: row_cells[i]) if row_cells is not None else (lambda i: t.rows[i].cells)
    start = None
    for r_idx in range(len(t.rows)):
        cells = cells_at(r_idx)
        if oname_col >= len(cells): continue
        if label in cells[oname_col].text.strip():
            start = r_idx + 1
            break
    if start is None: return []
    rows = []
    for r_idx in range(start, len(t.rows)):
        cells = cells_at(r_idx)
        if oname_col >= len(cells): break
        txt = cells[oname_col].text.strip()
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
    for c in range(3, min(ws_ext.max_column, 41)+1, 2):
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
            # The numbered shift is in the SAME column as the staff name
            # (e.g. C2=洪立建, C3=1); the status is in the next column.
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
                if '停休' in raw:
                    status = '停休'
                elif '公假' in raw:
                    status = '公假'
                elif any(k in raw for k in ['喪假','病假','事假','婚假','產假','陪產假','家庭照顧假','生理假','特休']):
                    status = raw
                else:
                    for s in ['輪','優','休','補','公']:
                        if s in raw:
                            status = s
                            break
                    if not status:
                        status = raw  # 保留其他假別及常訓等備考

            person = {'name': name, 'ban': ban, 'status': status}
            # Any kind of leave is off-duty; 停休 means leave was cancelled.
            off_statuses = ['輪','優','休','補','公']
            is_leave = (status in off_statuses or ('假' in str(status) and '停休' not in str(status)))
            if is_leave:
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

    # Read the fixed code/name directory (BQ:BS) on every upload.
    # Blank named positions are meaningful; unnumbered newcomers are excluded.
    fixed_roster = {}
    if '全隊大表' in wb.sheetnames:
        roster_sheet = wb['全隊大表']
        for rr in range(1, roster_sheet.max_row + 1):
            code_raw = roster_sheet.cell(rr, 69).value
            if not str(code_raw or '').strip().isdigit():
                continue
            code = int(code_raw)
            if 21 <= code <= 54:
                raw_name = roster_sheet.cell(rr, 71).value
                fixed_roster[code] = str(raw_name or '').strip()
    # The daily parser currently reads the external roster first. Never
    # publish a supposedly complete daily roster while fixed internal staff
    # (21-54) are absent from the day's parsed personnel. This deliberately
    # fails closed until daily internal-status extraction is implemented
    # against the approved workbook layout.
    if fixed_roster:
        for date_str, day_data in monthly_data.items():
            people = day_data['on_duty'] + day_data['off_duty']
            by_code = {p.get('ban'): p.get('name') for p in people
                       if p.get('ban') is not None}
            missing = [(code, name) for code, name in fixed_roster.items()
                       if name and code not in by_code]
            mismatched = [(code, name, by_code[code])
                          for code, name in fixed_roster.items()
                          if name and code in by_code and by_code[code] != name]
            if missing or mismatched:
                missing_text = ', '.join(f'{code}:{name}' for code, name in missing[:8])
                mismatch_text = ', '.join(
                    f'{code}:{name}!={actual}' for code, name, actual in mismatched[:8])
                raise ValueError(
                    f'{date_str} 內勤 21-54 尚未完整從每日 Excel 解析；'
                    f'缺少 [{missing_text}]，姓名衝突 [{mismatch_text}]。'
                    '禁止將不完整名冊儲存或生成勤務表'
                )
    # The workbook does not define the supervisor/secretary 60-63 mapping.
    # Keep those separately verified records, never infer them from row order.

    # ── 全隊大表：主管督勤欄（依職稱，不依外勤姓名） ──
    # Excel's fixed roster (BQ:BS) is authoritative for internal code/name.
    supervisor_titles = ('隊長', '林副隊長', '陳副隊長', '李秘書', '姚組長', '翁組長')
    if '全隊大表' in wb.sheetnames:
        ws_all = wb['全隊大表']
        fixed_staff = {}
        for roster_row in range(1, ws_all.max_row + 1):
            code_val = ws_all.cell(roster_row, 69).value  # BQ
            name_val = ws_all.cell(roster_row, 71).value  # BS
            try:
                code = int(code_val)
            except (TypeError, ValueError):
                continue
            name = str(name_val or '').strip()
            if 21 <= code <= 99 and 2 <= len(name) <= 5:
                fixed_staff[name] = code
        # Resolve supervisor titles through their Excel roster names.
        # Titles without a verified name-to-code mapping are skipped safely.
        supervisor_names = {
            '隊長': '張瑞忠', '林副隊長': '林志明',
            '陳副隊長': '陳彥文', '李秘書': '李安景',
            '姚組長': '姚宗志', '翁組長': '翁水成',
        }
        # Find the supervisor headers in the top rows; do not assume fixed columns.
        supervisor_cols = {}
        supervisor_header_row = None
        for header_row in range(1, min(ws_all.max_row, 50) + 1):
            for col in range(1, ws_all.max_column + 1):
                label = str(ws_all.cell(header_row, col).value or '').replace(' ', '').replace('　', '').replace('\n', '')
                if label == '長隊':
                    label = '隊長'
                if label in supervisor_titles:
                    supervisor_cols[label] = col
                    supervisor_header_row = header_row

        for row in range((supervisor_header_row or ws_all.max_row) + 1, ws_all.max_row + 1):
            date_val = ws_all.cell(row, 1).value
            if isinstance(date_val, datetime):
                date_obj = date_val
            else:
                try:
                    date_obj = datetime.strptime(str(date_val)[:10], '%Y-%m-%d')
                except (TypeError, ValueError):
                    # The full-team duty sheet commonly stores only day numbers (1..31).
                    # Resolve against the month already parsed from the external roster.
                    try:
                        day_number = int(date_val)
                    except (TypeError, ValueError):
                        continue
                    if 30000 <= day_number <= 80000:
                        # Excel serial date (e.g. 46328 = 2026-11-02).
                        date_obj = openpyxl.utils.datetime.from_excel(day_number, wb.epoch)
                    elif 1 <= day_number <= 31:
                        candidates = [k for k in monthly_data if k.endswith(f'月{day_number}日')]
                        if len(candidates) != 1:
                            continue
                        date_str = candidates[0]
                        date_obj = None
                    else:
                        continue
            if date_obj is not None:
                date_str = f'{date_obj.year - 1911}年{date_obj.month}月{date_obj.day}日'
            if date_str not in monthly_data:
                continue
            day_data = monthly_data[date_str]
            for title, col in supervisor_cols.items():
                raw = str(ws_all.cell(row, col).value or '').strip()
                if not raw:
                    continue
                name = supervisor_names[title]
                # Verified supervisor code crosswalk supplied for this schedule.
                # Excel remains authoritative for each day's duty/leave status.
                verified_supervisor_codes = {
                    '隊長': 63, '林副隊長': 62, '陳副隊長': 61,
                    '李秘書': 60, '姚組長': 35, '翁組長': 43,
                }
                code = verified_supervisor_codes[title]
                roster_code = fixed_staff.get(name)
                if roster_code is not None and roster_code != code:
                    raise ValueError(
                        f'{date_str} 主管 {name} Excel 名冊代號 {roster_code} '
                        f'與正式對照代號 {code} 不符'
                    )
                status = ''
                if '公假' in raw or '公休' in raw:
                    status = '公假'
                elif '輪休' in raw or raw == '輪':
                    status = '輪休'
                elif '休假' in raw or raw in ('休', '優', '補'):
                    status = '休假'
                if not status:
                    continue
                person = {'name': name, 'ban': code, 'status': status}
                # Codes 35/43 are already in the internal roster. Reconcile
                # the supervisor status with that same person, never add a
                # second record or silently discard a mismatched name.
                existing = [p for p in day_data['on_duty'] + day_data['off_duty']
                            if p.get('ban') == code or p.get('name') == name]
                if any(p.get('ban') != code or p.get('name') != name for p in existing):
                    raise ValueError(f'{date_str} 主管 {name}({code}) 與當日名冊姓名或代號衝突')
                if len(existing) > 1:
                    raise ValueError(f'{date_str} 主管 {name}({code}) 重複出現')
                if existing:
                    existing[0]['status'] = status
                    day_data['on_duty'] = [p for p in day_data['on_duty']
                                           if p is not existing[0]]
                    day_data['off_duty'] = [p for p in day_data['off_duty']
                                            if p is not existing[0]]
                    day_data['off_duty'].append(existing[0])
                else:
                    day_data['off_duty'].append(person)

    # ── 全隊大表：取締酒駕 ──────────────────────────────
    if '全隊大表' in wb.sheetnames:
        ws_all = wb['全隊大表']
        for row in range(1, ws_all.max_row+1):
            date_val = ws_all.cell(row=row, column=1).value
            ao_val = ws_all.cell(row=row, column=41).value
            # AO can contain only the hour range (e.g. 18-22); the
            # literal word 取締 is not required. Date may be a day number.
            sake_time = parse_sake_time(ao_val)
            if not date_val or not sake_time:
                continue
            if isinstance(date_val, datetime):
                date_obj = date_val
            else:
                try:
                    date_obj = datetime.strptime(str(date_val)[:10], '%Y-%m-%d')
                except (TypeError, ValueError):
                    try:
                        day_number = int(date_val)
                    except (TypeError, ValueError):
                        continue
                    if 30000 <= day_number <= 80000:
                        date_obj = openpyxl.utils.datetime.from_excel(day_number, wb.epoch)
                    elif 1 <= day_number <= 31:
                        matches = [key for key in monthly_data if key.endswith(f'月{day_number}日')]
                        if len(matches) != 1:
                            continue
                        date_str = matches[0]
                        date_obj = None
                    else:
                        continue
            if date_obj is not None:
                date_str = f'{date_obj.year - 1911}年{date_obj.month}月{date_obj.day}日'
            if date_str in monthly_data:
                sake_by_date[date_str] = sake_time
                monthly_data[date_str]['has_sake'] = True
                monthly_data[date_str]['sake_times'] = [sake_time]

    return {
        'monthlyData': monthly_data,
        'staffFromExcel': list(staff_from_excel.keys()),
        'fixedRoster': fixed_roster,
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
        response = sb_post('monthly_duty', batch)
        if not isinstance(response, list) or len(response) != len(batch):
            raise RuntimeError(f'資料庫寫入未確認：第 {i // 50 + 1} 批，回應 {str(response)[:300]}')
    return len(records)

def sync_staff(staff_names):
    """把大表人員同步到 Supabase staff"""
    existing = sb_get('staff', params={'type': 'eq.輪班', 'order': 'code'})
    existing_names = {s['name'] for s in existing if s.get('name')}
    new_names = [n for n in staff_names if n not in existing_names]
    return new_names  # 回傳新人員名單讓前端確認

def sync_fixed_roster(fixed_roster):
    """Synchronize only fixed positions explicitly numbered in the uploaded Excel."""
    if not fixed_roster:
        raise RuntimeError('全隊大表沒有可核對的固定番號名冊')
    existing = sb_get('staff', params={'select': 'code,name', 'order': 'code'})
    if not isinstance(existing, list):
        raise RuntimeError('資料庫名冊讀取失敗，拒絕同步')
    by_code = {int(item['code']): item for item in existing if str(item.get('code', '')).isdigit()}
    changed = []
    for code, name in fixed_roster.items():
        code = int(code)
        if code not in by_code:
            raise RuntimeError(f'資料庫缺少固定番號 {code}，停止同步')
        if str(by_code[code].get('name') or '').strip() != name:
            response = sb_patch('staff', 'code', code, {'name': name})
            if not isinstance(response, list) or len(response) != 1:
                raise RuntimeError(f'固定番號 {code} 更新未確認')
            changed.append(code)
    return changed

def generate_word(data):
    on_duty  = data['onDuty']
    off_duty = data['offDuty']
    roc_date = data['rocDate']
    weekday  = data['weekday']
    is_holiday = data['isHoliday']
    has_sake   = data.get('hasSake', False)
    sake_times = data.get('sakeTimes', [])

    tmpl_path = get_tmpl(is_holiday, has_sake)
    # Supabase is the authoritative template store. A local bundled copy is optional.
    # template_bytes() handles the Storage lookup and local fallback.

    # Column mapping is resolved after loading the document.

    lx, bx, gg = [], [], []
    # All numbered personnel, including internal staff and supervisors, belong
    # in the daily right-hand leave directory. Never reuse yesterday's entries.
    seen_leave_codes = set()
    for person in sorted(off_duty, key=lambda x: x.get('ban') if isinstance(x.get('ban'), int) else 999):
        ban = person.get('ban')
        if not isinstance(ban, int):
            continue  # Unnumbered new personnel must not receive invented codes.
        if ban in seen_leave_codes:
            raise ValueError(f'{roc_date} 右側假別代號 {ban} 重複')
        seen_leave_codes.add(ban)
        st = str(person.get('status') or '').strip()
        if '停休' in st:
            continue
        if st in ('輪', '輪休'):
            lx.append(person)
        elif st in ('公', '公假') or '公假' in st:
            gg.append(person)
        elif st in ('優', '休', '補', '休假', '補休') or '假' in st:
            bx.append(person)

    doc = Document(io.BytesIO(template_bytes(tmpl_path)))
    t = doc.tables[0]
    # Column positions depend on the actual approved template, not the day type.
    # Both 35-column and 37-column approved layouts are supported.
    CODE_COL, NAME_COL, NOTE_COL, OCODE_COL, ONAME_COL = (len(t.columns)-5, len(t.columns)-4, len(t.columns)-3, len(t.columns)-2, len(t.columns)-1)
    # Resolve each complex merged row only once instead of dozens of times.
    row_cells = [row.cells for row in t.rows]

    ban_to_row = {}
    for r_idx in range(4, len(t.rows)):
        if CODE_COL >= len(row_cells[r_idx]): continue
        ct = row_cells[r_idx][CODE_COL].text.strip()
        if ct.isdigit() and 1 <= int(ct) <= 63:
            ban_to_row[int(ct)] = r_idx

    # The approved layout leaves the secretary row blank: populate code 60
    # in that existing row without changing the original stored template.
    if 60 not in ban_to_row and 61 in ban_to_row:
        secretary_row = ban_to_row[61] - 1
        if secretary_row >= 4 and not row_cells[secretary_row][CODE_COL].text.strip():
            set_cell(row_cells[secretary_row][CODE_COL], '60')
            ban_to_row[60] = secretary_row

    roster = sb_get('staff', params={'select': 'code,name', 'order': 'code'})
    if not isinstance(roster, list) or not roster:
        raise RuntimeError('人員名冊讀取失敗')
    names_by_code = {}
    for entry in roster:
        try:
            code = int(entry.get('code'))
        except (TypeError, ValueError):
            continue
        if 1 <= code <= 63:
            names_by_code[code] = str(entry.get('name') or '').strip()
    # External personnel use the Excel daily number, which rotates.
    daily_external = {}
    for person in list(on_duty) + list(off_duty):
        code = person.get('ban')
        if isinstance(code, int) and 1 <= code <= 20 and person.get('name'):
            daily_external[code] = str(person['name']).strip()
    for code, row_index in ban_to_row.items():
        name_cell = row_cells[row_index][NAME_COL]
        name = daily_external.get(code, '') if code <= 20 else names_by_code.get(code, '')
        if name_cell.text.strip() != name:
            set_cell(name_cell, name)

    # Keep existing names and formatting. Only replace a name when the Excel
    # record explicitly supplies a valid numbered staff member.
    for person in list(on_duty) + list(off_duty):
        ban = person.get('ban')
        name = person.get('name')
        if not isinstance(ban, int) or ban not in ban_to_row or not name:
            continue
        row = row_cells[ban_to_row[ban]]
        # Excel is authoritative for numbered staff names; Word is layout only.
        if ban >= 21 and names_by_code.get(ban) != name:
            continue
        status = person.get('status') or ''
        # Do not overwrite permanent remarks such as 內部管理 on ordinary days.
        if isinstance(ban, int) and ban >= 21:
            leave_status = {'輪': '輪休', '休': '休假', '優': '休假',
                            '補': '休假', '公': '公假'}.get(status, status)
            # Preserve Excel's full leave text, including partial-day and mixed leave.
            # E.g. 8-12公假 13-17休假 must not be reduced to a single category.
            # For mixed leave types, display public leave first as the priority.
            if '公假' in leave_status:
                display_leave = '公假'
            elif '輪休' in leave_status:
                display_leave = '輪休'
            elif '休假' in leave_status:
                display_leave = '休假'
            elif '假' in leave_status and '停休' not in leave_status:
                display_leave = leave_status
            else:
                display_leave = ''
            if display_leave and row[NOTE_COL].text.strip() != display_leave:
                set_cell(row[NOTE_COL], display_leave)
        elif status and row[NOTE_COL].text.strip() != status:
            set_cell(row[NOTE_COL], status)

    lx_rows = find_section_rows(t, ONAME_COL, '輪   休', row_cells)
    bx_rows = find_section_rows(t, ONAME_COL, '休(補)假', row_cells)
    gg_rows = find_section_rows(t, ONAME_COL, '公假', row_cells)

    for rows in [lx_rows, bx_rows, gg_rows]:
        for r_idx in rows:
            set_cell_preserving_format(row_cells[r_idx][OCODE_COL], '')
            set_cell_preserving_format(row_cells[r_idx][ONAME_COL], '')

    # Never silently drop leave entries if a template section has too few rows.
    for label, people, rows in [('輪休', lx, lx_rows), ('休(補)假', bx, bx_rows), ('公假', gg, gg_rows)]:
        if not rows:
            raise ValueError(f'{roc_date} Word範本找不到「{label}」欄位')
        if len(people) > len(rows):
            raise ValueError(f'{roc_date}「{label}」共{len(people)}人，但範本只有{len(rows)}格，禁止漏列')
        for j, person in enumerate(people):
            set_cell_preserving_format(row_cells[rows[j]][OCODE_COL], str(person['ban']) if isinstance(person.get('ban'), int) else '')
            set_cell_preserving_format(row_cells[rows[j]][ONAME_COL], person['name'])

    day = roc_date.split('年')[1].split('日')[0] + '日'
    for cell in row_cells[0]:
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
            cells = row_cells[r_idx]
            if not cells: continue
            cell_txt = cells[0].text.strip()
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
                    # Do not draw borders across the personnel directory.
                    for cell in row_cells[time_to_row[h]][:CODE_COL]:
                        set_red_border_cells([cell])

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

def validate_monthly_data_before_write(monthly_data):
    """Fail closed before any roster or daily schedule is written to Supabase."""
    if not isinstance(monthly_data, dict) or not monthly_data:
        raise ValueError('Excel 沒有可用的每日勤務資料，未更新資料庫')
    for date_str, day in monthly_data.items():
        if not isinstance(day, dict) or day.get('date_str') != date_str:
            raise ValueError(f'{date_str} 日期資料不一致')
        people = list(day.get('on_duty') or []) + list(day.get('off_duty') or [])
        if not people:
            raise ValueError(f'{date_str} 沒有人員資料')
        seen_codes = set()
        seen_names = set()
        for person in people:
            if not isinstance(person, dict):
                raise ValueError(f'{date_str} 人員資料格式異常')
            name = str(person.get('name') or '').strip()
            if not name:
                raise ValueError(f'{date_str} 發現缺少姓名的人員')
            if name in seen_names:
                raise ValueError(f'{date_str} 姓名 {name} 重複出現，停止產檔')
            seen_names.add(name)
            code = person.get('ban')
            if code is None:
                continue  # Unnumbered newcomers stay unnumbered.
            if type(code) is not int or not 1 <= code <= 99:
                raise ValueError(f'{date_str} 人員代號異常：{code}')
            if code in seen_codes:
                raise ValueError(f'{date_str} 人員代號 {code} 重複，停止產檔')
            seen_codes.add(code)
            verified_supervisors = {
                35: '姚宗志', 43: '翁水成', 60: '李安景',
                61: '陳彥文', 62: '林志明', 63: '張瑞忠',
            }
            if code in verified_supervisors and name != verified_supervisors[code]:
                raise ValueError(f'{date_str} 主管代號 {code} 與姓名 {name} 不符')
        match = re.fullmatch(r'(\\d+)年(\\d+)月(\\d+)日', date_str)
        if not match:
            raise ValueError(f'{date_str} 日期格式異常')
        roc_year, month, day_number = map(int, match.groups())
        try:
            date_obj = datetime(roc_year + 1911, month, day_number)
        except ValueError as exc:
            raise ValueError(f'{date_str} 日期不存在') from exc
        expected_weekday = ['一', '二', '三', '四', '五', '六', '日'][date_obj.weekday()]
        if day.get('weekday') != expected_weekday or not isinstance(day.get('is_holiday'), bool):
            raise ValueError(f'{date_str} 星期或假日欄位異常')


@app.route('/api/parse-excel', methods=['POST'])
def parse_excel_api():
    f = request.files.get('file')
    if not f: return jsonify({'error': '未上傳檔案'}), 400
    try:
        excel_bytes = f.read(12 * 1024 * 1024 + 1)
        if len(excel_bytes) > 12 * 1024 * 1024:
            raise ValueError('Excel 超過 12MB 上限')
        if not zipfile.is_zipfile(io.BytesIO(excel_bytes)):
            raise ValueError('Excel 不是有效的 xlsx 檔案')
        result = parse_excel(excel_bytes)
        validate_monthly_data_before_write(result['monthlyData'])
        # The official generator must reread original Excel on every run.
        # Do not publish parsed rows if the authoritative source cannot be
        # saved in the private bucket first.
        if not SUPABASE_URL or not SUPABASE_KEY:
            raise RuntimeError('無法永久儲存原始 Excel，停止匯入')
        token = os.environ.get('TEMPLATE_UPLOAD_TOKEN', '')
        if not token or not hmac.compare_digest(request.headers.get('X-Template-Upload-Token', ''), token):
            raise PermissionError('上傳原始全隊大表需要管理授權')
        dates = list(result['monthlyData'])
        months = {tuple(map(int, m.groups()[:2])) for date in dates
                  if (m := re.fullmatch(r'(\\d+)年(\\d+)月(\\d+)日', date))}
        if len(months) != 1 or len(dates) != len(result['monthlyData']):
            raise ValueError('Excel 日期跨月或格式異常，停止匯入')
        roc_year, month = next(iter(months))
        object_name = f'source_excel/{roc_year + 1911}-{month:02d}.xlsx'
        url = f'{SUPABASE_URL}/storage/v1/object/{TEMPLATE_BUCKET}/{object_name}'
        headers = {
            'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'x-upsert': 'true'
        }
        uploaded = req.post(url, headers=headers, data=excel_bytes, timeout=40)
        if uploaded.status_code not in (200, 201):
            raise RuntimeError(f'原始 Excel 儲存失敗 HTTP {uploaded.status_code}，停止匯入')
        source_check = req.get(
            f'{SUPABASE_URL}/storage/v1/object/authenticated/{TEMPLATE_BUCKET}/{object_name}',
            headers={'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}'},
            timeout=40
        )
        if source_check.status_code != 200 or source_check.content != excel_bytes:
            raise RuntimeError('原始 Excel 回讀驗證失敗，停止匯入')
        # Sync authoritative fixed roster before publishing daily schedules.
        changed_codes = sync_fixed_roster(result['fixedRoster'])
        result['rosterUpdatedCodes'] = changed_codes
        # 自動存到 Supabase
        saved = save_to_supabase(result['monthlyData'])
        result['savedDays'] = saved
        # Report the actual parsed supervisor leave rows for audit.
        result['supervisorAudit'] = {
            date: [p for p in day.get('off_duty', [])
                   if p.get('ban') in (61, 62, 63)]
            for date, day in result['monthlyData'].items()
            if any(p.get('ban') in (61, 62, 63)
                   for p in day.get('off_duty', []))
        }
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

@app.route('/api/duty-month/latest', methods=['GET'])
def get_latest_duty_month():
    """Return the most recently uploaded duty month from persisted Supabase rows."""
    if not SUPABASE_URL:
        return jsonify({'year': None, 'month': None, 'days': []})
    recent = sb_get('monthly_duty', params={'select': 'date_str', 'order': 'created_at.desc,id.desc', 'limit': '1'})
    if not recent:
        return jsonify({'year': None, 'month': None, 'days': []})
    match = re.match(r'^([0-9]+)年([0-9]+)月[0-9]+日' + chr(36), recent[0].get('date_str', ''))
    if not match:
        return jsonify({'error': '儲存的勤務日期格式不正確'}), 500
    year, month = map(int, match.groups())
    rows = sb_get('monthly_duty', params={'date_str': f'like.{year}年{month}月%', 'order': 'id'})
    return jsonify({'year': year, 'month': month, 'days': rows})

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

def verified_generation_input(payload):
    """Reparse the authoritative Excel on every generation request.

    This checks source consistency only. Visual DOCX rendering verification is
    still missing, so official output remains blocked until that is implemented.
    """
    if not isinstance(payload, dict):
        raise ValueError('生成資料格式不正確')
    date_str = payload.get('rocDate')
    match = re.fullmatch(r'(\d+)年(\d+)月(\d+)日', str(date_str or ''))
    if not match:
        raise ValueError('缺少有效的勤務日期')
    roc_year, month, _ = map(int, match.groups())
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError('無法讀取原始 Excel')
    object_name = f'source_excel/{roc_year + 1911}-{month:02d}.xlsx'
    source = req.get(
        f'{SUPABASE_URL}/storage/v1/object/authenticated/{TEMPLATE_BUCKET}/{object_name}',
        headers={'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}'},
        timeout=40
    )
    if source.status_code != 200:
        raise RuntimeError(f'{date_str} 原始 Excel 不存在或讀取失敗，HTTP {source.status_code}')
    if len(source.content) > 12 * 1024 * 1024:
        raise ValueError('原始 Excel 超過 12MB')
    parsed = parse_excel(source.content)
    validate_monthly_data_before_write(parsed['monthlyData'])
    day = parsed['monthlyData'].get(date_str)
    if day is None:
        raise ValueError(f'{date_str} 不在原始 Excel 內')
    expected = {
        'rocDate': day['roc_date'], 'weekday': day['weekday'],
        'isHoliday': day['is_holiday'], 'hasSake': day['has_sake'],
        'sakeTimes': day['sake_times'], 'onDuty': day['on_duty'],
        'offDuty': day['off_duty'],
    }
    # Reject tampering or stale frontend data. The source Excel is authoritative.
    for key, value in expected.items():
        if payload.get(key) != value:
            raise ValueError(f'{date_str} {key} 與當次原始 Excel 解析不一致')
    raise RuntimeError('當次 Excel 重新解析已完成，但 Word 逐頁渲染與完整人員驗證尚未完成，禁止正式輸出')


@app.route('/api/generate', methods=['POST'])
def generate_api():
    data = request.json
    try:
        verified_generation_input(data)
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
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for d in days_data:
            try:
                verified_generation_input(d)
                buf = generate_word(d)
                fname = f"勤務表_{d.get('rocDate','')}.docx"
                zf.writestr(fname, buf.read())
            except Exception as e:
                # Never return a partial ZIP that silently omits failed dates.
                return jsonify({'error': '批次產生失敗', 'date': d.get('rocDate', ''), 'detail': str(e)}), 500
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
    storage_name = TEMPLATE_STORAGE_NAMES.get(name, name)
    url = f"{SUPABASE_URL}/storage/v1/object/{TEMPLATE_BUCKET}/{storage_name}"
    headers = {'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}',
               'Content-Type': DOCX_MIME, 'x-upsert': 'true',
               'Cache-Control': 'no-cache'}
    try:
        result = req.post(url, headers=headers, data=payload, timeout=30)
        if result.status_code not in (200, 201):
            return jsonify({'error': f'永久儲存失敗 HTTP {result.status_code}'}), 502
        saved = template_bytes(os.path.join(ASSETS, name))
        if saved != payload:
            return jsonify({'error': '儲存後內容驗證失敗'}), 502
    except Exception as exc:
        return jsonify({'error': f'範本儲存或驗證失敗：{exc}'}), 502
    return jsonify({'ok': True, 'type': tmpl_type, 'persistent': True})

if __name__ == '__main__':
    app.run(debug=True)
