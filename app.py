from flask import Flask, request, jsonify, send_file, render_template, g, has_request_context
from word_validation import table_row_cells, validate_generated_word
from docx import Document
from lxml import etree
import openpyxl, json, io, os, re, copy, zipfile, requests as req, hmac, uuid, gc
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
    '平日取酒': 'tmpl_平日取酒.docx', '假日取酒': 'tmpl_假日取酒.docx',
    '平日取酒18–22': 'tmpl_平日取酒18-22.docx',
}
# Supabase Storage object keys use stable ASCII names; user-facing Word filenames remain Chinese.
TEMPLATE_STORAGE_NAMES = {
    'tmpl_平日.docx': 'weekday.docx',
    'tmpl_假日.docx': 'holiday.docx',
    'tmpl_平日取酒.docx': 'weekday_sake.docx',
    'tmpl_假日取酒.docx': 'holiday_sake.docx',
    'tmpl_平日取酒18-22.docx': 'weekday_sake_18_22.docx',
}
SAKE_TEMPLATE_TYPES = {
    (False, ((19, 23),)): '平日取酒',
    (False, ((18, 22),)): '平日取酒18–22',
    (True, ((18, 22),)): '假日取酒',
}
DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

def template_bytes(tmpl_path):
    cache = g.get('batch_template_bytes') if has_request_context() else None
    if cache is None:
        payload = _read_template_bytes(tmpl_path)
    else:
        if tmpl_path not in cache:
            cache[tmpl_path] = _read_template_bytes(tmpl_path)
        payload = cache[tmpl_path]
    if has_request_context():
        g.generation_template = (tmpl_path, payload)
    return payload


def _read_template_bytes(tmpl_path):
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
    try:
        with open(tmpl_path, 'rb') as f:
            return f.read()
    except FileNotFoundError as exc:
        raise RuntimeError(f'缺少 Word 範本：{storage_name}') from exc

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
    rows = table_row_cells(table)
    codes = {int(row[code_col].text.strip()) for row in rows
             if row[code_col].text.strip().isdigit()}
    if not {19, 29}.issubset(codes):
        raise ValueError('Word範本缺少正式代號19或29')
    # The approved weekday-sake template places its leave headings beside
    # codes 15/26; other approved templates use 19/29. Validate the actual
    # leave column, not a cross-column row coincidence.
    leave_labels = [row[name_col].text.strip() for row in rows]
    for label in ('輪   休', '休(補)假', '公假'):
        if sum(text == label for text in leave_labels) != 1:
            raise ValueError(f'Word範本必須保留唯一「{label}」標題')
    return doc

def get_tmpl(is_holiday, has_sake, sake_times=(), roc_date=''):
    if not has_sake:
        return os.path.join(ASSETS, TEMPLATE_NAMES['假日' if is_holiday else '平日'])
    times = tuple(sorted((t['start'], t['end']) for t in sake_times))
    template_type = SAKE_TEMPLATE_TYPES.get((is_holiday, times))
    if template_type is None:
        day_type = '假日' if is_holiday else '平日'
        hours = '、'.join(f'{start:02d}–{end:02d}' for start, end in times) or '未提供'
        raise ValueError(f'{roc_date}（{day_type}）沒有對應的已驗收取締酒駕範本（Excel：{hours}）')
    return os.path.join(ASSETS, TEMPLATE_NAMES[template_type])

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

def set_cell_preserving_format(cell, text, new_run_rpr=None):
    """Edit existing runs in place; inherit cell formatting for a new run."""
    text = str(text) if text is not None else ''
    if cell.text == text:
        return
    runs = [run for para in cell.paragraphs for run in para.runs]
    if runs:
        # Keep text in its original run slots, including differently styled
        # date fragments. Empty leading runs must not absorb styled text.
        occupied = [i for i, run in enumerate(runs) if run.text]
        last = occupied[-1] if occupied else 0
        remaining = text
        for i, run in enumerate(runs):
            length = len(run.text)
            value = remaining if i == last else remaining[:length] if i < last else ''
            remaining = remaining[len(value):]
            if run.text != value:
                run.text = value
    elif text:
        para = cell.paragraphs[0]
        run = para.add_run(text)
        paragraph_rpr = para._p.find(f'{{{ns}}}pPr/{{{ns}}}rPr')
        if paragraph_rpr is not None:
            run._r.insert(0, copy.deepcopy(paragraph_rpr))
        elif new_run_rpr is not None:
            run._r.insert(0, copy.deepcopy(new_run_rpr))
        # New text inherits the original cell's size/weight. Only the required
        # font family is explicit when the empty cell has no existing run.
        rpr = run._r.get_or_add_rPr()
        fonts = rpr.find(f'{{{ns}}}rFonts')
        if fonts is None:
            fonts = etree.SubElement(rpr, f'{{{ns}}}rFonts')
        for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
            fonts.set(f'{{{ns}}}{attr}', '標楷體')


def apply_directory_text_format(cell, points):
    """Apply the approved column format without replacing paragraphs or runs.

    Accepted 11/6 directory numbers/names use 12pt regular Kai. Use black
    explicitly so stale template blue/theme colors cannot follow new text.
    """
    from docx.shared import Pt, RGBColor
    for para in cell.paragraphs:
        for run in para.runs:
            run.font.name = '標楷體'
            run.font.size = Pt(points)
            run.bold = False
            run.font.color.rgb = RGBColor(0, 0, 0)
            rpr = run._r.get_or_add_rPr()
            rpr.rFonts.set(f'{{{ns}}}eastAsia', '標楷體')
            for tag, value in (('szCs', str(points * 2)), ('bCs', '0')):
                node = rpr.find(f'{{{ns}}}{tag}')
                if node is None:
                    node = etree.SubElement(rpr, f'{{{ns}}}{tag}')
                node.set(f'{{{ns}}}val', value)


def clear_daily_hand_entry_codes(row_cells, code_col):
    """Clear personnel numbers only in time rows left of the directory.

    Template vehicle IDs (e.g. 908, BEW-7085), road headings, time labels,
    and duty text are outside this predicate. Merged cells are edited once.
    """
    seen = set()
    for cells in row_cells:
        if not re.fullmatch(r'\d{1,2}[－\-–]\d{1,2}', cells[0].text.strip()):
            continue
        for cell in cells[1:code_col]:
            if cell._tc in seen:
                continue
            seen.add(cell._tc)
            text = cell.text.strip()
            numbers = re.findall(r'\d+', text)
            if (numbers and re.fullmatch(r'[\d\s()（）,，、+＋/／\-－]+', text)
                    and all(1 <= int(number) <= 63 for number in numbers)):
                set_cell_preserving_format(cell, '')


# Persistent annotation confirmed in all four approved templates and the
# accepted 11/6–11/8 documents; the Excel BB daily cells do not encode it.
# Match both name and code so a replacement employee cannot inherit it.
PERSISTENT_STAFF_REMARKS = {(23, '林筠臻'): '留職停薪'}


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

def is_leave_status(status):
    status = str(status or '')
    if any(x in status for x in ('停休', '△', 'Δ', '假日交督', '假日 交督', '留職停薪')):
        return False
    return status in ('輪', '輪休', '優', '休', '補', '公') or '假' in status


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

    # 原始 115 年 11 月 Excel：第 2 列為姓名，第 3 列起為每日資料
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
        elif isinstance(date_val, (int, float)) and 30000 <= date_val <= 80000:
            date_obj = openpyxl.utils.datetime.from_excel(date_val, wb.epoch)
        else:
            try:
                date_obj = datetime.strptime(str(date_val)[:10], '%Y-%m-%d')
            except (TypeError, ValueError):
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

            person = {'name': name, 'ban': ban, 'status': status,
                      'sourceSheet': '外勤',
                      'sourceColumn': openpyxl.utils.get_column_letter(col)}
            # Any kind of leave is off-duty; 停休 means leave was cancelled.
            off_statuses = ['輪','優','休','補','公']
            is_leave = is_leave_status(status)
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

    # 全隊大表 AW:BV 全部是內勤；AW:AZ 為 50–53，不是外勤 17–20。
    # 外勤 17–20 已由外勤工作表 AI:AO 的姓名、代號及每日狀態解析。
    fixed_roster = {}
    if '全隊大表' in wb.sheetnames:
        ws_all = wb['全隊大表']
        internal_cols = {}
        for col in range(49, 75):  # AW:BV; BI is an empty separator.
            name = str(ws_all.cell(2, col).value or '').strip()
            if 2 <= len(name) <= 5 and name not in ('日期', '星期'):
                internal_cols[col] = name
                staff_from_excel[name] = True

        # Approved Word template's fixed name/code directory. Never derive
        # codes from worksheet column position. Names absent from this directory
        # (e.g. newcomer 周彥璋) remain unnumbered.
        approved_internal_codes = {
            '林筠臻': 23, '黃雅琳': 26, '洪君智': 27,
            '楊智為': 28, '陳淑卿': 29, '謝函儒': 31,
            '曾義雄': 33, '姚宗志': 35, '柯育婷': 36,
            '林裕善': 38, '王庭偉': 39, '林家豪': 41,
            '張銥凌': 42, '翁水成': 43, '洪呈欣': 44,
            '簡竹芳': 45, '沈宗漢': 46, '曾煙煌': 47,
            '陳駿騰': 48, '李智偉': 50, '林新慶': 51,
            '李金孔': 52, '董安民': 53, '陳昱靜': 54,
        }
        unknown = sorted(set(internal_cols.values()) -
                         set(approved_internal_codes) - {'周彥璋'})
        if unknown:
            raise ValueError('全隊大表有未核對代號的人員：' + '、'.join(unknown))
        # A new hire without an assigned code is retained for status
        # auditing, but must never receive an invented number.
        for row in range(3, ws_all.max_row + 1):
            raw_date = ws_all.cell(row, 47).value
            if isinstance(raw_date, datetime):
                d = raw_date
            elif isinstance(raw_date, (int, float)) and 30000 <= raw_date <= 80000:
                d = openpyxl.utils.datetime.from_excel(raw_date, wb.epoch)
            else:
                continue
            key = f'{d.year-1911}年{d.month}月{d.day}日'
            if key not in monthly_data:
                continue
            day_data = monthly_data[key]
            for col, name in internal_cols.items():
                if any(p.get('name') == name for p in
                       day_data['on_duty'] + day_data['off_duty']):
                    raise ValueError(f'{key} {name} 在外勤與全隊大表重複，需核對來源')
                raw = str(ws_all.cell(row, col).value or '').strip()
                if '停休' in raw or '△' in raw or 'Δ' in raw:
                    status = '停休'
                elif raw in ('輪', '輪休'):
                    status = '輪休'
                elif raw in ('休', '補', '優') or '休假' in raw or '補休' in raw or raw.endswith('休'):
                    status = '休假'
                elif raw in ('公', '公假') or '公假' in raw:
                    status = '公假'
                else:
                    status = raw
                person = {'name': name, 'ban': approved_internal_codes.get(name),
                          'status': status, 'sourceSheet': '全隊大表',
                          'sourceColumn': openpyxl.utils.get_column_letter(col)}
                if is_leave_status(status):
                    day_data['off_duty'].append(person)
                else:
                    day_data['on_duty'].append(person)

    # ── 全隊大表：主管督勤欄（依職稱，不依外勤姓名） ──
    # BQ:BS are daily-status cells, not a fixed code/name directory.
    supervisor_titles = ('隊長', '林副隊長', '陳副隊長', '李秘書', '姚組長', '翁組長')
    if '全隊大表' in wb.sheetnames:
        ws_all = wb['全隊大表']
        # BQ:BS contain named staff's daily statuses, not a code/name
        # directory. Do not derive supervisor codes from these cells.
        fixed_staff = {}
        # Resolve supervisor titles through their Excel roster names.
        # Titles without a verified name-to-code mapping are skipped safely.
        supervisor_names = {
            '隊長': '張瑞忠', '林副隊長': '林志明',
            '陳副隊長': '陳彥文', '李秘書': '李安景',
            '姚組長': '姚宗志', '翁組長': '翁水成',
        }
        # Headers may be spread across the top-right duty area. Use one
        # verified header row only; never mix columns from different rows.
        supervisor_cols = {}
        supervisor_header_row = None
        for header_row in range(1, min(ws_all.max_row, 50) + 1):
            found = {}
            for col in range(1, ws_all.max_column + 1):
                label = re.sub(r'\s+', '', str(ws_all.cell(header_row, col).value or ''))
                if label == '長隊':
                    label = '隊長'
                if label in supervisor_titles:
                    if label in found:
                        raise ValueError(f'主管表頭 {label} 在第 {header_row} 列重複')
                    found[label] = col
            if len(found) >= 4:
                supervisor_cols = found
                supervisor_header_row = header_row
                break
        if not supervisor_cols:
            raise ValueError('全隊大表未找到可核對的主管督勤表頭，拒絕略過主管資料')
        for row in range((supervisor_header_row or ws_all.max_row) + 1, ws_all.max_row + 1):
            # 督勤區日期在 A 欄；AU 是上半部內勤區日期。
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
                if '停休' in raw or '△' in raw or 'Δ' in raw:
                    status = '停休'
                elif '公假' in raw or '公休' in raw or raw == '公':
                    status = '公假'
                elif '輪休' in raw or raw == '輪':
                    status = '輪休'
                elif '休假' in raw or '補休' in raw or raw in ('休', '優', '補'):
                    status = '休假'
                elif '假' in raw:
                    status = raw
                if not status:
                    status = raw
                person = {'name': name, 'ban': code, 'status': status,
                          'remarks': raw, 'sourceSheet': '全隊大表督勤',
                          'sourceColumn': openpyxl.utils.get_column_letter(col),
                          'sourceRow': row}
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
                    if not raw:
                        continue  # 空白督勤不覆寫內勤原有狀態。
                    existing[0]['status'] = status
                    existing[0]['supervisorSource'] = dict(person)
                    day_data['on_duty'] = [p for p in day_data['on_duty']
                                           if p is not existing[0]]
                    day_data['off_duty'] = [p for p in day_data['off_duty']
                                            if p is not existing[0]]
                    target = 'off_duty' if is_leave_status(status) else 'on_duty'
                    day_data[target].append(existing[0])
                else:
                    target = 'off_duty' if is_leave_status(status) else 'on_duty'
                    day_data[target].append(person)

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

    tmpl_path = get_tmpl(is_holiday, has_sake, sake_times, roc_date)
    # Supabase is the authoritative template store. A local bundled copy is optional.
    # template_bytes() handles the Storage lookup and local fallback.

    # Column mapping is resolved after loading the document.

    lx, bx, gg = [], [], []
    # Right-hand rotating leave is code-based: 1–20, 50–53, 60–63 only.
    # Other leave categories and every person's own remarks are unchanged.
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
            if 1 <= ban <= 20 or 50 <= ban <= 53 or 60 <= ban <= 63:
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
    row_cells = table_row_cells(t)

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
            set_cell_preserving_format(row_cells[secretary_row][CODE_COL], '60')
            ban_to_row[60] = secretary_row

    daily_people = {p['ban']: p for p in list(on_duty) + list(off_duty)
                    if isinstance(p.get('ban'), int)}
    for code, row_index in ban_to_row.items():
        row = row_cells[row_index]
        person = daily_people.get(code, {})
        name = str(person.get('name') or '')
        set_cell_preserving_format(row[NAME_COL], name)
        status = str(person.get('status') or '').strip()
        persistent = PERSISTENT_STAFF_REMARKS.get((code, name), '')
        daily_note = str(person.get('remarks') or
                         person.get('supervisorSource', {}).get('remarks') or status).strip()
        note = ' '.join(dict.fromkeys(value for value in (persistent, daily_note) if value))
        # Empty daily status is a write too: remove yesterday's remarks.
        set_cell_preserving_format(row[NOTE_COL], note)
        for column in (CODE_COL, NAME_COL):
            apply_directory_text_format(row[column], 12)
        # Read from accepted 11/6 Word: ordinary remarks inherit Normal
        # 12pt; 留職停薪 explicitly uses 9pt, supervisor remarks 8pt.
        # Keep the user's black-text correction; do not reuse template blue.
        remark_points = {23: 9, 35: 8, 43: 8, 60: 8, 61: 8, 62: 8, 63: 8}.get(code, 12)
        apply_directory_text_format(row[NOTE_COL], remark_points)

    clear_daily_hand_entry_codes(row_cells, CODE_COL)

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
            apply_directory_text_format(row_cells[rows[j]][OCODE_COL], 12)
            apply_directory_text_format(row_cells[rows[j]][ONAME_COL], 12)

    date_match = re.fullmatch(r'(\d+)年(\d+)月(\d+)日', roc_date)
    if not date_match:
        raise ValueError('勤務日期格式異常')
    year, month, day = map(int, date_match.groups())
    date_obj = datetime(year + 1911, month, day)
    expected_weekday = '一二三四五六日'[date_obj.weekday()]
    if weekday != expected_weekday:
        raise ValueError('勤務日期與星期不一致')
    date_cells = {cell._tc: cell for cell in row_cells[0]
                  if '民國' in cell.text and '年' in cell.text}
    if len(date_cells) != 1:
        raise ValueError('Word範本無法定位唯一日期欄')
    date_cell = next(iter(date_cells.values()))
    set_cell_preserving_format(
        date_cell, f'中華民國{year}年{month}月{day}日【星期{weekday}】')

    # Approved templates own the duty-area borders, including the sake frame.
    # Painting four sides of every time-row cell destroys that frame and adds
    # borders to shared vertical cells containing the fixed blue annotation.
    # Keep every physical cell's original borders, widths and merges intact.

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
    if not staff_admin_authorized():
        return jsonify({'error': '修改人員資料需要有效的管理授權'}), 403
    return jsonify(sb_post('staff', request.json))

@app.route('/api/staff/<int:code>', methods=['PUT'])
def update_staff(code):
    if not staff_admin_authorized():
        return jsonify({'error': '修改人員資料需要有效的管理授權'}), 403
    return jsonify(sb_patch('staff', 'code', code, request.json))

@app.route('/api/staff/<int:code>', methods=['DELETE'])
def delete_staff(code):
    if not staff_admin_authorized():
        return jsonify({'error': '修改人員資料需要有效的管理授權'}), 403
    sb_delete('staff', 'code', code)
    return jsonify({'ok': True})

def staff_admin_authorized():
    token = os.environ.get('TEMPLATE_UPLOAD_TOKEN', '')
    supplied = request.headers.get('X-Template-Upload-Token', '')
    return bool(token) and hmac.compare_digest(supplied.encode('utf-8'), token.encode('utf-8'))

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
        # The 20 external names must be present exactly once each day.
        # Four designated external staff are reconciled, never appended.
        external_codes = [p.get('ban') for p in people
                          if p.get('sourceSheet') == '外勤']
        if external_codes:
            if len(external_codes) != 20:
                raise ValueError(f'{date_str} 外勤應有 20 人，實際 {len(external_codes)} 人')
            if sorted(c for c in external_codes if isinstance(c, int)) != list(range(1, 21)):
                raise ValueError(f'{date_str} 外勤輪動及固定代號不完整')
        internal_people = [p for p in people
                           if p.get('sourceSheet') == '全隊大表']
        if internal_people and len(internal_people) != 25:
            raise ValueError(
                f'{date_str} 全隊大表 AW:BV 應有 25 人，實際 {len(internal_people)} 人')
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
        match = re.fullmatch(r'(\d+)年(\d+)月(\d+)日', date_str)
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
    if not staff_admin_authorized():
        return jsonify({'error': '上傳原始全隊大表需要管理授權'}), 403
    f = request.files.get('file')
    if not f: return jsonify({'error': '未上傳檔案'}), 400
    source_saved = False
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
        dates = list(result['monthlyData'])
        months = {tuple(map(int, m.groups()[:2])) for date in dates
                  if (m := re.fullmatch(r'(\d+)年(\d+)月(\d+)日', date))}
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
        source_saved = True
        source_check = req.get(
            f'{SUPABASE_URL}/storage/v1/object/authenticated/{TEMPLATE_BUCKET}/{object_name}',
            headers={'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}'},
            timeout=40
        )
        if source_check.status_code != 200 or source_check.content != excel_bytes:
            raise RuntimeError('原始 Excel 回讀驗證失敗，停止匯入')
        # The uploaded workbook has no separate verified fixed-code
        # directory at BQ:BS. Do not mutate the staff table using a
        # fabricated roster. The daily parse retains known codes only.
        if result['fixedRoster']:
            changed_codes = sync_fixed_roster(result['fixedRoster'])
        else:
            changed_codes = []
        result['rosterUpdatedCodes'] = changed_codes
        # 自動存到 Supabase
        saved = save_to_supabase(result['monthlyData'])
        result['savedDays'] = saved
        result['year'] = roc_year
        result['month'] = month
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
    except PermissionError as e:
        return jsonify({'error': str(e)}), 403
    except Exception as e:
        if source_saved:
            return jsonify({'error': '原始 Excel 已保存，但匯入未完成；請勿使用舊月資料，請重新匯入相同 Excel。',
                            'detail': str(e), 'sourceSaved': True,
                            'persistenceState': 'incomplete', 'needsReimport': True}), 500
        return jsonify({'error': str(e)}), 400 if isinstance(e, ValueError) else 500

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
    """Verify a day against the source Excel snapshot for this request."""
    if not isinstance(payload, dict):
        raise ValueError('生成資料格式不正確')
    date_str = payload.get('rocDate')
    match = re.fullmatch(r'(\d+)年(\d+)月(\d+)日', str(date_str or ''))
    if not match:
        raise ValueError('缺少有效的勤務日期')
    roc_year, month, _ = map(int, match.groups())
    object_name = f'source_excel/{roc_year + 1911}-{month:02d}.xlsx'
    snapshot = g.get('batch_source_month') if has_request_context() else None
    if snapshot is not None and snapshot[0] == object_name:
        monthly_data = snapshot[1]
    else:
        monthly_data = read_generation_month(object_name, date_str)
        if snapshot is not None:
            g.batch_source_month = (object_name, monthly_data)
    day = monthly_data.get(date_str)
    if day is None:
        raise ValueError(f'{date_str} 不在原始 Excel 內')
    expected = {
        'rocDate': day['roc_date'], 'weekday': day['weekday'],
        'isHoliday': day['is_holiday'], 'hasSake': day['has_sake'],
        'sakeTimes': day['sake_times'], 'onDuty': day['on_duty'],
        'offDuty': day['off_duty'],
    }
    for timing in expected['sakeTimes']:
        if (not isinstance(timing, dict) or type(timing.get('start')) is not int
                or type(timing.get('end')) is not int
                or not 0 <= timing['start'] <= 23 or not 0 <= timing['end'] <= 24
                or timing['start'] % 24 == timing['end'] % 24):
            raise ValueError(f'{date_str} 原始 Excel 取酒時段異常')
    # Reject tampering or stale frontend data. The source Excel is authoritative.
    for key, value in expected.items():
        if payload.get(key) != value:
            raise ValueError(f'{date_str} {key} 與當次原始 Excel 解析不一致')
    return expected


def read_generation_month(object_name, date_str):
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError('無法讀取原始 Excel')
    source = req.get(
        f'{SUPABASE_URL}/storage/v1/object/authenticated/{TEMPLATE_BUCKET}/{object_name}',
        headers={'apikey': SUPABASE_KEY, 'Authorization': f'Bearer {SUPABASE_KEY}',
                 'Cache-Control': 'no-cache'},
        params={'v': uuid.uuid4().hex}, timeout=40
    )
    if source.status_code != 200:
        raise RuntimeError(f'{date_str} 原始 Excel 不存在或讀取失敗，HTTP {source.status_code}')
    if len(source.content) > 12 * 1024 * 1024:
        raise ValueError('原始 Excel 超過 12MB')
    monthly_data = parse_excel(source.content)['monthlyData']
    validate_monthly_data_before_write(monthly_data)
    return monthly_data


def generate_checked_word(data):
    g.generation_template = None
    buf = generate_word(data)
    snapshot = g.generation_template
    if snapshot is None or snapshot[0] != get_tmpl(
            data['isHoliday'], data['hasSake'], data['sakeTimes'], data['rocDate']):
        raise ValueError('Word 驗證失敗：缺少當次模板來源')
    validate_generated_word(buf.getvalue(), data, snapshot[1],
                            PERSISTENT_STAFF_REMARKS, find_section_rows)
    buf.seek(0)
    return buf


@app.route('/api/generate', methods=['POST'])
def generate_api():
    data = request.json
    try:
        data = verified_generation_input(data)
        buf = generate_checked_word(data)
        fname = f'勤務表_{data.get("rocDate","")}.docx'
        return send_file(buf, as_attachment=True, download_name=fname,
                        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate-batch', methods=['POST'])
def generate_batch():
    days_data = request.json
    if not isinstance(days_data, list) or not days_data or any(not isinstance(d, dict) for d in days_data):
        return jsonify({'error': '請提供有效的勤務日期清單'}), 400
    # Use one source-month snapshot and one copy of each template per batch.
    g.batch_source_month = (None, None)
    g.batch_template_bytes = {}
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for d in days_data:
            try:
                d = verified_generation_input(d)
                buf = generate_checked_word(d)
                fname = f"勤務表_{d.get('rocDate','')}.docx"
                zf.writestr(fname, buf.read())
                # In long batches, release cyclic Word objects before the next document.
                if len(days_data) > 5:
                    gc.collect()
            except Exception as e:
                # Never return a partial ZIP that silently omits failed dates.
                return jsonify({'error': '批次產生失敗', 'date': d.get('rocDate', ''), 'detail': str(e)}), 500
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True, download_name='勤務表批次.zip',
                    mimetype='application/zip')

@app.route('/api/upload-template', methods=['POST'])
def upload_template():
    if not staff_admin_authorized():
        return jsonify({'error': '範本上傳未授權，請輸入管理密碼'}), 403
    tmpl_type = request.form.get('type')
    f = request.files.get('file')
    if not f or tmpl_type not in TEMPLATE_NAMES:
        return jsonify({'error': '請選擇正確的範本類型與Word檔案'}), 400
    if not SUPABASE_URL or not SUPABASE_KEY:
        return jsonify({'error': '尚未設定永久儲存服務，未執行上傳'}), 503
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
