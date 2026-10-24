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
            else:
                for tag in ['sz','szCs']:
                    el = rpr.find(f'{{{ns}}}{tag}')
                    if el is not None: rpr.remove(el)
            return
    cell.paragraphs[0].add_run(text)

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
    for c in range(1, ws_ext.max_column+1):
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
            ban_val = ws_ext.cell(row=row, column=col-1).value
            status_val = ws_ext.cell(row=row, column=col+1).value

            ban = None
            if ban_val is not None:
                ban_str = str(ban_val).strip()
                if ban_str.isdigit():
                    ban = int(ban_str)

            status = ''
            if status_val and str(status_val).strip():
                raw = str(status_val).strip()
                for s in ['輪','優','休','補','公']:
                    if s in raw:
                        status = s
                        break
                if not status:
                    status = raw[:2]  # 常訓等備考

            person = {'name': name, 'ban': ban, 'status': status}
            off_statuses = ['輪','優','休','補','公']
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
    # Supabase is the authoritative template store. A local bundled copy is optional.
    # template_bytes() handles the Storage lookup and local fallback.

    # Column mapping is resolved after loading the document.

    lx, bx, gg = [], [], []
    # Only field staff appear in the right-hand leave directory.
    # Internal staff (21+) and supervisors keep leave status in their own note.
    for s in sorted(off_duty, key=lambda x: x.get('ban') or 99):
        ban = s.get('ban')
        if not isinstance(ban, int) or ban > 20:
            continue
        st = s.get('status') or ''
        if st == '輪': lx.append(s)
        elif st in ['優','休','補']: bx.append(s)
        elif st == '公': gg.append(s)

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
        if ct.isdigit() and 1 <= int(ct) <= 20:
            ban_to_row[int(ct)] = r_idx

    # Keep existing names and formatting. Only replace a name when the Excel
    # record explicitly supplies a valid numbered staff member.
    for person in list(on_duty) + list(off_duty):
        ban = person.get('ban')
        name = person.get('name')
        if not isinstance(ban, int) or ban not in ban_to_row or not name:
            continue
        row = row_cells[ban_to_row[ban]]
        if row[NAME_COL].text.strip() != name:
            set_cell(row[NAME_COL], name)
        status = person.get('status') or ''
        if status and row[NOTE_COL].text.strip() != status:
            set_cell(row[NOTE_COL], status)

    lx_rows = find_section_rows(t, ONAME_COL, '輪   休', row_cells)
    bx_rows = find_section_rows(t, ONAME_COL, '休(補)假', row_cells)
    gg_rows = find_section_rows(t, ONAME_COL, '公假', row_cells)

    for rows in [lx_rows, bx_rows, gg_rows]:
        for r_idx in rows:
            set_cell(row_cells[r_idx][OCODE_COL], '')
            set_cell(row_cells[r_idx][ONAME_COL], '')

    for people, rows in [(lx, lx_rows), (bx, bx_rows), (gg, gg_rows)]:
        for j, s in enumerate(people):
            if j >= len(rows): break
            set_cell(row_cells[rows[j]][OCODE_COL], str(s['ban']) if isinstance(s.get('ban'), int) else '')
            set_cell(row_cells[rows[j]][ONAME_COL], s['name'])

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
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for d in days_data:
            try:
                buf = generate_word(d)
                fname = f"勤務表_{d.get('rocDate','')}.docx"
                zf.writestr(fname, buf.read())
            except:
                pass
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
