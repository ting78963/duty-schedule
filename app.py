from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from lxml import etree
import openpyxl, json, io, os, re, copy, zipfile, requests as req
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
    if not os.path.exists(tmpl_path):
        raise FileNotFoundError(f'模板不存在: {tmpl_path}，請先上傳模板')

    CODE_COL, NAME_COL, NOTE_COL, OCODE_COL, ONAME_COL = get_cols(is_holiday)

    lx, bx, gg = [], [], []
    for s in sorted(off_duty, key=lambda x: x.get('ban') or 99):
        st = s.get('status') or ''
        if st == '輪': lx.append(s)
        elif st in ['優','休','補']: bx.append(s)
        elif st == '公': gg.append(s)

    doc = Document(tmpl_path)
    t = doc.tables[0]

    ban_to_row = {}
    for r_idx in range(4, len(t.rows)):
        if CODE_COL >= len(t.rows[r_idx].cells): continue
        ct = t.rows[r_idx].cells[CODE_COL].text.strip()
        if ct.isdigit() and 1 <= int(ct) <= 20:
            ban_to_row[int(ct)] = r_idx

    for ban, r_idx in ban_to_row.items():
        for ci in [NAME_COL, NOTE_COL]:
            for para in t.rows[r_idx].cells[ci].paragraphs:
                for run in para.runs: run.text = ''

    for s in on_duty:
        ban = s.get('ban')
        if not ban or ban not in ban_to_row: continue
        set_cell(t.rows[ban_to_row[ban]].cells[NAME_COL], s['name'])
        set_cell(t.rows[ban_to_row[ban]].cells[NOTE_COL], s.get('status') or '', size=24)

    for s in off_duty:
        ban = s.get('ban')
        if not ban or ban not in ban_to_row: continue
        set_cell(t.rows[ban_to_row[ban]].cells[NAME_COL], s['name'])
        set_cell(t.rows[ban_to_row[ban]].cells[NOTE_COL], s.get('status') or '', size=24)

    lx_rows = find_section_rows(t, ONAME_COL, '輪   休')
    bx_rows = find_section_rows(t, ONAME_COL, '休(補)假')
    gg_rows = find_section_rows(t, ONAME_COL, '公假')

    for rows in [lx_rows, bx_rows, gg_rows]:
        for r_idx in rows:
            set_cell(t.rows[r_idx].cells[OCODE_COL], '')
            set_cell(t.rows[r_idx].cells[ONAME_COL], '')

    for people, rows in [(lx, lx_rows), (bx, bx_rows), (gg, gg_rows)]:
        for j, s in enumerate(people):
            if j >= len(rows): break
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
    if not f or not tmpl_type:
        return jsonify({'error': '缺少參數'}), 400
    name_map = {
        '平日':'tmpl_平日.docx','假日':'tmpl_假日.docx',
        '平日取酒':'tmpl_平日取酒.docx','假日取酒':'tmpl_假日取酒.docx'
    }
    fname = name_map.get(tmpl_type)
    if not fname: return jsonify({'error': '無效類型'}), 400
    os.makedirs(ASSETS, exist_ok=True)
    f.save(os.path.join(ASSETS, fname))
    return jsonify({'ok': True})

if __name__ == '__main__':
    app.run(debug=True)