from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from lxml import etree
import openpyxl, json, io, os, re, copy, zipfile, requests as req

app = Flask(__name__)

# Supabase REST API
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
    r = req.post(f"{SUPABASE_URL}/rest/v1/{table}", headers=sb_headers(), json=data)
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

# ─── 模板路徑 ──────────────────────────────────────────────
def get_tmpl(is_holiday, has_sake):
    if is_holiday and has_sake: return os.path.join(ASSETS, 'tmpl_假日取酒.docx')
    if is_holiday:              return os.path.join(ASSETS, 'tmpl_假日.docx')
    if has_sake:                return os.path.join(ASSETS, 'tmpl_平日取酒.docx')
    return os.path.join(ASSETS, 'tmpl_平日.docx')

# ─── 欄位設定 ──────────────────────────────────────────────
def get_cols(is_holiday):
    if is_holiday:
        return 28, 29, 30, 31, 32   # CODE, NAME, NOTE, OCODE, ONAME
    return 30, 31, 32, 33, 34

# ─── 字體設定 ──────────────────────────────────────────────
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

# ─── 紅框設定（取締酒駕時段）──────────────────────────────
def set_red_border(row):
    """在指定列的所有儲存格加上紅色框線"""
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

# ─── 找輪休欄可用列 ────────────────────────────────────────
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

# ─── 從大表解析取締酒駕時段 ────────────────────────────────
def parse_sake_time(cell_value):
    """從儲存格文字解析時段，例如 '18-22 局頒 取締酒駕' → '18-22'"""
    if not cell_value: return None
    m = re.search(r'(\d{1,2})-(\d{1,2})', str(cell_value))
    if m: return f"{m.group(1)}-{m.group(2)}"
    return None

# ─── 產生 Word 主函數 ──────────────────────────────────────
def generate_word(data):
    on_duty  = data['onDuty']
    off_duty = data['offDuty']
    roc_date = data['rocDate']
    weekday  = data['weekday']
    is_holiday = data['isHoliday']
    has_sake   = data.get('hasSake', False)
    sake_times = data.get('sakeTimes', [])  # [{start:18, end:22}, ...]

    tmpl_path = get_tmpl(is_holiday, has_sake)
    CODE_COL, NAME_COL, NOTE_COL, OCODE_COL, ONAME_COL = get_cols(is_holiday)

    # 分類放假人員
    lx, bx, gg = [], [], []
    for s in sorted(off_duty, key=lambda x: x['ban']):
        st = s.get('status') or ''
        if st == '輪': lx.append(s)
        elif st in ['優','休','補']: bx.append(s)
        elif st == '公': gg.append(s)

    # 從資料庫取人員名冊
    staff_map = {}  # code -> name
    if supabase:
        res = supabase.table('staff').select('*').execute()
        for p in res.data:
            staff_map[p['code']] = p['name']

    doc = Document(tmpl_path)
    t = doc.tables[0]

    # 找代號列
    ban_to_row = {}
    for r_idx in range(4, len(t.rows)):
        if CODE_COL >= len(t.rows[r_idx].cells): continue
        ct = t.rows[r_idx].cells[CODE_COL].text.strip()
        if ct.isdigit() and 1 <= int(ct) <= 20:
            ban_to_row[int(ct)] = r_idx

    # 清空名冊
    for ban, r_idx in ban_to_row.items():
        for ci in [NAME_COL, NOTE_COL]:
            for para in t.rows[r_idx].cells[ci].paragraphs:
                for run in para.runs: run.text = ''

    # 填在班
    for s in on_duty:
        ban = s['ban']
        if ban not in ban_to_row: continue
        set_cell(t.rows[ban_to_row[ban]].cells[NAME_COL], s['name'])
        set_cell(t.rows[ban_to_row[ban]].cells[NOTE_COL], s.get('status') or '', size=24)

    # 填放假
    for s in off_duty:
        ban = s['ban']
        if ban not in ban_to_row: continue
        set_cell(t.rows[ban_to_row[ban]].cells[NAME_COL], s['name'])
        set_cell(t.rows[ban_to_row[ban]].cells[NOTE_COL], s.get('status') or '', size=24)

    # 輪休欄
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
            set_cell(t.rows[rows[j]].cells[OCODE_COL], str(s['ban']))
            set_cell(t.rows[rows[j]].cells[ONAME_COL], s['name'])

    # 日期
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

    # 取締酒駕紅框（支援跨夜：08:00 為一天的起點）
    if sake_times and has_sake:
        # 找時段列的對應（從表格第一欄找時間標記）
        time_to_rows = {}
        for r_idx in range(len(t.rows)):
            row = t.rows[r_idx]
            if not row.cells: continue
            cell_txt = row.cells[0].text.strip()
            m = re.match(r'(\d{1,2})[－\-–](\d{1,2})', cell_txt)
            if m:
                h = int(m.group(1))
                time_to_rows[h] = r_idx

        for st in sake_times:
            s_h = int(st['start'])
            e_h = int(st['end'])
            # 處理跨夜：08點為起點，超過24算次日
            hours = []
            h = s_h
            while True:
                hours.append(h % 24)
                h += 1
                if h % 24 == e_h % 24:
                    break
                if len(hours) > 16: break  # 防止無限迴圈

            for h in hours:
                if h in time_to_rows:
                    set_red_border(t.rows[time_to_rows[h]])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─── 解析大表 ─────────────────────────────────────────────
def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb['全隊大表'] if '全隊大表' in wb.sheetnames else wb.active

    # 找日期行和人員行
    # 找所有含「取締」的儲存格，解析日期和時段
    sake_by_date = {}
    date_col_map = {}

    for row in ws.iter_rows():
        for cell in row:
            if cell.value and '取締' in str(cell.value):
                time_str = parse_sake_time(cell.value)
                if time_str:
                    # 找同列的日期（第一欄）
                    date_val = ws.cell(row=cell.row, column=1).value
                    if date_val:
                        sake_by_date[str(date_val)] = time_str

    return {'sakeByDate': sake_by_date}

# ─── API 路由 ──────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/staff', methods=['GET'])
def get_staff():
    if not SUPABASE_URL:
        return jsonify([])
    data = sb_get('staff', params={'order': 'code'})
    return jsonify(data)

@app.route('/api/staff', methods=['POST'])
def add_staff():
    data = sb_post('staff', request.json)
    return jsonify(data)

@app.route('/api/staff/<int:code>', methods=['PUT'])
def update_staff(code):
    data = sb_patch('staff', 'code', code, request.json)
    return jsonify(data)

@app.route('/api/staff/<int:code>', methods=['DELETE'])
def delete_staff(code):
    sb_delete('staff', 'code', code)
    return jsonify({'ok': True})

@app.route('/api/parse-excel', methods=['POST'])
def parse_excel_api():
    f = request.files.get('file')
    if not f: return jsonify({'error': '未上傳檔案'}), 400
    result = parse_excel(f.read())
    return jsonify(result)

@app.route('/api/generate', methods=['POST'])
def generate_api():
    data = request.json
    buf = generate_word(data)
    roc_date = data.get('rocDate', '勤務表')
    fname = f'勤務表_{roc_date}.docx'
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/api/generate-batch', methods=['POST'])
def generate_batch():
    """批次產生多天，回傳 ZIP"""
    days_data = request.json  # [{day1_data}, {day2_data}, ...]
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for d in days_data:
            buf = generate_word(d)
            fname = f"勤務表_{d.get('rocDate','未知')}.docx"
            zf.writestr(fname, buf.read())
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True, download_name='勤務表批次.zip',
                     mimetype='application/zip')

@app.route('/api/upload-template', methods=['POST'])
def upload_template():
    tmpl_type = request.form.get('type')  # 平日/假日/平日取酒/假日取酒
    f = request.files.get('file')
    if not f or not tmpl_type:
        return jsonify({'error': '缺少參數'}), 400
    name_map = {'平日':'tmpl_平日.docx','假日':'tmpl_假日.docx','平日取酒':'tmpl_平日取酒.docx','假日取酒':'tmpl_假日取酒.docx'}
    fname = name_map.get(tmpl_type)
    if not fname: return jsonify({'error': '無效類型'}), 400
    f.save(os.path.join(ASSETS, fname))
    return jsonify({'ok': True})

if __name__ == '__main__':
    app.run(debug=True)