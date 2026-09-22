from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from lxml import etree
import openpyxl, json, io, os, re, copy, zipfile
from supabase import create_client

app = Flask(__name__)
SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')
supabase = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL else None
ASSETS = os.path.join(os.path.dirname(__file__), 'assets')
ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def get_tmpl(is_holiday, has_sake):
    if is_holiday and has_sake: return os.path.join(ASSETS, 'tmpl_假日取酒.docx')
    if is_holiday: return os.path.join(ASSETS, 'tmpl_假日.docx')
    if has_sake: return os.path.join(ASSETS, 'tmpl_平日取酒.docx')
    return os.path.join(ASSETS, 'tmpl_平日.docx')

def get_cols(is_holiday):
    return (28,29,30,31,32) if is_holiday else (30,31,32,33,34)

def set_cell(cell, text, size=None, color='000000'):
    text = str(text) if text else ''
    for para in cell.paragraphs:
        runs = para.runs
        if runs:
            runs[0].text = text
            for rr in runs[1:]: rr.text = ''
            rpr = runs[0]._r.find(f'{{{ns}}}rPr')
            if rpr is None:
                rpr = etree.SubElement(runs[0]._r, f'{{{ns}}}rPr'); runs[0]._r.insert(0,rpr)
            fonts = rpr.find(f'{{{ns}}}rFonts')
            if fonts is None: fonts = etree.SubElement(rpr,f'{{{ns}}}rFonts')
            for attr in ['ascii','hAnsi','eastAsia','cs']: fonts.set(f'{{{ns}}}{attr}','標楷體')
            col_el = rpr.find(f'{{{ns}}}color')
            if col_el is None: col_el=etree.SubElement(rpr,f'{{{ns}}}color')
            col_el.set(f'{{{ns}}}val',color)
            if size:
                for tag in ['sz','szCs']:
                    el=rpr.find(f'{{{ns}}}{tag}')
                    if el is None: el=etree.SubElement(rpr,f'{{{ns}}}{tag}')
                    el.set(f'{{{ns}}}val',str(size))
            return
    cell.paragraphs[0].add_run(text)

def set_red_border(row):
    for cell in row.cells:
        tc=cell._tc; tcPr=tc.find(f'{{{ns}}}tcPr')
        if tcPr is None: tcPr=etree.SubElement(tc,f'{{{ns}}}tcPr')
        borders=tcPr.find(f'{{{ns}}}tcBorders')
        if borders is None: borders=etree.SubElement(tcPr,f'{{{ns}}}tcBorders')
        for side in ['top','bottom','left','right']:
            el=borders.find(f'{{{ns}}}{side}')
            if el is None: el=etree.SubElement(borders,f'{{{ns}}}{side}')
            el.set(f'{{{ns}}}val','single'); el.set(f'{{{ns}}}sz','16'); el.set(f'{{{ns}}}color','FF0000')

def find_section_rows(t,oname_col,label):
    start=None
    for r_idx in range(len(t.rows)):
        if oname_col < len(t.rows[r_idx].cells) and label in t.rows[r_idx].cells[oname_col].text.strip():
            start=r_idx+1; break
    if start is None: return []
    rows=[]
    for r_idx in range(start,len(t.rows)):
        if oname_col>=len(t.rows[r_idx].cells): break
        txt=t.rows[r_idx].cells[oname_col].text.strip()
        if txt and any(x in txt for x in ['假','附記','巡邏','調適']): break
        rows.append(r_idx)
    return rows

def parse_sake_time(v):
    if not v: return None
    m=re.search(r'(\d{1,2})-(\d{1,2})',str(v))
    return f"{m.group(1)}-{m.group(2)}" if m else None

def generate_word(data):
    on_duty=data['onDuty']; off_duty=data['offDuty']; roc_date=data['rocDate']; weekday=data['weekday']
    is_holiday=data['isHoliday']; has_sake=data.get('hasSake',False); sake_times=data.get('sakeTimes',[])
    CODE_COL,NAME_COL,NOTE_COL,OCODE_COL,ONAME_COL=get_cols(is_holiday)
    lx,bx,gg=[],[],[]
    for s in sorted(off_duty,key=lambda x:x['ban']):
        st=s.get('status') or ''
        if st=='輪': lx.append(s)
        elif st in ['優','休','補']: bx.append(s)
        elif st=='公': gg.append(s)
    doc=Document(get_tmpl(is_holiday,has_sake)); t=doc.tables[0]
    ban_to_row={}
    for r_idx in range(4,len(t.rows)):
        if CODE_COL>=len(t.rows[r_idx].cells): continue
        ct=t.rows[r_idx].cells[CODE_COL].text.strip()
        if ct.isdigit() and 1<=int(ct)<=20: ban_to_row[int(ct)]=r_idx
    for ban,r_idx in ban_to_row.items():
        for ci in [NAME_COL,NOTE_COL]:
            for para in t.rows[r_idx].cells[ci].paragraphs:
                for run in para.runs: run.text=''
    for s in on_duty+off_duty:
        if s['ban'] in ban_to_row:
            set_cell(t.rows[ban_to_row[s['ban']]].cells[NAME_COL],s['name'])
            set_cell(t.rows[ban_to_row[s['ban']]].cells[NOTE_COL],s.get('status') or '',size=24)
    lx_rows=find_section_rows(t,ONAME_COL,'輪   休'); bx_rows=find_section_rows(t,ONAME_COL,'休(補)假'); gg_rows=find_section_rows(t,ONAME_COL,'公假')
    for rows in [lx_rows,bx_rows,gg_rows]:
        for r_idx in rows: set_cell(t.rows[r_idx].cells[OCODE_COL],''); set_cell(t.rows[r_idx].cells[ONAME_COL],'')
    for people,rows in [(lx,lx_rows),(bx,bx_rows),(gg,gg_rows)]:
        for j,s in enumerate(people):
            if j>=len(rows): break
            set_cell(t.rows[rows[j]].cells[OCODE_COL],str(s['ban'])); set_cell(t.rows[rows[j]].cells[ONAME_COL],s['name'])
    day=roc_date.split('年')[1].split('日')[0]+'日'
    for cell in t.rows[0].cells:
        if '民國' in cell.text and '年' in cell.text:
            for para in cell.paragraphs:
                for run in para.runs:
                    if '年月日【' in run.text: run.text=f'年{day}【'
                    elif run.text=='星期': run.text=f'星期{weekday}'
            break
    if sake_times and has_sake:
        time_to_rows={}
        for r_idx,row in enumerate(t.rows):
            if not row.cells: continue
            m=re.match(r'(\d{1,2})[－\-–](\d{1,2})',row.cells[0].text.strip())
            if m: time_to_rows[int(m.group(1))]=r_idx
        for st in sake_times:
            h=int(st['start']); e_h=int(st['end']); hours=[]
            while True:
                hours.append(h%24); h+=1
                if h%24==e_h%24 or len(hours)>16: break
            for h in hours:
                if h in time_to_rows: set_red_border(t.rows[time_to_rows[h]])
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf

def parse_excel(file_bytes):
    wb=openpyxl.load_workbook(io.BytesIO(file_bytes),data_only=True)
    ws=wb['全隊大表'] if '全隊大表' in wb.sheetnames else wb.active
    sake_by_date={}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and '取締' in str(cell.value):
                ts=parse_sake_time(cell.value)
                if ts:
                    dv=ws.cell(row=cell.row,column=1).value
                    if dv: sake_by_date[str(dv)]=ts
    return {'sakeByDate':sake_by_date}

@app.route('/')
def index(): return render_template('index.html')

@app.route('/api/staff',methods=['GET'])
def get_staff():
    if not SUPABASE_URL:
        return jsonify([])
    data = sb_get('staff', params={'order': 'code'})
    return jsonify(data)

@app.route('/api/staff',methods=['POST'])
def add_staff():
    data = sb_post('staff', request.json)
    return jsonify(data)

@app.route('/api/staff/<int:code>',methods=['PUT'])
def update_staff(code):
    data = sb_patch('staff', 'code', code, request.json)
    return jsonify(data)

@app.route('/api/staff/<int:code>',methods=['DELETE'])
def delete_staff(code):
    sb_delete('staff', 'code', code)
    return jsonify({'ok': True})

@app.route('/api/parse-excel',methods=['POST'])
def parse_excel_api():
    f=request.files.get('file')
    if not f: return jsonify({'error':'未上傳檔案'}),400
    return jsonify(parse_excel(f.read()))

@app.route('/api/generate',methods=['POST'])
def generate_api():
    data=request.json; buf=generate_word(data); fname=f"勤務表_{data.get('rocDate','勤務表')}.docx"
    return send_file(buf,as_attachment=True,download_name=fname,mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/api/generate-batch',methods=['POST'])
def generate_batch():
    zip_buf=io.BytesIO()
    with zipfile.ZipFile(zip_buf,'w') as zf:
        for d in request.json:
            buf=generate_word(d); zf.writestr(f"勤務表_{d.get('rocDate','未知')}.docx",buf.read())
    zip_buf.seek(0)
    return send_file(zip_buf,as_attachment=True,download_name='勤務表批次.zip',mimetype='application/zip')

@app.route('/api/upload-template',methods=['POST'])
def upload_template():
    tmpl_type=request.form.get('type'); f=request.files.get('file')
    if not f or not tmpl_type: return jsonify({'error':'缺少參數'}),400
    name_map={'平日':'tmpl_平日.docx','假日':'tmpl_假日.docx','平日取酒':'tmpl_平日取酒.docx','假日取酒':'tmpl_假日取酒.docx'}
    fname=name_map.get(tmpl_type)
    if not fname: return jsonify({'error':'無效類型'}),400
    os.makedirs(ASSETS,exist_ok=True); f.save(os.path.join(ASSETS,fname)); return jsonify({'ok':True})

if __name__=='__main__': app.run(debug=True)
