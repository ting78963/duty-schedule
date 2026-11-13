import copy
import io
import unittest
from unittest.mock import patch
from docx import Document
from docx.shared import Pt
from lxml import etree
import app


def template_bytes():
    doc=Document()
    table=doc.add_table(rows=19, cols=12)
    table.cell(0,0).text='南投縣政府警察局交通隊49人勤務分配表'
    # Date deliberately split over multiple differently styled runs.
    date=table.cell(0,7)
    date.text=''
    date.paragraphs[0].add_run('中華民國115年')
    date.paragraphs[0].add_run('11月2日【星期一】')
    codes=[1,11,23,50,51,52,53,54,60,61,62,63]
    names=['林伯曉','徐志瑋','林筠臻','李智偉','林新慶','李金孔','董安民','陳昱靜','李安景','陳彥文','林志明','張瑞忠']
    for i,(code,name) in enumerate(zip(codes,names),4):
        table.cell(i,7).text=str(code)
        table.cell(i,8).text=name
        table.cell(i,9).text={11:'常訓',23:'留職停薪',54:'8-12休'}.get(code,'前日備考')
    for r,label in [(4,'輪   休'),(8,'休(補)假'),(12,'公假')]:
        table.cell(r,11).text=label
    table.cell(5,0).text='08-09'
    for c,text in [(1,'5'),(2,'(5)'),(3,'908'),(4,'BEW-7085'),(5,'取締酒後駕車'),(6,'1、2')]:
        table.cell(5,c).text=text
    table.cell(3,1).text='5'  # Non-time row: must survive.
    for i,cell in enumerate(table._tbl.iter('{'+app.ns+'}tc')):
        from docx.table import _Cell
        for p in _Cell(cell,table).paragraphs:
            for j,run in enumerate(p.runs):
                run.font.name='標楷體'
                run._r.get_or_add_rPr().rFonts.set('{'+app.ns+'}eastAsia','標楷體')
                run.font.size=Pt(10+(i%3))
                run.bold=bool(j%2)
    # Empty note without a run must inherit paragraph properties.
    blank=table.cell(9,9)
    for run in list(blank.paragraphs[0].runs):
        blank.paragraphs[0]._p.remove(run._r)
    ppr=blank.paragraphs[0]._p.get_or_add_pPr()
    rpr=etree.SubElement(ppr,'{'+app.ns+'}rPr')
    etree.SubElement(rpr,'{'+app.ns+'}sz').set('{'+app.ns+'}val','22')
    buf=io.BytesIO();doc.save(buf)
    return buf.getvalue()


def input_data():
    names={1:'林伯曉',11:'徐志瑋',23:'林筠臻',50:'李智偉',51:'林新慶',52:'李金孔',53:'董安民',54:'陳昱靜',60:'李安景',61:'陳彥文',62:'林志明',63:'張瑞忠'}
    people=[{'ban':code,'name':name,'status':{1:'常訓',51:'輪休',52:'8-12業',53:'常訓',61:'公假'}.get(code,'')} for code,name in names.items()]
    people.append({'ban':None,'name':'周彥璋','status':''})
    return dict(rocDate='115年11月6日',weekday='五',isHoliday=False,hasSake=False,sakeTimes=[],onDuty=[p for p in people if p['status'] not in ('輪休','公假')],offDuty=[p for p in people if p['status'] in ('輪休','公假')])


class WordGenerationRegression(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source=template_bytes()
        cls.payload=input_data()
        with patch.object(app,'template_bytes',return_value=cls.source),patch.object(app,'sb_get',side_effect=AssertionError('Live roster forbidden')):
            cls.output=app.generate_word(copy.deepcopy(cls.payload)).getvalue()
        cls.doc=Document(io.BytesIO(cls.output))
        cls.rows=[r.cells for r in cls.doc.tables[0].rows]
        cls.bycode={int(r[7].text):(r[8].text,r[9].text) for r in cls.rows[4:] if r[7].text.isdigit()}

    def test_full_date_and_weekday_replace_split_runs(self):
        self.assertEqual(self.rows[0][7].text,'中華民國115年11月6日【星期五】')
        self.assertEqual(len(self.rows[0][7].paragraphs[0].runs),3)
        self.assertEqual([r.text for r in self.rows[0][7].paragraphs[0].runs],['','中華民國115年','11月6日【星期五】'])

    def test_all_notes_rebuilt_and_persistent_note_retained(self):
        for code,note in {11:'',54:'',52:'8-12業',53:'常訓',23:'留職停薪',50:'',60:'',61:'公假'}.items():
            self.assertEqual(self.bycode[code][1],note)
        self.assertNotIn('前日備考', '\n'.join(r[9].text for r in self.rows[4:16]))

    def test_hand_codes_only_and_fixed_fields_untouched(self):
        self.assertEqual([self.rows[5][i].text for i in (1,2,6)],['','',''])
        self.assertEqual([self.rows[5][i].text for i in (0,3,4,5)],['08-09','908','BEW-7085','取締酒後駕車'])
        self.assertEqual(self.rows[3][1].text,'5')
        self.assertEqual(self.rows[5][7].text,'11')

    def test_existing_runs_and_column_formats_preserved(self):
        before=Document(io.BytesIO(self.source)).tables[0]
        after=self.doc.tables[0]
        self.assertEqual(len(before.rows),len(after.rows))
        self.assertEqual([c.width for c in before.columns],[c.width for c in after.columns])
        for oldrow,newrow in zip(before.rows,after.rows):
            self.assertEqual(oldrow.height,newrow.height)
            for old,new in zip(oldrow.cells,newrow.cells):
                self.assertEqual(len(old.paragraphs),len(new.paragraphs))
                self.assertGreaterEqual(sum(len(p.runs) for p in new.paragraphs),sum(len(p.runs) for p in old.paragraphs))
        for row in self.rows[4:16]:
            size_by_code={23:9,35:8,43:8,60:8,61:8,62:8,63:8}
            note_size=size_by_code.get(int(row[7].text),12)
            for column,size in [(7,12),(8,12),(9,note_size)]:
                for para in row[column].paragraphs:
                    for run in para.runs:
                        self.assertEqual(run.font.name,'標楷體')
                        self.assertEqual(run.font.size.pt,size)
                        self.assertIs(run.bold,False)
                        self.assertEqual(str(run.font.color.rgb),'000000')
        for i in (5,13):
            for col in (10,11):
                for run in self.rows[i][col].paragraphs[0].runs:
                    self.assertEqual(run.font.size.pt,12)
                    self.assertIs(run.bold,False)
                    self.assertEqual(str(run.font.color.rgb),'000000')

    def test_blue_bold_and_wrong_size_are_replaced_without_rebuilding_runs(self):
        from docx.shared import RGBColor
        doc=Document()
        cell=doc.add_table(rows=1,cols=1).cell(0,0)
        run=cell.paragraphs[0].add_run('輪')
        run.font.color.rgb=RGBColor(0,0,255)
        run.font.size=Pt(18)
        run.bold=True
        app.apply_directory_text_format(cell,7)
        self.assertEqual(len(cell.paragraphs),1)
        self.assertEqual(len(cell.paragraphs[0].runs),1)
        self.assertEqual(cell.text,'輪')
        self.assertEqual(run.font.size.pt,7)
        self.assertIs(run.bold,False)
        self.assertEqual(str(run.font.color.rgb),'000000')

    def test_newcomer_not_assigned_or_inserted(self):
        self.assertNotIn('周彥璋','\n'.join(c.text for r in self.rows for c in r))
        self.assertEqual(len(self.bycode),12)
        self.assertIsNone(next(p['ban'] for p in self.payload['onDuty'] if p['name']=='周彥璋'))

    def test_persistent_remark_does_not_follow_code_to_replacement(self):
        payload=copy.deepcopy(self.payload)
        next(p for p in payload['onDuty'] if p['ban']==23)['name']='新人測試'
        with patch.object(app,'template_bytes',return_value=self.source):
            doc=Document(app.generate_word(payload))
        self.assertEqual(doc.tables[0].cell(6,9).text,'')

    def test_current_day_full_remarks_retained(self):
        payload=copy.deepcopy(self.payload)
        person=next(p for p in payload['onDuty'] if p['ban']==60)
        person.update(status='停休',remarks='停休 8-12督勤')
        with patch.object(app,'template_bytes',return_value=self.source):
            doc=Document(app.generate_word(payload))
        self.assertEqual(doc.tables[0].cell(12,9).text,'停休 8-12督勤')

    def test_timed_note_uses_accepted_1106_size_without_shrinking(self):
        doc=Document(io.BytesIO(self.source))
        cell=doc.tables[0].cell(9,9)
        cell.paragraphs[0]._p.remove(cell.paragraphs[0]._p.pPr)
        donor=doc.tables[0].cell(11,9).paragraphs[0].runs[0]
        donor.font.size=Pt(7)
        buf=io.BytesIO();doc.save(buf)
        with patch.object(app,'template_bytes',return_value=buf.getvalue()):
            result=Document(app.generate_word(copy.deepcopy(self.payload)))
        run=result.tables[0].cell(9,9).paragraphs[0].runs[0]
        self.assertEqual(run.text,'8-12業')
        self.assertEqual(run.font.size.pt,12)

    def test_right_rotating_leave_uses_only_actual_numeric_code(self):
        for code in (1,20,21,49,50,53,54,59,60,63,64):
            with self.subTest(code=code):
                payload=copy.deepcopy(self.payload)
                payload['offDuty'].append({'ban':code,'name':'測試人員','status':'輪','sourceSheet':'任意身分'})
                with patch.object(app,'template_bytes',return_value=self.source):
                    doc=Document(app.generate_word(payload))
                rows=[r.cells for r in doc.tables[0].rows]
                slots=app.find_section_rows(doc.tables[0],11,'輪   休',rows)
                codes=[int(rows[i][10].text) for i in slots if rows[i][10].text.isdigit()]
                self.assertEqual(code in codes,1<=code<=20 or 50<=code<=53 or 60<=code<=63)
        payload=copy.deepcopy(self.payload)
        payload['offDuty'].append({'ban':1,'name':'測試停休','status':'△輪'})
        with patch.object(app,'template_bytes',return_value=self.source):
            doc=Document(app.generate_word(payload))
        rows=[r.cells for r in doc.tables[0].rows]
        slots=app.find_section_rows(doc.tables[0],11,'輪   休',rows)
        self.assertNotIn('1',[rows[i][10].text for i in slots])

    def test_mismatched_weekday_rejected(self):
        payload=copy.deepcopy(self.payload);payload['weekday']='一'
        with patch.object(app,'template_bytes',return_value=self.source),self.assertRaisesRegex(ValueError,'星期'):
            app.generate_word(payload)

if __name__=='__main__': unittest.main()
