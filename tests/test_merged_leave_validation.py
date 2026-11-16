"""Regression for the 11/2 weekday template's vertically merged leave code."""
import io
import unittest
from unittest.mock import patch

from docx import Document

import app
from word_validation import table_row_cells, validate_generated_word


class MergedLeaveValidation(unittest.TestCase):
    def setUp(self):
        # Model the official grid without checking in operational Word files.
        doc = Document()
        table = doc.add_table(rows=16, cols=8)
        table.cell(0, 1).text = '中華民國115年11月2日【星期一】'
        for row, code in enumerate((2, 6, 13, 14, 15, 19, 54, 61), 4):
            table.cell(row, 3).text = str(code)
        for row, label in ((3, '輪   休'), (7, '休(補)假'), (12, '公假')):
            table.cell(row, 7).text = label
        table.cell(9, 6).merge(table.cell(10, 6))
        self.template = self.save(doc)
        self.data = dict(rocDate='115年11月2日', weekday='一', isHoliday=False,
                         hasSake=False, sakeTimes=[], onDuty=[], offDuty=[
                             dict(ban=code, name=f'測試{code}', status=status)
                             for code, status in ((2, '休'), (6, '輪'), (13, '輪'),
                                                  (54, '休假'), (61, '公假'))])
        with patch.object(app, 'template_bytes', return_value=self.template):
            self.output = app.generate_word(self.data).getvalue()

    @staticmethod
    def save(doc):
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    def validate(self, payload):
        validate_generated_word(payload, self.data, self.template,
                                app.PERSISTENT_STAFF_REMARKS, app.find_section_rows)

    def test_empty_continuation_is_not_a_second_leave_entry(self):
        self.validate(self.output)

    def test_two_people_do_not_share_the_merged_leave_code_slot(self):
        for earlier, later in ((14, 15), (15, 19)):
            with self.subTest(earlier=earlier, later=later):
                self.data['offDuty'] = [
                    dict(ban=code, name=f'測試{code}', status='休')
                    for code in (2, earlier, later)
                ]
                with patch.object(app, 'template_bytes', return_value=self.template):
                    output = app.generate_word(self.data).getvalue()
                self.validate(output)
                rows = table_row_cells(Document(io.BytesIO(output)).tables[0])
                self.assertEqual((rows[9][6].text, rows[9][7].text),
                                 (str(earlier), f'測試{earlier}'))
                self.assertEqual(rows[10][7].text, '')
                self.assertEqual((rows[11][6].text, rows[11][7].text),
                                 (str(later), f'測試{later}'))

    def test_actual_leave_errors_are_rejected(self):
        for kind in ('missing name', 'duplicate code', 'populated continuation',
                     'wrong section', 'wrong order'):
            with self.subTest(kind=kind):
                doc = Document(io.BytesIO(self.output))
                rows = table_row_cells(doc.tables[0])

                def write(i, j, value):
                    app.set_cell_preserving_format(rows[i][j], value)
                    app.apply_directory_text_format(rows[i][j], 12)

                if kind == 'missing name':
                    write(9, 7, '')
                elif kind == 'duplicate code':
                    write(11, 6, '54')
                    write(11, 7, '測試54')
                elif kind == 'populated continuation':
                    write(10, 7, '測試54')
                elif kind == 'wrong section':
                    write(13, 6, '54')
                    write(13, 7, '測試54')
                else:
                    write(4, 6, '13')
                    write(4, 7, '測試13')
                    write(5, 6, '6')
                    write(5, 7, '測試6')
                with self.assertRaisesRegex(ValueError, '假別名單、分類或輪休排序不符'):
                    self.validate(self.save(doc))
