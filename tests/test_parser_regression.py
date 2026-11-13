import copy
import io
import unittest
from datetime import datetime

import openpyxl
import app


def workbook_bytes():
    wb = openpyxl.Workbook()
    ext = wb.active
    ext.title = '外勤'
    full = wb.create_sheet('全隊大表')
    # Synthetic identities ensure classification does not depend on four names.
    internal = ['李智偉', '林新慶', '李金孔', '董安民', '陳昱靜',
                '林筠臻', '黃雅琳', '洪君智', '楊智為', '陳淑卿', '謝函儒',
                '曾義雄', '姚宗志', '柯育婷', '林裕善', '王庭偉', '周彥璋',
                '林家豪', '張銥凌', '翁水成', '洪呈欣', '簡竹芳', '沈宗漢', '曾煙煌', '陳駿騰']
    cols = [c for c in range(49, 75) if c != 61]
    for c, name in zip(cols, internal):
        full.cell(2, c, name)
    for i in range(20):
        ext.cell(2, 3 + 2*i, f'外勤人員{i}')
    for r, d in enumerate((1, 6, 7, 8), 3):
        date = datetime(2026, 11, d)
        ext.cell(r, 1, date)
        full.cell(r, 47, date)
        for i in range(20):
            ext.cell(r, 3+2*i, (i+d-1) % 16+1 if i < 16 else i+1)
        full.cell(r, 49, 'Δ 輪' if d == 1 else '常訓')
        full.cell(r, 50, '輪')
        full.cell(r, 51, '8-12業')
        full.cell(r, 52, '常訓')
    # Supervisor dates belong to column A in their own lower block.
    for c, title in enumerate(('隊長', '林副隊長', '陳副隊長', '李秘書'), 1):
        full.cell(39, c, title)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class ParserRegression(unittest.TestCase):
    def test_internal_aw_az_are_not_matched_as_external(self):
        result = app.parse_excel(workbook_bytes())
        app.validate_monthly_data_before_write(result['monthlyData'])
        for day in result['monthlyData'].values():
            people = day['on_duty'] + day['off_duty']
            self.assertEqual(len(people), 45)
            self.assertEqual(len({p['name'] for p in people}), 45)
            self.assertEqual(sorted(p['ban'] for p in people if p['sourceSheet'] == '外勤'), list(range(1, 21)))
            for name, code, column in [('李智偉', 50, 'AW'), ('林新慶', 51, 'AX'), ('李金孔', 52, 'AY'), ('董安民', 53, 'AZ')]:
                person, = [p for p in people if p['name'] == name]
                self.assertEqual((person['ban'], person['sourceSheet'], person['sourceColumn']), (code, '全隊大表', column))
            self.assertIsNone(next(p['ban'] for p in people if p['name'] == '周彥璋'))
        first = result['monthlyData']['115年11月1日']
        self.assertEqual(next(p['status'] for p in first['on_duty'] if p['name'] == '李智偉'), '停休')

    def test_supervisor_date_source_and_special_statuses(self):
        wb = openpyxl.load_workbook(io.BytesIO(workbook_bytes()))
        ws = wb['全隊大表']
        ws.cell(39, 1, '日期')
        for col, title in [(3, '長\n隊'), (6, '林副隊長'), (10, '陳副隊長'), (14, '李秘書')]:
            ws.cell(39, col, title)
        # Remove the old synthetic title cells.
        ws.cell(39, 2).value = None
        ws.cell(39, 4).value = None
        ws.cell(40, 1, datetime(2026, 11, 6))
        ws.cell(40, 3, '輪休')
        ws.cell(40, 6, '輪休')
        ws.cell(40, 10, '公假')
        ws.cell(40, 14, '假日交督 停休6-16')
        buf = io.BytesIO()
        wb.save(buf)
        result = app.parse_excel(buf.getvalue())['monthlyData']
        app.validate_monthly_data_before_write(result)
        day = result['115年11月6日']
        self.assertEqual({p['ban'] for p in day['off_duty'] if p['ban'] in (60, 61, 62, 63)}, {61, 62, 63})
        secretary, = [p for p in day['on_duty'] if p['ban'] == 60]
        self.assertEqual((secretary['status'], secretary['sourceColumn'], secretary['sourceRow']), ('停休', 'N', 40))
        self.assertFalse(app.is_leave_status('假日 交督'))
        self.assertFalse(app.is_leave_status('常訓'))
        self.assertTrue(app.is_leave_status('家庭照顧假'))

    def test_missing_internal_person_is_rejected(self):
        data = app.parse_excel(workbook_bytes())['monthlyData']
        day = data['115年11月6日']
        day['on_duty'] = [p for p in day['on_duty'] if p['name'] != '李智偉']
        with self.assertRaisesRegex(ValueError, '25'):
            app.validate_monthly_data_before_write(data)

    def test_valid_and_invalid_roc_dates(self):
        source = app.parse_excel(workbook_bytes())['monthlyData']['115年11月6日']
        for date, weekday in [('115年11月6日', '五'), ('115年11月06日', '五'), ('113年2月29日', '四')]:
            day = copy.deepcopy(source)
            day.update(date_str=date, weekday=weekday)
            app.validate_monthly_data_before_write({date: day})
        for date in ['115年11月31日', '115年2月29日', '115年11月6日junk', '2026-11-06']:
            day = copy.deepcopy(source)
            day['date_str'] = date
            with self.assertRaises(ValueError):
                app.validate_monthly_data_before_write({date: day})

    def test_duplicate_external_code_is_rejected(self):
        data = app.parse_excel(workbook_bytes())['monthlyData']
        day = data['115年11月6日']
        ext = [p for p in day['on_duty']+day['off_duty'] if p['sourceSheet'] == '外勤']
        ext[1]['ban'] = ext[0]['ban']
        with self.assertRaises(ValueError):
            app.validate_monthly_data_before_write(data)


if __name__ == '__main__':
    unittest.main()
