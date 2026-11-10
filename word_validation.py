"""Read-only delivery checks for the approved Word contract; no document edits."""
import io
import re
from collections import defaultdict
from itertools import combinations
from docx import Document
from docx.table import _Cell
from lxml import etree

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{' + NS + '}'


def table_row_cells(table):
    """Resolve the logical grid once, reusing each vertical merge from the previous row."""
    rows = []
    previous = {}
    for tr in table._tbl.tr_lst:
        offset = tr.grid_before
        current = {}
        row = []
        for tc in tr.tc_lst:
            span = tc.grid_span
            if tc.vMerge == 'continue':
                if offset not in previous:
                    raise ValueError('Word 表格的合併儲存格缺少上一列')
                cells = previous[offset]
            else:
                cells = (_Cell(tc, table),) * span
            current[offset] = cells
            row.extend(cells)
            offset += span
        rows.append(tuple(row))
        previous = current
    return rows


def sake_frame_times(table, rows, duty_columns):
    """Read closed red frames on the physical table grid, including merge continuations."""
    horizontal = defaultdict(set)
    vertical = defaultdict(set)
    slots = []
    for y, (tr, row) in enumerate(zip(table._tbl.tr_lst, rows)):
        match = re.fullmatch(r'(\d{1,2})[－\-–](\d{1,2})', row[0].text.strip())
        slots.append(tuple(map(int, match.groups())) if match else None)
        if not match:
            continue
        x = tr.grid_before
        for cell in tr.tc_lst:
            right = x + cell.grid_span
            if x >= duty_columns:
                break
            if right <= duty_columns:
                for border in cell.xpath('./w:tcPr/w:tcBorders/*'):
                    color = border.get(W + 'color', '').upper()
                    if color not in ('FF0000', 'EE0000') or border.get(W + 'val') != 'single':
                        continue
                    side = etree.QName(border).localname
                    if side in ('top', 'bottom'):
                        horizontal[color, y + (side == 'bottom')].update(range(x, right))
                    elif side in ('left', 'right'):
                        vertical[color, x if side == 'left' else right].add(y)
            x = right

    # Pair full-height sides before checking horizontal edges. A stray red
    # horizontal edge inside a split hour must not create another interval.
    sides_by_span = defaultdict(list)
    for (color, x), covered in vertical.items():
        for first in sorted(covered):
            if first - 1 in covered:
                continue
            end = first + 1
            while end in covered:
                end += 1
            sides_by_span[color, first, end].append(x)

    times = []
    for (color, first, end), sides in sides_by_span.items():
        if (first > 0 and slots[first - 1] == slots[first]
                or end < len(slots) and slots[end - 1] == slots[end]):
            continue
        hours = [slot for i, slot in enumerate(slots[first:end], first)
                 if i == first or slot != slots[i - 1]]
        if (any(slot is None for slot in hours)
                or any(not (0 <= start <= 23 and 0 <= stop <= 24 and (start + 1) % 24 == stop % 24)
                       for start, stop in hours)
                or any(a[1] % 24 != b[0] for a, b in zip(hours, hours[1:]))):
            continue
        top = horizontal[color, first]
        bottom = horizontal[color, end]
        for left, right in combinations(sorted(sides), 2):
            if set(range(left, right)) <= top & bottom:
                times.append((hours[0][0], hours[-1][1]))
    return sorted(times)


def validate_generated_word(payload, data, template, persistent_notes, section_rows):
    def require(ok, reason):
        if not ok:
            raise ValueError('Word 驗證失敗：' + reason)

    try:
        doc, source = Document(io.BytesIO(payload)), Document(io.BytesIO(template))
    except Exception as exc:
        raise ValueError('Word 驗證失敗：無法開啟產出文件') from exc
    require(len(doc.tables) == len(source.tables) > 0, '表格數量不符')
    require(etree.tostring(doc._element.body.sectPr) == etree.tostring(source._element.body.sectPr), '頁面設定改變')
    table, base = doc.tables[0], source.tables[0]
    require(len(table.rows) == len(base.rows) and len(table.columns) == len(base.columns), '表格列欄數改變')
    require(etree.tostring(table._tbl.tblGrid) == etree.tostring(base._tbl.tblGrid), '欄寬改變')
    rows, old = table_row_cells(table), table_row_cells(base)
    c = len(table.columns) - 5
    require(c >= 1, '名冊欄位不存在')
    mutable = set()
    date_cells = {cell._tc for cell in rows[0] if '民國' in cell.text and '年' in cell.text}
    require(len(date_cells) == 1, '日期欄缺漏或重複')
    match = re.fullmatch(r'(\d+)年(\d+)月(\d+)日', data['rocDate'])
    require(match is not None, '日期格式錯誤')
    y, m, d = map(int, match.groups())
    expected_date = f'中華民國{y}年{m}月{d}日【星期{data["weekday"]}】'
    require(any(cell.text == expected_date for cell in rows[0]), '日期或星期內容不符')
    mutable.update(date_cells)
    if data.get('hasSake'):
        expected = sorted((t['start'], t['end']) for t in data['sakeTimes'])
        actual = sake_frame_times(table, rows, c)
        # Excel and Word can use either 24 or 00 for midnight.
        normalized = lambda times: sorted((start, end % 24) for start, end in times)
        excel_times = '、'.join(f'{start:02d}–{end:02d}' for start, end in expected) or '未提供'
        word_times = '、'.join(f'{start:02d}–{end:02d}' for start, end in actual) or '無法辨識完整紅框'
        require(bool(expected) and normalized(expected) == normalized(actual),
                f'{data["rocDate"]} 取締酒駕時段不一致（Excel：{excel_times}，Word：{word_times}）')
    people = {p['ban']: p for p in data['onDuty'] + data['offDuty'] if isinstance(p.get('ban'), int)}
    directory = {}
    for i, row in enumerate(rows[4:], 4):
        text = row[c].text.strip()
        if text != old[i][c].text.strip():
            require(text == '60' and not old[i][c].text.strip()
                    and i+1 < len(old) and old[i+1][c].text.strip() == '61', '名冊代號位置改變')
        if text.isdigit() and 1 <= int(text) <= 63:
            code = int(text)
            require(code not in directory, '名冊代號重複')
            directory[code] = row
    require(set(people) <= set(directory), '名冊遺漏已編號人員')

    def font(cell, size):
        for para in cell.paragraphs:
            for run in para.runs:
                require(run.font.name == '標楷體' and run.font.size is not None
                        and run.font.size.pt == size and run.bold is False
                        and str(run.font.color.rgb) == '000000', '姓名、假別或備考字型不符')

    for code, row in directory.items():
        p = people.get(code, {})
        name = str(p.get('name') or '')
        note = str(p.get('remarks') or p.get('supervisorSource', {}).get('remarks') or p.get('status') or '').strip()
        note = ' '.join(dict.fromkeys(v for v in (persistent_notes.get((code, name), ''), note) if v))
        require(row[c+1].text == name and row[c+2].text.rstrip('\n') == note, '姓名或備考與當日來源不符')
        for column, size in ((c, 12), (c+1, 12), (c+2, {23:9,35:8,43:8,60:8,61:8,62:8,63:8}.get(code,12))):
            font(row[column], size)
            mutable.add(row[column]._tc)

    groups = {label: [] for label in ('輪   休', '休(補)假', '公假')}
    for p in sorted(data['offDuty'], key=lambda p: p.get('ban') if isinstance(p.get('ban'), int) else 999):
        code, status = p.get('ban'), str(p.get('status') or '').strip()
        if not isinstance(code, int) or '停休' in status:
            continue
        if status in ('輪', '輪休'):
            if 1 <= code <= 20 or 50 <= code <= 53 or 60 <= code <= 63:
                groups['輪   休'].append((str(code), p['name']))
        elif status in ('公', '公假') or '公假' in status:
            groups['公假'].append((str(code), p['name']))
        elif status in ('優', '休', '補', '休假', '補休') or '假' in status:
            groups['休(補)假'].append((str(code), p['name']))
    for label, expected in groups.items():
        indices = section_rows(base, c+4, label, old)
        require(bool(indices), '假別分類標題缺漏')
        actual = []
        for i in indices:
            pair = (rows[i][c+3].text.strip(), rows[i][c+4].text.strip())
            if any(pair):
                actual.append(pair)
                font(rows[i][c+3], 12)
                font(rows[i][c+4], 12)
            mutable.update((rows[i][c+3]._tc, rows[i][c+4]._tc))
        require(actual == expected, '假別名單、分類或輪休排序不符')

    def hand(text):
        numbers = re.findall(r'\d+', text)
        return bool(numbers and re.fullmatch(r'[\d\s()（）,，、+＋/／\-－]+', text)
                    and all(1 <= int(n) <= 63 for n in numbers))

    for tr, original_tr in zip(table._tbl.tr_lst, base._tbl.tr_lst):
        require((etree.tostring(tr.trPr) if tr.trPr is not None else None)
                == (etree.tostring(original_tr.trPr) if original_tr.trPr is not None else None), '列高或列設定改變')
        require(tr.grid_before == original_tr.grid_before and tr.grid_after == original_tr.grid_after,
                '實體儲存格位置改變')
        require(len(tr.tc_lst) == len(original_tr.tc_lst), '實體儲存格數量改變')
        for cell, original in zip(tr.tc_lst, original_tr.tc_lst):
            require((etree.tostring(cell.tcPr) if cell.tcPr is not None else None)
                    == (etree.tostring(original.tcPr) if original.tcPr is not None else None),
                    '實體儲存格邊線、合併、欄寬或文字方向改變')

    checked_fixed = set()
    for i, row in enumerate(rows):
        time_row = bool(re.fullmatch(r'\d{1,2}[－\-–]\d{1,2}', old[i][0].text.strip()))
        for j, cell in enumerate(row):
            original = old[i][j]
            if time_row and 0 < j < c:
                text = cell.text.strip()
                require(not hand(text), '勤務手填代號未留白')
                if hand(original.text.strip()):
                    require(not text, '手填欄出現非來源內容')
                    mutable.add(cell._tc)
            pair = (cell._tc, original._tc)
            if cell._tc not in mutable and pair not in checked_fixed:
                # Paragraph XML includes the fixed text as well as its formatting.
                require([etree.tostring(p._p) for p in cell.paragraphs]
                        == [etree.tostring(p._p) for p in original.paragraphs], '模板固定文字或格式改變')
                checked_fixed.add(pair)
    # The check above covers the main duty table. Other tables and non-table
    # body content must remain byte-for-byte identical to the actual template.
    for out, original in zip(doc.tables[1:], source.tables[1:]):
        require(etree.tostring(out._tbl) == etree.tostring(original._tbl), '其他表格被修改')
    require([etree.tostring(p._p) for p in doc.paragraphs]
            == [etree.tostring(p._p) for p in source.paragraphs], '表格外文字被修改')
