"""Conservative daily assignment blanking for validated duty Word layouts."""
import re
import json
from docx.oxml.ns import qn

STAFF = re.compile(r'^\s*\(?\d{1,2}\)?(?:[\s、,，/＋+]+\(?\d{1,2}\)?)*\s*$')
COMBINED = re.compile(r'^\d{3,4}
VEHICLE = re.compile(r'^(?:\d{3,}|[A-Z]{2,}-?\d{3,})$')

def blank_duty_document(doc):
    if not doc.tables:
        raise ValueError('勤務表沒有主表格')
    table = doc.tables[0]
    width = len(table.columns)
    if len(table.rows) != 64 or width not in (33, 35, 37):
        raise ValueError('未驗證的勤務表版型')
    boundary = width - 5
    changes = []
    review = []
    for row, tr in enumerate(table._tbl.tr_lst):
        if not 15 <= row <= 62:
            continue
        col = 0
        for tc in tr.tc_lst:
            span = tc.tcPr.gridSpan if tc.tcPr is not None else None
            n = int(span.val) if span is not None else 1
            start = col
            col += n
            if start == 0 or start >= boundary or col > boundary:
                continue
            nodes = list(tc.iter(qn('w:t')))
            val = ''.join(node.text or '' for node in nodes).strip()
            is_combined = (width == 37 and row <= 39 and start not in (4, 5, 6, 7)
                           and COMBINED.fullmatch(val) is not None)
            if not val or (VEHICLE.fullmatch(val) and not is_combined):
                continue
            if STAFF.fullmatch(val) or is_combined or DUTY_VEHICLE.fullmatch(val):
                changes.append(nodes)
            elif re.search(r'\d', val):
                review.append([row, start, val[:80]])
    if review:
        raise ValueError('勤務區含無法判斷的數字: ' + json.dumps(review, ensure_ascii=False))
    for nodes in changes:
        for node in nodes:
            node.text = ''
    return len(changes)
)
# Vehicle identifiers with direction arrows also belong to manually entered duty cells.
DUTY_VEHICLE = re.compile(r'^(?:[A-Z]{1,4}-?\d{1,4}[↑↓←→]?|[↑↓←→]\d{2,4})
VEHICLE = re.compile(r'^(?:\d{3,}|[A-Z]{2,}-?\d{3,})$')

def blank_duty_document(doc):
    if not doc.tables:
        raise ValueError('勤務表沒有主表格')
    table = doc.tables[0]
    width = len(table.columns)
    if len(table.rows) != 64 or width not in (33, 35, 37):
        raise ValueError('未驗證的勤務表版型')
    boundary = width - 5
    changes = []
    review = []
    for row, tr in enumerate(table._tbl.tr_lst):
        if not 15 <= row <= 62:
            continue
        col = 0
        for tc in tr.tc_lst:
            span = tc.tcPr.gridSpan if tc.tcPr is not None else None
            n = int(span.val) if span is not None else 1
            start = col
            col += n
            if start == 0 or start >= boundary or col > boundary:
                continue
            nodes = list(tc.iter(qn('w:t')))
            val = ''.join(node.text or '' for node in nodes).strip()
            is_combined = (width == 37 and row <= 39 and start not in (4, 5, 6, 7)
                           and COMBINED.fullmatch(val) is not None)
            if not val or (VEHICLE.fullmatch(val) and not is_combined):
                continue
            if STAFF.fullmatch(val) or is_combined:
                changes.append(nodes)
            elif re.search(r'\d', val):
                review.append([row, start, val[:80]])
    if review:
        raise ValueError('勤務區含無法判斷的數字: ' + json.dumps(review, ensure_ascii=False))
    for nodes in changes:
        for node in nodes:
            node.text = ''
    return len(changes)
)
VEHICLE = re.compile(r'^(?:\d{3,}|[A-Z]{2,}-?\d{3,})$')

def blank_duty_document(doc):
    if not doc.tables:
        raise ValueError('勤務表沒有主表格')
    table = doc.tables[0]
    width = len(table.columns)
    if len(table.rows) != 64 or width not in (33, 35, 37):
        raise ValueError('未驗證的勤務表版型')
    boundary = width - 5
    changes = []
    review = []
    for row, tr in enumerate(table._tbl.tr_lst):
        if not 15 <= row <= 62:
            continue
        col = 0
        for tc in tr.tc_lst:
            span = tc.tcPr.gridSpan if tc.tcPr is not None else None
            n = int(span.val) if span is not None else 1
            start = col
            col += n
            if start == 0 or start >= boundary or col > boundary:
                continue
            nodes = list(tc.iter(qn('w:t')))
            val = ''.join(node.text or '' for node in nodes).strip()
            is_combined = (width == 37 and row <= 39 and start not in (4, 5, 6, 7)
                           and COMBINED.fullmatch(val) is not None)
            if not val or (VEHICLE.fullmatch(val) and not is_combined):
                continue
            if STAFF.fullmatch(val) or is_combined:
                changes.append(nodes)
            elif re.search(r'\d', val):
                review.append([row, start, val[:80]])
    if review:
        raise ValueError('勤務區含無法判斷的數字: ' + json.dumps(review, ensure_ascii=False))
    for nodes in changes:
        for node in nodes:
            node.text = ''
    return len(changes)
