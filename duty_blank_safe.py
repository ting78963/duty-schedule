"""Clear personnel assignment codes only; retain vehicle identifiers."""
import re
from docx.oxml.ns import qn

PERSONNEL = re.compile(r'^\s*\(?\d{1,2}\)?(?:[\s、,，/＋+]+\(?\d{1,2}\)?)*\s*$')

def blank_duty_document(doc):
    if not doc.tables:
        raise ValueError('勤務表沒有主表格')
    table = doc.tables[0]
    width = len(table.columns)
    if len(table.rows) != 64 or width not in (33, 35, 37):
        raise ValueError(f'未驗證的勤務表版型: {len(table.rows)}列 {width}欄')
    boundary = width - 5
    targets = []
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
            value = ''.join(node.text or '' for node in nodes).strip()
            # Vehicle identifiers (3+ digits, alphanumeric codes, arrows)
            # are intentionally left unchanged.
            if value and PERSONNEL.fullmatch(value):
                targets.append(nodes)
    for nodes in targets:
        for node in nodes:
            node.text = ''
    return len(targets)
