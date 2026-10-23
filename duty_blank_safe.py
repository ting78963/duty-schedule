"""Blank staff assignments in known Word duty layouts, preserving vehicle identifiers."""
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
            # The 37-column layout also stores concatenated personnel codes
            # (e.g. 951 = 9/5/1). Only clear known, verified positions
            # and values; ordinary 3-4 digit vehicle IDs remain intact.
            combined_personnel = (width == 37 and row <= 39
                                  and start not in (4, 5, 6, 7)
                                  and value in ('951', '852', '1453', '1016', '810'))
            if value and (PERSONNEL.fullmatch(value) or combined_personnel):
                targets.append(nodes)
    for nodes in targets:
        for node in nodes:
            node.text = ''
    return len(targets)
