#!/usr/bin/env python3
"""Insert a property row at the end of a given month block.

Rebuilds the sheet as a permutation of its rows (plus the new one) so Rate PM2
formulas and merged month headers are remapped, not left dangling.
"""
import sys, copy, re, json
import openpyxl
from openpyxl.styles import Font
sys.path.insert(0, sys.argv[1])
import update_workbook as U

WB = sys.argv[2]
BLUE, MONEY = U.BLUE, U.MONEY

# sheet, month index, [addr, price, bed, bath, car, size]
INSERTS = [
    ('Randwick UNIT', 10, ['2/15-17 Park Avenue', None, 2, 1, 1, 71]),
]

wb = openpyxl.load_workbook(WB)
log = []
for sheet, midx, vals in INSERTS:
    ws = wb[sheet]
    sh = U.Sheet(ws)
    blk = next(b for b in sh.blocks if b[1] == midx)
    tmpl = sh.tmpl_by_block.get(blk[0]) or sh.tmpl_prop
    anchor = blk[4]
    snap = U.snapshot(ws, sh.max_row, sh.max_col)

    order = []
    for r in range(1, sh.max_row + 1):
        order.append(r)
        if r == anchor:
            order.append('NEW')
    omap = {old: new for new, old in enumerate(order, start=1) if old != 'NEW'}
    new_row_at = order.index('NEW') + 1

    hdr_txt = re.sub(r'\(\d+\)', '(%d)' % (blk[3] + 1),
                     str(snap[blk[0]]['cells'][1]['v'])) if blk[3] is not None else None

    old_merges = [(m.min_row, m.min_col, m.max_row, m.max_col)
                  for m in list(ws.merged_cells.ranges)]
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    for new_r, old_r in enumerate(order, start=1):
        if old_r == 'NEW':
            for c in range(1, sh.max_col + 1):
                cell = ws.cell(row=new_r, column=c)
                t = snap[tmpl]['cells'][c]
                cell._style = copy.copy(t['style'])
                cell.value = vals[c - 1] if c <= len(vals) else None
                cell.number_format = MONEY if c in (2, 7) else t['fmt']
                cell.hyperlink = None
                cell.font = Font(name='Arial', size=19,
                                 color=(BLUE if c == 1 else 'FF000000'))
            continue
        s = snap[old_r]
        for c in range(1, sh.max_col + 1):
            cell = ws.cell(row=new_r, column=c)
            cd = s['cells'][c]
            v = cd['v']
            if isinstance(v, str) and v.startswith('='):
                v = U.remap_formula(v, omap)
            if c == 1 and old_r == blk[0] and hdr_txt:
                v = hdr_txt
            cell.value = v
            cell._style = copy.copy(cd['style'])
            cell.number_format = cd['fmt']
            cell.hyperlink = copy.copy(cd['link']) if cd['link'] else None
        if s['h'] is not None:
            ws.row_dimensions[new_r].height = s['h']

    for mnr, mnc, mxr, mxc in old_merges:
        if mnr == mxr and mnr in omap:
            ws.merge_cells(start_row=omap[mnr], start_column=mnc,
                           end_row=omap[mnr], end_column=mxc)
    log.append({'sheet': sheet, 'addr': vals[0], 'month': U.MONTHS[midx - 1],
                'new_row': new_row_at, 'header': hdr_txt, 'styled_after_row': tmpl})

wb.save(WB)
print(json.dumps(log, indent=1))
