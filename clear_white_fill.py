#!/usr/bin/env python3
"""Strip solid white backgrounds from the workbook's data cells.

A solid white fill paints over the gridlines, so the cell renders as a
seamless blank gap rather than an empty bordered cell -- most visible in the
Sale Price column, where a price is often not yet disclosed. Nothing in this
workbook uses cell borders; the lines are gridlines showing through cells that
carry no fill.

Month headers (#93C47D) are left alone, as is every value, formula, font,
number format, width and merge. Formatting only.

    python3 clear_white_fill.py [workbook]
"""
import sys
import openpyxl
from openpyxl.styles import PatternFill

import update_workbook as U

WHITE = {'FFFFFFFF', '00FFFFFF', 'FFFFFF'}


def is_white(cell):
    f = cell.fill
    return (f is not None and f.patternType == 'solid'
            and f.fgColor is not None and f.fgColor.rgb in WHITE)


def clear(path=None):
    path = path or U.WB_PATH
    wb = openpyxl.load_workbook(path)
    hit = {}
    for name in wb.sheetnames:
        if name.strip().lower() in U.OFF_LIMITS:
            continue
        ws = wb[name]
        n = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                max_col=max(7, ws.max_column)):
            for cell in row:
                if U.is_green(cell):          # month headers keep their fill
                    continue
                if is_white(cell):
                    cell.fill = PatternFill(fill_type=None)
                    n += 1
        if n:
            hit[name] = n
    wb.save(path)
    return hit


if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else U.WB_PATH
    hit = clear(path)
    print('cleared %d white-filled cells across %d sheets'
          % (sum(hit.values()), len(hit)))
    for k, v in sorted(hit.items(), key=lambda x: -x[1]):
        print('  %-22s %d' % (k, v))
