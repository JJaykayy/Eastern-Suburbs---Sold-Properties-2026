#!/usr/bin/env python3
"""Identify the unprocessed RP Data export(s) in the repository root.

An export is a .numbers / .csv / .tsv / .xlsx file that is not the master
workbook and does not carry the "processed-" prefix. Everything else in the
root -- the .py tooling, .gitignore, __pycache__ -- is not an export, however
the weekly instructions are worded.

Prints one path per line, oldest first, so several unprocessed exports are
handled in chronological order. Exits 1 when there is nothing to process.

    python3 find_export.py
"""
import os, re, sys

import update_workbook as U

EXPORT_EXT = ('.numbers', '.csv', '.tsv', '.xlsx', '.xlsm')


def find(root='.'):
    out = []
    for name in os.listdir(root):
        if os.path.isdir(os.path.join(root, name)):
            continue
        if name == os.path.basename(U.WB_PATH):
            continue
        if name.startswith('processed-') or name.startswith('.'):
            continue
        if not name.lower().endswith(EXPORT_EXT):
            continue
        out.append(name)
    # sort by the timestamp in the filename where there is one, else by name
    def key(n):
        m = re.search(r'(\d{8,14})', n)
        return (0, m.group(1)) if m else (1, n.lower())
    return sorted(out, key=key)


if __name__ == '__main__':
    found = find(os.path.dirname(os.path.abspath(U.WB_PATH)) or '.')
    for f in found:
        print(f)
    if not found:
        print('no unprocessed export in the repository root', file=sys.stderr)
        sys.exit(1)
