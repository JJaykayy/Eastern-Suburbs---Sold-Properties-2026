# Eastern Suburbs - Sold Properties 2026

Weekly RP Data exports are merged into the master workbook
`Eastern Suburbs - Sold Properties 2026.xlsx`. The scheduled routine's own
prompt is the specification; this file is repository orientation only and does
not override it.

## Which file is the export

An export is a **`.numbers`, `.csv`, `.tsv` or `.xlsx` file that is not the
master workbook and has no `processed-` prefix**. Run `python3 find_export.py`
rather than guessing: it prints the unprocessed exports oldest first and exits
1 when there are none.

Everything else in the root is tooling, not an export, even though it has no
`processed-` prefix: `update_workbook.py`, `build_summary.py`,
`find_export.py`, `clear_white_fill.py`, `move_rows.py`, `insert_row.py`,
`CLAUDE.md`, `.gitignore`, `__pycache__/`.

Never delete an export. Rename it with a `processed-` prefix when done.

## Scripts

| Script | Purpose |
| --- | --- |
| `update_workbook.py` | The weekly merge. Use it; do not re-derive the logic. Rebuilds the Summary tab at the end of an `--apply` run. |
| `build_summary.py` | Rebuilds the `Summary` tab. Also runs standalone. |
| `find_export.py` | Locates the unprocessed export(s). |
| `clear_white_fill.py` | One-off: strips solid white fills that hide the gridlines. |
| `move_rows.py`, `insert_row.py` | One-off corrections, kept as the audit trail of past fixes. |

Usage: `python3 update_workbook.py <export> --apply --report report.json`
(omit `--apply` for a dry run).

Dependencies: `openpyxl` and `numbers-parser`. If missing, install with
`python3 -m pip install --ignore-installed openpyxl numbers-parser` -- plain
`pip install` may target a different interpreter than `python3`.

## Sheets

One sheet per suburb and type (`<Suburb> HOUSE` / `<Suburb> UNIT`; naming is
inconsistent, e.g. `Maroubra - HOUSE`), plus `Summary` last. Rows sit in month
blocks whose header is recognised by its green `#93C47D` fill, not its text:
eight sheets carry a count-less `August`, and `Woollahra HOUSE` row 2 reads
`Januray (1)`. Both are real headers -- leave the text alone.

`Suburb Details` and `Vacancy` are off limits: never read, write or create
them. Neither has ever existed here.

## Things that have bitten before

- A blank Sale Price is deliberate: the price is undisclosed. The late-fill
  rule only writes into blank cells, so never write `$0`/`$1` placeholders.
- A property can hold two rows in one year when it genuinely sold twice
  (25 Epping Road, Double Bay; 2/15-17 Park Avenue, Randwick). Matching on
  address alone has merged two sales into one before.
- New rows copy a template row's style. Keep data rows free of any solid fill:
  green reads as a month header, white hides the gridlines.
