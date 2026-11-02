# electric-sheet-rt

Tools that reorganise the **X3 Electric** workbook (PCBs, parts, stock, pinouts,
HV commissioning) safely and prove that nothing was lost.

> The workbook itself is **not** committed: this repository is public and the
> workbook contains Vector licence keys and links to internal Google Sheets.
> `*.xlsx` is git-ignored on purpose.

## What the reorganisation does

* **New tabs**
  * **Start**: guide to every tab (what it is for, where to edit, what it is
    linked with), how the PCB numbers flow, and every dropdown's allowed values,
    with click-through links.
  * **Dashboard**: live key numbers, PCB build status per board, design-status
    and physical-board-status breakdowns, and **health checks** that turn red
    when a formula breaks.
  * **Review Notes**: every formula that was changed (old and new formula, reason,
    effect) and a prioritised list of open issues that need a human decision.
* **Tab order and colours**: Overview, PCBs, Stock & parts, Wiring & pinouts,
  HV & accumulator, Admin, Archive. Hidden tabs stay hidden.
* **Small usability tweaks**: frozen header rows on 5 tabs, and header styling
  on 2 plain header rows.
* **Formula fixes** (74 cells), limited to clear copy-paste slips, for example
  references to the wrong row, swapped board columns, a total that skipped a
  board, #REF! cells, and SUMIF windows that slid down one row per line. Each fix
  first asserts that the cell still holds exactly the expected broken formula.

Nothing else changes. No typed value is edited, no tab is renamed or deleted,
and no row or column is inserted or removed.

## Usage

```bash
python tools/build_improved_workbook.py ORIGINAL.xlsx raw.xlsx changes.json
```

Optional: recalculate a copy with LibreOffice (slow, 20+ min) to refresh the stored results:

```bash
python <xlsx-skill>/scripts/recalc.py lo.xlsx 7000
```

```bash
python tools/postprocess_workbook.py raw.xlsx ORIGINAL.xlsx FINAL.xlsx [lo.xlsx]
```

```bash
python tools/verify_workbook.py ORIGINAL.xlsx FINAL.xlsx changes.json
```

`postprocess_workbook.py` restores what openpyxl drops on save: the saved
Google Sheets filter views and the stored formula results, so previews show
numbers. Excel and Google Sheets still recalculate everything when the file is
opened.

`verify_workbook.py` compares every cell (value, formula, style) and every
sheet feature (merged cells, dropdowns, conditional formatting, tables, images,
hyperlinks, frozen panes, column widths, row heights, hidden state, filters,
named ranges, filter views). It exits non-zero on any difference that is not
listed in `changes.json`.
