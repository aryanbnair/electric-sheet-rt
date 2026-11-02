"""Check rules Excel enforces on open (tables vs filters/merges, header names, ...). Pass xlsx paths."""
import sys, openpyxl, collections
from openpyxl.worksheet.cell_range import CellRange
wb = openpyxl.load_workbook(sys.argv[1]); probs = []
names = collections.Counter()
for ws in wb.worksheets:
    af = ws.auto_filter.ref
    tabs = list(ws.tables.values())
    for i, t in enumerate(tabs):
        names[t.name.lower()] += 1; names[t.displayName.lower()] += 0
        r = CellRange(t.ref)
        if af and not CellRange(af.replace("$", "")).isdisjoint(r): probs.append(f"{ws.title}: table {t.name} overlaps sheet filter {af}")
        for m in ws.merged_cells.ranges:
            if not m.isdisjoint(r): probs.append(f"{ws.title}: table {t.name} overlaps merged {m}")
        for u in tabs[i+1:]:
            if not CellRange(u.ref).isdisjoint(r): probs.append(f"{ws.title}: tables {t.name}/{u.name} overlap")
        if t.headerRowCount != 0:
            hdr = [ws.cell(r.min_row, c).value for c in range(r.min_col, r.max_col + 1)]
            cols = [c.name for c in t.tableColumns]
            if [str(h) if h is not None else None for h in hdr] != cols:
                probs.append(f"{ws.title}: table {t.name} header cells {hdr[:5]} != columns {cols[:5]}")
        if len(t.tableColumns) != r.max_col - r.min_col + 1: probs.append(f"{ws.title}: {t.name} column count")
    for dv in ws.data_validations.dataValidation:
        if dv.formula1 and len(dv.formula1) > 255: probs.append(f"{ws.title}: validation {dv.sqref} formula >255 chars")
    for m in ws.merged_cells.ranges:
        for n in ws.merged_cells.ranges:
            if m is not n and not m.isdisjoint(n): probs.append(f"{ws.title}: merged overlap {m} {n}")
    if len(ws.title) > 31: probs.append(f"sheet name too long {ws.title}")
probs += [f"duplicate table name {n}" for n, k in names.items() if k > 1]
print(sys.argv[1], "->", len(probs), "problems"); [print("  ", p) for p in probs[:30]]
