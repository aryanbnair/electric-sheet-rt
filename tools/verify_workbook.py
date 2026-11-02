"""Prove that an improved workbook did not lose or alter anything by accident.

Usage:
    python verify_workbook.py ORIGINAL.xlsx IMPROVED.xlsx [changes.json]

Compares, sheet by sheet (matched by name, so reordering tabs is fine):
  * every cell value / formula / array-formula (text and range)
  * cell styles (font, fill, border, alignment, number format, protection)
  * merged cells, data validations, conditional formatting, tables,
    images, hyperlinks, freeze panes, column widths/hidden, row heights/hidden,
    sheet visibility, auto-filter, sheet-scoped defined names
  * Google "filter views" (customSheetViews) and workbook custom views (raw XML)

Anything that differs must be listed in changes.json (written by the build
script) or the check fails.  Exit code 0 = no unexpected differences.
"""
import json
import re
import sys
import zipfile

import openpyxl


def cell_val(v):
    if hasattr(v, "text") and hasattr(v, "ref"):  # ArrayFormula
        return ("ARRAY", v.ref, v.text)
    if hasattr(v, "__iter__") and not isinstance(v, str) and type(v).__name__ == "CellRichText":
        return ("RICH", str(v), repr(v))
    return v


def _norm(obj):
    # openpyxl writes an explicit False back as "unset"; both mean the same in Excel/Sheets.
    s = re.sub(r"=False\b", "=None", repr(obj))
    return s.replace("vertical='bottom'", "vertical=None")  # bottom is the default


def style_key(c):
    return (
        _norm(c.font), _norm(c.fill), _norm(c.border), _norm(c.alignment),
        c.number_format, _norm(c.protection),
    )


def sheet_features(ws):
    f = {}
    f["state"] = ws.sheet_state
    f["freeze"] = ws.freeze_panes
    f["merged"] = sorted(str(r) for r in ws.merged_cells.ranges)
    f["dv"] = sorted(
        (str(d.sqref), d.type, d.operator, d.formula1, d.formula2, d.allow_blank,
         d.showDropDown, d.showErrorMessage, d.error, d.prompt)
        for d in ws.data_validations.dataValidation
    )
    cf = []
    for rng in ws.conditional_formatting:
        for r in rng.rules:
            dx = r.dxf
            cf.append((str(rng.sqref), r.type, r.operator, tuple(r.formula or ()), r.text,
                       r.priority, r.stopIfTrue,
                       repr(dx.font) if dx else None, repr(dx.fill) if dx else None))
    f["cf"] = sorted(cf, key=repr)
    f["tables"] = sorted((t.name, t.displayName, t.ref) for t in ws.tables.values())
    f["images"] = sorted(
        (repr(getattr(i.anchor, "_from", None).__dict__ if hasattr(i.anchor, "_from") else i.anchor),
         i.width, i.height) for i in ws._images
    )
    f["autofilter"] = ws.auto_filter.ref
    f["cols"] = sorted(
        (k, d.width, d.hidden, d.min, d.max, d.outlineLevel) for k, d in ws.column_dimensions.items()
    )
    f["rows"] = sorted(
        (k, d.height, d.hidden, d.outlineLevel) for k, d in ws.row_dimensions.items()
        if d.height is not None or d.hidden or d.outlineLevel
    )
    f["defined_names"] = sorted((n, d.attr_text) for n, d in ws.defined_names.items())
    f["hyperlinks"] = sorted(
        (c.coordinate, c.hyperlink.target, c.hyperlink.location)
        for row in ws.iter_rows() for c in row if c.hyperlink
    )
    return f


def raw_xml_by_sheet(path):
    """Return {sheet name: raw customSheetViews xml} and the workbook customWorkbookViews xml."""
    z = zipfile.ZipFile(path)
    wbx = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid2target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    rid2target.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)})
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wbx):
        tag = m.group(0)
        name = re.search(r'name="([^"]*)"', tag).group(1)
        rid = re.search(r'r:id="([^"]*)"', tag).group(1)
        target = rid2target[rid].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        x = z.read(target).decode("utf8")
        csv = re.search(r"<customSheetViews>.*?</customSheetViews>", x, re.S)
        out[name.replace("&amp;", "&")] = csv.group(0) if csv else None
    cwv = re.search(r"<customWorkbookViews>.*?</customWorkbookViews>", wbx, re.S)
    return out, (cwv.group(0) if cwv else None)


def main(orig_path, new_path, changes_path=None):
    allowed_cells, allowed_styles, new_sheets, allowed_features = set(), set(), set(), set()
    if changes_path:
        ch = json.load(open(changes_path, encoding="utf8"))
        allowed_cells = {(c["sheet"], c["cell"]) for c in ch.get("cells", [])}
        allowed_styles = {tuple(x) for x in ch.get("styled_cells", [])}
        new_sheets = set(ch.get("new_sheets", []))
        allowed_features = {tuple(x) for x in ch.get("features", [])}

    a = openpyxl.load_workbook(orig_path, rich_text=True)
    b = openpyxl.load_workbook(new_path, rich_text=True)
    problems = []
    a_names, b_names = set(a.sheetnames), set(b.sheetnames)
    if a_names - b_names:
        problems.append(f"MISSING SHEETS: {sorted(a_names - b_names)}")
    extra = b_names - a_names - new_sheets
    if extra:
        problems.append(f"UNEXPECTED NEW SHEETS: {sorted(extra)}")

    n_cells = n_formulas = 0
    used_cells = set()
    for name in a.sheetnames:
        if name not in b_names:
            continue
        wa, wb_ = a[name], b[name]
        max_r = max(wa.max_row, wb_.max_row)
        max_c = max(wa.max_column, wb_.max_column)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                ca, cb = wa.cell(r, c), wb_.cell(r, c)
                va, vb = cell_val(ca.value), cell_val(cb.value)
                if va is not None:
                    n_cells += 1
                    if isinstance(ca.value, str) and ca.value.startswith("=") or isinstance(va, tuple):
                        n_formulas += 1
                if va != vb:
                    if (name, ca.coordinate) in allowed_cells:
                        used_cells.add((name, ca.coordinate))
                    else:
                        problems.append(f"[{name}] {ca.coordinate}: {va!r} -> {vb!r}")
                if style_key(ca) != style_key(cb) and (name, ca.coordinate) not in allowed_styles:
                    problems.append(f"[{name}] {ca.coordinate}: STYLE changed")
        fa, fb = sheet_features(wa), sheet_features(wb_)
        for k in fa:
            if fa[k] != fb[k] and (name, k) not in allowed_features:
                problems.append(f"[{name}] feature '{k}' differs:\n    was {fa[k]!r}\n    now {fb[k]!r}")

    va_, vwa = raw_xml_by_sheet(orig_path)
    vb_, vwb = raw_xml_by_sheet(new_path)
    for name, x in va_.items():
        if vb_.get(name) != x:
            problems.append(f"[{name}] Google filter views (customSheetViews) differ")
    if vwa != vwb:
        problems.append("workbook customWorkbookViews differ")

    never_used = allowed_cells - used_cells
    print(f"Compared {len(a.sheetnames)} original sheets, {n_cells} non-empty cells ({n_formulas} formulas).")
    print(f"Intended cell changes applied: {len(used_cells)} / {len(allowed_cells)}")
    if never_used:
        problems.append(f"Changes listed but not found in output: {sorted(never_used)}")
    if problems:
        print(f"\n{len(problems)} UNEXPECTED DIFFERENCES:")
        for p in problems[:300]:
            print("  " + p)
        sys.exit(1)
    print("OK - no unexpected differences.")


if __name__ == "__main__":
    main(*sys.argv[1:])
