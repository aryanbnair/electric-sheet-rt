"""Put back what openpyxl cannot carry through a save.

Usage:
    python postprocess_workbook.py RAW.xlsx ORIGINAL.xlsx OUT.xlsx [VALUES.xlsx]

1. Google Sheets "filter views" (<customSheetViews> per sheet and
   <customWorkbookViews> in workbook.xml) are copied from ORIGINAL verbatim.
2. Every formula cell gets a cached result again, so previews (phone, Drive,
   mail) show numbers instead of blanks.  Results come from VALUES.xlsx (a
   LibreOffice-recalculated copy of RAW) when given; where LibreOffice cannot
   evaluate a Google-only function (IMPORTRANGE, UNIQUE, ...) and the cell's
   formula is unchanged, the original Google result is kept instead.
   Excel and Google Sheets still recalculate everything on open
   (fullCalcOnLoad stays on), so these values are only a display cache.
"""
import datetime
import io
import json
import sys
import zipfile

import openpyxl
from lxml import etree
from openpyxl.utils.datetime import to_excel

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ERRORS = {"#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"}
# elements that must come AFTER customSheetViews in a worksheet (ECMA-376 order)
AFTER_CSV = ["mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks",
             "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
             "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
             "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]
AFTER_CWV = ["pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr",
             "webPublishObjects", "extLst"]


def sheet_parts(z):
    wbx = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for s in wbx.iter(f"{{{NS}}}sheet"):
        t = rid2t[s.get(f"{{{RNS}}}id")].lstrip("/")
        out[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
    return out


def _rels(z, part):
    import posixpath
    rp = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if rp not in z.namelist():
        return {}
    out = {}
    for r in etree.fromstring(z.read(rp)):
        t = r.get("Target")
        t = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), t))
        out[r.get("Id")] = (r.get("Type").rsplit("/", 1)[-1], t)
    return out


def drawing_of(z, sheet_part):
    if not sheet_part:
        return None
    for typ, target in _rels(z, sheet_part).values():
        if typ == "drawing":
            return target
    return None


def images_of(z, drawing_part):
    """{relationship id: image bytes} of a drawing part."""
    return {rid: z.read(t) for rid, (typ, t) in _rels(z, drawing_part).items() if typ == "image"}


def insert_before(parent, new_el, after_names):
    for i, child in enumerate(parent):
        if etree.QName(child).localname in after_names:
            parent.insert(i, new_el)
            return
    parent.append(new_el)


def load_values(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    vals = {}
    for ws in wb.worksheets:
        d = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    d[c.coordinate] = c.value
        vals[ws.title] = d
    wb.close()
    return vals


def is_err(v):
    return isinstance(v, str) and v in ERRORS


def set_cached(c_el, value):
    for v in c_el.findall(f"{{{NS}}}v"):
        c_el.remove(v)
    for k in ("t",):
        if k in c_el.attrib:
            del c_el.attrib[k]
    if value is None:
        return
    v = etree.SubElement(c_el, f"{{{NS}}}v")
    if isinstance(value, bool):
        c_el.set("t", "b")
        v.text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        v.text = repr(float(value)) if isinstance(value, float) else str(value)
    elif isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        v.text = repr(float(to_excel(value)))
    elif is_err(value):
        c_el.set("t", "e")
        v.text = value
    else:
        c_el.set("t", "str")
        v.text = str(value)


def main(raw, original, out, values=None):
    zr, zo = zipfile.ZipFile(raw), zipfile.ZipFile(original)
    parts_r, parts_o = sheet_parts(zr), sheet_parts(zo)
    orig_vals = load_values(original)
    lo_vals = load_values(values) if values else {}
    orig_f = {}  # formulas of the original, to know which formulas are unchanged
    wbo = openpyxl.load_workbook(original, read_only=True)
    for ws in wbo.worksheets:
        orig_f[ws.title] = {c.coordinate: (c.value.text if hasattr(c.value, "text") else c.value)
                            for row in ws.iter_rows() for c in row if c.value is not None}
    wbo.close()

    new_parts = {}
    stats = {"filter_views": 0, "cached_lo": 0, "cached_orig": 0, "no_value": 0, "lo_err_kept_google": 0}
    for name, part in parts_r.items():
        root = etree.fromstring(zr.read(part))
        # 1. filter views
        if name in parts_o:
            oroot = etree.fromstring(zo.read(parts_o[name]))
            csv = oroot.find(f"{{{NS}}}customSheetViews")
            if csv is not None and root.find(f"{{{NS}}}customSheetViews") is None:
                insert_before(root, csv, AFTER_CSV)
                stats["filter_views"] += 1
        # 2. cached values
        lov, ov, of = lo_vals.get(name, {}), orig_vals.get(name, {}), orig_f.get(name, {})
        for c_el in root.iter(f"{{{NS}}}c"):
            f_el = c_el.find(f"{{{NS}}}f")
            if f_el is None:
                continue
            ref = c_el.get("r")
            ftext = "=" + (f_el.text or "")
            unchanged = ref in of and isinstance(of[ref], str) and of[ref] == ftext
            lo = lov.get(ref)
            # LibreOffice cannot evaluate some Google functions (UNIQUE, IMPORTRANGE): where it gives an
            # error but Google had a real result, keep Google's result as the display cache.
            if ref in lov and not (is_err(lo) and ref in ov and not is_err(ov.get(ref))):
                set_cached(c_el, lo)
                stats["cached_lo"] += 1
            elif ref in ov and (unchanged or ref in lov):
                set_cached(c_el, ov[ref])
                stats["cached_orig"] += 1
                if ref in lov:
                    stats["lo_err_kept_google"] += 1
            else:
                set_cached(c_el, None)
                stats["no_value"] += 1
        new_parts[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 3. pictures: openpyxl rewrites drawing XML with an element in the wrong namespace (Excel then
    #    "repairs" it). Put the original drawing XML back where the pictures are byte-identical.
    stats["drawings_restored"] = 0
    for name, part in parts_r.items():
        d_r, d_o = drawing_of(zr, part), drawing_of(zo, parts_o.get(name)) if name in parts_o else None
        if not d_r or not d_o:
            continue
        img_r, img_o = images_of(zr, d_r), images_of(zo, d_o)
        assert img_r == img_o, f"pictures of {name} differ - not restoring drawing"
        new_parts[d_r] = zo.read(d_o)
        stats["drawings_restored"] += 1

    # workbook-level custom views
    wbr = etree.fromstring(zr.read("xl/workbook.xml"))
    cwv = etree.fromstring(zo.read("xl/workbook.xml")).find(f"{{{NS}}}customWorkbookViews")
    if cwv is not None and wbr.find(f"{{{NS}}}customWorkbookViews") is None:
        insert_before(wbr, cwv, AFTER_CWV)
    new_parts["xl/workbook.xml"] = etree.tostring(wbr, xml_declaration=True, encoding="UTF-8", standalone=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zw:
        for item in zr.infolist():
            data = new_parts.get(item.filename, zr.read(item.filename))
            zw.writestr(item, data)
    open(out, "wb").write(buf.getvalue())
    print(json.dumps(stats))


if __name__ == "__main__":
    main(*sys.argv[1:])
