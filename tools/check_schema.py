"""Validate every sheet/table/drawing XML part against the ISO OOXML schemas (pass xlsx paths)."""
import sys, zipfile, re, collections
from lxml import etree
import os
SD = os.environ.get("OOXML_SCHEMAS", "/mnt/skills/public/xlsx/scripts/office/schemas/ISO-IEC29500-4_2016/")
sml = etree.XMLSchema(etree.parse(SD + "sml.xsd"))
dr = etree.XMLSchema(etree.parse(SD + "dml-spreadsheetDrawing.xsd"))
def check(path):
    z = zipfile.ZipFile(path); res = collections.OrderedDict()
    for n in z.namelist():
        if not n.endswith(".xml") or "_rels" in n or "Content_Types" in n: continue
        sch = dr if "/drawings/" in n else sml if re.search(r"xl/(worksheets|tables)/|xl/(workbook|styles|sharedStrings)\.xml", n) else None
        if sch is None: continue
        doc = etree.fromstring(z.read(n))
        # strip mc:Ignorable extension namespaces the strict-ish xsd does not know
        ok = sch.validate(etree.ElementTree(doc))
        errs = [] if ok else sorted({re.sub(r"line \d+", "", str(e.message))[:180] for e in sch.error_log if "XML/1998/namespace}space" not in str(e.message)})
        res[n] = errs
    return res
for p in sys.argv[1:]:
    r = check(p); bad = {k: v for k, v in r.items() if v}
    print(f"== {p}: {len(r)} parts checked, {len(bad)} with schema errors")
    for k, v in bad.items(): print("  ", k, v[:3])
