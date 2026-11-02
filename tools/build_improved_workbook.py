"""Build the reorganised X3-Electric workbook from the original export.

Usage:
    python build_improved_workbook.py ORIGINAL.xlsx OUT_RAW.xlsx CHANGES.json

What it does (and nothing else):
  1. Applies a small list of formula fixes.  Each fix first asserts that the
     cell still contains exactly the broken formula, so it can never overwrite
     something it does not expect.
  2. Adds three new tabs: Start (guide + links), Dashboard (live KPIs and
     health checks) and Review Notes (every change + open questions).
  3. Reorders tabs into groups and colours them.  Hidden tabs stay hidden.
  4. Small usability tweaks: frozen header rows on a few tabs and a header
     style on two plain header rows.
  5. Writes CHANGES.json, which verify_workbook.py uses to prove that nothing
     outside this list changed.

No data cell (anything typed by a person) is modified.
"""
import datetime
import json
import re
import sys

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.hyperlink import Hyperlink

TODAY = datetime.date.today().isoformat()
FONT = "Verdana"  # the workbook's existing default font

# ---------------------------------------------------------------- palette
NAVY = "1F3864"
GROUPS = {
    # key: (label, tab colour, light tint for cells)
    "overview": ("Overview", "1F3864", "D9E1F2"),
    "pcb": ("PCBs", "2E75B6", "DDEBF7"),
    "stock": ("Stock & parts", "548235", "E2EFDA"),
    "wiring": ("Wiring & pinouts", "C55A11", "FCE4D6"),
    "hv": ("HV & accumulator", "C00000", "F8D7DA"),
    "admin": ("Admin", "7F7F7F", "EDEDED"),
    "archive": ("Archive (hidden)", "BFBFBF", "F2F2F2"),
}
GOOD_FILL, GOOD_FONT = "C6EFCE", "006100"
BAD_FILL, BAD_FONT = "FFC7CE", "9C0006"
WARN_FILL, WARN_FONT = "FFEB9C", "9C5700"

thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def f(size=10, bold=False, color="000000", italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


# ---------------------------------------------------------------- tab plan
# (sheet, group, purpose, where to edit, linked with)
TABS = [
    ("Start", "overview", "This page. Guide to the workbook with links to every tab.",
     "Nothing to edit.", "-"),
    ("Dashboard", "overview",
     "Live summary: PCB build status, design/board status counts, stock shortages, HV commissioning "
     "progress and automatic health checks that flag broken formulas.",
     "Nothing to edit; everything is calculated.", "Reads Platinen, bestückte Platinen, Platinen Teile, RT Stock, HV Commisioning, Connectors Tracking"),
    ("Review Notes", "overview",
     "Every change made during the reorganisation (with the old formula), plus open issues that need a decision.",
     "Tick off / delete issues as you resolve them.", "-"),
    ("Platinen", "pcb",
     "Master list of all PCBs (rows 2-30) and bought devices (rows 31-38): quantity in car, ordered, "
     "assembled, owner, location, design status, version, layers, weight. Totals and status summary in rows 39-49.",
     "Columns B-F and H-W. Column G (Assembled) is calculated from 'bestückte Platinen' for own PCBs.",
     "Feeds bestückte Platinen, Platinen Teile, Platinen Tracking, Stecker, Std. Bauteile"),
    ("bestückte Platinen", "pcb",
     "One column per PCB type, one row per physical board (E16-1 ... E16-40). Choose each board's state from the dropdown.",
     "Rows 6-45 (dropdown). Rows 2-5 and 46-50 are calculated.",
     "Reads Platinen (names, quantities); feeds Platinen column G and Platinen Tracking"),
    ("Platinen Tracking", "pcb",
     "Test log per physical board: coating, test plan, test status, documentation and problems. "
     "Column C mirrors the state in 'bestückte Platinen'.",
     "Columns D-I. Column C is linked - change the state in 'bestückte Platinen'.",
     "Reads bestückte Platinen and Platinen"),
    ("Platinen Teile", "pcb",
     "BOM matrix: parts (rows) x boards (columns C-AF). Row 6 = boards still to build. "
     "AG total required, AH in stock, AI ordered, AJ surplus (negative = shortage), AK used, AL/AM storage positions.",
     "Part rows: package (A), value/part no. (B) and quantity per board (C-AF). AG-AM are calculated.",
     "Reads Platinen, bestückte Platinen, RT Stock; feeds RT Stock and AlTa HVM BOM"),
    ("RT Stock", "stock",
     "Inventory of the RT stock: description, package, part number, supplier, positions and counts in RT and GT stock. "
     "M-Q calculate initial, used, remaining, needed and surplus.",
     "Columns A-L and R-T. M-Q are calculated.",
     "Reads Platinen Teile; feeds Platinen Teile"),
    ("GT Stock", "stock",
     "Read-only copy of the GT stock list, imported live (IMPORTRANGE) from the GT stock Google Sheet.",
     "Do not edit - edit the source Google Sheet instead.",
     "Imported from another Google Sheet"),
    ("Connectors Tracking", "stock",
     "Harness connectors per connector name, summed per part number with +20 % spares, purchase status and cost. "
     "Kistler connector legend in columns N-R.",
     "Columns A-D (connectors), I-K (status, supplier, price).", "-"),
    ("Stecker", "stock",
     "Older connector list per board (X2 layout). Quantity totals are currently broken (#REF!) - see Review Notes.",
     "Use with care.", "Reads Platinen"),
    ("Std. Bauteile", "stock",
     "Older standard-part list per board (X2 layout). Quantity totals are currently broken (#REF!) - see Review Notes.",
     "Use with care.", "Reads Platinen"),
    ("Pinouts", "wiring",
     "Pinout of every harness connector, grouped by system (AS, LVS, HV, COM, HMI, MS, SDC, misc., splices) "
     "with test status per pin; CAN placement and Ethernet convention on the right.",
     "Everywhere.", "-"),
    ("AIR GEO pinout", "wiring",
     "Pinouts for the accumulator: Akku, BMS master, IVT, IMD, connector and SDC splices, with pictures.",
     "Everywhere.", "-"),
    ("com_switch X3", "wiring", "Channel assignment of the Ethernet switches (front / rear).", "Everywhere.", "-"),
    ("lvs_ecu X3", "wiring",
     "LVS ECU channel map: which device is on each channel, on/off per car state (Idle, HV, RTD, AS_ON, AS_RTD) "
     "and current per channel.", "Everywhere.", "-"),
    ("No_DV_lvs_ecu X3", "wiring",
     "Intended as the non-driverless variant of lvs_ecu X3 - currently an exact copy of it.", "Everywhere.", "-"),
    ("JOOOO", "wiring", "Older / alternative copy of Pinouts (differs from Pinouts).", "Everywhere.", "-"),
    ("HV Commisioning", "hv",
     "HV commissioning checklist: category, test, procedure, expected result and PASS / FAIL / UNTESTED.",
     "Columns B-G (P/F has a dropdown).", "Feeds Dashboard"),
    ("Akku Requirements", "hv",
     "LV accumulator sizing: energy budget per event (power x time) and comparison of candidate cells.",
     "Blue input values; totals are calculated.", "-"),
    ("Lizenzen Vector", "admin",
     "Vector CANoe licence keys and who is using them. Contains licence keys - keep sharing restricted.",
     "Columns B and E (user).", "-"),
    ("AlTa HVM BOM", "archive", "Old stock check for hvm accumulator hardware (hidden).", "-", "Reads Platinen Teile"),
    ("AlTa as ebs bom", "archive", "Old stock check for as_ebs (hidden).", "-", "-"),
    ("R", "archive", "Old resistor list per board (hidden, X2 layout).", "-", "Reads Platinen"),
    ("C", "archive", "Old capacitor list per board (hidden, X2 layout).", "-", "Reads Platinen"),
    ("D", "archive", "Old diode list per board (hidden, X2 layout).", "-", "Reads Platinen"),
    ("IC", "archive", "Old IC list per board (hidden, X2 layout).", "-", "Reads Platinen"),
]
NEW_SHEETS = ["Start", "Dashboard", "Review Notes"]

# ---------------------------------------------------------------- formula fixes
# (sheet, cell, find, replace, expected occurrences, why, effect)
FIXES = [
    ("Platinen", "D39", "SUM(D2:D29)", "SUM(D2:D30)", 1,
     "'Eigenentwickelt' total stopped at row 29 and left out hvm_bpr (row 30), which is an own PCB.",
     "Own PCBs in car: 84 -> 85."),
    ("Platinen", "E39", "SUM(E2:E29)", "SUM(E2:E30)", 1,
     "'Eigenentwickelt' order total stopped at row 29 and left out hvm_bpr (row 30).",
     "Own PCBs ordered: 226 -> 229."),
    ("Platinen", "G16", None, "='bestückte Platinen'!P$5", None,
     "Assembled count for lvs_pump was a typed 0 instead of the link every other own PCB uses "
     "('bestückte Platinen' column P is lvs_pump).",
     "Platinen G16: 0 -> 2 (two boards are tested in 'bestückte Platinen')."),
    ("bestückte Platinen", "E4", "E7:E45", "E6:E45", 1,
     "com_cm5 counted from row 7, so board E16-1 (row 6) was ignored. All other columns start at row 6.",
     "No change today; prevents an undercount if E16-1 is set to 'not needed'."),
    ("bestückte Platinen", "E5", "E7:E45", "E6:E45", 7,
     "Same as above: E16-1 of com_cm5 (status 'installed') was not counted as assembled.",
     "com_cm5 assembled: 2 -> 3 (also Platinen G5)."),
    ("bestückte Platinen", "AE4", "AE7:AE45", "AE6:AE45", 1,
     "hvm_bpr counted from row 7, ignoring board E16-1.", "No change today."),
    ("bestückte Platinen", "AE46", "AE7:AE45", "AE6:AE45", 1, "Same off-by-one row for hvm_bpr.", "No change today."),
    ("bestückte Platinen", "AE47", "AE7:AE45", "AE6:AE45", 1, "Same off-by-one row for hvm_bpr.", "No change today."),
    ("bestückte Platinen", "AE48", "AE7:AE45", "AE6:AE45", 1, "Same off-by-one row for hvm_bpr.", "No change today."),
    ("bestückte Platinen", "AE49", "AE7:AE45", "AE6:AE45", 1, "Same off-by-one row for hvm_bpr.", "No change today."),
    ("bestückte Platinen", "AE50", "AE7:AE45", "AE6:AE45", 6, "Same off-by-one row for hvm_bpr.", "No change today."),
    ("Platinen Teile", "I6", "=Platinen!$E9-Platinen!$G9", "=Platinen!$E10-Platinen!$G10", 1,
     "Column I is hmi_trb (I2 = Platinen!C10) but 'boards to build' used row 9 (hmi_tb). Swapped with column K.",
     "I6: -6 -> -4."),
    ("Platinen Teile", "K6", "=Platinen!$E10-Platinen!$G10", "=Platinen!$E9-Platinen!$G9", 1,
     "Column K is hmi_tb (K2 = Platinen!C9) but 'boards to build' used row 10 (hmi_trb). Swapped with column I.",
     "K6: -4 -> -6."),
    ("Platinen Teile", "AE6", "=Platinen!E20-Platinen!G20", "=Platinen!E19-Platinen!G19", 1,
     "Column AE is as_ebs (AE2 = Platinen!C19) but 'boards to build' used row 20 (as_relay).",
     "No change today (both are 0)."),
    ("Platinen Teile", "AH346", "B332", "B346", 2,
     "Stock lookup used the part name from row 332 instead of its own row (copy/paste slip).",
     "Stock of part 615008160421: 0 -> 14."),
    ("Platinen Teile", "AI345", "B331", "B345", 2,
     "Ordered-quantity lookup used row 331's part instead of its own row.", "Row 345 now looks up its own part."),
    ("Platinen Teile", "AI346", "B332", "B346", 2,
     "Ordered-quantity lookup used row 332's part instead of its own row.", "Row 346 now looks up its own part."),
    ("Platinen Teile", "AI347", "B333", "B347", 2,
     "Ordered-quantity lookup used row 333's part instead of its own row.", "Row 347 now looks up its own part."),
    ("Platinen Teile", "AL331", "=A331)*('RT Stock'!D$2:D$5024=#REF!)", "=B331)*('RT Stock'!D$2:D$5024=A331)", 1,
     "Position lookup had a #REF! and compared the part number with the package column. Now same as the rows around it.",
     "Still #N/A today: part 104031-0811 is listed in RT Stock with package 'Molex', not 'THT'."),
    ("Platinen Teile", "AL345", "('RT Stock'!B$2:B$5024=B331)", "('RT Stock'!E$2:E$5024=B345)", 1,
     "Position lookup used row 331's part. Fixed to its own part (SN74HC21), matched on the Part Number column "
     "because B345 is a part number.", "#N/A -> X-C48 (RT position of SN74HC21)."),
    ("Platinen Teile", "AL346", "('RT Stock'!B$2:B$5024=#REF!)*('RT Stock'!D$2:D$5024=#REF!)",
     "('RT Stock'!E$2:E$5024=B346)*('RT Stock'!D$2:D$5024=A346)", 1,
     "Position lookup was #REF!. Rebuilt like its neighbours, matched on Part Number (B346 is a Würth part number).",
     "#N/A -> G-F01 (RT position of 615008160421)."),
    ("Platinen Teile", "AL423", "=B355)", "=B423)", 1,
     "Position lookup used row 355's part instead of its own row.", "Row has no part yet, so no visible change."),
] + [
    ("Platinen Teile", f"AM{r}", "'RT Stock'!$G$2:$G$5024", "'RT Stock'!$I$2:$I$5024", 1,
     "'GT Stock Position' column read the RT position (RT Stock column G) instead of the GT position (column I).",
     "Showed the RT position 'X-B07, Samplebox'; now shows the GT position (none recorded for this part)."
     if r == 287 else "Shows the GT position now (none recorded today, so no visible change).")
    for r in (287, 307, 311, 312, 314, 317, 318, 321, 322)
] + [
    ("Connectors Tracking", f"G{r}", ("SLIDING", r - 2, r + 23), f"=SUMIF($B$4:$B$30,F{r},$C$4:$C$30)", 1,
     "QTY used SUMIF over a window that slid down one row per line (B1:B26, B2:B27, ...), so lower lines "
     "ignored connectors listed near the top and upper lines missed entries added at the bottom. Now always "
     "sums the whole list B4:B30 (the same range the UNIQUE list in column F reads).",
     "Same values today; prevents wrong QTY when the list grows.")
    for r in range(3, 22)
] + [
    ("Connectors Tracking", f"G{r}", ("SLIDING", None, None), f"=SUMIF($B$32:$B$58,F{r},$C$32:$C$58)", 1,
     "Same sliding SUMIF window in the 'Others' block. Now always sums the whole list B32:B58 (the range of "
     "the UNIQUE list in column F).",
     "Same values today; prevents wrong QTY when the list grows.")
    for r in range(33, 52)
] + [
    ("RT Stock", "M363", "=$H361+$J363", "=$H363+$J363", 1,
     "Initial stock added RT stock of row 361 to GT stock of row 363.", "Same value today (both rows have 30)."),
    ("RT Stock", "O363", "=$H361+$J363-$N363", "=$H363+$J363-$N363", 1,
     "Remaining stock used RT stock of row 361.", "Same value today."),
    ("Platinen Tracking", "A50", "=#REF!", "=Platinen!C4", 1,
     "Block title was #REF!. The block below reads 'bestückte Platinen' column D = com_jack (Platinen row 4).",
     "Shows 'com_jack' instead of #REF!."),
    ("Platinen Tracking", "C271", "'bestückte Platinen'!AE8", "'bestückte Platinen'!AE6", 1,
     "hvm_bpr board E16-1 showed the state of E16-3 (row 8).", "Same value today ('coated')."),
    ("Platinen Tracking", "C273", "=#REF!", "='bestückte Platinen'!AE8", 1,
     "hvm_bpr board E16-3 state was #REF!.", "Shows 'coated' instead of #REF!."),
]

# ---------------------------------------------------------------- open issues
ISSUES = [
    ("High", "Platinen Teile", "Row 6 (C6:AF6)",
     "'Boards still to build' goes negative when more boards are assembled than ordered (e.g. hmi_tb, hmi_trb). "
     "Column AG multiplies by this row, so a negative value SUBTRACTS parts needed by other boards and makes "
     "'Total Required' too low.",
     "Wrap each formula in MAX(0, ...), e.g. =MAX(0,Platinen!$E9-Platinen!$G9). Not changed because it alters many stock numbers - please confirm."),
    ("High", "Stecker, Std. Bauteile, R, C, D, IC", "Rows 2-3 and columns AD/AG",
     "Board columns D, J, O, Z, AC point to rows deleted from Platinen (#REF!). Because of that every 'needed' "
     "(AD) and 'to order' (AG) value on these tabs is #REF!.",
     "Re-point the header columns to the right Platinen rows, or archive these tabs if Platinen Teile has replaced them."),
    ("Medium", "bestückte Platinen", "Column AD, AF3, AF4",
     "Column AD belongs to a board that was deleted from Platinen (#REF!). Column AF (hvm_charger) has no row in "
     "Platinen, so its 'In Car' and 'Order' are #REF!.",
     "Delete column AD if the board is gone; add hvm_charger to Platinen and link AF3/AF4."),
    ("Medium", "bestückte Platinen", "Rows 46-50",
     "The summary counts the status 'problem', which is no longer in the dropdown (now 'fixable' / 'unfixable problem'). "
     "'Gesamt' (row 50) only counts unassembled/assembled/ordered/tested/installed/problem and ignores coated, in testing, fixable etc.",
     "Use the Dashboard table 'State of all physical boards', or update the COUNTIFs to the current dropdown values."),
    ("Medium", "Platinen", "K41:M49",
     "Status summary: 'Summe' (K49) leaves out 'not needed' and the statuses concept / schematic done / layout done / "
     "reviewed / Unclear, so the percentages in column M do not add up to 100 %. K41 counts from row 1 instead of row 2.",
     "The Dashboard table 'Design status of board types' has the complete breakdown. Update K41:K49 if you want it here too."),
    ("Medium", "Platinen Teile", "J5, S5, T5",
     "'Can we build' shows #VALUE! because of text in number cells: J95 = '?', S6 = ' ' (space), T351 = ' ' (space).",
     "Clear those cells or put a number in them."),
    ("Low", "bestückte Platinen", "J3, U3, N4, S4",
     "Typed numbers instead of links to Platinen (values match today, but will not update).",
     "Replace with the same formula as the neighbouring cells."),
    ("Low", "Platinen Teile", "Row 5 'Can we build board'",
     "Mix of formulas and typed TRUE values (F5, G5, N5, P5, R5, V5, W5, AB5), so some boards always show TRUE.",
     "Replace with the formula used in C5."),
    ("Low", "RT Stock", "P593",
     "'X3 - Needed' formula differs from all other rows (compares package with Part Number instead of Short Description).",
     "Check whether this is intentional for the RJ45 part."),
    ("Low", "No_DV_lvs_ecu X3", "whole tab", "Exact copy of 'lvs_ecu X3' - the non-driverless changes were never made.",
     "Update it or delete it."),
    ("Low", "JOOOO", "whole tab", "Unclear name; older / alternative copy of Pinouts.",
     "Rename (e.g. 'Pinouts old') or move to the archive. Renaming is safe - no formula refers to it."),
    ("Info", "Platinen Teile", "Columns AL / AM",
     "#N/A means the part was not found in RT Stock with the same package (the Dashboard health checks show how many).",
     "Add the part to RT Stock or fix the package / part number spelling."),
    ("Info", "Connectors Tracking", "F3:F14",
     "Uses UNIQUE(), which works in Google Sheets and Excel 365 but not in older Excel / LibreOffice.", "-"),
    ("Info", "GT Stock", "whole tab",
     "Live IMPORTRANGE from another Google Sheet. Outside Google Sheets you only see the last snapshot.", "-"),
    ("Info", "Lizenzen Vector", "whole tab", "Contains Vector licence keys in plain text.",
     "Keep the file's sharing restricted, or move the keys to a password manager."),
]

FREEZE = {  # sheet: new freeze cell (only where none existed)
    "HV Commisioning": "A3",
    "Connectors Tracking": "A3",
    "Pinouts": "A3",
    "JOOOO": "A3",
    "AIR GEO pinout": "A3",
}
# Google exports these "tables" (banded formatting, no header row) inside the sheet's filter range.
# Excel does not allow a table to overlap the sheet filter and deletes them on open with a
# "repaired" warning. Removing them here gives the same result without the warning; the filters stay.
TABLES_OVERLAPPING_FILTER = {
    ("Platinen Teile", "Table_1"), ("Platinen Teile", "Table_5"), ("Platinen Teile", "Table_6"),
    ("Std. Bauteile", "Table_36"), ("C", "Table_37"), ("D", "Table_38"),
}
HEADER_ROWS = {  # plain header rows that get the workbook's table-header look
    "Platinen Tracking": "B1:I1",
    "GT Stock": "A1:J1",
}


def norm_ws(s):
    return re.sub(r"\s+", " ", s or "").strip()


def apply_fixes(wb, log):
    for sheet, cell, find, repl, count, why, effect in FIXES:
        c = wb[sheet][cell]
        cur = c.value
        is_arr = isinstance(cur, ArrayFormula)
        text = cur.text if is_arr else cur
        if find is None:  # whole-value replacement (a typed constant)
            assert text in (0, 0.0), f"{sheet}!{cell}: expected 0, found {text!r}"
            new_text = repl
        elif isinstance(find, tuple) and find[0] == "SLIDING":
            row = int(re.sub(r"[A-Z]+", "", cell))
            m = re.fullmatch(rf"=SUMIF\(B(\d+):B(\d+),F{row},C(\d+):C(\d+)\)", text or "")
            assert m and m.group(1) == m.group(3) and m.group(2) == m.group(4), f"{sheet}!{cell}: unexpected {text!r}"
            assert int(m.group(1)) <= row <= int(m.group(2)), f"{sheet}!{cell}: window does not contain own row"
            new_text = repl
        else:
            assert isinstance(text, str), f"{sheet}!{cell}: expected a formula, found {text!r}"
            n = text.count(find)
            assert n == count, f"{sheet}!{cell}: expected {count}x {find!r}, found {n}: {text!r}"
            new_text = text.replace(find, repl)
        c.value = ArrayFormula(cur.ref, new_text) if is_arr else new_text
        log["cells"].append({
            "sheet": sheet, "cell": cell, "before": norm_ws(str(text)), "after": norm_ws(new_text),
            "why": why, "effect": effect,
        })


# ---------------------------------------------------------------- helpers for new sheets
def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def title(ws, text, sub=None):
    ws["B1"] = text
    ws["B1"].font = f(16, True, NAVY)
    ws.row_dimensions[1].height = 28
    if sub:
        ws["B2"] = sub
        ws["B2"].font = f(10, False, "595959", italic=True)


def section(ws, row, text, last_col):
    for col in range(2, last_col + 1):
        ws.cell(row, col).fill = fill(NAVY)
    ws.cell(row, 2, text).font = f(11, True, "FFFFFF")
    ws.row_dimensions[row].height = 20


def header(ws, row, labels, start_col=2):
    for i, lab in enumerate(labels):
        c = ws.cell(row, start_col + i, lab)
        c.font = f(10, True)
        c.fill = fill("D9E1F2")
        c.border = BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)


def body(c, wrap=True, bold=False, color="000000", align=None, fmt=None):
    c.font = f(10, bold, color)
    c.border = BOX
    c.alignment = Alignment(vertical="top", wrap_text=wrap, horizontal=align)
    if fmt:
        c.number_format = fmt
    return c


def goto(c, sheet_name, cell="A1"):
    # internal link (no external relationship) - works in Excel, LibreOffice and Google Sheets import
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{sheet_name}'!{cell}", display=str(c.value))


def link(c, sheet_name, text=None):
    c.value = text or sheet_name
    goto(c, sheet_name)
    c.font = f(10, True, "0563C1", underline="single")


def ok_rules(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'],
                                                  fill=fill(GOOD_FILL), font=Font(color=GOOD_FONT, bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"CHECK"'],
                                                  fill=fill(BAD_FILL), font=Font(color=BAD_FONT, bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"INFO"'],
                                                  fill=fill(WARN_FILL), font=Font(color=WARN_FONT, bold=True)))


def dv_list(wb, sheet, contains_cell):
    """Return the dropdown values (list validation) covering contains_cell."""
    ws = wb[sheet]
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and contains_cell in dv.sqref:
            raw = dv.formula1.replace('"&"', "")
            return [v for v in raw.strip('"').split(",") if v]
    raise KeyError(f"no list validation on {sheet}!{contains_cell}")


# ---------------------------------------------------------------- Start
def build_start(wb, ws):
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 2, "B": 18, "C": 24, "D": 62, "E": 44, "F": 40})
    title(ws, "X3 Electric workbook - start here",
          "Click a tab name to jump to it. Tabs are grouped and colour-coded. Nothing on this page needs editing.")
    ws["B3"] = (f"Reorganised {TODAY}. No data was changed; a few broken formulas were repaired - "
                "every change and every open question is listed on the 'Review Notes' tab.")
    ws["B3"].font = f(9, False, "595959")

    r = 5
    section(ws, r, "Tabs", 6)
    r += 1
    header(ws, r, ["Area", "Tab (click to open)", "What it is for", "Where to edit", "Linked with"])
    for name, grp, purpose, edit, linked in TABS:
        r += 1
        label, tab_col, tint = GROUPS[grp]
        a = body(ws.cell(r, 2, label), bold=True, color="FFFFFF")
        a.fill = fill(tab_col)
        c = ws.cell(r, 3)
        body(c)
        if grp == "archive":
            c.value = name
            c.font = f(10, False, "7F7F7F")
        else:
            link(c, name)
        for col, val in ((4, purpose), (5, edit), (6, linked)):
            body(ws.cell(r, col, val))

    r += 2
    section(ws, r, "How the PCB numbers flow", 6)
    steps = [
        ("1", "Platinen", "Enter each board's quantity In Car (D) and Order (E)."),
        ("2", "bestückte Platinen", "Set the state of every physical board (E16-1, E16-2, ...) from the dropdown. "
                                    "Row 5 counts assembled boards; that number appears in Platinen column G."),
        ("3", "Platinen Teile", "Row 6 = boards still to build (Order - Assembled). Total Required (AG) = parts per board x "
                                "boards to build. Stock (AH) and Ordered (AI) come from RT Stock; Surplus (AJ) < 0 means "
                                "parts are missing."),
        ("4", "RT Stock", "Used / Remaining / Needed / Surplus (M-Q) are calculated back from Platinen Teile."),
        ("5", "Dashboard", "Shows the result: which boards are still missing for the car, shortages and health checks."),
    ]
    r += 1
    header(ws, r, ["Step", "Tab", "What happens"])
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    for n, tab, txt in steps:
        r += 1
        body(ws.cell(r, 2, n), align="center")
        link(body(ws.cell(r, 3)), tab)
        body(ws.cell(r, 4, txt))
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 30

    r += 2
    section(ws, r, "Dropdown values (exactly as defined in the sheet)", 6)
    lists = [
        ("Platinen", "K2", "Design status (column K)"),
        ("bestückte Platinen", "B6", "Physical board state (rows 6-45)"),
        ("Platinen Tracking", "F3", "Testing status (column F)"),
        ("HV Commisioning", "F3", "Test result (column F)"),
        ("Connectors Tracking", "I3", "Purchase status (column I)"),
    ]
    r += 1
    header(ws, r, ["Tab", "Column", "Allowed values"])
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    for sheet, cell, what in lists:
        r += 1
        link(body(ws.cell(r, 2)), sheet)
        body(ws.cell(r, 3, what))
        body(ws.cell(r, 4, "  |  ".join(dv_list(wb, sheet, cell))))
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 30

    r += 2
    section(ws, r, "Tips", 6)
    tips = [
        "Cells with a formula are calculated - change the source tab instead (see 'Where to edit').",
        "Hidden tabs (archive) can be shown with right-click on any tab > Unhide / Show.",
        "The Dashboard's 'Health checks' turn red when a formula breaks (e.g. after deleting a row in Platinen).",
        "When you add a new board: add a row in Platinen, a column in 'bestückte Platinen' and in 'Platinen Teile', "
        "then check the Dashboard health checks.",
    ]
    for t in tips:
        r += 1
        ws.cell(r, 2, "•").font = f(10)
        ws.cell(r, 3, t).font = f(10)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- Dashboard
def build_dashboard(wb, ws):
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 2, "B": 30, "C": 22, "D": 13, "E": 22, "F": 11, "G": 11, "H": 12, "I": 14, "J": 12})
    title(ws, "Dashboard", "Everything here is calculated live from the other tabs - nothing to type.")
    link(ws["H1"], "Start", "◀ Back to Start")

    # ---- key numbers
    r = 4
    section(ws, r, "Key numbers", 10)
    kpis = [
        ("Own PCBs needed in car", "=Platinen!D39", "Platinen D39"),
        ("Own PCBs ordered", "=Platinen!E39", "Platinen E39"),
        ("Bought devices in car", "=Platinen!D40", "Platinen D40"),
        ("Board types with design status 'tested'", '=COUNTIF(Platinen!K2:K38,"tested")', "Platinen K"),
        ("Board types in total", "=COUNTA(Platinen!C2:C38)", "Platinen C"),
        ("Own board types still missing boards for the car",
         "=SUMPRODUCT(ISNUMBER(Platinen!G2:G30)*(Platinen!D2:D30>Platinen!G2:G30))", "In car > assembled"),
        ("Parts short for planned builds", "=COUNTIF('Platinen Teile'!AJ8:AJ2097,\"<0\")", "Platinen Teile AJ < 0"),
        ("RT Stock items with negative surplus", "=COUNTIF('RT Stock'!Q2:Q677,\"<0\")", "RT Stock Q < 0"),
        ("HV commissioning tests PASS", '=COUNTIF(\'HV Commisioning\'!F3:F200,"PASS")', "HV Commisioning F"),
        ("HV commissioning tests FAIL", '=COUNTIF(\'HV Commisioning\'!F3:F200,"FAIL")', "HV Commisioning F"),
        ("HV commissioning tests UNTESTED", '=COUNTIF(\'HV Commisioning\'!F3:F200,"UNTESTED")', "HV Commisioning F"),
        ("Connectors arrived / purchased",
         '=COUNTIF(\'Connectors Tracking\'!I3:I63,"arrived")&" / "&COUNTIF(\'Connectors Tracking\'!I3:I63,"purchased")',
         "Connectors Tracking I"),
    ]
    r += 1
    header(ws, r, ["Metric", "Value", "Source"])
    kpi_rows = {}
    for label, formula, src in kpis:
        r += 1
        body(ws.cell(r, 2, label))
        body(ws.cell(r, 3, formula), bold=True, align="right")
        body(ws.cell(r, 4, src), color="7F7F7F")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        kpi_rows[label] = r
    # colour the "problem" KPIs
    for label in ("Own board types still missing boards for the car", "Parts short for planned builds",
                  "RT Stock items with negative surplus", "HV commissioning tests FAIL"):
        ref = f"C{kpi_rows[label]}"
        ws.conditional_formatting.add(ref, CellIsRule(operator="greaterThan", formula=["0"],
                                                      fill=fill(BAD_FILL), font=Font(color=BAD_FONT, bold=True)))
        ws.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=["0"],
                                                      fill=fill(GOOD_FILL), font=Font(color=GOOD_FONT, bold=True)))

    # ---- health checks
    r += 2
    section(ws, r, "Health checks (should all be OK)", 10)
    checks = [
        ("Error cells in Platinen", "=SUMPRODUCT(--ISERROR(Platinen!B2:W49))", 0, "CHECK"),
        ("Error cells in 'bestückte Platinen' rows 2-5", "=SUMPRODUCT(--ISERROR('bestückte Platinen'!B2:AF5))", 0, "CHECK"),
        ("Board columns in 'bestückte Platinen' not found in Platinen",
         "=SUMPRODUCT(--ISERROR(MATCH('bestückte Platinen'!B2:AF2,Platinen!C2:C38,0)))", 0, "CHECK"),
        ("Own PCBs in Platinen without a column in 'bestückte Platinen'",
         "=SUMPRODUCT(--ISNA(MATCH(Platinen!C2:C30,'bestückte Platinen'!B2:AF2,0)))", 0, "CHECK"),
        ("Own PCBs with more assembled than ordered",
         "=SUMPRODUCT(ISNUMBER(Platinen!G2:G30)*(Platinen!G2:G30>Platinen!E2:E30))", 0, "INFO"),
        ("Error cells in 'Platinen Teile' summary rows 3-6", "=SUMPRODUCT(--ISERROR('Platinen Teile'!C3:AF6))", 0, "CHECK"),
        ("Negative 'boards to build' in 'Platinen Teile' row 6", "=COUNTIF('Platinen Teile'!C6:AF6,\"<0\")", 0, "CHECK"),
        ("Error cells in 'Platinen Teile' AG-AK", "=SUMPRODUCT(--ISERROR('Platinen Teile'!AG8:AK2097))", 0, "CHECK"),
        ("Parts in 'Platinen Teile' not found in RT Stock (#N/A in AL)",
         "=SUMPRODUCT(('Platinen Teile'!B8:B2097<>\"\")*ISNA('Platinen Teile'!AL8:AL2097))", 0, "INFO"),
        ("Error cells in RT Stock M-Q", "=SUMPRODUCT(--ISERROR('RT Stock'!M2:Q677))", 0, "CHECK"),
        ("Error cells in Platinen Tracking", "=SUMPRODUCT(--ISERROR('Platinen Tracking'!A1:I303))", 0, "CHECK"),
    ]
    r += 1
    header(ws, r, ["Check", "Result", "Expected", "Status"])
    first = r + 1
    for label, formula, expected, bad_word in checks:
        r += 1
        body(ws.cell(r, 2, label))
        body(ws.cell(r, 3, formula), bold=True, align="right")
        body(ws.cell(r, 4, expected), align="right")
        body(ws.cell(r, 5, f'=IF(C{r}=D{r},"OK","{bad_word}")'), bold=True, align="center")
    ok_rules(ws, f"E{first}:E{r}")
    r += 1
    ws.cell(r, 2, "CHECK = something is broken (see 'Review Notes'); INFO = worth a look, not necessarily wrong.").font = f(9, False, "595959", italic=True)

    # ---- PCB build status
    r += 2
    section(ws, r, "PCB build status (one row per board in Platinen)", 10)
    r += 1
    header(ws, r, ["Board", "Responsible", "In car", "Design status", "Ordered", "Assembled", "Missing for car", "Location"])
    first = r + 1
    for pr in range(2, 39):
        r += 1
        body(ws.cell(r, 2, f'=IF(Platinen!C{pr}="","",Platinen!C{pr})'), bold=True)
        body(ws.cell(r, 3, f'=IF(Platinen!I{pr}="","",Platinen!I{pr})'))
        body(ws.cell(r, 4, f'=IF(Platinen!D{pr}="","",Platinen!D{pr})'), align="right")
        body(ws.cell(r, 5, f'=IF(Platinen!K{pr}="","",Platinen!K{pr})'))
        body(ws.cell(r, 6, f'=IF(Platinen!E{pr}="","",Platinen!E{pr})'), align="right")
        body(ws.cell(r, 7, f'=IF(Platinen!G{pr}="","",Platinen!G{pr})'), align="right")
        body(ws.cell(r, 8, f'=IF(AND(ISNUMBER(Platinen!G{pr}),ISNUMBER(Platinen!D{pr})),MAX(0,Platinen!D{pr}-Platinen!G{pr}),"-")'),
             bold=True, align="right")
        body(ws.cell(r, 9, f'=IF(Platinen!J{pr}="","",Platinen!J{pr})'))
    last = r
    rng = f"H{first}:H{last}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"],
                                                  fill=fill(BAD_FILL), font=Font(color=BAD_FONT, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(H{first}),H{first}=0)'],
                                                   fill=fill(GOOD_FILL), font=Font(color=GOOD_FONT, bold=True)))
    srng = f"E{first}:E{last}"
    ws.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=['"tested"'], fill=fill(GOOD_FILL)))
    ws.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=['"assembled"'], fill=fill("E2EFDA")))
    r += 1
    ws.cell(r, 2, "Missing for car = In car - Assembled (never below 0). '-' = bought device, not tracked in 'bestückte Platinen'.").font = f(9, False, "595959", italic=True)

    # ---- design status breakdown
    r += 2
    section(ws, r, "Design status of board types (Platinen column K)", 10)
    r += 1
    header(ws, r, ["Status", "Board types", "Share"])
    first = r + 1
    for v in dv_list(wb, "Platinen", "K2"):
        r += 1
        body(ws.cell(r, 2, v))
        body(ws.cell(r, 3, f'=COUNTIF(Platinen!$K$2:$K$38,B{r})'), align="right")
        body(ws.cell(r, 4, f'=IF($C${first + 100}=0,0,C{r}/$C${first + 100})'), align="right", fmt="0%")
    last_status = r
    r += 1
    body(ws.cell(r, 2, "(empty or other value)"), color="7F7F7F")
    body(ws.cell(r, 3, f'=C{r + 1}-SUM(C{first}:C{last_status})'), align="right")
    body(ws.cell(r, 4, f'=IF($C${first + 100}=0,0,C{r}/$C${first + 100})'), align="right", fmt="0%")
    r += 1
    total_row = r
    body(ws.cell(r, 2, "Total board types"), bold=True)
    body(ws.cell(r, 3, "=COUNTA(Platinen!C2:C38)"), bold=True, align="right")
    # patch the share formulas to point at the real total row
    for rr in range(first, total_row):
        ws.cell(rr, 4).value = f'=IF($C${total_row}=0,0,C{rr}/$C${total_row})'

    # ---- physical board status
    r += 2
    section(ws, r, "State of all physical boards ('bestückte Platinen' rows 6-45)", 10)
    r += 1
    header(ws, r, ["State", "Boards"])
    first = r + 1
    for v in dv_list(wb, "bestückte Platinen", "B6"):
        r += 1
        body(ws.cell(r, 2, v))
        body(ws.cell(r, 3, f"=COUNTIF('bestückte Platinen'!$B$6:$AF$45,B{r})"), align="right")
    last_state = r
    r += 1
    body(ws.cell(r, 2, "(other value)"), color="7F7F7F")
    body(ws.cell(r, 3, f"=C{r + 1}-SUM(C{first}:C{last_state})"), align="right")
    r += 1
    body(ws.cell(r, 2, "Total boards entered"), bold=True)
    body(ws.cell(r, 3, "=COUNTA('bestückte Platinen'!B6:AF45)"), bold=True, align="right")

    ws.freeze_panes = "A4"


# ---------------------------------------------------------------- Review Notes
def build_notes(wb, ws, log, original_order):
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 2, "B": 6, "C": 20, "D": 16, "E": 48, "F": 48, "G": 46})
    title(ws, "Review notes", f"What was changed on {TODAY}, and what still needs a decision.")
    link(ws["G1"], "Start", "◀ Back to Start")

    r = 4
    section(ws, r, "1. Structure & formatting changes (no data touched)", 7)
    items = [
        "Added tabs: Start (guide with links), Dashboard (live summary + health checks), Review Notes (this tab).",
        "Tabs reordered into groups and colour-coded: Overview, PCBs, Stock & parts, Wiring & pinouts, HV & accumulator, "
        "Admin, Archive. Hidden tabs stay hidden. Formulas refer to tabs by name, so the order change does not affect them.",
        "Frozen header rows added on: " + ", ".join(FREEZE) + ".",
        "Header rows styled like the workbook's existing tables on: " + ", ".join(f"{k} ({v})" for k, v in HEADER_ROWS.items()) + ".",
        "Removed 6 header-less 'tables' from the Google export (Platinen Teile: Table_1, Table_5, Table_6; "
        "Std. Bauteile: Table_36; C: Table_37; D: Table_38). They overlapped the sheet filter, which Excel does not "
        "allow, so Excel deleted them with a 'repaired' warning anyway. Data, colours and filters are unchanged.",
        "No tab was renamed or deleted, no row or column was inserted or removed, and no typed value was changed.",
        "Dropdowns, colour rules, tables, pictures and the saved Google filter views (PCB Filter Personal, Filter 1, "
        "Filter 2, FeKru, FrGa) are all kept.",
    ]
    for t in items:
        r += 1
        ws.cell(r, 2, "•").font = f(10)
        ws.cell(r, 3, t).font = f(10)
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 28 if len(t) > 110 else 15

    r += 2
    groups = []  # consecutive fixes with the same reason are shown as one line
    for ch in log["cells"]:
        if groups and groups[-1][0]["sheet"] == ch["sheet"] and groups[-1][0]["why"] == ch["why"]:
            groups[-1].append(ch)
        else:
            groups.append([ch])
    section(ws, r, f"2. Formula fixes ({len(log['cells'])} cells) - each was a clear slip; the old formula is kept here", 7)
    r += 1
    header(ws, r, ["#", "Tab", "Cell(s)", "Why", "Before  ->  After", "Effect"])
    for i, grp in enumerate(groups, 1):
        ch = grp[0]
        r += 1
        body(ws.cell(r, 2, i), align="center")
        link(body(ws.cell(r, 3)), ch["sheet"])
        cells = ch["cell"] if len(grp) == 1 else (
            ", ".join(g["cell"] for g in grp) if len(grp) <= 6 else f"{grp[0]['cell']} ... {grp[-1]['cell']} ({len(grp)} cells)")
        c = body(ws.cell(r, 4, cells), bold=True)
        goto(c, ch["sheet"], ch["cell"])
        c.font = f(10, True, "0563C1", underline="single")
        body(ws.cell(r, 5, ch["why"]))
        example = "" if len(grp) == 1 else f"(first cell shown, the others follow the same pattern)\n"
        body(ws.cell(r, 6, f"{example}BEFORE: {ch['before']}\nAFTER:  {ch['after']}"))
        body(ws.cell(r, 7, ch["effect"]))
        longest = max(len(ch["before"]), len(ch["after"]), len(ch["why"]) * 45 // 48)
        ws.row_dimensions[r].height = max(45, 15 * (2 + longest // 45))

    r += 2
    section(ws, r, f"3. Needs a decision ({len(ISSUES)} items, not changed because the right fix is not certain)", 7)
    r += 1
    header(ws, r, ["#", "Priority", "Tab", "Where", "Issue", "Suggested fix"])
    first = r + 1
    for i, (prio, tab, where, issue, fix) in enumerate(ISSUES, 1):
        r += 1
        body(ws.cell(r, 2, i), align="center")
        body(ws.cell(r, 3, prio), bold=True)
        if tab in wb.sheetnames:
            link(body(ws.cell(r, 4)), tab)
        else:
            body(ws.cell(r, 4, tab))
        body(ws.cell(r, 5, f"{where}: {issue}"))
        body(ws.cell(r, 6, fix))
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        ws.row_dimensions[r].height = max(30, 15 * (1 + len(issue) // 50))
    prng = f"C{first}:C{r}"
    ws.conditional_formatting.add(prng, CellIsRule(operator="equal", formula=['"High"'], fill=fill(BAD_FILL), font=Font(color=BAD_FONT, bold=True)))
    ws.conditional_formatting.add(prng, CellIsRule(operator="equal", formula=['"Medium"'], fill=fill(WARN_FILL), font=Font(color=WARN_FONT, bold=True)))
    ws.conditional_formatting.add(prng, CellIsRule(operator="equal", formula=['"Low"'], fill=fill("DDEBF7")))

    r += 2
    section(ws, r, "4. Original tab order (for reference)", 7)
    for i, name in enumerate(original_order, 1):
        r += 1
        ws.cell(r, 2, i).font = f(10)
        ws.cell(r, 3, name).font = f(10)
    ws.freeze_panes = "A4"


# ---------------------------------------------------------------- main
def main(src, dst, changes_path):
    wb = openpyxl.load_workbook(src, rich_text=True)
    original_order = list(wb.sheetnames)
    plan_names = [t[0] for t in TABS]
    assert sorted(set(plan_names) - set(NEW_SHEETS)) == sorted(original_order), "tab plan does not cover every sheet"
    for n in NEW_SHEETS:
        assert n not in wb.sheetnames

    log = {"cells": [], "styled_cells": [], "new_sheets": NEW_SHEETS, "features": []}
    apply_fixes(wb, log)

    # tables Excel would delete anyway (overlap with the sheet filter)
    found = set()
    for ws in wb.worksheets:
        af = ws.auto_filter.ref
        for t in list(ws.tables.values()):
            if af and not CellRange(af.replace("$", "")).isdisjoint(CellRange(t.ref)):
                found.add((ws.title, t.name))
                del ws.tables[t.name]
                log["features"].append([ws.title, "tables"])
    assert found == TABLES_OVERLAPPING_FILTER, f"unexpected overlapping tables: {found}"

    # usability tweaks on existing tabs
    for sheet, cell in FREEZE.items():
        ws = wb[sheet]
        assert ws.freeze_panes is None, f"{sheet} already frozen"
        ws.freeze_panes = cell
        log["features"].append([sheet, "freeze"])
    for sheet, rng in HEADER_ROWS.items():
        ws = wb[sheet]
        for row in ws[rng]:
            for c in row:
                c.font = Font(name=c.font.name or FONT, size=c.font.size or 10, bold=True, color="FFFFFF")
                c.fill = fill("356854")  # same colour as the existing table headers (RT Stock, HV Commisioning)
                c.alignment = Alignment(vertical="center", wrap_text=True)
                log["styled_cells"].append([sheet, c.coordinate])

    # new tabs
    start = wb.create_sheet("Start")
    dash = wb.create_sheet("Dashboard")
    notes = wb.create_sheet("Review Notes")
    build_start(wb, start)
    build_dashboard(wb, dash)
    build_notes(wb, notes, log, original_order)

    # order + colours (visibility untouched)
    states_before = {ws.title: ws.sheet_state for ws in wb.worksheets}
    wb._sheets = [wb[n] for n in plan_names]
    for name, grp, *_ in TABS:
        wb[name].sheet_properties.tabColor = GROUPS[grp][1]
    for ws in wb.worksheets:
        assert ws.sheet_state == states_before[ws.title]
        ws.sheet_view.tabSelected = ws.title == "Start"
    wb.active = 0
    for v in wb.views:
        v.firstSheet = 0

    wb.save(dst)
    with open(changes_path, "w", encoding="utf8") as fh:
        json.dump(log, fh, ensure_ascii=False, indent=1)
    print(f"Saved {dst}: {len(log['cells'])} formula fixes, {len(log['styled_cells'])} restyled header cells, "
          f"{len(FREEZE)} freeze panes, new tabs {NEW_SHEETS}.")


if __name__ == "__main__":
    main(*sys.argv[1:])
