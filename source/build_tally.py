#!/usr/bin/env python3
r"""
build_tally.py - Generate ../Centerbeam_Tally_Recommender.xlsm

Hybrid build:
  1. openpyxl writes ALL structure / formatting / data-validation / freeze-panes
     to a temporary .xlsx (fast, headless, no Excel required).
  2. A short Excel COM pass (pywin32) injects the VBA module + the sheet sort
     event + the two macro buttons, seeds an example run, and saves as .xlsm.

The VBA lives in source/tally_recommender.bas and is loaded via AddFromString,
so the macro logic stays readable and version-controlled.

Requirements:
  - openpyxl, pywin32  (pip install openpyxl pywin32)
  - Microsoft Excel installed, with Trust Center > "Trust access to the VBA
    project object model" enabled (one-time; already on for this project).

Run (use the REAL interpreter -- on Windows a bare `python` may resolve to the
Microsoft Store stub; use the full path to the python.org install):
  "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" source\build_tally.py

Note: make sure no headless EXCEL.EXE is already running before building -- a
stale COM instance can make the pywin32 pass hang.
"""

import os
import sys
import tempfile

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------------------
# Paths
# ----------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
OUT = os.path.join(REPO, "Centerbeam_Tally_Recommender.xlsm")
BAS = os.path.join(HERE, "tally_recommender.bas")

# ----------------------------------------------------------------------------
# Palette (RGB hex - openpyxl uses normal RRGGBB, unlike COM's BGR integers)
# ----------------------------------------------------------------------------
NAVY, INK, WHITE = "1F4E79", "33414F", "FFFFFF"
HDRLT, STATUSF, NOTEF = "DDE6ED", "EAF0E6", "F0F4F8"
LC = ["BBD4EA", "C2E5C9", "FBE7B2", "E5CDEE", "FAC4BC", "BFEAE0", "DCE7AE"]  # 8..20 fills
LEN_NUMS = [8, 10, 12, 14, 16, 18, 20]
COLS = list("ABCDEFGHIJKL")
LEN_COLS = list("CDEFGHI")  # the 8'..20' columns

LEFT, CENTER, RIGHT = "left", "center", "right"


def font(bold=False, italic=False, size=11, color=INK):
    return Font(name="Calibri", size=size, bold=bold, italic=italic, color=color)


def solid(hex_):
    return PatternFill(fill_type="solid", fgColor=hex_)


def put(ws, addr, value=None, *, bold=False, italic=False, size=11, color=INK,
        fill=None, align=None, valign=None, wrap=False):
    c = ws[addr]
    if value is not None:
        c.value = value
    c.font = font(bold, italic, size, color)
    if fill:
        c.fill = solid(fill)
    if align or valign or wrap:
        c.alignment = Alignment(horizontal=align, vertical=valign, wrap_text=wrap)
    return c


def banner(ws, rng, value, *, fill=None, color=INK, bold=True, size=11,
           italic=False, align=LEFT):
    """Style every cell of a range, merge it, and set the top-left value."""
    f = font(bold, italic, size, color)
    pf = solid(fill) if fill else None
    for row in ws[rng]:
        for cc in row:
            cc.font = f
            if pf:
                cc.fill = pf
    ws.merge_cells(rng)
    top = rng.split(":")[0]
    ws[top] = value
    ws[top].alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)


def list_validation(ws, addr, items):
    dv = DataValidation(type="list", formula1='"%s"' % items, allow_blank=True,
                        showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(addr)


def setup_grid(ws):
    widths = {"A": 8, "B": 42, "C": 6, "D": 6, "E": 6, "F": 6, "G": 6, "H": 6,
              "I": 6, "J": 10, "K": 9, "L": 13}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def color_length_headers(ws, row):
    for i, col in enumerate(LEN_COLS):
        ws["%s%d" % (col, row)].fill = solid(LC[i])


def fmt_block(ws, rng, numfmt=None):
    """Center a block (and optionally apply an integer number format) so the
    count/total columns read consistently across the sheet."""
    for row in ws[rng]:
        for c in row:
            c.alignment = Alignment(horizontal=CENTER, vertical="center")
            if numfmt:
                c.number_format = numfmt


# ----------------------------------------------------------------------------
# Stage 1: openpyxl builds the structure
# ----------------------------------------------------------------------------
def build_layout(path):
    wb = Workbook()

    # ---- Recommender ----
    rec = wb.active
    rec.title = "Recommender"
    setup_grid(rec)

    banner(rec, "A1:L1", "Centerbeam 72-ft Car  -  Lumber Tally Recommender",
           fill=NAVY, color=WHITE, bold=True, size=16, align=CENTER)
    rec.row_dimensions[1].height = 30
    banner(rec, "A2:L2",
           "Pick the lengths you want, choose a match mode and car size, then click Recommend Tallies.",
           italic=True, size=10, align=LEFT)

    banner(rec, "A4:B4", "SETTINGS", fill=HDRLT, bold=True)
    put(rec, "A5", "Car size:", bold=True, align=RIGHT)
    put(rec, "B5", "720 ft (10 rows)")
    put(rec, "A6", "Match mode:", bold=True, align=RIGHT)
    put(rec, "B6", "Palette - use only selected")
    banner(rec, "A7:L7",
           "Palette = use only these lengths   /   Each must appear = every picked length shows up in the car   /   + fillers = may add other lengths to finish a 72-ft row",
           italic=True, size=9)

    list_validation(rec, "B5", "720 ft (10 rows),1008 ft (14 rows)")
    list_validation(rec, "B6", "Palette - use only selected,Each selected must appear,Each must appear + fillers")

    banner(rec, "A9:B9", "LENGTH PALETTE", fill=HDRLT, bold=True)
    put(rec, "A10", "Length", bold=True, align=CENTER)
    put(rec, "B10", "Include?  (check the box)", bold=True)
    for i, n in enumerate(LEN_NUMS):
        row = 11 + i
        put(rec, "A%d" % row, n, bold=True, fill=LC[i], align=CENTER)
        # A Form Control checkbox (added in the COM stage) links to this cell as
        # TRUE/FALSE; the ';;;' number format hides that text so only the box shows.
        cb = put(rec, "B%d" % row, None, align=LEFT)
        cb.number_format = ";;;"

    put(rec, "A19", "Status:", bold=True)
    banner(rec, "B19:L19", "", fill=STATUSF)
    rec["B19"].alignment = Alignment(horizontal=LEFT, vertical="center", wrap_text=True)
    rec.row_dimensions[19].height = 28

    banner(rec, "A21:L21", "RECOMMENDED FULL-CAR TALLIES", fill=NAVY, color=WHITE, bold=True)
    rec_hdr = ["#", "Tally (how to load the car)", "8'", "10'", "12'", "14'", "16'",
               "18'", "20'", "Total pcs", "Total ft", "OK?"]
    for c, txt in zip(COLS, rec_hdr):
        put(rec, "%s22" % c, txt, bold=True, fill=HDRLT)
    color_length_headers(rec, 22)

    banner(rec, "A39:L39",
           "Need the building blocks or to hand-build a custom blend?   ->   see the 'Row Patterns' tab.",
           fill=NOTEF, italic=True)

    # ---- Row Patterns ----
    pat = wb.create_sheet("Row Patterns")
    setup_grid(pat)
    banner(pat, "A1:J1",
           "VALID 72-FT ROW PATTERNS  -  building blocks  (set 'Rows to use' to hand-build; click a length header 8'-20' to sort)",
           fill=NAVY, color=WHITE, bold=True)
    pat.row_dimensions[1].height = 22
    pat_hdr = ["#", "Row = 72 ft", "8'", "10'", "12'", "14'", "16'", "18'", "20'",
               "Rows to use"]
    for c, txt in zip(COLS, pat_hdr):
        put(pat, "%s2" % c, txt, bold=True, fill=HDRLT)
    color_length_headers(pat, 2)
    pat.column_dimensions["J"].width = 12   # 'Rows to use' moved here (Pcs/Ft removed)
    pat.freeze_panes = "A3"   # <-- freeze the two header rows (one clean line)

    # ---- How to Use ----
    how = wb.create_sheet("How to Use")
    how.column_dimensions["A"].width = 104
    lines = [
        "CENTERBEAM 72-FT CAR  -  LUMBER TALLY RECOMMENDER",
        "",
        "WHAT IT DOES",
        "A 72-ft centerbeam car is loaded in rows; every row must be filled end-to-end to exactly 72 ft.",
        "This tool finds the length combinations that fill a 72-ft row, then recommends complete car tallies",
        "(720 ft = 10 rows, or 1008 ft = 14 rows) using the lengths you choose.",
        "",
        "THE TABS",
        "Recommender  - set your options here and read the recommended full-car tallies.",
        "Row Patterns - every way to fill a single 72-ft row (building blocks); hand-build a custom car here.",
        "How to Use   - this sheet.",
        "",
        "HOW TO USE (Recommender tab)",
        "1.  Car size: pick 720 ft (10 rows) or 1008 ft (14 rows).",
        "2.  Match mode:",
        "       - Palette - use only selected: tallies are built only from the lengths you turn on.",
        "       - Each selected must appear: every length you turn on must show up somewhere in the car.",
        "       - Each must appear + fillers: every selected length appears, other lengths may finish a row.",
        "3.  Length palette: check the box for each length you want; clear it to leave that length out.",
        "4.  Click Recommend Tallies.",
        "",
        "RECOMMENDED FULL-CAR TALLIES (Recommender tab)",
        "Ready-to-load cars. Each row shows the piece count for every length, the total pieces, the total feet",
        '(always rows x 72), and an OK check. "All 10 rows: 2x16, 2x20" means load every row the same way:',
        "two 16s + two 20s = 20 pcs 16 ft + 20 pcs 20 ft = 720 ft.",
        "",
        "ROW PATTERNS tab",
        "Every way to fill a single 72-ft row from your lengths. The two header rows stay frozen at the top as",
        "you scroll. Click a length column header (8 ft ... 20 ft) to sort the patterns by that length, most",
        "first. To hand-build a mixed car, type how many rows of each pattern in the Rows to use column; the",
        "YOUR HAND-BUILT TALLY line (directly below the last pattern) totals the pieces per length and the",
        "row count. Make the row count match your car size (10 or 14).",
        "",
        "NOTES",
        "-  Only length controls the 72-ft fit; product and grade are not part of this tool.",
        "-  Click Enable Content when you open the file so the buttons, checkboxes and sorting work.",
        "-  Up to 15 recommendations and up to 101 row patterns are listed.",
    ]
    for i, text in enumerate(lines):
        if text:
            put(how, "A%d" % (i + 1), text)
    put(how, "A1", lines[0], bold=True, size=14, color=NAVY)
    for r in (3, 8, 13, 22, 27, 34):
        put(how, "A%d" % r, lines[r - 1], bold=True, color=NAVY)

    # ---- consolidate numeric formatting across the data blocks ----
    # (count + total columns centered as integers; runtime VBA writes preserve it)
    fmt_block(rec, "A23:A37", "0")
    fmt_block(rec, "C23:K37", "0")
    fmt_block(rec, "L23:L37")            # OK? is text, just center it
    fmt_block(pat, "A3:A103", "0")
    fmt_block(pat, "C3:J103", "0")       # lengths (C-I) + Rows to use (J)

    wb.save(path)


# ----------------------------------------------------------------------------
# Stage 2: Excel COM injects VBA + buttons, seeds, saves as .xlsm
# ----------------------------------------------------------------------------
SHEET_EVENT = "\r\n".join([
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)",
    "    If Target.Cells.Count <> 1 Then Exit Sub",
    "    If Target.Row = 2 And Target.Column >= 3 And Target.Column <= 9 Then",
    "        SortPatternsBy Target.Column",
    "    End If",
    "End Sub",
])

XL_MACRO_ENABLED = 52  # xlOpenXMLWorkbookMacroEnabled


def inject_macros(src_xlsx, out_xlsm):
    import win32com.client as win32

    with open(BAS, "r", encoding="utf-8") as fh:
        module_code = "\r\n".join(fh.read().splitlines())

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow (allow macro Run)
    except Exception:
        pass

    wb = None
    try:
        print("      .. workbooks open before:", excel.Workbooks.Count, flush=True)
        print("      .. opening temp xlsx", flush=True)
        wb = excel.Workbooks.Open(src_xlsx)
        print("      .. opened; adding standard module", flush=True)

        # standard module
        comp = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = "TallyRecommender"
        comp.CodeModule.AddFromString(module_code)
        print("      .. module added; adding sheet event", flush=True)

        # sheet click-to-sort event on the Row Patterns sheet
        ws2 = wb.Worksheets("Row Patterns")
        wb.VBProject.VBComponents(ws2.CodeName).CodeModule.AddFromString(SHEET_EVENT)
        print("      .. event added; adding buttons", flush=True)

        # macro buttons on the Recommender sheet
        rec = wb.Worksheets("Recommender")
        b1 = rec.Buttons().Add(rec.Range("D4").Left, rec.Range("D4").Top, 160, 30)
        b1.Caption = "Recommend Tallies"
        b1.OnAction = "RecommendTallies"
        b1.Font.Size = 11
        b1.Font.Bold = True
        b2 = rec.Buttons().Add(rec.Range("D6").Left, rec.Range("D6").Top, 90, 24)
        b2.Caption = "Clear"
        b2.OnAction = "ClearOutputsButton"
        print("      .. buttons added; adding palette checkboxes", flush=True)

        # length-palette checkboxes (replace the old Yes/No dropdowns). Each links
        # to its B cell as TRUE/FALSE; the ';;;' format from the layout stage hides
        # the text so only the box shows next to the colored length chip in col A.
        xl_on, xl_off = 1, -4146
        default_checked = {16, 20}
        for i, n in enumerate(LEN_NUMS):
            row = 11 + i
            cell = rec.Range("B%d" % row)
            cb = rec.CheckBoxes().Add(cell.Left + 2, cell.Top + 1, 16, 13)
            cb.Caption = ""
            cb.LinkedCell = "$B$%d" % row
            cb.Value = xl_on if n in default_checked else xl_off
            cell.Value = (n in default_checked)
        print("      .. checkboxes added; seeding (Run)", flush=True)

        # seed an example run so the file is populated on open
        rec.Activate()
        excel.Run("RecommendTallies")
        print("      .. seeded; saving .xlsm", flush=True)

        if os.path.exists(out_xlsm):
            os.remove(out_xlsm)
        wb.SaveAs(out_xlsm, XL_MACRO_ENABLED)
        has_vba = bool(wb.HasVBProject)
        print("      .. saved; HasVBProject =", has_vba, flush=True)
        wb.Close(False)
        return has_vba
    finally:
        try:
            if wb is not None and excel.Workbooks.Count:
                pass
        except Exception:
            pass
        excel.Quit()


def main():
    tmp = os.path.join(tempfile.gettempdir(), "tally_build_tmp.xlsx")
    print("[1/2] openpyxl: building layout ->", os.path.basename(tmp))
    build_layout(tmp)

    print("[2/2] Excel COM: injecting VBA + buttons, seeding, saving .xlsm")
    has_vba = inject_macros(tmp, OUT)

    try:
        os.remove(tmp)
    except OSError:
        pass

    size_kb = round(os.path.getsize(OUT) / 1024, 1)
    print("DONE ->", OUT)
    print("      HasVBProject:", has_vba, "| size:", size_kb, "KB")
    if not has_vba:
        print("WARNING: macros not present after save!", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
