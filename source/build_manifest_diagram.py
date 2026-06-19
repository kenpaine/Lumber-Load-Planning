#!/usr/bin/env python3
r"""
build_manifest_diagram.py — add the "to scale, per side" car diagram to the
Manifest tab of Centerbeam_Lumber_Layout_Planner.xlsm.

The Layout Planner workbook is NOT cleanly regenerable here (its build_v3*.py
scripts use /home/claude paths and don't add the working macros), so this script
*modifies the existing .xlsm in place*, preserving its CenterbeamSolver macros:

  1. Re-titles the Manifest "Car Layout" header and CLEARS the old fixed-cell
     grid (rows 14-29) + its conditional formatting.
  2. Injects the ManifestDiagram module (source/manifest_diagram.bas).
  3. Injects a Worksheet_Activate event on the Manifest sheet so the diagram
     redraws whenever you open the tab (after a Solve), plus a Redraw button.
  4. Seeds the first draw and saves.

Idempotent — safe to re-run. Requires pywin32 + Excel; the .xlsm must be CLOSED.
Run:  "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" source\build_manifest_diagram.py
"""
import os
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
XLSM = os.path.join(REPO, "Centerbeam_Lumber_Layout_Planner.xlsm")
BAS = os.path.join(HERE, "manifest_diagram.bas")

MSO_ROUNDED_RECT = 5
XL_NONE = -4142

ACTIVATE_EVENT = "\r\n".join([
    "Private Sub Worksheet_Activate()",
    "    On Error Resume Next",
    "    DrawManifestDiagram",
    "End Sub",
    "",
    "Private Sub Worksheet_Change(ByVal Target As Range)",
    "    If Intersect(Target, Me.Range(\"N2\")) Is Nothing Then Exit Sub",
    "    Application.EnableEvents = False",
    "    On Error Resume Next",
    "    DrawManifestDiagram",
    "    On Error GoTo 0",
    "    Application.EnableEvents = True",
    "End Sub",
])


def main():
    with open(BAS, "r", encoding="utf-8") as fh:
        code = "\r\n".join(fh.read().splitlines())

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    wb = None
    try:
        print("workbooks open before:", excel.Workbooks.Count)
        wb = excel.Workbooks.Open(XLSM)
        mf = wb.Worksheets("Manifest")
        print("opened; Manifest codename =", mf.CodeName)

        # 1. retitle header (row 13, merged A13:L13) + clear old grid rows 14-29
        mf.Range("A13").Value = ("CAR LAYOUT — drawn to scale, split per side   "
                                 "(updates from the solved load; click Redraw if needed)")
        area = mf.Range("A14:L29")
        area.Clear()
        try:
            area.FormatConditions.Delete()
        except Exception as e:
            print("  (no CF to delete:", e, ")")
        for r in range(14, 30):
            mf.Rows(r).RowHeight = 19
        print("  cleared old fixed-cell grid; reserved rows 14-29 for the diagram")

        # 2. inject the ManifestDiagram standard module (idempotent)
        vbp = wb.VBProject
        for c in list(vbp.VBComponents):
            if c.Name == "ManifestDiagram":
                vbp.VBComponents.Remove(c)
        comp = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = "ManifestDiagram"
        comp.CodeModule.AddFromString(code)
        print("  injected ManifestDiagram module")

        # 3a. Worksheet_Activate on the Manifest sheet module (replace any existing)
        mcm = vbp.VBComponents(mf.CodeName).CodeModule
        if mcm.CountOfLines > 0:
            mcm.DeleteLines(1, mcm.CountOfLines)
        mcm.AddFromString(ACTIVATE_EVENT)
        print("  injected Worksheet_Activate ->", mf.CodeName)

        # 3b. Redraw button (idempotent), parked just right of the L-column print edge
        #     so the full-width diagram has the page to itself and the button never prints.
        for i in range(mf.Shapes.Count, 0, -1):
            if mf.Shapes(i).Name == "btn_redraw":
                mf.Shapes(i).Delete()
        left = mf.Range("M14").Left + 6      # column M starts past the A:L print area
        top = mf.Range("A13").Top
        btn = mf.Shapes.AddShape(MSO_ROUNDED_RECT, left, top, 86, 22)
        btn.Name = "btn_redraw"
        btn.Fill.Solid()
        btn.Fill.ForeColor.RGB = 0x4F8A3D  # green (BGR-ish; cosmetic)
        btn.Line.Visible = False
        tr = btn.TextFrame.Characters()
        tr.Text = "Redraw"
        tr.Font.Bold = True
        tr.Font.Size = 9
        tr.Font.Color = 0xFFFFFF
        btn.TextFrame.HorizontalAlignment = -4108  # center
        btn.TextFrame.VerticalAlignment = -4108
        btn.OnAction = "RedrawManifestButton"
        print("  added Redraw button")

        # 3c. Palette selector (off-print, columns M/N) — the 12 length-colour
        #     schemes shared with the browser apps. Changing it fires
        #     Worksheet_Change, which redraws the diagram + pick list.
        SCHEME_NAMES = ("Color (pastel),Vivid,Material,Tableau,Earth / lumberyard,"
                        "Jewel tones,Rainbow (warm to cool),Viridis (colour-safe),"
                        "Sunset (warm),Neon,High contrast,B & W (print)")
        lbl = mf.Range("M2")
        lbl.Value = "Palette:"
        lbl.Font.Bold = True
        lbl.HorizontalAlignment = -4152          # xlRight
        pal_cell = mf.Range("N2")
        try:
            pal_cell.Validation.Delete()
        except Exception:
            pass
        # 3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween
        pal_cell.Validation.Add(3, 1, 1, SCHEME_NAMES)
        pal_cell.Validation.IgnoreBlank = True
        pal_cell.Validation.InCellDropdown = True
        # migrate the old 3-item default ("Color") to the new name
        if str(pal_cell.Value or "").strip() in ("", "Color"):
            pal_cell.Value = "Color (pastel)"
        pal_cell.Interior.Color = 0xE8F0DF       # light tint so it reads as an input
        pal_cell.Font.Bold = True
        mf.Columns("M").ColumnWidth = 9
        mf.Columns("N").ColumnWidth = 20
        print("  added palette selector (N2):", pal_cell.Value)

        # 4. The Pick List is now drawn dynamically as shapes by VBA, directly under
        #    the diagram, wrapping into columns so it stays short. Remove the old
        #    cell-based Pick List (the shapes replace it) and reserve a blank band
        #    below the diagram so the shapes never overlap live cells. The Placed
        #    Packs / Summary tables stay below as data and are left off the print.
        RESERVE_ROWS = 16
        band_last = 29 + RESERVE_ROWS    # rows 30..45 reserved for the pick-list shapes

        def find_banner(substr):
            for r in range(1, 200):
                if substr in str(mf.Cells(r, 1).Value or "").upper():
                    return r
            return None

        # 4a. delete any leftover cell-based PICK LIST block (banner .. TOTAL row)
        pick_row = find_banner("PICK LIST")
        if pick_row:
            total_row = None
            for r in range(pick_row + 1, pick_row + 60):
                if str(mf.Cells(r, 1).Value or "").strip().upper() == "TOTAL":
                    total_row = r
                    break
            if total_row is None:
                total_row = pick_row + 22
            mf.Rows(f"{pick_row}:{total_row}").Delete()
            print(f"  removed old cell Pick List (rows {pick_row}-{total_row})")

        # 4b. reserve the blank band: push the first data table (Placed Packs) to row 46
        pp_row = find_banner("PLACED PACKS")
        if pp_row and pp_row < 30 + RESERVE_ROWS:
            cnt = 30 + RESERVE_ROWS - pp_row
            mf.Rows(f"{pp_row}:{pp_row + cnt - 1}").Insert(Shift=-4121)  # xlShiftDown
            print(f"  reserved rows 30-{band_last}; data tables start at row {30 + RESERVE_ROWS}")
        else:
            print("  Pick List band already reserved")

        # 4c. clean the reserve band so only the shapes show there (no stray fills)
        mf.Rows(f"30:{band_last}").ClearFormats()
        mf.Rows(f"30:{band_last}").RowHeight = 14.4

        # 5. seed the draw — DrawManifestDiagram now also draws the dynamic pick list
        #    and sets the fit-to-one-page landscape print area to the pick list's bottom.
        mf.Activate()
        excel.Run("DrawManifestDiagram")
        ndiag = sum(1 for i in range(1, mf.Shapes.Count + 1)
                    if str(mf.Shapes(i).Name).startswith("diag_"))
        npick = sum(1 for i in range(1, mf.Shapes.Count + 1)
                    if str(mf.Shapes(i).Name).startswith("pick_"))
        print(f"  seeded; diagram shapes: {ndiag}, pick-list shapes: {npick}")
        print("  print area:", mf.PageSetup.PrintArea)

        wb.Save()
        has_vba = bool(wb.HasVBProject)
        print("saved; HasVBProject =", has_vba)
        wb.Close(False)
        print("DONE")
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
