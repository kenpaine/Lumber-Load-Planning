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

        # 3b. Redraw button (idempotent), placed to the right of the diagram
        for i in range(mf.Shapes.Count, 0, -1):
            if mf.Shapes(i).Name == "btn_redraw":
                mf.Shapes(i).Delete
        left = mf.Range("A14").Left + 556
        top = mf.Range("A14").Top + 2
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

        # 4. seed the first draw
        mf.Activate()
        excel.Run("DrawManifestDiagram")
        ndiag = sum(1 for i in range(1, mf.Shapes.Count + 1)
                    if str(mf.Shapes(i).Name).startswith("diag_"))
        print("  seeded; diagram shapes drawn:", ndiag)

        # 5. move the PICK LIST directly under the car layout diagram (rows 14-29).
        #    The Placed Packs and Summary blocks reflow below it and are left off the
        #    print area (Placed Packs doesn't need to print). Idempotent: skips if the
        #    Pick List is already directly under the diagram.
        INSERT_AT = 30           # first row after the diagram band (rows 14-29)

        def find_banner(substr):
            for r in range(1, 200):
                if substr in str(mf.Cells(r, 1).Value or "").upper():
                    return r
            return None

        pick_row = find_banner("PICK LIST")
        if pick_row is None:
            print("  WARNING: could not find PICK LIST section")
        else:
            # bottom of the block = its TOTAL row (handles any inventory size)
            pick_end = None
            for r in range(pick_row + 1, pick_row + 60):
                if str(mf.Cells(r, 1).Value or "").strip().upper() == "TOTAL":
                    pick_end = r
                    break
            if pick_end is None:
                pick_end = pick_row + 22   # fallback: banner+header+20 items+total
            n_pick = pick_end - pick_row + 1

            if pick_row != INSERT_AT:
                mf.Rows(f"{pick_row}:{pick_end}").Cut()
                mf.Rows(INSERT_AT).Insert(Shift=-4121)  # xlShiftDown (insert cut cells)
                excel.CutCopyMode = False
                print(f"  moved Pick List: rows {pick_row}-{pick_end} -> {INSERT_AT}-{INSERT_AT+n_pick-1}")
            else:
                print(f"  Pick List already at row {INSERT_AT}")

            # 6. print area down to the bottom of the relocated Pick List;
            #    landscape Letter, fit to one page. Placed Packs now sits below and is excluded.
            pick_bottom = INSERT_AT + n_pick - 1
            ps = mf.PageSetup
            ps.PrintArea = f"A1:L{pick_bottom}"
            ps.Orientation = 2        # xlLandscape
            ps.PaperSize = 1          # xlPaperLetter
            ps.LeftMargin = excel.InchesToPoints(0.4)
            ps.RightMargin = excel.InchesToPoints(0.4)
            ps.TopMargin = excel.InchesToPoints(0.35)
            ps.BottomMargin = excel.InchesToPoints(0.35)
            ps.HeaderMargin = excel.InchesToPoints(0.15)
            ps.FooterMargin = excel.InchesToPoints(0.15)
            # FitTo via VBA — pywin32 dynamic dispatch can't set FitToPagesTall directly
            excel.Run("SetPrintFitToPage")
            print(f"  print area A1:L{pick_bottom}, landscape letter, fit to 1 page")

        wb.Save()
        has_vba = bool(wb.HasVBProject)
        print("saved; HasVBProject =", has_vba)
        wb.Close(False)
        print("DONE")
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
