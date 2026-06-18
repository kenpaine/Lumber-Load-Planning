#!/usr/bin/env python3
r"""
build_asym_planner.py - rebuild Centerbeam_Lumber_Layout_Planner.xlsm for the
asymmetric (per-side) car: a "Car layout" config (5+5 / 7+5 / 7+7) in C5 and a
per-line "Side" column (M). Injects the updated CenterbeamSolver + ManifestDiagram
VBA, re-solves the symmetric example, reseeds the manifest, saves; then opens a
throwaway pass to functionally verify a mixed 7+5 solve.

Requires pywin32 + Excel; the .xlsm must be CLOSED (it can run while Excel is open
on OTHER files - DispatchEx uses an isolated instance).
Run: "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" source\build_asym_planner.py
"""
import os
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
XLSM = os.path.join(REPO, "Centerbeam_Lumber_Layout_Planner.xlsm")
SOLVER = os.path.join(REPO, "CenterbeamSolver.bas")
MANIFEST = os.path.join(HERE, "manifest_diagram.bas")

XL_LIST, XL_STOP, XL_BETWEEN = 3, 1, 1


def read_code(path):
    with open(path, "r", encoding="utf-8") as fh:
        lines = fh.read().splitlines()
    lines = [ln for ln in lines if not ln.startswith("Attribute VB_Name")]
    return "\r\n".join(lines)


def inject(vbp, name, code):
    for c in list(vbp.VBComponents):
        if c.Name == name:
            vbp.VBComponents.Remove(c)
    comp = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = name
    comp.CodeModule.AddFromString(code)


def setval(rng, formula):
    try:
        rng.Validation.Delete()
    except Exception:
        pass
    rng.Validation.Add(XL_LIST, XL_STOP, XL_BETWEEN, formula)
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True


def row_products(ws, r):
    out = []
    for s in range(9):
        v = str(ws.Cells(r, 2 + s).Value or "").strip()
        if v:
            out.append(v.split("-")[0])
    return out


def main():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass
    try:
        # ===== Phase 1: modify + save =====
        wb = excel.Workbooks.Open(XLSM)
        try:
            ws = wb.Worksheets("Planner")
            cur = str(ws.Range("C5").Value).strip()
            ws.Range("A5").Value = "Car layout:"
            setval(ws.Range("C5"), "5+5,7+5,7+7")
            if cur in ("10", "10.0"):
                ws.Range("C5").Value = "5+5"
            elif cur in ("14", "14.0"):
                ws.Range("C5").Value = "7+7"
            elif cur not in ("5+5", "7+5", "7+7"):
                ws.Range("C5").Value = "5+5"
            print("  C5 (car layout) ->", ws.Range("C5").Value)

            ws.Range("N9").Value = "Side"
            ws.Range("N9").Font.Bold = True
            setval(ws.Range("N10:N29"), "7-row side,5-row side")
            ws.Range("N10:N29").Interior.Color = 0xF0F4E8
            ws.Columns("N").ColumnWidth = 12
            print("  added Side column (N)")

            vbp = wb.VBProject
            inject(vbp, "CenterbeamSolver", read_code(SOLVER))
            inject(vbp, "ManifestDiagram", read_code(MANIFEST))
            print("  injected CenterbeamSolver + ManifestDiagram")

            excel.Run("ApplySingleMode")
            excel.Run("SolveLayoutQuiet")
            full = sum(1 for r in range(34, 48) if ws.Cells(r, 11).Value == 72)
            print("  symmetric example re-solved; full 72-ft rows =", full)

            wb.Worksheets("Manifest").Activate()
            excel.Run("DrawManifestDiagram")
            ndiag = sum(1 for i in range(1, wb.Worksheets("Manifest").Shapes.Count + 1)
                        if str(wb.Worksheets("Manifest").Shapes(i).Name).startswith("diag_"))
            print("  manifest reseeded; diagram shapes =", ndiag)
            ws.Activate()
            wb.Save()
            print("  Phase 1 saved.")
        finally:
            wb.Close(False)

        # ===== Phase 2: functionally verify a mixed 7+5 solve (discarded) =====
        wb2 = excel.Workbooks.Open(XLSM)
        try:
            ws = wb2.Worksheets("Planner")
            ws.Range("C7").Value = "No"
            excel.Run("ApplySingleMode")
            ws.Range("B10:E29").ClearContents()      # inputs (B-E are not merged)
            ws.Range("N10:N29").ClearContents()      # Side column (N is free of the status box)
            ws.Range("C5").Value = "7+5"
            ws.Cells(10, 2).Value = "4x4"; ws.Cells(10, 3).Value = 12
            ws.Cells(10, 4).Value = "2";   ws.Cells(10, 5).Value = 42
            ws.Cells(10, 14).Value = "7-row side"      # 7 x 72 = 504 ft
            ws.Cells(11, 2).Value = "2x8"; ws.Cells(11, 3).Value = 12
            ws.Cells(11, 4).Value = "MSR"; ws.Cells(11, 5).Value = 30
            ws.Cells(11, 14).Value = "5-row side"      # 5 x 72 = 360 ft
            excel.Run("SolveLayoutQuiet")

            sideA, sideB = set(), set()
            for r in range(34, 41):          # rows 1-7 = the 7-row side
                sideA.update(row_products(ws, r))
            for r in range(41, 46):          # rows 8-12 = the 5-row side
                sideB.update(row_products(ws, r))
            totals = [ws.Cells(r, 11).Value for r in range(34, 46)]
            print("VERIFY mixed 7+5:")
            print("  side 1 (rows 1-7) products :", sorted(sideA))
            print("  side 2 (rows 8-12) products:", sorted(sideB))
            print("  row totals (1-12)          :", totals)
            ok = sideA == {"4x4"} and sideB == {"2x8"} and all(t == 72 for t in totals)
            print("  RESULT:", "PASS" if ok else "FAIL")
        finally:
            wb2.Close(False)   # discard the test scenario
    finally:
        excel.Quit()
    print("DONE")


if __name__ == "__main__":
    main()
