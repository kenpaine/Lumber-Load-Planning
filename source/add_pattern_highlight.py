#!/usr/bin/env python3
r"""
add_pattern_highlight.py - click-to-highlight on the Layout Planner workbook.

Adds the "active row" highlight to the **Pattern Library** sheet of
../Centerbeam_Lumber_Layout_Planner.xlsm: clicking any cell in a pattern row
paints that whole row light blue (BDD7EE) until you click another row, matching
the Tally Recommender workbook and both browser apps.

It is conditional formatting tied to a hidden helper cell $N$1 that a
Worksheet_SelectionChange on the Pattern Library sheet keeps in sync (0 = none).

Idempotent - safe to re-run (it clears its own CF and only injects the event
once). Only the Pattern Library sheet is touched; the Planner input grid is left
alone (it has the merged Load Status box + its own conditional formats, and
Excel already highlights the active cell there). The macro-free .xlsx is
intentionally skipped - with no events it can't track the cursor.

Requires Excel with "Trust access to the VBA project object model" (already on).
Run with the real interpreter:
  "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" source\add_pattern_highlight.py
"""
import os

import pythoncom
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
XLSM = os.path.join(REPO, "Centerbeam_Lumber_Layout_Planner.xlsm")
SHEET = "Pattern Library"
LIGHTBLUE = 15652797   # BGR of BDD7EE (matches the Tally workbook's active row)
TOP, BOT = 5, 107      # the pattern data rows (col A = pattern number)

SEL_EVENT = "\r\n".join([
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)",
    "    Dim rw As Long: rw = Target.Row",
    "    Dim newv As Long: newv = 0",
    "    If rw >= 5 And rw <= 107 And IsNumeric(Me.Cells(rw, 1).Value) _",
    "       And Len(CStr(Me.Cells(rw, 1).Value)) > 0 Then newv = rw",
    '    If Me.Range("N1").Value <> newv Then Me.Range("N1").Value = newv',
    "End Sub",
])


def main():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    wb = None
    try:
        wb = excel.Workbooks.Open(XLSM)
        ws = wb.Worksheets(SHEET)

        # hidden helper cell holding the active row (0 = none)
        ws.Range("N1").Value = 0
        ws.Range("N1").NumberFormat = ";;;"

        # conditional format: light-blue fill on the row where ROW() = $N$1
        rng = ws.Range("A%d:J%d" % (TOP, BOT))
        rng.FormatConditions.Delete()                       # idempotent
        # 2 = xlExpression; Operator is unused for an expression rule, so pass the
        # COM "missing" sentinel positionally (keyword args fail under late binding).
        fc = rng.FormatConditions.Add(2, pythoncom.Empty, "=ROW()=$N$1")
        fc.Interior.Color = LIGHTBLUE
        fc.StopIfTrue = True

        # inject the SelectionChange into the sheet's own code module (once)
        cm = wb.VBProject.VBComponents(ws.CodeName).CodeModule
        body = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        added = "Worksheet_SelectionChange" not in body
        if added:
            cm.AddFromString(SEL_EVENT)

        wb.Save()
        print("OK: active-row highlight on '%s' (SelectionChange added: %s)" % (SHEET, added))
    finally:
        if wb is not None:
            wb.Close(False)
        excel.Quit()


if __name__ == "__main__":
    main()
