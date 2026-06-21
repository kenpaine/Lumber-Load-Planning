"""Build Lumber_Loader.xlsm — the single combined workbook.

Starts from the Layout Planner workbook (the "Loader" — it owns the solver,
Manifest and Pattern Library), copies the Tally's two sheets to the front (Tally
comes first), imports the TallyRecommender module, and injects the LoaderTransfer
module — a "Send tally to Loader" button that takes the recommended tally on the
active row and lays it out in the Loader grid (the Excel twin of the web app's
"Load into car").
"""
import os
import shutil
from win32com.client import DispatchEx

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLANNER = os.path.join(ROOT, "Centerbeam_Lumber_Layout_Planner.xlsm")
TALLY = os.path.join(ROOT, "Centerbeam_Tally_Recommender.xlsm")
OUT = os.path.join(ROOT, "Lumber_Loader.xlsm")
BAS = os.path.join(ROOT, "source", "tally_recommender.bas")
XL_MACRO = 52  # xlOpenXMLWorkbookMacroEnabled

# The transfer macro: read the recommended tally on the active row of the Tally
# sheet and lay it into the Loader grid, then solve. Active-row pick.
TRANSFER_VBA = r'''Option Explicit

' ===== Tally -> Loader transfer (the Excel twin of the web "Load into car") =====
Private Const TALLY_SHEET As String = "Tally"
Private Const LOADER_SHEET As String = "Loader"
Private Const CAR_CELL As String = "B5"       ' Tally car-layout cell
Private Const RECA_TOP As Long = 23           ' table A (symmetric, or 7-row side) rows 23..33
Private Const RECA_BOT As Long = 33
Private Const RECB_TOP As Long = 37           ' table B (5-row side, mixed) rows 37..47
Private Const RECB_BOT As Long = 47
Private Const PLI_FIRST As Long = 10          ' Loader line-item rows 10..29
Private Const PLI_LAST As Long = 29
Private Const DEF_PROD As String = "2x6"      ' default product/grade (tally is length-only)
Private Const DEF_GRADE As String = "2"

Public Sub SendTallyToLoader()
    Dim t As Worksheet, ld As Worksheet
    Set t = ThisWorkbook.Worksheets(TALLY_SHEET)
    Set ld = ThisWorkbook.Worksheets(LOADER_SHEET)

    If ActiveSheet.Name <> TALLY_SHEET Then t.Activate
    Dim r As Long: r = ActiveCell.Row
    Dim isA As Boolean, isB As Boolean
    isA = (r >= RECA_TOP And r <= RECA_BOT)
    isB = (r >= RECB_TOP And r <= RECB_BOT)
    If Not (isA Or isB) Then
        MsgBox "Click a cell in a recommended-tally row first (a numbered row under a TALLIES header), then Send to Loader.", vbExclamation, "Send to Loader"
        Exit Sub
    End If
    If Len(Trim(CStr(t.Cells(r, 2).Value))) = 0 Then
        MsgBox "That row has no tally on it. Pick a numbered tally row.", vbExclamation, "Send to Loader"
        Exit Sub
    End If

    Dim lengths As Variant: lengths = Array(8, 10, 12, 14, 16, 18, 20)
    Dim cnt(0 To 6) As Long, i As Long, anyCnt As Boolean
    For i = 0 To 6
        cnt(i) = CLng(Val(CStr(t.Cells(r, 3 + i).Value)))   ' cols C..I = 8'..20'
        If cnt(i) > 0 Then anyCnt = True
    Next i
    If Not anyCnt Then
        MsgBox "That tally row has no piece counts to load.", vbExclamation, "Send to Loader"
        Exit Sub
    End If

    ' B5 is a descriptive string ("720 ft - 5+5 (10 rows)"); pull the layout token out
    ' so we write the plain "5+5"/"7+5"/"7+7" the Loader's C5 expects.
    Dim raw As String: raw = CStr(t.Range(CAR_CELL).Value)
    Dim carT As String
    If InStr(raw, "7+5") > 0 Then
        carT = "7+5"
    ElseIf InStr(raw, "7+7") > 0 Then
        carT = "7+7"
    Else
        carT = "5+5"
    End If
    Dim mixed As Boolean: mixed = (carT = "7+5")
    Dim sideTag As String
    If Not mixed Then
        sideTag = ""
    ElseIf isA Then
        sideTag = "7-row"
    Else
        sideTag = "5-row"
    End If

    Application.EnableEvents = False
    On Error GoTo cleanup
    ld.Range("C5").Value = carT

    ' Clear the rows this transfer owns: symmetric -> all rows; mixed -> everything
    ' except the OTHER side (so leftover/seed blank-side rows get cleared too, and the
    ' side we're not touching is preserved).
    Dim keepTag As String
    If mixed Then keepTag = IIf(sideTag = "7-row", "5-row", "7-row")
    Dim row As Long
    For row = PLI_FIRST To PLI_LAST
        If (Not mixed) Or (CStr(ld.Cells(row, 1).Value) <> keepTag) Then
            ld.Range(ld.Cells(row, 1), ld.Cells(row, 5)).ClearContents   ' inputs only; F/G/H are formulas
        End If
    Next row

    ' write the tally's lengths into the first empty input rows
    Dim wrow As Long: wrow = PLI_FIRST
    For i = 0 To 6
        If cnt(i) > 0 Then
            Do While wrow <= PLI_LAST And Len(Trim(CStr(ld.Cells(wrow, 2).Value))) > 0
                wrow = wrow + 1
            Loop
            If wrow > PLI_LAST Then Exit For
            If mixed Then ld.Cells(wrow, 1).Value = sideTag        ' A = Side
            ld.Cells(wrow, 2).Value = DEF_PROD                     ' B = Product
            ld.Cells(wrow, 3).Value = lengths(i)                   ' C = Length
            ld.Cells(wrow, 4).Value = DEF_GRADE                    ' D = Grade
            ld.Cells(wrow, 5).Value = cnt(i)                       ' E = Packs (F/G/H compute from formulas)
            wrow = wrow + 1
        End If
    Next i

cleanup:
    Application.EnableEvents = True
    SolveLayoutQuiet
    ld.Activate
    ld.Range("A1").Select
End Sub
'''


def add_module(wb, name, code):
    comp = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = name
    if comp.CodeModule.CountOfLines:  # drop any auto-inserted "Option Explicit"
        comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
    comp.CodeModule.AddFromString(code)


def replace_in_module(wb, name, find, repl):
    """In-place text replace inside an existing standard module's code."""
    cm = wb.VBProject.VBComponents(name).CodeModule
    n = cm.CountOfLines
    code = cm.Lines(1, n) if n else ""
    new = code.replace(find, repl)
    if new != code:
        cm.DeleteLines(1, n)
        cm.AddFromString(new)


def main():
    shutil.copyfile(PLANNER, OUT)
    with open(BAS, "r", encoding="utf-8") as f:
        tally_code = f.read()

    xl = DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    try:
        wb = xl.Workbooks.Open(OUT)
        src = xl.Workbooks.Open(TALLY, ReadOnly=True)
        src.Sheets("Row Patterns").Copy(Before=wb.Sheets(1))
        src.Sheets("Recommender").Copy(Before=wb.Sheets(1))
        src.Close(SaveChanges=False)

        # Renames: Recommender -> Tally, Planner -> Loader. Renaming a sheet auto-updates
        # all *formula* references; VBA *string literals* are updated by hand here.
        wb.Worksheets("Recommender").Name = "Tally"
        wb.Worksheets("Planner").Name = "Loader"
        replace_in_module(wb, "CenterbeamSolver", '"Planner"', '"Loader"')  # 5x Sheets("Planner")
        replace_in_module(wb, "ManifestDiagram", '"Planner"', '"Loader"')   # 1x Worksheets("Planner")
        tally_code = tally_code.replace(
            'Const REC_SHEET As String = "Recommender"',
            'Const REC_SHEET As String = "Tally"')

        add_module(wb, "TallyRecommender", tally_code)
        add_module(wb, "LoaderTransfer", TRANSFER_VBA)

        # "Send tally to Loader" button on the Tally sheet, right of the tally table
        rec = wb.Worksheets("Tally")
        anchor = rec.Range("N21")
        btn = rec.Buttons().Add(anchor.Left, anchor.Top, 165, 34)
        btn.Name = "btnSendToLoader"
        btn.Caption = "Send tally to Loader  >>"
        btn.OnAction = "SendTallyToLoader"

        # compile check
        xl.EnableEvents = True
        compiled = True
        try:
            xl.Run("ApplyTallyPalette")
        except Exception as e:
            compiled = False
            print("  ! ApplyTallyPalette failed:", e)
        xl.EnableEvents = False

        wb.Save()
        sheets = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
        modules = [c.Name for c in wb.VBProject.VBComponents]
        wb.Close(SaveChanges=True)
        print("compiled:", compiled)
        print("sheets:", sheets)
        print("modules:", modules)
    finally:
        xl.Quit()
    print("built ->", OUT, "(%.0f KB)" % (os.path.getsize(OUT) / 1024))


if __name__ == "__main__":
    main()
