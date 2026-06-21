"""Build Lumber_Loader.xlsm — the single combined workbook.

Starts from the Layout Planner workbook (the "Loader" — it owns the solver,
Manifest and Pattern Library), copies the Tally's two sheets to the front (Tally
comes first), renames the sheets (Recommender->Tally, Planner->Loader), imports
the TallyRecommender module, and injects LoaderTransfer — two "Send to Loader"
buttons (the Excel twin of the web app's "Load into car"):

  * Tally sheet  -> SendTallyToLoader     (the recommended tally on the active row)
  * Row Patterns -> SendHandBuildToLoader (the hand-built tally totals)
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

TRANSFER_VBA = r'''Option Explicit

' ===== Tally -> Loader transfer (the Excel twin of the web "Load into car") =====
Private Const TALLY_SHEET As String = "Tally"
Private Const PAT_SHEET As String = "Row Patterns"
Private Const LOADER_SHEET As String = "Loader"
Private Const CAR_CELL As String = "B5"       ' Tally car-layout cell
Private Const RECA_TOP As Long = 23           ' recommended table A rows 23..33 (symmetric / 7-row)
Private Const RECA_BOT As Long = 33
Private Const RECB_TOP As Long = 37           ' recommended table B rows 37..47 (5-row, mixed)
Private Const RECB_BOT As Long = 47
Private Const PLI_FIRST As Long = 10          ' Loader line-item rows 10..29
Private Const PLI_LAST As Long = 29
Private Const DEF_PROD As String = "2x6"      ' default product/grade (a tally is length-only)
Private Const DEF_GRADE As String = "2"

' Layout token ("5+5"/"7+5"/"7+7") parsed out of the Tally's descriptive B5 string.
Private Function CarToken() As String
    Dim raw As String
    raw = CStr(ThisWorkbook.Worksheets(TALLY_SHEET).Range(CAR_CELL).Value)
    If InStr(raw, "7+5") > 0 Then
        CarToken = "7+5"
    ElseIf InStr(raw, "7+7") > 0 Then
        CarToken = "7+7"
    Else
        CarToken = "5+5"
    End If
End Function

' Read the 7 per-length piece counts (cols C..I = 8'..20') from row r of ws.
' Returns True if any count > 0.
Private Function ReadCounts(ws As Worksheet, ByVal r As Long, ByRef cnt() As Long) As Boolean
    Dim i As Long
    ReadCounts = False
    For i = 0 To 6
        cnt(i) = CLng(Val(CStr(ws.Cells(r, 3 + i).Value)))
        If cnt(i) > 0 Then ReadCounts = True
    Next i
End Function

' Clear Loader input rows. keepTag="" clears all; otherwise keeps rows whose Side = keepTag.
Private Sub ClearLoaderLines(ld As Worksheet, ByVal keepTag As String)
    Dim row As Long
    For row = PLI_FIRST To PLI_LAST
        If keepTag = "" Or CStr(ld.Cells(row, 1).Value) <> keepTag Then
            ld.Range(ld.Cells(row, 1), ld.Cells(row, 5)).ClearContents   ' inputs only; F/G/H are formulas
        End If
    Next row
End Sub

' Write a length-count set into the first empty Loader rows (no clearing here).
Private Sub WriteCounts(ld As Worksheet, cnt() As Long, ByVal mixed As Boolean, ByVal sideTag As String)
    Dim lengths As Variant: lengths = Array(8, 10, 12, 14, 16, 18, 20)
    Dim i As Long, wrow As Long: wrow = PLI_FIRST
    For i = 0 To 6
        If cnt(i) > 0 Then
            Do While wrow <= PLI_LAST And Len(Trim(CStr(ld.Cells(wrow, 2).Value))) > 0
                wrow = wrow + 1
            Loop
            If wrow > PLI_LAST Then Exit For
            If mixed Then ld.Cells(wrow, 1).Value = sideTag
            ld.Cells(wrow, 2).Value = DEF_PROD
            ld.Cells(wrow, 3).Value = lengths(i)
            ld.Cells(wrow, 4).Value = DEF_GRADE
            ld.Cells(wrow, 5).Value = cnt(i)
            wrow = wrow + 1
        End If
    Next i
End Sub

' Solve, re-enable events, and (only when the car is complete) jump to the Loader.
Private Sub FinishLoad(ld As Worksheet, ByVal jumpToLoader As Boolean)
    Application.EnableEvents = True
    SolveLayoutQuiet
    If jumpToLoader Then
        ld.Activate
        ld.Range("A1").Select
    End If
End Sub

' Does any Loader line carry this side tag (i.e. has that side already been loaded)?
Private Function SideHasInventory(ld As Worksheet, ByVal sideTag As String) As Boolean
    Dim row As Long
    For row = PLI_FIRST To PLI_LAST
        If CStr(ld.Cells(row, 1).Value) = sideTag And Len(Trim(CStr(ld.Cells(row, 2).Value))) > 0 Then
            SideHasInventory = True
            Exit Function
        End If
    Next row
End Function

' ---- Send the recommended tally on the active row of the Tally sheet ----
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

    Dim cnt(0 To 6) As Long
    If Not ReadCounts(t, r, cnt) Then
        MsgBox "That tally row has no piece counts to load.", vbExclamation, "Send to Loader"
        Exit Sub
    End If

    Dim carT As String: carT = CarToken()
    Dim mixed As Boolean: mixed = (carT = "7+5")
    Dim sideTag As String, keepTag As String
    If Not mixed Then
        sideTag = "": keepTag = ""
    ElseIf isA Then
        sideTag = "7-row": keepTag = "5-row"
    Else
        sideTag = "5-row": keepTag = "7-row"
    End If

    Application.EnableEvents = False
    On Error GoTo done
    ld.Range("C5").Value = carT
    ClearLoaderLines ld, keepTag
    WriteCounts ld, cnt, mixed, sideTag
    Dim jump As Boolean
    If Not mixed Then
        jump = True                            ' symmetric: one send = the whole car
    Else
        jump = SideHasInventory(ld, keepTag)   ' compound: jump only once BOTH sides are loaded
    End If
done:
    FinishLoad ld, jump
End Sub

' ---- Send the hand-built tally from the Row Patterns tab (the whole car) ----
Public Sub SendHandBuildToLoader()
    Dim pat As Worksheet, ld As Worksheet
    Set pat = ThisWorkbook.Worksheets(PAT_SHEET)
    Set ld = ThisWorkbook.Worksheets(LOADER_SHEET)

    ' Find the hand-build total row(s): "YOUR HAND-BUILT TALLY ..." in col B.
    ' One on a symmetric car; two on a mixed car (7-row grid first, then 5-row).
    Dim totRow(0 To 3) As Long, nTot As Long: nTot = 0
    Dim r As Long, lastR As Long
    lastR = pat.Cells(pat.Rows.Count, 2).End(xlUp).Row
    For r = 1 To lastR
        If InStr(1, CStr(pat.Cells(r, 2).Value), "YOUR HAND-BUILT TALLY", vbTextCompare) > 0 Then
            If nTot <= 3 Then totRow(nTot) = r
            nTot = nTot + 1
        End If
    Next r
    If nTot = 0 Then
        MsgBox "Hand-build a tally first: on the Row Patterns tab type 'Rows to use' for the patterns you want, then Send to Loader.", vbExclamation, "Send to Loader"
        Exit Sub
    End If

    Dim carT As String: carT = CarToken()
    Dim mixed As Boolean: mixed = (carT = "7+5")
    Dim cntA(0 To 6) As Long, cntB(0 To 6) As Long
    Dim anyA As Boolean, anyB As Boolean
    anyA = ReadCounts(pat, totRow(0), cntA)
    If mixed And nTot >= 2 Then anyB = ReadCounts(pat, totRow(1), cntB)
    If (Not anyA) And (Not anyB) Then
        MsgBox "The hand-built tally is empty -- set some 'Rows to use' first.", vbExclamation, "Send to Loader"
        Exit Sub
    End If

    Application.EnableEvents = False
    On Error GoTo done
    ld.Range("C5").Value = carT
    ClearLoaderLines ld, ""          ' the hand-build is the whole car -> clear everything
    If Not mixed Then
        WriteCounts ld, cntA, False, ""
    Else
        If anyA Then WriteCounts ld, cntA, True, "7-row"
        If anyB Then WriteCounts ld, cntB, True, "5-row"
    End If
done:
    FinishLoad ld, True          ' hand-build is always the whole car
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


def add_button(ws, anchor_cell, w, h, name, caption, action):
    a = ws.Range(anchor_cell)
    b = ws.Buttons().Add(a.Left, a.Top, w, h)
    b.Name = name
    b.Caption = caption
    b.OnAction = action


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

        # Send-to-Loader buttons: recommended tally (Tally sheet) + hand-built (Row Patterns).
        add_button(wb.Worksheets("Tally"), "N21", 165, 34,
                   "btnSendToLoader", "Send tally to Loader  >>", "SendTallyToLoader")
        add_button(wb.Worksheets("Row Patterns"), "L1", 205, 30,
                   "btnSendHandBuild", "Send hand-built tally to Loader  >>", "SendHandBuildToLoader")

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
