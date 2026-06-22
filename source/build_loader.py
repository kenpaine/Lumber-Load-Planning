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
import pythoncom
from win32com.client import DispatchEx

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HERE = os.path.dirname(os.path.abspath(__file__))
PLANNER = os.path.join(HERE, "Centerbeam_Lumber_Layout_Planner.xlsm")
TALLY = os.path.join(HERE, "Centerbeam_Tally_Recommender.xlsm")
OUT = os.path.join(ROOT, "Lumber_Loader.xlsm")
BAS = os.path.join(HERE, "tally_recommender.bas")
XL_MACRO = 52  # xlOpenXMLWorkbookMacroEnabled

# Button styling — match the workbook's other macro buttons (navy rounded-rect shapes).
NAVY, WHITE = "1F4E79", "FFFFFF"
MSO_ROUNDED_RECT, XL_CENTER = 5, -4108


def bgr(hex_):
    """RRGGBB hex -> the BGR integer Excel COM expects for .RGB / .Color."""
    r, g, b = int(hex_[0:2], 16), int(hex_[2:4], 16), int(hex_[4:6], 16)
    return r + (g << 8) + (b << 16)

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

' ---- Send recommended tallies to the Loader ----
' Symmetric car: the active row's tally fills the whole car.
' Compound (7+5) car: the user picks a row in EACH table -- the 7-row table is
' tracked in N1, the 5-row table in N2 -- and one Send loads both sides.
Public Sub SendTallyToLoader()
    Dim t As Worksheet, ld As Worksheet
    Set t = ThisWorkbook.Worksheets(TALLY_SHEET)
    Set ld = ThisWorkbook.Worksheets(LOADER_SHEET)
    If ActiveSheet.Name <> TALLY_SHEET Then t.Activate

    Dim carT As String: carT = CarToken()

    If carT <> "7+5" Then
        ' --- symmetric: send the active row's tally ---
        Dim r As Long: r = ActiveCell.Row
        Dim cnt(0 To 6) As Long
        If Not (r >= RECA_TOP And r <= RECA_BOT) _
           Or Len(Trim(CStr(t.Cells(r, 2).Value))) = 0 Or Not ReadCounts(t, r, cnt) Then
            MsgBox "Click a cell in a recommended-tally row first, then Send to Loader.", vbExclamation, "Send to Loader"
            Exit Sub
        End If
        Application.EnableEvents = False
        On Error GoTo doneSym
        ld.Range("C5").Value = carT
        ClearLoaderLines ld, ""
        WriteCounts ld, cnt, False, ""
doneSym:
        FinishLoad ld, True
        Exit Sub
    End If

    ' --- compound 7+5: load BOTH tables' picks (N1 = 7-row table, N2 = 5-row table) ---
    Dim p7 As Long, p5 As Long
    p7 = CLng(Val(CStr(t.Range("N1").Value)))
    p5 = CLng(Val(CStr(t.Range("N2").Value)))
    Dim c7(0 To 6) As Long, c5(0 To 6) As Long
    Dim has7 As Boolean, has5 As Boolean
    If p7 >= RECA_TOP And p7 <= RECA_BOT Then has7 = ReadCounts(t, p7, c7)
    If p5 >= RECB_TOP And p5 <= RECB_BOT Then has5 = ReadCounts(t, p5, c5)
    If (Not has7) And (Not has5) Then
        MsgBox "Compound car: click a tally row in the 7-row table AND a row in the 5-row table, then Send to Loader.", vbExclamation, "Send to Loader"
        Exit Sub
    End If

    Application.EnableEvents = False
    On Error GoTo doneMix
    ld.Range("C5").Value = carT
    ClearLoaderLines ld, ""              ' both picks together are the whole car
    If has7 Then WriteCounts ld, c7, True, "7-row"
    If has5 Then WriteCounts ld, c5, True, "5-row"
doneMix:
    FinishLoad ld, (has7 And has5)       ' jump to the Loader once both sides are picked
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

# Replaces the copied Tally sheet's selection event so each recommended-tally table
# keeps its OWN highlighted pick (7-row table -> N1, 5-row table -> N2), letting you
# select a tally from both tables on a compound car. The reactive Worksheet_Change
# (palette / car-layout edits re-run the recommender) is preserved unchanged.
TALLY_SHEET_CODE = "\r\n".join([
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)",
    "    Dim rw As Long: rw = Target.Row",
    "    If rw >= 23 And rw <= 33 And IsNumeric(Me.Cells(rw, 1).Value) And Len(CStr(Me.Cells(rw, 1).Value)) > 0 Then",
    '        If Me.Range("N1").Value <> rw Then Me.Range("N1").Value = rw',
    "    ElseIf rw >= 37 And rw <= 47 And IsNumeric(Me.Cells(rw, 1).Value) And Len(CStr(Me.Cells(rw, 1).Value)) > 0 Then",
    '        If Me.Range("N2").Value <> rw Then Me.Range("N2").Value = rw',
    "    End If",
    "    If Target.Cells.Count <> 1 Then Exit Sub",
    "    If Target.Column < 3 Or Target.Column > 9 Then Exit Sub",
    "    If Target.Row = 22 Then SortRecsByRange Target.Column, 23, 33",
    "    If Target.Row = 36 Then SortRecsByRange Target.Column, 37, 47",
    "End Sub",
    "Private Sub Worksheet_Change(ByVal Target As Range)",
    '    If Not Intersect(Target, Me.Range("B8")) Is Nothing Then',
    '        ApplyPaletteEverywhere CStr(Me.Range("B8").Value)',
    "        Exit Sub",
    "    End If",
    '    If Intersect(Target, Me.Range("B5:B6")) Is Nothing Then Exit Sub',
    "    Application.EnableEvents = False",
    "    RecommendTallies",
    "    Application.EnableEvents = True",
    "End Sub",
])

# One colour scheme everywhere: both pickers (Tally B8 / Manifest N2) call this, which
# sets both, recolours the Tally + Row Patterns, the Loader grid's length-fill CF rules,
# and redraws the Manifest. Self-contained (combined workbook only) so the shared .bas
# files and the standalone Tally/Planner are untouched.
UNIFIED_PALETTE_VBA = r'''Option Explicit

Private Function UScheme(ByVal pal As Long, ByVal idx As Long) As Long
    Dim P As Variant
    P = Array( _
      Array("BBD4EA", "C2E5C9", "FBE7B2", "E5CDEE", "FAC4BC", "BFEAE0", "DCE7AE"), _
      Array("E53935", "FB8C00", "FDD835", "43A047", "00ACC1", "1E88E5", "8E24AA"), _
      Array("EF5350", "FFA726", "FFEE58", "66BB6A", "26C6DA", "42A5F5", "AB47BC"), _
      Array("4E79A7", "F28E2B", "E15759", "76B7B2", "59A14F", "EDC948", "B07AA1"), _
      Array("8C6239", "C9A66B", "7D8C4F", "B7410E", "2E5E4E", "D4A017", "5C3A21"), _
      Array("B71C1C", "E65100", "F9A825", "1B5E20", "00695C", "1A237E", "4A148C"), _
      Array("D32F2F", "F57C00", "FBC02D", "689F38", "0097A7", "1976D2", "7B1FA2"), _
      Array("440154", "443983", "31688E", "21918C", "35B779", "90D743", "FDE725"), _
      Array("5C1A33", "9D2B4A", "C44536", "E8590C", "F0A202", "F4C430", "F7E07A"), _
      Array("FF1744", "FF9100", "FFEA00", "00E676", "00E5FF", "2979FF", "D500F9"), _
      Array("1F77B4", "2CA02C", "FF7F0E", "9467BD", "D62728", "17BECF", "8C564B"), _
      Array("FFFFFF", "EEEEEE", "DDDDDD", "CCCCCC", "BBBBBB", "AAAAAA", "999999"))
    If pal < 0 Or pal > 11 Then pal = 0
    Dim h As String: h = P(pal)(idx)
    UScheme = RGB(CLng("&H" & Mid(h, 1, 2)), CLng("&H" & Mid(h, 3, 2)), CLng("&H" & Mid(h, 5, 2)))
End Function

Private Function USchemeIndex(ByVal nm As String) As Long
    Dim names As Variant
    names = Array("Color (pastel)", "Vivid", "Material", "Tableau", "Earth / lumberyard", _
        "Jewel tones", "Rainbow (warm to cool)", "Viridis (colour-safe)", "Sunset (warm)", _
        "Neon", "High contrast", "B & W (print)")
    Dim i As Long
    For i = 0 To 11
        If StrComp(Trim(nm), CStr(names(i)), vbTextCompare) = 0 Then USchemeIndex = i: Exit Function
    Next i
    USchemeIndex = 0
End Function

' Recolour the Loader grid's 7 length-fill conditional-format rules to the scheme.
Private Sub RecolorLoaderGrid(ld As Worksheet, ByVal pal As Long)
    Dim lengths As Variant: lengths = Array(8, 10, 12, 14, 16, 18, 20)
    Dim fcs As Object: Set fcs = ld.Range("B34:J47").FormatConditions
    Dim i As Long, k As Long, f As String
    For i = 1 To fcs.Count
        f = ""
        On Error Resume Next
        f = fcs(i).Formula1
        On Error GoTo 0
        If Len(f) > 0 Then
            For k = 0 To 6
                If InStr(f, "=""" & lengths(k) & """") > 0 Then
                    On Error Resume Next
                    fcs(i).Interior.Color = UScheme(pal, k)
                    On Error GoTo 0
                    Exit For
                End If
            Next k
        End If
    Next i
End Sub

' Black or white text, whichever reads better on the given fill (luminance).
Private Function UBestText(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&: g = (c \ &H100) And &HFF&: b = (c \ &H10000) And &HFF&
    If (0.299 * r + 0.587 * g + 0.114 * b) > 150 Then UBestText = RGB(0, 0, 0) Else UBestText = RGB(255, 255, 255)
End Function

' Column number -> letter (A, B, ... AA ...).
Private Function UColL(ByVal n As Long) As String
    Dim s As String, m As Long
    Do While n > 0
        m = (n - 1) Mod 26
        s = Chr(65 + m) & s
        n = (n - 1) \ 26
    Loop
    UColL = s
End Function

' Recolour the Pattern Library's length cells (header, count cells, legend) to the
' scheme so it follows the palette like the Tally chips and the Loader grid.
Private Sub RecolorPatternLibrary(ByVal pal As Long)
    Dim pl As Worksheet
    On Error Resume Next
    Set pl = ThisWorkbook.Worksheets("Pattern Library")
    On Error GoTo 0
    If pl Is Nothing Then Exit Sub
    Dim j As Long, col As Long, fillc As Long, r As Long
    For j = 0 To 6                                    ' header row 4, length cols C..I
        col = 3 + j: fillc = UScheme(pal, j)
        pl.Cells(4, col).Interior.Color = fillc
        pl.Cells(4, col).Font.Color = UBestText(fillc)
    Next j
    r = 5                                             ' data rows: colour the non-empty count cells
    Do While r < 500
        If Not IsNumeric(pl.Cells(r, 1).Value) Then Exit Do
        If Len(CStr(pl.Cells(r, 1).Value)) = 0 Then Exit Do
        For j = 0 To 6
            col = 3 + j
            If Len(CStr(pl.Cells(r, col).Value)) > 0 Then
                fillc = UScheme(pal, j)
                pl.Cells(r, col).Interior.Color = fillc
                pl.Cells(r, col).Font.Color = UBestText(fillc)
            End If
        Next j
        r = r + 1
    Loop
    Dim g As Long                                     ' colour the length legend just below the table
    For g = r To r + 4
        If InStr(1, CStr(pl.Cells(g, 2).Value), "Length colors", vbTextCompare) > 0 Then
            For j = 0 To 6
                col = 3 + j: fillc = UScheme(pal, j)
                pl.Cells(g, col).Interior.Color = fillc
                pl.Cells(g, col).Font.Color = UBestText(fillc)
            Next j
            Exit For
        End If
    Next g
End Sub

' Click a Pattern Library header (row 4, cols C..J) to sort by it; clicking the same
' column again flips ascending/descending. Sort state lives in hidden N2 (col) / N3 (dir).
Public Sub SortPatternLibrary(ByVal sortCol As Long)
    Dim pl As Worksheet: Set pl = ThisWorkbook.Worksheets("Pattern Library")
    Const HDR As Long = 4
    Dim firstR As Long: firstR = HDR + 1
    Dim lastR As Long: lastR = firstR - 1
    Dim r As Long: r = firstR
    Do While r < 500
        If Not IsNumeric(pl.Cells(r, 1).Value) Then Exit Do
        If Len(CStr(pl.Cells(r, 1).Value)) = 0 Then Exit Do
        lastR = r: r = r + 1
    Loop
    If lastR <= firstR Then Exit Sub
    Dim ord As Long
    If CLng(Val(pl.Range("N2").Value)) = sortCol And CStr(pl.Range("N3").Value) = "D" Then
        ord = xlAscending: pl.Range("N3").Value = "A"
    Else
        ord = xlDescending: pl.Range("N3").Value = "D"
    End If
    pl.Range("N2").Value = sortCol
    On Error GoTo cleanup
    Application.EnableEvents = False
    With pl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=pl.Range(UColL(sortCol) & firstR & ":" & UColL(sortCol) & lastR), _
                        SortOn:=xlSortOnValues, Order:=ord, DataOption:=xlSortNormal
        .SetRange pl.Range("B" & firstR & ":J" & lastR)
        .Header = xlNo
        .Apply
    End With
    Dim k As Long
    For k = firstR To lastR
        pl.Cells(k, 1).Value = k - firstR + 1                              ' renumber the # column
    Next k
    RecolorPatternLibrary USchemeIndex(CStr(ThisWorkbook.Worksheets("Loader").Range("J7").Value))
cleanup:
    Application.EnableEvents = True
End Sub

Public Sub ApplyPaletteEverywhere(ByVal schemeName As String)
    Dim t As Worksheet, mf As Worksheet, ld As Worksheet
    Set t = ThisWorkbook.Worksheets("Tally")
    Set mf = ThisWorkbook.Worksheets("Manifest")
    Set ld = ThisWorkbook.Worksheets("Loader")
    Dim ev As Boolean: ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    If CStr(t.Range("B8").Value) <> schemeName Then t.Range("B8").Value = schemeName
    If CStr(mf.Range("N2").Value) <> schemeName Then mf.Range("N2").Value = schemeName
    If CStr(ld.Range("J7").Value) <> schemeName Then ld.Range("J7").Value = schemeName
    ApplyTallyPalette
    RecolorLoaderGrid ld, USchemeIndex(schemeName)
    RecolorPatternLibrary USchemeIndex(schemeName)
    DrawManifestDiagram
    On Error GoTo 0
    Application.EnableEvents = ev
End Sub
'''

MANIFEST_SHEET_CODE = "\r\n".join([
    "Private Sub Worksheet_Activate()",
    "    On Error Resume Next",
    "    DrawManifestDiagram",
    "End Sub",
    "Private Sub Worksheet_Change(ByVal Target As Range)",
    '    If Intersect(Target, Me.Range("N2")) Is Nothing Then Exit Sub',
    '    ApplyPaletteEverywhere CStr(Me.Range("N2").Value)',
    "End Sub",
])

# Pattern Library sheet code: keep the active-row highlight (sets N1) AND make the
# header row (row 4) click-to-sort the table by the clicked length / Total column.
PATTERN_SHEET_CODE = "\r\n".join([
    "Private Sub Worksheet_SelectionChange(ByVal Target As Range)",
    "    Dim rw As Long: rw = Target.Row",
    "    Dim newv As Long: newv = 0",
    "    If rw >= 5 And rw <= 107 And IsNumeric(Me.Cells(rw, 1).Value) _",
    "       And Len(CStr(Me.Cells(rw, 1).Value)) > 0 Then newv = rw",
    '    If Me.Range("N1").Value <> newv Then Me.Range("N1").Value = newv',
    "    If Target.Cells.Count <> 1 Then Exit Sub",
    "    If rw = 4 And Target.Column >= 3 And Target.Column <= 10 Then SortPatternLibrary Target.Column",
    "End Sub",
])

# Loader sheet code: preserve the row-tag double-click + OnPlannerChange behaviour,
# but route the new J7 colour picker to ApplyPaletteEverywhere.
LOADER_SHEET_CODE = "\r\n".join([
    "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)",
    '    If Intersect(Target, Me.Range("A10:A29")) Is Nothing Then Exit Sub',
    "    Cancel = True",
    "    Application.EnableEvents = False",
    '    If InStr(CStr(Target.Value), "7") > 0 Then',
    '        Target.Value = "5-row"',
    "    Else",
    '        Target.Value = "7-row"',
    "    End If",
    "    Application.EnableEvents = True",
    "End Sub",
    "Private Sub Worksheet_Change(ByVal Target As Range)",
    '    If Not Intersect(Target, Me.Range("J7")) Is Nothing Then',
    '        ApplyPaletteEverywhere CStr(Me.Range("J7").Value)',
    "        Exit Sub",
    "    End If",
    "    OnPlannerChange Target",
    "End Sub",
])


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
    """A 'button' styled to match the workbook's other macro buttons: a navy
    rounded-rectangle shape with centered bold white text (Form Control buttons
    can't be coloured, so the rest of the workbook uses shapes via OnAction)."""
    a = ws.Range(anchor_cell)
    shp = ws.Shapes.AddShape(MSO_ROUNDED_RECT, a.Left, a.Top, w, h)
    shp.Name = name
    shp.Fill.Solid()
    shp.Fill.ForeColor.RGB = bgr(NAVY)
    shp.Line.Visible = False
    tf = shp.TextFrame
    tf.Characters().Text = caption
    tf.Characters().Font.Bold = True
    tf.Characters().Font.Size = 11
    tf.Characters().Font.Color = bgr(WHITE)
    tf.HorizontalAlignment = XL_CENTER
    tf.VerticalAlignment = XL_CENTER
    shp.OnAction = action


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
        add_module(wb, "UnifiedPalette", UNIFIED_PALETTE_VBA)
        # One palette everywhere: re-point the Manifest's N2 Change at ApplyPaletteEverywhere
        # (the Tally B8 Change is already pointed there via TALLY_SHEET_CODE).
        mfsheet = wb.Worksheets("Manifest")
        msm = wb.VBProject.VBComponents(mfsheet.CodeName).CodeModule
        if msm.CountOfLines:
            msm.DeleteLines(1, msm.CountOfLines)
        msm.AddFromString(MANIFEST_SHEET_CODE)

        # Loader sheet: re-inject its code so the new J7 colour picker fires
        # ApplyPaletteEverywhere (keeping its row-tag double-click + OnPlannerChange).
        ldsheet = wb.Worksheets("Loader")
        lsm = wb.VBProject.VBComponents(ldsheet.CodeName).CodeModule
        if lsm.CountOfLines:
            lsm.DeleteLines(1, lsm.CountOfLines)
        lsm.AddFromString(LOADER_SHEET_CODE)

        # Pattern Library: keep its active-row highlight, add click-to-sort headers.
        # N1 = active row (already used); N2 = last sort column, N3 = direction (hidden).
        plsheet = wb.Worksheets("Pattern Library")
        for c in ("N2", "N3"):
            plsheet.Range(c).NumberFormat = ";;;"
        psm = wb.VBProject.VBComponents(plsheet.CodeName).CodeModule
        if psm.CountOfLines:
            psm.DeleteLines(1, psm.CountOfLines)
        psm.AddFromString(PATTERN_SHEET_CODE)

        # Loader colour picker at J7 (just under the Solve/Clear buttons), matching the
        # Tally B8 / Manifest N2 pickers. Label above it, list-validated dropdown below.
        SCHEME_NAMES = ("Color (pastel),Vivid,Material,Tableau,Earth / lumberyard,"
                        "Jewel tones,Rainbow (warm to cool),Viridis (colour-safe),"
                        "Sunset (warm),Neon,High contrast,B & W (print)")
        ldsheet.Range("J6:L6").Merge()
        ldsheet.Range("J6").Value = "Color scheme:"
        ldsheet.Range("J6").Font.Bold = True
        ldsheet.Range("J7:L7").Merge()
        pal_cell = ldsheet.Range("J7")
        try:
            pal_cell.Validation.Delete()
        except Exception:
            pass
        pal_cell.Validation.Add(3, 1, 1, SCHEME_NAMES)   # xlValidateList, xlValidAlertStop, xlBetween
        pal_cell.Validation.IgnoreBlank = True
        pal_cell.Validation.InCellDropdown = True
        pal_cell.Value = "Color (pastel)"
        pal_cell.Interior.Color = bgr("DDE6ED")          # light input fill + box border
        pal_cell.Font.Bold = True
        ldsheet.Range("J7:L7").BorderAround(1, 2)        # xlContinuous, xlThin

        # Fix LOAD STATUS formulas: the Planner template references C5 as a numeric row
        # count, but Loader stores a text token ("5+5"/"7+5"/"7+7") in C5 instead, so
        # formulas like =C5*C6 evaluate to #VALUE!.  Replace with an IF-chain lookup.
        rows = 'IF(C5="5+5",10,IF(C5="7+5",12,IF(C5="7+7",14,10)))'
        ldsheet.Range("L11").Formula = "=" + rows + "*C6"
        ldsheet.Range("L12").Formula = "=F30-" + rows + "*C6"
        ldsheet.Range("L17").Formula = (
            '=IF(E30-COUNTA($B$34:$J$47)>0,'
            '"OVERLOADED ✕ "&(E30-COUNTA($B$34:$J$47))&" unplaced",'
            'IF(AND(COUNTIF($L$34:$L$47,"OK 72")=' + rows + ',' + rows + '>0),'
            '"COMPLETE ✔ all "&' + rows + '&" rows = 72 ft",'
            'IF(COUNTIF($L$34:$L$47,"OVER*")>0,'
            '"ROW OVER 72 ⚠ fix",'
            '"INCOMPLETE ↓ keep filling")))'
        )

        # Compound-car dual selection: the 7-row table tracks its picked row in N1, the
        # 5-row table in N2, so a tally can stay selected in BOTH tables at once.
        tw = wb.Worksheets("Tally")
        for c in ("N1", "N2"):
            tw.Range(c).Value = 0
            tw.Range(c).NumberFormat = ";;;"
        rb = tw.Range("A37:L47")                                            # 5-row table highlight -> N2
        rb.FormatConditions.Delete()
        fcb = rb.FormatConditions.Add(2, pythoncom.Empty, "=ROW()=$N$2")    # 2 = xlExpression
        fcb.Interior.Color = 15652797                                       # BGR of BDD7EE (light blue)
        fcb.StopIfTrue = True
        sm = wb.VBProject.VBComponents(tw.CodeName).CodeModule              # dual-pick SelectionChange
        if sm.CountOfLines:
            sm.DeleteLines(1, sm.CountOfLines)
        sm.AddFromString(TALLY_SHEET_CODE)

        # Send-to-Loader buttons: recommended tally (Tally sheet) + hand-built (Row Patterns).
        add_button(wb.Worksheets("Tally"), "J17", 165, 34,
                   "btnSendToLoader", "Send tally to Loader  >>", "SendTallyToLoader")
        add_button(wb.Worksheets("Row Patterns"), "L1", 205, 34,
                   "btnSendHandBuild", "Send hand-built tally to Loader  >>", "SendHandBuildToLoader")

        # compile check
        xl.EnableEvents = True
        compiled = True
        try:
            xl.Run("ApplyPaletteEverywhere", "Color (pastel)")
        except Exception as e:
            compiled = False
            print("  ! ApplyPaletteEverywhere failed:", e)
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
