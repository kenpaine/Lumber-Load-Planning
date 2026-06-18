Attribute VB_Name = "CenterbeamSolver"
Option Explicit
' === Centerbeam 72-ft Auto-Solver (product + length + grade, per-side) ===
' Supports an asymmetric car: the two sides can have different row heights
' (5+5 = 10 rows, 7+5 = 12 rows mixed, 7+7 = 14 rows). The car layout is in C5.
' Each inventory line (rows 10-29) carries a Side tag in column A ("7-row" / "5-row",
' colour-coded, double-click to flip); in the mixed case the two sides are solved
' INDEPENDENTLY to exact 72-ft rows and written into the grid as Side 1 then Side 2 rows.
Private Patterns As Collection
Private Found As Boolean
Private Chosen As Collection
Private RemCount() As Long
Private TypeLen() As Long
Private NumTypes As Long
Private mQuiet As Boolean        ' suppress dialogs (for automated testing)

Private Const LIFIRST As Long = 10
Private Const LILAST As Long = 29
Private Const GRIDFIRST As Long = 34
Private Const SLOTS As Long = 9
Private Const SIDECOL As Long = 1       ' column A = per-line Side (7 or 5)

' Automation/test entry: run the solve with no popup dialogs.
Public Sub SolveLayoutQuiet()
    mQuiet = True
    On Error GoTo done
    SolveLayout
done:
    mQuiet = False
End Sub

' Parse the car-layout config (C5: "5+5","7+5","7+7"; or legacy 10/14) into sides.
Private Sub CarSides(ws As Worksheet, ByRef aRows As Long, ByRef bRows As Long, _
                     ByRef aTag As Long, ByRef bTag As Long, ByRef mixed As Boolean)
    Dim s As String: s = Trim(CStr(ws.Range("C5").Value))
    Select Case s
        Case "7+5", "5+7": aRows = 7: bRows = 5: aTag = 7: bTag = 5: mixed = True
        Case "7+7":        aRows = 7: bRows = 7: aTag = 7: bTag = 7: mixed = False
        Case "5+5":        aRows = 5: bRows = 5: aTag = 5: bTag = 5: mixed = False
        Case Else
            Dim n As Long: If IsNumeric(s) Then n = CLng(Val(s))
            If n = 14 Then
                aRows = 7: bRows = 7: aTag = 7: bTag = 7
            Else
                aRows = 5: bRows = 5: aTag = 5: bTag = 5
            End If
            mixed = False
    End Select
End Sub

Private Function ParseSide(v As Variant, defSide As Long) As Long
    Dim s As String: s = Trim(CStr(v))
    If InStr(s, "7") > 0 Then
        ParseSide = 7
    ElseIf InStr(s, "5") > 0 Then
        ParseSide = 5
    Else
        ParseSide = defSide
    End If
End Function

Sub SolveLayout()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Planner")
    Dim prods As Variant: prods = Array("2x4", "2x6", "2x8", "2x10", "4x4", "4x6", "6x6")
    Dim grds As Variant:  grds = Array("1", "2", "3", "4", "2P", "MSR")
    Dim lens As Variant:  lens = Array(8, 10, 12, 14, 16, 18, 20)

    Dim aRows As Long, bRows As Long, aTag As Long, bTag As Long, mixed As Boolean
    CarSides ws, aRows, bRows, aTag, bTag, mixed
    Dim defSide As Long: defSide = IIf(Trim(CStr(ws.Range("C5").Value)) = "5+5", 5, 7)

    ' --- single product/grade mode ---
    Dim singleMode As Boolean
    singleMode = (UCase(Trim(CStr(ws.Range("C7").Value))) = "YES")
    Dim defP As String: defP = Trim(CStr(ws.Range("G5").Value))
    Dim defG As String: defG = Trim(CStr(ws.Range("G6").Value))
    If singleMode And defP = "" Then
        If Not mQuiet Then MsgBox "Single product/grade mode is ON, but no default Product is set." & vbCrLf & _
               "Pick a Product (and Grade) in the 'Single Product / Grade' box, " & _
               "or set the toggle to No.", vbExclamation
        Exit Sub
    End If

    ' --- read line items (incl. per-line Side) ---
    Dim nItems As Long: nItems = 0
    Dim itP() As String, itL() As Long, itG() As String, itQ() As Long, itS() As Long
    ReDim itP(1 To LILAST - LIFIRST + 1)
    ReDim itL(1 To LILAST - LIFIRST + 1)
    ReDim itG(1 To LILAST - LIFIRST + 1)
    ReDim itQ(1 To LILAST - LIFIRST + 1)
    ReDim itS(1 To LILAST - LIFIRST + 1)
    Dim r As Long
    For r = LIFIRST To LILAST
        Dim pp As String: pp = Trim(CStr(ws.Cells(r, 2).Value))
        Dim ll As Variant: ll = ws.Cells(r, 3).Value
        Dim gg As String: gg = Trim(CStr(ws.Cells(r, 4).Value))
        Dim qq As Variant: qq = ws.Cells(r, 5).Value
        If singleMode And IsNumeric(ll) And IsNumeric(qq) Then
            If CLng(qq) > 0 Then
                If pp = "" Then pp = defP
                If gg = "" Then gg = defG
            End If
        End If
        If pp <> "" And IsNumeric(ll) And IsNumeric(qq) Then
            If CLng(qq) > 0 Then
                nItems = nItems + 1
                itP(nItems) = pp: itL(nItems) = CLng(ll): itG(nItems) = gg: itQ(nItems) = CLng(qq)
                ' In a symmetric car the Side column is irrelevant -> force the single side.
                If mixed Then
                    itS(nItems) = ParseSide(ws.Cells(r, SIDECOL).Value, defSide)
                Else
                    itS(nItems) = aTag
                End If
                If singleMode Then
                    ws.Cells(r, 2).Value = pp
                    ws.Cells(r, 4).Value = gg
                End If
            End If
        End If
    Next r
    If nItems = 0 Then
        If Not mQuiet Then MsgBox "Enter some line items first."
        Exit Sub
    End If

    ' --- clear the whole grid first ---
    Dim s As Long
    For r = 0 To 13
        For s = 0 To SLOTS - 1: ws.Cells(GRIDFIRST + r, 2 + s).ClearContents: Next s
    Next r

    ' --- solve per side ---
    If mixed Then
        Dim okA As Boolean, okB As Boolean
        okA = SolveSideToGrid(ws, prods, grds, lens, itP, itL, itG, itQ, itS, nItems, aTag, aRows, GRIDFIRST)
        okB = SolveSideToGrid(ws, prods, grds, lens, itP, itL, itG, itQ, itS, nItems, bTag, bRows, GRIDFIRST + aRows)
        If okA And okB Then
            If Not mQuiet Then MsgBox "Solved a mixed car: a " & aRows & "-row side + a " & bRows & "-row side, " & _
                   "every row exactly 72 ft. Products are kept on their assigned side.", vbInformation
        Else
            Dim m As String: m = "Some packs could not be laid out:"
            If Not okA Then m = m & vbCrLf & " - the " & aRows & "-row side's packs don't form exact 72-ft rows."
            If Not okB Then m = m & vbCrLf & " - the " & bRows & "-row side's packs don't form exact 72-ft rows."
            m = m & vbCrLf & "(Each side must total rows x 72 ft from packs that partition into 72-ft rows.)"
            If Not mQuiet Then MsgBox m, vbExclamation
        End If
    Else
        If SolveSideToGrid(ws, prods, grds, lens, itP, itL, itG, itQ, itS, nItems, aTag, aRows + bRows, GRIDFIRST) Then
            If Not mQuiet Then MsgBox "Solved: " & (aRows + bRows) & " rows at 72 ft. Product+grade grouped & column-aligned.", vbInformation
        Else
            If Not mQuiet Then MsgBox "No exact 72-ft solution exists for these lengths and row count.", vbExclamation
        End If
    End If
End Sub

' Solve one side (filter items to sideTag, fill numRows rows to 72 ft, write into the
' grid starting at gridStart). Returns True if every requested row was filled exactly.
Private Function SolveSideToGrid(ws As Worksheet, prods As Variant, grds As Variant, lens As Variant, _
        itP() As String, itL() As Long, itG() As String, itQ() As Long, itS() As Long, _
        nItems As Long, sideTag As Long, numRows As Long, gridStart As Long) As Boolean
    Dim nl As Long: nl = UBound(lens) + 1
    Dim lenTot() As Long: ReDim lenTot(0 To nl - 1)
    Dim i As Long, li As Long, total As Long: total = 0
    For i = 1 To nItems
        If itS(i) = sideTag Then
            For li = 0 To nl - 1
                If lens(li) = itL(i) Then lenTot(li) = lenTot(li) + itQ(i): total = total + itQ(i) * itL(i): Exit For
            Next li
        End If
    Next i
    If total = 0 Then SolveSideToGrid = (numRows <= 0): Exit Function   ' nothing tagged to this side

    NumTypes = 0
    For li = 0 To nl - 1
        If lenTot(li) > 0 Then NumTypes = NumTypes + 1
    Next li
    ReDim TypeLen(0 To NumTypes - 1): ReDim RemCount(0 To NumTypes - 1)
    Dim k As Long: k = 0
    For li = nl - 1 To 0 Step -1
        If lenTot(li) > 0 Then TypeLen(k) = lens(li): RemCount(k) = lenTot(li): k = k + 1
    Next li

    Set Patterns = New Collection
    EnumPatterns 0, 72, ""
    SortPatterns
    Dim fillable As Long
    If numRows < Int(total / 72) Then fillable = numRows Else fillable = Int(total / 72)
    Set Chosen = New Collection
    Found = False
    Solve 0, fillable, 72
    If Not Found Then SolveSideToGrid = False: Exit Function
    If Chosen.Count < numRows Then SolveSideToGrid = False Else SolveSideToGrid = True

    Dim nR As Long: nR = Chosen.Count
    Dim ordered() As String: ReDim ordered(1 To nR)
    For i = 1 To nR: ordered(i) = Chosen(i): Next i
    Dim oa As Long, ob As Long, otmp As String
    For oa = 1 To nR - 1
        For ob = oa + 1 To nR
            If CompareRows(ordered(ob), ordered(oa)) < 0 Then otmp = ordered(oa): ordered(oa) = ordered(ob): ordered(ob) = otmp
        Next ob
    Next oa

    Dim maxQ As Long: maxQ = 0
    For li = 0 To nl - 1
        If lenTot(li) > maxQ Then maxQ = lenTot(li)
    Next li
    If maxQ < 1 Then maxQ = 1
    Dim qP() As String, qG() As String, qpos() As Long
    ReDim qP(0 To nl - 1, 1 To maxQ): ReDim qG(0 To nl - 1, 1 To maxQ): ReDim qpos(0 To nl - 1)
    Dim pix As Long, gix As Long, pos As Long, it As Long
    For li = 0 To nl - 1
        pos = 0
        For pix = 0 To UBound(prods)
            For gix = 0 To UBound(grds)
                For it = 1 To nItems
                    If itS(it) = sideTag And itL(it) = lens(li) And itP(it) = prods(pix) And itG(it) = grds(gix) Then
                        Dim c As Long
                        For c = 1 To itQ(it): pos = pos + 1: qP(li, pos) = prods(pix): qG(li, pos) = grds(gix): Next c
                    End If
                Next it
            Next gix
        Next pix
        qpos(li) = 0
    Next li

    Dim s As Long, r As Long
    For r = 1 To nR
        Dim arr() As String: arr = Split(Trim(ordered(r)), " ")
        For s = 0 To UBound(arr)
            If Len(arr(s)) > 0 Then
                Dim slotLen As Long: slotLen = CLng(arr(s))
                Dim lidx As Long: lidx = -1
                For li = 0 To nl - 1
                    If lens(li) = slotLen Then lidx = li: Exit For
                Next li
                If lidx >= 0 Then
                    If qpos(lidx) < lenTot(lidx) Then
                        qpos(lidx) = qpos(lidx) + 1
                        ws.Cells(gridStart + r - 1, 2 + s).Value = qP(lidx, qpos(lidx)) & "-" & slotLen & "-" & qG(lidx, qpos(lidx))
                    End If
                End If
            End If
        Next s
    Next r
End Function

Private Sub EnumPatterns(ti As Long, remn As Long, cur As String)
    If remn = 0 Then Patterns.Add cur: Exit Sub
    If ti >= NumTypes Then Exit Sub
    Dim t As Long: t = TypeLen(ti)
    Dim maxN As Long
    If RemCount(ti) < Int(remn / t) Then maxN = RemCount(ti) Else maxN = Int(remn / t)
    Dim n As Long, j As Long, addStr As String
    For n = maxN To 0 Step -1
        addStr = cur
        For j = 1 To n: addStr = addStr & t & " ": Next j
        EnumPatterns ti + 1, remn - t * n, addStr
    Next n
End Sub

Private Function PieceCount(patStr As String) As Long
    Dim arr() As String: arr = Split(Trim(patStr), " ")
    Dim s As Long, c As Long
    For s = 0 To UBound(arr)
        If Len(arr(s)) > 0 Then c = c + 1
    Next s
    PieceCount = c
End Function

Private Function DistinctCount(patStr As String) As Long
    Dim arr() As String: arr = Split(Trim(patStr), " ")
    Dim seen As String, s As Long, c As Long: seen = "|"
    For s = 0 To UBound(arr)
        If Len(arr(s)) > 0 Then
            If InStr(seen, "|" & arr(s) & "|") = 0 Then seen = seen & arr(s) & "|": c = c + 1
        End If
    Next s
    DistinctCount = c
End Function

Private Sub SortPatterns()
    Dim n As Long: n = Patterns.Count
    If n < 2 Then Exit Sub
    Dim arr() As String: ReDim arr(1 To n)
    Dim i As Long: For i = 1 To n: arr(i) = Patterns(i): Next i
    Dim a As Long, b As Long, tmp As String
    For a = 1 To n - 1
        For b = a + 1 To n
            If PatBetter(arr(b), arr(a)) Then tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
        Next b
    Next a
    Set Patterns = New Collection
    For i = 1 To n: Patterns.Add arr(i): Next i
End Sub

Private Function PatBetter(x As String, y As String) As Boolean
    Dim dx As Long, dy As Long: dx = DistinctCount(x): dy = DistinctCount(y)
    If dx <> dy Then PatBetter = (dx < dy): Exit Function
    Dim px As Long, py As Long: px = PieceCount(x): py = PieceCount(y)
    If px <> py Then PatBetter = (px < py): Exit Function
    PatBetter = (CompareRows(x, y) < 0)
End Function

Private Function CompareRows(x As String, y As String) As Long
    Dim ax() As String: ax = Split(Trim(x), " ")
    Dim ay() As String: ay = Split(Trim(y), " ")
    Dim nx As Long: nx = UBound(ax): Dim ny As Long: ny = UBound(ay)
    Dim mm As Long: If nx < ny Then mm = nx Else mm = ny
    Dim i As Long, vx As Long, vy As Long
    For i = 0 To mm
        vx = CLng(ax(i)): vy = CLng(ay(i))
        If vx <> vy Then CompareRows = IIf(vx > vy, -1, 1): Exit Function
    Next i
    If nx <> ny Then CompareRows = IIf(nx < ny, -1, 1) Else CompareRows = 0
End Function

Private Sub Solve(rowIdx As Long, fillable As Long, target As Long)
    If Found Then Exit Sub
    If rowIdx = fillable Then Found = True: Exit Sub
    Dim p As Variant, patStr As String
    For Each p In Patterns
        patStr = CStr(p)
        If CanUse(patStr) Then
            UsePat patStr, -1
            Chosen.Add patStr
            Solve rowIdx + 1, fillable, target
            If Found Then Exit Sub
            Chosen.Remove Chosen.Count
            UsePat patStr, 1
        End If
    Next p
End Sub

Private Function CanUse(patStr As String) As Boolean
    Dim need() As Long: ReDim need(0 To NumTypes - 1)
    Dim arr() As String: arr = Split(Trim(patStr), " ")
    Dim s As Long, i As Long, v As Long
    For s = 0 To UBound(arr)
        If Len(arr(s)) > 0 Then
            v = CLng(arr(s))
            For i = 0 To NumTypes - 1
                If TypeLen(i) = v Then need(i) = need(i) + 1: Exit For
            Next i
        End If
    Next s
    For i = 0 To NumTypes - 1
        If need(i) > RemCount(i) Then CanUse = False: Exit Function
    Next i
    CanUse = True
End Function

Private Sub UsePat(patStr As String, delta As Long)
    Dim arr() As String: arr = Split(Trim(patStr), " ")
    Dim s As Long, i As Long, v As Long
    For s = 0 To UBound(arr)
        If Len(arr(s)) > 0 Then
            v = CLng(arr(s))
            For i = 0 To NumTypes - 1
                If TypeLen(i) = v Then RemCount(i) = RemCount(i) + delta: Exit For
            Next i
        End If
    Next s
End Sub

Sub ClearGrid()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Planner")
    Dim r As Long, s As Long
    For r = 0 To 13
        For s = 0 To 8: ws.Cells(34 + r, 2 + s).ClearContents: Next s
    Next r
End Sub

Sub ClearAll()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Planner")
    If MsgBox("Clear the whole layout and all line items?", vbQuestion + vbYesNo, "Start Over") <> vbYes Then Exit Sub
    Dim r As Long, s As Long
    For r = 0 To 13
        For s = 0 To 8: ws.Cells(34 + r, 2 + s).ClearContents: Next s
    Next r
    For r = 10 To 29
        ws.Cells(r, 2).ClearContents: ws.Cells(r, 3).ClearContents
        ws.Cells(r, 4).ClearContents: ws.Cells(r, 5).ClearContents
    Next r
    ApplySingleMode
    MsgBox "Cleared. Enter new line items, then click Solve.", vbInformation
End Sub

' ===== Single product/grade live behavior =====
Public Sub ApplySingleMode()
    On Error GoTo cleanup
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Planner")
    Dim singleMode As Boolean
    singleMode = (UCase(Trim(CStr(ws.Range("C7").Value))) = "YES")
    Dim defP As String: defP = Trim(CStr(ws.Range("G5").Value))
    Dim defG As String: defG = Trim(CStr(ws.Range("G6").Value))
    Dim r As Long
    Application.EnableEvents = False
    If singleMode Then
        For r = 10 To 29
            If RowHasData(ws, r) Then
                ws.Cells(r, 2).Value = defP
                ws.Cells(r, 4).Value = defG
            Else
                ws.Cells(r, 2).ClearContents
                ws.Cells(r, 4).ClearContents
            End If
        Next r
        StyleInputCells ws.Range("B10:B29"), True
        StyleInputCells ws.Range("D10:D29"), True
        DropValidation ws.Range("B10:B29")
        DropValidation ws.Range("D10:D29")
    Else
        StyleInputCells ws.Range("B10:B29"), False
        StyleInputCells ws.Range("D10:D29"), False
        AddListValidation ws.Range("B10:B29"), "2x4,2x6,2x8,2x10,4x4,4x6,6x6"
        AddListValidation ws.Range("D10:D29"), "1,2,3,4,2P,MSR"
    End If
cleanup:
    Application.EnableEvents = True
End Sub

Public Sub OnPlannerChange(ByVal Target As Range)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Planner")
    If Not Application.Intersect(Target, ws.Range("C7")) Is Nothing Then
        ApplySingleMode
        Exit Sub
    End If
    Dim singleMode As Boolean
    singleMode = (UCase(Trim(CStr(ws.Range("C7").Value))) = "YES")
    If Not singleMode Then Exit Sub
    If Application.Intersect(Target, ws.Range("B10:E29,G5:G6")) Is Nothing Then Exit Sub
    On Error GoTo cleanup
    Dim defP As String: defP = Trim(CStr(ws.Range("G5").Value))
    Dim defG As String: defG = Trim(CStr(ws.Range("G6").Value))
    Dim r As Long
    Application.EnableEvents = False
    For r = 10 To 29
        If RowHasData(ws, r) Then
            If CStr(ws.Cells(r, 2).Value) <> defP Then ws.Cells(r, 2).Value = defP
            If CStr(ws.Cells(r, 4).Value) <> defG Then ws.Cells(r, 4).Value = defG
        Else
            ws.Cells(r, 2).ClearContents
            ws.Cells(r, 4).ClearContents
        End If
    Next r
cleanup:
    Application.EnableEvents = True
End Sub

Private Function RowHasData(ws As Worksheet, r As Long) As Boolean
    RowHasData = (Trim(CStr(ws.Cells(r, 3).Value)) <> "") Or (Trim(CStr(ws.Cells(r, 5).Value)) <> "")
End Function

Private Sub StyleInputCells(rng As Range, disabled As Boolean)
    If disabled Then
        rng.Interior.Color = RGB(230, 230, 225)
        rng.Font.Color = RGB(150, 150, 150)
        rng.Font.Italic = True
    Else
        rng.Interior.Color = RGB(255, 255, 153)
        rng.Font.Color = RGB(0, 0, 255)
        rng.Font.Italic = False
    End If
End Sub

Private Sub DropValidation(rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
End Sub

Private Sub AddListValidation(rng As Range, listStr As String)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listStr
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
End Sub
