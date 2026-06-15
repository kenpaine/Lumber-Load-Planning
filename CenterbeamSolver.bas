Attribute VB_Name = "CenterbeamSolver"
Option Explicit
' === Centerbeam 72-ft Auto-Solver (product + length + grade) ===
Private Patterns As Collection
Private Found As Boolean
Private Chosen As Collection
Private RemCount() As Long
Private TypeLen() As Long
Private NumTypes As Long

Sub SolveLayout()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Planner")
    Dim prods As Variant: prods = Array("2x4", "2x6", "2x8", "2x10", "4x4", "4x6", "6x6")
    Dim grds As Variant:  grds = Array("1", "2", "3", "4", "2P", "MSR")
    Dim lens As Variant:  lens = Array(8, 10, 12, 14, 16, 18, 20)
    Const LIFIRST As Long = 10
    Const LILAST As Long = 29
    Const GRIDFIRST As Long = 34
    Const SLOTS As Long = 9
    Dim numRows As Long: numRows = ws.Range("C5").Value
    Dim target As Long: target = ws.Range("C6").Value

    ' --- single product/grade mode ---
    ' When C7 = "Yes", the whole car uses one product + grade. The user only needs to
    ' enter Length + Packs on each line; blank Product/Grade cells are filled with the
    ' defaults set in G5 (product) and G6 (grade). Any value typed on a line is kept.
    Dim singleMode As Boolean
    singleMode = (UCase(Trim(CStr(ws.Range("C7").Value))) = "YES")
    Dim defP As String: defP = Trim(CStr(ws.Range("G5").Value))
    Dim defG As String: defG = Trim(CStr(ws.Range("G6").Value))
    If singleMode And defP = "" Then
        MsgBox "Single product/grade mode is ON, but no default Product is set." & vbCrLf & _
               "Pick a Product (and Grade) in the 'Single Product / Grade' box, " & _
               "or set the toggle to No.", vbExclamation
        Exit Sub
    End If

    ' --- read line items ---
    Dim nItems As Long: nItems = 0
    Dim itP() As String, itL() As Long, itG() As String, itQ() As Long
    ReDim itP(1 To LILAST - LIFIRST + 1)
    ReDim itL(1 To LILAST - LIFIRST + 1)
    ReDim itG(1 To LILAST - LIFIRST + 1)
    ReDim itQ(1 To LILAST - LIFIRST + 1)
    Dim r As Long
    For r = LIFIRST To LILAST
        Dim pp As String: pp = Trim(CStr(ws.Cells(r, 2).Value))
        Dim ll As Variant: ll = ws.Cells(r, 3).Value
        Dim gg As String: gg = Trim(CStr(ws.Cells(r, 4).Value))
        Dim qq As Variant: qq = ws.Cells(r, 5).Value
        ' In single mode, fill blank Product/Grade from the defaults (per-line entries win).
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
                ' Write resolved product/grade back so the tally + status formulas compute.
                If singleMode Then
                    ws.Cells(r, 2).Value = pp
                    ws.Cells(r, 4).Value = gg
                End If
            End If
        End If
    Next r
    If nItems = 0 Then MsgBox "Enter some line items first.": Exit Sub

    ' --- length totals ---
    Dim nl As Long: nl = UBound(lens) + 1
    Dim lenTot() As Long: ReDim lenTot(0 To nl - 1)
    Dim i As Long, li As Long, total As Long
    For i = 1 To nItems
        For li = 0 To nl - 1
            If lens(li) = itL(i) Then lenTot(li) = lenTot(li) + itQ(i): total = total + itQ(i) * itL(i): Exit For
        Next li
    Next i

    ' --- length types (descending) ---
    NumTypes = 0
    For li = 0 To nl - 1
        If lenTot(li) > 0 Then NumTypes = NumTypes + 1
    Next li
    ReDim TypeLen(0 To NumTypes - 1): ReDim RemCount(0 To NumTypes - 1)
    Dim k As Long: k = 0
    For li = nl - 1 To 0 Step -1
        If lenTot(li) > 0 Then TypeLen(k) = lens(li): RemCount(k) = lenTot(li): k = k + 1
    Next li

    ' --- solve length layout ---
    Set Patterns = New Collection
    EnumPatterns 0, target, ""
    SortPatterns
    Dim fillable As Long
    If numRows < Int(total / target) Then fillable = numRows Else fillable = Int(total / target)
    Set Chosen = New Collection
    Found = False
    Solve 0, fillable, target

    Dim s As Long
    For r = 0 To 13
        For s = 0 To SLOTS - 1: ws.Cells(GRIDFIRST + r, 2 + s).ClearContents: Next s
    Next r
    If Not Found Then MsgBox "No exact 72-ft solution exists for these lengths and row count.", vbExclamation: Exit Sub

    ' --- order rows for column stacking ---
    Dim nR As Long: nR = Chosen.Count
    Dim ordered() As String: ReDim ordered(1 To nR)
    For r = 1 To nR: ordered(r) = Chosen(r): Next r
    Dim oa As Long, ob As Long, otmp As String
    For oa = 1 To nR - 1
        For ob = oa + 1 To nR
            If CompareRows(ordered(ob), ordered(oa)) < 0 Then otmp = ordered(oa): ordered(oa) = ordered(ob): ordered(ob) = otmp
        Next ob
    Next oa

    ' --- build product+grade queue per length (grouped product order then grade order) ---
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
                    If itL(it) = lens(li) And itP(it) = prods(pix) And itG(it) = grds(gix) Then
                        Dim c As Long
                        For c = 1 To itQ(it): pos = pos + 1: qP(li, pos) = prods(pix): qG(li, pos) = grds(gix): Next c
                    End If
                Next it
            Next gix
        Next pix
        qpos(li) = 0
    Next li

    ' --- write rows ---
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
                        ws.Cells(GRIDFIRST + r - 1, 2 + s).Value = qP(lidx, qpos(lidx)) & "-" & slotLen & "-" & qG(lidx, qpos(lidx))
                    End If
                End If
            End If
        Next s
    Next r
    MsgBox "Solved: " & nR & " rows at 72 ft. Product+grade grouped & column-aligned.", vbInformation
End Sub

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
' When C7 = "Yes", the Product (col B) and Grade (col D) columns of the inventory are
' disabled (greyed, no dropdown) and auto-filled from the Single Product/Grade box
' (G5 = product, G6 = grade) for every line that has a Length or Packs entered.

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
    ' Toggle changed -> reconfigure (lock/unlock + auto-fill)
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

