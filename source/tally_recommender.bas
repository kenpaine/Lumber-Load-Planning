Option Explicit

' ===== Centerbeam 72-ft Car - Lumber Tally Recommender =====
' Tab 'Recommender' : settings, two length palettes, recommended tallies (two tables).
' Tab 'Row Patterns': every 72-ft row pattern (building blocks) + hand-build total.
'
' Car layout (B5): 5+5 (720 ft / 10 rows), 7+5 (864 ft / 12, mixed), 7+7 (1008 / 14).
' Symmetric cars use one combined palette (union of the two columns) and write the
' full-car tallies into table A. A mixed 7+5 car tallies each side INDEPENDENTLY: the
' 7-row side (504 ft) from the column-B palette into table A, and the 5-row side
' (360 ft) from the column-C palette into table B.
'
' Loaded into the workbook by source/build_tally.py via CodeModule.AddFromString.

Private Const REC_SHEET As String = "Recommender"
Private Const PAT_SHEET As String = "Row Patterns"
Private Const CAR_CELL As String = "B5"
Private Const MODE_CELL As String = "B6"
Private Const PAL_FIRST As Integer = 11
Private Const PAL7_COL As Integer = 2          ' column B = 7-row side / symmetric
Private Const PAL5_COL As Integer = 3          ' column C = 5-row side (mixed)
Private Const STATUS_CELL As String = "B19"
Private Const RECA_BANNER As String = "A21"
Private Const RECA_TOP As Integer = 23
Private Const RECA_MAX As Integer = 33
Private Const RECB_BANNER As String = "A35"
Private Const RECB_TOP As Integer = 37
Private Const RECB_MAX As Integer = 47
Private Const PAT_TOP As Integer = 3
Private Const PAT_MAX As Integer = 103

Private Function Lengths() As Variant
    Lengths = Array(8, 10, 12, 14, 16, 18, 20)
End Function

Private Sub CarLayout(ws As Worksheet, ByRef aRows As Integer, ByRef bRows As Integer, ByRef mixed As Boolean)
    Dim t As String: t = CStr(ws.Range(CAR_CELL).Value)
    If InStr(t, "7+5") > 0 Or InStr(t, "864") > 0 Then
        aRows = 7: bRows = 5: mixed = True
    ElseIf InStr(t, "7+7") > 0 Or InStr(t, "1008") > 0 Or InStr(t, "14") > 0 Then
        aRows = 7: bRows = 7: mixed = False
    Else
        aRows = 5: bRows = 5: mixed = False
    End If
End Sub

Private Function ReadMode(ws As Worksheet) As Integer
    Dim modeTxt As String: modeTxt = LCase(CStr(ws.Range(MODE_CELL).Value))
    If InStr(modeTxt, "filler") > 0 Then
        ReadMode = 3
    ElseIf InStr(modeTxt, "each") > 0 Or InStr(modeTxt, "must") > 0 Then
        ReadMode = 2
    Else
        ReadMode = 1
    End If
End Function

Private Sub ReadPalette(ws As Worksheet, col As Integer, ByRef sel() As Boolean)
    Dim i As Integer, pv As Variant
    For i = 0 To 6
        pv = ws.Cells(PAL_FIRST + i, col).Value
        sel(i) = (pv = True) Or (UCase(Trim(CStr(pv & ""))) = "TRUE") Or (UCase(Trim(CStr(pv & ""))) = "YES")
    Next i
End Sub

Private Function AnyTrue(sel() As Boolean) As Boolean
    Dim i As Integer
    For i = 0 To 6
        If sel(i) Then AnyTrue = True: Exit Function
    Next i
End Function

Public Sub RecommendTallies()
    Dim ws As Worksheet, ws2 As Worksheet
    Set ws = ThisWorkbook.Worksheets(REC_SHEET)
    Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    Application.ScreenUpdating = False

    Dim L As Variant: L = Lengths()
    Dim aRows As Integer, bRows As Integer, mixed As Boolean
    CarLayout ws, aRows, bRows, mixed
    Dim mode As Integer: mode = ReadMode(ws)

    Dim sel7(0 To 6) As Boolean, sel5(0 To 6) As Boolean
    ReadPalette ws, PAL7_COL, sel7
    ReadPalette ws, PAL5_COL, sel5

    ClearOutputs ws, ws2

    If mixed Then
        ws.Range(RECA_BANNER).Value = "7-ROW SIDE TALLIES   -   504 ft  (click a length header 8'-20' to sort)"
        ws.Range(RECB_BANNER).Value = "5-ROW SIDE TALLIES   -   360 ft  (click a length header 8'-20' to sort)"
        Dim okA As Boolean, okB As Boolean
        okA = RecommendInto(ws, L, sel7, 7, mode, RECA_TOP, RECA_MAX)
        okB = RecommendInto(ws, L, sel5, 5, mode, RECB_TOP, RECB_MAX)
        ' Row Patterns tab: union of both sides as a reference
        Dim u(0 To 6) As Boolean, i As Integer
        For i = 0 To 6: u(i) = sel7(i) Or sel5(i): Next i
        WriteUnionPatterns ws2, L, u, mode
        If AnyTrue(sel7) And AnyTrue(sel5) Then
            ws.Range(STATUS_CELL).Value = "Mixed 7+5 car: the 7-row side (504 ft) and 5-row side (360 ft) are tallied independently below. Check lengths for each side in its own palette column."
        Else
            ws.Range(STATUS_CELL).Value = "Mixed 7+5 car: check lengths for BOTH sides (7-row side = column B palette, 5-row side = column C palette)."
        End If
    Else
        ws.Range(RECA_BANNER).Value = "RECOMMENDED FULL-CAR TALLIES   (click a length header 8'-20' to sort)"
        ws.Range(RECB_BANNER).Value = "(switch Car layout to 7+5 for a mixed / asymmetric car - each side tallied separately)"
        Dim sel(0 To 6) As Boolean
        For i = 0 To 6: sel(i) = sel7(i) Or sel5(i): Next i
        If Not AnyTrue(sel) Then
            ws.Range(STATUS_CELL).Value = "Check at least one length box, then click Recommend Tallies."
            Application.ScreenUpdating = True: Exit Sub
        End If
        Dim okAll As Boolean
        okAll = RecommendInto(ws, L, sel, aRows + bRows, mode, RECA_TOP, RECA_MAX)
        WriteUnionPatterns ws2, L, sel, mode
        Dim carFt As Long: carFt = (aRows + bRows) * 72
        If okAll Then
            ws.Range(STATUS_CELL).Value = "Recommended for a " & carFt & "-ft car (" & (aRows + bRows) & " rows). Row patterns are on the 'Row Patterns' tab."
        Else
            ws.Range(STATUS_CELL).Value = "No single- or two-pattern car covers every selected length. Try 'Palette' mode, or hand-build on the 'Row Patterns' tab."
        End If
    End If

    Application.ScreenUpdating = True
End Sub

' Enumerate this selection's 72-ft row patterns and write recommended tallies for
' `rows` rows into the table region [top..maxr]. Returns True if any tally written.
Private Function RecommendInto(ws As Worksheet, L As Variant, sel() As Boolean, ByVal rows As Long, _
                               ByVal mode As Integer, ByVal top As Long, ByVal maxr As Long) As Boolean
    Dim i As Integer
    If Not AnyTrue(sel) Then Exit Function

    Dim allow(0 To 6) As Boolean
    For i = 0 To 6: allow(i) = IIf(mode = 3, True, sel(i)): Next i

    Dim patterns() As Variant: ReDim patterns(1 To 250)
    Dim nPat As Long: nPat = 0
    Dim cur(0 To 6) As Integer
    EnumRows L, allow, 0, 72, cur, patterns, nPat
    If nPat = 0 Then Exit Function

    Dim r As Long: r = top
    Dim p As Long, a As Long, b As Long
    For p = 1 To nPat
        If r > maxr Then Exit For
        Dim cnt As Variant: cnt = patterns(p)
        If CarCovers(cnt, cnt, sel, mode) Then
            Dim comb As Variant: ReDim comb(0 To 6)
            For i = 0 To 6: comb(i) = cnt(i) * rows: Next i
            WriteRecCounts ws, r, top, "All " & rows & " rows:  " & RowLabel(L, cnt), L, comb, rows
            r = r + 1
        End If
    Next p
    If mode >= 2 Then
        For a = 1 To nPat
            If r > maxr Then Exit For
            For b = a + 1 To nPat
                If r > maxr Then Exit For
                Dim ca As Variant: ca = patterns(a)
                Dim cb As Variant: cb = patterns(b)
                If CarCovers(ca, cb, sel, mode) Then
                    Dim na As Integer, nb As Integer
                    na = rows \ 2: If na < 1 Then na = 1
                    nb = rows - na
                    Dim comb2 As Variant: ReDim comb2(0 To 6)
                    For i = 0 To 6: comb2(i) = ca(i) * na + cb(i) * nb: Next i
                    WriteRecCounts ws, r, top, na & " rows (" & RowLabel(L, ca) & ")  +  " & nb & " rows (" & RowLabel(L, cb) & ")", L, comb2, rows
                    r = r + 1
                End If
            Next b
        Next a
    End If
    RecommendInto = (r > top)
End Function

Private Sub EnumRows(L As Variant, allow() As Boolean, idx As Integer, remaining As Integer, _
                     ByRef cur() As Integer, ByRef patterns() As Variant, ByRef nPat As Long)
    If remaining = 0 Then
        nPat = nPat + 1
        If nPat > UBound(patterns) Then ReDim Preserve patterns(1 To nPat + 100)
        Dim snap As Variant: ReDim snap(0 To 6)
        Dim j As Integer
        For j = 0 To 6: snap(j) = cur(j): Next j
        patterns(nPat) = snap
        Exit Sub
    End If
    If idx > 6 Then Exit Sub
    If Not allow(idx) Then
        EnumRows L, allow, idx + 1, remaining, cur, patterns, nPat
        Exit Sub
    End If
    Dim maxk As Integer, k As Integer
    maxk = remaining \ CInt(L(idx))
    For k = 0 To maxk
        cur(idx) = k
        EnumRows L, allow, idx + 1, remaining - k * CInt(L(idx)), cur, patterns, nPat
    Next k
    cur(idx) = 0
End Sub

Private Function RowLabel(L As Variant, cnt As Variant) As String
    Dim s As String, i As Integer
    For i = 0 To 6
        If cnt(i) > 0 Then
            If Len(s) > 0 Then s = s & ", "
            s = s & cnt(i) & "x" & L(i)
        End If
    Next i
    RowLabel = s
End Function

Private Function ColL(ByVal c As Integer) As String
    Dim s As String, m As Integer
    s = ""
    Do While c > 0
        m = (c - 1) Mod 26
        s = Chr(65 + m) & s
        c = (c - m - 1) \ 26
    Loop
    ColL = s
End Function

Private Sub WriteUnionPatterns(ws2 As Worksheet, L As Variant, sel() As Boolean, mode As Integer)
    Dim i As Integer
    Dim allow(0 To 6) As Boolean
    For i = 0 To 6: allow(i) = IIf(mode = 3, True, sel(i)): Next i
    Dim patterns() As Variant: ReDim patterns(1 To 250)
    Dim nPat As Long: nPat = 0
    Dim cur(0 To 6) As Integer
    EnumRows L, allow, 0, 72, cur, patterns, nPat
    WritePatterns ws2, L, patterns, nPat
End Sub

Private Sub WritePatterns(ws2 As Worksheet, L As Variant, patterns() As Variant, nPat As Long)
    Dim r As Long, p As Long, i As Integer, shown As Long
    r = PAT_TOP: shown = 0
    For p = 1 To nPat
        If r > PAT_MAX Then Exit For
        Dim cnt As Variant: cnt = patterns(p)
        ws2.Cells(r, 1).Value = p
        ws2.Cells(r, 2).Value = RowLabel(L, cnt)
        For i = 0 To 6
            If cnt(i) > 0 Then ws2.Cells(r, 3 + i).Value = cnt(i)
        Next i
        r = r + 1: shown = shown + 1
    Next p
    If shown = 0 Then Exit Sub

    Dim lastPat As Long: lastPat = PAT_TOP + shown - 1
    ws2.Range("J" & PAT_TOP & ":J" & lastPat).Interior.Color = RGB(255, 247, 214)

    Dim br As Long: br = lastPat + 1
    ws2.Cells(br, 2).Value = "YOUR HAND-BUILT TALLY (set 'Rows to use' above):"
    For i = 0 To 6
        ws2.Cells(br, 3 + i).Formula = _
            "=SUMPRODUCT(" & ColL(3 + i) & PAT_TOP & ":" & ColL(3 + i) & lastPat & ",$J$" & PAT_TOP & ":$J$" & lastPat & ")"
    Next i
    ws2.Cells(br, 10).Formula = "=SUM($J$" & PAT_TOP & ":$J$" & lastPat & ")&"" rows"""
    With ws2.Range("A" & br & ":J" & br)
        .Font.Bold = True
        .Interior.Color = RGB(221, 230, 237)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
End Sub

Public Sub SortPatternsBy(ByVal sortCol As Long)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    Dim lastPat As Long: lastPat = PAT_TOP - 1
    Dim r As Long: r = PAT_TOP
    Do While r <= PAT_MAX
        If Len(CStr(ws2.Cells(r, 1).Value)) = 0 Then Exit Do
        lastPat = r: r = r + 1
    Loop
    If lastPat <= PAT_TOP Then Exit Sub
    On Error GoTo cleanup
    Application.EnableEvents = False
    With ws2.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws2.Range(ColL(CInt(sortCol)) & PAT_TOP & ":" & ColL(CInt(sortCol)) & lastPat), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws2.Range("B" & PAT_TOP & ":J" & lastPat)
        .Header = xlNo
        .Apply
    End With
    Dim k As Long
    For k = PAT_TOP To lastPat
        ws2.Cells(k, 1).Value = k - PAT_TOP + 1
    Next k
cleanup:
    Application.EnableEvents = True
End Sub

' Sort one recommendation table (top..max) by a length column, most first.
Public Sub SortRecsByRange(ByVal sortCol As Long, ByVal top As Long, ByVal maxr As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(REC_SHEET)
    Dim lastRow As Long: lastRow = top - 1
    Dim r As Long: r = top
    Do While r <= maxr
        If Len(CStr(ws.Cells(r, 1).Value)) = 0 Then Exit Do
        lastRow = r: r = r + 1
    Loop
    If lastRow <= top Then Exit Sub
    On Error GoTo cleanup
    Application.EnableEvents = False
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ColL(CInt(sortCol)) & top & ":" & ColL(CInt(sortCol)) & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.Range("B" & top & ":L" & lastRow)
        .Header = xlNo
        .Apply
    End With
    Dim k As Long
    For k = top To lastRow
        ws.Cells(k, 1).Value = k - top + 1
    Next k
cleanup:
    Application.EnableEvents = True
End Sub

' Back-compat shim (older sheet event) -> sort table A.
Public Sub SortRecsBy(ByVal sortCol As Long)
    SortRecsByRange sortCol, RECA_TOP, RECA_MAX
End Sub

Private Function CarCovers(ca As Variant, cb As Variant, sel() As Boolean, mode As Integer) As Boolean
    Dim i As Integer
    If mode = 1 Then CarCovers = True: Exit Function
    For i = 0 To 6
        If sel(i) And (ca(i) = 0) And (cb(i) = 0) Then CarCovers = False: Exit Function
    Next i
    CarCovers = True
End Function

Private Sub WriteRecCounts(ws As Worksheet, ByVal r As Long, ByVal top As Long, label As String, L As Variant, comb As Variant, ByVal rows As Long)
    ws.Cells(r, 1).Value = r - top + 1
    ws.Cells(r, 2).Value = label
    Dim i As Integer, pcs As Long, ft As Long
    pcs = 0: ft = 0
    For i = 0 To 6
        If comb(i) > 0 Then ws.Cells(r, 3 + i).Value = comb(i)
        pcs = pcs + comb(i)
        ft = ft + comb(i) * CLng(L(i))
    Next i
    ws.Cells(r, 10).Value = pcs
    ws.Cells(r, 11).Value = ft
    ws.Cells(r, 12).Value = IIf(ft = rows * 72, "OK", "check")
End Sub

Private Sub ClearOutputs(ws As Worksheet, ws2 As Worksheet)
    ws.Range("A" & RECA_TOP & ":L" & RECA_MAX).ClearContents
    ws.Range("A" & RECB_TOP & ":L" & RECB_MAX).ClearContents
    Dim rng As Range
    Set rng = ws2.Range("A" & PAT_TOP & ":L" & (PAT_MAX + 20))
    rng.ClearContents
    rng.Interior.ColorIndex = xlNone
    rng.Font.Bold = False
    rng.Borders.LineStyle = xlNone
End Sub

Public Sub ClearOutputsButton()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(REC_SHEET)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    ClearOutputs ws, ws2
    ws.Range(STATUS_CELL).Value = "Cleared. Choose options and click Recommend Tallies."
End Sub
