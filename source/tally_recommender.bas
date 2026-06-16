Option Explicit

' ===== Centerbeam 72-ft Car - Lumber Tally Recommender =====
' Tab 'Recommender' : settings, palette, recommended full-car tallies.
' Tab 'Row Patterns': every 72-ft row pattern (building blocks) + hand-build total.
'
' Loaded into the workbook by source/build_tally.py via CodeModule.AddFromString.

Private Const REC_SHEET As String = "Recommender"
Private Const PAT_SHEET As String = "Row Patterns"
Private Const CAR_CELL As String = "B5"
Private Const MODE_CELL As String = "B6"
Private Const PAL_FIRST As Integer = 11
Private Const STATUS_CELL As String = "B19"
Private Const REC_TOP As Integer = 23
Private Const REC_MAX As Integer = 37
Private Const PAT_HDR As Integer = 2
Private Const PAT_TOP As Integer = 3
Private Const PAT_MAX As Integer = 103

Private Function Lengths() As Variant
    Lengths = Array(8, 10, 12, 14, 16, 18, 20)
End Function

Public Sub RecommendTallies()
    Dim ws As Worksheet, ws2 As Worksheet
    Set ws = ThisWorkbook.Worksheets(REC_SHEET)
    Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    Application.ScreenUpdating = False

    Dim L As Variant: L = Lengths()
    Dim i As Integer

    Dim carRows As Integer, carTxt As String
    carTxt = CStr(ws.Range(CAR_CELL).Value)
    If InStr(carTxt, "1008") > 0 Or InStr(carTxt, "14") > 0 Then carRows = 14 Else carRows = 10

    Dim modeTxt As String, mode As Integer
    modeTxt = LCase(CStr(ws.Range(MODE_CELL).Value))
    If InStr(modeTxt, "filler") > 0 Then
        mode = 3
    ElseIf InStr(modeTxt, "each") > 0 Or InStr(modeTxt, "must") > 0 Then
        mode = 2
    Else
        mode = 1
    End If

    ' Palette column B holds each checkbox's linked TRUE/FALSE (hidden via ';;;').
    ' Stay tolerant of legacy "Yes"/"True" text in case of an older sheet.
    Dim sel(0 To 6) As Boolean, anySel As Boolean
    Dim pv As Variant
    For i = 0 To 6
        pv = ws.Cells(PAL_FIRST + i, 2).Value
        sel(i) = (pv = True) Or (UCase(Trim(CStr(pv & ""))) = "TRUE") Or (UCase(Trim(CStr(pv & ""))) = "YES")
        If sel(i) Then anySel = True
    Next i

    ClearOutputs ws, ws2

    If Not anySel Then
        ws.Range(STATUS_CELL).Value = "Check at least one length box, then click Recommend Tallies."
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim allow(0 To 6) As Boolean
    For i = 0 To 6
        If mode = 3 Then allow(i) = True Else allow(i) = sel(i)
    Next i

    Dim patterns() As Variant: ReDim patterns(1 To 250)
    Dim nPat As Long: nPat = 0
    Dim cur(0 To 6) As Integer
    EnumRows L, allow, 0, 72, cur, patterns, nPat

    If nPat = 0 Then
        ws.Range(STATUS_CELL).Value = "No 72-ft row can be built from those lengths. Add another length and try again."
        Application.ScreenUpdating = True
        Exit Sub
    End If

    WritePatterns ws2, L, patterns, nPat
    BuildRecommendations ws, L, sel, patterns, nPat, carRows, mode

    Application.ScreenUpdating = True
End Sub

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

    Dim lastPat As Long: lastPat = PAT_TOP + shown - 1
    ' Rows-to-use input column is now J (the Pcs/Ft columns were removed)
    ws2.Range("J" & PAT_TOP & ":J" & lastPat).Interior.Color = RGB(255, 247, 214)

    ' Builder total placed directly below the last pattern row (reactive to row count)
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
    ' Click a length header (8'..20') to sort the pattern grid by that length, most first.
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

Private Sub BuildRecommendations(ws As Worksheet, L As Variant, sel() As Boolean, _
                                 patterns() As Variant, nPat As Long, carRows As Integer, mode As Integer)
    Dim r As Long: r = REC_TOP
    Dim p As Long, a As Long, b As Long, i As Integer

    For p = 1 To nPat
        If r > REC_MAX Then Exit For
        Dim cnt As Variant: cnt = patterns(p)
        If CarCovers(cnt, cnt, sel, mode) Then
            Dim comb As Variant: ReDim comb(0 To 6)
            For i = 0 To 6: comb(i) = cnt(i) * carRows: Next i
            WriteRecCounts ws, r, "All " & carRows & " rows:  " & RowLabel(L, cnt), L, comb, carRows
            r = r + 1
        End If
    Next p

    If mode >= 2 Then
        For a = 1 To nPat
            If r > REC_MAX Then Exit For
            For b = a + 1 To nPat
                If r > REC_MAX Then Exit For
                Dim ca As Variant: ca = patterns(a)
                Dim cb As Variant: cb = patterns(b)
                If CarCovers(ca, cb, sel, mode) Then
                    Dim na As Integer, nb As Integer
                    na = carRows \ 2: If na < 1 Then na = 1
                    nb = carRows - na
                    Dim comb2 As Variant: ReDim comb2(0 To 6)
                    For i = 0 To 6: comb2(i) = ca(i) * na + cb(i) * nb: Next i
                    WriteRecCounts ws, r, na & " rows (" & RowLabel(L, ca) & ")  +  " & nb & " rows (" & RowLabel(L, cb) & ")", L, comb2, carRows
                    r = r + 1
                End If
            Next b
        Next a
    End If

    Dim made As Long: made = r - REC_TOP
    Dim carFt As Long: carFt = carRows * 72
    If made = 0 Then
        ws.Range(STATUS_CELL).Value = "No single- or two-pattern car covers every selected length. Add a length, try 'Palette' mode, or hand-build on the 'Row Patterns' tab.  (" & nPat & " row patterns found.)"
    Else
        ws.Range(STATUS_CELL).Value = made & " tally(ies) recommended for a " & carFt & "-ft car (" & carRows & " rows).  " & nPat & " valid 72-ft row patterns are on the 'Row Patterns' tab."
    End If
End Sub

Private Function CarCovers(ca As Variant, cb As Variant, sel() As Boolean, mode As Integer) As Boolean
    Dim i As Integer
    If mode = 1 Then CarCovers = True: Exit Function
    For i = 0 To 6
        If sel(i) And (ca(i) = 0) And (cb(i) = 0) Then CarCovers = False: Exit Function
    Next i
    CarCovers = True
End Function

Private Sub WriteRecCounts(ws As Worksheet, r As Long, label As String, L As Variant, comb As Variant, carRows As Integer)
    ws.Cells(r, 1).Value = r - REC_TOP + 1
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
    ws.Cells(r, 12).Value = IIf(ft = carRows * 72, "OK", "check")
End Sub

Private Sub ClearOutputs(ws As Worksheet, ws2 As Worksheet)
    ws.Range("A" & REC_TOP & ":L" & REC_MAX).ClearContents
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
