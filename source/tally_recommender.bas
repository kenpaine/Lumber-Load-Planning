Option Explicit

' ===== Centerbeam 72-ft Car - Lumber Tally Recommender =====
' Tab 'Recommender' : settings, length palette(s), recommended tallies.
' Tab 'Row Patterns': hand-build grid(s) - every 72-ft row pattern + a target-aware total.
'
' Car layout (B5): 5+5 (720 ft / 10 rows), 7+5 (864 ft / 12, mixed), 7+7 (1008 / 14).
' The UI is REACTIVE to B5:
'   * Symmetric (5+5 / 7+7): ONE palette column is shown (the 5-row checkboxes and the
'     5-row-side table are hidden), and the Row Patterns tab shows ONE hand-build grid
'     that aims for the proper number of rows (10 or 14).
'   * Mixed 7+5: BOTH palette columns are shown and each side is tallied INDEPENDENTLY -
'     the 7-row side (504 ft) from the column-B palette into table A, and the 5-row side
'     (360 ft) from the column-C palette into table B - and the Row Patterns tab shows
'     TWO hand-build grids (aim for 7 rows / aim for 5 rows).
'
' Loaded into the workbook by source/build_tally.py via CodeModule.AddFromString.

Private Const REC_SHEET As String = "Recommender"
Private Const PAT_SHEET As String = "Row Patterns"
Private Const CAR_CELL As String = "B5"
Private Const MODE_CELL As String = "B6"
Private Const PAL_CELL As String = "B8"        ' colour-scheme picker
Private Const PAL_FIRST As Integer = 11
Private Const PAL7_COL As Integer = 2          ' column B = 7-row side / symmetric
Private Const PAL5_COL As Integer = 3          ' column C = 5-row side (mixed)
Private Const STATUS_CELL As String = "B19"
Private Const RECA_BANNER As String = "A21"
Private Const RECA_TOP As Integer = 23
Private Const RECA_MAX As Integer = 33
Private Const RECB_BANNER As String = "A35"
Private Const RECB_ROW1 As Integer = 35        ' first row of the 5-row-side table block
Private Const RECB_TOP As Integer = 37
Private Const RECB_MAX As Integer = 47
Private Const PAT_TOP As Integer = 3           ' first hand-build grid starts at row 1 (banner)
Private Const PAT_REGION As Integer = 240      ' Row Patterns area cleared/managed by VBA
Private Const GRID_CAP As Integer = 101        ' max patterns shown per hand-build grid

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

' Show/hide the reactive bits to match the car layout.
Private Sub ApplyLayoutUI(ws As Worksheet, ByVal mixed As Boolean)
    Dim i As Integer, L As Variant: L = Lengths()
    On Error Resume Next
    For i = 0 To 6
        ws.CheckBoxes("cb5_" & L(i)).Visible = mixed
        ws.CheckBoxes("cb7_" & L(i)).Caption = IIf(mixed, "7-row", "")
    Next i
    On Error GoTo 0
    If mixed Then
        ws.Range("B10").Value = "Mixed 7+5:  LEFT box = 7-row side  /  RIGHT box = 5-row side  (each side tallied separately)"
    Else
        ws.Range("B10").Value = "Check each length to load in this car"
    End If
    ' Only show the 5-row-side table for a mixed car.
    ws.Rows(RECB_ROW1 & ":" & RECB_MAX).Hidden = Not mixed
End Sub

Public Sub RecommendTallies()
    Dim ws As Worksheet, ws2 As Worksheet
    Set ws = ThisWorkbook.Worksheets(REC_SHEET)
    Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    Application.ScreenUpdating = False
    ws.Range("N1").Value = 0: ws2.Range("N1").Value = 0   ' reset active-row highlight

    Dim L As Variant: L = Lengths()
    Dim aRows As Integer, bRows As Integer, mixed As Boolean
    CarLayout ws, aRows, bRows, mixed
    ApplyLayoutUI ws, mixed
    Dim mode As Integer: mode = ReadMode(ws)

    Dim sel7(0 To 6) As Boolean, sel5(0 To 6) As Boolean
    ReadPalette ws, PAL7_COL, sel7
    ReadPalette ws, PAL5_COL, sel5

    ClearRecTables ws
    ClearPatterns ws2

    If mixed Then
        ws.Range(RECA_BANNER).Value = "7-ROW SIDE TALLIES   -   504 ft  (click a length header 8'-20' to sort)"
        ws.Range(RECB_BANNER).Value = "5-ROW SIDE TALLIES   -   360 ft  (click a length header 8'-20' to sort)"
        RecommendInto ws, L, sel7, 7, mode, RECA_TOP, RECA_MAX
        RecommendInto ws, L, sel5, 5, mode, RECB_TOP, RECB_MAX
        ' Two independent hand-build grids on the Row Patterns tab.
        Dim g1 As Long
        g1 = WriteGrid(ws2, 1, "7-ROW SIDE  -  hand-build 72-ft rows  (aim for 7 rows)   [click a length header 8'-20' to sort]", L, sel7, mode, 7)
        WriteGrid ws2, g1 + 2, "5-ROW SIDE  -  hand-build 72-ft rows  (aim for 5 rows)   [click a length header 8'-20' to sort]", L, sel5, mode, 5
        If AnyTrue(sel7) And AnyTrue(sel5) Then
            ws.Range(STATUS_CELL).Value = "Mixed 7+5 car: the 7-row side (504 ft) and 5-row side (360 ft) are tallied independently. Hand-build each side on the 'Row Patterns' tab."
        Else
            ws.Range(STATUS_CELL).Value = "Mixed 7+5 car: check lengths for BOTH sides (LEFT box = 7-row side, RIGHT box = 5-row side)."
        End If
    Else
        ws.Range(RECA_BANNER).Value = "RECOMMENDED FULL-CAR TALLIES   (click a length header 8'-20' to sort)"
        Dim i As Integer
        If Not AnyTrue(sel7) Then
            ws.Range(STATUS_CELL).Value = "Check at least one length box, then click Recommend Tallies."
            Application.ScreenUpdating = True: Exit Sub
        End If
        Dim okAll As Boolean
        okAll = RecommendInto(ws, L, sel7, aRows + bRows, mode, RECA_TOP, RECA_MAX)
        WriteGrid ws2, 1, "VALID 72-FT ROW PATTERNS  -  hand-build a full car  (aim for " & (aRows + bRows) & " rows)   [click a length header 8'-20' to sort]", L, sel7, mode, aRows + bRows
        Dim carFt As Long: carFt = (aRows + bRows) * 72
        If okAll Then
            ws.Range(STATUS_CELL).Value = "Recommended for a " & carFt & "-ft car (" & (aRows + bRows) & " rows). Hand-build a custom blend on the 'Row Patterns' tab."
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

' Color the 8'..20' header cells (cols 3..9) on row `hr` to match the length palette.
Private Sub ColorLenHeaders(ws2 As Worksheet, ByVal hr As Long)
    Dim pal As Long: pal = TallyPalette()
    Dim Lv As Variant: Lv = Lengths()
    Dim i As Integer
    For i = 0 To 6
        ColorOneLen ws2.Cells(hr, 3 + i), pal, CLng(Lv(i))
    Next i
End Sub

' ===== Colour schemes (12 selectable palettes; B8 picker, match the apps) =====
' Length drives the fill. The picker recolors the length chips (A11:A17), both
' recommendation tables' length headers, and every hand-build grid header.

Private Function SchemeNames() As Variant
    SchemeNames = Array("Color (pastel)", "Vivid", "Material", "Tableau", _
        "Earth / lumberyard", "Jewel tones", "Rainbow (warm to cool)", _
        "Viridis (colour-safe)", "Sunset (warm)", "Neon", "High contrast", "B & W (print)")
End Function

Private Function TallyPalette() As Long
    On Error Resume Next
    Dim v As String: v = Trim(CStr(ThisWorkbook.Worksheets(REC_SHEET).Range(PAL_CELL).Value & ""))
    Dim nm As Variant: nm = SchemeNames()
    Dim i As Long
    For i = 0 To UBound(nm)
        If StrComp(v, CStr(nm(i)), vbTextCompare) = 0 Then TallyPalette = i: Exit Function
    Next i
    TallyPalette = 0
End Function

Private Function PaletteHex(ByVal pal As Long) As Variant
    Select Case pal
        Case 1:  PaletteHex = Array("E53935", "FB8C00", "FDD835", "43A047", "00ACC1", "1E88E5", "8E24AA")
        Case 2:  PaletteHex = Array("EF5350", "FFA726", "FFEE58", "66BB6A", "26C6DA", "42A5F5", "AB47BC")
        Case 3:  PaletteHex = Array("4E79A7", "F28E2B", "E15759", "76B7B2", "59A14F", "EDC948", "B07AA1")
        Case 4:  PaletteHex = Array("8C6239", "C9A66B", "7D8C4F", "B7410E", "2E5E4E", "D4A017", "5C3A21")
        Case 5:  PaletteHex = Array("B71C1C", "E65100", "F9A825", "1B5E20", "00695C", "1A237E", "4A148C")
        Case 6:  PaletteHex = Array("D32F2F", "F57C00", "FBC02D", "689F38", "0097A7", "1976D2", "7B1FA2")
        Case 7:  PaletteHex = Array("440154", "443983", "31688E", "21918C", "35B779", "90D743", "FDE725")
        Case 8:  PaletteHex = Array("5C1A33", "9D2B4A", "C44536", "E8590C", "F0A202", "F4C430", "F7E07A")
        Case 9:  PaletteHex = Array("FF1744", "FF9100", "FFEA00", "00E676", "00E5FF", "2979FF", "D500F9")
        Case 10: PaletteHex = Array("1F77B4", "2CA02C", "FF7F0E", "9467BD", "D62728", "17BECF", "8C564B")
        Case 11: PaletteHex = Array("FFFFFF", "EEEEEE", "DDDDDD", "CCCCCC", "BBBBBB", "AAAAAA", "999999")
        Case Else: PaletteHex = Array("BBD4EA", "C2E5C9", "FBE7B2", "E5CDEE", "FAC4BC", "BFEAE0", "DCE7AE")
    End Select
End Function

Private Function LenIdx(ByVal L As Long) As Long
    Select Case L
        Case 8: LenIdx = 0
        Case 10: LenIdx = 1
        Case 12: LenIdx = 2
        Case 14: LenIdx = 3
        Case 16: LenIdx = 4
        Case 18: LenIdx = 5
        Case 20: LenIdx = 6
        Case Else: LenIdx = -1
    End Select
End Function

Private Function HexBGR(ByVal h As String) As Long
    HexBGR = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

Private Function LenColor(ByVal pal As Long, ByVal L As Long) As Long
    Dim idx As Long: idx = LenIdx(L)
    If idx < 0 Then LenColor = RGB(240, 240, 240): Exit Function
    Dim arr As Variant: arr = PaletteHex(pal)
    LenColor = HexBGR(CStr(arr(idx)))
End Function

Private Function BestText(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&
    g = (c \ &H100) And &HFF&
    b = (c \ &H10000) And &HFF&
    If (0.299 * r + 0.587 * g + 0.114 * b) > 150 Then BestText = RGB(0, 0, 0) Else BestText = RGB(255, 255, 255)
End Function

Private Sub ColorOneLen(cell As Range, ByVal pal As Long, ByVal L As Long)
    Dim col As Long: col = LenColor(pal, L)
    cell.Interior.Color = col
    cell.Font.Color = BestText(col)
End Sub

' Recolor every length-coloured cell to the scheme picked in B8.
Public Sub ApplyTallyPalette()
    Dim ws As Worksheet, ws2 As Worksheet
    Set ws = ThisWorkbook.Worksheets(REC_SHEET)
    Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    Dim pal As Long: pal = TallyPalette()
    Dim L As Variant: L = Lengths()
    Dim i As Long, r As Long
    Application.ScreenUpdating = False
    For i = 0 To 6
        ColorOneLen ws.Cells(PAL_FIRST + i, 1), pal, CLng(L(i))      ' length chips A11:A17
        ColorOneLen ws.Cells(22, 3 + i), pal, CLng(L(i))            ' table A header
        ColorOneLen ws.Cells(36, 3 + i), pal, CLng(L(i))            ' table B header
    Next i
    For r = 1 To PAT_REGION                                          ' every hand-build grid header
        If CStr(ws2.Cells(r, 2).Value) = "Row = 72 ft" Then
            For i = 0 To 6
                ColorOneLen ws2.Cells(r, 3 + i), pal, CLng(L(i))
            Next i
        End If
    Next r
    Application.ScreenUpdating = True
End Sub

' Write one hand-build grid (banner + header + 72-ft row patterns + a target-aware
' total) starting at `startRow`. Returns the row index of the total line.
Private Function WriteGrid(ws2 As Worksheet, ByVal startRow As Long, title As String, _
                           L As Variant, sel() As Boolean, ByVal mode As Integer, ByVal target As Long) As Long
    Dim i As Integer
    ' banner
    ws2.Range("A" & startRow & ":J" & startRow).Merge
    With ws2.Range("A" & startRow)
        .Value = title
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' header row
    Dim hr As Long: hr = startRow + 1
    Dim hdr As Variant
    hdr = Array("#", "Row = 72 ft", "8'", "10'", "12'", "14'", "16'", "18'", "20'", "Rows to use")
    For i = 0 To 9
        With ws2.Cells(hr, i + 1)
            .Value = hdr(i)
            .Font.Bold = True
            .Interior.Color = RGB(221, 230, 237)
        End With
    Next i
    ColorLenHeaders ws2, hr

    ' enumerate this side's 72-ft row patterns
    Dim allow(0 To 6) As Boolean
    For i = 0 To 6: allow(i) = IIf(mode = 3, True, sel(i)): Next i
    Dim patterns() As Variant: ReDim patterns(1 To 250)
    Dim nPat As Long: nPat = 0
    Dim cur(0 To 6) As Integer
    EnumRows L, allow, 0, 72, cur, patterns, nPat

    Dim firstPat As Long: firstPat = hr + 1
    If Not AnyTrue(sel) Then
        ws2.Cells(firstPat, 2).Value = "(check this side's lengths in the palette on the Recommender tab)"
        WriteGrid = firstPat
        Exit Function
    End If
    If nPat = 0 Then
        ws2.Cells(firstPat, 2).Value = "(no 72-ft row can be built from this side's lengths - add another length)"
        WriteGrid = firstPat
        Exit Function
    End If

    Dim r As Long: r = firstPat
    Dim shown As Long: shown = 0
    Dim p As Long
    For p = 1 To nPat
        If shown >= GRID_CAP Then Exit For
        Dim cnt As Variant: cnt = patterns(p)
        ws2.Cells(r, 1).Value = shown + 1
        ws2.Cells(r, 2).Value = RowLabel(L, cnt)
        For i = 0 To 6
            If cnt(i) > 0 Then ws2.Cells(r, 3 + i).Value = cnt(i)
        Next i
        r = r + 1: shown = shown + 1
    Next p
    Dim lastPat As Long: lastPat = firstPat + shown - 1

    ' highlight the editable "Rows to use" column
    ws2.Range("J" & firstPat & ":J" & lastPat).Interior.Color = RGB(255, 247, 214)

    ' target-aware total line (e.g. "5 / 7 rows" + " OK" when it matches)
    Dim br As Long: br = lastPat + 1
    ws2.Cells(br, 2).Value = "YOUR HAND-BUILT TALLY (set 'Rows to use' above):"
    For i = 0 To 6
        ws2.Cells(br, 3 + i).Formula = _
            "=SUMPRODUCT(" & ColL(3 + i) & firstPat & ":" & ColL(3 + i) & lastPat & ",$J$" & firstPat & ":$J$" & lastPat & ")"
    Next i
    ws2.Cells(br, 10).Formula = _
        "=SUM($J$" & firstPat & ":$J$" & lastPat & ")&"" / " & target & " rows""&IF(SUM($J$" & firstPat & ":$J$" & lastPat & ")=" & target & ","" OK"","""")"
    With ws2.Range("A" & br & ":J" & br)
        .Font.Bold = True
        .Interior.Color = RGB(221, 230, 237)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    WriteGrid = br
End Function

' Click a length-column header (8'..20') of any hand-build grid to sort that grid.
Public Sub SortPatternsBy(ByVal sortCol As Long, ByVal headerRow As Long)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    Dim firstPat As Long: firstPat = headerRow + 1
    Dim lastPat As Long: lastPat = firstPat - 1
    Dim r As Long: r = firstPat
    Do While r <= PAT_REGION
        If Len(CStr(ws2.Cells(r, 1).Value)) = 0 Then Exit Do
        If Not IsNumeric(ws2.Cells(r, 1).Value) Then Exit Do
        lastPat = r: r = r + 1
    Loop
    If lastPat <= firstPat Then Exit Sub
    On Error GoTo cleanup
    Application.EnableEvents = False
    With ws2.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws2.Range(ColL(CInt(sortCol)) & firstPat & ":" & ColL(CInt(sortCol)) & lastPat), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws2.Range("B" & firstPat & ":J" & lastPat)
        .Header = xlNo
        .Apply
    End With
    Dim k As Long
    For k = firstPat To lastPat
        ws2.Cells(k, 1).Value = k - firstPat + 1
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

Private Sub ClearRecTables(ws As Worksheet)
    ws.Range("A" & RECA_TOP & ":L" & RECA_MAX).ClearContents
    ws.Range("A" & RECB_TOP & ":L" & RECB_MAX).ClearContents
End Sub

Private Sub ClearPatterns(ws2 As Worksheet)
    Dim rng As Range
    Set rng = ws2.Range("A1:L" & PAT_REGION)
    rng.UnMerge
    rng.ClearContents
    rng.Interior.ColorIndex = xlNone
    rng.Font.Bold = False
    rng.Font.Color = RGB(51, 65, 79)
    rng.Borders.LineStyle = xlNone
End Sub

Public Sub ClearOutputsButton()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(REC_SHEET)
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Worksheets(PAT_SHEET)
    ClearRecTables ws
    ClearPatterns ws2
    ws.Range("N1").Value = 0: ws2.Range("N1").Value = 0   ' reset active-row highlight
    ws.Range(STATUS_CELL).Value = "Cleared. Choose options and click Recommend Tallies."
End Sub
