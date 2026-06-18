Option Explicit

' ===== Manifest car diagram — to scale, per side =====
' Draws the solved load on the Manifest tab as proportional rectangles
' (rectangle width = pack length in feet), split into the car's two sides:
'   Side 1 = tiers 1 .. N\2,  Side 2 = tiers N\2+1 .. N.
' Each pack's colors are read from the matching Planner grid cell's
' DisplayFormat (length = fill, product = border, grade = text), so the
' diagram always matches the grid's palette. Redrawn by the Manifest sheet's
' Worksheet_Activate event and by the Redraw button.
'
' Loaded into the workbook by source/build_manifest_diagram.py via AddFromString.

Private Const GF As Long = 34           ' Planner grid first data row (R1)
Private Const SLOTS As Long = 9         ' P1..P9 = Planner columns B..J
Private Const TARGETFT As Double = 72
Private Const DIAG As String = "diag_"  ' every diagram shape gets this name prefix
Private Const PICK As String = "pick_"  ' every pick-list shape gets this name prefix
Private Const PLI_FIRST As Long = 10    ' Planner line-item first row
Private Const PLI_LAST As Long = 29     ' Planner line-item last row
Private mSeq As Long
Private mPick As Long

' Parse the Planner car-layout config (C5: "5+5","7+5","7+7"; or legacy 10/14)
' into the two sides' row counts. A mixed 7+5 car => Side 1 = 7, Side 2 = 5.
Private Sub ParseCarSides(ByVal s As String, ByRef aRows As Long, ByRef bRows As Long)
    s = Trim(s)
    Select Case s
        Case "7+5", "5+7": aRows = 7: bRows = 5
        Case "7+7":        aRows = 7: bRows = 7
        Case "5+5":        aRows = 5: bRows = 5
        Case Else
            Dim n As Long: If IsNumeric(s) Then n = CLng(Val(s))
            If n = 14 Then
                aRows = 7: bRows = 7
            Else
                aRows = 5: bRows = 5
            End If
    End Select
End Sub

Public Sub DrawManifestDiagram()
    On Error GoTo done
    Dim wsP As Worksheet, mf As Worksheet
    Set wsP = ThisWorkbook.Worksheets("Planner")
    Set mf = ThisWorkbook.Worksheets("Manifest")
    Application.ScreenUpdating = False

    ClearDiagram mf

    Dim pal As Long: pal = CurrentPalette()
    ApplyLegend mf, pal

    Dim aRows As Long, bRows As Long
    ParseCarSides CStr(wsP.Range("C5").Value), aRows, bRows
    Dim nRows As Long: nRows = aRows + bRows

    ' --- geometry: fit inside the cleared band (Manifest row 14 .. row 30) ---
    Dim x0 As Double, y0 As Double, availH As Double
    x0 = mf.Range("A14").Left
    y0 = mf.Range("A14").Top
    availH = mf.Range("A30").Top - y0
    If availH < 140 Then availH = 320

    Dim labelW As Double: labelW = 28      ' left gutter for "R1" labels
    Dim rightPad As Double: rightPad = 14  ' room for the "72'" ruler label overhang
    ' fill the full print width (columns A..L) instead of a fixed 520 pt
    Dim carW As Double: carW = mf.Range("A14:L14").Width - labelW - rightPad
    If carW < 300 Then carW = 520          ' fallback if columns are unexpectedly narrow
    Dim pxft As Double: pxft = carW / TARGETFT
    Dim sideHdrH As Double: sideHdrH = 14
    Dim rulerH As Double: rulerH = 11
    Dim sideGap As Double: sideGap = 8

    Dim chrome As Double: chrome = 2 * (sideHdrH + rulerH) + sideGap
    Dim bandTotal As Double: bandTotal = availH - chrome
    If bandTotal < nRows * 9 Then bandTotal = nRows * 9
    Dim slot As Double: slot = bandTotal / nRows
    If slot > 28 Then slot = 28
    Dim bandGap As Double: bandGap = IIf(slot > 16, 2, 1)
    Dim bandH As Double: bandH = slot - bandGap

    Dim y As Double: y = y0
    Dim side As Long
    For side = 1 To 2
        Dim sideRows As Long, startTier As Long
        If side = 1 Then
            sideRows = aRows: startTier = 1
        Else
            sideRows = bRows: startTier = aRows + 1
        End If
        AddLabel mf, x0, y, labelW + carW, sideHdrH, _
                 "SIDE " & side & "   " & sideRows & " rows  (tiers " & startTier & "-" & (startTier + sideRows - 1) & ")", True
        y = y + sideHdrH
        DrawRuler mf, x0 + labelW, y, pxft, rulerH
        y = y + rulerH
        Dim t As Long
        For t = 1 To sideRows
            Dim tier As Long: tier = startTier + t - 1
            Dim pr As Long: pr = GF + tier - 1
            AddLabel mf, x0, y, labelW, bandH, "R" & tier, False
            Dim bg As Shape
            Set bg = mf.Shapes.AddShape(msoShapeRectangle, x0 + labelW, y, carW, bandH)
            NameShape bg
            bg.Fill.ForeColor.RGB = RGB(244, 244, 242)
            bg.Line.ForeColor.RGB = RGB(205, 205, 205)
            bg.Line.Weight = 0.5
            bg.Shadow.Visible = msoFalse

            Dim cur As Double: cur = 0
            Dim anyPack As Boolean: anyPack = False
            Dim s As Long
            For s = 0 To SLOTS - 1
                Dim v As String: v = CStr(wsP.Cells(pr, 2 + s).Value & "")
                If Len(v) > 0 Then
                    Dim L As Double: L = ParseLen(v)
                    If L > 0 Then
                        anyPack = True
                        Dim cl As Range: Set cl = wsP.Cells(pr, 2 + s)
                        Dim fcol As Long, bcol As Long, tcol As Long
                        PackColors cl, v, pal, fcol, bcol, tcol
                        Dim rc As Shape
                        Set rc = mf.Shapes.AddShape(msoShapeRectangle, x0 + labelW + cur * pxft, y, L * pxft, bandH)
                        NameShape rc
                        rc.Fill.ForeColor.RGB = fcol
                        rc.Line.ForeColor.RGB = bcol
                        rc.Line.Weight = IIf(pal = 0, 1.75, 1#)
                        rc.Shadow.Visible = msoFalse
                        Dim tr As Object: Set tr = rc.TextFrame2.TextRange
                        If L * pxft > 50 And bandH >= 12 Then
                            tr.Text = v
                        ElseIf L * pxft > 17 Then
                            tr.Text = CStr(CLng(L)) & "'"
                        End If
                        On Error Resume Next
                        tr.Font.Size = 7
                        tr.Font.Bold = msoTrue
                        tr.Font.Fill.ForeColor.RGB = tcol
                        rc.TextFrame2.VerticalAnchor = msoAnchorMiddle
                        tr.ParagraphFormat.Alignment = msoAlignCenter
                        rc.TextFrame2.WordWrap = msoFalse
                        rc.TextFrame2.AutoSize = msoAutoSizeNone
                        rc.TextFrame2.MarginLeft = 0
                        rc.TextFrame2.MarginRight = 0
                        On Error GoTo 0
                        cur = cur + L
                    End If
                End If
            Next s

            If Not anyPack Then
                Dim em As Object: Set em = bg.TextFrame2.TextRange
                em.Text = "— empty —"
                em.Font.Size = 7
                em.Font.Italic = msoTrue
                em.Font.Fill.ForeColor.RGB = RGB(170, 170, 170)
                bg.TextFrame2.VerticalAnchor = msoAnchorMiddle
                em.ParagraphFormat.Alignment = msoAlignCenter
            End If
            y = y + bandH + bandGap
        Next t
        y = y + sideGap
    Next side

    DrawPickList mf, wsP

done:
    Application.ScreenUpdating = True
End Sub

Private Sub ClearDiagram(mf As Worksheet)
    Dim i As Long
    For i = mf.Shapes.Count To 1 Step -1
        If Left(mf.Shapes(i).Name, Len(DIAG)) = DIAG Then mf.Shapes(i).Delete
    Next i
End Sub

Private Sub NameShape(s As Shape)
    mSeq = mSeq + 1
    s.Name = DIAG & mSeq
End Sub

Private Sub AddLabel(mf As Worksheet, x As Double, y As Double, w As Double, h As Double, txt As String, isHdr As Boolean)
    Dim tb As Shape
    Set tb = mf.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    NameShape tb
    tb.Fill.Visible = msoFalse
    tb.Line.Visible = msoFalse
    Dim tr As Object: Set tr = tb.TextFrame2.TextRange
    tr.Text = txt
    tr.Font.Bold = msoTrue
    If isHdr Then
        tr.Font.Size = 9
        tr.Font.Fill.ForeColor.RGB = RGB(46, 94, 32)
        tr.ParagraphFormat.Alignment = msoAlignLeft
    Else
        tr.Font.Size = 8
        tr.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
        tr.ParagraphFormat.Alignment = msoAlignCenter
    End If
    tb.TextFrame2.VerticalAnchor = msoAnchorMiddle
    tb.TextFrame2.MarginLeft = 1
    tb.TextFrame2.MarginRight = 1
    tb.TextFrame2.MarginTop = 0
    tb.TextFrame2.MarginBottom = 0
    tb.TextFrame2.WordWrap = msoFalse
End Sub

Private Sub DrawRuler(mf As Worksheet, x As Double, y As Double, pxft As Double, h As Double)
    Dim ft As Long
    For ft = 0 To 72 Step 12
        Dim tb As Shape
        Set tb = mf.Shapes.AddTextbox(msoTextOrientationHorizontal, x + ft * pxft - 13, y, 26, h)
        NameShape tb
        tb.Fill.Visible = msoFalse
        tb.Line.Visible = msoFalse
        Dim tr As Object: Set tr = tb.TextFrame2.TextRange
        tr.Text = ft & "'"
        tr.Font.Size = 6
        tr.Font.Fill.ForeColor.RGB = RGB(150, 150, 150)
        tr.ParagraphFormat.Alignment = msoAlignCenter
        tb.TextFrame2.VerticalAnchor = msoAnchorMiddle
        tb.TextFrame2.WordWrap = msoFalse
    Next ft
End Sub

Private Function ParseLen(ByVal v As String) As Double
    Dim parts() As String
    parts = Split(v, "-")
    If UBound(parts) >= 1 Then ParseLen = Val(parts(1)) Else ParseLen = 0
End Function

Public Sub RedrawManifestButton()
    DrawManifestDiagram
End Sub

' ===== Color palettes =====
' The user picks a palette in cell N2 (off the print area). Three options:
'   0 = Color (current pastel "Sorbet" scheme — reads the live Planner grid colors)
'   1 = High contrast (saturated, distinct fills with black borders/auto text)
'   2 = Black & white (grayscale fills for clean printing)
' Length drives the fill (the primary cue); product/grade ride along in each
' pack's label. The length LEGEND (row 10) is recolored to match the palette.

Public Function CurrentPalette() As Long
    On Error Resume Next
    Dim v As String
    v = UCase(Trim(CStr(ThisWorkbook.Worksheets("Manifest").Range("N2").Value & "")))
    If InStr(v, "B & W") > 0 Or InStr(v, "BLACK") > 0 Or v = "BW" Then
        CurrentPalette = 2
    ElseIf InStr(v, "CONTRAST") > 0 Or InStr(v, "HIGH") > 0 Then
        CurrentPalette = 1
    Else
        CurrentPalette = 0
    End If
End Function

' Fill/border/text for one pack, given its grid cell and "prod-len-grd" code.
Private Sub PackColors(cl As Range, v As String, pal As Long, _
                       ByRef fillCol As Long, ByRef borderCol As Long, ByRef textCol As Long)
    If pal = 0 Then
        ' current scheme: live colors from the Planner grid cell
        fillCol = cl.DisplayFormat.Interior.Color
        borderCol = RGB(120, 120, 120)
        On Error Resume Next
        borderCol = cl.DisplayFormat.Borders(xlEdgeLeft).Color
        textCol = cl.DisplayFormat.Font.Color
        On Error GoTo 0
    Else
        Dim L As Long: L = CLng(ParseLen(v))
        fillCol = IIf(pal = 1, HCLen(L), BWLen(L))
        borderCol = RGB(0, 0, 0)
        textCol = BestText(fillCol)
    End If
End Sub

' Pastel "Sorbet" length fills (palette 0) — measured from the original legend.
Private Function Pal0Len(ByVal L As Long) As Long
    Select Case L
        Case 8:  Pal0Len = 15389883
        Case 10: Pal0Len = 13231554
        Case 12: Pal0Len = 11724795
        Case 14: Pal0Len = 15650277
        Case 16: Pal0Len = 12371194
        Case 18: Pal0Len = 14740159
        Case 20: Pal0Len = 11462620
        Case Else: Pal0Len = RGB(240, 240, 240)
    End Select
End Function

' High-contrast qualitative length fills (palette 1).
Private Function HCLen(ByVal L As Long) As Long
    Select Case L
        Case 8:  HCLen = RGB(31, 119, 180)    ' blue
        Case 10: HCLen = RGB(44, 160, 44)     ' green
        Case 12: HCLen = RGB(255, 127, 14)    ' orange
        Case 14: HCLen = RGB(148, 103, 189)   ' purple
        Case 16: HCLen = RGB(214, 39, 40)     ' red
        Case 18: HCLen = RGB(23, 190, 207)    ' cyan
        Case 20: HCLen = RGB(140, 86, 75)     ' brown
        Case Else: HCLen = RGB(90, 90, 90)
    End Select
End Function

' Grayscale length fills, light->dark, all readable with black text (palette 2).
Private Function BWLen(ByVal L As Long) As Long
    Select Case L
        Case 8:  BWLen = RGB(255, 255, 255)
        Case 10: BWLen = RGB(238, 238, 238)
        Case 12: BWLen = RGB(221, 221, 221)
        Case 14: BWLen = RGB(204, 204, 204)
        Case 16: BWLen = RGB(187, 187, 187)
        Case 18: BWLen = RGB(170, 170, 170)
        Case 20: BWLen = RGB(153, 153, 153)
        Case Else: BWLen = RGB(255, 255, 255)
    End Select
End Function

' Black or white text, whichever reads better on the given fill.
Private Function BestText(ByVal c As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&
    g = (c \ &H100) And &HFF&
    b = (c \ &H10000) And &HFF&
    If (0.299 * r + 0.587 * g + 0.114 * b) > 150 Then
        BestText = RGB(0, 0, 0)
    Else
        BestText = RGB(255, 255, 255)
    End If
End Function

' Recolor the LENGTH legend (row 10, cols B..H) to match the active palette so
' the diagram and pick list always agree with the legend.
Private Sub ApplyLegend(mf As Worksheet, pal As Long)
    Dim lv As Variant: lv = Array(8, 10, 12, 14, 16, 18, 20)
    Dim i As Long, col As Long
    For i = 0 To 6
        Select Case pal
            Case 0: col = Pal0Len(CLng(lv(i)))
            Case 1: col = HCLen(CLng(lv(i)))
            Case Else: col = BWLen(CLng(lv(i)))
        End Select
        With mf.Cells(10, 2 + i)
            .Interior.Color = col
            .Font.Color = BestText(col)
        End With
    Next i
End Sub

' ===== Dynamic, column-wrapping Pick List (drawn below the diagram) =====
' Reads the loaded line items from the Planner and lays them out as a compact
' table directly under the car diagram. A long list wraps into additional
' columns to the RIGHT (instead of growing downward), so the diagram always
' prints as large as possible and the whole sheet fits one landscape page.
' The print area is set to end at the bottom of whatever the pick list needs.

Public Sub DrawPickList(mf As Worksheet, wsP As Worksheet)
    ClearPick mf
    Dim pal As Long: pal = CurrentPalette()

    ' --- gather loaded line items (product set, packs placed > 0) ---
    Dim cap As Long: cap = PLI_LAST - PLI_FIRST + 1
    Dim prod() As String, lng() As Double, grd() As String, pcs() As Double
    ReDim prod(1 To cap): ReDim lng(1 To cap): ReDim grd(1 To cap): ReDim pcs(1 To cap)
    Dim n As Long: n = 0
    Dim r As Long
    For r = PLI_FIRST To PLI_LAST
        Dim pv As String: pv = Trim(CStr(wsP.Cells(r, 2).Value & ""))
        Dim gv As Variant: gv = wsP.Cells(r, 7).Value          ' placed packs (col G)
        If pv <> "" And IsNumeric(gv) Then
            If CLng(gv) > 0 Then
                n = n + 1
                prod(n) = pv
                lng(n) = Val(wsP.Cells(r, 3).Value)
                grd(n) = Trim(CStr(wsP.Cells(r, 4).Value & ""))
                pcs(n) = CDbl(gv)
            End If
        End If
    Next r

    ' --- length -> fill color from the LENGTH legend (row 10, cols B..H), which
    '     ApplyLegend has already recolored for the active palette ---
    Dim lenVal As Variant: lenVal = Array(8, 10, 12, 14, 16, 18, 20)
    Dim lenClr(0 To 6) As Long
    Dim i As Long
    For i = 0 To 6
        lenClr(i) = mf.Cells(10, 2 + i).DisplayFormat.Interior.Color
    Next i

    ' --- geometry ---
    Dim x0 As Double: x0 = mf.Range("A30").Left
    Dim y0 As Double: y0 = mf.Range("A30").Top
    Dim rowH As Double: rowH = 13
    Dim bannerH As Double: bannerH = 15
    Dim wNum As Double: wNum = 22
    Dim wProd As Double: wProd = 46
    Dim wLen As Double: wLen = 34
    Dim wGrd As Double: wGrd = 26
    Dim wPcs As Double: wPcs = 40
    Dim colW As Double: colW = wNum + wProd + wLen + wGrd + wPcs   ' 168
    Dim colGap As Double: colGap = 10

    Dim totalPcs As Double: totalPcs = 0
    For i = 1 To n: totalPcs = totalPcs + pcs(i): Next i

    ' choose number of columns: wrap into up to 4 columns (~6 rows each) so a long
    ' list grows to the right instead of pushing the diagram down.
    Dim numCols As Long
    If n <= 1 Then
        numCols = 1
    Else
        numCols = (n + 5) \ 6
        If numCols < 1 Then numCols = 1
        If numCols > 4 Then numCols = 4
    End If
    Dim rowsPerCol As Long
    If n <= 0 Then rowsPerCol = 0 Else rowsPerCol = (n + numCols - 1) \ numCols

    Dim bannerW As Double: bannerW = numCols * colW + (numCols - 1) * colGap
    PickShape mf, x0, y0, bannerW, bannerH, _
        "PICK LIST  -  packs to pull" & _
        IIf(n > 0, "   (" & n & " items  /  " & CLng(totalPcs) & " packs)", "   (no packs placed)"), _
        RGB(34, 40, 32), True, vbWhite, True, 9, 1

    Dim hdrY As Double: hdrY = y0 + bannerH
    Dim itemY0 As Double: itemY0 = hdrY + rowH

    Dim c As Long
    For c = 0 To numCols - 1
        Dim bx As Double: bx = x0 + c * (colW + colGap)
        ' per-column header
        PickShape mf, bx, hdrY, wNum, rowH, "#", RGB(43, 49, 38), True, vbWhite, True, 7.5, 2
        PickShape mf, bx + wNum, hdrY, wProd, rowH, "Product", RGB(43, 49, 38), True, vbWhite, True, 7.5, 2
        PickShape mf, bx + wNum + wProd, hdrY, wLen, rowH, "Len", RGB(43, 49, 38), True, vbWhite, True, 7.5, 2
        PickShape mf, bx + wNum + wProd + wLen, hdrY, wGrd, rowH, "Grd", RGB(43, 49, 38), True, vbWhite, True, 7.5, 2
        PickShape mf, bx + wNum + wProd + wLen + wGrd, hdrY, wPcs, rowH, "Packs", RGB(43, 49, 38), True, vbWhite, True, 7.5, 2
        Dim k As Long
        For k = 1 To rowsPerCol
            Dim idx As Long: idx = c * rowsPerCol + k
            If idx <= n Then
                Dim iy As Double: iy = itemY0 + (k - 1) * rowH
                Dim lc As Long: lc = RGB(120, 120, 120)
                Dim li As Long
                For li = 0 To 6
                    If lenVal(li) = lng(idx) Then lc = lenClr(li): Exit For
                Next li
                Dim lt As Long: lt = IIf(pal = 0, vbWhite, BestText(lc))
                PickShape mf, bx, iy, wNum, rowH, CStr(idx), vbWhite, True, RGB(90, 90, 90), False, 7.5, 2
                PickShape mf, bx + wNum, iy, wProd, rowH, prod(idx), vbWhite, True, RGB(30, 30, 30), True, 7.5, 2
                PickShape mf, bx + wNum + wProd, iy, wLen, rowH, CStr(CLng(lng(idx))) & "'", lc, True, lt, True, 7.5, 2
                PickShape mf, bx + wNum + wProd + wLen, iy, wGrd, rowH, grd(idx), vbWhite, True, RGB(60, 60, 60), False, 7.5, 2
                PickShape mf, bx + wNum + wProd + wLen + wGrd, iy, wPcs, rowH, CStr(CLng(pcs(idx))), RGB(244, 244, 240), True, RGB(20, 20, 20), True, 7.5, 2
            End If
        Next k
    Next c

    Dim pickBottom As Double
    If n = 0 Then
        pickBottom = hdrY + 4
    Else
        pickBottom = itemY0 + rowsPerCol * rowH + 4
    End If

    SetManifestPrintArea mf, pickBottom
End Sub

Private Sub ClearPick(mf As Worksheet)
    Dim i As Long
    For i = mf.Shapes.Count To 1 Step -1
        If Left(mf.Shapes(i).Name, Len(PICK)) = PICK Then mf.Shapes(i).Delete
    Next i
End Sub

Private Sub PickShape(mf As Worksheet, x As Double, y As Double, w As Double, h As Double, _
                      txt As String, fillColor As Long, hasFill As Boolean, _
                      fontColor As Long, bold As Boolean, fsize As Single, align As Long)
    Dim s As Shape
    Set s = mf.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    mPick = mPick + 1
    s.Name = PICK & mPick
    If hasFill Then
        s.Fill.Solid
        s.Fill.ForeColor.RGB = fillColor
    Else
        s.Fill.Visible = msoFalse
    End If
    s.Line.ForeColor.RGB = RGB(208, 208, 208)
    s.Line.Weight = 0.5
    s.Shadow.Visible = msoFalse
    Dim tr As Object: Set tr = s.TextFrame2.TextRange
    tr.Text = txt
    On Error Resume Next
    tr.Font.Size = fsize
    tr.Font.Bold = IIf(bold, msoTrue, msoFalse)
    tr.Font.Fill.ForeColor.RGB = fontColor
    s.TextFrame2.VerticalAnchor = msoAnchorMiddle
    tr.ParagraphFormat.Alignment = align        ' 1 = left, 2 = center
    s.TextFrame2.WordWrap = msoFalse
    s.TextFrame2.AutoSize = msoAutoSizeNone
    s.TextFrame2.MarginLeft = 2
    s.TextFrame2.MarginRight = 2
    s.TextFrame2.MarginTop = 0
    s.TextFrame2.MarginBottom = 0
    On Error GoTo 0
End Sub

Private Sub SetManifestPrintArea(mf As Worksheet, bottomY As Double)
    ' last Excel row whose top is above bottomY -> the print area ends there
    Dim r As Long: r = 30
    Do While r < 250
        If mf.Rows(r + 1).Top >= bottomY Then Exit Do
        r = r + 1
    Loop
    With mf.PageSetup
        .PrintArea = "A1:L" & r
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.1)
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub
