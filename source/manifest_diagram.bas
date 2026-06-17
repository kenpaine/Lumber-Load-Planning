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

Public Sub DrawManifestDiagram()
    On Error GoTo done
    Dim wsP As Worksheet, mf As Worksheet
    Set wsP = ThisWorkbook.Worksheets("Planner")
    Set mf = ThisWorkbook.Worksheets("Manifest")
    Application.ScreenUpdating = False

    ClearDiagram mf

    Dim nRows As Long: nRows = CLng(Val(wsP.Range("C5").Value))
    If nRows < 2 Then nRows = 10
    Dim perSide As Long: perSide = nRows \ 2
    If perSide < 1 Then perSide = 1

    ' --- geometry: fit inside the cleared band (Manifest row 14 .. row 30) ---
    Dim x0 As Double, y0 As Double, availH As Double
    x0 = mf.Range("A14").Left
    y0 = mf.Range("A14").Top
    availH = mf.Range("A30").Top - y0
    If availH < 140 Then availH = 320

    Dim labelW As Double: labelW = 28      ' left gutter for "R1" labels
    Dim carW As Double: carW = 520         ' 72 ft is drawn this wide
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
        AddLabel mf, x0, y, labelW + carW, sideHdrH, _
                 "SIDE " & side & "   tiers " & ((side - 1) * perSide + 1) & "-" & (side * perSide), True
        y = y + sideHdrH
        DrawRuler mf, x0 + labelW, y, pxft, rulerH
        y = y + rulerH
        Dim t As Long
        For t = 1 To perSide
            Dim tier As Long: tier = (side - 1) * perSide + t
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
                        Dim bc As Long: bc = RGB(120, 120, 120)
                        On Error Resume Next
                        bc = cl.DisplayFormat.Borders(xlEdgeLeft).Color
                        On Error GoTo 0
                        Dim rc As Shape
                        Set rc = mf.Shapes.AddShape(msoShapeRectangle, x0 + labelW + cur * pxft, y, L * pxft, bandH)
                        NameShape rc
                        rc.Fill.ForeColor.RGB = cl.DisplayFormat.Interior.Color
                        rc.Line.ForeColor.RGB = bc
                        rc.Line.Weight = 1.75
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
                        tr.Font.Fill.ForeColor.RGB = cl.DisplayFormat.Font.Color
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

' ===== Dynamic, column-wrapping Pick List (drawn below the diagram) =====
' Reads the loaded line items from the Planner and lays them out as a compact
' table directly under the car diagram. A long list wraps into additional
' columns to the RIGHT (instead of growing downward), so the diagram always
' prints as large as possible and the whole sheet fits one landscape page.
' The print area is set to end at the bottom of whatever the pick list needs.

Public Sub DrawPickList(mf As Worksheet, wsP As Worksheet)
    ClearPick mf

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

    ' --- length -> fill color from the LENGTH legend (row 10, cols B..H) ---
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
                PickShape mf, bx, iy, wNum, rowH, CStr(idx), vbWhite, True, RGB(90, 90, 90), False, 7.5, 2
                PickShape mf, bx + wNum, iy, wProd, rowH, prod(idx), vbWhite, True, RGB(30, 30, 30), True, 7.5, 2
                PickShape mf, bx + wNum + wProd, iy, wLen, rowH, CStr(CLng(lng(idx))) & "'", lc, True, vbWhite, True, 7.5, 2
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
