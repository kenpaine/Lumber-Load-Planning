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
Private mSeq As Long

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
