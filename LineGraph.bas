Option Explicit
' ===== style, taken from resize.m (550x400 px figure) =====
Private Const FONT_NAME     As String = "Arial"
Private Const MAT_FIG_W_PT  As Double = 412.5    ' 550 px @ 96 dpi
Private Const MAT_ASPECT    As Double = 1.375    ' 550/400
Private Const MAT_FS_TICKS  As Double = 13.5     ' resize.m: axes FontSize
Private Const MAT_FS_LABEL  As Double = 16       ' resize.m: x/ylabel FontSize
Private Const MAT_LINEW     As Double = 0.5      ' MATLAB default line width
Private Const MAT_PLOTW     As Double = 2        ' xgrapher: plot LineWidth
Private Const MAT_MS        As Double = 18       ' xgrapher: ms (MarkerSize, MATLAB pts)
Private Const CH_WIDTH_IN   As Double = 5        ' <-- set your output size here
Private Const PA_LEFT       As Double = 0.1505   ' resize.m axes Position
Private Const PA_TOP        As Double = 0.071
Private Const PA_WIDTH      As Double = 0.8
Private Const PA_HEIGHT     As Double = 0.75
Private Const LG_RIGHT      As Double = 0.938    ' 'northeast'
Private Const LG_TOP        As Double = 0.086
Private Const LG_LINESP     As Double = 1.45
Private Const TICK_OFFSET   As Long = 20
Private Const DO_LABELS     As Boolean = True    ' xgrapher: dolabels
Private Const LABEL_FS      As Double = 10       ' point-label font size (pt, not scaled)
' settings cells on the Buttons sheet of this workbook
Private Const SETTINGS_SHEET As String = "Buttons"
Private Const LINE_CELL      As String = "K6"     ' 1 = connecting lines, 0 = markers only
Private Const SCHEME_CELL    As String = "K7"     ' standard / stoplight / green->red
' =========================================================
Sub GenerateLineGraph()
    Dim src As Range, tws As Worksheet, scope As Range, ur As Range, f As Range
    Dim tagR As Long, tagC As Long
    Dim xC As Long, yC As Long, e1C As Long, e2C As Long, nameC As Long
    Dim r As Long, j As Long, lastR As Long, p As Long
    Dim blkRow() As Long, rc() As Long, blkLab() As String
    Dim nBlk As Long, nPts As Long
    Dim inBlk As Boolean
    Dim b As Variant, c As Variant
    Dim tag As String, xLog As Boolean, yLog As Boolean
    Dim xTitle As String, yTitle As String
    Dim hasE1 As Boolean, hasE2 As Boolean
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim gotX As Boolean, gotY As Boolean
    Dim dv As Double
    Dim co As ChartObject, ch As Chart, mks As Variant
    Dim xRng As Range, yRng As Range, e1Rng As Range, e2Rng As Range
    Dim W As Double, H As Double, sc As Double
    Dim fsT As Double, fsA As Double, fsL As Double
    Dim wAX As Double, wPL As Double, msz As Double
    Dim axMax As Double, axStep As Double, axMin As Double
    Dim lblC() As Long, fsTx As Double
    Dim useLine As Boolean, scheme As String, clr() As Long

    '--- 1. pick the range (you can switch workbooks in this dialog) --------
    On Error Resume Next
    Set src = Application.InputBox( _
        Prompt:="Switch to the workbook you want, then select/drag the data block." & vbCrLf & _
                "The selection must include the 'graph' tag cell.", _
        Title:="Generate Line Graph", Type:=8)
    On Error GoTo 0
    If src Is Nothing Then Exit Sub
    Set tws = src.Worksheet
    Set ur = tws.UsedRange
    Set scope = Application.Intersect(src, ur)
    If scope Is Nothing Then MsgBox "That selection contains no data.", vbExclamation: Exit Sub

    '--- 2. orient on the tag ----------------------------------------------
    Set f = scope.Find(What:="graph*", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    Do While Not f Is Nothing
        tag = LCase$(Trim$(CStr(f.Value)))
        If tag = "graph" Or tag = "graph linlog" Or tag = "graph loglin" Or tag = "graph loglog" Then Exit Do
        Set f = scope.FindNext(f)
        If f Is Nothing Then Exit Do
        If f.Row = tagR And f.Column = tagC Then Exit Do
    Loop
    If f Is Nothing Then
        MsgBox "Could not find a cell containing 'graph' inside the selection.", vbExclamation
        Exit Sub
    End If
    tagR = f.Row: tagC = f.Column
    tag = LCase$(Trim$(CStr(f.Value)))
    If Len(tag) = 12 Then
        xLog = (Mid$(tag, 7, 3) = "log")
        yLog = (Mid$(tag, 10, 3) = "log")
    End If

    xC = tagC: yC = tagC + 1: e1C = tagC + 2: e2C = tagC + 3: nameC = tagC - 1
    If nameC < 1 Then MsgBox "The 'graph' tag needs a legend column to its left.", vbExclamation: Exit Sub

    ' axis labels live one row below the tag (xgrapher: xdata{i+1,j}, xdata{i+1,j+1})
    b = tws.Cells(tagR + 1, xC).Value
    If Not IsNum(b) Then xTitle = CellText(b)
    b = tws.Cells(tagR + 1, yC).Value
    If Not IsNum(b) Then yTitle = CellText(b)

    lastR = Application.Min(scope.Row + scope.Rows.Count - 1, ur.Row + ur.Rows.Count - 1)
    If lastR < tagR + 2 Then MsgBox "No data rows found below the tag.", vbExclamation: Exit Sub

    '--- 3. blocks = SERIES (blank X+Y pair separates blocks) ---------------
    ReDim blkRow(1 To lastR - tagR + 2)
    ReDim blkLab(1 To lastR - tagR + 2)
    ReDim rc(1 To lastR - tagR + 2)
    ReDim lblC(1 To lastR - tagR + 2)
    inBlk = False
    For r = tagR + 2 To lastR
        b = tws.Cells(r, xC).Value
        c = tws.Cells(r, yC).Value
        If IsNum(b) And IsNum(c) Then
            If Not inBlk Then
                nBlk = nBlk + 1
                blkRow(nBlk) = r
                blkLab(nBlk) = CellText(tws.Cells(r, nameC).Value)
                rc(nBlk) = 0
                inBlk = True
            End If
            rc(nBlk) = rc(nBlk) + 1
            dv = CDbl(b)
            If Not gotX Then minX = dv: maxX = dv: gotX = True
            If dv < minX Then minX = dv
            If dv > maxX Then maxX = dv
            dv = CDbl(c)
            If Not gotY Then minY = dv: maxY = dv: gotY = True
            If dv < minY Then minY = dv
            If dv > maxY Then maxY = dv
            If IsNum(tws.Cells(r, e1C).Value) Then hasE1 = True
            If IsNum(tws.Cells(r, e2C).Value) Then hasE2 = True
        Else
            inBlk = False
        End If
    Next r
    If nBlk = 0 Then MsgBox "No numeric X/Y pairs found below the tag.", vbExclamation: Exit Sub
    For j = 1 To nBlk
        If rc(j) > nPts Then nPts = rc(j)
        ' text labels sit in the first non-numeric column after X,Y (xgrapher:
        ' j+2 if no error columns, j+3 after +/-Y, j+4 after +/-Y and +/-X)
        If ColHasText(tws, blkRow(j), rc(j), e1C) Then
            lblC(j) = e1C
        ElseIf ColHasText(tws, blkRow(j), rc(j), e2C) Then
            lblC(j) = e2C
        ElseIf Application.CountA(tws.Cells(blkRow(j), e2C + 1).Resize(rc(j), 1)) > 0 Then
            lblC(j) = e2C + 1
        End If
    Next j

    '--- 4. size + scaled type / line weights ------------------------------
    W = CH_WIDTH_IN * 72: H = W / MAT_ASPECT
    sc = W / MAT_FIG_W_PT
    fsT = Application.Max(1, Round(MAT_FS_TICKS * sc, 1))
    fsA = Application.Max(1, Round(MAT_FS_LABEL * sc, 1))
    fsL = fsT
    fsTx = LABEL_FS
    wAX = Application.Max(0.25, Round(MAT_LINEW * sc, 2))
    wPL = Application.Max(0.25, Round(MAT_PLOTW * sc, 2))
    msz = Application.Min(72, Application.Max(2, Round(MAT_MS * sc * 0.5, 0)))

    '--- 4b. line / colour settings from the Buttons sheet ----------------
    ReadSettings useLine, scheme
    clr = SeriesColors(scheme, nBlk)

    '--- 5. new chart (nothing existing is deleted) ------------------------
    Set co = tws.ChartObjects.Add( _
        Left:=tws.Cells(scope.Row, scope.Column + scope.Columns.Count).Left + 12, _
        Top:=tws.Cells(scope.Row, 1).Top, Width:=W, Height:=H)
    co.Name = "LineGraph_" & Format$(Now, "yyyymmdd_hhmmss")
    co.Placement = xlFreeFloating
    Set ch = co.Chart
    If useLine Then ch.ChartType = xlXYScatterLines Else ch.ChartType = xlXYScatter
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ' xgrapher marker order 'o^sdv<>'
    mks = Array(xlMarkerStyleCircle, xlMarkerStyleTriangle, xlMarkerStyleSquare, _
                xlMarkerStyleDiamond, xlMarkerStyleTriangle, xlMarkerStyleX, xlMarkerStyleStar)

    '--- 6. one series per block, wired to the source cells ----------------
    For j = 1 To nBlk
        Set xRng = tws.Cells(blkRow(j), xC).Resize(rc(j), 1)
        Set yRng = tws.Cells(blkRow(j), yC).Resize(rc(j), 1)
        Set e1Rng = tws.Cells(blkRow(j), e1C).Resize(rc(j), 1)
        Set e2Rng = tws.Cells(blkRow(j), e2C).Resize(rc(j), 1)
        With ch.SeriesCollection.NewSeries
            If Len(blkLab(j)) > 0 Then
                .Name = "='" & tws.Name & "'!" & tws.Cells(blkRow(j), nameC).Address(True, True)
            Else
                .Name = "Series " & j
            End If
            .XValues = xRng
            .Values = yRng
            .MarkerStyle = mks((j - 1) Mod 7)
            .MarkerSize = msz
            .MarkerForegroundColor = clr(j)
            ' MATLAB MarkerFaceColor = 0.5 + 0.5*color  (lightened fill)
            .MarkerBackgroundColor = RGB( _
                128 + (clr(j) And &HFF&) \ 2, _
                128 + ((clr(j) \ &H100&) And &HFF&) \ 2, _
                128 + ((clr(j) \ &H10000) And &HFF&) \ 2)
            If useLine Then
                .Format.Line.Visible = msoTrue
                .Format.Line.ForeColor.RGB = clr(j)
                .Format.Line.Weight = wPL
            End If
            ' col 3 = +/- Y error, col 4 = +/- X error (xgrapher errorbar signature)
            If hasE1 And lblC(j) <> e1C Then
                .ErrorBar Direction:=xlY, Include:=xlBoth, _
                          Type:=xlErrorBarTypeCustom, Amount:=e1Rng, MinusValues:=e1Rng
                With .ErrorBars
                    .EndStyle = xlCap
                    .Format.Line.ForeColor.RGB = clr(j)
                    .Format.Line.Weight = wAX
                End With
            End If
            If hasE2 And lblC(j) <> e1C And lblC(j) <> e2C Then
                .ErrorBar Direction:=xlX, Include:=xlBoth, _
                          Type:=xlErrorBarTypeCustom, Amount:=e2Rng, MinusValues:=e2Rng
            End If
        End With
    Next j

    '--- 7. chart area / fonts --------------------------------------------
    With ch.ChartArea
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Visible = msoFalse
    End With
    With ch.ChartArea.Font
        .Name = FONT_NAME: .Size = fsT: .Color = RGB(0, 0, 0): .Bold = False
    End With

    '--- 7b. point text labels (after the ChartArea font, which would resize them)
    If DO_LABELS Then
        For j = 1 To nBlk
            If lblC(j) > 0 Then
                AddPointLabels ch.SeriesCollection(j), _
                               tws.Cells(blkRow(j), lblC(j)).Resize(rc(j), 1), fsTx
            End If
        Next j
    End If

    '--- 8. axes ----------------------------------------------------------
    With ch.Axes(xlValue)
        If yLog Then
            .ScaleType = xlScaleLogarithmic
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
        Else
            NiceScale maxY, axMax, axStep
            If minY < 0 Then
                .MinimumScaleIsAuto = True
            Else
                .MinimumScale = 0
            End If
            .MaximumScale = axMax
            .MajorUnit = axStep
        End If
        .HasMajorGridlines = False
        .HasMinorGridlines = False
        .MajorTickMark = xlTickMarkInside
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionNextToAxis
        .TickLabels.Font.Name = FONT_NAME
        .TickLabels.Font.Size = fsT
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = wAX
        .HasTitle = (Len(yTitle) > 0)
        If Len(yTitle) > 0 Then
            With .AxisTitle
                .Orientation = xlUpward
                .Text = yTitle
                .Font.Name = FONT_NAME
                .Font.Size = fsA
                .Font.Bold = False
                .Font.Color = RGB(0, 0, 0)
            End With
        End If
    End With
    On Error Resume Next
    ch.Axes(xlValue).TickLabels.Offset = TICK_OFFSET
    On Error GoTo 0

    With ch.Axes(xlCategory)
        If xLog Then
            .ScaleType = xlScaleLogarithmic
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
        Else
            ' fitted to the data; starts at 0 only when that adds little
            NiceRange minX, maxX, axMin, axMax, axStep
            .MaximumScale = axMax
            .MinimumScale = axMin
            .MajorUnit = axStep
        End If
        .HasMajorGridlines = False
        .HasMinorGridlines = False
        .MajorTickMark = xlTickMarkInside
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionNextToAxis
        .TickLabels.Font.Name = FONT_NAME
        .TickLabels.Font.Size = fsT
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = wAX
        .HasTitle = (Len(xTitle) > 0)
        If Len(xTitle) > 0 Then
            With .AxisTitle
                .Text = xTitle
                .Font.Name = FONT_NAME
                .Font.Size = fsA
                .Font.Bold = False
                .Font.Color = RGB(0, 0, 0)
            End With
        End If
    End With
    On Error Resume Next
    ch.Axes(xlCategory).TickLabels.Offset = TICK_OFFSET
    On Error GoTo 0

    '--- 9. legend --------------------------------------------------------
    ch.HasLegend = True
    With ch.Legend
        .Position = xlLegendPositionCorner
        .IncludeInLayout = False
        .Font.Name = FONT_NAME
        .Font.Size = fsL
        .Font.Bold = False
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = wAX
    End With
    On Error Resume Next
    ch.Legend.Height = nBlk * fsL * LG_LINESP + 4
    On Error GoTo 0

    '--- 10. plot box -----------------------------------------------------
    With ch.PlotArea
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = wAX
    End With

    Application.ScreenUpdating = True
    On Error Resume Next
    tws.Parent.Activate
    tws.Activate
    co.Activate
    For p = 1 To 2
        With ch.PlotArea
            .InsideLeft = PA_LEFT * W
            .InsideTop = PA_TOP * H
            .InsideWidth = PA_WIDTH * W
            .InsideHeight = PA_HEIGHT * H
        End With
    Next p
    ch.Legend.Left = LG_RIGHT * W - ch.Legend.Width
    ch.Legend.Top = LG_TOP * H
    On Error GoTo 0

    ' pick the corner / column count that covers no data, raising the axis
    ' maxima if no corner is clear (falls back to top-right on any error)
    PlaceLegend ch, tws, W, H, nBlk, fsL * LG_LINESP, msz, fsTx, useLine, xLog, yLog, _
                blkRow, rc, xC, yC, IIf(hasE1, e1C, 0), lblC

    On Error Resume Next
    For p = 1 To 2
        With ch.PlotArea
            .InsideLeft = PA_LEFT * W
            .InsideTop = PA_TOP * H
            .InsideWidth = PA_WIDTH * W
            .InsideHeight = PA_HEIGHT * H
        End With
    Next p
    src.Select
    On Error GoTo 0

    On Error Resume Next
    ch.HasTitle = False
    ch.SetElement msoElementChartTitleNone
    On Error GoTo 0
    co.Width = W: co.Height = H
    ch.Refresh                               ' draw the final layout now
End Sub

' True if any cell in column c, rows r0..r0+n-1, holds non-blank text.
Private Function ColHasText(ws As Worksheet, r0 As Long, n As Long, c As Long) As Boolean
    Dim i As Long, v As Variant
    For i = 0 To n - 1
        v = ws.Cells(r0 + i, c).Value
        If VarType(v) = vbString Then
            If Len(Trim$(v)) > 0 Then ColHasText = True: Exit Function
        End If
    Next i
End Function

' Data labels to the right of each point, text taken from lblRng (xgrapher:
' text(x, y, label, 'HorizontalAlignment', 'left', 'VerticalAlignment', 'middle')).
Private Sub AddPointLabels(ser As Series, lblRng As Range, fs As Double)
    Dim i As Long, v As Variant, linked As Boolean

    ser.HasDataLabels = True
    ' Excel 2013+: "Value From Cells", so labels follow later edits to the sheet
    On Error Resume Next
    ser.DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, _
        "='" & Replace(lblRng.Worksheet.Name, "'", "''") & "'!" & lblRng.Address, 0
    linked = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    With ser.DataLabels
        If linked Then
            .ShowRange = True
            .ShowValue = False
        End If
        .ShowSeriesName = False
        .ShowCategoryName = False
        .ShowLegendKey = False
        .Position = xlLabelPositionRight
        .Font.Name = FONT_NAME
        .Font.Size = fs
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
    End With

    ' error cells get no label (xgrapher skips "ActiveX VT_ERROR:"); on older
    ' Excel, copy each cell's text in as a static label instead
    For i = 1 To Application.Min(lblRng.Rows.Count, ser.Points.Count)
        v = lblRng.Cells(i, 1).Value
        If IsError(v) Then
            ser.Points(i).HasDataLabel = False
        ElseIf Not linked Then
            If Len(Trim$(CStr(v))) = 0 Then
                ser.Points(i).HasDataLabel = False
            Else
                ser.Points(i).DataLabel.Text = CStr(v)
            End If
        End If
    Next i
End Sub

' Axis max rounded up to a nice value with <= 6 steps (same as the bar graph's;
' each module keeps its own copy so the two never depend on each other).
Private Sub NiceScale(ByVal raw As Double, ByRef axMax As Double, ByRef axStep As Double)
    Dim e As Double, m As Variant, i As Long, s As Double, n As Long
    If raw <= 0 Then axMax = 1: axStep = 0.2: Exit Sub
    e = 10 ^ Int(Log(raw) / Log(10#))
    m = Array(0.1, 0.2, 0.25, 0.5, 1, 2, 2.5, 5, 10)
    For i = LBound(m) To UBound(m)
        s = e * m(i)
        n = -Int(-raw / s)
        If n <= 6 Then axStep = s: axMax = s * n: Exit Sub
    Next i
    axStep = e: axMax = e * (-Int(-raw / e))
End Sub

' True for a number (or numeric text); False for blank, text or an error
' value such as #N/A, which would otherwise stop the macro on "<> """.
Private Function IsNum(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(v)) = 0 Then Exit Function
    End If
    IsNum = IsNumeric(v)
End Function

' Cell value as trimmed text; "" for an error value.
Private Function CellText(v As Variant) As String
    If Not IsError(v) Then CellText = Trim$(CStr(v))
End Function

' K6: 0 = markers only, anything else (1, blank) = connecting lines.
' K7: "stoplight", "green->red" (anything starting "green"), else standard.
' Missing sheet or cells fall back to lines + standard.
Private Sub ReadSettings(ByRef useLine As Boolean, ByRef scheme As String)
    Dim ws As Worksheet, v As Variant
    useLine = True: scheme = "standard"
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    v = ws.Range(LINE_CELL).Value
    If IsNum(v) Then
        If CDbl(v) = 0 Then useLine = False
    End If

    v = LCase$(CellText(ws.Range(SCHEME_CELL).Value))
    If v = "stoplight" Then
        scheme = "stoplight"
    ElseIf Left$(v, 5) = "green" Then
        scheme = "gradient"
    End If
End Sub

' One colour per series (1..n).
'   standard  - xgrapher color1 palette, repeating after 13
'   stoplight - green, yellow, red, dark red, repeating after 4
'   gradient  - hue slides from green (first series) to red (last) via yellow
Private Function SeriesColors(scheme As String, n As Long) As Long()
    Dim c() As Long, pal As Variant, i As Long, t As Double

    If n < 1 Then n = 1
    ReDim c(1 To n)
    Select Case scheme
        Case "stoplight"
            pal = Array(RGB(0, 150, 0), RGB(230, 190, 0), RGB(220, 20, 20), RGB(128, 0, 0))
        Case "gradient"
            For i = 1 To n
                If n > 1 Then t = (i - 1) / (n - 1) Else t = 0
                c(i) = HsvToRgb(120# * (1# - t), 0.95, 0.55)
            Next i
            SeriesColors = c
            Exit Function
        Case Else
            pal = Array(RGB(4, 65, 215), RGB(225, 0, 21), RGB(20, 20, 20), RGB(23, 105, 14), _
                        RGB(5, 66, 105), RGB(140, 87, 5), RGB(115, 5, 5), RGB(50, 21, 79), _
                        RGB(36, 186, 120), RGB(184, 93, 33), RGB(237, 143, 136), _
                        RGB(224, 219, 110), RGB(157, 162, 248))
    End Select
    For i = 1 To n
        c(i) = pal((i - 1) Mod (UBound(pal) + 1))
    Next i
    SeriesColors = c
End Function

' hue in degrees (0 = red, 60 = yellow, 120 = green), s and v in 0..1
Private Function HsvToRgb(hue As Double, s As Double, v As Double) As Long
    Dim c As Double, x As Double, m As Double, r As Double, g As Double, b As Double
    c = v * s
    x = c * (1 - Abs(((hue / 60#) - 2 * Int((hue / 60#) / 2)) - 1))
    m = v - c
    Select Case hue
        Case Is < 60:  r = c: g = x: b = 0
        Case Is < 120: r = x: g = c: b = 0
        Case Else:     r = 0: g = c: b = x
    End Select
    HsvToRgb = RGB(CInt((r + m) * 255), CInt((g + m) * 255), CInt((b + m) * 255))
End Function

'--- legend placement ---------------------------------------------------
' Tries every corner (top-right, top-left, bottom-right, bottom-left) with
' 1..3 legend columns and keeps the one that covers no markers, point
' labels, error bars or connecting lines, preferring top-right and one
' column. If none is clear it raises the Y maximum (up to 3 major steps),
' then the X maximum (up to 2), to open space; if that never clears a
' corner, the axes are put back and the least-covering option is used.
Private Sub PlaceLegend(ch As Chart, tws As Worksheet, ByVal W As Double, ByVal H As Double, _
                        ByVal n As Long, ByVal lineH As Double, ByVal msz As Double, _
                        ByVal fsTx As Double, ByVal useLine As Boolean, _
                        ByVal xLog As Boolean, ByVal yLog As Boolean, _
                        blkRow() As Long, rc() As Long, ByVal xC As Long, ByVal yC As Long, _
                        ByVal eC As Long, lblC() As Long)
    Dim oX() As Double, oY() As Double, offL() As Double, offT() As Double
    Dim offR() As Double, offB() As Double, nO As Long
    Dim sX1() As Double, sY1() As Double, sX2() As Double, sY2() As Double, nS As Long
    Dim j As Long, r As Long, tot As Long, x As Variant, y As Variant, e As Variant
    Dim txt As String, half As Double, hasPrev As Boolean, prevX As Double, prevY As Double
    Dim pL As Double, pT As Double, pW As Double, pH As Double, insX As Double, insY As Double
    Dim w1 As Double, h1 As Double, xLo As Double, xHi As Double, yLo As Double, yHi As Double
    Dim xHi0 As Double, yHi0 As Double, kMax As Long
    Dim iter As Long, k As Long, corner As Long, rws As Long, hits As Long, sc As Long
    Dim lw As Double, lh As Double, lx As Double, ly As Double
    Dim bSc As Long, bHits As Long, bW As Double, bH As Double, bX As Double, bY As Double
    Dim fSc As Long, fHits As Long, fW As Double, fH As Double, fX As Double, fY As Double

    On Error GoTo Done

    ' obstacles: point boxes (data anchor + point offsets) and line segments
    For j = 1 To n: tot = tot + rc(j): Next j
    If tot = 0 Then Exit Sub
    ReDim oX(1 To 2 * tot): ReDim oY(1 To 2 * tot): ReDim offL(1 To 2 * tot)
    ReDim offT(1 To 2 * tot): ReDim offR(1 To 2 * tot): ReDim offB(1 To 2 * tot)
    ReDim sX1(1 To 2 * tot): ReDim sY1(1 To 2 * tot): ReDim sX2(1 To 2 * tot): ReDim sY2(1 To 2 * tot)
    half = msz / 2 + 1
    For j = 1 To n
        hasPrev = False
        For r = blkRow(j) To blkRow(j) + rc(j) - 1
            x = tws.Cells(r, xC).Value: y = tws.Cells(r, yC).Value
            If IsNum(x) And IsNum(y) Then
                nO = nO + 1
                oX(nO) = CDbl(x): oY(nO) = CDbl(y)
                offL(nO) = -half: offR(nO) = half: offT(nO) = -half: offB(nO) = half
                If DO_LABELS And lblC(j) > 0 Then
                    txt = CellText(tws.Cells(r, lblC(j)).Value)
                    If Len(txt) > 0 Then
                        nO = nO + 1
                        oX(nO) = CDbl(x): oY(nO) = CDbl(y)
                        offL(nO) = half: offR(nO) = half + 4 + Len(txt) * fsTx * 0.55
                        offT(nO) = -fsTx * 0.6: offB(nO) = fsTx * 0.6
                    End If
                End If
                If eC > 0 And lblC(j) <> eC Then
                    e = tws.Cells(r, eC).Value
                    If IsNum(e) Then
                        nS = nS + 1
                        sX1(nS) = CDbl(x): sY1(nS) = CDbl(y) - Abs(CDbl(e))
                        sX2(nS) = CDbl(x): sY2(nS) = CDbl(y) + Abs(CDbl(e))
                    End If
                End If
                If useLine And hasPrev Then
                    nS = nS + 1
                    sX1(nS) = prevX: sY1(nS) = prevY: sX2(nS) = CDbl(x): sY2(nS) = CDbl(y)
                End If
                prevX = CDbl(x): prevY = CDbl(y): hasPrev = True
            End If
        Next r
    Next j

    pL = PA_LEFT * W: pT = PA_TOP * H: pW = PA_WIDTH * W: pH = PA_HEIGHT * H
    insX = (PA_LEFT + PA_WIDTH - LG_RIGHT) * W
    insY = (LG_TOP - PA_TOP) * H
    ' Measure the legend as Excel lays it out on its own (one column, every
    ' entry shown) instead of estimating line heights, which clipped the
    ' last entry. IncludeInLayout = False keeps the plot area where it is.
    With ch.Legend
        .Position = xlLegendPositionRight
        .IncludeInLayout = False
    End With
    ch.Refresh                               ' let Excel lay the legend out now,
    DoEvents                                 ' not after the macro ends
    With ch.Legend
        w1 = .Width
        h1 = .Height
    End With
    If h1 < n * lineH Then h1 = n * lineH + 4     ' Excel squeezed it: fall back
    kMax = n: If kMax > 3 Then kMax = 3
    xHi0 = ch.Axes(xlCategory).MaximumScale
    yHi0 = ch.Axes(xlValue).MaximumScale

    For iter = 0 To 5
        xLo = ch.Axes(xlCategory).MinimumScale: xHi = ch.Axes(xlCategory).MaximumScale
        yLo = ch.Axes(xlValue).MinimumScale: yHi = ch.Axes(xlValue).MaximumScale
        bSc = 2147483647
        For k = 1 To kMax
            rws = -Int(-n / k)
            lw = k * w1                          ' full width per column: no wrapping
            lh = rws * h1 / n + 2                ' measured height per entry
            If lw <= pW - 2 * insX And lh <= pH - 2 * insY Then
                For corner = 0 To 3                  ' 0 TR, 1 TL, 2 BR, 3 BL
                    If corner Mod 2 = 0 Then lx = pL + pW - insX - lw Else lx = pL + insX
                    If corner < 2 Then ly = pT + insY Else ly = pT + pH - insY - lh
                    hits = LegendHits(lx, ly, lw, lh, nO, oX, oY, offL, offT, offR, offB, _
                                      nS, sX1, sY1, sX2, sY2, xLo, xHi, yLo, yHi, _
                                      xLog, yLog, pL, pT, pW, pH)
                    sc = hits * 1000 + (k - 1) * 4 + corner
                    If sc < bSc Then
                        bSc = sc: bHits = hits: bW = lw: bH = lh: bX = lx: bY = ly
                    End If
                Next corner
            End If
        Next k
        If bSc = 2147483647 Then Exit Sub        ' legend too big for any layout
        If iter = 0 Then fSc = bSc: fHits = bHits: fW = bW: fH = bH: fX = bX: fY = bY
        If bHits = 0 Then Exit For

        If iter < 3 Then                          ' room at the top first
            With ch.Axes(xlValue)
                If yLog Then .MaximumScale = yHi * 10 Else .MaximumScale = yHi + .MajorUnit
            End With
        ElseIf iter < 5 Then                      ' then room at the right
            With ch.Axes(xlCategory)
                If xLog Then .MaximumScale = xHi * 10 Else .MaximumScale = xHi + .MajorUnit
            End With
        End If
    Next iter

    If bHits > 0 And bHits >= fHits Then          ' bumping didn't help: undo it
        ch.Axes(xlCategory).MaximumScale = xHi0
        ch.Axes(xlValue).MaximumScale = yHi0
        bW = fW: bH = fH: bX = fX: bY = fY
    End If

    With ch.Legend
        .Width = bW
        .Height = bH
        .Left = bX
        .Top = bY
    End With
Done:
End Sub

' Number of obstacles the legend box (plus a 2 pt margin) would cover.
Private Function LegendHits(ByVal lx As Double, ByVal ly As Double, ByVal lw As Double, _
                            ByVal lh As Double, ByVal nO As Long, oX() As Double, oY() As Double, _
                            offL() As Double, offT() As Double, offR() As Double, offB() As Double, _
                            ByVal nS As Long, sX1() As Double, sY1() As Double, _
                            sX2() As Double, sY2() As Double, _
                            ByVal xLo As Double, ByVal xHi As Double, _
                            ByVal yLo As Double, ByVal yHi As Double, _
                            ByVal xLog As Boolean, ByVal yLog As Boolean, _
                            ByVal pL As Double, ByVal pT As Double, _
                            ByVal pW As Double, ByVal pH As Double) As Long
    Const PAD As Double = 2
    Dim i As Long, m As Long, p1x As Double, p1y As Double, p2x As Double, p2y As Double
    Dim qx As Double, qy As Double, x0 As Double, x1 As Double, y0 As Double, y1 As Double

    x0 = lx - PAD: x1 = lx + lw + PAD: y0 = ly - PAD: y1 = ly + lh + PAD
    For i = 1 To nO
        If MapV(oX(i), xLo, xHi, xLog, pL, pW, False, p1x) And _
           MapV(oY(i), yLo, yHi, yLog, pT, pH, True, p1y) Then
            If p1x + offR(i) > x0 And p1x + offL(i) < x1 And p1y + offB(i) > y0 And p1y + offT(i) < y1 Then
                LegendHits = LegendHits + 1
            End If
        End If
    Next i
    For i = 1 To nS
        If MapV(sX1(i), xLo, xHi, xLog, pL, pW, False, p1x) And _
           MapV(sY1(i), yLo, yHi, yLog, pT, pH, True, p1y) And _
           MapV(sX2(i), xLo, xHi, xLog, pL, pW, False, p2x) And _
           MapV(sY2(i), yLo, yHi, yLog, pT, pH, True, p2y) Then
            For m = 0 To 20
                qx = p1x + (p2x - p1x) * m / 20: qy = p1y + (p2y - p1y) * m / 20
                If qx > x0 And qx < x1 And qy > y0 And qy < y1 Then
                    LegendHits = LegendHits + 1
                    Exit For
                End If
            Next m
        End If
    Next i
End Function

' Data value -> chart position along one axis (flip for Y, which grows
' downward on screen). False for values the axis can't show.
Private Function MapV(ByVal v As Double, ByVal lo As Double, ByVal hi As Double, _
                      ByVal isLog As Boolean, ByVal p0 As Double, ByVal span As Double, _
                      ByVal flip As Boolean, ByRef outP As Double) As Boolean
    Dim f As Double
    If isLog Then
        If v <= 0 Or lo <= 0 Or hi <= lo Then Exit Function
        f = (Log(v) - Log(lo)) / (Log(hi) - Log(lo))
    Else
        If hi <= lo Then Exit Function
        f = (v - lo) / (hi - lo)
    End If
    If flip Then outP = p0 + (1 - f) * span Else outP = p0 + f * span
    MapV = True
End Function

' Axis bounds fitted to lo..hi: the smallest step of 1, 2, 2.5 or 5 x 10^n
' that spans the data in at most 6 steps, with the bounds on whole steps.
' For non-negative data the minimum drops to 0 when that stretches the
' axis by no more than a quarter of its span; otherwise it stays near the
' data (e.g. 0.33..0.55 -> 0.30..0.55, but 0.05..0.55 -> 0..0.6).
Private Sub NiceRange(ByVal lo As Double, ByVal hi As Double, ByRef axMin As Double, _
                      ByRef axMax As Double, ByRef axStep As Double)
    Dim span As Double, e As Double, m As Variant, i As Long, s As Double
    Dim a As Double, b As Double

    If hi < lo Then a = lo: lo = hi: hi = a
    span = hi - lo
    If span <= 0 Then span = Abs(hi)
    If span <= 0 Then span = 1
    e = 10 ^ Int(Log(span) / Log(10#))
    m = Array(0.1, 0.2, 0.25, 0.5, 1, 2, 2.5, 5, 10, 20)
    For i = LBound(m) To UBound(m)
        s = e * m(i)
        a = s * Int(lo / s + 0.000000001)          ' floor to a whole step
        b = -s * Int(-(hi / s) + 0.000000001)      ' ceiling to a whole step
        If b <= a Then b = a + s
        If (b - a) / s <= 6.000000001 Then Exit For
    Next i

    If lo >= 0 And a > 0 And a <= 0.25 * (b - a) Then
        NiceRange 0, hi, axMin, axMax, axStep
        Exit Sub
    End If
    axMin = Round(a, 12): axMax = Round(b, 12): axStep = Round(s, 12)
End Sub
