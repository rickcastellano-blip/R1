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
    Dim co As ChartObject, ch As Chart, cols As Variant, mks As Variant
    Dim xRng As Range, yRng As Range, e1Rng As Range, e2Rng As Range
    Dim W As Double, H As Double, sc As Double
    Dim fsT As Double, fsA As Double, fsL As Double
    Dim wAX As Double, wPL As Double, msz As Double
    Dim axMax As Double, axStep As Double, axMin As Double

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
    If Not IsNumeric(b) Then xTitle = Trim$(CStr(b))
    b = tws.Cells(tagR + 1, yC).Value
    If Not IsNumeric(b) Then yTitle = Trim$(CStr(b))

    lastR = Application.Min(scope.Row + scope.Rows.Count - 1, ur.Row + ur.Rows.Count - 1)
    If lastR < tagR + 2 Then MsgBox "No data rows found below the tag.", vbExclamation: Exit Sub

    '--- 3. blocks = SERIES (blank X+Y pair separates blocks) ---------------
    ReDim blkRow(1 To lastR - tagR + 2)
    ReDim blkLab(1 To lastR - tagR + 2)
    ReDim rc(1 To lastR - tagR + 2)
    inBlk = False
    For r = tagR + 2 To lastR
        b = tws.Cells(r, xC).Value
        c = tws.Cells(r, yC).Value
        If IsNumeric(b) And b <> "" And IsNumeric(c) And c <> "" Then
            If Not inBlk Then
                nBlk = nBlk + 1
                blkRow(nBlk) = r
                blkLab(nBlk) = Trim$(CStr(tws.Cells(r, nameC).Value))
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
            If IsNumeric(tws.Cells(r, e1C).Value) And tws.Cells(r, e1C).Value <> "" Then hasE1 = True
            If IsNumeric(tws.Cells(r, e2C).Value) And tws.Cells(r, e2C).Value <> "" Then hasE2 = True
        Else
            inBlk = False
        End If
    Next r
    If nBlk = 0 Then MsgBox "No numeric X/Y pairs found below the tag.", vbExclamation: Exit Sub
    For j = 1 To nBlk
        If rc(j) > nPts Then nPts = rc(j)
    Next j

    '--- 4. size + scaled type / line weights ------------------------------
    W = CH_WIDTH_IN * 72: H = W / MAT_ASPECT
    sc = W / MAT_FIG_W_PT
    fsT = Application.Max(1, Round(MAT_FS_TICKS * sc, 1))
    fsA = Application.Max(1, Round(MAT_FS_LABEL * sc, 1))
    fsL = fsT
    wAX = Application.Max(0.25, Round(MAT_LINEW * sc, 2))
    wPL = Application.Max(0.25, Round(MAT_PLOTW * sc, 2))
    msz = Application.Min(72, Application.Max(2, Round(MAT_MS * sc * 0.5, 0)))

    '--- 5. new chart (nothing existing is deleted) ------------------------
    Set co = tws.ChartObjects.Add( _
        Left:=tws.Cells(scope.Row, scope.Column + scope.Columns.Count).Left + 12, _
        Top:=tws.Cells(scope.Row, 1).Top, Width:=W, Height:=H)
    co.Name = "LineGraph_" & Format$(Now, "yyyymmdd_hhmmss")
    co.Placement = xlFreeFloating
    Set ch = co.Chart
    ch.ChartType = xlXYScatterLines
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ' xgrapher color1 palette
    cols = Array(RGB(4, 65, 215), RGB(225, 0, 21), RGB(20, 20, 20), RGB(23, 105, 14), _
                 RGB(5, 66, 105), RGB(140, 87, 5), RGB(115, 5, 5), RGB(50, 21, 79), _
                 RGB(36, 186, 120), RGB(184, 93, 33), RGB(237, 143, 136), _
                 RGB(224, 219, 110), RGB(157, 162, 248))
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
            .MarkerForegroundColor = cols((j - 1) Mod 13)
            ' MATLAB MarkerFaceColor = 0.5 + 0.5*color  (lightened fill)
            .MarkerBackgroundColor = RGB( _
                128 + (cols((j - 1) Mod 13) And &HFF&) \ 2, _
                128 + ((cols((j - 1) Mod 13) \ &H100&) And &HFF&) \ 2, _
                128 + ((cols((j - 1) Mod 13) \ &H10000) And &HFF&) \ 2)
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = cols((j - 1) Mod 13)
            .Format.Line.Weight = wPL
            ' col 3 = +/- Y error, col 4 = +/- X error (xgrapher errorbar signature)
            If hasE1 Then
                .ErrorBar Direction:=xlY, Include:=xlBoth, _
                          Type:=xlErrorBarTypeCustom, Amount:=e1Rng, MinusValues:=e1Rng
                With .ErrorBars
                    .EndStyle = xlCap
                    .Format.Line.ForeColor.RGB = cols((j - 1) Mod 13)
                    .Format.Line.Weight = wAX
                End With
            End If
            If hasE2 Then
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
            NiceScale maxX, axMax, axStep
            If minX < 0 Then
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
    src.Select
    On Error GoTo 0

    On Error Resume Next
    ch.HasTitle = False
    ch.SetElement msoElementChartTitleNone
    On Error GoTo 0
    co.Width = W: co.Height = H

    MsgBox "Chart created in '" & tws.Parent.Name & "' on '" & tws.Name & "'." & vbCrLf & _
           nBlk & " series, " & nPts & " points max.", vbInformation
End Sub
