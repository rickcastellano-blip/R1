Option Explicit

' ===== style, taken from resize.m (550x400 px figure) =====
Private Const FONT_NAME     As String = "Arial"
Private Const MAT_FIG_W_PT  As Double = 412.5    ' 550 px @ 96 dpi
Private Const MAT_ASPECT    As Double = 1.375    ' 550/400
Private Const MAT_FS_TICKS  As Double = 13.5     ' resize.m: axes FontSize
Private Const MAT_FS_LABEL  As Double = 16       ' resize.m: x/ylabel FontSize
Private Const MAT_LINEW     As Double = 0.5      ' MATLAB default line width
Private Const CH_WIDTH_IN   As Double = 5     ' <-- set your output size here
Private Const PA_LEFT       As Double = 0.1505   ' resize.m axes Position
Private Const PA_TOP        As Double = 0.071
Private Const PA_WIDTH      As Double = 0.8
Private Const PA_HEIGHT     As Double = 0.75
Private Const LG_RIGHT      As Double = 0.938    ' 'northeast'
Private Const LG_TOP        As Double = 0.086
Private Const LG_LINESP     As Double = 1.45
Private Const TICK_OFFSET   As Long = 20
Private Const GAP_GROUPED   As Long = 150
Private Const GAP_SINGLE    As Long = 80
' =========================================================

Sub GenerateStrengthBarGraph()
    Dim src As Range, tws As Worksheet, scope As Range, ur As Range, f As Range
    Dim tagR As Long, tagC As Long, xlR As Long, xlC As Long
    Dim labC As Long, valC As Long, errC As Long, nameC As Long
    Dim r As Long, i As Long, j As Long, lastR As Long, p As Long
    Dim blkRow() As Long, blkLab() As String, rc() As Long
    Dim nBlk As Long, nRow As Long, nLab As Long
    Dim inBlk As Boolean
    Dim b As Variant, c As Variant, dv As Double, de As Double
    Dim yTitle As String, maxVal As Double
    Dim co As ChartObject, ch As Chart, cols As Variant
    Dim vRng As Range, eRng As Range, catRng As Range
    Dim W As Double, H As Double, sc As Double
    Dim fsT As Double, fsA As Double, fsL As Double
    Dim wAX As Double, axMax As Double, axStep As Double
    Dim maxLen As Long, rot As Long, paH As Double

    '--- 1. pick the range (you can switch workbooks in this dialog) --------
    On Error Resume Next
    Set src = Application.InputBox( _
        Prompt:="Switch to the workbook you want, then select/drag the data block." & vbCrLf & _
                "The selection must include the 'bargraph' tag cell.", _
        Title:="Generate Bar Graph", Type:=8)
    On Error GoTo 0
    If src Is Nothing Then Exit Sub

    Set tws = src.Worksheet
    Set ur = tws.UsedRange
    Set scope = Application.Intersect(src, ur)
    If scope Is Nothing Then MsgBox "That selection contains no data.", vbExclamation: Exit Sub

    '--- 2. orient on the tags ---------------------------------------------
    Set f = scope.Find(What:="bargraph", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        MsgBox "Could not find a cell containing 'bargraph' inside the selection.", vbExclamation
        Exit Sub
    End If
    tagR = f.Row: tagC = f.Column

    Set f = scope.Find(What:="xlabels", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then xlR = f.Row: xlC = f.Column

    valC = tagC: errC = tagC + 1: labC = tagC - 1
    If labC < 1 Then MsgBox "The 'bargraph' tag needs a label column to its left.", vbExclamation: Exit Sub
    If xlC > 1 Then nameC = xlC - 1 Else nameC = labC

    b = tws.Cells(tagR + 1, valC).Value
    If Not IsNumeric(b) Then yTitle = Trim$(CStr(b))
    If Len(yTitle) = 0 Then yTitle = "Value"

    If xlR > 0 Then
        lastR = xlR - 1
    Else
        lastR = Application.Min(scope.Row + scope.Rows.Count - 1, ur.Row + ur.Rows.Count - 1)
    End If
    If lastR < tagR + 2 Then MsgBox "No data rows found below the tag.", vbExclamation: Exit Sub

    '--- 3. blocks = SERIES (xgrapher: bars(:,plotnum) = one block) ---------
    ReDim blkRow(1 To lastR - tagR + 2)
    ReDim blkLab(1 To lastR - tagR + 2)
    ReDim rc(1 To lastR - tagR + 2)
    inBlk = False
    For r = tagR + 2 To lastR
        b = tws.Cells(r, valC).Value
        c = tws.Cells(r, errC).Value
        If IsNumeric(b) And b <> "" Then
            If Not inBlk Then
                nBlk = nBlk + 1
                blkRow(nBlk) = r
                blkLab(nBlk) = Trim$(CStr(tws.Cells(r, labC).Value))
                rc(nBlk) = 0
                inBlk = True
            End If
            rc(nBlk) = rc(nBlk) + 1
            dv = CDbl(b): de = 0
            If IsNumeric(c) And c <> "" Then de = CDbl(c)
            If dv + de > maxVal Then maxVal = dv + de
        Else
            inBlk = False
        End If
    Next r
    If nBlk = 0 Then MsgBox "No numeric data found below the tag.", vbExclamation: Exit Sub
    For j = 1 To nBlk
        If rc(j) > nRow Then nRow = rc(j)
    Next j

    '--- 4. xlabels = CATEGORIES (x axis ticks) ----------------------------
    Set catRng = Nothing
    If xlR > 0 Then
        nLab = 0
        Do While Len(Trim$(CStr(tws.Cells(xlR + nLab, nameC).Value))) > 0 And nLab < nRow
            nLab = nLab + 1
        Loop
        If nLab > 0 Then Set catRng = tws.Cells(xlR, nameC).Resize(nLab, 1)
        maxLen = 0
        For i = 0 To nLab - 1
            If Len(CStr(tws.Cells(xlR + i, nameC).Value)) > maxLen Then _
                maxLen = Len(CStr(tws.Cells(xlR + i, nameC).Value))
        Next i
    End If

    rot = 0: paH = PA_HEIGHT

    '--- 5. size + scaled type / line weights ------------------------------
    W = CH_WIDTH_IN * 72: H = W / MAT_ASPECT
    sc = W / MAT_FIG_W_PT
    fsT = Application.Max(1, Round(MAT_FS_TICKS * sc, 1))
    fsA = Application.Max(1, Round(MAT_FS_LABEL * sc, 1))
    fsL = fsT
    wAX = Application.Max(0.25, Round(MAT_LINEW * sc, 2))

    '--- 6. new chart (nothing existing is deleted) ------------------------
    Set co = tws.ChartObjects.Add( _
        Left:=tws.Cells(scope.Row, scope.Column + scope.Columns.Count).Left + 12, _
        Top:=tws.Cells(scope.Row, 1).Top, Width:=W, Height:=H)
    co.Name = "BarGraph_" & Format$(Now, "yyyymmdd_hhmmss")
    co.Placement = xlFreeFloating
    Set ch = co.Chart
    ch.ChartType = xlColumnClustered
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    cols = Array(RGB(0, 114, 189), RGB(217, 83, 25), RGB(237, 177, 32), _
                 RGB(126, 47, 142), RGB(119, 172, 48), RGB(77, 190, 238), RGB(162, 20, 47))

    '--- 7. one series per block, wired to the source cells ----------------
    For j = 1 To nBlk
        Set vRng = tws.Cells(blkRow(j), valC).Resize(rc(j), 1)
        Set eRng = tws.Cells(blkRow(j), errC).Resize(rc(j), 1)
        With ch.SeriesCollection.NewSeries
            If Len(blkLab(j)) > 0 Then
                .Name = "='" & tws.Name & "'!" & tws.Cells(blkRow(j), labC).Address(True, True)
            Else
                .Name = "Series " & j
            End If
            .Values = vRng
            If Not catRng Is Nothing Then .XValues = catRng
            .Format.Fill.Visible = msoTrue
            .Format.Fill.ForeColor.RGB = cols((j - 1) Mod 7)
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = wAX
            .ErrorBar Direction:=xlY, Include:=xlBoth, _
                      Type:=xlErrorBarTypeCustom, Amount:=eRng, MinusValues:=eRng
            With .ErrorBars
                .EndStyle = xlCap
                .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
                .Format.Line.Weight = wAX
            End With
        End With
    Next j

    With ch.ChartGroups(1)
        If nBlk = 1 Then .GapWidth = GAP_SINGLE Else .GapWidth = GAP_GROUPED
        .Overlap = 0
    End With

    '--- 8. chart area / fonts --------------------------------------------
    With ch.ChartArea
        .Format.Fill.Visible = msoTrue
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Visible = msoFalse
    End With
    With ch.ChartArea.Font
        .Name = FONT_NAME: .Size = fsT: .Color = RGB(0, 0, 0): .Bold = False
    End With

    '--- 9. axes ----------------------------------------------------------
    NiceScale maxVal, axMax, axStep
    With ch.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = axMax
        .MajorUnit = axStep
        .HasMajorGridlines = False
        .HasMinorGridlines = False
        .MajorTickMark = xlTickMarkInside
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionNextToAxis
        .TickLabels.NumberFormat = "0"
        .TickLabels.Font.Name = FONT_NAME
        .TickLabels.Font.Size = fsT
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = wAX
        .HasTitle = True
        With .AxisTitle
            .Orientation = xlUpward
            .Text = yTitle
            .Font.Name = FONT_NAME
            .Font.Size = fsA
            .Font.Bold = False
            .Font.Color = RGB(0, 0, 0)
        End With
    End With
    On Error Resume Next
    ch.Axes(xlValue).TickLabels.Offset = TICK_OFFSET
    On Error GoTo 0

    With ch.Axes(xlCategory)
        .HasTitle = False
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionLow
        .TickLabels.Font.Name = FONT_NAME
        .TickLabels.Font.Size = fsT
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Format.Line.Weight = wAX
    End With
    On Error Resume Next
    ch.Axes(xlCategory).TickLabels.Orientation = rot
    ch.Axes(xlCategory).TickLabels.Offset = TICK_OFFSET
    On Error GoTo 0

    '--- 10. legend -------------------------------------------------------
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

    '--- 11. plot box -----------------------------------------------------
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
            .InsideHeight = paH * H
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
           nRow & " x-labels x " & nBlk & " series.", vbInformation
End Sub


Public Sub NiceScale(ByVal raw As Double, ByRef axMax As Double, ByRef axStep As Double)
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
