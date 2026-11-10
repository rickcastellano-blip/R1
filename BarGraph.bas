Option Explicit

' ===== style, taken from resize.m (550x400 px figure) =====
Private Const FONT_NAME     As String = "Arial"
Private Const MAT_FIG_W_PT  As Double = 412.5    ' 550 px @ 96 dpi
Private Const MAT_ASPECT    As Double = 1.375    ' 550/400
Private Const MAT_FS_TICKS  As Double = 13.5     ' resize.m: axes FontSize
Private Const MAT_FS_LABEL  As Double = 16       ' resize.m: x/ylabel FontSize
Private Const MAT_LINEW     As Double = 0.5      ' MATLAB default line width
Private Const CH_WIDTH_IN   As Double = 5        ' default figure width (in) when the Width cell is blank
Private Const WIDTH_CELL    As String = "J3"     ' figure width in inches (1-20)
Private Const PA_LEFT       As Double = 0.1505   ' resize.m axes Position
Private Const PA_TOP        As Double = 0.071
Private Const PA_WIDTH      As Double = 0.8
Private Const PA_HEIGHT     As Double = 0.75
Private Const LG_RIGHT      As Double = 0.938    ' 'northeast'
Private Const LG_TOP        As Double = 0.086
Private Const LG_LINESP     As Double = 1.45
Private Const TICK_OFFSET   As Long = 20
Private Const GAP_GROUPED   As Long = 150
' colour scheme cell on the Buttons sheet of this workbook: standard / darkening pairs
Private Const SETTINGS_SHEET As String = "Buttons"
Private Const FIG_FOLDER    As String = "Fig Pix"   ' PNG of every chart, next to the workbook
Private Const SCHEME_CELL    As String = "G3"
Private Const LABELDIR_CELL  As String = "H3"   ' bar labels: horiz or vert
Private Const LABELFS_CELL   As String = "I3"   ' bar-label font size (pt); blank = LABEL_FS
Private Const GAP_SINGLE    As Long = 80
Private Const LABEL_FS      As Double = 10       ' default bar-label font size (pt, not scaled)
' =========================================================

Sub GenerateStrengthBarGraph()
    Dim src As Range, tws As Worksheet, scope As Range, ur As Range, f As Range
    Dim tagR As Long, tagC As Long, xlR As Long, xlC As Long
    Dim labC As Long, valC As Long, errC As Long, nameC As Long
    Dim r As Long, i As Long, j As Long, lastR As Long, p As Long
    Dim blkRow() As Long, blkLab() As String, rc() As Long, lblC() As Long
    Dim nBlk As Long, nRow As Long, nLab As Long
    Dim inBlk As Boolean
    Dim b As Variant, c As Variant, dv As Double, de As Double
    Dim yTitle As String, maxVal As Double, gotNum As Boolean
    Dim co As ChartObject, ch As Chart, clr() As Long
    Dim vRng As Range, eRng As Range, catRng As Range
    Dim W As Double, H As Double, sc As Double
    Dim fsT As Double, fsA As Double, fsL As Double
    Dim wAX As Double, axMax As Double, axStep As Double
    Dim maxLen As Long, rot As Long, paH As Double
    Dim legCorner As Long, legExtraW As Double, legExtraH As Double, vertLbl As Boolean, lblFs As Double

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
    If Not IsNum(b) Then yTitle = CellText(b)
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
    ' A block runs over consecutive numeric or error (#N/A) values; an error
    ' is a group with no bar, so the rows still line up with the x labels.
    ' A blank (or text) cell ends the block.
    For r = tagR + 2 To lastR
        b = tws.Cells(r, valC).Value
        c = tws.Cells(r, errC).Value
        If IsNum(b) Or IsError(b) Then
            If Not inBlk Then
                nBlk = nBlk + 1
                blkRow(nBlk) = r
                blkLab(nBlk) = CellText(tws.Cells(r, labC).Value)
                rc(nBlk) = 0
                inBlk = True
            End If
            rc(nBlk) = rc(nBlk) + 1
            If IsNum(b) Then
                gotNum = True
                dv = CDbl(b): de = 0
                If IsNum(c) Then de = CDbl(c)
                If dv + de > maxVal Then maxVal = dv + de
            End If
        Else
            inBlk = False
        End If
    Next r
    If Not gotNum Then nBlk = 0
    If nBlk = 0 Then MsgBox "No numeric data found below the tag.", vbExclamation: Exit Sub
    For j = 1 To nBlk
        If rc(j) > nRow Then nRow = rc(j)
    Next j

    '--- 4. xlabels = CATEGORIES (x axis ticks) ----------------------------
    Set catRng = Nothing
    If xlR > 0 Then
        nLab = 0
        Do While Len(CellText(tws.Cells(xlR + nLab, nameC).Value)) > 0 And nLab < nRow
            nLab = nLab + 1
        Loop
        If nLab > 0 Then Set catRng = tws.Cells(xlR, nameC).Resize(nLab, 1)
        maxLen = 0
        For i = 0 To nLab - 1
            If Len(CellText(tws.Cells(xlR + i, nameC).Value)) > maxLen Then _
                maxLen = Len(CellText(tws.Cells(xlR + i, nameC).Value))
        Next i
    End If

    rot = 0: paH = PA_HEIGHT

    '--- 5. size + scaled type / line weights ------------------------------
    ' Only the width follows the Width cell: the height, type, line weights and
    ' the plot's left and right margins stay those of the default width, so a
    ' wider figure just has a wider plot (whole pixels, so every chart of a
    ' width is the same size)
    W = Px(ReadWidthIn() * 72): H = Px(BaseW() / MAT_ASPECT)
    sc = BaseW() / MAT_FIG_W_PT
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

    ' Bar labels: text in the +/- column itself (then no error bars), or
    ' anything (text, or a number such as a % shown as is) in the column after
    ' a numeric +/- column. An empty column ends the search, so a table further
    ' right is never taken as labels.
    ReDim lblC(1 To nBlk)
    For j = 1 To nBlk
        If ColHasText(tws, blkRow(j), rc(j), errC) Then
            lblC(j) = errC
        ElseIf ColIsBlank(tws, blkRow(j), rc(j), errC) Then
            lblC(j) = 0
        ElseIf Not ColIsBlank(tws, blkRow(j), rc(j), errC + 1) Then
            lblC(j) = errC + 1
        End If
    Next j

    clr = BarColors(ReadScheme(), nBlk)
    vertLbl = ReadVertLabels()
    lblFs = ReadLabelSize()

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
            .Format.Fill.ForeColor.RGB = clr(j)
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = wAX
            If lblC(j) <> errC Then
                .ErrorBar Direction:=xlY, Include:=xlBoth, _
                          Type:=xlErrorBarTypeCustom, Amount:=eRng, MinusValues:=eRng
                With .ErrorBars
                    .EndStyle = xlCap
                    .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
                    .Format.Line.Weight = wAX
                End With
            End If
        End With
    Next j

    With ch.ChartGroups(1)
        If nBlk = 1 Then .GapWidth = GAP_SINGLE Else .GapWidth = GAP_GROUPED
        .Overlap = 0
    End With

    '--- 8. chart area / fonts --------------------------------------------
    With ch.ChartArea                            ' figure background: no fill
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With
    With ch.ChartArea.Font
        .Name = FONT_NAME: .Size = fsT: .Color = RGB(0, 0, 0): .Bold = False
    End With

    '--- 8b. bar labels (after the ChartArea font, which would resize them) --
    For j = 1 To nBlk
        If lblC(j) > 0 Then
            AddBarLabels ch.SeriesCollection(j), tws.Cells(blkRow(j), lblC(j)).Resize(rc(j), 1), vertLbl, lblFs
        End If
    Next j

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
            ' linked only when the cell has a title (otherwise it stays "Value")
            If Len(CellText(tws.Cells(tagR + 1, valC).Value)) > 0 Then LinkTitle ch.Axes(xlValue).AxisTitle, tws.Cells(tagR + 1, valC)
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
    With ch.PlotArea                             ' axis background: white
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
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
            .InsideLeft = Px(PA_LEFT * BaseW())
            .InsideTop = Px(PA_TOP * H)
            .InsideWidth = Px(W - (1 - PA_WIDTH) * BaseW())
            .InsideHeight = Px(paH * H)
        End With
    Next p
    ch.Legend.Left = W - (1 - LG_RIGHT) * BaseW() - ch.Legend.Width
    ch.Legend.Top = LG_TOP * H
    On Error GoTo 0

    ' pick the corner / column count that covers no bar, raising the axis
    ' maximum if no corner is clear, or put the legend outside the plot
    PlaceBarLegend ch, tws, W, H, nBlk, nRow, ch.ChartGroups(1).GapWidth, fsL * LG_LINESP, _
                   blkRow, rc, valC, errC, lblC, vertLbl, lblFs, legCorner, legExtraW, legExtraH

    On Error Resume Next
    For p = 1 To 2
        With ch.PlotArea
            .InsideLeft = Px(PA_LEFT * BaseW())
            .InsideTop = Px(PA_TOP * H)
            .InsideWidth = Px(W - (1 - PA_WIDTH) * BaseW())
            .InsideHeight = Px(paH * H)
        End With
    Next p
    src.Select
    On Error GoTo 0

    On Error Resume Next
    ch.HasTitle = False
    ch.SetElement msoElementChartTitleNone
    On Error GoTo 0

    co.Width = W + Px(legExtraW): co.Height = H + Px(legExtraH)   ' wider/taller only for an outside legend
    AnchorLegend ch, legCorner, W, H         ' last, once nothing else will move
    FitBarLabels ch, tws, nBlk, blkRow, rc, valC, errC, lblC, legCorner, W, H
    ch.Refresh                               ' draw the final layout now
    SaveFigure co                            ' PNG copy in <workbook folder>\Fig Pix
End Sub

' Keeps every bar label inside the plot box and off the legend: while any
' label pokes out of the top or touches the legend, the Y maximum goes up a
' step (up to 6) and the labels are placed again.
Private Sub FitBarLabels(ch As Chart, tws As Worksheet, ByVal n As Long, blkRow() As Long, _
                         rc() As Long, ByVal valC As Long, ByVal errC As Long, lblC() As Long, _
                         ByVal legCorner As Long, ByVal W As Double, ByVal H As Double)
    Dim iter As Long, j As Long, i As Long, bad As Boolean, hasLbl As Boolean
    Dim pT As Double, gL As Double, gT As Double, gR As Double, gB As Double
    Dim dl As DataLabel
    For j = 1 To n
        If lblC(j) > 0 Then hasLbl = True
    Next j
    If Not hasLbl Then Exit Sub
    On Error Resume Next
    For iter = 0 To 6
        If iter > 0 Then
            With ch.Axes(xlValue)
                .MaximumScale = .MaximumScale + .MajorUnit
            End With
            AnchorLegend ch, legCorner, W, H
        End If
        LiftBarLabels ch, tws, n, blkRow, rc, valC, errC, lblC
        ch.Refresh
        DoEvents
        pT = ch.PlotArea.InsideTop
        If legCorner >= 0 And legCorner <= 3 Then
            gL = ch.Legend.Left: gT = ch.Legend.Top
            gR = gL + ch.Legend.Width: gB = gT + ch.Legend.Height
        End If
        bad = False
        For j = 1 To n
            If lblC(j) > 0 Then
                For i = 1 To ch.SeriesCollection(j).Points.Count
                    If ch.SeriesCollection(j).Points(i).HasDataLabel Then
                        Set dl = ch.SeriesCollection(j).Points(i).DataLabel
                        If dl.Top < pT Then bad = True
                        If legCorner >= 0 And legCorner <= 3 Then
                            If dl.Left < gR And dl.Left + dl.Width > gL And _
                               dl.Top < gB And dl.Top + dl.Height > gT Then bad = True
                        End If
                    End If
                    If bad Then Exit For
                Next i
            End If
            If bad Then Exit For
        Next j
        If Not bad Then Exit For
    Next iter
End Sub

' A label sits on the bar's top by default, where the error bar's cap runs
' through it: move each one up to sit just above its error bar instead.
Private Sub LiftBarLabels(ch As Chart, tws As Worksheet, ByVal n As Long, blkRow() As Long, _
                          rc() As Long, ByVal valC As Long, ByVal errC As Long, lblC() As Long)
    Dim j As Long, i As Long, v As Variant, e As Variant
    Dim yLo As Double, yHi As Double, pT As Double, pH As Double, yTop As Double
    On Error Resume Next
    ch.Refresh
    DoEvents
    yLo = ch.Axes(xlValue).MinimumScale: yHi = ch.Axes(xlValue).MaximumScale
    pT = ch.PlotArea.InsideTop: pH = ch.PlotArea.InsideHeight
    If yHi <= yLo Or pH <= 0 Then Exit Sub
    For j = 1 To n
        If lblC(j) > 0 And lblC(j) <> errC Then
            For i = 1 To rc(j)
                v = tws.Cells(blkRow(j) + i - 1, valC).Value
                e = tws.Cells(blkRow(j) + i - 1, errC).Value
                If IsNum(v) And IsNum(e) Then
                    If ch.SeriesCollection(j).Points(i).HasDataLabel Then
                        yTop = pT + (1 - (CDbl(v) + Abs(CDbl(e)) - yLo) / (yHi - yLo)) * pH
                        With ch.SeriesCollection(j).Points(i).DataLabel
                            .Position = xlLabelPositionOutsideEnd   ' back to the bar top first
                            If .Top + .Height > yTop - 1 Then .Top = yTop - .Height - 1
                        End With
                    End If
                End If
            Next i
        End If
    Next j
End Sub


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

' True if every cell in column c, rows r0..r0+n-1, is empty.
Private Function ColIsBlank(ws As Worksheet, r0 As Long, n As Long, c As Long) As Boolean
    ColIsBlank = (Application.CountA(ws.Cells(r0, c).Resize(n, 1)) = 0)
End Function

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

' Data labels above each bar, text taken from lblRng (as the Line Graph's
' point labels: linked to the cells on Excel 2013+, static text before that).
Private Sub AddBarLabels(ser As Series, lblRng As Range, ByVal vertLbl As Boolean, _
                         ByVal lblFs As Double)
    Dim i As Long, v As Variant, linked As Boolean

    ser.HasDataLabels = True
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
        .Position = xlLabelPositionOutsideEnd
        If vertLbl Then .Orientation = xlUpward Else .Orientation = xlHorizontal
        .Font.Name = FONT_NAME
        .Font.Size = lblFs
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
    End With

    ' error and blank cells get no label
    For i = 1 To Application.Min(lblRng.Rows.Count, ser.Points.Count)
        v = lblRng.Cells(i, 1).Value
        If IsError(v) Then
            ser.Points(i).HasDataLabel = False
        ElseIf Len(Trim$(CStr(v))) = 0 Then
            ser.Points(i).HasDataLabel = False
        ElseIf Not linked Then
            ser.Points(i).DataLabel.Text = lblRng.Cells(i, 1).Text   ' as displayed
        End If
    Next i
End Sub

' The default figure width in points (CH_WIDTH_IN), whole pixels: the
' height, type, weights and plot margins are set for this width.
Private Function BaseW() As Double
    BaseW = Px(CH_WIDTH_IN * 72)
End Function

' Figure width in inches from the Width cell on the Buttons sheet (1-20);
' blank or not a number gives CH_WIDTH_IN.
Private Function ReadWidthIn() As Double
    Dim ws As Worksheet, v As Variant
    ReadWidthIn = CH_WIDTH_IN
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    v = ws.Range(WIDTH_CELL).Value
    If IsNum(v) Then
        If CDbl(v) > 0 Then ReadWidthIn = Application.Min(20, Application.Max(1, CDbl(v)))
    End If
End Function

' Saves the finished chart as a PNG in the "Fig Pix" folder next to the
' workbook (made if missing), named <sheet> <chart name>.png. A workbook saved
' to OneDrive / SharePoint reports a web address as its path, so that is
' turned back into the local synced folder. Says so only if it fails.
Private Sub SaveFigure(co As ChartObject)
    Dim dir0 As String, fName As String, bad As Variant

    dir0 = LocalPath(ThisWorkbook.Path)
    If Len(dir0) = 0 Then
        MsgBox "The chart was made, but no PNG was saved: the workbook has no local folder" & _
               " (save it first, or check that OneDrive is syncing).", vbExclamation
        Exit Sub
    End If
    dir0 = dir0 & "\" & FIG_FOLDER
    fName = co.Parent.Name & " " & co.Name
    For Each bad In Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
        fName = Replace(fName, bad, "_")
    Next bad

    On Error GoTo Fail
    If Len(Dir(dir0, vbDirectory)) = 0 Then MkDir dir0
    co.Chart.Export Filename:=dir0 & "\" & fName & ".png", FilterName:="PNG"
    Exit Sub
Fail:
    MsgBox "The chart was made, but the PNG could not be saved to" & vbCrLf & dir0 & _
           vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub

' A local folder for a workbook path; an https OneDrive / SharePoint address
' is mapped onto the synced OneDrive folder (personal: after the 4th "/",
' business: after "/Documents"). "" if no local folder is found.
Private Function LocalPath(ByVal p As String) As String
    Dim k As Long, i As Long, rest As String, root As Variant
    If LCase$(Left$(p, 4)) <> "http" Then LocalPath = p: Exit Function
    k = InStr(1, p, "/Documents/", vbTextCompare)
    If k > 0 Then
        rest = Mid$(p, k + Len("/Documents/"))
    Else
        k = 0
        For i = 1 To 4                           ' https://d.docs.live.net/<id>/...
            k = InStr(k + 1, p, "/")
            If k = 0 Then Exit Function
        Next i
        rest = Mid$(p, k + 1)
    End If
    rest = Replace(rest, "/", "\")
    For Each root In Array(Environ$("OneDriveCommercial"), Environ$("OneDriveConsumer"), Environ$("OneDrive"))
        If Len(root) > 0 Then
            If Len(Dir(root & "\" & rest, vbDirectory)) > 0 Then
                LocalPath = root & "\" & rest
                Exit Function
            End If
        End If
    Next root
End Function

' Round a size in points to whole screen pixels (0.75 pt at 96 dpi), so
' Excel never has to round it differently from one chart to the next.
Private Function Px(ByVal v As Double) As Double
    Px = Round(v / 0.75, 0) * 0.75
End Function

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

' Link an axis title to the cell it was read from, so editing the cell
' updates the chart (same as typing =Sheet!$A$1 in the title). If Excel
' can't link it (older versions), the title keeps its plain text.
Private Sub LinkTitle(t As AxisTitle, c As Range)
    On Error Resume Next
    t.Formula = "='" & Replace(c.Worksheet.Name, "'", "''") & "'!" & c.Address(True, True)
End Sub

'--- legend placement (same approach as the line graph's) -------------------
' Tries every corner (top-right, top-left, bottom-right, bottom-left) with 1
' or 2 legend columns and keeps the one that covers no bar (error bars
' included), preferring top-right and one column. If none is clear it raises
' the Y maximum one major step at a time (up to 3) and takes the first level
' where any layout is clear. If no level clears a corner, or the legend is too
' big for the plot, the axis is put back and the legend goes outside the plot
' on the right, with the chart widened to fit it.
Private Sub PlaceBarLegend(ch As Chart, tws As Worksheet, ByVal W As Double, ByVal H As Double, _
                           ByVal n As Long, ByVal nCat As Long, ByVal gapPct As Double, _
                           ByVal lineH As Double, blkRow() As Long, rc() As Long, _
                           ByVal valC As Long, ByVal errC As Long, lblC() As Long, _
                           ByVal vertLbl As Boolean, ByVal lblFs As Double, ByRef legCorner As Long, ByRef extraW As Double, ByRef extraH As Double)
    Dim bx0() As Double, bx1() As Double, by0() As Double, by1() As Double, nB As Long
    Dim bL() As Double, t As Variant
    Dim j As Long, cIdx As Long, v As Variant, e As Variant, barW As Double, barTop As Double
    Dim pL As Double, pT As Double, pW As Double, pH As Double, insX As Double, insY As Double
    Dim w1 As Double, h1 As Double, eW As Double, eH As Double, padW As Double, padH As Double
    Dim i As Long, eWh As Double, gapH As Double, kMax As Long, yHi0 As Double
    Dim yLo As Double, yHi As Double, iter As Long, k As Long, corner As Long, rws As Long
    Dim hits As Long, sc As Long, lw As Double, lh As Double, lx As Double, ly As Double
    Dim bSc As Long, bHits As Long, bK As Long, bC As Long
    Dim bW As Double, bH As Double, bX As Double, bY As Double

    On Error GoTo Done
    legCorner = 0: extraW = 0: extraH = 0
    If n < 1 Or nCat < 1 Then Exit Sub

    ' bars as rectangles: x in category units (category c spans c-0.5..c+0.5),
    ' y from 0 to value + error
    ReDim bx0(1 To n * nCat): ReDim bx1(1 To n * nCat)
    ReDim by0(1 To n * nCat): ReDim by1(1 To n * nCat)
    ReDim bL(1 To n * nCat)
    barW = 1# / (n + gapPct / 100#)
    For j = 1 To n
        For cIdx = 1 To rc(j)
            v = tws.Cells(blkRow(j) + cIdx - 1, valC).Value
            If IsNum(v) Then
                e = tws.Cells(blkRow(j) + cIdx - 1, errC).Value
                barTop = CDbl(v)
                If IsNum(e) And lblC(j) <> errC Then barTop = barTop + Abs(CDbl(e))
                nB = nB + 1
                If lblC(j) > 0 Then                  ' a label sits above the bar
                    t = tws.Cells(blkRow(j) + cIdx - 1, lblC(j)).Value
                    If Len(CellText(t)) > 0 Then
                        If vertLbl Then            ' rotated: as tall as the text is long
                            bL(nB) = Len(tws.Cells(blkRow(j) + cIdx - 1, lblC(j)).Text) * lblFs * 0.6 + 4
                        Else
                            bL(nB) = lblFs * 1.3 + 2
                        End If
                    End If
                End If
                bx0(nB) = cIdx - 0.5 + (gapPct / 100#) * barW / 2# + (j - 1) * barW
                bx1(nB) = bx0(nB) + barW
                If barTop >= 0 Then by0(nB) = 0: by1(nB) = barTop Else by0(nB) = barTop: by1(nB) = 0
            End If
        Next cIdx
    Next j

    pL = Px(PA_LEFT * BaseW()): pT = Px(PA_TOP * H): pW = Px(W - (1 - PA_WIDTH) * BaseW()): pH = Px(PA_HEIGHT * H)
    insX = (PA_LEFT + PA_WIDTH - LG_RIGHT) * BaseW()
    insY = (LG_TOP - PA_TOP) * H

    ' measure the legend as Excel lays it out (one column, every entry shown)
    With ch.Legend
        .Position = xlLegendPositionRight
        .IncludeInLayout = False
    End With
    ch.Refresh
    DoEvents
    w1 = ch.Legend.Width: h1 = ch.Legend.Height
    If h1 < n * lineH Then h1 = n * lineH + 4
    On Error Resume Next
    For i = 1 To ch.Legend.LegendEntries.Count
        If ch.Legend.LegendEntries(i).Width > eW Then eW = ch.Legend.LegendEntries(i).Width
        If ch.Legend.LegendEntries(i).Height > eH Then eH = ch.Legend.LegendEntries(i).Height
    Next i
    On Error GoTo Done
    If eW <= 0 Or eW > w1 Or eH <= 0 Or n * eH > h1 Then eW = w1: eH = h1 / n
    padW = w1 - eW
    padH = h1 - n * eH
    ' Excel can report entries spread over a tall box; one line of text is
    ' all an entry needs, so the rows are never taller than that
    If eH > lineH Then eH = lineH
    If padH > 8 Then padH = 8
    If padH < 4 Then padH = 4

    ' horizontal layout (used for 2 columns) spaces entries differently
    eWh = eW: gapH = 12
    If n >= 2 Then
        On Error Resume Next
        ch.Legend.Position = xlLegendPositionTop
        ch.Legend.IncludeInLayout = False
        ch.Refresh
        DoEvents
        eWh = 0
        For i = 1 To ch.Legend.LegendEntries.Count
            If ch.Legend.LegendEntries(i).Width > eWh Then eWh = ch.Legend.LegendEntries(i).Width
        Next i
        If ch.Legend.LegendEntries(2).Top = ch.Legend.LegendEntries(1).Top Then
            gapH = ch.Legend.LegendEntries(2).Left - ch.Legend.LegendEntries(1).Left - eWh
        End If
        ch.Legend.Position = xlLegendPositionRight
        ch.Legend.IncludeInLayout = False
        On Error GoTo Done
        If eWh <= 0 Then eWh = eW
        If gapH < 4 Or gapH > 40 Then gapH = 12
    End If
    kMax = n: If kMax > 2 Then kMax = 2
    yHi0 = ch.Axes(xlValue).MaximumScale

    For iter = 0 To 3
        yLo = ch.Axes(xlValue).MinimumScale: yHi = ch.Axes(xlValue).MaximumScale
        bSc = 2147483647
        For k = 1 To kMax
            rws = -Int(-n / k)
            If k = 1 Then lw = eW + padW Else lw = k * eWh + (k - 1) * gapH + padW + 4
            lh = rws * eH + padH + 1
            If lw <= pW - 2 * insX And lh <= pH - 2 * insY Then
                For corner = 0 To 3                  ' 0 TR, 1 TL, 2 BR, 3 BL
                    If corner Mod 2 = 0 Then lx = pL + pW - insX - lw Else lx = pL + insX
                    If corner < 2 Then ly = pT + insY Else ly = pT + pH - insY - lh
                    hits = BarHits(lx, ly, lw, lh, nB, bx0, bx1, by0, by1, bL, nCat, _
                                   yLo, yHi, pL, pT, pW, pH)
                    sc = hits * 1000 + (k - 1) * 4 + corner
                    If sc < bSc Then
                        bSc = sc: bHits = hits: bK = k: bC = corner: bW = lw: bH = lh: bX = lx: bY = ly
                    End If
                Next corner
            End If
        Next k
        If bSc = 2147483647 Then GoTo Outside    ' legend too big for any layout
        If bHits = 0 Then Exit For
        If iter < 3 Then                          ' room at the top
            With ch.Axes(xlValue)
                .MaximumScale = yHi + .MajorUnit
            End With
        End If
    Next iter
    If bHits > 0 Then GoTo Outside              ' no corner is clear at any level

    ' a 2-column legend only wraps into rows in horizontal (top) layout
    If bK > 1 Then
        ch.Legend.Position = xlLegendPositionTop
        ch.Legend.IncludeInLayout = False
        ch.Refresh
    End If
    With ch.Legend
        .Width = bW
        .Height = bH
        .Left = bX
        .Top = bY
    End With
    legCorner = bC
    Exit Sub

Outside:
    ' legend outside the plot on the right, chart widened (and made taller if
    ' needed) first, sized from the longest series name and widened until no
    ' name wraps
    Dim fs As Double, maxLen As Long, nm As String
    Dim wrapTry As Long, hMinE As Double, hMaxE As Double, hEnt As Double
    ch.Axes(xlValue).MaximumScale = yHi0
    fs = lineH / LG_LINESP
    For j = 1 To ch.SeriesCollection.Count
        nm = ch.SeriesCollection(j).Name
        If Len(nm) > maxLen Then maxLen = Len(nm)
    Next j
    lw = 30 + maxLen * fs * 0.52
    lh = n * fs * LG_LINESP + 12
    extraW = lw + 24
    If pT + lh + 8 > H Then extraH = pT + lh + 8 - H
    ch.Parent.Width = W + extraW
    ch.Parent.Height = H + extraH
    For j = 1 To 2
        With ch.PlotArea
            .InsideLeft = pL: .InsideTop = pT: .InsideWidth = pW: .InsideHeight = pH
        End With
    Next j
    With ch.Legend
        .Position = xlLegendPositionRight
        .IncludeInLayout = False
        .Left = pL + pW + 18
        .Top = pT
        .Width = lw
        .Height = lh
    End With
    For wrapTry = 1 To 4
        ch.Refresh
        DoEvents
        hMinE = 0: hMaxE = 0
        On Error Resume Next
        For j = 1 To ch.Legend.LegendEntries.Count
            hEnt = ch.Legend.LegendEntries(j).Height
            If hMinE = 0 Or hEnt < hMinE Then hMinE = hEnt
            If hEnt > hMaxE Then hMaxE = hEnt
        Next j
        On Error GoTo Done
        If hMinE <= 0 Or hMaxE < hMinE * 1.5 Then Exit For
        lw = lw * 1.15
        extraW = lw + 24
        ch.Parent.Width = W + extraW
        With ch.Legend
            .Left = pL + pW + 18
            .Width = lw
            .Height = lh
        End With
    Next wrapTry
    legCorner = 4
Done:
End Sub

' Number of bars (with their labels) the legend box plus a 2 pt margin would cover.
Private Function BarHits(ByVal lx As Double, ByVal ly As Double, ByVal lw As Double, _
                         ByVal lh As Double, ByVal nB As Long, bx0() As Double, bx1() As Double, _
                         by0() As Double, by1() As Double, bL() As Double, _
                         ByVal nCat As Long, _
                         ByVal yLo As Double, ByVal yHi As Double, ByVal pL As Double, _
                         ByVal pT As Double, ByVal pW As Double, ByVal pH As Double) As Long
    Const PAD As Double = 2
    Dim i As Long, xa As Double, xb As Double, ya As Double, yb As Double
    If yHi <= yLo Then Exit Function
    For i = 1 To nB
        xa = pL + (bx0(i) - 0.5) / nCat * pW
        xb = pL + (bx1(i) - 0.5) / nCat * pW
        ya = pT + (1 - (by1(i) - yLo) / (yHi - yLo)) * pH      ' bar top on screen
        yb = pT + (1 - (by0(i) - yLo) / (yHi - yLo)) * pH      ' bar bottom
        If bL(i) > 0 Then                                       ' label above the bar
            ya = ya - bL(i)                                     ' (labels sit above the error bar)
        End If
        If xb > lx - PAD And xa < lx + lw + PAD And yb > ly - PAD And ya < ly + lh + PAD Then
            BarHits = BarHits + 1
        End If
    Next i
End Function

' Align the legend to its corner of the plot box as finally drawn (or beside
' it for corner 4), using the size Excel actually gave it.
' corner: 0 top-right, 1 top-left, 2 bottom-right, 3 bottom-left, 4 outside right.
Private Sub AnchorLegend(ch As Chart, ByVal corner As Long, ByVal W As Double, ByVal H As Double)
    Dim insX As Double, insY As Double, p As Long
    On Error Resume Next
    ch.Refresh
    DoEvents
    insX = (PA_LEFT + PA_WIDTH - LG_RIGHT) * BaseW()
    insY = (LG_TOP - PA_TOP) * H
    ' Excel can adjust one side of the plot box when another is set, so it is
    ' set until it reads back as asked (every chart gets the same plot box)
    For p = 1 To 4
        With ch.PlotArea
            .InsideLeft = Px(PA_LEFT * BaseW())
            .InsideTop = Px(PA_TOP * H)
            .InsideWidth = Px(W - (1 - PA_WIDTH) * BaseW())
            .InsideHeight = Px(PA_HEIGHT * H)
        End With
        If Abs(ch.PlotArea.InsideLeft - Px(PA_LEFT * BaseW())) < 0.4 And _
           Abs(ch.PlotArea.InsideTop - Px(PA_TOP * H)) < 0.4 And _
           Abs(ch.PlotArea.InsideWidth - Px(W - (1 - PA_WIDTH) * BaseW())) < 0.4 And _
           Abs(ch.PlotArea.InsideHeight - Px(PA_HEIGHT * H)) < 0.4 Then Exit For
    Next p
    With ch.Legend
        If corner = 4 Then
            .Left = ch.PlotArea.InsideLeft + ch.PlotArea.InsideWidth + 18
            .Top = ch.PlotArea.InsideTop
            ch.Refresh
            Exit Sub
        End If
        If corner Mod 2 = 0 Then
            .Left = ch.PlotArea.InsideLeft + ch.PlotArea.InsideWidth - insX - .Width
        Else
            .Left = ch.PlotArea.InsideLeft + insX
        End If
        If corner < 2 Then
            .Top = ch.PlotArea.InsideTop + insY
        Else
            .Top = ch.PlotArea.InsideTop + ch.PlotArea.InsideHeight - insY - .Height
        End If
    End With
    ch.Refresh
End Sub

' Buttons!G3: "darkening pairs" (any case); anything else, a blank cell or a
' missing sheet means standard.
' Bar-label direction from Buttons!H3: "vert" turns them upright, anything
' else (horiz, blank) leaves them horizontal.
Private Function ReadVertLabels() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ReadVertLabels = (LCase$(CellText(ws.Range(LABELDIR_CELL).Value)) = "vert")
End Function

' Bar-label font size from Buttons!I3 (points, 4-72); blank or not a number
' gives LABEL_FS.
Private Function ReadLabelSize() As Double
    Dim ws As Worksheet, v As Variant
    ReadLabelSize = LABEL_FS
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    v = ws.Range(LABELFS_CELL).Value
    If IsNum(v) Then
        If CDbl(v) > 0 Then ReadLabelSize = Application.Min(72, Application.Max(4, CDbl(v)))
    End If
End Function

Private Function ReadScheme() As String
    Dim ws As Worksheet
    ReadScheme = "standard"
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Select Case LCase$(CellText(ws.Range(SCHEME_CELL).Value))
        Case "darkening pairs": ReadScheme = "pairs"
        Case "darkening triplets": ReadScheme = "triplets"
        Case "pairs": ReadScheme = "lightpairs"
        Case "triplets": ReadScheme = "lighttriplets"
    End Select
End Function

' One fill colour per series (1..n).
'   standard - the xgrapher.m color1 list (as Line Graph), repeating after 13
'   pairs    - light blue, light red, blue, red, dark blue, dark red, repeating
'   triplets - light blue, light red, light gray, blue, red, gray,
'              dark blue, dark red, black, repeating
'   Blue and red are the standard ones; light = halfway to white, dark = 60 %
'   of the colour. Gray is (72,72,72), light gray halfway to white, and the
'   darkest is the standard black. A single pair / triplet (n up to 2 / 3) is
'   just the middle row (blue, red, gray).
'   lightpairs - ("pairs") each standard colour as a light / full pair, in the
'              standard order: light blue, blue, light red, red, gray, black,
'              ..., repeating after 26; light = halfway to white
'   lighttriplets - ("triplets") the same with a dark one after each pair:
'              light blue, blue, dark blue, light red, red, dark red, ...,
'              repeating after 39; dark = 60 % of the colour. Black, which can't
'              darken, is light gray, gray, black as in darkening triplets
Private Function BarColors(scheme As String, n As Long) As Long()
    Dim c() As Long, pal As Variant, hues As Variant, i As Long, k As Long
    Dim hR As Long, hG As Long, hB As Long

    If n < 1 Then n = 1
    ReDim c(1 To n)
    If scheme = "pairs" Or scheme = "triplets" Then
        hues = Array(RGB(4, 65, 215), RGB(225, 0, 21), RGB(72, 72, 72))
        If scheme = "pairs" Then k = 2 Else k = 3
        ReDim pal(0 To 3 * k - 1)
        For i = 0 To k - 1
            hR = hues(i) And &HFF&
            hG = (hues(i) \ &H100&) And &HFF&
            hB = (hues(i) \ &H10000) And &HFF&
            pal(i) = RGB((hR + 255) \ 2, (hG + 255) \ 2, (hB + 255) \ 2)
            pal(k + i) = hues(i)
            pal(2 * k + i) = RGB(hR * 0.6, hG * 0.6, hB * 0.6)
        Next i
        If k = 3 Then pal(8) = RGB(20, 20, 20)   ' gray darkens to the standard black
        If n <= k Then                           ' a single group: the plain colours
            For i = 1 To n
                c(i) = hues(i - 1)
            Next i
            BarColors = c
            Exit Function
        End If
    Else
        pal = Array(RGB(4, 65, 215), RGB(225, 0, 21), RGB(20, 20, 20), RGB(23, 105, 14), _
                    RGB(5, 66, 105), RGB(140, 87, 5), RGB(115, 5, 5), RGB(50, 21, 79), _
                    RGB(36, 186, 120), RGB(184, 93, 33), RGB(237, 143, 136), _
                    RGB(224, 219, 110), RGB(157, 162, 248))
        If scheme = "lightpairs" Or scheme = "lighttriplets" Then
            If scheme = "lightpairs" Then i = 2 Else i = 3     ' colours per group
            hues = pal
            ReDim pal(0 To i * (UBound(hues) + 1) - 1)
            For k = 0 To UBound(hues)
                hR = hues(k) And &HFF&
                hG = (hues(k) \ &H100&) And &HFF&
                hB = (hues(k) \ &H10000) And &HFF&
                pal(i * k) = RGB((hR + 255) \ 2, (hG + 255) \ 2, (hB + 255) \ 2)
                pal(i * k + 1) = hues(k)
                If i = 3 Then pal(i * k + 2) = RGB(hR * 0.6, hG * 0.6, hB * 0.6)
                If i = 3 And hues(k) = RGB(20, 20, 20) Then
                    ' black can't darken: light gray, gray, black, as in darkening triplets
                    pal(i * k) = RGB(163, 163, 163)
                    pal(i * k + 1) = RGB(72, 72, 72)
                    pal(i * k + 2) = RGB(20, 20, 20)
                End If
            Next k
        End If
    End If
    For i = 1 To n
        c(i) = pal((i - 1) Mod (UBound(pal) + 1))
    Next i
    BarColors = c
End Function
