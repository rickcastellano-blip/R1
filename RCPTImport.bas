Option Explicit

' RCPT (ASTM C1202) import for TTi CPX400DP logger CSVs: one row (and current
' chart) per 6 h run on the RCPT sheet; a log holding several runs gives
' several rows. This module also holds the shared analysis (RCPTAnalyzeLog),
' which the RCPT Recovered button calls after reading its own file format.

'=========================== CONFIG ===========================
Private Const SHEET_NAME    As String = "RCPT"
Private Const LAST_COL      As Long = 16         ' A:P - used to find an empty row
Private Const TEST_SECONDS  As Double = 21600#   ' 6 hours
Private Const V_THRESHOLD   As Double = 30#      ' voltage-on detection
Private Const H_OFFSET_SEC  As Double = 30#      ' initial-current sample point
Private Const TAIL_POINTS   As Long = 10         ' samples averaged to extrapolate a short log
Private Const OFF_GAP_SEC   As Double = 60#      ' voltage off this long ends a run; a file can
                                                 ' hold several runs, one row each
Private Const WRITE_FORMULAS       As Boolean = True
Private Const WRITE_AREA_RATIO     As Boolean = True
Private Const DEFAULT_AREA_RATIO   As Double = 1#
Private Const CHART_WIDTH_PT       As Double = 320#
Private Const CHART_HEIGHT_PT      As Double = 170#
Private Const CHART_MAX_POINTS     As Long = 1200    ' plotted points; keeps the SERIES
                                                    ' formula under Excel's 8192-char cap
Private Const PARTIAL_FILL         As Long = 13431551  ' RGB(255,242,204) light yellow

'====================== BUTTON: Add RCPT Run ==================
Public Sub AddRCPTRunFromCSV()
    Dim path As String, msg As String
    Dim tS() As Double, vA() As Double, aA() As Double, n As Long

    path = PickCSVFile()
    If Len(path) = 0 Then Exit Sub

    Application.StatusBar = "Reading " & Dir(path) & " ..."
    msg = ReadLog(path, tS, vA, aA, n)
    Application.StatusBar = False
    If Len(msg) > 0 Then MsgBox msg, vbExclamation, "Add RCPT Run": Exit Sub

    RCPTAnalyzeLog tS, vA, aA, n, SampleFromFileName(path)
End Sub

'================ shared analysis (also used by RCPT Recovered) ============
' Everything after a log is read: each 6 h run becomes a row on the RCPT sheet
' with its chart, then one summary message. tS = seconds since 1899-12-30,
' vA = volts, aA = amps, 1..n; sample = sample name. RCPTRecovered calls
' this, so keep its arguments unchanged unless both modules are updated together.
Public Sub RCPTAnalyzeLog(tS() As Double, vA() As Double, aA() As Double, _
                          ByVal n As Long, ByVal sample As String)
    Dim dTest As Date, msg As String
    Dim hA As Double, i6A As Double, charge As Double, onHours As Double
    Dim nUsed As Long
    Dim ws As Worksheet, r As Long
    Dim tArr() As Double, aArr() As Double, nPts As Long
    Dim isPartial As Boolean, availHrs As Double, estQ As Double, cls As String
    Dim iFrom As Long, iPrev As Long, nRuns As Long, summary As String

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    ' one row per run: each pass starts where the last run ended
    iFrom = 1
    Do While iFrom <= n
        iPrev = iFrom
        msg = AnalyzeRun(tS, vA, aA, n, iFrom, dTest, hA, i6A, charge, onHours, nUsed, _
                         tArr, aArr, nPts, isPartial, availHrs)
        If Len(msg) > 0 Then
            If nRuns = 0 Then
                Application.ScreenUpdating = True
                MsgBox msg, vbExclamation, "Add RCPT Run"
                Exit Sub
            End If
            ' no further run in the log (or a fragment too short to use)
            If InStr(1, msg, "No sample above", vbTextCompare) = 1 Then Exit Do
            summary = summary & vbCrLf & "Run " & nRuns + 1 & " skipped: " & msg & vbCrLf
        Else
            nRuns = nRuns + 1

            ' --- short log: add it flagged as partial ----------------------
            If isPartial Then
                estQ = charge + i6A * (TEST_SECONDS - availHrs * 3600#)
                cls = RCPTClass(estQ)
            Else
                estQ = charge
                cls = RCPTClass(charge)
            End If

            r = FirstEmptyRow(ws)

            ws.Cells(r, 1).Value = dTest
            ws.Cells(r, 2).Value = sample
            ws.Cells(r, 8).Value = Round(hA * 1000#, 0)
            ws.Cells(r, 9).Value = i6A / hA
            ws.Cells(r, 12).Value = Round(estQ, 0)
            ws.Cells(r, 14).Value = cls
            ws.Cells(r, 16).Value = onHours

            If WRITE_AREA_RATIO Then
                If Len(ws.Cells(r, 7).Value) = 0 Then ws.Cells(r, 7).Value = DEFAULT_AREA_RATIO
            End If

            If WRITE_FORMULAS Then
                ws.Cells(r, 10).Formula = "=H" & r & "/1000*60*60*6*(I" & r & "+1)/2*G" & r
                ws.Cells(r, 13).Formula = "=L" & r & "*G" & r
                ws.Cells(r, 15).Formula = "=M" & r
            End If

            ws.Cells(r, 1).NumberFormat = "m/d/yyyy"
            ws.Cells(r, 8).NumberFormat = "0"
            ws.Cells(r, 9).NumberFormat = "0.000"
            ws.Range(ws.Cells(r, 10), ws.Cells(r, 13)).NumberFormat = "0"
            ws.Cells(r, 15).NumberFormat = "0"
            ws.Cells(r, 16).NumberFormat = "0.00"
            ws.Range(ws.Cells(r, 10), ws.Cells(r, 13)).HorizontalAlignment = xlCenter
            ws.Cells(r, 16).HorizontalAlignment = xlCenter

            If isPartial Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, LAST_COL)).Interior.Color = PARTIAL_FILL
                ws.Cells(r, 12).ClearComments
                ws.Cells(r, 12).AddComment "Partial log: measured over " & Format(availHrs, "0.00") & _
                    " h only." & vbLf & "Measured charge: " & Format(charge, "#,##0") & " C" & vbLf & _
                    "Value shown is extrapolated to 6 h, holding the mean of the last " & _
                    TAIL_POINTS & " samples constant." & vbLf & _
                    "Increase factor (col I) uses that same mean."
            End If

            If nPts > 0 Then
                AddCurrentChart ws, r, sample, estQ, tArr, aArr
            End If

            summary = summary & IIf(nRuns > 1, vbCrLf, "") & _
                   "Run " & nRuns & ": row " & r & vbCrLf & _
                   IIf(isPartial, "*** PARTIAL: only " & Format(availHrs, "0.00") & " h of data ***" & vbCrLf, "") & _
                   "Test date:        " & Format(dTest, "m/d/yyyy") & vbCrLf & _
                   "Initial current:  " & Format(hA * 1000#, "0") & " mA  (at " & H_OFFSET_SEC & " s)" & vbCrLf & _
                   "Increase factor:  " & Format(i6A / hA, "0.000") & vbCrLf & _
                   "Integral (" & IIf(isPartial, Format(availHrs, "0.00"), "6") & " h):   " & Format(charge, "#,##0") & " C" & vbCrLf & _
                   IIf(isPartial, "Est. 6 h integral: " & Format(estQ, "#,##0") & " C" & vbCrLf, "") & _
                   "Classification:   " & cls & vbCrLf & _
                   "Voltage duration: " & Format(onHours, "0.00") & " hr" & vbCrLf & _
                   "Samples used:     " & Format(nUsed, "#,##0") & vbCrLf

        End If
        If iFrom <= iPrev Then Exit Do               ' safety: always move on
    Loop

    Application.ScreenUpdating = True
    MsgBox "Sample " & sample & ": " & nRuns & IIf(nRuns = 1, " run", " runs") & " added." & _
           vbCrLf & vbCrLf & summary, vbInformation, "Done"
    Exit Sub

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' One run, from reading iFrom on: the first reading above V_THRESHOLD starts
' it, the 6 h integral and currents are taken as before, and it ends once the
' voltage has been off for OFF_GAP_SEC. iFrom returns where the next run
' search starts (n + 1 at the end of the log).
Private Function AnalyzeRun(tS() As Double, vA() As Double, aA() As Double, ByVal n As Long, _
                            ByRef iFrom As Long, ByRef dTest As Date, _
                            ByRef hA As Double, ByRef i6A As Double, _
                            ByRef charge As Double, ByRef onHours As Double, _
                            ByRef nUsed As Long, ByRef tArr() As Double, _
                            ByRef aArr() As Double, ByRef nPts As Long, _
                            ByRef isPartial As Boolean, ByRef availHrs As Double) As String
    Dim i As Long, nextFrom As Long, offIdx As Long
    Dim v As Double, a As Double, tsec As Double
    Dim t0 As Double, s As Double, prevS As Double, prevA As Double
    Dim prevT As Double, prevV As Double, onSec As Double, offStart As Double
    Dim found0 As Boolean, hFound As Boolean, endFound As Boolean, havePrev As Boolean
    Dim cap As Long

    charge = 0: onSec = 0: nUsed = 0: nPts = 0
    isPartial = False: availHrs = 0: hA = 0: i6A = 0
    cap = 1024
    ReDim tArr(1 To cap)
    ReDim aArr(1 To cap)
    nextFrom = n + 1

    For i = iFrom To n
        v = vA(i): a = aA(i): tsec = tS(i)

        ' end of this run: voltage off for OFF_GAP_SEC
        If found0 Then
            If v > V_THRESHOLD Then
                offStart = 0
            ElseIf offStart = 0 Then
                offStart = tsec: offIdx = i
            ElseIf tsec - offStart >= OFF_GAP_SEC Then
                nextFrom = offIdx + 1
                Exit For
            End If
        End If

        If havePrev Then
            If v > V_THRESHOLD And prevV > V_THRESHOLD Then onSec = onSec + (tsec - prevT)
        End If

        If Not found0 Then
            If v > V_THRESHOLD Then
                found0 = True
                t0 = tsec
                dTest = Int(tsec / 86400#)
                prevS = 0
                prevA = a
                nUsed = 1
                AddPoint tArr, aArr, nPts, cap, 0#, a
            End If
        Else
            s = tsec - t0
            If s <= TEST_SECONDS Then
                ' only voltage-on readings count, so a run that stops
                ' early isn't diluted by the switched-off tail
                If v > V_THRESHOLD Then
                    charge = charge + (prevA + a) / 2# * (s - prevS)
                    prevS = s
                    prevA = a
                    nUsed = nUsed + 1
                    AddPoint tArr, aArr, nPts, cap, s / 3600#, a
                    If Not hFound Then
                        If s >= H_OFFSET_SEC Then
                            hA = a
                            hFound = True
                        End If
                    End If
                End If
            ElseIf Not endFound Then
                ' current at 6 h; if the supply already switched off
                ' (a timed 6 h run), the last reading within 6 h
                If v > V_THRESHOLD Then i6A = a Else i6A = prevA
                endFound = True
            End If
        End If

        prevT = tsec
        prevV = v
        havePrev = True
    Next i
    iFrom = nextFrom
    onHours = onSec / 3600#

    If nPts > 0 Then
        ReDim Preserve tArr(1 To nPts)
        ReDim Preserve aArr(1 To nPts)
    End If

    If Not found0 Then
        AnalyzeRun = "No sample above " & V_THRESHOLD & " V was found - is this an RCPT log?"
    ElseIf Not hFound Then
        AnalyzeRun = "The log ends less than " & H_OFFSET_SEC & " s after voltage was applied."
    ElseIf hA <= 0 Then
        AnalyzeRun = "Initial current read as zero - cannot compute the increase factor."
    ElseIf Not endFound Then
        ' short log: extrapolated and flagged by the caller
        isPartial = True
        availHrs = prevS / 3600#
        i6A = TailMean(aArr, nPts, TAIL_POINTS)
    End If
End Function

'====================== CSV read ==============================
' Reads every row with a timestamp, a voltage and a current into tS / vA /
' aA (1..n). Columns come from the "#TimeStamp,..." header (5- and 9-column
' layouts); without one, TimeStamp,Volts,_,Amps.
Private Function ReadLog(path As String, ByRef tS() As Double, ByRef vA() As Double, _
                         ByRef aA() As Double, ByRef n As Long) As String
    Dim ff As Integer, txt As String, lines() As String, ln As String, f() As String
    Dim i As Long, nRows As Long, hdrMsg As String
    Dim tCol As Long, vCol As Long, aCol As Long, maxCol As Long

    On Error GoTo Fail
    ff = FreeFile
    Open path For Input As #ff
    txt = Input$(LOF(ff), ff)
    Close #ff

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    txt = ""

    ReDim tS(1 To UBound(lines) + 1)
    ReDim vA(1 To UBound(lines) + 1)
    ReDim aA(1 To UBound(lines) + 1)
    tCol = 0: vCol = 1: aCol = 3: maxCol = 3
    n = 0
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 20 Then
            If StrComp(Left$(ln, 10), "#TimeStamp", vbTextCompare) = 0 Then
                hdrMsg = MapColumns(ln, tCol, vCol, aCol)
                If Len(hdrMsg) > 0 Then ReadLog = hdrMsg: Exit Function
                maxCol = tCol
                If vCol > maxCol Then maxCol = vCol
                If aCol > maxCol Then maxCol = aCol
            ElseIf Left$(ln, 1) <> "#" Then
                nRows = nRows + 1
                f = Split(ln, ",")
                If UBound(f) >= maxCol Then
                    If Len(Trim$(f(vCol))) > 0 And Len(Trim$(f(aCol))) > 0 And Len(f(tCol)) >= 19 Then
                        n = n + 1
                        tS(n) = StampToSeconds(f(tCol))
                        vA(n) = Val(f(vCol))
                        aA(n) = Val(f(aCol))
                    End If
                End If
            End If
        End If
        If (i And 65535) = 0 Then
            Application.StatusBar = "Parsing ... " & Format(i, "#,##0") & " lines"
            DoEvents
        End If
    Next i
    Erase lines

    If n = 0 And nRows > 0 Then
        ReadLog = "This file has no readings - its " & Format(nRows, "#,##0") & _
                  " rows are timestamps only, so the logger stored no volts or amps."
    ElseIf n < 2 Then
        ReadLog = "No data rows found - is this a TTi measurement CSV?"
    End If
    Exit Function

Fail:
    On Error Resume Next
    Close #ff
    ReadLog = "Could not read the file:" & vbCrLf & path & vbCrLf & vbCrLf & Err.Description
End Function

' Finds the first Volts and first Amps columns in the "#TimeStamp,..." header row,
' and the TimeStamp column that belongs to that Amps column. Handles the 5-column
' layout (TimeStamp,Volts,TimeStamp,Amps) and the older 9-column one
' (TimeStamp,Volts,TimeStamp,Volts,TimeStamp,Amps,TimeStamp,Amps).
Private Function MapColumns(hdrLine As String, ByRef tCol As Long, _
                            ByRef vCol As Long, ByRef aCol As Long) As String
    Dim h() As String, k As Long

    h = Split(Mid$(hdrLine, 2), ",")
    tCol = -1: vCol = -1: aCol = -1
    For k = 0 To UBound(h)
        Select Case LCase$(Trim$(h(k)))
            Case "volts": If vCol < 0 Then vCol = k
            Case "amps":  If aCol < 0 Then aCol = k
        End Select
    Next k

    If vCol < 0 Or aCol < 0 Then
        MapColumns = "The header row has no " & IIf(vCol < 0, "Volts", "Amps") & _
                     " column - is this a TTi RCPT log?" & vbCrLf & vbCrLf & hdrLine
        Exit Function
    End If

    For k = aCol - 1 To 0 Step -1
        If LCase$(Trim$(h(k))) = "timestamp" Then tCol = k: Exit For
    Next k
    If tCol < 0 Then tCol = 0
End Function

Private Sub AddPoint(ByRef tArr() As Double, ByRef aArr() As Double, _
                     ByRef nPts As Long, ByRef cap As Long, _
                     tVal As Double, aVal As Double)
    nPts = nPts + 1
    If nPts > cap Then
        cap = cap * 2
        ReDim Preserve tArr(1 To cap)
        ReDim Preserve aArr(1 To cap)
    End If
    tArr(nPts) = tVal
    aArr(nPts) = aVal
End Sub

Private Function TailMean(aArr() As Double, nPts As Long, nTail As Long) As Double
    Dim i As Long, k As Long, s As Double
    k = nTail
    If k > nPts Then k = nPts
    For i = nPts - k + 1 To nPts
        s = s + aArr(i)
    Next i
    TailMean = s / k
End Function

Private Function StampToSeconds(t As String) As Double
    StampToSeconds = CDbl(DateSerial(CLng(Mid$(t, 1, 4)), CLng(Mid$(t, 6, 2)), CLng(Mid$(t, 9, 2)))) * 86400# _
                   + CLng(Mid$(t, 12, 2)) * 3600# _
                   + CLng(Mid$(t, 15, 2)) * 60# _
                   + Val(Mid$(t, 18))
End Function

'====================== Chart: first 6 h of current ==============
Private Sub AddCurrentChart(ws As Worksheet, r As Long, sample As String, _
                            rcptC As Double, tArr() As Double, aArr() As Double)
    Dim cht As ChartObject
    Dim anchorCell As Range
    Dim xArr() As Double, yArr() As Double
    Dim nSrc As Long, stride As Long, nOut As Long, i As Long, k As Long

    nSrc = UBound(tArr)
    stride = 1
    If nSrc > CHART_MAX_POINTS Then stride = -Int(-nSrc / CHART_MAX_POINTS)

    nOut = 0
    For i = 1 To nSrc Step stride
        nOut = nOut + 1
    Next i
    If ((nSrc - 1) Mod stride) <> 0 Then nOut = nOut + 1

    ReDim xArr(1 To nOut)
    ReDim yArr(1 To nOut)
    k = 0
    For i = 1 To nSrc Step stride
        k = k + 1
        xArr(k) = Round(tArr(i), 4)
        yArr(k) = Round(aArr(i), 6)
    Next i
    If k < nOut Then
        k = k + 1
        xArr(k) = Round(tArr(nSrc), 4)
        yArr(k) = Round(aArr(nSrc), 6)
    End If

    On Error Resume Next
    ws.ChartObjects("RCPT_" & r).Delete
    On Error GoTo 0

    Set anchorCell = ws.Cells(r, LAST_COL + 1)

    Set cht = ws.ChartObjects.Add(anchorCell.Left, anchorCell.Top, CHART_WIDTH_PT, CHART_HEIGHT_PT)
    cht.Name = "RCPT_" & r
    cht.Placement = xlMove

    With cht.Chart
        .ChartType = xlXYScatterLines
        With .SeriesCollection.NewSeries
            .XValues = xArr
            .Values = yArr
            .Name = sample
            .MarkerStyle = xlMarkerStyleNone
            .Format.Line.Weight = 1.5
        End With

        .HasTitle = True
        .ChartTitle.Text = sample & " - (" & Format(rcptC, "#,##0") & " C)"

        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Time (hours)"
            .MinimumScale = 0
            .MaximumScale = 6
        End With

        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Current (A)"
        End With

        .HasLegend = False
    End With
End Sub

'=========================== Helpers ==========================
Private Function RCPTClass(q As Double) As String
    Select Case q
        Case Is > 4000: RCPTClass = "High"
        Case Is > 2000: RCPTClass = "Moderate"
        Case Is > 1000: RCPTClass = "Low"
        Case Is > 100:  RCPTClass = "Very Low"
        Case Else:      RCPTClass = "Negligible"
    End Select
End Function

Private Function SampleFromFileName(path As String) As String
    Dim f As String, k As Long
    f = Dir(path)
    k = InStrRev(f, ".")
    If k > 0 Then f = Left$(f, k - 1)
    k = InStr(1, f, "TTiMeasurement ", vbTextCompare)
    If k > 0 Then f = Mid$(f, k + Len("TTiMeasurement "))
    k = InStr(f, " (")
    If k > 0 Then f = Left$(f, k - 1)
    SampleFromFileName = Trim$(f)
End Function

Private Function FirstEmptyRow(ws As Worksheet) As Long
    Dim r As Long
    r = 2
    Do While Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, LAST_COL))) > 0
        r = r + 1
    Loop
    FirstEmptyRow = r
End Function

Private Function PickCSVFile() As String
    Dim v As Variant
    v = Application.GetOpenFilename("CSV Files (*.csv),*.csv,All Files (*.*),*.*", 1, _
                                    "Select an RCPT measurement CSV")
    If VarType(v) = vbBoolean Then PickCSVFile = "" Else PickCSVFile = CStr(v)
End Function
