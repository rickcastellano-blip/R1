Option Explicit

' NT BUILD 492 - read a TTi CPX400DP logger CSV and write one row per
' specimen. Columns A:M match the "Data Summary" sheet of the NT492
' Measurement workbook, so a row copies straight across; that workbook
' works out depths, Dnssm and the categories. N:P add an RCPT-equivalent
' charge, its ASTM C1202 class, and bulk resistivity from I0.

'=========================== CONFIG ===========================
Private Const SHEET_NAME    As String = "NT492"
Private Const LAST_COL      As Long = 16         ' A:M as Data Summary, then N:P
Private Const NACL_FRAC     As Double = 0.1      ' 10 % NaCl catholyte
Private Const V_ON          As Double = 5#       ' readings at/above this are "voltage on"
Private Const U_TOL         As Double = 1#       ' V; a reading within this of a level is "at" it
Private Const SETTLE_SEC    As Double = 5#       ' initial currents are read this long after switch-on
Private Const T_TOL_HRS     As Double = 0.5      ' duration tolerance for the Table 1 check
Private Const CD_N          As Double = 0.07     ' N, colour-change concentration (OPC)
Private Const CHART_WIDTH_PT   As Double = 320#
Private Const CHART_HEIGHT_PT  As Double = 170#
Private Const CHART_MAX_POINTS As Long = 1200
Private Const RCPT_SEC      As Double = 21600#   ' C1202: 6 h at 60 V ...
Private Const RCPT_V        As Double = 60#      ' ... so integrate I over 6 h * 60/U at U
Private Const RCPT_AREA     As Double = 0.9073   ' C1202 diameter correction (95.25/100 mm)^2
Private Const TAIL_POINTS   As Long = 10         ' readings averaged to extend a short run
Private Const SPEC_DIAM_MM  As Double = 100#     ' NT BUILD 492 specimen diameter
Private Const FLAG_FILL        As Long = 13431551  ' RGB(255,242,204) light yellow

' columns (Data Summary order)
Private Const C_DATE As Long = 1, C_SPEC As Long = 2, C_U As Long = 3, C_NACL As Long = 4
Private Const C_CCL As Long = 5, C_ERF As Long = 6, C_TEMP As Long = 7, C_L As Long = 8
Private Const C_T As Long = 9, C_I30 As Long = 10, C_I0 As Long = 11, C_IFIN As Long = 12
Private Const C_CHK As Long = 13, C_QEQ As Long = 14, C_QCLS As Long = 15, C_RHO As Long = 16

'==================== BUTTON: Analyze NT Build 492 ==============
Public Sub AnalyzeNTBuild492()
    Dim path As String, spec As String, msg As String
    Dim tS() As Double, vA() As Double, aA() As Double, n As Long
    Dim dTest As Date, i30 As Double, has30 As Boolean
    Dim U As Double, i0 As Double, iFin As Double, hrs As Double
    Dim iStart As Long, iEnd As Long
    Dim uExp As Double, tExp As Double, chk As String
    Dim ws As Worksheet, r As Long
    Dim qEq As Double, qHrs As Double, qShort As Boolean

    path = PickCSVFile()
    If Len(path) = 0 Then Exit Sub

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.StatusBar = "Reading " & Dir(path) & " ..."

    msg = ReadLog(path, tS, vA, aA, n)
    If Len(msg) = 0 Then msg = FindRun(tS, vA, aA, n, iStart, iEnd, U, i30, has30, i0, iFin)
    If Len(msg) > 0 Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox msg, vbExclamation, "NT Build 492"
        Exit Sub
    End If

    hrs = (tS(iEnd) - tS(iStart)) / 3600#
    qEq = RCPTEquivalent(tS, aA, iStart, iEnd, U, qHrs, qShort)
    dTest = Int(tS(iStart) / 86400#)
    spec = SpecimenFromFileName(path)

    ' Table 1 is keyed on the 30 V current; without a 30 V reading use I0
    Table1 IIf(has30, i30, i0) * 1000#, uExp, tExp
    If Abs(U - uExp) <= U_TOL And Abs(hrs - tExp) <= T_TOL_HRS Then
        chk = "OK"
    Else
        chk = "Table 1: " & uExp & " V, " & tExp & " h"
    End If

    Set ws = GetSheet()
    r = FirstEmptyRow(ws)

    ws.Cells(r, C_DATE).Value = dTest
    ws.Cells(r, C_SPEC).Value = spec
    ws.Cells(r, C_NACL).Value = NACL_FRAC
    If has30 Then ws.Cells(r, C_I30).Value = Round(i30 * 1000#, 0) Else ws.Cells(r, C_I30).ClearContents
    ws.Cells(r, C_U).Value = Round(U, 1)
    ws.Cells(r, C_I0).Value = Round(i0 * 1000#, 0)
    ws.Cells(r, C_IFIN).Value = Round(iFin * 1000#, 0)
    ws.Cells(r, C_T).Value = Round(hrs, 2)
    ws.Cells(r, C_CHK).Value = chk
    ws.Cells(r, C_QEQ).Value = Round(qEq, 0)
    WriteFormulas ws, r
    FormatRow ws, r, (chk <> "OK")
    If qShort Then
        ws.Cells(r, C_QEQ).Interior.Color = FLAG_FILL
        ws.Cells(r, C_QEQ).AddComment "Run is shorter than the " & Format(qHrs, "0.0") & _
            " h needed at " & Format(U, "0") & " V." & vbLf & _
            "Extended to " & Format(qHrs, "0.0") & " h holding the mean of the last " & _
            TAIL_POINTS & " readings constant."
    End If

    AddCurrentChart ws, r, spec, U, tS, aA, iStart, iEnd

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ws.Activate
    ws.Cells(r, C_TEMP).Select      ' next: type temperature and thickness
    Exit Sub

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "NT Build 492"
End Sub

'====================== CSV read ==============================
' TimeStamp,Volts,TimeStamp,Amps,... ; "#" lines are headers.
Private Function ReadLog(path As String, ByRef tS() As Double, ByRef vA() As Double, _
                         ByRef aA() As Double, ByRef n As Long) As String
    Dim ff As Integer, txt As String, lines() As String, ln As String, f() As String
    Dim i As Long

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
    n = 0
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 20 And Left$(ln, 1) <> "#" Then
            f = Split(ln, ",")
            If UBound(f) >= 3 Then
                n = n + 1
                tS(n) = StampToSeconds(f(0))
                vA(n) = Val(f(1))
                aA(n) = Val(f(3))
            End If
        End If
        If (i And 32767) = 0 Then
            Application.StatusBar = "Parsing ... " & Format(i, "#,##0") & " lines"
            DoEvents
        End If
    Next i
    Erase lines

    If n < 2 Then ReadLog = "No data rows found - is this a TTi measurement CSV?"
    Exit Function

Fail:
    On Error Resume Next
    Close #ff
    ReadLog = "Could not read the file:" & vbCrLf & path & vbCrLf & vbCrLf & Err.Description
End Function

'====================== run detection =========================
' Main run = longest continuous stretch with V >= V_ON. U = mean voltage of
' its second half. The run starts at its first reading within U_TOL of U, so
' a 30 V check that steps straight to U without switching off is excluded,
' and ends at its last such reading, so switch-off transients are too.
' I30V comes from the first 30 V stretch at or before the run start.
Private Function FindRun(tS() As Double, vA() As Double, aA() As Double, n As Long, _
                         ByRef iStart As Long, ByRef iEnd As Long, ByRef U As Double, _
                         ByRef i30 As Double, ByRef has30 As Boolean, _
                         ByRef i0 As Double, ByRef iFin As Double) As String
    Dim i As Long, s0 As Long, best0 As Long, best1 As Long, bestDur As Double
    Dim sumV As Double, k As Long, iMid As Long

    s0 = 0
    For i = 1 To n + 1
        If i <= n Then
            If vA(i) >= V_ON Then
                If s0 = 0 Then s0 = i
                GoTo NextI
            End If
        End If
        If s0 > 0 Then
            If tS(i - 1) - tS(s0) > bestDur Then
                bestDur = tS(i - 1) - tS(s0): best0 = s0: best1 = i - 1
            End If
            s0 = 0
        End If
NextI:
    Next i
    If best0 = 0 Or bestDur <= 0 Then
        FindRun = "No stretch with voltage on (>= " & V_ON & " V) was found.": Exit Function
    End If

    iMid = (best0 + best1) \ 2
    For i = iMid To best1
        sumV = sumV + vA(i): k = k + 1
    Next i
    U = sumV / k

    iStart = best0
    Do While iStart < best1 And Abs(vA(iStart) - U) > U_TOL
        iStart = iStart + 1
    Loop
    ' likewise drop switch-off transients at the end (e.g. one 12 V / 131 mA
    ' reading logged while the supply ramps down)
    iEnd = best1
    Do While iEnd > iStart And Abs(vA(iEnd) - U) > U_TOL
        iEnd = iEnd - 1
    Loop

    ' 30 V check: first reading within U_TOL of 30 V up to the run start
    has30 = False
    For i = 1 To iStart
        If Abs(vA(i) - 30#) <= U_TOL Then
            i30 = SettledCurrent(tS, vA, aA, i, iEnd, 30#)
            has30 = True
            Exit For
        End If
    Next i

    i0 = SettledCurrent(tS, vA, aA, iStart, iEnd, U)
    iFin = aA(iEnd)
    If i0 <= 0 Then FindRun = "Initial current read as zero."
End Function

' Current SETTLE_SEC after index i0, staying within the stretch at level lvl;
' if that stretch is shorter, its last reading.
Private Function SettledCurrent(tS() As Double, vA() As Double, aA() As Double, _
                                i0 As Long, iMax As Long, lvl As Double) As Double
    Dim i As Long
    i = i0
    Do While i < iMax
        If Abs(vA(i + 1) - lvl) > U_TOL Then Exit Do
        If tS(i) - tS(i0) >= SETTLE_SEC Then Exit Do
        i = i + 1
    Loop
    SettledCurrent = aA(i)
End Function

' NT BUILD 492 Appendix 2, Table 1: initial current at 30 V (mA) -> U (V), t (h)
Private Sub Table1(mA As Double, ByRef uV As Double, ByRef tH As Double)
    Select Case mA
        Case Is < 5:    uV = 60: tH = 96
        Case Is < 10:   uV = 60: tH = 48
        Case Is < 15:   uV = 60: tH = 24
        Case Is < 20:   uV = 50: tH = 24
        Case Is < 30:   uV = 40: tH = 24
        Case Is < 40:   uV = 35: tH = 24
        Case Is < 60:   uV = 30: tH = 24
        Case Is < 90:   uV = 25: tH = 24
        Case Is < 120:  uV = 20: tH = 24
        Case Is < 180:  uV = 15: tH = 24
        Case Is < 360:  uV = 10: tH = 24
        Case Else:      uV = 10: tH = 6
    End Select
End Sub

' RCPT-equivalent charge: at U the same charge takes 60/U times longer to
' pass than at 60 V, so integrate the current over 6 h * 60/U from the run
' start (12 h at 30 V), then apply the C1202 diameter correction. A run
' shorter than that is extended with the mean of its last readings.
Private Function RCPTEquivalent(tS() As Double, aA() As Double, iStart As Long, _
                                iEnd As Long, U As Double, ByRef qHrs As Double, _
                                ByRef isShort As Boolean) As Double
    Dim tEnd As Double, q As Double, i As Long, s0 As Double, s1 As Double
    Dim aEnd As Double, k As Long, tail As Double

    tEnd = RCPT_SEC * RCPT_V / U
    qHrs = tEnd / 3600#
    isShort = False
    For i = iStart + 1 To iEnd
        s0 = tS(i - 1) - tS(iStart)
        s1 = tS(i) - tS(iStart)
        If s1 >= tEnd Then
            ' last partial interval, current interpolated at tEnd
            aEnd = aA(i - 1) + (aA(i) - aA(i - 1)) * (tEnd - s0) / (s1 - s0)
            q = q + (aA(i - 1) + aEnd) / 2# * (tEnd - s0)
            RCPTEquivalent = q * RCPT_AREA
            Exit Function
        End If
        q = q + (aA(i - 1) + aA(i)) / 2# * (s1 - s0)
    Next i

    isShort = True
    For i = iEnd To iStart Step -1
        tail = tail + aA(i): k = k + 1
        If k = TAIL_POINTS Then Exit For
    Next i
    q = q + tail / k * (tEnd - (tS(iEnd) - tS(iStart)))
    RCPTEquivalent = q * RCPT_AREA
End Function

'====================== sheet ================================
Private Function GetSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_NAME
    End If
    Set GetSheet = ws
End Function

' Every import goes on the first empty row below the header row.
Private Function FirstEmptyRow(ws As Worksheet) As Long
    Dim r As Long
    r = 2
    Do While Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, LAST_COL))) > 0
        r = r + 1
    Loop
    FirstEmptyRow = r
End Function

' c_Cl- (N) from the NaCl fraction (10 % -> 2.00, 3 % -> 0.60) and
' erf^-1(1 - 2cd/c0); Excel has no inverse erf, so
' erfinv(y) = NORM.S.INV((1+y)/2)/sqrt(2). Relative references, so they
' still work after the row is pasted into Data Summary.
' Also the C1202 class of the equivalent charge, and bulk resistivity
' rho = (U / I0) * A / L in kOhm-cm, which fills in once Thickness is typed.
Private Sub WriteFormulas(ws As Worksheet, r As Long)
    Dim q As String
    q = Adr(ws, r, C_QEQ)
    ws.Cells(r, C_QCLS).Formula = "=IF(" & q & "="""",""""," & _
        "IF(" & q & ">4000,""High"",IF(" & q & ">2000,""Moderate""," & _
        "IF(" & q & ">1000,""Low"",IF(" & q & ">100,""Very Low"",""Negligible"")))))"
    ws.Cells(r, C_RHO).Formula = "=IF(OR(" & Adr(ws, r, C_L) & "=""""," & Adr(ws, r, C_I0) & "=""""),""""," & _
        Adr(ws, r, C_U) & "/(" & Adr(ws, r, C_I0) & "/1000)*PI()*(" & Num(SPEC_DIAM_MM) & "/20)^2/(" & _
        Adr(ws, r, C_L) & "/10)/1000)"
    ws.Cells(r, C_CCL).Formula = "=" & Adr(ws, r, C_NACL) & "*20"
    ws.Cells(r, C_ERF).Formula = "=NORM.S.INV((2-2*" & Num(CD_N) & "/" & Adr(ws, r, C_CCL) & ")/2)/SQRT(2)"
End Sub

Private Function Adr(ws As Worksheet, r As Long, c As Long) As String
    Adr = ws.Cells(r, c).Address(False, False)
End Function

' Number as formula text with a "." decimal point, whatever the Windows locale
Private Function Num(x As Double) As String
    Num = Trim$(Str$(x))
End Function

Private Sub FormatRow(ws As Worksheet, r As Long, flagged As Boolean)
    ws.Cells(r, C_DATE).NumberFormat = "m/d/yyyy"
    ws.Cells(r, C_U).NumberFormat = "0"
    ws.Cells(r, C_NACL).NumberFormat = "0%"
    ws.Cells(r, C_CCL).NumberFormat = "0.00"
    ws.Cells(r, C_ERF).NumberFormat = "0.000"
    ws.Cells(r, C_TEMP).NumberFormat = "0.0"
    ws.Cells(r, C_L).NumberFormat = "0"
    ws.Cells(r, C_T).NumberFormat = "General"
    ws.Range(ws.Cells(r, C_I30), ws.Cells(r, C_IFIN)).NumberFormat = "0"
    ws.Cells(r, C_QEQ).NumberFormat = "#,##0"
    ws.Cells(r, C_RHO).NumberFormat = "0.0"
    ws.Range(ws.Cells(r, C_U), ws.Cells(r, C_RHO)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r, C_TEMP), ws.Cells(r, C_L)).Interior.Color = RGB(221, 235, 247)
    If flagged Then
        ws.Cells(r, C_CHK).Interior.Color = FLAG_FILL
    Else
        ws.Cells(r, C_CHK).Interior.ColorIndex = xlColorIndexNone
    End If
End Sub

'====================== chart: current during the run ===========
Private Sub AddCurrentChart(ws As Worksheet, r As Long, spec As String, U As Double, _
                            tS() As Double, aA() As Double, iStart As Long, iEnd As Long)
    Dim cht As ChartObject, anchorCell As Range
    Dim xArr() As Double, yArr() As Double
    Dim nSrc As Long, stride As Long, nOut As Long, i As Long, k As Long

    nSrc = iEnd - iStart + 1
    stride = 1
    If nSrc > CHART_MAX_POINTS Then stride = -Int(-nSrc / CHART_MAX_POINTS)
    nOut = (nSrc - 1) \ stride + 1
    If ((nSrc - 1) Mod stride) <> 0 Then nOut = nOut + 1

    ReDim xArr(1 To nOut)
    ReDim yArr(1 To nOut)
    k = 0
    For i = iStart To iEnd Step stride
        k = k + 1
        xArr(k) = Round((tS(i) - tS(iStart)) / 3600#, 4)
        yArr(k) = Round(aA(i) * 1000#, 1)
    Next i
    If k < nOut Then
        k = k + 1
        xArr(k) = Round((tS(iEnd) - tS(iStart)) / 3600#, 4)
        yArr(k) = Round(aA(iEnd) * 1000#, 1)
    End If

    On Error Resume Next
    ws.ChartObjects("NT492_" & r).Delete
    On Error GoTo 0

    Set anchorCell = ws.Cells(r, LAST_COL + 2)   ' one spare column after P
    Set cht = ws.ChartObjects.Add(anchorCell.Left, anchorCell.Top, CHART_WIDTH_PT, CHART_HEIGHT_PT)
    cht.Name = "NT492_" & r
    cht.Placement = xlMove

    With cht.Chart
        .ChartType = xlXYScatterLines
        With .SeriesCollection.NewSeries
            .XValues = xArr
            .Values = yArr
            .Name = spec
            .MarkerStyle = xlMarkerStyleNone
            .Format.Line.Weight = 1.5
        End With
        .HasTitle = True
        .ChartTitle.Text = spec & " - " & Format(U, "0.0") & " V"
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Time (hours)"
            .MinimumScale = 0
            .MaximumScale = -Int(-xArr(k))
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Current (mA)"
            .MinimumScale = 0
        End With
        .HasLegend = False
    End With
End Sub

'=========================== helpers ==========================
' "yyyy/mm/dd hh:mm:ss.ss" (or yyyy-mm-dd) -> seconds since 1899-12-30
Private Function StampToSeconds(t As String) As Double
    StampToSeconds = CDbl(DateSerial(CLng(Mid$(t, 1, 4)), CLng(Mid$(t, 6, 2)), CLng(Mid$(t, 9, 2)))) * 86400# _
                   + CLng(Mid$(t, 12, 2)) * 3600# _
                   + CLng(Mid$(t, 15, 2)) * 60# _
                   + Val(Mid$(t, 18))
End Function

' "20260923_114233_TTiMeasurement_11129I.CSV" -> "11129I"
Private Function SpecimenFromFileName(path As String) As String
    Dim f As String, k As Long
    f = Dir(path)
    k = InStrRev(f, ".")
    If k > 0 Then f = Left$(f, k - 1)
    k = InStr(1, f, "TTiMeasurement", vbTextCompare)
    If k > 0 Then f = Mid$(f, k + Len("TTiMeasurement"))
    Do While Left$(f, 1) = "_" Or Left$(f, 1) = " "
        f = Mid$(f, 2)
    Loop
    k = InStr(f, " (")
    If k > 0 Then f = Left$(f, k - 1)
    SpecimenFromFileName = Trim$(f)
End Function

Private Function PickCSVFile() As String
    Dim v As Variant
    v = Application.GetOpenFilename("CSV Files (*.csv),*.csv,All Files (*.*),*.*", 1, _
                                    "Select an NT Build 492 measurement CSV")
    If VarType(v) = vbBoolean Then PickCSVFile = "" Else PickCSVFile = CStr(v)
End Function
