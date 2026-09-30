Option Explicit

' TTi CPX400DP logger CSV reading, shared by RCPTImport and NTBuild492.
' Handles the TTiMeasurement export (TimeStamp,Volts,TimeStamp,Amps,...),
' the older 9-column one, and Test Bridge "recovered" logs (18 columns,
' "yyyy-mm-dd" timestamps, most rows timestamp-only).

' Reads every row that has a timestamp, a voltage and a current into
' tS (seconds since 1899-12-30), vA (V) and aA (A), 1..n. Returns "" or an
' error message.
Public Function TTiReadLog(path As String, ByRef tS() As Double, ByRef vA() As Double, _
                           ByRef aA() As Double, ByRef n As Long) As String
    Dim ff As Integer, txt As String, lines() As String, ln As String, f() As String
    Dim i As Long, nRows As Long
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

    ' Default layout TimeStamp,Volts,TimeStamp,Amps (0-based columns);
    ' replaced by the #TimeStamp header row when it names both columns
    tCol = 0: vCol = 1: aCol = 3: maxCol = 3
    n = 0
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 20 Then
            If StrComp(Left$(ln, 10), "#TimeStamp", vbTextCompare) = 0 Then
                MapColumns ln, tCol, vCol, aCol
                maxCol = tCol
                If vCol > maxCol Then maxCol = vCol
                If aCol > maxCol Then maxCol = aCol
            ElseIf Left$(ln, 1) <> "#" Then
                nRows = nRows + 1
                f = Split(ln, ",")
                If UBound(f) >= maxCol Then
                    If Len(Trim$(f(vCol))) > 0 And Len(Trim$(f(aCol))) > 0 And Len(f(tCol)) >= 19 Then
                        n = n + 1
                        tS(n) = TTiStampToSeconds(f(tCol))
                        vA(n) = Val(f(vCol))
                        aA(n) = Val(f(aCol))
                    End If
                End If
            End If
        End If
        If (i And 32767) = 0 Then
            Application.StatusBar = "Parsing ... " & Format(i, "#,##0") & " lines"
            DoEvents
        End If
    Next i
    Erase lines

    If n = 0 And nRows > 0 Then
        TTiReadLog = "This file has no readings - its " & Format(nRows, "#,##0") & _
                     " rows are timestamps only, so the logger stored no volts or amps."
    ElseIf n < 2 Then
        TTiReadLog = "No data rows found - is this a TTi measurement CSV?"
    End If
    Exit Function

Fail:
    On Error Resume Next
    Close #ff
    TTiReadLog = "Could not read the file:" & vbCrLf & path & vbCrLf & vbCrLf & Err.Description
End Function

' First Volts and first Amps column in the "#TimeStamp,..." header, and the
' TimeStamp column before that Amps column. Test Bridge headers can label
' every channel "Volts"; with no Amps (or no Volts) the defaults are kept.
Private Sub MapColumns(hdrLine As String, ByRef tCol As Long, _
                       ByRef vCol As Long, ByRef aCol As Long)
    Dim h() As String, k As Long, v As Long, a As Long, t As Long

    h = Split(Mid$(hdrLine, 2), ",")
    v = -1: a = -1: t = 0
    For k = 0 To UBound(h)
        Select Case LCase$(Trim$(h(k)))
            Case "volts": If v < 0 Then v = k
            Case "amps":  If a < 0 Then a = k
        End Select
    Next k
    If v < 0 Or a < 0 Then Exit Sub

    For k = a - 1 To 0 Step -1
        If LCase$(Trim$(h(k))) = "timestamp" Then t = k: Exit For
    Next k
    tCol = t: vCol = v: aCol = a
End Sub

' "yyyy/mm/dd hh:mm:ss.ss" or "yyyy-mm-dd hh:mm:ss.sss" -> seconds since 1899-12-30
Public Function TTiStampToSeconds(t As String) As Double
    TTiStampToSeconds = CDbl(DateSerial(CLng(Mid$(t, 1, 4)), CLng(Mid$(t, 6, 2)), CLng(Mid$(t, 9, 2)))) * 86400# _
                      + CLng(Mid$(t, 12, 2)) * 3600# _
                      + CLng(Mid$(t, 15, 2)) * 60# _
                      + Val(Mid$(t, 18))
End Function

' Sample name from the file name:
'   "20260923_114233_TTiMeasurement 11129I.CSV" -> "11129I"
'   "Recovered_TTi 11127I.csv", "recovered_S3260.csv" -> "11127I", "S3260"
' A trailing " (2)" copy suffix is dropped.
Public Function TTiSampleName(path As String) As String
    Dim f As String, k As Long
    f = Dir(path)
    k = InStrRev(f, ".")
    If k > 0 Then f = Left$(f, k - 1)
    k = InStr(f, " (")
    If k > 0 Then f = Left$(f, k - 1)

    k = InStr(1, f, "TTiMeasurement", vbTextCompare)
    If k > 0 Then
        f = Mid$(f, k + Len("TTiMeasurement"))
    ElseIf StrComp(Left$(f, 9), "recovered", vbTextCompare) = 0 Then
        f = StripSeparators(Mid$(f, 10))
        If StrComp(Left$(f, 3), "TTi", vbTextCompare) = 0 Then f = Mid$(f, 4)
    End If
    TTiSampleName = Trim$(StripSeparators(f))
End Function

Private Function StripSeparators(s As String) As String
    Do While Left$(s, 1) = "_" Or Left$(s, 1) = " "
        s = Mid$(s, 2)
    Loop
    StripSeparators = s
End Function

Public Function TTiPickCSV(title As String) As String
    Dim v As Variant
    v = Application.GetOpenFilename("CSV Files (*.csv),*.csv,All Files (*.*),*.*", 1, title)
    If VarType(v) = vbBoolean Then TTiPickCSV = "" Else TTiPickCSV = CStr(v)
End Function
