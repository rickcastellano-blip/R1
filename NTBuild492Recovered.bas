Option Explicit

' NT BUILD 492 for Test Bridge "recovered" TTi logs (e.g. "20653-2_recovered.csv",
' "Recovered_TTi 11127I.csv"): yyyy-mm-dd timestamps, 17-18 columns, most rows
' timestamp-only. This module only reads the file and names the sample; the
' analysis, row and chart come from NT492AnalyzeLog in the NTBuild492 module,
' so analysis changes arrive with the NT492 Update button.

'============= BUTTON: Analyze recovered NT Build 492 log ===========
Public Sub AnalyzeNTBuild492Recovered()
    Dim path As String, msg As String
    Dim tS() As Double, vA() As Double, aA() As Double, n As Long

    path = PickCSVFile()
    If Len(path) = 0 Then Exit Sub

    Application.StatusBar = "Reading " & Dir(path) & " ..."
    msg = ReadLog(path, tS, vA, aA, n)
    Application.StatusBar = False
    If Len(msg) > 0 Then MsgBox msg, vbExclamation, "NT Build 492 Recovered": Exit Sub

    NTBuild492.NT492AnalyzeLog tS, vA, aA, n, SpecimenFromFileName(path)
End Sub

'====================== CSV read ==============================
' Fixed layout TimeStamp,Volts,_,Amps (fields 1, 2 and 4); "#" lines and
' rows with a blank reading are skipped.
Private Function ReadLog(path As String, ByRef tS() As Double, ByRef vA() As Double, _
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
    tCol = 0: vCol = 1: aCol = 3: maxCol = 3
    n = 0
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 20 Then
            ' Test Bridge headers don't match the data (one labels field 5
            ' "Amps" while the current is in field 4), so they're ignored:
            ' recovered logs always hold TimeStamp,Volts,_,Amps
            If Left$(ln, 1) <> "#" Then
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
        If (i And 32767) = 0 Then
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

'=========================== helpers ==========================
' "yyyy/mm/dd hh:mm:ss.ss" (or yyyy-mm-dd) -> seconds since 1899-12-30
Private Function StampToSeconds(t As String) As Double
    StampToSeconds = CDbl(DateSerial(CLng(Mid$(t, 1, 4)), CLng(Mid$(t, 6, 2)), CLng(Mid$(t, 9, 2)))) * 86400# _
                   + CLng(Mid$(t, 12, 2)) * 3600# _
                   + CLng(Mid$(t, 15, 2)) * 60# _
                   + Val(Mid$(t, 18))
End Function

' "Recovered_TTi 11127I.csv" -> "11127I", "recovered_S3248.csv" -> "S3248",
' "20653-2_recovered.csv" -> "20653-2" (a TTiMeasurement name still works);
' a trailing " (2)" is dropped
Private Function SpecimenFromFileName(path As String) As String
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
    ElseIf StrComp(Right$(f, 9), "recovered", vbTextCompare) = 0 Then
        f = Left$(f, Len(f) - 9)                 ' "20653-2_recovered" -> "20653-2"
        Do While Right$(f, 1) = "_" Or Right$(f, 1) = " "
            f = Left$(f, Len(f) - 1)
        Loop
    End If
    SpecimenFromFileName = Trim$(StripSeparators(f))
End Function

Private Function StripSeparators(ByVal s As String) As String
    Do While Left$(s, 1) = "_" Or Left$(s, 1) = " "
        s = Mid$(s, 2)
    Loop
    StripSeparators = s
End Function

Private Function PickCSVFile() As String
    Dim v As Variant
    v = Application.GetOpenFilename("CSV Files (*.csv),*.csv,All Files (*.*),*.*", 1, _
                                    "Select a recovered NT Build 492 CSV")
    If VarType(v) = vbBoolean Then PickCSVFile = "" Else PickCSVFile = CStr(v)
End Function
