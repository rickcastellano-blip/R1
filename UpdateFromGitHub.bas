Option Explicit

'=========================== CONFIG ===========================
Private Const GH_REPO    As String = "rickcastellano-blip/R1"
Private Const GH_BRANCH  As String = "claude/tender-ramanujan-047vzz"
' module name in this workbook = file in the repo, separated by ";"
Private Const MODULE_MAP As String = _
    "Module4=GenerateLineGraph.bas;" & _
    "Module3=GenerateStrengthBarGraph.bas"
Private Const NT492_MAP  As String = "NTBuild492=NTBuild492.bas"
'==============================================================

' Button macros: pull the latest code for every module in a map from GitHub
' and replace that module's code in this workbook (a missing module is
' created). This module is never overwritten, so keep it separate from the
' ones it updates. Needs File > Options > Trust Center > Macro Settings >
' "Trust access to the VBA project object model".
Public Sub UpdateFromGitHub()
    UpdateModules MODULE_MAP
End Sub

Public Sub UpdateNTBuild492()
    UpdateModules NT492_MAP
End Sub

Private Sub UpdateModules(moduleMap As String)
    Dim pairs As Variant, kv As Variant, i As Long
    Dim modName As String, fileName As String, code As String, msg As String
    Dim vbp As Object, report As String, nOK As Long, nAll As Long

    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    If Not vbp Is Nothing Then i = vbp.VBComponents.Count
    If vbp Is Nothing Or Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Excel is blocking access to the VBA project." & vbCrLf & vbCrLf & _
               "Turn on: File > Options > Trust Center > Trust Center Settings > " & _
               "Macro Settings > 'Trust access to the VBA project object model', " & _
               "then run Update again.", vbExclamation, "Update from GitHub"
        Exit Sub
    End If
    On Error GoTo 0

    pairs = Split(moduleMap, ";")
    For i = LBound(pairs) To UBound(pairs)
        If Len(Trim$(pairs(i))) > 0 Then
            nAll = nAll + 1
            kv = Split(pairs(i), "=")
            modName = Trim$(kv(0)): fileName = Trim$(kv(1))
            Application.StatusBar = "Downloading " & fileName & " ..."
            msg = ""
            code = DownloadFile(fileName, msg)
            If Len(msg) = 0 Then msg = ReplaceModuleCode(vbp, modName, code)
            If Len(msg) = 0 Then
                nOK = nOK + 1
                report = report & "  OK      " & modName & "  <-  " & fileName & vbCrLf
            Else
                report = report & "  FAILED  " & modName & "  <-  " & fileName & ":  " & msg & vbCrLf
            End If
        End If
    Next i
    Application.StatusBar = False

    ' silent on success; only report when a module failed
    If nOK < nAll Then
        MsgBox "Branch: " & GH_BRANCH & vbCrLf & vbCrLf & report & vbCrLf & _
               nOK & " of " & nAll & " module(s) updated.", vbExclamation, "Update from GitHub"
    End If
End Sub

' GitHub contents API with the raw media type: returns the file text and,
' unlike raw.githubusercontent.com, isn't cached for minutes after a push.
Private Function DownloadFile(fileName As String, ByRef errMsg As String) As String
    Dim http As Object, url As String

    url = "https://api.github.com/repos/" & GH_REPO & "/contents/" & fileName & _
          "?ref=" & GH_BRANCH
    On Error GoTo Fail
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "Accept", "application/vnd.github.raw"
    http.SetRequestHeader "User-Agent", "Excel-VBA-Updater"
    http.SetRequestHeader "Cache-Control", "no-cache"
    http.Send
    If http.Status <> 200 Then
        errMsg = "HTTP " & http.Status & " " & http.StatusText
        Exit Function
    End If
    DownloadFile = http.ResponseText
    ' refuse anything that isn't VBA source (e.g. an HTML error page)
    If InStr(1, DownloadFile, "Sub ", vbTextCompare) = 0 And _
       InStr(1, DownloadFile, "Function ", vbTextCompare) = 0 Then
        errMsg = "download doesn't look like VBA code"
        DownloadFile = ""
    End If
    Exit Function
Fail:
    errMsg = "download failed (" & Err.Description & ")"
End Function

' Replace all code in module modName (created if missing) with code.
Private Function ReplaceModuleCode(vbp As Object, modName As String, code As String) As String
    Dim comp As Object, cm As Object, lines As Variant, i As Long, body As String

    ' drop the Attribute lines an exported .bas can carry; they aren't valid code
    code = Replace(code, vbCrLf, vbLf)
    lines = Split(code, vbLf)
    For i = LBound(lines) To UBound(lines)
        If Left$(lines(i), 10) <> "Attribute " Then body = body & lines(i) & vbCrLf
    Next i

    On Error Resume Next
    Set comp = vbp.VBComponents(modName)
    On Error GoTo Fail
    If comp Is Nothing Then
        Set comp = vbp.VBComponents.Add(1)          ' vbext_ct_StdModule
        comp.Name = modName
    End If
    Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then
        If InStr(1, cm.Lines(1, cm.CountOfLines), "Sub UpdateFromGitHub(", vbTextCompare) > 0 Then
            ReplaceModuleCode = "that module holds this updater; won't overwrite it": Exit Function
        End If
    End If
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString body
    Exit Function
Fail:
    ReplaceModuleCode = Err.Description
End Function
