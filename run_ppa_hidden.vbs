Option Explicit

Dim shell, fso, base, batPath, checkUrl, openUrl
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

base = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = base & "\run_ppa_hidden.bat"
checkUrl = "http://127.0.0.1:8765/api/schema"
openUrl = "http://127.0.0.1:8765"

If Not fso.FileExists(batPath) Then
    MsgBox "BAT file not found:" & vbCrLf & batPath, 16, "PPA Error"
    WScript.Quit 1
End If

shell.Run Chr(34) & batPath & Chr(34), 0, False

If WaitForServer(checkUrl, 60000, 500) Then
    shell.Run openUrl, 1, False
Else
    shell.Run openUrl, 1, False
End If

Function WaitForServer(targetUrl, timeoutMs, intervalMs)
    Dim http, startedAt, elapsedMs

    startedAt = Timer

    Do
        On Error Resume Next
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", targetUrl, False
        http.Send

        If Err.Number = 0 Then
            If http.Status = 200 Then
                WaitForServer = True
                Exit Function
            End If
        End If

        Err.Clear
        On Error GoTo 0

        WScript.Sleep intervalMs
        elapsedMs = (Timer - startedAt) * 1000
        If elapsedMs >= timeoutMs Then Exit Do
    Loop

    WaitForServer = False
End Function