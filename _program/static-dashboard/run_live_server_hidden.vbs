' ==============================================================================
' run_live_server_hidden.vbs - PPA 실시간 입력 서버를 콘솔 창 없이 백그라운드로
'   띄웁니다.
'
'   run_live_server.bat 과 하는 일은 같습니다(xlsm 자동 탐색, python 위치
'   확인, ppa_liveserver.py 실행) - 다른 점은 그 모든 과정과 서버 프로세스
'   자체를 화면에 아무 창도 띄우지 않고 진행한다는 것입니다. 정상적으로
'   시작되면(브라우저가 자동으로 열리는 것 말고는) 아무것도 뜨지 않고 조용히
'   끝나고, 서버는 백그라운드에서 계속 실행됩니다. 문제가 있을 때만(xlsm/
'   python을 못 찾음, 서버가 확인 시간 안에 응답하지 않음) 안내 창이 뜹니다.
'
'   콘솔 출력이 없으므로 서버의 표준출력/오류는 live_server.log 파일로
'   저장됩니다(_run_server_hidden_worker.bat 을 통해 리다이렉트) - 실패 시 이
'   파일의 마지막 부분을 안내 창에 그대로 보여줍니다.
'
'   사용법: wscript.exe run_live_server_hidden.vbs ["엑셀파일경로.xlsm"]
'           (인자를 생략하면 상위 폴더에서 xlsm을 자동으로 하나 찾습니다)
' ==============================================================================
Option Explicit

Dim fso, shell, scriptDir, engineDir, topDir, xlsmPath, pyExe, logPath, workerBat, cmd
Dim ok, startTime, title
Const READY_TIMEOUT_SEC = 30

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
' engineDir = _program (static-dashboard/python-embed가 나란히 있는 폴더)
' topDir    = 작업폴더 (xlsm, 결과 html이 있는 최상위)
engineDir = fso.GetParentFolderName(scriptDir)
topDir = fso.GetParentFolderName(engineDir)
title = "실시간 입력 서버"

' ---- 1) xlsm 경로 ----------------------------------------------------------
If WScript.Arguments.Count >= 1 Then
    xlsmPath = WScript.Arguments(0)
Else
    xlsmPath = FindSingleXlsm(topDir)
    If xlsmPath = "" Then
        MsgBox "작업폴더(" & topDir & ")에서 xlsm 파일을 정확히 하나 찾지 " & _
               "못했습니다." & vbCrLf & vbCrLf & _
               "파일이 없거나 여러 개 있을 수 있습니다 - 직접 경로를 지정해서 " & _
               "실행해주세요.", vbExclamation, title
        WScript.Quit 1
    End If
End If

If Not fso.FileExists(xlsmPath) Then
    MsgBox "엑셀 파일을 찾을 수 없습니다:" & vbCrLf & xlsmPath, vbExclamation, title
    WScript.Quit 1
End If

' ---- 2) python 위치 (PATH 우선, 없으면 동봉된 python-embed) ----------------
pyExe = FindPython(engineDir)
If pyExe = "" Then
    MsgBox "python을 찾을 수 없습니다." & vbCrLf & vbCrLf & _
           "PATH에 python이 있거나, 이 폴더 바로 위에 python-embed 폴더가 " & _
           "있어야 합니다. static-dashboard\README.md 의 Python 설치 안내를 " & _
           "참고해주세요.", vbExclamation, title
    WScript.Quit 1
End If

workerBat = scriptDir & "\_run_server_hidden_worker.bat"
If Not fso.FileExists(workerBat) Then
    MsgBox "내부 실행 파일을 찾을 수 없습니다:" & vbCrLf & workerBat & vbCrLf & vbCrLf & _
           "static-dashboard 폴더가 통째로 옮겨졌는지 확인해주세요.", vbExclamation, title
    WScript.Quit 1
End If

' ---- 3) 서버를 완전히 숨겨서 실행 ------------------------------------------
' 실제 리다이렉트(> 로그파일 2>&1)는 _run_server_hidden_worker.bat 안에서
' 처리합니다 - 여기서 cmd /c 안에 따옴표를 겹겹이 넣는 것보다, 이미 검증된
' 배치파일 인자 전달 방식(공백 포함 경로도 "%~1"로 안전하게 받음)을 그대로
' 재사용하는 편이 훨씬 안전합니다.
logPath = scriptDir & "\live_server.log"
On Error Resume Next
If fso.FileExists(logPath) Then fso.DeleteFile logPath, True
On Error Goto 0

cmd = """" & workerBat & """ """ & pyExe & """ """ & scriptDir & "\ppa_liveserver.py"" """ & _
      xlsmPath & """ """ & logPath & """"
shell.Run cmd, 0, False

' ---- 4) 준비될 때까지 폴링 - 서버가 응답하면 성공으로 판단 ----------------
ok = False
startTime = Timer
Do While Timer - startTime < READY_TIMEOUT_SEC
    WScript.Sleep 700
    If IsServerReady() Then
        ok = True
        Exit Do
    End If
Loop

If Not ok Then
    MsgBox "실시간 입력 서버가 " & READY_TIMEOUT_SEC & "초 안에 준비 확인이 " & _
           "안 됐습니다." & vbCrLf & _
           "(브라우저 탭이 이미 열려 있다면 거기서 좀 더 기다려도 됩니다 - 엑셀 " & _
           "파일이 커서 여는 데 시간이 걸리는 것뿐일 수도 있습니다. 이 확인 " & _
           "시간을 넘겼다고 해서 반드시 실패한 것은 아닙니다.)" & vbCrLf & vbCrLf & _
           "최근 로그(" & logPath & "):" & vbCrLf & ReadTail(logPath, 3500), _
           vbExclamation, title
    WScript.Quit 1
End If

' 성공 - 아무 창도 띄우지 않고 조용히 끝냅니다(브라우저는 ppa_liveserver.py
' 자신이 엽니다). 서버는 백그라운드에서 계속 실행됩니다.
WScript.Quit 0

' ==============================================================================
' 헬퍼 함수
' ==============================================================================

Function FindSingleXlsm(baseDir)
    Dim folder, f, match, cnt
    match = ""
    cnt = 0
    On Error Resume Next
    Set folder = fso.GetFolder(baseDir)
    If Err.Number <> 0 Then
        FindSingleXlsm = ""
        Exit Function
    End If
    On Error Goto 0
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "xlsm" Then
            match = f.Path
            cnt = cnt + 1
        End If
    Next
    If cnt = 1 Then
        FindSingleXlsm = match
    Else
        FindSingleXlsm = ""
    End If
End Function

Function FindPython(baseDir)
    Dim ret, embedPath

    ' PATH에 python이 있는지 - 창 없이 확인(종료 코드만 봄, 따옴표 없는
    ' 인자라 cmd /c 안에 넣어도 이스케이프 문제가 없습니다)
    ret = shell.Run("cmd /c where python >nul 2>nul", 0, True)
    If ret = 0 Then
        FindPython = "python"
        Exit Function
    End If

    embedPath = baseDir & "\python-embed\python.exe"
    If fso.FileExists(embedPath) Then
        FindPython = embedPath
        Exit Function
    End If

    FindPython = ""
End Function

Function IsServerReady()
    Dim xhr
    IsServerReady = False
    On Error Resume Next
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", "http://127.0.0.1:8842/api/ready", False
    xhr.Send
    If Err.Number = 0 And xhr.Status = 200 Then
        IsServerReady = (InStr(xhr.ResponseText, """ready"": true") > 0 Or _
                          InStr(xhr.ResponseText, """ready"":true") > 0)
    End If
    Err.Clear
    On Error Goto 0
End Function

Function ReadTail(path, maxChars)
    Dim ts, allText
    On Error Resume Next
    If Not fso.FileExists(path) Then
        ReadTail = "(로그 파일이 아직 없습니다 - python/서버 실행 자체가 안 됐을 수 있습니다)"
        Exit Function
    End If
    Set ts = fso.OpenTextFile(path, 1, False)
    allText = ts.ReadAll
    ts.Close
    If Err.Number <> 0 Then
        ReadTail = "(로그 파일을 읽지 못했습니다: " & Err.Description & ")"
        Exit Function
    End If
    On Error Goto 0
    If Len(allText) > maxChars Then
        ReadTail = "..." & Right(allText, maxChars)
    Else
        ReadTail = allText
    End If
End Function
