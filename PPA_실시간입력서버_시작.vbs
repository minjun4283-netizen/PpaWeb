' ==============================================================================
' PPA_실시간입력서버_시작.vbs
'
'   이 파일을 실제 PPA 엑셀 파일과 같은 폴더(=_program 폴더가 있는 바로 그
'   폴더)에 두고 더블클릭하면, 엑셀을 직접 열지 않고도 실시간 입력 서버
'   (=실시간 대시보드)가 백그라운드로 바로 시작됩니다.
'
'   내부적으로 _program\static-dashboard\run_live_server_hidden.vbs 를 그대로
'   호출합니다 - 엑셀 안의 [실시간 입력 서버 시작] 버튼과 완전히 동일한
'   방식으로 동작하며, 어떤 xlsm을 열지는 그 스크립트가 이 폴더에서 자동으로
'   찾습니다(파일이 정확히 하나여야 함).
'
'   정상적으로 시작되면 아무 창도 뜨지 않고 몇 초 뒤 브라우저만 자동으로
'   열립니다. 문제가 있을 때만(엑셀/python을 못 찾음, xlsm이 여러 개/없음 등)
'   안내 창이 뜹니다.
' ==============================================================================
Option Explicit

Dim fso, shell, myDir, target

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
myDir = fso.GetParentFolderName(WScript.ScriptFullName)
target = myDir & "\_program\static-dashboard\run_live_server_hidden.vbs"

If Not fso.FileExists(target) Then
    MsgBox "_program\static-dashboard\run_live_server_hidden.vbs 파일을 찾을 수 " & _
           "없습니다." & vbCrLf & vbCrLf & _
           "이 파일(PPA_실시간입력서버_시작.vbs)이 엑셀 파일 및 _program 폴더와 " & _
           "같은 위치에 있는지 확인해주세요(최신 버전으로 폴더를 통째로 다시 " & _
           "받아보는 것도 방법입니다).", vbExclamation, "실시간 입력 서버 시작"
    WScript.Quit 1
End If

shell.Run "wscript.exe """ & target & """", 0, False
