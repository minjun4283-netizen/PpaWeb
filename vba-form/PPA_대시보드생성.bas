Attribute VB_Name = "PPA_DashboardGen"
'==============================================================================
' PPA 대시보드 생성 / 실시간 입력 서버 버튼
'   ※ 이 파일은 ANSI(CP949)로 저장돼 있습니다. VBE [파일 가져오기]가 한국어
'      Windows 기본 코드페이지로 읽기 때문입니다. 다시 저장할 일이 있으면
'      인코딩을 반드시 'ANSI'로 두세요 (UTF-8로 저장하면 한글이 깨집니다).
'------------------------------------------------------------------------------
' 이 통합문서 옆의 static-dashboard 폴더에 있는 두 배치파일을 실행하는
' 버튼 두 개를 만듭니다.
'
'   [대시보드 생성]        : dashboard_recreate.bat 실행 (한 번만 새로 만들고
'                            끝 - 서버 없음, 조회 전용 PPA현황.html 생성)
'   [실시간 입력 서버 시작] : run_live_server.bat 실행 (계속 떠 있는 서버 -
'                            브라우저 화면에서 입력/저장하면 Windows COM으로
'                            엑셀에 바로 반영됨, pywin32 필요)
'
' 데이터 입력/수정은 여전히 PPA_InputForm(입력폼)/PPA_Explorer(탐색) 시트
' 에서도 할 수 있습니다 - 실시간 입력 서버는 그 대안(브라우저에서도 입력
' 가능)이지 대체가 아닙니다. 자세한 차이는 static-dashboard/README.md 참고.
'
' 폴더 구조 (이 순서를 벗어나면 아래 매크로가 배치파일을 못 찾습니다):
'   (작업 폴더)\
'     PPA파일.xlsm            <- 이 통합문서
'     static-dashboard\
'       dashboard_recreate.bat
'       run_live_server.bat
'       ppa_liveserver.py, excel_com.py, dashboard_form.js, ppa_*.py, vendor\...
'==============================================================================
Option Explicit

Private Const SH_NAME  As String = "대시보드생성"
Private Const BTN_W    As Double = 190
Private Const BTN_H    As Double = 32
Private Const C_INK    As Long = 2565654      ' RGB(22,38,43)
Private Const C_SUB    As Long = 7367516      ' RGB(92,107,110)
Private Const C_TEAL   As Long = 8093710      ' RGB(14,124,123)
Private Const C_LINE   As Long = 14213347     ' RGB(227,225,216)
Private Const C_FAIL   As Long = 3818163      ' RGB(179,58,58)

'---- 지금 어느 단계인지 (오류가 나면 이 값을 같이 보여줍니다) ---------------
Private g_단계 As String

Private Sub 단계(ByVal s As String)
    g_단계 = s
End Sub

Private Sub 오류보고(ByVal 어디 As String)
    MsgBox "실행 중 예상치 못한 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "위치 : " & 어디 & vbCrLf & _
           "단계 : " & g_단계 & vbCrLf & _
           "오류 " & Err.Number & " : " & Err.Description & vbCrLf & vbCrLf & _
           "이 창을 캡처해서 보내주시면 원인을 바로 확인할 수 있습니다.", _
           vbExclamation, "대시보드 생성 - 오류"
End Sub

Private Function 시트(ByVal 이름 As String) As Worksheet
    On Error Resume Next
    Set 시트 = ThisWorkbook.Worksheets(이름)
    On Error GoTo 0
End Function

'---- AutoSave로 OneDrive/SharePoint에 클라우드 경로로 열려 있으면
'     ThisWorkbook.Path가 "C:\..." 가 아니라 "https://..." 를 돌려줘서
'     Dir$() 가 오류를 냅니다. 두 버튼이 공통으로 쓰는 사전 점검입니다. ----
Private Function 로컬경로_확인() As Boolean
    If Left$(ThisWorkbook.Path, 4) = "http" Then
        MsgBox "이 통합문서가 OneDrive/SharePoint에 클라우드 경로로 열려 있어 " & _
               "로컬 폴더를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "지금 경로: " & ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "해결 방법 중 하나를 시도해주세요:" & vbCrLf & _
               "1) 파일 → 정보 에서 이 파일의 자동 저장(AutoSave)을 끄고 다시 시도" & vbCrLf & _
               "2) 즐겨찾기/최근 문서가 아니라, 탐색기의 OneDrive 동기화 폴더에서 " & _
               "이 파일을 직접 더블클릭해서 열고 다시 시도", _
               vbExclamation, "대시보드 생성"
        로컬경로_확인 = False
    Else
        로컬경로_확인 = True
    End If
End Function

'==============================================================================
' 최초 1회 실행 - 안내문 + 버튼 두 개가 있는 시트를 만듭니다
'==============================================================================
Public Sub 대시보드_만들기()
    Dim ws As Worksheet, shp As Shape, j As Long

    Application.ScreenUpdating = False
    On Error GoTo 정리

    Set ws = 시트(SH_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SH_NAME
    End If
    ws.Cells.Clear
    On Error Resume Next
    For j = ws.Shapes.Count To 1 Step -1
        ws.Shapes(j).Delete
    Next j
    On Error GoTo 0

    With ws.Cells
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
    End With
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 82
    ws.Rows(2).RowHeight = 26

    With ws.Cells(2, 2)
        .Value = "PPA 대시보드 생성"
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = C_INK
    End With
    With ws.Cells(3, 2)
        .Value = "입력/수정은 [입력폼] 또는 [탐색] 시트에서 하고, 다 저장한 뒤 " & _
                 "여기서 버튼을 누르면 최신 내용으로 PPA현황.html 이 다시 만들어집니다."
        .Font.Color = C_SUB
        .Font.Size = 10.5
    End With
    ws.Rows(3).RowHeight = 18

    Call 버튼만들기(ws, shp, "BTN_DASHGEN", 5, "대시보드_생성", "대시보드 생성")

    With ws.Cells(7, 2)
        .Value = "상태: 아직 생성한 적 없음"
        .Font.Color = C_SUB
        .Font.Size = 10
    End With

    With ws.Cells(9, 2)
        .Value = "찾는 파일: " & 배치파일경로()
        .Font.Color = C_SUB
        .Font.Size = 9
    End With

    With ws.Cells(11, 2)
        .Value = "실시간 입력 서버 - 웹 화면(브라우저)에서 입력/저장하면 지금 열려 있는 " & _
                 "엑셀에 Windows COM으로 바로 반영됩니다 (pywin32 필요, static-dashboard/README.md 참고)."
        .Font.Color = C_SUB
        .Font.Size = 10.5
    End With
    ws.Rows(11).RowHeight = 32

    Call 버튼만들기(ws, shp, "BTN_LIVESERVER", 13, "웹서버_시작", "실시간 입력 서버 시작")

    With ws.Cells(15, 2)
        .Value = "상태: 아직 시작한 적 없음"
        .Font.Color = C_SUB
        .Font.Size = 10
    End With

    With ws.Cells(17, 2)
        .Value = "찾는 파일: " & 웹서버경로()
        .Font.Color = C_SUB
        .Font.Size = 9
    End With

    ws.Cells(1, 1).Select
    ActiveWindow.DisplayGridlines = False

정리:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "시트를 만드는 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Sub 버튼만들기(ByVal ws As Worksheet, ByRef shp As Shape, ByVal shapeName As String, _
                        ByVal row As Long, ByVal onAction As String, ByVal caption As String)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(row, 2).Left, ws.Cells(row, 2).Top, BTN_W, BTN_H)
    shp.Name = shapeName
    shp.OnAction = onAction
    On Error Resume Next
    With shp
        .Fill.Solid
        .Fill.ForeColor.RGB = C_TEAL
        .Line.Visible = msoFalse
        .Placement = xlMove
        .Adjustments.Item(1) = 0.18
    End With
    With shp.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Name = "맑은 고딕"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    If Len(shp.TextFrame2.TextRange.Text) = 0 Then shp.TextFrame.Characters.Text = caption
    On Error GoTo 0
    ws.Rows(row).RowHeight = BTN_H + 6
End Sub

Private Function 배치파일경로() As String
    배치파일경로 = ThisWorkbook.Path & "\static-dashboard\dashboard_recreate.bat"
End Function

Private Function 웹서버경로() As String
    웹서버경로 = ThisWorkbook.Path & "\static-dashboard\run_live_server.bat"
End Function

Private Sub 상태쓰기(ByVal s As String, Optional ByVal 실패 As Boolean = False, Optional ByVal row As Long = 7)
    Dim ws As Worksheet, c As Range
    Set ws = 시트(SH_NAME)
    If ws Is Nothing Then Exit Sub
    Set c = ws.Cells(row, 2)
    c.Value = "상태: " & s
    c.Font.Color = IIf(실패, C_FAIL, C_SUB)
    DoEvents
End Sub

'==============================================================================
' [대시보드 생성] 버튼에 연결된 동작 - 한 번 만들고 끝(서버 없음)
'==============================================================================
Public Sub 대시보드_생성()
    Dim batPath As String, xlsmPath As String, cmd As String
    Dim ret As Long, ans As VbMsgBoxResult, found As String

    g_단계 = ""
    On Error GoTo 처리안됨

    단계 "경로 확인"
    If Not 로컬경로_확인() Then Exit Sub

    batPath = 배치파일경로()

    단계 "배치파일 확인"
    found = Dir$(batPath)
    If Len(found) = 0 Then
        MsgBox "dashboard_recreate.bat 를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "찾은 경로: " & batPath & vbCrLf & vbCrLf & _
               "이 통합문서 옆에 static-dashboard 폴더가 있고 그 안에 " & _
               "dashboard_recreate.bat 가 있는지 확인해주세요.", vbExclamation, "대시보드 생성"
        Exit Sub
    End If

    If Not ThisWorkbook.Saved Then
        ans = MsgBox("저장하지 않은 변경사항이 있습니다." & vbCrLf & _
                      "대시보드는 디스크에 저장된 내용을 기준으로 만들어집니다." & vbCrLf & vbCrLf & _
                      "지금 저장할까요?", vbYesNoCancel + vbQuestion, "대시보드 생성")
        If ans = vbCancel Then Exit Sub
        If ans = vbYes Then
            단계 "저장"
            ThisWorkbook.Save
        End If
    End If

    xlsmPath = ThisWorkbook.FullName
    cmd = """" & batPath & """ """ & xlsmPath & """"

    상태쓰기 "생성 중... (완료될 때까지 이 창이 유지됩니다)", False, 7

    단계 "배치파일 실행"
    ret = CreateObject("WScript.Shell").Run(cmd, 1, True)

    If ret = 0 Then
        상태쓰기 "마지막 생성 " & Format$(Now, "yyyy-mm-dd hh:nn") & " - 성공", False, 7
        MsgBox "대시보드를 다시 만들었습니다." & vbCrLf & _
               "PPA현황.html 이 자동으로 열렸어야 합니다(안 열렸으면 " & _
               "static-dashboard 폴더에서 직접 열어주세요).", vbInformation, "대시보드 생성"
    Else
        상태쓰기 "마지막 시도 " & Format$(Now, "yyyy-mm-dd hh:nn") & " - 실패(코드 " & ret & ")", True, 7
        MsgBox "대시보드 생성이 실패했습니다(종료 코드 " & ret & ")." & vbCrLf & vbCrLf & _
               "static-dashboard 폴더에서 dashboard_recreate.bat 를 직접 더블클릭해서 " & _
               "실행하면 자세한 오류 메시지를 볼 수 있습니다.", vbExclamation, "대시보드 생성"
    End If
    Exit Sub

처리안됨:
    상태쓰기 "오류 - " & Err.Description, True, 7
    오류보고 "대시보드_생성"
End Sub

'==============================================================================
' [실시간 입력 서버 시작] 버튼에 연결된 동작 - 계속 떠 있는 서버를 띄웁니다
'   (완료를 기다리지 않고 바로 돌아옵니다 - 서버는 검은 콘솔 창에서 계속 실행)
'==============================================================================
Public Sub 웹서버_시작()
    Dim batPath As String, xlsmPath As String, cmd As String
    Dim found As String

    g_단계 = ""
    On Error GoTo 처리안됨

    단계 "경로 확인"
    If Not 로컬경로_확인() Then Exit Sub

    batPath = 웹서버경로()

    단계 "배치파일 확인"
    found = Dir$(batPath)
    If Len(found) = 0 Then
        MsgBox "run_live_server.bat 를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "찾은 경로: " & batPath & vbCrLf & vbCrLf & _
               "이 통합문서 옆에 static-dashboard 폴더가 있고 그 안에 " & _
               "run_live_server.bat 가 있는지 확인해주세요.", vbExclamation, "실시간 입력 서버"
        Exit Sub
    End If

    ' 이 서버는 지금 화면(메모리)의 상태를 COM으로 직접 읽으므로, 대시보드 생성과
    ' 달리 저장을 먼저 하라고 요구하지 않습니다.
    xlsmPath = ThisWorkbook.FullName
    cmd = """" & batPath & """ """ & xlsmPath & """"

    단계 "서버 실행 요청"
    CreateObject("WScript.Shell").Run cmd, 1, False

    상태쓰기 "시작 요청됨 " & Format$(Now, "yyyy-mm-dd hh:nn") & " (성공 여부는 새로 뜬 검은 창에서 확인)", False, 15

    MsgBox "실시간 입력 서버를 시작했습니다." & vbCrLf & vbCrLf & _
           "몇 초 안에 브라우저가 자동으로 열립니다(안 열리면 " & _
           "http://127.0.0.1:8842 를 직접 열어주세요)." & vbCrLf & vbCrLf & _
           "새로 뜬 검은 콘솔 창이 서버입니다 - 그 창을 닫으면 서버도 함께 " & _
           "종료됩니다. pywin32가 없으면 그 창에 안내가 뜨고 바로 멈춥니다.", _
           vbInformation, "실시간 입력 서버"
    Exit Sub

처리안됨:
    상태쓰기 "오류 - " & Err.Description, True, 15
    오류보고 "웹서버_시작"
End Sub
