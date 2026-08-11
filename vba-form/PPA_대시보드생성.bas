Attribute VB_Name = "PPA_DashboardGen"
'==============================================================================
' PPA 대시보드 생성 버튼
'   ※ 이 파일은 ANSI(CP949)로 저장돼 있습니다. VBE [파일 가져오기]가 한국어
'      Windows 기본 코드페이지로 읽기 때문입니다. 다시 저장할 일이 있으면
'      인코딩을 반드시 'ANSI'로 두세요 (UTF-8로 저장하면 한글이 깨집니다).
'------------------------------------------------------------------------------
' 이 통합문서 옆의 static-dashboard 폴더에 있는 dashboard_recreate.bat 를
' 실행해서 조회 전용 HTML 대시보드(PPA현황.html)를 다시 만듭니다.
'
'   데이터 입력/수정   : PPA_InputForm(입력폼) 또는 PPA_Explorer(탐색) 시트
'   대시보드 보기용 생성 : 이 모듈의 [대시보드 생성] 버튼
'
' 값을 저장하는 기능은 여기 없습니다 - 입력폼/탐색 매크로가 이미 그 역할을
' 하고 있고, 이 모듈은 "지금 엑셀에 저장된 내용을 HTML로 다시 뽑아내는"
' 다리 역할만 합니다. HTML 대시보드는 여전히 조회 전용입니다.
'
' 폴더 구조 (이 순서를 벗어나면 아래 매크로가 배치파일을 못 찾습니다):
'   (작업 폴더)\
'     PPA파일.xlsm            <- 이 통합문서
'     static-dashboard\
'       dashboard_recreate.bat
'       build_dashboard.py, ppa_*.py, vendor\...
'==============================================================================
Option Explicit

Private Const SH_NAME  As String = "대시보드생성"
Private Const BTN_W    As Double = 150
Private Const BTN_H    As Double = 32
Private Const C_INK    As Long = 2565654      ' RGB(22,38,43)
Private Const C_SUB    As Long = 7367516      ' RGB(92,107,110)
Private Const C_TEAL   As Long = 8093710      ' RGB(14,124,123)
Private Const C_LINE   As Long = 14213347     ' RGB(227,225,216)
Private Const C_FAIL   As Long = 3818163      ' RGB(179,58,58)

Private Function 시트(ByVal 이름 As String) As Worksheet
    On Error Resume Next
    Set 시트 = ThisWorkbook.Worksheets(이름)
    On Error GoTo 0
End Function

'==============================================================================
' 최초 1회 실행 - 안내문 + 버튼이 있는 시트를 만듭니다
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

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(5, 2).Left, ws.Cells(5, 2).Top, BTN_W, BTN_H)
    shp.Name = "BTN_DASHGEN"
    shp.OnAction = "대시보드_생성"
    On Error Resume Next
    With shp
        .Fill.Solid
        .Fill.ForeColor.RGB = C_TEAL
        .Line.Visible = msoFalse
        .Placement = xlMove
        .Adjustments.Item(1) = 0.18
    End With
    With shp.TextFrame2
        .TextRange.Text = "대시보드 생성"
        .TextRange.Font.Name = "맑은 고딕"
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    If Len(shp.TextFrame2.TextRange.Text) = 0 Then shp.TextFrame.Characters.Text = "대시보드 생성"
    On Error GoTo 0

    ws.Rows(5).RowHeight = BTN_H + 6

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

    ws.Cells(1, 1).Select
    ActiveWindow.DisplayGridlines = False

정리:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "시트를 만드는 중 오류가 발생했습니다: " & Err.Description, vbExclamation
End Sub

Private Function 배치파일경로() As String
    배치파일경로 = ThisWorkbook.Path & "\static-dashboard\dashboard_recreate.bat"
End Function

Private Sub 상태쓰기(ByVal s As String, Optional ByVal 실패 As Boolean = False)
    Dim ws As Worksheet, c As Range
    Set ws = 시트(SH_NAME)
    If ws Is Nothing Then Exit Sub
    Set c = ws.Cells(7, 2)
    c.Value = "상태: " & s
    c.Font.Color = IIf(실패, C_FAIL, C_SUB)
    DoEvents
End Sub

'==============================================================================
' 버튼에 연결된 실제 동작
'==============================================================================
Public Sub 대시보드_생성()
    Dim batPath As String, xlsmPath As String, cmd As String
    Dim ret As Long, ans As VbMsgBoxResult

    batPath = 배치파일경로()
    If Len(Dir$(batPath)) = 0 Then
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
            On Error GoTo 저장실패
            ThisWorkbook.Save
            On Error GoTo 0
        End If
    End If

    xlsmPath = ThisWorkbook.FullName
    cmd = """" & batPath & """ """ & xlsmPath & """"

    상태쓰기 "생성 중... (완료될 때까지 이 창이 유지됩니다)"

    On Error GoTo 실행실패
    ret = CreateObject("WScript.Shell").Run(cmd, 1, True)
    On Error GoTo 0

    If ret = 0 Then
        상태쓰기 "마지막 생성 " & Format$(Now, "yyyy-mm-dd hh:nn") & " - 성공"
        MsgBox "대시보드를 다시 만들었습니다." & vbCrLf & _
               "PPA현황.html 이 자동으로 열렸어야 합니다(안 열렸으면 " & _
               "static-dashboard 폴더에서 직접 열어주세요).", vbInformation, "대시보드 생성"
    Else
        상태쓰기 "마지막 시도 " & Format$(Now, "yyyy-mm-dd hh:nn") & " - 실패(코드 " & ret & ")", True
        MsgBox "대시보드 생성이 실패했습니다(종료 코드 " & ret & ")." & vbCrLf & vbCrLf & _
               "static-dashboard 폴더에서 dashboard_recreate.bat 를 직접 더블클릭해서 " & _
               "실행하면 자세한 오류 메시지를 볼 수 있습니다.", vbExclamation, "대시보드 생성"
    End If
    Exit Sub

저장실패:
    MsgBox "저장 중 오류가 발생했습니다: " & Err.Description & vbCrLf & _
           "저장 후 다시 눌러주세요.", vbExclamation, "대시보드 생성"
    Exit Sub

실행실패:
    상태쓰기 "실행 실패 - " & Err.Description, True
    MsgBox "배치파일을 실행하지 못했습니다: " & Err.Description & vbCrLf & vbCrLf & _
           "회사 보안 정책이 매크로의 외부 프로그램 실행을 막고 있을 수 있습니다." & vbCrLf & _
           "이 경우 static-dashboard 폴더의 dashboard_recreate.bat 를 " & _
           "직접 더블클릭해서 실행해주세요.", vbExclamation, "대시보드 생성"
End Sub
