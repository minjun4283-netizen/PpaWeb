Attribute VB_Name = "PPA_DashboardGen"
'==============================================================================
' PPA 대시보드 생성 / 실시간 입력 서버 버튼
'   ※ 이 파일은 ANSI(CP949)로 저장돼 있습니다. VBE [파일 가져오기]가 한국어
'      Windows 기본 코드페이지로 읽기 때문입니다. 다시 저장할 일이 있으면
'      인코딩을 반드시 'ANSI'로 두세요 (UTF-8로 저장하면 한글이 깨집니다).
'------------------------------------------------------------------------------
' 이 통합문서 옆의 _program\static-dashboard 폴더에 있는 두 배치파일을
' 실행하는 버튼 두 개를 만듭니다.
'
'   [대시보드 생성]        : dashboard_recreate.bat 실행 (한 번만 새로 만들고
'                            끝 - 서버 없음, 조회 전용 PPA현황.html 생성)
'   [실시간 입력 서버 시작] : run_live_server.bat 실행 (계속 떠 있는 서버 -
'                            브라우저 화면에서 입력/저장하면 Windows COM으로
'                            엑셀에 바로 반영됨, pywin32 필요)
'
' 데이터 입력/수정은 여전히 PPA_InputForm(입력폼)/PPA_Explorer(탐색) 시트
' 에서도 할 수 있습니다 - 실시간 입력 서버는 그 대안(브라우저에서도 입력
' 가능)이지 대체가 아닙니다. 자세한 차이는 _program/static-dashboard/README.md
' 참고.
'
' 폴더 구조 (이 순서를 벗어나면 아래 매크로가 배치파일을 못 찾습니다):
'   (작업 폴더)\
'     PPA파일.xlsm            <- 이 통합문서
'     PPA현황.html            <- 생성된 대시보드(최상위에 생성됨)
'     _program\
'       archive\              <- 저장/종료 시 자동 백업(섹션 참고)
'       static-dashboard\
'         dashboard_recreate.bat
'         run_live_server.bat
'         ppa_liveserver.py, excel_com.py, dashboard_form.js, ppa_*.py, vendor\...
'
' OneDrive/SharePoint 조직 정책으로 AutoSave를 강제로 켜두는 경우, 로컬 폴더
' 에서 다시 열어도 ThisWorkbook.FullName이 계속 클라우드 주소(https://...)로
' 나올 수 있습니다. 이때는 OneDrive 로컬 동기화 루트(Environ$의 OneDrive*
' 환경변수) 밑에서 같은 파일명을 자동으로 검색해 실제 로컬 경로를 찾아냅니다
' (실제_통합문서_전체경로 참고) - 파일이 아주 많은 팀 사이트면 검색에 시간이
' 걸릴 수 있습니다.
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
Private Const ARCHIVE_KEEP As Long = 100   ' 확장자별 보관할 백업 최대 개수

'---- 지금 어느 단계인지 (오류가 나면 이 값을 같이 보여줍니다) ---------------
Private g_단계 As String
Private g_임시파일번호 As Long

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

'==============================================================================
' 실제 통합문서 경로 확보
'   1) 로컬 경로면 그대로 씁니다(가장 흔한 경우, 빠름).
'   2) 클라우드 경로(https://...)면 OneDrive 로컬 동기화 루트 밑에서 같은
'      파일명을 검색해서 찾아냅니다 - AutoSave를 꺼도, 로컬 폴더에서 다시
'      열어도 조직 정책으로 계속 클라우드 경로가 나오는 경우를 위한 대비책.
'   둘 다 실패하면 빈 문자열을 돌려주고, 이 함수 안에서 안내 메시지를 띄웁니다.
'==============================================================================
Private Function 실제_통합문서_전체경로() As String
    Dim found As String

    If Left$(ThisWorkbook.FullName, 4) <> "http" Then
        실제_통합문서_전체경로 = ThisWorkbook.FullName
        Exit Function
    End If

    found = OneDrive_에서_찾기(ThisWorkbook.Name)

    If Len(found) = 0 Then
        MsgBox "이 통합문서가 OneDrive/SharePoint에 클라우드 경로로 열려 있어 " & _
               "로컬 폴더를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "지금 경로: " & ThisWorkbook.FullName & vbCrLf & vbCrLf & _
               "이 컴퓨터의 OneDrive 동기화 폴더에서 같은 이름의 파일을 자동으로 " & _
               "찾아봤지만 찾지 못했습니다. 아래를 확인해주세요:" & vbCrLf & _
               "1) 파일 → 정보 에서 자동 저장(AutoSave)을 끄고 다시 시도" & vbCrLf & _
               "2) OneDrive 앱에서 이 파일이 있는 폴더까지 동기화가 끝났는지 확인" & vbCrLf & _
               "3) 그래도 안 되면, 조직 정책으로 이 SharePoint 문서함이 항상 " & _
               "클라우드 모드로 열리도록 강제돼 있을 수 있습니다 - IT 담당자에게 " & _
               "문의하거나, 이 파일을 OneDrive가 아닌 팀이 접근 가능한 다른 " & _
               "공유 위치로 옮기는 방법을 검토해주세요.", _
               vbExclamation, "대시보드 생성"
        실제_통합문서_전체경로 = ""
        Exit Function
    End If

    실제_통합문서_전체경로 = found
End Function

Private Function 폴더경로(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        폴더경로 = Left$(fullPath, p - 1)
    Else
        폴더경로 = fullPath
    End If
End Function

'---- OneDrive 로컬 동기화 루트 목록 (Environ$ 변수 이름이 OneDrive로 시작하는 것 전부) ----
Private Function OneDrive_루트목록() As Collection
    Dim col As New Collection
    Dim i As Long, s As String, eqPos As Long, nm As String, val As String

    i = 1
    Do
        s = Environ$(i)
        If Len(s) = 0 Then Exit Do
        eqPos = InStr(s, "=")
        If eqPos > 0 Then
            nm = Left$(s, eqPos - 1)
            val = Mid$(s, eqPos + 1)
            If Len(val) > 0 And LCase$(Left$(nm, 8)) = "onedrive" Then
                On Error Resume Next
                col.Add val
                On Error GoTo 0
            End If
        End If
        i = i + 1
    Loop

    Set OneDrive_루트목록 = col
End Function

'==============================================================================
' 경로가 OneDrive 동기화 폴더(사용자 로컬 OneDrive 루트 중 하나) 아래에
' 있는지만 판정합니다. 'Files On-Demand' 로 인해 탐색기에는 파일이 보여도
' 실제로는 이 컴퓨터에 내려받아지지 않은 상태(구름 아이콘)일 수 있는데,
' 이 경우 Dir$() 로도 존재를 확인하지 못하는 경우가 있어 경로진단() 에서
' 안내 문구를 붙일지 판단하는 데만 사용합니다.
'==============================================================================
Private Function OneDrive_안인지(ByVal 경로 As String) As Boolean
    Dim roots As Collection, root As Variant
    Dim 경로소문자 As String, 루트소문자 As String

    On Error Resume Next
    경로소문자 = LCase$(경로)
    Set roots = OneDrive_루트목록()
    For Each root In roots
        루트소문자 = LCase$(CStr(root))
        If Len(루트소문자) > 0 Then
            If Left$(경로소문자, Len(루트소문자)) = 루트소문자 Then
                OneDrive_안인지 = True
                Exit Function
            End If
        End If
    Next root
    OneDrive_안인지 = False
End Function

'==============================================================================
' 배치/실행 파일을 못 찾았을 때, 단순히 '찾은 경로: ...' 한 줄만 보여주는
' 대신 작업폴더부터 상대경로를 한 단계씩 따라가며 정확히 어느 폴더/파일
' 단계에서 끊기는지 짚어주고, 흔한 원인(경로 길이 제한, OneDrive 온라인
' 전용 파일, 백신 격리, ZIP 안에서 직접 열기/예전 '최근 항목')을 함께
' 안내하는 진단 메시지를 만듭니다. 실제 Dir$() 판정 로직 자체는 바꾸지
' 않고, 실패했을 때 보여줄 메시지만 더 자세하게 만드는 용도입니다.
'==============================================================================
Private Function 경로진단(ByVal 작업폴더 As String, ByVal 상대경로 As String) As String
    Dim 조각() As String
    Dim 현재경로 As String, i As Long
    Dim 메시지 As String, 전체경로 As String

    전체경로 = 작업폴더 & 상대경로
    조각 = Split(상대경로, "\")

    현재경로 = 작업폴더
    For i = LBound(조각) To UBound(조각)
        If Len(조각(i)) > 0 Then
            현재경로 = 현재경로 & "\" & 조각(i)
            If i < UBound(조각) Then
                If Len(Dir$(현재경로, vbDirectory)) = 0 Then
                    메시지 = "다음 폴더를 찾을 수 없습니다:" & vbCrLf & 현재경로
                    Exit For
                End If
            Else
                If Len(Dir$(현재경로)) = 0 Then
                    메시지 = "다음 파일을 찾을 수 없습니다:" & vbCrLf & 현재경로
                    Exit For
                End If
            End If
        End If
    Next i

    If Len(메시지) = 0 Then
        메시지 = "경로상 모든 폴더/파일이 확인되었습니다(원인 불명):" & vbCrLf & 전체경로
    End If

    메시지 = 메시지 & vbCrLf & vbCrLf & "찾은 전체 경로: " & 전체경로

    If Len(전체경로) > 250 Then
        메시지 = 메시지 & vbCrLf & vbCrLf & _
            "참고: 경로 길이가 " & Len(전체경로) & "자로 깁니다. Windows는 260자 " & _
            "안팎의 제한이 있을 수 있어, 폴더 이름을 줄이거나 더 상위 위치로 " & _
            "옮기면 해결될 수 있습니다."
    End If

    If OneDrive_안인지(작업폴더) Then
        메시지 = 메시지 & vbCrLf & vbCrLf & _
            "참고: 이 폴더는 OneDrive 동기화 폴더 안에 있습니다. 탐색기에 파일이 " & _
            "보여도 구름 아이콘(온라인 전용)이면 실제로는 이 컴퓨터에 내려받아지지 " & _
            "않은 상태일 수 있습니다. 탐색기에서 _program 폴더를 우클릭한 뒤 " & _
            "'항상 이 디바이스에 유지'를 선택해 완전히 내려받고 나서 다시 시도해보세요."
    End If

    메시지 = 메시지 & vbCrLf & vbCrLf & _
        "그 밖의 흔한 원인: 최근에 받은 .bat/.vbs 파일을 백신 프로그램이 " & _
        "격리했을 수 있고(백신 알림/격리함을 확인해보세요), 또는 이 통합문서를 " & _
        "압축(zip) 파일 안에서 바로 열었거나 '최근 항목'에 남은 예전 위치를 " & _
        "열었을 수 있습니다 - 파일 탐색기에서 xlsm 파일이 실제로 있는 폴더로 " & _
        "이동해 그 폴더 안의 파일을 직접 더블클릭해서 열어보세요."

    경로진단 = 메시지
End Function

Private Function OneDrive_에서_찾기(ByVal fileName As String) As String
    Dim roots As Collection
    Dim root As Variant
    Dim found As String

    Set roots = OneDrive_루트목록()
    For Each root In roots
        found = 파일찾기_재귀(CStr(root), fileName)
        If Len(found) > 0 Then
            OneDrive_에서_찾기 = found
            Exit Function
        End If
    Next root

    OneDrive_에서_찾기 = ""
End Function

'---- root 밑을 재귀적으로 뒤져 fileName과 이름이 같은 첫 파일의 전체 경로를
'     돌려줍니다. VBA의 Dir$은 재귀 검색을 못 해서, cmd의 dir /s /b 결과를
'     임시 파일로 받아 읽는 방식을 씁니다. ------------------------------------
Private Function 파일찾기_재귀(ByVal root As String, ByVal fileName As String) As String
    Dim tmp As String, cmd As String, result As String
    Dim f As Integer

    If Len(Dir$(root, vbDirectory)) = 0 Then
        파일찾기_재귀 = ""
        Exit Function
    End If

    g_임시파일번호 = g_임시파일번호 + 1
    tmp = Environ$("TEMP") & "\ppa_find_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(g_임시파일번호) & ".txt"

    cmd = "cmd /c dir /s /b """ & root & "\" & fileName & """ > """ & tmp & """ 2>nul"
    CreateObject("WScript.Shell").Run cmd, 0, True

    result = ""
    If Len(Dir$(tmp)) > 0 Then
        f = FreeFile
        Open tmp For Input As #f
        If Not EOF(f) Then Line Input #f, result
        Close #f
        On Error Resume Next
        Kill tmp
        On Error GoTo 0
    End If

    파일찾기_재귀 = result
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
        .Value = "통합문서 경로: " & ThisWorkbook.FullName
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
        .Value = "버튼을 누르면 static-dashboard 폴더를 자동으로 찾습니다" & _
                 "(클라우드 경로로 열려 있으면 OneDrive에서 자동 검색)."
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
    Dim batPath As String, xlsmPath As String, cmd As String, 작업폴더 As String
    Dim ret As Long, ans As VbMsgBoxResult, found As String

    g_단계 = ""
    On Error GoTo 처리안됨

    단계 "경로 확인"
    If Left$(ThisWorkbook.FullName, 4) = "http" Then
        상태쓰기 "클라우드 경로 감지됨 - 로컬 폴더 검색 중... (파일이 많으면 시간이 걸릴 수 있습니다)", False, 7
    End If
    xlsmPath = 실제_통합문서_전체경로()
    If Len(xlsmPath) = 0 Then
        상태쓰기 "경로를 찾지 못해 중단됨", True, 7
        Exit Sub
    End If

    작업폴더 = 폴더경로(xlsmPath)
    batPath = 작업폴더 & "\_program\static-dashboard\dashboard_recreate.bat"

    단계 "배치파일 확인"
    found = Dir$(batPath)
    If Len(found) = 0 Then
        MsgBox "dashboard_recreate.bat 를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               경로진단(작업폴더, "\_program\static-dashboard\dashboard_recreate.bat"), _
               vbExclamation, "대시보드 생성"
        상태쓰기 "배치파일을 찾지 못해 중단됨", True, 7
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

    cmd = """" & batPath & """ """ & xlsmPath & """"

    상태쓰기 "생성 중... (완료될 때까지 이 창이 유지됩니다)", False, 7

    단계 "배치파일 실행"
    ret = CreateObject("WScript.Shell").Run(cmd, 1, True)

    If ret = 0 Then
        상태쓰기 "마지막 생성 " & Format$(Now, "yyyy-mm-dd hh:nn") & " - 성공", False, 7
        단계 "백업 저장"
        아카이브_백업 작업폴더, xlsmPath, 작업폴더 & "\PPA현황.html"
        MsgBox "대시보드를 다시 만들었습니다." & vbCrLf & _
               "PPA현황.html 이 자동으로 열렸어야 합니다(안 열렸으면 " & _
               "작업 폴더에서 직접 열어주세요).", vbInformation, "대시보드 생성"
    Else
        상태쓰기 "마지막 시도 " & Format$(Now, "yyyy-mm-dd hh:nn") & " - 실패(코드 " & ret & ")", True, 7
        MsgBox "대시보드 생성이 실패했습니다(종료 코드 " & ret & ")." & vbCrLf & vbCrLf & _
               "_program\static-dashboard 폴더에서 dashboard_recreate.bat 를 직접 " & _
               "더블클릭해서 " & _
               "실행하면 자세한 오류 메시지를 볼 수 있습니다.", vbExclamation, "대시보드 생성"
    End If
    Exit Sub

처리안됨:
    상태쓰기 "오류 - " & Err.Description, True, 7
    오류보고 "대시보드_생성"
End Sub

'==============================================================================
' 저장/생성 시점마다 xlsm + html을 archive 폴더에 타임스탬프로 복사해두는
' 간단한 백업입니다. 개수가 늘어나기만 하면 디스크가 계속 차오르므로,
' 매번 복사한 뒤 확장자별로 오래된 것부터 정리해 ARCHIVE_KEEP개만 남깁니다.
' 오류가 나도(디스크 꽉 참 등) 대시보드 생성 자체는 실패시키지 않도록
' On Error Resume Next로 조용히 넘어갑니다 - 백업은 있으면 좋은 안전장치일
' 뿐, 이것 때문에 정작 원래 하려던 대시보드 생성이 막히면 안 됩니다.
'==============================================================================
Private Sub 아카이브_백업(ByVal 작업폴더 As String, ByVal xlsmPath As String, ByVal htmlPath As String)
    Dim archiveDir As String, ts As String
    On Error Resume Next
    archiveDir = 작업폴더 & "\_program\archive"
    If Len(Dir$(archiveDir, vbDirectory)) = 0 Then MkDir archiveDir
    ts = Format$(Now, "yyyymmdd_hhnnss")
    If Len(Dir$(xlsmPath)) > 0 Then FileCopy xlsmPath, archiveDir & "\PPA파일_" & ts & ".xlsm"
    If Len(Dir$(htmlPath)) > 0 Then FileCopy htmlPath, archiveDir & "\PPA현황_" & ts & ".html"
    아카이브_정리 archiveDir, "*.xlsm"
    아카이브_정리 archiveDir, "*.html"
    On Error GoTo 0
End Sub

' archiveDir 안에서 패턴에 맞는 파일 중 ARCHIVE_KEEP개를 넘는 만큼, 이름이
' 가장 앞서는(=타임스탬프가 파일명 안에 있어 이름순 = 시간순) 것부터 지웁니다.
Private Sub 아카이브_정리(ByVal archiveDir As String, ByVal 패턴 As String)
    Dim col As Collection, f As String, i As Long, j As Long
    Dim arr() As String, n As Long, tmp As String
    Set col = New Collection
    f = Dir$(archiveDir & "\" & 패턴)
    Do While Len(f) > 0
        col.Add f
        f = Dir$()
    Loop
    n = col.Count
    If n <= ARCHIVE_KEEP Then Exit Sub
    ReDim arr(n - 1)
    For i = 1 To n
        arr(i - 1) = col(i)
    Next i
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(j) > arr(j + 1) Then
                tmp = arr(j): arr(j) = arr(j + 1): arr(j + 1) = tmp
            End If
        Next j
    Next i
    For i = 0 To n - ARCHIVE_KEEP - 1
        Kill archiveDir & "\" & arr(i)
    Next i
End Sub

'==============================================================================
' [실시간 입력 서버 시작] 버튼에 연결된 동작 - 계속 떠 있는 서버를 띄웁니다
'   (완료를 기다리지 않고 바로 돌아옵니다). run_live_server_hidden.vbs 를
'   통해 콘솔 창 없이 백그라운드로 띄웁니다 - 정상적으로 시작되면 브라우저가
'   자동으로 열리는 것 외에는 아무 창도 뜨지 않습니다. 문제가 있을 때만(엑셀/
'   python을 못 찾음, 서버가 응답하지 않음) 그 vbs가 스스로 안내 창을 띄웁니다
'   - 이 매크로는 그 결과를 기다리지 않으므로(Wait:=False) 엑셀 화면이
'   멈추지 않습니다.
'==============================================================================
Public Sub 웹서버_시작()
    Dim vbsPath As String, batPath As String, xlsmPath As String, cmd As String
    Dim 작업폴더 As String, found As String

    g_단계 = ""
    On Error GoTo 처리안됨

    단계 "경로 확인"
    If Left$(ThisWorkbook.FullName, 4) = "http" Then
        상태쓰기 "클라우드 경로 감지됨 - 로컬 폴더 검색 중... (파일이 많으면 시간이 걸릴 수 있습니다)", False, 15
    End If
    xlsmPath = 실제_통합문서_전체경로()
    If Len(xlsmPath) = 0 Then
        상태쓰기 "경로를 찾지 못해 중단됨", True, 15
        Exit Sub
    End If

    작업폴더 = 폴더경로(xlsmPath)
    vbsPath = 작업폴더 & "\_program\static-dashboard\run_live_server_hidden.vbs"
    batPath = 작업폴더 & "\_program\static-dashboard\run_live_server.bat"

    단계 "실행 파일 확인"
    found = Dir$(vbsPath)
    If Len(found) > 0 Then
        ' 콘솔 창 없이 백그라운드로 - 정상 작동 시 아무 창도 뜨지 않습니다.
        cmd = "wscript.exe """ & vbsPath & """ """ & xlsmPath & """"
        단계 "서버 실행 요청(숨김)"
        CreateObject("WScript.Shell").Run cmd, 0, False
        상태쓰기 "시작 요청됨 " & Format$(Now, "yyyy-mm-dd hh:nn") & _
                 " - 몇 초 안에 브라우저가 자동으로 열립니다(문제가 있으면 알림 창이 뜹니다)", False, 15
        Exit Sub
    End If

    ' run_live_server_hidden.vbs 가 없는 예전 static-dashboard 폴더 대비 -
    ' 예전처럼 검은 콘솔 창이 뜨는 방식으로 대체합니다.
    found = Dir$(batPath)
    If Len(found) = 0 Then
        MsgBox "run_live_server_hidden.vbs 도 run_live_server.bat 도 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               경로진단(작업폴더, "\_program\static-dashboard\run_live_server.bat") & vbCrLf & vbCrLf & _
               "(최신 버전으로 폴더를 통째로 다시 받아보는 것도 방법입니다.)", _
               vbExclamation, "실시간 입력 서버"
        상태쓰기 "실행 파일을 찾지 못해 중단됨", True, 15
        Exit Sub
    End If

    cmd = """" & batPath & """ """ & xlsmPath & """"
    단계 "서버 실행 요청(콘솔 표시)"
    CreateObject("WScript.Shell").Run cmd, 1, False
    상태쓰기 "시작 요청됨 " & Format$(Now, "yyyy-mm-dd hh:nn") & " (구버전 방식 - 검은 콘솔 창에서 확인)", False, 15
    MsgBox "실시간 입력 서버를 시작했습니다(구버전 방식 - 콘솔 창이 뜹니다)." & vbCrLf & vbCrLf & _
           "_program\static-dashboard 폴더를 최신 버전으로 갱신하면 콘솔 창 없이 시작할 수 있습니다.", _
           vbInformation, "실시간 입력 서버"
    Exit Sub

처리안됨:
    상태쓰기 "오류 - " & Err.Description, True, 15
    오류보고 "웹서버_시작"
End Sub
