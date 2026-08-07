Attribute VB_Name = "PPA_Explorer"
'==============================================================================
' PPA 관계형 데이터 탐색 · 편집 폼
'   ※ 이 파일은 ANSI(CP949)로 저장돼 있습니다. VBE [파일 가져오기]가 한국어
'      Windows 기본 코드페이지로 읽기 때문입니다. 다시 저장할 일이 있으면
'      인코딩을 반드시 'ANSI'로 두세요 (UTF-8로 저장하면 한글이 깨집니다).
'------------------------------------------------------------------------------
' HTML 대시보드의 [탐색] 탭을 엑셀 안으로 옮긴 것입니다. 다른 점 하나 -
' 결과를 보기만 하는 게 아니라 그 자리에서 고쳐 원본 시트에 반영합니다.
'
'   ① 설정 반영 : 기준 표를 고르고 연결할 표를 체크한 뒤 누릅니다
'   ② 조회 실행 : 여러 표가 합쳐진 목록을 만듭니다
'   수정 저장    : 결과 칸에서 고친 값을 원본에 반영합니다
'
' 속도 설계
'   시트를 셀 단위로 오가면 수만 번의 왕복이 생겨 느립니다. 그래서
'     · 각 표를 배열로 한 번만 읽어 캐시하고 (표배열)
'     · PK/FK 를 Dictionary 색인으로 만들어 조인에서 전체 스캔을 없애고
'     · 행 대응은 문자열이 아니라 Long 배열로 다루고
'     · 결과는 2차원 배열을 한 번에 붙여넣고
'     · 줄무늬/누락 표시는 셀마다 칠하지 않고 조건부 서식 규칙으로 처리합니다.
'
' 편집 안전장치
'   · PK 열은 수정 대상에서 제외 (ID를 바꾸면 FK가 전부 깨지므로)
'   · FK 열은 드롭다운으로만 고르게 하고, 저장 시 실제 존재 여부를 다시 검사
'   · 1:N 전개로 같은 원본 행이 여러 줄에 보일 때 서로 다른 값으로 고치면 중단
'   · 저장 전에 무엇이 어떻게 바뀌는지 보여주고 확인
'   · 고친 칸은 저장 전까지 빨간 굵은 글씨로 표시 (더티 마킹)
'
' 이 모듈은 단독으로 동작합니다 (PPA_InputForm 이 없어도 됩니다).
' 이름(Name) 은 전부 PPAX_ 접두어를 써서 다른 모듈의 정의와 충돌하지 않습니다.
'==============================================================================
Option Explicit

'---- 시트 -------------------------------------------------------------------
Private Const EXP_SHEET As String = "탐색"
Private Const MAP_SHEET As String = "_탐색맵"
Private Const NAME_PREFIX As String = "PPAX_"

Private Const SH_PLANT  As String = "T_발전소"
Private Const SH_BUY    As String = "T_구매계약"
Private Const SH_DEMAND As String = "T_수요기업"
Private Const SH_SELL   As String = "T_판매계약"
Private Const SH_SITE   As String = "T_전기사용지"
Private Const SH_MATCH  As String = "T_수급매칭"

'---- 레이아웃 (전부 고정 - 체크박스가 늘 같은 칸에 오도록) -----------------
Private Const COL_KEY As Long = 1
Private Const COL_LBL As Long = 2
Private Const COL_VAL As Long = 3

Private Const SLOT_W As Double = 19      ' 값/결과 열 너비 (한 번 정하고 바꾸지 않음)
Private Const SLOTS  As Long = 44        ' 균일 폭을 적용할 열 개수

Private Const ROW_BTN     As Long = 1    ' 버튼 (틀 고정으로 항상 보임)
Private Const ROW_TITLE   As Long = 2
Private Const ROW_PRESET  As Long = 4
Private Const ROW_BASE    As Long = 5
Private Const ROW_JOIN    As Long = 7
Private Const ROW_COLHEAD As Long = 9
Private Const ROW_COL1    As Long = 10   ' 10~15 : 표별 컬럼 체크박스 6줄
Private Const ROW_COND    As Long = 17
Private Const ROW_STATUS  As Long = 18
Private Const ROW_HEAD    As Long = 20
Private Const ROW_DATA    As Long = 21

Private Const BTN_W As Double = 118
Private Const BTN_H As Double = 27
Private Const BTN_GAP As Double = 6
Private Const BTN_GROUPGAP As Double = 24

Private Const MAX_ROWS As Long = 20000
Private Const MISS_MARK As String = "-"
Private Const FK_LIST_COL As Long = 40   ' 맵시트에서 FK 목록을 쌓는 시작 열

Private Const F_ALL  As String = "전체 (연결 여부 무관)"
Private Const F_ANY  As String = "한 곳이라도 빠진 행만 (누락)"
Private Const F_NONE As String = "전부 연결된 행만"
Private Const P_NONE As String = "(직접 설정)"

'---- 색 (한 곳에서 관리) ----------------------------------------------------
Private Const C_INK    As Long = 2565654      ' RGB(22,38,43)
Private Const C_SUB    As Long = 7367516      ' RGB(92,107,110)
Private Const C_TEAL   As Long = 8093710      ' RGB(14,124,123)
Private Const C_TEALD  As Long = 5856266      ' RGB(10,90,89)
Private Const C_PURPLE As Long = 12012627     ' RGB(83,74,183)
Private Const C_LINE   As Long = 14213347     ' RGB(227,225,216)

'---- 캐시 (작업 한 번 동안만 유지) ------------------------------------------
Private mHdr As Object      ' "시트|헤더" -> 열번호
Private mArrD As Object     ' 시트명      -> 2차원 배열(1행=머리글)
Private mPkD As Object      ' 시트명      -> Dictionary(PK값 -> 행)
Private mFkD As Object      ' "시트|컬럼" -> Dictionary(값 -> Collection(행))
Private mTypD As Object     ' "시트|헤더" -> "TEXT"/"BOOL"/"DATE"/"NUM"

'==============================================================================
' 1. 스키마
'==============================================================================
Private Function 표_컬럼목록(ByVal 시트명 As String) As Variant
    Select Case 시트명
    Case SH_PLANT
        표_컬럼목록 = Array("발전소ID", "발전소명", "발전법인명", "설비용량(MW)", _
                            "발전원", "Readiness", "MGA_Supply")
    Case SH_BUY
        표_컬럼목록 = Array("구매계약ID", "발전소ID", "구매계약용량(MW)", "구매단가(원/kWh)", _
                            "공급기한_구매", "계약기간(년)", "수요기업 미확보", "구매 담당자")
    Case SH_DEMAND
        표_컬럼목록 = Array("수요기업ID", "기업명")
    Case SH_SELL
        표_컬럼목록 = Array("판매계약ID", "수요기업ID", "판매계약용량(MW)", "계약일", _
                            "공급기한_판매", "계약유형", "판매단가(원/kWh)", "공급자원 미확보", _
                            "판매 담당자", "계약기간(년)", "Requirement", "MGA_Demand")
    Case SH_SITE
        표_컬럼목록 = Array("전기사용지ID", "판매계약ID", "전기사용지명", "전기사용지계약용량(MW)")
    Case SH_MATCH
        표_컬럼목록 = Array("수급매칭ID", "전기사용지ID", "구매계약ID", "현황")
    Case Else
        표_컬럼목록 = Array()
    End Select
End Function

Private Function 표_PK(ByVal 시트명 As String) As String
    Dim c As Variant
    c = 표_컬럼목록(시트명)
    If UBound(c) < 0 Then Exit Function
    표_PK = CStr(c(0))
End Function

Private Function 표_순서() As Variant
    표_순서 = Array(SH_PLANT, SH_BUY, SH_DEMAND, SH_SELL, SH_SITE, SH_MATCH)
End Function

Private Function 라벨(ByVal 시트명 As String) As String
    라벨 = Replace(시트명, "T_", "")
End Function

Private Function 시트명찾기(ByVal 라벨값 As String) As String
    Dim t As Variant
    For Each t In 표_순서()
        If 라벨(CStr(t)) = 라벨값 Then
            시트명찾기 = CStr(t)
            Exit Function
        End If
    Next t
End Function

Private Function FK_참조(ByVal 시트명 As String, ByVal 컬럼 As String) As String
    Select Case 시트명 & "|" & 컬럼
    Case SH_BUY & "|발전소ID":       FK_참조 = SH_PLANT
    Case SH_SELL & "|수요기업ID":    FK_참조 = SH_DEMAND
    Case SH_SITE & "|판매계약ID":    FK_참조 = SH_SELL
    Case SH_MATCH & "|전기사용지ID": FK_참조 = SH_SITE
    Case SH_MATCH & "|구매계약ID":   FK_참조 = SH_BUY
    Case Else:                       FK_참조 = ""
    End Select
End Function

' "상대시트|방향|컬럼" (parent = 이 표가 상대를 참조 / child = 상대가 이 표를 참조)
Private Function 이웃(ByVal 시트명 As String) As Collection
    Dim c As New Collection, t As Variant, cols As Variant, i As Long, 참조 As String
    cols = 표_컬럼목록(시트명)
    For i = LBound(cols) To UBound(cols)
        참조 = FK_참조(시트명, CStr(cols(i)))
        If Len(참조) > 0 Then c.Add 참조 & "|parent|" & cols(i)
    Next i
    For Each t In 표_순서()
        cols = 표_컬럼목록(CStr(t))
        For i = LBound(cols) To UBound(cols)
            If FK_참조(CStr(t), CStr(cols(i))) = 시트명 Then c.Add CStr(t) & "|child|" & cols(i)
        Next i
    Next t
    Set 이웃 = c
End Function

'==============================================================================
' 2. 캐시 · 배열 입출력
'==============================================================================
Private Sub 캐시비우기()
    Set mHdr = CreateObject("Scripting.Dictionary")
    Set mArrD = CreateObject("Scripting.Dictionary")
    Set mPkD = CreateObject("Scripting.Dictionary")
    Set mFkD = CreateObject("Scripting.Dictionary")
    Set mTypD = CreateObject("Scripting.Dictionary")
End Sub

Private Function 시트(ByVal 이름 As String) As Worksheet
    On Error Resume Next
    Set 시트 = ThisWorkbook.Worksheets(이름)
    On Error GoTo 0
End Function

' 표 전체를 배열로 한 번만 읽어 캐시 (1행 = 머리글)
Private Function 표배열(ByVal 시트명 As String) As Variant
    Dim ws As Worksheet, 끝행 As Long, 끝열 As Long, pk열 As Long
    If mArrD Is Nothing Then 캐시비우기
    If mArrD.Exists(시트명) Then
        표배열 = mArrD(시트명)
        Exit Function
    End If

    Set ws = 시트(시트명)
    If ws Is Nothing Then Exit Function
    끝열 = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If 끝열 < 2 Then 끝열 = 2
    pk열 = 머리열스캔(ws, 표_PK(시트명), 끝열)
    If pk열 = 0 Then pk열 = 1
    끝행 = ws.Cells(ws.Rows.Count, pk열).End(xlUp).Row
    If 끝행 < 2 Then 끝행 = 2

    표배열 = ws.Range(ws.Cells(1, 1), ws.Cells(끝행, 끝열)).Value
    mArrD(시트명) = 표배열
End Function

Private Function 머리열스캔(ByVal ws As Worksheet, ByVal 헤더 As String, _
                            ByVal 끝열 As Long) As Long
    Dim j As Long
    For j = 1 To 끝열
        If Trim$(CStr(ws.Cells(1, j).Value)) = 헤더 Then
            머리열스캔 = j
            Exit Function
        End If
    Next j
End Function

Private Function 헤더열(ByVal 시트명 As String, ByVal 헤더 As String) As Long
    Dim k As String, a As Variant, j As Long
    If mHdr Is Nothing Then 캐시비우기
    k = 시트명 & "|" & 헤더
    If mHdr.Exists(k) Then
        헤더열 = mHdr(k)
        Exit Function
    End If
    a = 표배열(시트명)
    If IsEmpty(a) Then Exit Function
    For j = 1 To UBound(a, 2)
        If Trim$(CStr(a(1, j) & "")) = 헤더 Then
            헤더열 = j
            Exit For
        End If
    Next j
    mHdr(k) = 헤더열
End Function

Private Function 값문자열(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        값문자열 = ""
    ElseIf VarType(v) = vbBoolean Then
        값문자열 = IIf(v, "TRUE", "FALSE")
    ElseIf IsDate(v) Then
        값문자열 = Format$(v, "yyyy-mm-dd")
    Else
        값문자열 = Trim$(CStr(v))
    End If
End Function

Private Function 셀값(ByVal 시트명 As String, ByVal r As Long, ByVal 헤더 As String) As String
    Dim a As Variant, j As Long
    If r < 2 Then Exit Function
    j = 헤더열(시트명, 헤더)
    If j = 0 Then Exit Function
    a = 표배열(시트명)
    If IsEmpty(a) Then Exit Function
    If r > UBound(a, 1) Then Exit Function
    셀값 = 값문자열(a(r, j))
End Function

Private Function PK색인(ByVal 시트명 As String) As Object
    Dim d As Object, a As Variant, pk As Long, r As Long, v As String
    If mPkD Is Nothing Then 캐시비우기
    If mPkD.Exists(시트명) Then
        Set PK색인 = mPkD(시트명)
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    pk = 헤더열(시트명, 표_PK(시트명))
    a = 표배열(시트명)
    If pk > 0 And Not IsEmpty(a) Then
        For r = 2 To UBound(a, 1)
            v = 값문자열(a(r, pk))
            If Len(v) > 0 Then
                If Not d.Exists(v) Then d(v) = r
            End If
        Next r
    End If
    Set mPkD(시트명) = d
    Set PK색인 = d
End Function

Private Function FK색인(ByVal 시트명 As String, ByVal 컬럼 As String) As Object
    Dim d As Object, a As Variant, j As Long, r As Long, v As String
    Dim c As Collection, k As String
    If mFkD Is Nothing Then 캐시비우기
    k = 시트명 & "|" & 컬럼
    If mFkD.Exists(k) Then
        Set FK색인 = mFkD(k)
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    j = 헤더열(시트명, 컬럼)
    a = 표배열(시트명)
    If j > 0 And Not IsEmpty(a) Then
        For r = 2 To UBound(a, 1)
            v = 값문자열(a(r, j))
            If Len(v) > 0 Then
                If d.Exists(v) Then
                    Set c = d(v)
                Else
                    Set c = New Collection
                    Set d(v) = c
                End If
                c.Add r
            End If
        Next r
    End If
    Set mFkD(k) = d
    Set FK색인 = d
End Function

Private Function 열형식(ByVal 시트명 As String, ByVal 헤더 As String) As String
    Dim k As String, a As Variant, j As Long, r As Long
    If mTypD Is Nothing Then 캐시비우기
    k = 시트명 & "|" & 헤더
    If mTypD.Exists(k) Then
        열형식 = mTypD(k)
        Exit Function
    End If

    If 헤더 = 표_PK(시트명) Or Len(FK_참조(시트명, 헤더)) > 0 Then
        열형식 = "TEXT"                       ' ID 는 "001" 이 1 로 바뀌면 안 됨
    Else
        j = 헤더열(시트명, 헤더)
        a = 표배열(시트명)
        열형식 = "TEXT"
        If j > 0 And Not IsEmpty(a) Then
            For r = 2 To UBound(a, 1)
                If Not IsEmpty(a(r, j)) Then
                    If VarType(a(r, j)) = vbBoolean Then
                        열형식 = "BOOL"
                    ElseIf IsDate(a(r, j)) Then
                        열형식 = "DATE"
                    ElseIf IsNumeric(a(r, j)) Then
                        열형식 = "NUM"
                    End If
                    Exit For
                End If
            Next r
        End If
    End If
    mTypD(k) = 열형식
End Function

Private Sub 셀쓰기(ByVal 시트명 As String, ByVal r As Long, ByVal 헤더 As String, _
                   ByVal v As String)
    Dim ws As Worksheet, j As Long, t As String
    Set ws = 시트(시트명)
    If ws Is Nothing Then Exit Sub
    j = 헤더열(시트명, 헤더)
    If j = 0 Then Exit Sub

    If Len(v) = 0 Then
        ws.Cells(r, j).ClearContents
        Exit Sub
    End If

    t = 열형식(시트명, 헤더)
    Select Case t
    Case "TEXT"
        ws.Cells(r, j).NumberFormatLocal = "@"
        ws.Cells(r, j).Value = v
    Case "BOOL"
        ws.Cells(r, j).Value = (UCase$(v) = "TRUE" Or v = "예" Or UCase$(v) = "Y")
    Case "DATE"
        If IsDate(v) Then
            ws.Cells(r, j).Value = CDate(v)
            ws.Cells(r, j).NumberFormatLocal = "yyyy-mm-dd"
        Else
            ws.Cells(r, j).Value = v
        End If
    Case "NUM"
        If IsNumeric(v) Then ws.Cells(r, j).Value = CDbl(v) Else ws.Cells(r, j).Value = v
    Case Else
        ws.Cells(r, j).Value = v
    End Select
End Sub

'---- 화면 갱신 잠시 끄기 ----------------------------------------------------
Private Function 빠르게시작() As Variant
    Dim s(0 To 3) As Variant
    s(0) = Application.ScreenUpdating
    s(1) = Application.EnableEvents
    s(2) = xlCalculationAutomatic
    On Error Resume Next
    s(2) = Application.Calculation
    On Error GoTo 0
    s(3) = Application.DisplayStatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    On Error GoTo 0
    Application.DisplayStatusBar = True
    빠르게시작 = s
End Function

Private Sub 빠르게끝(ByVal s As Variant)
    On Error Resume Next
    Application.Calculation = s(2)
    On Error GoTo 0
    Application.ScreenUpdating = s(0)
    Application.EnableEvents = s(1)
    Application.StatusBar = False
    Application.DisplayStatusBar = s(3)
End Sub

'==============================================================================
' 3. 화면 만들기
'==============================================================================
Public Sub 탐색_만들기()
    Dim ws As Worksheet, s As Variant

    캐시비우기
    If Not 원본표_점검() Then Exit Sub

    s = 빠르게시작()
    Application.DisplayAlerts = False
    On Error GoTo 정리

    Set ws = 시트(EXP_SHEET)
    If Not ws Is Nothing Then ws.Delete
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = EXP_SHEET

    With ws.Cells
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    ws.Columns(COL_KEY).ColumnWidth = 0.1
    ws.Columns(COL_LBL).ColumnWidth = 19
    ' 값/체크박스/결과 열은 모두 같은 너비 - 이후 절대 바꾸지 않습니다.
    ' (폭이 바뀌면 체크박스와 격자가 어긋나기 때문)
    ws.Range(ws.Columns(COL_VAL), ws.Columns(COL_VAL + SLOTS - 1)).ColumnWidth = SLOT_W

    ws.Rows(ROW_BTN).RowHeight = 38
    ws.Rows(ROW_TITLE).RowHeight = 26
    ws.Rows(ROW_JOIN).RowHeight = 22
    ws.Range(ws.Rows(ROW_COL1), ws.Rows(ROW_COL1 + 5)).RowHeight = 22
    ws.Rows(ROW_PRESET).RowHeight = 20
    ws.Rows(ROW_BASE).RowHeight = 20
    ws.Rows(ROW_COND).RowHeight = 20

    With ws.Cells(ROW_TITLE, COL_LBL)
        .Value = "관계형 데이터 탐색 · 편집"
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = C_INK
    End With
    ws.Cells(ROW_TITLE, COL_VAL).Value = "① 설정 반영 → ② 조회 실행 → 흰 칸을 고치고 [수정 저장]"
    ws.Cells(ROW_TITLE, COL_VAL).Font.Color = C_SUB

    구역라벨 ws, ROW_PRESET, "빠른 조회", C_PURPLE
    구역라벨 ws, ROW_BASE, "1. 기준 표", C_TEAL
    구역라벨 ws, ROW_JOIN, "2. 연결할 표", C_TEAL
    구역라벨 ws, ROW_COLHEAD, "3. 출력 컬럼", C_TEAL
    구역라벨 ws, ROW_COND, "4. 조건", C_TEAL
    구역라벨 ws, ROW_STATUS, "상태", C_SUB

    ws.Cells(ROW_BASE, COL_VAL).Value = 라벨(SH_PLANT)
    ws.Cells(ROW_BASE, COL_VAL + 1).Value = "이 표의 각 행이 결과의 기준이 됩니다"
    ws.Cells(ROW_BASE, COL_VAL + 1).Font.Color = C_SUB
    ws.Cells(ROW_COLHEAD, COL_VAL).Value = "보고 싶은 컬럼만 체크하세요 (연결하지 않은 표는 흐리게 표시됩니다)"
    ws.Cells(ROW_COLHEAD, COL_VAL).Font.Color = C_SUB
    ws.Cells(ROW_PRESET, COL_VAL + 1).Value = "고르면 기준 표 · 연결 표 · 누락 조건이 한 번에 맞춰집니다"
    ws.Cells(ROW_PRESET, COL_VAL + 1).Font.Color = C_SUB

    ws.Cells(ROW_COND, COL_VAL).Value = F_ALL
    ws.Cells(ROW_COND, COL_VAL + 2).Value = "검색어"
    ws.Cells(ROW_COND, COL_VAL + 2).Font.Bold = True
    ws.Cells(ROW_COND, COL_VAL + 2).HorizontalAlignment = xlRight
    With ws.Cells(ROW_COND, COL_VAL + 3)
        .Interior.Color = RGB(255, 249, 219)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 196, 180)
        .NumberFormatLocal = "@"
    End With
    ws.Cells(ROW_COND, COL_VAL + 4).Value = "출력 컬럼 안에서 찾습니다 (대소문자 무시)"
    ws.Cells(ROW_COND, COL_VAL + 4).Font.Color = C_SUB
    ws.Cells(ROW_PRESET, COL_VAL).Value = P_NONE

    버튼만들기 ws
    드롭다운만들기 ws
    체크박스_다시그리기 ws
    상태쓰기 "준비됐습니다. 기준 표와 연결할 표를 고르고 [① 설정 반영] → [② 조회 실행]."

    ' 버튼 줄만 고정 - 결과를 아래로 훑어도 버튼이 항상 보입니다
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Cells(ROW_BASE, COL_VAL).Select

정리:
    Application.DisplayAlerts = True
    빠르게끝 s
    If Err.Number <> 0 Then
        MsgBox "탐색 시트를 만드는 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation
        Exit Sub
    End If

    MsgBox "탐색 시트를 만들었습니다." & vbCrLf & vbCrLf & _
           "· 기준 표나 연결할 표를 바꾸면 [① 설정 반영]을 누르세요." & vbCrLf & _
           "· 결과의 흰 칸을 고치면 빨갛게 표시되고, [수정 저장]으로 원본에 반영됩니다.", _
           vbInformation, "PPA 탐색 · 편집"
End Sub

Private Sub 구역라벨(ByVal ws As Worksheet, ByVal r As Long, ByVal 글 As String, ByVal 색 As Long)
    With ws.Cells(r, COL_LBL)
        .Value = 글
        .Font.Bold = True
        .Font.Color = 색
        .HorizontalAlignment = xlRight
    End With
End Sub

' 버튼: 둥근 셰이프 · 크기 동일 · 역할별로 묶어 배치 · 맨 윗줄 고정
Private Sub 버튼만들기(ByVal ws As Worksheet)
    Dim 정의 As Variant, i As Long, x As Double, y As Double
    Dim 그룹 As String, 이전그룹 As String, sh As Shape, j As Long

    ' 이 시트의 기존 버튼만 지웁니다 (체크박스는 건드리지 않음)
    On Error Resume Next
    For j = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(j).Name, 4) = "BTN_" Then ws.Shapes(j).Delete
    Next j
    On Error GoTo 0

    ' "그룹|캡션|매크로|색"
    정의 = Array( _
        "A|① 설정 반영|탐색_설정적용|" & CStr(C_TEAL), _
        "A|② 조회 실행|탐색_실행|" & CStr(RGB(30, 99, 168)), _
        "B|수정 저장|탐색_변경저장|" & CStr(RGB(31, 122, 84)), _
        "B|수정 취소|탐색_되돌리기|" & CStr(RGB(176, 120, 23)), _
        "C|새 파일로 내보내기|탐색_내보내기|" & CStr(C_SUB), _
        "C|결과 지우기|탐색_초기화|" & CStr(C_SUB))

    x = ws.Cells(ROW_BTN, COL_LBL).Left
    y = ws.Cells(ROW_BTN, COL_LBL).Top + 5
    이전그룹 = ""
    For i = LBound(정의) To UBound(정의)
        그룹 = Split(정의(i), "|")(0)
        If Len(이전그룹) > 0 And 그룹 <> 이전그룹 Then x = x + BTN_GROUPGAP
        셰이프버튼 ws, "BTN_" & i, x, y, Split(정의(i), "|")(1), _
                   Split(정의(i), "|")(2), CLng(Split(정의(i), "|")(3))
        x = x + BTN_W + BTN_GAP
        이전그룹 = 그룹
    Next i
End Sub

Private Sub 셰이프버튼(ByVal ws As Worksheet, ByVal 이름 As String, _
                       ByVal x As Double, ByVal y As Double, _
                       ByVal 캡션 As String, ByVal 액션 As String, ByVal 색 As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, BTN_W, BTN_H)
    shp.Name = 이름
    With shp
        .OnAction = 액션
        .Fill.ForeColor.RGB = 색
        .Fill.Solid
        .Line.Visible = msoFalse
        .Placement = xlMove              ' 행 높이만 따라가고 크기는 유지
        On Error Resume Next
        .Adjustments.Item(1) = 0.18
        On Error GoTo 0
        With .TextFrame2
            .TextRange.Text = 캡션
            .TextRange.Font.Name = "맑은 고딕"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 0
            .MarginRight = 0
        End With
    End With
End Sub

Private Sub 드롭다운만들기(ByVal ws As Worksheet)
    Dim 표목록 As String, 필터목록 As String, 프리셋 As String
    Dim t As Variant, c As Collection, i As Long

    For Each t In 표_순서()
        표목록 = 표목록 & IIf(Len(표목록) > 0, ",", "") & 라벨(CStr(t))
    Next t
    필터목록 = F_ALL & "," & F_ANY & "," & F_NONE

    Set c = 프리셋목록()
    프리셋 = P_NONE
    For i = 1 To c.Count
        프리셋 = 프리셋 & "," & Split(CStr(c(i)), "|")(2)
    Next i

    목록달기 ws.Cells(ROW_BASE, COL_VAL), 표목록
    목록달기 ws.Cells(ROW_COND, COL_VAL), 필터목록
    목록달기 ws.Cells(ROW_PRESET, COL_VAL), 프리셋
End Sub

Private Sub 목록달기(ByVal 셀 As Range, ByVal 목록 As String)
    On Error Resume Next
    With 셀
        .Interior.Color = RGB(238, 237, 254)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 196, 180)
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                 Operator:=xlBetween, Formula1:=목록
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = False
        End With
    End With
    On Error GoTo 0
End Sub

Private Function 원본표_점검() As Boolean
    Dim t As Variant, cols As Variant, i As Long, ws As Worksheet, 문제 As String
    For Each t In 표_순서()
        Set ws = 시트(CStr(t))
        If ws Is Nothing Then
            문제 = 문제 & "· 시트 없음: " & t & vbCrLf
        Else
            cols = 표_컬럼목록(CStr(t))
            For i = LBound(cols) To UBound(cols)
                If 헤더열(CStr(t), CStr(cols(i))) = 0 Then
                    문제 = 문제 & "· " & t & " 시트에 '" & cols(i) & "' 열이 없습니다" & vbCrLf
                End If
            Next i
        End If
    Next t
    If Len(문제) > 0 Then
        MsgBox "표 구성이 예상과 달라 만들 수 없습니다." & vbCrLf & vbCrLf & 문제, _
               vbExclamation, "PPA 탐색"
        원본표_점검 = False
    Else
        원본표_점검 = True
    End If
End Function

'==============================================================================
' 4. 체크박스 - 셀 격자에 정확히 맞추고 6개 표를 늘 같은 자리에
'==============================================================================
Private Function 기준표() As String
    Dim ws As Worksheet, v As String
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then
        기준표 = SH_PLANT
        Exit Function
    End If
    v = Trim$(CStr(ws.Cells(ROW_BASE, COL_VAL).Value))
    기준표 = 시트명찾기(v)
    If Len(기준표) = 0 Then 기준표 = SH_PLANT
End Function

Public Sub 탐색_설정적용()
    Dim ws As Worksheet, s As Variant
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then
        MsgBox "먼저 [탐색_만들기]를 실행해주세요.", vbInformation
        Exit Sub
    End If
    캐시비우기
    s = 빠르게시작()
    체크박스_다시그리기 ws
    빠르게끝 s
    상태쓰기 "설정을 반영했습니다. 컬럼을 확인하고 [② 조회 실행]을 누르세요."
End Sub

Private Sub 체크박스_다시그리기(ByVal ws As Worksheet)
    Dim cb As Object, t As Variant, 표이름 As String
    Dim i As Long, k As Long, r As Long, cols As Variant
    Dim 기준 As String, 거리 As Object, 켠표 As Object, 켠컬럼 As Object
    Dim 유효 As Collection, 사용중 As Object, 이표선택 As Boolean

    기준 = 기준표()

    ' 지금 켜진 상태를 기억했다가 복원
    Set 켠표 = CreateObject("Scripting.Dictionary")
    Set 켠컬럼 = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If Left$(cb.Name, 4) = "TBL^" Then
            If cb.Value = xlOn Then 켠표(Mid$(cb.Name, 5)) = True
        ElseIf Left$(cb.Name, 4) = "COL^" Then
            If cb.Value = xlOn Then 켠컬럼(Mid$(cb.Name, 5)) = True
        End If
    Next cb
    ws.CheckBoxes.Delete
    On Error GoTo 0

    Set 거리 = 최단거리(기준)

    ' --- 2. 연결할 표 : 6개를 항상 같은 칸(C~H)에 ---
    k = 0
    For Each t In 표_순서()
        표이름 = CStr(t)
        Set cb = 체크박스놓기(ws, ROW_JOIN, COL_VAL + k)
        If 표이름 = 기준 Then
            cb.Caption = 라벨(표이름) & " (기준)"
            cb.Name = "BASEMARK"
            cb.Value = xlOn
            cb.Enabled = False
        Else
            cb.Caption = 라벨(표이름) & IIf(거리.Exists(표이름), " " & 거리(표이름) & "단계", "")
            cb.Name = "TBL^" & 표이름
            If 켠표.Exists(표이름) Then cb.Value = xlOn Else cb.Value = xlOff
            cb.Enabled = 거리.Exists(표이름)
        End If
        k = k + 1
    Next t

    ' --- 3. 출력 컬럼 : 6개 표를 늘 6줄로. 연결 안 한 표도 자리를 지킵니다 ---
    Set 유효 = 유효표목록(ws)
    Set 사용중 = CreateObject("Scripting.Dictionary")
    For k = 1 To 유효.Count
        사용중(CStr(유효(k))) = True
    Next k

    r = ROW_COL1
    For Each t In 표_순서()
        표이름 = CStr(t)
        With ws.Cells(r, COL_LBL)
            .Value = 라벨(표이름)
            .Font.Bold = 사용중.Exists(표이름)
            .Font.Color = IIf(사용중.Exists(표이름), C_INK, RGB(176, 174, 166))
            .HorizontalAlignment = xlRight
        End With

        cols = 표_컬럼목록(표이름)
        이표선택 = False
        For i = LBound(cols) To UBound(cols)
            If 켠컬럼.Exists(표이름 & "^" & cols(i)) Then 이표선택 = True
        Next i

        For i = LBound(cols) To UBound(cols)
            Set cb = 체크박스놓기(ws, r, COL_VAL + i)
            cb.Caption = CStr(cols(i))
            cb.Name = "COL^" & 표이름 & "^" & cols(i)
            If Not 사용중.Exists(표이름) Then
                cb.Value = xlOff
                cb.Enabled = False
            Else
                cb.Enabled = True
                If Not 이표선택 Then
                    cb.Value = IIf(i <= 1, xlOn, xlOff)     ' 새로 붙인 표는 앞 2개
                ElseIf 켠컬럼.Exists(표이름 & "^" & cols(i)) Then
                    cb.Value = xlOn
                Else
                    cb.Value = xlOff
                End If
            End If
        Next i
        r = r + 1
    Next t
End Sub

' 셀 하나에 딱 맞춰 체크박스를 놓습니다 (격자 정렬의 핵심)
Private Function 체크박스놓기(ByVal ws As Worksheet, ByVal r As Long, _
                              ByVal c As Long) As Object
    Dim 셀 As Range, cb As Object
    Set 셀 = ws.Cells(r, c)
    Set cb = ws.CheckBoxes.Add(셀.Left + 2, 셀.Top + 2, 셀.Width - 4, 셀.Height - 4)
    cb.Placement = xlMoveAndSize
    cb.Font.Name = "맑은 고딕"
    cb.Font.Size = 9
    Set 체크박스놓기 = cb
End Function

Private Function 최단거리(ByVal 기준 As String) As Object
    Dim d As Object, q As Collection, cur As String, nb As Collection
    Dim i As Long, 상대 As String
    Set d = CreateObject("Scripting.Dictionary")
    Set q = New Collection
    d(기준) = 0
    q.Add 기준
    Do While q.Count > 0
        cur = CStr(q(1))
        q.Remove 1
        Set nb = 이웃(cur)
        For i = 1 To nb.Count
            상대 = Split(CStr(nb(i)), "|")(0)
            If Not d.Exists(상대) Then
                d(상대) = d(cur) + 1
                q.Add 상대
            End If
        Next i
    Loop
    Set 최단거리 = d
End Function

' 최단경로에서 각 표에 붙는 단계 정보 "앞표|방향|컬럼"
Private Function 경로단계(ByVal 기준 As String) As Object
    Dim d As Object, p As Object, q As Collection, cur As String
    Dim nb As Collection, i As Long, 상대 As String
    Set d = CreateObject("Scripting.Dictionary")
    Set p = CreateObject("Scripting.Dictionary")
    Set q = New Collection
    d(기준) = 0
    q.Add 기준
    Do While q.Count > 0
        cur = CStr(q(1))
        q.Remove 1
        Set nb = 이웃(cur)
        For i = 1 To nb.Count
            상대 = Split(CStr(nb(i)), "|")(0)
            If Not d.Exists(상대) Then
                d(상대) = d(cur) + 1
                p(상대) = cur & "|" & Split(CStr(nb(i)), "|")(1) & "|" & Split(CStr(nb(i)), "|")(2)
                q.Add 상대
            End If
        Next i
    Loop
    Set 경로단계 = p
End Function

' 기준 + 체크한 표 + 경로상 거쳐야 하는 표 (거리 순)
Private Function 유효표목록(ByVal ws As Worksheet) As Collection
    Dim cb As Object, 기준 As String, 거리 As Object, 단계 As Object
    Dim 필요 As Object, t As Variant, cur As String, 결과 As New Collection
    Dim d As Long, 최대 As Long

    기준 = 기준표()
    Set 거리 = 최단거리(기준)
    Set 단계 = 경로단계(기준)
    Set 필요 = CreateObject("Scripting.Dictionary")
    필요(기준) = True

    If Not ws Is Nothing Then
        On Error Resume Next
        For Each cb In ws.CheckBoxes
            If Left$(cb.Name, 4) = "TBL^" Then
                If cb.Value = xlOn Then
                    cur = Mid$(cb.Name, 5)
                    Do While Len(cur) > 0 And cur <> 기준
                        필요(cur) = True
                        If 단계.Exists(cur) Then cur = Split(CStr(단계(cur)), "|")(0) Else cur = ""
                    Loop
                End If
            End If
        Next cb
        On Error GoTo 0
    End If

    For Each t In 표_순서()
        If 필요.Exists(CStr(t)) Then
            If 거리.Exists(CStr(t)) Then
                If CLng(거리(CStr(t))) > 최대 Then 최대 = CLng(거리(CStr(t)))
            End If
        End If
    Next t
    For d = 0 To 최대
        For Each t In 표_순서()
            If 필요.Exists(CStr(t)) Then
                If 거리.Exists(CStr(t)) Then
                    If CLng(거리(CStr(t))) = d Then 결과.Add CStr(t)
                End If
            End If
        Next t
    Next d
    Set 유효표목록 = 결과
End Function

' 체크된 출력 컬럼 ("표^컬럼") - 표 순서 → 스키마 컬럼 순서
Private Function 출력컬럼(ByVal ws As Worksheet) As Collection
    Dim cb As Object, 켠 As Object, 유효 As Collection
    Dim k As Long, 표이름 As String, cols As Variant, i As Long
    Dim 결과 As New Collection, 기준PK As String, 있음 As Boolean, j As Long

    Set 켠 = CreateObject("Scripting.Dictionary")
    If Not ws Is Nothing Then
        On Error Resume Next
        For Each cb In ws.CheckBoxes
            If Left$(cb.Name, 4) = "COL^" Then
                If cb.Value = xlOn Then 켠(Mid$(cb.Name, 5)) = True
            End If
        Next cb
        On Error GoTo 0
    End If

    Set 유효 = 유효표목록(ws)
    For k = 1 To 유효.Count
        표이름 = CStr(유효(k))
        cols = 표_컬럼목록(표이름)
        For i = LBound(cols) To UBound(cols)
            If 켠.Exists(표이름 & "^" & cols(i)) Then 결과.Add 표이름 & "^" & cols(i)
        Next i
    Next k

    ' 기준 표의 PK 는 편집 대상을 특정하는 기준이라 항상 포함
    기준PK = 기준표() & "^" & 표_PK(기준표())
    For j = 1 To 결과.Count
        If CStr(결과(j)) = 기준PK Then 있음 = True
    Next j
    If Not 있음 Then
        If 결과.Count = 0 Then 결과.Add 기준PK Else 결과.Add 기준PK, , 1
    End If

    Set 출력컬럼 = 결과
End Function

'==============================================================================
' 5. 조회 실행 (조인 - Long 배열 + Dictionary 색인)
'==============================================================================
Public Sub 탐색_실행()
    Dim ws As Worksheet, 유효 As Collection, 컬럼들 As Collection
    Dim 기준 As String, 단계 As Object, s As Variant
    Dim mp() As Long, cnt As Long, 넘침 As Boolean
    Dim t0 As Double

    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then
        MsgBox "먼저 [탐색_만들기]를 실행해주세요.", vbInformation
        Exit Sub
    End If

    캐시비우기
    기준 = 기준표()
    Set 유효 = 유효표목록(ws)
    Set 컬럼들 = 출력컬럼(ws)
    If 컬럼들.Count = 0 Then
        MsgBox "출력할 컬럼을 하나 이상 체크해주세요.", vbInformation, "조회 실행"
        Exit Sub
    End If

    t0 = Timer
    s = 빠르게시작()
    On Error GoTo 정리

    Application.StatusBar = "조회 중… 표 " & 유효.Count & "개 연결"
    Set 단계 = 경로단계(기준)
    조인전개 기준, 유효, 단계, mp, cnt, 넘침
    조건적용 ws, 유효, 컬럼들, mp, cnt

    Application.StatusBar = "화면에 그리는 중… " & Format$(cnt, "#,##0") & "행"
    FK목록준비 유효                       ' 드롭다운 원본을 먼저 만들고
    결과_그리기 ws, 유효, 컬럼들, mp, cnt  ' 그 다음에 유효성 검사를 겁니다

정리:
    빠르게끝 s
    If Err.Number <> 0 Then
        MsgBox "조회 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation
        Exit Sub
    End If

    상태쓰기 라벨(기준) & " 기준 " & Format$(cnt, "#,##0") & "행 · 연결 표 " & _
             (유효.Count - 1) & "개 · 컬럼 " & 컬럼들.Count & "개 · " & _
             Format$(Timer - t0, "0.0") & "초"
    If 넘침 Then
        MsgBox "결과가 " & Format$(MAX_ROWS, "#,##0") & "행을 넘어 잘랐습니다." & vbCrLf & _
               "연결할 표를 줄이거나 조건으로 좁혀주세요.", vbExclamation, "조회 실행"
    End If
End Sub

' mp(표순서, 결과행) = 원본 행번호 (없으면 0)
Private Sub 조인전개(ByVal 기준 As String, ByVal 유효 As Collection, _
                     ByVal 단계 As Object, ByRef mp() As Long, ByRef cnt As Long, _
                     ByRef 넘침 As Boolean)
    Dim nt As Long, cap As Long, i As Long, j As Long, k As Long, r As Long
    Dim aB As Variant, pk열 As Long
    Dim 표이름 As String, 정보 As Variant, 앞표 As String, 방향 As String, 컬럼 As String
    Dim 앞위치 As Long, 색인 As Object, aP As Variant, 앞키열 As Long
    Dim np() As Long, ncap As Long, ncnt As Long
    Dim 부모행 As Long, 키 As String, 자식 As Collection

    nt = 유효.Count
    cap = 4096
    cnt = 0
    ReDim mp(1 To nt, 1 To cap)

    aB = 표배열(기준)
    pk열 = 헤더열(기준, 표_PK(기준))
    For r = 2 To UBound(aB, 1)
        If Len(값문자열(aB(r, pk열))) > 0 Then
            cnt = cnt + 1
            If cnt > cap Then
                cap = cap * 2
                ReDim Preserve mp(1 To nt, 1 To cap)
            End If
            mp(1, cnt) = r
        End If
    Next r

    For k = 2 To nt
        표이름 = CStr(유효(k))
        If Not 단계.Exists(표이름) Then GoTo 다음표
        정보 = Split(CStr(단계(표이름)), "|")
        앞표 = CStr(정보(0)): 방향 = CStr(정보(1)): 컬럼 = CStr(정보(2))

        앞위치 = 0
        For i = 1 To nt
            If CStr(유효(i)) = 앞표 Then 앞위치 = i
        Next i
        If 앞위치 = 0 Then GoTo 다음표

        aP = 표배열(앞표)
        If 방향 = "child" Then
            Set 색인 = FK색인(표이름, 컬럼)        ' 값 -> 자식 행들
            앞키열 = 헤더열(앞표, 표_PK(앞표))
        Else
            Set 색인 = PK색인(표이름)               ' PK -> 행
            앞키열 = 헤더열(앞표, 컬럼)
        End If

        ncap = cnt + 1024
        ncnt = 0
        ReDim np(1 To nt, 1 To ncap)

        For i = 1 To cnt
            부모행 = mp(앞위치, i)
            키 = ""
            If 부모행 > 0 Then 키 = 값문자열(aP(부모행, 앞키열))

            If Len(키) = 0 Or Not 색인.Exists(키) Then
                ncnt = ncnt + 1
                If ncnt > ncap Then
                    ncap = ncap * 2
                    ReDim Preserve np(1 To nt, 1 To ncap)
                End If
                For j = 1 To nt
                    np(j, ncnt) = mp(j, i)
                Next j
                np(k, ncnt) = 0
            ElseIf 방향 = "parent" Then
                ncnt = ncnt + 1
                If ncnt > ncap Then
                    ncap = ncap * 2
                    ReDim Preserve np(1 To nt, 1 To ncap)
                End If
                For j = 1 To nt
                    np(j, ncnt) = mp(j, i)
                Next j
                np(k, ncnt) = CLng(색인(키))
            Else
                Set 자식 = 색인(키)
                For j = 1 To 자식.Count
                    If ncnt >= MAX_ROWS Then
                        넘침 = True
                        Exit For
                    End If
                    ncnt = ncnt + 1
                    If ncnt > ncap Then
                        ncap = ncap * 2
                        ReDim Preserve np(1 To nt, 1 To ncap)
                    End If
                    Dim jj As Long
                    For jj = 1 To nt
                        np(jj, ncnt) = mp(jj, i)
                    Next jj
                    np(k, ncnt) = CLng(자식(j))
                Next j
            End If
            If 넘침 Then Exit For
        Next i

        ReDim mp(1 To nt, 1 To IIf(ncnt = 0, 1, ncnt))
        For i = 1 To ncnt
            For j = 1 To nt
                mp(j, i) = np(j, i)
            Next j
        Next i
        cnt = ncnt
다음표:
    Next k
End Sub

Private Sub 조건적용(ByVal ws As Worksheet, ByVal 유효 As Collection, _
                     ByVal 컬럼들 As Collection, ByRef mp() As Long, ByRef cnt As Long)
    Dim 필터 As String, 검색 As String, nt As Long
    Dim i As Long, k As Long, c As Long, j As Long
    Dim 빠짐 As Boolean, 걸림 As Boolean, 남김 As Boolean
    Dim np() As Long, ncnt As Long
    Dim 대상 As String, 대상위치 As Long
    Dim 조 As Variant, 표이름 As String, 위치 As Long, 원본 As Long

    If cnt = 0 Then Exit Sub
    nt = 유효.Count
    필터 = Trim$(CStr(ws.Cells(ROW_COND, COL_VAL).Value))
    검색 = Trim$(CStr(ws.Cells(ROW_COND, COL_VAL + 3).Value))
    If 필터 = F_ALL And Len(검색) = 0 Then Exit Sub

    대상위치 = 0
    If Left$(필터, 1) = "[" Then
        대상 = Mid$(필터, 2, InStr(필터, "]") - 2)
        For k = 1 To nt
            If 라벨(CStr(유효(k))) = 대상 Then 대상위치 = k
        Next k
    End If

    ReDim np(1 To nt, 1 To cnt)
    ncnt = 0

    For i = 1 To cnt
        남김 = True
        빠짐 = False
        For k = 2 To nt
            If mp(k, i) = 0 Then 빠짐 = True
        Next k
        If 필터 = F_ANY And Not 빠짐 Then 남김 = False
        If 필터 = F_NONE And 빠짐 Then 남김 = False
        If 대상위치 > 0 Then
            If mp(대상위치, i) <> 0 Then 남김 = False
        End If

        If 남김 And Len(검색) > 0 Then
            걸림 = False
            For c = 1 To 컬럼들.Count
                조 = Split(CStr(컬럼들(c)), "^")
                표이름 = CStr(조(0))
                위치 = 0
                For k = 1 To nt
                    If CStr(유효(k)) = 표이름 Then 위치 = k
                Next k
                If 위치 > 0 Then
                    원본 = mp(위치, i)
                    If 원본 > 0 Then
                        If InStr(1, 셀값(표이름, 원본, CStr(조(1))), 검색, vbTextCompare) > 0 Then
                            걸림 = True
                            Exit For
                        End If
                    End If
                End If
            Next c
            If Not 걸림 Then 남김 = False
        End If

        If 남김 Then
            ncnt = ncnt + 1
            For j = 1 To nt
                np(j, ncnt) = mp(j, i)
            Next j
        End If
    Next i

    ReDim mp(1 To nt, 1 To IIf(ncnt = 0, 1, ncnt))
    For i = 1 To ncnt
        For j = 1 To nt
            mp(j, i) = np(j, i)
        Next j
    Next i
    cnt = ncnt
End Sub

'==============================================================================
' 6. 결과 그리기 (배열 한 번에 붙여넣기 + 조건부 서식)
'==============================================================================
' FK 드롭다운 원본을 맵시트에 준비하고 이름(Name)으로 등록합니다.
' 이름은 PPAX_ 접두어를 써서 다른 모듈의 정의를 건드리지 않습니다.
Private Sub FK목록준비(ByVal 유효 As Collection)
    Dim wsM As Worksheet, k As Long, 표이름 As String
    Dim a As Variant, pk As Long, r As Long, n As Long, 열 As Long
    Dim 지울 As Collection, nm As Name, i As Long, 주소 As String

    Set wsM = 맵시트()

    ' 이름 삭제는 먼저 모아둔 뒤에 (컬렉션을 돌면서 지우면 건너뜁니다)
    Set 지울 = New Collection
    On Error Resume Next
    For Each nm In ThisWorkbook.Names
        If Left$(nm.Name, Len(NAME_PREFIX)) = NAME_PREFIX Then 지울.Add nm.Name
    Next nm
    For i = 1 To 지울.Count
        ThisWorkbook.Names(CStr(지울(i))).Delete
    Next i
    On Error GoTo 0

    열 = FK_LIST_COL
    For k = 1 To 유효.Count
        표이름 = CStr(유효(k))
        pk = 헤더열(표이름, 표_PK(표이름))
        a = 표배열(표이름)
        If pk > 0 And Not IsEmpty(a) Then
            wsM.Columns(열).ClearContents
            n = 0
            Dim out() As Variant
            ReDim out(1 To UBound(a, 1), 1 To 1)
            For r = 2 To UBound(a, 1)
                If Len(값문자열(a(r, pk))) > 0 Then
                    n = n + 1
                    out(n, 1) = 값문자열(a(r, pk))
                End If
            Next r
            If n > 0 Then
                wsM.Range(wsM.Cells(1, 열), wsM.Cells(n, 열)).Value = out
                주소 = wsM.Range(wsM.Cells(1, 열), wsM.Cells(n, 열)).Address(True, True)
                On Error Resume Next
                ThisWorkbook.Names.Add Name:=NAME_PREFIX & 라벨(표이름), _
                    RefersTo:="='" & MAP_SHEET & "'!" & 주소, Visible:=False
                On Error GoTo 0
            End If
            열 = 열 + 1
        End If
    Next k
End Sub

Private Sub 결과_그리기(ByVal ws As Worksheet, ByVal 유효 As Collection, _
                        ByVal 컬럼들 As Collection, ByRef mp() As Long, ByVal cnt As Long)
    Dim wsM As Worksheet, nt As Long, nc As Long
    Dim out() As Variant, hdr() As Variant, mapOut() As Variant, mark() As Variant
    Dim c As Long, i As Long, k As Long, 위치 As Long, 원본 As Long
    Dim 조 As Variant, 표이름 As String, 컬럼 As String
    Dim aT As Variant, jc As Long, 마지막열 As Long
    Dim rngData As Range, rngHead As Range, 참조 As String

    nt = 유효.Count
    nc = 컬럼들.Count
    마지막열 = COL_VAL + nc - 1
    결과_지우기 ws

    ' --- 머리글 (표 이름 + 컬럼 이름 2줄) ---
    ReDim hdr(1 To 1, 1 To nc)
    For c = 1 To nc
        조 = Split(CStr(컬럼들(c)), "^")
        hdr(1, c) = 라벨(CStr(조(0))) & Chr$(10) & CStr(조(1))
    Next c
    Set rngHead = ws.Range(ws.Cells(ROW_HEAD, COL_VAL), ws.Cells(ROW_HEAD, 마지막열))
    rngHead.Value = hdr
    With rngHead
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 9
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(239, 238, 231)
        .Font.Color = RGB(70, 70, 70)
    End With
    ws.Rows(ROW_HEAD).RowHeight = 32
    With ws.Cells(ROW_HEAD, COL_LBL)
        .Value = "결과"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With

    ' PK / FK 머리글만 색 구분
    For c = 1 To nc
        조 = Split(CStr(컬럼들(c)), "^")
        If CStr(조(1)) = 표_PK(CStr(조(0))) Then
            With ws.Cells(ROW_HEAD, COL_VAL + c - 1)
                .Interior.Color = C_TEAL
                .Font.Color = RGB(255, 255, 255)
            End With
        ElseIf Len(FK_참조(CStr(조(0)), CStr(조(1)))) > 0 Then
            With ws.Cells(ROW_HEAD, COL_VAL + c - 1)
                .Interior.Color = RGB(238, 237, 254)
                .Font.Color = C_PURPLE
            End With
        End If
    Next c

    If cnt = 0 Then
        ws.Cells(ROW_DATA, COL_VAL).Value = "조건에 맞는 데이터가 없습니다."
        ws.Cells(ROW_DATA, COL_VAL).Font.Color = C_SUB
        누락필터갱신 ws, 유효
        Exit Sub
    End If

    ' --- 데이터: 열 단위로 채워 배열 복사를 컬럼 수만큼만 ---
    ReDim out(1 To cnt, 1 To nc)
    For c = 1 To nc
        조 = Split(CStr(컬럼들(c)), "^")
        표이름 = CStr(조(0)): 컬럼 = CStr(조(1))
        위치 = 0
        For k = 1 To nt
            If CStr(유효(k)) = 표이름 Then 위치 = k
        Next k
        aT = 표배열(표이름)
        jc = 헤더열(표이름, 컬럼)
        For i = 1 To cnt
            원본 = 0
            If 위치 > 0 Then 원본 = mp(위치, i)
            If 원본 = 0 Or jc = 0 Then
                out(i, c) = MISS_MARK
            Else
                out(i, c) = 값문자열(aT(원본, jc))
            End If
        Next i
    Next c

    Set rngData = ws.Range(ws.Cells(ROW_DATA, COL_VAL), ws.Cells(ROW_DATA + cnt - 1, 마지막열))
    rngData.NumberFormatLocal = "@"
    rngData.Value = out                                  ' 한 번에 붙여넣기
    With rngData
        .Font.Size = 9.5
        .Font.Color = C_INK
        .Font.Bold = False
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .HorizontalAlignment = xlLeft
        .RowHeight = 18
    End With

    ' PK 열은 편집 불가 - 열 단위로 한 번씩만
    For c = 1 To nc
        조 = Split(CStr(컬럼들(c)), "^")
        If CStr(조(1)) = 표_PK(CStr(조(0))) Then
            With ws.Range(ws.Cells(ROW_DATA, COL_VAL + c - 1), _
                          ws.Cells(ROW_DATA + cnt - 1, COL_VAL + c - 1))
                .Interior.Color = RGB(240, 245, 244)
                .Font.Color = C_TEALD
                .Font.Bold = True
            End With
        End If
    Next c

    ' 조건부 서식: 누락(-) 먼저, 그 다음 줄무늬 (규칙 하나로 전체 처리)
    On Error Resume Next
    rngData.FormatConditions.Delete
    With rngData.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
                                      Formula1:="=""" & MISS_MARK & """")
        .Interior.Color = RGB(250, 238, 218)
        .Font.Color = RGB(176, 120, 23)
        .StopIfTrue = True
    End With
    With rngData.FormatConditions.Add(Type:=xlExpression, _
                                      Formula1:="=MOD(ROW(),2)=0")
        .Interior.Color = RGB(250, 249, 245)
    End With
    On Error GoTo 0

    ' FK 열은 드롭다운으로만 고르게 (잘못된 참조를 애초에 못 넣게)
    For c = 1 To nc
        조 = Split(CStr(컬럼들(c)), "^")
        표이름 = CStr(조(0)): 컬럼 = CStr(조(1))
        참조 = FK_참조(표이름, 컬럼)
        If Len(참조) > 0 Then
            그리드유효성 ws, c, cnt, "=" & NAME_PREFIX & 라벨(참조), _
                         "연결된 " & 라벨(참조) & " 에 없는 ID 입니다."
        ElseIf 열형식(표이름, 컬럼) = "BOOL" Then
            그리드유효성 ws, c, cnt, "TRUE,FALSE", "TRUE 또는 FALSE 만 넣을 수 있습니다."
        End If
    Next c

    ' --- 결과 ↔ 원본 행 대응표 ---
    Set wsM = 맵시트()
    wsM.Range(wsM.Columns(1), wsM.Columns(FK_LIST_COL - 1)).ClearContents
    wsM.Cells(1, 1).Value = 유효표문자열(유효)
    wsM.Cells(2, 1).Value = 컬럼문자열(컬럼들)
    ReDim mapOut(1 To cnt, 1 To nt)
    For i = 1 To cnt
        For k = 1 To nt
            mapOut(i, k) = mp(k, i)
        Next k
    Next i
    wsM.Range(wsM.Cells(4, 1), wsM.Cells(3 + cnt, nt)).Value = mapOut

    ReDim mark(1 To cnt, 1 To 1)
    For i = 1 To cnt
        mark(i, 1) = "R" & i
    Next i
    ws.Range(ws.Cells(ROW_DATA, COL_KEY), ws.Cells(ROW_DATA + cnt - 1, COL_KEY)).Value = mark

    ' 엑셀 기본 필터
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(ROW_HEAD, COL_VAL), ws.Cells(ROW_DATA + cnt - 1, 마지막열)).AutoFilter
    On Error GoTo 0

    누락필터갱신 ws, 유효
End Sub

Private Sub 그리드유효성(ByVal ws As Worksheet, ByVal c As Long, ByVal cnt As Long, _
                         ByVal 목록 As String, ByVal 안내 As String)
    On Error Resume Next
    With ws.Range(ws.Cells(ROW_DATA, COL_VAL + c - 1), _
                  ws.Cells(ROW_DATA + cnt - 1, COL_VAL + c - 1)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=목록
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "참조 확인"
        .ErrorMessage = 안내
    End With
    On Error GoTo 0
End Sub

Private Sub 결과_지우기(ByVal ws As Worksheet)
    Dim 끝행 As Long
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
    끝행 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    If 끝행 < ROW_HEAD Then 끝행 = ROW_HEAD
    With ws.Range(ws.Rows(ROW_HEAD), ws.Rows(끝행 + 3))
        On Error Resume Next
        .Validation.Delete
        .FormatConditions.Delete
        On Error GoTo 0
        .Clear
    End With
End Sub

Private Sub 누락필터갱신(ByVal ws As Worksheet, ByVal 유효 As Collection)
    Dim s As String, k As Long, 현재 As String
    s = F_ALL & "," & F_ANY & "," & F_NONE
    For k = 2 To 유효.Count
        s = s & ",[" & 라벨(CStr(유효(k))) & "] 이(가) 없는 행만"
    Next k
    현재 = Trim$(CStr(ws.Cells(ROW_COND, COL_VAL).Value))
    목록달기 ws.Cells(ROW_COND, COL_VAL), s
    If InStr(1, "," & s & ",", "," & 현재 & ",") = 0 Then
        ws.Cells(ROW_COND, COL_VAL).Value = F_ALL
    End If
End Sub

Private Function 맵시트() As Worksheet
    Set 맵시트 = 시트(MAP_SHEET)
    If 맵시트 Is Nothing Then
        Set 맵시트 = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        맵시트.Name = MAP_SHEET
        맵시트.Visible = xlSheetHidden
    End If
End Function

Private Function 유효표문자열(ByVal 유효 As Collection) As String
    Dim k As Long, s As String
    For k = 1 To 유효.Count
        s = s & IIf(Len(s) > 0, ";", "") & CStr(유효(k))
    Next k
    유효표문자열 = s
End Function

Private Function 컬럼문자열(ByVal 컬럼들 As Collection) As String
    Dim k As Long, s As String
    For k = 1 To 컬럼들.Count
        s = s & IIf(Len(s) > 0, ";", "") & CStr(컬럼들(k))
    Next k
    컬럼문자열 = s
End Function

'==============================================================================
' 7. 수정 저장
'==============================================================================
Public Sub 탐색_변경저장()
    Dim ws As Worksheet, wsM As Worksheet, s As Variant
    Dim 유효 As Variant, 컬럼들 As Variant
    Dim cnt As Long, nt As Long, nc As Long
    Dim gv As Variant, mv As Variant
    Dim c As Long, i As Long, k As Long, 위치 As Long, 원본 As Long
    Dim 조 As Variant, 표이름 As String, 컬럼 As String
    Dim aT As Variant, jc As Long
    Dim 화면값 As String, 원래값 As String
    Dim 변경 As Object, 키 As String
    Dim 오류 As String, 경고 As String, 미리보기 As String
    Dim 답 As VbMsgBoxResult, 건수 As Long

    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then
        MsgBox "먼저 [탐색_만들기]를 실행해주세요.", vbInformation, "수정 저장"
        Exit Sub
    End If
    Set wsM = 시트(MAP_SHEET)
    If wsM Is Nothing Then
        MsgBox "먼저 [② 조회 실행]으로 결과를 만들어주세요.", vbInformation, "수정 저장"
        Exit Sub
    End If
    If Len(Trim$(CStr(wsM.Cells(1, 1).Value))) = 0 Then
        MsgBox "먼저 [② 조회 실행]으로 결과를 만들어주세요.", vbInformation, "수정 저장"
        Exit Sub
    End If

    캐시비우기
    유효 = Split(CStr(wsM.Cells(1, 1).Value), ";")
    컬럼들 = Split(CStr(wsM.Cells(2, 1).Value), ";")
    nt = UBound(유효) - LBound(유효) + 1
    nc = UBound(컬럼들) - LBound(컬럼들) + 1
    cnt = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row - ROW_DATA + 1
    If cnt < 1 Then
        MsgBox "결과가 없습니다.", vbInformation, "수정 저장"
        Exit Sub
    End If

    s = 빠르게시작()
    Set 변경 = CreateObject("Scripting.Dictionary")
    On Error GoTo 정리

    ' 화면과 대응표를 각각 한 번에 읽습니다
    gv = ws.Range(ws.Cells(ROW_DATA, COL_VAL), ws.Cells(ROW_DATA + cnt - 1, COL_VAL + nc - 1)).Value
    mv = wsM.Range(wsM.Cells(4, 1), wsM.Cells(3 + cnt, nt)).Value

    For c = 1 To nc
        조 = Split(CStr(컬럼들(c - 1 + LBound(컬럼들))), "^")
        표이름 = CStr(조(0)): 컬럼 = CStr(조(1))
        위치 = 0
        For k = LBound(유효) To UBound(유효)
            If CStr(유효(k)) = 표이름 Then 위치 = k - LBound(유효) + 1
        Next k
        aT = 표배열(표이름)
        jc = 헤더열(표이름, 컬럼)

        For i = 1 To cnt
            원본 = 0
            If 위치 > 0 Then 원본 = CLng(Val(mv(i, 위치) & ""))
            화면값 = Trim$(CStr(gv(i, c) & ""))

            If 원본 = 0 Then
                If Len(화면값) > 0 And 화면값 <> MISS_MARK Then
                    경고 = 경고 & "· " & (ROW_DATA + i - 1) & "행 " & 라벨(표이름) & "." & 컬럼 & _
                           " : 연결된 레코드가 없어 저장할 수 없습니다." & vbCrLf
                End If
            ElseIf jc > 0 Then
                원래값 = 값문자열(aT(원본, jc))
                If 화면값 <> 원래값 Then
                    If 컬럼 = 표_PK(표이름) Then
                        오류 = 오류 & "· " & (ROW_DATA + i - 1) & "행 " & 라벨(표이름) & "." & 컬럼 & _
                               " : ID는 여기서 바꿀 수 없습니다 (" & 원래값 & " → " & 화면값 & ")" & vbCrLf
                    Else
                        키 = 표이름 & "^" & 원본 & "^" & 컬럼
                        If 변경.Exists(키) Then
                            If Split(CStr(변경(키)), vbTab)(1) <> 화면값 Then
                                오류 = 오류 & "· " & 라벨(표이름) & " " & _
                                       셀값(표이름, 원본, 표_PK(표이름)) & " 의 " & 컬럼 & _
                                       " 을(를) 서로 다른 값으로 고쳤습니다 (" & _
                                       Split(CStr(변경(키)), vbTab)(1) & " / " & 화면값 & ")" & vbCrLf
                            End If
                        Else
                            변경(키) = 원래값 & vbTab & 화면값
                        End If
                    End If
                End If
            End If
        Next i
    Next c

    ' FK 로 고친 값이 실제로 있는지
    Dim ky As Variant, 조각 As Variant, 참조 As String, 새값 As String
    For Each ky In 변경.Keys
        조각 = Split(CStr(ky), "^")
        참조 = FK_참조(CStr(조각(0)), CStr(조각(2)))
        If Len(참조) > 0 Then
            새값 = Split(CStr(변경(ky)), vbTab)(1)
            If Len(새값) = 0 Then
                오류 = 오류 & "· " & 라벨(CStr(조각(0))) & "." & 조각(2) & " 는 비워둘 수 없습니다." & vbCrLf
            ElseIf Not PK색인(참조).Exists(새값) Then
                오류 = 오류 & "· " & 라벨(CStr(조각(0))) & "." & 조각(2) & " = '" & 새값 & _
                       "' 에 해당하는 " & 라벨(참조) & " 가 없습니다." & vbCrLf
            End If
        End If
    Next ky

정리:
    빠르게끝 s
    If Err.Number <> 0 Then
        MsgBox "변경 내용을 확인하는 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation
        Exit Sub
    End If

    If Len(오류) > 0 Then
        MsgBox "저장할 수 없습니다." & vbCrLf & vbCrLf & 오류, vbExclamation, "수정 저장"
        Exit Sub
    End If
    If 변경.Count = 0 Then
        MsgBox "바뀐 내용이 없습니다." & IIf(Len(경고) > 0, vbCrLf & vbCrLf & 경고, ""), _
               vbInformation, "수정 저장"
        Exit Sub
    End If

    건수 = 0
    For Each ky In 변경.Keys
        조각 = Split(CStr(ky), "^")
        건수 = 건수 + 1
        If 건수 <= 40 Then
            미리보기 = 미리보기 & "· " & 라벨(CStr(조각(0))) & " " & _
                       셀값(CStr(조각(0)), CLng(조각(1)), 표_PK(CStr(조각(0)))) & _
                       " / " & 조각(2) & " : " & _
                       IIf(Len(Split(CStr(변경(ky)), vbTab)(0)) = 0, "(공란)", _
                           Split(CStr(변경(ky)), vbTab)(0)) & _
                       " → " & Split(CStr(변경(ky)), vbTab)(1) & vbCrLf
        End If
    Next ky
    If 건수 > 40 Then 미리보기 = 미리보기 & "  … 외 " & (건수 - 40) & "건" & vbCrLf

    답 = MsgBox("아래 " & 건수 & "건을 원본 시트에 반영합니다." & vbCrLf & vbCrLf & _
                미리보기 & IIf(Len(경고) > 0, vbCrLf & "[저장 안 되는 항목]" & vbCrLf & 경고, "") & _
                vbCrLf & "계속할까요?", vbQuestion + vbYesNo, "수정 저장 확인")
    If 답 <> vbYes Then
        상태쓰기 "저장을 취소했습니다.", True
        Exit Sub
    End If

    s = 빠르게시작()
    For Each ky In 변경.Keys
        조각 = Split(CStr(ky), "^")
        셀쓰기 CStr(조각(0)), CLng(조각(1)), CStr(조각(2)), Split(CStr(변경(ky)), vbTab)(1)
    Next ky
    빠르게끝 s

    상태쓰기 건수 & "건을 저장했습니다."
    탐색_실행
    MsgBox 건수 & "건을 원본 시트에 반영했습니다.", vbInformation, "수정 저장"
End Sub

Public Sub 탐색_되돌리기()
    Dim 답 As VbMsgBoxResult
    답 = MsgBox("화면에서 고친 내용을 버리고 원본 값으로 다시 불러옵니다." & vbCrLf & _
                "계속할까요?", vbQuestion + vbYesNo, "수정 취소")
    If 답 <> vbYes Then Exit Sub
    탐색_실행
    상태쓰기 "원본 값으로 되돌렸습니다."
End Sub

'==============================================================================
' 8. 빠른 조회 프리셋
'==============================================================================
Private Function 프리셋목록() As Collection
    Dim c As New Collection, t As Variant, cols As Variant, i As Long, 참조 As String
    For Each t In 표_순서()
        cols = 표_컬럼목록(CStr(t))
        For i = LBound(cols) To UBound(cols)
            참조 = FK_참조(CStr(t), CStr(cols(i)))
            If Len(참조) > 0 Then
                c.Add 참조 & "|" & CStr(t) & "|" & 라벨(참조) & " 중 " & 라벨(CStr(t)) & " 없음"
            End If
        Next i
    Next t
    Set 프리셋목록 = c
End Function

Public Sub 탐색_프리셋적용()
    Dim ws As Worksheet, c As Collection, i As Long, 고른 As String
    Dim 기준 As String, 연결 As String, cb As Object, s As Variant

    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    고른 = Trim$(CStr(ws.Cells(ROW_PRESET, COL_VAL).Value))
    If Len(고른) = 0 Or 고른 = P_NONE Then
        MsgBox "빠른 조회 항목을 먼저 고르세요.", vbInformation, "빠른 조회"
        Exit Sub
    End If

    Set c = 프리셋목록()
    For i = 1 To c.Count
        If Split(CStr(c(i)), "|")(2) = 고른 Then
            기준 = Split(CStr(c(i)), "|")(0)
            연결 = Split(CStr(c(i)), "|")(1)
            Exit For
        End If
    Next i
    If Len(기준) = 0 Then Exit Sub

    캐시비우기
    s = 빠르게시작()
    ws.Cells(ROW_BASE, COL_VAL).Value = 라벨(기준)
    체크박스_다시그리기 ws
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If cb.Name = "TBL^" & 연결 Then cb.Value = xlOn
    Next cb
    On Error GoTo 0
    체크박스_다시그리기 ws
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If cb.Name = "COL^" & 연결 & "^" & 표_PK(연결) Then cb.Value = xlOn
    Next cb
    On Error GoTo 0
    ws.Cells(ROW_COND, COL_VAL).Value = "[" & 라벨(연결) & "] 이(가) 없는 행만"
    빠르게끝 s

    탐색_실행
End Sub

'==============================================================================
' 9. 내보내기 / 초기화 / 상태
'==============================================================================
Public Sub 탐색_내보내기()
    Dim ws As Worksheet, 새책 As Workbook, 끝행 As Long, 끝열 As Long
    Dim v As Variant, 시트새 As Worksheet
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    끝행 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    If 끝행 < ROW_DATA Then
        MsgBox "내보낼 결과가 없습니다. 먼저 [② 조회 실행]을 눌러주세요.", vbInformation, "내보내기"
        Exit Sub
    End If
    끝열 = ws.Cells(ROW_HEAD, ws.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False
    ' 클립보드를 건드리지 않고 값만 옮깁니다
    v = ws.Range(ws.Cells(ROW_HEAD, COL_VAL), ws.Cells(끝행, 끝열)).Value
    Set 새책 = Workbooks.Add
    Set 시트새 = 새책.Worksheets(1)
    시트새.Range(시트새.Cells(1, 1), _
                 시트새.Cells(UBound(v, 1), UBound(v, 2))).Value = v
    With 시트새.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(239, 238, 231)
        .RowHeight = 30
        .WrapText = True
    End With
    시트새.Range("A2").Select
    ActiveWindow.FreezePanes = True
    시트새.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "새 통합 문서로 내보냈습니다. 원하는 이름으로 저장하세요.", vbInformation, "내보내기"
End Sub

Public Sub 탐색_초기화()
    Dim ws As Worksheet
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    결과_지우기 ws
    ws.Cells(ROW_COND, COL_VAL + 3).ClearContents
    ws.Cells(ROW_COND, COL_VAL).Value = F_ALL
    ws.Cells(ROW_PRESET, COL_VAL).Value = P_NONE
    On Error Resume Next
    맵시트().Cells.ClearContents
    On Error GoTo 0
    상태쓰기 "결과를 지웠습니다."
End Sub

Private Sub 상태쓰기(ByVal msg As String, Optional ByVal 경고 As Boolean = False)
    Dim ws As Worksheet
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    With ws.Cells(ROW_STATUS, COL_VAL)
        .Value = msg & "   (" & Format$(Now, "hh:nn:ss") & ")"
        .Font.Bold = True
        .Font.Color = IIf(경고, RGB(178, 58, 58), RGB(31, 122, 84))
    End With
End Sub

'==============================================================================
' 10. (선택) 시트 이벤트
'   탐색 시트 코드창에 아래를 붙여넣으면
'     · 기준 표 / 빠른 조회를 고르는 즉시 반영되고
'     · 결과에서 고친 칸이 빨간 굵은 글씨로 표시됩니다 (저장 전까지)
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       탐색_변경감지 Target
'   End Sub
'==============================================================================
Public Sub 탐색_변경감지(ByVal Target As Range)
    Dim ws As Worksheet, 이전 As Boolean, 끝행 As Long
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> EXP_SHEET Then Exit Sub

    ' 결과 영역을 고쳤으면 "아직 저장 안 됨" 표시만 남깁니다
    끝행 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    If Target.Row >= ROW_DATA And Target.Row <= 끝행 And Target.Column >= COL_VAL Then
        이전 = Application.EnableEvents
        Application.EnableEvents = False
        On Error Resume Next
        With Target
            .Font.Color = RGB(178, 58, 58)
            .Font.Bold = True
        End With
        On Error GoTo 0
        Application.EnableEvents = 이전
        상태쓰기 "수정한 칸이 있습니다. [수정 저장]을 눌러 반영하세요.", True
        Exit Sub
    End If

    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Column <> COL_VAL Then Exit Sub

    If Target.Row = ROW_BASE Then
        이전 = Application.EnableEvents
        Application.EnableEvents = False
        On Error Resume Next
        캐시비우기
        체크박스_다시그리기 ws
        On Error GoTo 0
        Application.EnableEvents = 이전
        상태쓰기 "기준 표를 바꿨습니다. 컬럼을 확인하고 [② 조회 실행]을 누르세요."
    ElseIf Target.Row = ROW_PRESET Then
        탐색_프리셋적용
    End If
End Sub
