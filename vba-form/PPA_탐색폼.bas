Attribute VB_Name = "PPA_Explorer"
'==============================================================================
' PPA 관계형 데이터 탐색 · 편집 폼
'   ※ 이 파일은 ANSI(CP949)로 저장돼 있습니다. VBE의 [파일 가져오기]가 한국어
'      Windows 기본 코드페이지로 읽기 때문입니다. 메모장 등으로 다시 저장할 일이
'      있으면 인코딩을 반드시 'ANSI'로 두세요. UTF-8로 저장하면 가져올 때 한글이
'      깨지면서 "허가된 개체 이름이 아닙니다" 오류가 납니다.
'------------------------------------------------------------------------------
' HTML 대시보드의 [탐색] 탭을 엑셀 안으로 옮긴 것입니다. 다른 점은 하나 -
' 여기서는 결과를 보기만 하는 게 아니라 그 자리에서 고쳐서 원본에 반영할 수
' 있습니다.
'
'   1. 기준 표를 고르고
'   2. 연결할 표를 체크하면 (경로상 필요한 표는 자동으로 따라옵니다)
'   3. 보고 싶은 컬럼만 체크해서
'   4. 여러 표가 합쳐진 목록을 만들고, 그 위에서 바로 수정합니다.
'
' 연결 상대가 없는 행도 빈칸(-)으로 남기기 때문에 "구매계약이 없는 발전소" 같은
' 누락 데이터를 그대로 찾아낼 수 있습니다(LEFT JOIN).
'
' 편집 안전장치
'   · PK 열은 수정 대상에서 제외합니다 (ID를 바꾸면 FK가 전부 깨지므로)
'   · FK 열을 고치면 그 값이 실제로 있는지 검사합니다
'   · 같은 원본 행이 여러 줄에 걸쳐 보일 때(1:N 전개) 서로 다른 값으로 고치면
'     충돌로 잡아내고 저장을 멈춥니다
'   · 저장 전에 "무엇이 어떻게 바뀌는지"를 보여주고 확인을 받습니다
'   · 연결 상대가 없어 비어 있는 칸은 쓸 대상이 없으므로 편집을 막습니다
'
' 사용법
'   1) VBE(Alt+F11) → 파일 → 파일 가져오기 → 이 .bas 선택
'   2) Alt+F8 → "탐색_만들기" 실행
'   3) 이후로는 탐색 시트의 버튼만 누르면 됩니다.
'
' 이 모듈은 단독으로 동작합니다 (PPA_InputForm 이 없어도 됩니다).
'==============================================================================
Option Explicit

'---- 시트 이름 --------------------------------------------------------------
Private Const EXP_SHEET As String = "탐색"
Private Const MAP_SHEET As String = "_탐색맵"      ' 결과 ↔ 원본 행 대응표(숨김)

Private Const SH_PLANT  As String = "T_발전소"
Private Const SH_BUY    As String = "T_구매계약"
Private Const SH_DEMAND As String = "T_수요기업"
Private Const SH_SELL   As String = "T_판매계약"
Private Const SH_SITE   As String = "T_전기사용지"
Private Const SH_MATCH  As String = "T_수급매칭"

'---- 화면 레이아웃 (고정 행) ------------------------------------------------
Private Const COL_KEY   As Long = 1      ' 숨김 키 열
Private Const COL_LBL   As Long = 2      ' 라벨
Private Const COL_VAL   As Long = 3      ' 값 / 결과 첫 컬럼

Private Const ROW_TITLE   As Long = 2
Private Const ROW_PRESET  As Long = 3    ' 빠른 조회
Private Const ROW_BASE    As Long = 4    ' 1. 기준 표
Private Const ROW_JOIN    As Long = 6    ' 2. 연결할 표 (체크박스)
Private Const ROW_COLHEAD As Long = 8    ' 3. 출력 컬럼 안내
Private Const ROW_COL1    As Long = 9    ' 표별 컬럼 체크박스 6줄 (9~14)
Private Const ROW_COND    As Long = 16   ' 4. 조건 (누락필터 / 검색어)
Private Const ROW_STATUS  As Long = 17   ' 상태
Private Const ROW_HEAD    As Long = 19   ' 결과 머리글
Private Const ROW_DATA    As Long = 20   ' 결과 첫 데이터 행

Private Const MAX_ROWS As Long = 5000    ' 결과 행 상한
Private Const MISS_MARK As String = "-"  ' 연결 상대 없음 표시

Private Const F_ALL   As String = "전체 (연결 여부 무관)"
Private Const F_ANY   As String = "한 곳이라도 빠진 행만 (누락)"
Private Const F_NONE  As String = "전부 연결된 행만"
Private Const P_NONE  As String = "(직접 설정)"

'==============================================================================
' 스키마 - static-dashboard/ppa_schema.py, backend의 tableDefs.ts와 동일
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

' FK: 해당 컬럼이 참조하는 시트 (아니면 "")
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

' 표 사이 연결 (부모/자식 양방향). 반환: "상대시트|방향|컬럼" 목록
'   방향 parent = 이 표가 상대를 참조 / child = 상대가 이 표를 참조
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
' 공통 도우미
'==============================================================================
Private Function 시트(ByVal 이름 As String) As Worksheet
    On Error Resume Next
    Set 시트 = ThisWorkbook.Worksheets(이름)
    On Error GoTo 0
End Function

Private Function 헤더열(ByVal ws As Worksheet, ByVal 헤더 As String) As Long
    Dim 마지막 As Long, j As Long
    If ws Is Nothing Then Exit Function
    마지막 = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To 마지막
        If Trim$(CStr(ws.Cells(1, j).Value)) = 헤더 Then
            헤더열 = j
            Exit Function
        End If
    Next j
End Function

Private Function 마지막행(ByVal ws As Worksheet, ByVal pk열 As Long) As Long
    If ws Is Nothing Or pk열 = 0 Then Exit Function
    마지막행 = ws.Cells(ws.Rows.Count, pk열).End(xlUp).Row
    If 마지막행 < 1 Then 마지막행 = 1
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

Private Function 셀값(ByVal ws As Worksheet, ByVal r As Long, ByVal 헤더 As String) As String
    Dim j As Long
    If ws Is Nothing Or r < 2 Then Exit Function
    j = 헤더열(ws, 헤더)
    If j = 0 Then Exit Function
    셀값 = 값문자열(ws.Cells(r, j).Value)
End Function

' 원본 열의 기존 형식을 따라 값 쓰기 (ID 는 항상 문자열)
Private Sub 셀쓰기(ByVal ws As Worksheet, ByVal r As Long, ByVal 헤더 As String, ByVal v As String)
    Dim j As Long, 견본 As Variant, rr As Long, pk열 As Long, 끝 As Long, 견본있음 As Boolean

    j = 헤더열(ws, 헤더)
    If j = 0 Then Exit Sub
    If Len(v) = 0 Then
        ws.Cells(r, j).ClearContents
        Exit Sub
    End If

    If 헤더 = 표_PK(ws.Name) Or Len(FK_참조(ws.Name, 헤더)) > 0 Then
        ws.Cells(r, j).NumberFormatLocal = "@"
        ws.Cells(r, j).Value = v
        Exit Sub
    End If

    pk열 = 헤더열(ws, 표_PK(ws.Name))
    끝 = 마지막행(ws, pk열)
    For rr = 2 To 끝
        If rr <> r Then
            If Len(Trim$(CStr(ws.Cells(rr, j).Value))) > 0 Then
                견본 = ws.Cells(rr, j).Value
                견본있음 = True
                Exit For
            End If
        End If
    Next rr

    If 견본있음 And VarType(견본) = vbBoolean Then
        ws.Cells(r, j).Value = (UCase$(v) = "TRUE" Or v = "예" Or UCase$(v) = "Y")
    ElseIf 견본있음 And IsDate(견본) And IsDate(v) Then
        ws.Cells(r, j).Value = CDate(v)
        ws.Cells(r, j).NumberFormatLocal = "yyyy-mm-dd"
    ElseIf UCase$(v) = "TRUE" Then
        ws.Cells(r, j).Value = True
    ElseIf UCase$(v) = "FALSE" Then
        ws.Cells(r, j).Value = False
    ElseIf IsNumeric(v) And Not IsDate(v) Then
        ws.Cells(r, j).Value = CDbl(v)
    Else
        ws.Cells(r, j).Value = v
    End If
End Sub

Private Function PK행찾기(ByVal ws As Worksheet, ByVal PK값 As String) As Long
    Dim pk열 As Long, r As Long, 끝 As Long
    If ws Is Nothing Then Exit Function
    If Len(Trim$(PK값)) = 0 Then Exit Function
    pk열 = 헤더열(ws, 표_PK(ws.Name))
    If pk열 = 0 Then Exit Function
    끝 = 마지막행(ws, pk열)
    For r = 2 To 끝
        If Trim$(CStr(ws.Cells(r, pk열).Value)) = Trim$(PK값) Then
            PK행찾기 = r
            Exit Function
        End If
    Next r
End Function

'==============================================================================
' 화면 만들기
'==============================================================================
Public Sub 탐색_만들기()
    Dim ws As Worksheet

    If Not 원본표_점검() Then Exit Sub

    Application.ScreenUpdating = False
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
    ws.Columns(COL_LBL).ColumnWidth = 17

    With ws.Cells(ROW_TITLE, COL_LBL)
        .Value = "관계형 데이터 탐색 · 편집"
        .Font.Size = 15
        .Font.Bold = True
    End With
    ws.Cells(ROW_TITLE, COL_VAL).Value = "기준 표를 고르고 연결할 표와 컬럼을 체크한 뒤 [조회 실행]"
    ws.Cells(ROW_TITLE, COL_VAL).Font.Color = RGB(110, 110, 110)

    구역라벨 ws, ROW_PRESET, "빠른 조회", RGB(83, 74, 183)
    구역라벨 ws, ROW_BASE, "1. 기준 표", RGB(14, 124, 123)
    구역라벨 ws, ROW_JOIN, "2. 연결할 표", RGB(14, 124, 123)
    구역라벨 ws, ROW_COLHEAD, "3. 출력 컬럼", RGB(14, 124, 123)
    구역라벨 ws, ROW_COND, "4. 조건", RGB(14, 124, 123)

    ws.Cells(ROW_BASE, COL_VAL).Value = 라벨(SH_PLANT)
    ws.Cells(ROW_COLHEAD, COL_VAL).Value = "보고 싶은 컬럼만 체크하세요 (기준 표의 PK는 항상 포함됩니다)"
    ws.Cells(ROW_COLHEAD, COL_VAL).Font.Color = RGB(110, 110, 110)
    ws.Cells(ROW_JOIN, COL_VAL + 9).Value = "회색 = 경로상 자동으로 거쳐가는 표"
    ws.Cells(ROW_JOIN, COL_VAL + 9).Font.Color = RGB(110, 110, 110)

    ws.Cells(ROW_COND, COL_VAL).Value = F_ALL
    ws.Cells(ROW_COND, COL_VAL + 3).Value = "검색어"
    ws.Cells(ROW_COND, COL_VAL + 3).Font.Bold = True
    With ws.Cells(ROW_COND, COL_VAL + 4)
        .Interior.Color = RGB(255, 249, 219)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 196, 180)
        .NumberFormatLocal = "@"
    End With
    ws.Cells(ROW_PRESET, COL_VAL).Value = P_NONE

    ws.Cells(ROW_STATUS, COL_LBL).Value = "상태"
    ws.Cells(ROW_STATUS, COL_LBL).Font.Bold = True

    ws.Rows(ROW_JOIN).RowHeight = 20
    ws.Rows(ROW_TITLE).RowHeight = 24
    ws.Rows(1).RowHeight = 28

    버튼만들기 ws
    드롭다운만들기 ws
    체크박스_다시그리기
    상태쓰기 "준비됐습니다. 기준 표와 연결할 표를 고르고 [조회 실행]을 누르세요."

    ws.Activate
    ws.Cells(ROW_BASE, COL_VAL).Select

정리:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "탐색 시트를 만드는 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation
        Exit Sub
    End If

    MsgBox "탐색 시트를 만들었습니다." & vbCrLf & vbCrLf & _
           "· 기준 표를 바꾸면 [설정 적용]을 눌러 체크박스를 새로 그리세요." & vbCrLf & _
           "· 결과 칸을 직접 고친 뒤 [변경 저장]을 누르면 원본 시트에 반영됩니다.", _
           vbInformation, "PPA 탐색 · 편집"
End Sub

Private Sub 구역라벨(ByVal ws As Worksheet, ByVal r As Long, ByVal 글 As String, ByVal 색 As Long)
    With ws.Cells(r, COL_LBL)
        .Value = 글
        .Font.Bold = True
        .Font.Color = 색
    End With
End Sub

Private Sub 버튼만들기(ByVal ws As Worksheet)
    Dim b As Object, 정의 As Variant, i As Long, x As Double
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    정의 = Array("설정 적용|탐색_설정적용", "조회 실행|탐색_실행", "변경 저장|탐색_변경저장", _
                 "되돌리기|탐색_되돌리기", "엑셀로 내보내기|탐색_내보내기", "초기화|탐색_초기화")
    x = ws.Cells(1, COL_LBL).Left
    For i = LBound(정의) To UBound(정의)
        Set b = ws.Buttons.Add(x, ws.Cells(1, 1).Top + 3, 92, 22)
        b.Caption = Split(정의(i), "|")(0)
        b.OnAction = Split(정의(i), "|")(1)
        x = x + 96
    Next i
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
                If 헤더열(ws, CStr(cols(i))) = 0 Then
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
' 체크박스 (연결할 표 / 출력 컬럼)
'==============================================================================
Private Function 기준표() As String
    Dim ws As Worksheet, v As String
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Function
    v = Trim$(CStr(ws.Cells(ROW_BASE, COL_VAL).Value))
    기준표 = 시트명찾기(v)
    If Len(기준표) = 0 Then 기준표 = SH_PLANT
End Function

Public Sub 탐색_설정적용()
    체크박스_다시그리기
    상태쓰기 "설정을 적용했습니다. 컬럼을 고르고 [조회 실행]을 누르세요."
End Sub

Private Sub 체크박스_다시그리기()
    Dim ws As Worksheet, cb As Object, t As Variant, 표이름 As String
    Dim x As Double, y As Double, i As Long, cols As Variant
    Dim 기준 As String, 거리 As Object, 선택 As Object, 켠컬럼 As Object
    Dim r As Long, w As Double

    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    기준 = 기준표()

    ' 지금 켜져 있는 상태를 기억했다가 다시 그린 뒤 복원
    Set 선택 = CreateObject("Scripting.Dictionary")
    Set 켠컬럼 = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If Left$(cb.Name, 4) = "TBL^" Then
            If cb.Value = xlOn Then 선택(Mid$(cb.Name, 5)) = True
        ElseIf Left$(cb.Name, 4) = "COL^" Then
            If cb.Value = xlOn Then 켠컬럼(Mid$(cb.Name, 5)) = True
        End If
    Next cb
    ws.CheckBoxes.Delete
    On Error GoTo 0

    Set 거리 = 최단거리(기준)

    ' --- 2. 연결할 표 ---
    x = ws.Cells(ROW_JOIN, COL_VAL).Left
    y = ws.Cells(ROW_JOIN, COL_VAL).Top
    For Each t In 표_순서()
        표이름 = CStr(t)
        w = 26 + Len(라벨(표이름)) * 11
        If 표이름 = 기준 Then
            ws.Cells(ROW_JOIN, COL_VAL).Value = ""
            Set cb = ws.CheckBoxes.Add(x, y, w, 17)
            cb.Caption = 라벨(표이름) & " (기준)"
            cb.Name = "BASEMARK"
            cb.Value = xlOn
            cb.Enabled = False
            w = w + 40
        ElseIf 거리.Exists(표이름) Then
            Set cb = ws.CheckBoxes.Add(x, y, w, 17)
            cb.Caption = 라벨(표이름) & " " & 거리(표이름)
            cb.Name = "TBL^" & 표이름
            cb.OnAction = ""
            If 선택.Exists(표이름) Then cb.Value = xlOn Else cb.Value = xlOff
        Else
            w = 0
        End If
        x = x + w + 6
    Next t

    ' --- 3. 출력 컬럼 (표별 한 줄) ---
    Dim 유효 As Collection, k As Long, 사용중 As Object
    Set 유효 = 유효표목록()
    Set 사용중 = CreateObject("Scripting.Dictionary")
    For k = 1 To 유효.Count
        사용중(CStr(유효(k))) = True
    Next k

    r = ROW_COL1
    For Each t In 표_순서()
        표이름 = CStr(t)
        ws.Cells(r, COL_LBL).Value = 라벨(표이름)
        ws.Cells(r, COL_LBL).Font.Bold = 사용중.Exists(표이름)
        ws.Cells(r, COL_LBL).Font.Color = IIf(사용중.Exists(표이름), RGB(22, 38, 43), RGB(180, 178, 170))
        ws.Rows(r).RowHeight = 19

        If 사용중.Exists(표이름) Then
            cols = 표_컬럼목록(표이름)
            ' 이 표에 살아남은 선택이 하나도 없으면 앞쪽 2개를 기본으로 켭니다
            Dim 이표선택 As Boolean
            이표선택 = False
            For i = LBound(cols) To UBound(cols)
                If 켠컬럼.Exists(표이름 & "^" & cols(i)) Then 이표선택 = True
            Next i
            x = ws.Cells(r, COL_VAL).Left
            y = ws.Cells(r, COL_VAL).Top
            For i = LBound(cols) To UBound(cols)
                w = 26 + Len(CStr(cols(i))) * 10
                Set cb = ws.CheckBoxes.Add(x, y, w, 16)
                cb.Caption = CStr(cols(i))
                cb.Name = "COL^" & 표이름 & "^" & cols(i)
                If Not 이표선택 Then
                    ' 처음이거나 새로 붙인 표 - 알아보기 쉬운 앞쪽 컬럼을 기본 선택
                    cb.Value = IIf(i <= 1, xlOn, xlOff)
                ElseIf 켠컬럼.Exists(표이름 & "^" & cols(i)) Then
                    cb.Value = xlOn
                Else
                    cb.Value = xlOff
                End If
                x = x + w + 4
            Next i
        End If
        r = r + 1
    Next t
End Sub

' 기준 표에서 각 표까지의 최단 거리 (연결 안 되면 없음)
Private Function 최단거리(ByVal 기준 As String) As Object
    Dim d As Object, q As Collection, cur As String, e As Variant, nb As Collection
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

' 기준 + 사용자가 고른 표 + 경로상 거쳐야 하는 표 (거리 순)
Private Function 유효표목록() As Collection
    Dim ws As Worksheet, cb As Object, 기준 As String
    Dim 거리 As Object, 앞 As Object, 필요 As Object
    Dim t As Variant, cur As String, 결과 As New Collection
    Dim d As Long, 최대 As Long

    Set ws = 시트(EXP_SHEET)
    기준 = 기준표()
    Set 거리 = 최단거리(기준)
    Set 앞 = 경로부모(기준)
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
                        If 앞.Exists(cur) Then cur = CStr(앞(cur)) Else cur = ""
                    Loop
                End If
            End If
        Next cb
        On Error GoTo 0
    End If

    ' 거리 순으로 정렬 (부모가 먼저 붙어야 조인이 성립)
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

' 최단경로에서 각 표의 바로 앞 표
Private Function 경로부모(ByVal 기준 As String) As Object
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
    ' 값에서 앞 표 이름만 남기기
    Dim k As Variant, out As Object
    Set out = CreateObject("Scripting.Dictionary")
    For Each k In p.Keys
        out(CStr(k)) = Split(CStr(p(k)), "|")(0)
    Next k
    Set 경로부모 = out
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

' 체크된 출력 컬럼 ("표^컬럼") - 표 순서 → 컬럼 순서로 정렬
Private Function 출력컬럼() As Collection
    Dim ws As Worksheet, cb As Object, 켠 As Object
    Dim 유효 As Collection, k As Long, 표이름 As String
    Dim cols As Variant, i As Long, 결과 As New Collection

    Set ws = 시트(EXP_SHEET)
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

    Set 유효 = 유효표목록()
    For k = 1 To 유효.Count
        표이름 = CStr(유효(k))
        cols = 표_컬럼목록(표이름)
        For i = LBound(cols) To UBound(cols)
            If 켠.Exists(표이름 & "^" & cols(i)) Then 결과.Add 표이름 & "^" & cols(i)
        Next i
    Next k

    ' 기준 표의 PK 는 편집 대상을 특정하는 기준이라 항상 넣습니다
    Dim 기준PK As String
    기준PK = 기준표() & "^" & 표_PK(기준표())
    Dim 있음 As Boolean, j As Long
    For j = 1 To 결과.Count
        If CStr(결과(j)) = 기준PK Then 있음 = True
    Next j
    If Not 있음 Then 결과.Add 기준PK, , 1

    Set 출력컬럼 = 결과
End Function

'==============================================================================
' 조회 실행 (LEFT JOIN 전개)
'==============================================================================
Public Sub 탐색_실행()
    Dim ws As Worksheet, 유효 As Collection, 컬럼들 As Collection
    Dim 기준 As String, 단계 As Object
    Dim 행맵 As Collection            ' 각 결과행: "원본행1|원본행2|..." (유효표 순서)
    Dim i As Long, k As Long

    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then
        MsgBox "먼저 [탐색_만들기]를 실행해주세요.", vbInformation
        Exit Sub
    End If

    기준 = 기준표()
    Set 유효 = 유효표목록()
    Set 컬럼들 = 출력컬럼()
    If 컬럼들.Count = 0 Then
        MsgBox "출력할 컬럼을 하나 이상 체크해주세요.", vbInformation, "조회 실행"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set 단계 = 경로단계(기준)
    Set 행맵 = 조인전개(기준, 유효, 단계)
    Set 행맵 = 조건적용(행맵, 유효, 컬럼들)

    결과_그리기 ws, 유효, 컬럼들, 행맵
    Application.ScreenUpdating = True

    상태쓰기 라벨(기준) & " 기준 " & 행맵.Count & "행 · 연결 표 " & (유효.Count - 1) & _
             "개 · 출력 컬럼 " & 컬럼들.Count & "개" & _
             "   (칸을 고친 뒤 [변경 저장])"
End Sub

' 기준 표의 각 행에서 시작해 유효표를 순서대로 붙입니다 (상대 없으면 0)
Private Function 조인전개(ByVal 기준 As String, ByVal 유효 As Collection, _
                          ByVal 단계 As Object) As Collection
    Dim 결과 As New Collection, 임시 As Collection
    Dim wsB As Worksheet, pk열 As Long, r As Long, 끝 As Long
    Dim k As Long, 표이름 As String, 정보 As Variant
    Dim 앞표 As String, 방향 As String, 컬럼 As String
    Dim 앞위치 As Long, i As Long, 조각 As Variant
    Dim ws앞 As Worksheet, ws현 As Worksheet
    Dim 부모행 As Long, 자식 As Collection, j As Long
    Dim 새줄 As String, 넘침 As Boolean

    Set wsB = 시트(기준)
    pk열 = 헤더열(wsB, 표_PK(기준))
    끝 = 마지막행(wsB, pk열)
    For r = 2 To 끝
        If Len(Trim$(CStr(wsB.Cells(r, pk열).Value))) > 0 Then 결과.Add CStr(r)
    Next r

    For k = 2 To 유효.Count
        표이름 = CStr(유효(k))
        If Not 단계.Exists(표이름) Then GoTo 다음표
        정보 = Split(CStr(단계(표이름)), "|")
        앞표 = CStr(정보(0)): 방향 = CStr(정보(1)): 컬럼 = CStr(정보(2))
        앞위치 = 0
        For i = 1 To 유효.Count
            If CStr(유효(i)) = 앞표 Then 앞위치 = i
        Next i
        If 앞위치 = 0 Then GoTo 다음표

        Set ws앞 = 시트(앞표)
        Set ws현 = 시트(표이름)
        Set 임시 = New Collection
        For i = 1 To 결과.Count
            조각 = Split(CStr(결과(i)), "|")
            부모행 = CLng(Val(조각(앞위치 - 1)))
            Set 자식 = New Collection
            If 부모행 > 0 Then
                If 방향 = "child" Then
                    Set 자식 = 자식행찾기(ws현, 컬럼, 셀값(ws앞, 부모행, 표_PK(앞표)))
                Else
                    Dim v As String, rr As Long
                    v = 셀값(ws앞, 부모행, 컬럼)
                    rr = PK행찾기(ws현, v)
                    If rr > 0 Then 자식.Add CStr(rr)
                End If
            End If
            If 자식.Count = 0 Then
                임시.Add CStr(결과(i)) & "|0"
            Else
                For j = 1 To 자식.Count
                    If 임시.Count >= MAX_ROWS Then
                        넘침 = True
                        Exit For
                    End If
                    임시.Add CStr(결과(i)) & "|" & 자식(j)
                Next j
            End If
            If 넘침 Then Exit For
        Next i
        Set 결과 = 임시
다음표:
    Next k

    If 넘침 Then
        MsgBox "결과가 " & MAX_ROWS & "행을 넘어 잘랐습니다." & vbCrLf & _
               "연결할 표를 줄이거나 조건으로 좁혀주세요.", vbExclamation, "조회 실행"
    End If
    Set 조인전개 = 결과
End Function

Private Function 자식행찾기(ByVal ws As Worksheet, ByVal FK컬럼 As String, _
                            ByVal 값 As String) As Collection
    Dim c As New Collection, j As Long, pk열 As Long, r As Long, 끝 As Long
    Set 자식행찾기 = c
    If ws Is Nothing Or Len(값) = 0 Then Exit Function
    j = 헤더열(ws, FK컬럼)
    pk열 = 헤더열(ws, 표_PK(ws.Name))
    If j = 0 Or pk열 = 0 Then Exit Function
    끝 = 마지막행(ws, pk열)
    For r = 2 To 끝
        If Trim$(CStr(ws.Cells(r, j).Value)) = 값 Then c.Add CStr(r)
    Next r
End Function

' 누락 필터 + 검색어
Private Function 조건적용(ByVal 행맵 As Collection, ByVal 유효 As Collection, _
                          ByVal 컬럼들 As Collection) As Collection
    Dim ws As Worksheet, 결과 As New Collection
    Dim 필터 As String, 검색 As String
    Dim i As Long, k As Long, 조각 As Variant
    Dim 빠짐 As Boolean, 걸림 As Boolean
    Dim c As Long, 조 As Variant, 표이름 As String, 위치 As Long, 원본 As Long

    Set ws = 시트(EXP_SHEET)
    필터 = Trim$(CStr(ws.Cells(ROW_COND, COL_VAL).Value))
    검색 = Trim$(CStr(ws.Cells(ROW_COND, COL_VAL + 4).Value))

    For i = 1 To 행맵.Count
        조각 = Split(CStr(행맵(i)), "|")
        빠짐 = False
        For k = 2 To 유효.Count
            If CLng(Val(조각(k - 1))) = 0 Then 빠짐 = True
        Next k

        If 필터 = F_ANY And Not 빠짐 Then GoTo 다음
        If 필터 = F_NONE And 빠짐 Then GoTo 다음
        If Left$(필터, 1) = "[" Then
            ' [표이름] 이(가) 없는 행만
            Dim 대상 As String
            대상 = Mid$(필터, 2, InStr(필터, "]") - 2)
            위치 = 0
            For k = 1 To 유효.Count
                If 라벨(CStr(유효(k))) = 대상 Then 위치 = k
            Next k
            If 위치 = 0 Then GoTo 다음
            If CLng(Val(조각(위치 - 1))) <> 0 Then GoTo 다음
        End If

        If Len(검색) > 0 Then
            걸림 = False
            For c = 1 To 컬럼들.Count
                조 = Split(CStr(컬럼들(c)), "^")
                표이름 = CStr(조(0))
                위치 = 0
                For k = 1 To 유효.Count
                    If CStr(유효(k)) = 표이름 Then 위치 = k
                Next k
                If 위치 > 0 Then
                    원본 = CLng(Val(조각(위치 - 1)))
                    If 원본 > 0 Then
                        If InStr(1, 셀값(시트(표이름), 원본, CStr(조(1))), 검색, vbTextCompare) > 0 Then
                            걸림 = True
                            Exit For
                        End If
                    End If
                End If
            Next c
            If Not 걸림 Then GoTo 다음
        End If

        결과.Add CStr(행맵(i))
다음:
    Next i
    Set 조건적용 = 결과
End Function

'==============================================================================
' 결과 그리기
'==============================================================================
Private Sub 결과_그리기(ByVal ws As Worksheet, ByVal 유효 As Collection, _
                        ByVal 컬럼들 As Collection, ByVal 행맵 As Collection)
    Dim wsM As Worksheet, i As Long, c As Long, r As Long
    Dim 조 As Variant, 표이름 As String, 컬럼 As String, 위치 As Long
    Dim 조각 As Variant, 원본 As Long, k As Long
    Dim 마지막열 As Long, 기준 As String

    기준 = 기준표()
    결과_지우기 ws

    마지막열 = COL_VAL + 컬럼들.Count - 1

    ' 머리글 (표 이름 + 컬럼 이름 2줄 효과)
    For c = 1 To 컬럼들.Count
        조 = Split(CStr(컬럼들(c)), "^")
        표이름 = CStr(조(0)): 컬럼 = CStr(조(1))
        With ws.Cells(ROW_HEAD, COL_VAL + c - 1)
            .Value = 라벨(표이름) & Chr$(10) & 컬럼
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .Font.Bold = True
            .Font.Size = 9
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(190, 186, 172)
            If 컬럼 = 표_PK(표이름) Then
                .Interior.Color = RGB(14, 124, 123)
                .Font.Color = RGB(255, 255, 255)
            ElseIf Len(FK_참조(표이름, 컬럼)) > 0 Then
                .Interior.Color = RGB(238, 237, 254)
                .Font.Color = RGB(83, 74, 183)
            Else
                .Interior.Color = RGB(239, 238, 231)
                .Font.Color = RGB(70, 70, 70)
            End If
        End With
        ws.Columns(COL_VAL + c - 1).ColumnWidth = 시각폭(컬럼)
    Next c
    ws.Rows(ROW_HEAD).RowHeight = 30
    ws.Cells(ROW_HEAD, COL_LBL).Value = "결과"
    ws.Cells(ROW_HEAD, COL_LBL).Font.Bold = True

    ' 데이터 + 원본 행 대응표
    Set wsM = 맵시트()
    wsM.Cells.ClearContents
    wsM.Cells(1, 1).Value = 유효표문자열(유효)
    wsM.Cells(2, 1).Value = 컬럼문자열(컬럼들)

    For i = 1 To 행맵.Count
        r = ROW_DATA + i - 1
        조각 = Split(CStr(행맵(i)), "|")
        ws.Cells(r, COL_KEY).Value = "R" & i
        For k = 1 To 유효.Count
            wsM.Cells(i + 3, k).Value = CLng(Val(조각(k - 1)))
        Next k

        For c = 1 To 컬럼들.Count
            조 = Split(CStr(컬럼들(c)), "^")
            표이름 = CStr(조(0)): 컬럼 = CStr(조(1))
            위치 = 0
            For k = 1 To 유효.Count
                If CStr(유효(k)) = 표이름 Then 위치 = k
            Next k
            원본 = 0
            If 위치 > 0 Then 원본 = CLng(Val(조각(위치 - 1)))

            With ws.Cells(r, COL_VAL + c - 1)
                .NumberFormatLocal = "@"
                .Font.Size = 9.5
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(226, 223, 212)
                If 원본 = 0 Then
                    .Value = MISS_MARK
                    .Interior.Color = RGB(250, 238, 218)      ' 누락 - 편집 불가
                    .Font.Color = RGB(176, 120, 23)
                    .HorizontalAlignment = xlCenter
                ElseIf 컬럼 = 표_PK(표이름) Then
                    .Value = 셀값(시트(표이름), 원본, 컬럼)
                    .Interior.Color = RGB(240, 245, 244)      ' PK - 편집 불가
                    .Font.Color = RGB(10, 90, 89)
                    .Font.Bold = True
                Else
                    .Value = 셀값(시트(표이름), 원본, 컬럼)
                    .Interior.Color = RGB(255, 255, 255)      ' 편집 가능
                    .Font.Color = RGB(22, 38, 43)
                End If
            End With
        Next c
    Next i

    ' 보기 편의: 자동필터 · 틀 고정
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If 행맵.Count > 0 Then
        ws.Range(ws.Cells(ROW_HEAD, COL_VAL), ws.Cells(ROW_DATA + 행맵.Count - 1, 마지막열)).AutoFilter
    End If
    ws.Activate
    ws.Cells(ROW_HEAD, COL_LBL).Select
    On Error GoTo 0

    If 행맵.Count = 0 Then
        ws.Cells(ROW_DATA, COL_VAL).Value = "조건에 맞는 데이터가 없습니다."
        ws.Cells(ROW_DATA, COL_VAL).Font.Color = RGB(120, 120, 120)
    End If

    ' 누락 필터 드롭다운을 현재 연결 표에 맞춰 갱신
    누락필터갱신 ws, 유효
End Sub

Private Function 시각폭(ByVal 컬럼 As String) As Double
    시각폭 = Len(컬럼) * 1.6 + 6
    If 시각폭 < 11 Then 시각폭 = 11
    If 시각폭 > 26 Then 시각폭 = 26
End Function

Private Sub 결과_지우기(ByVal ws As Worksheet)
    Dim 끝행 As Long, 끝열 As Long
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
    끝행 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    If 끝행 < ROW_HEAD Then 끝행 = ROW_HEAD
    끝열 = ws.Cells(ROW_HEAD, ws.Columns.Count).End(xlToLeft).Column
    If 끝열 < COL_VAL Then 끝열 = COL_VAL
    ws.Range(ws.Rows(ROW_HEAD), ws.Rows(끝행 + 3)).Clear
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
' 변경 저장 (그리드에서 고친 값을 원본 시트에 반영)
'==============================================================================
Public Sub 탐색_변경저장()
    Dim ws As Worksheet, wsM As Worksheet
    Dim 유효 As Variant, 컬럼들 As Variant
    Dim i As Long, c As Long, k As Long, r As Long
    Dim 조 As Variant, 표이름 As String, 컬럼 As String, 위치 As Long, 원본 As Long
    Dim 화면값 As String, 원래값 As String
    Dim 변경 As Object, 미리보기 As String, 경고 As String, 오류 As String
    Dim 키 As String, 답 As VbMsgBoxResult, 건수 As Long

    Set ws = 시트(EXP_SHEET)
    Set wsM = 시트(MAP_SHEET)
    If ws Is Nothing Or wsM Is Nothing Then
        MsgBox "먼저 [조회 실행]으로 결과를 만들어주세요.", vbInformation, "변경 저장"
        Exit Sub
    End If
    If Len(Trim$(CStr(wsM.Cells(1, 1).Value))) = 0 Then
        MsgBox "먼저 [조회 실행]으로 결과를 만들어주세요.", vbInformation, "변경 저장"
        Exit Sub
    End If

    유효 = Split(CStr(wsM.Cells(1, 1).Value), ";")
    컬럼들 = Split(CStr(wsM.Cells(2, 1).Value), ";")
    Set 변경 = CreateObject("Scripting.Dictionary")

    i = 0
    Do
        i = i + 1
        r = ROW_DATA + i - 1
        If CStr(ws.Cells(r, COL_KEY).Value) <> "R" & i Then Exit Do

        For c = LBound(컬럼들) To UBound(컬럼들)
            조 = Split(CStr(컬럼들(c)), "^")
            표이름 = CStr(조(0)): 컬럼 = CStr(조(1))
            위치 = 0
            For k = LBound(유효) To UBound(유효)
                If CStr(유효(k)) = 표이름 Then 위치 = k + 1
            Next k
            원본 = 0
            If 위치 > 0 Then 원본 = CLng(Val(wsM.Cells(i + 3, 위치).Value))

            화면값 = Trim$(CStr(ws.Cells(r, COL_VAL + c - LBound(컬럼들)).Value))

            If 원본 = 0 Then
                If Len(화면값) > 0 And 화면값 <> MISS_MARK Then
                    경고 = 경고 & "· " & r & "행 " & 라벨(표이름) & "." & 컬럼 & _
                           " : 연결된 " & 라벨(표이름) & " 레코드가 없어 저장할 수 없습니다." & vbCrLf
                End If
            Else
                원래값 = 셀값(시트(표이름), 원본, 컬럼)
                If 화면값 <> 원래값 Then
                    If 컬럼 = 표_PK(표이름) Then
                        오류 = 오류 & "· " & r & "행 " & 라벨(표이름) & "." & 컬럼 & _
                               " : PK(ID)는 여기서 바꿀 수 없습니다 (" & 원래값 & " → " & 화면값 & ")" & vbCrLf
                    Else
                        키 = 표이름 & "^" & 원본 & "^" & 컬럼
                        If 변경.Exists(키) Then
                            If Split(CStr(변경(키)), vbTab)(1) <> 화면값 Then
                                오류 = 오류 & "· " & 라벨(표이름) & " " & _
                                       셀값(시트(표이름), 원본, 표_PK(표이름)) & " 의 " & 컬럼 & _
                                       " 을(를) 서로 다른 값으로 고쳤습니다 (" & _
                                       Split(CStr(변경(키)), vbTab)(1) & " / " & 화면값 & ")" & vbCrLf
                            End If
                        Else
                            변경(키) = 원래값 & vbTab & 화면값
                        End If
                    End If
                End If
            End If
        Next c
    Loop While i < MAX_ROWS

    ' FK 로 고친 값이 실제로 있는지 확인
    Dim ky As Variant, 조각2 As Variant, 참조 As String, 새값 As String
    For Each ky In 변경.Keys
        조각2 = Split(CStr(ky), "^")
        참조 = FK_참조(CStr(조각2(0)), CStr(조각2(2)))
        If Len(참조) > 0 Then
            새값 = Split(CStr(변경(ky)), vbTab)(1)
            If Len(새값) = 0 Then
                오류 = 오류 & "· " & 라벨(CStr(조각2(0))) & "." & 조각2(2) & " 는 비워둘 수 없습니다." & vbCrLf
            ElseIf PK행찾기(시트(참조), 새값) = 0 Then
                오류 = 오류 & "· " & 라벨(CStr(조각2(0))) & "." & 조각2(2) & " = '" & 새값 & _
                       "' 에 해당하는 " & 라벨(참조) & " 가 없습니다." & vbCrLf
            End If
        End If
    Next ky

    If Len(오류) > 0 Then
        MsgBox "저장할 수 없습니다." & vbCrLf & vbCrLf & 오류, vbExclamation, "변경 저장"
        Exit Sub
    End If
    If 변경.Count = 0 Then
        MsgBox "바뀐 내용이 없습니다." & IIf(Len(경고) > 0, vbCrLf & vbCrLf & 경고, ""), _
               vbInformation, "변경 저장"
        Exit Sub
    End If

    건수 = 0
    For Each ky In 변경.Keys
        조각2 = Split(CStr(ky), "^")
        건수 = 건수 + 1
        If 건수 <= 40 Then
            미리보기 = 미리보기 & "· " & 라벨(CStr(조각2(0))) & " " & _
                       셀값(시트(CStr(조각2(0))), CLng(조각2(1)), 표_PK(CStr(조각2(0)))) & _
                       " / " & 조각2(2) & " : " & _
                       IIf(Len(Split(CStr(변경(ky)), vbTab)(0)) = 0, "(공란)", Split(CStr(변경(ky)), vbTab)(0)) & _
                       " → " & Split(CStr(변경(ky)), vbTab)(1) & vbCrLf
        End If
    Next ky
    If 건수 > 40 Then 미리보기 = 미리보기 & "  … 외 " & (건수 - 40) & "건" & vbCrLf

    답 = MsgBox("아래 " & 건수 & "건을 원본 시트에 반영합니다." & vbCrLf & vbCrLf & _
                미리보기 & IIf(Len(경고) > 0, vbCrLf & "[저장 안 되는 항목]" & vbCrLf & 경고, "") & _
                vbCrLf & "계속할까요?", vbQuestion + vbYesNo, "변경 저장 확인")
    If 답 <> vbYes Then
        상태쓰기 "저장을 취소했습니다.", True
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For Each ky In 변경.Keys
        조각2 = Split(CStr(ky), "^")
        셀쓰기 시트(CStr(조각2(0))), CLng(조각2(1)), CStr(조각2(2)), Split(CStr(변경(ky)), vbTab)(1)
    Next ky
    Application.ScreenUpdating = True

    상태쓰기 건수 & "건을 저장했습니다. (" & Format$(Now, "hh:nn:ss") & ")"
    MsgBox 건수 & "건을 원본 시트에 반영했습니다.", vbInformation, "변경 저장"
    탐색_실행
End Sub

' 고치던 것을 버리고 원본 값으로 다시 그리기
Public Sub 탐색_되돌리기()
    Dim 답 As VbMsgBoxResult
    답 = MsgBox("화면에서 고친 내용을 버리고 원본 값으로 다시 불러옵니다." & vbCrLf & _
                "계속할까요?", vbQuestion + vbYesNo, "되돌리기")
    If 답 <> vbYes Then Exit Sub
    탐색_실행
    상태쓰기 "원본 값으로 되돌렸습니다."
End Sub

'==============================================================================
' 빠른 조회 프리셋 (스키마의 FK 에서 자동으로 뽑은 "빠진 것 찾기")
'==============================================================================
Private Function 프리셋목록() As Collection
    Dim c As New Collection, t As Variant, cols As Variant, i As Long, 참조 As String
    For Each t In 표_순서()
        cols = 표_컬럼목록(CStr(t))
        For i = LBound(cols) To UBound(cols)
            참조 = FK_참조(CStr(t), CStr(cols(i)))
            If Len(참조) > 0 Then
                ' 기준표|연결표|표시이름
                c.Add 참조 & "|" & CStr(t) & "|" & 라벨(참조) & " 중 " & 라벨(CStr(t)) & " 없음"
            End If
        Next i
    Next t
    Set 프리셋목록 = c
End Function

' 기준 표 / 연결 표 / 누락 필터를 프리셋에 맞춰 세팅
Public Sub 탐색_프리셋적용()
    Dim ws As Worksheet, c As Collection, i As Long, 고른 As String
    Dim 기준 As String, 연결 As String, cb As Object

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

    ws.Cells(ROW_BASE, COL_VAL).Value = 라벨(기준)
    체크박스_다시그리기
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If cb.Name = "TBL^" & 연결 Then cb.Value = xlOn
    Next cb
    On Error GoTo 0
    체크박스_다시그리기          ' 유효표가 바뀌었으니 컬럼 줄도 다시
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If cb.Name = "COL^" & 연결 & "^" & 표_PK(연결) Then cb.Value = xlOn
    Next cb
    On Error GoTo 0
    ws.Cells(ROW_COND, COL_VAL).Value = "[" & 라벨(연결) & "] 이(가) 없는 행만"
    탐색_실행
    상태쓰기 고른 & " - " & ws.Cells(ROW_STATUS, COL_VAL).Value
End Sub

'==============================================================================
' 내보내기 / 초기화 / 상태
'==============================================================================
Public Sub 탐색_내보내기()
    Dim ws As Worksheet, 새책 As Workbook, 끝행 As Long, 끝열 As Long
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    끝행 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    If 끝행 < ROW_DATA Then
        MsgBox "내보낼 결과가 없습니다. 먼저 [조회 실행]을 눌러주세요.", vbInformation, "내보내기"
        Exit Sub
    End If
    끝열 = ws.Cells(ROW_HEAD, ws.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False
    Set 새책 = Workbooks.Add
    ws.Range(ws.Cells(ROW_HEAD, COL_VAL), ws.Cells(끝행, 끝열)).Copy
    새책.Worksheets(1).Range("A1").PasteSpecial xlPasteValues
    새책.Worksheets(1).Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    새책.Worksheets(1).Rows(1).RowHeight = 30
    새책.Worksheets(1).Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "새 통합 문서로 내보냈습니다. 원하는 이름으로 저장하세요.", vbInformation, "내보내기"
End Sub

Public Sub 탐색_초기화()
    Dim ws As Worksheet
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    결과_지우기 ws
    ws.Cells(ROW_COND, COL_VAL + 4).ClearContents
    ws.Cells(ROW_COND, COL_VAL).Value = F_ALL
    ws.Cells(ROW_PRESET, COL_VAL).Value = P_NONE
    On Error Resume Next
    맵시트().Cells.ClearContents
    On Error GoTo 0
    상태쓰기 "초기화했습니다."
End Sub

Private Sub 상태쓰기(ByVal msg As String, Optional ByVal 경고 As Boolean = False)
    Dim ws As Worksheet
    Set ws = 시트(EXP_SHEET)
    If ws Is Nothing Then Exit Sub
    With ws.Cells(ROW_STATUS, COL_VAL)
        .Value = msg
        .Font.Bold = True
        .Font.Color = IIf(경고, RGB(178, 58, 58), RGB(31, 122, 84))
    End With
End Sub

'==============================================================================
' (선택) 시트 이벤트 연결용 - 탐색 시트 코드창에 아래를 붙여넣으면
' 기준 표나 빠른 조회를 고르는 즉시 반영됩니다.
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       탐색_변경감지 Target
'   End Sub
'==============================================================================
Public Sub 탐색_변경감지(ByVal Target As Range)
    Dim 이전 As Boolean
    If Target.Worksheet.Name <> EXP_SHEET Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Column <> COL_VAL Then Exit Sub

    이전 = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    If Target.Row = ROW_BASE Then
        체크박스_다시그리기
        상태쓰기 "기준 표를 바꿨습니다. 컬럼을 확인하고 [조회 실행]을 누르세요."
    ElseIf Target.Row = ROW_PRESET Then
        탐색_프리셋적용
    End If
    On Error GoTo 0
    Application.EnableEvents = 이전
End Sub
