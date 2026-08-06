Attribute VB_Name = "PPA_통합입력폼"
'==============================================================================
' PPA 통합 입력/조회 폼
'------------------------------------------------------------------------------
' 목적
'   지금은 한 건의 공급-수요 건을 등록하려면 T_발전소 → T_구매계약 → T_수급매칭 →
'   T_전기사용지 → T_판매계약 → T_수요기업 시트를 오가며 ID를 손으로 맞춰 넣어야
'   합니다. 이 모듈은 "입력폼" 시트 하나에서 그 6개 표를 한 번에 입력/조회/수정할
'   수 있게 해줍니다.
'
' 특징
'   - UserForm(.frm)을 쓰지 않습니다. 이 모듈(.bas) 하나만 가져오면 폼 시트를
'     스스로 만들어냅니다 (VBE 폼 디자이너 작업 불필요).
'   - ID 칸은 드롭다운에서 고르고, 고르면 나머지 항목이 자동으로 채워집니다.
'     발전소를 고르면 구매계약 목록이 그 발전소 것만 남는 식으로 연동됩니다.
'   - 저장 전에 "무엇이 새로 생기고 무엇이 어떻게 바뀌는지"를 먼저 보여주고
'     확인을 받습니다. 실수로 기존 데이터를 덮어쓰는 일을 막기 위한 장치입니다.
'   - 저장 시 기존 매크로의 검증 규칙(PK 공란/중복, FK 참조, 조합중복)을 그대로
'     적용합니다.
'   - 폼에서 비워 둔 칸은 기존 값을 지우지 않습니다(실수로 값이 날아가는 것 방지).
'
' 사용 순서
'   1) VBE(Alt+F11) → 파일 → 파일 가져오기 → 이 .bas 선택
'   2) Alt+F8 → "폼_만들기" 실행  (입력폼 시트가 생성됩니다)
'   3) 이후로는 입력폼 시트 위쪽의 버튼만 누르면 됩니다.
'
'   * 드롭다운을 고르는 즉시 자동으로 채워지게 하려면(선택 사항)
'     입력폼 시트 코드창에 아래 3줄을 붙여넣으세요.
'         Private Sub Worksheet_Change(ByVal Target As Range)
'             폼_변경감지 Target
'         End Sub
'     붙여넣지 않아도 [불러오기] 버튼으로 동일하게 동작합니다.
'==============================================================================
Option Explicit

'---- 시트 이름 --------------------------------------------------------------
Private Const FORM_SHEET  As String = "입력폼"
Private Const LIST_SHEET  As String = "_폼목록"      ' 드롭다운 원본 + 조회결과(숨김)

Private Const SH_PLANT    As String = "T_발전소"
Private Const SH_BUY      As String = "T_구매계약"
Private Const SH_DEMAND   As String = "T_수요기업"
Private Const SH_SELL     As String = "T_판매계약"
Private Const SH_SITE     As String = "T_전기사용지"
Private Const SH_MATCH    As String = "T_수급매칭"

'---- 폼 레이아웃 ------------------------------------------------------------
' A열: 필드키("시트명|컬럼명", 숨김) / B열: 항목명 / C열: 값 / D열: 안내
Private Const COL_KEY   As Long = 1
Private Const COL_LABEL As Long = 2
Private Const COL_VALUE As Long = 3
Private Const COL_NOTE  As Long = 4

Private Const ROW_SEARCH As Long = 2   ' 검색어 입력칸 (C2)
Private Const ROW_STATUS As Long = 3   ' 상태 표시줄 (C3)
Private Const FIRST_FIELD_ROW As Long = 5

'---- 숨김 목록 시트의 열 배치 -----------------------------------------------
Private Const LC_PLANT As Long = 1     ' 표별 전체 ID 목록
Private Const LC_BUY   As Long = 2
Private Const LC_DEM   As Long = 3
Private Const LC_SELL  As Long = 4
Private Const LC_SITE  As Long = 5
Private Const LC_MATCH As Long = 6
Private Const LC_BUY_F  As Long = 8    ' 상위 선택에 따라 좁혀진 목록
Private Const LC_SELL_F As Long = 9
Private Const LC_SITE_F As Long = 10
Private Const LC_RESULT As Long = 12   ' 조회 결과(수급매칭 행번호)
Private Const LC_POS    As Long = 13   ' 조회 결과 내 현재 위치

'==============================================================================
' 스키마 정의 — static-dashboard/ppa_schema.py, backend의 tableDefs.ts와 동일
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
    표_PK = CStr(c(0))                  ' 6개 표 모두 첫 컬럼이 PK
End Function

' 이 폼이 다루는 표를 입력 순서대로 (부모 → 자식)
Private Function 표_순서() As Variant
    표_순서 = Array(SH_PLANT, SH_BUY, SH_DEMAND, SH_SELL, SH_SITE, SH_MATCH)
End Function

' FK: 해당 컬럼이 참조하는 시트 (아니면 "")
Private Function FK_참조시트(ByVal 시트명 As String, ByVal 컬럼 As String) As String
    Select Case 시트명 & "|" & 컬럼
    Case SH_BUY & "|발전소ID":       FK_참조시트 = SH_PLANT
    Case SH_SELL & "|수요기업ID":    FK_참조시트 = SH_DEMAND
    Case SH_SITE & "|판매계약ID":    FK_참조시트 = SH_SELL
    Case SH_MATCH & "|전기사용지ID": FK_참조시트 = SH_SITE
    Case SH_MATCH & "|구매계약ID":   FK_참조시트 = SH_BUY
    Case Else:                       FK_참조시트 = ""
    End Select
End Function

' ID 성격의 컬럼인가 (숫자로 바꾸면 안 되는 값)
Private Function ID컬럼인가(ByVal 시트명 As String, ByVal 컬럼 As String) As Boolean
    ID컬럼인가 = (컬럼 = 표_PK(시트명)) Or (Len(FK_참조시트(시트명, 컬럼)) > 0)
End Function

'==============================================================================
' 공통 도우미
'==============================================================================
Private Function 시트(ByVal 이름 As String) As Worksheet
    On Error Resume Next
    Set 시트 = ThisWorkbook.Worksheets(이름)
    On Error GoTo 0
End Function

' 헤더 텍스트로 열 번호 찾기 (열 순서가 바뀌어도 안전)
Private Function 헤더열(ByVal ws As Worksheet, ByVal 헤더 As String) As Long
    Dim 마지막열 As Long, j As Long
    If ws Is Nothing Then Exit Function
    마지막열 = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To 마지막열
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

' PK 값으로 데이터 행 찾기 (없으면 0)
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

' 셀 값을 화면 표기용 문자열로 (날짜 → yyyy-mm-dd, 불린 → TRUE/FALSE)
Private Function 값을문자열(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        값을문자열 = ""
    ElseIf VarType(v) = vbBoolean Then
        값을문자열 = IIf(v, "TRUE", "FALSE")
    ElseIf IsDate(v) Then
        값을문자열 = Format$(v, "yyyy-mm-dd")
    Else
        값을문자열 = Trim$(CStr(v))
    End If
End Function

Private Function 셀값(ByVal ws As Worksheet, ByVal r As Long, ByVal 헤더 As String) As String
    Dim j As Long
    If ws Is Nothing Or r < 2 Then Exit Function
    j = 헤더열(ws, 헤더)
    If j = 0 Then Exit Function
    셀값 = 값을문자열(ws.Cells(r, j).Value)
End Function

' 값 쓰기 — 그 열에 이미 들어있는 데이터의 형식을 따라갑니다
' (불린 열에는 True/False, 날짜 열에는 진짜 날짜로). ID 컬럼은 항상 문자열.
Private Sub 셀쓰기(ByVal ws As Worksheet, ByVal r As Long, ByVal 헤더 As String, ByVal v As String)
    Dim j As Long, 견본 As Variant, rr As Long, pk열 As Long, 끝 As Long
    Dim 견본있음 As Boolean

    j = 헤더열(ws, 헤더)
    If j = 0 Then Exit Sub

    If Len(v) = 0 Then
        ws.Cells(r, j).ClearContents
        Exit Sub
    End If

    ' ID(PK/FK)는 "001" 같은 값이 숫자 1로 바뀌지 않도록 항상 문자열로
    If ID컬럼인가(ws.Name, 헤더) Then
        ws.Cells(r, j).NumberFormatLocal = "@"
        ws.Cells(r, j).Value = v
        Exit Sub
    End If

    ' 같은 열에서 비어있지 않은 기존 값 하나를 찾아 형식 판단
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
    ElseIf IsDate(v) And InStr(1, 헤더, "일") + InStr(1, 헤더, "기한") > 0 Then
        ws.Cells(r, j).Value = CDate(v)
        ws.Cells(r, j).NumberFormatLocal = "yyyy-mm-dd"
    Else
        ws.Cells(r, j).Value = v
    End If
End Sub

'==============================================================================
' 폼 만들기
'==============================================================================
Public Sub 폼_만들기()
    Dim ws As Worksheet, r As Long, t As Variant, cols As Variant, i As Long
    Dim 표이름 As String

    If Not 원본표_점검() Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo 정리

    Set ws = 시트(FORM_SHEET)
    If Not ws Is Nothing Then ws.Delete
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = FORM_SHEET

    With ws.Cells
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
    End With
    ws.Columns(COL_KEY).ColumnWidth = 0.1        ' 필드키(사실상 숨김)
    ws.Columns(COL_LABEL).ColumnWidth = 22
    ws.Columns(COL_VALUE).ColumnWidth = 34
    ws.Columns(COL_NOTE).ColumnWidth = 48

    With ws.Cells(1, COL_LABEL)
        .Value = "PPA 통합 입력 · 조회 폼"
        .Font.Size = 15
        .Font.Bold = True
    End With

    ws.Cells(ROW_SEARCH, COL_LABEL).Value = "조회 (ID / 이름)"
    ws.Cells(ROW_SEARCH, COL_LABEL).Font.Bold = True
    With ws.Cells(ROW_SEARCH, COL_VALUE)
        .Interior.Color = RGB(255, 249, 219)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 196, 180)
        .NumberFormatLocal = "@"
    End With
    ws.Cells(ROW_SEARCH, COL_NOTE).Value = "값을 넣고 [조회]. 여러 건이면 [이전]/[다음]으로 넘깁니다."
    ws.Cells(ROW_SEARCH, COL_NOTE).Font.Color = RGB(110, 110, 110)

    ws.Cells(ROW_STATUS, COL_LABEL).Value = "상태"
    ws.Cells(ROW_STATUS, COL_LABEL).Font.Bold = True

    ' 표별 입력 항목 배치
    r = FIRST_FIELD_ROW
    For Each t In 표_순서()
        표이름 = CStr(t)
        ws.Cells(r, COL_LABEL).Value = "■ " & Replace(표이름, "T_", "")
        With ws.Range(ws.Cells(r, COL_LABEL), ws.Cells(r, COL_NOTE))
            .Interior.Color = RGB(14, 124, 123)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
        r = r + 1

        cols = 표_컬럼목록(표이름)
        For i = LBound(cols) To UBound(cols)
            ws.Cells(r, COL_KEY).Value = 표이름 & "|" & cols(i)
            ws.Cells(r, COL_LABEL).Value = cols(i)
            With ws.Cells(r, COL_VALUE)
                .Interior.Color = RGB(255, 255, 255)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(214, 210, 196)
                .HorizontalAlignment = xlLeft
                .NumberFormatLocal = "@"          ' 폼에서는 전부 텍스트로 다룸
            End With
            If i = 0 Then
                ws.Cells(r, COL_LABEL).Font.Bold = True
                ws.Cells(r, COL_NOTE).Value = "PK · 비워두면 저장할 때 자동 생성"
            ElseIf Len(FK_참조시트(표이름, CStr(cols(i)))) > 0 Then
                ws.Cells(r, COL_LABEL).Font.Color = RGB(83, 74, 183)
                ws.Cells(r, COL_NOTE).Value = "FK → " & Replace(FK_참조시트(표이름, CStr(cols(i))), "T_", "")
            End If
            ws.Cells(r, COL_NOTE).Font.Color = RGB(110, 110, 110)
            r = r + 1
        Next i
        r = r + 1
    Next t

    폼_버튼만들기 ws
    ws.Rows(1).RowHeight = 26
    폼_목록갱신
    폼_초기화
    ws.Activate
    ws.Cells(ROW_SEARCH, COL_VALUE).Select

정리:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "폼을 만드는 중 오류가 발생했습니다." & vbCrLf & Err.Description, vbExclamation
        Exit Sub
    End If

    MsgBox "입력폼 시트를 만들었습니다." & vbCrLf & vbCrLf & _
           "· ID 칸의 드롭다운에서 기존 항목을 고르고 [불러오기]를 누르면 나머지가 채워집니다." & vbCrLf & _
           "· 새 건이면 값만 채우고 [저장]을 누르세요. 저장 전에 확인 창이 뜹니다.", _
           vbInformation, "PPA 통합 입력폼"
End Sub

Private Sub 폼_버튼만들기(ByVal ws As Worksheet)
    Dim b As Object, 정의 As Variant, i As Long, x As Double

    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0

    정의 = Array("조회|폼_조회", "이전|폼_이전", "다음|폼_다음", "불러오기|폼_불러오기", _
                 "새로 만들기|폼_초기화", "저장|폼_저장", "삭제|폼_삭제", "목록 새로고침|폼_목록갱신")
    x = ws.Cells(1, COL_VALUE).Left
    For i = LBound(정의) To UBound(정의)
        Set b = ws.Buttons.Add(x, ws.Cells(1, 1).Top + 2, 74, 22)
        b.Caption = Split(정의(i), "|")(0)
        b.OnAction = Split(정의(i), "|")(1)
        x = x + 78
    Next i
End Sub

' 원본 표 시트와 헤더가 제대로 있는지 미리 확인
Private Function 원본표_점검() As Boolean
    Dim t As Variant, cols As Variant, i As Long, ws As Worksheet
    Dim 문제 As String

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
        MsgBox "표 구성이 예상과 달라 폼을 만들 수 없습니다." & vbCrLf & vbCrLf & 문제 & vbCrLf & _
               "시트 이름과 1행 머리글을 확인해주세요.", vbExclamation, "PPA 통합 입력폼"
        원본표_점검 = False
    Else
        원본표_점검 = True
    End If
End Function

'==============================================================================
' 폼 값 읽기/쓰기
'==============================================================================
Private Function 필드행(ByVal ws As Worksheet, ByVal 필드키 As String) As Long
    Dim r As Long, 끝 As Long
    If ws Is Nothing Then Exit Function
    끝 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For r = FIRST_FIELD_ROW To 끝
        If CStr(ws.Cells(r, COL_KEY).Value) = 필드키 Then
            필드행 = r
            Exit Function
        End If
    Next r
End Function

Private Function 폼값(ByVal 시트명 As String, ByVal 컬럼 As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Function
    r = 필드행(ws, 시트명 & "|" & 컬럼)
    If r = 0 Then Exit Function
    폼값 = Trim$(CStr(ws.Cells(r, COL_VALUE).Value))
End Function

' 폼에 값을 넣을 때는 Change 이벤트가 다시 돌지 않도록 잠급니다.
' (이전 상태를 저장해뒀다 돌려놓아야 중첩 호출에서 꼬이지 않습니다)
Private Sub 폼값쓰기(ByVal 시트명 As String, ByVal 컬럼 As String, ByVal v As String)
    Dim ws As Worksheet, r As Long, 이전 As Boolean
    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Sub
    r = 필드행(ws, 시트명 & "|" & 컬럼)
    If r = 0 Then Exit Sub
    이전 = Application.EnableEvents
    Application.EnableEvents = False
    ws.Cells(r, COL_VALUE).Value = v
    Application.EnableEvents = 이전
End Sub

Private Sub 상태쓰기(ByVal msg As String, Optional ByVal 경고 As Boolean = False)
    Dim ws As Worksheet
    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Sub
    With ws.Cells(ROW_STATUS, COL_VALUE)
        .Value = msg
        .Font.Color = IIf(경고, RGB(178, 58, 58), RGB(31, 122, 84))
        .Font.Bold = True
    End With
End Sub

'==============================================================================
' 새로 만들기
'==============================================================================
Public Sub 폼_초기화()
    Dim ws As Worksheet, r As Long, 끝 As Long, 이전 As Boolean
    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Sub

    이전 = Application.EnableEvents
    Application.EnableEvents = False
    끝 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For r = FIRST_FIELD_ROW To 끝
        If Len(CStr(ws.Cells(r, COL_KEY).Value)) > 0 Then
            ws.Cells(r, COL_VALUE).ClearContents
            ws.Cells(r, COL_VALUE).Interior.Color = RGB(255, 255, 255)
        End If
    Next r
    ws.Cells(ROW_SEARCH, COL_VALUE).ClearContents
    Application.EnableEvents = 이전

    조회결과_지우기
    폼_기존신규표시
    상태쓰기 "새 입력 — 값을 채우고 [저장]을 누르세요."
End Sub

Private Sub 폼_값만지우기()
    Dim ws As Worksheet, r As Long, 끝 As Long, 이전 As Boolean
    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Sub
    이전 = Application.EnableEvents
    Application.EnableEvents = False
    끝 = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For r = FIRST_FIELD_ROW To 끝
        If Len(CStr(ws.Cells(r, COL_KEY).Value)) > 0 Then ws.Cells(r, COL_VALUE).ClearContents
    Next r
    Application.EnableEvents = 이전
End Sub

'==============================================================================
' 조회
'==============================================================================
Public Sub 폼_조회()
    Dim ws As Worksheet, wsL As Worksheet, 검색어 As String
    Dim 찾은행 As Collection, i As Long

    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Sub
    검색어 = Trim$(CStr(ws.Cells(ROW_SEARCH, COL_VALUE).Value))
    If Len(검색어) = 0 Then
        MsgBox "조회할 ID나 이름을 입력해주세요.", vbInformation, "조회"
        Exit Sub
    End If

    Set 찾은행 = 매칭검색(검색어)
    If 찾은행.Count = 0 Then
        조회결과_지우기
        상태쓰기 "'" & 검색어 & "' 로 찾은 항목이 없습니다.", True
        Exit Sub
    End If

    Set wsL = 목록시트()
    wsL.Columns(LC_RESULT).ClearContents
    For i = 1 To 찾은행.Count
        wsL.Cells(i, LC_RESULT).Value = 찾은행(i)
    Next i
    wsL.Cells(1, LC_POS).Value = 1
    폼_결과표시
End Sub

Public Sub 폼_다음()
    조회위치_이동 1
End Sub

Public Sub 폼_이전()
    조회위치_이동 -1
End Sub

Private Sub 조회위치_이동(ByVal 증감 As Long)
    Dim wsL As Worksheet, 개수 As Long, 현재 As Long
    Set wsL = 목록시트()
    개수 = 조회결과_개수()
    If 개수 = 0 Then
        상태쓰기 "먼저 [조회]를 실행해주세요.", True
        Exit Sub
    End If
    현재 = CLng(Val(wsL.Cells(1, LC_POS).Value)) + 증감
    If 현재 < 1 Then 현재 = 1
    If 현재 > 개수 Then 현재 = 개수
    wsL.Cells(1, LC_POS).Value = 현재
    폼_결과표시
End Sub

Private Function 조회결과_개수() As Long
    Dim wsL As Worksheet
    Set wsL = 목록시트()
    If wsL Is Nothing Then Exit Function
    조회결과_개수 = wsL.Cells(wsL.Rows.Count, LC_RESULT).End(xlUp).Row
    If Len(Trim$(CStr(wsL.Cells(1, LC_RESULT).Value))) = 0 Then 조회결과_개수 = 0
End Function

Private Sub 조회결과_지우기()
    Dim wsL As Worksheet
    Set wsL = 목록시트()
    If wsL Is Nothing Then Exit Sub
    wsL.Columns(LC_RESULT).ClearContents
    wsL.Cells(1, LC_POS).ClearContents
End Sub

' 검색: 수급매칭을 기준으로, 그와 이어진 상위 표의 ID/이름까지 뒤져서 찾습니다
Private Function 매칭검색(ByVal 검색어 As String) As Collection
    Dim 결과 As New Collection
    Dim wsM As Worksheet, wsSite As Worksheet, wsBuy As Worksheet
    Dim wsPlant As Worksheet, wsSell As Worksheet, wsDem As Worksheet
    Dim r As Long, 끝 As Long, pk열 As Long
    Dim 사용지ID As String, 구매ID As String, 판매ID As String, 수요ID As String, 발전소ID As String
    Dim 합침 As String, 키워드 As String
    Dim rs As Long, rb As Long, rp As Long, rl As Long, rd As Long

    키워드 = Trim$(검색어)
    Set wsM = 시트(SH_MATCH): Set wsSite = 시트(SH_SITE): Set wsBuy = 시트(SH_BUY)
    Set wsPlant = 시트(SH_PLANT): Set wsSell = 시트(SH_SELL): Set wsDem = 시트(SH_DEMAND)
    If wsM Is Nothing Then
        Set 매칭검색 = 결과
        Exit Function
    End If

    pk열 = 헤더열(wsM, 표_PK(SH_MATCH))
    끝 = 마지막행(wsM, pk열)

    For r = 2 To 끝
        사용지ID = 셀값(wsM, r, "전기사용지ID")
        구매ID = 셀값(wsM, r, "구매계약ID")
        합침 = 셀값(wsM, r, "수급매칭ID") & "|" & 사용지ID & "|" & 구매ID & "|" & 셀값(wsM, r, "현황")

        rs = PK행찾기(wsSite, 사용지ID)
        If rs > 0 Then
            합침 = 합침 & "|" & 셀값(wsSite, rs, "전기사용지명")
            판매ID = 셀값(wsSite, rs, "판매계약ID")
            rl = PK행찾기(wsSell, 판매ID)
            If rl > 0 Then
                합침 = 합침 & "|" & 판매ID & "|" & 셀값(wsSell, rl, "판매 담당자")
                수요ID = 셀값(wsSell, rl, "수요기업ID")
                rd = PK행찾기(wsDem, 수요ID)
                If rd > 0 Then 합침 = 합침 & "|" & 수요ID & "|" & 셀값(wsDem, rd, "기업명")
            End If
        End If

        rb = PK행찾기(wsBuy, 구매ID)
        If rb > 0 Then
            합침 = 합침 & "|" & 셀값(wsBuy, rb, "구매 담당자")
            발전소ID = 셀값(wsBuy, rb, "발전소ID")
            rp = PK행찾기(wsPlant, 발전소ID)
            If rp > 0 Then
                합침 = 합침 & "|" & 발전소ID & "|" & 셀값(wsPlant, rp, "발전소명") & _
                       "|" & 셀값(wsPlant, rp, "발전법인명")
            End If
        End If

        If InStr(1, 합침, 키워드, vbTextCompare) > 0 Then 결과.Add r
    Next r

    Set 매칭검색 = 결과
End Function

' 현재 위치의 수급매칭 행을 폼 전체에 펼쳐 보여줍니다
Private Sub 폼_결과표시()
    Dim wsL As Worksheet, wsM As Worksheet
    Dim 위치 As Long, 개수 As Long, 행 As Long

    Set wsL = 목록시트()
    개수 = 조회결과_개수()
    If 개수 = 0 Then Exit Sub
    위치 = CLng(Val(wsL.Cells(1, LC_POS).Value))
    If 위치 < 1 Then 위치 = 1
    행 = CLng(Val(wsL.Cells(위치, LC_RESULT).Value))
    If 행 < 2 Then Exit Sub

    Set wsM = 시트(SH_MATCH)
    폼_값만지우기
    표값_폼에채우기 SH_MATCH, 행
    폼_불러오기 True

    상태쓰기 "조회 결과 " & 위치 & " / " & 개수 & "건 — 수급매칭 " & 셀값(wsM, 행, "수급매칭ID")
End Sub

Private Sub 표값_폼에채우기(ByVal 시트명 As String, ByVal 행 As Long)
    Dim ws As Worksheet, cols As Variant, i As Long
    If 행 < 2 Then Exit Sub
    Set ws = 시트(시트명)
    If ws Is Nothing Then Exit Sub
    cols = 표_컬럼목록(시트명)
    For i = LBound(cols) To UBound(cols)
        폼값쓰기 시트명, CStr(cols(i)), 셀값(ws, 행, CStr(cols(i)))
    Next i
End Sub

'==============================================================================
' 불러오기 — 지금 입력된 ID들을 기준으로 연결된 표를 모두 채웁니다
'==============================================================================
Public Sub 폼_불러오기(Optional ByVal 조용히 As Boolean = False)
    Dim r As Long, v As String

    ' 수급매칭 PK 가 있으면 그것부터
    v = 폼값(SH_MATCH, "수급매칭ID")
    If Len(v) > 0 Then
        r = PK행찾기(시트(SH_MATCH), v)
        If r > 0 Then 표값_폼에채우기 SH_MATCH, r
    End If

    ' 수요측: 전기사용지 → 판매계약 → 수요기업
    v = 폼값(SH_MATCH, "전기사용지ID")
    If Len(v) = 0 Then v = 폼값(SH_SITE, "전기사용지ID")
    If Len(v) > 0 Then
        r = PK행찾기(시트(SH_SITE), v)
        If r > 0 Then 표값_폼에채우기 SH_SITE, r
    End If
    v = 폼값(SH_SITE, "판매계약ID")
    If Len(v) = 0 Then v = 폼값(SH_SELL, "판매계약ID")
    If Len(v) > 0 Then
        r = PK행찾기(시트(SH_SELL), v)
        If r > 0 Then 표값_폼에채우기 SH_SELL, r
    End If
    v = 폼값(SH_SELL, "수요기업ID")
    If Len(v) = 0 Then v = 폼값(SH_DEMAND, "수요기업ID")
    If Len(v) > 0 Then
        r = PK행찾기(시트(SH_DEMAND), v)
        If r > 0 Then 표값_폼에채우기 SH_DEMAND, r
    End If

    ' 공급측: 구매계약 → 발전소
    v = 폼값(SH_MATCH, "구매계약ID")
    If Len(v) = 0 Then v = 폼값(SH_BUY, "구매계약ID")
    If Len(v) > 0 Then
        r = PK행찾기(시트(SH_BUY), v)
        If r > 0 Then 표값_폼에채우기 SH_BUY, r
    End If
    v = 폼값(SH_BUY, "발전소ID")
    If Len(v) = 0 Then v = 폼값(SH_PLANT, "발전소ID")
    If Len(v) > 0 Then
        r = PK행찾기(시트(SH_PLANT), v)
        If r > 0 Then 표값_폼에채우기 SH_PLANT, r
    End If

    폼_연동목록갱신
    폼_기존신규표시
    If Not 조용히 Then 상태쓰기 "연결된 항목을 불러왔습니다."
End Sub

' 각 표의 PK 칸에 기존/신규 여부를 색과 안내로 표시
Private Sub 폼_기존신규표시()
    Dim ws As Worksheet, t As Variant, 표이름 As String
    Dim r As Long, PK값 As String
    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Sub

    For Each t In 표_순서()
        표이름 = CStr(t)
        r = 필드행(ws, 표이름 & "|" & 표_PK(표이름))
        If r > 0 Then
            PK값 = Trim$(CStr(ws.Cells(r, COL_VALUE).Value))
            If Len(PK값) = 0 Then
                ws.Cells(r, COL_NOTE).Value = "PK · 비워두면 저장할 때 자동 생성"
                ws.Cells(r, COL_NOTE).Font.Color = RGB(110, 110, 110)
                ws.Cells(r, COL_VALUE).Interior.Color = RGB(255, 255, 255)
            ElseIf PK행찾기(시트(표이름), PK값) > 0 Then
                ws.Cells(r, COL_NOTE).Value = "기존 항목 — 값을 고치면 원본이 수정됩니다"
                ws.Cells(r, COL_NOTE).Font.Color = RGB(30, 99, 168)
                ws.Cells(r, COL_VALUE).Interior.Color = RGB(230, 239, 248)
            Else
                ws.Cells(r, COL_NOTE).Value = "신규 — 저장하면 새로 추가됩니다"
                ws.Cells(r, COL_NOTE).Font.Color = RGB(31, 122, 84)
                ws.Cells(r, COL_VALUE).Interior.Color = RGB(231, 243, 236)
            End If
        End If
    Next t
End Sub

'==============================================================================
' 드롭다운 목록
'   유효성 검사 목록은 문자열로 넣으면 255자 제한이 있어 발전소 589건 같은 목록을
'   담지 못합니다. 그래서 숨김 시트에 값을 쓰고 "이름(Name)"으로 참조합니다.
'==============================================================================
Private Function 목록시트() As Worksheet
    Set 목록시트 = 시트(LIST_SHEET)
    If 목록시트 Is Nothing Then
        Set 목록시트 = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        목록시트.Name = LIST_SHEET
        목록시트.Visible = xlSheetHidden
    End If
End Function

Public Sub 폼_목록갱신()
    Dim wsL As Worksheet
    Application.ScreenUpdating = False
    Set wsL = 목록시트()

    ID목록쓰기 wsL, LC_PLANT, "목록_발전소", ID수집(SH_PLANT, "", "")
    ID목록쓰기 wsL, LC_BUY, "목록_구매계약", ID수집(SH_BUY, "", "")
    ID목록쓰기 wsL, LC_DEM, "목록_수요기업", ID수집(SH_DEMAND, "", "")
    ID목록쓰기 wsL, LC_SELL, "목록_판매계약", ID수집(SH_SELL, "", "")
    ID목록쓰기 wsL, LC_SITE, "목록_전기사용지", ID수집(SH_SITE, "", "")
    ID목록쓰기 wsL, LC_MATCH, "목록_수급매칭", ID수집(SH_MATCH, "", "")

    폼_연동목록갱신
    Application.ScreenUpdating = True
End Sub

' 상위에서 고른 값에 따라 하위 목록을 좁혀 다시 겁니다
Private Sub 폼_연동목록갱신()
    Dim wsF As Worksheet, wsL As Worksheet
    Set wsF = 시트(FORM_SHEET)
    If wsF Is Nothing Then Exit Sub
    Set wsL = 목록시트()

    ID목록쓰기 wsL, LC_BUY_F, "목록_구매계약_F", ID수집(SH_BUY, "발전소ID", 폼값(SH_PLANT, "발전소ID"))
    ID목록쓰기 wsL, LC_SELL_F, "목록_판매계약_F", ID수집(SH_SELL, "수요기업ID", 폼값(SH_DEMAND, "수요기업ID"))
    ID목록쓰기 wsL, LC_SITE_F, "목록_전기사용지_F", ID수집(SH_SITE, "판매계약ID", 폼값(SH_SELL, "판매계약ID"))

    유효성적용 wsF, SH_PLANT, "발전소ID", "목록_발전소"
    유효성적용 wsF, SH_DEMAND, "수요기업ID", "목록_수요기업"
    유효성적용 wsF, SH_MATCH, "수급매칭ID", "목록_수급매칭"

    유효성적용 wsF, SH_BUY, "구매계약ID", "목록_구매계약_F"
    유효성적용 wsF, SH_BUY, "발전소ID", "목록_발전소"
    유효성적용 wsF, SH_SELL, "판매계약ID", "목록_판매계약_F"
    유효성적용 wsF, SH_SELL, "수요기업ID", "목록_수요기업"
    유효성적용 wsF, SH_SITE, "전기사용지ID", "목록_전기사용지_F"
    유효성적용 wsF, SH_SITE, "판매계약ID", "목록_판매계약_F"
    유효성적용 wsF, SH_MATCH, "전기사용지ID", "목록_전기사용지_F"
    유효성적용 wsF, SH_MATCH, "구매계약ID", "목록_구매계약_F"
End Sub

' 조건이 비어 있으면 전체 목록
Private Function ID수집(ByVal 시트명 As String, ByVal 조건컬럼 As String, _
                        ByVal 조건값 As String) As Collection
    Dim 결과 As New Collection
    Dim ws As Worksheet, pk열 As Long, 조건열 As Long, r As Long, 끝 As Long, v As String

    Set ws = 시트(시트명)
    If ws Is Nothing Then
        Set ID수집 = 결과
        Exit Function
    End If
    pk열 = 헤더열(ws, 표_PK(시트명))
    If Len(조건컬럼) > 0 And Len(조건값) > 0 Then 조건열 = 헤더열(ws, 조건컬럼)
    끝 = 마지막행(ws, pk열)

    For r = 2 To 끝
        v = Trim$(CStr(ws.Cells(r, pk열).Value))
        If Len(v) > 0 Then
            If 조건열 = 0 Then
                결과.Add v
            ElseIf Trim$(CStr(ws.Cells(r, 조건열).Value)) = 조건값 Then
                결과.Add v
            End If
        End If
    Next r
    Set ID수집 = 결과
End Function

Private Sub ID목록쓰기(ByVal wsL As Worksheet, ByVal 열 As Long, ByVal 이름 As String, _
                       ByVal 값들 As Collection)
    Dim i As Long
    wsL.Columns(열).ClearContents
    For i = 1 To 값들.Count
        wsL.Cells(i, 열).Value = 값들(i)
    Next i

    On Error Resume Next
    ThisWorkbook.Names(이름).Delete
    On Error GoTo 0
    If 값들.Count > 0 Then
        ThisWorkbook.Names.Add Name:=이름, _
            RefersTo:="='" & LIST_SHEET & "'!" & _
                      wsL.Range(wsL.Cells(1, 열), wsL.Cells(값들.Count, 열)).Address(True, True), _
            Visible:=False
    End If
End Sub

Private Sub 유효성적용(ByVal wsF As Worksheet, ByVal 시트명 As String, _
                       ByVal 컬럼 As String, ByVal 이름 As String)
    Dim r As Long, 있음 As Boolean, nm As Name
    r = 필드행(wsF, 시트명 & "|" & 컬럼)
    If r = 0 Then Exit Sub

    On Error Resume Next
    Set nm = ThisWorkbook.Names(이름)
    있음 = (Err.Number = 0) And (Not nm Is Nothing)
    Err.Clear

    With wsF.Cells(r, COL_VALUE).Validation
        .Delete
        If 있음 Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                 Operator:=xlBetween, Formula1:="=" & 이름
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = False      ' 새 ID를 직접 타이핑하는 것도 허용
        End If
    End With
    On Error GoTo 0
End Sub

'==============================================================================
' 자동 변경 감지 (선택 사항 — 입력폼 시트 모듈에서 호출)
'==============================================================================
Public Sub 폼_변경감지(ByVal Target As Range)
    Dim ws As Worksheet, 이전 As Boolean
    Set ws = 시트(FORM_SHEET)
    If ws Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> FORM_SHEET Then Exit Sub
    If Target.Column <> COL_VALUE Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Len(CStr(ws.Cells(Target.Row, COL_KEY).Value)) = 0 Then Exit Sub

    이전 = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    폼_불러오기 True
    On Error GoTo 0
    Application.EnableEvents = 이전
End Sub

'==============================================================================
' 저장 — 검증 → 미리보기 확인 → 6개 표에 반영
'==============================================================================
Public Sub 폼_저장()
    Dim 오류 As String, 미리보기 As String
    Dim t As Variant, 표이름 As String, PK값 As String
    Dim 신규ID As Object
    Dim 답 As VbMsgBoxResult

    If 시트(FORM_SHEET) Is Nothing Then Exit Sub
    Set 신규ID = CreateObject("Scripting.Dictionary")

    ' 1) 저장 대상 판단 + PK 자동 생성
    For Each t In 표_순서()
        표이름 = CStr(t)
        If 표_입력있음(표이름) Then
            PK값 = 폼값(표이름, 표_PK(표이름))
            If Len(PK값) = 0 Then 신규ID(표이름) = 다음ID(표이름)
        End If
    Next t

    ' 2) 검증
    오류 = 저장전_검증(신규ID)
    If Len(오류) > 0 Then
        MsgBox "저장할 수 없습니다." & vbCrLf & vbCrLf & 오류, vbExclamation, "검증 오류"
        상태쓰기 "검증 오류로 저장하지 않았습니다.", True
        Exit Sub
    End If

    ' 3) 미리보기 확인
    미리보기 = 저장_미리보기(신규ID)
    If Len(미리보기) = 0 Then
        MsgBox "저장할 내용이 없습니다. 값을 입력해주세요.", vbInformation, "저장"
        Exit Sub
    End If
    답 = MsgBox(미리보기 & vbCrLf & "이대로 저장할까요?", vbQuestion + vbYesNo, "저장 확인")
    If 답 <> vbYes Then
        상태쓰기 "저장을 취소했습니다."
        Exit Sub
    End If

    ' 4) 반영 (부모 → 자식 순서 — 부모가 먼저 있어야 FK 가 성립)
    Application.ScreenUpdating = False
    For Each t In 표_순서()
        표이름 = CStr(t)
        If 표_입력있음(표이름) Then 표_저장 신규ID, 표이름
    Next t

    ' 자동 생성된 ID를 폼에 되돌려 보여주고, 나머지 칸도 저장 결과로 새로 고침
    For Each t In 표_순서()
        표이름 = CStr(t)
        If 신규ID.Exists(표이름) Then 폼값쓰기 표이름, 표_PK(표이름), CStr(신규ID(표이름))
    Next t
    폼_목록갱신
    폼_불러오기 True
    Application.ScreenUpdating = True

    상태쓰기 "저장했습니다. (" & Format$(Now, "hh:nn:ss") & ")"
    MsgBox "저장을 완료했습니다.", vbInformation, "저장"
End Sub

' 저장에 실제로 쓰일 값을 계산합니다.
'   - PK 를 비워뒀으면 이번에 자동 생성된 ID
'   - FK 를 비워뒀는데 그 부모를 이번에 같이 만든다면 그 부모의 새 ID
' 폼에 직접 써넣지 않고 계산만 하므로, 사용자가 저장을 취소해도 폼이 더럽혀지지
' 않습니다.
Private Function 값해석(ByVal 신규ID As Object, ByVal 시트명 As String, _
                        ByVal 컬럼 As String) As String
    Dim v As String, 참조 As String
    v = 폼값(시트명, 컬럼)

    If 컬럼 = 표_PK(시트명) Then
        If Len(v) = 0 And 신규ID.Exists(시트명) Then v = CStr(신규ID(시트명))
        값해석 = v
        Exit Function
    End If

    참조 = FK_참조시트(시트명, 컬럼)
    If Len(참조) > 0 And Len(v) = 0 Then
        If 신규ID.Exists(참조) Then v = CStr(신규ID(참조))
    End If
    값해석 = v
End Function

' 그 표에 사용자가 값을 하나라도 넣었는지
Private Function 표_입력있음(ByVal 시트명 As String) As Boolean
    Dim cols As Variant, i As Long
    cols = 표_컬럼목록(시트명)
    For i = LBound(cols) To UBound(cols)
        If Len(폼값(시트명, CStr(cols(i)))) > 0 Then
            표_입력있음 = True
            Exit Function
        End If
    Next i
End Function

' 기존 매크로의 검증 규칙과 동일: PK 공란/중복, FK 공란/참조, 조합중복
Private Function 저장전_검증(ByVal 신규ID As Object) As String
    Dim t As Variant, 표이름 As String, cols As Variant, i As Long
    Dim 오류 As String, PK값 As String, 원래PK As String
    Dim FK값 As String, 참조 As String, 라벨 As String

    For Each t In 표_순서()
        표이름 = CStr(t)
        If 표_입력있음(표이름) Then
            라벨 = Replace(표이름, "T_", "")
            원래PK = 폼값(표이름, 표_PK(표이름))
            PK값 = 원래PK
            If 신규ID.Exists(표이름) Then PK값 = CStr(신규ID(표이름))

            If Len(PK값) = 0 Then
                오류 = 오류 & "· " & 라벨 & ": " & 표_PK(표이름) & " 가 비어 있습니다." & vbCrLf
            End If

            cols = 표_컬럼목록(표이름)
            For i = LBound(cols) To UBound(cols)
                참조 = FK_참조시트(표이름, CStr(cols(i)))
                If Len(참조) > 0 Then
                    FK값 = 폼값(표이름, CStr(cols(i)))
                    ' 이번에 같이 만들어질 부모라면 그 새 ID로 채워집니다
                    If Len(FK값) = 0 And 신규ID.Exists(참조) Then FK값 = CStr(신규ID(참조))
                    If Len(FK값) = 0 Then
                        오류 = 오류 & "· " & 라벨 & ": " & cols(i) & " 가 비어 있습니다." & vbCrLf
                    ElseIf PK행찾기(시트(참조), FK값) = 0 Then
                        If Not (신규ID.Exists(참조) And FK값 = CStr(신규ID(참조))) Then
                            오류 = 오류 & "· " & 라벨 & ": " & cols(i) & " '" & FK값 & _
                                   "' 에 해당하는 " & Replace(참조, "T_", "") & " 가 없습니다." & vbCrLf
                        End If
                    End If
                End If
            Next i
        End If
    Next t

    ' 수급매칭 조합중복 (전기사용지ID + 구매계약ID)
    If 표_입력있음(SH_MATCH) Then
        Dim 사용지 As String, 구매 As String, 내PK As String
        Dim wsM As Worksheet, r As Long, 끝 As Long, pk열 As Long

        사용지 = 폼값(SH_MATCH, "전기사용지ID")
        If Len(사용지) = 0 And 신규ID.Exists(SH_SITE) Then 사용지 = CStr(신규ID(SH_SITE))
        구매 = 폼값(SH_MATCH, "구매계약ID")
        If Len(구매) = 0 And 신규ID.Exists(SH_BUY) Then 구매 = CStr(신규ID(SH_BUY))
        내PK = 폼값(SH_MATCH, "수급매칭ID")
        If 신규ID.Exists(SH_MATCH) Then 내PK = CStr(신규ID(SH_MATCH))

        If Len(사용지) > 0 And Len(구매) > 0 Then
            Set wsM = 시트(SH_MATCH)
            pk열 = 헤더열(wsM, "수급매칭ID")
            끝 = 마지막행(wsM, pk열)
            For r = 2 To 끝
                If 셀값(wsM, r, "전기사용지ID") = 사용지 And 셀값(wsM, r, "구매계약ID") = 구매 Then
                    If 셀값(wsM, r, "수급매칭ID") <> 내PK Then
                        오류 = 오류 & "· 수급매칭: 같은 (전기사용지 " & 사용지 & " + 구매계약 " & 구매 & _
                               ") 조합이 이미 있습니다 [" & 셀값(wsM, r, "수급매칭ID") & "]." & vbCrLf
                        Exit For
                    End If
                End If
            Next r
        End If
    End If

    저장전_검증 = 오류
End Function

' 무엇이 새로 생기고 무엇이 바뀌는지 사람이 읽을 수 있게 정리
Private Function 저장_미리보기(ByVal 신규ID As Object) As String
    Dim t As Variant, 표이름 As String, cols As Variant, i As Long
    Dim s As String, ws As Worksheet, r As Long
    Dim 새값 As String, 옛값 As String, 변경 As String, PK값 As String, 라벨 As String

    For Each t In 표_순서()
        표이름 = CStr(t)
        If 표_입력있음(표이름) Then
            라벨 = Replace(표이름, "T_", "")
            Set ws = 시트(표이름)
            PK값 = 폼값(표이름, 표_PK(표이름))
            If 신규ID.Exists(표이름) Then PK값 = CStr(신규ID(표이름))
            r = PK행찾기(ws, PK값)

            If r = 0 Then
                s = s & "[신규] " & 라벨 & "  " & PK값 & vbCrLf
            Else
                변경 = ""
                cols = 표_컬럼목록(표이름)
                For i = LBound(cols) To UBound(cols)
                    새값 = 값해석(신규ID, 표이름, CStr(cols(i)))
                    옛값 = 셀값(ws, r, CStr(cols(i)))
                    If Len(새값) > 0 And 새값 <> 옛값 Then
                        변경 = 변경 & "      · " & cols(i) & ": " & _
                               IIf(Len(옛값) = 0, "(공란)", 옛값) & " → " & 새값 & vbCrLf
                    End If
                Next i
                If Len(변경) > 0 Then
                    s = s & "[수정] " & 라벨 & "  " & PK값 & vbCrLf & 변경
                Else
                    s = s & "[변경없음] " & 라벨 & "  " & PK값 & vbCrLf
                End If
            End If
        End If
    Next t

    If Len(s) > 0 Then 저장_미리보기 = "아래 내용으로 반영됩니다." & vbCrLf & vbCrLf & s
End Function

' 한 표를 실제로 쓰기 (있으면 수정, 없으면 새 행 추가)
Private Sub 표_저장(ByVal 신규ID As Object, ByVal 시트명 As String)
    Dim ws As Worksheet, cols As Variant, i As Long
    Dim PK값 As String, r As Long, pk열 As Long, v As String

    Set ws = 시트(시트명)
    If ws Is Nothing Then Exit Sub
    PK값 = 값해석(신규ID, 시트명, 표_PK(시트명))
    If Len(PK값) = 0 Then Exit Sub

    r = PK행찾기(ws, PK값)
    If r = 0 Then
        pk열 = 헤더열(ws, 표_PK(시트명))
        r = 마지막행(ws, pk열) + 1
    End If

    cols = 표_컬럼목록(시트명)
    For i = LBound(cols) To UBound(cols)
        v = 값해석(신규ID, 시트명, CStr(cols(i)))
        If i = 0 Then
            셀쓰기 ws, r, CStr(cols(i)), PK값
        ElseIf Len(v) > 0 Then
            ' 폼에서 비워 둔 칸은 기존 값을 지우지 않습니다
            셀쓰기 ws, r, CStr(cols(i)), v
        End If
    Next i
End Sub

'==============================================================================
' PK 자동 생성 — 기존 ID들이 쓰고 있는 규칙을 그대로 따라갑니다
'   예) 기존이 "구매-0001..0615" 면 다음은 "구매-0616"
'       기존이 "P001..P589"      면 다음은 "P590"
'==============================================================================
Private Function 다음ID(ByVal 시트명 As String) As String
    Dim ws As Worksheet, pk열 As Long, r As Long, 끝 As Long, p As Long
    Dim v As String, 접두 As String, 숫자부 As String, 공통접두 As String
    Dim 최대 As Long, 자릿수 As Long, n As Long, 첫번째 As Boolean

    Set ws = 시트(시트명)
    If ws Is Nothing Then Exit Function
    pk열 = 헤더열(ws, 표_PK(시트명))
    끝 = 마지막행(ws, pk열)
    첫번째 = True

    For r = 2 To 끝
        v = Trim$(CStr(ws.Cells(r, pk열).Value))
        If Len(v) > 0 Then
            숫자부 = ""
            For p = Len(v) To 1 Step -1
                If Mid$(v, p, 1) Like "#" Then
                    숫자부 = Mid$(v, p, 1) & 숫자부
                Else
                    Exit For
                End If
            Next p
            If Len(숫자부) > 0 Then
                접두 = Left$(v, Len(v) - Len(숫자부))
                If 첫번째 Then
                    공통접두 = 접두
                    첫번째 = False
                ElseIf 공통접두 <> 접두 Then
                    공통접두 = "?"          ' 규칙이 하나가 아님
                End If
                n = CLng(Val(숫자부))
                If n > 최대 Then 최대 = n
                If Len(숫자부) > 자릿수 Then 자릿수 = Len(숫자부)
            End If
        End If
    Next r

    If 자릿수 = 0 Then 자릿수 = 3
    If Len(공통접두) = 0 Or 공통접두 = "?" Then
        Select Case 시트명                     ' 규칙을 못 찾으면 기본 접두
        Case SH_PLANT:  공통접두 = "P"
        Case SH_BUY:    공통접두 = "구매-"
        Case SH_DEMAND: 공통접두 = "D"
        Case SH_SELL:   공통접두 = "판매-"
        Case SH_SITE:   공통접두 = "전기사용지-"
        Case SH_MATCH:  공통접두 = "매칭-"
        End Select
    End If

    Do
        최대 = 최대 + 1
        다음ID = 공통접두 & Format$(최대, String$(자릿수, "0"))
    Loop While PK행찾기(ws, 다음ID) > 0
End Function

'==============================================================================
' 삭제 — 현재 폼의 수급매칭 행만 지웁니다 (상위 표는 건드리지 않음)
'==============================================================================
Public Sub 폼_삭제()
    Dim wsM As Worksheet, 매칭ID As String, r As Long
    Dim 답 As VbMsgBoxResult

    매칭ID = 폼값(SH_MATCH, "수급매칭ID")
    If Len(매칭ID) = 0 Then
        MsgBox "삭제할 수급매칭ID가 폼에 없습니다." & vbCrLf & _
               "먼저 [조회]로 대상을 불러와주세요.", vbInformation, "삭제"
        Exit Sub
    End If

    Set wsM = 시트(SH_MATCH)
    r = PK행찾기(wsM, 매칭ID)
    If r = 0 Then
        MsgBox "'" & 매칭ID & "' 은(는) " & SH_MATCH & " 에 없습니다.", vbInformation, "삭제"
        Exit Sub
    End If

    답 = MsgBox("수급매칭 '" & 매칭ID & "' 행을 삭제합니다." & vbCrLf & vbCrLf & _
                "전기사용지: " & 셀값(wsM, r, "전기사용지ID") & vbCrLf & _
                "구매계약: " & 셀값(wsM, r, "구매계약ID") & vbCrLf & vbCrLf & _
                "발전소·구매계약·판매계약 등 상위 항목은 그대로 둡니다." & vbCrLf & _
                "계속할까요?", vbExclamation + vbYesNo, "삭제 확인")
    If 답 <> vbYes Then Exit Sub

    wsM.Rows(r).Delete
    조회결과_지우기                     ' 행 번호가 밀리므로 조회 결과는 무효
    폼_목록갱신
    상태쓰기 "수급매칭 " & 매칭ID & " 을(를) 삭제했습니다. 다시 [조회]해주세요."
    MsgBox "삭제했습니다.", vbInformation, "삭제"
End Sub
