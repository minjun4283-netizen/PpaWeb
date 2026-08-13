Option Explicit

Dim args, xlsmPath, outPath
Set args = WScript.Arguments

If args.Count < 2 Then
    WScript.Echo "usage: extract_schema.vbs <xlsm_path> <out_json_path>"
    WScript.Quit 1
End If

xlsmPath = args.Item(0)
outPath = args.Item(1)

Dim targets
targets = Array( _
    "T_발전소", _
    "T_구매계약", _
    "T_수요기업", _
    "T_판매계약", _
    "T_전기사용지", _
    "T_수급매칭" _
)

Dim xl, wb, jsonText, i, t, sheetName

On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WriteUtf8 outPath, MakeErrorJson("Excel.Application 생성 실패: " & Err.Description)
    WScript.Quit 2
End If
Err.Clear
On Error GoTo 0

xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
xl.EnableEvents = False

On Error Resume Next
xl.AutomationSecurity = 3
On Error GoTo 0

On Error Resume Next
Set wb = xl.Workbooks.Open(xlsmPath, 0, True)
If Err.Number <> 0 Then
    WriteUtf8 outPath, MakeErrorJson("엑셀 파일 열기 실패: " & Err.Description)
    Cleanup xl, wb
    WScript.Quit 3
End If
Err.Clear
On Error GoTo 0

jsonText = "{" & QQ("ok") & ":true," & QQ("schema") & ":{"

For i = 0 To UBound(targets)
    t = targets(i)

    If i > 0 Then
        jsonText = jsonText & ","
    End If

    jsonText = jsonText & QQ(t) & ":"

    sheetName = FindSheetName(wb, t)
    If sheetName = "" Then
        jsonText = jsonText & "[]"
    Else
        jsonText = jsonText & BuildHeadersJson(wb.Worksheets(sheetName))
    End If
Next

jsonText = jsonText & "}}"

WriteUtf8 outPath, jsonText
Cleanup xl, wb
WScript.Quit 0


Sub Cleanup(xlApp, wbObj)
    On Error Resume Next

    If Not wbObj Is Nothing Then
        wbObj.Close False
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If
End Sub


Sub WriteUtf8(filePath, textValue)
    Dim stm
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText textValue
    stm.SaveToFile filePath, 2
    stm.Close
End Sub


Function DQ()
    DQ = Chr(34)
End Function


Function QQ(s)
    QQ = DQ() & JsonEscape(CStr(s)) & DQ()
End Function


Function MakeErrorJson(msg)
    MakeErrorJson = "{" & QQ("ok") & ":false," & QQ("error") & ":" & QQ(msg) & "}"
End Function


Function GetText(v)
    If IsError(v) Then
        GetText = ""
    ElseIf IsNull(v) Then
        GetText = ""
    ElseIf IsEmpty(v) Then
        GetText = ""
    Else
        GetText = Trim(CStr(v))
    End If
End Function


Function NormalizeName(s)
    Dim t
    t = LCase(GetText(s))
    t = Replace(t, " ", "")
    t = Replace(t, "_", "")
    t = Replace(t, "-", "")
    NormalizeName = t
End Function


Function JsonEscape(s)
    Dim t
    t = CStr(s)
    t = Replace(t, "\", "\\")
    t = Replace(t, DQ(), "\" & DQ())
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    JsonEscape = t
End Function


Function IsHelperColumn(name)
    Dim n
    n = GetText(name)

    If n = "" Then
        IsHelperColumn = True
        Exit Function
    End If

    If n = "PK중복" Or n = "PK공란" Or n = "조합중복" Or n = "열1" Then
        IsHelperColumn = True
        Exit Function
    End If

    If Right(n, 2) = "참조" Or Right(n, 2) = "공란" Or Right(n, 2) = "중복" Then
        IsHelperColumn = True
        Exit Function
    End If

    IsHelperColumn = False
End Function


Function FindSheetName(wbObj, targetName)
    Dim ws, targetNorm, targetNoT, wsNorm

    For Each ws In wbObj.Worksheets
        If GetText(ws.Name) = targetName Then
            FindSheetName = ws.Name
            Exit Function
        End If
    Next

    targetNorm = NormalizeName(targetName)
    targetNoT = NormalizeName(Replace(targetName, "T_", ""))

    For Each ws In wbObj.Worksheets
        wsNorm = NormalizeName(ws.Name)
        If wsNorm = targetNorm Or wsNorm = targetNoT Then
            FindSheetName = ws.Name
            Exit Function
        End If
    Next

    For Each ws In wbObj.Worksheets
        wsNorm = NormalizeName(ws.Name)
        If InStr(1, wsNorm, targetNoT, vbTextCompare) > 0 Or InStr(1, targetNoT, wsNorm, vbTextCompare) > 0 Then
            FindSheetName = ws.Name
            Exit Function
        End If
    Next

    FindSheetName = ""
End Function


Function BuildHeadersJson(ws)
    Dim usedLastCol, r, c, cnt, bestCnt, bestRow
    Dim val, outText
    Dim seen

    Set seen = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    usedLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If Err.Number <> 0 Then
        Err.Clear
        usedLastCol = 100
    End If
    On Error GoTo 0

    If usedLastCol < 1 Then
        usedLastCol = 100
    End If

    bestCnt = -1
    bestRow = 1

    For r = 1 To 20
        cnt = 0
        For c = 1 To usedLastCol
            val = GetText(ws.Cells(r, c).Value)
            If val <> "" Then
                cnt = cnt + 1
            End If
        Next

        If cnt > bestCnt Then
            bestCnt = cnt
            bestRow = r
        End If
    Next

    outText = ""

    For c = 1 To usedLastCol
        val = GetText(ws.Cells(bestRow, c).Value)

        If val <> "" Then
            If Not IsHelperColumn(val) Then
                If Not seen.Exists(val) Then
                    seen.Add val, True

                    If outText <> "" Then
                        outText = outText & ","
                    End If

                    outText = outText & QQ(val)
                End If
            End If
        End If
    Next

    BuildHeadersJson = "[" & outText & "]"