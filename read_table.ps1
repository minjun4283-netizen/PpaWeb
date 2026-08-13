Option Explicit

Dim args
Set args = WScript.Arguments

If args.Count < 3 Then
    WScript.Quit 1
End If

Dim xlsmPath, tableName, outJsonPath
xlsmPath = args.Item(0)
tableName = args.Item(1)
outJsonPath = args.Item(2)

Dim xl, wb, ws, headerRow, headersJson, rowsJson

On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Excel.Application create failed: " & Err.Description)
    WScript.Quit 2
End If
Err.Clear
On Error GoTo 0

xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
xl.EnableEvents = False

On Error Resume Next
Set wb = xl.Workbooks.Open(xlsmPath, 0, True)
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Workbook open failed: " & Err.Description)
    Cleanup xl, wb
    WScript.Quit 3
End If
Err.Clear
On Error GoTo 0

Set ws = FindWorksheet(wb, tableName)
If ws Is Nothing Then
    WriteUtf8 outJsonPath, MakeErrorJson("Worksheet not found: " & tableName)
    Cleanup xl, wb
    WScript.Quit 4
End If

headerRow = DetectHeaderRow(ws)
headersJson = BuildHeadersJson(ws, headerRow)
rowsJson = BuildRowsJson(ws, headerRow)

WriteUtf8 outJsonPath, "{" & QQ("ok") & ":true," & QQ("headers") & ":" & headersJson & "," & QQ("rows") & ":" & rowsJson & "}"

Cleanup xl, wb
WScript.Quit 0


Sub Cleanup(xlApp, wbObj)
    On Error Resume Next
    If Not wbObj Is Nothing Then wbObj.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
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


Function FindWorksheet(wbObj, targetName)
    Dim ws, targetNorm, targetNoT, wsNorm

    For Each ws In wbObj.Worksheets
        If GetText(ws.Name) = targetName Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next

    targetNorm = NormalizeName(targetName)
    targetNoT = NormalizeName(Replace(targetName, "T_", ""))

    For Each ws In wbObj.Worksheets
        wsNorm = NormalizeName(ws.Name)
        If wsNorm = targetNorm Or wsNorm = targetNoT Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next

    For Each ws In wbObj.Worksheets
        wsNorm = NormalizeName(ws.Name)
        If targetNoT <> "" Then
            If InStr(1, wsNorm, targetNoT, vbTextCompare) > 0 Or InStr(1, targetNoT, wsNorm, vbTextCompare) > 0 Then
                Set FindWorksheet = ws
                Exit Function
            End If
        End If
    Next

    Set FindWorksheet = Nothing
End Function


Function DetectHeaderRow(ws)
    Dim usedLastCol, r, c, cnt, bestCnt, bestRow
    Dim val

    On Error Resume Next
    usedLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If Err.Number <> 0 Then
        Err.Clear
        usedLastCol = 100
    End If
    On Error GoTo 0

    If usedLastCol < 1 Then usedLastCol = 100

    bestCnt = -1
    bestRow = 1

    For r = 1 To 20
        cnt = 0
        For c = 1 To usedLastCol
            val = GetText(ws.Cells(r, c).Value)
            If val <> "" Then cnt = cnt + 1
        Next

        If cnt > bestCnt Then
            bestCnt = cnt
            bestRow = r
        End If
    Next

    DetectHeaderRow = bestRow
End Function


Function BuildHeadersJson(ws, headerRow)
    Dim usedLastCol, c, val, outText
    Dim seen

    Set seen = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    usedLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If Err.Number <> 0 Then
        Err.Clear
        usedLastCol = 100
    End If
    On Error GoTo 0

    If usedLastCol < 1 Then usedLastCol = 100

    outText = ""

    For c = 1 To usedLastCol
        val = GetText(ws.Cells(headerRow, c).Value)
        If val <> "" Then
            If Not seen.Exists(val) Then
                seen.Add val, True
                If outText <> "" Then outText = outText & ","
                outText = outText & QQ(val)
            End If
        End If
    Next

    BuildHeadersJson = "[" & outText & "]"
End Function


Function BuildRowsJson(ws, headerRow)
    Dim usedLastCol, usedLastRow
    Dim c, r, key, val, rowText, outText
    Dim headerCols, headerOrder
    Dim idx, nonEmpty

    Set headerCols = CreateObject("Scripting.Dictionary")
    Set headerOrder = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    usedLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If Err.Number <> 0 Then
        Err.Clear
        usedLastCol = 100
        usedLastRow = headerRow
    End If
    On Error GoTo 0

    If usedLastCol < 1 Then usedLastCol = 100
    If usedLastRow < headerRow Then usedLastRow = headerRow

    idx = 0
    For c = 1 To usedLastCol
        key = GetText(ws.Cells(headerRow, c).Value)
        If key <> "" Then
            If Not headerCols.Exists(CStr(c)) Then
                headerCols.Add CStr(c), key
                headerOrder.Add CStr(idx), CStr(c)
                idx = idx + 1
            End If
        End If
    Next

    outText = ""

    For r = headerRow + 1 To usedLastRow
        rowText = ""
        nonEmpty = 0

        For idx = 0 To headerOrder.Count - 1
            c = CLng(headerOrder(CStr(idx)))
            key = headerCols(CStr(c))
            val = GetText(ws.Cells(r, c).Text)

            If val <> "" Then nonEmpty = nonEmpty + 1

            If rowText <> "" Then rowText = rowText & ","
            rowText = rowText & QQ(key) & ":" & QQ(val)
        Next

        If nonEmpty > 0 Then
            If outText <> "" Then outText = outText & ","
            outText = outText & "{" & rowText & "}"
        End If
    Next

    BuildRowsJson = "[" & outText & "]"
End Function