Option Explicit

Dim args
Set args = WScript.Arguments

If args.Count < 5 Then
    WScript.Quit 1
End If

Dim xlsmPath, tableName, pkName, payloadPath, outJsonPath
xlsmPath = args.Item(0)
tableName = args.Item(1)
pkName = args.Item(2)
payloadPath = args.Item(3)
outJsonPath = args.Item(4)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim tempRoot, tempDir, tempXlsm
tempRoot = GetTempRoot()
tempDir = tempRoot & "\" & fso.GetTempName
tempXlsm = tempDir & "\source.xlsm"

Dim xl, wb, ws
Dim dict, pkValue
Dim headerRow, usedLastCol, usedLastRow
Dim headerMap, c, key
Dim pkCol, targetRow, r, existing, actionText
Dim k

On Error Resume Next
fso.CreateFolder tempDir
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Temp folder create failed: " & Err.Description)
    WScript.Quit 10
End If
Err.Clear

fso.CopyFile xlsmPath, tempXlsm, True
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Source copy failed: " & Err.Description)
    CleanupTemp fso, tempDir
    WScript.Quit 11
End If
Err.Clear

Set wb = GetObject(tempXlsm)
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Workbook open failed: " & Err.Description)
    CleanupTemp fso, tempDir
    WScript.Quit 12
End If
Err.Clear

Set xl = wb.Application
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Excel application attach failed: " & Err.Description)
    CleanupExcel wb, xl
    CleanupTemp fso, tempDir
    WScript.Quit 13
End If
Err.Clear
On Error GoTo 0

xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
xl.EnableEvents = False

Set ws = FindWorksheet(wb, tableName)
If ws Is Nothing Then
    WriteUtf8 outJsonPath, MakeErrorJson("Worksheet not found: " & tableName)
    CleanupExcel wb, xl
    CleanupTemp fso, tempDir
    WScript.Quit 14
End If

Set dict = ParsePayloadLines(ReadUtf8(payloadPath))

If Not dict.Exists(pkName) Then
    WriteUtf8 outJsonPath, MakeErrorJson("PK value missing: " & pkName)
    CleanupExcel wb, xl
    CleanupTemp fso, tempDir
    WScript.Quit 15
End If

pkValue = Trim(CStr(dict(pkName)))
If pkValue = "" Then
    WriteUtf8 outJsonPath, MakeErrorJson("PK value blank: " & pkName)
    CleanupExcel wb, xl
    CleanupTemp fso, tempDir
    WScript.Quit 16
End If

Set headerMap = CreateObject("Scripting.Dictionary")

headerRow = DetectHeaderRow(ws)

On Error Resume Next
usedLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
If Err.Number <> 0 Then
    Err.Clear
    usedLastCol = 100
    usedLastRow = headerRow
End If
On Error GoTo 0

For c = 1 To usedLastCol
    key = GetText(ws.Cells(headerRow, c).Value)
    If key <> "" Then
        If Not headerMap.Exists(key) Then
            headerMap.Add key, c
        End If
    End If
Next

If Not headerMap.Exists(pkName) Then
    WriteUtf8 outJsonPath, MakeErrorJson("PK column not found: " & pkName)
    CleanupExcel wb, xl
    CleanupTemp fso, tempDir
    WScript.Quit 17
End If

pkCol = CLng(headerMap(pkName))
targetRow = 0

For r = headerRow + 1 To usedLastRow
    existing = GetText(ws.Cells(r, pkCol).Text)
    If existing = pkValue Then
        targetRow = r
        Exit For
    End If
Next

If targetRow = 0 Then
    targetRow = usedLastRow + 1
    If targetRow <= headerRow Then targetRow = headerRow + 1
    actionText = "inserted"
Else
    actionText = "updated"
End If

For Each k In dict.Keys
    If headerMap.Exists(k) Then
        ws.Cells(targetRow, CLng(headerMap(k))).Value = CStr(dict(k))
    End If
Next

On Error Resume Next
wb.Save
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Workbook save failed: " & Err.Description)
    CleanupExcel wb, xl
    CleanupTemp fso, tempDir
    WScript.Quit 18
End If
Err.Clear
On Error GoTo 0

CleanupExcel wb, xl

On Error Resume Next
fso.CopyFile tempXlsm, xlsmPath, True
If Err.Number <> 0 Then
    WriteUtf8 outJsonPath, MakeErrorJson("Final overwrite failed. Close the original Excel file first: " & Err.Description)
    CleanupTemp fso, tempDir
    WScript.Quit 19
End If
Err.Clear
On Error GoTo 0

WriteUtf8 outJsonPath, _
    "{" & QQ("ok") & ":true," & QQ("result") & ":{" & _
    QQ("action") & ":" & QQ(actionText) & "," & _
    QQ("row") & ":" & CStr(targetRow) & "," & _
    QQ("pk_name") & ":" & QQ(pkName) & "," & _
    QQ("pk_value") & ":" & QQ(pkValue) & "," & _
    QQ("table") & ":" & QQ(tableName) & "}}"

CleanupTemp fso, tempDir
WScript.Quit 0


Sub CleanupExcel(wbObj, xlApp)
    On Error Resume Next
    If Not wbObj Is Nothing Then wbObj.Close True
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

Sub CleanupTemp(fsoObj, folderPath)
    On Error Resume Next
    If Not fsoObj Is Nothing Then
        If fsoObj.FolderExists(folderPath) Then
            fsoObj.DeleteFolder folderPath, True
        End If
    End If
End Sub

Function GetTempRoot()
    Dim sh
    Set sh = CreateObject("WScript.Shell")
    GetTempRoot = sh.ExpandEnvironmentStrings("%TEMP%")
End Function

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

Function ReadUtf8(filePath)
    Dim stm
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadUtf8 = stm.ReadText
    stm.Close
End Function

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

Function ParsePayloadLines(textValue)
    Dim dict, normalized, lines, i, line, parts, key, value
    Set dict = CreateObject("Scripting.Dictionary")

    normalized = Replace(textValue, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    lines = Split(normalized, vbLf)

    For i = 0 To UBound(lines)
        line = lines(i)
        If line <> "" Then
            parts = Split(line, vbTab, 2)
            key = ""
            value = ""

            If UBound(parts) >= 0 Then key = parts(0)
            If UBound(parts) >= 1 Then value = parts(1)

            If key <> "" Then
                dict(key) = value
            End If
        End If
    Next

    Set ParsePayloadLines = dict
End Function