param(
    [Parameter(Mandatory = $true)][string]$XlsmPath,
    [Parameter(Mandatory = $true)][string]$OutJsonPath
)

$ErrorActionPreference = "Stop"

function Get-HeaderInfo($ws) {
    $used = $ws.UsedRange
    $lastCol = [Math]::Max(1, $used.Column + $used.Columns.Count - 1)

    $bestRow = 1
    $bestCnt = -1

    for ($r = 1; $r -le 20; $r++) {
        $cnt = 0

        for ($c = 1; $c -le $lastCol; $c++) {
            $v = [string]$ws.Cells.Item($r, $c).Text
            if (-not [string]::IsNullOrWhiteSpace($v)) {
                $cnt++
            }
        }

        if ($cnt -gt $bestCnt) {
            $bestCnt = $cnt
            $bestRow = $r
        }
    }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $headers = New-Object System.Collections.ArrayList
    $cols = New-Object System.Collections.ArrayList

    for ($c = 1; $c -le $lastCol; $c++) {
        $name = [string]$ws.Cells.Item($bestRow, $c).Text
        $name = $name.Trim()

        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($seen.Contains($name)) { continue }

        [void]$seen.Add($name)
        [void]$headers.Add($name)
        [void]$cols.Add($c)
    }

    return [ordered]@{
        HeaderRow = $bestRow
        Headers   = @($headers)
        Cols      = @($cols)
    }
}

function Get-Rows($ws, $headerInfo) {
    $used = $ws.UsedRange
    $lastRow = [Math]::Max(1, $used.Row + $used.Rows.Count - 1)

    $headerRow = [int]$headerInfo.HeaderRow
    $headers = @($headerInfo.Headers)
    $cols = @($headerInfo.Cols)

    $rows = New-Object System.Collections.ArrayList

    for ($r = ($headerRow + 1); $r -le $lastRow; $r++) {
        $obj = [ordered]@{}
        $nonEmpty = 0

        for ($i = 0; $i -lt $headers.Count; $i++) {
            $h = [string]$headers[$i]
            $c = [int]$cols[$i]
            $v = [string]$ws.Cells.Item($r, $c).Text
            $v = $v.Trim()

            if (-not [string]::IsNullOrWhiteSpace($v)) {
                $nonEmpty++
            }

            $obj[$h] = $v
        }

        if ($nonEmpty -gt 0) {
            [void]$rows.Add($obj)
        }
    }

    return @($rows)
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($XlsmPath, 0, $true)

    $sheets = @()

    foreach ($ws in $workbook.Worksheets) {
        $headerInfo = Get-HeaderInfo $ws

        $sheets += [ordered]@{
            name    = [string]$ws.Name
            headers = @($headerInfo.Headers)
            rows    = @(Get-Rows $ws $headerInfo)
        }
    }

    $obj = [ordered]@{
        ok     = $true
        sheets = $sheets
    }

    $json = $obj | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($OutJsonPath, $json, [System.Text.UTF8Encoding]::new($false))
    exit 0
}
catch {
    $obj = [ordered]@{
        ok    = $false
        error = $_.Exception.Message
    }

    $json = $obj | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($OutJsonPath, $json, [System.Text.UTF8Encoding]::new($false))
    exit 1
}
finally {
    if ($workbook -ne $null) {
        try { $workbook.Close($false) } catch {}
    }

    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
    }
}