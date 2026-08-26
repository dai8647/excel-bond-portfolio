$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open("C:\Users\dai86\Downloads\BondRV_Model.xlsm")
    $xl.CalculateFull()
    $lines = @()
    # scan every sheet for error values
    foreach ($ws in $wb.Sheets) {
        $used = $ws.UsedRange
        $errCount = 0
        $errSamples = @()
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count
        # limit scan to avoid huge ranges
        $maxR = [Math]::Min($rows, 1400)
        $maxC = [Math]::Min($cols, 30)
        for ($r=1; $r -le $maxR; $r++) {
            for ($c=1; $c -le $maxC; $c++) {
                $cell = $used.Cells.Item($r,$c)
                $v = $cell.Value2
                if ($v -is [int] -or $v -is [double]) {
                    # error values come back as specific ints; check via .Text for # prefix
                }
                $t = $cell.Text
                if ($t -match "^#") {
                    $errCount++
                    if ($errSamples.Count -lt 5) {
                        $errSamples += ("  " + $ws.Name + "!" + $cell.Address + " = " + $t)
                    }
                }
            }
        }
        $lines += ($ws.Name + " errors=" + $errCount)
        foreach ($s in $errSamples) { $lines += $s }
    }
    $lines | Out-File -FilePath "C:\Users\dai86\.zcode\workspace\default\bond_rv\qa_scan.txt" -Encoding utf8
    Write-Output "scan done"
    $wb.Close($false)
} catch {
    Write-Output ("FAILED: " + $_.Exception.Message)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
