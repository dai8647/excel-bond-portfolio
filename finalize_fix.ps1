$ErrorActionPreference = "Stop"
$xlsm = "C:\Users\dai86\Downloads\Capula_BondRV_Model.xlsm"
$bas  = "C:\Users\dai86\.zcode\workspace\default\capula_rv\DataUpdater.bas"
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlsm)

# --- 1) re-inject fixed VBA (portable log path) ---
$vbp = $wb.VBProject
foreach ($comp in $vbp.VBComponents) {
    if ($comp.Name -eq "DataUpdater") { $vbp.VBComponents.Remove($comp) }
}
$vbp.VBComponents.Import($bas) | Out-Null
Write-Output ("VBA re-imported. components=" + $vbp.VBComponents.Count)

# --- 2) fix the update button (was 13pt tall = invisible sliver) ---
$btn = $null
$dash = $null
foreach ($ws in $wb.Worksheets) {
    foreach ($sh in $ws.Shapes) {
        if ($sh.Name -eq "CapulaUpdateBtn") { $btn = $sh; $dash = $ws }
    }
}
if ($btn -ne $null) {
    $dash.Rows.Item(4).RowHeight = 32
    $anchor = $dash.Range("B4")
    $btn.Left = $anchor.Left
    $btn.Top = $anchor.Top + 1
    $btn.Width = 260
    $btn.Height = 30
    try { $btn.Fill.ForeColor.RGB = (84 + 130*256 + 53*65536) } catch {}
    try { $btn.Line.Visible = $false } catch {}
    try {
        $btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (255 + 255*256 + 255*65536)
        $btn.TextFrame2.TextRange.Font.Bold = $true
        $btn.TextFrame2.TextRange.Font.Size = 13
        $btn.TextFrame2.VerticalAnchor = 1
        $btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    } catch {}
    $btn.OnAction = "UpdateData"
    try { $btn.ZOrder(0) } catch {}
    Write-Output ("button fixed: top=" + $btn.Top + " left=" + $btn.Left + " w=" + $btn.Width + " h=" + $btn.Height + " onaction=" + $btn.OnAction)
} else {
    Write-Output "BUTTON NOT FOUND"
}

$wb.Save()
Write-Output "saved"
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "done"
