$ErrorActionPreference = "Stop"
$dst = "C:\Users\dai86\Downloads\BondRV_Model.xlsm"
$bas = "C:\Users\dai86\.zcode\workspace\default\bond_rv\DataUpdater.bas"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($dst)
    Write-Output ("opened: sheets=" + $wb.Sheets.Count)

    # re-import updated VBA module
    $vbp = $wb.VBProject
    foreach ($comp in $vbp.VBComponents) {
        if ($comp.Name -eq "DataUpdater") { $vbp.VBComponents.Remove($comp) }
    }
    $vbp.VBComponents.Import($bas) | Out-Null
    Write-Output ("VBA re-imported. components=" + $vbp.VBComponents.Count)

    # run the silent update (fetches live data, rewrites history, recalcs)
    Write-Output "running UpdateDataSilent..."
    $xl.Run("UpdateDataSilent")
    Write-Output "macro returned"

    $h = $wb.Sheets.Item("履歴データ")
    Write-Output ("result S1 = " + $h.Range("S1").Text)
    # last data row
    $lastRow = $h.Cells($h.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
    Write-Output ("last data row = " + $lastRow)
    Write-Output ("last date A" + $lastRow + " = " + $h.Range("A" + $lastRow).Text)
    Write-Output ("UST10 D" + $lastRow + " = " + $h.Range("D" + $lastRow).Text)
    Write-Output ("JGB10 H" + $lastRow + " = " + $h.Range("H" + $lastRow).Text)
    Write-Output ("USDJPY O" + $lastRow + " = " + $h.Range("O" + $lastRow).Text)

    # KPI after recalc
    $d = $wb.Sheets.Item("ダッシュボード")
    Write-Output ("NAV B9 = " + $d.Range("B9").Text)
    Write-Output ("Sharpe B11 = " + $d.Range("B11").Text)

    # save
    $wb.Save()
    Write-Output "saved"
    $wb.Close($false)
} catch {
    Write-Output ("FAILED: " + $_.Exception.Message)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
