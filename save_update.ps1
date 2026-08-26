$ErrorActionPreference = "Stop"
$src = "C:\Users\dai86\Downloads\BondRV_Model.xlsm"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($src)
    Write-Output "running UpdateDataSilent..."
    $xl.Application.Run("UpdateDataSilent")
    $wb.Save()
    Write-Output "saved"
    $ws = $wb.Sheets.Item(8)
    $status = $ws.Range("T1").Text
    Write-Output ("Status: " + $status)
    $wsAsw = $wb.Sheets.Item(11)
    Write-Output ("ASW treasury10Y: " + $wsAsw.Cells(13, 2).Text)
    Write-Output ("ASW oisSwap10Y: " + $wsAsw.Cells(18, 2).Text)
    $wb.Close($false)
} catch {
    Write-Output ("FAILED: " + $_.Exception.Message)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
