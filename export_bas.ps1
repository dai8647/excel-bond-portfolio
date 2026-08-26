$ErrorActionPreference = "Stop"
$src = "C:\Users\dai86\Downloads\BondRV_Model.xlsm"
$out = "C:\Users\dai86\.zcode\workspace\default\bond_rv\DataUpdater_original.bas"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($src)
    $vbp = $wb.VBProject
    $comp = $vbp.VBComponents.Item("DataUpdater")
    $comp.Export($out)
    Write-Output ("exported to " + $out)
    $wb.Close($false)
} catch {
    Write-Output ("FAILED: " + $_.Exception.Message)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
