$ErrorActionPreference = "Stop"
$src = "C:\Users\dai86\Downloads\Capula_BondRV_Model.xlsm"
$tmp = "C:\Users\dai86\.zcode\workspace\default\capula_rv\standalone_test.xlsm"
$log = "C:\Users\dai86\.zcode\workspace\default\capula_rv\vba_progress.log"
$out = "C:\Users\dai86\.zcode\workspace\default\capula_rv\standalone_result.txt"
Copy-Item $src $tmp -Force
if (Test-Path $log) { Remove-Item $log -Force }
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow = enable macros for this test
$wb = $excel.Workbooks.Open($tmp)
$lines = @()
$lines += ("opened copy, sheets=" + $wb.Worksheets.Count)
$t0 = Get-Date
$excel.Run("standalone_test.xlsm!UpdateDataSilent")
$elapsed = (Get-Date) - $t0
$lines += ("UpdateDataSilent returned after " + [int]$elapsed.TotalSeconds + "s")
$hist = $wb.Worksheets.Item(9)
$lines += ("hist sheet name=" + $hist.Name)
$lines += ("T1 result=" + $hist.Range("T1").Text)
$lines += ("colA count=" + $excel.WorksheetFunction.CountA($hist.Columns.Item(1)))
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
$lines += ""
$lines += "=== vba_progress.log ==="
if (Test-Path $log) { $lines += (Get-Content $log) } else { $lines += "(no log file)" }
$lines | Out-File -FilePath $out -Encoding utf8
Remove-Item $tmp -Force
Write-Host "done"
