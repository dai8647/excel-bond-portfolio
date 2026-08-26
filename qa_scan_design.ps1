$ErrorActionPreference = "Continue"
$xlsm = "C:\Users\dai86\Downloads\Capula_BondRV_Model.xlsm"
$rep  = "C:\Users\dai86\.zcode\workspace\default\capula_rv\qa_design_report.txt"
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlsm)
$L = New-Object System.Collections.Generic.List[string]

function Add($s) { $L.Add($s) }

Add ("=== QA DESIGN " + (Get-Date -Format "yyyy-MM-dd HH:mm") + " ===")
try { Add ("workbook normal style font = " + $wb.Styles.Item("Normal").Font.Name + " size=" + $wb.Styles.Item("Normal").Font.Size) } catch { try { Add ("workbook normal style font = " + $wb.Styles.Item(1).Font.Name) } catch { Add "normal style err" } }

# ---- per-sheet basics ----
for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
    $ws = $wb.Worksheets.Item($i)
    $tab = try { $ws.Tab.Color } catch { "?" }
    $grid = try { $ws.DisplayGridlines } catch { "?" }
    $frz = try { if ($excel.ActiveWindow -ne $null) { "" } else { "" } } catch { "" }
    $charts = try { $ws.ChartObjects().Count } catch { 0 }
    $shapes = try { $ws.Shapes.Count } catch { 0 }
    Add ("S" + $i + " | " + $ws.Name + " | tab=" + $tab + " grid=" + $grid + " charts=" + $charts + " shapes=" + $shapes)
}

# ---- freeze panes per sheet via activewindow ----
for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
    $ws = $wb.Worksheets.Item($i)
    try {
        $ws.Activate() | Out-Null
        $aw = $excel.ActiveWindow
        Add ("S" + $i + " freeze=" + $aw.FreezePanes + " splitRow=" + $aw.SplitRow + " zoom=" + $aw.Zoom)
    } catch { Add ("S" + $i + " freeze err") }
}

# ---- charts detail ----
foreach ($idx in @(2,4,5,10,11)) {
    $ws = $wb.Worksheets.Item($idx)
    $cos = $ws.ChartObjects()
    for ($c = 1; $c -le $cos.Count; $c++) {
        $ch = $cos.Item($c).Chart
        $sn = try { $ch.SeriesCollection().Count } catch { -1 }
        $tt = "none"
        try { if ($ch.HasTitle) { $tt = $ch.ChartTitle.Text } } catch { $tt = "err" }
        Add ("CHART S" + $idx + "#" + $c + " type=" + $ch.ChartType + " series=" + $sn + " title=" + $tt)
        if ($sn -gt 0 -and $sn -le 6) {
            for ($s2 = 1; $s2 -le $sn; $s2++) {
                try {
                    $ser = $ch.SeriesCollection($s2)
                    Add ("  ser" + $s2 + " name=" + $ser.Name + " formula=" + $ser.Formula)
                } catch { Add ("  ser" + $s2 + " err") }
            }
        }
    }
}

# ---- workbook names ----
Add "--- names ---"
foreach ($n in $wb.Names) {
    try { Add ($n.Name + " -> " + $n.RefersTo) } catch { Add ($n.Name + " -> err") }
}

# ---- conditional formatting spot checks ----
Add "--- conditional formats ---"
$spots = @(
    @(2, "D9:E9"), @(2, "G16:G23"),
    @(5, "E10:G16"),
    @(6, "C6:E14"),
    @(7, "I4:I20"),
    @(14, "C4:C12"),
    @(16, "F13:F20"),
    @(16, "B46:F54")
)
foreach ($sp in $spots) {
    try {
        $fc = $wb.Worksheets.Item($sp[0]).Range($sp[1]).FormatConditions.Count
        Add ("S" + $sp[0] + " " + $sp[1] + " fc=" + $fc)
    } catch { Add ("S" + $sp[0] + " " + $sp[1] + " fc err") }
}

# ---- key cell checks ----
Add "--- key cells ---"
try { Add ("dash B10 label = " + $wb.Worksheets.Item(2).Range("B10").Text) } catch {}
try { Add ("dash A2 formula = " + $wb.Worksheets.Item(2).Range("A2").Formula) } catch {}
try { Add ("ledger H5 label = " + $wb.Worksheets.Item(6).Range("H5").Text) } catch {}
try { Add ("intro B4 = " + $wb.Worksheets.Item(1).Range("B4").Text) } catch {}
try { Add ("sheet8 A1 hdr = " + $wb.Worksheets.Item(8).Range("A1").Text + " / B1 = " + $wb.Worksheets.Item(8).Range("B1").Text) } catch {}
try { Add ("sheet9 A1 hdr = " + $wb.Worksheets.Item(9).Range("A1").Text) } catch {}
try { Add ("yield B3 = " + $wb.Worksheets.Item(10).Range("B3").Text) } catch {}

# ---- numberformat spot checks ----
Add "--- numberformats ---"
try { Add ("dash B9 fmt = " + $wb.Worksheets.Item(2).Range("B9").NumberFormat) } catch {}
try { Add ("sig C5 fmt = " + $wb.Worksheets.Item(5).Range("C5").NumberFormat) } catch {}
try { Add ("risk C6 fmt = " + $wb.Worksheets.Item(16).Range("C6").NumberFormat) } catch {}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.IO.File]::WriteAllLines($rep, $L, [System.Text.Encoding]::UTF8)
Write-Output "QA DONE"
