$ErrorActionPreference = "Stop"
$src = "C:\Users\dai86\Downloads\BondRV_Model.xlsx"
$dst = "C:\Users\dai86\Downloads\BondRV_Model.xlsm"
$bas = "C:\Users\dai86\.zcode\workspace\default\bond_rv\DataUpdater.bas"
if (Test-Path $dst) { Remove-Item $dst -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($src)
    Write-Output ("opened: sheets=" + $wb.Sheets.Count)

    # --- VBA module import ---
    $vbp = $wb.VBProject
    # remove existing DataUpdater if present
    foreach ($comp in $vbp.VBComponents) {
        if ($comp.Name -eq "DataUpdater") { $vbp.VBComponents.Remove($comp) }
    }
    $vbp.VBComponents.Import($bas) | Out-Null
    Write-Output ("VBA imported. components=" + $vbp.VBComponents.Count)

    # --- wire button on dashboard ---
    $d = $wb.Sheets.Item("ダッシュボード")
    # clear the placeholder cell visual (we will draw a real shape button)
    $d.Range("B4:C4").Clear()
    # remove any old shapes named RVUpdateBtn
    for ($i = $d.Shapes.Count; $i -ge 1; $i--) {
        if ($d.Shapes.Item($i).Name -eq "RVUpdateBtn") { $d.Shapes.Item($i).Delete() }
    }
    $d.Rows.Item(4).RowHeight = 32
    $rng = $d.Range("B4")
    $shp = $d.Shapes.AddShape(5, $rng.Left, ($rng.Top + 1), 260, 30)  # 5 = msoShapeRoundedRectangle
    $shp.Name = "RVUpdateBtn"
    $shp.Fill.ForeColor.RGB = 0x00805435  # will set properly below
    # green 548235 -> RGB(84,130,53); Excel RGB = R + G*256 + B*65536
    $shp.Fill.ForeColor.RGB = (84 + 130*256 + 53*65536)
    $shp.Line.Visible = $false
    $shp.TextFrame2.TextRange.Text = "▶ データ更新(ヒストリカル取得)"
    $shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (255 + 255*256 + 255*65536)
    $shp.TextFrame2.TextRange.Font.Bold = $true
    $shp.TextFrame2.TextRange.Font.Size = 13
    $shp.TextFrame2.VerticalAnchor = 1  # middle
    $shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # center
    $shp.OnAction = "UpdateData"
    $shp.ZOrder(0)  # bring to front
    Write-Output "button wired to UpdateData"

    # --- save as .xlsm (52 = xlOpenXMLWorkbookMacroEnabled) ---
    $wb.SaveAs($dst, 52)
    Write-Output ("saved: " + $dst)
    $wb.Close($false)
} catch {
    Write-Output ("FAILED: " + $_.Exception.Message)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
