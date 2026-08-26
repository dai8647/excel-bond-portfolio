$ErrorActionPreference = "Stop"
$xlsm = "C:\Users\dai86\Downloads\BondRV_Model.xlsm"
$bak  = "C:\Users\dai86\Downloads\BondRV_Model.bak-design.xlsm"
Copy-Item $xlsm $bak -Force
Write-Output "backup ok"

# palette (COM Color = R + G*256 + B*65536)
$NAVY  = 6567967    # 1F3864
$BLUE  = 12874308   # 4472C4
$LIGHT = 15917785   # D9E2F2
$PALE  = 16512755   # F3F6FB
$GRAYF = 5855577    # 595959
$BORD  = 12566463   # BFBFBF
$GREEN = 3506772    # 548235
$RED   = 192        # C00000
$LGRAY = 14277081   # D9D9D9
$MGRAY = 8421504    # 808080
$WHITE = 16777215
$YELL  = 13431551   # FFF2CC input

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$wb = $excel.Workbooks.Open($xlsm)
$WBN = $wb.Name

function Band($ws, $a1, $a2) {
    try {
        $r = $ws.Range($a1 + ":" + $a2)
        $r.Interior.Color = $NAVY
        $r.Font.Color = $WHITE
        $r.Font.Bold = $true
        $r.Font.Size = 10.5
        $r.VerticalAlignment = -4108
        $ws.Rows.Item($r.Row).RowHeight = 18
    } catch { Write-Output ("band err " + $a1) }
}
function THead($ws, $a1, $a2) {
    try {
        $r = $ws.Range($a1 + ":" + $a2)
        $r.Interior.Color = $NAVY
        $r.Font.Color = $WHITE
        $r.Font.Bold = $true
        $r.Font.Size = 10
        $r.HorizontalAlignment = -4108
        $r.VerticalAlignment = -4108
        $r.WrapText = $true
    } catch { Write-Output ("thead err " + $a1) }
}
function Borders($ws, $a1, $a2) {
    try {
        $r = $ws.Range($a1 + ":" + $a2)
        $r.Borders.LineStyle = 1
        $r.Borders.Weight = 2
        $r.Borders.Color = $BORD
    } catch { Write-Output ("bord err " + $a1) }
}
function RedNeg($ws, $a1, $a2) {
    try {
        $r = $ws.Range($a1 + ":" + $a2)
        $r.FormatConditions.Delete()
        $fc = $r.FormatConditions.Add(1, 6, "0")
        $fc.Font.Color = 192
    } catch { Write-Output ("redneg err " + $a1) }
}
function NoGrid($ws) {
    try { $ws.Activate(); $excel.ActiveWindow.DisplayGridlines = $false } catch {}
}
function Freeze($ws, $row, $col) {
    try {
        $ws.Activate()
        $aw = $excel.ActiveWindow
        $aw.SplitRow = $row
        $aw.SplitColumn = $col
        $aw.FreezePanes = $true
    } catch { Write-Output "freeze err" }
}
function StyleChart($ch, $title, $tcolor) {
    try {
        $ch.ChartArea.Font.Name = "Meiryo UI"
        $ch.ChartArea.Font.Size = 9
        $ch.ChartArea.Format.Line.ForeColor.RGB = 15132390
        $ch.ChartArea.Format.Line.Weight = 0.75
        if ($title -ne "") {
            $ch.HasTitle = $true
            $ch.ChartTitle.Text = $title
            $ch.ChartTitle.Font.Size = 11
            $ch.ChartTitle.Font.Bold = $true
            $ch.ChartTitle.Font.Color = $tcolor
        }
        try {
            $ch.Axes(2).HasMajorGridlines = $true
            $ch.Axes(2).MajorGridlines.Format.Line.Color = $LGRAY
            $ch.Axes(2).TickLabels.NumberFormat = "General"
            $ch.Axes(1).TickLabels.Font.Size = 8
        } catch {}
        try { $ch.PlotArea.Format.Line.Visible = 0 } catch {}
    } catch { Write-Output ("chartstyle err " + $title) }
}

# ============ 0. global ============
try {
    $wb.Styles.Item("Normal").Font.Name = "Meiryo UI"
    $wb.Styles.Item("Normal").Font.Size = 10
} catch { Write-Output "normal style err" }
foreach ($ws in $wb.Worksheets) {
    try { $ws.UsedRange.Font.Name = "Meiryo UI" } catch { Write-Output ("font err " + $ws.Name) }
}
Write-Output "global font ok"

# ============ 1. tab colors ============
$tabs = @{1=5855577; 2=6567967; 3=12874308; 4=1137349; 5=3243501; 6=3506515; 7=4697456; 8=11572868; 9=11572868; 10=12553984; 11=49407; 12=10921638; 13=6740479; 14=9359529; 15=10921638; 16=192}
foreach ($ws in $wb.Worksheets) {
    try { $ws.Tab.Color = $tabs[[int]$ws.Index] } catch {}
}
Write-Output "tabs ok"

# dynamic date subtitle
$d = '="債券RV — 仮NAV・リスク指標・戦略サマリ / 基準日 "&TEXT(MAX(履歴データ!$A:$A),"yyyy-mm-dd")'
$d2 = '="v3.0 / マクロ(VBA)でデータ更新 / 基準日 "&TEXT(MAX(履歴データ!$A:$A),"yyyy-mm-dd")'

# ============ 2. はじめに (sheet1) ============
$ws = $wb.Worksheets.Item(1)
NoGrid $ws
$ws.Columns.Item("A").ColumnWidth = 3
$ws.Columns.Item("B").ColumnWidth = 112
$ws.Columns.Item("C").ColumnWidth = 46
try {
    $ws.Range("A1").Font.Size = 18
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1").Font.Color = $NAVY
    $ws.Rows.Item(1).RowHeight = 30
    $ws.Range("A2").Formula = $d2
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Range("A2").Font.Color = $GRAYF
    foreach ($hr in @(4,7,10,13,16,19,22,25,28)) {
        $c = $ws.Range("B" + $hr)
        $c.Font.Bold = $true
        $c.Font.Size = 11.5
        $c.Font.Color = $NAVY
    }
    foreach ($br in @(5,8,11,14,17,20,23,26,29)) {
        $c = $ws.Range("B" + $br)
        $c.Font.Color = $GRAYF
        $c.Font.Size = 10
        $c.WrapText = $true
        $c.VerticalAlignment = -4160
    }
    $hts = @(@(5,32), @(8,46), @(11,32), @(14,46), @(17,46), @(20,32), @(23,46), @(26,46), @(29,46))
    foreach ($hh in $hts) { $ws.Rows.Item($hh[0]).RowHeight = $hh[1] }
} catch { Write-Output ("intro err " + $_.Exception.Message) }
# TOC
try {
    $ws.Range("B31").Value2 = "■ シート一覧"
    Band $ws "B31" "C31"
    $ws.Range("B32").Value2 = "シート"
    $ws.Range("C32").Value2 = "役割"
    THead $ws "B32" "C32"
    $toc = @(
        @("ダッシュボード", "全体状況・KPI・データ更新ボタン"),
        @("戦略解説", "8戦略のロジックとヘッジの考え方"),
        @("エントリーシグナル", "各戦略のZスコアと±2σバンド"),
        @("シグナル検証", "±2σルールのバックテスト(勝率・DD)"),
        @("トレード台帳", "戦略別収益表・建玉・クローズド取引"),
        @("ポートフォリオ", "レッグ別建玉とDV01内訳"),
        @("分析データ", "スプレッド/P&L計算エンジン(数式)"),
        @("履歴データ", "唯一のデータソース(VBAが更新)"),
        @("イールドカーブ", "日米カーブの最新値"),
        @("ASW", "アセットスワップスプレッド監視"),
        @("ASWデータ", "Pensford取得値(自動更新)"),
        @("ファンディング", "レポ/資金調達コスト想定"),
        @("損益", "キャリー内訳(年率)"),
        @("銘柄マスター", "銘柄・リスク係数・BBGティッカー"),
        @("リスク", "VaR/ストレス/グリークス/リミット")
    )
    $i = 33
    foreach ($t in $toc) {
        $ws.Range("B" + $i).Value2 = $t[0]
        $ws.Range("C" + $i).Value2 = $t[1]
        $i++
    }
    Borders $ws "B32" "C47"
    for ($r = 34; $r -le 47; $r += 2) { $ws.Range("B" + $r + ":C" + $r).Interior.Color = $PALE }
    foreach ($r in 33..47) { $ws.Range("B" + $r).Font.Bold = $true; $ws.Range("B" + $r).Font.Color = $NAVY; $ws.Range("C" + $r).Font.Size = 9.5; $ws.Range("C" + $r).Font.Color = $GRAYF }
    $ws.Rows.Item(31).RowHeight = 18
} catch { Write-Output ("toc err " + $_.Exception.Message) }
Write-Output "sheet1 ok"

# ============ 3. ダッシュボード (sheet2) ============
$ws = $wb.Worksheets.Item(2)
NoGrid $ws
Freeze $ws 2 0
$ws.Columns.Item("A").ColumnWidth = 2.5
$ws.Columns.Item("B").ColumnWidth = 33
$ws.Columns.Item("C").ColumnWidth = 13
$ws.Columns.Item("D").ColumnWidth = 13
$ws.Columns.Item("E").ColumnWidth = 11
$ws.Columns.Item("F").ColumnWidth = 21
$ws.Columns.Item("G").ColumnWidth = 13
$ws.Columns.Item("H").ColumnWidth = 8
try {
    $ws.Range("A1").Value2 = "債券RVモデル — ダッシュボード"
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Formula = $d
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch { Write-Output ("dash title err " + $_.Exception.Message) }
try { $ws.Range("B5").Font.Italic = $true; $ws.Range("B5").Font.Size = 9; $ws.Range("B5").Font.Color = $GRAYF } catch {}
Band $ws "B7" "G7"
Band $ws "B14" "G14"
# KPI cards
try {
    $cards = @(@("B","C"), @("D","E"), @("F","G"))
    foreach ($pair in $cards) {
        $k = $pair[0]
        $v = $pair[1]
        $lab = $ws.Range($k + "8:" + $v + "8")
        $lab.Interior.Color = $BLUE
        $lab.Font.Color = $WHITE
        $lab.Font.Bold = $true
        $lab.Font.Size = 9
        $lab.HorizontalAlignment = -4108
        $val = $ws.Range($k + "9:" + $v + "9")
        $val.Interior.Color = $LIGHT
        $val.Font.Color = $NAVY
        $val.Font.Bold = $true
        $val.Font.Size = 15
        $val.HorizontalAlignment = -4108
        $lab2 = $ws.Range($k + "10:" + $v + "10")
        $lab2.Interior.Color = $BLUE
        $lab2.Font.Color = $WHITE
        $lab2.Font.Bold = $true
        $lab2.Font.Size = 9
        $lab2.HorizontalAlignment = -4108
        $val2 = $ws.Range($k + "11:" + $v + "11")
        $val2.Interior.Color = $LIGHT
        $val2.Font.Color = $NAVY
        $val2.Font.Bold = $true
        $val2.Font.Size = 15
        $val2.HorizontalAlignment = -4108
    }
    $ws.Range("B9").NumberFormat = "#,##0.0"
    $ws.Range("D9").NumberFormat = "0.0%;-0.0%"
    $ws.Range("F9").NumberFormat = "0.0%;-0.0%"
    $ws.Range("B11").NumberFormat = "0.00;-0.00"
    $ws.Range("D11").NumberFormat = "0.0%;-0.0%"
    $ws.Range("F11").NumberFormat = "0.00"
    $ws.Rows.Item(9).RowHeight = 26
    $ws.Rows.Item(11).RowHeight = 26
    $ws.Range("B10").Value2 = "シャープレシオ"
    $ws.Range("B8:C8").MergeCells = $false
    $ws.Range("B12").RowHeight = 8
    $ws.Range("B13").RowHeight = 6
} catch { Write-Output ("kpi err " + $_.Exception.Message) }
# strategy table
try {
    THead $ws "B15" "G15"
    Borders $ws "B16" "G23"
    for ($r = 17; $r -le 23; $r += 2) { $ws.Range("B" + $r + ":G" + $r).Interior.Color = $PALE }
    $ws.Range("B16:G23").Font.Size = 10
    $ws.Range("C16:D23").HorizontalAlignment = -4108
    $ws.Range("E16:E23").HorizontalAlignment = -4108
    $ws.Range("E16:E23").Font.Bold = $true
    $ws.Range("G16:G23").NumberFormat = "#,##0;-#,##0"
    $ws.Range("G16:G23").HorizontalAlignment = -4152
    $ws.Range("D16:D23").NumberFormat = "0.00"
    foreach ($r in 16..23) { $ws.Rows.Item($r).RowHeight = 17 }
    $ws.Range("D16:D23").FormatConditions.Delete()
    $cs = $ws.Range("D16:D23").FormatConditions.AddColorScale(3)
    $cs.ColorScaleCriteria(1).FormatColor.Color = 16279915
    $cs.ColorScaleCriteria(2).FormatColor.Color = $WHITE
    $cs.ColorScaleCriteria(3).FormatColor.Color = 13011546
} catch { Write-Output ("stable err " + $_.Exception.Message) }
RedNeg $ws "D9" "E9"
RedNeg $ws "F9" "G9"
RedNeg $ws "B11" "C11"
RedNeg $ws "D11" "E11"
RedNeg $ws "G16" "G23"
try { $ws.Range("B25").Font.Italic = $true; $ws.Range("B25").Font.Size = 9; $ws.Range("B25").Font.Color = $GRAYF } catch {}
Write-Output "sheet2 cells ok"

# ---- dynamic named ranges for charts (auto-extend with data) ----
$nerrs = 0
foreach ($n in @(
    @("DATE_X",  '=分析データ!$A$2:INDEX(分析データ!$A:$A,COUNT(分析データ!$A:$A)+1)'),
    @("NAV_Y",   '=分析データ!$AE$2:INDEX(分析データ!$AE:$AE,COUNT(分析データ!$A:$A)+1)'),
    @("DD_Y",    '=分析データ!$AH$2:INDEX(分析データ!$AH:$AH,COUNT(分析データ!$A:$A)+1)'),
    @("ESPR_Y",  '=分析データ!$B$2:INDEX(分析データ!$B:$B,COUNT(分析データ!$A:$A)+1)'),
    @("EBAND_U", '=分析データ!$AJ$2:INDEX(分析データ!$AJ:$AJ,COUNT(分析データ!$A:$A)+1)'),
    @("EBAND_L", '=分析データ!$AK$2:INDEX(分析データ!$AK:$AK,COUNT(分析データ!$A:$A)+1)')
)) {
    try { $wb.Names.Add($n[0], $n[1]) | Out-Null } catch { Write-Output ("name err " + $n[0]); $nerrs++ }
}
Write-Output ("names done errors=" + $nerrs)

# ---- KPI card merges ----
foreach ($p in @('B8:C8','B9:C9','D8:E8','D9:E9','F8:G8','F9:G9','B10:C10','B11:C11','D10:E10','D11:E11','F10:G10','F11:G11')) {
    try { $ws.Range($p).Merge() | Out-Null } catch {}
}

# ---- dashboard charts: delete & rebuild ----
$ws2 = $wb.Worksheets.Item(2)
$wsD = $wb.Worksheets.Item(8)
$wsL = $wb.Worksheets.Item(6)
try {
    $old = @($ws2.ChartObjects())
    for ($i = $old.Count - 1; $i -ge 0; $i--) { $old[$i].Delete() }
    $x1 = [double]$ws2.Range("B27").Left
    $y1 = [double]$ws2.Range("B27").Top
    $x3 = [double]$ws2.Range("I27").Left
    $y2 = [double]$ws2.Range("B45").Top

    $co = $ws2.ChartObjects().Add($x1, $y1, 470, 232)
    $ch = $co.Chart
    $ch.ChartType = 4
    $ch.SetSourceData($wsD.Range("AE1:AE911"), 2)
    $ser = $ch.SeriesCollection(1)
    $ser.Formula = ('=SERIES(分析データ!$AE$1,' + $WBN + '!DATE_X,' + $WBN + '!NAV_Y,1)')
    StyleChart $ch '仮NAV推移 ($mm) — バックテスト' $NAVY
    $ch.HasLegend = $false
    $ser.Format.Line.ForeColor.RGB = $NAVY
    $ser.Format.Line.Weight = 2.25
    try { $ser.MarkerStyle = -4142 } catch {}
    try { $ch.Axes(1).TickLabels.NumberFormat = 'yy-mm' } catch {}
    try { $ch.Axes(2).TickLabels.NumberFormat = '#,##0' } catch {}

    $co = $ws2.ChartObjects().Add($x1, $y2, 470, 212)
    $ch = $co.Chart
    $ch.ChartType = 4
    $ch.SetSourceData($wsD.Range("AH1:AH911"), 2)
    $ser = $ch.SeriesCollection(1)
    $ser.Formula = ('=SERIES(分析データ!$AH$1,' + $WBN + '!DATE_X,' + $WBN + '!DD_Y,1)')
    StyleChart $ch 'ドローダウン (%) — NAVピーク比' $RED
    $ch.HasLegend = $false
    $ser.Format.Line.ForeColor.RGB = $RED
    $ser.Format.Line.Weight = 1.75
    try { $ser.MarkerStyle = -4142 } catch {}
    try { $ch.Axes(1).TickLabels.NumberFormat = 'yy-mm' } catch {}

    $co = $ws2.ChartObjects().Add($x3, $y1, 430, 232)
    $ch = $co.Chart
    $ch.ChartType = 51
    $ch.SetSourceData($wsL.Range("E5:E13"), 2)
    $ser = $ch.SeriesCollection(1)
    try { $ser.XValues = $wsL.Range("B6:B13") } catch {}
    StyleChart $ch '戦略別 合計損益 ($, バックテスト)' $NAVY
    $ch.HasLegend = $false
    try {
        $ser.Format.Fill.ForeColor.RGB = $BLUE
        $ch.ChartGroups(1).GapWidth = 60
        for ($i = 1; $i -le 8; $i++) {
            $v = $wsL.Cells.Item(5 + $i, 5).Value2
            if ($v -lt 0) { $ser.Points.Item($i).Format.Fill.ForeColor.RGB = $RED }
        }
    } catch {}
    Write-Output "dash charts ok"
} catch { Write-Output ("dash charts ERR " + $_.Exception.Message) }

# ---- entry signal chart: dynamic formulas + restyle ----
$ws4 = $wb.Worksheets.Item(4)
try {
    $ch4 = $ws4.ChartObjects(1).Chart
    $s1 = $ch4.SeriesCollection(1); $s2 = $ch4.SeriesCollection(2); $s3 = $ch4.SeriesCollection(3)
    $s1.Formula = ('=SERIES(分析データ!$B$1,' + $WBN + '!DATE_X,' + $WBN + '!ESPR_Y,1)')
    $s2.Formula = ('=SERIES(分析データ!$AJ$1,' + $WBN + '!DATE_X,' + $WBN + '!EBAND_U,2)')
    $s3.Formula = ('=SERIES(分析データ!$AK$1,' + $WBN + '!DATE_X,' + $WBN + '!EBAND_L,3)')
    $ch4.ChartArea.Font.Name = "Meiryo UI"
    $ch4.ChartArea.Font.Size = 9
    $ch4.HasTitle = $true
    $ch4.ChartTitle.Font.Size = 11
    $ch4.ChartTitle.Font.Bold = $true
    $ch4.ChartTitle.Font.Color = $NAVY
    $s1.Format.Line.ForeColor.RGB = $NAVY
    $s1.Format.Line.Weight = 2
    foreach ($sb in @($s2, $s3)) {
        $sb.Format.Line.ForeColor.RGB = $MGRAY
        $sb.Format.Line.Weight = 1
        try { $sb.Format.Line.DashStyle = 4 } catch {}
    }
    foreach ($sx in @($s1, $s2, $s3)) { try { $sx.MarkerStyle = -4142 } catch {} }
    try { $ch4.Legend.Position = -4107 } catch {}
    try { $ch4.Axes(1).TickLabels.NumberFormat = 'yy-mm' } catch {}
    Write-Output "entry chart ok"
} catch { Write-Output ("entry chart ERR " + $_.Exception.Message) }

# ---- sigverify chart: rebuild with cumulative series ----
$ws5 = $wb.Worksheets.Item(5)
try {
    $olds = @($ws5.ChartObjects())
    for ($i = $olds.Count - 1; $i -ge 0; $i--) { $olds[$i].Delete() }
    $co = $ws5.ChartObjects().Add([double]$ws5.Range("K9").Left, [double]$ws5.Range("K9").Top, 520, 300)
    $ch = $co.Chart
    $ch.ChartType = 4
    $ch.PlotVisibleOnly = $false
    $ch.SetSourceData($ws5.Range("CE2:CE1307"), 2)
    $ser = $ch.SeriesCollection(1)
    $ser.Formula = '=SERIES("累積損益",シグナル検証!$CH$2:$CH$1307,シグナル検証!$CE$2:$CE$1307,1)'
    $ch.DisplayBlanksAs = 1
    StyleChart $ch 'シグナル戦略 累積損益($) — 全戦略合算・参考' $GREEN
    $ch.HasLegend = $false
    $ser.Format.Line.ForeColor.RGB = $GREEN
    $ser.Format.Line.Weight = 2
    try { $ser.MarkerStyle = -4142 } catch {}
    try { $ch.Axes(1).TickLabels.NumberFormat = 'yy-mm' } catch {}
    Write-Output "sig chart ok"
} catch { Write-Output ("sig chart ERR " + $_.Exception.Message) }

# ============ sheet10 イールドカーブ ============
$ws10 = $wb.Worksheets.Item(10)
NoGrid $ws10
$ws10.Columns.Item("A").ColumnWidth = 2
$ws10.Columns.Item("B").ColumnWidth = 10
$ws10.Columns.Item("C").ColumnWidth = 11
$ws10.Columns.Item("D").ColumnWidth = 11
$ws10.Columns.Item("E").ColumnWidth = 15
try {
    $ws10.Range("A1:A2").Interior.Color = $NAVY
    $ws10.Range("A1").Font.Color = $WHITE
    $ws10.Range("A1").Font.Size = 13
    $ws10.Range("A1").Font.Bold = $true
    $ws10.Rows.Item(1).RowHeight = 26
    $ws10.Range("A2").Font.Color = $LIGHT
    $ws10.Range("A2").Font.Italic = $true
    $ws10.Range("A2").Font.Size = 9
    $ws10.Rows.Item(2).RowHeight = 15
} catch {}
$ws10.Range("B3").Value2 = "年限"
$ws10.Range("C3").Value2 = "米国債%"
$ws10.Range("D3").Value2 = "JGB%"
$ws10.Range("E3").Value2 = "スプレッド(bp)"
THead $ws10 "B3" "E3"
Borders $ws10 "B3" "E8"
for ($r = 5; $r -le 7; $r += 2) { $ws10.Range("B" + $r + ":E" + $r).Interior.Color = $PALE }
$ws10.Range("C4:D8").NumberFormat = "0.000"
$ws10.Range("E4:E8").NumberFormat = "#,##0.0"
foreach ($r in 4..8) { $ws10.Rows.Item($r).RowHeight = 18 }
try { $ws10.Range("B10").Font.Italic = $true; $ws10.Range("B10").Font.Size = 9; $ws10.Range("B10").Font.Color = $GRAYF } catch {}
try {
    $oldc = @($ws10.ChartObjects())
    for ($i = $oldc.Count - 1; $i -ge 0; $i--) { $oldc[$i].Delete() }
    $co = $ws10.ChartObjects().Add([double]$ws10.Range("G3").Left, [double]$ws10.Range("G3").Top, 440, 260)
    $ch = $co.Chart
    $ch.ChartType = 4
    $ch.SetSourceData($ws10.Range("B3:D8"), 2)
    StyleChart $ch 'イールドカーブ (%) — 米国債 vs JGB' $NAVY
    try { $ch.Legend.Position = -4107 } catch {}
    $ch.SeriesCollection(1).Format.Line.ForeColor.RGB = $BLUE
    $ch.SeriesCollection(1).Format.Line.Weight = 2
    $ch.SeriesCollection(2).Format.Line.ForeColor.RGB = $RED
    $ch.SeriesCollection(2).Format.Line.Weight = 2
    foreach ($sx in @($ch.SeriesCollection(1), $ch.SeriesCollection(2))) { try { $sx.MarkerStyle = -4142 } catch {} }
    try { $ch.Axes(2).TickLabels.NumberFormat = '0.0' } catch {}
    Write-Output "yield ok"
} catch { Write-Output ("yield ERR " + $_.Exception.Message) }

# ---- ASW chart restyle ----
$ws11 = $wb.Worksheets.Item(11)
try {
    $cha = $ws11.ChartObjects(1).Chart
    StyleChart $cha '' 0
    $cha.ChartTitle.Font.Color = $NAVY
    $cha.ChartTitle.Font.Bold = $true
    $cha.ChartTitle.Font.Size = 11
    $sera = $cha.SeriesCollection(1)
    $sera.Format.Fill.ForeColor.RGB = $BLUE
    try { $cha.ChartGroups(1).GapWidth = 50 } catch {}
    try { $cha.Axes(1).TickLabels.Font.Size = 8 } catch {}
    Write-Output "asw chart ok"
} catch { Write-Output ("asw chart ERR " + $_.Exception.Message) }

# ---- strip stray shapes inside any chart ----
foreach ($wsx in $wb.Worksheets) {
    foreach ($cox in @($wsx.ChartObjects())) {
        try {
            $shc = $cox.Chart.Shapes
            if ($shc.Count -gt 0) {
                Write-Output ("inner shapes sheet=" + $wsx.Index + " n=" + $shc.Count)
                for ($i = $shc.Count; $i -ge 1; $i--) { try { Write-Output ("  del " + $shc.Item($i).Name); $shc.Item($i).Delete() } catch {} }
            }
        } catch {}
    }
}

# ============ sheet3 戦略解説 ============
$ws = $wb.Worksheets.Item(3)
NoGrid $ws
$ws.Columns.Item("A").ColumnWidth = 2
$ws.Columns.Item("B").ColumnWidth = 26
$ws.Columns.Item("C").ColumnWidth = 34
$ws.Columns.Item("D").ColumnWidth = 38
$ws.Columns.Item("E").ColumnWidth = 38
$ws.Columns.Item("F").ColumnWidth = 13
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
$ws.Range("B3").Value2 = "戦略"
$ws.Range("C3").Value2 = "ポジション"
$ws.Range("D3").Value2 = "なぜ儲かるか"
$ws.Range("E3").Value2 = "ヘッジ・リスク"
$ws.Range("F3").Value2 = "キャリー$/年"
THead $ws "B3" "F3"
Borders $ws "B3" "F11"
for ($r = 5; $r -le 11; $r += 2) { $ws.Range("B" + $r + ":F" + $r).Interior.Color = $PALE }
$ws.Range("B4:B11").Font.Bold = $true
$ws.Range("B4:B11").Font.Color = $NAVY
$ws.Range("C4:E11").WrapText = $true
$ws.Range("C4:E11").VerticalAlignment = -4160
foreach ($r in 4..11) { $ws.Rows.Item($r).RowHeight = 42 }
$ws.Range("F4:F11").NumberFormat = "#,##0;-#,##0"
$ws.Range("F4:F11").HorizontalAlignment = -4152
RedNeg $ws "F4" "F11"
try { $ws.Range("B13").Font.Italic = $true; $ws.Range("B13").Font.Size = 9; $ws.Range("B13").Font.Color = $GRAYF } catch {}
Write-Output "sheet3 ok"

# ============ sheet4 エントリーシグナル (cells) ============
$ws = $ws4
NoGrid $ws
$ws.Columns.Item("A").ColumnWidth = 2
$ws.Columns.Item("B").ColumnWidth = 30
$ws.Columns.Item("C").ColumnWidth = 11
$ws.Columns.Item("D").ColumnWidth = 11
$ws.Columns.Item("E").ColumnWidth = 11
$ws.Columns.Item("F").ColumnWidth = 10
$ws.Columns.Item("G").ColumnWidth = 10
$ws.Columns.Item("H").ColumnWidth = 10
$ws.Columns.Item("I").ColumnWidth = 10
$ws.Columns.Item("J").ColumnWidth = 56
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
Borders $ws "B4" "J9"
for ($r = 5; $r -le 9; $r += 2) { $ws.Range("B" + $r + ":J" + $r).Interior.Color = $PALE }
$ws.Range("B4:B9").Font.Bold = $true
$ws.Range("B4:B9").Font.Color = $NAVY
$ws.Range("C4:E9").NumberFormat = "#,##0.0"
$ws.Range("F4:F9").NumberFormat = "0.00"
$ws.Range("F4:F9").HorizontalAlignment = -4108
$ws.Range("F4:F9").Font.Bold = $true
$ws.Range("G4:H9").NumberFormat = "#,##0.0"
$ws.Range("I4:I9").HorizontalAlignment = -4108
$ws.Range("I4:I9").Font.Bold = $true
$ws.Range("J4:J9").WrapText = $true
$ws.Range("J4:J9").Font.Size = 9
$ws.Range("J4:J9").Font.Color = $GRAYF
$ws.Range("J4:J9").VerticalAlignment = -4160
foreach ($r in 4..9) { $ws.Rows.Item($r).RowHeight = 32 }
try { $ws.Range("B11:B12").Font.Italic = $true; $ws.Range("B11:B12").Font.Size = 9; $ws.Range("B11:B12").Font.Color = $GRAYF } catch {}
Write-Output "sheet4 ok"

# ============ sheet5 シグナル検証 (cells) ============
$ws = $ws5
NoGrid $ws
try {
    $ws.Range("B1").Font.Size = 14
    $ws.Range("B1").Font.Bold = $true
    $ws.Range("B1").Font.Color = $NAVY
    $ws.Range("B2").Font.Italic = $true
    $ws.Range("B2").Font.Size = 9
    $ws.Range("B2").Font.Color = $GRAYF
} catch {}
Band $ws "B4" "I4"
Band $ws "B18" "I18"
try {
    $ws.Range("B5:B7").Font.Bold = $true
    $ws.Range("C5:C7").Interior.Color = $LIGHT
    $ws.Range("C5:C7").HorizontalAlignment = -4108
    $ws.Range("C5:C7").Font.Bold = $true
    Borders $ws "C5" "C7"
} catch {}
THead $ws "B9" "I9"
Borders $ws "B9" "I16"
for ($r = 11; $r -le 15; $r += 2) { $ws.Range("B" + $r + ":I" + $r).Interior.Color = $PALE }
$ws.Range("C10:C16").NumberFormat = "#,##0"
$ws.Range("D10:D16").NumberFormat = "0.0%"
$ws.Range("E10:G16").NumberFormat = "#,##0;-#,##0"
$ws.Range("H10:H16").NumberFormat = "0.0"
RedNeg $ws "E10" "G16"
$ws.Range("I10:I16").Font.Bold = $true
$ws.Range("I10:I16").HorizontalAlignment = -4108
$ws.Range("B10:B16").Font.Bold = $true
try {
    $ws.Range("B16:I16").Font.Bold = $true
    $ws.Range("B16:I16").Borders.Item(8).Weight = 3
} catch {}
try {
    $ws.Range("B19:B20").Font.Bold = $true
    $ws.Range("C19").NumberFormat = "#,##0;-#,##0"
    $ws.Range("C20").NumberFormat = "0.0%;-0.0%"
    $ws.Range("C19:C20").Font.Bold = $true
    $ws.Range("C19:C20").Font.Color = $RED
} catch {}
try { $ws.Range("B24:B27").Font.Italic = $true; $ws.Range("B24:B27").Font.Size = 9; $ws.Range("B24:B27").Font.Color = $GRAYF } catch {}
try { $ws.Columns.Item("AT").EntireColumn.Hidden = $true } catch {}
Write-Output "sheet5 ok"

# ============ sheet6 トレード台帳 ============
$ws = $wsL
NoGrid $ws
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
try { $ws.Range("H5").Value2 = "シャープレシオ(日次)" } catch {}
Band $ws "B3" "K3"
Band $ws "B16" "K16"
Band $ws "B26" "K26"
THead $ws "B5" "J5"
Borders $ws "B5" "J14"
for ($r = 7; $r -le 13; $r += 2) { $ws.Range("B" + $r + ":J" + $r).Interior.Color = $PALE }
$ws.Range("C6:E14").NumberFormat = "#,##0;-#,##0"
$ws.Range("F6:F13").NumberFormat = "0.0%"
$ws.Range("G6:G13").NumberFormat = "#,##0"
$ws.Range("H6:H13").NumberFormat = "0.00"
$ws.Range("I6:I13").NumberFormat = "#,##0;-#,##0"
RedNeg $ws "C6" "E14"
RedNeg $ws "I6" "I13"
try {
    $ws.Range("B14:J14").Font.Bold = $true
    $ws.Range("B14:J14").Borders.Item(8).Weight = 3
} catch {}
THead $ws "B17" "I17"
Borders $ws "B17" "I24"
for ($r = 19; $r -le 23; $r += 2) { $ws.Range("B" + $r + ":I" + $r).Interior.Color = $PALE }
try {
    $ws.Range("C18:C24").Interior.Color = $YELL
    $ws.Range("C18:C24").Font.Italic = $true
    $ws.Range("C18:C24").Font.Size = 9
} catch {}
$ws.Range("D18:D24").NumberFormat = "#,##0.0"
$ws.Range("E18:E24").NumberFormat = "#,##0;-#,##0"
$ws.Range("F18:F24").NumberFormat = "#,##0"
RedNeg $ws "E18" "E24"
THead $ws "B27" "K27"
Borders $ws "B27" "K36"
try { $ws.Range("B28:K36").Interior.Color = $YELL } catch {}
$ws.Range("H28:J36").NumberFormat = "#,##0;-#,##0"
RedNeg $ws "H28" "J36"
try { $ws.Range("B37").Font.Italic = $true; $ws.Range("B37").Font.Size = 9; $ws.Range("B37").Font.Color = $GRAYF } catch {}
Freeze $ws 5 0
Write-Output "sheet6 ok"

# ============ sheet7 ポートフォリオ ============
$ws = $wb.Worksheets.Item(7)
NoGrid $ws
$ws.Columns.Item("A").ColumnWidth = 2
$ws.Columns.Item("B").ColumnWidth = 26
$ws.Columns.Item("C").ColumnWidth = 12
$ws.Columns.Item("D").ColumnWidth = 26
$ws.Columns.Item("E").ColumnWidth = 7
$ws.Columns.Item("F").ColumnWidth = 9
$ws.Columns.Item("G").ColumnWidth = 9
$ws.Columns.Item("H").ColumnWidth = 10
$ws.Columns.Item("I").ColumnWidth = 11
$ws.Columns.Item("J").ColumnWidth = 22
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
$ws.Range("B3").Value2 = "戦略"
$ws.Range("C3").Value2 = "ティッカー"
$ws.Range("D3").Value2 = "レッグ"
$ws.Range("E3").Value2 = "方向"
$ws.Range("F3").Value2 = "価格"
$ws.Range("G3").Value2 = "リスク係数"
$ws.Range("H3").Value2 = "DV01($/bp)"
$ws.Range("I3").Value2 = "ポジDV01($)"
$ws.Range("J3").Value2 = "役割"
THead $ws "B3" "J3"
Borders $ws "B3" "J20"
foreach ($r in @(7,8,9,12,13,15,16,19,20)) { $ws.Range("B" + $r + ":J" + $r).Interior.Color = $PALE }
$ws.Range("B4:B20").Font.Bold = $true
$ws.Range("B4:B20").Font.Color = $NAVY
$ws.Range("C4:C20").Font.Name = "Consolas"
$ws.Range("C4:C20").Font.Size = 9
$ws.Range("F4:G20").NumberFormat = "0.000"
$ws.Range("H4:H20").NumberFormat = "#,##0"
$ws.Range("I4:I20").NumberFormat = "#,##0;-#,##0"
$ws.Range("E4:E20").HorizontalAlignment = -4108
RedNeg $ws "I4" "I20"
Freeze $ws 3 0
Write-Output "sheet7 ok"

# ============ sheet8/9 data sheets ============
$ws = $wsD
try {
    $cnt = $ws.UsedRange.Columns.Count
    $h = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $cnt))
    $h.Interior.Color = $NAVY
    $h.Font.Color = $WHITE
    $h.Font.Bold = $true
} catch { Write-Output "hdr8 err" }
Freeze $ws 1 1
$ws = $wb.Worksheets.Item(9)
try {
    $cnt = $ws.UsedRange.Columns.Count
    $h = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $cnt))
    $h.Interior.Color = $NAVY
    $h.Font.Color = $WHITE
    $h.Font.Bold = $true
} catch { Write-Output "hdr9 err" }
Freeze $ws 1 1
Write-Output "sheet89 ok"

# ============ sheet11 ASW ============
$ws = $ws11
NoGrid $ws
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
foreach ($br in @(4,12,15,18,21,23)) { Band $ws ("B" + $br) ("H" + $br) }
THead $ws "B5" "E5"
Borders $ws "B5" "E10"
for ($r = 7; $r -le 9; $r += 2) { $ws.Range("B" + $r + ":E" + $r).Interior.Color = $PALE }
try { $ws.Range("B13").Font.Bold = $true; $ws.Range("B16").Font.Bold = $true } catch {}
try {
    $ws.Range("C14").Font.Bold = $true
    $ws.Range("C14").Font.Size = 12
    $ws.Range("C14").Font.Color = $NAVY
    $ws.Range("C17").Font.Bold = $true
    $ws.Range("C17").Font.Size = 12
    $ws.Range("C17").Font.Color = $NAVY
} catch {}
try { $ws.Range("B11:D11").Font.Italic = $true; $ws.Range("B11:D11").Font.Size = 9; $ws.Range("B11:D11").Font.Color = $GRAYF } catch {}
try { $ws.Range("D14:D17").Font.Italic = $true; $ws.Range("D14:D17").Font.Size = 9; $ws.Range("D14:D17").Font.Color = $GRAYF } catch {}
try {
    $ws.Range("B24:B28").Font.Bold = $true
    $ws.Range("B24:B28").Font.Color = $NAVY
    $ws.Range("C24:C28").Font.Color = 12673797
    $ws.Range("C24:C28").Font.Underline = 2
    $ws.Range("C24:C28").Font.Size = 9
    $ws.Range("D24:D28").Font.Size = 9
    $ws.Range("D24:D28").Font.Color = $GRAYF
    $ws.Range("D24:D28").WrapText = $true
    $ws.Range("B24:D28").Borders.Item(12).LineStyle = 1
    $ws.Range("B24:D28").Borders.Item(12).Weight = 2
    $ws.Range("B24:D28").Borders.Item(12).Color = $BORD
} catch {}
Write-Output "sheet11 ok"

# ============ sheet12 ASWデータ ============
$ws = $wb.Worksheets.Item(12)
NoGrid $ws
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
THead $ws "B3" "D3"
Borders $ws "B3" "D20"
for ($r = 4; $r -le 20; $r += 2) { $ws.Range("B" + $r + ":D" + $r).Interior.Color = $PALE }
$ws.Range("C4:C20").Font.Bold = $true
$ws.Range("C4:C20").Font.Color = $NAVY
$ws.Range("D4:D20").Font.Size = 9
$ws.Range("D4:D20").Font.Color = $GRAYF
$ws.Columns.Item("B").ColumnWidth = 16
$ws.Columns.Item("C").ColumnWidth = 12
$ws.Columns.Item("D").ColumnWidth = 42
Write-Output "sheet12 ok"

# ============ sheet13 ファンディング ============
$ws = $wb.Worksheets.Item(13)
NoGrid $ws
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
Band $ws "B9" "D9"
Borders $ws "B4" "D7"
for ($r = 5; $r -le 7; $r += 2) { $ws.Range("B" + $r + ":D" + $r).Interior.Color = $PALE }
$ws.Range("B4:B7").Font.Bold = $true
$ws.Range("B4:B7").Font.Color = $NAVY
$ws.Range("C4:C7").NumberFormat = "0.00"
$ws.Range("C4:C7").HorizontalAlignment = -4152
$ws.Range("C4:C7").Font.Bold = $true
$ws.Range("D4:D7").Font.Size = 9
$ws.Range("D4:D7").Font.Color = $GRAYF
Borders $ws "B11" "D15"
$ws.Range("B11:B15").Font.Bold = $true
try {
    $ws.Range("C11:C15").Font.Bold = $true
    $ws.Range("C11:C15").HorizontalAlignment = -4108
    $ws.Range("C11:C15").NumberFormat = "0.00"
} catch {}
$ws.Range("D11:D15").Font.Size = 9
$ws.Range("D11:D15").Font.Color = $GRAYF
try { $ws.Range("B17").Font.Italic = $true; $ws.Range("B17").Font.Size = 9; $ws.Range("B17").Font.Color = $GRAYF } catch {}
$ws.Columns.Item("B").ColumnWidth = 30
$ws.Columns.Item("C").ColumnWidth = 10
$ws.Columns.Item("D").ColumnWidth = 52
Write-Output "sheet13 ok"

# ============ sheet14 損益 ============
$ws = $wb.Worksheets.Item(14)
NoGrid $ws
$ws.Columns.Item("A").ColumnWidth = 2
$ws.Columns.Item("B").ColumnWidth = 26
$ws.Columns.Item("C").ColumnWidth = 13
$ws.Columns.Item("D").ColumnWidth = 2
$ws.Columns.Item("E").ColumnWidth = 48
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
$ws.Range("B3").Value2 = "戦略"
$ws.Range("C3").Value2 = "キャリー$/年"
$ws.Range("E3").Value2 = "備考"
THead $ws "B3" "C3"
THead $ws "E3" "E3"
Borders $ws "B3" "C12"
Borders $ws "E4" "E11"
for ($r = 5; $r -le 11; $r += 2) { $ws.Range("B" + $r + ":C" + $r).Interior.Color = $PALE }
$ws.Range("B4:B11").Font.Bold = $true
$ws.Range("B4:B11").Font.Color = $NAVY
$ws.Range("C4:C12").NumberFormat = "#,##0;-#,##0"
$ws.Range("C4:C12").HorizontalAlignment = -4152
RedNeg $ws "C4" "C12"
try {
    $ws.Range("B12:C12").Font.Bold = $true
    $ws.Range("B12:C12").Borders.Item(8).Weight = 3
} catch {}
$ws.Range("E4:E11").Font.Size = 9
$ws.Range("E4:E11").Font.Color = $GRAYF
$ws.Range("E4:E11").WrapText = $true
try { $ws.Range("B14").Font.Italic = $true; $ws.Range("B14").Font.Size = 9; $ws.Range("B14").Font.Color = $GRAYF } catch {}
Write-Output "sheet14 ok"

# ============ sheet15 銘柄マスター ============
$ws = $wb.Worksheets.Item(15)
NoGrid $ws
$ws.Columns.Item("A").ColumnWidth = 2
$ws.Columns.Item("B").ColumnWidth = 12
$ws.Columns.Item("C").ColumnWidth = 26
$ws.Columns.Item("D").ColumnWidth = 8
$ws.Columns.Item("E").ColumnWidth = 7
$ws.Columns.Item("F").ColumnWidth = 10
$ws.Columns.Item("G").ColumnWidth = 11
$ws.Columns.Item("H").ColumnWidth = 12
$ws.Columns.Item("I").ColumnWidth = 20
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
$ws.Range("B3").Value2 = "銘柄"
$ws.Range("C3").Value2 = "名称"
$ws.Range("D3").Value2 = "資産"
$ws.Range("E3").Value2 = "年限"
$ws.Range("F3").Value2 = "利回り%"
$ws.Range("G3").Value2 = "リスク係数"
$ws.Range("H3").Value2 = "DV01($/bp)"
$ws.Range("I3").Value2 = "BBGティッカー"
THead $ws "B3" "I3"
Borders $ws "B3" "I16"
for ($r = 5; $r -le 15; $r += 2) { $ws.Range("B" + $r + ":I" + $r).Interior.Color = $PALE }
$ws.Range("B4:B16").Font.Bold = $true
$ws.Range("B4:B16").Font.Color = $NAVY
$ws.Range("F4:F16").NumberFormat = "0.000"
$ws.Range("G4:G16").NumberFormat = "0.00"
$ws.Range("H4:H16").NumberFormat = "#,##0"
$ws.Range("I4:I16").Font.Name = "Consolas"
$ws.Range("I4:I16").Font.Size = 9
$ws.Range("I4:I16").Font.Color = $GRAYF
try { $ws.Range("B18").Font.Italic = $true; $ws.Range("B18").Font.Size = 9; $ws.Range("B18").Font.Color = $GRAYF } catch {}
Write-Output "sheet15 ok"

# ============ sheet16 リスク ============
$ws = $wb.Worksheets.Item(16)
NoGrid $ws
try {
    $ws.Range("A1:A2").Interior.Color = $NAVY
    $ws.Range("A1").Font.Color = $WHITE
    $ws.Range("A1").Font.Size = 13
    $ws.Range("A1").Font.Bold = $true
    $ws.Rows.Item(1).RowHeight = 26
    $ws.Range("A2").Font.Color = $LIGHT
    $ws.Range("A2").Font.Italic = $true
    $ws.Range("A2").Font.Size = 9
    $ws.Rows.Item(2).RowHeight = 15
} catch {}
$CL = @("A","B","C","D","E","F","G","H","I","J","K","L","M")
foreach ($c in $ws.Range("B1:B80").Cells) {
    $v = $c.Value2
    if (($null -ne $v) -and ($v -is [string]) -and $v.StartsWith([char]0x25A0)) {
        Band $ws ("B" + $c.Row) ("M" + $c.Row)
    }
}
THead $ws "B5" "D5"
THead $ws "B12" "F12"
Borders $ws "B5" "D9"
Borders $ws "B12" "F20"
Borders $ws "B25" "M33"
Borders $ws "B36" "D43"
Borders $ws "B46" "F54"
$ws.Range("C6:C9").NumberFormat = "0.00"
$ws.Range("C13:C20").Font.Bold = $true
$ws.Range("C13:C20").HorizontalAlignment = -4108
$ws.Range("F13:F20").NumberFormat = "#,##0;-#,##0"
RedNeg $ws "F13" "F20"
try { $ws.Range("D6:D9").WrapText = $true; $ws.Range("D6:D9").Font.Size = 9; $ws.Range("D6:D9").Font.Color = $GRAYF } catch {}
try { $ws.Columns.Item("O").Hidden = $true } catch {}
$ws.Columns.Item("B").ColumnWidth = 32
$ws.Columns.Item("C").ColumnWidth = 12
$ws.Columns.Item("D").ColumnWidth = 58
$ws.Columns.Item("E").ColumnWidth = 24
$ws.Columns.Item("F").ColumnWidth = 16
Freeze $ws 2 0
Write-Output "sheet16 ok"

# ============ finalize ============
foreach ($sh in $ws2.Shapes) { if ($sh.Type -eq 1) { try { $sh.ZOrder(0) } catch {} } }
$excel.ScreenUpdating = $true
try { $null = $excel.Goto($ws2.Range("A1"), $true) } catch {}
try { $excel.CalculateFull() } catch {}
$wb.Save()
Write-Output "SAVED"
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "DONE"
