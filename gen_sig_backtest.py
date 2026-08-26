# -*- coding: utf-8 -*-
# シグナルバックテストシート「シグナル検証」を既存xlsmにCOMで追加するPSを生成
import io
from openpyxl.utils import get_column_letter as gcl

ANA = "\u5206\u6790\u30c7\u30fc\u30bf"  # 分析データ
SHEET = "\u30b7\u30b0\u30ca\u30eb\u691c\u8a3c"  # シグナル検証
ENTRY = "\u30a8\u30f3\u30c8\u30ea\u30fc\u30b7\u30b0\u30ca\u30eb"  # エントリーシグナル

CAP = 1307  # 容量最終行
START = 46
strats = [
    ("S1 UST 2s10s30s BF", "B", 80000),
    ("S2 JGB 2s10s30s BF", "C", 60000),
    ("S3 UST-JGB 10\u5e74\u30d9\u30fc\u30b7\u30b9", "D", 50000),
    ("S4 UST 5s10s\u30b9\u30ed\u30fc\u30d7", "E", 40000),
    ("S7 Bund-UST 10\u5e74\u30d9\u30fc\u30b7\u30b9", "F", 40000),
    ("S8 JGB-Bund 10\u5e74\u30d9\u30fc\u30b7\u30b9", "G", 40000),
]

cols = {}
trade_cols = []
base = START
for name, L, V in strats:
    cols[L] = dict(z=base, sig=base+1, trade=base+2, cum=base+3, peak=base+4, dd=base+5)
    trade_cols.append(gcl(base+2))
    base += 6
agg = dict(daily=base, cum=base+1, peak=base+2, dd=base+3)
AGC = gcl(agg["cum"]); AGD = gcl(agg["daily"]); AGP = gcl(agg["peak"]); AGDD = gcl(agg["dd"])
DATEH = base + 4  # ローカル日付ヘルパー列(チャートX軸用)
DH = gcl(DATEH)
LASTH = base + 4  # last helper col index

def z_f(L):
    return ('=IF(ROW()<($C$7+1),"",IFERROR(('
            f'{ANA}!{L}2-AVERAGE(OFFSET({ANA}!{L}$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1))'
            f')/STDEV(OFFSET({ANA}!{L}$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)),0))')

def sig_f(zc):
    return f'=IF(ISNUMBER({zc}2),IF(ABS({zc}2)>=$C$6,1,0),0)'

def trade_f(sc, zc, L, V):
    return (f'=IF({sc}2=1,IF(INDEX({ANA}!$A:$A,ROW()+$C$5)<>"",'
            f'{V}*SIGN({zc}2)*({ANA}!{L}2-INDEX({ANA}!{L}:{L},ROW()+$C$5)),""),"")')

def cum_f(cc, sc, tc):
    return f'={cc}1+IF(AND({sc}2=1,ISNUMBER({tc}2)),{tc}2,0)'

def peak_f(pc, cc):
    return f'=MAX({pc}1,{cc}2)'

def dd_f(cc, pc):
    return f'={cc}2-{pc}2'

ps = []
ps.append('$ErrorActionPreference = "Stop"')
ps.append('$src = "C:\\Users\\dai86\\Downloads\\BondRV_Model.xlsm"')
ps.append('$xl = New-Object -ComObject Excel.Application')
ps.append('$xl.Visible = $false')
ps.append('$xl.DisplayAlerts = $false')
ps.append('try {')
ps.append('    $wb = $xl.Workbooks.Open($src)')
# 同名シートが残っていれば削除(再実行ガード)
ps.append(f'    foreach ($s2 in $wb.Sheets) {{ if ($s2.Name -eq "{SHEET}") {{ $s2.Delete() }} }}')
# シート追加(エントリーシグナルの後ろ)
ps.append(f'    $after = $wb.Sheets.Item("{ENTRY}")')
ps.append('    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $after)')
ps.append(f'    $ws.Name = "{SHEET}"')

# タイトル
ps.append(f'    $ws.Range("B1").Value = "\u30b7\u30b0\u30ca\u30eb\u691c\u8a3c \u2014 \u00b12\u03c3\u30a8\u30f3\u30c8\u30ea\u30fc\u30eb\u30fc\u30eb\u306e\u30d0\u30c3\u30af\u30c6\u30b9\u30c8"')
ps.append('    $ws.Range("B1").Font.Bold = $true')
ps.append('    $ws.Range("B1").Font.Size = 14')
ps.append(f'    $ws.Range("B2").Value = "\u5404\u6226\u7565\u30b9\u30d7\u30ec\u30c3\u30c9\u306e252\u65e5Z\u30b9\u30b3\u30a2\u304c\u00b1\u95be\u5024\u3092\u8d85\u3048\u305f\u65e5\u306b\u30a8\u30f3\u30c8\u30ea\u30fc(\u5e73\u5747\u56de\u5e30)\u3001n\u55b6\u696d\u65e5\u5f8c\u30af\u30ed\u30fc\u30ba\u3002\u5206\u6790\u30c7\u30fc\u30bf\u53c2\u7167\u306e\u305f\u3081\u66f4\u65b0\u5f8c\u81ea\u52d5\u518d\u8a08\u7b97\u3002"')
ps.append('    $ws.Range("B2").Font.Size = 9')

# パラメータ
ps.append('    $ws.Range("B4").Value = "\u25a0 \u30d1\u30e9\u30e1\u30fc\u30bf(\u7a7a\u8272=\u7de8\u96c6\u53ef)"')
ps.append('    $ws.Range("B4").Font.Bold = $true')
ps.append('    $ws.Range("B5").Value = "\u4fdd\u6709\u671f\u9593(\u55b6\u696d\u65e5)"')
ps.append('    $ws.Range("C5").Value = 10')
ps.append('    $ws.Range("B6").Value = "\u30a8\u30f3\u30c8\u30ea\u30fc\u95be\u5024(\u03c3)"')
ps.append('    $ws.Range("C6").Value = 2')
ps.append('    $ws.Range("B7").Value = "Z\u30b9\u30b3\u30a2\u30fb\u30eb\u30c3\u30af\u30d0\u30c3\u30af(\u65e5)"')
ps.append('    $ws.Range("C7").Value = 252')
ps.append('    $pr = $ws.Range("C5:C7")')
ps.append('    $pr.Interior.Color = 49407')
ps.append('    $pr.Borders.LineStyle = 1')

# サマリーヘッダ
hdrs = ["\u6226\u7565", "\u30b7\u30b0\u30ca\u30eb\u56de\u6570", "\u52dd\u7387", "\u5e73\u5747\u640d\u76ca($)", "\u5408\u8a08\u640d\u76ca($)", "\u6700\u5927\u5358\u56de\u640d\u5931($)", "\u5e73\u5747\u4fdd\u6709(bp)", "\u5224\u5b9a"]
for i, h in enumerate(hdrs):
    col = gcl(2 + i)
    ps.append(f'    $ws.Range("{col}9").Value = "{h}"')
ps.append('    $hd = $ws.Range("B9:I9")')
ps.append('    $hd.Font.Bold = $true')
ps.append('    $hd.Interior.Color = 15189684')
ps.append('    $hd.Borders.LineStyle = 1')

# サマリー行(各戦略)
for idx, (name, L, V) in enumerate(strats):
    sr = 10 + idx
    tc = gcl(cols[L]["trade"])
    rng = f'{tc}$2:{tc}${CAP}'
    ps.append(f'    $ws.Range("B{sr}").Value = "{name}"')
    ps.append(f'    $ws.Range("C{sr}").Formula = \'=COUNT({rng})\'')
    ps.append(f'    $ws.Range("D{sr}").Formula = \'=IFERROR(COUNTIF({rng},">0")/COUNT({rng}),"")\'')
    ps.append(f'    $ws.Range("E{sr}").Formula = \'=IFERROR(AVERAGE({rng}),"")\'')
    ps.append(f'    $ws.Range("F{sr}").Formula = \'=SUM({rng})\'')
    ps.append(f'    $ws.Range("G{sr}").Formula = \'=IFERROR(MIN({rng}),"")\'')
    ps.append(f'    $ws.Range("H{sr}").Formula = \'=IFERROR(AVERAGE({rng})/{V},"")\'')
    ps.append(f'    $ws.Range("I{sr}").Formula = \'=IF(COUNT({rng})=0,"\u30c7\u30fc\u30bf\u4e0d\u8db3",IF(AND(F{sr}>0,D{sr}>=0.5),"\u63a1\u7528\u5019\u88dc",IF(F{sr}>0,"\u8981\u89b3\u5bdf","\u898b\u9001\u308a")))\'')

# 合計行(16)
num_win = "+".join(f'COUNTIF({t}$2:{t}${CAP},">0")' for t in trade_cols)
min_all = ",".join(f'{t}$2:{t}${CAP}' for t in trade_cols)
ps.append('    $ws.Range("B16").Value = "\u5168\u6226\u7565\u5408\u8a08"')
ps.append('    $ws.Range("C16").Formula = \'=SUM(C10:C15)\'')
ps.append(f'    $ws.Range("D16").Formula = \'=IFERROR(({num_win})/C16,"")\'')
ps.append('    $ws.Range("E16").Formula = \'=IFERROR(F16/C16,"")\'')
ps.append('    $ws.Range("F16").Formula = \'=SUM(F10:F15)\'')
ps.append(f'    $ws.Range("G16").Formula = \'=MIN({min_all})\'')
ps.append('    $ws.Range("I16").Formula = \'=IF(AND(F16>0,D16>=0.5),"\u63a1\u7528\u5019\u88dc",IF(F16>0,"\u8981\u89b3\u5bdf","\u898b\u9001\u308a"))\'')
ps.append('    $sm = $ws.Range("B10:I16")')
ps.append('    $sm.Borders.LineStyle = 1')
ps.append('    $ws.Range("B16:I16").Font.Bold = $true')

# リスク指標ブロック
ps.append('    $ws.Range("B18").Value = "\u25a0 \u30dd\u30fc\u30c8\u30d5\u30a9\u30ea\u30aa(\u5168\u6226\u7565\u5408\u7b97)\u306e\u30ea\u30b9\u30af\u6307\u6a19"')
ps.append('    $ws.Range("B18").Font.Bold = $true')
ps.append('    $ws.Range("B19").Value = "\u6700\u5927\u30c9\u30ed\u30fc\u30c0\u30a6\u30f3($)"')
ps.append(f'    $ws.Range("C19").Formula = \'=MIN({AGDD}$2:{AGDD}${CAP})\'')
ps.append('    $ws.Range("B20").Value = "\u6700\u5927DD(\u4eeeNAV $100mm\u6bd4 %)"')
ps.append('    $ws.Range("C20").Formula = \'=C19/100000000\'')
ps.append('    $ws.Range("B21").Value = "\u7d2f\u7a4d\u30c8\u30ec\u30fc\u30c9\u640d\u76ca($)"')
ps.append('    $ws.Range("C21").Formula = \'=F16\'')
ps.append('    $ws.Range("B22").Value = "\u7dcf\u30c8\u30ec\u30fc\u30c9\u56de\u6570"')
ps.append('    $ws.Range("C22").Formula = \'=C16\'')
ps.append('    $ws.Range("C19").NumberFormat = "#,##0"')
ps.append('    $ws.Range("C20").NumberFormat = "0.00%"')
ps.append('    $ws.Range("C21").NumberFormat = "#,##0"')

# 手法注記
notes = [
    "\u3010\u65b9\u6cd5\u3011\u5404\u6226\u7565\u30b9\u30d7\u30ec\u30c3\u30c9\u306e252\u65e5Z\u30b9\u30b3\u30a2\u304c\u00b1\u95be\u5024\u03c3\u3092\u8d85\u3048\u305f\u65e5\u306b\u30a8\u30f3\u30c8\u30ea\u30fc(+\u03c3=\u30b7\u30e7\u30fc\u30c8\u3001\u2212\u03c3=\u30ed\u30f3\u30b0\u306e\u5e73\u5747\u56de\u5e30)\u3001n\u55b6\u696d\u65e5\u5f8c\u30af\u30ed\u30fc\u30ba\u3002",
    "\u30c8\u30ec\u30fc\u30c9\u640d\u76ca = DV01 \u00d7 SIGN(Z\u30a8\u30f3\u30c8\u30ea\u30fc) \u00d7 (\u30a8\u30f3\u30c8\u30ea\u30fc\u6642\u30b9\u30d7\u30ec\u30c3\u30c9 \u2212 \u30af\u30ed\u30fc\u30ba\u6642\u30b9\u30d7\u30ec\u30c3\u30c9)\u3002\u53d6\u5f15\u30b3\u30b9\u30c8\u30fb\u30dd\u30b8\u30b7\u30e7\u30f3\u4e0a\u9650\u306f\u672a\u53cd\u6620\u3002",
    "\u7d2f\u7a4d\u30ab\u30fc\u30d6\u306f\u5404\u30c8\u30ec\u30fc\u30c9\u640d\u76ca\u3092\u30a8\u30f3\u30c8\u30ea\u30fc\u65e5\u306b\u5408\u7b97(\u30c8\u30ec\u30fc\u30c9\u91cd\u8907\u8a31\u5bb9)\u3057\u305f\u53c2\u8003\u5024\u3002S4\u306f\u30b9\u30ef\u30c3\u30d7\u5c65\u6b74\u7121\u305f\u30815s10s\u30b9\u30ed\u30fc\u30d7\u3067\u4ee3\u7528\u3002",
    "\u53f3\u5074\u306e\u975e\u8868\u793a\u5217\u306b\u65e5\u6b21Z\u30b9\u30b3\u30a2\u30fb\u30b7\u30b0\u30ca\u30eb\u30fb\u30c8\u30ec\u30fc\u30c9\u640d\u76ca\u30fb\u7d2f\u7a4d\u30ab\u30fc\u30d6\u3092\u4fdd\u6301\u3002\u30d1\u30e9\u30e1\u30fc\u30bf(C5:C7)\u3092\u5909\u3048\u308b\u3068\u5168\u3066\u518d\u8a08\u7b97\u3002",
]
for i, nt in enumerate(notes):
    r = 24 + i
    ps.append(f'    $ws.Range("B{r}").Value = "{nt}"')
    ps.append(f'    $ws.Range("B{r}").Font.Size = 9')

# ヘルパー列: row1シード
for name, L, V in strats:
    cc = gcl(cols[L]["cum"]); pc = gcl(cols[L]["peak"]); dc = gcl(cols[L]["dd"])
    ps.append(f'    $ws.Range("{cc}1").Value = 0')
    ps.append(f'    $ws.Range("{pc}1").Value = 0')
    ps.append(f'    $ws.Range("{dc}1").Value = 0')
ps.append(f'    $ws.Range("{AGC}1").Value = 0')
ps.append(f'    $ws.Range("{gcl(agg["peak"])}1").Value = 0')
ps.append(f'    $ws.Range("{AGDD}1").Value = 0')

# ヘルパー列: 各戦略の式を範囲代入(相対参照で自動展開)
for name, L, V in strats:
    zc = gcl(cols[L]["z"]); sc = gcl(cols[L]["sig"]); tc = gcl(cols[L]["trade"])
    cc = gcl(cols[L]["cum"]); pc = gcl(cols[L]["peak"]); dc = gcl(cols[L]["dd"])
    ps.append(f"    $ws.Range(\"{zc}2:{zc}{CAP}\").Formula = '{z_f(L)}'")
    ps.append(f"    $ws.Range(\"{sc}2:{sc}{CAP}\").Formula = '{sig_f(zc)}'")
    ps.append(f"    $ws.Range(\"{tc}2:{tc}{CAP}\").Formula = '{trade_f(sc, zc, L, V)}'")
    ps.append(f"    $ws.Range(\"{cc}2:{cc}{CAP}\").Formula = '{cum_f(cc, sc, tc)}'")
    ps.append(f"    $ws.Range(\"{pc}2:{pc}{CAP}\").Formula = '{peak_f(pc, cc)}'")
    ps.append(f"    $ws.Range(\"{dc}2:{dc}{CAP}\").Formula = '{dd_f(cc, pc)}'")

# 集約列
daily_terms = "+".join(f'N({t}2)' for t in trade_cols)
ps.append(f"    $ws.Range(\"{AGD}2:{AGD}{CAP}\").Formula = '={daily_terms}'")
ps.append(f"    $ws.Range(\"{AGC}2:{AGC}{CAP}\").Formula = '={AGC}1+{AGD}2'")
ps.append(f"    $ws.Range(\"{AGP}2:{AGP}{CAP}\").Formula = '=MAX({AGP}1,{AGC}2)'")
ps.append(f"    $ws.Range(\"{AGDD}2:{AGDD}{CAP}\").Formula = '={AGC}2-{AGP}2'")
ps.append(f"    $ws.Range(\"{DH}2:{DH}{CAP}\").Formula = '={ANA}!A2'")
ps.append(f"    $ws.Range(\"{DH}2:{DH}{CAP}\").NumberFormat = \"yyyy-mm-dd\"")

# ヘルパー列を非表示
ps.append(f'    $ws.Range("{gcl(START)}:{gcl(LASTH)}").EntireColumn.Hidden = $true')

# 数式書式(サマリー)
ps.append('    $ws.Range("C10:C16").NumberFormat = "0"')
ps.append('    $ws.Range("D10:D16").NumberFormat = "0.0%"')
ps.append('    $ws.Range("E10:E16").NumberFormat = "#,##0"')
ps.append('    $ws.Range("F10:F16").NumberFormat = "#,##0"')
ps.append('    $ws.Range("G10:G16").NumberFormat = "#,##0"')
ps.append('    $ws.Range("H10:H15").NumberFormat = "0.0"')
ps.append('    $ws.Columns("B").ColumnWidth = 26')
ps.append('    $ws.Columns("C:I").ColumnWidth = 14')

# チャート(累積カーブ) K9アンカー — 失敗しても保存は進める
ps.append('    try {')
ps.append('        $anchor = $ws.Range("K9")')
ps.append('        $co = $ws.ChartObjects().Add($anchor.Left, $anchor.Top, 540, 300)')
ps.append('        $ch = $co.Chart')
ps.append(f"        $ch.SetSourceData($ws.Range('{AGC}$2:{AGC}${CAP}'))")
ps.append('        $ch.ChartType = 4')
ps.append(f"        try {{ $ch.SeriesCollection(1).XValues = $ws.Range('{DH}$2:{DH}${CAP}') }} catch {{}}")
ps.append("        try { $ch.SeriesCollection(1).Name = '\u7d2f\u7a4d\u30c8\u30ec\u30fc\u30c9\u640d\u76ca($)' } catch {}")
ps.append('        $ch.HasTitle = $true')
ps.append("        $ch.ChartTitle.Text = '\u30b7\u30b0\u30ca\u30eb\u6226\u7565 \u7d2f\u7a4d\u640d\u76ca(\u5168\u6226\u7565\u5408\u7b97\u30fb\u53c2\u8003)'")
ps.append('        Write-Output "chart OK"')
ps.append('    } catch {')
ps.append('        Write-Output ("chart failed (non-fatal): " + $_.Exception.Message)')
ps.append('    }')

ps.append('    $wb.Save()')
ps.append('    Write-Output ("sheet added: " + $ws.Name)')
ps.append('    Write-Output ("S1 count=" + $ws.Range("C10").Text + " win=" + $ws.Range("D10").Text + " sum=" + $ws.Range("F10").Text)')
ps.append('    Write-Output ("portfolio sum=" + $ws.Range("F16").Text + " maxDD=" + $ws.Range("C19").Text)')
ps.append('    $wb.Close($true)')
ps.append('} catch {')
ps.append('    Write-Output ("FAILED: " + $_.Exception.Message)')
ps.append('} finally {')
ps.append('    $xl.Quit()')
ps.append('    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null')
ps.append('}')

io.open('sig_backtest.ps1', 'w', encoding='utf-8-sig').write("\n".join(ps))
print("sig_backtest.ps1 written, lines:", len(ps))
print("agg cum col:", AGC, "daily:", AGD, "dd:", AGDD, "last helper:", gcl(LASTH))
