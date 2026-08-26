$ErrorActionPreference = "Stop"
$src = "C:\Users\dai86\Downloads\BondRV_Model.xlsm"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($src)
    foreach ($s2 in $wb.Sheets) { if ($s2.Name -eq "シグナル検証") { $s2.Delete() } }
    $after = $wb.Sheets.Item("エントリーシグナル")
    $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $after)
    $ws.Name = "シグナル検証"
    $ws.Range("B1").Value = "シグナル検証 — ±2σエントリールールのバックテスト"
    $ws.Range("B1").Font.Bold = $true
    $ws.Range("B1").Font.Size = 14
    $ws.Range("B2").Value = "各戦略スプレッドの252日Zスコアが±閾値を超えた日にエントリー(平均回帰)、n営業日後クローズ。分析データ参照のため更新後自動再計算。"
    $ws.Range("B2").Font.Size = 9
    $ws.Range("B4").Value = "■ パラメータ(空色=編集可)"
    $ws.Range("B4").Font.Bold = $true
    $ws.Range("B5").Value = "保有期間(営業日)"
    $ws.Range("C5").Value = 10
    $ws.Range("B6").Value = "エントリー閾値(σ)"
    $ws.Range("C6").Value = 2
    $ws.Range("B7").Value = "Zスコア・ルックバック(日)"
    $ws.Range("C7").Value = 252
    $pr = $ws.Range("C5:C7")
    $pr.Interior.Color = 49407
    $pr.Borders.LineStyle = 1
    $ws.Range("B9").Value = "戦略"
    $ws.Range("C9").Value = "シグナル回数"
    $ws.Range("D9").Value = "勝率"
    $ws.Range("E9").Value = "平均損益($)"
    $ws.Range("F9").Value = "合計損益($)"
    $ws.Range("G9").Value = "最大単回損失($)"
    $ws.Range("H9").Value = "平均保有(bp)"
    $ws.Range("I9").Value = "判定"
    $hd = $ws.Range("B9:I9")
    $hd.Font.Bold = $true
    $hd.Interior.Color = 15189684
    $hd.Borders.LineStyle = 1
    $ws.Range("B10").Value = "S1 UST 2s10s30s BF"
    $ws.Range("C10").Formula = '=COUNT(AV$2:AV$1307)'
    $ws.Range("D10").Formula = '=IFERROR(COUNTIF(AV$2:AV$1307,">0")/COUNT(AV$2:AV$1307),"")'
    $ws.Range("E10").Formula = '=IFERROR(AVERAGE(AV$2:AV$1307),"")'
    $ws.Range("F10").Formula = '=SUM(AV$2:AV$1307)'
    $ws.Range("G10").Formula = '=IFERROR(MIN(AV$2:AV$1307),"")'
    $ws.Range("H10").Formula = '=IFERROR(AVERAGE(AV$2:AV$1307)/80000,"")'
    $ws.Range("I10").Formula = '=IF(COUNT(AV$2:AV$1307)=0,"データ不足",IF(AND(F10>0,D10>=0.5),"採用候補",IF(F10>0,"要観察","見送り")))'
    $ws.Range("B11").Value = "S2 JGB 2s10s30s BF"
    $ws.Range("C11").Formula = '=COUNT(BB$2:BB$1307)'
    $ws.Range("D11").Formula = '=IFERROR(COUNTIF(BB$2:BB$1307,">0")/COUNT(BB$2:BB$1307),"")'
    $ws.Range("E11").Formula = '=IFERROR(AVERAGE(BB$2:BB$1307),"")'
    $ws.Range("F11").Formula = '=SUM(BB$2:BB$1307)'
    $ws.Range("G11").Formula = '=IFERROR(MIN(BB$2:BB$1307),"")'
    $ws.Range("H11").Formula = '=IFERROR(AVERAGE(BB$2:BB$1307)/60000,"")'
    $ws.Range("I11").Formula = '=IF(COUNT(BB$2:BB$1307)=0,"データ不足",IF(AND(F11>0,D11>=0.5),"採用候補",IF(F11>0,"要観察","見送り")))'
    $ws.Range("B12").Value = "S3 UST-JGB 10年ベーシス"
    $ws.Range("C12").Formula = '=COUNT(BH$2:BH$1307)'
    $ws.Range("D12").Formula = '=IFERROR(COUNTIF(BH$2:BH$1307,">0")/COUNT(BH$2:BH$1307),"")'
    $ws.Range("E12").Formula = '=IFERROR(AVERAGE(BH$2:BH$1307),"")'
    $ws.Range("F12").Formula = '=SUM(BH$2:BH$1307)'
    $ws.Range("G12").Formula = '=IFERROR(MIN(BH$2:BH$1307),"")'
    $ws.Range("H12").Formula = '=IFERROR(AVERAGE(BH$2:BH$1307)/50000,"")'
    $ws.Range("I12").Formula = '=IF(COUNT(BH$2:BH$1307)=0,"データ不足",IF(AND(F12>0,D12>=0.5),"採用候補",IF(F12>0,"要観察","見送り")))'
    $ws.Range("B13").Value = "S4 UST 5s10sスロープ"
    $ws.Range("C13").Formula = '=COUNT(BN$2:BN$1307)'
    $ws.Range("D13").Formula = '=IFERROR(COUNTIF(BN$2:BN$1307,">0")/COUNT(BN$2:BN$1307),"")'
    $ws.Range("E13").Formula = '=IFERROR(AVERAGE(BN$2:BN$1307),"")'
    $ws.Range("F13").Formula = '=SUM(BN$2:BN$1307)'
    $ws.Range("G13").Formula = '=IFERROR(MIN(BN$2:BN$1307),"")'
    $ws.Range("H13").Formula = '=IFERROR(AVERAGE(BN$2:BN$1307)/40000,"")'
    $ws.Range("I13").Formula = '=IF(COUNT(BN$2:BN$1307)=0,"データ不足",IF(AND(F13>0,D13>=0.5),"採用候補",IF(F13>0,"要観察","見送り")))'
    $ws.Range("B14").Value = "S7 Bund-UST 10年ベーシス"
    $ws.Range("C14").Formula = '=COUNT(BT$2:BT$1307)'
    $ws.Range("D14").Formula = '=IFERROR(COUNTIF(BT$2:BT$1307,">0")/COUNT(BT$2:BT$1307),"")'
    $ws.Range("E14").Formula = '=IFERROR(AVERAGE(BT$2:BT$1307),"")'
    $ws.Range("F14").Formula = '=SUM(BT$2:BT$1307)'
    $ws.Range("G14").Formula = '=IFERROR(MIN(BT$2:BT$1307),"")'
    $ws.Range("H14").Formula = '=IFERROR(AVERAGE(BT$2:BT$1307)/40000,"")'
    $ws.Range("I14").Formula = '=IF(COUNT(BT$2:BT$1307)=0,"データ不足",IF(AND(F14>0,D14>=0.5),"採用候補",IF(F14>0,"要観察","見送り")))'
    $ws.Range("B15").Value = "S8 JGB-Bund 10年ベーシス"
    $ws.Range("C15").Formula = '=COUNT(BZ$2:BZ$1307)'
    $ws.Range("D15").Formula = '=IFERROR(COUNTIF(BZ$2:BZ$1307,">0")/COUNT(BZ$2:BZ$1307),"")'
    $ws.Range("E15").Formula = '=IFERROR(AVERAGE(BZ$2:BZ$1307),"")'
    $ws.Range("F15").Formula = '=SUM(BZ$2:BZ$1307)'
    $ws.Range("G15").Formula = '=IFERROR(MIN(BZ$2:BZ$1307),"")'
    $ws.Range("H15").Formula = '=IFERROR(AVERAGE(BZ$2:BZ$1307)/40000,"")'
    $ws.Range("I15").Formula = '=IF(COUNT(BZ$2:BZ$1307)=0,"データ不足",IF(AND(F15>0,D15>=0.5),"採用候補",IF(F15>0,"要観察","見送り")))'
    $ws.Range("B16").Value = "全戦略合計"
    $ws.Range("C16").Formula = '=SUM(C10:C15)'
    $ws.Range("D16").Formula = '=IFERROR((COUNTIF(AV$2:AV$1307,">0")+COUNTIF(BB$2:BB$1307,">0")+COUNTIF(BH$2:BH$1307,">0")+COUNTIF(BN$2:BN$1307,">0")+COUNTIF(BT$2:BT$1307,">0")+COUNTIF(BZ$2:BZ$1307,">0"))/C16,"")'
    $ws.Range("E16").Formula = '=IFERROR(F16/C16,"")'
    $ws.Range("F16").Formula = '=SUM(F10:F15)'
    $ws.Range("G16").Formula = '=MIN(AV$2:AV$1307,BB$2:BB$1307,BH$2:BH$1307,BN$2:BN$1307,BT$2:BT$1307,BZ$2:BZ$1307)'
    $ws.Range("I16").Formula = '=IF(AND(F16>0,D16>=0.5),"採用候補",IF(F16>0,"要観察","見送り"))'
    $sm = $ws.Range("B10:I16")
    $sm.Borders.LineStyle = 1
    $ws.Range("B16:I16").Font.Bold = $true
    $ws.Range("B18").Value = "■ ポートフォリオ(全戦略合算)のリスク指標"
    $ws.Range("B18").Font.Bold = $true
    $ws.Range("B19").Value = "最大ドローダウン($)"
    $ws.Range("C19").Formula = '=MIN(CG$2:CG$1307)'
    $ws.Range("B20").Value = "最大DD(仮NAV $100mm比 %)"
    $ws.Range("C20").Formula = '=C19/100000000'
    $ws.Range("B21").Value = "累積トレード損益($)"
    $ws.Range("C21").Formula = '=F16'
    $ws.Range("B22").Value = "総トレード回数"
    $ws.Range("C22").Formula = '=C16'
    $ws.Range("C19").NumberFormat = "#,##0"
    $ws.Range("C20").NumberFormat = "0.00%"
    $ws.Range("C21").NumberFormat = "#,##0"
    $ws.Range("B24").Value = "【方法】各戦略スプレッドの252日Zスコアが±閾値σを超えた日にエントリー(+σ=ショート、−σ=ロングの平均回帰)、n営業日後クローズ。"
    $ws.Range("B24").Font.Size = 9
    $ws.Range("B25").Value = "トレード損益 = DV01 × SIGN(Zエントリー) × (エントリー時スプレッド − クローズ時スプレッド)。取引コスト・ポジション上限は未反映。"
    $ws.Range("B25").Font.Size = 9
    $ws.Range("B26").Value = "累積カーブは各トレード損益をエントリー日に合算(トレード重複許容)した参考値。S4はスワップ履歴無ため5s10sスロープで代用。"
    $ws.Range("B26").Font.Size = 9
    $ws.Range("B27").Value = "右側の非表示列に日次Zスコア・シグナル・トレード損益・累積カーブを保持。パラメータ(C5:C7)を変えると全て再計算。"
    $ws.Range("B27").Font.Size = 9
    $ws.Range("AW1").Value = 0
    $ws.Range("AX1").Value = 0
    $ws.Range("AY1").Value = 0
    $ws.Range("BC1").Value = 0
    $ws.Range("BD1").Value = 0
    $ws.Range("BE1").Value = 0
    $ws.Range("BI1").Value = 0
    $ws.Range("BJ1").Value = 0
    $ws.Range("BK1").Value = 0
    $ws.Range("BO1").Value = 0
    $ws.Range("BP1").Value = 0
    $ws.Range("BQ1").Value = 0
    $ws.Range("BU1").Value = 0
    $ws.Range("BV1").Value = 0
    $ws.Range("BW1").Value = 0
    $ws.Range("CA1").Value = 0
    $ws.Range("CB1").Value = 0
    $ws.Range("CC1").Value = 0
    $ws.Range("CE1").Value = 0
    $ws.Range("CF1").Value = 0
    $ws.Range("CG1").Value = 0
    $ws.Range("AT2:AT1307").Formula = '=IF(ROW()<($C$7+1),"",IFERROR((分析データ!B2-AVERAGE(OFFSET(分析データ!B$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)))/STDEV(OFFSET(分析データ!B$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)),0))'
    $ws.Range("AU2:AU1307").Formula = '=IF(ISNUMBER(AT2),IF(ABS(AT2)>=$C$6,1,0),0)'
    $ws.Range("AV2:AV1307").Formula = '=IF(AU2=1,IF(INDEX(分析データ!$A:$A,ROW()+$C$5)<>"",80000*SIGN(AT2)*(分析データ!B2-INDEX(分析データ!B:B,ROW()+$C$5)),""),"")'
    $ws.Range("AW2:AW1307").Formula = '=AW1+IF(AND(AU2=1,ISNUMBER(AV2)),AV2,0)'
    $ws.Range("AX2:AX1307").Formula = '=MAX(AX1,AW2)'
    $ws.Range("AY2:AY1307").Formula = '=AW2-AX2'
    $ws.Range("AZ2:AZ1307").Formula = '=IF(ROW()<($C$7+1),"",IFERROR((分析データ!C2-AVERAGE(OFFSET(分析データ!C$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)))/STDEV(OFFSET(分析データ!C$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)),0))'
    $ws.Range("BA2:BA1307").Formula = '=IF(ISNUMBER(AZ2),IF(ABS(AZ2)>=$C$6,1,0),0)'
    $ws.Range("BB2:BB1307").Formula = '=IF(BA2=1,IF(INDEX(分析データ!$A:$A,ROW()+$C$5)<>"",60000*SIGN(AZ2)*(分析データ!C2-INDEX(分析データ!C:C,ROW()+$C$5)),""),"")'
    $ws.Range("BC2:BC1307").Formula = '=BC1+IF(AND(BA2=1,ISNUMBER(BB2)),BB2,0)'
    $ws.Range("BD2:BD1307").Formula = '=MAX(BD1,BC2)'
    $ws.Range("BE2:BE1307").Formula = '=BC2-BD2'
    $ws.Range("BF2:BF1307").Formula = '=IF(ROW()<($C$7+1),"",IFERROR((分析データ!D2-AVERAGE(OFFSET(分析データ!D$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)))/STDEV(OFFSET(分析データ!D$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)),0))'
    $ws.Range("BG2:BG1307").Formula = '=IF(ISNUMBER(BF2),IF(ABS(BF2)>=$C$6,1,0),0)'
    $ws.Range("BH2:BH1307").Formula = '=IF(BG2=1,IF(INDEX(分析データ!$A:$A,ROW()+$C$5)<>"",50000*SIGN(BF2)*(分析データ!D2-INDEX(分析データ!D:D,ROW()+$C$5)),""),"")'
    $ws.Range("BI2:BI1307").Formula = '=BI1+IF(AND(BG2=1,ISNUMBER(BH2)),BH2,0)'
    $ws.Range("BJ2:BJ1307").Formula = '=MAX(BJ1,BI2)'
    $ws.Range("BK2:BK1307").Formula = '=BI2-BJ2'
    $ws.Range("BL2:BL1307").Formula = '=IF(ROW()<($C$7+1),"",IFERROR((分析データ!E2-AVERAGE(OFFSET(分析データ!E$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)))/STDEV(OFFSET(分析データ!E$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)),0))'
    $ws.Range("BM2:BM1307").Formula = '=IF(ISNUMBER(BL2),IF(ABS(BL2)>=$C$6,1,0),0)'
    $ws.Range("BN2:BN1307").Formula = '=IF(BM2=1,IF(INDEX(分析データ!$A:$A,ROW()+$C$5)<>"",40000*SIGN(BL2)*(分析データ!E2-INDEX(分析データ!E:E,ROW()+$C$5)),""),"")'
    $ws.Range("BO2:BO1307").Formula = '=BO1+IF(AND(BM2=1,ISNUMBER(BN2)),BN2,0)'
    $ws.Range("BP2:BP1307").Formula = '=MAX(BP1,BO2)'
    $ws.Range("BQ2:BQ1307").Formula = '=BO2-BP2'
    $ws.Range("BR2:BR1307").Formula = '=IF(ROW()<($C$7+1),"",IFERROR((分析データ!F2-AVERAGE(OFFSET(分析データ!F$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)))/STDEV(OFFSET(分析データ!F$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)),0))'
    $ws.Range("BS2:BS1307").Formula = '=IF(ISNUMBER(BR2),IF(ABS(BR2)>=$C$6,1,0),0)'
    $ws.Range("BT2:BT1307").Formula = '=IF(BS2=1,IF(INDEX(分析データ!$A:$A,ROW()+$C$5)<>"",40000*SIGN(BR2)*(分析データ!F2-INDEX(分析データ!F:F,ROW()+$C$5)),""),"")'
    $ws.Range("BU2:BU1307").Formula = '=BU1+IF(AND(BS2=1,ISNUMBER(BT2)),BT2,0)'
    $ws.Range("BV2:BV1307").Formula = '=MAX(BV1,BU2)'
    $ws.Range("BW2:BW1307").Formula = '=BU2-BV2'
    $ws.Range("BX2:BX1307").Formula = '=IF(ROW()<($C$7+1),"",IFERROR((分析データ!G2-AVERAGE(OFFSET(分析データ!G$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)))/STDEV(OFFSET(分析データ!G$1,MAX(1,ROW()-$C$7),0,MIN($C$7,ROW()-1),1)),0))'
    $ws.Range("BY2:BY1307").Formula = '=IF(ISNUMBER(BX2),IF(ABS(BX2)>=$C$6,1,0),0)'
    $ws.Range("BZ2:BZ1307").Formula = '=IF(BY2=1,IF(INDEX(分析データ!$A:$A,ROW()+$C$5)<>"",40000*SIGN(BX2)*(分析データ!G2-INDEX(分析データ!G:G,ROW()+$C$5)),""),"")'
    $ws.Range("CA2:CA1307").Formula = '=CA1+IF(AND(BY2=1,ISNUMBER(BZ2)),BZ2,0)'
    $ws.Range("CB2:CB1307").Formula = '=MAX(CB1,CA2)'
    $ws.Range("CC2:CC1307").Formula = '=CA2-CB2'
    $ws.Range("CD2:CD1307").Formula = '=N(AV2)+N(BB2)+N(BH2)+N(BN2)+N(BT2)+N(BZ2)'
    $ws.Range("CE2:CE1307").Formula = '=CE1+CD2'
    $ws.Range("CF2:CF1307").Formula = '=MAX(CF1,CE2)'
    $ws.Range("CG2:CG1307").Formula = '=CE2-CF2'
    $ws.Range("CH2:CH1307").Formula = '=分析データ!A2'
    $ws.Range("CH2:CH1307").NumberFormat = "yyyy-mm-dd"
    $ws.Range("AT:CH").EntireColumn.Hidden = $true
    $ws.Range("C10:C16").NumberFormat = "0"
    $ws.Range("D10:D16").NumberFormat = "0.0%"
    $ws.Range("E10:E16").NumberFormat = "#,##0"
    $ws.Range("F10:F16").NumberFormat = "#,##0"
    $ws.Range("G10:G16").NumberFormat = "#,##0"
    $ws.Range("H10:H15").NumberFormat = "0.0"
    $ws.Columns("B").ColumnWidth = 26
    $ws.Columns("C:I").ColumnWidth = 14
    try {
        $anchor = $ws.Range("K9")
        $co = $ws.ChartObjects().Add($anchor.Left, $anchor.Top, 540, 300)
        $ch = $co.Chart
        $ch.SetSourceData($ws.Range('CE$2:CE$1307'))
        $ch.ChartType = 4
        try { $ch.SeriesCollection(1).XValues = $ws.Range('CH$2:CH$1307') } catch {}
        try { $ch.SeriesCollection(1).Name = '累積トレード損益($)' } catch {}
        $ch.HasTitle = $true
        $ch.ChartTitle.Text = 'シグナル戦略 累積損益(全戦略合算・参考)'
        Write-Output "chart OK"
    } catch {
        Write-Output ("chart failed (non-fatal): " + $_.Exception.Message)
    }
    $wb.Save()
    Write-Output ("sheet added: " + $ws.Name)
    Write-Output ("S1 count=" + $ws.Range("C10").Text + " win=" + $ws.Range("D10").Text + " sum=" + $ws.Range("F10").Text)
    Write-Output ("portfolio sum=" + $ws.Range("F16").Text + " maxDD=" + $ws.Range("C19").Text)
    $wb.Close($true)
} catch {
    Write-Output ("FAILED: " + $_.Exception.Message)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}