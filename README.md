# 債券レラティブバリュー（RV）モデル

海外大手ヘッジファンドのスタイルを参考にした、**債券 RV ポートフォリオ管理モデル**です。
成果物は Excel マクロ有効ブック `BondRV_Model.xlsm`（16 シート・8 戦略）。
本リポジトリには、そのモデルを**ゼロから再生成・日次更新するための一式**を収めています。

> ⚠️ 本モデルは教育・研究用の自作サンプルです。投資助言ではありません。

---

## 成果物（xlsm）の構成

| シート | 内容 |
|---|---|
| はじめに | モデルの概要・使い方 |
| ダッシュボード | 各戦略の最新 Z スコア・シグナル一覧 |
| 戦略解説 | 8 戦略のロジック説明 |
| エントリーシグナル | ±2σ バンドとエントリー判定 |
| シグナル検証 | エントリールールのバックテスト（勝率・平均リターン・最大 DD） |
| トレード台帳 | 建玉・クローズ管理 |
| ポートフォリオ | ポジション集約・NAV |
| 分析データ | 各戦略の日次スプレッドと P&L（Z スコア計算の元データ） |
| 履歴データ | 唯一のデータソース（19 列 A〜S、R=BUND_FUT / S=JGB_FUT） |
| イールドカーブ | 各国イールドカーブ |
| ASW / ASWデータ | アセットスワップスプレッド |
| ファンディング | ファンディングコスト |
| 損益 | 損益集計 |
| 銘柄マスター | 銘柄・係数マスタ |
| リスク | VaR・ストレステスト・リミット（Phase 2） |

### 8 戦略
- **S1** 米国債 10y−2y イールドカーブ
- **S2** 米国債 30y−10y イールドカーブ
- **S3** ドイツ Bund − 米国債 10y スプレッド
- **S4** スワップスプレッド（日米）
- **S7** Bund 先物 ベーシス／インプライド
- **S8** JGB 先物 ベーシス／インプライド
- ほか 2 戦略（詳細は「戦略解説」シート参照）

---

## データソース（すべて公開データ）

| 区分 | ソース |
|---|---|
| 米国債イールド | U.S. Treasury |
| 短期金利（SOFR/EFFR/BGCR/TGCR） | NY Fed |
| 日本国債イールドカーブ | 財務省（MOF、cp932） |
| 先物・ETF・為替 | Yahoo Finance（2510.T / Bund ETF / EURUSD / USDJPY ほか） |
| Bund 先物インプライド | EUREX 系ページより収集（`data/BUND_implied_yield.csv`） |
| ASW 参照 | Pensford（公開レート） |

生データは `data/` に CSV で保存。VBA マクロがこれらを取得・更新します。

---

## 使い方

### 1) モデルを再生成する（フルビルド）
```bash
python build_v3.py
```
- `BondRV_Model.xlsx` を生成 → その後 VBA を注入して `.xlsm` 化します。
- ビルド時のデータ行数・容量定数は `build_v3.py` 冒頭を参照。

### 2) VBA を注入する（.xlsm 化）
```powershell
powershell -ExecutionPolicy Bypass -File inject_vba.ps1
```
- `DataUpdater.bas`（**cp932 / Shift-JIS 必須**）をブックに組み込みます。

### 3) 日次データを更新する
ブック内の更新ボタン（VBA `DataUpdater`）を押すか、PowerShell で実行：
```powershell
powershell -ExecutionPolicy Bypass -File run_update.ps1     # 更新マクロ実行
powershell -ExecutionPolicy Bypass -File save_update.ps1    # 上書き保存
```

### 4) 品質チェック（#エラー走査）
```powershell
powershell -ExecutionPolicy Bypass -File qa_scan.ps1
```

### 5) エントリーシグナルのバックテスト
```bash
python gen_sig_backtest.py   # → sig_backtest.ps1 を生成
powershell -ExecutionPolicy Bypass -File sig_backtest.ps1
```

---

## ファイル構成

```
build_v3.py            # 本体ビルダー（16 シートを生成）
master_hist.py         # 履歴データ収集・整形
parse_data.py          # データパース補助
dv01.py                # DV01 計算
DataUpdater.bas        # 日次更新 VBA（cp932）
inject_vba.ps1         # VBA 注入（.xlsm 化）
run_update.ps1         # 更新マクロ実行
save_update.ps1        # 保存
export_bas.ps1         # VBA 書き出し
qa_scan.ps1            # #エラー走査
gen_sig_backtest.py    # シグナル検証シート生成スクリプト
sig_backtest.ps1       # シグナル検証シート適用（COM）
data/                  # 生データ CSV 一式
BondRV_Model.xlsm  # 成果物（VBA 込み）
```

---

## 技術メモ（落とし穴）

- **`.bas` は必ず cp932（Shift-JIS）** で保存。UTF-8/UTF-16 だと VBA 取込が壊れます。
- COM 操作の前に **`EXCEL.EXE` を必ず kill** すること。
- 数式参照をまたぐ範囲での `Rows.Delete` は禁止。
- PowerShell で日本語を書く場合は **utf-8-sig（BOM 付き）** の `.ps1` を Python で生成するのが安全。
- 文字列内の `$A` は PowerShell でドライブ修飾変数と解釈されるため、Range 指定は**シングルクォート**で。

---

## ライセンス / 免責

本リポジトリは個人の利用・学習目的の自作サンプルです。市場データは各公開ソースの利用条件に従ってください。本モデルを用いた取引結果について、作成者は一切責任を負いません。
