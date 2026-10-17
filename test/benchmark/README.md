# ベンチマーク用のテストデータ

起動・表示・スクロールのベンチマーク（`scripts/benchmark.sh`）で開く、公開してよい重いファイルです。社名・人名・製品名・数値は、pptx を除いてすべて架空です（実在の会社や製品のデータとは関係がありません）。

| ファイル | 中身 | 作り方 |
|---|---|---|
| `earnings-presentation.pptx` | アイビスの決算説明会資料の英語版（FY2026/12 2Q、31 スライド、画像と埋め込みフォントが多く 35,902,214 バイト）。2026-10-07 に、日本語版から英語版に替えた | 機密情報ではない資料から、人や端末を特定できるメタ情報を取り除いた（`scripts/benchmark-data/scrub_ooxml.py --drop-chart-data`）：作成者・更新者・印刷日時、コメントの作者の名前・ID（`ppt/authors.xml` は 1 人の匿名の作者に）、アドインの登録、グラフの元データのブック（埋め込みのブックの作者・プリンターの設定も）、EMF の画像（Illustrator の EPS 注釈）の中の作成者名とフォルダーのパス（同じ長さの `x` に置き換え、描画は変えない）。zip の中身の全部を、名前（漢字・ローマ字）・メール・`/Users/`・userId で探して 0 件であることを確かめた。スライドの文面にある役員の名前と、スライド 31 の公開の URL（IR のページと開示の PDF）は資料の内容なので残してある。Microsoft Forms の URL（アンケートの QR コードの画像、「IR Feedback Survey」の文字の枠、ハイパーリンク）は 2026-10-07 に削った（スライド 31 だけが変わる）。ibisOffice の描画は元の 31 枚と同じ画素 |
| `coffee-chain-sales-report.xlsx` | 架空のコーヒーチェーン「Nimbus Coffee」の日次の売上。商品の分類 4 つ × 地域 8 つ × 販売経路 2 つ（店舗・オンライン）の売上・注文数・客単価のシート 17 枚（2020-01-01〜2026-07-31、2,404 行）と、生データのシート 4 枚（日付 × 販売経路、4,808 行）、日次のシートを参照する重いグラフ 29 個（折れ線と、最大 16 系列の 100% 積み上げ面、各 2,404 点）の「Overview」シート。日次のシートは、生データから列全体を参照する SUMIFS で集計する数式と、シートをまたいで割る客単価の数式（計 58.9 万個、計算済みの値つき）。計算エンジンの版（calcId）は 191028 | `scripts/benchmark-data/gen_xlsx.py` |
| `securities-registration-statement.docx` | 架空のコーヒーチェーン「Nimbus Coffee, Inc.」の英文の有価証券届出書（Securities Registration Statement）。証券情報、事業の概要・リスク・経営者による分析、店舗と焙煎所、役員、ストックオプション、財務諸表（貸借対照表・損益計算書・キャッシュ・フロー計算書と注記を 7 期分、単位は千ドル）など。本文 Times New Roman 10.5pt・見出し Arial、A4 で 140 ページ（ibisOffice のレイアウト）、罫線の表が 159 個 | `scripts/benchmark-data/gen_docx.py` |
| `vba-module-demo.xlsm` | VBA プラグイン（`crates/ibis-vba`）の標準モジュールだけのテスト。`Module1.BuildReport`：`Cells` のループで A1:A5 に数値を書き、`B1` に `=SUM(A1:A5)`、`B2` に文字。実行すると A1..A5 = 10..50、B1 = 150 になる | `pyopenvba` の `ExcelFile.create_new` と `add_module`（下の戻し方） |
| `vba-class-events-demo.xlsm` | 同じく VBA プラグインの、クラスとシートイベント入りのテスト。`Module1.BuildReport` は `MyClass` を `New` してフィールドを読む（B2 に "class says 5"）、`MyClass` はクラスモジュール（`X As Long`）、`Sheet1` に `Worksheet_Change` のイベント | 同上 |

xlsx と docx は、生成のたびに同じファイルになる（乱数の種を固定している）。作り直すとき：

```sh
python3 scripts/benchmark-data/gen_xlsx.py test/benchmark/coffee-chain-sales-report.xlsx
python3 scripts/benchmark-data/gen_docx.py test/benchmark/securities-registration-statement.docx
```

VBA のテストファイル（`vba-*.xlsm`）の戻し方（`pyopenvba` を pip で入れて）：

```python
from pyopenvba import ExcelFile, VBAModuleKind
with ExcelFile.create_new("test/benchmark/vba-module-demo.xlsm") as wb:
    wb.set_module("Module1", SOURCES["Module1"])
    wb.save()
with ExcelFile.create_new("test/benchmark/vba-class-events-demo.xlsm") as wb:
    wb.set_module("Module1", SOURCES["Module1"])
    p = wb.vba_project()
    p.add_module("MyClass", SOURCES["MyClass"], kind=VBAModuleKind.other)
    wb.set_module("Sheet1", SOURCES["Sheet1"])
    wb.save()
```
