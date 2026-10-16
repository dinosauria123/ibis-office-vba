# ibis-vba — ibisOffice VBA プラグイン

`ibisOffice` に VBA マクロ機能を提供する Rust クレートです。マクロ有効ファイル
(`.xlsm` / `.xlsb` / `.docm` / `.pptm`) の VBA プロジェクトを読み取り、一覧表示と、
**データ処理系マクロのヘッドレス実行**を行えます。

上位リポジトリの既定仕様は「マクロは実行しない ( kept as they are )」です。
本クレートは、ユーザーが明示的に要求したときだけ(`--run-macro` コマンドライン
オプション)実行を行うプラグインとして設計されています。

## 依存ライブラリ

両ライブラリとも MIT ライセンスで、リポジトリの依存ポリシー (`deny.toml`: Apache-2.0 / MIT / BSD のみ) に準拠しています。

| クレート | 役割 |
|---|---|
| [`ovba`](https://crates.io/crates/ovba) | [MS-OVBA] 仕様実装の `vbaProject.bin` パーサ(100% safe Rust)。モジュール名・種別・ソースコードを抽出 |
| [`elixcee`](https://crates.io/crates/elixcee) | ヘッドレス Excel VBA ランタイム(独自 VM)。ワークブック読込 → `Sub` エントリポイント実行 → 保存 |

## 使い方(CLI)

```bash
# ① マクロ一覧(モジュール名・種別・ソース長を表示)
target/release/ibisOffice --list-macros book.xlsm

# ② ヘッドレス実行(Module1 の BuildReport を実行 → book-macro-run.xlsx に結果を保存)
target/release/ibisOffice --run-macro Module1.BuildReport book.xlsm

# ③ 出力先を指定
target/release/ibisOffice --run-macro BuildReport --macro-out result.xlsx book.xlsm
```

記法は `Sub名` または `Module名.Sub名` です(`--run-macro` の検索対象は標準モジュール)。

## 使い方(Rust API)

```rust
use ibis_vba::{read_modules, run_macro};

// マクロの読み取り(実行なし)
for module in read_modules(std::path::Path::new("book.xlsm"))? {
    println!("{} ({})", module.name, module.kind);
    println!("{}", module.source);
}

// オプションで明示された実行
let report = run_macro(
    std::path::Path::new("book.xlsm"),
    "Module1.BuildReport",
    std::path::Path::new("result.xlsx"),
)?;
println!("sheets: {:?}", report.sheets);   // 書込先: report.saved
```

`run_macro` はウィンドウを開きません。elixcee の VM 上でワークブックをコピーし、
実行した結果を自身のパイプラインで保存する方式です。

## 対応スコープと作業範囲

本プラグインの想定用途は **Excelブックのデータ処理系 VBA** に限定しています。

### 対応

- `If` / `For` / `Do` / `Select Case` / `With`、配列、ワークシート関数
- `Range` / `Cells` 、複数シート参照
- 保存: 同じ形式 (`.xlsx` / `.xlsm` / `.ods`) で出力

### 非対応

- `Word`/`PowerPoint` の Application オブジェクトモデル(ドキュメント/スライドの COM 自動化)
- UserForm の表示、UI 操作(Select 例外、MsgBox や MsgBox のボタン操作を含む)
- 添付型 COM オブジェクト(`CreateObject`、シェル/ファイルシステム操作、外部 DLL)
- ドキュメント/クラス/デザイナモジュールの実行(ソース閲覧は可能、`--list-macros` 参照)

未対応の操作を書いたマクロを実行したときは、elixcee が明示的にエラー名を提示します
(黙ってゼロや `#NAME?` を返すことはありません)。

## セキュリティ

VBA マクロは任意コード実行の載も持つため、以下の設計にしています。

- **実行は明示的なユーザーアクションのみ** — `--run-macro` コマンドで直接指定されたとき
  (アプリウィンドウ内の自動起動・自動実行は存在しません)
- **サンドボックス** — elixcee の VM 側で、外部副作用操作 (`CreateObject` / シェル /
  ファイルシステム操作 / `Save` / `Close`) を拒否。命令数・呼び出し深さ・文字列 /
  配列 / セル数に既定上限が設定されます (Rust API で `vm.max_instructions = None` に
  すると上限なしにできます)。`run_macro` では上限なしにしていますが、将来的には
  `ibisOffice` 側の設定へ移す予定です

マクロの状態を格納したファイルを開くだけで何かが実行されることはありません。

## ビルド

リポジトリ本体の `cargo build` に含まれます。

```bash
git clone --recurse-submodules https://github.com/ibisjp/ibis-office.git
cd ibis-office
cargo build --release
```

### Linux (Ubuntu) について

本ツリーには既存の Linux 移植向け修正 (v0.8.1 + Ubuntu 26.10 動作確認済) が混在しています:

- `version.txt` / `crates/ibis-office-app/build.rs` — Linux版のバージョン表記
- `crates/ibis-office-app/Cargo.toml` — `image` クレートのターゲット拡大(GIF書出パス)
- 描画バックエンド: `gpui-kit` → `gpui-pre-platform` → `gpui-pre-linux`
  (Wayland / X11 + wgpu/Vulkan) — crates.io 公開済みの公式 crate を利用しています

ビルドに必要なシステムライブラリ(Debian/Ubuntu):

```bash
sudo apt-get install -y pkg-config build-essential \
    libxkbcommon-dev libxkbcommon-x11-dev \
    libfontconfig1-dev libfreetype6-dev libvulkan-dev
```

## プロジェクト構成

```
crates/ibis-vba/
├── Cargo.toml   # ovba + zip(deflate) + elixcee
└── src/
    └── lib.rs   # read_modules / run_macro
```
