# Excel エビデンス貼り付けツール

画面の一部を撮影し、Excelのシートへ画像として順番に追加するWindows用ツールです。

## すぐに使う

**`EvidenceToolExcel.exe` を使えば、PythonやRustのインストール、コマンド操作は不要です。**

必要なものはWindowsとデスクトップ版Microsoft Excelだけです。

1. Excelで画像を貼り付けたいブックを開きます。
2. `EvidenceToolExcel.exe` をダブルクリックします。
3. 画面の一覧から、貼り付け先のExcelブックとシートを選択します。
4. 「スクリーンショットを撮る」を押します。
5. 暗くなった画面で、撮影したい範囲をドラッグします。
6. Excelへ画像が追加されたら、Excel側でブックを保存します。

2回目以降の画像は、シートの下方向へ順番に追加されます。

`Ctrl+Shift+E` でも撮影を開始できます。範囲選択をやめる場合は `Esc` を押してください。

> [!IMPORTANT]
> 画像は、ツール画面で選択したブックとシートへ追加されます。ツールはExcelファイルを自動保存しません。

## 主な機能

- マウスドラッグによる範囲撮影
- 開いているExcelブックとシートを画面から選択
- 選択したシートへ画像を埋め込み
- 既存のセル内容と画像の下へ自動配置
- `Ctrl+Shift+E` のグローバルホットキー
- マルチモニター対応
- exeへの専用アイコン埋め込み

## 困ったとき

- Excelと対象ブックを先に開いてください。
- Excelを後から開いた場合やブック・シートを追加した場合は、「再読込」を押してください。
- Excelでセルを編集中の場合は、`Enter` または `Esc` で編集を終えてから撮影してください。
- 複数のブックを開いている場合は、ツール画面の一覧から貼り付け先を選んでください。
- 画像追加後はExcel側でブックを保存してください。
- `Ctrl+Shift+E` が反応しない場合は、別のアプリが同じショートカットを使用している可能性があります。画面のボタンから撮影できます。

## 開発者向け

本ツールはRustで実装されています。Pythonやvenvは使用しません。

### 必要な環境

- Windows 10またはWindows 11
- デスクトップ版Microsoft Excel
- RustのMSVCツールチェーン

[Rust公式サイト](https://www.rust-lang.org/tools/install)からRustを導入後、プロジェクトのフォルダーで実行します。

```powershell
cargo run --release
```

コンパイル確認は次のコマンドで行います。

```powershell
cargo check
```

配布用exeの作成方法は[exe化手順書](docs/EXE_BUILD.md)を参照してください。

## ファイル構成

```text
├─ src/
│  └─ main.rs              Rustソース
├─ ico/
│  └─ evidence_tool.ico    exeへ埋め込むアイコン
├─ docs/
│  └─ EXE_BUILD.md         exe化手順書
├─ Cargo.toml
├─ Cargo.lock
├─ build.rs
├─ README.md
└─ .gitignore
```

`target/`、`dist/`、exeなどの生成物は `.gitignore` で除外しています。GitHubでexeを配布する場合は、リポジトリへ直接登録する代わりにReleasesへ添付してください。
