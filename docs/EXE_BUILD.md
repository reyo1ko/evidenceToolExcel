# Rust版 exe化手順書

Windows上でRustのリリースビルドを行い、単体で実行できるexeを作成します。Python、venv、PyInstallerは使用しません。

## 1. Rustを準備する

[Rust公式サイト](https://www.rust-lang.org/tools/install)から `rustup-init.exe` を取得し、既定のMSVCツールチェーンをインストールします。

新しいPowerShellを開き、次のコマンドで確認します。

```powershell
rustc --version
cargo --version
```

どちらもバージョンが表示されれば準備完了です。

## 2. アイコンを確認する

次の場所にICOファイルを配置します。

```text
ico/evidence_tool.ico
```

`build.rs` がビルド時にこのアイコンをexeへ埋め込みます。ファイル名や場所を変更する場合は、`build.rs` の指定も変更してください。

## 3. コンパイル確認

PowerShellでプロジェクトのルートを開きます。

```powershell
cargo check
```

初回は必要なRustクレートが自動でダウンロードされます。

## 4. 配布用exeをビルドする

```powershell
cargo build --release
```

完成したexeはこちらです。

```text
target/release/EvidenceToolExcel.exe
```

このexeだけで起動できます。配布先にRustやPythonをインストールする必要はありません。Windowsとデスクトップ版Microsoft Excelは必要です。

## 5. 配布用フォルダーへコピーする

必要に応じて `dist` フォルダーへコピーします。

```powershell
New-Item -ItemType Directory -Force dist | Out-Null
Copy-Item .\target\release\EvidenceToolExcel.exe .\dist\EvidenceToolExcel.exe
```

`target/`、`dist/`、`*.exe` は `.gitignore` で除外されています。GitHubで配布する場合は、`dist/EvidenceToolExcel.exe` をReleasesへ添付してください。

## 6. 動作確認

1. Excelでテスト用ブックを開きます。
2. `EvidenceToolExcel.exe` を起動します。
3. ツール画面で対象のブックとシートを選択します。
4. 「スクリーンショットを撮る」を押し、範囲を選択します。
5. 選択したシートへ画像が追加されることを確認します。
6. 2回撮影し、画像が下方向へ並ぶことを確認します。
7. Excelで保存して開き直し、画像が残っていることを確認します。
8. `Ctrl+Shift+E` でも撮影を開始できることを確認します。

## ビルドし直す場合

通常は、そのまま `cargo build --release` を再実行すれば更新されます。中間生成物も含めて作り直す場合だけ、次を実行します。

```powershell
cargo clean
cargo build --release
```

`cargo clean` は `target/` 内の生成物を削除するため、実行後のビルドには時間がかかります。
