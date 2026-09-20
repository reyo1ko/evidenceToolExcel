# exe化手順書（Windows）

Windows上でPyInstallerを使用し、単一のexeを作成します。以下はPowerShell用です。すべてプロジェクトのルート（`evidenceToolExcel.py` があるフォルダー）で実行してください。

## 1. 環境を準備する

[README](../README.md)に従ってPythonとvenvを用意します。作成済みのvenvを有効化し、依存パッケージとビルドツールをインストールします。

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install --upgrade pyinstaller
```

activateが使えない場合は、以降の `python` を `.\.venv\Scripts\python.exe` に置き換えます。

Excelでテスト用ブックを開き、ソース版で撮影・画像追加ができることを確認してから終了します。

```powershell
python evidenceToolExcel.py
```

## 2. アイコンを配置する

用意した `.ico` ファイルを、ソースと同じ場所に `evidence_tool.ico` という名前で置きます。別名の場合はビルドコマンドの `--icon` も変更してください。

```text
プロジェクト/
├─ evidenceToolExcel.py
├─ evidence_tool.ico       ← 用意するアイコン
├─ requirements.txt
├─ README.md
├─ .gitignore
└─ docs/
   └─ EXE_BUILD.md
```

画像の拡張子を変更するだけではICO形式になりません。ICO形式で書き出したファイルを使用してください。`.ico` はGit管理対象です。

`--icon` はexeファイルのアイコンを設定します。現在のタスクトレイアイコンはコードで描画しているため、この指定では変わりません。ウィンドウやトレイにも同じアイコンを使う場合は、別途アプリ側の読み込み処理が必要です。

## 3. ビルドする

起動中のツールを終了してから実行します。

```powershell
python -m PyInstaller --clean --noconfirm --onefile --noconsole --name EvidenceToolExcel --icon "evidence_tool.ico" evidenceToolExcel.py
```

| オプション | 内容 |
| --- | --- |
| `--clean` | ビルドキャッシュを削除してから作成 |
| `--noconfirm` | 既存のビルド出力を確認なしで置き換え |
| `--onefile` | 単一のexeにまとめる |
| `--noconsole` | コンソールを表示しない |
| `--name` | 出力するexeの名前 |
| `--icon` | exeのICOファイル |

アイコンが未準備の場合は `--icon "evidence_tool.ico"` を省略して仮ビルドできます。

生成物は次のとおりです。

```text
dist/EvidenceToolExcel.exe   配布用ファイル
build/                      ビルド中間ファイル
EvidenceToolExcel.spec       PyInstaller設定ファイル
```

これらは `.gitignore` で除外しています。この手順では毎回ソースとコマンドからビルドします。今後 `.spec` を手動編集して管理する場合は、`.gitignore` の `*.spec` を見直してください。

## 4. 動作確認する

ソース版を終了し、Excelでテスト用ブックを開いて実行します。

```powershell
.\dist\EvidenceToolExcel.exe
```

- exeにアイコンが表示され、ツールが起動すること
- 貼り付け先の選択と `Ctrl+Shift+E` による範囲選択ができること
- 画像を2回追加すると縦に並ぶこと
- シート切り替え後は切り替え先に追加されること
- Excelで保存して開き直しても画像が残ること
- ツールを終了して再起動できること

## 5. 配布する

動作確認済みの `dist/EvidenceToolExcel.exe` を配布します。配布先にPythonやvenvは不要ですが、Windowsとデスクトップ版Excelが必要です。exeのアイコンはビルド時に埋め込まれるため、`.ico` の同梱は不要です。

GitHubではソースをリポジトリに登録し、exeはReleasesなどに添付できます。

作業が終わったらvenvを解除します。

```powershell
deactivate
```

## 起動しない場合の調査用ビルド

コンソール表示ありのexeを作成し、PowerShellから実行してエラーを確認します。

```powershell
python -m PyInstaller --clean --noconfirm --onefile --console --name EvidenceToolExcel-debug --icon "evidence_tool.ico" evidenceToolExcel.py
.\dist\EvidenceToolExcel-debug.exe
```

ソース版とexe版は同時起動できません。「起動済み」と表示されたら既存のツールを終了してください。

オプションの詳細は[PyInstaller公式ドキュメント](https://www.pyinstaller.org/en/stable/usage.html)を参照してください。
