# エビデンス自動貼り付けツール

画面の選択範囲をキャプチャして、Excelのアクティブシートに画像を縦並びで追加するWindows用ツールです。画像はブックに埋め込まれます。

## すぐに使う

**配布されている `EvidenceToolExcel.exe` を使えば、Pythonのインストールやコマンド操作は不要です。**

必要なものは、Windowsとデスクトップ版Microsoft Excelだけです。

1. Excelで画像を貼り付けたいブックを開きます。
2. `EvidenceToolExcel.exe` をダブルクリックします。
3. `Ctrl+Shift+E` または「スクショ起動」を押し、撮影したい範囲をドラッグします。
4. Excelに画像が追加されたら、ブックを保存します。

繰り返し撮影すると、画像はシートの下方向に順番に追加されます。

> [!NOTE]
> Windowsから警告が表示された場合は、配布元が信頼できることを確認してから実行してください。

## 機能

- `Ctrl+Shift+E` または「スクショ起動」ボタンで範囲選択
- 接続先Excelで開かれているブックから貼り付け先を選択
- 選択したブックのアクティブシートに追記
- 既存の図形と使用セル範囲の下に余白を空けて配置
- マルチモニター対応

画像サイズは1ピクセルを0.75ポイントとして配置します。

## 動作環境

- Windows
- デスクトップ版Microsoft Excel
- exeを利用する場合、Pythonのインストールは不要
- ソースから実行・開発する場合はPythonとpip（開発環境ではPython 3.14.7を使用）

## 詳しい使い方

- 貼り付け先のブックが1つの場合は、自動で選択されます。
- 複数のブックを開いている場合は、「貼り付け先を変更」から選択します。
- シートを切り替えると、次の画像は切り替え後のシートに追加されます。
- 範囲選択を取り消すには、`Esc` または「キャンセル」を押します。
- ツールを終了するには、「終了」ボタンまたはタスクトレイのメニューを使います。
- Excelブックは自動保存されません。作業後にExcelで保存してください。

## 開発者向け：ソースから実行する

以下はWindows PowerShellの手順です。取得したプロジェクトの `evidenceToolExcel.py` があるフォルダーをPowerShellで開いて実行します。

### 1. venvを作成する（初回のみ）

```powershell
py -m venv .venv
```

`py` が見つからず `python` が利用できる場合は、`python -m venv .venv` を使用してください。

### 2. venvを有効にする（activate）

```powershell
.\.venv\Scripts\Activate.ps1
```

新しくPowerShellを開くたびに実行します。有効化するとプロンプトの先頭に `(.venv)` が表示されます。

スクリプトの実行が制限されている環境では、activateせずにvenvのPythonを直接指定できます。

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe evidenceToolExcel.py
```

### 3. pipで依存パッケージをインストールする

venvを有効にした状態で実行します。

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

`requirements.txt` を変更した場合も、2行目を再実行してください。

### 4. 起動する

先にExcelで貼り付け先のブックを開きます。

```powershell
python evidenceToolExcel.py
```

### 5. venvを無効にする（deactivate）

ツールを終了し、PowerShellで実行します。

```powershell
deactivate
```

仮想環境は削除されません。次回はactivateして利用できます。

コマンドプロンプト（cmd）では、有効化は `.venv\Scripts\activate.bat`、無効化は `deactivate` です。

詳細は[Python公式のvenv手順](https://docs.python.org/3/library/venv.html)を参照してください。

## エラーが出る場合

- Excelのセル編集やダイアログを終了して再試行してください。
- 対象ブックを閉じた場合は開き直し、「貼り付け先を変更」で再選択してください。
- 複数のExcelプロセスがあると、接続先以外のブックが一覧に表示されない場合があります。
- ソースを更新した場合は、ツールを終了して起動し直してください。

## exe化・アイコンの設定

[exe化手順書](docs/EXE_BUILD.md)に、PyInstallerのインストール、`.ico` の配置、ビルド、配布前の確認方法を記載しています。

## GitHubに公開するファイル

ソース、`requirements.txt`、README、`docs/`、`.gitignore`、用意した `.ico` を管理します。

`.gitignore` でvenv、ローカル検証用テスト、キャッシュ、ビルド生成物などを除外しています。テストは手元に残ります。すでにGit管理されているファイルには除外設定が適用されないため、該当する場合は `git rm --cached <ファイル名>` で追跡を解除してください。
