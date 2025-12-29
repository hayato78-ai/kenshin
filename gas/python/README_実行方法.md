# 人間ドック帳票生成システム - 実行方法

## 概要

GASのボタンを押すと、自動的にCSVからExcel帳票が生成されます。
このシステムを動作させるには、PCでPythonプログラムを起動しておく必要があります。

---

## 初回セットアップ（管理者がPCごとに1回だけ実施）

### 手順1: Pythonのインストール

| OS | インストール方法 |
|----|------------------|
| Windows | https://www.python.org/downloads/ からダウンロードしてインストール |
| Mac | 標準でインストール済み（追加作業不要） |

※ Windowsの場合、インストール時に「Add Python to PATH」にチェックを入れてください

### 手順2: 必要なライブラリのインストール

コマンドプロンプト（Windows）またはターミナル（Mac）で以下を実行：

```bash
pip install openpyxl pyyaml watchdog
```

### 手順3: 自動起動の設定（推奨）

自動起動を設定すると、PC起動時に監視が自動で開始されます。
**ユーザーは何も操作する必要がなくなります。**

#### Windows の場合

1. `Win + R` キーを押す
2. `shell:startup` と入力してEnter
3. 開いたフォルダに `start_watcher.bat` のショートカットを作成

#### Mac の場合

1. 「システム設定」→「一般」→「ログイン項目」を開く
2. 「+」ボタンをクリック
3. `start_watcher_mac.command` を選択して追加

---

## 日常の使い方

### 自動起動を設定済みの場合

**何もする必要はありません。**
PCを起動すれば自動的に監視が開始されます。

### 手動で起動する場合

| OS | 方法 |
|----|------|
| Windows | `start_watcher.bat` をダブルクリック |
| Mac | `start_watcher_mac.command` をダブルクリック |

黒い画面（コマンドプロンプト/ターミナル）が表示され「監視中...」と表示されたらOKです。

---

## 動作の流れ

```
1. ユーザーがGASの「Excel出力」ボタンを押す
      ↓
2. GASがリクエストJSONを pending フォルダに作成
      ↓
3. 監視プログラムが自動検知
      ↓
4. CSVを読み込み、Excelに転記
      ↓
5. 02_出力フォルダ にExcelファイルが生成される
```

---

## トラブルシューティング

### Excelが生成されない場合

1. 監視プログラムが起動しているか確認
   - タスクバー（Windows）またはDock（Mac）に黒い画面があるか
2. `python/pending` フォルダにJSONファイルが残っていないか確認
3. `python/error` フォルダにエラーファイルがないか確認

### エラーが発生した場合

`drive_watcher.log` ファイルを確認してください。

### 監視プログラムが起動しない場合

- Pythonがインストールされているか確認
- 必要なライブラリがインストールされているか確認

---

## フォルダ構成

```
python/
├── pending/      ← GASがリクエストJSONを置く場所
├── processed/    ← 処理完了したJSON
├── error/        ← エラー発生時のJSON
├── drive_watcher.py       ← 監視プログラム
├── unified_transcriber.py ← 転記エンジン
├── settings.yaml          ← 設定ファイル
├── start_watcher.bat      ← Windows起動用
└── start_watcher_mac.command ← Mac起動用
```

---

## 管理者向け情報

### 手動実行（デバッグ用）

```bash
# CSVから直接実行
python3 unified_transcriber.py --csv /path/to/BML.csv

# JSON経由で実行
python3 unified_transcriber.py request.json --type HUMAN_DOCK
```

### 設定ファイル

`settings.yaml` でテンプレートパス、出力先などを変更できます。
