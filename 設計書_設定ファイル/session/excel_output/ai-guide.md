# Excel出力システム ai-guide

## 概要
GAS（Web UI）からExcel帳票を出力する機能。GAS → Google Drive → Python の非同期連携。

## 更新履歴
| 日付 | セッション | 変更内容 |
|------|------------|----------|
| 2025-12-26 | #1 | 初版作成。ポーリング方式実装、フォルダID直接参照修正 |
| 2025-12-26 | #2 | クロスプラットフォーム対応（macOS/Windows）追加 |

## ナレッジベース

### [2025-12-26] Google Drive for Desktop と watchdog の非互換
**問題**: watchdog のファイル作成イベントが発火しない
**原因**: Google Drive for Desktop の同期方式が watchdog のinotify/FSEventsと互換性なし
**解決策**: ポーリング方式（5秒間隔でフォルダスキャン）に変更
**学び**: クラウド同期フォルダでは watchdog ではなくポーリングを使用

### [2025-12-26] GAS から Drive フォルダが見つからない問題
**問題**: Python処理完了後、GASの「状態確認」でファイルが見つからない
**原因**: `pendingFolder.getParents().next()` で取得した親フォルダが期待と異なる
**解決策**: フォルダIDを設定シートに直接登録し、`DriveApp.getFolderById()` で参照
**学び**: 親フォルダ経由の検索は信頼性が低い。フォルダIDを直接指定すべき

### [2025-12-26] クロスプラットフォーム対応（macOS/Windows）
**課題**: macOS と Windows でGoogle Driveのパスが異なる
**パス例**:
- macOS: `/Users/{user}/Library/CloudStorage/GoogleDrive-xxx/マイドライブ`
- Windows: `G:\マイドライブ` または `C:\Users\{user}\Google Drive\マイドライブ`

**解決策**: テンプレート方式
1. `settings_template.yaml` - `${GOOGLE_DRIVE_BASE}` プレースホルダー使用
2. `setup.sh` (macOS) / `setup.ps1` (Windows) - 環境検出＆設定生成
3. `settings.yaml` は `.gitignore` で除外（環境固有）

**セットアップ手順**:
```bash
# macOS
cd python && bash setup.sh

# Windows (PowerShell)
cd python
.\setup.ps1
```

**学び**: 環境固有設定はテンプレート＋セットアップスクリプトで管理

## 実装パターン

### システムアーキテクチャ
```
[GAS] → pending/*.json → [Python監視] → Excel生成 → processed/, output/
                                                              ↓
[GAS] ← 状態確認 ← processed/*_result.json ←────────────────┘
```

### Python監視起動
```bash
cd python
python3 unified_transcriber.py --watch
```

### 設定シート必須項目
| キー | 用途 |
|------|------|
| PYTHON_PENDING_FOLDER_ID | リクエストJSON格納先 |
| PYTHON_PROCESSED_FOLDER_ID | 処理完了JSON格納先 |
| PYTHON_OUTPUT_FOLDER_ID | Excel出力先 |

## 関連ファイル
| ファイル | 役割 |
|---------|------|
| `gas/reportExporter.js` | JSON生成・状態確認API |
| `gas/excelExportBridge.js` | Python連携ブリッジ |
| `python/drive_watcher.py` | フォルダ監視（ポーリング） |
| `python/unified_transcriber.py` | Excel転記エンジン |
| `python/settings.yaml` | Python側設定（.gitignore対象） |
| `python/settings_template.yaml` | 設定テンプレート（Git管理） |
| `python/setup.sh` | macOSセットアップスクリプト |
| `python/setup.ps1` | Windowsセットアップスクリプト |

## 設計書との乖離記録
| 検出日 | 乖離内容 | 対応状況 |
|--------|----------|----------|
| - | なし | - |
