# CLAUDE.md - CDmedical健診システム プロジェクトルール

## プロジェクト概要

CDmedical渋谷駅前クリニックの健診結果管理システム。
Google Apps Script (GAS) + Google Sheets で構築。

---

## 開発原則
- 速度より正確性を優先
- 設計書を常に参照し、乖離があれば実装前に報告
- 推測で進めず、不明点は必ず質問
- 各実装ブロック完了後、動作確認してから次へ
- データ構造は正規化を基本とし、非正規化は根拠を明示

## 🚨 重要: 開発フロー（必ず守ること）

### GASコード編集時の鉄則

```
1. ローカルで編集 → 2. clasp push → 3. 動作確認 → 4. clasp deploy
```

**絶対にやってはいけないこと:**
- GASエディタ（ブラウザ）での直接編集
- clasp push せずにローカルだけ編集して終わり
- リモートの状態を確認せずに push

### 編集前の確認コマンド

```bash
# 1. リモートとローカルの差分確認
clasp pull --force  # リモートの状態を取得（上書き注意）
clasp status        # 変更ファイル確認

# 2. 安全な編集開始
git status          # Gitの状態確認
git pull            # 最新を取得
```

### 編集後の反映コマンド

```bash
# 1. ローカル変更をリモートに反映
clasp push

# 2. 動作確認（開発版URL）
# https://script.google.com/macros/s/{SCRIPT_ID}/dev

# 3. 本番デプロイ（確認後）
clasp deploy --description "変更内容の説明"
```

---

## ディレクトリ構造

```
05_development/
├── CLAUDE.md            # このファイル（プロジェクトルール）
├── gas/                 # ★ GASソースコード（この中だけ編集）
│   ├── .clasp.json      # clasp設定（Script ID）
│   ├── appsscript.json  # GAS設定
│   ├── portal.js        # Webアプリエントリーポイント
│   ├── portalApi.js     # ポータルAPI
│   ├── master_examItem.js  # 検査項目マスタCRUD ★NEW
│   ├── Config.js        # 設定値
│   ├── templates/       # HTMLテンプレート
│   │   ├── portal.html
│   │   ├── css.html
│   │   └── js.html
│   └── ...
├── python/              # Python補助スクリプト
└── 設計書_設定ファイル/  # 設計書・設定ファイル
    ├── 1220_new/        # 最新設計書v2
    ├── id-heart/        # iD-Heart参考資料
    └── *.json           # マッピング設定
```

---

## GASファイルの役割

| ファイル | 役割 | 編集頻度 |
|----------|------|----------|
| `portal.js` | Webアプリ doGet() エントリーポイント | 低 |
| `portalApi.js` | フロントエンドから呼ばれるAPI | 高 |
| `master_examItem.js` | 検査項目マスタCRUD | 中 |
| `patientManager.js` | 受診者CRUD | 中 |
| `billingManager.js` | 請求管理 | 中 |
| `Config.js` | DB設定・定数 | 低 |
| `templates/portal.html` | メインUI | 高 |
| `templates/css.html` | スタイル | 中 |
| `templates/js.html` | フロントエンドJS | 高 |

---

## コーディング規約

### GAS (JavaScript)

```javascript
// 関数コメント必須
/**
 * 受診者を検索
 * @param {Object} criteria - 検索条件
 * @returns {Array<Object>} 検索結果
 */
function searchPatients(criteria) {
  // ...
}

// ログ出力
logInfo('処理開始: ' + patientId);
logError('エラー', error);

// エラーハンドリング
try {
  // 処理
} catch (e) {
  logError('functionName', e);
  throw e;
}
```

### HTML/CSS

- CSSは `templates/css.html` に集約
- JavaScriptは `templates/js.html` に集約
- インラインスタイル禁止

---

## データベース（スプレッドシート）

### シート一覧

| シート名 | 用途 | 状態 |
|----------|------|------|
| T_受診者 | 受診者基本情報 | ✅ |
| T_受診記録 | 受診ごとの記録 | ✅ |
| T_検査結果 | 検査結果（縦持ち） | ✅ |
| M_検査項目 | 検査項目マスタ（46列） | ✅ NEW |
| M_検査所見 | 所見テンプレート | 予定 |
| M_コース | 健診コース | 予定 |
| M_団体 | 企業・団体 | ✅ |
| 設定 | システム設定 | ✅ |

### 重要: スプレッドシートの直接編集

- 構造変更（列追加・削除）は必ず設計書を更新してから
- データ編集はWebアプリ経由を推奨
- 直接編集した場合はログに記録

---

## デプロイ情報

| 環境 | URL | 用途 |
|------|-----|------|
| 本番 | `https://script.google.com/macros/s/{DEPLOY_ID}/exec` | ユーザー向け |
| 開発 | `https://script.google.com/macros/s/{SCRIPT_ID}/dev` | テスト用 |

### デプロイ手順

```bash
# 1. 変更をpush
clasp push

# 2. 開発版で確認
open "https://script.google.com/macros/s/{SCRIPT_ID}/dev"

# 3. 本番デプロイ
clasp deploy --description "v1.x.x - 変更内容"

# 4. デプロイ一覧確認
clasp deployments
```

---

## トラブルシューティング

### リモートとローカルが乖離した場合

```bash
# 1. リモートの状態をバックアップ
mkdir backup_remote_$(date +%Y%m%d)
cd backup_remote_$(date +%Y%m%d)
clasp clone {SCRIPT_ID}

# 2. 差分確認
diff -r ../gas ./

# 3. ローカルを正として反映
cd ..
clasp push --force
```

### UIが表示されない場合

1. `portal.js` の `doGet()` 関数を確認
2. `templates/portal.html` のパス確認
3. デプロイが最新か確認 (`clasp deployments`)

---

## 🚨 UI/テンプレート関連の厳格ルール

### 絶対禁止（複数UI問題の防止）

1. **新しいHTMLファイルの作成禁止**
   - `templates/` フォルダ以外にHTMLを作成しない
   - 新しいUIが必要な場合は「提案のみ」して確認を求める
   - 勝手に dashboard.html, index.html 等を作成しない

2. **既存UIの複製禁止**
   - `portal.html` のコピーを作成しない
   - 「別バージョン」「新デザイン」のUIを勝手に作成しない
   - UIの大幅変更は必ず事前確認

3. **ui/ フォルダへの追加禁止**
   - `_unused_ui/` は触らない（アーカイブ済み）
   - 新しい `ui/` フォルダを作成しない
   - UIコードは `templates/` のみに配置

4. **doGet() 関数の変更禁止**
   - `portal.js` の `doGet()` は変更しない
   - 別のエントリーポイントを作成しない
   - 変更が必要な場合は必ず事前確認

### 正式なUI構成（これ以外は作成禁止）

```
gas/templates/          ← 唯一のUIフォルダ
├── portal.html         ← メインUI（タブ形式）
├── css.html            ← スタイル
└── js.html             ← JavaScript

gas/_unused_ui/         ← 触らない（アーカイブ）
```

---

## 🔒 確認ルール（シンプル版）

### 🔴 確認必須（この2つだけ）

| 操作 | 理由 |
|------|------|
| ファイル削除 | 復元が面倒 |
| 新規HTMLファイル作成 | 複数UI問題の防止 |

### 🟢 報告のみでOK（確認不要）

- 既存ファイルの編集
- `clasp push`
- `clasp deploy` ← **push成功後は自動実行**
- `git commit / push`
- コメント追加、軽微な修正

### 📦 標準デプロイフロー（clasp push後は必ず実行）

```bash
# 1. push成功後
clasp push

# 2. 即座に本番デプロイ（確認不要）
clasp deploy --description "変更内容"

# 3. 結果報告
clasp deployments | head -5
```

### 作業完了時の報告フォーマット

```
実行完了:
- portal.html: 〇〇を修正
- css.html: 〇〇を追加
- clasp push 完了
```

---

## ❌ 禁止事項

### コード管理
1. GASエディタでの直接編集
2. clasp push せずにローカル編集のみ
3. 本番デプロイ前のテスト省略

### UI関連（複数UI問題の防止）
4. `templates/` 以外へのHTML作成
5. `portal.html` の複製・別バージョン作成
6. `doGet()` 関数の無断変更
7. `_unused_ui/` フォルダへのアクセス

---

## 🔀 Git バージョン管理ルール

### 基本原則

- **こまめなコミット**: 機能単位・修正単位でコミット
- **機能実装後は必ずpush**: ローカルのみに留めない
- **意味のあるコミットメッセージ**: 何を変更したか明確に

### コミットタイミング

| タイミング | 必須度 |
|------------|--------|
| 機能実装完了 | ✅ 必須 |
| バグ修正完了 | ✅ 必須 |
| clasp deploy 後 | ✅ 必須 |
| 大きな変更の途中経過 | 推奨 |
| 作業終了時 | ✅ 必須 |

### コミットメッセージ規約

```bash
# プレフィックス
feat:     # 新機能
fix:      # バグ修正
refactor: # リファクタリング
docs:     # ドキュメント
style:    # スタイル変更（機能に影響なし）

# 例
git commit -m "feat: 検査項目マスタUI追加"
git commit -m "fix: item_code先頭0消失問題を修正"
```

### 標準ワークフロー

```bash
# 1. 作業開始前
git status
git pull

# 2. 機能実装
# ... コード編集 ...

# 3. clasp反映
clasp push
clasp deploy --description "変更内容"

# 4. git commit & push（機能単位で）
git add -A
git commit -m "feat: 機能説明"
git push

# 5. 確認
git log --oneline -3
```

### 注意事項

- **mainブランチで直接作業OK**（小規模プロジェクトのため）
- **push前にstatus確認**: 不要なファイルが含まれていないか
- **diverged時の対処**: rebaseよりmerge優先

---

## 連絡先・参考

- スプレッドシートID: `健診結果入力システム_マスタ`
- GAS Script ID: `.clasp.json` を参照
