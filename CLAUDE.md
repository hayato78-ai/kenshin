# CLAUDE.md - CDmedical健診システム プロジェクトルール

## プロジェクト概要

健診結果管理システム。
Google Apps Script (GAS) + Google Sheets で構築。

---

## 開発原則
- 速度より正確性を優先
- 設計書を常に参照し、乖離があれば実装前に報告
- 推測で進めず、不明点は必ず質問
- 各実装ブロック完了後、動作確認してから次へ
- データ構造は実運用に合わせて作成

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
├── docs/                # ドキュメント
│   ├── ARCHITECTURE.md       # ★ アーキテクチャ設計書（統合版）
│   ├── FUNCTION_INVENTORY.md # ★ 機能棚卸し表（移植判定）
│   ├── cross-cutting-concerns.md  # 横断的ナレッジ
│   └── archive/              # 旧設計書（参照のみ）
├── gas/                 # ★ GASソースコード（この中だけ編集）
│   ├── .clasp.json      # clasp設定（Script ID）
│   ├── appsscript.json  # GAS設定
│   ├── portal.js        # Webアプリエントリーポイント
│   ├── portalApi.js     # ポータルAPI
│   ├── master_examItem.js  # 検査項目マスタCRUD
│   ├── Config.js        # 設定値
│   ├── templates/       # HTMLテンプレート
│   │   ├── portal.html
│   │   ├── css.html
│   │   └── js.html
│   └── ...
├── python/              # Python補助スクリプト
└── 設計書_設定ファイル/  # 設計書・設定ファイル
    ├── 1220_new/        # ★ 最新設計書v2（DATA_STRUCTURE_DESIGN_v2.md）
    ├── id-heart/        # iD-Heart参考資料
    ├── session/         # 機能別ナレッジ格納
    │   ├── excel_output/
    │   │   └── ai-guide.md  # ← 機能別ナレッジ
    │   ├── csv_import/
    │   │   └── ai-guide.md
    │   └── [機能名]/
    │       └── ai-guide.md
    └── *.json           # マッピング設定
```

---

## 📚 ドキュメント参照ルール

### 正式ドキュメント（必ず参照）

| ファイル | 役割 | 参照タイミング |
|----------|------|----------------|
| CLAUDE.md | プロジェクトルール | セッション開始時 |
| docs/ARCHITECTURE.md | アーキテクチャ全体像 | 設計判断時 |
| docs/FUNCTION_INVENTORY.md | 機能一覧・移植判定 | 機能実装時 |
| 設計書_設定ファイル/1220_new/DATA_STRUCTURE_DESIGN_v2.md | データ構造定義 | DB操作時 |

### アーカイブ（参照のみ・編集禁止）

`docs/archive/` 配下のファイルは旧設計書です。
- 新規実装の参考として参照可能
- 編集・更新は禁止
- 正式版との矛盾がある場合は正式版を優先

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

適宜変更があるため、この部分を参照・編集は許可を求めること

| シート名 | 用途 | 状態 |
|----------|------|------|
| T_受診者 | 受診者基本情報 | ✅ |
| T_受診記録 | 受診ごとの記録 | ✅ |
| 検査結果 | 検査結果（**横持ち**: 1患者1行、101列） | ✅ |
| T_所見 | 所見データ（縦持ち） | ✅ |
| M_検査項目 | 検査項目マスタ（46列） | ✅ |
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

### 🔴 確認必須（）

| 操作 | 理由 |
|------|------|
| ファイル削除 | 復元が面倒 |
| 新規HTMLファイル作成 | 複数UI問題の防止 |
|新規.mdや設計書の作成| ファイルの膨大を防止|


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

## 📝 セッション追跡 & ai-guide 運用ルール

### 概要
セッション中の変更・学びを記録し、ドメイン知識として蓄積する。
ai-guideは機能/領域ごとに配置し、継続的に育てる。

### ai-guide の配置ルール
- **ベースパス**: `設計書_設定ファイル/session/`（プロジェクトルートからの相対パス）
- **ファイル名**: `ai-guide.md`（統一）
- **フルパス**: `設計書_設定ファイル/session/[機能名]/ai-guide.md`
- **例**: `設計書_設定ファイル/session/excel_output/ai-guide.md`

### セッション開始時
**ユーザーが指示**:
```
作業対象: [機能名]
```

**Claudeが自動解決**:
1. `設計書_設定ファイル/session/[機能名]/ai-guide.md` を参照
2. ディレクトリが存在しない場合 → 新規作成を提案（`_template/ai-guide-template.md` をコピー）
3. ai-guideが存在しない場合 → 初期化を提案
4. 前回から1週間以上経過していれば → 要約更新を提案

### セッション終了時
**ユーザーが終了を宣言** → **Claudeが以下を実行**:

1. セッション内容を要約
2. ai-guide.md への追記内容を**提案**（ユーザー承認待ち）:
   - 実装した機能・修正内容
   - 遭遇した問題と解決策
   - 新しいパターン・設計判断
   - 設計書との乖離検出（修正は別途判断）
3. 承認後、更新を実行

### ai-guide.md フォーマット
```markdown
# [機能名] ai-guide

## 概要
[この機能の目的・スコープ]

## 更新履歴
| 日付 | セッション | 変更内容 |
|------|------------|----------|
| 2024-12-26 | #1 | 初版作成 |

## ナレッジベース
### [YYYY-MM-DD] [トピック]
**問題**:
**原因**:
**解決策**:
**学び**:

## 実装パターン
[この機能で使用するパターン・規約]

## 関連ファイル
[主要ファイルとその役割]

## 設計書との乖離記録
| 検出日 | 乖離内容 | 対応状況 |
|--------|----------|----------|
```

### 更新トリガー
| タイミング | 必須度 | 実行者 |
|------------|--------|--------|
| セッション終了時 | ✅ 必須 | Claude（提案）→ ユーザー（承認） |
| 重要な設計変更時 | ✅ 必須 | Claude（即座に記録） |
| 1週間以上経過後の開始時 | 推奨 | Claude（要約提案） |
| ナレッジ10件超過時 | 推奨 | Claude（統合提案） |

### ai-guide 陳腐化防止
- セッション開始時に前回からの経過日数をチェック
- 実装パターンに矛盾を検出したら即座に指摘
- 新しい概念が既存を置き換えたら統合を提案

### 横断的なナレッジの扱い
複数機能に影響する学びは、以下のルールで管理する：

**横断的ナレッジの定義**（以下のすべてを満たす場合のみ）:
- 3つ以上の機能/領域に影響する
- プロジェクト全体の設計方針に関わる
- 個別のai-guideに記載すると重複が多くなる

**管理方法**:
- `/docs/cross-cutting-concerns.md` に統合ナレッジを記録
- 各ai-guideからは参照リンクのみ設置（内容を複製しない）
- 例：「BMLコード標準化方針」「命名規則」「エラーハンドリング方針」

**横断的ナレッジにしないもの**:
- 2つ以下の機能にしか関係しない → 各ai-guideに個別記載
- 一時的な回避策 → 該当機能のai-guideのみに記載
- 特定技術の詳細 → 該当機能のai-guideのみに記載

### セッション番号の採番ルール
- 各ai-guideごとに独立して採番（#1, #2, #3...）
- 更新履歴に記録し、ナレッジベースの各エントリからも参照可能にする
- 例：「セッション#3で解決」のように記載

---

## 連絡先・参考

- スプレッドシートID: `健診結果入力システム_マスタ`
- GAS Script ID: `.clasp.json` を参照
