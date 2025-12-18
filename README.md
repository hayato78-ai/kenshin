# 健診結果DB 統合システム - GASコード

## 概要

健診結果DB_設計書_v1.md に基づいて作成されたGoogle Apps Scriptコードです。

- Phase 1（基盤構築）: スプレッドシートDBの作成、マスタデータ投入、CRUD操作、判定自動計算
- Phase 2（UI開発）: Webアプリ画面（メニュー、検索、一覧、詳細、結果表示、新規登録、結果入力）

## ファイル構成

### サーバーサイド（.gs）

| ファイル | 説明 |
|---------|------|
| Config.gs | 設定・定数・シート定義 |
| SetupDB.gs | DB初期セットアップ・シート作成 |
| CRUD.gs | CRUD操作・ID自動採番 |
| JudgmentEngine.gs | 判定自動計算エンジン |
| Main.gs | メニュー・UI・エントリーポイント |
| UI.gs | Webアプリ用サーバー関数 |

### クライアントサイド（ui/）

| ファイル | 説明 |
|---------|------|
| ui/Index.html | メインエントリーポイント |
| ui/Styles.html | 共通CSSスタイル |
| ui/Common.html | 共通コンポーネント（ヘッダー、ナビ、フッター） |
| ui/JavaScript.html | 共通JavaScript（画面遷移、API呼び出し） |
| ui/ScreenMenu.html | メニュー画面 (SCR-002) |
| ui/ScreenSearch.html | 受診者検索画面 (SCR-003) |
| ui/ScreenList.html | 受診者一覧画面 (SCR-004) |
| ui/ScreenDetail.html | 受診者詳細画面 (SCR-005) |
| ui/ScreenResult.html | 検査結果表示画面 (SCR-006) |
| ui/ScreenRegister.html | 新規受診登録画面 (SCR-007) |
| ui/ScreenInput.html | 検査結果入力画面 (SCR-008) |

## デプロイ手順

### 1. 新規スプレッドシートを作成

Google Driveで新しいスプレッドシートを作成し、名前を「健診結果DB_2025」に設定します。

### 2. Apps Scriptを開く

メニューから「拡張機能」→「Apps Script」を選択します。

### 3. ファイルをコピー

このフォルダ内の各 `.gs` ファイルの内容を、Apps Scriptエディタにコピーします。

- 最初に表示される `コード.gs` を `Config.gs` に名前変更し、Config.gs の内容を貼り付け
- 「+」ボタンから新規スクリプトファイルを追加し、他のファイルも同様に作成

### 4. 初期セットアップを実行

1. `SetupDB.gs` の `setupDatabase` 関数を実行
2. 権限の承認ダイアログが表示されたら許可
3. セットアップ完了のダイアログが表示されればOK

### 5. 動作確認

- スプレッドシートを再読み込み
- メニューに「健診DB」が追加されていることを確認
- 「健診DB」→「セットアップ」→「DB整合性チェック」で問題がないことを確認

## 機能一覧

### セットアップ

- `setupDatabase()` - DB初期セットアップ（6シート作成、マスタ投入）
- `resetMasterData()` - マスタデータのリセット
- `validateDatabase()` - DB整合性チェック

### 受診者CRUD

- `createPatient(data)` - 受診者作成（ID自動採番: P00001形式）
- `getPatientById(patientId)` - 受診者取得
- `searchPatients(criteria)` - 受診者検索
- `updatePatient(patientId, data)` - 受診者更新
- `deletePatient(patientId)` - 受診者削除

### 受診記録CRUD

- `createVisitRecord(data)` - 受診記録作成（ID自動採番: YYYYMMDD-NNN形式）
- `getVisitRecordById(visitId)` - 受診記録取得
- `getVisitRecordsByPatientId(patientId)` - 受診履歴取得
- `updateVisitRecord(visitId, data)` - 受診記録更新

### 検査結果CRUD

- `createTestResult(data)` - 検査結果作成（ID自動採番: R00001形式）
- `getTestResultsByVisitId(visitId)` - 検査結果取得
- `updateTestResult(resultId, data)` - 検査結果更新
- `deleteTestResultsByVisitId(visitId)` - 検査結果削除

### 判定エンジン

- `calculateJudgment(itemId, value, gender)` - 個別判定
- `calculateOverallJudgment(visitId)` - 総合判定
- `recalculateAllJudgments(visitId, gender)` - 全判定再計算
- `inputTestResultWithJudgment(visitId, itemId, value, gender)` - 結果入力+自動判定

## シート構成

設計書4章に基づく6シート（+コースマスタ）：

| シート名 | 用途 |
|---------|------|
| 受診者マスタ | 受診者基本情報（P00001形式ID） |
| 受診記録 | 受診ごとの記録（YYYYMMDD-NNN形式ID） |
| 検査結果 | 縦持ち形式の結果データ（R00001形式ID） |
| 項目マスタ | 検査項目と判定基準 |
| 検診種別マスタ | 検診種別定義（DOCK/REGULAR/EMPLOY/ROSAI/SPECIFIC） |
| コースマスタ | 人間ドックコース定義 |
| 保健指導記録 | 保健指導履歴（G00001形式ID） |

## 判定基準

設計書5章に基づくABCD判定：

| 判定 | 意味 | 色 |
|------|------|-----|
| A | 異常なし | 緑 #e8f5e9 |
| B | 軽度異常 | 黄 #fff8e1 |
| C | 要経過観察 | 橙 #fff3e0 |
| D | 要精密検査 | 赤 #ffebee |

## 関連ドキュメント

- 設計書: `00_管理ドキュメント/健診結果DB_設計書_v1.md`
- UI設計書: `00_管理ドキュメント/健診結果DB_UI設計書_v1.md`
- 作業指示書: `00_管理ドキュメント/LLM_INSTRUCTIONS.md`
- ドキュメント一覧: `00_管理ドキュメント/INDEX.md`

## Webアプリのデプロイ手順

### 6. UIファイルをコピー

`ui/` フォルダ内のHTMLファイルも同様にApps Scriptに追加します。

- 「+」ボタン → 「HTMLファイル」を選択
- ファイル名を `ui/Index` のように入力（`.html` は自動付与）
- 各HTMLファイルの内容をコピー

### 7. Webアプリとしてデプロイ

1. Apps Scriptエディタで「デプロイ」→「新しいデプロイ」
2. 種類を「ウェブアプリ」に設定
3. 次のユーザーとして実行: 「自分」
4. アクセスできるユーザー: 「自分のみ」または組織内
5. 「デプロイ」をクリック

### 8. Webアプリにアクセス

デプロイ完了後に表示されるURLにアクセスすると、健診結果管理システムのWebアプリが利用できます。

## 画面一覧

| 画面ID | 画面名 | 説明 |
|--------|--------|------|
| SCR-002 | メニュー | ダッシュボード・メニュー画面 |
| SCR-003 | 受診者検索 | 検索条件入力 |
| SCR-004 | 受診者一覧 | 検索結果一覧 |
| SCR-005 | 受診者詳細 | 基本情報・受診履歴 |
| SCR-006 | 検査結果表示 | カテゴリ別検査結果 |
| SCR-007 | 新規受診登録 | ステップ形式の登録 |
| SCR-008 | 検査結果入力 | リアルタイム判定付き入力 |

## clasp開発フロー

### 前提条件

- Node.js インストール済み
- clasp インストール済み (`npm install -g @google/clasp`)
- clasp ログイン済み (`clasp login`)

### コード編集後のデプロイ

**重要**: コードを編集したら必ずデプロイを実行してください。

```bash
# 1. プロジェクトディレクトリに移動
cd /path/to/gas_integrated

# 2. GASにプッシュ（デプロイ）
clasp push

# 3. 確認（任意）
clasp open  # GASエディタを開く
```

### GitHub Actions自動デプロイ

mainブランチへのpushで自動デプロイされます。

```bash
# コード編集後
git add -A
git commit -m "変更内容"
git push origin main
# → GitHub Actionsが自動でclasp pushを実行
```

### デプロイが必要なタイミング

| 変更内容 | デプロイ必要 |
|---------|-------------|
| .gsファイルの編集 | ✅ 必要 |
| ui/*.htmlファイルの編集 | ✅ 必要 |
| README.mdの編集 | ❌ 不要 |
| .clasp.jsonの編集 | ✅ 必要 |

### トラブルシューティング

```bash
# デプロイ状況確認
clasp deployments

# ログ確認
clasp logs

# 強制プッシュ（競合時）
clasp push --force
```

## 更新履歴

| 日付 | バージョン | 内容 |
|------|-----------|------|
| 2025/12/18 | 2.1.0 | clasp開発フロー追加、ナビゲーション修正 |
| 2025/12/14 | 2.0.0 | Phase 2 UI開発完了 |
| 2025/12/14 | 1.0.0 | Phase 1 初版作成 |
