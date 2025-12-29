# 健診結果入力システム 実装ワークフロー

## 概要

設計書 `GAS_SYSTEM_DESIGN.md` に基づく実装手順と作業項目

---

## Phase 1: スプレッドシート・GAS構築（目安: 1週間）

### 1.1 スプレッドシート作成 ⏳

#### Task 1.1.1: ベーススプレッドシート作成
- [ ] Google Driveで新規スプレッドシート作成
- [ ] ファイル名: `健診結果入力システム_マスタ`
- [ ] 共有設定: 関係者に編集権限付与

#### Task 1.1.2: データシート作成（7シート）
```
作成するシート:
├── 受診者マスタ（16列）
├── 身体測定（23列）
├── 血液検査（28列）
├── 所見（16列）
├── 判定マスタ（14列）
├── 所見テンプレート（4列）
└── 設定（システム設定用）
```

- [ ] 受診者マスタシート作成（列A〜P）
- [ ] 身体測定シート作成（列A〜W）
- [ ] 血液検査シート作成（列A〜AB）
- [ ] 所見シート作成（列A〜P）
- [ ] 判定マスタシート作成（列A〜N）
- [ ] 所見テンプレートシート作成（列A〜D）
- [ ] 設定シート作成

#### Task 1.1.3: 出力用テンプレートシート作成（5シート）
- [ ] template.xlsmを開く
- [ ] 1ページ〜5ページのレイアウトをスプレッドシートに再現
- [ ] セル結合・書式設定を維持
- [ ] データ参照用のセル位置をマッピング

#### Task 1.1.4: マスタデータ投入
- [ ] 判定マスタ: template.xlsmの判定マスタをインポート
- [ ] 所見テンプレート: 初期コメントを登録
- [ ] 設定: フォルダID等の初期設定

---

### 1.2 GAS開発 ⏳

#### Task 1.2.1: プロジェクト設定
- [ ] スプレッドシートからApps Scriptを開く
- [ ] プロジェクト名: `健診結果入力システム`
- [ ] タイムゾーン: Asia/Tokyo

#### Task 1.2.2: ユーティリティモジュール（utils.gs）
```javascript
// 実装する関数
- getSpreadsheet(): Spreadsheet
- getSheet(sheetName): Sheet
- getFolderById(folderId): Folder
- formatDate(date): string
- generatePatientId(): string
- sendNotification(subject, body): void
```
- [ ] utils.gs ファイル作成
- [ ] 基本ユーティリティ関数実装
- [ ] テスト実行

#### Task 1.2.3: CSV解析モジュール（csvParser.gs）
```javascript
// 実装する関数
- parseCSV(fileId): PatientData[]
- detectEncoding(blob): string
- extractPatientInfo(rows): PatientInfo
- extractTestResults(rows): TestResult[]
- findNewCsvFiles(): File[]
- markCsvAsProcessed(file): void
```
- [ ] csvParser.gs ファイル作成
- [ ] BML CSVフォーマット解析実装
- [ ] エンコーディング自動検出（Shift_JIS/UTF-8）
- [ ] 既存Pythonロジック移植
- [ ] テストCSVで動作確認

#### Task 1.2.4: 判定エンジンモジュール（judgmentEngine.gs）
```javascript
// 実装する関数
- judge(itemId, value, gender): string
- getJudgmentCriteria(itemId, gender): Criteria
- calculateOverallJudgment(judgments): string
- getCategoryJudgments(patientId): CategoryJudgments
- updateAllJudgments(patientId): void
```
- [ ] judgmentEngine.gs ファイル作成
- [ ] 判定ロジック実装（A/B/C/D）
- [ ] 性別依存判定対応
- [ ] 総合判定計算
- [ ] 既存Python JudgmentEngineロジック移植
- [ ] テストケースで検証

#### Task 1.2.5: 所見生成モジュール（findingsGenerator.gs）
```javascript
// 実装する関数
- generateFindings(patientId): Findings
- generateCategoryFindings(category, judgments): string
- combineFindings(categoryFindings): string
- getFindingsTemplate(category, judgment): string
- updateFindings(patientId): void
```
- [ ] findingsGenerator.gs ファイル作成
- [ ] カテゴリ別所見生成
- [ ] テンプレート参照ロジック
- [ ] 総合所見結合
- [ ] 動作確認

#### Task 1.2.6: Excel出力モジュール（excelExporter.gs）
```javascript
// 実装する関数
- exportToExcel(patientId): string
- fillTemplateSheet(patientData): void
- copyTemplateSheets(): Spreadsheet
- convertToExcel(spreadsheet, fileName): File
- saveToOutputFolder(file): string
```
- [ ] excelExporter.gs ファイル作成
- [ ] テンプレートシートへのデータ転記
- [ ] Spreadsheet → Excel変換
- [ ] 出力フォルダへの保存
- [ ] ファイル命名規則（result_YYYYMMDD_依頼ID.xlsx）

#### Task 1.2.7: メインモジュール（main.gs）
```javascript
// 実装する関数
- onCsvUploaded(): void（トリガー）
- processPatient(patientId): void
- processAll(): void
- exportPatientToExcel(patientId): string
- setupTriggers(): void
- removeTriggers(): void
```
- [ ] main.gs ファイル作成
- [ ] CSV監視トリガー実装
- [ ] 一連の処理フロー結合
- [ ] エラーハンドリング
- [ ] 通知機能（処理完了/エラー時）

#### Task 1.2.8: トリガー設定
- [ ] CSV監視トリガー（毎時）設定
- [ ] 日次アラートトリガー（8:00）設定
- [ ] トリガーのテスト実行

---

### 1.3 単体テスト ⏳

#### Task 1.3.1: CSV解析テスト
- [ ] テスト用CSVを01_csv/に配置
- [ ] parseCSV()実行
- [ ] 解析結果をログ確認
- [ ] エラーケーステスト

#### Task 1.3.2: 判定エンジンテスト
- [ ] 各検査項目の判定テスト
- [ ] 境界値テスト（A/B/C/D境界）
- [ ] 性別依存項目テスト
- [ ] 総合判定テスト

#### Task 1.3.3: 所見生成テスト
- [ ] 判定B以下の項目がある患者データ
- [ ] 各カテゴリの所見生成確認
- [ ] 総合所見フォーマット確認

#### Task 1.3.4: Excel出力テスト
- [ ] 単一患者の出力テスト
- [ ] レイアウト確認
- [ ] 複数患者の連続出力

---

## Phase 2: AppSheet構築（目安: 3日）

### 2.1 AppSheet基本設定 ⏳

#### Task 2.1.1: アプリ作成
- [ ] AppSheetでアプリ新規作成
- [ ] スプレッドシート接続
- [ ] データソース設定

#### Task 2.1.2: テーブル設定
- [ ] 受診者マスタテーブル設定
- [ ] 身体測定テーブル設定
- [ ] 血液検査テーブル設定
- [ ] 所見テーブル設定
- [ ] リレーションシップ設定（受診ID連携）

---

### 2.2 画面作成 ⏳

#### Task 2.2.1: ホーム画面
- [ ] ダッシュボードビュー作成
- [ ] ステータス別件数表示
- [ ] クイックアクションボタン配置

#### Task 2.2.2: 受診者一覧画面
- [ ] テーブルビュー作成
- [ ] 検索・フィルタ設定
- [ ] ステータスアイコン表示

#### Task 2.2.3: 受診者詳細画面
- [ ] フォームビュー作成
- [ ] 基本情報入力フィールド
- [ ] タブ切り替え設定

#### Task 2.2.4: 身体測定画面
- [ ] 入力フォーム作成
- [ ] 数値入力バリデーション
- [ ] BMI自動計算表示

#### Task 2.2.5: 検査結果画面
- [ ] 表形式表示
- [ ] 判定表示（色分け）
- [ ] 編集可能フィールド設定

#### Task 2.2.6: 所見入力画面
- [ ] カテゴリ別表示
- [ ] テキストエリア編集
- [ ] 自動生成所見の確認

#### Task 2.2.7: Excel出力画面
- [ ] 出力ボタン配置
- [ ] 出力履歴表示
- [ ] GAS関数呼び出し設定

---

### 2.3 アクション設定 ⏳

#### Task 2.3.1: ステータス変更アクション
- [ ] 入力中→確認待ち
- [ ] 確認待ち→完了
- [ ] 差戻し機能

#### Task 2.3.2: Excel出力アクション
- [ ] ボタンからGAS呼び出し
- [ ] 出力完了通知
- [ ] 出力日時記録

#### Task 2.3.3: データ更新アクション
- [ ] 保存時の判定自動更新
- [ ] 所見自動再生成

---

### 2.4 UIカスタマイズ ⏳

#### Task 2.4.1: 外観設定
- [ ] アプリアイコン設定
- [ ] カラーテーマ設定
- [ ] ロゴ追加

#### Task 2.4.2: ナビゲーション設定
- [ ] メニュー構成
- [ ] 画面遷移フロー

---

## Phase 3: 並行運用・検証（目安: 2週間）

### 3.1 検証準備 ⏳

#### Task 3.1.1: テストデータ準備
- [ ] 過去のCSVデータを準備（5〜10件）
- [ ] 期待結果を手動計算で用意

#### Task 3.1.2: 検証環境構築
- [ ] 本番と同じフォルダ構成
- [ ] テスト用スプレッドシート複製

---

### 3.2 機能検証 ⏳

#### Task 3.2.1: CSV取込検証
- [ ] CSV配置→自動取込確認
- [ ] データ正確性検証
- [ ] エラーCSV処理確認

#### Task 3.2.2: 判定検証
- [ ] 判定結果を既存Pythonと比較
- [ ] 差異がある場合は原因調査・修正
- [ ] 性別依存項目の検証

#### Task 3.2.3: 所見生成検証
- [ ] 自動生成所見の内容確認
- [ ] カテゴリ分類の正確性
- [ ] テンプレート適用の確認

#### Task 3.2.4: Excel出力検証
- [ ] 出力ファイルのレイアウト確認
- [ ] データ転記位置の確認
- [ ] 既存template.xlsmとの比較

#### Task 3.2.5: AppSheet操作検証
- [ ] 各画面の操作テスト
- [ ] 入力バリデーション確認
- [ ] アクション動作確認

---

### 3.3 並行運用 ⏳

#### Task 3.3.1: 運用開始
- [ ] 実際の業務データで処理開始
- [ ] 既存Python処理と並行実行
- [ ] 結果比較・差異確認

#### Task 3.3.2: 問題対応
- [ ] 発生した問題の記録
- [ ] 修正対応
- [ ] 再検証

#### Task 3.3.3: 運用手順確定
- [ ] 操作マニュアル作成
- [ ] FAQ作成
- [ ] 担当者への説明

---

## Phase 4: 本番移行（目安: 1日）

### 4.1 移行準備 ⏳

#### Task 4.1.1: 最終確認
- [ ] 全機能の動作確認
- [ ] 権限設定確認
- [ ] バックアップ取得

#### Task 4.1.2: 切り替え準備
- [ ] Python処理の停止手順確認
- [ ] フォルダ構成の最終確認

---

### 4.2 本番切り替え ⏳

#### Task 4.2.1: 切り替え実行
- [ ] GASトリガー有効化
- [ ] Python処理停止
- [ ] 運用開始宣言

#### Task 4.2.2: 初期運用監視
- [ ] 最初の数件を重点監視
- [ ] 問題発生時の即時対応
- [ ] 運用ログ記録

---

## 成果物一覧

| 成果物 | 形式 | 配置場所 |
|--------|------|----------|
| スプレッドシート（マスタ） | Google Spreadsheet | Google Drive |
| GASプロジェクト | Google Apps Script | スプレッドシート紐付け |
| AppSheetアプリ | AppSheet | Google Workspace |
| 操作マニュアル | Google Docs/PDF | 共有フォルダ |
| 設計書 | Markdown | 設計書_設定ファイル/ |

---

## リスクと対策

| リスク | 影響度 | 対策 |
|--------|--------|------|
| GASの実行時間制限（6分） | 中 | バッチ処理分割、進捗保存 |
| CSV文字化け | 低 | エンコーディング自動検出 |
| Excel変換のレイアウト崩れ | 低 | 事前検証で確認済み |
| AppSheet操作ミス | 中 | バリデーション、確認ダイアログ |

---

## 次のステップ

Phase 1の実装を開始するには:

1. **スプレッドシート作成から開始**
   - Google Driveで新規スプレッドシート作成
   - 各シートの列定義を設計書通りに設定

2. **または GAS開発から開始**
   - 既存Pythonコードを参考にGAS実装
   - csvParser.gs から開始推奨

**推奨**: スプレッドシート構造を先に作成し、GAS開発時にすぐテストできる状態にする

---

**文書情報**
- 作成日: 2025-11-29
- 対応設計書: GAS_SYSTEM_DESIGN.md v1.0
