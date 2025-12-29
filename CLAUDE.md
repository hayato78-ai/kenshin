# CLAUDE.md - プロジェクト開発ガイドライン

## プロジェクト概要

- **アプリ名**: 健診結果管理システム
- **概要**: 健診結果の入力・判定・Excel出力を行うGAS + Pythonシステム
- **技術スタック**: GAS, Python, Google Sheets, Google Drive, Claude API

---

## 参照ルール（厳守）

| 対象 | 参照先 | 備考 |
|------|--------|------|
| 設計情報 | `docs/ARCHITECTURE.md` | 唯一の設計書 |
| 機能一覧 | `docs/FUNCTION_INVENTORY.md` | 移植判定の参照 |
| データ構造 | `設計書_設定ファイル/1220_new/DATA_STRUCTURE_DESIGN_v2.md` | DB操作時に参照 |
| セッションナレッジ | `設計書_設定ファイル/session/*/ai-guide.md` | 機能別の学び |
| 横断的ナレッジ | `docs/cross-cutting-concerns.md` | 複数機能に影響 |
| 旧設計書 | `docs/archive/` | **参照のみ・編集禁止** |

---

## 開発原則

### 1. デフォルト・標準設定を優先する

- ライブラリのデフォルト設定を尊重し、独自実装を避ける
- GAS標準APIを優先使用
- 理由: AIは公開情報量が多いほど精度が上がるため

### 2. 統一感を持たせる

- コードスタイル、命名規則、コミットメッセージを統一する
- 認知負荷を下げ、AI・人間両方の読み取りコストを削減する

### 3. 勝手な変更をしない

- 指示された範囲外のファイルを変更しない
- 既存のコードスタイルや設計方針を勝手に変えない
- 不明点があれば必ず確認を求める

### 4. セキュリティを意識する

- APIキーは設定シートまたはPropertiesServiceで管理
- 個人情報を含むログ出力は禁止
- Claude API呼び出し時のプロンプトインジェクション対策

---

## 環境前提

| 項目 | 内容 |
|------|------|
| 開発環境 | macOS |
| 利用環境 | Windows |
| パス設定 | `python/settings.yaml` で環境別パスを設定 |

### クロスプラットフォーム対応ルール

- **パスのハードコード禁止**
- `settings.yaml` で環境別のパスを設定
- 詳細: `docs/cross-cutting-concerns.md` の DP-001 を参照

---

## ディレクトリ構成

```
/
├── CLAUDE.md                    ← このファイル
├── docs/
│   ├── ARCHITECTURE.md          ← 唯一の設計書
│   ├── FUNCTION_INVENTORY.md    ← 機能棚卸し・移植判定
│   ├── cross-cutting-concerns.md ← 横断的ナレッジ
│   └── archive/                 ← 旧設計書（参照のみ）
├── gas/
│   ├── .clasp.json              ← clasp設定
│   ├── *.js                     ← GASソースコード
│   └── templates/               ← HTML/CSS/JS（ポータルUI）
│       ├── portal.html
│       ├── css.html
│       └── js.html
├── python/
│   ├── *.py                     ← Pythonソースコード
│   └── settings.yaml            ← 環境固有設定（.gitignore）
└── 設計書_設定ファイル/
    ├── 1220_new/
    │   └── DATA_STRUCTURE_DESIGN_v2.md ← データ構造定義
    └── session/
        └── {機能名}/ai-guide.md ← 機能別ナレッジ
```

---

## GAS開発フロー（必須）

### 編集の鉄則

```
1. ローカルで編集 → 2. clasp push → 3. 動作確認 → 4. clasp deploy
```

### 編集前の確認

```bash
git status && git pull
```

### 編集後の反映

```bash
# 1. GASに反映
clasp push

# 2. 動作確認（開発版URL）

# 3. 本番デプロイ
clasp deploy --description "変更内容"
```

### 絶対禁止

- GASエディタ（ブラウザ）での直接編集
- clasp push せずにローカル編集のみ
- 本番デプロイ前のテスト省略

---

## UI開発ルール

### 正式なUI構成（これ以外は作成禁止）

```
gas/templates/          ← 唯一のUIフォルダ
├── portal.html         ← メインUI（タブ形式）
├── css.html            ← スタイル
└── js.html             ← JavaScript
```

### 禁止事項

- `templates/` 以外へのHTML作成
- `portal.html` の複製・別バージョン作成
- `doGet()` 関数の無断変更
- スプレッドシートメニューへのUI機能追加

---

## コーディング規約

### GAS（JavaScript）

| 対象 | 規則 | 例 |
|------|------|-----|
| 関数 | camelCase | `getPatientData()` |
| 定数 | SCREAMING_SNAKE_CASE | `CONFIG.SHEETS.PATIENT` |
| クラス | PascalCase | `PatientManager` |
| ファイル | camelCase | `patientManager.js` |

### Python

| 対象 | 規則 | 例 |
|------|------|-----|
| 関数・変数 | snake_case | `get_patient_data()` |
| 定数 | SCREAMING_SNAKE_CASE | `MAX_RETRY_COUNT` |
| クラス | PascalCase | `HumanDockTranscriber` |
| ファイル | snake_case | `unified_transcriber.py` |

---

## Git運用

### ブランチ命名

```
feature/[機能名]       # 新機能
fix/[バグ内容]         # バグ修正
refactor/[対象]        # リファクタリング
```

### コミットメッセージ

```
[種別]([スコープ]): [変更内容の要約]

種別:
- feat: 新機能
- fix: バグ修正
- refactor: リファクタリング
- docs: ドキュメント
- style: フォーマット修正
- chore: 設定変更など

例:
feat(excel): Python連携Excel出力追加
fix(judgment): BMI判定基準修正
docs(arch): ARCHITECTURE.md更新
```

### コミット粒度

- **1作業項目 = 1コミット** を原則とする
- 作業項目が完了するたびに確認を求め、承認後にコミット

---

## 実装時のルール

### 確認を求めるタイミング

1. 1つの作業項目が完了したとき
2. 設計判断が必要なとき
3. 既存コードの大幅な変更が必要なとき
4. 新しいパッケージの追加が必要なとき
5. 想定外のエラーが発生したとき

### やってはいけないこと

- [ ] 指示にないファイルの変更
- [ ] 新規設計書の作成（ARCHITECTURE.mdを更新すること）
- [ ] パスのハードコード
- [ ] console.log / Logger.log の残留（デバッグ用除く）
- [ ] エラーハンドリングの省略
- [ ] docs/archive/ の編集

### 推奨事項

- [ ] 型安全性を意識（JSDoc / TypeHint）
- [ ] 早期リターンでネストを浅く
- [ ] 関数は単一責任に
- [ ] マジックナンバーは定数化
- [ ] セッションで学んだことは ai-guide.md に記録

---

## 移植対象機能

`docs/FUNCTION_INVENTORY.md` の移植判定を参照。

| 記号 | 意味 | 対応 |
|------|------|------|
| ○ | 移植する | 実装対象 |
| △ | 検討 | 要判断 |
| × | 不要 | 対象外 |

---

## Issue実装の進め方

### 1. 準備フェーズ

```
1. Issueの内容を確認
2. docs/ARCHITECTURE.md を読む
3. 関連する session/*/ai-guide.md を読む
4. 実装方針を提示して確認を求める
```

### 2. 実装フェーズ

```
1. 作業項目を1つずつ実装
2. 完了ごとに差分を提示して確認
3. 承認後にコミット
4. 次の作業項目へ
```

### 3. 完了フェーズ

```
1. 全作業項目の完了を報告
2. 動作確認の手順を提示
3. ai-guide.md に学びを記録
```

---

## セッション追跡

### ai-guide の配置

- **パス**: `設計書_設定ファイル/session/[機能名]/ai-guide.md`
- **例**: `設計書_設定ファイル/session/excel_output/ai-guide.md`

### セッション終了時

1. セッション内容を要約
2. ai-guide.md への追記内容を提案（ユーザー承認待ち）
3. 承認後、更新を実行

---

## 備考

- 不明点は推測せず質問する
- パフォーマンスより可読性を優先（最適化は後で）
- 設計書との乖離を発見したら ARCHITECTURE.md の「乖離記録」に追記
