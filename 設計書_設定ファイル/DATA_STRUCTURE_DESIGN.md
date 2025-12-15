# データ構造設計書

## 1. 概要

### 1.1 目的
本設計書は、健診結果入力システムにおけるデータ構造を定義し、以下を実現する：
- iD-Heart帳票フローを参考とした一貫性あるデータモデル
- 人間ドック学会2025年度判定基準への準拠
- 人間ドック・定期健康診断（労安法）両対応
- PDF/Excel出力に最適化された構造

### 1.2 対象システム
| 項目 | 内容 |
|------|------|
| 健診種別 | 人間ドック、定期健康診断（労安法） |
| 入力方式 | BML CSV取込 + AppSheet手入力 |
| 出力形式 | PDF、Excel |
| データ基盤 | Google Sheets + Apps Script |

### 1.3 設計方針
- **iD-Heart参照**: フロー全体（予約→受診→結果入力→報告書→フォロー）のデータ構造を参考
- **既存システム統合**: 現行GASシステム（spreadsheet_setup.gs）との互換性維持
- **拡張性**: 将来の機能追加に対応できる柔軟な構造

---

## 2. エンティティ関連図 (ER図)

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   事業所マスタ   │     │   受診者マスタ   │     │   コースマスタ   │
│  (Organizations) │     │   (Patients)    │     │    (Courses)    │
├─────────────────┤     ├─────────────────┤     ├─────────────────┤
│ PK: 事業所ID     │     │ PK: 受診者ID    │     │ PK: コースID    │
│    事業所名      │◄────│ FK: 事業所ID    │     │    コース名      │
│    郵便番号      │     │    氏名         │────►│    健診種別      │
│    住所          │     │    カナ         │     │    基本料金      │
│    電話番号      │     │    生年月日     │     │    検査項目群    │
│    担当者名      │     │    性別         │     └─────────────────┘
└─────────────────┘     │    連絡先       │              │
                        └─────────────────┘              │
                                 │                        │
                                 │ 1:N                    │
                                 ▼                        ▼
                        ┌─────────────────┐     ┌─────────────────┐
                        │   受診記録       │     │   検査項目マスタ  │
                        │  (Visits)       │     │  (TestItems)    │
                        ├─────────────────┤     ├─────────────────┤
                        │ PK: 受診ID      │     │ PK: 項目ID      │
                        │ FK: 受診者ID    │     │    項目名        │
                        │ FK: コースID    │     │    BMLコード     │
                        │    受診日       │     │    カテゴリ      │
                        │    ステータス   │     │    単位          │
                        │    総合判定     │     │    基準値       │
                        └─────────────────┘     └─────────────────┘
                                 │                        │
           ┌─────────────────────┼─────────────────────────┤
           │                     │                        │
           ▼                     ▼                        ▼
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   身体測定       │     │   検査結果       │     │   判定マスタ     │
│ (Measurements)  │     │ (TestResults)   │     │ (JudgmentRules) │
├─────────────────┤     ├─────────────────┤     ├─────────────────┤
│ PK: 測定ID      │     │ PK: 結果ID      │     │ PK: 判定ID      │
│ FK: 受診ID      │     │ FK: 受診ID      │     │ FK: 項目ID      │
│    身長・体重   │     │ FK: 項目ID      │     │    性別条件      │
│    血圧         │     │    検査値       │     │    A/B/C/D範囲   │
│    視力・聴力   │     │    判定         │     └─────────────────┘
└─────────────────┘     │    フラグ       │
                        └─────────────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │   所見データ     │
                        │  (Findings)     │
                        ├─────────────────┤
                        │ PK: 所見ID      │
                        │ FK: 受診ID      │
                        │    カテゴリ別所見 │
                        │    総合所見     │
                        │    医師コメント  │
                        └─────────────────┘
```

---

## 3. マスタデータ定義

### 3.1 事業所マスタ (M_Organizations)
団体健診時の企業・団体情報を管理

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 事業所ID | VARCHAR(10) | NO | PK | 自動採番 (ORG-NNNNN) |
| 事業所名 | VARCHAR(100) | NO | | 企業名・団体名 |
| 事業所名カナ | VARCHAR(200) | YES | | ソート・検索用 |
| 郵便番号 | VARCHAR(8) | YES | | NNN-NNNN形式 |
| 住所1 | VARCHAR(100) | YES | | 都道府県・市区町村 |
| 住所2 | VARCHAR(100) | YES | | 番地・建物名 |
| 電話番号 | VARCHAR(15) | YES | | |
| FAX番号 | VARCHAR(15) | YES | | |
| 担当者名 | VARCHAR(50) | YES | | 健診窓口担当者 |
| 担当者メール | VARCHAR(100) | YES | | |
| 請求先区分 | ENUM | YES | | 個人/団体一括/代行機関 |
| 契約コースID | VARCHAR(10)[] | YES | FK | 利用可能コース |
| 備考 | TEXT | YES | | |
| 作成日時 | DATETIME | NO | | |
| 更新日時 | DATETIME | NO | | |

### 3.2 受診コースマスタ (M_Courses)
人間ドック・定期健診のコース定義

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| コースID | VARCHAR(10) | NO | PK | 自動採番 (CRS-NNNNN) |
| コース名 | VARCHAR(50) | NO | | 表示名 |
| コース略称 | VARCHAR(20) | YES | | 帳票用短縮名 |
| 健診種別 | ENUM | NO | | 人間ドック/定期健診/雇入健診/特殊健診 |
| 基本料金 | INTEGER | YES | | 税抜金額 |
| 所要時間 | INTEGER | YES | | 分単位 |
| 含有検査項目 | JSON | NO | | 検査項目ID配列 |
| 有効フラグ | BOOLEAN | NO | | |
| 表示順 | INTEGER | YES | | |
| 作成日時 | DATETIME | NO | | |

**健診種別ENUM定義**:
```
人間ドック: DOCK
定期健康診断: REGULAR
雇入時健康診断: EMPLOYMENT
特殊健康診断: SPECIAL
```

### 3.3 検査項目マスタ (M_TestItems)
検査項目の基本情報

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 項目ID | VARCHAR(10) | NO | PK | 内部ID (TI-NNNNN) |
| 項目名 | VARCHAR(50) | NO | | 標準名称 |
| 項目名略称 | VARCHAR(20) | YES | | 帳票用 |
| BMLコード | VARCHAR(10) | YES | | BML検査コード |
| カテゴリ | ENUM | NO | | 検査カテゴリ |
| サブカテゴリ | VARCHAR(30) | YES | | 詳細分類 |
| 単位 | VARCHAR(20) | YES | | mg/dL, % 等 |
| データ型 | ENUM | NO | | 数値/文字列/選択肢 |
| 小数桁数 | INTEGER | YES | | 数値の場合 |
| 入力方式 | ENUM | NO | | CSV/手入力/機器連携 |
| 必須フラグ | BOOLEAN | NO | | |
| 表示順 | INTEGER | YES | | |
| 有効フラグ | BOOLEAN | NO | | |

**カテゴリENUM定義** (iD-Heart参照):
```
身体計測: BODY
循環器系: CIRCULATORY
消化器系: DIGESTIVE
代謝系_糖: METABOLISM_GLUCOSE
代謝系_脂質: METABOLISM_LIPID
腎機能: RENAL
血液系: HEMATOLOGY
尿検査: URINALYSIS
眼科: OPHTHALMOLOGY
聴力: HEARING
画像検査: IMAGING
その他: OTHER
```

### 3.4 判定基準マスタ (M_JudgmentRules)
人間ドック学会2025年度基準に準拠

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 判定ID | VARCHAR(15) | NO | PK | 項目ID + 性別 |
| 項目ID | VARCHAR(10) | NO | FK | 検査項目ID |
| 性別条件 | ENUM | NO | | 共通/男性/女性 |
| 基準値下限 | DECIMAL | YES | | 正常範囲下限 |
| 基準値上限 | DECIMAL | YES | | 正常範囲上限 |
| A判定下限 | DECIMAL | YES | | 異常なし |
| A判定上限 | DECIMAL | YES | | |
| B判定下限 | DECIMAL | YES | | 軽度異常 |
| B判定上限 | DECIMAL | YES | | |
| C判定下限 | DECIMAL | YES | | 要経過観察/要精検 |
| C判定上限 | DECIMAL | YES | | |
| D判定下限 | DECIMAL | YES | | 要治療/要精検 |
| D判定上限 | DECIMAL | YES | | |
| 判定ロジック | ENUM | NO | | 範囲/閾値/複合 |
| 年度 | VARCHAR(4) | NO | | 基準年度 (2025等) |
| 有効フラグ | BOOLEAN | NO | | |

**判定基準例** (人間ドック学会2025年度):
| 項目 | A判定 | B判定 | C判定 | D判定 |
|------|-------|-------|-------|-------|
| AST | ≤30 | 31-35 | 36-50 | >50 |
| ALT | ≤30 | 31-40 | 41-50 | >50 |
| γ-GTP | ≤50 | 51-80 | 81-100 | >100 |
| HbA1c | ≤5.5 | 5.6-5.9 | 6.0-6.4 | ≥6.5 |
| LDL-C | 60-119 | 120-139 | 140-179 | ≥180 |
| HDL-C | ≥40 | - | 30-39 | <30 |
| 空腹時血糖 | ≤99 | 100-109 | 110-125 | ≥126 |
| eGFR | ≥60 | 45-59 | 30-44 | <30 |

### 3.5 所見テンプレートマスタ (M_FindingTemplates)
カテゴリ・判定別の所見文テンプレート

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| テンプレートID | VARCHAR(15) | NO | PK | 自動採番 |
| カテゴリ | ENUM | NO | | 検査カテゴリ |
| 判定 | ENUM | NO | | B/C/D |
| 項目パターン | VARCHAR(50) | YES | | 特定項目用 |
| コメント | TEXT | NO | | 所見文テンプレート |
| 指導コメント | TEXT | YES | | 生活指導文 |
| 優先度 | INTEGER | YES | | 表示順 |
| 有効フラグ | BOOLEAN | NO | | |

---

## 4. トランザクションデータ定義

### 4.1 受診者データ (T_Patients)
受診者の基本情報（個人マスタ兼用）

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 受診者ID | VARCHAR(15) | NO | PK | 自動採番 (P-YYYYMMDD-NNN) |
| 事業所ID | VARCHAR(10) | YES | FK | 団体健診時 |
| 氏名 | VARCHAR(50) | NO | | |
| 氏名カナ | VARCHAR(100) | NO | | 全角カタカナ |
| 生年月日 | DATE | NO | | |
| 性別 | ENUM | NO | | 男性/女性 |
| 郵便番号 | VARCHAR(8) | YES | | |
| 住所1 | VARCHAR(100) | YES | | |
| 住所2 | VARCHAR(100) | YES | | |
| 電話番号 | VARCHAR(15) | YES | | |
| メールアドレス | VARCHAR(100) | YES | | |
| 保険証番号 | VARCHAR(20) | YES | | |
| 所属部署 | VARCHAR(50) | YES | | 団体健診時 |
| 備考 | TEXT | YES | | |
| 作成日時 | DATETIME | NO | | |
| 更新日時 | DATETIME | NO | | |

### 4.2 受診記録 (T_Visits)
受診1回ごとの記録（中心テーブル）

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 受診ID | VARCHAR(20) | NO | PK | YYYYMMDD-NNN形式 |
| 受診者ID | VARCHAR(15) | NO | FK | |
| コースID | VARCHAR(10) | NO | FK | |
| 受診日 | DATE | NO | | |
| 受診時刻 | TIME | YES | | |
| ステータス | ENUM | NO | | 処理状態 |
| 事業所ID | VARCHAR(10) | YES | FK | 団体健診時 |
| 予約ID | VARCHAR(20) | YES | | 予約システム連携用 |
| 総合判定 | ENUM | YES | | A/B/C/D/E |
| 医師ID | VARCHAR(10) | YES | | 担当医師 |
| CSV取込日時 | DATETIME | YES | | |
| 最終更新日時 | DATETIME | NO | | |
| 出力日時 | DATETIME | YES | | 帳票出力時 |
| 備考 | TEXT | YES | | |

**ステータスENUM定義** (iD-Heartフロー参照):
```
予約済: RESERVED
受付済: CHECKED_IN
検査中: IN_PROGRESS
入力中: DATA_ENTRY
確認待ち: PENDING_REVIEW
完了: COMPLETED
報告済: REPORTED
```

### 4.3 身体測定データ (T_Measurements)
手入力・機器連携の身体測定結果

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 測定ID | VARCHAR(25) | NO | PK | 受診ID + 連番 |
| 受診ID | VARCHAR(20) | NO | FK | |
| 身長 | DECIMAL(5,1) | YES | | cm |
| 体重 | DECIMAL(5,1) | YES | | kg |
| 標準体重 | DECIMAL(5,1) | YES | | 自動計算 |
| BMI | DECIMAL(4,1) | YES | | 自動計算 |
| BMI判定 | ENUM | YES | | A/B/C/D |
| 体脂肪率 | DECIMAL(4,1) | YES | | % |
| 腹囲 | DECIMAL(5,1) | YES | | cm |
| 腹囲判定 | ENUM | YES | | |
| 血圧_収縮期_1 | INTEGER | YES | | mmHg |
| 血圧_拡張期_1 | INTEGER | YES | | mmHg |
| 血圧_収縮期_2 | INTEGER | YES | | |
| 血圧_拡張期_2 | INTEGER | YES | | |
| 血圧判定 | ENUM | YES | | |
| 脈拍 | INTEGER | YES | | /分 |
| 視力_裸眼_右 | DECIMAL(3,1) | YES | | |
| 視力_裸眼_左 | DECIMAL(3,1) | YES | | |
| 視力_矯正_右 | DECIMAL(3,1) | YES | | |
| 視力_矯正_左 | DECIMAL(3,1) | YES | | |
| 視力判定 | ENUM | YES | | |
| 聴力_右_1000Hz | ENUM | YES | | 所見あり/なし |
| 聴力_左_1000Hz | ENUM | YES | | |
| 聴力_右_4000Hz | ENUM | YES | | |
| 聴力_左_4000Hz | ENUM | YES | | |
| 聴力判定 | ENUM | YES | | |
| 眼圧_右 | DECIMAL(4,1) | YES | | mmHg |
| 眼圧_左 | DECIMAL(4,1) | YES | | |
| 眼底_右 | VARCHAR(100) | YES | | 所見文 |
| 眼底_左 | VARCHAR(100) | YES | | |
| 入力日時 | DATETIME | NO | | |
| 更新日時 | DATETIME | NO | | |

### 4.4 検査結果データ (T_TestResults)
血液検査等のCSV取込・手入力結果

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 結果ID | VARCHAR(30) | NO | PK | 受診ID + 項目ID |
| 受診ID | VARCHAR(20) | NO | FK | |
| 項目ID | VARCHAR(10) | NO | FK | |
| BMLコード | VARCHAR(10) | YES | | CSV取込時 |
| 検査値 | VARCHAR(50) | YES | | 数値または文字列 |
| 検査値数値 | DECIMAL(10,3) | YES | | 数値変換後 |
| 単位 | VARCHAR(20) | YES | | |
| 判定 | ENUM | YES | | A/B/C/D/E |
| 異常フラグ | ENUM | YES | | H(高)/L(低)/空白 |
| コメント | VARCHAR(200) | YES | | 検査所コメント |
| 入力方式 | ENUM | NO | | CSV/手入力/機器 |
| 入力日時 | DATETIME | NO | | |
| 更新日時 | DATETIME | NO | | |

### 4.5 所見データ (T_Findings)
カテゴリ別・総合所見

| 列名 | データ型 | NULL | PK/FK | 説明 |
|------|----------|------|-------|------|
| 所見ID | VARCHAR(25) | NO | PK | 受診ID + 連番 |
| 受診ID | VARCHAR(20) | NO | FK | |
| 既往歴 | TEXT | YES | | 手入力 |
| 自覚症状 | TEXT | YES | | |
| 他覚症状 | TEXT | YES | | |
| 所見_循環器系 | TEXT | YES | | 自動生成+編集 |
| 所見_消化器系 | TEXT | YES | | |
| 所見_代謝系_糖 | TEXT | YES | | |
| 所見_代謝系_脂質 | TEXT | YES | | |
| 所見_腎機能 | TEXT | YES | | |
| 所見_血液系 | TEXT | YES | | |
| 所見_尿検査 | TEXT | YES | | |
| 所見_その他 | TEXT | YES | | |
| 総合所見 | TEXT | YES | | 自動結合+編集 |
| 医師コメント | TEXT | YES | | 医師追記 |
| 指導コメント | TEXT | YES | | 生活指導 |
| 心電図_今回 | TEXT | YES | | |
| 心電図_前回 | TEXT | YES | | 前回参照 |
| 胸部X線_今回 | TEXT | YES | | |
| 胸部X線_前回 | TEXT | YES | | |
| 腹部超音波_今回 | TEXT | YES | | |
| 腹部超音波_前回 | TEXT | YES | | |
| 生成日時 | DATETIME | YES | | 自動生成時 |
| 編集日時 | DATETIME | YES | | 手動編集時 |

---

## 5. iD-Heart帳票フロー対応マッピング

### 5.1 フロー別データ利用

| iD-Heartフロー段階 | 対応帳票例 | 主要データソース |
|-------------------|-----------|-----------------|
| 予約・案内 | コース日程表、受診案内 | M_Courses, T_Visits(予約) |
| 受診受付 | 受診票、問診票 | T_Patients, T_Visits |
| 検査実施 | 検査チケット | T_Visits, M_TestItems |
| 結果入力 | (内部処理) | T_TestResults, T_Measurements |
| 結果報告 | 健診結果報告書 | 全テーブル結合 |
| 団体集計 | 一覧表、集計表 | T_Visits + 事業所GROUP BY |
| 請求 | 請求書、明細 | T_Visits, M_Courses |
| 事後フォロー | 精検依頼書、追跡リスト | T_Findings(C/D判定) |

### 5.2 優先度Aデータ構造 (受診者向け結果報告書)

```
結果報告書データビュー:
├─ T_Visits (受診基本情報)
│   ├─ 受診ID, 受診日, 総合判定
│   └─ → T_Patients (氏名, 生年月日, 性別)
│
├─ T_Measurements (身体測定)
│   └─ 身長, 体重, BMI, 血圧, 視力, 聴力 + 各判定
│
├─ T_TestResults (検査結果)
│   ├─ 血液検査 (WBC, RBC, Hb, 肝機能, 脂質, 血糖...)
│   ├─ 尿検査 (蛋白, 糖, 潜血...)
│   └─ 各項目判定 (A/B/C/D)
│
├─ T_Findings (所見)
│   ├─ カテゴリ別所見
│   ├─ 総合所見
│   └─ 医師コメント
│
└─ 前回比較データ (T_Visits JOIN 過去分)
```

### 5.3 優先度A (団体向け一覧・集計)

```
団体一覧データビュー:
├─ T_Visits WHERE 事業所ID = X
│   └─ GROUP BY 受診日 (日付別)
│
├─ 横展開 (8名/ページ等):
│   ├─ 受診者名 x 検査項目マトリクス
│   └─ 判定集計 (A:n名, B:n名, C:n名, D:n名)
│
└─ 集計サマリ:
    ├─ 受診者数 (総数, 男女別)
    ├─ 判定分布 (項目別)
    └─ 有所見率 (カテゴリ別)
```

---

## 6. BMLコードマッピング

### 6.1 コード対応表 (既存システム互換)

| BMLコード | 項目ID | 項目名 | カテゴリ |
|-----------|--------|--------|---------|
| 0000301 | TI-00001 | WBC | 血液系 |
| 0000302 | TI-00002 | RBC | 血液系 |
| 0000303 | TI-00003 | Hb | 血液系 |
| 0000304 | TI-00004 | Ht | 血液系 |
| 0000305 | TI-00005 | MCV | 血液系 |
| 0000306 | TI-00006 | MCH | 血液系 |
| 0000307 | TI-00007 | MCHC | 血液系 |
| 0000308 | TI-00008 | PLT | 血液系 |
| 0000401 | TI-00010 | TP | 消化器系 |
| 0000402 | TI-00011 | ALB | 消化器系 |
| 0000407 | TI-00012 | UA | その他 |
| 0000410 | TI-00013 | LDL-C | 代謝系_脂質 |
| 0000413 | TI-00014 | Cr | 腎機能 |
| 0000450 | TI-00015 | TC | 代謝系_脂質 |
| 0000454 | TI-00016 | TG | 代謝系_脂質 |
| 0000460 | TI-00017 | HDL-C | 代謝系_脂質 |
| 0000472 | TI-00018 | T-Bil | 消化器系 |
| 0000481 | TI-00019 | AST | 消化器系 |
| 0000482 | TI-00020 | ALT | 消化器系 |
| 0000484 | TI-00021 | γ-GTP | 消化器系 |
| 0000503 | TI-00022 | FBS | 代謝系_糖 |
| 0000658 | TI-00023 | CRP | その他 |
| 0002696 | TI-00024 | eGFR | 腎機能 |
| 0003317 | TI-00025 | HbA1c | 代謝系_糖 |

### 6.2 CSVフォーマット (BML)

```
[ヘッダ部]
施設コード,依頼ID,カナ,受診日,受診時刻,検査セット,保険番号,,,区分,,医師,性別,生年月日,...

[データ部]
検査コード,検査値,フラグ,コメント, (4項目×検査数)
```

---

## 7. 既存GASシステムとの互換性

### 7.1 シート名マッピング

| 既存シート名 | 新データ構造 | 備考 |
|-------------|-------------|------|
| 受診者マスタ | T_Patients + T_Visits | 分離して正規化 |
| 身体測定 | T_Measurements | 互換維持 |
| 血液検査 | T_TestResults | 項目ID方式へ |
| 所見 | T_Findings | 互換維持 |
| 判定マスタ | M_JudgmentRules | 構造拡張 |
| 所見テンプレート | M_FindingTemplates | 互換維持 |

### 7.2 移行時の注意点

1. **受診ID形式**: 既存 `YYYYMMDD-NNN` を維持
2. **BMLコード**: `utils.gs`の`CODE_TO_CRITERIA`と同期
3. **判定基準**: `judgmentEngine.gs`の`JUDGMENT_CRITERIA`と同期
4. **カテゴリ分類**: `CONFIG.CATEGORIES`と同期

---

## 8. 拡張予定

### 8.1 Phase 2 (定期健診対応)

追加テーブル:
- `M_LegalRequirements`: 労安法検査項目定義
- `T_LegalResults`: 法定健診結果

### 8.2 Phase 3 (事後フォロー)

追加テーブル:
- `T_FollowUp`: 精検依頼・追跡管理
- `T_Appointments`: 再検査予約

---

## 付録

### A. ENUM値一覧

```javascript
// 性別
GENDER = { MALE: '男性', FEMALE: '女性' }

// 判定
JUDGMENT = { A: 'A', B: 'B', C: 'C', D: 'D', E: 'E' }

// ステータス
STATUS = {
  RESERVED: '予約済',
  CHECKED_IN: '受付済',
  IN_PROGRESS: '検査中',
  DATA_ENTRY: '入力中',
  PENDING_REVIEW: '確認待ち',
  COMPLETED: '完了',
  REPORTED: '報告済'
}

// カテゴリ
CATEGORY = {
  BODY: '身体計測',
  CIRCULATORY: '循環器系',
  DIGESTIVE: '消化器系',
  METABOLISM_GLUCOSE: '代謝系_糖',
  METABOLISM_LIPID: '代謝系_脂質',
  RENAL: '腎機能',
  HEMATOLOGY: '血液系',
  URINALYSIS: '尿検査',
  OPHTHALMOLOGY: '眼科',
  HEARING: '聴力',
  IMAGING: '画像検査',
  OTHER: 'その他'
}
```

### B. 作成情報

- 作成日: 2025-12-02
- バージョン: 1.0
- 参照: iD-Heart帳票サンプル、既存GASシステム
- 判定基準: 人間ドック学会2025年度基準
