# 帳票マスタ設計書

## 1. 概要

### 1.1 目的
本設計書は、健診結果入力システムにおける帳票出力機能を定義し、以下を実現する：
- iD-Heartを参考とした帳票体系の整備
- No-Code設計によるテンプレート管理
- PDF/Excel出力の統一的な仕組み

### 1.2 帳票優先度

| 優先度 | カテゴリ | 帳票例 |
|--------|---------|--------|
| **A** | 受診者向け結果報告書 | 健診結果報告書、人間ドック結果報告書 |
| **A** | 団体向け一覧・集計 | 8名リスト、事業所別集計表 |
| **B** | 請求・会計帳票 | 請求書、領収書、明細書 |
| **B** | 予約・案内帳票 | コース日程表、受診案内、問診票 |
| **C** | 事後フォロー帳票 | 精検依頼書、追跡管理リスト |

### 1.3 設計方針
- **テンプレート駆動**: Excelテンプレート + セルマッピングJSON
- **No-Code運用**: テンプレート変更はExcel編集のみ
- **マルチ出力**: 同一データから PDF/Excel/印刷 対応

---

## 2. 帳票マスタ構造

### 2.1 帳票定義テーブル (M_Forms)

| 列名 | データ型 | NULL | 説明 |
|------|----------|------|------|
| 帳票ID | VARCHAR(15) | NO | PK: FRM-NNNNN |
| 帳票名 | VARCHAR(50) | NO | 表示名 |
| 帳票カテゴリ | ENUM | NO | 結果報告/一覧集計/請求/案内/フォロー |
| 帳票種別 | ENUM | NO | 個人/団体/管理 |
| 健診種別 | ENUM[] | NO | 対応する健診種別 |
| テンプレートファイル | VARCHAR(100) | NO | Google Driveファイル名 |
| テンプレートID | VARCHAR(50) | YES | Google DriveファイルID |
| マッピングJSON | VARCHAR(100) | NO | セルマッピング定義ファイル |
| ページ数 | INTEGER | NO | 想定ページ数 |
| 出力形式 | ENUM[] | NO | PDF/Excel/印刷 |
| 用紙サイズ | ENUM | NO | A4/A3/B4/B5 |
| 印刷方向 | ENUM | NO | 縦/横 |
| 有効フラグ | BOOLEAN | NO | |
| 表示順 | INTEGER | YES | |
| 備考 | TEXT | YES | |

### 2.2 帳票カテゴリENUM

```javascript
FORM_CATEGORY = {
  RESULT_REPORT: '結果報告',      // 優先度A
  LIST_SUMMARY: '一覧集計',        // 優先度A
  BILLING: '請求会計',             // 優先度B
  GUIDANCE: '予約案内',            // 優先度B
  FOLLOW_UP: '事後フォロー'        // 優先度C
}
```

---

## 3. 帳票一覧 (iD-Heart参照)

### 3.1 優先度A: 受診者向け結果報告書

#### FRM-00001: 人間ドック結果報告書
```yaml
帳票ID: FRM-00001
帳票名: 人間ドック結果報告書
カテゴリ: 結果報告
種別: 個人
健診種別: [人間ドック]
テンプレート: template_dock_result.xlsx
マッピング: mapping_dock_result.json
ページ数: 5
出力形式: [PDF, Excel]
用紙: A4
印刷: 縦

構成:
  1ページ目: 基本情報、総合判定、身体測定
  2ページ目: 血液検査結果（肝機能、脂質、腎機能）
  3ページ目: 血液検査結果（血糖、血液学的検査）
  4ページ目: 尿検査、画像検査所見
  5ページ目: 総合所見、医師コメント、生活指導
```

#### FRM-00002: 定期健康診断結果報告書
```yaml
帳票ID: FRM-00002
帳票名: 定期健康診断結果報告書
カテゴリ: 結果報告
種別: 個人
健診種別: [定期健診, 雇入健診]
テンプレート: template_regular_result.xlsx
マッピング: mapping_regular_result.json
ページ数: 2
出力形式: [PDF, Excel]
用紙: A4
印刷: 縦

構成:
  1ページ目: 基本情報、身体測定、血圧、視力・聴力
  2ページ目: 血液検査、尿検査、総合判定、医師意見
```

#### FRM-00003: 健診結果通知書（簡易版）
```yaml
帳票ID: FRM-00003
帳票名: 健診結果通知書
カテゴリ: 結果報告
種別: 個人
健診種別: [人間ドック, 定期健診, 雇入健診]
テンプレート: template_result_notice.xlsx
マッピング: mapping_result_notice.json
ページ数: 1
出力形式: [PDF, Excel]
用紙: A4
印刷: 縦

用途: 速報用、要再検通知等
```

### 3.2 優先度A: 団体向け一覧・集計

#### FRM-00010: 健診結果一覧表（8名リスト）
```yaml
帳票ID: FRM-00010
帳票名: 健診結果一覧表
カテゴリ: 一覧集計
種別: 団体
健診種別: [人間ドック, 定期健診, 雇入健診]
テンプレート: template_list_8.xlsx
マッピング: mapping_list_8.json
ページ数: 動的
出力形式: [PDF, Excel]
用紙: A3
印刷: 横

構成:
  - 横軸: 受診者8名/ページ
  - 縦軸: 検査項目（身体測定、血液検査、尿検査）
  - 判定: A/B/C/D色分け
  - フッター: 判定集計
```

#### FRM-00011: 事業所別集計表
```yaml
帳票ID: FRM-00011
帳票名: 事業所別集計表
カテゴリ: 一覧集計
種別: 団体
健診種別: [人間ドック, 定期健診]
テンプレート: template_org_summary.xlsx
マッピング: mapping_org_summary.json
ページ数: 1-2
出力形式: [PDF, Excel]
用紙: A4
印刷: 縦

構成:
  - 受診者数（総数、男女別、年代別）
  - 判定分布（A/B/C/D人数・割合）
  - 項目別有所見率
  - 前年比較（オプション）
```

#### FRM-00012: 有所見者リスト
```yaml
帳票ID: FRM-00012
帳票名: 有所見者リスト
カテゴリ: 一覧集計
種別: 団体
健診種別: [人間ドック, 定期健診]
テンプレート: template_abnormal_list.xlsx
マッピング: mapping_abnormal_list.json
ページ数: 動的
出力形式: [PDF, Excel]
用紙: A4
印刷: 横

構成:
  - C/D判定者一覧
  - 氏名、該当項目、検査値、判定
  - カテゴリ別グループ化
```

### 3.3 優先度B: 請求・会計帳票

#### FRM-00020: 請求書
```yaml
帳票ID: FRM-00020
帳票名: 請求書
カテゴリ: 請求会計
種別: 個人/団体
健診種別: [人間ドック, 定期健診, 雇入健診]
テンプレート: template_invoice.xlsx
マッピング: mapping_invoice.json
ページ数: 1
出力形式: [PDF, Excel, 印刷]
用紙: A4
印刷: 縦

構成:
  - 宛先（個人名または事業所名）
  - 請求日、請求番号
  - 明細（コース名、金額、オプション）
  - 合計金額（税抜、消費税、税込）
```

#### FRM-00021: 領収書
```yaml
帳票ID: FRM-00021
帳票名: 領収書
カテゴリ: 請求会計
種別: 個人/団体
テンプレート: template_receipt.xlsx
マッピング: mapping_receipt.json
ページ数: 1
出力形式: [PDF, 印刷]
用紙: A4
印刷: 縦
```

#### FRM-00022: 健診費用明細書
```yaml
帳票ID: FRM-00022
帳票名: 健診費用明細書
カテゴリ: 請求会計
種別: 個人
テンプレート: template_cost_detail.xlsx
マッピング: mapping_cost_detail.json
ページ数: 1
出力形式: [PDF, Excel]
用紙: A4
印刷: 縦

用途: 医療費控除用明細、保険請求用
```

### 3.4 優先度B: 予約・案内帳票

#### FRM-00030: コース日程表
```yaml
帳票ID: FRM-00030
帳票名: コース日程表
カテゴリ: 予約案内
種別: 管理
テンプレート: template_schedule.xlsx
マッピング: mapping_schedule.json
ページ数: 1
出力形式: [PDF, Excel]
用紙: A4
印刷: 横

構成:
  - 日付別コース枠
  - 予約状況（○△×）
  - 定員/予約数
```

#### FRM-00031: 受診案内
```yaml
帳票ID: FRM-00031
帳票名: 受診案内
カテゴリ: 予約案内
種別: 個人
テンプレート: template_guidance.xlsx
マッピング: mapping_guidance.json
ページ数: 1
出力形式: [PDF, 印刷]
用紙: A4
印刷: 縦

構成:
  - 受診者名、受診日時
  - コース名、検査内容
  - 注意事項（食事、服薬等）
  - アクセス情報
```

#### FRM-00032: 問診票
```yaml
帳票ID: FRM-00032
帳票名: 問診票
カテゴリ: 予約案内
種別: 個人
テンプレート: template_questionnaire.xlsx
マッピング: mapping_questionnaire.json
ページ数: 2
出力形式: [PDF, 印刷]
用紙: A4
印刷: 縦

構成:
  1ページ目: 基本情報、既往歴、現病歴
  2ページ目: 自覚症状、生活習慣
```

#### FRM-00033: 受診票
```yaml
帳票ID: FRM-00033
帳票名: 受診票
カテゴリ: 予約案内
種別: 個人
テンプレート: template_exam_ticket.xlsx
マッピング: mapping_exam_ticket.json
ページ数: 1
出力形式: [印刷]
用紙: A4
印刷: 縦

用途: 当日携帯用、検査チェックリスト
```

### 3.5 優先度C: 事後フォロー帳票

#### FRM-00040: 精密検査依頼書
```yaml
帳票ID: FRM-00040
帳票名: 精密検査依頼書
カテゴリ: 事後フォロー
種別: 個人
テンプレート: template_referral.xlsx
マッピング: mapping_referral.json
ページ数: 1
出力形式: [PDF, 印刷]
用紙: A4
印刷: 縦

構成:
  - 紹介先医療機関（選択式）
  - 受診者情報
  - 依頼内容（C/D判定項目）
  - 検査結果抜粋
```

#### FRM-00041: 精検受診勧奨通知
```yaml
帳票ID: FRM-00041
帳票名: 精検受診勧奨通知
カテゴリ: 事後フォロー
種別: 個人
テンプレート: template_followup_notice.xlsx
マッピング: mapping_followup_notice.json
ページ数: 1
出力形式: [PDF, 印刷]
用紙: A4
印刷: 縦

用途: C/D判定者への精検案内
```

#### FRM-00042: 追跡管理リスト
```yaml
帳票ID: FRM-00042
帳票名: 追跡管理リスト
カテゴリ: 事後フォロー
種別: 管理
テンプレート: template_tracking_list.xlsx
マッピング: mapping_tracking_list.json
ページ数: 動的
出力形式: [Excel]
用紙: A4
印刷: 横

構成:
  - C/D判定者リスト
  - 精検受診状況（未/済）
  - フォロー実施状況
  - 次回予定
```

---

## 4. セルマッピング設計

### 4.1 マッピングJSON構造

```json
{
  "version": "1.0.0",
  "form_id": "FRM-00001",
  "description": "人間ドック結果報告書マッピング",

  "sheets": {
    "1ページ目": {
      "patient_info": {
        "name": { "cell": "D5", "source": "T_Patients.氏名" },
        "name_kana": { "cell": "D6", "source": "T_Patients.氏名カナ" },
        "birthdate": { "cell": "D8", "source": "T_Patients.生年月日", "format": "YYYY年MM月DD日" },
        "age": { "cell": "L8", "source": "T_Patients.生年月日", "transform": "age" },
        "gender": { "cell": "N8", "source": "T_Patients.性別" },
        "exam_date": { "cell": "D10", "source": "T_Visits.受診日", "format": "YYYY年MM月DD日" },
        "visit_id": { "cell": "W5", "source": "T_Visits.受診ID" }
      },

      "overall_judgment": {
        "judgment": { "cell": "C15", "source": "T_Visits.総合判定" },
        "comment": { "cell": "E15:W18", "source": "T_Findings.総合所見" }
      },

      "measurements": {
        "height": { "cell": "F22", "source": "T_Measurements.身長" },
        "weight": { "cell": "F23", "source": "T_Measurements.体重" },
        "bmi": { "cell": "F24", "source": "T_Measurements.BMI" },
        "bmi_judgment": { "cell": "H24", "source": "T_Measurements.BMI判定" },
        "bp_systolic": { "cell": "F26", "source": "T_Measurements.血圧_収縮期_1" },
        "bp_diastolic": { "cell": "F27", "source": "T_Measurements.血圧_拡張期_1" },
        "bp_judgment": { "cell": "H26", "source": "T_Measurements.血圧判定" }
      }
    },

    "2ページ目": {
      "blood_tests": {
        "ast": {
          "value": { "cell": "F5", "source": "T_TestResults[項目ID=TI-00019].検査値数値" },
          "judgment": { "cell": "H5", "source": "T_TestResults[項目ID=TI-00019].判定" }
        },
        "alt": {
          "value": { "cell": "F6", "source": "T_TestResults[項目ID=TI-00020].検査値数値" },
          "judgment": { "cell": "H6", "source": "T_TestResults[項目ID=TI-00020].判定" }
        }
        // ... 他の検査項目
      }
    }
  },

  "transforms": {
    "age": "calcAge(birthdate, examDate)",
    "judgment_color": {
      "A": "#FFFFFF",
      "B": "#FFFFCC",
      "C": "#FFCC99",
      "D": "#FF9999"
    }
  },

  "conditions": {
    "show_previous": "T_Visits.前回受診ID IS NOT NULL"
  }
}
```

### 4.2 データソース参照形式

| 参照形式 | 説明 | 例 |
|---------|------|-----|
| `テーブル.列名` | 単純参照 | `T_Patients.氏名` |
| `テーブル[条件].列名` | 条件付き参照 | `T_TestResults[項目ID=TI-00019].検査値数値` |
| `テーブル.列名 + transform` | 変換付き参照 | `T_Patients.生年月日`, transform: "age" |
| `テーブル.列名 + format` | フォーマット参照 | format: "YYYY年MM月DD日" |

### 4.3 変換関数 (transforms)

| 関数名 | 説明 | 入力 | 出力 |
|--------|------|------|------|
| `age` | 年齢計算 | 生年月日, 基準日 | 年齢（整数） |
| `judgment_color` | 判定色 | A/B/C/D | HTMLカラーコード |
| `gender_ja` | 性別日本語化 | M/F/1/2 | 男性/女性 |
| `date_jp` | 和暦変換 | 日付 | 令和X年X月X日 |
| `number_format` | 数値フォーマット | 数値, 小数桁 | フォーマット済文字列 |

---

## 5. テンプレート管理

### 5.1 フォルダ構造

```
Google Drive/
└─ 50_健診結果入力/
   └─ 20_人間ドック/
      └─ 結果入力/
         ├─ 設計書_設定ファイル/
         │   ├─ FORM_MASTER_DESIGN.md  (本設計書)
         │   └─ mappings/
         │       ├─ mapping_dock_result.json
         │       ├─ mapping_regular_result.json
         │       ├─ mapping_list_8.json
         │       └─ ...
         │
         └─ templates/
             ├─ 01_結果報告/
             │   ├─ template_dock_result.xlsx
             │   ├─ template_regular_result.xlsx
             │   └─ template_result_notice.xlsx
             │
             ├─ 02_一覧集計/
             │   ├─ template_list_8.xlsx
             │   ├─ template_org_summary.xlsx
             │   └─ template_abnormal_list.xlsx
             │
             ├─ 03_請求会計/
             │   ├─ template_invoice.xlsx
             │   ├─ template_receipt.xlsx
             │   └─ template_cost_detail.xlsx
             │
             ├─ 04_予約案内/
             │   ├─ template_schedule.xlsx
             │   ├─ template_guidance.xlsx
             │   ├─ template_questionnaire.xlsx
             │   └─ template_exam_ticket.xlsx
             │
             └─ 05_事後フォロー/
                 ├─ template_referral.xlsx
                 ├─ template_followup_notice.xlsx
                 └─ template_tracking_list.xlsx
```

### 5.2 テンプレート命名規則

```
template_[カテゴリ略称]_[帳票略称].xlsx

例:
- template_dock_result.xlsx    (人間ドック結果報告書)
- template_list_8.xlsx         (8名リスト)
- template_invoice.xlsx        (請求書)
```

### 5.3 テンプレート作成ガイドライン

1. **セル配置**
   - データ挿入セルは罫線で囲む
   - 結合セルは最小限に
   - 印刷範囲を設定

2. **書式設定**
   - フォント: 游ゴシック または メイリオ
   - サイズ: 本文10pt、見出し12pt
   - 判定欄: 条件付き書式で色分け

3. **名前付き範囲**
   - 反復データ領域に名前付け
   - 例: `DATA_AREA`, `HEADER_ROW`

---

## 6. 出力処理フロー

### 6.1 個人帳票出力フロー

```
┌────────────────┐
│ AppSheet/GAS   │
│ 出力ボタン押下  │
└───────┬────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 1. パラメータ取得                        │
│    - 受診ID                             │
│    - 帳票ID                             │
│    - 出力形式 (PDF/Excel)               │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 2. マッピング定義読込                    │
│    - mappings/mapping_xxx.json          │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 3. データ取得                           │
│    - T_Patients                         │
│    - T_Visits                           │
│    - T_Measurements                     │
│    - T_TestResults                      │
│    - T_Findings                         │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 4. テンプレートコピー                    │
│    - templates/template_xxx.xlsx        │
│    - → 一時ファイル作成                 │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 5. データ転記                           │
│    - マッピング定義に従いセル設定        │
│    - 変換関数適用                        │
│    - 条件付き処理                        │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 6. 出力生成                             │
│    - Excel: そのまま保存                │
│    - PDF: ExportAsでPDF変換             │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 7. ファイル保存                         │
│    - 02_出力フォルダ/                   │
│    - ファイル名: 受診ID_帳票名_日付.xxx │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 8. 完了処理                             │
│    - T_Visits.出力日時 更新              │
│    - ダウンロードURL返却                 │
└────────────────────────────────────────┘
```

### 6.2 団体帳票出力フロー (一覧表)

```
┌────────────────────────────────────────┐
│ 1. パラメータ取得                        │
│    - 事業所ID                           │
│    - 期間（開始日〜終了日）              │
│    - 帳票ID                             │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 2. 対象データ抽出                        │
│    SELECT * FROM T_Visits               │
│    WHERE 事業所ID = X                   │
│      AND 受診日 BETWEEN 開始 AND 終了   │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 3. ページング処理                        │
│    - 8名/ページで分割                    │
│    - ページ数 = CEILING(件数/8)          │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 4. ページ別データ転記                    │
│    FOR each page:                       │
│      - テンプレートシートコピー          │
│      - 8名分のデータ転記                 │
│      - 判定集計計算                      │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 5. 集計シート生成（オプション）          │
│    - 全体集計                           │
│    - 判定分布グラフ                      │
└───────┬────────────────────────────────┘
        │
        ▼
┌────────────────────────────────────────┐
│ 6. 出力・保存                           │
└────────────────────────────────────────┘
```

---

## 7. GAS実装インターフェース

### 7.1 エクスポート関数

```javascript
/**
 * 個人帳票出力
 * @param {string} visitId - 受診ID
 * @param {string} formId - 帳票ID
 * @param {string} outputFormat - 'PDF' | 'Excel'
 * @returns {string} ダウンロードURL
 */
function exportIndividualForm(visitId, formId, outputFormat) {
  // 実装
}

/**
 * 団体帳票出力（一覧表）
 * @param {string} orgId - 事業所ID
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @param {string} formId - 帳票ID
 * @param {string} outputFormat - 'PDF' | 'Excel'
 * @returns {string} ダウンロードURL
 */
function exportGroupForm(orgId, startDate, endDate, formId, outputFormat) {
  // 実装
}

/**
 * バッチ出力（複数帳票一括）
 * @param {string[]} visitIds - 受診ID配列
 * @param {string} formId - 帳票ID
 * @param {string} outputFormat - 'PDF' | 'Excel'
 * @returns {string} ZIPダウンロードURL
 */
function exportBatch(visitIds, formId, outputFormat) {
  // 実装
}
```

### 7.2 マッピング処理関数

```javascript
/**
 * マッピング定義読込
 * @param {string} formId - 帳票ID
 * @returns {Object} マッピング定義オブジェクト
 */
function loadMapping(formId) {
  // mappings/mapping_xxx.json を読込
}

/**
 * データ転記実行
 * @param {Spreadsheet} spreadsheet - 出力先スプレッドシート
 * @param {Object} mapping - マッピング定義
 * @param {Object} data - 転記データ
 */
function applyMapping(spreadsheet, mapping, data) {
  // マッピング定義に従いデータ転記
}

/**
 * 変換関数適用
 * @param {any} value - 元値
 * @param {string} transform - 変換関数名
 * @param {Object} context - コンテキスト（受診日等）
 * @returns {any} 変換後の値
 */
function applyTransform(value, transform, context) {
  // 変換処理
}
```

---

## 8. AppSheet連携

### 8.1 出力ボタン設定

```yaml
# 個人帳票出力アクション
Action:
  Name: "結果報告書出力"
  For: T_Visits
  Do: "Call a script"
  Script:
    Name: exportIndividualForm
    Args:
      - [_THISROW].[受診ID]
      - "FRM-00001"
      - "PDF"
```

### 8.2 出力画面デザイン

```
┌─────────────────────────────────────────┐
│ 帳票出力                          × 閉じる│
├─────────────────────────────────────────┤
│                                         │
│ 出力対象: 田中 太郎 (2025/01/15受診)    │
│                                         │
│ 帳票選択:                               │
│ ┌─────────────────────────────────────┐ │
│ │ ○ 人間ドック結果報告書              │ │
│ │ ○ 健診結果通知書（簡易）            │ │
│ │ ○ 精密検査依頼書                    │ │
│ └─────────────────────────────────────┘ │
│                                         │
│ 出力形式:                               │
│ ┌──────────┐  ┌──────────┐             │
│ │ ● PDF   │  │ ○ Excel │             │
│ └──────────┘  └──────────┘             │
│                                         │
│         ┌──────────────────┐            │
│         │    📄 出力実行    │            │
│         └──────────────────┘            │
└─────────────────────────────────────────┘
```

---

## 9. 運用ガイド

### 9.1 テンプレート変更手順

1. **テンプレートファイル編集**
   - Google Driveからダウンロード
   - Excelで編集（レイアウト、文言変更）
   - 同名でアップロード（上書き）

2. **マッピング変更が必要な場合**
   - `mappings/mapping_xxx.json` を編集
   - セル位置の更新
   - テスト出力で確認

3. **新規帳票追加**
   - テンプレートファイル作成
   - マッピングJSON作成
   - M_Formsにレコード追加
   - GAS関数にて帳票ID追加

### 9.2 よくあるカスタマイズ

| カスタマイズ | 対応方法 |
|-------------|---------|
| ロゴ変更 | テンプレートのロゴ画像を差替 |
| 文言変更 | テンプレートのテキスト編集 |
| 項目追加 | テンプレート編集 + マッピング追加 |
| レイアウト変更 | テンプレート編集 + マッピングセル位置更新 |
| 新規帳票 | テンプレート新規作成 + マッピング新規作成 |

### 9.3 トラブルシューティング

| 症状 | 原因 | 対処 |
|------|------|------|
| 出力されない | テンプレートファイルなし | ファイルパス確認 |
| データが空 | マッピング不一致 | JSON定義確認 |
| レイアウト崩れ | セル位置ずれ | マッピング更新 |
| PDF変換エラー | 権限不足 | Drive権限確認 |

---

## 10. 将来拡張

### 10.1 Phase 2: 電子署名対応

```yaml
機能:
  - 医師電子署名欄
  - タイムスタンプ付与
  - 改ざん検知

対応帳票:
  - 結果報告書
  - 精密検査依頼書
```

### 10.2 Phase 3: 多言語対応

```yaml
機能:
  - 英語版テンプレート
  - 中国語版テンプレート
  - 言語切替パラメータ

対応帳票:
  - 結果報告書
  - 受診案内
```

### 10.3 Phase 4: Web配信

```yaml
機能:
  - 受診者ポータル連携
  - メール添付送信
  - QRコードアクセス

対応帳票:
  - 結果報告書
  - 受診案内
```

---

## 付録

### A. 帳票ID一覧

| 帳票ID | 帳票名 | カテゴリ | 優先度 |
|--------|--------|---------|--------|
| FRM-00001 | 人間ドック結果報告書 | 結果報告 | A |
| FRM-00002 | 定期健康診断結果報告書 | 結果報告 | A |
| FRM-00003 | 健診結果通知書 | 結果報告 | A |
| FRM-00010 | 健診結果一覧表 | 一覧集計 | A |
| FRM-00011 | 事業所別集計表 | 一覧集計 | A |
| FRM-00012 | 有所見者リスト | 一覧集計 | A |
| FRM-00020 | 請求書 | 請求会計 | B |
| FRM-00021 | 領収書 | 請求会計 | B |
| FRM-00022 | 健診費用明細書 | 請求会計 | B |
| FRM-00030 | コース日程表 | 予約案内 | B |
| FRM-00031 | 受診案内 | 予約案内 | B |
| FRM-00032 | 問診票 | 予約案内 | B |
| FRM-00033 | 受診票 | 予約案内 | B |
| FRM-00040 | 精密検査依頼書 | 事後フォロー | C |
| FRM-00041 | 精検受診勧奨通知 | 事後フォロー | C |
| FRM-00042 | 追跡管理リスト | 事後フォロー | C |

### B. 作成情報

- 作成日: 2025-12-02
- バージョン: 1.0
- 参照: iD-Heart帳票サンプル、既存GASシステム
- 対応健診種別: 人間ドック、定期健康診断（労安法）
