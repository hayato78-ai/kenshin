# CDmedical 健診システム データ構造設計書
## シート・カラム詳細定義（iD-Heart II準拠）

> **バージョン**: 2.3
> **作成日**: 2025-12-20
> **最終更新**: 2025-12-25
> **ステータス**: 🟢 マスタードキュメント（正式版）

---

## 1. シート一覧

### マスタ系

| シート名 | 説明 | レコード数目安 |
|----------|------|---------------|
| M_検査項目 | 検査項目定義・判定基準 | 200-500 |
| M_検査所見 | 所見テンプレート | 300-1000 |
| M_コース | 健診コース | 10-50 |
| M_コース項目 | コース-検査項目関連 | 500-2000 |
| M_判定項目 | 部位別判定項目 | 20-50 |
| M_団体 | 企業・団体 | 50-500 |
| M_保険者 | 保険者 | 10-100 |

### トランザクション系

| シート名 | 説明 | レコード数目安 | 実装状態 |
|----------|------|---------------|----------|
| T_受診者 | 受診者基本情報 | 1000-50000 | ✅ 実装済（受診者マスタ） |
| T_受診記録 | 受診ごとの情報 | 5000-100000 | ✅ 実装済 |
| 検査結果 | 検査結果（**横持ち**: 1患者1行、101列） | 1000-50000 | ✅ 実装済 |
| T_所見 | 所見データ（縦持ち） | 10000-100000 | ✅ 実装済 |
| T_判定 | 判定結果 | 50000-500000 | 予定 |
| T_進捗 | 進捗管理 | 5000-100000 | 予定 |

---

## 2. M_検査項目 詳細

### カラム定義（主要カラム）

```
A: item_code        - 項目コード (PK) 例: 011001
B: jlac10_code      - JLAC10コード
C: item_name        - 項目名 例: 身長
D: item_name_kana   - カナ
E: category         - 区分 (身体情報/検体検査/画像診断)
F: sub_category     - 小分類 (身体計測/脂質/血液一般...)
G: data_type        - データ型 (numeric/text/select)
H: decimal_places   - 小数桁数
I: unit             - 単位 (cm, kg, mg/dl...)
J-K: valid_min/max  - 有効範囲
L-M: std_male/female - 帳票基準値表記
N-X: criteria_*_m   - 男性判定基準 (F低〜F高)
Y-AI: criteria_*_f  - 女性判定基準 (F低〜F高)
AJ: default_value   - 初期値
AK: finding_ref     - 所見参照初期値
AL-AO: price_*      - オプション単価
AP-AQ: lab_code     - 検査会社コード
AR: display_order   - 表示順
AS: notes           - 備考
AT: is_active       - 有効フラグ
```

### サンプルデータ

| item_code | item_name | category | unit | std_male |
|-----------|-----------|----------|------|----------|
| 011001 | 身長 | 身体情報 | cm | |
| 011002 | 体重 | 身体情報 | kg | |
| 011025 | BMI | 身体情報 | | 25未満 |
| 011014 | 腹囲 | 身体情報 | cm | 85未満 |
| 011005 | 血圧(最高) | 身体情報 | mmHg | 130未満 |
| 021018 | 総コレステロール | 検体検査 | mg/dl | 150〜219 |
| 021022 | LDL-C | 検体検査 | mg/dl | 120未満 |
| 021020 | HDL-C | 検体検査 | mg/dl | 40以上 |
| 021037 | 空腹時血糖 | 検体検査 | mg/dl | 100未満 |
| 012003 | 心電図 | 画像診断 | | |
| 012012 | 胸部X線 | 画像診断 | | |
| 050000 | eGFR | 検体検査 | | 60以上 |

---

## 3. M_検査所見 詳細

### カラム定義

```
A: finding_code      - 所見コード (PK)
B: finding_name      - 所見名
C: default_judgment  - デフォルト判定
D: exam_item_refs    - 紐づく検査項目
E: is_history        - 既往歴フラグ
F: notes             - 備考
G: display_order     - 表示順
H: is_active         - 有効フラグ
```

### サンプルデータ（iD-Heart参照）

| code | name | judgment | notes |
|------|------|----------|-------|
| 0001 | 異常なし | A | |
| 0002 | 異常あり | C | |
| 0003 | 未提出 | | |
| 0005 | 陳旧性炎症性変化 | B | |
| 0006 | 高血圧 | | 既往歴 |
| 0007 | 心臓病 | | 既往歴 |
| 0116 | ST上昇 | C | |
| 0117 | ST下降 | C | |
| 0118 | 低電位 | B | |
| 0121 | 右室肥大 | C | |
| 0122 | 左室肥大 | C | |

---

## 4. 検査結果シート 詳細（横持ち形式）

> **✅ 正式採用: 横持ち形式**
>
> 本システムでは**横持ち形式**（1行=1患者、各検査項目を列として配置）を正式採用しています。
>
> - **シート名**: `検査結果`
> - **定義場所**: `portalApi.js` の `LAB_RESULT_COLUMNS`
> - **データ構造**: 1患者1行、検査項目ごとに列を配置
> - **総列数**: 101列（基本情報5列 + 検査項目96項目）
>
> **採用理由**:
> - BML CSV取込との整合性（横展開フォーマット）
> - 1患者の全検査項目を1行で管理（読み書き効率）
> - Excel帳票への直接マッピング（セル位置指定）

### カラム構造

#### 基本情報（5列）

| 列 | ID | 名前 | 説明 |
|----|-----|------|------|
| A | PATIENT_ID | 受診者ID | P-00001形式 |
| B | KARTE_NO | カルテNo | 6桁 |
| C | EXAM_DATE | 受診日 | 日付 |
| D | BML_PATIENT_ID | BML患者ID | BML通信用 |
| E | CSV_IMPORT_DATE | CSV取込日時 | タイムスタンプ |

#### 検査項目（96項目） - 2025-12-25更新

| カテゴリ | 項目数 | 項目例 |
|----------|--------|--------|
| 血液学検査 | 13 | WBC, RBC, Hb, Ht, PLT, MCV, MCH, MCHC, 好中球, 好塩基球, 好酸球, リンパ球, 単球 |
| 肝胆膵機能 | 12 | TP, ALB, AST, ALT, γ-GTP, ALP, LDH, T-Bil, Ch-E, アミラーゼ |
| 脂質検査 | 4 | TC, TG, HDL-C, LDL-C |
| 糖代謝 | 2 | FBS, HbA1c |
| 腎機能 | 7 | Cr, BUN, eGFR, UA, Na, K, Cl |
| 電解質・鉄代謝 | 4 | Ca, Fe, TIBC |
| 尿検査 | 10 | 尿蛋白, 尿糖, 尿潜血, ウロビリノーゲン, 尿PH, 尿ビリルビン, アセトン体 |
| 尿沈渣 | 4 | 白血球, 赤血球, 扁平上皮, 細菌 |
| 身体測定 | 4 | 身長, 体重, BMI, 腹囲 |
| 血圧 | 5 | 収縮期血圧, 拡張期血圧, 収縮期血圧_2回目, 拡張期血圧_2回目, 脈拍 |
| 便検査 | 4 | 便潜血1回目, 便潜血2回目, 便ヘモグロビン1回目, 便ヘモグロビン2回目 |
| 凝固 | 2 | PT, APTT |
| 腫瘍マーカー | 12 | PSA, CEA, CA19-9, CA125, NSE, エラスターゼ1, 抗p53抗体, CYFRA21-1, SCC, ProGRP, AFP, PIVKA II |
| 感染症 | 10 | HBs抗原, HCV抗体, TPHA, RPR, HBs抗体, HIV, STS, 梅毒トレポネーマ抗体 |
| 心臓・甲状腺 | 4 | NT-proBNP, FT3, FT4, TSH |
| 血液型 | 2 | ABO式, Rho・D |
| ピロリ菌 | 2 | 尿素呼気試験, ピロリ抗体 |
| 自己抗体 | 1 | RF定量 |

#### BMLコード対応表（主要項目）

| BMLコード | 項目名 | カテゴリ |
|-----------|--------|----------|
| 0000301 | 白血球数 | 血液学検査 |
| 0000302 | 赤血球数 | 血液学検査 |
| 0000303 | 血色素量 | 血液学検査 |
| 0000453 | 総コレステロール | 脂質検査 |
| 0000460 | HDLコレステロール | 脂質検査 |
| 0000410 | LDLコレステロール | 脂質検査 |
| 0000503 | 空腹時血糖 | 糖代謝 |
| 0003317 | HbA1c | 糖代謝 |
| 0000413 | クレアチニン | 腎機能 |
| 0005005 | PSA | 腫瘍マーカー |
| 0005001 | CEA | 腫瘍マーカー |
| 0004007 | TSH | 甲状腺 |

> **詳細なBMLコード対応は `portalApi.js` の `LAB_RESULT_COLUMNS.ITEMS` を参照**

<!--
### カラム定義（未実装・将来設計案）

```
A: result_id         - 結果ID (PK)
B: visit_id          - 受診ID (FK)
C: item_code         - 検査項目コード (FK)
D: result_value      - 結果値
E: result_text       - 結果文字（定性）
F: judgment          - 判定
G: judgment_auto     - 自動判定
H: finding_codes     - 所見コード
I: finding_text      - 所見自由記述
J: prev_value        - 前回値
K: prev_prev_value   - 前々回値
L: is_abnormal       - 異常フラグ
M: notes             - 備考
N: created_at        - 作成日時
O: updated_at        - 更新日時
```

### データ例（未実装・将来設計案）

| visit_id | item_code | result_value | judgment |
|----------|-----------|--------------|----------|
| 20251217-001 | 011001 | 167.7 | * |
| 20251217-001 | 011002 | 66.1 | * |
| 20251217-001 | 011025 | 23.5 | A |
| 20251217-001 | 011005 | 114 | A |
| 20251217-001 | 021022 | 142 | C |
-->

---

## 5. T_所見 詳細

### 概要

所見コメント入力が必要な検査項目（画像診断、心電図、内視鏡など）のデータを保存するシート。
**縦持ち・検査項目別構造**を採用し、1レコード = 1検査項目の所見となる。

> **注意**: カラム構成は運用に応じて増減する可能性があるため、Config.js の FINDINGS_DEF を正とする。

### カラム定義（現行）

```
A: finding_id      - 所見ID (PK) F00001形式
B: patient_id      - 受診者ID (FK) P00001形式
C: karte_no        - カルテNo（6桁、クエリ用）
D: item_id         - 項目ID (FK) H02xxxx形式
E: finding_text    - 所見テキスト（自由記述またはテンプレート）
F: judgment        - 判定 (A/B/C/D/E/F)
G: template_id     - 使用テンプレートID
H: exam_date       - 検査実施日
I: input_by        - 入力者
J: created_at      - 作成日時
K: updated_at      - 更新日時
```

### 対象検査項目

| カテゴリ | サブカテゴリ | 例 |
|----------|--------------|-----|
| 画像診断 | X線 | 胸部X線、腹部X線 |
| 画像診断 | 超音波検査 | 腹部超音波、頸部超音波、心臓超音波 |
| 画像診断 | CT | 胸部CT、腹部CT |
| 画像診断 | MRI | 脳MRI、脳MRA、腹部MRI等 |
| 心電図 | - | 心電図 |
| 内視鏡 | - | 上部消化管内視鏡、下部消化管内視鏡 |

### データ例

| finding_id | patient_id | karte_no | item_id | finding_text | judgment |
|------------|------------|----------|---------|--------------|----------|
| F00001 | P00123 | 001234 | H020010 | 異常なし | A |
| F00002 | P00123 | 001234 | H020020 | 洞性徐脈 | B |
| F00003 | P00123 | 001234 | H020040 | 脂肪肝（軽度） | C |

### 関連CRUD関数（CRUD.js）

- `createFinding(data)` - 所見作成
- `getFindingsByKarteNo(karteNo)` - カルテNoで取得
- `getFindingsByPatientId(patientId)` - 受診者IDで取得
- `updateFinding(findingId, data)` - 所見更新
- `deleteFinding(findingId)` - 所見削除
- `upsertFindings(karteNo, findingsData)` - 一括作成/更新

---

## 6. T_進捗 詳細

### カラム定義

```
A: progress_id       - 進捗ID (PK)
B: visit_id          - 受診ID (FK)
C: pre_form          - 事前帳票 (完了/途中/未)
D: reception         - 受付
E: day_form          - 当日帳票
F: questionnaire     - 問診結果
G: physical          - 身体情報
H: lab_test          - 検体検査
I: imaging           - 画像診断
J: specific_health   - 特定健診
K: judgment          - 判定項目
L: findings          - 所見文章
M: report            - 報告書
N: billing_org       - 請求先請求
O: billing_person    - 個人請求
P: payment           - 個人入金
Q: updated_at        - 更新日時
```

### ステータス値

| 値 | 表示 | 説明 |
|----|------|------|
| 完了 | 💚 | 処理完了 |
| 途中 | 💛 | 処理中 |
| 未 | ❤️ | 未処理 |
| 対象外 | ⬜ | 対象外 |

---

## 7. 既存GASコードとの対応

### portalApi.js 対応

```javascript
// 既存 PATIENT_COL → 新構造
const PATIENT_COL = {
  PATIENT_ID: 0,   // A列: P-00001形式
  NAME: 1,         // B列: 氏名
  KANA: 2,         // C列: カナ
  BIRTHDATE: 3,    // D列: 生年月日
  GENDER: 4,       // E列: 性別
  // ...
};

// 既存 VISIT_COL → 新構造
const VISIT_COL = {
  VISIT_ID: 0,     // A列: 20251217-001形式
  PATIENT_ID: 1,   // B列: 受診者ID (FK)
  VISIT_DATE: 2,   // C列: 受診日
  // ...
};
```

### Config.js シート名対応

```javascript
const DB_CONFIG = {
  SHEETS: {
    // マスタ
    EXAM_ITEM: 'M_検査項目',
    FINDING: 'M_検査所見',
    COURSE: 'M_コース',
    ORG: 'M_団体',

    // トランザクション
    PATIENT: 'T_受診者',
    VISIT: 'T_受診記録',
    TEST_RESULT: '検査結果',  // ✅ 横持ち形式（1患者1行、101列）
    FINDINGS: 'T_所見',       // 縦持ち形式（1レコード = 1検査項目の所見）
    PROGRESS: 'T_進捗'
  }
};
```

> **✅ 確定**: 検査結果は `検査結果` シート（横持ち形式: 1患者1行、101列）を正式採用。
> 詳細は `portalApi.js` の `LAB_RESULT_COLUMNS` を参照。

---

## 8. 判定ロジック

### 8.1 判定区分定義（統一基準）

> **⚠️ 重要**: 本システムにおける判定区分の**唯一の正式定義**です。
> 他の設計書・コードはこの定義に従ってください。

| 判定 | 名称 | 説明 | 総合所見への反映 | 優先度 |
|------|------|------|-----------------|--------|
| A | 異常なし | 基準値内、健康 | 反映しない | 6（最低） |
| B | 軽度異常 | わずかな逸脱、経過観察不要 | カテゴリ別コメント生成 | 5 |
| C | 要経過観察 | 定期的なフォローアップ推奨 | カテゴリ別コメント生成（優先度高） | 4 |
| D | 要精密検査 | 追加検査が必要 | カテゴリ別コメント生成（最優先） | 3 |
| E | 要治療 | 医療機関での治療開始推奨 | 必須記載 | 2 |
| F | 治療中 | 現在治療継続中 | 必要に応じて記載 | 1（最高） |
| * | 判定不可 | 未実施・判定基準なし・欠測 | 反映しない | - |

```javascript
// 判定区分の正式定義（Config.jsに実装）
const JUDGMENT_CODES = {
  A: { code: 'A', name: '異常なし',     priority: 6, needsComment: false },
  B: { code: 'B', name: '軽度異常',     priority: 5, needsComment: true },
  C: { code: 'C', name: '要経過観察',   priority: 4, needsComment: true },
  D: { code: 'D', name: '要精密検査',   priority: 3, needsComment: true },
  E: { code: 'E', name: '要治療',       priority: 2, needsComment: true },
  F: { code: 'F', name: '治療中',       priority: 1, needsComment: true },
  '*': { code: '*', name: '判定不可',   priority: 99, needsComment: false }
};
```

### 8.2 数値判定ロジック

```javascript
/**
 * 検査値から判定を自動算出
 * @param {number} value - 検査値
 * @param {string} gender - 性別（'男性'/'女性'）
 * @param {Object} criteria - M_検査項目から取得した判定基準
 * @returns {string} 判定コード（A/B/C/D/E/F/*）
 */
function autoJudge(value, gender, criteria) {
  if (value === null || value === undefined || value === '') return '*';

  const c = gender === '男性' ? criteria.male : criteria.female;
  if (!c) return '*'; // 判定基準なし

  // 低値側判定
  if (c.e_low !== null && value < c.e_low) return 'E';
  if (c.d_low !== null && value < c.d_low) return 'D';
  if (c.c_low !== null && value < c.c_low) return 'C';
  if (c.b_low !== null && value < c.b_low) return 'B';

  // 正常範囲
  if (value >= c.a_low && value <= c.a_high) return 'A';

  // 高値側判定
  if (c.b_high !== null && value <= c.b_high) return 'B';
  if (c.c_high !== null && value <= c.c_high) return 'C';
  if (c.d_high !== null && value <= c.d_high) return 'D';
  if (c.e_high !== null && value <= c.e_high) return 'E';

  // 基準値外（異常高値/低値）
  return 'E';
}
```

### 8.3 総合判定ロジック

```javascript
/**
 * 全検査項目の判定から総合判定を算出
 * @param {Array<string>} judgments - 各項目の判定コード配列
 * @returns {string} 総合判定コード
 */
function calculateOverallJudgment(judgments) {
  // 判定不可・未実施を除外
  const validJudgments = judgments.filter(j => j && j !== '*');

  if (validJudgments.length === 0) return '*';

  // 優先度順（数値が小さいほど優先）
  const priority = { 'F': 1, 'E': 2, 'D': 3, 'C': 4, 'B': 5, 'A': 6 };

  // 最も悪い（優先度が高い）判定を採用
  return validJudgments.sort((a, b) =>
    priority[a] - priority[b]
  )[0];
}

// 総合判定ルール
// ・1つでもE → 総合E
// ・EなしでD → 総合D
// ・D以下なしでC → 総合C
// ・C以下なしでB → 総合B
// ・全てA → 総合A
// ・F（治療中）は総合判定に含めない場合あり（設定による）
```

### 8.4 部位別判定ロジック

```javascript
/**
 * 部位（カテゴリ）別の総合判定を算出
 * @param {Object} results - 検査結果オブジェクト {item_code: {value, judgment}}
 * @param {Object} judgmentItem - M_判定項目の1レコード
 * @returns {string} 部位別判定コード
 */
function getAreaJudgment(results, judgmentItem) {
  // 関連する検査項目の判定を集計
  const judgments = judgmentItem.related_items
    .map(code => results[code]?.judgment)
    .filter(j => j && j !== '*');

  if (judgments.length === 0) return '*';

  // 最も悪い判定を採用
  const priority = { 'F': 1, 'E': 2, 'D': 3, 'C': 4, 'B': 5, 'A': 6 };
  return judgments.sort((a, b) =>
    priority[a] - priority[b]
  )[0];
}
```

### 8.5 性別依存項目

以下の検査項目は性別によって判定基準が異なります。

| 項目コード | 項目名 | 男性基準 | 女性基準 | 備考 |
|-----------|--------|---------|---------|------|
| 021003 | Hb（血色素量） | 13.1-16.3 g/dL | 12.1-14.5 g/dL | |
| 021002 | RBC（赤血球数） | 400-539 万/μL | 360-489 万/μL | |
| 021004 | Ht（ヘマトクリット） | 38.5-48.9 % | 35.5-43.9 % | |
| 011014 | 腹囲 | 85cm未満 | 90cm未満 | メタボ基準 |
| 021014 | Cr（クレアチニン） | 0.6-1.1 mg/dL | 0.4-0.8 mg/dL | |
| 050000 | eGFR | 60以上 | 60以上 | 計算式に性別使用 |

> **実装上の注意**: `autoJudge()` 呼び出し時は必ず性別を渡すこと。
> 性別未設定の場合はエラーとするか、共通基準（存在する場合）を適用。

### 8.6 判定基準例（人間ドック学会2025年度）

| 項目 | A（異常なし） | B（軽度異常） | C（要経過観察） | D（要精密検査） | E（要治療） |
|------|--------------|--------------|----------------|----------------|------------|
| AST | ≤30 | 31-35 | 36-50 | 51-100 | >100 |
| ALT | ≤30 | 31-40 | 41-50 | 51-100 | >100 |
| γ-GTP | ≤50 | 51-80 | 81-100 | 101-200 | >200 |
| HbA1c | ≤5.5 | 5.6-5.9 | 6.0-6.4 | 6.5-6.9 | ≥7.0 |
| LDL-C | 60-119 | 120-139 | 140-159 | 160-179 | ≥180 |
| HDL-C | ≥40 | 35-39 | 30-34 | - | <30 |
| 空腹時血糖 | ≤99 | 100-109 | 110-125 | - | ≥126 |
| eGFR | ≥60 | 45-59 | 30-44 | 15-29 | <15 |
| 収縮期血圧 | <130 | 130-139 | 140-159 | 160-179 | ≥180 |
| 拡張期血圧 | <85 | 85-89 | 90-99 | 100-109 | ≥110 |

> **注意**: 上記は参考値です。実際の判定基準は `M_検査項目` シートの N-AI列を正とします。

### 8.7 カテゴリ別検査項目グループ

所見生成時のカテゴリ分類：

| カテゴリ | 含まれる検査項目 | 判定対象 |
|---------|-----------------|---------|
| 循環器系 | 血圧（収縮期/拡張期）、心電図 | ✅ |
| 消化器系 | AST、ALT、γ-GTP、ALP、T-Bil、腹部超音波 | ✅ |
| 代謝系（糖） | FBS、HbA1c | ✅ |
| 代謝系（脂質） | TC、HDL-C、LDL-C、TG、non-HDL-C | ✅ |
| 腎機能 | Cr、BUN、eGFR、尿蛋白、尿潜血 | ✅ |
| 血液系 | WBC、RBC、Hb、Ht、PLT、MCV、MCH、MCHC | ✅ |
| 尿酸代謝 | UA（尿酸） | ✅ |
| 身体計測 | BMI、腹囲 | ✅ |
| 眼科 | 視力、眼圧、眼底 | ✅ |
| 聴力 | 聴力（1000Hz/4000Hz） | ✅ |

---

## 9. 旧設計書との対応

### 9.1 設計書の位置づけ

| 設計書 | ステータス | 役割 | 備考 |
|--------|----------|------|------|
| **DATA_STRUCTURE_DESIGN_v2.md** | 🟢 **正式版** | データ構造の唯一の定義 | 本ドキュメント |
| SYSTEM_DESIGN_SPECIFICATION.md | 🟡 参照用 | UI/UX・API設計 | 将来UI_DESIGN.mdに分割予定 |
| GAS_SYSTEM_DESIGN.md | 🟠 アーカイブ予定 | 旧運用設計 | 一部をOPERATIONS.mdに移行予定 |
| DATA_STRUCTURE_DESIGN.md | 🔴 非推奨 | 旧データ構造 | v2に統合済 |

### 9.2 シート名対応表

| 本設計書（v2） | GAS_SYSTEM_DESIGN | DATA_STRUCTURE_DESIGN | 実装（GAS） | 状態 |
|---------------|-------------------|----------------------|-------------|------|
| T_受診者 | 受診者マスタ | T_Patients | 受診者マスタ | ✅ |
| T_受診記録 | （受診者マスタに含む） | T_Visits | T_受診記録 | ✅ |
| **検査結果（横持ち）** | 身体測定+血液検査 | ~~T_TestResults（縦持ち）~~ | 検査結果 | ✅ **正式採用** |
| T_所見（縦持ち） | 所見 | T_Findings | T_所見 | ✅ |
| M_検査項目 | 判定マスタ（一部） | M_JudgmentRules | M_検査項目 | ✅ |
| M_検査所見 | 所見テンプレート | M_FindingTemplates | M_検査所見 | ✅ |
| T_進捗 | - | - | T_進捗 | 予定 |

### 9.3 判定区分対応表

| 本設計書（v2） | GAS_SYSTEM_DESIGN | DATA_STRUCTURE_DESIGN | 対応 |
|---------------|-------------------|----------------------|------|
| A: 異常なし | A: 異常なし | A: 異常なし | ✅ 一致 |
| B: 軽度異常 | B: 軽度異常 | B: 軽度異常 | ✅ 一致 |
| C: 要経過観察 | C: 要精密検査 | C: 要経過観察/要精検 | ⚠️ 名称統一 |
| D: 要精密検査 | D: 要治療 | D: 要治療/要精検 | ⚠️ 意味変更 |
| E: 要治療 | E: 治療中 | E: - | ⚠️ 新規追加 |
| F: 治療中 | - | - | ⚠️ 新規追加 |
| *: 判定不可 | - | - | ⚠️ 新規追加 |

> **移行時の注意**: 旧システムでC=要精検、D=要治療としていた場合、
> 本設計書ではD=要精密検査、E=要治療に変更されています。
> データ移行時は判定コードの読み替えが必要です。

### 9.4 関数名対応表

| 本設計書（v2） | GAS_SYSTEM_DESIGN | SYSTEM_DESIGN_SPEC | 実装推奨 |
|---------------|-------------------|-------------------|---------|
| `autoJudge(value, gender, criteria)` | `judge(itemId, value, gender)` | `autoJudge(visitId, itemCode)` | v2版を採用 |
| `calculateOverallJudgment(judgments)` | `calculateOverallJudgment(judgments)` | - | 一致 |
| `getAreaJudgment(results, judgmentItem)` | - | - | v2版を採用 |

---

## 10. Excel出力システム（Python連携）

### 10.1 システム概要

Web UI（GAS）からExcel帳票を出力する機能。GASからリクエストJSONを生成し、ローカルPythonが監視・処理してExcelを生成する非同期方式。

```
[Web UI (GAS)]          [Google Drive]           [Python]
     │                       │                       │
  患者選択                   │                       │
     ↓                       │                       │
exportIndividualReport() ────→ pending/{id}.json     │
     │                       │                       │
     │                       │←──── 5秒ごとにスキャン
     │                       │                       ↓
     │                       │              unified_transcriber.py
     │                       │                       │
     │                       │←── processed/{id}_result.json
     │                       │←── output/{id}.xlsx
     ↓                       │                       │
checkExportStatus() ←────────┘                       │
     ↓
ダウンロードリンク表示
```

### 10.2 フォルダ構成

| フォルダ | 用途 | 管理 |
|---------|------|------|
| `pending/` | リクエストJSON待機 | GASが作成 → Pythonが取得 |
| `processed/` | 処理完了JSONを移動 | Pythonが移動 |
| `error/` | エラー発生時に移動 | Pythonが移動 |
| `output/` | 生成Excel出力先 | Pythonが作成 |

### 10.3 運用手順

#### Python監視の起動

```bash
cd python
python3 unified_transcriber.py --watch
```

**オプション:**
- `--settings settings.yaml`: 設定ファイル指定
- `--watch`: 監視モード（5秒ごとにpendingスキャン）

**停止:** `Ctrl+C`

#### Web UIからの出力

1. 帳票出力タブを開く
2. 検索条件を入力して患者を検索
3. 対象患者をチェック
4. 「出力」ボタンをクリック
5. 「Python処理中」表示を確認
6. 「状態確認」ボタンで処理状況を確認
7. 完了後、ダウンロードリンクをクリック

### 10.4 トラブルシューティング

| 症状 | 原因 | 対処 |
|------|------|------|
| JSON作成されない | フォルダ権限・パス設定 | 設定シートのフォルダID確認 |
| Pythonが検知しない | 監視未起動 | `--watch` オプション確認 |
| Excel生成失敗 | テンプレート・マッピング不備 | logs/ のエラーログ確認 |
| 「処理中」のまま | Drive同期遅延 | 数分待って再確認 |

### 10.5 設定ファイル

**python/settings.yaml:**
```yaml
poll_interval: 5  # ポーリング間隔（秒）

folders:
  pending: /path/to/pending
  processed: /path/to/processed
  error: /path/to/error

output_dir: /path/to/output
```

### 10.6 関連ファイル

| ファイル | 役割 |
|---------|------|
| `gas/reportExporter.js` | リクエストJSON生成 |
| `gas/excelExportBridge.js` | Python連携ブリッジ |
| `python/drive_watcher.py` | フォルダ監視（ポーリング） |
| `python/unified_transcriber.py` | Excel転記エンジン |

---

## 変更履歴

| 日付 | Ver | 内容 |
|------|-----|------|
| 2025-12-20 | 2.0 | 初版（iD-Heart分析に基づく）|
| 2025-12-25 | 2.1 | T_所見セクション追加（縦持ち・検査項目別構造）|
| 2025-12-25 | 2.2 | 判定ロジック大幅拡充（8章）、判定区分統一定義、性別依存項目、旧設計書対応表（9章）追加。マスタードキュメントとして正式化 |
| 2025-12-25 | 2.3 | **検査結果シートを横持ち形式として正式採用**。4章を「検査結果シート詳細（横持ち形式）」に改訂、縦持ち設計案（未実装）の記載を削除 |
| 2025-12-25 | 2.4 | **Excel出力システム（Python連携）** 10章追加。ポーリング方式監視、運用手順、トラブルシューティング |
