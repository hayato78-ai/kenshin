# CLAUDE_CODE_INSTRUCTIONS.md
# Claude Code 実装指示書 - マスタ管理機能追加

> **作成日**: 2025-12-20
> **目的**: 検査項目マスタ・コース項目管理機能の追加
> **重要**: 既存コードへの影響を最小限に

---

## 🎯 実装目標

iD-Heart II システムを参考に、以下の機能を追加：
1. 検査項目マスタの CRUD
2. 検査所見マスタの CRUD
3. コース-検査項目マッピングの管理
4. マスタ管理UI（ポータルにタブ追加）

---

## ⚠️ 重要ルール（既存コードとの統合）

### 絶対に守ること

1. **既存ファイルの大幅変更禁止**
   - `portalApi.js`: 関数追加のみ、既存関数は変更しない
   - `MasterData.js`: 参照のみ、構造変更しない
   - `portal.html`: タブ追加のみ、既存タブは変更しない

2. **新規ファイルの命名規則**
   ```
   master_[機能名].js    # 例: master_examItem.js
   ```

3. **API関数の命名規則**
   ```javascript
   // 既存パターンに合わせる
   portalGetXxx()      // 取得
   portalSaveXxx()     // 保存
   portalDeleteXxx()   // 削除
   portalSearchXxx()   // 検索
   ```

4. **シート操作**
   - `Config.js` の `DB_CONFIG` を使用
   - 直接シート名をハードコードしない

---

## 📁 新規作成ファイル

### 1. master_examItem.js（検査項目マスタCRUD）

```javascript
/**
 * master_examItem.js - 検査項目マスタ管理
 * 
 * 依存: Config.js, CRUD.js
 * シート: M_検査項目（新規作成 or 検査項目マスタを使用）
 */

// ============================================
// 検査項目マスタ CRUD
// ============================================

/**
 * 検査項目一覧を取得
 * @param {Object} criteria - 検索条件 {category, keyword}
 * @returns {Object} {success, data, count}
 */
function portalGetExamItems(criteria) {
  // 実装
}

/**
 * 検査項目を保存（新規/更新）
 * @param {Object} itemData - 検査項目データ
 * @returns {Object} {success, data}
 */
function portalSaveExamItem(itemData) {
  // 実装
}

/**
 * 検査項目を削除（論理削除）
 * @param {string} itemCode - 項目コード
 * @returns {Object} {success}
 */
function portalDeleteExamItem(itemCode) {
  // 実装
}
```

### 2. master_finding.js（検査所見マスタCRUD）

```javascript
/**
 * master_finding.js - 検査所見マスタ管理
 * 
 * シート: M_検査所見（新規作成）
 */

function portalGetFindings(criteria) { }
function portalSaveFinding(findingData) { }
function portalDeleteFinding(findingCode) { }
```

### 3. master_course.js（コース-項目マッピング）

```javascript
/**
 * master_course.js - コース管理・項目マッピング
 * 
 * シート: M_コース, M_コース項目（新規作成）
 */

function portalGetCourses() { }
function portalSaveCourse(courseData) { }
function portalGetCourseItems(courseId) { }
function portalSaveCourseItems(courseId, itemCodes) { }
```

---

## 📊 新規シート定義

### M_検査項目（拡張版）

既存の `検査項目マスタ` シートを拡張、または新規作成：

| 列 | カラム名 | 型 | 説明 |
|----|---------|-----|------|
| A | item_code | STRING | 項目コード（PK）例: 011001 |
| B | item_name | STRING | 項目名 |
| C | category | STRING | 区分（身体情報/検体検査/画像診断）|
| D | sub_category | STRING | 小分類 |
| E | data_type | STRING | データ型（numeric/text/select）|
| F | unit | STRING | 単位 |
| G | std_male | STRING | 男性基準値表記 |
| H | std_female | STRING | 女性基準値表記 |
| I-T | criteria_* | NUMBER | 判定基準（男女別 A〜F）|
| U | display_order | NUMBER | 表示順 |
| V | is_active | BOOLEAN | 有効フラグ |
| W | created_at | DATETIME | 作成日時 |
| X | updated_at | DATETIME | 更新日時 |

### M_検査所見

| 列 | カラム名 | 型 | 説明 |
|----|---------|-----|------|
| A | finding_code | STRING | 所見コード（PK）|
| B | finding_name | STRING | 所見名 |
| C | default_judgment | STRING | デフォルト判定（A〜F）|
| D | exam_item_refs | STRING | 紐づく検査項目（カンマ区切り）|
| E | is_history | BOOLEAN | 既往歴フラグ |
| F | display_order | NUMBER | 表示順 |
| G | is_active | BOOLEAN | 有効フラグ |

### M_コース項目

| 列 | カラム名 | 型 | 説明 |
|----|---------|-----|------|
| A | mapping_id | STRING | マッピングID（PK）|
| B | course_id | STRING | コースID（FK）|
| C | item_code | STRING | 検査項目コード（FK）|
| D | is_required | BOOLEAN | 必須フラグ |
| E | display_order | NUMBER | 表示順 |

---

## 🔧 portalApi.js への追加（最小限）

既存ファイルの末尾に以下を追加：

```javascript
// ============================================
// Phase 2: マスタ管理API（master_*.js から呼び出し）
// ============================================

// 検査項目マスタ
// → master_examItem.js で実装

// 検査所見マスタ
// → master_finding.js で実装

// コース管理
// → master_course.js で実装
```

**注意**: 実際のロジックは新規ファイルに実装し、portalApi.jsは薄いラッパーとして使用可能。

---

## 🖥️ UI追加（portal.html）

### タブ追加位置

既存のタブ構造を維持し、新しいタブを追加：

```html
<!-- 既存タブの後に追加 -->
<li class="nav-item">
  <a class="nav-link" id="master-tab" data-bs-toggle="tab" href="#master">
    <i class="bi bi-gear"></i> マスタ管理
  </a>
</li>
```

### マスタ管理タブの内容

```html
<div class="tab-pane fade" id="master">
  <!-- サブタブ: 検査項目 | 検査所見 | コース管理 -->
  <ul class="nav nav-pills mb-3">
    <li class="nav-item">
      <a class="nav-link active" data-subtab="examItem">検査項目</a>
    </li>
    <li class="nav-item">
      <a class="nav-link" data-subtab="finding">検査所見</a>
    </li>
    <li class="nav-item">
      <a class="nav-link" data-subtab="course">コース管理</a>
    </li>
  </ul>
  
  <!-- 各サブタブのコンテンツ -->
  <div id="master-content"></div>
</div>
```

---

## 📋 実装順序

### Phase 2-1: 基盤（1日目）
1. [ ] `master_examItem.js` 作成 - 検査項目CRUD
2. [ ] `M_検査項目` シート作成スクリプト追加（SetupDB.js拡張）
3. [ ] 動作確認

### Phase 2-2: 所見マスタ（2日目）
1. [ ] `master_finding.js` 作成
2. [ ] `M_検査所見` シート作成
3. [ ] 動作確認

### Phase 2-3: コース管理（3日目）
1. [ ] `master_course.js` 作成
2. [ ] `M_コース項目` シート作成
3. [ ] 既存コースデータ移行

### Phase 2-4: UI（4-5日目）
1. [ ] portal.html にマスタ管理タブ追加
2. [ ] js.html にマスタ管理用JavaScript追加
3. [ ] css.html にスタイル追加（必要に応じて）

---

## ✅ 完了チェックリスト

各機能の完了時に確認：

- [ ] 既存機能が壊れていないこと
- [ ] clasp push でエラーがないこと
- [ ] 開発版URLで動作確認
- [ ] エラーハンドリングが適切
- [ ] ログ出力が適切

---

## 📚 参考ドキュメント

- `SYSTEM_DESIGN_SPECIFICATION.md` - 画面設計・機能仕様
- `DATA_STRUCTURE_DESIGN_v2.md` - データ構造詳細
- `CLAUDE.md` - プロジェクトルール（必読）
- `Config.js` - 設定値・カラム定義

---

## 🚫 やってはいけないこと

1. `MasterData.js` の `EXAM_ITEM_MASTER_DATA` を直接編集
2. `portalApi.js` の既存関数を変更
3. `portal.html` の既存タブ構造を変更
4. シート名をハードコード
5. テストなしで clasp deploy

---

## 💡 実装のヒント

### 既存パターンを参考にする

```javascript
// portalApi.js の既存パターン
function portalSearchPatients(criteria) {
  try {
    // 1. バリデーション
    // 2. シート取得
    // 3. データ処理
    // 4. 結果返却
    return { success: true, data: results };
  } catch (error) {
    console.error('portalSearchPatients error:', error);
    return { success: false, error: error.message };
  }
}
```

### CRUD.js の活用

```javascript
// 汎用CRUD関数を使用
const result = genericSearch(sheetName, criteria);
const saved = genericSave(sheetName, data, keyColumn);
```

---

**この指示書に従って実装を進めてください。**
**不明点があれば、実装前に確認してください。**
