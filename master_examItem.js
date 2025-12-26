/**
 * master_examItem.js - 検査項目マスタ管理
 * CDmedical 健診結果管理システム
 *
 * @description 検査項目マスタのCRUD操作
 * @version 1.0.0
 * @date 2025-12-20
 *
 * 依存: Config.js, portalApi.js (getPortalSpreadsheet, safeString, safeNumber)
 * シート: M_検査項目（DATA_STRUCTURE_DESIGN_v2.md 準拠・46列）
 */

// ============================================
// 定数定義
// ============================================

/**
 * シート名
 */
const M_EXAM_ITEM_SHEET_NAME = 'M_検査項目';

/**
 * 列定義（DATA_STRUCTURE_DESIGN_v2.md 準拠・46列）
 */
const M_EXAM_ITEM_COL = {
  // 基本情報
  ITEM_CODE: 0,       // A: 項目コード (PK) 例: 011001
  JLAC10_CODE: 1,     // B: JLAC10コード
  ITEM_NAME: 2,       // C: 項目名
  ITEM_NAME_KANA: 3,  // D: カナ
  CATEGORY: 4,        // E: 区分 (身体情報/検体検査/画像診断)
  SUB_CATEGORY: 5,    // F: 小分類
  DATA_TYPE: 6,       // G: データ型 (numeric/text/select)
  DECIMAL_PLACES: 7,  // H: 小数桁数
  UNIT: 8,            // I: 単位
  VALID_MIN: 9,       // J: 有効範囲下限
  VALID_MAX: 10,      // K: 有効範囲上限
  STD_MALE: 11,       // L: 帳票基準値表記（男性）
  STD_FEMALE: 12,     // M: 帳票基準値表記（女性）
  // 判定基準（男性）N-X列
  CRITERIA_F_LOW_M: 13,
  CRITERIA_E_LOW_M: 14,
  CRITERIA_D_LOW_M: 15,
  CRITERIA_C_LOW_M: 16,
  CRITERIA_B_LOW_M: 17,
  CRITERIA_A_LOW_M: 18,
  CRITERIA_A_HIGH_M: 19,
  CRITERIA_B_HIGH_M: 20,
  CRITERIA_C_HIGH_M: 21,
  CRITERIA_D_HIGH_M: 22,
  CRITERIA_E_HIGH_M: 23,
  // 判定基準（女性）Y-AI列
  CRITERIA_F_LOW_F: 24,
  CRITERIA_E_LOW_F: 25,
  CRITERIA_D_LOW_F: 26,
  CRITERIA_C_LOW_F: 27,
  CRITERIA_B_LOW_F: 28,
  CRITERIA_A_LOW_F: 29,
  CRITERIA_A_HIGH_F: 30,
  CRITERIA_B_HIGH_F: 31,
  CRITERIA_C_HIGH_F: 32,
  CRITERIA_D_HIGH_F: 33,
  CRITERIA_E_HIGH_F: 34,
  // その他
  DEFAULT_VALUE: 35,  // AJ: 初期値
  FINDING_REF: 36,    // AK: 所見参照初期値
  PRICE_DOCK: 37,     // AL: オプション単価（ドック）
  PRICE_REGULAR: 38,  // AM: オプション単価（定期）
  PRICE_EMPLOY: 39,   // AN: オプション単価（雇入）
  PRICE_ROSAI: 40,    // AO: オプション単価（労災）
  LAB_CODE_BML: 41,   // AP: BMLコード
  LAB_CODE_OTHER: 42, // AQ: その他検査会社コード
  DISPLAY_ORDER: 43,  // AR: 表示順
  NOTES: 44,          // AS: 備考
  IS_ACTIVE: 45       // AT: 有効フラグ
};

/**
 * ヘッダー定義（46列）
 */
const M_EXAM_ITEM_HEADERS = [
  '項目コード', 'JLAC10コード', '項目名', 'カナ', '区分', '小分類',
  'データ型', '小数桁数', '単位', '有効下限', '有効上限',
  '基準値表記(男)', '基準値表記(女)',
  // 判定基準（男性）
  'F低_M', 'E低_M', 'D低_M', 'C低_M', 'B低_M', 'A低_M',
  'A高_M', 'B高_M', 'C高_M', 'D高_M', 'E高_M',
  // 判定基準（女性）
  'F低_F', 'E低_F', 'D低_F', 'C低_F', 'B低_F', 'A低_F',
  'A高_F', 'B高_F', 'C高_F', 'D高_F', 'E高_F',
  // その他
  '初期値', '所見参照', '単価_ドック', '単価_定期', '単価_雇入', '単価_労災',
  'BMLコード', 'その他コード', '表示順', '備考', '有効'
];

// ============================================
// シート作成・初期化
// ============================================

/**
 * M_検査項目シートを作成（既存削除→新規作成）
 * GASエディタから手動実行
 *
 * @returns {Object} {success, message}
 */
function setupExamItemMasterSheet() {
  try {
    const ss = getPortalSpreadsheet();
    const sheetName = M_EXAM_ITEM_SHEET_NAME;

    // 既存シート削除
    const existing = ss.getSheetByName(sheetName);
    if (existing) {
      ss.deleteSheet(existing);
      console.log('既存シート削除: ' + sheetName);
    }

    // 新規作成
    const sheet = ss.insertSheet(sheetName);

    // ヘッダー設定
    sheet.getRange(1, 1, 1, M_EXAM_ITEM_HEADERS.length).setValues([M_EXAM_ITEM_HEADERS]);

    // ヘッダー行を固定
    sheet.setFrozenRows(1);

    // ヘッダー行のスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, M_EXAM_ITEM_HEADERS.length);
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    // 列幅設定（主要列のみ）
    sheet.setColumnWidth(1, 100);  // 項目コード
    sheet.setColumnWidth(3, 150);  // 項目名
    sheet.setColumnWidth(5, 100);  // 区分
    sheet.setColumnWidth(6, 100);  // 小分類
    sheet.setColumnWidth(9, 60);   // 単位

    // 項目コード列（A列）をテキスト形式に設定（先頭0の欠落防止）
    sheet.getRange('A:A').setNumberFormat('@');
    // JLAC10コード列（B列）もテキスト形式
    sheet.getRange('B:B').setNumberFormat('@');
    // BMLコード列（AP列）もテキスト形式
    sheet.getRange('AP:AP').setNumberFormat('@');

    console.log('M_検査項目シート作成完了: ' + M_EXAM_ITEM_HEADERS.length + '列');

    return {
      success: true,
      message: 'M_検査項目シート作成完了',
      columnCount: M_EXAM_ITEM_HEADERS.length
    };

  } catch (error) {
    console.error('setupExamItemMasterSheet error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================
// CRUD: Read（一覧取得）
// ============================================

/**
 * 検査項目一覧を取得
 *
 * @param {Object} criteria - 検索条件
 *   - category: カテゴリでフィルタ（身体情報/検体検査/画像診断）
 *   - subCategory: 小分類でフィルタ
 *   - keyword: 項目名で部分一致検索
 *   - activeOnly: 有効フラグがtrueのみ（デフォルト: true）
 *   - limit: 最大取得件数（デフォルト: 500）
 * @returns {Object} {success, data, count, _debug}
 */
function portalGetExamItems(criteria) {
  const debugInfo = {
    version: '2025-12-20-v1',
    criteriaReceived: criteria ? JSON.stringify(criteria) : 'null',
    timestamp: new Date().toISOString()
  };

  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName(M_EXAM_ITEM_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        error: M_EXAM_ITEM_SHEET_NAME + 'シートが見つかりません。setupExamItemMasterSheet()を実行してください。',
        data: [],
        _debug: debugInfo
      };
    }

    const data = sheet.getDataRange().getValues();
    debugInfo.totalRows = data.length;

    if (data.length <= 1) {
      return {
        success: true,
        data: [],
        count: 0,
        _debug: debugInfo
      };
    }

    // 検索条件の取得
    const filterCategory = criteria && criteria.category ? String(criteria.category).trim() : '';
    const filterSubCategory = criteria && criteria.subCategory ? String(criteria.subCategory).trim() : '';
    const filterKeyword = criteria && criteria.keyword ? String(criteria.keyword).trim().toLowerCase() : '';
    const activeOnly = criteria && criteria.activeOnly !== undefined ? criteria.activeOnly : true;
    const limit = criteria && criteria.limit ? Number(criteria.limit) : 500;

    debugInfo.filters = { filterCategory, filterSubCategory, filterKeyword, activeOnly, limit };

    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 空行スキップ
      if (!row[M_EXAM_ITEM_COL.ITEM_CODE]) continue;

      // 有効フラグチェック
      if (activeOnly) {
        const isActive = row[M_EXAM_ITEM_COL.IS_ACTIVE];
        // 空、false、0、"FALSE" の場合はスキップ
        if (isActive === false || isActive === 0 || isActive === 'FALSE' || isActive === '') {
          continue;
        }
      }

      // カテゴリフィルタ
      if (filterCategory) {
        const category = safeString(row[M_EXAM_ITEM_COL.CATEGORY]);
        if (category !== filterCategory) continue;
      }

      // 小分類フィルタ
      if (filterSubCategory) {
        const subCategory = safeString(row[M_EXAM_ITEM_COL.SUB_CATEGORY]);
        if (subCategory !== filterSubCategory) continue;
      }

      // キーワード検索（項目名の部分一致）
      if (filterKeyword) {
        const itemName = safeString(row[M_EXAM_ITEM_COL.ITEM_NAME]).toLowerCase();
        const itemKana = safeString(row[M_EXAM_ITEM_COL.ITEM_NAME_KANA]).toLowerCase();
        if (!itemName.includes(filterKeyword) && !itemKana.includes(filterKeyword)) {
          continue;
        }
      }

      // 結果に追加
      results.push(convertRowToExamItem(row, i + 1));

      // 最大件数制限
      if (results.length >= limit) break;
    }

    debugInfo.resultsCount = results.length;

    return {
      success: true,
      data: results,
      count: results.length,
      _debug: debugInfo
    };

  } catch (error) {
    console.error('portalGetExamItems error:', error);
    debugInfo.error = error.message;
    return {
      success: false,
      error: error.message,
      data: [],
      _debug: debugInfo
    };
  }
}

// ============================================
// CRUD: Read（1件取得）
// ============================================

/**
 * 検査項目を1件取得
 *
 * @param {string} itemCode - 項目コード
 * @returns {Object} {success, data}
 */
function portalGetExamItemById(itemCode) {
  const debugInfo = {
    version: '2025-12-20-v1',
    itemCode: itemCode,
    timestamp: new Date().toISOString()
  };

  try {
    if (!itemCode) {
      return {
        success: false,
        error: '項目コードが指定されていません',
        _debug: debugInfo
      };
    }

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName(M_EXAM_ITEM_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        error: M_EXAM_ITEM_SHEET_NAME + 'シートが見つかりません',
        _debug: debugInfo
      };
    }

    const data = sheet.getDataRange().getValues();
    const searchCode = String(itemCode).trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const code = safeString(row[M_EXAM_ITEM_COL.ITEM_CODE]);

      if (code === searchCode) {
        debugInfo.foundAtRow = i + 1;
        return {
          success: true,
          data: convertRowToExamItem(row, i + 1),
          _debug: debugInfo
        };
      }
    }

    return {
      success: false,
      error: '項目コード ' + searchCode + ' が見つかりません',
      _debug: debugInfo
    };

  } catch (error) {
    console.error('portalGetExamItemById error:', error);
    debugInfo.error = error.message;
    return {
      success: false,
      error: error.message,
      _debug: debugInfo
    };
  }
}

// ============================================
// CRUD: Create/Update（保存）
// ============================================

/**
 * 検査項目を保存（新規/更新）
 *
 * @param {Object} itemData - 検査項目データ
 *   - item_code: 項目コード（必須）
 *   - item_name: 項目名（必須）
 *   - category: 区分
 *   - その他の項目...
 * @returns {Object} {success, data, message, isNew}
 */
function portalSaveExamItem(itemData) {
  const debugInfo = {
    version: '2025-12-20-v1',
    timestamp: new Date().toISOString()
  };

  try {
    // バリデーション
    if (!itemData) {
      return { success: false, error: 'データが指定されていません', _debug: debugInfo };
    }

    if (!itemData.item_code) {
      return { success: false, error: '項目コードは必須です', _debug: debugInfo };
    }

    if (!itemData.item_name) {
      return { success: false, error: '項目名は必須です', _debug: debugInfo };
    }

    debugInfo.itemCode = itemData.item_code;

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName(M_EXAM_ITEM_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        error: M_EXAM_ITEM_SHEET_NAME + 'シートが見つかりません。setupExamItemMasterSheet()を実行してください。',
        _debug: debugInfo
      };
    }

    const data = sheet.getDataRange().getValues();
    const searchCode = String(itemData.item_code).trim();

    // 既存データ検索
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      const code = safeString(data[i][M_EXAM_ITEM_COL.ITEM_CODE]);
      if (code === searchCode) {
        existingRow = i + 1;
        break;
      }
    }

    // 行データを作成
    const rowData = convertExamItemToRow(itemData);

    if (existingRow > 0) {
      // 更新
      const range = sheet.getRange(existingRow, 1, 1, rowData.length);
      // 項目コード列をテキスト形式に設定してから値をセット
      sheet.getRange(existingRow, 1).setNumberFormat('@');
      range.setValues([rowData]);
      debugInfo.action = 'update';
      debugInfo.rowIndex = existingRow;

      return {
        success: true,
        data: convertRowToExamItem(rowData, existingRow),
        message: '検査項目を更新しました: ' + searchCode,
        isNew: false,
        _debug: debugInfo
      };

    } else {
      // 新規追加
      const newRowIndex = sheet.getLastRow() + 1;
      // 項目コード列をテキスト形式に設定してから値をセット
      sheet.getRange(newRowIndex, 1).setNumberFormat('@');
      sheet.getRange(newRowIndex, 1, 1, rowData.length).setValues([rowData]);
      debugInfo.action = 'create';
      debugInfo.rowIndex = newRowIndex;

      return {
        success: true,
        data: convertRowToExamItem(rowData, newRowIndex),
        message: '検査項目を追加しました: ' + searchCode,
        isNew: true,
        _debug: debugInfo
      };
    }

  } catch (error) {
    console.error('portalSaveExamItem error:', error);
    debugInfo.error = error.message;
    return {
      success: false,
      error: error.message,
      _debug: debugInfo
    };
  }
}

// ============================================
// CRUD: Delete（削除）
// ============================================

/**
 * 検査項目を削除（論理削除）
 *
 * @param {string} itemCode - 項目コード
 * @param {boolean} physical - 物理削除するか（デフォルト: false=論理削除）
 * @returns {Object} {success, message}
 */
function portalDeleteExamItem(itemCode, physical) {
  const debugInfo = {
    version: '2025-12-20-v1',
    itemCode: itemCode,
    physical: physical || false,
    timestamp: new Date().toISOString()
  };

  try {
    if (!itemCode) {
      return { success: false, error: '項目コードが指定されていません', _debug: debugInfo };
    }

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName(M_EXAM_ITEM_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        error: M_EXAM_ITEM_SHEET_NAME + 'シートが見つかりません',
        _debug: debugInfo
      };
    }

    const data = sheet.getDataRange().getValues();
    const searchCode = String(itemCode).trim();

    // 対象行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const code = safeString(data[i][M_EXAM_ITEM_COL.ITEM_CODE]);
      if (code === searchCode) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow < 0) {
      return {
        success: false,
        error: '項目コード ' + searchCode + ' が見つかりません',
        _debug: debugInfo
      };
    }

    debugInfo.targetRow = targetRow;

    if (physical) {
      // 物理削除
      sheet.deleteRow(targetRow);
      debugInfo.action = 'physical_delete';

      return {
        success: true,
        message: '検査項目を物理削除しました: ' + searchCode,
        _debug: debugInfo
      };

    } else {
      // 論理削除（有効フラグをfalseに）
      const isActiveCol = M_EXAM_ITEM_COL.IS_ACTIVE + 1; // 1-indexed
      sheet.getRange(targetRow, isActiveCol).setValue(false);
      debugInfo.action = 'logical_delete';

      return {
        success: true,
        message: '検査項目を論理削除しました: ' + searchCode,
        _debug: debugInfo
      };
    }

  } catch (error) {
    console.error('portalDeleteExamItem error:', error);
    debugInfo.error = error.message;
    return {
      success: false,
      error: error.message,
      _debug: debugInfo
    };
  }
}

// ============================================
// ユーティリティ関数
// ============================================

/**
 * 行データをオブジェクトに変換
 *
 * @param {Array} row - 行データ
 * @param {number} rowIndex - 行番号（1-indexed）
 * @returns {Object} 検査項目オブジェクト
 */
function convertRowToExamItem(row, rowIndex) {
  return {
    item_code: safeString(row[M_EXAM_ITEM_COL.ITEM_CODE]),
    jlac10_code: safeString(row[M_EXAM_ITEM_COL.JLAC10_CODE]),
    item_name: safeString(row[M_EXAM_ITEM_COL.ITEM_NAME]),
    item_name_kana: safeString(row[M_EXAM_ITEM_COL.ITEM_NAME_KANA]),
    category: safeString(row[M_EXAM_ITEM_COL.CATEGORY]),
    sub_category: safeString(row[M_EXAM_ITEM_COL.SUB_CATEGORY]),
    data_type: safeString(row[M_EXAM_ITEM_COL.DATA_TYPE]),
    decimal_places: safeNumber(row[M_EXAM_ITEM_COL.DECIMAL_PLACES]),
    unit: safeString(row[M_EXAM_ITEM_COL.UNIT]),
    valid_min: safeNumber(row[M_EXAM_ITEM_COL.VALID_MIN]),
    valid_max: safeNumber(row[M_EXAM_ITEM_COL.VALID_MAX]),
    std_male: safeString(row[M_EXAM_ITEM_COL.STD_MALE]),
    std_female: safeString(row[M_EXAM_ITEM_COL.STD_FEMALE]),
    // 判定基準（男性）
    criteria_male: {
      f_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_F_LOW_M]),
      e_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_E_LOW_M]),
      d_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_D_LOW_M]),
      c_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_C_LOW_M]),
      b_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_B_LOW_M]),
      a_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_A_LOW_M]),
      a_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_A_HIGH_M]),
      b_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_B_HIGH_M]),
      c_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_C_HIGH_M]),
      d_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_D_HIGH_M]),
      e_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_E_HIGH_M])
    },
    // 判定基準（女性）
    criteria_female: {
      f_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_F_LOW_F]),
      e_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_E_LOW_F]),
      d_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_D_LOW_F]),
      c_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_C_LOW_F]),
      b_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_B_LOW_F]),
      a_low: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_A_LOW_F]),
      a_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_A_HIGH_F]),
      b_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_B_HIGH_F]),
      c_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_C_HIGH_F]),
      d_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_D_HIGH_F]),
      e_high: safeNumber(row[M_EXAM_ITEM_COL.CRITERIA_E_HIGH_F])
    },
    // その他
    default_value: safeString(row[M_EXAM_ITEM_COL.DEFAULT_VALUE]),
    finding_ref: safeString(row[M_EXAM_ITEM_COL.FINDING_REF]),
    price_dock: safeNumber(row[M_EXAM_ITEM_COL.PRICE_DOCK]),
    price_regular: safeNumber(row[M_EXAM_ITEM_COL.PRICE_REGULAR]),
    price_employ: safeNumber(row[M_EXAM_ITEM_COL.PRICE_EMPLOY]),
    price_rosai: safeNumber(row[M_EXAM_ITEM_COL.PRICE_ROSAI]),
    lab_code_bml: safeString(row[M_EXAM_ITEM_COL.LAB_CODE_BML]),
    lab_code_other: safeString(row[M_EXAM_ITEM_COL.LAB_CODE_OTHER]),
    display_order: safeNumber(row[M_EXAM_ITEM_COL.DISPLAY_ORDER]),
    notes: safeString(row[M_EXAM_ITEM_COL.NOTES]),
    is_active: row[M_EXAM_ITEM_COL.IS_ACTIVE] !== false && row[M_EXAM_ITEM_COL.IS_ACTIVE] !== 'FALSE',
    _rowIndex: rowIndex
  };
}

/**
 * オブジェクトを行データに変換
 *
 * @param {Object} item - 検査項目オブジェクト
 * @returns {Array} 行データ
 */
function convertExamItemToRow(item) {
  const criteriaMale = item.criteria_male || {};
  const criteriaFemale = item.criteria_female || {};

  return [
    item.item_code || '',
    item.jlac10_code || '',
    item.item_name || '',
    item.item_name_kana || '',
    item.category || '',
    item.sub_category || '',
    item.data_type || 'numeric',
    item.decimal_places || '',
    item.unit || '',
    item.valid_min || '',
    item.valid_max || '',
    item.std_male || '',
    item.std_female || '',
    // 判定基準（男性）
    criteriaMale.f_low || '',
    criteriaMale.e_low || '',
    criteriaMale.d_low || '',
    criteriaMale.c_low || '',
    criteriaMale.b_low || '',
    criteriaMale.a_low || '',
    criteriaMale.a_high || '',
    criteriaMale.b_high || '',
    criteriaMale.c_high || '',
    criteriaMale.d_high || '',
    criteriaMale.e_high || '',
    // 判定基準（女性）
    criteriaFemale.f_low || '',
    criteriaFemale.e_low || '',
    criteriaFemale.d_low || '',
    criteriaFemale.c_low || '',
    criteriaFemale.b_low || '',
    criteriaFemale.a_low || '',
    criteriaFemale.a_high || '',
    criteriaFemale.b_high || '',
    criteriaFemale.c_high || '',
    criteriaFemale.d_high || '',
    criteriaFemale.e_high || '',
    // その他
    item.default_value || '',
    item.finding_ref || '',
    item.price_dock || '',
    item.price_regular || '',
    item.price_employ || '',
    item.price_rosai || '',
    item.lab_code_bml || '',
    item.lab_code_other || '',
    item.display_order || '',
    item.notes || '',
    item.is_active !== false ? true : false
  ];
}

// ============================================
// カテゴリ・小分類取得
// ============================================

/**
 * カテゴリ一覧を取得
 *
 * @returns {Object} {success, data}
 */
function portalGetExamItemCategories() {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName(M_EXAM_ITEM_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        error: M_EXAM_ITEM_SHEET_NAME + 'シートが見つかりません',
        data: []
      };
    }

    const data = sheet.getDataRange().getValues();
    const categories = new Set();

    for (let i = 1; i < data.length; i++) {
      const category = safeString(data[i][M_EXAM_ITEM_COL.CATEGORY]);
      if (category) {
        categories.add(category);
      }
    }

    return {
      success: true,
      data: Array.from(categories).sort()
    };

  } catch (error) {
    console.error('portalGetExamItemCategories error:', error);
    return {
      success: false,
      error: error.message,
      data: []
    };
  }
}

/**
 * 小分類一覧を取得
 *
 * @param {string} category - カテゴリ（省略時は全て）
 * @returns {Object} {success, data}
 */
function portalGetExamItemSubCategories(category) {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName(M_EXAM_ITEM_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        error: M_EXAM_ITEM_SHEET_NAME + 'シートが見つかりません',
        data: []
      };
    }

    const data = sheet.getDataRange().getValues();
    const subCategories = new Set();
    const filterCategory = category ? String(category).trim() : '';

    for (let i = 1; i < data.length; i++) {
      const rowCategory = safeString(data[i][M_EXAM_ITEM_COL.CATEGORY]);
      const subCategory = safeString(data[i][M_EXAM_ITEM_COL.SUB_CATEGORY]);

      if (subCategory) {
        if (!filterCategory || rowCategory === filterCategory) {
          subCategories.add(subCategory);
        }
      }
    }

    return {
      success: true,
      data: Array.from(subCategories).sort()
    };

  } catch (error) {
    console.error('portalGetExamItemSubCategories error:', error);
    return {
      success: false,
      error: error.message,
      data: []
    };
  }
}

// ============================================
// 人間ドック項目マッピング照合
// ============================================

/**
 * 人間ドックテンプレート全項目リスト
 * 1221_template_new_default.xlsm 準拠
 *
 * input_type:
 *   - 'auto': BML CSVから自動転記
 *   - 'manual': 手入力
 *   - 'calc': 計算値（他項目から算出）
 *
 * page: テンプレートのページ番号
 */
const HUMAN_DOCK_ALL_ITEMS = {
  // =============================================
  // 2ページ: 所見手入力（画像診断・検査所見）
  // =============================================
  page2_findings: [
    { item_code: '020001', display: '診察所見', category: '画像診断', sub_category: '診察', input_type: 'manual', data_type: 'text' },
    { item_code: '020002', display: '胸部X線', category: '画像診断', sub_category: 'X線', input_type: 'manual', data_type: 'text' },
    { item_code: '020003', display: '腹部X線', category: '画像診断', sub_category: 'X線', input_type: 'manual', data_type: 'text' },
    { item_code: '020004', display: '心電図', category: '画像診断', sub_category: '生理検査', input_type: 'manual', data_type: 'text' },
    { item_code: '020005', display: '上部消化管内視鏡検査(胃カメラ)', category: '画像診断', sub_category: '内視鏡', input_type: 'manual', data_type: 'text' },
    { item_code: '020006', display: '下部消化管内視鏡検査(大腸カメラ)', category: '画像診断', sub_category: '内視鏡', input_type: 'manual', data_type: 'text' },
    { item_code: '020007', display: '腹部超音波', category: '画像診断', sub_category: '超音波', input_type: 'manual', data_type: 'text' },
    { item_code: '020008', display: '頸部超音波', category: '画像診断', sub_category: '超音波', input_type: 'manual', data_type: 'text' },
    { item_code: '020009', display: '心臓超音波検査', category: '画像診断', sub_category: '超音波', input_type: 'manual', data_type: 'text' },
    { item_code: '020010', display: '甲状腺超音波検査', category: '画像診断', sub_category: '超音波', input_type: 'manual', data_type: 'text' },
    { item_code: '020011', display: '胸部CT', category: '画像診断', sub_category: 'CT', input_type: 'manual', data_type: 'text' },
    { item_code: '020012', display: '腹部CT', category: '画像診断', sub_category: 'CT', input_type: 'manual', data_type: 'text' },
    { item_code: '020013', display: '腹部MRI(肝胆膵)+MRCP', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020014', display: '脳MRI', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020015', display: '脳MRA', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020016', display: 'DWI', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020017', display: '頸動脈MRA', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020018', display: '上腹部MRI', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020019', display: 'MRCP', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020020', display: '子宮・卵巣MRI', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020021', display: 'マンモMRI', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' },
    { item_code: '020022', display: '前立腺MRI', category: '画像診断', sub_category: 'MRI', input_type: 'manual', data_type: 'text' }
  ],

  // =============================================
  // 3ページ: 所見手入力（身体計測・生理検査）
  // =============================================
  page3_physical: [
    { item_code: '011001', display: '身長', category: '身体情報', sub_category: '身体計測', input_type: 'manual', data_type: 'numeric', unit: 'cm', decimal: 1 },
    { item_code: '011002', display: '体重', category: '身体情報', sub_category: '身体計測', input_type: 'manual', data_type: 'numeric', unit: 'kg', decimal: 1 },
    { item_code: '011003', display: '標準体重', category: '身体情報', sub_category: '身体計測', input_type: 'calc', data_type: 'numeric', unit: 'kg', decimal: 1 },
    { item_code: '011004', display: 'BMI', category: '身体情報', sub_category: '身体計測', input_type: 'calc', data_type: 'numeric', unit: '', decimal: 1 },
    { item_code: '011005', display: '体脂肪率', category: '身体情報', sub_category: '身体計測', input_type: 'manual', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '011006', display: '腹囲', category: '身体情報', sub_category: '身体計測', input_type: 'manual', data_type: 'numeric', unit: 'cm', decimal: 1 },
    { item_code: '012001', display: '血圧1回目(収縮期)', category: '身体情報', sub_category: '血圧', input_type: 'manual', data_type: 'numeric', unit: 'mmHg', decimal: 0 },
    { item_code: '012002', display: '血圧1回目(拡張期)', category: '身体情報', sub_category: '血圧', input_type: 'manual', data_type: 'numeric', unit: 'mmHg', decimal: 0 },
    { item_code: '012003', display: '血圧2回目(収縮期)', category: '身体情報', sub_category: '血圧', input_type: 'manual', data_type: 'numeric', unit: 'mmHg', decimal: 0 },
    { item_code: '012004', display: '血圧2回目(拡張期)', category: '身体情報', sub_category: '血圧', input_type: 'manual', data_type: 'numeric', unit: 'mmHg', decimal: 0 },
    { item_code: '013001', display: '視力 裸眼 右', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'numeric', unit: '', decimal: 1 },
    { item_code: '013002', display: '視力 裸眼 左', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'numeric', unit: '', decimal: 1 },
    { item_code: '013003', display: '視力 矯正 右', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'numeric', unit: '', decimal: 1 },
    { item_code: '013004', display: '視力 矯正 左', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'numeric', unit: '', decimal: 1 },
    { item_code: '013005', display: '眼圧 右', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'numeric', unit: 'mmHg', decimal: 0 },
    { item_code: '013006', display: '眼圧 左', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'numeric', unit: 'mmHg', decimal: 0 },
    { item_code: '013007', display: '眼底 右', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'text' },
    { item_code: '013008', display: '眼底 左', category: '身体情報', sub_category: '眼科', input_type: 'manual', data_type: 'text' },
    { item_code: '014001', display: '聴力右 1000Hz', category: '身体情報', sub_category: '聴力', input_type: 'manual', data_type: 'select' },
    { item_code: '014002', display: '聴力左 1000Hz', category: '身体情報', sub_category: '聴力', input_type: 'manual', data_type: 'select' },
    { item_code: '014003', display: '聴力右 4000Hz', category: '身体情報', sub_category: '聴力', input_type: 'manual', data_type: 'select' },
    { item_code: '014004', display: '聴力左 4000Hz', category: '身体情報', sub_category: '聴力', input_type: 'manual', data_type: 'select' },
    { item_code: '015001', bml_code: '0002151',display: '便ヘモグロビン1回目', category: '検体検査', sub_category: '便検査', input_type: 'auto', data_type: 'select' },
    { item_code: '015002',bml_code: '0002152', display: '便ヘモグロビン2回目', category: '検体検査', sub_category: '便検査', input_type: 'auto', data_type: 'select' }
  ],

  // =============================================
  // 3ページ: 尿検査（自動転記）
  // =============================================
  page3_urine: [
    { item_code: '031001', bml_code: '0000911', display: '尿糖(定性)', category: '検体検査', sub_category: '尿検査', input_type: 'auto', data_type: 'select' },
    { item_code: '031002', bml_code: '0000703', display: '尿蛋白', category: '検体検査', sub_category: '尿検査', input_type: 'auto', data_type: 'select' },
    { item_code: '031003', bml_code: '0000705', display: '尿潜血', category: '検体検査', sub_category: '尿検査', input_type: 'auto', data_type: 'select' },
    { item_code: '031004', bml_code: '0000707', display: 'ウロビリノーゲン', category: '検体検査', sub_category: '尿検査', input_type: 'auto', data_type: 'select' },
    { item_code: '031005', bml_code: '0000709', display: '尿PH', category: '検体検査', sub_category: '尿検査', input_type: 'auto', data_type: 'numeric' },
    { item_code: '031006', bml_code: '0000711', display: '尿ビリルビン', category: '検体検査', sub_category: '尿検査', input_type: 'auto', data_type: 'select' },
    { item_code: '031007', bml_code: '0000713', display: 'アセトン体', category: '検体検査', sub_category: '尿検査', input_type: 'auto', data_type: 'select' },
    { item_code: '031008', bml_code: '0000721', display: '尿沈渣白血球', category: '検体検査', sub_category: '尿沈渣', input_type: 'auto', data_type: 'text' },
    { item_code: '031009', bml_code: '0000723', display: '尿沈渣赤血球', category: '検体検査', sub_category: '尿沈渣', input_type: 'auto', data_type: 'text' },
    { item_code: '031010', bml_code: '0000725', display: '尿沈渣扁平上皮', category: '検体検査', sub_category: '尿沈渣', input_type: 'auto', data_type: 'text' },
    { item_code: '031011', bml_code: '0000727', display: '尿沈渣細菌', category: '検体検査', sub_category: '尿沈渣', input_type: 'auto', data_type: 'text' }
  ],

  // =============================================
  // 4ページ: 血液検査（自動転記）- 血液学
  // =============================================
  page4_blood_cell: [
    { item_code: '041001', bml_code: '0000301', display: '白血球数', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: '/μL', decimal: 0 },
    { item_code: '041002', bml_code: '0000302', display: '赤血球数', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: '万/μL', decimal: 0 },
    { item_code: '041003', bml_code: '0000303', display: '血色素量(ヘモグロビン)', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: 'g/dL', decimal: 1 },
    { item_code: '041004', bml_code: '0000304', display: 'ヘマトクリット', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '041005', bml_code: '0000308', display: '血小板(PLT)', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: '万/μL', decimal: 1 },
    { item_code: '041006', bml_code: '0000305', display: 'MCV', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: 'fL', decimal: 1 },
    { item_code: '041007', bml_code: '0000306', display: 'MCH', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: 'pg', decimal: 1 },
    { item_code: '041008', bml_code: '0000307', display: 'MCHC', category: '検体検査', sub_category: '血液学検査', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '041009', bml_code: '0001885', display: 'NEUT(好中球)', category: '検体検査', sub_category: '白血球像', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '041010', bml_code: '0001881', display: 'BASO(好塩基球)', category: '検体検査', sub_category: '白血球像', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '041011', bml_code: '0001882', display: 'EOS(好酸球)', category: '検体検査', sub_category: '白血球像', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '041012', bml_code: '0001889', display: 'LYMPHO(リンパ球)', category: '検体検査', sub_category: '白血球像', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '041013', bml_code: '0001886', display: 'MONO(単球)', category: '検体検査', sub_category: '白血球像', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 }
  ],

  // =============================================
  // 4ページ: 血液検査（自動転記）- 凝固
  // =============================================
  page4_coagulation: [
    { item_code: '042001', bml_code: '0002091', display: 'プロトロンビン時間(PT)', category: '検体検査', sub_category: '凝固', input_type: 'auto', data_type: 'numeric', unit: '秒', decimal: 1 },
    { item_code: '042002', bml_code: '0002093', display: '活性化部分トロンボプラスチン時間(APTT)', category: '検体検査', sub_category: '凝固', input_type: 'auto', data_type: 'numeric', unit: '秒', decimal: 1 }
  ],

  // =============================================
  // 4ページ: 血液検査（自動転記）- 生化学
  // =============================================
  page4_biochemistry: [
    { item_code: '043001', bml_code: '0000401', display: '総蛋白(TP)', category: '検体検査', sub_category: '蛋白', input_type: 'auto', data_type: 'numeric', unit: 'g/dL', decimal: 1 },
    { item_code: '043002', bml_code: '0000417', display: 'アルブミン', category: '検体検査', sub_category: '蛋白', input_type: 'auto', data_type: 'numeric', unit: 'g/dL', decimal: 1 },
    { item_code: '043003', bml_code: '0000481', display: 'AST(GOT)', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '043004', bml_code: '0000482', display: 'ALT(GPT)', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '043005', bml_code: '0000484', display: 'γ-GTP', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '043006', bml_code: '0013067', display: 'ALP(IFCC)', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '043007', bml_code: '0000497', display: 'LDH(乳酸脱水素酵素)', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '043008', bml_code: '0000495', display: 'コリンエステラーゼ(Ch-E)', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '043009', bml_code: '0000501', display: '血清アミラーゼ', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '043010', bml_code: '0000472', display: '総ビリルビン(T-Bil)', category: '検体検査', sub_category: '肝胆膵機能', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 1 }
  ],

  // =============================================
  // 4ページ: 血液検査（自動転記）- 脂質
  // =============================================
  page4_lipid: [
    { item_code: '044001', bml_code: '0000453', display: '総コレステロール', category: '検体検査', sub_category: '脂質検査', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 0 },
    { item_code: '044002', bml_code: '0000454', display: '中性脂肪(TG)', category: '検体検査', sub_category: '脂質検査', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 0 },
    { item_code: '044003', bml_code: '0000460', display: 'HDLコレステロール', category: '検体検査', sub_category: '脂質検査', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 0 },
    { item_code: '044004', bml_code: '0000410', display: 'LDLコレステロール', category: '検体検査', sub_category: '脂質検査', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 0 },
    { item_code: '044005', display: 'non HDLコレステロール', category: '検体検査', sub_category: '脂質検査', input_type: 'calc', data_type: 'numeric', unit: 'mg/dL', decimal: 0 }
  ],

  // =============================================
  // 4ページ: 血液検査（自動転記）- 糖代謝
  // =============================================
  page4_diabetes: [
    { item_code: '045001', bml_code: '0000503', display: '空腹時血糖', category: '検体検査', sub_category: '糖代謝', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 0 },
    { item_code: '045002', bml_code: '0003317', display: 'HbA1c', category: '検体検査', sub_category: '糖代謝', input_type: 'auto', data_type: 'numeric', unit: '%', decimal: 1 }
  ],

  // =============================================
  // 4ページ: 血液検査（自動転記）- 腎機能
  // =============================================
  page4_kidney: [
    { item_code: '046001', bml_code: '0000413', display: 'クレアチニン', category: '検体検査', sub_category: '腎機能', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 2 },
    { item_code: '046002', bml_code: '0000491', display: '尿素窒素(BUN)', category: '検体検査', sub_category: '腎機能', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 1 },
    { item_code: '046003', bml_code: '0002696', display: 'eGFR', category: '検体検査', sub_category: '腎機能', input_type: 'auto', data_type: 'numeric', unit: 'mL/min/1.73m²', decimal: 1 },
    { item_code: '046004', bml_code: '0000407', display: '尿酸(UA)', category: '検体検査', sub_category: '腎機能', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 1 }
  ],

  // =============================================
  // 4ページ: 血液検査（自動転記）- その他生化学
  // =============================================
  page4_other_biochem: [
    { item_code: '047001', bml_code: '0003845', display: 'クレアチニンキナーゼ(CK)', category: '検体検査', sub_category: '酵素', input_type: 'auto', data_type: 'numeric', unit: 'U/L', decimal: 0 },
    { item_code: '047002', bml_code: '0003550', display: 'ナトリウム(Na)', category: '検体検査', sub_category: '電解質', input_type: 'auto', data_type: 'numeric', unit: 'mEq/L', decimal: 0 },
    { item_code: '047003', bml_code: '0000421', display: 'カリウム(K)', category: '検体検査', sub_category: '電解質', input_type: 'auto', data_type: 'numeric', unit: 'mEq/L', decimal: 1 },
    { item_code: '047004', bml_code: '0000425', display: 'クロール(Cl)', category: '検体検査', sub_category: '電解質', input_type: 'auto', data_type: 'numeric', unit: 'mEq/L', decimal: 0 },
    { item_code: '047005', bml_code: '0000429', display: 'カルシウム(Ca)', category: '検体検査', sub_category: '電解質', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 1 },
    { item_code: '047006', bml_code: '0000433', display: '血清鉄(Fe)', category: '検体検査', sub_category: '鉄代謝', input_type: 'auto', data_type: 'numeric', unit: 'μg/dL', decimal: 0 },
    { item_code: '047007', bml_code: '0000435', display: '総鉄結合能(TIBC)', category: '検体検査', sub_category: '鉄代謝', input_type: 'auto', data_type: 'numeric', unit: 'μg/dL', decimal: 0 },
    { item_code: '047008', bml_code: '0000658', display: 'CRP定量', category: '検体検査', sub_category: '炎症マーカー', input_type: 'auto', data_type: 'numeric', unit: 'mg/dL', decimal: 2 },
    { item_code: '047009', bml_code: '0002117', display: 'リウマトイド因子定量', category: '検体検査', sub_category: '自己抗体', input_type: 'auto', data_type: 'numeric', unit: 'IU/mL', decimal: 0 }
  ],

  // =============================================
  // 5ページ: 腫瘍マーカー
  // =============================================
  page5_tumor_markers: [
    { item_code: '051001', bml_code: '0005005', display: 'PSA', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'ng/mL', decimal: 2 },
    { item_code: '051002', bml_code: '0005001', display: 'CEA', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'ng/mL', decimal: 1 },
    { item_code: '051003', bml_code: '0005003', display: 'CA19-9', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'U/mL', decimal: 1 },
    { item_code: '051004', bml_code: '0005007', display: 'CA125', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'U/mL', decimal: 1 },
    { item_code: '051005', bml_code: '0005009', display: 'NSE', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'ng/mL', decimal: 1 },
    { item_code: '051006', bml_code: '0005011', display: 'エラスターゼ1', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'ng/dL', decimal: 0 },
    { item_code: '051007', bml_code: '0005013', display: '抗p53抗体', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'U/mL', decimal: 1 },
    { item_code: '051008', bml_code: '0005015', display: 'CYFRA21-1', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'ng/mL', decimal: 1 },
    { item_code: '051009', bml_code: '0005017', display: 'SCC', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'ng/mL', decimal: 1 },
    { item_code: '051010', bml_code: '0005019', display: 'ProGRP', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'pg/mL', decimal: 1 },
    { item_code: '051011', bml_code: '0005021', display: 'AFP', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'ng/mL', decimal: 1 },
    { item_code: '051012', bml_code: '0005023', display: 'PIVKA II', category: '検体検査', sub_category: '腫瘍マーカー', input_type: 'auto', data_type: 'numeric', unit: 'mAU/mL', decimal: 0 }
  ],

  // =============================================
  // 5ページ: 感染症
  // =============================================
  page5_infectious: [
    { item_code: '052001', bml_code: '0006001', display: 'TPHA', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' },
    { item_code: '052002', bml_code: '0006003', display: 'RPR定性', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' },
    { item_code: '052003', bml_code: '0006005', display: 'HBs抗原', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' },
    { item_code: '052004', bml_code: '0006007', display: 'HBs抗体', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' },
    { item_code: '052005', bml_code: '0006009', display: 'HCV抗体', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' },
    { item_code: '052006', bml_code: '0006011', display: 'HIV-1抗体', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' }
  ],

  // =============================================
  // 5ページ: 心臓・甲状腺
  // =============================================
  page5_cardiac_thyroid: [
    { item_code: '053001', bml_code: '0004001', display: 'NT-proBNP', category: '検体検査', sub_category: '心臓マーカー', input_type: 'auto', data_type: 'numeric', unit: 'pg/mL', decimal: 0 },
    { item_code: '053002', bml_code: '0004003', display: 'FT3', category: '検体検査', sub_category: '甲状腺', input_type: 'auto', data_type: 'numeric', unit: 'pg/mL', decimal: 2 },
    { item_code: '053003', bml_code: '0004005', display: 'FT4', category: '検体検査', sub_category: '甲状腺', input_type: 'auto', data_type: 'numeric', unit: 'ng/dL', decimal: 2 },
    { item_code: '053004', bml_code: '0004007', display: 'TSH', category: '検体検査', sub_category: '甲状腺', input_type: 'auto', data_type: 'numeric', unit: 'μIU/mL', decimal: 2 }
  ],

  // =============================================
  // 5ページ: 血液型・その他
  // =============================================
  page5_blood_type: [
    { item_code: '054001', bml_code: '0007001', display: '血液型 ABO式', category: '検体検査', sub_category: '血液型', input_type: 'auto', data_type: 'text' },
    { item_code: '054002', bml_code: '0007003', display: '血液型 Rho・D', category: '検体検査', sub_category: '血液型', input_type: 'auto', data_type: 'text' }
  ],

  // =============================================
  // 5ページ: 呼吸器・その他検査
  // =============================================
  page5_respiratory: [
    { item_code: '055001', display: '喀痰細胞診', category: '検体検査', sub_category: '細胞診', input_type: 'manual', data_type: 'text' },
    { item_code: '055002', display: '肺活量', category: '生理検査', sub_category: '呼吸機能', input_type: 'manual', data_type: 'numeric', unit: 'L', decimal: 2 },
    { item_code: '055003', display: '1秒量', category: '生理検査', sub_category: '呼吸機能', input_type: 'manual', data_type: 'numeric', unit: 'L', decimal: 2 },
    { item_code: '055004', display: '%肺活量', category: '生理検査', sub_category: '呼吸機能', input_type: 'manual', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '055005', display: '1秒率', category: '生理検査', sub_category: '呼吸機能', input_type: 'manual', data_type: 'numeric', unit: '%', decimal: 1 },
    { item_code: '055006', display: '%1秒量', category: '生理検査', sub_category: '呼吸機能', input_type: 'manual', data_type: 'numeric', unit: '%', decimal: 1 }
  ],

  // =============================================
  // 5ページ: ピロリ菌関連
  // =============================================
  page5_pylori: [
    { item_code: '056001', bml_code: '0008001', display: '尿素呼気試験', category: '検体検査', sub_category: 'ピロリ菌', input_type: 'auto', data_type: 'numeric' },
    { item_code: '056002', display: 'ウレアーゼ試験検査', category: '検体検査', sub_category: 'ピロリ菌', input_type: 'manual', data_type: 'select' },
    { item_code: '056003', bml_code: '0008003', display: 'ピロリ抗体', category: '検体検査', sub_category: 'ピロリ菌', input_type: 'auto', data_type: 'numeric' }
  ],

  // =============================================
  // 5ページ: その他
  // =============================================
  page5_others: [
    { item_code: '057001', display: '色覚検査', category: '生理検査', sub_category: '眼科', input_type: 'manual', data_type: 'text' },
    { item_code: '057002', display: '末梢血液一般', category: '検体検査', sub_category: '血液学検査', input_type: 'manual', data_type: 'text' },
    { item_code: '057003', bml_code: '0006013', display: 'STS定性', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' },
    { item_code: '057004', bml_code: '0006015', display: '梅毒トレポネーマ抗体定性', category: '検体検査', sub_category: '感染症', input_type: 'auto', data_type: 'select' }
  ]
};

/**
 * 全項目をフラット配列に変換（BMLコード付きのみ抽出可能）
 */
function getAllHumanDockItems(bmlOnly) {
  const all = [];
  for (const category of Object.values(HUMAN_DOCK_ALL_ITEMS)) {
    for (const item of category) {
      if (bmlOnly && !item.bml_code) continue;
      all.push(item);
    }
  }
  return all;
}

/**
 * 旧互換: BMLコード付き項目のみ（元の配列形式）
 */
const HUMAN_DOCK_REQUIRED_ITEMS = getAllHumanDockItems(true);

/**
 * 人間ドック必要項目とマスタを照合（全項目版）
 * GASエディタから実行して確認
 *
 * @param {boolean} bmlOnly - BMLコード付き項目のみチェック（デフォルト: false）
 * @returns {Object} {success, registered, missing, total, report}
 */
function verifyHumanDockItems(bmlOnly) {
  console.log('=== 人間ドック検査項目照合（' + (bmlOnly ? 'BMLのみ' : '全項目') + '） ===');

  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName(M_EXAM_ITEM_SHEET_NAME);

    if (!sheet) {
      return {
        success: false,
        error: M_EXAM_ITEM_SHEET_NAME + 'シートが見つかりません。setupExamItemMasterSheet()を実行してください。'
      };
    }

    const data = sheet.getDataRange().getValues();

    // マスタのBMLコード・項目コードを収集
    const masterBmlCodes = new Map();
    const masterItemCodes = new Map();
    for (let i = 1; i < data.length; i++) {
      const bmlCode = safeString(data[i][M_EXAM_ITEM_COL.LAB_CODE_BML]);
      const itemCode = safeString(data[i][M_EXAM_ITEM_COL.ITEM_CODE]);
      const itemName = safeString(data[i][M_EXAM_ITEM_COL.ITEM_NAME]);
      if (bmlCode) {
        masterBmlCodes.set(bmlCode, { itemCode, itemName, rowIndex: i + 1 });
      }
      if (itemCode) {
        masterItemCodes.set(itemCode, { itemName, rowIndex: i + 1 });
      }
    }

    console.log('マスタ登録数: BMLコード=' + masterBmlCodes.size + ', 項目コード=' + masterItemCodes.size);

    // 対象項目取得
    const allItems = getAllHumanDockItems(bmlOnly || false);
    console.log('照合対象項目数: ' + allItems.length);

    const registered = [];
    const missing = [];

    // 必要項目と照合
    for (const required of allItems) {
      let found = null;

      // BMLコードで照合
      if (required.bml_code) {
        found = masterBmlCodes.get(required.bml_code);
      }

      // 項目コードで照合（BMLコードがない場合）
      if (!found && required.item_code) {
        const byCode = masterItemCodes.get(required.item_code);
        if (byCode) {
          found = { itemCode: required.item_code, itemName: byCode.itemName, rowIndex: byCode.rowIndex };
        }
      }

      if (found) {
        registered.push({
          item_code: required.item_code,
          bml_code: required.bml_code || '',
          required_name: required.display,
          master_item_code: found.itemCode,
          master_item_name: found.itemName,
          row: found.rowIndex,
          input_type: required.input_type
        });
      } else {
        missing.push({
          item_code: required.item_code,
          bml_code: required.bml_code || '',
          display: required.display,
          category: required.category,
          sub_category: required.sub_category,
          input_type: required.input_type,
          data_type: required.data_type,
          unit: required.unit || ''
        });
      }
    }

    // レポート生成
    console.log('\n=== 照合結果 ===');
    console.log('必要項目数: ' + allItems.length);
    console.log('登録済み: ' + registered.length);
    console.log('未登録: ' + missing.length);

    // カテゴリ別集計
    const missingByCategory = {};
    for (const item of missing) {
      const cat = item.sub_category || item.category;
      if (!missingByCategory[cat]) missingByCategory[cat] = [];
      missingByCategory[cat].push(item);
    }

    if (missing.length > 0) {
      console.log('\n【未登録項目一覧（カテゴリ別）】');
      for (const [cat, items] of Object.entries(missingByCategory)) {
        console.log('\n  [' + cat + '] ' + items.length + '件');
        for (const item of items) {
          const codeInfo = item.bml_code ? 'BML:' + item.bml_code : 'ID:' + item.item_code;
          console.log('    - ' + codeInfo + ' ' + item.display + ' (' + item.input_type + ')');
        }
      }
    }

    return {
      success: true,
      total: allItems.length,
      registeredCount: registered.length,
      missingCount: missing.length,
      registered: registered,
      missing: missing,
      missingByCategory: missingByCategory,
      report: {
        summary: '必要' + allItems.length + '項目中、登録済み' + registered.length + '件、未登録' + missing.length + '件',
        allRegistered: missing.length === 0
      }
    };

  } catch (error) {
    console.error('verifyHumanDockItems error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * BMLコード付き項目のみ照合（簡易版）
 */
function verifyHumanDockBmlItems() {
  return verifyHumanDockItems(true);
}

/**
 * 人間ドック不足項目を一括登録
 * GASエディタから実行（verifyHumanDockItems()で不足確認後）
 *
 * @param {boolean} bmlOnly - BMLコード付き項目のみ登録
 * @returns {Object} {success, addedCount, results}
 */
function registerMissingHumanDockItems(bmlOnly) {
  console.log('=== 人間ドック不足項目登録（' + (bmlOnly ? 'BMLのみ' : '全項目') + '） ===');

  try {
    // まず照合
    const verifyResult = verifyHumanDockItems(bmlOnly);
    if (!verifyResult.success) {
      return verifyResult;
    }

    if (verifyResult.missingCount === 0) {
      console.log('不足項目はありません');
      return {
        success: true,
        message: '全項目登録済み',
        addedCount: 0
      };
    }

    const results = [];
    let addedCount = 0;

    // 不足項目を登録
    for (const item of verifyResult.missing) {
      // 定義済みの項目コードを使用
      const itemCode = item.item_code;

      const itemData = {
        item_code: itemCode,
        item_name: item.display,
        item_name_kana: '',
        category: item.category,
        sub_category: item.sub_category,
        data_type: item.data_type || 'numeric',
        decimal_places: item.decimal || 1,
        unit: item.unit || '',
        lab_code_bml: item.bml_code || '',
        is_active: true,
        notes: '人間ドック用項目（自動登録・' + item.input_type + '）'
      };

      const saveResult = portalSaveExamItem(itemData);
      const codeInfo = item.bml_code ? 'BML:' + item.bml_code : 'ID:' + itemCode;

      results.push({
        item_code: itemCode,
        bml_code: item.bml_code || '',
        item_name: item.display,
        input_type: item.input_type,
        success: saveResult.success,
        message: saveResult.message || saveResult.error
      });

      if (saveResult.success) {
        addedCount++;
        console.log('✓ 登録: ' + codeInfo + ' ' + item.display);
      } else {
        console.log('✗ 失敗: ' + codeInfo + ' - ' + saveResult.error);
      }
    }

    console.log('\n登録完了: ' + addedCount + '/' + verifyResult.missingCount + '件');

    return {
      success: true,
      addedCount: addedCount,
      totalMissing: verifyResult.missingCount,
      results: results
    };

  } catch (error) {
    console.error('registerMissingHumanDockItems error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================
// テスト関数
// ============================================

/**
 * 検査項目マスタCRUDテスト
 * GASエディタから実行
 */
function testExamItemMasterCRUD() {
  console.log('=== 検査項目マスタCRUDテスト開始 ===');

  // 1. シート作成
  console.log('\n--- シート作成 ---');
  const setupResult = setupExamItemMasterSheet();
  console.log(JSON.stringify(setupResult, null, 2));

  // 2. テストデータ追加
  console.log('\n--- テストデータ追加 ---');
  const testItem = {
    item_code: '011001',
    item_name: '身長',
    item_name_kana: 'シンチョウ',
    category: '身体情報',
    sub_category: '身体計測',
    data_type: 'numeric',
    decimal_places: 1,
    unit: 'cm',
    is_active: true
  };
  const saveResult = portalSaveExamItem(testItem);
  console.log(JSON.stringify(saveResult, null, 2));

  // 3. 一覧取得
  console.log('\n--- 一覧取得 ---');
  const listResult = portalGetExamItems({});
  console.log(JSON.stringify(listResult, null, 2));

  // 4. 1件取得
  console.log('\n--- 1件取得 ---');
  const getResult = portalGetExamItemById('011001');
  console.log(JSON.stringify(getResult, null, 2));

  // 5. 更新
  console.log('\n--- 更新 ---');
  testItem.notes = 'テスト更新';
  const updateResult = portalSaveExamItem(testItem);
  console.log(JSON.stringify(updateResult, null, 2));

  // 6. カテゴリ一覧
  console.log('\n--- カテゴリ一覧 ---');
  const categoriesResult = portalGetExamItemCategories();
  console.log(JSON.stringify(categoriesResult, null, 2));

  console.log('\n=== テスト完了 ===');
}
