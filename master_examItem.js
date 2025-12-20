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
