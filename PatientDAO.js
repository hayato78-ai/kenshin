/**
 * PatientDAO.js - 受診者データアクセス層
 *
 * 設計原則:
 * - 物理的な列位置（Config.js）をここだけで参照
 * - ビジネスロジック層にはエンティティオブジェクトを返す
 * - UI層は直接アクセスしない
 *
 * @version 1.0.0
 * @date 2025-12-22
 */

// ============================================
// PatientDAO - データアクセスオブジェクト
// ============================================

// 受診者マスタ列定義（17列構造）
// ※ COLUMN_DEFINITIONS参照問題を回避するため直接定義
const PATIENT_COLS = {
  PATIENT_ID: 0,      // A: 受診者ID
  KARTE_NO: 1,        // B: カルテNo
  STATUS: 2,          // C: ステータス
  VISIT_DATE: 3,      // D: 受診日
  NAME: 4,            // E: 氏名
  KANA: 5,            // F: カナ
  GENDER: 6,          // G: 性別
  BIRTHDATE: 7,       // H: 生年月日
  AGE: 8,             // I: 年齢
  COURSE: 9,          // J: 受診コース
  COMPANY: 10,        // K: 事業所名
  DEPARTMENT: 11,     // L: 所属
  OVERALL_JUDGMENT: 12, // M: 総合判定
  CSV_IMPORT_DATE: 13,  // N: CSV取込日時
  UPDATED_AT: 14,     // O: 最終更新日時
  EXPORT_DATE: 15,    // P: 出力日時
  BML_PATIENT_ID: 16  // Q: BML患者ID
};
const PATIENT_COL_COUNT = 17;

const PatientDAO = {

  /**
   * シートを取得
   * @returns {Sheet} 受診者マスタシート
   */
  getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheetByName('受診者マスタ');
  },

  /**
   * 列定義を取得（内部用）
   * @returns {Object} 列インデックス定義
   */
  getColumns() {
    return PATIENT_COLS;
  },

  /**
   * ヘッダー数を取得
   * @returns {number} 列数
   */
  getColumnCount() {
    return PATIENT_COL_COUNT;
  },

  // ============================================
  // 変換メソッド（コア機能）
  // ============================================

  /**
   * 1行の生データ → エンティティに変換
   * @param {Array} row - シートの1行データ
   * @param {number} rowIndex - 行番号（1始まり、オプション）
   * @returns {Object} 受診者エンティティ
   */
  rowToEntity(row, rowIndex = null) {
    const cols = this.getColumns();
    return {
      // 識別子
      patientId: row[cols.PATIENT_ID] || '',
      karteNo: row[cols.KARTE_NO] || '',

      // 受診情報
      status: row[cols.STATUS] || '',
      visitDate: row[cols.VISIT_DATE] || '',

      // 個人情報
      name: row[cols.NAME] || '',
      kana: row[cols.KANA] || '',
      gender: row[cols.GENDER] || '',
      birthdate: row[cols.BIRTHDATE] || '',
      age: row[cols.AGE] || '',

      // 所属情報
      course: row[cols.COURSE] || '',
      company: row[cols.COMPANY] || '',
      department: row[cols.DEPARTMENT] || '',

      // 判定・処理情報
      overallJudgment: row[cols.OVERALL_JUDGMENT] || '',
      csvImportDate: row[cols.CSV_IMPORT_DATE] || '',
      updatedAt: row[cols.UPDATED_AT] || '',
      exportDate: row[cols.EXPORT_DATE] || '',
      bmlPatientId: row[cols.BML_PATIENT_ID] || '',

      // メタ情報
      _rowIndex: rowIndex
    };
  },

  /**
   * エンティティ → 行データに変換
   * @param {Object} entity - 受診者エンティティ
   * @returns {Array} シート用の行データ
   */
  entityToRow(entity) {
    const cols = this.getColumns();
    const row = new Array(this.getColumnCount()).fill('');

    row[cols.PATIENT_ID] = entity.patientId || '';
    row[cols.KARTE_NO] = entity.karteNo || '';
    row[cols.STATUS] = entity.status || '';
    row[cols.VISIT_DATE] = entity.visitDate || '';
    row[cols.NAME] = entity.name || '';
    row[cols.KANA] = entity.kana || '';
    row[cols.GENDER] = entity.gender || '';
    row[cols.BIRTHDATE] = entity.birthdate || '';
    row[cols.AGE] = entity.age || '';
    row[cols.COURSE] = entity.course || '';
    row[cols.COMPANY] = entity.company || '';
    row[cols.DEPARTMENT] = entity.department || '';
    row[cols.OVERALL_JUDGMENT] = entity.overallJudgment || '';
    row[cols.CSV_IMPORT_DATE] = entity.csvImportDate || '';
    row[cols.UPDATED_AT] = entity.updatedAt || new Date();
    row[cols.EXPORT_DATE] = entity.exportDate || '';
    row[cols.BML_PATIENT_ID] = entity.bmlPatientId || '';

    return row;
  },

  // ============================================
  // CRUD操作
  // ============================================

  /**
   * 全件取得
   * @returns {Array<Object>} エンティティ配列
   */
  getAll() {
    const sheet = this.getSheet();
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, this.getColumnCount()).getValues();
    return data.map((row, i) => this.rowToEntity(row, i + 2));
  },

  /**
   * ID指定で取得
   * @param {string} patientId - 受診者ID
   * @returns {Object|null} エンティティまたはnull
   */
  getById(patientId) {
    if (!patientId) return null;

    const sheet = this.getSheet();
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;

    const cols = this.getColumns();
    const data = sheet.getRange(2, 1, lastRow - 1, this.getColumnCount()).getValues();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][cols.PATIENT_ID]) === String(patientId)) {
        return this.rowToEntity(data[i], i + 2);
      }
    }
    return null;
  },

  /**
   * カルテNo指定で取得
   * @param {string} karteNo - カルテNo
   * @returns {Object|null} エンティティまたはnull
   */
  getByKarteNo(karteNo) {
    if (!karteNo) return null;

    const sheet = this.getSheet();
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;

    const cols = this.getColumns();
    const data = sheet.getRange(2, 1, lastRow - 1, this.getColumnCount()).getValues();
    const karteNoStr = String(karteNo).trim();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][cols.KARTE_NO]).trim() === karteNoStr) {
        return this.rowToEntity(data[i], i + 2);
      }
    }
    return null;
  },

  /**
   * 氏名と生年月日で検索
   * @param {string} name - 氏名
   * @param {Date|string} birthdate - 生年月日
   * @returns {Object|null} エンティティまたはnull
   */
  getByNameAndBirthdate(name, birthdate) {
    if (!name || !birthdate) return null;

    const sheet = this.getSheet();
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;

    const cols = this.getColumns();
    const data = sheet.getRange(2, 1, lastRow - 1, this.getColumnCount()).getValues();

    const normalizedName = normalizeString(name);
    const targetBirthStr = formatDateToString(birthdate);

    for (let i = 0; i < data.length; i++) {
      const rowName = normalizeString(data[i][cols.NAME]);
      const rowBirthStr = formatDateToString(data[i][cols.BIRTHDATE]);

      if (rowName === normalizedName && rowBirthStr === targetBirthStr) {
        return this.rowToEntity(data[i], i + 2);
      }
    }
    return null;
  },

  /**
   * 条件検索
   * @param {Object} criteria - 検索条件
   * @returns {Array<Object>} エンティティ配列
   */
  search(criteria) {
    const all = this.getAll();
    if (!criteria || Object.keys(criteria).length === 0) {
      return all;
    }

    return all.filter(entity => {
      // カルテNo検索
      if (criteria.karteNo) {
        const searchKarte = String(criteria.karteNo).trim();
        const entityKarte = String(entity.karteNo || '').trim();
        if (!entityKarte.includes(searchKarte)) return false;
      }

      // 氏名検索（部分一致）
      if (criteria.name) {
        const searchName = normalizeString(criteria.name);
        const entityName = normalizeString(entity.name || '');
        if (!entityName.includes(searchName)) return false;
      }

      // カナ検索（部分一致）
      if (criteria.kana) {
        const searchKana = normalizeString(criteria.kana);
        const entityKana = normalizeString(entity.kana || '');
        if (!entityKana.includes(searchKana)) return false;
      }

      // ステータス検索
      if (criteria.status && entity.status !== criteria.status) {
        return false;
      }

      // 受診日範囲検索
      if (criteria.visitDateFrom || criteria.visitDateTo) {
        const visitDate = entity.visitDate ? new Date(entity.visitDate) : null;
        if (!visitDate) return false;

        if (criteria.visitDateFrom && visitDate < new Date(criteria.visitDateFrom)) {
          return false;
        }
        if (criteria.visitDateTo && visitDate > new Date(criteria.visitDateTo)) {
          return false;
        }
      }

      return true;
    });
  },

  /**
   * 新規保存
   * @param {Object} entity - 受診者エンティティ
   * @returns {Object} 結果 {success, patientId, error}
   */
  save(entity) {
    try {
      const sheet = this.getSheet();
      if (!sheet) {
        return { success: false, error: '受診者マスタシートが見つかりません' };
      }

      // IDが未設定の場合は生成
      if (!entity.patientId) {
        entity.patientId = this.generateNextId(sheet);
      }

      // 更新日時を設定
      entity.updatedAt = new Date();

      const row = this.entityToRow(entity);
      sheet.appendRow(row);

      logInfo(`PatientDAO.save: ${entity.patientId} (${entity.name})`);

      return { success: true, patientId: entity.patientId };
    } catch (e) {
      logError('PatientDAO.save', e);
      return { success: false, error: e.message };
    }
  },

  /**
   * 更新
   * @param {Object} entity - 受診者エンティティ（_rowIndex または patientId 必須）
   * @returns {Object} 結果 {success, error}
   */
  update(entity) {
    try {
      const sheet = this.getSheet();
      if (!sheet) {
        return { success: false, error: '受診者マスタシートが見つかりません' };
      }

      let rowIndex = entity._rowIndex;

      // rowIndexがない場合はpatientIdで検索
      if (!rowIndex && entity.patientId) {
        const existing = this.getById(entity.patientId);
        if (existing) {
          rowIndex = existing._rowIndex;
        }
      }

      if (!rowIndex) {
        return { success: false, error: '更新対象の行が見つかりません' };
      }

      // 更新日時を設定
      entity.updatedAt = new Date();

      const row = this.entityToRow(entity);
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);

      logInfo(`PatientDAO.update: ${entity.patientId} (row ${rowIndex})`);

      return { success: true };
    } catch (e) {
      logError('PatientDAO.update', e);
      return { success: false, error: e.message };
    }
  },

  /**
   * 特定フィールドのみ更新
   * @param {string} patientId - 受診者ID
   * @param {Object} fields - 更新するフィールド
   * @returns {Object} 結果 {success, error}
   */
  updateFields(patientId, fields) {
    try {
      const existing = this.getById(patientId);
      if (!existing) {
        return { success: false, error: '受診者が見つかりません' };
      }

      // 既存データにフィールドをマージ
      const updated = { ...existing, ...fields };
      return this.update(updated);
    } catch (e) {
      logError('PatientDAO.updateFields', e);
      return { success: false, error: e.message };
    }
  },

  // ============================================
  // ユーティリティ
  // ============================================

  /**
   * 次のIDを生成
   * @param {Sheet} sheet - シート
   * @returns {string} P-00001形式のID
   */
  generateNextId(sheet) {
    sheet = sheet || this.getSheet();
    if (!sheet || sheet.getLastRow() < 2) {
      return 'P-00001';
    }

    const cols = this.getColumns();
    const data = sheet.getRange(2, cols.PATIENT_ID + 1, sheet.getLastRow() - 1, 1).getValues();
    let maxNum = 0;

    for (const row of data) {
      const id = String(row[0]);
      if (id.startsWith('P-')) {
        const num = parseInt(id.replace('P-', ''), 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }

    return `P-${String(maxNum + 1).padStart(5, '0')}`;
  },

  /**
   * 件数取得
   * @returns {number} データ件数
   */
  count() {
    const sheet = this.getSheet();
    if (!sheet) return 0;
    return Math.max(0, sheet.getLastRow() - 1);
  }
};

// ============================================
// ヘルパー関数（PatientDAO内部用）
// ============================================

/**
 * 文字列を正規化（空白除去、全角半角統一）
 * @param {*} value - 値
 * @returns {string} 正規化された文字列
 */
function normalizeString(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().replace(/\s+/g, ' ');
}

/**
 * 日付を文字列に変換
 * @param {*} value - 日付値
 * @returns {string} YYYY/MM/DD形式
 */
function formatDateToString(value) {
  if (!value) return '';
  try {
    const date = value instanceof Date ? value : new Date(value);
    if (isNaN(date.getTime())) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  } catch (e) {
    return '';
  }
}
