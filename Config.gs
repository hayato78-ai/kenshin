/**
 * 健診結果DB 統合システム - 設定ファイル
 *
 * @description 設計書: 健診結果DB_設計書_v1.md に基づく
 * @version 1.0.0
 * @date 2025-12-14
 */

// ============================================
// グローバル設定
// ============================================

const DB_CONFIG = {
  // スプレッドシートID（デプロイ時に設定）
  SPREADSHEET_ID: '',

  // シート名定義（設計書4章準拠）
  SHEETS: {
    PATIENT_MASTER: '受診者マスタ',
    VISIT_RECORD: '受診記録',
    TEST_RESULT: '検査結果',
    ITEM_MASTER: '項目マスタ',
    EXAM_TYPE_MASTER: '検診種別マスタ',
    COURSE_MASTER: 'コースマスタ',
    GUIDANCE_RECORD: '保健指導記録'
  },

  // ステータス定義（設計書4.2準拠）
  STATUS: {
    INPUT: '入力中',
    CONFIRMED: '確定',
    REPORTED: '報告済'
  },

  // 判定定義（設計書5章準拠）
  JUDGMENT: {
    A: 'A',  // 異常なし
    B: 'B',  // 軽度異常
    C: 'C',  // 要経過観察
    D: 'D'   // 要精密検査
  },

  // 判定の色定義
  JUDGMENT_COLORS: {
    A: '#e8f5e9',  // 緑
    B: '#fff8e1',  // 黄
    C: '#fff3e0',  // 橙
    D: '#ffebee'   // 赤
  },

  // ID採番プレフィックス（設計書6章準拠）
  ID_PREFIX: {
    PATIENT: 'P',      // 受診者ID: P00001
    RESULT: 'R',       // 結果ID: R00001
    GUIDANCE: 'G'      // 指導ID: G00001
  },

  // 検診種別ID（設計書4.5準拠）
  EXAM_TYPE: {
    DOCK: 'DOCK',           // 人間ドック
    REGULAR: 'REGULAR',     // 定期検診
    EMPLOY: 'EMPLOY',       // 雇入検診
    ROSAI: 'ROSAI',         // 労災二次
    SPECIFIC: 'SPECIFIC'    // 特定健診
  }
};

// ============================================
// シート列定義（設計書4章準拠）
// ============================================

const COLUMN_DEFINITIONS = {
  // 受診者マスタ（設計書4.1）
  PATIENT_MASTER: {
    headers: [
      '受診者ID', '氏名', 'カナ', '生年月日', '性別',
      '郵便番号', '住所', '電話番号', 'メール', '所属企業',
      '備考', '作成日時', '更新日時'
    ],
    columns: {
      PATIENT_ID: 0,    // A: 受診者ID
      NAME: 1,          // B: 氏名
      KANA: 2,          // C: カナ
      BIRTHDATE: 3,     // D: 生年月日
      GENDER: 4,        // E: 性別 (M/F)
      POSTAL_CODE: 5,   // F: 郵便番号
      ADDRESS: 6,       // G: 住所
      PHONE: 7,         // H: 電話番号
      EMAIL: 8,         // I: メール
      COMPANY: 9,       // J: 所属企業
      NOTES: 10,        // K: 備考
      CREATED_AT: 11,   // L: 作成日時
      UPDATED_AT: 12    // M: 更新日時
    },
    columnWidths: {
      A: 100, B: 100, C: 120, D: 100, E: 50,
      F: 80, G: 200, H: 120, I: 150, J: 150,
      K: 200, L: 150, M: 150
    }
  },

  // 受診記録（設計書4.2）
  VISIT_RECORD: {
    headers: [
      '受診ID', '受診者ID', '検診種別ID', 'コースID', '受診日',
      '年齢', '総合判定', '医師所見', 'ステータス', '作成日時', '更新日時'
    ],
    columns: {
      VISIT_ID: 0,        // A: 受診ID (YYYYMMDD-NNN)
      PATIENT_ID: 1,      // B: 受診者ID (FK)
      EXAM_TYPE_ID: 2,    // C: 検診種別ID (FK)
      COURSE_ID: 3,       // D: コースID (FK)
      VISIT_DATE: 4,      // E: 受診日
      AGE: 5,             // F: 年齢
      OVERALL_JUDGMENT: 6, // G: 総合判定
      DOCTOR_NOTES: 7,    // H: 医師所見
      STATUS: 8,          // I: ステータス
      CREATED_AT: 9,      // J: 作成日時
      UPDATED_AT: 10      // K: 更新日時
    },
    columnWidths: {
      A: 130, B: 100, C: 100, D: 100, E: 100,
      F: 50, G: 80, H: 300, I: 80, J: 150, K: 150
    }
  },

  // 検査結果（設計書4.3 - 縦持ち形式）
  TEST_RESULT: {
    headers: [
      '結果ID', '受診ID', '項目ID', '値', '数値変換',
      '判定', '所見', '作成日時'
    ],
    columns: {
      RESULT_ID: 0,      // A: 結果ID (R00001)
      VISIT_ID: 1,       // B: 受診ID (FK)
      ITEM_ID: 2,        // C: 項目ID (FK)
      VALUE: 3,          // D: 値
      NUMERIC_VALUE: 4,  // E: 数値変換
      JUDGMENT: 5,       // F: 判定
      NOTES: 6,          // G: 所見
      CREATED_AT: 7      // H: 作成日時
    },
    columnWidths: {
      A: 100, B: 130, C: 80, D: 80, E: 80,
      F: 50, G: 200, H: 150
    }
  },

  // 項目マスタ（設計書4.4）
  ITEM_MASTER: {
    headers: [
      '項目ID', '項目名', 'カテゴリ', '単位', 'データ型',
      '性別差', '判定方法', 'A下限', 'A上限', 'B下限', 'B上限',
      'C下限', 'C上限', 'D条件', 'A下限_F', 'A上限_F', '表示順', '有効'
    ],
    columns: {
      ITEM_ID: 0,        // A: 項目ID
      ITEM_NAME: 1,      // B: 項目名
      CATEGORY: 2,       // C: カテゴリ
      UNIT: 3,           // D: 単位
      DATA_TYPE: 4,      // E: データ型
      GENDER_DIFF: 5,    // F: 性別差
      JUDGMENT_METHOD: 6, // G: 判定方法
      A_MIN: 7,          // H: A下限
      A_MAX: 8,          // I: A上限
      B_MIN: 9,          // J: B下限
      B_MAX: 10,         // K: B上限
      C_MIN: 11,         // L: C下限
      C_MAX: 12,         // M: C上限
      D_CONDITION: 13,   // N: D条件
      A_MIN_F: 14,       // O: A下限_F
      A_MAX_F: 15,       // P: A上限_F
      DISPLAY_ORDER: 16, // Q: 表示順
      IS_ACTIVE: 17      // R: 有効
    },
    columnWidths: {
      A: 80, B: 120, C: 100, D: 60, E: 60,
      F: 60, G: 80, H: 60, I: 60, J: 60, K: 60,
      L: 60, M: 60, N: 100, O: 60, P: 60, Q: 60, R: 50
    }
  },

  // 検診種別マスタ（設計書4.5）
  EXAM_TYPE_MASTER: {
    headers: [
      '種別ID', '種別名', 'コース必須', '表示順', '有効'
    ],
    columns: {
      TYPE_ID: 0,        // A: 種別ID
      TYPE_NAME: 1,      // B: 種別名
      COURSE_REQUIRED: 2, // C: コース必須
      DISPLAY_ORDER: 3,  // D: 表示順
      IS_ACTIVE: 4       // E: 有効
    },
    columnWidths: {
      A: 100, B: 150, C: 80, D: 60, E: 50
    }
  },

  // コースマスタ（設計書4.6）
  COURSE_MASTER: {
    headers: [
      'コースID', 'コース名', '料金', '検査項目', '表示順', '有効'
    ],
    columns: {
      COURSE_ID: 0,      // A: コースID
      COURSE_NAME: 1,    // B: コース名
      PRICE: 2,          // C: 料金
      TEST_ITEMS: 3,     // D: 検査項目
      DISPLAY_ORDER: 4,  // E: 表示順
      IS_ACTIVE: 5       // F: 有効
    },
    columnWidths: {
      A: 100, B: 150, C: 80, D: 300, E: 60, F: 50
    }
  },

  // 保健指導記録（設計書4.7）
  GUIDANCE_RECORD: {
    headers: [
      '指導ID', '受診ID', '指導日', '指導区分', '指導内容',
      '担当者', '次回予定', '作成日時'
    ],
    columns: {
      GUIDANCE_ID: 0,    // A: 指導ID (G00001)
      VISIT_ID: 1,       // B: 受診ID (FK)
      GUIDANCE_DATE: 2,  // C: 指導日
      GUIDANCE_TYPE: 3,  // D: 指導区分
      CONTENT: 4,        // E: 指導内容
      STAFF: 5,          // F: 担当者
      NEXT_DATE: 6,      // G: 次回予定
      CREATED_AT: 7      // H: 作成日時
    },
    columnWidths: {
      A: 100, B: 130, C: 100, D: 120, E: 300,
      F: 100, G: 100, H: 150
    }
  }
};

// ============================================
// ユーティリティ関数
// ============================================

/**
 * スプレッドシートを取得
 * @returns {Spreadsheet} スプレッドシート
 */
function getSpreadsheet() {
  if (DB_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(DB_CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * シートを取得
 * @param {string} sheetName - シート名
 * @returns {Sheet|null} シート
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * ログ出力（INFO）
 * @param {string} message - メッセージ
 */
function logInfo(message) {
  console.log(`[INFO] ${new Date().toLocaleString('ja-JP')}: ${message}`);
}

/**
 * ログ出力（ERROR）
 * @param {string} funcName - 関数名
 * @param {Error} error - エラー
 */
function logError(funcName, error) {
  console.error(`[ERROR] ${funcName}: ${error.message}`);
  console.error(error.stack);
}

/**
 * 列文字をインデックスに変換
 * @param {string} letter - 列文字（A, B, ..., AA, AB, ...）
 * @returns {number} 列インデックス（1始まり）
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

/**
 * 日付をYYYYMMDD形式に変換
 * @param {Date} date - 日付
 * @returns {string} YYYYMMDD形式
 */
function formatDateYYYYMMDD(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

/**
 * 年齢を計算
 * @param {Date} birthDate - 生年月日
 * @param {Date} targetDate - 基準日（省略時は今日）
 * @returns {number} 年齢
 */
function calculateAge(birthDate, targetDate = new Date()) {
  const birth = new Date(birthDate);
  const target = new Date(targetDate);

  let age = target.getFullYear() - birth.getFullYear();
  const monthDiff = target.getMonth() - birth.getMonth();

  if (monthDiff < 0 || (monthDiff === 0 && target.getDate() < birth.getDate())) {
    age--;
  }

  return age;
}

// ============================================
// CSVインポート設定（Phase 3）
// ============================================

const CSV_CONFIG = {
  // 最大行数
  MAX_ROWS: 1000,

  // 対応エンコーディング
  SUPPORTED_ENCODINGS: ['UTF-8', 'Shift_JIS'],

  // AI信頼度閾値
  AI_CONFIDENCE_THRESHOLD: 0.7,

  // データ種別
  DATA_TYPES: {
    PATIENT: 'PATIENT',
    TEST_RESULT: 'TEST_RESULT'
  },

  // インポートオプションデフォルト
  DEFAULT_OPTIONS: {
    skipErrors: true,
    allowDuplicates: false,
    savePattern: true
  }
};

// 受診者インポート用スキーマ
const PATIENT_IMPORT_SCHEMA = [
  { id: 'name', name: '氏名', aliases: ['お名前', '患者名', '氏名（漢字）', 'NAME'], required: true },
  { id: 'kana', name: 'カナ', aliases: ['フリガナ', 'ふりがな', 'カナ', 'KANA'], required: true },
  { id: 'birthdate', name: '生年月日', aliases: ['生年月日', '誕生日', 'BIRTH', 'BIRTHDAY'], required: true },
  { id: 'gender', name: '性別', aliases: ['性別', 'SEX', 'GENDER'], required: true },
  { id: 'phone', name: '電話番号', aliases: ['電話', 'TEL', 'PHONE', '携帯'], required: false },
  { id: 'postalCode', name: '郵便番号', aliases: ['郵便番号', '〒', 'ZIP'], required: false },
  { id: 'address', name: '住所', aliases: ['住所', 'ADDRESS'], required: false },
  { id: 'email', name: 'メール', aliases: ['メール', 'EMAIL', 'E-MAIL'], required: false },
  { id: 'company', name: '所属企業', aliases: ['企業名', '会社名', '所属', 'COMPANY', '勤務先'], required: false }
];

// 検査結果インポート用スキーマ（縦持ち形式）
const TEST_RESULT_IMPORT_SCHEMA = [
  { id: 'visitDate', name: '受診日', aliases: ['受診日', '検査日', 'DATE'], required: true },
  { id: 'name', name: '氏名', aliases: ['氏名', 'お名前', '患者名'], required: true },
  { id: 'gender', name: '性別', aliases: ['性別', 'SEX'], required: true },
  { id: 'itemId', name: '項目ID', aliases: ['項目ID', '検査項目ID', 'ITEM_ID'], required: true },
  { id: 'value', name: '値', aliases: ['値', '結果値', 'VALUE', '検査値'], required: true },
  { id: 'judgment', name: '判定', aliases: ['判定', 'JUDGMENT', '結果判定'], required: false }
];

// 検査結果インポート用スキーマ（横持ち形式）
const TEST_RESULT_HORIZONTAL_SCHEMA = [
  { id: 'visitDate', name: '受診日', aliases: ['受診日', '検査日', 'DATE'], required: true },
  { id: 'name', name: '氏名', aliases: ['氏名', 'お名前', '患者名'], required: true },
  { id: 'kana', name: 'カナ', aliases: ['フリガナ', 'ふりがな', 'カナ'], required: false },
  { id: 'birthdate', name: '生年月日', aliases: ['生年月日', '誕生日'], required: true },
  { id: 'gender', name: '性別', aliases: ['性別', 'SEX'], required: true },
  { id: 'company', name: '企業名', aliases: ['企業名', '会社名', '所属'], required: false }
  // 以降は項目マスタから動的に追加
];

// マッピングパターンシート定義
const MAPPING_PATTERN_SHEET = {
  NAME: 'M_MappingPattern',
  HEADERS: [
    'パターンID', 'ソース名', 'ヘッダーハッシュ', 'データ種別',
    'マッピングJSON', '値変換JSON', '使用回数', '作成日時', '更新日時'
  ],
  COLUMNS: {
    PATTERN_ID: 0,
    SOURCE_NAME: 1,
    HEADERS_HASH: 2,
    DATA_TYPE: 3,
    MAPPINGS_JSON: 4,
    VALUE_TRANSFORMS_JSON: 5,
    USE_COUNT: 6,
    CREATED_AT: 7,
    UPDATED_AT: 8
  }
};
