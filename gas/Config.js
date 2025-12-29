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
  SPREADSHEET_ID: '16KtctyT2gd7oJZdcu84kUtuP-D9jB9KtLxxzxXx_wdk',

  // シート名定義（設計書4章準拠 + マスタ拡張）
  // ✅ 検査結果シートは「横持ち形式」（1患者1行、101列）を正式採用
  //    列定義は portalApi.js の LAB_RESULT_COLUMNS を参照
  SHEETS: {
    PATIENT_MASTER: '受診者マスタ',
    VISIT_RECORD: '受診記録',
    TEST_RESULT: '検査結果',  // 横持ち形式: 1患者1行、基本5列 + 検査項目96列 = 101列
    // ITEM_MASTER: 削除済み → EXAM_ITEM_MASTER に統一
    EXAM_TYPE_MASTER: '検診種別マスタ',
    COURSE_MASTER: 'コースマスタ',
    GUIDANCE_RECORD: '保健指導記録',
    // Phase 1追加: 検査項目マスタ拡張
    EXAM_ITEM_MASTER: '検査項目マスタ',      // 150項目の詳細定義
    JUDGMENT_CRITERIA: '判定基準マスタ',      // 判定基準（人間ドック学会2025）
    SELECT_OPTIONS: '選択肢マスタ',           // 定性検査の選択肢
    EXAM_COURSE_MASTER: '健診コースマスタ',  // 6コース定義
    // Phase 2追加: 結果入力機能用シート（iD-Heart準拠）
    FINDING_TEMPLATE: 'M_検査所見マスタ',    // 所見テンプレート
    ORGANIZATION_MASTER: 'M_団体マスタ',     // 企業・団体情報
    COURSE_ITEM: 'M_コース項目マスタ',       // コースと検査項目の関連
    JUDGMENT_RESULT: 'T_判定結果',           // 3レベル判定結果
    FINDINGS: 'T_所見',                      // 所見記録
    // Phase 3追加: 帳票マッピング（検査項目とは分離）
    REPORT_MAPPING: 'M_ReportMapping'        // 検診種別×テンプレート→セル位置
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
  // 受診者マスタ（17列構造）
  // ※年度内1回受診のため、受診情報も含む非正規化構造
  // ※カルテNo = クリニック患者ID（CSV取込の主キー）
  PATIENT_MASTER: {
    headers: [
      '受診者ID', 'カルテNo', 'ステータス', '受診日', '氏名', 'カナ',
      '性別', '生年月日', '年齢', '受診コース', '事業所名',
      '所属', '総合判定', 'CSV取込日時', '最終更新日時', '出力日時', 'BML患者ID'
    ],
    columns: {
      PATIENT_ID: 0,      // A: 受診者ID (内部ID: P-00001形式) ※UIで非表示
      KARTE_NO: 1,        // B: カルテNo (クリニック患者ID: 6桁) ★CSV紐付け用
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
      BML_PATIENT_ID: 16  // Q: BML患者ID（BML通信用・トレーサビリティ）
    },
    columnWidths: {
      A: 100, B: 80, C: 80, D: 100, E: 100, F: 120,
      G: 50, H: 100, I: 50, J: 150, K: 150,
      L: 100, M: 80, N: 150, O: 150, P: 150, Q: 100
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
  // ⚠️ 注意: この定義は設計書の縦持ち形式用。
  //    実際の運用では portalApi.js の LAB_RESULT_COLUMNS（横持ち形式）を使用。
  //    縦持ち形式は CRUD.js で一部使用されるが、メインの検査結果管理は横持ち。
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

  // 項目マスタ（旧版）は削除済み
  // → EXAM_ITEM_MASTER（検査項目マスタ）+ JUDGMENT_CRITERIA（判定基準マスタ）に統一
  // 定義は下部の EXAM_ITEM_MASTER_DEF, JUDGMENT_CRITERIA_DEF を参照

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

// ============================================
// Phase 1: 検査項目マスタ拡張シート定義
// ============================================

// 検査項目マスタシート定義（150項目対応）
const EXAM_ITEM_MASTER_DEF = {
  headers: [
    '項目ID', '項目名', 'カテゴリ', 'サブカテゴリ', 'データ型',
    '単位', '人間ドック必須', '定期健診必須', '労災二次必須', '表示順'
  ],
  columns: {
    ITEM_ID: 0,
    ITEM_NAME: 1,
    CATEGORY: 2,
    SUBCATEGORY: 3,
    DATA_TYPE: 4,
    UNIT: 5,
    REQUIRED_DOCK: 6,
    REQUIRED_REGULAR: 7,
    REQUIRED_SECONDARY: 8,
    DISPLAY_ORDER: 9
  },
  columnWidths: {
    A: 120, B: 200, C: 100, D: 100, E: 80,
    F: 80, G: 80, H: 80, I: 80, J: 60
  }
};

// 判定基準マスタシート定義（人間ドック学会2025年度版準拠）
const JUDGMENT_CRITERIA_DEF = {
  headers: [
    '項目ID', '項目名', '性別', '単位',
    'A下限', 'A上限', 'B下限', 'B上限', 'C下限', 'C上限', 'D下限', 'D上限', '備考'
  ],
  columns: {
    ITEM_ID: 0,
    ITEM_NAME: 1,
    GENDER: 2,
    UNIT: 3,
    A_MIN: 4,
    A_MAX: 5,
    B_MIN: 6,
    B_MAX: 7,
    C_MIN: 8,
    C_MAX: 9,
    D_MIN: 10,
    D_MAX: 11,
    NOTE: 12
  },
  columnWidths: {
    A: 120, B: 150, C: 60, D: 100,
    E: 60, F: 60, G: 60, H: 60, I: 60, J: 60, K: 60, L: 60, M: 150
  }
};

// 選択肢マスタシート定義（定性検査用）
const SELECT_OPTIONS_DEF = {
  headers: [
    '項目ID', '選択肢', '説明'
  ],
  columns: {
    ITEM_ID: 0,
    OPTIONS: 1,
    DESCRIPTION: 2
  },
  columnWidths: {
    A: 150, B: 250, C: 200
  }
};

// 健診コースマスタシート定義
const EXAM_COURSE_MASTER_DEF = {
  headers: [
    'コースID', 'コース名', '料金', '説明', '項目数', '必須項目リスト'
  ],
  columns: {
    COURSE_ID: 0,
    COURSE_NAME: 1,
    PRICE: 2,
    DESCRIPTION: 3,
    ITEM_COUNT: 4,
    REQUIRED_ITEMS: 5
  },
  columnWidths: {
    A: 150, B: 180, C: 80, D: 250, E: 60, F: 400
  }
};

// 判定ラベル定義
const JUDGMENT_LABELS = {
  'A': '異常なし',
  'B': '軽度異常',
  'C': '要再検査・生活改善',
  'D': '要精密検査・治療',
  'E': '治療中',
  'F': '経過観察中'
};

// ============================================
// Phase 2: 結果入力機能用シート定義（iD-Heart準拠）
// ============================================

// M_検査所見マスタ（所見テンプレート）
const FINDING_TEMPLATE_DEF = {
  headers: [
    '所見ID', '項目ID', 'カテゴリ', '判定', '所見テンプレート', '優先順位', '有効'
  ],
  columns: {
    FINDING_ID: 0,      // A: 所見ID
    ITEM_ID: 1,         // B: 項目ID（FK）
    CATEGORY: 2,        // C: カテゴリ（身体計測、血圧等）
    JUDGMENT: 3,        // D: 判定（A/B/C/D/E/F）
    TEMPLATE: 4,        // E: 所見テンプレート文章
    PRIORITY: 5,        // F: 優先順位（同一判定内の表示順）
    IS_ACTIVE: 6        // G: 有効フラグ
  },
  columnWidths: {
    A: 100, B: 120, C: 100, D: 60, E: 400, F: 80, G: 60
  }
};

// M_団体マスタ（企業・団体情報）
const ORGANIZATION_MASTER_DEF = {
  headers: [
    '団体ID', '団体名', '郵便番号', '住所', '電話番号',
    '担当者', '契約コース', '請求方法', '備考', '有効'
  ],
  columns: {
    ORG_ID: 0,          // A: 団体ID
    ORG_NAME: 1,        // B: 団体名
    POSTAL_CODE: 2,     // C: 郵便番号
    ADDRESS: 3,         // D: 住所
    PHONE: 4,           // E: 電話番号
    CONTACT: 5,         // F: 担当者
    CONTRACT_COURSE: 6, // G: 契約コース（カンマ区切り）
    BILLING_METHOD: 7,  // H: 請求方法（一括/個別/都度）
    NOTES: 8,           // I: 備考
    IS_ACTIVE: 9        // J: 有効フラグ
  },
  columnWidths: {
    A: 100, B: 200, C: 100, D: 300, E: 120,
    F: 100, G: 150, H: 100, I: 200, J: 60
  }
};

// M_コース項目マスタ（コースと検査項目の多対多リレーション）
const COURSE_ITEM_DEF = {
  headers: [
    'コースID', '項目ID', '必須フラグ', '表示順'
  ],
  columns: {
    COURSE_ID: 0,       // A: コースID（FK）
    ITEM_ID: 1,         // B: 項目ID（FK）
    IS_REQUIRED: 2,     // C: 必須フラグ（TRUE/FALSE）
    DISPLAY_ORDER: 3    // D: 表示順
  },
  columnWidths: {
    A: 150, B: 150, C: 100, D: 80
  }
};

// T_判定結果（3レベル判定結果）
const JUDGMENT_RESULT_DEF = {
  headers: [
    '受診ID', '項目ID', 'Lv1判定', 'Lv2カテゴリ', 'Lv2判定',
    'Lv3総合判定', '判定日時', '判定者'
  ],
  columns: {
    VISIT_ID: 0,        // A: 受診ID（FK）
    ITEM_ID: 1,         // B: 項目ID（FK、Lv1のみ）
    LV1_JUDGMENT: 2,    // C: Lv1判定（A/B/C/D/E/F）
    LV2_CATEGORY: 3,    // D: Lv2カテゴリ名
    LV2_JUDGMENT: 4,    // E: Lv2判定（A/B/C/D/E/F）
    LV3_JUDGMENT: 5,    // F: Lv3総合判定（A/B/C/D/E/F）
    JUDGED_AT: 6,       // G: 判定日時
    JUDGED_BY: 7        // H: 判定者
  },
  columnWidths: {
    A: 130, B: 120, C: 80, D: 120, E: 80,
    F: 100, G: 150, H: 100
  }
};

// T_所見（所見記録）- 縦持ち・検査項目別
const FINDINGS_DEF = {
  headers: [
    '所見ID', '受診者ID', 'カルテNo', '項目ID',
    '所見テキスト', '判定', 'テンプレートID',
    '検査日', '入力者', '作成日時', '更新日時'
  ],
  columns: {
    FINDING_ID: 0,     // A: F00001形式
    PATIENT_ID: 1,     // B: P00001形式
    KARTE_NO: 2,       // C: 6桁カルテNo（クエリ用）
    ITEM_ID: 3,        // D: H02xxxx形式
    FINDING_TEXT: 4,   // E: 所見テキスト
    JUDGMENT: 5,       // F: A/B/C/D/E/F
    TEMPLATE_ID: 6,    // G: 使用テンプレートID
    EXAM_DATE: 7,      // H: 検査実施日
    INPUT_BY: 8,       // I: 入力者名
    CREATED_AT: 9,     // J: 作成日時
    UPDATED_AT: 10     // K: 更新日時
  },
  columnWidths: {
    A: 100, B: 100, C: 80, D: 100,
    E: 400, F: 60, G: 100,
    H: 100, I: 100, J: 150, K: 150
  }
};

// 判定カラー定義（iD-Heart準拠、6段階）
const JUDGMENT_COLORS_EXTENDED = {
  A: { bg: '#e8f5e9', text: '#2e7d32' },  // 緑 - 異常なし
  B: { bg: '#fff8e1', text: '#f57f17' },  // 黄 - 軽度異常
  C: { bg: '#fff3e0', text: '#e65100' },  // 橙 - 要経過観察
  D: { bg: '#ffebee', text: '#c62828' },  // 赤 - 要精密検査
  E: { bg: '#f3e5f5', text: '#6a1b9a' },  // 紫 - 治療中
  F: { bg: '#e3f2fd', text: '#1565c0' }   // 青 - 経過観察中
};

// 入力画面タブ定義（iD-Heart準拠）
const INPUT_TABS = [
  { id: 'interview', name: '問診結果', categories: ['基本情報', '問診'] },
  { id: 'physical', name: '身体情報', categories: ['身体測定', '血圧', '眼科', '聴力'] },
  { id: 'laboratory', name: '検体検査', categories: ['尿検査', '血液学検査', '肝胆膵機能', '脂質検査', '糖代謝', '腎機能'] },
  { id: 'imaging', name: '画像診断', categories: ['画像診断', '心電図', '胸部X線'] },
  { id: 'judgment', name: '判定項目', view: 'judgment' },
  { id: 'findings', name: '所見文章', view: 'findings' }
];

// カテゴリ別判定定義（Lv2判定用）
const JUDGMENT_CATEGORIES = [
  { id: 'body', name: '身体計測', items: ['身長', '体重', 'BMI', '腹囲'] },
  { id: 'bp', name: '血圧', items: ['収縮期血圧', '拡張期血圧'] },
  { id: 'hearing', name: '聴力', items: ['1000Hz右', '1000Hz左', '4000Hz右', '4000Hz左'] },
  { id: 'eye', name: '眼科', items: ['視力右', '視力左', '眼底', '眼圧'] },
  { id: 'urine', name: '尿検査', items: ['尿蛋白', '尿糖', '尿潜血'] },
  { id: 'blood', name: '血液学', items: ['白血球', '赤血球', 'ヘモグロビン', 'ヘマトクリット', '血小板'] },
  { id: 'liver', name: '肝胆膵', items: ['AST', 'ALT', 'γ-GTP', 'ALP', '総ビリルビン'] },
  { id: 'lipid', name: '脂質', items: ['総コレステロール', 'HDLコレステロール', 'LDLコレステロール', '中性脂肪'] },
  { id: 'diabetes', name: '糖尿病', items: ['空腹時血糖', 'HbA1c'] },
  { id: 'ecg', name: '心電図', items: ['心電図'] },
  { id: 'chest', name: '胸部', items: ['胸部X線'] }
];

// ============================================
// M_ReportMapping（帳票マッピング）
// 設計原則: 検査項目マスタと分離
// - 単一責任: 項目定義は項目マスタ、出力位置はマッピングテーブル
// - 疎結合: 検診種別追加時に項目マスタ変更不要
// - 正規化: 項目ID参照でDRY原則遵守
// ============================================
const REPORT_MAPPING_DEF = {
  headers: [
    'mapping_id',    // A: マッピングID（PK）
    'exam_type',     // B: 検診種別（DOCK/REGULAR/EMPLOY/ROSAI）
    'template_id',   // C: テンプレートID（template_new_default等）
    'sheet_name',    // D: シート名（1ページ/4ページ等）
    'item_id',       // E: 項目ID（M_検査項目のitem_code FK）
    'bml_code',      // F: BMLコード（照合用）
    'value_cell',    // G: 値セル（M6等）
    'judgment_cell', // H: 判定セル（K6等）
    'flag_cell',     // I: フラグセル（O6等）
    'format',        // J: 出力形式（numeric/text/date）
    'decimal_places',// K: 小数桁数
    'notes',         // L: 備考
    'is_active'      // M: 有効フラグ
  ],
  columns: {
    MAPPING_ID: 0,
    EXAM_TYPE: 1,
    TEMPLATE_ID: 2,
    SHEET_NAME: 3,
    ITEM_ID: 4,
    BML_CODE: 5,
    VALUE_CELL: 6,
    JUDGMENT_CELL: 7,
    FLAG_CELL: 8,
    FORMAT: 9,
    DECIMAL_PLACES: 10,
    NOTES: 11,
    IS_ACTIVE: 12
  },
  columnWidths: {
    A: 100, B: 80, C: 150, D: 80, E: 100,
    F: 80, G: 60, H: 60, I: 60, J: 80,
    K: 60, L: 200, M: 60
  }
};

// Auto-deploy test: 2025-12-18 07:21:10
// Workflow test: 2025-12-18 07:28:47
// Re-deploy: 2025-12-18 13:04:08
