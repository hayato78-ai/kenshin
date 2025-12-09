/**
 * ユーティリティモジュール
 * 共通関数とシステム設定
 */

// ============================================
// システム設定
// ============================================
const CONFIG = {
  // フォルダID（実際のIDに置き換えが必要）
  CSV_FOLDER_ID: 'YOUR_CSV_FOLDER_ID',      // 01_csv/
  OUTPUT_FOLDER_ID: 'YOUR_OUTPUT_FOLDER_ID', // 02_出力フォルダ/

  // スプレッドシートID（実際のIDに置き換えが必要）
  MASTER_SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',

  // ============================================
  // 健診種別パラメータ
  // ============================================

  /**
   * 現在の健診種別を取得
   * @returns {string} 'DOCK' | 'ROSAI_SECONDARY'
   */
  getExamType: function() {
    return PropertiesService.getScriptProperties().getProperty('EXAM_TYPE') || 'DOCK';
  },

  /**
   * 健診種別を設定
   * @param {string} examType - 'DOCK' | 'ROSAI_SECONDARY'
   */
  setExamType: function(examType) {
    if (!this.EXAM_PROFILES[examType]) {
      throw new Error('不正な健診種別: ' + examType);
    }
    PropertiesService.getScriptProperties().setProperty('EXAM_TYPE', examType);
  },

  /**
   * 現在のプロファイルを取得
   * @returns {Object} 健診プロファイル
   */
  getProfile: function() {
    return this.EXAM_PROFILES[this.getExamType()] || this.EXAM_PROFILES.DOCK;
  },

  /**
   * 項目が現在のプロファイルで有効かチェック
   * @param {string} itemCode - 検査項目コード
   * @returns {boolean}
   */
  isItemEnabled: function(itemCode) {
    const profile = this.getProfile();
    return profile.enabledItems.includes(itemCode);
  },

  // 健診プロファイル定義
  EXAM_PROFILES: {
    // 人間ドック
    DOCK: {
      name: '人間ドック',
      code: 'DOCK',
      csvFormat: 'BML',
      templateFileId: '',  // 設定シートから読み込み
      sheetName: '血液検査',
      enabledItems: [
        'FBS', 'HBA1C', 'TC', 'TG', 'HDL_CHOLESTEROL', 'LDL_CHOLESTEROL',
        'AST_GOT', 'ALT_GPT', 'GAMMA_GTP', 'TOTAL_PROTEIN', 'T_BIL', 'ALB',
        'CREATININE', 'EGFR', 'URIC_ACID', 'BUN',
        'WBC', 'RBC', 'HEMOGLOBIN', 'HT', 'MCV', 'MCH', 'MCHC', 'PLT',
        'CRP', 'BMI', 'WAIST', 'BLOOD_PRESSURE_SYS', 'BLOOD_PRESSURE_DIA'
      ],
      judgmentOverrides: {}
    },

    // 労災二次検診
    ROSAI_SECONDARY: {
      name: '労災二次検診',
      code: 'ROSAI',
      csvFormat: 'ROSAI',
      templateFileId: '',  // 設定シートから読み込み
      sheetName: '二次検診結果',
      enabledItems: [
        'FBS', 'HBA1C', 'TG', 'HDL_CHOLESTEROL', 'LDL_CHOLESTEROL',
        'AST_GOT', 'ALT_GPT', 'GAMMA_GTP',
        'CREATININE', 'URIC_ACID',
        'BMI', 'WAIST', 'BLOOD_PRESSURE_SYS', 'BLOOD_PRESSURE_DIA'
      ],
      judgmentOverrides: {
        // 労災二次検診では両側異常判定なし
        CREATININE: { hasBothSideAbnormal: false }
      }
    }
  },

  // シート名
  SHEETS: {
    PATIENT: '受診者マスタ',
    PHYSICAL: '身体測定',
    BLOOD_TEST: '血液検査',
    FINDINGS: '所見',
    JUDGMENT_MASTER: '判定マスタ',
    FINDINGS_TEMPLATE: '所見テンプレート',
    SETTINGS: '設定',
    OUTPUT_PAGE1: '出力用_1ページ',
    OUTPUT_PAGE2: '出力用_2ページ',
    OUTPUT_PAGE3: '出力用_3ページ',
    OUTPUT_PAGE4: '出力用_4ページ',
    OUTPUT_PAGE5: '出力用_5ページ',
    // 労災二次検診用
    ROSAI_INPUT: '労災二次検診_入力'
  },

  // 設定シート必須項目（労災二次検診）
  // 設定シートに以下の行を追加してください:
  // | ROSAI_CASE_FOLDER_ID | [10_案件フォルダのID] |
  //
  // フォルダID取得方法:
  // 1. Google Driveで「10_案件」フォルダを開く
  // 2. URLの /folders/ の後の文字列がフォルダID
  //    例: https://drive.google.com/drive/folders/XXXXXX → XXXXXX がID

  // CSV設定
  CSV: {
    SUPPORTED_ENCODINGS: ['Shift_JIS', 'UTF-8'],
    MIN_FIELDS: 12,
    TEST_START_INDEX: 11,
    TEST_FIELD_COUNT: 4
  },

  // 判定グレード
  JUDGMENT_GRADES: ['A', 'B', 'C', 'D'],

  // ステータス
  STATUS: {
    INPUT: '入力中',
    PENDING: '確認待ち',
    COMPLETE: '完了'
  },

  // カテゴリ
  CATEGORIES: {
    CIRCULATORY: '循環器系',
    DIGESTIVE: '消化器系',
    METABOLIC_SUGAR: '代謝系（糖）',
    METABOLIC_LIPID: '代謝系（脂質）',
    RENAL: '腎機能',
    BLOOD: '血液系',
    OTHER: 'その他'
  }
};

// ============================================
// BMLコード → 判定基準キー マッピング
// ============================================
const CODE_TO_CRITERIA = {
  '0000481': 'AST_GOT',
  '0000482': 'ALT_GPT',
  '0000484': 'GAMMA_GTP',
  '0000460': 'HDL_CHOLESTEROL',
  '0000410': 'LDL_CHOLESTEROL',
  '0000454': 'TRIGLYCERIDES',
  '0000503': 'FASTING_GLUCOSE',
  '0003317': 'HBA1C',
  '0000658': 'CRP',
  '0002696': 'EGFR',
  '0000401': 'TOTAL_PROTEIN',
  '0000407': 'URIC_ACID',
  '0000303': 'HEMOGLOBIN',
  '0000413': 'CREATININE',
  '0000301': 'WBC',
  '0000302': 'RBC',
  '0000304': 'HT',
  '0000308': 'PLT'
};

// 性別依存の検査コード
const GENDER_DEPENDENT_CODES = ['0000303', '0000413'];

// BMLコード → カテゴリ マッピング
const CODE_TO_CATEGORY = {
  // 循環器系
  'BLOOD_PRESSURE_SYS': '循環器系',
  'BLOOD_PRESSURE_DIA': '循環器系',

  // 消化器系
  '0000481': '消化器系', // AST
  '0000482': '消化器系', // ALT
  '0000484': '消化器系', // γ-GTP
  '0000401': '消化器系', // TP
  '0000472': '消化器系', // T-Bil

  // 代謝系（糖）
  '0000503': '代謝系（糖）', // FBS
  '0003317': '代謝系（糖）', // HbA1c

  // 代謝系（脂質）
  '0000460': '代謝系（脂質）', // HDL
  '0000410': '代謝系（脂質）', // LDL
  '0000454': '代謝系（脂質）', // TG

  // 腎機能
  '0000413': '腎機能', // Cr
  '0002696': '腎機能', // eGFR

  // 血液系
  '0000301': '血液系', // WBC
  '0000302': '血液系', // RBC
  '0000303': '血液系', // Hb
  '0000304': '血液系', // Ht
  '0000308': '血液系', // PLT

  // その他
  '0000407': 'その他', // UA
  '0000658': 'その他', // CRP
  'BMI': 'その他',
  'WAIST': 'その他'
};

// 性別変換
const GENDER_TRANSFORMS = {
  '1': '男',
  '2': '女',
  'M': '男',
  'F': '女'
};

// CSV性別コード → 内部表現
const GENDER_CODE_TO_INTERNAL = {
  '1': 'M',
  '2': 'F'
};

// ============================================
// ユーティリティ関数
// ============================================

/**
 * マスタースプレッドシートを取得
 * @returns {Spreadsheet}
 */
function getSpreadsheet() {
  // スプレッドシートに紐付けられている場合は自動取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    return ss;
  }
  // 外部から実行する場合はIDで取得
  return SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
}

/**
 * シートを取得
 * @param {string} sheetName - シート名
 * @returns {Sheet}
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * フォルダを取得
 * @param {string} folderId - フォルダID
 * @returns {Folder}
 */
function getFolderById(folderId) {
  return DriveApp.getFolderById(folderId);
}

/**
 * CSVフォルダを取得
 * 一時フォルダ設定（案件指定）を優先し、なければ設定シートのデフォルトを使用
 * @returns {Folder}
 */
function getCsvFolder() {
  // 一時フォルダID（案件指定）を優先
  const tempFolderId = PropertiesService.getScriptProperties().getProperty('TEMP_CSV_FOLDER_ID');
  if (tempFolderId) {
    return DriveApp.getFolderById(tempFolderId);
  }

  // デフォルトフォルダ（設定シート）
  const folderId = getSettingValue('CSV_FOLDER_ID');
  if (!folderId || folderId === 'YOUR_CSV_FOLDER_ID') {
    throw new Error('CSV_FOLDER_IDが設定されていません。\n「案件フォルダ」メニューからフォルダを選択するか、\n設定シートにCSV_FOLDER_IDを設定してください。');
  }
  return DriveApp.getFolderById(folderId);
}

/**
 * 出力フォルダを取得
 * 設定シートからフォルダIDを読み込む
 * @returns {Folder}
 */
function getOutputFolder() {
  const folderId = getSettingValue('OUTPUT_FOLDER_ID');
  if (!folderId || folderId === 'YOUR_OUTPUT_FOLDER_ID') {
    throw new Error('OUTPUT_FOLDER_IDが設定されていません。設定シートを確認してください。');
  }
  return DriveApp.getFolderById(folderId);
}

/**
 * 設定シートから値を取得
 * @param {string} key - 設定キー
 * @returns {string} 設定値
 */
function getSettingValue(key) {
  const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    throw new Error('設定シートが見つかりません');
  }

  const data = sheet.getDataRange().getValues();
  for (const row of data) {
    if (row[0] === key) {
      return row[1];
    }
  }
  return null;
}

/**
 * 日付をフォーマット
 * @param {Date} date - 日付
 * @param {string} format - フォーマット（デフォルト: YYYY/MM/DD）
 * @returns {string}
 */
function formatDate(date, format = 'YYYY/MM/DD') {
  if (!date) return '';

  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');

  return format
    .replace('YYYY', year)
    .replace('MM', month)
    .replace('DD', day);
}

/**
 * YYYYMMDD形式の文字列をYYYY/MM/DD形式に変換
 * @param {string} dateStr - YYYYMMDD形式の日付文字列
 * @returns {string}
 */
function formatDateString(dateStr) {
  if (!dateStr || dateStr.length !== 8) return dateStr;
  return `${dateStr.slice(0, 4)}/${dateStr.slice(4, 6)}/${dateStr.slice(6, 8)}`;
}

/**
 * 受診IDを生成
 * @param {Date} examDate - 受診日
 * @returns {string}
 */
function generatePatientId(examDate) {
  const dateStr = formatDate(examDate, 'YYYYMMDD');
  const timestamp = Date.now().toString().slice(-4);
  return `${dateStr}-${timestamp}`;
}

/**
 * 通知を送信
 * @param {string} subject - 件名
 * @param {string} body - 本文
 */
function sendNotification(subject, body) {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      GmailApp.sendEmail(email, subject, body);
    }
  } catch (e) {
    Logger.log('通知送信エラー: ' + e.message);
  }
}

/**
 * エラーログを記録
 * @param {string} functionName - 関数名
 * @param {Error} error - エラー
 */
function logError(functionName, error) {
  Logger.log(`[ERROR] ${functionName}: ${error.message}`);
  Logger.log(error.stack);
}

/**
 * 処理ログを記録
 * @param {string} message - メッセージ
 */
function logInfo(message) {
  Logger.log(`[INFO] ${message}`);
}

/**
 * 数値に変換（変換不可の場合はnull）
 * @param {*} value - 値
 * @returns {number|null}
 */
function toNumber(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  const num = parseFloat(value);
  return isNaN(num) ? null : num;
}

/**
 * 年齢を計算
 * @param {Date} birthDate - 生年月日
 * @param {Date} targetDate - 基準日
 * @returns {number}
 */
function calculateAge(birthDate, targetDate) {
  if (!birthDate || !targetDate) return 0;

  const birth = new Date(birthDate);
  const target = new Date(targetDate);

  let age = target.getFullYear() - birth.getFullYear();
  const monthDiff = target.getMonth() - birth.getMonth();

  if (monthDiff < 0 || (monthDiff === 0 && target.getDate() < birth.getDate())) {
    age--;
  }

  return age;
}

/**
 * BMIを計算
 * @param {number} height - 身長（cm）
 * @param {number} weight - 体重（kg）
 * @returns {number|null}
 */
function calculateBMI(height, weight) {
  if (!height || !weight || height <= 0 || weight <= 0) {
    return null;
  }
  const heightM = height / 100;
  return Math.round(weight / (heightM * heightM) * 10) / 10;
}

/**
 * 標準体重を計算
 * @param {number} height - 身長（cm）
 * @returns {number|null}
 */
function calculateStandardWeight(height) {
  if (!height || height <= 0) return null;
  const heightM = height / 100;
  return Math.round(22 * heightM * heightM * 10) / 10;
}

/**
 * 処理済みマークをファイル名に追加
 * @param {File} file - ファイル
 */
function markFileAsProcessed(file) {
  const name = file.getName();
  if (!name.startsWith('[済]')) {
    file.setName(`[済]${name}`);
  }
}

/**
 * ファイルが処理済みかチェック
 * @param {File} file - ファイル
 * @returns {boolean}
 */
function isFileProcessed(file) {
  return file.getName().startsWith('[済]');
}
