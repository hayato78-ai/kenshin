/**
 * 保健指導入力モジュール
 *
 * 労災二次検診の特定保健指導入力をサポート
 * - Part/Stepアコーディオン形式のUI
 * - Step完了フラグによる進捗管理
 * - Claude AI による指導アシスト生成
 * - 自動保存（デバウンス付き）
 *
 * @version 2.0
 * @date 2025-12-08
 */

// ============================================
// 定数
// ============================================

const GUIDANCE_SHEET_NAME = '保健指導_データ';

const GUIDANCE_COLUMNS = {
  // ============================================
  // 基本情報 (1-11)
  // ============================================
  RECORD_ID: 1,
  CREATED_AT: 2,
  UPDATED_AT: 3,
  CASE_ID: 4,
  PATIENT_NO: 5,
  PATIENT_NAME: 6,
  BIRTH_DATE: 7,        // 生年月日（H23.2.3形式）- 恒久的ID用
  AGE: 8,
  GENDER: 9,
  HEIGHT: 10,
  WEIGHT: 11,

  // ============================================
  // 質問票_就労の状況 (12-45)
  // ============================================
  // 職種
  Q_OCCUPATION_TYPE: 12,           // 屋内/屋外
  Q_OCCUPATION_DESK: 13,           // デスクワーク
  Q_OCCUPATION_INDOOR_OTHER: 14,   // 屋内その他
  Q_OCCUPATION_INDOOR_DETAIL: 15,  // 屋内その他詳細
  Q_OCCUPATION_OUTDOOR_DETAIL: 16, // 屋外作業詳細
  // 時間外労働
  Q_OVERTIME_AVG: 17,              // 直近6か月平均
  Q_OVERTIME_MAX: 18,              // 最大時間の月
  Q_OVERTIME_MIN: 19,              // 最小時間の月
  Q_OVERTIME_UNCERTAIN: 20,        // 判断困難
  // 不規則な勤務
  Q_IRREGULAR_WORK: 21,            // 有・無
  Q_IRREGULAR_WORK_DETAIL: 22,     // 具体的に
  Q_IRREGULAR_WORK_UNCERTAIN: 23,  // 判断困難
  // 出張の多い業務
  Q_BUSINESS_TRIP: 24,             // 有・無
  Q_BUSINESS_TRIP_DETAIL: 25,      // 具体的に
  Q_BUSINESS_TRIP_UNCERTAIN: 26,   // 判断困難
  // 交替制勤務・深夜勤務
  Q_SHIFT_WORK: 27,                // 有・無
  Q_SHIFT_WORK_DETAIL: 28,         // 具体的に
  Q_SHIFT_WORK_UNCERTAIN: 29,      // 判断困難
  // 高温・低温等の環境
  Q_EXTREME_TEMP: 30,              // 有・無
  Q_EXTREME_TEMP_DETAIL: 31,       // 具体的に
  Q_EXTREME_TEMP_UNCERTAIN: 32,    // 判断困難
  // 時差を伴う業務
  Q_TIME_ZONE_DIFF: 33,            // 有・無
  Q_TIME_ZONE_DIFF_DETAIL: 34,     // 具体的に
  Q_TIME_ZONE_DIFF_UNCERTAIN: 35,  // 判断困難
  // 精神的緊張を伴う業務
  Q_MENTAL_STRESS: 36,             // 有・無
  Q_MENTAL_STRESS_DETAIL: 37,      // 具体的に
  Q_MENTAL_STRESS_UNCERTAIN: 38,   // 判断困難
  Q_EXCESSIVE_QUOTA: 39,           // 過大なノルマ
  Q_EXCESSIVE_QUOTA_UNCERTAIN: 40, // 過大なノルマ判断困難
  Q_CUSTOMER_TROUBLE: 41,          // 顧客トラブル
  Q_CUSTOMER_TROUBLE_UNCERTAIN: 42,// 顧客トラブル判断困難
  Q_LIFE_CRITICAL: 43,             // 人の生命に関わる業務
  Q_LIFE_CRITICAL_UNCERTAIN: 44,   // 人の生命に関わる業務判断困難
  Q_MENTAL_STRESS_OTHER: 45,       // その他

  // ============================================
  // 質問票_通勤・休暇 (46-60)
  // ============================================
  Q_COMMUTE_CAR: 46,               // 自家用車
  Q_COMMUTE_PUBLIC: 47,            // 公共機関
  Q_COMMUTE_PUBLIC_DETAIL: 48,     // 公共機関詳細
  Q_COMMUTE_WALK: 49,              // 徒歩
  Q_COMMUTE_OTHER: 50,             // その他
  Q_COMMUTE_OTHER_DETAIL: 51,      // その他詳細
  Q_COMMUTE_TIME: 52,              // 通勤時間（分）
  Q_COMMUTE_UNCERTAIN: 53,         // 判断困難
  Q_DAYS_OFF_PER_WEEK: 54,         // 週休日数
  Q_DAYS_OFF_STATUS: 55,           // 取得状況
  Q_DAYS_OFF_UNCERTAIN: 56,        // 判断困難
  Q_PAID_LEAVE_STATUS: 57,         // 年次有給休暇取得状況
  Q_PAID_LEAVE_UNCERTAIN: 58,      // 判断困難
  Q_BREAK_TIME_STATUS: 59,         // 休憩時間取得状況
  Q_BREAK_TIME_UNCERTAIN: 60,      // 判断困難

  // ============================================
  // 質問票_睡眠・日常生活 (61-80)
  // ============================================
  Q_WORK_OTHER_NOTES: 61,          // 就労その他
  Q_SLEEP_HOURS: 62,               // 睡眠時間
  Q_MEALS_REGULAR: 63,             // 3食規則正しい
  Q_SNACKS_EXISTS: 64,             // 間食有無
  Q_SNACKS_PER_WEEK: 65,           // 間食週回数
  Q_SNACKS_PER_DAY: 66,            // 間食1日回数
  Q_ALCOHOL_DAYS: 67,              // 飲酒週日数
  Q_ALCOHOL_AMOUNT: 68,            // 1回飲酒量（合）
  Q_EXERCISE_FREQ: 69,             // 運動頻度
  Q_EXERCISE_TYPE: 70,             // 運動種目
  Q_SMOKING_STATUS: 71,            // 喫煙状況
  Q_SMOKING_PER_DAY: 72,           // 1日本数
  Q_SMOKING_YEARS: 73,             // 喫煙歴年数
  Q_WEIGHT_10Y_KG: 74,             // 10年前より増減kg
  Q_WEIGHT_10Y_DIRECTION: 75,      // 10年前より増・減
  Q_WEIGHT_20Y_KG: 76,             // 20年前より増減kg
  Q_WEIGHT_20Y_DIRECTION: 77,      // 20年前より増・減
  Q_SPECIAL_ATTENTION: 78,         // 特に注意していること
  Q_MEMO: 79,                      // メモ（自由記述）

  // ============================================
  // 検査結果 (80-95)
  // ============================================
  LAB_BP_SYS: 80,                  // 収縮期血圧
  LAB_BP_DIA: 81,                  // 拡張期血圧
  LAB_HDL: 82,                     // HDLコレステロール
  LAB_LDL: 83,                     // LDLコレステロール
  LAB_TG: 84,                      // 中性脂肪
  LAB_FBS: 85,                     // 空腹時血糖
  LAB_HBA1C: 86,                   // HbA1c
  LAB_ACR: 87,                     // 尿中アルブミン/Cre比
  LAB_EGFR: 88,                    // eGFR
  LAB_UA: 89,                      // 尿酸
  LAB_GOT: 90,                     // GOT
  LAB_GPT: 91,                     // GPT
  LAB_GGT: 92,                     // γ-GTP
  LAB_HEART_RESULT: 93,            // 心臓超音波結果
  LAB_CAROTID_RESULT: 94,          // 頸動脈超音波結果
  LAB_MEMO: 95,                    // 検査メモ

  // ============================================
  // AI出力 (96-99)
  // ============================================
  AI_GENERATED_AT: 96,
  AI_QUESTIONS: 97,
  AI_POINTS: 98,
  AI_GOALS: 99,

  // ============================================
  // 日常生活_重点指導 (100-132)
  // ============================================
  G_NUTRITION_AMOUNT: 100,
  G_NUTRITION_SALT: 101,
  G_NUTRITION_FIBER: 102,
  G_NUTRITION_OIL: 103,
  G_NUTRITION_EATING_OUT: 104,
  G_NUTRITION_EATING_OUT_TXT: 105,
  G_ALCOHOL: 106,
  G_ALCOHOL_TYPE: 107,
  G_ALCOHOL_AMOUNT: 108,
  G_ALCOHOL_FREQ: 109,
  G_SNACK: 110,
  G_SNACK_TYPE: 111,
  G_SNACK_FREQ: 112,
  G_EATING_STYLE: 113,
  G_EATING_SLOW: 114,
  G_EATING_STYLE_TXT: 115,
  G_MEAL_TIME: 116,
  G_NUTRITION_OTHER: 117,
  G_NUTRITION_OTHER_TXT: 118,
  G_EXERCISE_RX: 119,
  G_EXERCISE_TYPE: 120,
  G_EXERCISE_DURATION: 121,
  G_EXERCISE_FREQ: 122,
  G_EXERCISE_INTENSITY: 123,
  G_ACTIVITY_INCREASE: 124,
  G_ACTIVITY_TXT: 125,
  G_EXERCISE_CAUTION: 126,
  G_EXERCISE_CAUTION_TXT: 127,
  G_QUIT_SMOKING: 128,
  G_QUIT_METHOD: 129,
  G_HOME_MEASURE: 130,
  G_LIFESTYLE_OTHER: 131,
  G_LIFESTYLE_OTHER_TXT: 132,
  DAILY_PROBLEM: 133,

  // ============================================
  // 就労_重点指導 (134-142)
  // ============================================
  WG_WORK_HOURS: 134,
  WG_WORK_STYLE: 135,
  WG_ENVIRONMENT: 136,
  WG_ENVIRONMENT_TXT: 137,
  WG_SLEEP: 138,
  WG_LEISURE: 139,
  WG_OTHER: 140,
  WG_OTHER_TXT: 141,
  WORK_PROBLEM: 142,

  // ============================================
  // Step完了フラグ (143-147)
  // ============================================
  STEP_COMPLETE_DAILY_NUTRITION: 143,  // Part1-Step1: 日常生活_重点指導-栄養
  STEP_COMPLETE_DAILY_EXERCISE: 144,   // Part1-Step2: 日常生活_重点指導-運動
  STEP_COMPLETE_DAILY_LIFE: 145,       // Part1-Step3: 日常生活_重点指導-生活
  STEP_COMPLETE_WORK_GUIDANCE: 146,    // Part2-Step1: 就労_重点指導
  STEP_COMPLETE_LAB_RESULTS: 147,      // 検査結果

  // ============================================
  // ステータス (148)
  // ============================================
  STATUS: 148,

  // ============================================
  // 企業情報 (149-150)
  // ============================================
  COMPANY_ID: 149,
  COMPANY_NAME: 150
};

const GUIDANCE_HEADERS = [
  // 基本情報 (1-11)
  'record_id', 'created_at', 'updated_at', 'case_id', 'patient_no', 'patient_name', 'birth_date', 'age', 'gender', 'height', 'weight',
  // 質問票_就労の状況 (12-45)
  'q_occupation_type', 'q_occupation_desk', 'q_occupation_indoor_other', 'q_occupation_indoor_detail', 'q_occupation_outdoor_detail',
  'q_overtime_avg', 'q_overtime_max', 'q_overtime_min', 'q_overtime_uncertain',
  'q_irregular_work', 'q_irregular_work_detail', 'q_irregular_work_uncertain',
  'q_business_trip', 'q_business_trip_detail', 'q_business_trip_uncertain',
  'q_shift_work', 'q_shift_work_detail', 'q_shift_work_uncertain',
  'q_extreme_temp', 'q_extreme_temp_detail', 'q_extreme_temp_uncertain',
  'q_time_zone_diff', 'q_time_zone_diff_detail', 'q_time_zone_diff_uncertain',
  'q_mental_stress', 'q_mental_stress_detail', 'q_mental_stress_uncertain',
  'q_excessive_quota', 'q_excessive_quota_uncertain',
  'q_customer_trouble', 'q_customer_trouble_uncertain',
  'q_life_critical', 'q_life_critical_uncertain', 'q_mental_stress_other',
  // 質問票_通勤・休暇 (46-60)
  'q_commute_car', 'q_commute_public', 'q_commute_public_detail', 'q_commute_walk', 'q_commute_other', 'q_commute_other_detail',
  'q_commute_time', 'q_commute_uncertain',
  'q_days_off_per_week', 'q_days_off_status', 'q_days_off_uncertain',
  'q_paid_leave_status', 'q_paid_leave_uncertain',
  'q_break_time_status', 'q_break_time_uncertain',
  // 質問票_睡眠・日常生活 (61-79)
  'q_work_other_notes', 'q_sleep_hours',
  'q_meals_regular', 'q_snacks_exists', 'q_snacks_per_week', 'q_snacks_per_day',
  'q_alcohol_days', 'q_alcohol_amount', 'q_exercise_freq', 'q_exercise_type',
  'q_smoking_status', 'q_smoking_per_day', 'q_smoking_years',
  'q_weight_10y_kg', 'q_weight_10y_direction', 'q_weight_20y_kg', 'q_weight_20y_direction',
  'q_special_attention', 'q_memo',
  // 検査結果 (80-95)
  'lab_bp_sys', 'lab_bp_dia', 'lab_hdl', 'lab_ldl', 'lab_tg', 'lab_fbs', 'lab_hba1c', 'lab_acr',
  'lab_egfr', 'lab_ua', 'lab_got', 'lab_gpt', 'lab_ggt', 'lab_heart_result', 'lab_carotid_result', 'lab_memo',
  // AI出力 (96-99)
  'ai_generated_at', 'ai_questions', 'ai_points', 'ai_goals',
  // 日常生活_重点指導 (100-133)
  'g_nutrition_amount', 'g_nutrition_salt', 'g_nutrition_fiber', 'g_nutrition_oil', 'g_nutrition_eating_out', 'g_nutrition_eating_out_txt',
  'g_alcohol', 'g_alcohol_type', 'g_alcohol_amount', 'g_alcohol_freq',
  'g_snack', 'g_snack_type', 'g_snack_freq',
  'g_eating_style', 'g_eating_slow', 'g_eating_style_txt', 'g_meal_time',
  'g_nutrition_other', 'g_nutrition_other_txt',
  'g_exercise_rx', 'g_exercise_type', 'g_exercise_duration', 'g_exercise_freq', 'g_exercise_intensity',
  'g_activity_increase', 'g_activity_txt', 'g_exercise_caution', 'g_exercise_caution_txt',
  'g_quit_smoking', 'g_quit_method', 'g_home_measure', 'g_lifestyle_other', 'g_lifestyle_other_txt',
  'daily_problem',
  // 就労_重点指導 (134-142)
  'wg_work_hours', 'wg_work_style', 'wg_environment', 'wg_environment_txt', 'wg_sleep', 'wg_leisure', 'wg_other', 'wg_other_txt',
  'work_problem',
  // Step完了フラグ (143-147)
  'step_complete_daily_nutrition', 'step_complete_daily_exercise', 'step_complete_daily_life',
  'step_complete_work_guidance', 'step_complete_lab_results',
  // ステータス (148)
  'status',
  // 企業情報 (149-150)
  'company_id', 'company_name'
];

// ============================================
// データ整合性検証
// ============================================

/**
 * GUIDANCE_COLUMNS と GUIDANCE_HEADERS の整合性を検証
 * デバッグ・開発時に使用
 * @returns {Object} {valid, errors}
 */
function validateGuidanceDataStructure() {
  const errors = [];
  const columnKeys = Object.keys(GUIDANCE_COLUMNS);
  const expectedCount = GUIDANCE_COLUMNS.STATUS; // 最後の列番号

  // 列数チェック
  if (GUIDANCE_HEADERS.length !== expectedCount) {
    errors.push(`列数不一致: HEADERS=${GUIDANCE_HEADERS.length}, 期待値=${expectedCount}`);
  }

  // 各カラムの列番号が連続しているか確認
  const columnNumbers = Object.values(GUIDANCE_COLUMNS).sort((a, b) => a - b);
  for (let i = 0; i < columnNumbers.length; i++) {
    if (columnNumbers[i] !== i + 1) {
      errors.push(`列番号の欠損または重複: 位置${i + 1}に${columnNumbers[i]}が存在`);
      break;
    }
  }

  // カラム名とヘッダー名の対応チェック（スネークケース変換）
  columnKeys.forEach((key) => {
    const colNum = GUIDANCE_COLUMNS[key];
    const expectedHeader = key.toLowerCase();
    const actualHeader = GUIDANCE_HEADERS[colNum - 1];
    if (actualHeader !== expectedHeader) {
      // 完全一致でなくても警告レベル（命名規則の違いは許容）
      // errors.push(`名前不一致: ${key} → 期待=${expectedHeader}, 実際=${actualHeader}`);
    }
  });

  return {
    valid: errors.length === 0,
    columnCount: columnKeys.length,
    headerCount: GUIDANCE_HEADERS.length,
    errors: errors
  };
}

/**
 * シート初期化時にデータ構造を検証（開発用）
 */
function debugValidateStructure() {
  const result = validateGuidanceDataStructure();
  if (result.valid) {
    Logger.log(`✓ データ構造OK: ${result.headerCount}列`);
  } else {
    Logger.log(`✗ データ構造エラー:`);
    result.errors.forEach(e => Logger.log(`  - ${e}`));
  }
  return result;
}

/**
 * UI入力フィールド一覧を取得（GUIDANCE_HEADERSからシステムフィールドを除外）
 * クライアント側のcollectDataで使用
 * @returns {string[]} UI入力可能なフィールド名の配列
 */
function getUIInputFields() {
  // システム管理フィールド（サーバー側で設定）
  const systemFields = ['record_id', 'created_at', 'updated_at'];
  // PATIENT_DATAから取得するフィールド（collectDataで個別に処理）
  const patientDataFields = ['case_id', 'patient_no', 'patient_name', 'birth_date', 'age', 'gender', 'height', 'weight'];
  // AI出力フィールド（AI生成時に設定）
  const aiFields = ['ai_generated_at', 'ai_questions', 'ai_points', 'ai_goals'];
  // ステータス（collectDataで個別に処理）
  const statusFields = ['status'];

  const excludeFields = new Set([...systemFields, ...patientDataFields, ...aiFields, ...statusFields]);

  return GUIDANCE_HEADERS.filter(field => !excludeFields.has(field));
}

// ============================================
// メイン関数
// ============================================

/**
 * 選択行の保健指導入力ダイアログを表示
 */
function showGuidanceInputForSelectedRow() {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== '労災二次検診_入力') {
    SpreadsheetApp.getUi().alert('エラー', '労災二次検診_入力シートを選択してください', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const activeRow = ss.getActiveRange().getRow();

  if (activeRow < 6) {
    SpreadsheetApp.getUi().alert('エラー', 'データ行（6行目以降）を選択してください', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const name = sheet.getRange(activeRow, 2).getValue();
  if (!name) {
    SpreadsheetApp.getUi().alert('エラー', '選択した行にデータがありません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  showGuidanceInputDialog(activeRow);
}

/**
 * 指定行の保健指導入力ダイアログを表示（ナビゲーション用）
 * @param {number} rowIndex - 入力シートの行番号
 */
function showGuidanceInputForRow(rowIndex) {
  showGuidanceInputDialog(rowIndex);
}

/**
 * 保健指導入力ダイアログを表示
 * @param {number} rowIndex - 入力シートの行番号
 */
function showGuidanceInputDialog(rowIndex) {
  const ss = getSpreadsheet();
  const inputSheet = ss.getSheetByName('労災二次検診_入力');

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('エラー', '入力シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 案件情報を取得
  const caseInfo = inputSheet.getRange('B1').getValue();

  // 患者データを取得（19列に拡張: A-S）
  const rowData = inputSheet.getRange(rowIndex, 1, 1, 19).getValues()[0];

  // 年齢：E列の値があればそれを使用、なければ生年月日から自動計算
  const birthDateStr = rowData[3];  // D列
  let age = rowData[4];             // E列
  if (!age && birthDateStr) {
    age = calculateAgeFromWareki(birthDateStr);
  }

  const patientData = {
    rowIndex: rowIndex,
    caseId: caseInfo,
    no: rowData[0],           // A列
    name: rowData[1],         // B列
    kana: rowData[2],         // C列
    birthDate: birthDateStr,  // D列（生年月日: H23.2.3形式）
    age: age,                 // E列 or 自動計算
    gender: rowData[5],       // F列
    height: '',               // 入力シートにない場合は空
    weight: ''                // 入力シートにない場合は空
  };

  // 既存の保健指導データを取得（恒久的ID: 名前+生年月日で検索）
  const guidanceData = getGuidanceDataByPatient(caseInfo, patientData.name, patientData.birthDate);

  // 前後の行情報
  const lastRow = inputSheet.getLastRow();
  const hasPrev = rowIndex > 6;
  const hasNext = rowIndex < lastRow;

  // HTMLダイアログを生成
  const html = createGuidanceDialogHtml(patientData, guidanceData, hasPrev, hasNext);

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(1200)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '保健指導入力');
}

/**
 * 指定行の患者データと保健指導データを取得（ダイアログ内ナビゲーション用）
 * @param {number} rowIndex - 入力シートの行番号
 * @returns {Object} {patientData, guidanceData, hasPrev, hasNext}
 */
function getPatientDataForDialog(rowIndex) {
  const ss = getSpreadsheet();
  const inputSheet = ss.getSheetByName('労災二次検診_入力');

  if (!inputSheet) {
    return { error: '入力シートが見つかりません' };
  }

  const lastRow = inputSheet.getLastRow();

  // 範囲外の場合はループ（最初→最後、最後→最初）
  let targetRow = rowIndex;
  if (targetRow < 6) {
    targetRow = lastRow;
  } else if (targetRow > lastRow) {
    targetRow = 6;
  }

  // 案件情報を取得
  const caseInfo = inputSheet.getRange('B1').getValue();

  // 患者データを取得（19列に拡張: A-S）
  const rowData = inputSheet.getRange(targetRow, 1, 1, 19).getValues()[0];

  // 名前がなければ範囲外
  if (!rowData[1]) {
    return { error: 'データがありません' };
  }

  // 年齢：E列の値があればそれを使用、なければ生年月日から自動計算
  const birthDateStr = rowData[3];  // D列
  let age = rowData[4];             // E列
  if (!age && birthDateStr) {
    age = calculateAgeFromWareki(birthDateStr);
  }

  const patientData = {
    rowIndex: targetRow,
    caseId: caseInfo,
    no: rowData[0],           // A列
    name: rowData[1],         // B列
    kana: rowData[2],         // C列
    birthDate: birthDateStr,  // D列（生年月日: H23.2.3形式）
    age: age,                 // E列 or 自動計算
    gender: rowData[5],       // F列
    height: '',
    weight: ''
  };

  // 既存の保健指導データを取得（恒久的ID: 名前+生年月日で検索）
  const guidanceData = getGuidanceDataByPatient(caseInfo, patientData.name, patientData.birthDate) || {};

  return {
    patientData: patientData,
    guidanceData: guidanceData,
    hasPrev: targetRow > 6,
    hasNext: targetRow < lastRow
  };
}

// ============================================
// データ操作
// ============================================

/**
 * 保健指導データシートを取得（なければ作成）
 * @returns {Sheet}
 */
function getGuidanceDataSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(GUIDANCE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(GUIDANCE_SHEET_NAME);
    // ヘッダー設定
    sheet.getRange(1, 1, 1, GUIDANCE_HEADERS.length).setValues([GUIDANCE_HEADERS]);
    sheet.getRange(1, 1, 1, GUIDANCE_HEADERS.length)
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    logInfo('保健指導データシートを作成しました');
  }

  return sheet;
}

// ============================================
// 企業マスタ管理
// ============================================

const COMPANY_SHEET_NAME = '企業マスタ';
const COMPANY_HEADERS = [
  'company_id', 'company_name', 'company_kana', 'contact_person', 'contact_email',
  'contact_phone', 'address', 'notes', 'created_at', 'updated_at', 'is_active'
];
const COMPANY_COLUMNS = {
  COMPANY_ID: 1,
  COMPANY_NAME: 2,
  COMPANY_KANA: 3,
  CONTACT_PERSON: 4,
  CONTACT_EMAIL: 5,
  CONTACT_PHONE: 6,
  ADDRESS: 7,
  NOTES: 8,
  CREATED_AT: 9,
  UPDATED_AT: 10,
  IS_ACTIVE: 11
};

/**
 * 企業マスタシートを取得（なければ作成）
 * @returns {Sheet}
 */
function getCompanyMasterSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(COMPANY_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(COMPANY_SHEET_NAME);
    // ヘッダー設定
    sheet.getRange(1, 1, 1, COMPANY_HEADERS.length).setValues([COMPANY_HEADERS]);
    sheet.getRange(1, 1, 1, COMPANY_HEADERS.length)
      .setBackground('#388e3c')
      .setFontColor('white')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 列幅調整
    sheet.setColumnWidth(1, 80);   // company_id
    sheet.setColumnWidth(2, 200);  // company_name
    sheet.setColumnWidth(3, 150);  // company_kana
    sheet.setColumnWidth(4, 100);  // contact_person
    sheet.setColumnWidth(5, 180);  // contact_email
    sheet.setColumnWidth(6, 120);  // contact_phone
    sheet.setColumnWidth(7, 250);  // address
    sheet.setColumnWidth(8, 150);  // notes
    logInfo('企業マスタシートを作成しました');
  }

  return sheet;
}

/**
 * 新しい企業IDを生成（E001形式）
 * @returns {string}
 */
function generateCompanyId() {
  const sheet = getCompanyMasterSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'E001';
  }

  // 既存の最大IDを取得
  const ids = sheet.getRange(2, COMPANY_COLUMNS.COMPANY_ID, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0]);
    if (id.match(/^E(\d+)$/)) {
      const num = parseInt(RegExp.$1, 10);
      if (num > maxNum) maxNum = num;
    }
  }

  return 'E' + String(maxNum + 1).padStart(3, '0');
}

// 企業マスタキャッシュ
let _companyDataCache = null;
let _companyDataCacheTime = 0;
const COMPANY_CACHE_TTL_MS = 60000; // 60秒

/**
 * 企業マスタデータをキャッシュ付きで取得
 * @param {boolean} forceRefresh - 強制リフレッシュ
 * @returns {Array} 全企業データ配列
 */
function getCompanyDataWithCache(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && _companyDataCache && (now - _companyDataCacheTime) < COMPANY_CACHE_TTL_MS) {
    return _companyDataCache;
  }

  const sheet = getCompanyMasterSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    _companyDataCache = [];
    _companyDataCacheTime = now;
    return [];
  }

  _companyDataCache = sheet.getRange(2, 1, lastRow - 1, COMPANY_HEADERS.length).getValues();
  _companyDataCacheTime = now;
  return _companyDataCache;
}

/**
 * 企業マスタキャッシュを無効化
 */
function invalidateCompanyCache() {
  _companyDataCache = null;
  _companyDataCacheTime = 0;
}

/**
 * 行データから企業オブジェクトを構築
 * @param {Array} row - 行データ配列
 * @returns {Object} 企業情報オブジェクト
 */
function buildCompanyObject(row) {
  return {
    companyId: row[COMPANY_COLUMNS.COMPANY_ID - 1],
    companyName: row[COMPANY_COLUMNS.COMPANY_NAME - 1],
    companyKana: row[COMPANY_COLUMNS.COMPANY_KANA - 1],
    contactPerson: row[COMPANY_COLUMNS.CONTACT_PERSON - 1],
    contactEmail: row[COMPANY_COLUMNS.CONTACT_EMAIL - 1],
    contactPhone: row[COMPANY_COLUMNS.CONTACT_PHONE - 1],
    address: row[COMPANY_COLUMNS.ADDRESS - 1],
    notes: row[COMPANY_COLUMNS.NOTES - 1],
    isActive: row[COMPANY_COLUMNS.IS_ACTIVE - 1]
  };
}

/**
 * 企業名から企業情報を取得
 * @param {string} companyName - 企業名
 * @returns {Object|null} 企業情報
 */
function getCompanyByName(companyName) {
  if (!companyName) return null;

  try {
    const data = getCompanyDataWithCache();
    if (data.length === 0) return null;

    const normalizedName = String(companyName).trim();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][COMPANY_COLUMNS.COMPANY_NAME - 1]).trim() === normalizedName) {
        return buildCompanyObject(data[i]);
      }
    }
  } catch (e) {
    logError('getCompanyByName', e);
  }

  return null;
}

/**
 * 企業IDから企業情報を取得
 * @param {string} companyId - 企業ID
 * @returns {Object|null} 企業情報
 */
function getCompanyById(companyId) {
  if (!companyId) return null;

  try {
    const data = getCompanyDataWithCache();
    if (data.length === 0) return null;

    for (let i = 0; i < data.length; i++) {
      if (data[i][COMPANY_COLUMNS.COMPANY_ID - 1] === companyId) {
        return buildCompanyObject(data[i]);
      }
    }
  } catch (e) {
    logError('getCompanyById', e);
  }

  return null;
}

/**
 * 企業を登録または取得（存在しなければ新規作成）
 * @param {string} companyName - 企業名
 * @returns {Object} 企業情報（companyId, companyName）
 */
function registerOrGetCompany(companyName) {
  if (!companyName || String(companyName).trim() === '') {
    return { companyId: '', companyName: '' };
  }

  const normalizedName = String(companyName).trim();

  try {
    // 既存企業を検索
    const existing = getCompanyByName(normalizedName);
    if (existing) {
      return { companyId: existing.companyId, companyName: existing.companyName };
    }

    // 新規登録
    const sheet = getCompanyMasterSheet();
    const newId = generateCompanyId();
    const now = new Date();

    const newRow = [
      newId,           // company_id
      normalizedName,  // company_name
      '',              // company_kana
      '',              // contact_person
      '',              // contact_email
      '',              // contact_phone
      '',              // address
      '',              // notes
      now,             // created_at
      now,             // updated_at
      true             // is_active
    ];

    sheet.appendRow(newRow);
    invalidateCompanyCache(); // キャッシュを無効化
    logInfo(`企業を新規登録しました: ${newId} - ${normalizedName}`);

    return { companyId: newId, companyName: normalizedName };
  } catch (e) {
    logError('registerOrGetCompany', e);
    return { companyId: '', companyName: normalizedName };
  }
}

/**
 * アクティブ状態かどうかを判定
 * @param {*} value - is_activeフィールドの値
 * @returns {boolean}
 */
function isCompanyActive(value) {
  return value === true || value === 'TRUE' || value === '' || value === null || value === undefined;
}

/**
 * 全企業リストを取得（アクティブのみ）
 * @returns {Array} 企業リスト [{companyId, companyName, companyKana}, ...]
 */
function getActiveCompanyList() {
  try {
    const data = getCompanyDataWithCache();
    if (data.length === 0) return [];

    const result = [];

    for (let i = 0; i < data.length; i++) {
      if (isCompanyActive(data[i][COMPANY_COLUMNS.IS_ACTIVE - 1])) {
        result.push({
          companyId: data[i][COMPANY_COLUMNS.COMPANY_ID - 1],
          companyName: data[i][COMPANY_COLUMNS.COMPANY_NAME - 1],
          companyKana: data[i][COMPANY_COLUMNS.COMPANY_KANA - 1]
        });
      }
    }

    // 企業名でソート
    result.sort((a, b) => String(a.companyName).localeCompare(String(b.companyName), 'ja'));

    return result;
  } catch (e) {
    logError('getActiveCompanyList', e);
    return [];
  }
}

// キャッシュ（セッション中のみ有効）
let _guidanceDataCache = null;
let _guidanceDataCacheTime = 0;
const CACHE_TTL_MS = 30000; // 30秒

/**
 * 保健指導データをキャッシュ付きで取得
 * @param {boolean} forceRefresh - 強制リフレッシュ
 * @returns {Array} 全データ配列
 */
function getGuidanceDataWithCache(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && _guidanceDataCache && (now - _guidanceDataCacheTime) < CACHE_TTL_MS) {
    return _guidanceDataCache;
  }

  const sheet = getGuidanceDataSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    _guidanceDataCache = [];
    _guidanceDataCacheTime = now;
    return [];
  }

  _guidanceDataCache = sheet.getRange(2, 1, lastRow - 1, GUIDANCE_HEADERS.length).getValues();
  _guidanceDataCacheTime = now;
  return _guidanceDataCache;
}

/**
 * キャッシュを無効化（データ更新後に呼び出す）
 */
function invalidateGuidanceCache() {
  _guidanceDataCache = null;
  _guidanceDataCacheTime = 0;
}

/**
 * 保健指導データを取得（従来版: patient_noで検索、キャッシュ対応）
 * @param {string} caseId - 案件ID
 * @param {string} patientNo - 患者No
 * @returns {Object|null} 保存済みデータ
 * @deprecated getGuidanceDataByPatientを使用
 */
function getGuidanceData(caseId, patientNo) {
  const data = getGuidanceDataWithCache();
  if (data.length === 0) return null;

  for (let i = 0; i < data.length; i++) {
    if (data[i][GUIDANCE_COLUMNS.CASE_ID - 1] === caseId &&
        String(data[i][GUIDANCE_COLUMNS.PATIENT_NO - 1]) === patientNo) {
      // 行データをオブジェクトに変換
      const result = {};
      for (let j = 0; j < GUIDANCE_HEADERS.length; j++) {
        result[GUIDANCE_HEADERS[j]] = data[i][j];
      }
      result._rowIndex = i + 2; // シート上の行番号
      return result;
    }
  }

  return null;
}

/**
 * 恒久的ID（名前+生年月日）で保健指導データを取得（キャッシュ対応）
 * @param {string} caseId - 案件ID
 * @param {string} patientName - 患者名
 * @param {string} birthDate - 生年月日（H23.2.3形式）
 * @returns {Object|null} 保存済みデータ
 */
function getGuidanceDataByPatient(caseId, patientName, birthDate) {
  const data = getGuidanceDataWithCache();
  if (data.length === 0) return null;

  const normalizedName = normalizePatientName(patientName);
  const normalizedBirth = normalizeBirthDate(birthDate);

  for (let i = 0; i < data.length; i++) {
    const rowCaseId = data[i][GUIDANCE_COLUMNS.CASE_ID - 1];
    const rowName = normalizePatientName(data[i][GUIDANCE_COLUMNS.PATIENT_NAME - 1]);
    const rowBirth = normalizeBirthDate(data[i][GUIDANCE_COLUMNS.BIRTH_DATE - 1] || '');

    // 案件ID + 名前 + 生年月日で一致判定
    if (rowCaseId === caseId && rowName === normalizedName && rowBirth === normalizedBirth) {
      const result = {};
      for (let j = 0; j < GUIDANCE_HEADERS.length; j++) {
        result[GUIDANCE_HEADERS[j]] = data[i][j];
      }
      result._rowIndex = i + 2;
      return result;
    }
  }

  return null;
}

/**
 * 患者名を正規化（スペース除去、全角半角統一）
 * @param {string} name - 患者名
 * @returns {string} 正規化された名前
 */
function normalizePatientName(name) {
  if (!name) return '';
  return String(name)
    .replace(/[\s　]+/g, '')  // 全角・半角スペース除去
    .trim();
}

/**
 * 生年月日を正規化（H23.2.3 → 230203形式）
 * @param {string} birthDate - 生年月日
 * @returns {string} 正規化された生年月日（6桁数字）
 */
function normalizeBirthDate(birthDate) {
  if (!birthDate) return '';
  const str = String(birthDate);

  // H23.2.3 形式をパース
  const match = str.match(/^([MTSHR]?)(\d{1,2})\.(\d{1,2})\.(\d{1,2})$/i);
  if (match) {
    const year = match[2].padStart(2, '0');
    const month = match[3].padStart(2, '0');
    const day = match[4].padStart(2, '0');
    return year + month + day;  // 例: 230203
  }

  // すでに数字のみの場合
  const numOnly = str.replace(/\D/g, '');
  if (numOnly.length >= 6) {
    return numOnly.slice(-6);  // 末尾6桁を取得
  }

  return numOnly;
}

/**
 * 和暦生年月日から年齢を計算
 * @param {string} warekiDate - 和暦生年月日（H23.2.3形式）
 * @returns {number|null} 年齢
 */
function calculateAgeFromWareki(warekiDate) {
  if (!warekiDate) return null;

  const str = String(warekiDate);

  // H23.2.3 形式をパース
  const match = str.match(/^([MTSHR])(\d{1,2})\.(\d{1,2})\.(\d{1,2})$/i);
  if (!match) return null;

  const era = match[1].toUpperCase();
  const eraYear = parseInt(match[2], 10);
  const month = parseInt(match[3], 10);
  const day = parseInt(match[4], 10);

  // 元号から西暦年を計算
  const eraStartYears = {
    'M': 1868,  // 明治
    'T': 1912,  // 大正
    'S': 1926,  // 昭和
    'H': 1989,  // 平成
    'R': 2019   // 令和
  };

  const startYear = eraStartYears[era];
  if (!startYear) return null;

  const birthYear = startYear + eraYear - 1;
  const birthDate = new Date(birthYear, month - 1, day);

  // 年齢計算
  const today = new Date();
  let age = today.getFullYear() - birthDate.getFullYear();

  // 誕生日前なら1歳引く
  const todayMonthDay = today.getMonth() * 100 + today.getDate();
  const birthMonthDay = (month - 1) * 100 + day;
  if (todayMonthDay < birthMonthDay) {
    age--;
  }

  return age;
}

/**
 * 保健指導データを保存（永久ID対応版）
 * @param {Object} data - 入力データ
 * @returns {Object} {success, error}
 */
function saveGuidanceData(data) {
  try {
    const sheet = getGuidanceDataSheet();
    const now = new Date();

    // 永久ID（名前+生年月日）で既存データを検索
    // birth_dateがある場合は永久IDで検索、なければ従来のpatient_noで検索
    let existing = null;
    if (data.birth_date) {
      existing = getGuidanceDataByPatient(data.case_id, data.patient_name, data.birth_date);
    }
    // 永久IDで見つからない場合、従来のpatient_noでも検索（後方互換性）
    if (!existing && data.patient_no) {
      existing = getGuidanceData(data.case_id, String(data.patient_no));
    }

    // 行データを構築
    const rowData = GUIDANCE_HEADERS.map(header => {
      if (header === 'updated_at') return now;
      if (header === 'created_at') return existing ? existing.created_at : now;
      if (header === 'record_id') {
        // 永久ID形式: case_id_正規化名前_正規化生年月日_タイムスタンプ
        if (existing) return existing.record_id;
        const normalizedName = normalizePatientName(data.patient_name);
        const normalizedBirth = normalizeBirthDate(data.birth_date || '');
        return `${data.case_id}_${normalizedName}_${normalizedBirth}_${now.getTime()}`;
      }
      return data[header] !== undefined ? data[header] : '';
    });

    if (existing && existing._rowIndex) {
      // 更新
      sheet.getRange(existing._rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 新規追加
      sheet.appendRow(rowData);
    }

    // キャッシュを無効化（次回取得時に最新データを取得）
    invalidateGuidanceCache();

    return { success: true };
  } catch (e) {
    logError('saveGuidanceData', e);
    return { success: false, error: e.message };
  }
}

// ============================================
// AI連携
// ============================================

/**
 * AI指導アシストを生成
 * @param {Object} questionnaireData - 質問票データ
 * @param {Object} patientInfo - 患者情報
 * @param {string} additionalMemo - 追加メモ
 * @returns {Object} {success, questions, points, goals, error}
 */
function generateGuidanceAssist(questionnaireData, patientInfo, additionalMemo) {
  try {
    // 患者情報テキスト
    const patientText = `
氏名: ${patientInfo.name}
年齢: ${patientInfo.age}歳
性別: ${patientInfo.gender}
身長: ${patientInfo.height || '不明'}cm
体重: ${patientInfo.weight || '不明'}kg
BMI: ${patientInfo.height && patientInfo.weight ? (patientInfo.weight / Math.pow(patientInfo.height / 100, 2)).toFixed(1) : '不明'}
`.trim();

    // 質問票データテキスト（新フィールド対応）
    const occupationType = questionnaireData.q_occupation_type || '未入力';
    const occupationDetail = questionnaireData.q_occupation_desk ? 'デスクワーク' :
                             (questionnaireData.q_occupation_indoor_detail || questionnaireData.q_occupation_outdoor_detail || '');

    const questionnaireText = `
【就労状況】
職種: ${occupationType}${occupationDetail ? '（' + occupationDetail + '）' : ''}
時間外労働: ${questionnaireData.q_overtime_avg || '未入力'}時間/月（平均）
不規則勤務: ${questionnaireData.q_irregular_work || '未入力'}
交替制・深夜勤務: ${questionnaireData.q_shift_work || '未入力'}
精神的緊張: ${questionnaireData.q_mental_stress || '未入力'}

【睡眠】
睡眠時間: ${questionnaireData.q_sleep_hours || '未入力'}

【生活習慣】
喫煙: ${questionnaireData.q_smoking_status || '未入力'}${questionnaireData.q_smoking_per_day ? '（' + questionnaireData.q_smoking_per_day + '本/日）' : ''}
飲酒: ${questionnaireData.q_alcohol_days ? '週' + questionnaireData.q_alcohol_days + '日' : '未入力'}${questionnaireData.q_alcohol_amount ? '（' + questionnaireData.q_alcohol_amount + '合/回）' : ''}
運動習慣: ${questionnaireData.q_exercise_freq || '未入力'}${questionnaireData.q_exercise_type ? '（' + questionnaireData.q_exercise_type + '）' : ''}
3食規則正しい: ${questionnaireData.q_meals_regular ? 'はい' : 'いいえ'}
間食: ${questionnaireData.q_snacks_exists ? 'あり' : 'なし'}${questionnaireData.q_snacks_per_week ? '（週' + questionnaireData.q_snacks_per_week + '回）' : ''}
`.trim();

    const systemPrompt = `あなたは労災二次健診の保健指導を支援するAIアシスタントです。
検査値と生活習慣データから、保健師による指導を支援するための情報を提供します。
以下のルールに従ってください：
- 医学的に正確な表現を使用
- 簡潔で分かりやすい文章
- 具体的な数値目標を含める
- 実現可能な改善策を提案
日本語で回答してください。`;

    const userMessage = `【受診者情報】
${patientText}

【質問票データ】
${questionnaireText}

${additionalMemo ? `【追加の聞き取り内容】\n${additionalMemo}` : ''}

以下の3項目を簡潔に出力してください。各項目は箇条書きで3-5個程度にしてください。


## 目標数値案
- 達成可能な具体的目標と指導ポイントをを提案してください`;

    // Claude API呼び出し（既存のclaudeApi.gsを使用）
    const response = callClaudeApi(systemPrompt, userMessage);

    if (!response || !response.success || !response.content) {
      throw new Error(response.error || 'AIからの応答がありません');
    }

    const content = response.content;

    // レスポンスをパース
    const sections = parseAiResponse(content);

    return {
      success: true,
      questions: sections.questions || '',
      points: sections.points || '',
      goals: sections.goals || ''
    };

  } catch (e) {
    logError('generateGuidanceAssist', e);
    return { success: false, error: e.message };
  }
}

/**
 * AIレスポンスをパース
 * @param {string} content - AIレスポンス
 * @returns {Object} {questions, points, goals}
 */
function parseAiResponse(content) {
  const result = {
    questions: '',
    points: '',
    goals: ''
  };

  // セクションを分割
  const questionsMatch = content.match(/##\s*追加で確認すべき項目([\s\S]*?)(?=##|$)/);
  const pointsMatch = content.match(/##\s*重点指導ポイント([\s\S]*?)(?=##|$)/);
  const goalsMatch = content.match(/##\s*目標数値案([\s\S]*?)(?=##|$)/);

  if (questionsMatch) result.questions = questionsMatch[1].trim();
  if (pointsMatch) result.points = pointsMatch[1].trim();
  if (goalsMatch) result.goals = goalsMatch[1].trim();

  return result;
}

// ============================================
// HTML生成
// ============================================

/**
 * ダイアログHTMLを生成
 * @param {Object} patientData - 患者データ
 * @param {Object} guidanceData - 既存の指導データ
 * @param {boolean} hasPrev - 前の患者がいるか
 * @param {boolean} hasNext - 次の患者がいるか
 * @returns {string} HTML
 */
function createGuidanceDialogHtml(patientData, guidanceData, hasPrev, hasNext) {
  const data = guidanceData || {};
  const bmi = patientData.height && patientData.weight
    ? (patientData.weight / Math.pow(patientData.height / 100, 2)).toFixed(1)
    : '-';

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      background: #f5f5f5;
    }
    .header {
      background: #1a73e8;
      color: white;
      padding: 10px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .header h1 { font-size: 15px; font-weight: 500; }
    .header .patient-info { font-size: 12px; opacity: 0.9; }
    .save-status { font-size: 12px; }
    .save-status.saving { color: #ffc107; }
    .save-status.saved { color: #a5d6a7; }

    .main-container {
      display: flex;
      height: calc(100vh - 90px);
    }

    /* メインペイン */
    .main-pane {
      flex: 1;
      background: white;
      overflow-y: auto;
      padding: 12px 15px;
    }

    /* 右サイドペイン（フリーメモ） */
    .side-pane {
      width: 280px;
      background: #fafafa;
      border-left: 1px solid #ddd;
      padding: 12px;
      display: flex;
      flex-direction: column;
    }
    .side-pane h3 {
      font-size: 13px;
      color: #1a73e8;
      margin-bottom: 8px;
      padding-bottom: 5px;
      border-bottom: 2px solid #1a73e8;
    }
    .side-pane textarea {
      flex: 1;
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 12px;
      resize: none;
    }

    /* 左サイドペイン（質問票） */
    .left-pane {
      width: 320px;
      background: #fefefe;
      border-right: 1px solid #ddd;
      overflow-y: auto;
      padding: 12px;
    }
    .left-pane h3 {
      font-size: 13px;
      color: #1a73e8;
      margin-bottom: 8px;
      padding-bottom: 5px;
      border-bottom: 2px solid #1a73e8;
      position: sticky;
      top: 0;
      background: #fefefe;
      z-index: 10;
    }

    /* AI指導アシストセクション */
    .ai-assist-section {
      background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
      border: 1px solid #2196f3;
      border-radius: 8px;
      padding: 10px;
      margin-bottom: 12px;
    }
    .ai-assist-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }
    .ai-assist-header span {
      font-size: 12px;
      font-weight: bold;
      color: #1565c0;
    }
    .btn-ai-generate {
      background: #1a73e8;
      color: white;
      border: none;
      border-radius: 4px;
      padding: 4px 12px;
      font-size: 11px;
      cursor: pointer;
      transition: background 0.2s;
    }
    .btn-ai-generate:hover { background: #1557b0; }
    .btn-ai-generate:disabled {
      background: #90caf9;
      cursor: not-allowed;
    }
    .btn-ai-generate.loading {
      background: #90caf9;
      pointer-events: none;
    }
    .ai-assist-content {
      background: white;
      border-radius: 4px;
      padding: 8px;
      font-size: 11px;
      max-height: 200px;
      overflow-y: auto;
    }
    .ai-placeholder {
      color: #888;
      font-style: italic;
      text-align: center;
      padding: 10px;
    }
    .ai-result-section {
      margin-bottom: 8px;
    }
    .ai-result-section h5 {
      font-size: 11px;
      color: #1565c0;
      margin-bottom: 4px;
      padding-bottom: 2px;
      border-bottom: 1px solid #e3f2fd;
    }
    .ai-result-section ul {
      margin: 0;
      padding-left: 16px;
    }
    .ai-result-section li {
      margin-bottom: 2px;
      line-height: 1.4;
    }
    .ai-error {
      color: #d32f2f;
      text-align: center;
      padding: 10px;
    }

    .questionnaire-section {
      background: #f9f9f9;
      border-radius: 6px;
      padding: 10px;
      margin-bottom: 10px;
    }
    .questionnaire-section h4 {
      font-size: 12px;
      color: #333;
      margin-bottom: 8px;
      padding-bottom: 4px;
      border-bottom: 1px solid #e0e0e0;
    }

    /* 問題点セクション（色付け） */
    .problem-section {
      background: #fff8e1;
      border: 1px solid #ffcc02;
      border-radius: 6px;
      padding: 10px;
      margin: 10px 0;
    }
    .problem-section.daily {
      background: #e8f5e9;
      border-color: #4caf50;
    }
    .problem-section.work {
      background: #fff3e0;
      border-color: #ff9800;
    }
    .problem-section h4 {
      font-size: 12px;
      font-weight: bold;
      margin-bottom: 6px;
    }
    .problem-section.daily h4 { color: #2e7d32; }
    .problem-section.work h4 { color: #e65100; }

    /* 問題点ブロック（Stepと同列表示） */
    .problem-block {
      border-radius: 8px;
      padding: 12px 15px;
      margin: 8px 0;
    }
    .problem-block.daily {
      background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
      border-left: 4px solid #4caf50;
    }
    .problem-block.work {
      background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
      border-left: 4px solid #ff9800;
    }
    .problem-block-header {
      font-size: 13px;
      font-weight: bold;
      margin-bottom: 8px;
    }
    .problem-block.daily .problem-block-header { color: #2e7d32; }
    .problem-block.work .problem-block-header { color: #e65100; }
    .problem-block .problem-textarea {
      width: 100%;
      min-height: 80px;
      border: 1px solid rgba(0,0,0,0.15);
      border-radius: 4px;
      padding: 8px;
      font-size: 13px;
      resize: vertical;
    }

    /* スティッキーヘッダー */
    .part-header.sticky {
      position: sticky;
      top: 0;
      z-index: 20;
    }

    /* Part ヘッダー */
    .part-header {
      background: linear-gradient(135deg, #1a73e8 0%, #4285f4 100%);
      color: white;
      padding: 10px 15px;
      margin: 15px 0 10px 0;
      border-radius: 6px;
      font-weight: bold;
      font-size: 14px;
    }
    .part-header:first-child { margin-top: 0; }

    /* Step アコーディオン */
    .step-accordion {
      border: 1px solid #ddd;
      border-radius: 6px;
      margin-bottom: 8px;
      background: white;
    }
    .step-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 12px;
      background: #f8f9fa;
      cursor: pointer;
      border-radius: 6px 6px 0 0;
    }
    .step-header:hover { background: #e8eaed; }
    .step-header-left {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .step-header .toggle-icon {
      color: #666;
      font-size: 12px;
      transition: transform 0.2s;
    }
    .step-accordion.open .step-header .toggle-icon { transform: rotate(90deg); }
    .step-title { font-weight: 500; font-size: 13px; }
    .step-header-right { display: flex; align-items: center; gap: 8px; }
    .btn-complete {
      padding: 4px 12px;
      border: 1px solid #34a853;
      background: white;
      color: #34a853;
      border-radius: 4px;
      font-size: 11px;
      cursor: pointer;
    }
    .btn-complete:hover { background: #e8f5e9; }
    .btn-complete.completed {
      background: #34a853;
      color: white;
    }
    .step-content {
      display: none;
      padding: 12px;
      border-top: 1px solid #ddd;
    }
    .step-accordion.open .step-content { display: block; }

    /* フォーム要素 */
    .form-group { margin-bottom: 8px; }
    .form-group label {
      display: block;
      font-size: 11px;
      color: #666;
      margin-bottom: 2px;
    }
    .form-group select, .form-group input, .form-group textarea {
      width: 100%;
      padding: 5px 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 12px;
    }
    .form-group textarea { resize: vertical; min-height: 50px; }
    .form-row {
      display: flex;
      gap: 8px;
    }
    .form-row .form-group { flex: 1; margin-bottom: 0; }
    .form-row .form-group.small { flex: 0 0 70px; }
    .form-row .form-group.medium { flex: 0 0 100px; }

    .checkbox-inline {
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
    }
    .checkbox-inline label {
      display: flex;
      align-items: center;
      gap: 3px;
      font-size: 12px;
      color: #333;
    }

    /* 判断困難ラベル */
    .uncertain-label {
      display: flex;
      align-items: center;
      gap: 3px;
      font-size: 10px;
      color: #888;
      margin-left: auto;
      white-space: nowrap;
    }
    .uncertain-label input[type="checkbox"] {
      width: 12px;
      height: 12px;
    }

    .check-section {
      background: #f9f9f9;
      border-radius: 6px;
      padding: 10px;
      margin-bottom: 10px;
    }
    .check-section h4 {
      font-size: 12px;
      color: #333;
      margin-bottom: 8px;
      padding-bottom: 4px;
      border-bottom: 1px solid #e0e0e0;
    }
    .check-group {
      display: flex;
      flex-wrap: wrap;
      gap: 6px 12px;
      margin-bottom: 6px;
    }
    .check-group label {
      display: flex;
      align-items: center;
      gap: 3px;
      font-size: 12px;
    }
    .check-with-input {
      display: flex;
      align-items: center;
      gap: 6px;
      margin-top: 4px;
    }
    .check-with-input input[type="text"] {
      flex: 1;
      padding: 4px 6px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 11px;
    }
    .check-with-input input[type="number"] {
      width: 55px;
      padding: 4px 6px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 11px;
    }

    .work-item {
      background: white;
      border: 1px solid #e0e0e0;
      border-radius: 4px;
      padding: 8px;
      margin-bottom: 6px;
    }
    .work-item-title {
      font-weight: 500;
      font-size: 11px;
      color: #555;
      margin-bottom: 5px;
    }

    .problem-textarea {
      width: 100%;
      min-height: 60px;
      padding: 6px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 12px;
      resize: vertical;
    }

    /* 検査結果セクション（Part表示なし） */
    .lab-section {
      border: 1px solid #ddd;
      border-radius: 6px;
      margin: 15px 0;
      background: white;
    }
    .lab-header {
      padding: 10px 12px;
      background: #f8f9fa;
      border-radius: 6px 6px 0 0;
      display: flex;
      justify-content: space-between;
      align-items: center;
      cursor: pointer;
    }
    .lab-header:hover { background: #e8eaed; }
    .lab-header-left {
      display: flex;
      align-items: center;
      gap: 8px;
      font-weight: 500;
      font-size: 13px;
    }
    .lab-content {
      display: none;
      padding: 12px;
      border-top: 1px solid #ddd;
    }
    .lab-section.open .lab-content { display: block; }

    /* フッター */
    .footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 20px;
      background: white;
      border-top: 1px solid #ddd;
    }
    .btn {
      padding: 8px 20px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 13px;
    }
    .btn-primary { background: #1a73e8; color: white; }
    .btn-secondary { background: #f1f3f4; color: #333; }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .nav-buttons { display: flex; gap: 8px; }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1>${escapeHtml(patientData.name)}${patientData.kana ? '（' + escapeHtml(patientData.kana) + '）' : ''}　${patientData.age}歳・${patientData.gender}</h1>
      <div class="patient-info">
        身長: ${patientData.height || '-'}cm / 体重: ${patientData.weight || '-'}kg / BMI: ${bmi}
      </div>
    </div>
    <div class="save-status" id="saveStatus">💾 準備完了</div>
  </div>

  <div class="main-container">
    <!-- 左サイドペイン: 質問票 -->
    <div class="left-pane">
      <h3>📋 質問票（事前入力）</h3>

      <!-- AI指導アシスト -->
      <div class="ai-assist-section">
        <div class="ai-assist-header">
          <span>💡 AI指導アシスト</span>
          <button type="button" class="btn-ai-generate" id="btnGenerateAi" onclick="generateAiAssist()">生成</button>
        </div>
        <div class="ai-assist-content" id="aiAssistContent">
          <div class="ai-placeholder">質問票入力後、「生成」ボタンで指導ポイントを提案します</div>
        </div>
      </div>

      <!-- 基本情報 -->
      <div class="questionnaire-section">
        <h4>基本情報</h4>
        <div class="form-row">
          <div class="form-group">
            <label>身長 (cm)</label>
            <input type="number" id="height" value="${data.height || patientData.height || ''}" step="0.1" onchange="onInputChange()">
          </div>
          <div class="form-group">
            <label>体重 (kg)</label>
            <input type="number" id="weight" value="${data.weight || patientData.weight || ''}" step="0.1" onchange="onInputChange()">
          </div>
        </div>
      </div>

      <!-- 睡眠・食事・運動 -->
      <div class="questionnaire-section">
        <h4>睡眠・食事・運動</h4>
        <div class="form-group">
          <label>睡眠時間</label>
          <select id="q_sleep_hours" onchange="onInputChange()">
            <option value="">選択</option>
            <option value="4時間以下" ${data.q_sleep_hours === '4時間以下' ? 'selected' : ''}>4時間以下</option>
            <option value="5時間" ${data.q_sleep_hours === '5時間' ? 'selected' : ''}>5時間</option>
            <option value="6時間" ${data.q_sleep_hours === '6時間' ? 'selected' : ''}>6時間</option>
            <option value="7時間" ${data.q_sleep_hours === '7時間' ? 'selected' : ''}>7時間</option>
            <option value="8時間以上" ${data.q_sleep_hours === '8時間以上' ? 'selected' : ''}>8時間以上</option>
          </select>
        </div>
        <div class="checkbox-inline" style="margin: 6px 0;">
          <label><input type="checkbox" id="q_meals_regular" ${data.q_meals_regular ? 'checked' : ''} onchange="onInputChange()"> 3食規則正しい</label>
          <label><input type="checkbox" id="q_snacks_exists" ${data.q_snacks_exists ? 'checked' : ''} onchange="onInputChange()"> 間食あり</label>
        </div>
        <div class="form-row">
          <div class="form-group small">
            <label>間食 週</label>
            <input type="number" id="q_snacks_per_week" value="${data.q_snacks_per_week || ''}" placeholder="回" onchange="onInputChange()">
          </div>
          <div class="form-group small">
            <label>飲酒 週</label>
            <input type="number" id="q_alcohol_days" value="${data.q_alcohol_days || ''}" placeholder="日" onchange="onInputChange()">
          </div>
          <div class="form-group small">
            <label>飲酒量</label>
            <input type="number" id="q_alcohol_amount" value="${data.q_alcohol_amount || ''}" step="0.5" placeholder="合" onchange="onInputChange()">
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>運動頻度</label>
            <select id="q_exercise_freq" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="週1～2日" ${data.q_exercise_freq === '週1～2日' ? 'selected' : ''}>週1～2日</option>
              <option value="週3～4日" ${data.q_exercise_freq === '週3～4日' ? 'selected' : ''}>週3～4日</option>
              <option value="週5日以上" ${data.q_exercise_freq === '週5日以上' ? 'selected' : ''}>週5日以上</option>
              <option value="なし" ${data.q_exercise_freq === 'なし' ? 'selected' : ''}>なし</option>
            </select>
          </div>
          <div class="form-group">
            <label>運動種目</label>
            <input type="text" id="q_exercise_type" value="${escapeHtml(data.q_exercise_type || '')}" placeholder="ウォーキング等" onchange="onInputChange()">
          </div>
        </div>
      </div>

      <!-- 喫煙 -->
      <div class="questionnaire-section">
        <h4>喫煙</h4>
        <div class="form-row">
          <div class="form-group">
            <label>状況</label>
            <select id="q_smoking_status" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="有" ${data.q_smoking_status === '有' ? 'selected' : ''}>有</option>
              <option value="無" ${data.q_smoking_status === '無' ? 'selected' : ''}>無</option>
              <option value="過去に喫煙" ${data.q_smoking_status === '過去に喫煙' ? 'selected' : ''}>過去に喫煙</option>
            </select>
          </div>
          <div class="form-group small">
            <label>本/日</label>
            <input type="number" id="q_smoking_per_day" value="${data.q_smoking_per_day || ''}" onchange="onInputChange()">
          </div>
          <div class="form-group small">
            <label>喫煙歴</label>
            <input type="number" id="q_smoking_years" value="${data.q_smoking_years || ''}" placeholder="年" onchange="onInputChange()">
          </div>
        </div>
      </div>

      <!-- 体重の増減 -->
      <div class="questionnaire-section">
        <h4>体重の増減</h4>
        <div class="form-row">
          <div class="form-group small">
            <label>10年前より</label>
            <input type="number" id="q_weight_10y_kg" value="${data.q_weight_10y_kg || ''}" placeholder="kg" onchange="onInputChange()">
          </div>
          <div class="form-group small">
            <select id="q_weight_10y_direction" onchange="onInputChange()">
              <option value="">増減</option>
              <option value="増" ${data.q_weight_10y_direction === '増' ? 'selected' : ''}>増</option>
              <option value="減" ${data.q_weight_10y_direction === '減' ? 'selected' : ''}>減</option>
            </select>
          </div>
          <div class="form-group small">
            <label>20年前より</label>
            <input type="number" id="q_weight_20y_kg" value="${data.q_weight_20y_kg || ''}" placeholder="kg" onchange="onInputChange()">
          </div>
          <div class="form-group small">
            <select id="q_weight_20y_direction" onchange="onInputChange()">
              <option value="">増減</option>
              <option value="増" ${data.q_weight_20y_direction === '増' ? 'selected' : ''}>増</option>
              <option value="減" ${data.q_weight_20y_direction === '減' ? 'selected' : ''}>減</option>
            </select>
          </div>
        </div>
      </div>

      <!-- 就労状況 -->
      <div class="questionnaire-section">
        <h4>就労状況</h4>
        <div class="checkbox-inline" style="margin-bottom:6px;">
          <label><input type="radio" name="q_occupation_type" value="屋内" ${data.q_occupation_type === '屋内' ? 'checked' : ''} onchange="onInputChange()"> 屋内</label>
          <label><input type="radio" name="q_occupation_type" value="屋外" ${data.q_occupation_type === '屋外' ? 'checked' : ''} onchange="onInputChange()"> 屋外</label>
          <label><input type="checkbox" id="q_occupation_desk" ${data.q_occupation_desk ? 'checked' : ''} onchange="onInputChange()"> デスクワーク</label>
        </div>
        <div class="form-row">
          <div class="form-group small">
            <label>残業 平均</label>
            <input type="number" id="q_overtime_avg" value="${data.q_overtime_avg || ''}" placeholder="h/月" onchange="onInputChange()">
          </div>
          <div class="form-group small">
            <label>最大</label>
            <input type="number" id="q_overtime_max" value="${data.q_overtime_max || ''}" placeholder="h/月" onchange="onInputChange()">
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_overtime_uncertain" ${data.q_overtime_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>不規則な勤務</label>
            <select id="q_irregular_work" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="有" ${data.q_irregular_work === '有' ? 'selected' : ''}>有</option>
              <option value="無" ${data.q_irregular_work === '無' ? 'selected' : ''}>無</option>
            </select>
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_irregular_work_uncertain" ${data.q_irregular_work_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>交替制・深夜</label>
            <select id="q_shift_work" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="有" ${data.q_shift_work === '有' ? 'selected' : ''}>有</option>
              <option value="無" ${data.q_shift_work === '無' ? 'selected' : ''}>無</option>
            </select>
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_shift_work_uncertain" ${data.q_shift_work_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>精神的緊張</label>
            <select id="q_mental_stress" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="有" ${data.q_mental_stress === '有' ? 'selected' : ''}>有</option>
              <option value="無" ${data.q_mental_stress === '無' ? 'selected' : ''}>無</option>
            </select>
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_mental_stress_uncertain" ${data.q_mental_stress_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
        <!-- 精神的緊張の詳細 -->
        <div class="checkbox-inline mental-stress-details" style="margin:4px 0 4px 10px;font-size:11px;">
          <label><input type="checkbox" id="q_excessive_quota" ${data.q_excessive_quota ? 'checked' : ''} onchange="onInputChange()"> 過大なノルマ</label>
          <label><input type="checkbox" id="q_customer_trouble" ${data.q_customer_trouble ? 'checked' : ''} onchange="onInputChange()"> 顧客トラブル</label>
          <label><input type="checkbox" id="q_life_critical" ${data.q_life_critical ? 'checked' : ''} onchange="onInputChange()"> 人の生命に関わる</label>
        </div>
        <div class="form-group" style="margin-top:4px;">
          <label style="font-size:11px;">精神的緊張その他</label>
          <input type="text" id="q_mental_stress_other" value="${escapeHtml(data.q_mental_stress_other || '')}" placeholder="その他の緊張要因" style="font-size:11px;" onchange="onInputChange()">
        </div>
      </div>

      <!-- 通勤手段・時間 -->
      <div class="questionnaire-section">
        <h4>通勤</h4>
        <div class="checkbox-inline" style="margin-bottom:6px;">
          <label><input type="checkbox" id="q_commute_car" ${data.q_commute_car ? 'checked' : ''} onchange="onInputChange()"> 自家用車</label>
          <label><input type="checkbox" id="q_commute_public" ${data.q_commute_public ? 'checked' : ''} onchange="onInputChange()"> 公共機関</label>
          <label><input type="checkbox" id="q_commute_walk" ${data.q_commute_walk ? 'checked' : ''} onchange="onInputChange()"> 徒歩</label>
          <label><input type="checkbox" id="q_commute_other" ${data.q_commute_other ? 'checked' : ''} onchange="onInputChange()"> その他</label>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label style="font-size:11px;">公共機関詳細</label>
            <input type="text" id="q_commute_public_detail" value="${escapeHtml(data.q_commute_public_detail || '')}" placeholder="電車+バス等" style="font-size:11px;" onchange="onInputChange()">
          </div>
          <div class="form-group small">
            <label>通勤時間</label>
            <input type="number" id="q_commute_time" value="${data.q_commute_time || ''}" placeholder="分" onchange="onInputChange()">
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_commute_uncertain" ${data.q_commute_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
      </div>

      <!-- 休暇取得状況 -->
      <div class="questionnaire-section">
        <h4>休暇取得状況</h4>
        <div class="form-row">
          <div class="form-group small">
            <label>週休日数</label>
            <input type="number" id="q_days_off_per_week" value="${data.q_days_off_per_week || ''}" step="0.5" placeholder="日" onchange="onInputChange()">
          </div>
          <div class="form-group">
            <label>所定休日</label>
            <select id="q_days_off_status" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="ほぼ取得" ${data.q_days_off_status === 'ほぼ取得' ? 'selected' : ''}>ほぼ取得</option>
              <option value="半分程度" ${data.q_days_off_status === '半分程度' ? 'selected' : ''}>半分程度</option>
              <option value="ほぼ取得できず" ${data.q_days_off_status === 'ほぼ取得できず' ? 'selected' : ''}>ほぼ取得できず</option>
            </select>
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_days_off_uncertain" ${data.q_days_off_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>年次有給</label>
            <select id="q_paid_leave_status" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="ほぼ取得" ${data.q_paid_leave_status === 'ほぼ取得' ? 'selected' : ''}>ほぼ取得</option>
              <option value="半分程度" ${data.q_paid_leave_status === '半分程度' ? 'selected' : ''}>半分程度</option>
              <option value="ほぼ取得できず" ${data.q_paid_leave_status === 'ほぼ取得できず' ? 'selected' : ''}>ほぼ取得できず</option>
            </select>
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_paid_leave_uncertain" ${data.q_paid_leave_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>休憩時間</label>
            <select id="q_break_time_status" onchange="onInputChange()">
              <option value="">選択</option>
              <option value="ほぼ取得" ${data.q_break_time_status === 'ほぼ取得' ? 'selected' : ''}>ほぼ取得</option>
              <option value="半分程度" ${data.q_break_time_status === '半分程度' ? 'selected' : ''}>半分程度</option>
              <option value="ほぼ取得できず" ${data.q_break_time_status === 'ほぼ取得できず' ? 'selected' : ''}>ほぼ取得できず</option>
            </select>
          </div>
          <label class="uncertain-label"><input type="checkbox" id="q_break_time_uncertain" ${data.q_break_time_uncertain ? 'checked' : ''} onchange="onInputChange()"> 判断困難</label>
        </div>
      </div>

      <!-- 特記事項 -->
      <div class="questionnaire-section">
        <h4>特記事項</h4>
        <textarea id="q_special_attention" onchange="onInputChange()" style="width:100%;min-height:60px;font-size:12px;">${escapeHtml(data.q_special_attention || '')}</textarea>
      </div>
    </div>

    <!-- メインペイン -->
    <div class="main-pane">
      <!-- ═══════════════════════════════════════════════════════ -->
      <!-- Part1: 日常生活に関する事項 -->
      <!-- ═══════════════════════════════════════════════════════ -->
      <div class="part-header sticky">Part1: 日常生活に関する事項</div>

      <!-- Step1: 重点指導 - 栄養 -->
      <div class="step-accordion" id="step-daily-nutrition">
        <div class="step-header" onclick="toggleStep(this.parentElement)">
          <div class="step-header-left">
            <span class="toggle-icon">▶</span>
            <span class="step-title">Step1: 重点指導 - 栄養</span>
          </div>
          <div class="step-header-right">
            <button type="button" class="btn-complete ${data.step_complete_daily_nutrition ? 'completed' : ''}" onclick="event.stopPropagation(); toggleComplete('step_complete_daily_nutrition', this)">
              ${data.step_complete_daily_nutrition ? '完了済' : '完了'}
            </button>
          </div>
        </div>
        <div class="step-content" onclick="event.stopPropagation()">
          <div class="check-group">
            <label><input type="checkbox" id="g_nutrition_amount" ${data.g_nutrition_amount ? 'checked' : ''} onchange="onInputChange()"> 食事摂取量を適正にする</label>
            <label><input type="checkbox" id="g_nutrition_salt" ${data.g_nutrition_salt ? 'checked' : ''} onchange="onInputChange()"> 食塩・調味料を控える</label>
            <label><input type="checkbox" id="g_nutrition_fiber" ${data.g_nutrition_fiber ? 'checked' : ''} onchange="onInputChange()"> 食物繊維を増やす</label>
            <label><input type="checkbox" id="g_nutrition_oil" ${data.g_nutrition_oil ? 'checked' : ''} onchange="onInputChange()"> 油料理を減らす</label>
            <label><input type="checkbox" id="g_meal_time" ${data.g_meal_time ? 'checked' : ''} onchange="onInputChange()"> 食事時間を規則正しく</label>
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_nutrition_eating_out" ${data.g_nutrition_eating_out ? 'checked' : ''} onchange="onInputChange()"> 外食の注意</label>
            <input type="text" id="g_nutrition_eating_out_txt" value="${escapeHtml(data.g_nutrition_eating_out_txt || '')}" placeholder="具体的に" onchange="onInputChange()">
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_alcohol" ${data.g_alcohol ? 'checked' : ''} onchange="onInputChange()"> 節酒</label>
            <input type="text" id="g_alcohol_type" value="${escapeHtml(data.g_alcohol_type || '')}" placeholder="種類" style="width:70px" onchange="onInputChange()">
            <input type="text" id="g_alcohol_amount" value="${escapeHtml(data.g_alcohol_amount || '')}" placeholder="量" style="width:70px" onchange="onInputChange()">
            <span>週</span><input type="number" id="g_alcohol_freq" value="${data.g_alcohol_freq || ''}" style="width:45px" onchange="onInputChange()"><span>回に</span>
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_snack" ${data.g_snack ? 'checked' : ''} onchange="onInputChange()"> 間食を減らす</label>
            <input type="text" id="g_snack_type" value="${escapeHtml(data.g_snack_type || '')}" placeholder="種類" style="width:80px" onchange="onInputChange()">
            <span>週</span><input type="number" id="g_snack_freq" value="${data.g_snack_freq || ''}" style="width:45px" onchange="onInputChange()"><span>回</span>
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_eating_style" ${data.g_eating_style ? 'checked' : ''} onchange="onInputChange()"> 食べ方</label>
            <label><input type="checkbox" id="g_eating_slow" ${data.g_eating_slow ? 'checked' : ''} onchange="onInputChange()"> ゆっくり</label>
            <input type="text" id="g_eating_style_txt" value="${escapeHtml(data.g_eating_style_txt || '')}" placeholder="その他" onchange="onInputChange()">
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_nutrition_other" ${data.g_nutrition_other ? 'checked' : ''} onchange="onInputChange()"> その他</label>
            <input type="text" id="g_nutrition_other_txt" value="${escapeHtml(data.g_nutrition_other_txt || '')}" onchange="onInputChange()">
          </div>
        </div>
      </div>

      <!-- Step2: 重点指導 - 運動 -->
      <div class="step-accordion" id="step-daily-exercise">
        <div class="step-header" onclick="toggleStep(this.parentElement)">
          <div class="step-header-left">
            <span class="toggle-icon">▶</span>
            <span class="step-title">Step2: 重点指導 - 運動</span>
          </div>
          <div class="step-header-right">
            <button type="button" class="btn-complete ${data.step_complete_daily_exercise ? 'completed' : ''}" onclick="event.stopPropagation(); toggleComplete('step_complete_daily_exercise', this)">
              ${data.step_complete_daily_exercise ? '完了済' : '完了'}
            </button>
          </div>
        </div>
        <div class="step-content" onclick="event.stopPropagation()">
          <div class="check-with-input">
            <label><input type="checkbox" id="g_exercise_rx" ${data.g_exercise_rx ? 'checked' : ''} onchange="onInputChange()"> 運動処方</label>
            <input type="text" id="g_exercise_type" value="${escapeHtml(data.g_exercise_type || '')}" placeholder="種類" style="width:80px" onchange="onInputChange()">
            <input type="text" id="g_exercise_duration" value="${escapeHtml(data.g_exercise_duration || '')}" placeholder="時間" style="width:50px" onchange="onInputChange()">
            <input type="text" id="g_exercise_freq" value="${escapeHtml(data.g_exercise_freq || '')}" placeholder="頻度" style="width:50px" onchange="onInputChange()">
            <input type="text" id="g_exercise_intensity" value="${escapeHtml(data.g_exercise_intensity || '')}" placeholder="強度" style="width:60px" onchange="onInputChange()">
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_activity_increase" ${data.g_activity_increase ? 'checked' : ''} onchange="onInputChange()"> 活動量増加</label>
            <input type="text" id="g_activity_txt" value="${escapeHtml(data.g_activity_txt || '')}" placeholder="例: 1日8000歩" onchange="onInputChange()">
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_exercise_caution" ${data.g_exercise_caution ? 'checked' : ''} onchange="onInputChange()"> 運動時の注意</label>
            <input type="text" id="g_exercise_caution_txt" value="${escapeHtml(data.g_exercise_caution_txt || '')}" onchange="onInputChange()">
          </div>
        </div>
      </div>

      <!-- Step3: 重点指導 - 生活 -->
      <div class="step-accordion" id="step-daily-life">
        <div class="step-header" onclick="toggleStep(this.parentElement)">
          <div class="step-header-left">
            <span class="toggle-icon">▶</span>
            <span class="step-title">Step3: 重点指導 - 生活</span>
          </div>
          <div class="step-header-right">
            <button type="button" class="btn-complete ${data.step_complete_daily_life ? 'completed' : ''}" onclick="event.stopPropagation(); toggleComplete('step_complete_daily_life', this)">
              ${data.step_complete_daily_life ? '完了済' : '完了'}
            </button>
          </div>
        </div>
        <div class="step-content" onclick="event.stopPropagation()">
          <div class="check-group">
            <label><input type="checkbox" id="g_quit_smoking" ${data.g_quit_smoking ? 'checked' : ''} onchange="onInputChange()"> 禁煙・節煙の有効性</label>
            <label><input type="checkbox" id="g_quit_method" ${data.g_quit_method ? 'checked' : ''} onchange="onInputChange()"> 禁煙の実施方法</label>
            <label><input type="checkbox" id="g_home_measure" ${data.g_home_measure ? 'checked' : ''} onchange="onInputChange()"> 家庭での計測</label>
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="g_lifestyle_other" ${data.g_lifestyle_other ? 'checked' : ''} onchange="onInputChange()"> その他</label>
            <input type="text" id="g_lifestyle_other_txt" value="${escapeHtml(data.g_lifestyle_other_txt || '')}" onchange="onInputChange()">
          </div>
        </div>
      </div>

      <!-- 生活上の問題点（Stepと同列で強調表示） -->
      <div class="problem-block daily">
        <div class="problem-block-header">【生活上の問題点】</div>
        <textarea class="problem-textarea" id="daily_problem" onchange="onInputChange()">${escapeHtml(data.daily_problem || '')}</textarea>
      </div>

      <!-- ═══════════════════════════════════════════════════════ -->
      <!-- Part2: 就労に関する事項 -->
      <!-- ═══════════════════════════════════════════════════════ -->
      <div class="part-header sticky">Part2: 就労に関する事項</div>

      <!-- Step1: 重点指導 -->
      <div class="step-accordion" id="step-work-guidance">
        <div class="step-header" onclick="toggleStep(this.parentElement)">
          <div class="step-header-left">
            <span class="toggle-icon">▶</span>
            <span class="step-title">Step1: 重点指導</span>
          </div>
          <div class="step-header-right">
            <button type="button" class="btn-complete ${data.step_complete_work_guidance ? 'completed' : ''}" onclick="event.stopPropagation(); toggleComplete('step_complete_work_guidance', this)">
              ${data.step_complete_work_guidance ? '完了済' : '完了'}
            </button>
          </div>
        </div>
        <div class="step-content" onclick="event.stopPropagation()">
          <div class="check-group">
            <label><input type="checkbox" id="wg_work_hours" ${data.wg_work_hours ? 'checked' : ''} onchange="onInputChange()"> 労働時間</label>
            <label><input type="checkbox" id="wg_work_style" ${data.wg_work_style ? 'checked' : ''} onchange="onInputChange()"> 勤務形態</label>
            <label><input type="checkbox" id="wg_sleep" ${data.wg_sleep ? 'checked' : ''} onchange="onInputChange()"> 睡眠の確保</label>
            <label><input type="checkbox" id="wg_leisure" ${data.wg_leisure ? 'checked' : ''} onchange="onInputChange()"> 余暇</label>
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="wg_environment" ${data.wg_environment ? 'checked' : ''} onchange="onInputChange()"> 作業環境</label>
            <input type="text" id="wg_environment_txt" value="${escapeHtml(data.wg_environment_txt || '')}" onchange="onInputChange()">
          </div>
          <div class="check-with-input">
            <label><input type="checkbox" id="wg_other" ${data.wg_other ? 'checked' : ''} onchange="onInputChange()"> その他</label>
            <input type="text" id="wg_other_txt" value="${escapeHtml(data.wg_other_txt || '')}" onchange="onInputChange()">
          </div>
        </div>
      </div>

      <!-- 就労上の問題点（Stepと同列で強調表示） -->
      <div class="problem-block work">
        <div class="problem-block-header">【就労上の問題点】</div>
        <textarea class="problem-textarea" id="work_problem" onchange="onInputChange()">${escapeHtml(data.work_problem || '')}</textarea>
      </div>

      <!-- ═══════════════════════════════════════════════════════ -->
      <!-- Part3: 検査結果 -->
      <!-- ═══════════════════════════════════════════════════════ -->
      <div class="part-header sticky">Part3: 検査結果</div>

      <div class="step-accordion open" id="lab-section">
        <div class="step-header" onclick="toggleStep(this.parentElement)">
          <div class="step-header-left">
            <span class="toggle-icon">▶</span>
            <span class="step-title">検査データ入力</span>
          </div>
          <div class="step-header-right">
            <button type="button" class="btn-complete ${data.step_complete_lab_results ? 'completed' : ''}" onclick="event.stopPropagation(); toggleComplete('step_complete_lab_results', this)">
              ${data.step_complete_lab_results ? '完了済' : '完了'}
            </button>
          </div>
        </div>
        <div class="step-content" onclick="event.stopPropagation()">
          <div class="check-section">
            <h4>血圧</h4>
            <div class="form-row">
              <div class="form-group medium">
                <label>収縮期 (mmHg)</label>
                <input type="number" id="lab_bp_sys" value="${data.lab_bp_sys || ''}" onchange="onInputChange()">
              </div>
              <div class="form-group medium">
                <label>拡張期 (mmHg)</label>
                <input type="number" id="lab_bp_dia" value="${data.lab_bp_dia || ''}" onchange="onInputChange()">
              </div>
            </div>
          </div>

          <div class="check-section">
            <h4>脂質</h4>
            <div class="form-row">
              <div class="form-group">
                <label>HDL</label>
                <input type="number" id="lab_hdl" value="${data.lab_hdl || ''}" onchange="onInputChange()">
              </div>
              <div class="form-group">
                <label>LDL</label>
                <input type="number" id="lab_ldl" value="${data.lab_ldl || ''}" onchange="onInputChange()">
              </div>
              <div class="form-group">
                <label>TG</label>
                <input type="number" id="lab_tg" value="${data.lab_tg || ''}" onchange="onInputChange()">
              </div>
            </div>
          </div>

          <div class="check-section">
            <h4>血糖</h4>
            <div class="form-row">
              <div class="form-group">
                <label>空腹時血糖</label>
                <input type="number" id="lab_fbs" value="${data.lab_fbs || ''}" onchange="onInputChange()">
              </div>
              <div class="form-group">
                <label>HbA1c</label>
                <input type="number" id="lab_hba1c" value="${data.lab_hba1c || ''}" step="0.1" onchange="onInputChange()">
              </div>
            </div>
          </div>

          <div class="check-section">
            <h4>腎機能</h4>
            <div class="form-row">
              <div class="form-group">
                <label>ACR</label>
                <input type="number" id="lab_acr" value="${data.lab_acr || ''}" step="0.1" onchange="onInputChange()">
              </div>
              <div class="form-group">
                <label>eGFR</label>
                <input type="number" id="lab_egfr" value="${data.lab_egfr || ''}" step="0.1" onchange="onInputChange()">
              </div>
              <div class="form-group">
                <label>尿酸</label>
                <input type="number" id="lab_ua" value="${data.lab_ua || ''}" step="0.1" onchange="onInputChange()">
              </div>
            </div>
          </div>

          <div class="check-section">
            <h4>肝機能</h4>
            <div class="form-row">
              <div class="form-group">
                <label>GOT</label>
                <input type="number" id="lab_got" value="${data.lab_got || ''}" onchange="onInputChange()">
              </div>
              <div class="form-group">
                <label>GPT</label>
                <input type="number" id="lab_gpt" value="${data.lab_gpt || ''}" onchange="onInputChange()">
              </div>
              <div class="form-group">
                <label>γ-GTP</label>
                <input type="number" id="lab_ggt" value="${data.lab_ggt || ''}" onchange="onInputChange()">
              </div>
            </div>
          </div>

          <div class="check-section">
            <h4>超音波検査</h4>
            <div class="form-group">
              <label>心臓超音波</label>
              <textarea id="lab_heart_result" placeholder="所見..." onchange="onInputChange()" style="min-height:50px;">${escapeHtml(data.lab_heart_result || '')}</textarea>
            </div>
            <div class="form-group">
              <label>頸動脈超音波</label>
              <textarea id="lab_carotid_result" placeholder="所見..." onchange="onInputChange()" style="min-height:50px;">${escapeHtml(data.lab_carotid_result || '')}</textarea>
            </div>
            <div class="form-group">
              <label>検査メモ</label>
              <textarea id="lab_memo" placeholder="その他検査関連メモ..." onchange="onInputChange()" style="min-height:40px;">${escapeHtml(data.lab_memo || '')}</textarea>
            </div>
          </div>
        </div>
      </div>
    </div><!-- /main-pane -->

    <!-- 右サイドペイン: フリーメモ -->
    <div class="side-pane">
      <h3>📝 フリーメモ</h3>
      <textarea id="q_memo" placeholder="指導中に聞き取った内容、メモなど..." onchange="onInputChange()">${escapeHtml(data.q_memo || '')}</textarea>
    </div>
  </div>

  <div class="footer">
    <div class="nav-buttons">
      <button type="button" class="btn btn-secondary" onclick="goToPrev()">← 前</button>
      <button type="button" class="btn btn-secondary" onclick="goToNext()">次 →</button>
    </div>
    <button type="button" class="btn btn-primary" onclick="closeDialog()">閉じる</button>
  </div>

  <script>
    let PATIENT_DATA = ${JSON.stringify(patientData)};
    const UI_INPUT_FIELDS = ${JSON.stringify(getUIInputFields())};
    let saveTimeout = null;
    let isSaving = false;

    // Step アコーディオン切り替え
    function toggleStep(el) {
      el.classList.toggle('open');
    }

    // AI指導アシスト生成
    function generateAiAssist() {
      const btn = document.getElementById('btnGenerateAi');
      const content = document.getElementById('aiAssistContent');

      // ボタンをローディング状態に
      btn.classList.add('loading');
      btn.textContent = '生成中...';
      content.innerHTML = '<div class="ai-placeholder">AIが指導ポイントを分析中...</div>';

      // 質問票データを収集
      const questionnaireData = collectData();
      const memo = document.getElementById('q_memo') ? document.getElementById('q_memo').value : '';

      google.script.run
        .withSuccessHandler(function(result) {
          btn.classList.remove('loading');
          btn.textContent = '生成';

          if (result.success) {
            let html = '';
            if (result.questions) {
              html += '<div class="ai-result-section"><h5>📝 追加確認項目</h5><ul>';
              result.questions.split('\\n').filter(q => q.trim()).forEach(q => {
                html += '<li>' + escapeHtmlClient(q.replace(/^[-・]\\s*/, '')) + '</li>';
              });
              html += '</ul></div>';
            }
            if (result.points) {
              html += '<div class="ai-result-section"><h5>🎯 重点指導ポイント</h5><ul>';
              result.points.split('\\n').filter(p => p.trim()).forEach(p => {
                html += '<li>' + escapeHtmlClient(p.replace(/^[-・]\\s*/, '')) + '</li>';
              });
              html += '</ul></div>';
            }
            if (result.goals) {
              html += '<div class="ai-result-section"><h5>📊 目標数値案</h5><ul>';
              result.goals.split('\\n').filter(g => g.trim()).forEach(g => {
                html += '<li>' + escapeHtmlClient(g.replace(/^[-・]\\s*/, '')) + '</li>';
              });
              html += '</ul></div>';
            }
            content.innerHTML = html || '<div class="ai-placeholder">結果を取得できませんでした</div>';
          } else {
            content.innerHTML = '<div class="ai-error">エラー: ' + (result.error || '不明なエラー') + '</div>';
          }
        })
        .withFailureHandler(function(error) {
          btn.classList.remove('loading');
          btn.textContent = '生成';
          content.innerHTML = '<div class="ai-error">エラー: ' + error.message + '</div>';
        })
        .generateGuidanceAssist(questionnaireData, PATIENT_DATA, memo);
    }

    // クライアント側HTMLエスケープ
    function escapeHtmlClient(str) {
      if (!str) return '';
      return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // 完了ボタントグル
    function toggleComplete(fieldId, btn) {
      const isCompleted = btn.classList.contains('completed');
      if (isCompleted) {
        btn.classList.remove('completed');
        btn.textContent = '完了';
      } else {
        btn.classList.add('completed');
        btn.textContent = '完了済';
      }
      onInputChange();
    }

    // 入力変更時（デバウンス付き自動保存）
    function onInputChange() {
      if (saveTimeout) clearTimeout(saveTimeout);
      updateSaveStatus('saving');
      saveTimeout = setTimeout(saveData, 500);
    }

    // 保存状態表示
    function updateSaveStatus(status) {
      const el = document.getElementById('saveStatus');
      if (status === 'saving') {
        el.textContent = '💾 保存中...';
        el.className = 'save-status saving';
      } else if (status === 'saved') {
        el.textContent = '💾 保存済み ✓';
        el.className = 'save-status saved';
      }
    }

    // データ収集
    function collectData() {
      const data = {
        case_id: PATIENT_DATA.caseId,
        patient_no: PATIENT_DATA.no,
        patient_name: PATIENT_DATA.name,
        birth_date: PATIENT_DATA.birthDate || '',
        age: PATIENT_DATA.age,
        gender: PATIENT_DATA.gender,
        height: document.getElementById('height').value || PATIENT_DATA.height,
        weight: document.getElementById('weight').value || PATIENT_DATA.weight,
        status: '入力中'
      };

      // 職種タイプ（ラジオボタン）
      const occupationRadio = document.querySelector('input[name="q_occupation_type"]:checked');
      if (occupationRadio) data.q_occupation_type = occupationRadio.value;

      // 完了フラグを収集（新構造: 問診Step削除済み）
      data.step_complete_daily_nutrition = document.querySelector('#step-daily-nutrition .btn-complete').classList.contains('completed');
      data.step_complete_daily_exercise = document.querySelector('#step-daily-exercise .btn-complete').classList.contains('completed');
      data.step_complete_daily_life = document.querySelector('#step-daily-life .btn-complete').classList.contains('completed');
      data.step_complete_work_guidance = document.querySelector('#step-work-guidance .btn-complete').classList.contains('completed');
      data.step_complete_lab_results = document.querySelector('#lab-section .btn-complete').classList.contains('completed');

      // UI_INPUT_FIELDSから全フィールドを自動収集
      UI_INPUT_FIELDS.forEach(field => {
        if (field === 'q_occupation_type') return;
        if (field.startsWith('step_complete_')) return;

        const el = document.getElementById(field);
        if (el) {
          if (el.type === 'checkbox') {
            data[field] = el.checked;
          } else {
            data[field] = el.value;
          }
        }
      });

      return data;
    }

    // 保存実行（Promise版）
    function saveDataAsync() {
      return new Promise(function(resolve, reject) {
        const data = collectData();
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              updateSaveStatus('saved');
              resolve(result);
            } else {
              reject(new Error(result.error));
            }
          })
          .withFailureHandler(reject)
          .saveGuidanceData(data);
      });
    }

    // 保存実行（自動保存用）
    function saveData() {
      if (isSaving) return;
      isSaving = true;

      const data = collectData();

      google.script.run
        .withSuccessHandler(function(result) {
          isSaving = false;
          if (result.success) {
            updateSaveStatus('saved');
          } else {
            alert('保存エラー: ' + result.error);
          }
        })
        .withFailureHandler(function(error) {
          isSaving = false;
          alert('保存エラー: ' + error.message);
        })
        .saveGuidanceData(data);
    }

    // AI生成
    function generateAi() {
      const btn = document.getElementById('btnAiGenerate');
      btn.disabled = true;
      btn.textContent = '生成中...';

      const questionnaireData = {
        q_sleep_hours: document.getElementById('q_sleep_hours').value,
        q_smoking_status: document.getElementById('q_smoking_status').value,
        q_alcohol_days: document.getElementById('q_alcohol_days').value,
        q_exercise_freq: document.getElementById('q_exercise_freq').value,
        q_meals_regular: document.getElementById('q_meals_regular').checked,
        q_snacks_exists: document.getElementById('q_snacks_exists').checked
      };

      const additionalMemo = document.getElementById('q_memo').value;

      google.script.run
        .withSuccessHandler(function(result) {
          btn.disabled = false;
          btn.textContent = '指導アシスト再生成';

          if (result.success) {
            document.getElementById('aiOutput').innerHTML =
              '<div class="ai-output"><h4>📌 追加で確認すべき項目</h4><pre>' + escapeHtmlJs(result.questions) + '</pre></div>' +
              '<div class="ai-output"><h4>📌 重点指導ポイント</h4><pre>' + escapeHtmlJs(result.points) + '</pre></div>' +
              '<div class="ai-output"><h4>📌 目標数値案</h4><pre>' + escapeHtmlJs(result.goals) + '</pre></div>';

            const data = collectData();
            data.ai_generated_at = new Date().toISOString();
            data.ai_questions = result.questions;
            data.ai_points = result.points;
            data.ai_goals = result.goals;
            google.script.run.saveGuidanceData(data);
          } else {
            alert('生成エラー: ' + result.error);
          }
        })
        .withFailureHandler(function(error) {
          btn.disabled = false;
          btn.textContent = '指導アシスト生成';
          alert('生成エラー: ' + error.message);
        })
        .generateGuidanceAssist(questionnaireData, PATIENT_DATA, additionalMemo);
    }

    function escapeHtmlJs(text) {
      if (!text) return '';
      return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // ナビゲーション
    function goToPrev() { navigateTo(PATIENT_DATA.rowIndex - 1); }
    function goToNext() { navigateTo(PATIENT_DATA.rowIndex + 1); }

    let isNavigating = false;

    function navigateTo(newRowIndex) {
      if (isNavigating) return;
      isNavigating = true;
      document.querySelectorAll('.nav-buttons button').forEach(b => b.disabled = true);
      updateSaveStatus('saving');

      if (saveTimeout) {
        clearTimeout(saveTimeout);
        saveTimeout = null;
      }

      saveDataAsync()
        .catch(function(err) { console.log('保存スキップ:', err.message); })
        .finally(function() { fetchPatientData(newRowIndex); });
    }

    function fetchPatientData(newRowIndex) {
      // 直接新しいダイアログを開く（現在のダイアログは自動で閉じられる）
      google.script.run
        .withSuccessHandler(function() {
          // 新しいダイアログが開いたら現在のダイアログを閉じる
          google.script.host.close();
        })
        .withFailureHandler(function(error) {
          isNavigating = false;
          alert('エラー: ' + error.message);
          document.querySelectorAll('.nav-buttons button').forEach(b => b.disabled = false);
        })
        .showGuidanceInputForRow(newRowIndex);
    }

    function closeDialog() {
      saveData();
      setTimeout(function() { google.script.host.close(); }, 300);
    }
  </script>
</body>
</html>
`;
}

// ============================================
// A31セル転記関数（労災二次検診用）
// ============================================

/**
 * 指導項目フィールドとラベルのマッピング
 */
const GUIDANCE_FIELD_LABELS = {
  // 栄養指導
  g_nutrition_amount: '食事摂取量を適正にする',
  g_nutrition_salt: '食塩・調味料を控える',
  g_nutrition_fiber: '食物繊維を増やす',
  g_nutrition_oil: '油料理を減らす',
  g_meal_time: '食事時間を規則正しく',
  g_nutrition_eating_out: '外食の注意',
  g_alcohol: '節酒',
  g_snack: '間食を減らす',
  g_eating_style: '食べ方',
  g_nutrition_other: 'その他（栄養）',
  // 運動指導
  g_exercise_rx: '運動処方',
  g_activity_increase: '活動量増加',
  g_exercise_caution: '運動時の注意',
  // 生活指導
  g_quit_smoking: '禁煙・節煙の有効性',
  g_quit_method: '禁煙の実施方法',
  g_home_measure: '家庭での計測',
  g_lifestyle_other: 'その他（生活）',
  // 就労指導
  wg_work_hours: '労働時間',
  wg_work_style: '勤務形態',
  wg_sleep: '睡眠の確保',
  wg_leisure: '余暇',
  wg_environment: '作業環境',
  wg_other: 'その他（就労）'
};

// ヘッダーインデックスキャッシュ（起動時に一度だけ構築）
let _headerIndexMap = null;

/**
 * ヘッダーインデックスマップを取得（キャッシュ付き）
 * @returns {Object} ヘッダー名→インデックスのマップ
 */
function getHeaderIndexMap() {
  if (_headerIndexMap) return _headerIndexMap;

  _headerIndexMap = {};
  for (let i = 0; i < GUIDANCE_HEADERS.length; i++) {
    _headerIndexMap[GUIDANCE_HEADERS[i]] = i;
  }
  return _headerIndexMap;
}

/**
 * 値が有効（チェック済み）かどうかを判定
 * @param {*} value - 判定する値
 * @returns {boolean}
 */
function isGuidanceFieldChecked(value) {
  return value && value !== '' && value !== false && value !== 'FALSE';
}

/**
 * 保健指導データからA31セル用のテキストを生成
 * @param {Array} guidanceData - 保健指導データ配列
 * @returns {string} A31セル用のテキスト
 */
function generateA31CellText(guidanceData) {
  if (!guidanceData) return '';

  const indexMap = getHeaderIndexMap();
  const sections = [];

  // 1. 指導項目セクション
  const guidanceItems = [];

  for (const [field, label] of Object.entries(GUIDANCE_FIELD_LABELS)) {
    const headerIndex = indexMap[field];
    if (headerIndex === undefined) continue;

    const value = guidanceData[headerIndex];
    if (!isGuidanceFieldChecked(value)) continue;

    let itemText = label;

    // 運動処方の場合は詳細情報を追加
    if (field === 'g_exercise_rx') {
      const details = [];
      const typeIdx = indexMap['g_exercise_type'];
      const durationIdx = indexMap['g_exercise_duration'];
      const freqIdx = indexMap['g_exercise_freq'];

      if (typeIdx !== undefined && guidanceData[typeIdx]) {
        details.push(guidanceData[typeIdx]);
      }
      if (durationIdx !== undefined && guidanceData[durationIdx]) {
        details.push(guidanceData[durationIdx] + '分');
      }
      if (freqIdx !== undefined && guidanceData[freqIdx]) {
        details.push('週' + guidanceData[freqIdx] + '回');
      }

      if (details.length > 0) {
        itemText += '（' + details.join(' ') + '）';
      }
    }
    // その他テキストがある場合
    else {
      const txtIndex = indexMap[field + '_txt'];
      if (txtIndex !== undefined && guidanceData[txtIndex]) {
        itemText += '（' + guidanceData[txtIndex] + '）';
      }
    }

    guidanceItems.push('・' + itemText);
  }

  if (guidanceItems.length > 0) {
    sections.push('■指導項目\n' + guidanceItems.join('\n'));
  }

  // 2. 生活上の問題点セクション
  const dailyProblemIdx = indexMap['daily_problem'];
  if (dailyProblemIdx !== undefined && guidanceData[dailyProblemIdx]) {
    sections.push('■生活上の問題点\n' + guidanceData[dailyProblemIdx]);
  }

  // 3. 就労上の問題点セクション
  const workProblemIdx = indexMap['work_problem'];
  if (workProblemIdx !== undefined && guidanceData[workProblemIdx]) {
    sections.push('■就労上の問題点\n' + guidanceData[workProblemIdx]);
  }

  return sections.join('\n\n');
}

/**
 * 患者IDから保健指導データを取得してA31テキストを生成
 * @param {string} caseId - 案件ID
 * @param {string} patientName - 患者名
 * @param {string} birthDate - 生年月日
 * @returns {string} A31セル用のテキスト
 */
function getA31TextForPatient(caseId, patientName, birthDate) {
  const guidanceData = getGuidanceDataByPatient(caseId, patientName, birthDate);
  if (!guidanceData) return '';

  // guidanceDataからヘッダーインデックスに基づく配列を構築
  const dataArray = GUIDANCE_HEADERS.map(header => guidanceData[header] || '');
  return generateA31CellText(dataArray);
}

/**
 * 保健指導データIDからA31テキストを生成
 * @param {string} recordId - 保健指導レコードID
 * @returns {string} A31セル用のテキスト
 */
function getA31TextByRecordId(recordId) {
  const sheet = getGuidanceDataSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  const data = sheet.getRange(2, 1, lastRow - 1, GUIDANCE_HEADERS.length).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][GUIDANCE_COLUMNS.RECORD_ID - 1] === recordId) {
      return generateA31CellText(data[i]);
    }
  }

  return '';
}
