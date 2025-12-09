/**
 * AppSheet ブリッジモジュール
 *
 * 既存GAS関数とAppSheetシステムを連携
 * appSheetWebhook.gs から呼び出される統合関数
 */

// ============================================
// 設定定数
// ============================================

const APPSHEET_CONFIG = {
  // AS_設定シートから読み込み（またはここで直接設定）
  SPREADSHEET_ID: null,  // セットアップ時に自動取得

  // シート名（AS_プレフィックス）
  SHEETS: {
    CASES: 'AS_案件',
    PATIENTS: 'AS_受診者',
    BLOOD_TESTS: 'AS_血液検査',
    ULTRASOUND: 'AS_超音波',
    GUIDANCE: 'AS_保健指導',
    WORKFLOW: 'AS_ワークフロー',
    FINDINGS_TEMPLATE: 'AS_所見テンプレート',
    JUDGMENT_CRITERIA: 'AS_判定基準',
    SETTINGS: 'AS_設定'
  }
};

// ============================================
// 既存関数ラッパー
// ============================================

/**
 * 既存のprocessCsvFile()をAppSheet用にラップ
 * @param {File} file - CSVファイル
 * @returns {Object} 処理結果
 */
function processCsvFileForAppSheet(file) {
  // 既存関数を呼び出し
  const result = processCsvFile(file);

  if (!result.success) {
    return result;
  }

  // AppSheet用のpatient_idリストを構築
  const patientIds = result.savedPatients || [];

  return {
    success: true,
    patientIds: patientIds,
    processedCount: patientIds.length,
    errors: result.errors || []
  };
}

/**
 * 既存のfindLatestJudgmentCsv()を使用してCSVを検索
 * @param {string} folderId - フォルダID
 * @returns {Object} 検索結果
 */
function findCsvInFolder(folderId) {
  try {
    // 既存関数を呼び出し
    return findLatestJudgmentCsv(folderId);
  } catch (error) {
    logError('findCsvInFolder', error);
    return { success: false, error: error.message };
  }
}

/**
 * 既存のgetRosaiCaseList()をラップ
 * @returns {Array} 案件リスト
 */
function getCaseListFromExisting() {
  try {
    return getRosaiCaseList();
  } catch (error) {
    logError('getCaseListFromExisting', error);
    return [];
  }
}

/**
 * 既存のjudge()関数を使用
 * @param {string} itemCode - 検査項目コード
 * @param {number} value - 検査値
 * @param {string} gender - 性別 (M/F)
 * @returns {string} 判定結果 (A/B/C/D)
 */
function calculateJudgment(itemCode, value, gender) {
  return judge(itemCode, toNumber(value), gender);
}

/**
 * 既存のgetHighLowFlag()を使用
 * @param {string} itemCode - 検査項目コード
 * @param {number} value - 検査値
 * @param {string} gender - 性別
 * @returns {string} H/L/空文字
 */
function calculateHighLowFlag(itemCode, value, gender) {
  return getHighLowFlag(itemCode, value, gender);
}

// ============================================
// AppSheetスプレッドシート操作
// ============================================

/**
 * AppSheet用スプレッドシートを取得
 * @returns {Spreadsheet}
 */
function getAppSheetSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // AS_設定シートがあるかチェック
  const settingsSheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) {
    throw new Error('AS_設定シートが見つかりません。setupAppSheetTables()を実行してください。');
  }

  return ss;
}

/**
 * 設定値を取得
 * @param {string} key - 設定キー
 * @returns {string} 設定値
 */
function getAppSheetSetting(key) {
  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.SETTINGS);

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }

  return null;
}

/**
 * 設定値を保存
 * @param {string} key - 設定キー
 * @param {string} value - 設定値
 */
function setAppSheetSetting(key, value) {
  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.SETTINGS);

  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }

  // キーが見つからない場合は新規追加
  sheet.appendRow([key, value, '', '']);
  return true;
}

// ============================================
// データ変換関数
// ============================================

/**
 * 既存データをAppSheetフォーマットに変換
 * @param {Object} existingData - 既存形式のデータ
 * @returns {Object} AppSheet形式のデータ
 */
function convertToAppSheetFormat(existingData) {
  return {
    // 受診者
    patient: {
      patient_id: existingData.patientId || existingData.id,
      patient_no: existingData.chartNo || existingData.patientNo,
      name: existingData.name,
      name_kana: existingData.nameKana || existingData.kana,
      gender: convertGender(existingData.gender),
      birth_date: existingData.birthDate,
      age: existingData.age,
      exam_date: existingData.examDate,
      status: '未入力',
      current_step: 'STEP_1',
      created_at: new Date(),
      updated_at: new Date()
    },

    // 血液検査
    blood_test: convertBloodTestData(existingData),

    // 超音波（初期値）
    ultrasound: {
      ultrasound_id: `US_${existingData.patientId || Date.now()}`,
      patient_id: existingData.patientId,
      abd_judgment: '',
      abd_findings: '',
      carotid_judgment: '',
      carotid_findings: '',
      created_at: new Date()
    }
  };
}

/**
 * 血液検査データを変換
 * @param {Object} data - 元データ
 * @returns {Object} AppSheet形式
 */
function convertBloodTestData(data) {
  const gender = convertGender(data.gender);
  const bloodTest = data.bloodTest || data.testResults || {};

  // 検査項目マッピング
  const items = {
    fbs: { code: 'FASTING_GLUCOSE', value: bloodTest.fbs || bloodTest.FBS },
    hba1c: { code: 'HBA1C', value: bloodTest.hba1c || bloodTest.HBA1C },
    hdl: { code: 'HDL_CHOLESTEROL', value: bloodTest.hdl || bloodTest.HDL },
    ldl: { code: 'LDL_CHOLESTEROL', value: bloodTest.ldl || bloodTest.LDL },
    tg: { code: 'TRIGLYCERIDES', value: bloodTest.tg || bloodTest.TG },
    ast: { code: 'AST_GOT', value: bloodTest.ast || bloodTest.AST },
    alt: { code: 'ALT_GPT', value: bloodTest.alt || bloodTest.ALT },
    ggt: { code: 'GAMMA_GTP', value: bloodTest.ggt || bloodTest.GGT },
    cr: { code: 'CREATININE', value: bloodTest.cr || bloodTest.Cr },
    egfr: { code: 'EGFR', value: bloodTest.egfr || bloodTest.eGFR },
    ua: { code: 'URIC_ACID', value: bloodTest.ua || bloodTest.UA }
  };

  const result = {
    blood_test_id: `BT_${data.patientId || Date.now()}`,
    patient_id: data.patientId,
    data_source: 'CSV',
    verified: false,
    created_at: new Date(),
    updated_at: new Date()
  };

  // 各項目の値と判定を設定
  for (const [key, item] of Object.entries(items)) {
    const value = item.value;
    result[`${key}_value`] = value || '';

    if (value !== null && value !== undefined && value !== '') {
      result[`${key}_judgment`] = calculateJudgment(item.code, value, gender);
    } else {
      result[`${key}_judgment`] = '';
    }
  }

  return result;
}

/**
 * 性別を変換
 * @param {string} gender - 元の性別表記
 * @returns {string} M または F
 */
function convertGender(gender) {
  if (!gender) return 'M';

  const g = String(gender).toUpperCase().trim();

  if (g === 'M' || g === '男' || g === '1' || g === 'MALE') return 'M';
  if (g === 'F' || g === '女' || g === '2' || g === 'FEMALE') return 'F';

  return 'M';  // デフォルト
}

// ============================================
// 同期関数
// ============================================

/**
 * 既存データからAppSheetテーブルへ同期
 * @param {string} caseId - 案件ID
 * @param {Array<string>} patientIds - 患者IDリスト
 * @returns {Object} 同期結果
 */
function syncExistingToAppSheet(caseId, patientIds) {
  const ss = getAppSheetSpreadsheet();
  const results = {
    synced: 0,
    errors: []
  };

  for (const patientId of patientIds) {
    try {
      // 既存データを取得
      const existingData = getPatientFromExisting(patientId);

      if (!existingData) {
        results.errors.push(`${patientId}: データが見つかりません`);
        continue;
      }

      // AppSheet形式に変換
      const appSheetData = convertToAppSheetFormat(existingData);
      appSheetData.patient.case_id = caseId;

      // 各シートに保存
      saveToAppSheetTable(ss, APPSHEET_CONFIG.SHEETS.PATIENTS, appSheetData.patient, 'patient_id');
      saveToAppSheetTable(ss, APPSHEET_CONFIG.SHEETS.BLOOD_TESTS, appSheetData.blood_test, 'blood_test_id');
      saveToAppSheetTable(ss, APPSHEET_CONFIG.SHEETS.ULTRASOUND, appSheetData.ultrasound, 'ultrasound_id');

      results.synced++;

    } catch (error) {
      results.errors.push(`${patientId}: ${error.message}`);
    }
  }

  return results;
}

/**
 * 既存スプレッドシートから患者データを取得
 * @param {string} patientId - 患者ID
 * @returns {Object|null} 患者データ
 */
function getPatientFromExisting(patientId) {
  try {
    // 受診者マスタから取得
    const sheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    // patientIdで検索
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(patientId).trim()) {
        const row = data[i];
        const patient = {};

        headers.forEach((h, idx) => {
          patient[h] = row[idx];
        });

        // 血液検査データも取得
        patient.bloodTest = getBloodTestFromExisting(patientId);

        return patient;
      }
    }

    return null;

  } catch (error) {
    logError('getPatientFromExisting', error);
    return null;
  }
}

/**
 * 既存の血液検査データを取得
 * @param {string} patientId - 患者ID
 * @returns {Object|null} 血液検査データ
 */
function getBloodTestFromExisting(patientId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(patientId).trim()) {
        const row = data[i];
        const bloodTest = {};

        headers.forEach((h, idx) => {
          bloodTest[h] = row[idx];
        });

        return bloodTest;
      }
    }

    return null;

  } catch (error) {
    logError('getBloodTestFromExisting', error);
    return null;
  }
}

/**
 * AppSheetテーブルにデータを保存
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @param {Object} data - データ
 * @param {string} keyColumn - キーカラム名
 */
function saveToAppSheetTable(ss, sheetName, data, keyColumn) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シートが見つかりません: ${sheetName}`);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyIndex = headers.indexOf(keyColumn);

  if (keyIndex === -1) {
    throw new Error(`キーカラムが見つかりません: ${keyColumn}`);
  }

  // 既存行を検索
  const keyValue = data[keyColumn];
  let existingRow = null;

  if (sheet.getLastRow() > 1) {
    const keyValues = sheet.getRange(2, keyIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < keyValues.length; i++) {
      if (String(keyValues[i][0]).trim() === String(keyValue).trim()) {
        existingRow = i + 2;
        break;
      }
    }
  }

  // 行データを構築
  const rowData = headers.map(h => {
    if (data.hasOwnProperty(h)) {
      return data[h];
    }
    return '';
  });

  // 保存
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

// ============================================
// 超音波所見自動入力
// ============================================

/**
 * A判定時に「異常なし」を自動入力
 * @param {string} patientId - 患者ID
 * @param {string} findingType - 所見タイプ（abd/carotid/echo）
 * @param {string} judgment - 判定
 */
function autoFillNormalFinding(patientId, findingType, judgment) {
  if (judgment !== 'A') return;

  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.ULTRASOUND);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const patientIdIndex = headers.indexOf('patient_id');
  const findingsColumn = `${findingType}_findings`;
  const findingsIndex = headers.indexOf(findingsColumn);

  if (patientIdIndex === -1 || findingsIndex === -1) return;

  // 患者行を検索
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][patientIdIndex]).trim() === String(patientId).trim()) {
      const currentValue = data[i][findingsIndex];

      // 空欄の場合のみ自動入力
      if (!currentValue || currentValue === '') {
        sheet.getRange(i + 2, findingsIndex + 1).setValue('異常なし');
      }
      break;
    }
  }
}

// ============================================
// 案件管理
// ============================================

/**
 * 新規案件を作成
 * @param {Object} caseData - 案件データ
 * @returns {string} 案件ID
 */
function createNewCase(caseData) {
  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.CASES);

  const caseId = `CASE_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd')}_${Math.random().toString(36).substring(2, 6).toUpperCase()}`;

  const now = new Date();

  const newCase = {
    case_id: caseId,
    case_name: caseData.case_name || '',
    exam_type: caseData.exam_type || 'ROSAI_SECONDARY',
    client_name: caseData.client_name || '',
    exam_date: caseData.exam_date || '',
    csv_file_id: caseData.csv_file_id || '',
    status: '未着手',
    patient_count: 0,
    completed_count: 0,
    current_step: 'STEP_1',
    created_at: now,
    updated_at: now,
    notes: caseData.notes || ''
  };

  saveToAppSheetTable(ss, APPSHEET_CONFIG.SHEETS.CASES, newCase, 'case_id');

  return caseId;
}

/**
 * 案件のステータスを更新
 * @param {string} caseId - 案件ID
 * @param {string} status - 新しいステータス
 */
function updateCaseStatus(caseId, status) {
  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.CASES);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const caseIdIndex = headers.indexOf('case_id');
  const statusIndex = headers.indexOf('status');
  const updatedAtIndex = headers.indexOf('updated_at');

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][caseIdIndex]).trim() === String(caseId).trim()) {
      sheet.getRange(i + 2, statusIndex + 1).setValue(status);
      sheet.getRange(i + 2, updatedAtIndex + 1).setValue(new Date());
      break;
    }
  }
}

/**
 * 案件の受診者数を更新
 * @param {string} caseId - 案件ID
 */
function updateCasePatientCount(caseId) {
  const ss = getAppSheetSpreadsheet();

  // 受診者数をカウント
  const patientsSheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.PATIENTS);
  const patientsData = patientsSheet.getDataRange().getValues();
  const patientHeaders = patientsData[0];
  const caseIdColIndex = patientHeaders.indexOf('case_id');
  const statusColIndex = patientHeaders.indexOf('status');

  let patientCount = 0;
  let completedCount = 0;

  for (let i = 1; i < patientsData.length; i++) {
    if (String(patientsData[i][caseIdColIndex]).trim() === String(caseId).trim()) {
      patientCount++;
      if (patientsData[i][statusColIndex] === '完了') {
        completedCount++;
      }
    }
  }

  // 案件シートを更新
  const casesSheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.CASES);
  const casesHeaders = casesSheet.getRange(1, 1, 1, casesSheet.getLastColumn()).getValues()[0];
  const caseCaseIdIndex = casesHeaders.indexOf('case_id');
  const patientCountIndex = casesHeaders.indexOf('patient_count');
  const completedCountIndex = casesHeaders.indexOf('completed_count');

  const casesData = casesSheet.getRange(2, 1, casesSheet.getLastRow() - 1, casesSheet.getLastColumn()).getValues();

  for (let i = 0; i < casesData.length; i++) {
    if (String(casesData[i][caseCaseIdIndex]).trim() === String(caseId).trim()) {
      casesSheet.getRange(i + 2, patientCountIndex + 1).setValue(patientCount);
      casesSheet.getRange(i + 2, completedCountIndex + 1).setValue(completedCount);
      break;
    }
  }
}

// ============================================
// スクリーニング判定（労災二次対象者選別）
// ============================================

/**
 * 労災二次検診の実施基準
 * 一次検診結果から二次検診対象者かを判定
 *
 * 対象条件（いずれかに該当）:
 * 1. 脳血管疾患リスク: 血圧 + (血糖 or 脂質) 異常
 * 2. 心臓疾患リスク: 血圧 + (血糖 or 脂質) 異常
 */
const SCREENING_CRITERIA = {
  // 血圧異常
  bp: {
    sys: 140,  // 収縮期 ≥140
    dia: 90    // 拡張期 ≥90
  },
  // 血糖異常（いずれか）
  glucose: {
    fbs: 126,     // 空腹時血糖 ≥126
    hba1c: 6.5    // HbA1c ≥6.5
  },
  // 脂質異常（いずれか）
  lipid: {
    ldl: 140,     // LDL ≥140
    hdl: 40,      // HDL <40（逆向き）
    tg: 150       // TG ≥150
  },
  // 肥満
  obesity: {
    bmi: 25,      // BMI ≥25
    waist_m: 85,  // 腹囲(男) ≥85
    waist_f: 90   // 腹囲(女) ≥90
  }
};

/**
 * 一次検診結果からスクリーニング判定
 * @param {Object} primaryData - 一次検診データ
 * @returns {Object} 判定結果
 */
function calculateScreeningResult(primaryData) {
  const result = {
    isTarget: false,
    reason: [],
    details: {}
  };

  // 血圧チェック
  const bpAbnormal = (
    toNumber(primaryData.primary_bp_sys) >= SCREENING_CRITERIA.bp.sys ||
    toNumber(primaryData.primary_bp_dia) >= SCREENING_CRITERIA.bp.dia
  );
  result.details.bp = bpAbnormal;
  if (bpAbnormal) result.reason.push('血圧異常');

  // 血糖チェック
  const glucoseAbnormal = (
    toNumber(primaryData.primary_fbs) >= SCREENING_CRITERIA.glucose.fbs ||
    toNumber(primaryData.primary_hba1c) >= SCREENING_CRITERIA.glucose.hba1c
  );
  result.details.glucose = glucoseAbnormal;
  if (glucoseAbnormal) result.reason.push('血糖異常');

  // 脂質チェック
  const lipidAbnormal = (
    toNumber(primaryData.primary_ldl) >= SCREENING_CRITERIA.lipid.ldl ||
    toNumber(primaryData.primary_hdl) < SCREENING_CRITERIA.lipid.hdl ||
    toNumber(primaryData.primary_tg) >= SCREENING_CRITERIA.lipid.tg
  );
  result.details.lipid = lipidAbnormal;
  if (lipidAbnormal) result.reason.push('脂質異常');

  // 肥満チェック
  const gender = primaryData.gender || 'M';
  const waistThreshold = gender === 'F' ? SCREENING_CRITERIA.obesity.waist_f : SCREENING_CRITERIA.obesity.waist_m;
  const obesityAbnormal = (
    toNumber(primaryData.primary_bmi) >= SCREENING_CRITERIA.obesity.bmi ||
    toNumber(primaryData.primary_waist) >= waistThreshold
  );
  result.details.obesity = obesityAbnormal;
  if (obesityAbnormal) result.reason.push('肥満');

  // 総合判定: 血圧異常 + (血糖 or 脂質) で対象
  if (bpAbnormal && (glucoseAbnormal || lipidAbnormal)) {
    result.isTarget = true;
  }

  return result;
}

/**
 * 受診者のスクリーニング判定を更新
 * @param {string} patientId - 受診者ID
 * @returns {Object} 更新結果
 */
function updateScreeningResult(patientId) {
  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.PATIENTS);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();

  // patient_idで行を検索
  const patientIdIdx = headers.indexOf('patient_id');
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][patientIdIdx]).trim() === String(patientId).trim()) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, error: '受診者が見つかりません' };
  }

  // 一次検診データを取得
  const primaryData = {};
  const primaryFields = [
    'primary_hdl', 'primary_ldl', 'primary_tg',
    'primary_fbs', 'primary_hba1c',
    'primary_bp_sys', 'primary_bp_dia',
    'primary_bmi', 'primary_waist', 'gender'
  ];

  primaryFields.forEach(field => {
    const idx = headers.indexOf(field);
    if (idx >= 0) {
      primaryData[field] = data[rowIndex][idx];
    }
  });

  // スクリーニング判定
  const screening = calculateScreeningResult(primaryData);

  // 結果を更新
  const screeningResultIdx = headers.indexOf('screening_result');
  if (screeningResultIdx >= 0) {
    const resultText = screening.isTarget ? '対象' : '非対象';
    sheet.getRange(rowIndex + 1, screeningResultIdx + 1).setValue(resultText);
  }

  return {
    success: true,
    patientId: patientId,
    isTarget: screening.isTarget,
    reason: screening.reason,
    details: screening.details
  };
}

/**
 * 案件内の全受診者のスクリーニング判定を一括更新
 * @param {string} caseId - 案件ID
 */
function updateAllScreeningResults(caseId) {
  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.PATIENTS);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();

  const caseIdIdx = headers.indexOf('case_id');
  const patientIdIdx = headers.indexOf('patient_id');
  const results = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][caseIdIdx]).trim() === String(caseId).trim()) {
      const patientId = data[i][patientIdIdx];
      const result = updateScreeningResult(patientId);
      results.push(result);
    }
  }

  Logger.log(`スクリーニング更新: ${results.length}件`);
  return results;
}

// ============================================
// スケジュール表連携
// ============================================

/**
 * スケジュール表の設定
 * AS_設定シートの SCHEDULE_SPREADSHEET_ID に設定するか、ここで直接指定
 */
const SCHEDULE_CONFIG = {
  // スケジュール表のスプレッドシートID
  SPREADSHEET_ID: '1FGlVDK2PIUl4yIhKkR6R1389Uqm0rY8MgpyeDh2FMDA',

  // シート名（実際のシート名に合わせて変更）
  SHEET_NAME: 'スケジュール',

  // カラムマッピング（実際の列位置に合わせて変更）
  COLUMNS: {
    DATE: 'A',           // 検診日
    CLIENT_NAME: 'B',    // 事業所名
    PATIENT_COUNT: 'C',  // 予定人数
    STATUS: 'D',         // ステータス
    NOTES: 'E'           // 備考
  }
};

/**
 * スケジュール表から案件をインポート
 * @param {string} targetDate - 対象日付（YYYY-MM-DD形式、省略時は今日）
 * @returns {Object} インポート結果
 */
function importFromSchedule(targetDate = null) {
  try {
    const scheduleId = getAppSheetSetting('SCHEDULE_SPREADSHEET_ID') || SCHEDULE_CONFIG.SPREADSHEET_ID;
    const scheduleSs = SpreadsheetApp.openById(scheduleId);
    const scheduleSheet = scheduleSs.getSheetByName(SCHEDULE_CONFIG.SHEET_NAME);

    if (!scheduleSheet) {
      return { success: false, error: 'スケジュールシートが見つかりません' };
    }

    const data = scheduleSheet.getDataRange().getValues();
    const headers = data[0];
    const imported = [];

    // 対象日付の設定
    const targetDateObj = targetDate ? new Date(targetDate) : new Date();
    const targetDateStr = Utilities.formatDate(targetDateObj, 'Asia/Tokyo', 'yyyy-MM-dd');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = row[0];  // A列=日付

      // 日付フィルター
      if (rowDate) {
        const rowDateStr = Utilities.formatDate(new Date(rowDate), 'Asia/Tokyo', 'yyyy-MM-dd');

        if (!targetDate || rowDateStr === targetDateStr) {
          // 案件データを構築
          const caseData = {
            case_name: `${row[1] || '案件'}_${rowDateStr}`,
            exam_type: 'ROSAI_SECONDARY',
            client_name: row[1] || '',
            exam_date: rowDate,
            notes: `スケジュール表Row${i + 1}からインポート`,
            schedule_row: i + 1  // 元の行番号を記録
          };

          // 既存チェック（同じ日付・事業所名の案件があるかどうか）
          if (!caseExistsByDateAndClient(rowDateStr, caseData.client_name)) {
            const caseId = createNewCase(caseData);
            imported.push({
              caseId: caseId,
              clientName: caseData.client_name,
              examDate: rowDateStr
            });
          }
        }
      }
    }

    return {
      success: true,
      message: `${imported.length}件の案件をインポートしました`,
      imported: imported
    };

  } catch (error) {
    logError('importFromSchedule', error);
    return { success: false, error: error.message };
  }
}

/**
 * 同じ日付・事業所名の案件が存在するかチェック
 */
function caseExistsByDateAndClient(dateStr, clientName) {
  const ss = getAppSheetSpreadsheet();
  const sheet = ss.getSheetByName(APPSHEET_CONFIG.SHEETS.CASES);

  if (!sheet || sheet.getLastRow() < 2) return false;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateIndex = headers.indexOf('exam_date');
  const clientIndex = headers.indexOf('client_name');

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][dateIndex];
    const rowClient = data[i][clientIndex];

    if (rowDate && rowClient) {
      const rowDateStr = Utilities.formatDate(new Date(rowDate), 'Asia/Tokyo', 'yyyy-MM-dd');
      if (rowDateStr === dateStr && rowClient === clientName) {
        return true;
      }
    }
  }

  return false;
}

/**
 * スケジュール表のステータスを更新
 * @param {number} rowNum - スケジュール表の行番号
 * @param {string} status - 新しいステータス
 */
function updateScheduleStatus(rowNum, status) {
  try {
    const scheduleId = getAppSheetSetting('SCHEDULE_SPREADSHEET_ID') || SCHEDULE_CONFIG.SPREADSHEET_ID;
    const scheduleSs = SpreadsheetApp.openById(scheduleId);
    const scheduleSheet = scheduleSs.getSheetByName(SCHEDULE_CONFIG.SHEET_NAME);

    if (scheduleSheet) {
      const statusCol = SCHEDULE_CONFIG.COLUMNS.STATUS.charCodeAt(0) - 64;
      scheduleSheet.getRange(rowNum, statusCol).setValue(status);
    }

  } catch (error) {
    logError('updateScheduleStatus', error);
  }
}

/**
 * 今日のスケジュールをインポート（メニュー用）
 */
function importTodaySchedule() {
  const result = importFromSchedule();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * 指定日のスケジュールをインポート（テスト用）
 */
function testImportSchedule() {
  // テスト用の日付を指定
  const result = importFromSchedule('2025-12-10');
  Logger.log(JSON.stringify(result, null, 2));
}

// ============================================
// ユーティリティ関数（外部依存のフォールバック）
// ============================================

/**
 * 数値変換（toNumberが未定義の場合のフォールバック）
 */
if (typeof toNumber !== 'function') {
  function toNumber(value) {
    if (value === null || value === undefined || value === '') return 0;
    const num = parseFloat(String(value).replace(/,/g, ''));
    return isNaN(num) ? 0 : num;
  }
}

/**
 * ログ出力（logInfoが未定義の場合のフォールバック）
 */
if (typeof logInfo !== 'function') {
  function logInfo(message) {
    Logger.log(`[INFO] ${message}`);
  }
}

/**
 * エラーログ（logErrorが未定義の場合のフォールバック）
 */
if (typeof logError !== 'function') {
  function logError(functionName, error) {
    Logger.log(`[ERROR] ${functionName}: ${error.message}`);
    if (error.stack) {
      Logger.log(error.stack);
    }
  }
}

/**
 * シート取得（getSheetが未定義の場合のフォールバック）
 */
if (typeof getSheet !== 'function') {
  function getSheet(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss ? ss.getSheetByName(sheetName) : null;
  }
}

/**
 * 汎用行検索ユーティリティ
 * @param {Sheet} sheet - シート
 * @param {string} columnName - 検索するカラム名
 * @param {*} searchValue - 検索値
 * @returns {number} 行番号（見つからない場合は-1）
 */
function findRowByColumnValue(sheet, columnName, searchValue) {
  if (!sheet || sheet.getLastRow() < 2) return -1;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(columnName);

  if (colIndex === -1) return -1;

  const data = sheet.getRange(2, colIndex + 1, sheet.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(searchValue).trim()) {
      return i + 2;  // 1-indexed + header offset
    }
  }

  return -1;
}

/**
 * シートからオブジェクト配列を取得
 * @param {string} sheetName - シート名
 * @param {Object} options - オプション { filter: (row) => boolean }
 * @returns {Array<Object>} データ配列
 */
function getSheetDataAsObjects(sheetName, options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  let results = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  if (options.filter) {
    results = results.filter(options.filter);
  }

  return results;
}
