/**
 * 健診結果DB 統合システム - CRUD操作
 *
 * @description 設計書: 健診結果DB_設計書_v1.md Phase 1 基盤構築
 * @version 1.0.0
 * @date 2025-12-14
 */

// ============================================
// ID採番関数
// ============================================

/**
 * 受診者ID（P00001形式）を生成
 * @returns {string} 新しい受診者ID
 */
function generatePatientId() {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 'P00001';
  }

  // 最後のIDを取得
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (const row of ids) {
    const id = row[0];
    if (id && id.startsWith('P')) {
      const num = parseInt(id.substring(1), 10);
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }

  return 'P' + String(maxNum + 1).padStart(5, '0');
}

/**
 * 受診ID（YYYYMMDD-NNN形式）を生成
 * @param {Date} visitDate - 受診日
 * @returns {string} 新しい受診ID
 */
function generateVisitId(visitDate) {
  const dateStr = formatDateYYYYMMDD(visitDate);
  const sheet = getSheet(DB_CONFIG.SHEETS.VISIT_RECORD);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return `${dateStr}-001`;
  }

  // 同日の最大番号を取得
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (const row of ids) {
    const id = row[0];
    if (id && id.startsWith(dateStr)) {
      const parts = id.split('-');
      if (parts.length === 2) {
        const num = parseInt(parts[1], 10);
        if (num > maxNum) {
          maxNum = num;
        }
      }
    }
  }

  return `${dateStr}-${String(maxNum + 1).padStart(3, '0')}`;
}

/**
 * 結果ID（R00001形式）を生成
 * @returns {string} 新しい結果ID
 */
function generateResultId() {
  const sheet = getSheet(DB_CONFIG.SHEETS.TEST_RESULT);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 'R00001';
  }

  // 最後のIDを取得
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (const row of ids) {
    const id = row[0];
    if (id && id.startsWith('R')) {
      const num = parseInt(id.substring(1), 10);
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }

  return 'R' + String(maxNum + 1).padStart(5, '0');
}

/**
 * 指導ID（G00001形式）を生成
 * @returns {string} 新しい指導ID
 */
function generateGuidanceId() {
  const sheet = getSheet(DB_CONFIG.SHEETS.GUIDANCE_RECORD);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 'G00001';
  }

  // 最後のIDを取得
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (const row of ids) {
    const id = row[0];
    if (id && id.startsWith('G')) {
      const num = parseInt(id.substring(1), 10);
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }

  return 'G' + String(maxNum + 1).padStart(5, '0');
}

// ============================================
// 受診者マスタ CRUD
// ============================================

/**
 * 受診者を作成
 * @param {Object} data - 受診者データ
 * @returns {string} 作成された受診者ID
 */
function createPatient(data) {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const cols = COLUMN_DEFINITIONS.PATIENT_MASTER.columns;

  const patientId = generatePatientId();
  const now = new Date();

  const row = new Array(COLUMN_DEFINITIONS.PATIENT_MASTER.headers.length).fill('');
  row[cols.PATIENT_ID] = patientId;
  row[cols.NAME] = data.name || '';
  row[cols.KANA] = data.kana || data.nameKana || '';
  row[cols.BIRTHDATE] = data.birthdate || data.birthDate || '';
  row[cols.GENDER] = data.gender || '';
  row[cols.POSTAL_CODE] = data.postalCode || '';
  row[cols.ADDRESS] = data.address || '';
  row[cols.PHONE] = data.phone || '';
  row[cols.EMAIL] = data.email || '';
  row[cols.COMPANY] = data.company || '';
  row[cols.NOTES] = data.notes || '';
  row[cols.CREATED_AT] = now;
  row[cols.UPDATED_AT] = now;

  sheet.appendRow(row);
  logInfo(`受診者作成: ${patientId} (${data.name})`);

  return patientId;
}

/**
 * 受診者を取得（ID指定）
 * @param {string} patientId - 受診者ID
 * @returns {Object|null} 受診者データ
 */
function getPatientById(patientId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.PATIENT_MASTER.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.PATIENT_MASTER.columns;

  for (const row of data) {
    if (row[cols.PATIENT_ID] === patientId) {
      return {
        patientId: row[cols.PATIENT_ID],
        name: row[cols.NAME],
        kana: row[cols.KANA],
        birthdate: row[cols.BIRTHDATE],
        gender: row[cols.GENDER],
        postalCode: row[cols.POSTAL_CODE],
        address: row[cols.ADDRESS],
        phone: row[cols.PHONE],
        email: row[cols.EMAIL],
        company: row[cols.COMPANY],
        notes: row[cols.NOTES],
        createdAt: row[cols.CREATED_AT],
        updatedAt: row[cols.UPDATED_AT]
      };
    }
  }

  return null;
}

/**
 * 受診者を検索
 * @param {Object} criteria - 検索条件 {name, kana, company}
 * @returns {Array<Object>} 受診者リスト
 */
function searchPatients(criteria) {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.PATIENT_MASTER.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.PATIENT_MASTER.columns;
  const results = [];

  for (const row of data) {
    let match = true;

    if (criteria.name && !row[cols.NAME].includes(criteria.name)) {
      match = false;
    }
    if (criteria.kana && !row[cols.KANA].includes(criteria.kana)) {
      match = false;
    }
    if (criteria.company && !row[cols.COMPANY].includes(criteria.company)) {
      match = false;
    }

    if (match) {
      results.push({
        patientId: row[cols.PATIENT_ID],
        name: row[cols.NAME],
        kana: row[cols.KANA],
        birthdate: row[cols.BIRTHDATE],
        gender: row[cols.GENDER],
        company: row[cols.COMPANY]
      });
    }
  }

  return results;
}

/**
 * 受診者を更新
 * @param {string} patientId - 受診者ID
 * @param {Object} data - 更新データ
 * @returns {boolean} 成功/失敗
 */
function updatePatient(patientId, data) {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return false;

  const allData = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.PATIENT_MASTER.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.PATIENT_MASTER.columns;

  for (let i = 0; i < allData.length; i++) {
    if (allData[i][cols.PATIENT_ID] === patientId) {
      const rowNum = i + 2;

      if (data.name !== undefined) sheet.getRange(rowNum, cols.NAME + 1).setValue(data.name);
      if (data.kana !== undefined) sheet.getRange(rowNum, cols.KANA + 1).setValue(data.kana);
      if (data.birthdate !== undefined) sheet.getRange(rowNum, cols.BIRTHDATE + 1).setValue(data.birthdate);
      if (data.gender !== undefined) sheet.getRange(rowNum, cols.GENDER + 1).setValue(data.gender);
      if (data.postalCode !== undefined) sheet.getRange(rowNum, cols.POSTAL_CODE + 1).setValue(data.postalCode);
      if (data.address !== undefined) sheet.getRange(rowNum, cols.ADDRESS + 1).setValue(data.address);
      if (data.phone !== undefined) sheet.getRange(rowNum, cols.PHONE + 1).setValue(data.phone);
      if (data.email !== undefined) sheet.getRange(rowNum, cols.EMAIL + 1).setValue(data.email);
      if (data.company !== undefined) sheet.getRange(rowNum, cols.COMPANY + 1).setValue(data.company);
      if (data.notes !== undefined) sheet.getRange(rowNum, cols.NOTES + 1).setValue(data.notes);

      sheet.getRange(rowNum, cols.UPDATED_AT + 1).setValue(new Date());

      logInfo(`受診者更新: ${patientId}`);
      return true;
    }
  }

  return false;
}

/**
 * 受診者を削除
 * @param {string} patientId - 受診者ID
 * @returns {boolean} 成功/失敗
 */
function deletePatient(patientId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return false;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === patientId) {
      sheet.deleteRow(i + 2);
      logInfo(`受診者削除: ${patientId}`);
      return true;
    }
  }

  return false;
}

// ============================================
// 受診記録 CRUD
// ============================================

/**
 * 受診記録を作成
 * @param {Object} data - 受診記録データ
 * @returns {string} 作成された受診ID
 */
function createVisitRecord(data) {
  const sheet = getSheet(DB_CONFIG.SHEETS.VISIT_RECORD);
  const cols = COLUMN_DEFINITIONS.VISIT_RECORD.columns;

  const visitId = generateVisitId(data.visitDate || new Date());
  const now = new Date();

  // 年齢を計算
  let age = '';
  if (data.patientId) {
    const patient = getPatientById(data.patientId);
    if (patient && patient.birthdate) {
      age = calculateAge(patient.birthdate, data.visitDate || new Date());
    }
  }

  const row = new Array(COLUMN_DEFINITIONS.VISIT_RECORD.headers.length).fill('');
  row[cols.VISIT_ID] = visitId;
  row[cols.PATIENT_ID] = data.patientId || '';
  row[cols.EXAM_TYPE_ID] = data.examTypeId || '';
  row[cols.COURSE_ID] = data.courseId || '';
  row[cols.VISIT_DATE] = data.visitDate || '';
  row[cols.AGE] = age;
  row[cols.OVERALL_JUDGMENT] = data.overallJudgment || '';
  row[cols.DOCTOR_NOTES] = data.doctorNotes || '';
  row[cols.STATUS] = data.status || DB_CONFIG.STATUS.INPUT;
  row[cols.CREATED_AT] = now;
  row[cols.UPDATED_AT] = now;

  sheet.appendRow(row);
  logInfo(`受診記録作成: ${visitId}`);

  return visitId;
}

/**
 * 受診記録を取得（ID指定）
 * @param {string} visitId - 受診ID
 * @returns {Object|null} 受診記録データ
 */
function getVisitRecordById(visitId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.VISIT_RECORD);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.VISIT_RECORD.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.VISIT_RECORD.columns;

  for (const row of data) {
    if (row[cols.VISIT_ID] === visitId) {
      return {
        visitId: row[cols.VISIT_ID],
        patientId: row[cols.PATIENT_ID],
        examTypeId: row[cols.EXAM_TYPE_ID],
        courseId: row[cols.COURSE_ID],
        visitDate: row[cols.VISIT_DATE],
        age: row[cols.AGE],
        overallJudgment: row[cols.OVERALL_JUDGMENT],
        doctorNotes: row[cols.DOCTOR_NOTES],
        status: row[cols.STATUS],
        createdAt: row[cols.CREATED_AT],
        updatedAt: row[cols.UPDATED_AT]
      };
    }
  }

  return null;
}

/**
 * 受診者の受診履歴を取得
 * @param {string} patientId - 受診者ID
 * @returns {Array<Object>} 受診記録リスト
 */
function getVisitRecordsByPatientId(patientId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.VISIT_RECORD);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.VISIT_RECORD.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.VISIT_RECORD.columns;
  const results = [];

  for (const row of data) {
    if (row[cols.PATIENT_ID] === patientId) {
      results.push({
        visitId: row[cols.VISIT_ID],
        patientId: row[cols.PATIENT_ID],
        examTypeId: row[cols.EXAM_TYPE_ID],
        courseId: row[cols.COURSE_ID],
        visitDate: row[cols.VISIT_DATE],
        age: row[cols.AGE],
        overallJudgment: row[cols.OVERALL_JUDGMENT],
        status: row[cols.STATUS]
      });
    }
  }

  // 受診日で降順ソート
  results.sort((a, b) => new Date(b.visitDate) - new Date(a.visitDate));

  return results;
}

/**
 * 受診記録を更新
 * @param {string} visitId - 受診ID
 * @param {Object} data - 更新データ
 * @returns {boolean} 成功/失敗
 */
function updateVisitRecord(visitId, data) {
  const sheet = getSheet(DB_CONFIG.SHEETS.VISIT_RECORD);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return false;

  const allData = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.VISIT_RECORD.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.VISIT_RECORD.columns;

  for (let i = 0; i < allData.length; i++) {
    if (allData[i][cols.VISIT_ID] === visitId) {
      const rowNum = i + 2;

      if (data.examTypeId !== undefined) sheet.getRange(rowNum, cols.EXAM_TYPE_ID + 1).setValue(data.examTypeId);
      if (data.courseId !== undefined) sheet.getRange(rowNum, cols.COURSE_ID + 1).setValue(data.courseId);
      if (data.overallJudgment !== undefined) sheet.getRange(rowNum, cols.OVERALL_JUDGMENT + 1).setValue(data.overallJudgment);
      if (data.doctorNotes !== undefined) sheet.getRange(rowNum, cols.DOCTOR_NOTES + 1).setValue(data.doctorNotes);
      if (data.status !== undefined) sheet.getRange(rowNum, cols.STATUS + 1).setValue(data.status);

      sheet.getRange(rowNum, cols.UPDATED_AT + 1).setValue(new Date());

      logInfo(`受診記録更新: ${visitId}`);
      return true;
    }
  }

  return false;
}

// ============================================
// 検査結果 CRUD
// ============================================

/**
 * 検査結果を作成
 * @param {Object} data - 検査結果データ
 * @returns {string} 作成された結果ID
 */
function createTestResult(data) {
  const sheet = getSheet(DB_CONFIG.SHEETS.TEST_RESULT);
  const cols = COLUMN_DEFINITIONS.TEST_RESULT.columns;

  const resultId = generateResultId();
  const now = new Date();

  // 数値変換
  let numericValue = '';
  if (data.value && !isNaN(parseFloat(data.value))) {
    numericValue = parseFloat(data.value);
  }

  const row = new Array(COLUMN_DEFINITIONS.TEST_RESULT.headers.length).fill('');
  row[cols.RESULT_ID] = resultId;
  row[cols.VISIT_ID] = data.visitId || '';
  row[cols.ITEM_ID] = data.itemId || '';
  row[cols.VALUE] = data.value || '';
  row[cols.NUMERIC_VALUE] = numericValue;
  row[cols.JUDGMENT] = data.judgment || '';
  row[cols.NOTES] = data.notes || '';
  row[cols.CREATED_AT] = now;

  sheet.appendRow(row);

  return resultId;
}

/**
 * 受診の検査結果を取得
 * @param {string} visitId - 受診ID
 * @returns {Array<Object>} 検査結果リスト
 */
function getTestResultsByVisitId(visitId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.TEST_RESULT);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.TEST_RESULT.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.TEST_RESULT.columns;
  const results = [];

  for (const row of data) {
    if (row[cols.VISIT_ID] === visitId) {
      results.push({
        resultId: row[cols.RESULT_ID],
        visitId: row[cols.VISIT_ID],
        itemId: row[cols.ITEM_ID],
        value: row[cols.VALUE],
        numericValue: row[cols.NUMERIC_VALUE],
        judgment: row[cols.JUDGMENT],
        notes: row[cols.NOTES],
        createdAt: row[cols.CREATED_AT]
      });
    }
  }

  return results;
}

/**
 * 検査結果を更新
 * @param {string} resultId - 結果ID
 * @param {Object} data - 更新データ
 * @returns {boolean} 成功/失敗
 */
function updateTestResult(resultId, data) {
  const sheet = getSheet(DB_CONFIG.SHEETS.TEST_RESULT);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return false;

  const allData = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.TEST_RESULT.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.TEST_RESULT.columns;

  for (let i = 0; i < allData.length; i++) {
    if (allData[i][cols.RESULT_ID] === resultId) {
      const rowNum = i + 2;

      if (data.value !== undefined) {
        sheet.getRange(rowNum, cols.VALUE + 1).setValue(data.value);
        // 数値変換も更新
        const numericValue = !isNaN(parseFloat(data.value)) ? parseFloat(data.value) : '';
        sheet.getRange(rowNum, cols.NUMERIC_VALUE + 1).setValue(numericValue);
      }
      if (data.judgment !== undefined) sheet.getRange(rowNum, cols.JUDGMENT + 1).setValue(data.judgment);
      if (data.notes !== undefined) sheet.getRange(rowNum, cols.NOTES + 1).setValue(data.notes);

      return true;
    }
  }

  return false;
}

/**
 * 受診の全検査結果を削除
 * @param {string} visitId - 受診ID
 * @returns {number} 削除件数
 */
function deleteTestResultsByVisitId(visitId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.TEST_RESULT);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return 0;

  const ids = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let deleteCount = 0;

  // 後ろから削除（インデックスずれ防止）
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][1] === visitId) {
      sheet.deleteRow(i + 2);
      deleteCount++;
    }
  }

  if (deleteCount > 0) {
    logInfo(`検査結果削除: ${visitId} (${deleteCount}件)`);
  }

  return deleteCount;
}

// ============================================
// マスタデータ取得関数
// ============================================

/**
 * 項目マスタを取得
 * @param {string} itemId - 項目ID（省略で全件）
 * @returns {Object|Array<Object>} 項目データ
 */
function getItemMaster(itemId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.ITEM_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return itemId ? null : [];

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.ITEM_MASTER.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.ITEM_MASTER.columns;

  const mapRow = (row) => ({
    itemId: row[cols.ITEM_ID],
    itemName: row[cols.ITEM_NAME],
    category: row[cols.CATEGORY],
    unit: row[cols.UNIT],
    dataType: row[cols.DATA_TYPE],
    genderDiff: row[cols.GENDER_DIFF],
    judgmentMethod: row[cols.JUDGMENT_METHOD],
    aMin: row[cols.A_MIN],
    aMax: row[cols.A_MAX],
    bMin: row[cols.B_MIN],
    bMax: row[cols.B_MAX],
    cMin: row[cols.C_MIN],
    cMax: row[cols.C_MAX],
    dCondition: row[cols.D_CONDITION],
    aMinF: row[cols.A_MIN_F],
    aMaxF: row[cols.A_MAX_F],
    displayOrder: row[cols.DISPLAY_ORDER],
    isActive: row[cols.IS_ACTIVE]
  });

  if (itemId) {
    for (const row of data) {
      if (row[cols.ITEM_ID] === itemId && row[cols.IS_ACTIVE]) {
        return mapRow(row);
      }
    }
    return null;
  }

  return data.filter(row => row[cols.IS_ACTIVE]).map(mapRow);
}

/**
 * 検診種別マスタを取得
 * @returns {Array<Object>} 検診種別リスト
 */
function getExamTypeMaster() {
  const sheet = getSheet(DB_CONFIG.SHEETS.EXAM_TYPE_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.EXAM_TYPE_MASTER.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.EXAM_TYPE_MASTER.columns;

  return data
    .filter(row => row[cols.IS_ACTIVE])
    .map(row => ({
      examTypeId: row[cols.TYPE_ID],
      name: row[cols.TYPE_NAME],
      courseRequired: row[cols.COURSE_REQUIRED],
      displayOrder: row[cols.DISPLAY_ORDER]
    }))
    .sort((a, b) => a.displayOrder - b.displayOrder);
}

/**
 * コースマスタを取得
 * @returns {Array<Object>} コースリスト
 */
function getCourseMaster() {
  const sheet = getSheet(DB_CONFIG.SHEETS.COURSE_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.COURSE_MASTER.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.COURSE_MASTER.columns;

  return data
    .filter(row => row[cols.IS_ACTIVE])
    .map(row => ({
      courseId: row[cols.COURSE_ID],
      name: row[cols.COURSE_NAME],
      price: row[cols.PRICE],
      testItems: row[cols.TEST_ITEMS],
      displayOrder: row[cols.DISPLAY_ORDER]
    }))
    .sort((a, b) => a.displayOrder - b.displayOrder);
}

// ============================================
// デバッグ用テスト関数
// ============================================

/**
 * 検診種別マスタのデバッグ
 * GASエディタで実行してログを確認
 */
function testGetExamTypeMaster() {
  const sheetName = DB_CONFIG.SHEETS.EXAM_TYPE_MASTER;
  console.log('シート名:', sheetName);

  const sheet = getSheet(sheetName);
  if (!sheet) {
    console.log('エラー: シートが見つかりません');
    return;
  }

  const lastRow = sheet.getLastRow();
  console.log('最終行:', lastRow);

  if (lastRow <= 1) {
    console.log('データなし（ヘッダーのみ）');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  console.log('生データ:', JSON.stringify(data));

  const result = getExamTypeMaster();
  console.log('getExamTypeMaster結果:', JSON.stringify(result));
}
