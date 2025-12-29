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
 * 受診者を作成（17列構造対応 - カルテNo追加版）
 * @param {Object} data - 受診者データ
 * @returns {Object} 作成結果 {success, patientId, error}
 */
function createPatient(data) {
  try {
    const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
    const cols = COLUMN_DEFINITIONS.PATIENT_MASTER.columns;

    const patientId = generatePatientId();
    const now = new Date();

    // 17列構造: 受診者ID, カルテNo, ステータス, 受診日, 氏名, カナ, 性別, 生年月日, 年齢,
    //           受診コース, 事業所名, 所属, 総合判定, CSV取込日時, 最終更新日時, 出力日時, BML患者ID
    const row = new Array(COLUMN_DEFINITIONS.PATIENT_MASTER.headers.length).fill('');
    row[cols.PATIENT_ID] = patientId;
    row[cols.KARTE_NO] = data.karteNo || '';  // ★カルテNo（クリニック患者ID）
    row[cols.STATUS] = data.status || '入力中';
    row[cols.VISIT_DATE] = data.visitDate || data.examDate || '';
    row[cols.NAME] = data.name || '';
    row[cols.KANA] = data.kana || data.nameKana || '';
    row[cols.GENDER] = data.gender || '';
    row[cols.BIRTHDATE] = data.birthdate || data.birthDate || '';
    row[cols.AGE] = data.age || '';
    row[cols.COURSE] = data.course || data.courseId || '';
    row[cols.COMPANY] = data.company || data.companyName || '';
    row[cols.DEPARTMENT] = data.department || '';
    row[cols.OVERALL_JUDGMENT] = data.overallJudgment || '';
    row[cols.CSV_IMPORT_DATE] = data.csvImportDate || '';
    row[cols.UPDATED_AT] = now;
    row[cols.EXPORT_DATE] = '';
    row[cols.BML_PATIENT_ID] = data.bmlPatientId || '';

    sheet.appendRow(row);
    logInfo(`受診者作成: ${patientId} カルテNo:${data.karteNo || '-'} (${data.name})`);

    return { success: true, patientId: patientId };
  } catch (e) {
    logError('createPatient', e);
    return { success: false, error: e.message };
  }
}

/**
 * 受診者を取得（ID指定）（17列構造対応 - カルテNo追加版）
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
        karteNo: row[cols.KARTE_NO],  // ★カルテNo（クリニック患者ID）
        status: row[cols.STATUS],
        visitDate: row[cols.VISIT_DATE],
        name: row[cols.NAME],
        kana: row[cols.KANA],
        gender: row[cols.GENDER],
        birthdate: row[cols.BIRTHDATE],
        age: row[cols.AGE],
        course: row[cols.COURSE],
        company: row[cols.COMPANY],
        department: row[cols.DEPARTMENT],
        overallJudgment: row[cols.OVERALL_JUDGMENT],
        csvImportDate: row[cols.CSV_IMPORT_DATE],
        updatedAt: row[cols.UPDATED_AT],
        exportDate: row[cols.EXPORT_DATE],
        bmlPatientId: row[cols.BML_PATIENT_ID]
      };
    }
  }

  return null;
}

/**
 * 受診者をカルテNoで検索（CSV取込用）（17列構造対応）
 * @param {string} karteNo - カルテNo（6桁のクリニック患者ID）
 * @returns {Object|null} 受診者データ
 */
function getPatientByKarteNo(karteNo) {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMN_DEFINITIONS.PATIENT_MASTER.headers.length).getValues();
  const cols = COLUMN_DEFINITIONS.PATIENT_MASTER.columns;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (String(row[cols.KARTE_NO]) === String(karteNo)) {
      return {
        rowIndex: i + 2,  // シート上の行番号（更新用）
        patientId: row[cols.PATIENT_ID],
        karteNo: row[cols.KARTE_NO],
        status: row[cols.STATUS],
        visitDate: row[cols.VISIT_DATE],
        name: row[cols.NAME],
        kana: row[cols.KANA],
        gender: row[cols.GENDER],
        birthdate: row[cols.BIRTHDATE],
        age: row[cols.AGE],
        course: row[cols.COURSE],
        company: row[cols.COMPANY],
        department: row[cols.DEPARTMENT],
        overallJudgment: row[cols.OVERALL_JUDGMENT],
        csvImportDate: row[cols.CSV_IMPORT_DATE],
        updatedAt: row[cols.UPDATED_AT],
        exportDate: row[cols.EXPORT_DATE],
        bmlPatientId: row[cols.BML_PATIENT_ID]
      };
    }
  }

  return null;
}

/**
 * 受診者を検索（17列構造対応 - カルテNo追加版）
 * @param {Object} criteria - 検索条件 {name, kana, company, status, karteNo}
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

    if (criteria.karteNo && String(row[cols.KARTE_NO]) !== String(criteria.karteNo)) {
      match = false;
    }
    if (criteria.name && !String(row[cols.NAME]).includes(criteria.name)) {
      match = false;
    }
    if (criteria.kana && !String(row[cols.KANA]).includes(criteria.kana)) {
      match = false;
    }
    if (criteria.company && !String(row[cols.COMPANY]).includes(criteria.company)) {
      match = false;
    }
    if (criteria.status && row[cols.STATUS] !== criteria.status) {
      match = false;
    }

    if (match) {
      results.push({
        patientId: row[cols.PATIENT_ID],
        karteNo: row[cols.KARTE_NO],  // ★カルテNo（クリニック患者ID）
        status: row[cols.STATUS],
        visitDate: row[cols.VISIT_DATE],
        name: row[cols.NAME],
        kana: row[cols.KANA],
        gender: row[cols.GENDER],
        birthdate: row[cols.BIRTHDATE],
        age: row[cols.AGE],
        course: row[cols.COURSE],
        company: row[cols.COMPANY],
        department: row[cols.DEPARTMENT],
        overallJudgment: row[cols.OVERALL_JUDGMENT]
      });
    }
  }

  return results;
}

/**
 * 受診者を更新（17列構造対応 - カルテNo追加版）
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

      // 17列構造の各フィールドを更新
      if (data.karteNo !== undefined) sheet.getRange(rowNum, cols.KARTE_NO + 1).setValue(data.karteNo);  // ★カルテNo
      if (data.status !== undefined) sheet.getRange(rowNum, cols.STATUS + 1).setValue(data.status);
      if (data.visitDate !== undefined) sheet.getRange(rowNum, cols.VISIT_DATE + 1).setValue(data.visitDate);
      if (data.name !== undefined) sheet.getRange(rowNum, cols.NAME + 1).setValue(data.name);
      if (data.kana !== undefined) sheet.getRange(rowNum, cols.KANA + 1).setValue(data.kana);
      if (data.gender !== undefined) sheet.getRange(rowNum, cols.GENDER + 1).setValue(data.gender);
      if (data.birthdate !== undefined) sheet.getRange(rowNum, cols.BIRTHDATE + 1).setValue(data.birthdate);
      if (data.age !== undefined) sheet.getRange(rowNum, cols.AGE + 1).setValue(data.age);
      if (data.course !== undefined) sheet.getRange(rowNum, cols.COURSE + 1).setValue(data.course);
      if (data.company !== undefined) sheet.getRange(rowNum, cols.COMPANY + 1).setValue(data.company);
      if (data.department !== undefined) sheet.getRange(rowNum, cols.DEPARTMENT + 1).setValue(data.department);
      if (data.overallJudgment !== undefined) sheet.getRange(rowNum, cols.OVERALL_JUDGMENT + 1).setValue(data.overallJudgment);
      if (data.csvImportDate !== undefined) sheet.getRange(rowNum, cols.CSV_IMPORT_DATE + 1).setValue(data.csvImportDate);
      if (data.exportDate !== undefined) sheet.getRange(rowNum, cols.EXPORT_DATE + 1).setValue(data.exportDate);
      if (data.bmlPatientId !== undefined) sheet.getRange(rowNum, cols.BML_PATIENT_ID + 1).setValue(data.bmlPatientId);

      // 更新日時を自動設定
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
 * @deprecated 項目マスタは EXAM_ITEM_MASTER + JUDGMENT_CRITERIA に統一済み
 * MasterData.js の EXAM_ITEM_MASTER_DATA を使用してください。
 * または getExamItemById() / getJudgmentCriteria() を使用してください。
 *
 * @param {string} itemId - 項目ID（省略で全件）
 * @returns {Object|Array<Object>} 項目データ
 */
function getItemMaster(itemId) {
  // EXAM_ITEM_MASTER_DATA から取得（後方互換性）
  if (typeof EXAM_ITEM_MASTER_DATA === 'undefined') {
    logError('getItemMaster', new Error('EXAM_ITEM_MASTER_DATA is not defined'));
    return itemId ? null : [];
  }

  if (itemId) {
    const item = EXAM_ITEM_MASTER_DATA.find(i => i.item_id === itemId);
    if (!item) return null;
    // 旧フォーマットに変換（後方互換性）
    return {
      itemId: item.item_id,
      itemName: item.item_name,
      category: item.category,
      unit: item.unit || '',
      dataType: item.data_type || '数値',
      displayOrder: item.display_order || 0,
      isActive: true
    };
  }

  return EXAM_ITEM_MASTER_DATA.map(item => ({
    itemId: item.item_id,
    itemName: item.item_name,
    category: item.category,
    unit: item.unit || '',
    dataType: item.data_type || '数値',
    displayOrder: item.display_order || 0,
    isActive: true
  }));
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

/**
 * 登録処理のデバッグ関数
 */
function testRegistration() {
  console.log('=== 登録テスト開始 ===');

  // 1. シート確認
  const patientSheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const visitSheet = getSheet(DB_CONFIG.SHEETS.VISIT_RECORD);

  console.log('受診者マスタシート:', patientSheet ? '存在' : '未作成');
  console.log('受診記録シート:', visitSheet ? '存在' : '未作成');

  if (patientSheet) {
    const headers = patientSheet.getRange(1, 1, 1, 13).getValues()[0];
    console.log('受診者マスタヘッダー:', JSON.stringify(headers));
    console.log('期待するヘッダー:', JSON.stringify(COLUMN_DEFINITIONS.PATIENT_MASTER.headers));
  }

  if (visitSheet) {
    const headers = visitSheet.getRange(1, 1, 1, 11).getValues()[0];
    console.log('受診記録ヘッダー:', JSON.stringify(headers));
    console.log('期待するヘッダー:', JSON.stringify(COLUMN_DEFINITIONS.VISIT_RECORD.headers));
  }

  // 2. テストデータで登録試行
  const testPatient = {
    name: 'テスト太郎',
    nameKana: 'テストタロウ',
    birthDate: '1965-11-11',
    gender: '男性'
  };

  const testVisit = {
    visitDate: '2025-12-17',
    examTypeId: 'DOCK',
    courseId: 'SPECIFIC'
  };

  console.log('テストデータ（患者）:', JSON.stringify(testPatient));
  console.log('テストデータ（受診）:', JSON.stringify(testVisit));

  try {
    const patientId = createPatient(testPatient);
    console.log('患者作成成功: ' + patientId);

    testVisit.patientId = patientId;
    const visitId = createVisitRecord(testVisit);
    console.log('受診記録作成成功: ' + visitId);

  } catch (e) {
    console.log('エラー発生: ' + e.message);
    console.log('スタックトレース: ' + e.stack);
  }

  console.log('=== 登録テスト完了 ===');
}

// ============================================
// 所見CRUD関数（縦持ち・検査項目別構造）
// ============================================

/**
 * 所見ID（F00001形式）を生成
 * @returns {string} 新しい所見ID
 */
function generateFindingId() {
  const sheet = getSheet(DB_CONFIG.SHEETS.FINDINGS);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 'F00001';
  }

  // 最後のIDを取得
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (const row of ids) {
    const id = row[0];
    if (id && id.startsWith('F')) {
      const num = parseInt(id.substring(1), 10);
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }

  return 'F' + String(maxNum + 1).padStart(5, '0');
}

/**
 * 所見を作成
 * @param {Object} data - 所見データ
 * @param {string} data.patientId - 受診者ID（P00001形式）
 * @param {string} data.karteNo - カルテNo（6桁）
 * @param {string} data.itemId - 項目ID（H02xxxx形式）
 * @param {string} data.findingText - 所見テキスト
 * @param {string} data.judgment - 判定（A/B/C/D/E/F）
 * @param {string} [data.templateId] - テンプレートID
 * @param {Date|string} [data.examDate] - 検査日
 * @param {string} [data.inputBy] - 入力者
 * @returns {string} 作成された所見ID
 */
function createFinding(data) {
  if (!data.patientId || !data.karteNo || !data.itemId) {
    throw new Error('必須フィールドが不足しています: patientId, karteNo, itemId');
  }

  const sheet = getSheet(DB_CONFIG.SHEETS.FINDINGS);
  const findingId = generateFindingId();
  const now = new Date();
  const cols = FINDINGS_DEF.columns;

  const rowData = new Array(FINDINGS_DEF.headers.length).fill('');
  rowData[cols.FINDING_ID] = findingId;
  rowData[cols.PATIENT_ID] = data.patientId;
  rowData[cols.KARTE_NO] = String(data.karteNo).padStart(6, '0');
  rowData[cols.ITEM_ID] = data.itemId;
  rowData[cols.FINDING_TEXT] = data.findingText || '';
  rowData[cols.JUDGMENT] = data.judgment || '';
  rowData[cols.TEMPLATE_ID] = data.templateId || '';
  rowData[cols.EXAM_DATE] = data.examDate || now;
  rowData[cols.INPUT_BY] = data.inputBy || '';
  rowData[cols.CREATED_AT] = now;
  rowData[cols.UPDATED_AT] = now;

  sheet.appendRow(rowData);
  logInfo(`所見作成: ${findingId} (患者: ${data.patientId}, 項目: ${data.itemId})`);

  return findingId;
}

/**
 * カルテNoで所見を取得
 * @param {string} karteNo - カルテNo（6桁）
 * @returns {Array<Object>} 所見データの配列
 */
function getFindingsByKarteNo(karteNo) {
  const normalizedKarteNo = String(karteNo).padStart(6, '0');
  const sheet = getSheet(DB_CONFIG.SHEETS.FINDINGS);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, FINDINGS_DEF.headers.length).getValues();
  const cols = FINDINGS_DEF.columns;
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowKarteNo = String(row[cols.KARTE_NO]).padStart(6, '0');

    if (rowKarteNo === normalizedKarteNo) {
      results.push({
        findingId: row[cols.FINDING_ID],
        patientId: row[cols.PATIENT_ID],
        karteNo: rowKarteNo,
        itemId: row[cols.ITEM_ID],
        findingText: row[cols.FINDING_TEXT],
        judgment: row[cols.JUDGMENT],
        templateId: row[cols.TEMPLATE_ID],
        examDate: row[cols.EXAM_DATE],
        inputBy: row[cols.INPUT_BY],
        createdAt: row[cols.CREATED_AT],
        updatedAt: row[cols.UPDATED_AT],
        rowIndex: i + 2
      });
    }
  }

  return results;
}

/**
 * 受診者IDで所見を取得
 * @param {string} patientId - 受診者ID（P00001形式）
 * @returns {Array<Object>} 所見データの配列
 */
function getFindingsByPatientId(patientId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.FINDINGS);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, FINDINGS_DEF.headers.length).getValues();
  const cols = FINDINGS_DEF.columns;
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (row[cols.PATIENT_ID] === patientId) {
      results.push({
        findingId: row[cols.FINDING_ID],
        patientId: row[cols.PATIENT_ID],
        karteNo: String(row[cols.KARTE_NO]).padStart(6, '0'),
        itemId: row[cols.ITEM_ID],
        findingText: row[cols.FINDING_TEXT],
        judgment: row[cols.JUDGMENT],
        templateId: row[cols.TEMPLATE_ID],
        examDate: row[cols.EXAM_DATE],
        inputBy: row[cols.INPUT_BY],
        createdAt: row[cols.CREATED_AT],
        updatedAt: row[cols.UPDATED_AT],
        rowIndex: i + 2
      });
    }
  }

  return results;
}

/**
 * 所見IDで所見を取得
 * @param {string} findingId - 所見ID（F00001形式）
 * @returns {Object|null} 所見データ
 */
function getFindingById(findingId) {
  const sheet = getSheet(DB_CONFIG.SHEETS.FINDINGS);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return null;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, FINDINGS_DEF.headers.length).getValues();
  const cols = FINDINGS_DEF.columns;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (row[cols.FINDING_ID] === findingId) {
      return {
        findingId: row[cols.FINDING_ID],
        patientId: row[cols.PATIENT_ID],
        karteNo: String(row[cols.KARTE_NO]).padStart(6, '0'),
        itemId: row[cols.ITEM_ID],
        findingText: row[cols.FINDING_TEXT],
        judgment: row[cols.JUDGMENT],
        templateId: row[cols.TEMPLATE_ID],
        examDate: row[cols.EXAM_DATE],
        inputBy: row[cols.INPUT_BY],
        createdAt: row[cols.CREATED_AT],
        updatedAt: row[cols.UPDATED_AT],
        rowIndex: i + 2
      };
    }
  }

  return null;
}

/**
 * 所見を更新
 * @param {string} findingId - 所見ID（F00001形式）
 * @param {Object} data - 更新データ
 * @returns {boolean} 更新成功
 */
function updateFinding(findingId, data) {
  const existing = getFindingById(findingId);
  if (!existing) {
    throw new Error(`所見が見つかりません: ${findingId}`);
  }

  const sheet = getSheet(DB_CONFIG.SHEETS.FINDINGS);
  const cols = FINDINGS_DEF.columns;
  const rowIndex = existing.rowIndex;

  // 更新可能なフィールド
  if (data.findingText !== undefined) {
    sheet.getRange(rowIndex, cols.FINDING_TEXT + 1).setValue(data.findingText);
  }
  if (data.judgment !== undefined) {
    sheet.getRange(rowIndex, cols.JUDGMENT + 1).setValue(data.judgment);
  }
  if (data.templateId !== undefined) {
    sheet.getRange(rowIndex, cols.TEMPLATE_ID + 1).setValue(data.templateId);
  }
  if (data.examDate !== undefined) {
    sheet.getRange(rowIndex, cols.EXAM_DATE + 1).setValue(data.examDate);
  }
  if (data.inputBy !== undefined) {
    sheet.getRange(rowIndex, cols.INPUT_BY + 1).setValue(data.inputBy);
  }

  // 更新日時を更新
  sheet.getRange(rowIndex, cols.UPDATED_AT + 1).setValue(new Date());

  logInfo(`所見更新: ${findingId}`);
  return true;
}

/**
 * 所見を削除
 * @param {string} findingId - 所見ID（F00001形式）
 * @returns {boolean} 削除成功
 */
function deleteFinding(findingId) {
  const existing = getFindingById(findingId);
  if (!existing) {
    throw new Error(`所見が見つかりません: ${findingId}`);
  }

  const sheet = getSheet(DB_CONFIG.SHEETS.FINDINGS);
  sheet.deleteRow(existing.rowIndex);

  logInfo(`所見削除: ${findingId}`);
  return true;
}

/**
 * 所見テンプレートを取得
 * @param {string} [itemId] - 項目ID（指定しない場合は全件）
 * @param {string} [judgment] - 判定（A/B/C/D/E/F）
 * @returns {Array<Object>} テンプレート配列
 */
function getFindingTemplates(itemId, judgment) {
  const sheet = getSheet(DB_CONFIG.SHEETS.FINDING_TEMPLATE);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, FINDING_TEMPLATE_DEF.headers.length).getValues();
  const cols = FINDING_TEMPLATE_DEF.columns;
  const results = [];

  for (const row of data) {
    // 有効フラグチェック
    if (row[cols.IS_ACTIVE] !== true && row[cols.IS_ACTIVE] !== 'TRUE') {
      continue;
    }

    // 項目IDフィルタ
    if (itemId && row[cols.ITEM_ID] !== itemId) {
      continue;
    }

    // 判定フィルタ
    if (judgment && row[cols.JUDGMENT] !== judgment) {
      continue;
    }

    results.push({
      templateId: row[cols.TEMPLATE_ID],
      itemId: row[cols.ITEM_ID],
      category: row[cols.CATEGORY],
      judgment: row[cols.JUDGMENT],
      templateText: row[cols.TEMPLATE_TEXT],
      priority: row[cols.PRIORITY],
      isActive: row[cols.IS_ACTIVE]
    });
  }

  // 優先順位でソート
  results.sort((a, b) => (a.priority || 999) - (b.priority || 999));

  return results;
}

/**
 * 所見を一括作成/更新（upsert）
 * @param {string} karteNo - カルテNo
 * @param {Array<Object>} findingsData - 所見データの配列
 * @returns {Object} 処理結果 { created: number, updated: number }
 */
function upsertFindings(karteNo, findingsData) {
  const existingFindings = getFindingsByKarteNo(karteNo);
  const existingByItemId = {};

  for (const f of existingFindings) {
    existingByItemId[f.itemId] = f;
  }

  let created = 0;
  let updated = 0;

  for (const data of findingsData) {
    const existing = existingByItemId[data.itemId];

    if (existing) {
      // 更新
      updateFinding(existing.findingId, data);
      updated++;
    } else {
      // 新規作成
      data.karteNo = karteNo;
      createFinding(data);
      created++;
    }
  }

  logInfo(`所見一括処理: カルテNo=${karteNo}, 作成=${created}, 更新=${updated}`);
  return { created, updated };
}
