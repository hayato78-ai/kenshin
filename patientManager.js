/**
 * 受診者管理モジュール
 *
 * 機能:
 * - 受診者検索・一覧表示（SCR-003/004）
 * - 新規受診者登録（SCR-007）
 * - 企業別フィルタ・ソート機能
 *
 * 画面仕様:
 * - SCR-003: 受診者検索画面
 * - SCR-004: 受診者一覧画面
 * - SCR-007: 新規受診登録画面
 */

// ============================================
// 定数定義
// ============================================

const PATIENT_MANAGER_CONFIG = {
  // 検索結果の最大件数
  MAX_SEARCH_RESULTS: 500,

  // ステータス定義
  STATUS: {
    INPUT: '入力中',
    COMPLETE: '完了',
    PENDING: '保留'
  },

  // ID採番プレフィックス
  ID_PREFIX: 'P',

  // 必須フィールド（新規登録時）
  REQUIRED_FIELDS: ['name', 'examDate', 'courseId'],

  // 検索可能フィールド
  SEARCHABLE_FIELDS: ['name', 'nameKana', 'companyName', 'patientId', 'karteNo', 'bmlPatientId']
};

// ============================================
// 受診者検索機能
// ============================================

/**
 * 受診者を検索（旧版 - 使用しない）
 * @deprecated CRUD.gs の searchPatients を使用
 * @param {Object} criteria - 検索条件
 * @returns {Array} 検索結果
 */
function searchPatients_legacy(criteria) {
  logInfo('受診者検索開始: ' + JSON.stringify(criteria));

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const data = patientSheet.getDataRange().getValues();
    const headers = data[0];
    const results = [];

    // ヘッダーのインデックスを取得（17列構造 - カルテNo追加版）
    const colIndex = {
      patientId: 0,      // A: 受診者ID
      karteNo: 1,        // B: カルテNo（クリニック患者ID）★追加
      status: 2,         // C: ステータス
      examDate: 3,       // D: 受診日
      name: 4,           // E: 氏名
      nameKana: 5,       // F: カナ
      gender: 6,         // G: 性別
      birthDate: 7,      // H: 生年月日
      age: 8,            // I: 年齢
      course: 9,         // J: 受診コース
      company: 10,       // K: 事業所名
      department: 11,    // L: 所属
      overallJudgment: 12, // M: 総合判定
      csvImportDate: 13, // N: CSV取込日時
      lastUpdated: 14,   // O: 最終更新日時
      exportDate: 15,    // P: 出力日時
      bmlPatientId: 16   // Q: BML患者ID
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 空行スキップ
      if (!row[colIndex.patientId]) continue;

      // フィルタリング
      if (!matchesCriteria(row, colIndex, criteria)) continue;

      results.push({
        patientId: row[colIndex.patientId],
        karteNo: row[colIndex.karteNo] || '',  // ★カルテNo追加
        status: row[colIndex.status],
        examDate: row[colIndex.examDate] ? formatDate(row[colIndex.examDate]) : '',
        name: row[colIndex.name],
        nameKana: row[colIndex.nameKana],
        gender: row[colIndex.gender],
        birthDate: row[colIndex.birthDate] ? formatDate(row[colIndex.birthDate]) : '',
        age: row[colIndex.age],
        course: row[colIndex.course],
        company: row[colIndex.company],
        department: row[colIndex.department],
        overallJudgment: row[colIndex.overallJudgment],
        bmlPatientId: row[colIndex.bmlPatientId] || '',
        rowIndex: i + 1
      });

      // 最大件数制限
      if (results.length >= PATIENT_MANAGER_CONFIG.MAX_SEARCH_RESULTS) {
        break;
      }
    }

    // ソート（受診日降順）
    results.sort((a, b) => {
      if (!a.examDate) return 1;
      if (!b.examDate) return -1;
      return new Date(b.examDate) - new Date(a.examDate);
    });

    logInfo(`受診者検索完了: ${results.length}件`);
    return results;

  } catch (e) {
    logError('searchPatients', e);
    throw e;
  }
}

/**
 * 検索条件にマッチするかチェック
 * @param {Array} row - データ行
 * @param {Object} colIndex - カラムインデックス
 * @param {Object} criteria - 検索条件
 * @returns {boolean}
 */
function matchesCriteria(row, colIndex, criteria) {
  // 名前検索（部分一致）
  if (criteria.name) {
    const searchName = criteria.name.toLowerCase();
    const name = String(row[colIndex.name] || '').toLowerCase();
    const nameKana = String(row[colIndex.nameKana] || '').toLowerCase();
    if (!name.includes(searchName) && !nameKana.includes(searchName)) {
      return false;
    }
  }

  // 企業フィルタ
  if (criteria.companyId || criteria.companyName) {
    const company = String(row[colIndex.company] || '');
    if (criteria.companyId && company !== criteria.companyId) {
      // 企業名でも一致をチェック
      if (criteria.companyName && company !== criteria.companyName) {
        return false;
      }
    }
    if (criteria.companyName && !company.includes(criteria.companyName)) {
      return false;
    }
  }

  // 日付範囲フィルタ
  if (criteria.dateFrom || criteria.dateTo) {
    const examDate = row[colIndex.examDate];
    if (examDate) {
      const examDateObj = new Date(examDate);
      if (criteria.dateFrom && examDateObj < new Date(criteria.dateFrom)) {
        return false;
      }
      if (criteria.dateTo && examDateObj > new Date(criteria.dateTo)) {
        return false;
      }
    } else if (criteria.dateFrom || criteria.dateTo) {
      // 日付が空で日付範囲指定がある場合は除外
      return false;
    }
  }

  // ステータスフィルタ
  if (criteria.status && criteria.status !== 'all') {
    if (row[colIndex.status] !== criteria.status) {
      return false;
    }
  }

  // カルテNoフィルタ（完全一致）★追加
  if (criteria.karteNo) {
    if (String(row[colIndex.karteNo] || '') !== String(criteria.karteNo)) {
      return false;
    }
  }

  // コースフィルタ
  if (criteria.courseId) {
    if (row[colIndex.course] !== criteria.courseId) {
      return false;
    }
  }

  // BML患者IDフィルタ（完全一致）
  if (criteria.bmlPatientId) {
    if (String(row[colIndex.bmlPatientId] || '') !== String(criteria.bmlPatientId)) {
      return false;
    }
  }

  return true;
}

/**
 * 受診者詳細を取得
 * @param {string} patientId - 受診者ID
 * @returns {Object} 受診者詳細データ
 */
function getPatientDetail(patientId) {
  logInfo('受診者詳細取得: ' + patientId);

  try {
    // 基本情報を取得
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const data = patientSheet.getDataRange().getValues();
    let patientRow = null;
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === patientId) {
        patientRow = data[i];
        rowIndex = i + 1;
        break;
      }
    }

    if (!patientRow) {
      return null;
    }

    // 17列構造（カルテNo追加版）
    const patient = {
      patientId: patientRow[0],      // A: 受診者ID
      karteNo: patientRow[1] || '',  // B: カルテNo ★追加
      status: patientRow[2],         // C: ステータス
      examDate: patientRow[3] ? formatDate(patientRow[3]) : '',  // D: 受診日
      name: patientRow[4],           // E: 氏名
      nameKana: patientRow[5],       // F: カナ
      gender: patientRow[6],         // G: 性別
      birthDate: patientRow[7] ? formatDate(patientRow[7]) : '',  // H: 生年月日
      age: patientRow[8],            // I: 年齢
      course: patientRow[9],         // J: 受診コース
      company: patientRow[10],       // K: 事業所名
      department: patientRow[11],    // L: 所属
      overallJudgment: patientRow[12],  // M: 総合判定
      bmlPatientId: patientRow[16] || '',  // Q: BML患者ID
      rowIndex: rowIndex,
      physical: {},
      blood: {}
    };

    // 身体測定データを取得
    const physicalSheet = getSheet(CONFIG.SHEETS.PHYSICAL);
    if (physicalSheet) {
      const physicalData = physicalSheet.getDataRange().getValues();
      for (let i = 1; i < physicalData.length; i++) {
        if (physicalData[i][0] === patientId) {
          patient.physical = {
            height: physicalData[i][1],
            weight: physicalData[i][2],
            standardWeight: physicalData[i][3],
            BMI: physicalData[i][4],
            bodyFat: physicalData[i][5],
            waist: physicalData[i][6],
            bpSys1: physicalData[i][7],
            bpDia1: physicalData[i][8],
            bpSys2: physicalData[i][9],
            bpDia2: physicalData[i][10]
          };
          break;
        }
      }
    }

    // 血液検査データを取得
    const bloodSheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
    if (bloodSheet) {
      const bloodData = bloodSheet.getDataRange().getValues();
      const bloodHeaders = bloodData[0];
      for (let i = 1; i < bloodData.length; i++) {
        if (bloodData[i][0] === patientId) {
          for (let j = 1; j < bloodHeaders.length; j++) {
            patient.blood[bloodHeaders[j]] = bloodData[i][j];
          }
          break;
        }
      }
    }

    return patient;

  } catch (e) {
    logError('getPatientDetail', e);
    throw e;
  }
}

// ============================================
// 新規受診者登録機能
// ============================================

/**
 * 新規受診者を登録
 * @param {Object} data - 登録データ
 * @returns {Object} 登録結果 {success, patientId, error}
 */
function registerNewPatient(data) {
  logInfo('新規受診者登録開始');

  try {
    // 必須フィールドチェック
    if (!data.name || !data.name.trim()) {
      throw new Error('氏名は必須です');
    }
    if (!data.examDate) {
      throw new Error('受診日は必須です');
    }

    // 受診者IDを生成
    const patientId = generatePatientId();

    // 受診者マスタに登録
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    // 年齢計算
    let age = '';
    if (data.birthDate) {
      age = calculateAge(new Date(data.birthDate), new Date(data.examDate));
    }

    // 新規行を追加（17列構造 - カルテNo追加版）
    const newRow = [
      patientId,                                    // A: 受診者ID
      data.karteNo || '',                           // B: カルテNo ★追加
      PATIENT_MANAGER_CONFIG.STATUS.INPUT,          // C: ステータス
      new Date(data.examDate),                      // D: 受診日
      data.name,                                    // E: 氏名
      data.nameKana || '',                          // F: カナ
      data.gender || '',                            // G: 性別
      data.birthDate ? new Date(data.birthDate) : '', // H: 生年月日
      age,                                          // I: 年齢
      data.courseName || data.courseId || '',       // J: 受診コース
      data.companyName || data.companyId || '',     // K: 事業所名
      data.department || '',                        // L: 所属
      '',                                           // M: 総合判定
      '',                                           // N: CSV取込日時
      new Date(),                                   // O: 最終更新日時
      '',                                           // P: 出力日時
      data.bmlPatientId || ''                       // Q: BML患者ID
    ];

    patientSheet.appendRow(newRow);

    // 身体測定シートに空行を追加（データ構造を準備）
    const physicalSheet = getSheet(CONFIG.SHEETS.PHYSICAL);
    if (physicalSheet) {
      const physicalRow = [patientId];
      // 残りの列は空
      for (let i = 1; i < 23; i++) {
        physicalRow.push('');
      }
      physicalSheet.appendRow(physicalRow);
    }

    // 血液検査シートに空行を追加
    const bloodSheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
    if (bloodSheet) {
      const bloodRow = [patientId];
      // 残りの列は空
      for (let i = 1; i < 28; i++) {
        bloodRow.push('');
      }
      bloodSheet.appendRow(bloodRow);
    }

    logInfo(`新規受診者登録完了: ${patientId}`);

    return {
      success: true,
      patientId: patientId,
      message: `受診者を登録しました（ID: ${patientId}）`
    };

  } catch (e) {
    logError('registerNewPatient', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 受診者IDを生成
 * @returns {string} 新しい受診者ID（P00001形式）
 */
function generatePatientId() {
  const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
  if (!patientSheet) {
    return 'P00001';
  }

  const data = patientSheet.getDataRange().getValues();
  let maxNum = 0;

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '');
    if (id.startsWith('P')) {
      const num = parseInt(id.substring(1), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  const newNum = maxNum + 1;
  return 'P' + String(newNum).padStart(5, '0');
}

/**
 * 年齢を計算
 * @param {Date} birthDate - 生年月日
 * @param {Date} targetDate - 基準日
 * @returns {number} 年齢
 */
function calculateAge(birthDate, targetDate) {
  let age = targetDate.getFullYear() - birthDate.getFullYear();
  const monthDiff = targetDate.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && targetDate.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
}

/**
 * 受診者情報を更新
 * @param {string} patientId - 受診者ID
 * @param {Object} data - 更新データ
 * @returns {Object} 更新結果
 */
function updatePatient(patientId, data) {
  logInfo('受診者情報更新: ' + patientId);

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const allData = patientSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === patientId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('受診者が見つかりません: ' + patientId);
    }

    // 更新する列（17列構造 - カルテNo追加版）
    // 列番号: A=1, B=2(カルテNo), C=3(ステータス), D=4(受診日), E=5(氏名), F=6(カナ), G=7(性別), H=8(生年月日), I=9(年齢), J=10(コース), K=11(事業所), L=12(所属), M=13(総合判定), N=14(CSV取込), O=15(最終更新), P=16(出力), Q=17(BML患者ID)
    if (data.karteNo !== undefined) patientSheet.getRange(rowIndex, 2).setValue(data.karteNo);  // B: カルテNo
    if (data.status !== undefined) patientSheet.getRange(rowIndex, 3).setValue(data.status);   // C: ステータス
    if (data.name !== undefined) patientSheet.getRange(rowIndex, 5).setValue(data.name);       // E: 氏名
    if (data.nameKana !== undefined) patientSheet.getRange(rowIndex, 6).setValue(data.nameKana); // F: カナ
    if (data.gender !== undefined) patientSheet.getRange(rowIndex, 7).setValue(data.gender);   // G: 性別
    if (data.birthDate !== undefined) patientSheet.getRange(rowIndex, 8).setValue(data.birthDate ? new Date(data.birthDate) : ''); // H: 生年月日
    if (data.course !== undefined) patientSheet.getRange(rowIndex, 10).setValue(data.course);  // J: コース
    if (data.company !== undefined) patientSheet.getRange(rowIndex, 11).setValue(data.company); // K: 事業所
    if (data.department !== undefined) patientSheet.getRange(rowIndex, 12).setValue(data.department); // L: 所属
    if (data.bmlPatientId !== undefined) patientSheet.getRange(rowIndex, 17).setValue(data.bmlPatientId); // Q: BML患者ID

    // 最終更新日時を更新 (O列 = 15番目)
    patientSheet.getRange(rowIndex, 15).setValue(new Date());

    // 年齢を再計算
    if (data.birthDate !== undefined) {
      const examDate = patientSheet.getRange(rowIndex, 4).getValue();  // D列 = 受診日
      if (examDate && data.birthDate) {
        const age = calculateAge(new Date(data.birthDate), new Date(examDate));
        patientSheet.getRange(rowIndex, 9).setValue(age);  // I列 = 年齢
      }
    }

    logInfo('受診者情報更新完了');

    return {
      success: true,
      message: '受診者情報を更新しました'
    };

  } catch (e) {
    logError('updatePatient', e);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================
// BML患者ID関連関数
// ============================================

/**
 * BML患者IDで受診者を検索（17列構造対応）
 * @param {string} bmlPatientId - BML患者ID（例: 999991）
 * @returns {Object|null} 受診者データまたはnull
 */
function findPatientByBmlId(bmlPatientId) {
  if (!bmlPatientId) return null;

  logInfo('BML患者IDで検索: ' + bmlPatientId);

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const data = patientSheet.getDataRange().getValues();
    const bmlIdCol = 16; // Q列: BML患者ID列（0始まり）★16に変更

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 空行スキップ
      if (!row[0]) continue;

      // BML患者IDで照合（文字列比較）
      if (String(row[bmlIdCol] || '') === String(bmlPatientId)) {
        logInfo('BML患者ID一致: ' + row[0]);
        // 17列構造（カルテNo追加版）
        return {
          patientId: row[0],       // A: 受診者ID
          karteNo: row[1] || '',   // B: カルテNo ★追加
          status: row[2],          // C: ステータス
          examDate: row[3] ? formatDate(row[3]) : '',  // D: 受診日
          name: row[4],            // E: 氏名
          nameKana: row[5],        // F: カナ
          gender: row[6],          // G: 性別
          birthDate: row[7] ? formatDate(row[7]) : '',  // H: 生年月日
          age: row[8],             // I: 年齢
          course: row[9],          // J: 受診コース
          company: row[10],        // K: 事業所名
          department: row[11],     // L: 所属
          overallJudgment: row[12], // M: 総合判定
          bmlPatientId: row[bmlIdCol] || '',  // Q: BML患者ID
          rowIndex: i + 1
        };
      }
    }

    logInfo('BML患者ID該当なし: ' + bmlPatientId);
    return null;

  } catch (e) {
    logError('findPatientByBmlId', e);
    throw e;
  }
}

/**
 * BML患者IDの重複チェック（17列構造対応）
 * @param {string} bmlPatientId - チェックするBML患者ID
 * @param {string} excludePatientId - 除外する受診者ID（更新時用）
 * @returns {boolean} 重複している場合true
 */
function isBmlPatientIdDuplicate(bmlPatientId, excludePatientId) {
  if (!bmlPatientId) return false;

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) return false;

    const data = patientSheet.getDataRange().getValues();
    const bmlIdCol = 16;  // Q列: BML患者ID（0始まり）★16に変更

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      // 自分自身は除外
      if (excludePatientId && row[0] === excludePatientId) continue;

      if (String(row[bmlIdCol] || '') === String(bmlPatientId)) {
        return true;
      }
    }

    return false;

  } catch (e) {
    logError('isBmlPatientIdDuplicate', e);
    return false;
  }
}

// ============================================
// UI関連関数
// ============================================

/**
 * 受診者検索ダイアログを表示
 */
function showPatientSearchDialog() {
  const html = HtmlService.createHtmlOutput(getPatientSearchDialogHtml())
    .setWidth(900)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, '受診者検索');
}

/**
 * 受診者検索ダイアログのHTML
 * @returns {string} HTML文字列
 */
function getPatientSearchDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 15px;
      margin: 0;
      line-height: 1.5;
    }
    h3 {
      margin: 0 0 15px 0;
      color: #1a73e8;
      border-bottom: 2px solid #1a73e8;
      padding-bottom: 8px;
    }
    .search-panel {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      margin-bottom: 15px;
    }
    .search-row {
      display: flex;
      gap: 15px;
      flex-wrap: wrap;
      margin-bottom: 10px;
    }
    .search-field {
      flex: 1;
      min-width: 150px;
    }
    .search-field label {
      display: block;
      margin-bottom: 4px;
      font-weight: 500;
      font-size: 12px;
      color: #555;
    }
    .search-field input, .search-field select {
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
    }
    .btn-row {
      display: flex;
      gap: 10px;
      justify-content: flex-end;
      margin-top: 10px;
    }
    .btn {
      padding: 8px 20px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 13px;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    .btn-primary:hover {
      background: #1557b0;
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-success {
      background: #0f9d58;
      color: white;
    }
    .btn-success:hover {
      background: #0b8043;
    }
    .results-panel {
      border: 1px solid #ddd;
      border-radius: 8px;
      overflow: hidden;
    }
    .results-header {
      background: #f1f3f4;
      padding: 10px 15px;
      font-weight: bold;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .results-count {
      font-size: 12px;
      color: #666;
      font-weight: normal;
    }
    .results-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }
    .results-table th {
      background: #f8f9fa;
      padding: 10px 8px;
      text-align: left;
      border-bottom: 2px solid #ddd;
      position: sticky;
      top: 0;
      font-weight: 600;
    }
    .results-table td {
      padding: 10px 8px;
      border-bottom: 1px solid #eee;
    }
    .results-table tr:hover {
      background: #f5f8ff;
    }
    .results-table tr.selected {
      background: #e8f0fe;
    }
    .results-body {
      max-height: 350px;
      overflow-y: auto;
    }
    .status-badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 10px;
      font-size: 11px;
    }
    .status-complete { background: #d4edda; color: #155724; }
    .status-input { background: #fff3cd; color: #856404; }
    .status-pending { background: #f8d7da; color: #721c24; }
    .judgment-badge {
      display: inline-block;
      width: 24px;
      height: 24px;
      line-height: 24px;
      text-align: center;
      border-radius: 50%;
      font-weight: bold;
      font-size: 12px;
    }
    .judgment-A { background: #e8f5e9; color: #2e7d32; }
    .judgment-B { background: #fff8e1; color: #f9a825; }
    .judgment-C { background: #fff3e0; color: #ef6c00; }
    .judgment-D { background: #ffebee; color: #c62828; }
    .action-btn {
      padding: 4px 10px;
      font-size: 11px;
      border: 1px solid #1a73e8;
      background: white;
      color: #1a73e8;
      border-radius: 4px;
      cursor: pointer;
    }
    .action-btn:hover {
      background: #e8f0fe;
    }
    .loading {
      text-align: center;
      padding: 40px;
      color: #666;
    }
    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #1a73e8;
      border-radius: 50%;
      width: 30px;
      height: 30px;
      animation: spin 1s linear infinite;
      margin: 0 auto 10px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    .error { color: #d93025; margin-top: 10px; }
    .no-results {
      text-align: center;
      padding: 40px;
      color: #666;
    }
  </style>
</head>
<body>
  <h3>🔍 受診者検索</h3>

  <div class="search-panel">
    <div class="search-row">
      <div class="search-field">
        <label>氏名・カナ</label>
        <input type="text" id="searchName" placeholder="部分一致検索">
      </div>
      <div class="search-field">
        <label>企業</label>
        <select id="searchCompany">
          <option value="">すべて</option>
        </select>
      </div>
      <div class="search-field">
        <label>ステータス</label>
        <select id="searchStatus">
          <option value="all">すべて</option>
          <option value="入力中">入力中</option>
          <option value="完了">完了</option>
          <option value="保留">保留</option>
        </select>
      </div>
    </div>
    <div class="search-row">
      <div class="search-field">
        <label>受診日（From）</label>
        <input type="date" id="searchDateFrom">
      </div>
      <div class="search-field">
        <label>受診日（To）</label>
        <input type="date" id="searchDateTo">
      </div>
      <div class="search-field">
        <label>カルテNo</label>
        <input type="text" id="searchKarteNo" placeholder="999999">
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-secondary" onclick="clearSearch()">クリア</button>
      <button class="btn btn-primary" onclick="executeSearch()">検索</button>
    </div>
  </div>

  <div class="results-panel">
    <div class="results-header">
      <span>検索結果</span>
      <span class="results-count" id="resultsCount">0件</span>
    </div>
    <div class="results-body" id="resultsBody">
      <div class="no-results">検索条件を入力して「検索」ボタンをクリックしてください</div>
    </div>
  </div>

  <div style="margin-top: 15px; text-align: right;">
    <button class="btn btn-secondary" onclick="google.script.host.close()">閉じる</button>
    <button class="btn btn-success" onclick="openNewRegistration()">＋ 新規登録</button>
  </div>

  <script>
    let searchResults = [];

    // 企業リストを読み込み
    google.script.run
      .withSuccessHandler((companies) => {
        const select = document.getElementById('searchCompany');
        companies.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.name;
          opt.textContent = c.name;
          select.appendChild(opt);
        });
      })
      .getCompanyListForDropdown();

    function executeSearch() {
      const criteria = {
        name: document.getElementById('searchName').value,
        companyName: document.getElementById('searchCompany').value,
        status: document.getElementById('searchStatus').value,
        dateFrom: document.getElementById('searchDateFrom').value,
        dateTo: document.getElementById('searchDateTo').value,
        karteNo: document.getElementById('searchKarteNo').value  // ★カルテNoで検索
      };

      document.getElementById('resultsBody').innerHTML =
        '<div class="loading"><div class="spinner"></div>検索中...</div>';

      google.script.run
        .withSuccessHandler(renderResults)
        .withFailureHandler((e) => {
          document.getElementById('resultsBody').innerHTML =
            '<div class="error">エラー: ' + e.message + '</div>';
        })
        .searchPatients(criteria);
    }

    function renderResults(results) {
      searchResults = results;
      document.getElementById('resultsCount').textContent = results.length + '件';

      if (results.length === 0) {
        document.getElementById('resultsBody').innerHTML =
          '<div class="no-results">該当する受診者が見つかりません</div>';
        return;
      }

      // ★内部ID（受診者ID）は非表示、カルテNoを表示
      let html = '<table class="results-table"><thead><tr>' +
        '<th>カルテNo</th><th>氏名</th><th>企業</th><th>受診日</th>' +
        '<th>コース</th><th>判定</th><th>ステータス</th><th>操作</th>' +
        '</tr></thead><tbody>';

      results.forEach((p, idx) => {
        const statusClass = p.status === '完了' ? 'status-complete' :
                           p.status === '保留' ? 'status-pending' : 'status-input';
        const judgmentClass = p.overallJudgment ? 'judgment-' + p.overallJudgment : '';

        // ★カルテNoを表示（内部IDはdata属性に保持）
        html += '<tr onclick="selectRow(' + idx + ')" data-idx="' + idx + '" data-patient-id="' + p.patientId + '">' +
          '<td>' + (p.karteNo || '-') + '</td>' +
          '<td><strong>' + p.name + '</strong><br><small style="color:#666">' + (p.nameKana || '') + '</small></td>' +
          '<td>' + (p.company || '-') + '</td>' +
          '<td>' + (p.examDate || '-') + '</td>' +
          '<td>' + (p.course || '-') + '</td>' +
          '<td>' + (p.overallJudgment ? '<span class="judgment-badge ' + judgmentClass + '">' + p.overallJudgment + '</span>' : '-') + '</td>' +
          '<td><span class="status-badge ' + statusClass + '">' + p.status + '</span></td>' +
          '<td><button class="action-btn" onclick="viewDetail(\\'' + p.patientId + '\\')">詳細</button></td>' +
          '</tr>';
      });

      html += '</tbody></table>';
      document.getElementById('resultsBody').innerHTML = html;
    }

    function selectRow(idx) {
      document.querySelectorAll('.results-table tr').forEach(tr => tr.classList.remove('selected'));
      const row = document.querySelector('tr[data-idx="' + idx + '"]');
      if (row) row.classList.add('selected');
    }

    function viewDetail(patientId) {
      alert('受診者詳細画面は現在準備中です。\\n受診者ID: ' + patientId);
      // 将来的にはここで詳細ダイアログを開く
    }

    function clearSearch() {
      document.getElementById('searchName').value = '';
      document.getElementById('searchCompany').value = '';
      document.getElementById('searchStatus').value = 'all';
      document.getElementById('searchDateFrom').value = '';
      document.getElementById('searchDateTo').value = '';
      document.getElementById('searchKarteNo').value = '';  // ★カルテNo
    }

    function openNewRegistration() {
      google.script.host.close();
      google.script.run.showNewPatientDialog();
    }

    // Enterキーで検索実行
    document.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') executeSearch();
    });
  </script>
</body>
</html>
`;
}

/**
 * 新規受診者登録ダイアログを表示
 */
function showNewPatientDialog() {
  const html = HtmlService.createHtmlOutput(getNewPatientDialogHtml())
    .setWidth(600)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, '新規受診者登録');
}

/**
 * 新規受診者登録ダイアログのHTML
 * @returns {string} HTML文字列
 */
function getNewPatientDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 20px;
      margin: 0;
      line-height: 1.6;
    }
    h3 {
      margin: 0 0 20px 0;
      color: #0f9d58;
      border-bottom: 2px solid #0f9d58;
      padding-bottom: 8px;
    }
    .form-section {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 20px;
    }
    .section-title {
      font-weight: bold;
      margin-bottom: 15px;
      color: #333;
      font-size: 14px;
    }
    .form-row {
      display: flex;
      gap: 15px;
      margin-bottom: 15px;
    }
    .form-group {
      flex: 1;
    }
    .form-group.wide {
      flex: 2;
    }
    .form-group label {
      display: block;
      margin-bottom: 5px;
      font-weight: 500;
    }
    .form-group label .required {
      color: #d93025;
      margin-left: 3px;
    }
    .form-group input, .form-group select {
      width: 100%;
      padding: 10px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
    }
    .form-group input:focus, .form-group select:focus {
      outline: none;
      border-color: #0f9d58;
      box-shadow: 0 0 0 2px rgba(15, 157, 88, 0.1);
    }
    .form-group input.error {
      border-color: #d93025;
    }
    .hint {
      font-size: 11px;
      color: #666;
      margin-top: 4px;
    }
    .btn-container {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 20px;
      padding-top: 20px;
      border-top: 1px solid #eee;
    }
    .btn {
      padding: 12px 30px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
    }
    .btn-primary {
      background: #0f9d58;
      color: white;
    }
    .btn-primary:hover {
      background: #0b8043;
    }
    .btn-primary:disabled {
      background: #ccc;
      cursor: not-allowed;
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-link {
      background: none;
      color: #1a73e8;
      text-decoration: underline;
      padding: 5px;
    }
    .message {
      padding: 15px;
      border-radius: 4px;
      margin-bottom: 15px;
      display: none;
    }
    .message.error {
      background: #fce8e6;
      color: #c5221f;
      display: block;
    }
    .message.success {
      background: #e6f4ea;
      color: #137333;
      display: block;
    }
    .loading {
      display: none;
      text-align: center;
      padding: 30px;
    }
    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #0f9d58;
      border-radius: 50%;
      width: 30px;
      height: 30px;
      animation: spin 1s linear infinite;
      margin: 0 auto 10px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <h3>＋ 新規受診者登録</h3>

  <div class="message" id="messageBox"></div>

  <div id="formContainer">
    <div class="form-section">
      <div class="section-title">基本情報</div>

      <div class="form-row">
        <div class="form-group wide">
          <label>氏名<span class="required">*</span></label>
          <input type="text" id="name" placeholder="山田 太郎">
        </div>
        <div class="form-group">
          <label>性別</label>
          <select id="gender">
            <option value="">選択してください</option>
            <option value="男">男</option>
            <option value="女">女</option>
          </select>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group wide">
          <label>フリガナ</label>
          <input type="text" id="nameKana" placeholder="ヤマダ タロウ">
        </div>
        <div class="form-group">
          <label>生年月日</label>
          <input type="date" id="birthDate">
        </div>
      </div>
    </div>

    <div class="form-section">
      <div class="section-title">受診情報</div>

      <div class="form-row">
        <div class="form-group">
          <label>受診日<span class="required">*</span></label>
          <input type="date" id="examDate">
        </div>
        <div class="form-group">
          <label>受診コース</label>
          <select id="courseId">
            <option value="">選択してください</option>
          </select>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>企業・事業所</label>
          <select id="companyId">
            <option value="">選択してください</option>
          </select>
        </div>
        <div class="form-group">
          <label>所属・部署</label>
          <input type="text" id="department" placeholder="営業部">
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>カルテNo</label>
          <input type="text" id="karteNo" placeholder="999999" maxlength="6">
          <div class="hint">※ クリニックの患者番号（6桁）- CSV取込時の主キー</div>
        </div>
        <div class="form-group">
          <label>BML患者ID</label>
          <input type="text" id="bmlPatientId" placeholder="457973">
          <div class="hint">※ BML検査所の患者ID（トレーサビリティ用）</div>
        </div>
      </div>

      <div class="hint">※ 検査結果は登録後に入力できます</div>
    </div>
  </div>

  <div class="loading" id="loading">
    <div class="spinner"></div>
    <div>登録中...</div>
  </div>

  <div class="btn-container">
    <button class="btn btn-link" onclick="openSearch()">← 検索に戻る</button>
    <div>
      <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
      <button class="btn btn-primary" id="submitBtn" onclick="submitForm()">登録する</button>
    </div>
  </div>

  <script>
    // 本日の日付をデフォルトに設定
    document.getElementById('examDate').valueAsDate = new Date();

    // 企業リストを読み込み
    google.script.run
      .withSuccessHandler((companies) => {
        const select = document.getElementById('companyId');
        companies.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.id;
          opt.textContent = c.name;
          opt.dataset.name = c.name;
          select.appendChild(opt);
        });
      })
      .getCompanyListForDropdown();

    // コースリストを読み込み
    google.script.run
      .withSuccessHandler((courses) => {
        const select = document.getElementById('courseId');
        courses.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.id;
          opt.textContent = c.name;
          opt.dataset.name = c.name;
          select.appendChild(opt);
        });
      })
      .getCourseListForDropdown();

    function submitForm() {
      // バリデーション
      const name = document.getElementById('name').value.trim();
      const examDate = document.getElementById('examDate').value;

      if (!name) {
        showMessage('error', '氏名を入力してください');
        document.getElementById('name').classList.add('error');
        return;
      }
      if (!examDate) {
        showMessage('error', '受診日を入力してください');
        document.getElementById('examDate').classList.add('error');
        return;
      }

      // データ収集
      const companySelect = document.getElementById('companyId');
      const courseSelect = document.getElementById('courseId');

      const data = {
        name: name,
        nameKana: document.getElementById('nameKana').value.trim(),
        gender: document.getElementById('gender').value,
        birthDate: document.getElementById('birthDate').value,
        examDate: examDate,
        courseId: courseSelect.value,
        courseName: courseSelect.selectedOptions[0]?.dataset?.name || courseSelect.value,
        companyId: companySelect.value,
        companyName: companySelect.selectedOptions[0]?.dataset?.name || companySelect.value,
        department: document.getElementById('department').value.trim(),
        karteNo: document.getElementById('karteNo').value.trim(),  // ★カルテNo追加
        bmlPatientId: document.getElementById('bmlPatientId').value.trim()
      };

      showLoading(true);
      hideMessage();

      google.script.run
        .withSuccessHandler(handleResult)
        .withFailureHandler(handleError)
        .registerNewPatient(data);
    }

    function handleResult(result) {
      showLoading(false);

      if (result.success) {
        showMessage('success', result.message);
        document.getElementById('submitBtn').disabled = true;

        // 3秒後に閉じる
        setTimeout(() => {
          google.script.host.close();
        }, 2000);
      } else {
        showMessage('error', result.error || '登録に失敗しました');
      }
    }

    function handleError(error) {
      showLoading(false);
      showMessage('error', 'エラー: ' + error.message);
    }

    function showLoading(show) {
      document.getElementById('loading').style.display = show ? 'block' : 'none';
      document.getElementById('formContainer').style.display = show ? 'none' : 'block';
    }

    function showMessage(type, text) {
      const box = document.getElementById('messageBox');
      box.className = 'message ' + type;
      box.textContent = text;
    }

    function hideMessage() {
      document.getElementById('messageBox').className = 'message';
      document.querySelectorAll('input.error').forEach(el => el.classList.remove('error'));
    }

    function openSearch() {
      google.script.host.close();
      google.script.run.showPatientSearchDialog();
    }
  </script>
</body>
</html>
`;
}

// ============================================
// ドロップダウン用データ取得関数
// ============================================

/**
 * コースリストを取得（ドロップダウン用）
 * @returns {Array} コースリスト [{id, name}]
 */
function getCourseListForDropdown() {
  const result = [];

  try {
    const courseSheet = getSheet(CONFIG.SHEETS.COURSE);
    if (!courseSheet) {
      // コースマスタがない場合はデフォルト値を返す
      return [
        { id: 'CRS001', name: '生活習慣病ドック' },
        { id: 'CRS002', name: '人間ドック標準' },
        { id: 'CRS003', name: 'がんドック' },
        { id: 'CRS004', name: '定期健診A' }
      ];
    }

    const data = courseSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 有効フラグがtrueのコースのみ
      if (row[5] !== false) {
        result.push({
          id: row[0],
          name: row[1]
        });
      }
    }
  } catch (e) {
    logError('getCourseListForDropdown', e);
  }

  return result;
}

// ============================================
// テスト関数
// ============================================

/**
 * 受診者検索のテスト
 */
function testSearchPatients() {
  const results = searchPatients({
    name: '山田',
    status: 'all'
  });
  logInfo('検索結果: ' + results.length + '件');
  if (results.length > 0) {
    logInfo('最初の結果: ' + JSON.stringify(results[0]));
  }
}

/**
 * 新規登録のテスト
 */
function testRegisterNewPatient() {
  const result = registerNewPatient({
    name: 'テスト太郎',
    nameKana: 'テストタロウ',
    gender: '男',
    birthDate: '1980-01-15',
    examDate: '2025-12-16',
    courseName: '生活習慣病ドック',
    companyName: 'テスト株式会社'
  });
  logInfo('登録結果: ' + JSON.stringify(result));
}

/**
 * 検索ダイアログのテスト表示
 */
function testShowPatientSearchDialog() {
  showPatientSearchDialog();
}

/**
 * 新規登録ダイアログのテスト表示
 */
function testShowNewPatientDialog() {
  showNewPatientDialog();
}
