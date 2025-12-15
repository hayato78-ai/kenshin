/**
 * AppSheet Webhook エンドポイント
 *
 * AppSheetからのリクエストを受け付けてGAS処理を実行
 * デプロイ後のURLをAppSheetのアクションに設定
 *
 * デプロイ手順:
 * 1. GASエディタで「デプロイ」→「新しいデプロイ」
 * 2. 種類: ウェブアプリ
 * 3. 実行ユーザー: 自分
 * 4. アクセスできるユーザー: 全員（または組織内全員）
 * 5. デプロイ後のURLをAppSheetアクションのWebhook URLに設定
 */

// ============================================
// Webhook エントリーポイント
// ============================================

/**
 * POST リクエストを処理
 * @param {Object} e - イベントオブジェクト
 * @returns {TextOutput} JSON レスポンス
 */
function doPost(e) {
  const startTime = new Date();

  try {
    // リクエストパラメータを解析
    const params = JSON.parse(e.postData.contents);
    const action = params.action;

    logInfo(`[Webhook] Action: ${action}, Params: ${JSON.stringify(params)}`);

    let result;

    switch (action) {
      case 'import_csv':
        result = handleImportCsv(params);
        break;

      case 'generate_guidance':
        result = handleGenerateGuidance(params);
        break;

      case 'export_excel':
        result = handleExportExcel(params);
        break;

      case 'sync_patient_data':
        result = handleSyncPatientData(params);
        break;

      case 'calculate_judgments':
        result = handleCalculateJudgments(params);
        break;

      case 'get_case_list':
        result = handleGetCaseList(params);
        break;

      case 'import_schedule':
        result = handleImportSchedule(params);
        break;

      case 'calculate_screening':
        result = handleCalculateScreening(params);
        break;

      case 'health_check':
        result = { success: true, message: 'Webhook is running', timestamp: new Date().toISOString() };
        break;

      default:
        result = { success: false, error: `Unknown action: ${action}` };
    }

    // 処理時間を記録
    const endTime = new Date();
    result.processingTime = endTime - startTime + 'ms';

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    logError('doPost', error);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message,
        stack: error.stack
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET リクエストを処理（ヘルスチェック用）
 * @param {Object} e - イベントオブジェクト
 * @returns {TextOutput} JSON レスポンス
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      message: 'AppSheet Webhook Endpoint',
      version: '1.0.0',
      availableActions: [
        'import_csv',
        'generate_guidance',
        'export_excel',
        'sync_patient_data',
        'calculate_judgments',
        'get_case_list',
        'import_schedule',
        'calculate_screening',
        'health_check'
      ]
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// アクションハンドラー
// ============================================

/**
 * CSV取込処理
 * @param {Object} params - { case_id, folder_id }
 * @returns {Object} 処理結果
 */
function handleImportCsv(params) {
  const caseId = params.case_id;
  const folderId = params.folder_id;

  if (!folderId) {
    return { success: false, error: 'folder_id is required' };
  }

  try {
    // 案件フォルダからCSVを検索
    const csvResult = findLatestJudgmentCsv(folderId);
    if (!csvResult.success) {
      return { success: false, error: csvResult.error };
    }

    // CSV処理実行
    const result = processCsvFile(csvResult.file);

    if (!result.success) {
      return { success: false, error: result.errors.join(', ') };
    }

    // AppSheet用スプレッドシートにデータ同期
    const syncResult = syncToAppSheetData(caseId, result.patientIds);

    return {
      success: true,
      message: `${result.patientIds.length}名のデータを取り込みました`,
      patientCount: result.patientIds.length,
      patientIds: result.patientIds,
      syncResult: syncResult
    };

  } catch (error) {
    logError('handleImportCsv', error);
    return { success: false, error: error.message };
  }
}

/**
 * 保健指導AI生成
 * @param {Object} params - { patient_id }
 * @returns {Object} 処理結果
 */
function handleGenerateGuidance(params) {
  const patientId = params.patient_id;

  if (!patientId) {
    return { success: false, error: 'patient_id is required' };
  }

  try {
    // 患者データを取得
    const patientData = getPatientDataForGuidance(patientId);

    if (!patientData) {
      return { success: false, error: 'Patient data not found' };
    }

    // Claude APIで指導文生成
    const guidance = generateGuidanceWithClaude(patientData);

    // AppSheetのguidanceシートに保存
    saveGuidanceToAppSheet(patientId, guidance);

    return {
      success: true,
      message: '保健指導文を生成しました',
      guidance: guidance
    };

  } catch (error) {
    logError('handleGenerateGuidance', error);
    return { success: false, error: error.message };
  }
}

/**
 * Excel出力処理
 * @param {Object} params - { patient_id, output_type }
 * @returns {Object} 処理結果
 */
function handleExportExcel(params) {
  const patientId = params.patient_id;
  const outputType = params.output_type || 'idheart';

  if (!patientId) {
    return { success: false, error: 'patient_id is required' };
  }

  try {
    // 患者データを収集
    const patientData = collectPatientDataForExport(patientId);

    if (!patientData) {
      return { success: false, error: 'Patient data not found' };
    }

    // Excel出力
    let result;
    if (CONFIG.getExamType() === 'ROSAI_SECONDARY') {
      result = generateRosaiExcelWithHL(patientData, patientData.caseInfo);
    } else {
      result = generateExcelFromTemplate(patientData);
    }

    if (result.success) {
      // ステータスを更新
      updatePatientStep(patientId, 6, 'COMPLETED');
    }

    return result;

  } catch (error) {
    logError('handleExportExcel', error);
    return { success: false, error: error.message };
  }
}

/**
 * 患者データ同期
 * @param {Object} params - { patient_id }
 * @returns {Object} 処理結果
 */
function handleSyncPatientData(params) {
  const patientId = params.patient_id;

  try {
    // 既存スプレッドシートからデータを取得
    const data = getPatientFullData(patientId);

    // AppSheetスプレッドシートに同期
    const result = syncSinglePatientToAppSheet(patientId, data);

    return {
      success: true,
      message: 'データを同期しました',
      data: data
    };

  } catch (error) {
    logError('handleSyncPatientData', error);
    return { success: false, error: error.message };
  }
}

/**
 * 判定計算
 * @param {Object} params - { patient_id }
 * @returns {Object} 処理結果
 */
function handleCalculateJudgments(params) {
  const patientId = params.patient_id;

  if (!patientId) {
    return { success: false, error: 'patient_id is required' };
  }

  try {
    // AppSheetからblood_testsデータを取得
    const bloodData = getBloodTestFromAppSheet(patientId);

    if (!bloodData) {
      return { success: false, error: 'Blood test data not found' };
    }

    // 判定計算
    const gender = bloodData.gender || 'M';
    const judgments = {};

    const items = [
      { key: 'hdl', code: 'HDL_CHOLESTEROL' },
      { key: 'ldl', code: 'LDL_CHOLESTEROL' },
      { key: 'tg', code: 'TRIGLYCERIDES' },
      { key: 'fbs', code: 'FASTING_GLUCOSE' },
      { key: 'hba1c', code: 'HBA1C' },
      { key: 'acr', code: 'ACR' }
    ];

    for (const item of items) {
      const value = bloodData[item.key];
      if (value !== null && value !== undefined && value !== '') {
        judgments[item.key + '_judgment'] = judge(item.code, toNumber(value), gender);
        judgments[item.key + '_hl'] = getHighLowFlag(item.code, value, gender);
      }
    }

    // AppSheetのblood_testsシートを更新
    updateBloodTestJudgments(patientId, judgments);

    return {
      success: true,
      message: '判定を計算しました',
      judgments: judgments
    };

  } catch (error) {
    logError('handleCalculateJudgments', error);
    return { success: false, error: error.message };
  }
}

/**
 * 案件一覧取得
 * @param {Object} params - { status_filter }
 * @returns {Object} 処理結果
 */
function handleGetCaseList(params) {
  try {
    const cases = getRosaiCaseList();

    // ステータスフィルター
    const statusFilter = params.status_filter;
    let filteredCases = cases;

    if (statusFilter && statusFilter !== 'ALL') {
      filteredCases = cases.filter(c => c.status === statusFilter);
    }

    return {
      success: true,
      cases: filteredCases,
      totalCount: cases.length,
      filteredCount: filteredCases.length
    };

  } catch (error) {
    logError('handleGetCaseList', error);
    return { success: false, error: error.message };
  }
}

/**
 * スケジュール取込処理（appSheetBridge.gsの関数を呼び出し）
 * @param {Object} params - { case_id, exam_date }
 * @returns {Object} 処理結果
 */
function handleImportSchedule(params) {
  const caseId = params.case_id;
  const examDate = params.exam_date;

  if (!caseId) {
    return { success: false, error: 'case_id is required' };
  }

  try {
    // appSheetBridge.gsのimportScheduleToCase関数を呼び出し
    if (typeof importScheduleToCase === 'function') {
      const result = importScheduleToCase(caseId, examDate);
      return result;
    } else {
      return { success: false, error: 'importScheduleToCase function not found' };
    }
  } catch (error) {
    logError('handleImportSchedule', error);
    return { success: false, error: error.message };
  }
}

/**
 * 二次検診対象者判定（appSheetBridge.gsの関数を呼び出し）
 * @param {Object} params - { patient_id }
 * @returns {Object} 処理結果
 */
function handleCalculateScreening(params) {
  const patientId = params.patient_id;

  if (!patientId) {
    return { success: false, error: 'patient_id is required' };
  }

  try {
    // appSheetBridge.gsのcalculateScreeningResult関数を呼び出し
    if (typeof calculateAndSaveScreeningResult === 'function') {
      const result = calculateAndSaveScreeningResult(patientId);
      return result;
    } else if (typeof calculateScreeningResult === 'function') {
      // 直接計算関数を呼び出し
      const primaryData = getPrimaryExamData(patientId);
      if (!primaryData) {
        return { success: false, error: 'Primary exam data not found' };
      }
      const result = calculateScreeningResult(primaryData);
      return {
        success: true,
        screening: result
      };
    } else {
      return { success: false, error: 'calculateScreeningResult function not found' };
    }
  } catch (error) {
    logError('handleCalculateScreening', error);
    return { success: false, error: error.message };
  }
}

// ============================================
// データ取得・変換関数
// ============================================

/**
 * 保健指導生成用の患者データを取得
 * @param {string} patientId - 患者ID
 * @returns {Object|null} 患者データ
 */
function getPatientDataForGuidance(patientId) {
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);

  // patientsシートから基本情報
  const patientsSheet = ss.getSheetByName('AS_受診者');
  const patientRow = findRowById(patientsSheet, patientId, 'A');

  if (!patientRow) return null;

  const patientData = getRowAsObject(patientsSheet, patientRow, [
    'patient_id', 'case_id', 'chart_no', 'name', 'name_kana',
    'birth_date', 'age', 'gender', 'current_step'
  ]);

  // blood_testsシートから検査値
  const bloodSheet = ss.getSheetByName('AS_血液検査');
  const bloodRow = findRowByPatientId(bloodSheet, patientId);

  if (bloodRow) {
    const bloodData = getRowAsObject(bloodSheet, bloodRow, [
      'test_id', 'patient_id', 'hdl', 'hdl_judgment', 'ldl', 'ldl_judgment',
      'tg', 'tg_judgment', 'fbs', 'fbs_judgment', 'hba1c', 'hba1c_judgment', 'acr', 'acr_judgment'
    ]);
    Object.assign(patientData, bloodData);
  }

  // ultrasoundシートから超音波所見
  const ultrasoundSheet = ss.getSheetByName('AS_超音波');
  const ultrasoundRow = findRowByPatientId(ultrasoundSheet, patientId);

  if (ultrasoundRow) {
    const ultrasoundData = getRowAsObject(ultrasoundSheet, ultrasoundRow, [
      'ultrasound_id', 'patient_id', 'cardiac_judgment', 'cardiac_findings',
      'carotid_judgment', 'carotid_findings'
    ]);
    Object.assign(patientData, ultrasoundData);
  }

  return patientData;
}

/**
 * Claude APIで保健指導文を生成
 * @param {Object} patientData - 患者データ
 * @returns {Object} 生成された指導文
 */
function generateGuidanceWithClaude(patientData) {
  const prompt = `以下の健診結果から、特定保健指導の指導内容を作成してください。

【患者情報】
- 年齢: ${patientData.age || '不明'}歳
- 性別: ${patientData.gender === 'M' ? '男性' : '女性'}

【血液検査結果】
- HDL-C: ${patientData.hdl || '-'} mg/dL (${patientData.hdl_judgment || '-'})
- LDL-C: ${patientData.ldl || '-'} mg/dL (${patientData.ldl_judgment || '-'})
- 中性脂肪: ${patientData.tg || '-'} mg/dL (${patientData.tg_judgment || '-'})
- 空腹時血糖: ${patientData.fbs || '-'} mg/dL (${patientData.fbs_judgment || '-'})
- HbA1c: ${patientData.hba1c || '-'}% (${patientData.hba1c_judgment || '-'})
${patientData.acr ? `- 尿中アルブミン/Cre比: ${patientData.acr} mg/gCr (${patientData.acr_judgment || '-'})` : ''}

【超音波検査所見】
- 心臓超音波: ${patientData.cardiac_judgment || '-'} - ${patientData.cardiac_findings || '所見なし'}
- 頸動脈超音波: ${patientData.carotid_judgment || '-'} - ${patientData.carotid_findings || '所見なし'}

【指示】
1. B判定以上の項目に対して、具体的で実行可能な指導内容を作成
2. 患者の年齢・性別を考慮した現実的な目標設定
3. 専門用語を避け、わかりやすい表現を使用

【出力形式】JSON
{
  "nutrition": "栄養指導（100〜150文字程度）",
  "exercise": "運動指導（100〜150文字程度）",
  "lifestyle": "生活習慣指導（100〜150文字程度、必要な場合のみ）",
  "priority_items": ["優先的に改善すべき項目1", "項目2"]
}`;

  try {
    const response = callClaudeApi(
      'あなたは特定保健指導の専門家です。労災二次健診の結果に基づき、具体的で実行可能な指導内容を作成します。医学的に正確で、かつ患者が実践しやすい内容を心がけてください。',
      prompt,
      { max_tokens: 1024 }
    );

    // JSONをパース
    const content = response.content || response;
    const jsonMatch = content.match(/\{[\s\S]*\}/);

    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }

    // JSONパース失敗時はテキストとして返す
    return {
      nutrition: content,
      exercise: '',
      lifestyle: '',
      priority_items: []
    };

  } catch (error) {
    logError('generateGuidanceWithClaude', error);
    throw error;
  }
}

/**
 * 保健指導をAppSheetに保存
 * @param {string} patientId - 患者ID
 * @param {Object} guidance - 指導内容
 */
function saveGuidanceToAppSheet(patientId, guidance) {
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);
  const sheet = ss.getSheetByName('AS_保健指導');

  // 既存レコードを検索
  const existingRow = findRowByPatientId(sheet, patientId);

  const now = new Date();
  const guidanceId = existingRow ? sheet.getRange(existingRow, 1).getValue() : `GUID-${patientId}-${Date.now()}`;

  const rowData = [
    guidanceId,
    patientId,
    guidance.nutrition || '',
    guidance.exercise || '',
    guidance.lifestyle || '',
    true,  // is_ai_generated
    JSON.stringify(guidance)  // ai_suggestion (元のJSON全体)
  ];

  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

// ============================================
// AppSheetデータ同期
// ============================================

/**
 * AppSheet用スプレッドシートにデータ同期
 * @param {string} caseId - 案件ID
 * @param {Array<string>} patientIds - 患者IDリスト
 * @returns {Object} 同期結果
 */
function syncToAppSheetData(caseId, patientIds) {
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);

  const results = {
    patients: 0,
    bloodTests: 0,
    ultrasound: 0,
    errors: []
  };

  for (const patientId of patientIds) {
    try {
      // 既存データから取得
      const patientData = getPatientFullData(patientId);

      if (patientData) {
        // patientsシートに追加/更新
        syncPatientToSheet(ss, patientData, caseId);
        results.patients++;

        // blood_testsシートに追加/更新
        if (patientData.bloodTest) {
          syncBloodTestToSheet(ss, patientId, patientData.bloodTest);
          results.bloodTests++;
        }

        // ultrasoundシートに初期行を追加
        initUltrasoundRow(ss, patientId);
        results.ultrasound++;
      }

    } catch (error) {
      results.errors.push(`${patientId}: ${error.message}`);
    }
  }

  return results;
}

/**
 * 患者データをpatientsシートに同期
 */
function syncPatientToSheet(ss, patientData, caseId) {
  const sheet = ss.getSheetByName('AS_受診者');
  const existingRow = findRowById(sheet, patientData.patientId, 'A');

  const now = new Date();
  const rowData = [
    patientData.patientId,
    caseId,
    patientData.chartNo || '',
    patientData.name || '',
    patientData.nameKana || '',
    patientData.birthDate || '',
    patientData.age || '',
    patientData.gender || '',
    1,  // current_step
    'NOT_STARTED',  // step_status
    now,  // created_at
    now   // updated_at
  ];

  if (existingRow) {
    // 更新（created_atは維持）
    rowData[10] = sheet.getRange(existingRow, 11).getValue();
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

/**
 * 血液検査データをblood_testsシートに同期
 */
function syncBloodTestToSheet(ss, patientId, bloodData) {
  const sheet = ss.getSheetByName('AS_血液検査');
  const existingRow = findRowByPatientId(sheet, patientId);

  const gender = bloodData.gender || 'M';

  const rowData = [
    existingRow ? sheet.getRange(existingRow, 1).getValue() : `BT-${patientId}`,
    patientId,
    bloodData.hdl || '',
    bloodData.hdl ? judge('HDL_CHOLESTEROL', toNumber(bloodData.hdl), gender) : '',
    '',  // hdl_prev
    bloodData.ldl || '',
    bloodData.ldl ? judge('LDL_CHOLESTEROL', toNumber(bloodData.ldl), gender) : '',
    '',  // ldl_prev
    bloodData.tg || '',
    bloodData.tg ? judge('TRIGLYCERIDES', toNumber(bloodData.tg), gender) : '',
    '',  // tg_prev
    bloodData.fbs || '',
    bloodData.fbs ? judge('FASTING_GLUCOSE', toNumber(bloodData.fbs), gender) : '',
    '',  // fbs_prev
    bloodData.hba1c || '',
    bloodData.hba1c ? judge('HBA1C', toNumber(bloodData.hba1c), gender) : '',
    '',  // hba1c_prev
    bloodData.acr || '',
    bloodData.acr ? judge('ACR', toNumber(bloodData.acr), gender) : ''
  ];

  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

/**
 * 超音波検査の初期行を作成
 */
function initUltrasoundRow(ss, patientId) {
  const sheet = ss.getSheetByName('AS_超音波');
  const existingRow = findRowByPatientId(sheet, patientId);

  if (!existingRow) {
    sheet.appendRow([
      `US-${patientId}`,
      patientId,
      '',  // cardiac_judgment
      '',  // cardiac_findings
      '',  // carotid_judgment
      ''   // carotid_findings
    ]);
  }
}

// ============================================
// ユーティリティ関数
// ============================================

/**
 * AppSheet用スプレッドシートIDを取得
 * @returns {string} スプレッドシートID
 */
function getAppSheetSpreadsheetId() {
  // 現在のスプレッドシートを使用
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/**
 * 指定IDで行を検索
 * @param {Sheet} sheet - シート
 * @param {string} id - 検索ID
 * @param {string} column - 検索列（A, B, ...）
 * @returns {number|null} 行番号（見つからない場合はnull）
 */
function findRowById(sheet, id, column) {
  const col = column.charCodeAt(0) - 64;  // A=1, B=2, ...
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(id).trim()) {
      return i + 2;
    }
  }

  return null;
}

/**
 * patient_idで行を検索
 * @param {Sheet} sheet - シート
 * @param {string} patientId - 患者ID
 * @returns {number|null} 行番号
 */
function findRowByPatientId(sheet, patientId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const patientIdCol = headers.indexOf('patient_id') + 1;

  if (patientIdCol === 0) {
    // patient_idカラムがない場合はB列を検索
    return findRowById(sheet, patientId, 'B');
  }

  const values = sheet.getRange(2, patientIdCol, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(patientId).trim()) {
      return i + 2;
    }
  }

  return null;
}

/**
 * 行データをオブジェクトとして取得
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号
 * @param {Array<string>} columns - 取得するカラム名
 * @returns {Object} データオブジェクト
 */
function getRowAsObject(sheet, row, columns) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const obj = {};

  for (const col of columns) {
    const index = headers.indexOf(col);
    if (index >= 0) {
      obj[col] = values[index];
    }
  }

  return obj;
}

/**
 * 患者の全データを既存スプレッドシートから取得
 * @param {string} patientId - 患者ID
 * @returns {Object|null} 患者データ
 */
function getPatientFullData(patientId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // AS_受診者シートから基本情報
    const patientSheet = ss.getSheetByName('AS_受診者');
    if (!patientSheet) return null;

    const patientRow = findRowById(patientSheet, patientId, 'A');

    if (!patientRow) return null;

    const headers = patientSheet.getRange(1, 1, 1, patientSheet.getLastColumn()).getValues()[0];
    const patientValues = patientSheet.getRange(patientRow, 1, 1, patientSheet.getLastColumn()).getValues()[0];

    const data = {
      patientId: patientValues[headers.indexOf('patient_id')] || patientValues[0],
      name: patientValues[headers.indexOf('name')] || '',
      nameKana: patientValues[headers.indexOf('name_kana')] || '',
      gender: patientValues[headers.indexOf('gender')] || '',
      birthDate: patientValues[headers.indexOf('birth_date')] || '',
      age: patientValues[headers.indexOf('age')] || '',
      chartNo: patientValues[headers.indexOf('chart_no')] || patientValues[0]
    };

    // 血液検査シートから検査値
    const bloodSheet = ss.getSheetByName('AS_血液検査');
    if (bloodSheet) {
      const bloodRow = findRowByPatientId(bloodSheet, patientId);

      if (bloodRow) {
        const bloodHeaders = bloodSheet.getRange(1, 1, 1, bloodSheet.getLastColumn()).getValues()[0];
        const bloodValues = bloodSheet.getRange(bloodRow, 1, 1, bloodSheet.getLastColumn()).getValues()[0];

        data.bloodTest = {
          hdl: bloodValues[bloodHeaders.indexOf('hdl')] || '',
          ldl: bloodValues[bloodHeaders.indexOf('ldl')] || '',
          tg: bloodValues[bloodHeaders.indexOf('tg')] || '',
          fbs: bloodValues[bloodHeaders.indexOf('fbs')] || '',
          hba1c: bloodValues[bloodHeaders.indexOf('hba1c')] || '',
          gender: data.gender === '男' ? 'M' : 'F'
        };
      }
    }

    return data;

  } catch (error) {
    logError('getPatientFullData', error);
    return null;
  }
}

/**
 * Excel出力用の患者データを収集
 * @param {string} patientId - 患者ID
 * @returns {Object|null} 出力用データ
 */
function collectPatientDataForExport(patientId) {
  // AppSheetのデータから収集
  const patientData = getPatientDataForGuidance(patientId);

  if (!patientData) return null;

  // 案件情報を取得
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);
  const casesSheet = ss.getSheetByName('AS_案件');
  const caseRow = findRowById(casesSheet, patientData.case_id, 'A');

  let caseInfo = {};
  if (caseRow) {
    const caseData = getRowAsObject(casesSheet, caseRow, [
      'case_id', 'company_name', 'exam_date', 'doctor_name'
    ]);
    caseInfo = {
      companyName: caseData.company_name,
      examDate: caseData.exam_date,
      doctorName: caseData.doctor_name
    };
  }

  // guidanceを取得
  const guidanceSheet = ss.getSheetByName('AS_保健指導');
  const guidanceRow = findRowByPatientId(guidanceSheet, patientId);

  if (guidanceRow) {
    const guidanceData = getRowAsObject(guidanceSheet, guidanceRow, [
      'nutrition_text', 'exercise_text', 'lifestyle_text'
    ]);
    patientData.guidance = guidanceData.nutrition_text;
    patientData.summaryFindings = `${guidanceData.nutrition_text}\n${guidanceData.exercise_text}`;
  }

  patientData.caseInfo = caseInfo;

  return patientData;
}

/**
 * 患者のステップを更新
 * @param {string} patientId - 患者ID
 * @param {number} step - 新しいステップ
 * @param {string} status - 新しいステータス
 */
function updatePatientStep(patientId, step, status) {
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);
  const sheet = ss.getSheetByName('AS_受診者');

  const row = findRowById(sheet, patientId, 'A');
  if (row) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const stepCol = headers.indexOf('current_step') + 1;
    const statusCol = headers.indexOf('step_status') + 1;
    const updatedCol = headers.indexOf('updated_at') + 1;

    if (stepCol > 0) sheet.getRange(row, stepCol).setValue(step);
    if (statusCol > 0) sheet.getRange(row, statusCol).setValue(status);
    if (updatedCol > 0) sheet.getRange(row, updatedCol).setValue(new Date());
  }
}

/**
 * AppSheetから血液検査データを取得
 * @param {string} patientId - 患者ID
 * @returns {Object|null} 血液検査データ
 */
function getBloodTestFromAppSheet(patientId) {
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);
  const sheet = ss.getSheetByName('AS_血液検査');

  const row = findRowByPatientId(sheet, patientId);
  if (!row) return null;

  return getRowAsObject(sheet, row, [
    'test_id', 'patient_id', 'hdl', 'ldl', 'tg', 'fbs', 'hba1c', 'acr'
  ]);
}

/**
 * 血液検査の判定を更新
 * @param {string} patientId - 患者ID
 * @param {Object} judgments - 判定データ
 */
function updateBloodTestJudgments(patientId, judgments) {
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);
  const sheet = ss.getSheetByName('AS_血液検査');

  const row = findRowByPatientId(sheet, patientId);
  if (!row) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (const [key, value] of Object.entries(judgments)) {
    const col = headers.indexOf(key) + 1;
    if (col > 0) {
      sheet.getRange(row, col).setValue(value);
    }
  }
}

/**
 * 一次検診データを取得（スクリーニング用）
 * @param {string} patientId - 患者ID
 * @returns {Object|null} 一次検診データ
 */
function getPrimaryExamData(patientId) {
  const appSheetSsId = getAppSheetSpreadsheetId();
  const ss = SpreadsheetApp.openById(appSheetSsId);
  const sheet = ss.getSheetByName('AS_受診者');

  const row = findRowById(sheet, patientId, 'A');
  if (!row) return null;

  return getRowAsObject(sheet, row, [
    'patient_id', 'gender',
    'primary_exam_date', 'primary_hdl', 'primary_ldl', 'primary_tg',
    'primary_fbs', 'primary_hba1c', 'primary_bp_sys', 'primary_bp_dia',
    'primary_bmi', 'primary_waist'
  ]);
}

// ============================================
// テスト用関数
// ============================================

/**
 * Webhookのテスト実行
 */
function testWebhook() {
  const testParams = {
    action: 'health_check'
  };

  const mockEvent = {
    postData: {
      contents: JSON.stringify(testParams)
    }
  };

  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}

/**
 * CSV取込テスト
 */
function testImportCsv() {
  const testParams = {
    action: 'import_csv',
    case_id: 'TEST-CASE-001',
    folder_id: 'YOUR_TEST_FOLDER_ID'  // テスト用フォルダIDに置き換え
  };

  const mockEvent = {
    postData: {
      contents: JSON.stringify(testParams)
    }
  };

  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}

/**
 * 保健指導生成テスト
 */
function testGenerateGuidance() {
  const testParams = {
    action: 'generate_guidance',
    patient_id: 'YOUR_TEST_PATIENT_ID'  // テスト用患者IDに置き換え
  };

  const mockEvent = {
    postData: {
      contents: JSON.stringify(testParams)
    }
  };

  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}
