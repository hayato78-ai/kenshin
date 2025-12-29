/**
 * 帳票出力モジュール（企業一覧表）
 *
 * 機能:
 * - 企業別健診結果一覧表のExcel出力
 * - テンプレート選択（標準/8名リスト/オプション用）
 * - セルマッピング定義に従った動的転記
 * - 判定色付けオプション
 *
 * 画面仕様:
 * - SCR-010: レポート・帳票出力画面
 * - SCR-010-LIST: 企業一覧表出力画面
 */

// ============================================
// 定数定義
// ============================================

const REPORT_CONFIG = {
  // テンプレート種別
  TEMPLATES: {
    TPL_STANDARD: {
      id: 'TPL_STANDARD',
      name: '健康診断結果一覧表',
      fileName: 'フォーマット健康診断結果一覧表.xlsm',
      description: '標準企業向け一覧表',
      dataStartRow: 5,
      maxRowsPerSheet: 50
    },
    TPL_8LIST: {
      id: 'TPL_8LIST',
      name: '8名リスト',
      fileName: '結果表(8名リスト＆オプション).xlsm',
      description: '定型8名単位リスト',
      dataStartRow: 5,
      maxRowsPerSheet: 8
    },
    TPL_OPTION: {
      id: 'TPL_OPTION',
      name: 'オプション用一覧表',
      fileName: 'フォーマット健康診断結果一覧表(オプション用).xlsm',
      description: 'オプション検査含む一覧表',
      dataStartRow: 5,
      maxRowsPerSheet: 50
    }
  },

  // 判定色設定
  JUDGMENT_COLORS: {
    'A': '#e8f5e9',  // 緑
    'B': '#fff8e1',  // 黄
    'C': '#fff3e0',  // 橙
    'D': '#ffebee'   // 赤
  },

  // デフォルトセルマッピング（テンプレートごとに個別設定可能）
  DEFAULT_HEADER_MAPPING: {
    company_name: 'B2',
    exam_date: 'E2',
    output_date: 'H2',
    total_count: 'K2'
  },

  // デフォルトデータ列マッピング
  DEFAULT_DATA_COLUMNS: {
    no: 'A',
    name: 'B',
    name_kana: 'C',
    birth_date: 'D',
    age: 'E',
    gender: 'F',
    // 検査項目はM_ReportMappingから読み込み
  },

  // 検査項目のデフォルト列（テンプレートによって異なる）
  DEFAULT_ITEM_COLUMNS: {
    height: 'G',
    weight: 'H',
    BMI: 'I',
    waist: 'J',
    bp_sys: 'K',
    bp_dia: 'L',
    HDL: 'M',
    LDL: 'N',
    TG: 'O',
    FBS: 'P',
    HbA1c: 'Q',
    AST: 'R',
    ALT: 'S',
    γGTP: 'T',
    Cr: 'U',
    eGFR: 'V',
    UA: 'W',
    overall_judgment: 'X'
  }
};

// ============================================
// メイン出力関数
// ============================================

/**
 * 企業一覧表を出力
 * @param {Object} options - 出力オプション
 * @returns {Object} 出力結果 {success, fileUrl, fileName, error}
 */
function exportCompanyReport(options) {
  logInfo('企業一覧表出力開始');

  try {
    // オプションのバリデーション
    if (!options.companyId) {
      throw new Error('企業IDが指定されていません');
    }

    const templateId = options.templateId || 'TPL_STANDARD';
    const template = REPORT_CONFIG.TEMPLATES[templateId];

    if (!template) {
      throw new Error('不正なテンプレートID: ' + templateId);
    }

    // 1. 対象データを取得
    const reportData = collectReportData(options);

    if (reportData.patients.length === 0) {
      return {
        success: false,
        error: '出力対象のデータがありません'
      };
    }

    // 2. セルマッピングを取得
    const mappings = getReportMappings(templateId);

    // 3. テンプレートをコピー
    const outputSpreadsheet = copyTemplate(template, options);

    // 4. データを転記
    fillReportData(outputSpreadsheet, reportData, mappings, template, options);

    // 5. Excelファイルとして出力
    const file = convertReportToExcel(outputSpreadsheet, reportData, options);

    logInfo(`企業一覧表出力完了: ${file.getName()}`);

    return {
      success: true,
      fileUrl: file.getUrl(),
      fileName: file.getName(),
      patientCount: reportData.patients.length
    };

  } catch (e) {
    logError('exportCompanyReport', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * レポート用データを収集
 * @param {Object} options - オプション（companyId, year, examType, dateFrom, dateTo, patientIds）
 * @returns {Object} レポートデータ
 */
function collectReportData(options) {
  const result = {
    company: null,
    patients: [],
    year: options.year || new Date().getFullYear(),
    dateRange: {
      from: options.dateFrom,
      to: options.dateTo
    }
  };

  // 企業情報を取得
  const companySheet = getSheet(CONFIG.SHEETS.COMPANY);
  if (companySheet) {
    const companyData = companySheet.getDataRange().getValues();
    for (let i = 1; i < companyData.length; i++) {
      if (companyData[i][0] === options.companyId) {
        result.company = {
          companyId: companyData[i][0],
          name: companyData[i][2],
          code: companyData[i][3]
        };
        break;
      }
    }
  }

  // 受診者データを取得
  const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
  if (!patientSheet) {
    return result;
  }

  const patientData = patientSheet.getDataRange().getValues();
  const headers = patientData[0];

  // 各受診者のデータを収集
  for (let i = 1; i < patientData.length; i++) {
    const row = patientData[i];
    const patientId = row[0];
    const examDate = row[2];
    const company = row[9];  // 事業所名列

    // フィルタリング
    // 特定の受診者IDが指定されている場合
    if (options.patientIds && options.patientIds.length > 0) {
      if (!options.patientIds.includes(patientId)) {
        continue;
      }
    }

    // 企業でフィルタ（企業名または企業IDで一致）
    if (options.companyId && company !== options.companyId && company !== result.company?.name) {
      // 企業名での一致チェック
      if (result.company && company !== result.company.name) {
        continue;
      }
    }

    // 日付範囲フィルタ
    if (examDate) {
      const examDateObj = new Date(examDate);
      if (options.dateFrom && examDateObj < new Date(options.dateFrom)) {
        continue;
      }
      if (options.dateTo && examDateObj > new Date(options.dateTo)) {
        continue;
      }
    }

    // ステータスフィルタ（完了のみ出力オプション）
    if (options.completedOnly && row[1] !== CONFIG.STATUS.COMPLETE) {
      continue;
    }

    // 受診者の詳細データを収集
    const patientRecord = collectPatientReportData(patientId, row);
    if (patientRecord) {
      result.patients.push(patientRecord);
    }
  }

  // 受診日でソート
  result.patients.sort((a, b) => {
    if (!a.examDate) return 1;
    if (!b.examDate) return -1;
    return new Date(a.examDate) - new Date(b.examDate);
  });

  return result;
}

/**
 * 個別受診者のレポートデータを収集
 * @param {string} patientId - 受診ID
 * @param {Array} basicData - 基本情報行
 * @returns {Object} 受診者レポートデータ
 */
function collectPatientReportData(patientId, basicData) {
  const record = {
    patientId: basicData[0],
    status: basicData[1],
    examDate: basicData[2],
    name: basicData[3],
    nameKana: basicData[4],
    gender: basicData[5],
    birthDate: basicData[6],
    age: basicData[7],
    course: basicData[8],
    company: basicData[9],
    department: basicData[10],
    overallJudgment: basicData[11],
    physical: {},
    blood: {},
    judgments: {}
  };

  // 身体測定データを取得
  const physicalSheet = getSheet(CONFIG.SHEETS.PHYSICAL);
  if (physicalSheet) {
    const physicalData = physicalSheet.getDataRange().getValues();
    for (let i = 1; i < physicalData.length; i++) {
      if (physicalData[i][0] === patientId) {
        record.physical = {
          height: physicalData[i][1],
          weight: physicalData[i][2],
          standardWeight: physicalData[i][3],
          BMI: physicalData[i][4],
          bodyFat: physicalData[i][5],
          waist: physicalData[i][6],
          bp_sys_1: physicalData[i][7],
          bp_dia_1: physicalData[i][8],
          bp_sys_2: physicalData[i][9],
          bp_dia_2: physicalData[i][10]
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
          record.blood[bloodHeaders[j]] = bloodData[i][j];
        }
        break;
      }
    }
  }

  // 判定を計算
  record.judgments = calculateJudgments(record);

  return record;
}

/**
 * 各検査項目の判定を計算
 * @param {Object} record - 受診者レコード
 * @returns {Object} 判定結果
 */
function calculateJudgments(record) {
  const gender = record.gender === '女' || record.gender === 'F' ? 'F' : 'M';
  const judgments = {};

  // 血液検査の判定
  const bloodItems = {
    'HDL-C': 'HDL_CHOLESTEROL',
    'LDL-C': 'LDL_CHOLESTEROL',
    'TG': 'TRIGLYCERIDES',
    'FBS': 'FASTING_GLUCOSE',
    'HbA1c': 'HBA1C',
    'AST': 'AST_GOT',
    'ALT': 'ALT_GPT',
    'γ-GTP': 'GAMMA_GTP',
    'Cr': 'CREATININE',
    'eGFR': 'EGFR',
    'UA': 'URIC_ACID'
  };

  for (const [key, code] of Object.entries(bloodItems)) {
    const value = record.blood[key];
    if (value !== undefined && value !== '' && value !== null) {
      try {
        judgments[key] = judge(code, toNumber(value), gender);
      } catch (e) {
        // 判定エラーは無視
      }
    }
  }

  // BMI判定
  if (record.physical.BMI) {
    try {
      judgments['BMI'] = judge('BMI', toNumber(record.physical.BMI), gender);
    } catch (e) { }
  }

  return judgments;
}

// ============================================
// セルマッピング管理
// ============================================

/**
 * テンプレートのセルマッピングを取得
 * @param {string} templateId - テンプレートID
 * @returns {Object} マッピング定義
 */
function getReportMappings(templateId) {
  const mappings = {
    headers: {},
    data: {},
    items: {}
  };

  // M_ReportMappingシートから読み込み
  const mappingSheet = getSheet(CONFIG.SHEETS.REPORT_MAPPING || 'M_ReportMapping');

  if (mappingSheet) {
    const data = mappingSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowTemplateId = row[1];
      const fieldType = row[2];
      const fieldName = row[3];
      const cellRef = row[4];
      const columnRef = row[5];
      const dataStartRow = row[6];

      if (rowTemplateId !== templateId) continue;

      switch (fieldType) {
        case 'HEADER':
          mappings.headers[fieldName] = cellRef;
          break;
        case 'DATA':
          mappings.data[fieldName] = { column: columnRef, startRow: dataStartRow };
          break;
        case 'ITEM':
          mappings.items[fieldName] = { column: columnRef, startRow: dataStartRow };
          break;
      }
    }
  }

  // マッピングが空の場合はデフォルト値を使用
  if (Object.keys(mappings.headers).length === 0) {
    mappings.headers = { ...REPORT_CONFIG.DEFAULT_HEADER_MAPPING };
  }
  if (Object.keys(mappings.data).length === 0) {
    const template = REPORT_CONFIG.TEMPLATES[templateId];
    const startRow = template?.dataStartRow || 5;
    Object.entries(REPORT_CONFIG.DEFAULT_DATA_COLUMNS).forEach(([key, col]) => {
      mappings.data[key] = { column: col, startRow: startRow };
    });
  }
  if (Object.keys(mappings.items).length === 0) {
    const template = REPORT_CONFIG.TEMPLATES[templateId];
    const startRow = template?.dataStartRow || 5;
    Object.entries(REPORT_CONFIG.DEFAULT_ITEM_COLUMNS).forEach(([key, col]) => {
      mappings.items[key] = { column: col, startRow: startRow };
    });
  }

  return mappings;
}

// ============================================
// テンプレート処理
// ============================================

/**
 * テンプレートファイルをコピー
 * @param {Object} template - テンプレート設定
 * @param {Object} options - オプション
 * @returns {Spreadsheet} コピーされたスプレッドシート
 */
function copyTemplate(template, options) {
  // テンプレートファイルIDを設定シートから取得
  const templateFileId = getSettingValue(`TEMPLATE_${template.id}`) ||
    getSettingValue('REPORT_TEMPLATE_FOLDER_ID');

  if (!templateFileId) {
    // テンプレートがない場合は新規スプレッドシートを作成
    logInfo('テンプレートが設定されていないため、新規スプレッドシートを作成します');
    return SpreadsheetApp.create(`一覧表_${options.companyId}_${Date.now()}`);
  }

  try {
    const templateFile = DriveApp.getFileById(templateFileId);
    const copyName = `一覧表_${options.companyId}_${Date.now()}`;
    const copiedFile = templateFile.makeCopy(copyName);

    return SpreadsheetApp.openById(copiedFile.getId());
  } catch (e) {
    logInfo(`テンプレートのコピーに失敗: ${e.message}。新規スプレッドシートを作成します`);
    return SpreadsheetApp.create(`一覧表_${options.companyId}_${Date.now()}`);
  }
}

/**
 * レポートデータを転記
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Object} reportData - レポートデータ
 * @param {Object} mappings - セルマッピング
 * @param {Object} template - テンプレート設定
 * @param {Object} options - オプション
 */
function fillReportData(ss, reportData, mappings, template, options) {
  const sheet = ss.getSheets()[0];

  // ヘッダー情報を転記
  fillHeaderData(sheet, reportData, mappings.headers, options);

  // データ行を転記
  fillPatientRows(sheet, reportData.patients, mappings, template, options);
}

/**
 * ヘッダー情報を転記
 * @param {Sheet} sheet - シート
 * @param {Object} reportData - レポートデータ
 * @param {Object} headerMappings - ヘッダーマッピング
 * @param {Object} options - オプション
 */
function fillHeaderData(sheet, reportData, headerMappings, options) {
  // 企業名
  if (headerMappings.company_name && reportData.company) {
    sheet.getRange(headerMappings.company_name).setValue(reportData.company.name);
  }

  // 出力日
  if (headerMappings.output_date) {
    sheet.getRange(headerMappings.output_date).setValue(formatDate(new Date()));
  }

  // 検診日（範囲）
  if (headerMappings.exam_date) {
    let dateStr = '';
    if (reportData.dateRange.from && reportData.dateRange.to) {
      dateStr = `${formatDate(reportData.dateRange.from)} ～ ${formatDate(reportData.dateRange.to)}`;
    } else if (reportData.patients.length > 0) {
      const dates = reportData.patients.map(p => p.examDate).filter(d => d);
      if (dates.length > 0) {
        const minDate = new Date(Math.min(...dates.map(d => new Date(d))));
        const maxDate = new Date(Math.max(...dates.map(d => new Date(d))));
        dateStr = minDate.getTime() === maxDate.getTime()
          ? formatDate(minDate)
          : `${formatDate(minDate)} ～ ${formatDate(maxDate)}`;
      }
    }
    sheet.getRange(headerMappings.exam_date).setValue(dateStr);
  }

  // 総人数
  if (headerMappings.total_count) {
    sheet.getRange(headerMappings.total_count).setValue(reportData.patients.length + '名');
  }
}

/**
 * 受診者データ行を転記
 * @param {Sheet} sheet - シート
 * @param {Array} patients - 受診者配列
 * @param {Object} mappings - マッピング
 * @param {Object} template - テンプレート設定
 * @param {Object} options - オプション
 */
function fillPatientRows(sheet, patients, mappings, template, options) {
  const dataStartRow = template.dataStartRow || 5;

  patients.forEach((patient, index) => {
    const rowNum = dataStartRow + index;

    // 基本情報
    if (mappings.data.no) {
      sheet.getRange(rowNum, columnToNumber(mappings.data.no.column)).setValue(index + 1);
    }
    if (mappings.data.name) {
      sheet.getRange(rowNum, columnToNumber(mappings.data.name.column)).setValue(patient.name || '');
    }
    if (mappings.data.name_kana) {
      sheet.getRange(rowNum, columnToNumber(mappings.data.name_kana.column)).setValue(patient.nameKana || '');
    }
    if (mappings.data.birth_date) {
      sheet.getRange(rowNum, columnToNumber(mappings.data.birth_date.column)).setValue(
        patient.birthDate ? formatDate(patient.birthDate) : ''
      );
    }
    if (mappings.data.age) {
      sheet.getRange(rowNum, columnToNumber(mappings.data.age.column)).setValue(patient.age || '');
    }
    if (mappings.data.gender) {
      sheet.getRange(rowNum, columnToNumber(mappings.data.gender.column)).setValue(patient.gender || '');
    }

    // 身体測定
    if (mappings.items.height && patient.physical.height) {
      sheet.getRange(rowNum, columnToNumber(mappings.items.height.column)).setValue(patient.physical.height);
    }
    if (mappings.items.weight && patient.physical.weight) {
      sheet.getRange(rowNum, columnToNumber(mappings.items.weight.column)).setValue(patient.physical.weight);
    }
    if (mappings.items.BMI && patient.physical.BMI) {
      const cell = sheet.getRange(rowNum, columnToNumber(mappings.items.BMI.column));
      cell.setValue(patient.physical.BMI);
      if (options.colorCoding && patient.judgments.BMI) {
        applyJudgmentColor(cell, patient.judgments.BMI);
      }
    }
    if (mappings.items.waist && patient.physical.waist) {
      sheet.getRange(rowNum, columnToNumber(mappings.items.waist.column)).setValue(patient.physical.waist);
    }
    if (mappings.items.bp_sys && patient.physical.bp_sys_1) {
      sheet.getRange(rowNum, columnToNumber(mappings.items.bp_sys.column)).setValue(patient.physical.bp_sys_1);
    }
    if (mappings.items.bp_dia && patient.physical.bp_dia_1) {
      sheet.getRange(rowNum, columnToNumber(mappings.items.bp_dia.column)).setValue(patient.physical.bp_dia_1);
    }

    // 血液検査
    fillBloodItems(sheet, rowNum, patient, mappings.items, options);

    // 総合判定
    if (mappings.items.overall_judgment) {
      const cell = sheet.getRange(rowNum, columnToNumber(mappings.items.overall_judgment.column));
      cell.setValue(patient.overallJudgment || '');
      if (options.colorCoding && patient.overallJudgment) {
        applyJudgmentColor(cell, patient.overallJudgment);
      }
    }
  });
}

/**
 * 血液検査項目を転記
 * @param {Sheet} sheet - シート
 * @param {number} rowNum - 行番号
 * @param {Object} patient - 受診者データ
 * @param {Object} itemMappings - 項目マッピング
 * @param {Object} options - オプション
 */
function fillBloodItems(sheet, rowNum, patient, itemMappings, options) {
  const bloodItemMap = {
    'HDL': 'HDL-C',
    'LDL': 'LDL-C',
    'TG': 'TG',
    'FBS': 'FBS',
    'HbA1c': 'HbA1c',
    'AST': 'AST',
    'ALT': 'ALT',
    'γGTP': 'γ-GTP',
    'Cr': 'Cr',
    'eGFR': 'eGFR',
    'UA': 'UA'
  };

  for (const [mappingKey, bloodKey] of Object.entries(bloodItemMap)) {
    if (itemMappings[mappingKey] && patient.blood[bloodKey] !== undefined) {
      const cell = sheet.getRange(rowNum, columnToNumber(itemMappings[mappingKey].column));
      cell.setValue(patient.blood[bloodKey]);

      if (options.colorCoding && patient.judgments[bloodKey]) {
        applyJudgmentColor(cell, patient.judgments[bloodKey]);
      }
    }
  }
}

/**
 * 判定に応じた背景色を適用
 * @param {Range} cell - セル
 * @param {string} judgment - 判定（A/B/C/D）
 */
function applyJudgmentColor(cell, judgment) {
  const color = REPORT_CONFIG.JUDGMENT_COLORS[judgment];
  if (color) {
    cell.setBackground(color);
  }
}

/**
 * 列文字を列番号に変換
 * @param {string} column - 列文字（A, B, AA等）
 * @returns {number} 列番号（1始まり）
 */
function columnToNumber(column) {
  if (!column) return 1;

  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

// ============================================
// Excel出力
// ============================================

/**
 * スプレッドシートをExcelファイルに変換
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Object} reportData - レポートデータ
 * @param {Object} options - オプション
 * @returns {File} Excelファイル
 */
function convertReportToExcel(ss, reportData, options) {
  try {
    // ファイル名を生成
    const companyName = reportData.company?.name || options.companyId;
    const dateStr = formatDate(new Date(), 'YYYYMMDD');
    const fileName = options.fileName ||
      `${companyName}_健診結果一覧_${dateStr}.xlsx`;

    // Excelとしてエクスポート
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();

    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Excelエクスポートに失敗: ' + response.getContentText());
    }

    const blob = response.getBlob().setName(fileName);

    // 出力フォルダに保存
    const outputFolder = getOutputFolder();
    const file = outputFolder.createFile(blob);

    // 一時スプレッドシートを削除
    DriveApp.getFileById(ss.getId()).setTrashed(true);

    return file;

  } catch (e) {
    // エラー時も一時ファイルは削除
    try {
      DriveApp.getFileById(ss.getId()).setTrashed(true);
    } catch (deleteError) { }
    throw e;
  }
}

// ============================================
// UI関連関数
// ============================================

/**
 * 企業一覧表出力ダイアログを表示
 */
function showCompanyReportDialog() {
  const html = HtmlService.createHtmlOutput(getCompanyReportDialogHtml())
    .setWidth(700)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, '企業一覧表出力');
}

/**
 * 企業一覧表出力ダイアログのHTML
 * @returns {string} HTML文字列
 */
function getCompanyReportDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 20px;
      margin: 0;
      line-height: 1.6;
    }
    h3 {
      margin: 0 0 15px 0;
      color: #1a73e8;
      border-bottom: 2px solid #1a73e8;
      padding-bottom: 8px;
    }
    .step {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      margin-bottom: 15px;
    }
    .step-title {
      font-weight: bold;
      margin-bottom: 10px;
      color: #333;
    }
    .form-group {
      margin-bottom: 12px;
    }
    label {
      display: block;
      margin-bottom: 5px;
      font-weight: 500;
    }
    select, input[type="text"], input[type="date"] {
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
      box-sizing: border-box;
    }
    .radio-group {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .radio-group label {
      display: flex;
      align-items: center;
      font-weight: normal;
    }
    .radio-group input {
      margin-right: 8px;
    }
    .checkbox-group {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .checkbox-group label {
      display: flex;
      align-items: center;
      font-weight: normal;
    }
    .checkbox-group input {
      margin-right: 8px;
    }
    .patient-list {
      max-height: 200px;
      overflow-y: auto;
      border: 1px solid #ddd;
      border-radius: 4px;
      padding: 10px;
      background: white;
    }
    .patient-item {
      display: flex;
      align-items: center;
      padding: 5px 0;
      border-bottom: 1px solid #eee;
    }
    .patient-item:last-child {
      border-bottom: none;
    }
    .patient-item input {
      margin-right: 10px;
    }
    .patient-status {
      margin-left: auto;
      font-size: 11px;
      padding: 2px 6px;
      border-radius: 3px;
    }
    .status-complete { background: #d4edda; color: #155724; }
    .status-input { background: #fff3cd; color: #856404; }
    .btn-container {
      text-align: right;
      margin-top: 20px;
      padding-top: 15px;
      border-top: 1px solid #eee;
    }
    .btn {
      padding: 10px 24px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
      margin-left: 10px;
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
    .btn-link {
      background: none;
      color: #1a73e8;
      text-decoration: underline;
      padding: 5px 10px;
    }
    .btn:disabled {
      background: #ccc;
      cursor: not-allowed;
    }
    .loading {
      display: none;
      text-align: center;
      padding: 20px;
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
    .error {
      color: #d93025;
      margin-top: 10px;
    }
    .success {
      color: #0f9d58;
      margin-top: 10px;
    }
    .info-text {
      font-size: 11px;
      color: #666;
      margin-top: 5px;
    }
    .count-badge {
      background: #e8f0fe;
      color: #1a73e8;
      padding: 2px 8px;
      border-radius: 10px;
      font-size: 12px;
      margin-left: 10px;
    }
  </style>
</head>
<body>
  <h3>📊 企業一覧表出力</h3>

  <div id="formContainer">
    <div class="step">
      <div class="step-title">Step 1: 出力条件設定</div>

      <div class="form-group">
        <label>テンプレート:</label>
        <div class="radio-group">
          <label><input type="radio" name="template" value="TPL_STANDARD" checked> 健康診断結果一覧表（標準）</label>
          <label><input type="radio" name="template" value="TPL_8LIST"> 8名リスト形式</label>
          <label><input type="radio" name="template" value="TPL_OPTION"> オプション検査用一覧表</label>
        </div>
      </div>

      <div class="form-group">
        <label>対象企業: <span style="color:red">*</span></label>
        <select id="companySelect" onchange="loadPatients()">
          <option value="">選択してください</option>
        </select>
      </div>

      <div class="form-group">
        <label>実施日範囲（任意）:</label>
        <div style="display: flex; gap: 10px; align-items: center;">
          <input type="date" id="dateFrom" style="flex: 1;">
          <span>～</span>
          <input type="date" id="dateTo" style="flex: 1;">
        </div>
      </div>
    </div>

    <div class="step">
      <div class="step-title">Step 2: 出力対象確認 <span id="patientCount" class="count-badge">0名</span></div>

      <div style="margin-bottom: 10px;">
        <button class="btn-link" onclick="selectAll()">全選択</button>
        <button class="btn-link" onclick="selectCompleted()">完了者のみ</button>
        <button class="btn-link" onclick="deselectAll()">選択解除</button>
      </div>

      <div class="patient-list" id="patientList">
        <div style="color: #666; text-align: center;">企業を選択してください</div>
      </div>
    </div>

    <div class="step">
      <div class="step-title">Step 3: 出力オプション</div>

      <div class="checkbox-group">
        <label><input type="checkbox" id="colorCoding" checked> 判定色付け（A緑/B黄/C橙/D赤）</label>
        <label><input type="checkbox" id="completedOnly"> 完了者のみ出力</label>
      </div>

      <div class="form-group" style="margin-top: 15px;">
        <label>ファイル名:</label>
        <input type="text" id="fileName" placeholder="自動生成されます">
        <div class="info-text">空欄の場合は「企業名_健診結果一覧_日付.xlsx」で生成</div>
      </div>
    </div>
  </div>

  <div class="loading" id="loading">
    <div class="spinner"></div>
    <div>出力中...</div>
  </div>

  <div class="error" id="errorMsg"></div>
  <div class="success" id="successMsg"></div>

  <div class="btn-container">
    <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
    <button class="btn btn-primary" id="exportBtn" onclick="startExport()" disabled>出力実行</button>
  </div>

  <script>
    let patients = [];

    // 企業リストを読み込み
    google.script.run
      .withSuccessHandler((companies) => {
        const select = document.getElementById('companySelect');
        companies.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.id;
          opt.textContent = c.name;
          select.appendChild(opt);
        });
      })
      .getCompanyListForDropdown();

    function loadPatients() {
      const companyId = document.getElementById('companySelect').value;
      if (!companyId) {
        document.getElementById('patientList').innerHTML =
          '<div style="color: #666; text-align: center;">企業を選択してください</div>';
        document.getElementById('patientCount').textContent = '0名';
        document.getElementById('exportBtn').disabled = true;
        return;
      }

      document.getElementById('patientList').innerHTML =
        '<div style="color: #666; text-align: center;">読み込み中...</div>';

      google.script.run
        .withSuccessHandler(renderPatients)
        .withFailureHandler((e) => {
          document.getElementById('patientList').innerHTML =
            '<div style="color: red;">読み込みエラー: ' + e.message + '</div>';
        })
        .getPatientsByCompany(companyId);
    }

    function renderPatients(data) {
      patients = data;
      const list = document.getElementById('patientList');

      if (patients.length === 0) {
        list.innerHTML = '<div style="color: #666; text-align: center;">該当者なし</div>';
        document.getElementById('patientCount').textContent = '0名';
        document.getElementById('exportBtn').disabled = true;
        return;
      }

      list.innerHTML = patients.map((p, i) => {
        const statusClass = p.status === '完了' ? 'status-complete' : 'status-input';
        const statusText = p.status === '完了' ? '✅ 完了' : '🔄 入力中';
        return '<div class="patient-item">' +
          '<input type="checkbox" class="patient-cb" data-index="' + i + '" checked>' +
          '<span>' + p.name + '</span>' +
          '<span style="margin-left: 10px; color: #666;">' + (p.examDate || '') + '</span>' +
          '<span class="patient-status ' + statusClass + '">' + statusText + '</span>' +
          '</div>';
      }).join('');

      document.getElementById('patientCount').textContent = patients.length + '名';
      document.getElementById('exportBtn').disabled = false;
    }

    function selectAll() {
      document.querySelectorAll('.patient-cb').forEach(cb => cb.checked = true);
    }

    function selectCompleted() {
      document.querySelectorAll('.patient-cb').forEach((cb, i) => {
        cb.checked = patients[i].status === '完了';
      });
    }

    function deselectAll() {
      document.querySelectorAll('.patient-cb').forEach(cb => cb.checked = false);
    }

    function startExport() {
      const companyId = document.getElementById('companySelect').value;
      if (!companyId) {
        showError('企業を選択してください');
        return;
      }

      const selectedIds = [];
      document.querySelectorAll('.patient-cb:checked').forEach(cb => {
        const idx = parseInt(cb.dataset.index);
        selectedIds.push(patients[idx].patientId);
      });

      if (selectedIds.length === 0) {
        showError('出力対象を選択してください');
        return;
      }

      const options = {
        companyId: companyId,
        templateId: document.querySelector('input[name="template"]:checked').value,
        dateFrom: document.getElementById('dateFrom').value,
        dateTo: document.getElementById('dateTo').value,
        colorCoding: document.getElementById('colorCoding').checked,
        completedOnly: document.getElementById('completedOnly').checked,
        fileName: document.getElementById('fileName').value,
        patientIds: selectedIds
      };

      showLoading(true);
      hideError();

      google.script.run
        .withSuccessHandler(handleExportResult)
        .withFailureHandler(handleError)
        .exportCompanyReport(options);
    }

    function handleExportResult(result) {
      showLoading(false);

      if (result.success) {
        showSuccess('出力完了: ' + result.patientCount + '名');
        window.open(result.fileUrl, '_blank');
      } else {
        showError(result.error || '出力に失敗しました');
      }
    }

    function handleError(error) {
      showLoading(false);
      showError(error.message);
    }

    function showLoading(show) {
      document.getElementById('loading').style.display = show ? 'block' : 'none';
      document.getElementById('formContainer').style.display = show ? 'none' : 'block';
    }

    function showError(msg) {
      document.getElementById('errorMsg').textContent = msg;
      document.getElementById('successMsg').textContent = '';
    }

    function showSuccess(msg) {
      document.getElementById('successMsg').textContent = msg;
      document.getElementById('errorMsg').textContent = '';
    }

    function hideError() {
      document.getElementById('errorMsg').textContent = '';
      document.getElementById('successMsg').textContent = '';
    }
  </script>
</body>
</html>
`;
}

/**
 * 企業別受診者リストを取得
 * @param {string} companyId - 企業ID
 * @returns {Array} 受診者リスト
 */
function getPatientsByCompany(companyId) {
  const result = [];

  // 企業情報を取得
  let companyName = companyId;
  const companySheet = getSheet(CONFIG.SHEETS.COMPANY);
  if (companySheet) {
    const companyData = companySheet.getDataRange().getValues();
    for (let i = 1; i < companyData.length; i++) {
      if (companyData[i][0] === companyId) {
        companyName = companyData[i][2];
        break;
      }
    }
  }

  // 受診者を検索
  const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
  if (!patientSheet) return result;

  const data = patientSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const company = row[9];

    // 企業IDまたは企業名で一致
    if (company === companyId || company === companyName) {
      result.push({
        patientId: row[0],
        status: row[1],
        examDate: row[2] ? formatDate(row[2]) : '',
        name: row[3],
        gender: row[5],
        overallJudgment: row[11]
      });
    }
  }

  return result;
}

// ============================================
// テスト関数
// ============================================

/**
 * 企業一覧表出力のテスト
 */
function testCompanyReportExport() {
  // テスト用企業ID（実際の企業IDに置き換え）
  const result = exportCompanyReport({
    companyId: 'CO00001',
    templateId: 'TPL_STANDARD',
    colorCoding: true
  });

  logInfo('テスト結果: ' + JSON.stringify(result, null, 2));
}

/**
 * レポートダイアログのテスト表示
 */
function testShowCompanyReportDialog() {
  showCompanyReportDialog();
}

// ============================================
// メニュー用ユーティリティ関数
// ============================================

/**
 * ドロップダウン用企業リストを取得
 * @returns {Array} 企業リスト [{id, name}]
 */
function getCompanyListForDropdown() {
  const result = [];
  const companySheet = getSheet(CONFIG.SHEETS.COMPANY);

  if (!companySheet) {
    return result;
  }

  const data = companySheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // 有効フラグがtrueの企業のみ
    if (row[10] !== false) {
      result.push({
        id: row[0],
        name: row[2]
      });
    }
  }

  return result;
}

/**
 * テンプレート一覧を表示
 */
function showReportTemplateList() {
  const ui = SpreadsheetApp.getUi();

  let message = '【利用可能なテンプレート】\n\n';

  for (const [id, template] of Object.entries(REPORT_CONFIG.TEMPLATES)) {
    message += `■ ${template.name}\n`;
    message += `  ID: ${id}\n`;
    message += `  ファイル: ${template.fileName}\n`;
    message += `  データ開始行: ${template.dataStartRow}\n`;
    message += `  最大行数/シート: ${template.maxRowsPerSheet}\n\n`;
  }

  message += '\n【設定方法】\n';
  message += '設定シートにTEMPLATE_TPL_STANDARDなどの行を追加し、\n';
  message += 'テンプレートファイルのIDを設定してください。';

  ui.alert('レポートテンプレート一覧', message, ui.ButtonSet.OK);
}

/**
 * マッピング設定一覧を表示
 */
function showReportMappingList() {
  const ui = SpreadsheetApp.getUi();

  const mappingSheet = getSheet(CONFIG.SHEETS.REPORT_MAPPING || 'M_ReportMapping');

  if (!mappingSheet) {
    ui.alert('情報', 'M_ReportMappingシートが見つかりません。\n「初期セットアップ」を実行してシートを作成してください。', ui.ButtonSet.OK);
    return;
  }

  const data = mappingSheet.getDataRange().getValues();

  if (data.length <= 1) {
    ui.alert('情報', 'マッピング設定がありません。\nシートにマッピング定義を追加してください。', ui.ButtonSet.OK);
    return;
  }

  // テンプレート別に集計
  const templateStats = {};

  for (let i = 1; i < data.length; i++) {
    const templateId = data[i][1];
    if (!templateStats[templateId]) {
      templateStats[templateId] = 0;
    }
    templateStats[templateId]++;
  }

  let message = '【マッピング設定状況】\n\n';

  for (const [templateId, count] of Object.entries(templateStats)) {
    const templateName = REPORT_CONFIG.TEMPLATES[templateId]?.name || templateId;
    message += `■ ${templateName}\n`;
    message += `  マッピング数: ${count}件\n\n`;
  }

  message += `\n総マッピング数: ${data.length - 1}件\n`;
  message += '\n【編集方法】\n';
  message += 'M_ReportMappingシートを直接編集してください。';

  ui.alert('レポートマッピング設定', message, ui.ButtonSet.OK);
}

// ============================================
// 個人レポート出力機能（1221_template_new_default.xlsm用）
// ============================================

/**
 * 個人レポート出力設定
 */
const INDIVIDUAL_REPORT_CONFIG = {
  templateFileId: null,  // 設定シートから取得
  templateFileName: '1221_template_new_default.xlsm',
  outputFolderName: '個人レポート出力',
  patientInfoSheet: '1ページ',
  testResultSheet: '４ページ'
};

/**
 * 個人レポートを出力
 * @param {string} patientId - 受診者ID
 * @param {Object} options - オプション {templateId, includeJudgment}
 * @returns {Object} 結果 {success, fileUrl, fileName, error}
 */
function exportIndividualReport(patientId, options) {
  options = options || {};
  logInfo(`個人レポート出力開始: ${patientId}`);

  try {
    // 1. 受診者情報を取得
    const patientInfo = getPatientInfoForReport(patientId);
    if (!patientInfo) {
      return { success: false, patientId: patientId, error: '受診者情報が見つかりません' };
    }

    // 2. 検査結果を取得
    const testResults = getTestResultsForReport(patientId);
    logInfo(`検査結果取得: ${Object.keys(testResults).length}件`);

    // 3. Python処理用にJSONリクエストを作成してpendingフォルダに出力
    const requestId = `REQ_${patientId}_${Date.now()}`;
    const requestData = {
      request_id: requestId,
      exam_type: 'HUMAN_DOCK',
      patient_id: patientId,
      patient_info: patientInfo,
      test_results: testResults,
      options: options,
      created_at: new Date().toISOString()
    };

    // pendingフォルダにJSONを出力（Python監視で処理）
    const pendingFolder = getPythonPendingFolder();
    if (!pendingFolder) {
      return { success: false, patientId: patientId, error: 'Python連携フォルダの作成に失敗しました。Drive権限を確認してください。' };
    }

    const fileName = `${requestId}.json`;
    const file = pendingFolder.createFile(fileName, JSON.stringify(requestData, null, 2), 'application/json');

    logInfo(`Python処理リクエスト作成: ${fileName}`);

    return {
      success: true,
      pending: true,  // Python処理待ちフラグ
      requestId: requestId,
      message: 'Python処理をリクエストしました。「状態確認」ボタンで完了を確認してください。',
      patientId: patientId,
      patientName: patientInfo.name
    };

  } catch (e) {
    logError('exportIndividualReport', e);
    return {
      success: false,
      patientId: patientId,
      error: e.message
    };
  }
}

/**
 * Python連携用pendingフォルダを取得（なければ自動作成）
 * @returns {Folder|null} pendingフォルダ
 */
function getPythonPendingFolder() {
  try {
    // 設定シートからフォルダIDを取得
    const ss = getPortalSpreadsheet();
    const settingsSheet = ss.getSheetByName('設定');
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'PYTHON_PENDING_FOLDER_ID') {
          const folderId = data[i][1];
          if (folderId) {
            try {
              return DriveApp.getFolderById(folderId);
            } catch (folderError) {
              logError('getPythonPendingFolder', 'フォルダID無効: ' + folderId);
              // フォルダが見つからない場合は後続の処理へ
            }
          }
        }
      }
    }

    // フォールバック: スプレッドシートと同じフォルダ内のpendingフォルダを探す
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    const pendingFolders = parentFolder.getFoldersByName('pending');
    if (pendingFolders.hasNext()) {
      return pendingFolders.next();
    }

    // pendingフォルダがなければ自動作成
    logInfo('pendingフォルダを自動作成します');
    const newPendingFolder = parentFolder.createFolder('pending');

    // 設定シートにフォルダIDを記録（次回以降高速化）
    if (settingsSheet) {
      const lastRow = settingsSheet.getLastRow();
      settingsSheet.getRange(lastRow + 1, 1, 1, 2).setValues([
        ['PYTHON_PENDING_FOLDER_ID', newPendingFolder.getId()]
      ]);
      logInfo('設定シートにPYTHON_PENDING_FOLDER_IDを追加: ' + newPendingFolder.getId());
    }

    return newPendingFolder;
  } catch (e) {
    logError('getPythonPendingFolder', e);
    return null;
  }
}

/**
 * レポート用受診者情報を取得
 * @param {string} patientId - 受診者ID
 * @returns {Object|null} 受診者情報
 */
function getPatientInfoForReport(patientId) {
  // まずportalApiを試す
  if (typeof portalGetPatient === 'function') {
    const result = portalGetPatient(patientId);
    // portalGetPatientは {success: true, data: {...}} を返す
    if (result && result.success && result.data) {
      const d = result.data;
      logInfo(`getPatientInfoForReport: portalGetPatientから取得成功 - ${d['氏名'] || d['受診者ID']}`);
      return {
        patientId: d['受診者ID'] || patientId,
        name: d['氏名'] || '',
        nameKana: d['カナ'] || '',
        gender: d['性別'] || '',
        birthDate: d['生年月日'] || '',
        examDate: d['受診日'] || '',
        course: d['受診コース'] || '',
        company: d['所属企業'] || '',
        bmlPatientId: d['BML患者ID'] || ''
      };
    } else {
      logInfo(`getPatientInfoForReport: portalGetPatient失敗 - ${result ? result.error : '結果なし'}`);
    }
  }

  // 次にpatientManagerを試す
  if (typeof getPatientDetail === 'function') {
    const result = getPatientDetail(patientId);
    if (result) {
      return {
        patientId: result.patientId,
        name: result.name,
        nameKana: result.kana,
        gender: result.gender,
        birthDate: result.birthDate,
        examDate: result.examDate,
        course: result.course,
        company: result.company,
        bmlPatientId: result.bmlPatientId
      };
    }
  }

  // 直接シートから取得
  const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
  if (!patientSheet) return null;

  const data = patientSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === patientId) {
      return {
        patientId: data[i][0],
        name: data[i][3],
        nameKana: data[i][4],
        gender: data[i][5],
        birthDate: data[i][6],
        examDate: data[i][2],
        course: data[i][8],
        company: data[i][9],
        bmlPatientId: data[i][15] || ''
      };
    }
  }

  return null;
}

/**
 * レポート用検査結果を取得（BMLコードをキーとした辞書）
 * @param {string} patientId - 受診者ID
 * @returns {Object} 検査結果 {bmlCode: {value, flag, judgment}}
 */
function getTestResultsForReport(patientId) {
  const results = {};

  // BMLコードマッピング（項目名→BMLコード）- 検査結果シートのヘッダー名に対応
  const itemToBmlCode = {
    // 血液一般
    'WBC': '0000301', '白血球数': '0000301',
    'RBC': '0000302', '赤血球数': '0000302',
    'Hb': '0000303', 'ヘモグロビン': '0000303',
    'Ht': '0000304', 'ヘマトクリット': '0000304',
    'PLT': '0000308', '血小板数': '0000308',
    'MCV': '0000305', 'MCH': '0000306', 'MCHC': '0000307',
    // 生化学
    'TP': '0000401', '総蛋白': '0000401',
    'ALB': '0000417', 'アルブミン': '0000417',
    'AST': '0000481', 'GOT': '0000481',
    'ALT': '0000482', 'GPT': '0000482',
    'γ-GTP': '0000484', 'γGTP': '0000484', 'γ-GT': '0000484',
    'ALP': '0013067',
    'LDH': '0000497',
    'T-Bil': '0000472', '総ビリルビン': '0000472',
    'TC': '0000453', '総コレステロール': '0000453',
    'TG': '0000454', '中性脂肪': '0000454',
    'HDL': '0000460', 'HDL-C': '0000460', 'HDLコレステロール': '0000460',
    'LDL': '0000410', 'LDL-C': '0000410', 'LDLコレステロール': '0000410',
    'FBS': '0000503', '空腹時血糖': '0000503', '血糖': '0000503',
    'HbA1c': '0003317', 'ヘモグロビンA1c': '0003317',
    'Cre': '0000413', 'クレアチニン': '0000413', 'Cr': '0000413',
    'BUN': '0000491', '尿素窒素': '0000491',
    'eGFR': '0002696',
    'UA': '0000407', '尿酸': '0000407',
    'CK': '0003845', 'CPK': '0003845',
    'Na': '0003550', 'ナトリウム': '0003550',
    'K': '0000421', 'カリウム': '0000421',
    'Cl': '0000425', 'クロール': '0000425',
    'CRP': '0000658'
  };

  // まず患者情報を取得してカルテNoを得る
  let karteNo = null;
  if (typeof portalGetPatient === 'function') {
    const patientResult = portalGetPatient(patientId);
    if (patientResult && patientResult.success && patientResult.data) {
      karteNo = patientResult.data['カルテNo'];
      logInfo(`getTestResultsForReport: カルテNo=${karteNo}`);
    }
  }

  // 検査結果シートから取得（横持ち形式 - 各列が検査項目）
  const ss = getPortalSpreadsheet();
  const resultSheet = ss.getSheetByName('検査結果');

  if (resultSheet) {
    const data = resultSheet.getDataRange().getValues();
    const headers = data[0];
    const karteNoColIdx = headers.indexOf('カルテNo');
    const patientIdColIdx = headers.indexOf('受診者ID') >= 0 ? headers.indexOf('受診者ID') : headers.indexOf('患者ID');

    logInfo(`getTestResultsForReport: headers=${headers.slice(0, 10).join(',')}`);
    logInfo(`getTestResultsForReport: karteNoColIdx=${karteNoColIdx}, patientIdColIdx=${patientIdColIdx}`);

    // カルテNoまたは受診者IDで行を検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let isMatch = false;

      // カルテNoで検索
      if (karteNo && karteNoColIdx >= 0 && String(row[karteNoColIdx]).trim() === String(karteNo).trim()) {
        isMatch = true;
      }
      // 受診者IDでも検索
      if (!isMatch && patientIdColIdx >= 0 && String(row[patientIdColIdx]).trim() === String(patientId).trim()) {
        isMatch = true;
      }

      if (isMatch) {
        logInfo(`getTestResultsForReport: 検査結果行を発見 (row ${i + 1})`);

        // 各カラムをBMLコードにマッピング
        for (let j = 0; j < headers.length; j++) {
          const header = String(headers[j]).trim();
          const value = row[j];

          // 値がある場合のみ処理
          if (value !== '' && value !== null && value !== undefined) {
            // ヘッダーがBMLコードそのものの場合
            if (/^\d{7}$/.test(header)) {
              results[header] = {
                value: value,
                flag: '',
                judgment: ''
              };
            }
            // 項目名からBMLコードに変換
            else if (itemToBmlCode[header]) {
              results[itemToBmlCode[header]] = {
                value: value,
                flag: '',
                judgment: ''
              };
            }
          }
        }

        logInfo(`getTestResultsForReport: ${Object.keys(results).length}件の検査結果を取得`);
        break;
      }
    }
  } else {
    logInfo('getTestResultsForReport: 検査結果シートが見つかりません');
  }

  return results;
}

/**
 * 個人レポート用テンプレートをコピー
 * @param {Object} patientInfo - 受診者情報
 * @param {Object} options - オプション
 * @returns {Spreadsheet|null} コピーされたスプレッドシート
 */
function copyIndividualTemplate(patientInfo, options) {
  // テンプレートファイルIDを設定から取得
  let templateFileId = getSettingValue('TEMPLATE_INDIVIDUAL_1221') ||
                       getSettingValue('INDIVIDUAL_REPORT_TEMPLATE_ID');

  if (!templateFileId) {
    // テンプレートフォルダから探す
    const folderId = getSettingValue('TEMPLATE_FOLDER_ID');
    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        const files = folder.getFilesByName(INDIVIDUAL_REPORT_CONFIG.templateFileName);
        if (files.hasNext()) {
          templateFileId = files.next().getId();
        }
      } catch (e) {
        logInfo('テンプレートフォルダからの検索に失敗: ' + e.message);
      }
    }
  }

  if (!templateFileId) {
    logInfo('個人レポートテンプレートが設定されていません。新規スプレッドシートを作成します');
    const examDate = patientInfo.examDate ? formatDate(patientInfo.examDate, 'YYYYMMDD') : formatDate(new Date(), 'YYYYMMDD');
    return SpreadsheetApp.create(`個人レポート_${patientInfo.name}_${examDate}`);
  }

  try {
    const templateFile = DriveApp.getFileById(templateFileId);
    const examDate = patientInfo.examDate ? formatDate(patientInfo.examDate, 'YYYYMMDD') : formatDate(new Date(), 'YYYYMMDD');
    const copyName = `個人レポート_${patientInfo.name}_${examDate}`;
    const copiedFile = templateFile.makeCopy(copyName);

    return SpreadsheetApp.openById(copiedFile.getId());
  } catch (e) {
    logInfo(`テンプレートのコピーに失敗: ${e.message}`);
    return null;
  }
}

/**
 * 個人レポートにデータを転記
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Object} patientInfo - 受診者情報
 * @param {Object} testResults - 検査結果
 * @param {Object} mapping - マッピング定義
 * @param {Object} options - オプション
 */
function fillIndividualReportData(ss, patientInfo, testResults, mapping, options) {
  // 1ページ目（患者基本情報）
  const page1 = ss.getSheetByName(INDIVIDUAL_REPORT_CONFIG.patientInfoSheet);
  if (page1 && mapping.patientInfo) {
    if (mapping.patientInfo.name) {
      page1.getRange(mapping.patientInfo.name.cell).setValue(patientInfo.name || '');
    }
    if (mapping.patientInfo.examDate) {
      const examDateValue = patientInfo.examDate ? formatDate(patientInfo.examDate) : '';
      page1.getRange(mapping.patientInfo.examDate.cell).setValue(examDateValue);
    }
  }

  // 4ページ目（検査結果）
  const page4 = ss.getSheetByName(INDIVIDUAL_REPORT_CONFIG.testResultSheet);
  if (page4 && mapping.testItems) {
    for (const [bmlCode, cellMapping] of Object.entries(mapping.testItems)) {
      const result = testResults[bmlCode];
      if (!result) continue;

      // 値を転記
      if (cellMapping.value) {
        page4.getRange(cellMapping.value).setValue(result.value);
      }

      // 判定を転記
      if (cellMapping.judgment && result.judgment) {
        page4.getRange(cellMapping.judgment).setValue(result.judgment);
      }

      // フラグを転記
      if (cellMapping.flag && result.flag) {
        page4.getRange(cellMapping.flag).setValue(result.flag);
      }
    }
  }

  logInfo(`データ転記完了: 検査項目${Object.keys(testResults).length}件`);
}

// convertIndividualReportToExcel は削除済み（Python方式に移行）

/**
 * 複数受診者の個人レポートを一括出力
 * @param {Array} patientIds - 受診者IDリスト
 * @param {Object} options - オプション
 * @returns {Object} 結果 {success, files, errors}
 */
function exportMultipleIndividualReports(patientIds, options) {
  options = options || {};
  const results = {
    success: true,
    files: [],
    errors: []
  };

  for (const patientId of patientIds) {
    const result = exportIndividualReport(patientId, options);
    if (result.success) {
      results.files.push({
        patientId: patientId,
        fileName: result.fileName,
        fileUrl: result.fileUrl
      });
    } else {
      results.errors.push({
        patientId: patientId,
        error: result.error
      });
      results.success = false;
    }
  }

  return results;
}

/**
 * ポータルから個人レポート出力（UI用）
 * @param {string} patientId - 受診者ID
 * @returns {Object} 結果
 */
function portalExportIndividualReport(patientId) {
  return exportIndividualReport(patientId, {
    templateId: 'TPL_INDIVIDUAL_1221',
    includeJudgment: true
  });
}

/**
 * ポータルから出力ステータスを確認（UI用）
 * フォルダIDを設定シートから直接取得して確実に参照
 * @param {string} requestId - リクエストID
 * @returns {Object} ステータス {status, fileUrl, error}
 */
function portalCheckExportStatus(requestId) {
  try {
    logInfo(`ステータス確認: ${requestId}`);

    // 設定シートからフォルダIDを直接取得（親フォルダ経由の検索をやめる）
    const processedFolderId = getSettingValue('PYTHON_PROCESSED_FOLDER_ID');
    const outputFolderId = getSettingValue('PYTHON_OUTPUT_FOLDER_ID');
    const pendingFolderId = getSettingValue('PYTHON_PENDING_FOLDER_ID');

    // 1. processedフォルダを確認（Pythonが結果を保存する場所）
    if (processedFolderId) {
      try {
        const processedFolder = DriveApp.getFolderById(processedFolderId);
        const resultFiles = processedFolder.getFilesByName(`${requestId}_result.json`);

        if (resultFiles.hasNext()) {
          const file = resultFiles.next();
          const content = file.getBlob().getDataAsString();
          const resultData = JSON.parse(content);

          logInfo(`ステータス結果: ${JSON.stringify(resultData).substring(0, 200)}`);

          if (resultData.status === 'completed' && resultData.result) {
            const outputPath = resultData.result.output_path;
            if (outputPath) {
              const fileName = outputPath.split('/').pop();
              logInfo(`出力ファイル: ${fileName}`);

              // outputフォルダを確認
              if (outputFolderId) {
                try {
                  const outputFolder = DriveApp.getFolderById(outputFolderId);
                  const excelFiles = outputFolder.getFilesByName(fileName);
                  if (excelFiles.hasNext()) {
                    const excelFile = excelFiles.next();
                    return {
                      status: 'completed',
                      fileUrl: excelFile.getUrl(),
                      fileName: excelFile.getName()
                    };
                  }
                  logInfo(`outputフォルダにファイル未発見: ${fileName}（同期待ちの可能性）`);
                } catch (outputError) {
                  logInfo(`outputフォルダアクセスエラー: ${outputError.message}`);
                }
              }

              // フォールバック: ファイル名で全体検索
              try {
                const searchResults = DriveApp.searchFiles(`title = "${fileName}"`);
                if (searchResults.hasNext()) {
                  const foundFile = searchResults.next();
                  return {
                    status: 'completed',
                    fileUrl: foundFile.getUrl(),
                    fileName: foundFile.getName()
                  };
                }
              } catch (searchError) {
                logInfo(`全体検索エラー: ${searchError.message}`);
              }

              // ファイルが見つからない場合（同期待ち）
              return {
                status: 'completed',
                fileUrl: null,
                filePath: outputPath,
                fileName: fileName,
                message: 'ファイル出力完了。Google Drive同期後にダウンロード可能になります。'
              };
            }
          } else if (resultData.status === 'error') {
            return {
              status: 'error',
              error: resultData.error || '処理中にエラーが発生しました'
            };
          }
        }
      } catch (processedError) {
        logInfo(`processedフォルダアクセスエラー: ${processedError.message}`);
      }
    }

    // 2. pendingフォルダにまだある場合は処理待ち
    if (pendingFolderId) {
      try {
        const pendingFolder = DriveApp.getFolderById(pendingFolderId);
        const pendingFiles = pendingFolder.getFilesByName(`${requestId}.json`);
        if (pendingFiles.hasNext()) {
          return { status: 'pending', message: 'Python処理待ち' };
        }
      } catch (pendingError) {
        logInfo(`pendingフォルダアクセスエラー: ${pendingError.message}`);
      }
    }

    // 3. どこにもない場合
    return { status: 'unknown', message: 'リクエストが見つかりません。処理完了済みか、フォルダID設定を確認してください。' };

  } catch (e) {
    logError('portalCheckExportStatus', e);
    return { status: 'error', error: e.message };
  }
}
