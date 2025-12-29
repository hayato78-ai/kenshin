/**
 * Excel出力モジュール
 * スプレッドシートからExcelファイルを生成
 */

// ============================================
// Excel出力設定
// ============================================
const EXCEL_CONFIG = {
  // 出力ファイル名フォーマット
  FILE_NAME_FORMAT: 'result_{DATE}_{ID}.xlsx',

  // MIMEタイプ
  EXCEL_MIME: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

  // 出力用シートのセルマッピング（template.xlsmのレイアウトに対応）
  CELL_MAPPING: {
    PAGE1: {
      // 基本情報
      EXAM_DATE: 'AG5',
      PATIENT_ID: 'G8',
      GENDER: 'S9',
      NAME: 'G10',
      KANA: 'G11',
      BIRTH_DATE: 'G12',
      AGE: 'S12',

      // 身体測定
      HEIGHT: 'M28',
      WEIGHT: 'M29',
      BMI: 'M30',
      WAIST: 'M31',
      BP_SYS: 'M32',
      BP_DIA: 'Q32',

      // 検査結果セル（BMLコード→セル）
      TEST_CELLS: {
        // 脂質
        '0000460': { value: 'M35', judgment: 'Q35' },  // HDL
        '0000410': { value: 'M36', judgment: 'Q36' },  // LDL
        '0000454': { value: 'M37', judgment: 'Q37' },  // TG

        // 血液
        '0000301': { value: 'M38', judgment: 'Q38' },  // WBC
        '0000302': { value: 'M39', judgment: 'Q39' },  // RBC
        '0000303': { value: 'M40', judgment: 'Q40' },  // Hb
        '0000304': { value: 'M41', judgment: 'Q41' },  // Ht
        '0000305': { value: 'M42', judgment: 'Q42' },  // MCV
        '0000306': { value: 'M43', judgment: 'Q43' },  // MCH
        '0000307': { value: 'M44', judgment: 'Q44' },  // MCHC
        '0000308': { value: 'M45', judgment: 'Q45' },  // PLT

        // 生化学
        '0000401': { value: 'AD34', judgment: 'AH34' }, // TP
        '0000472': { value: 'AD35', judgment: 'AH35' }, // T-Bil
        '0000481': { value: 'AD36', judgment: 'AH36' }, // AST
        '0000482': { value: 'AD37', judgment: 'AH37' }, // ALT
        '0000484': { value: 'AD38', judgment: 'AH38' }, // γ-GTP
        '0000658': { value: 'AD40', judgment: 'AH40' }, // CRP
        '0000407': { value: 'AD41', judgment: 'AH41' }, // UA
        '0000503': { value: 'AD42', judgment: 'AH42' }, // FBS
        '0003317': { value: 'AD43', judgment: 'AH43' }, // HbA1c
        '0000413': { value: 'AD44', judgment: 'AH44' }, // Cr
        '0002696': { value: 'AD45', judgment: 'AH45' }  // eGFR
      }
    },
    PAGE5: {
      // 総合所見
      COMBINED_FINDINGS: 'B10',
      OVERALL_JUDGMENT: 'AH5'
    }
  }
};

// ============================================
// Excel出力関数
// ============================================

/**
 * 患者データをExcelに出力
 * @param {string} patientId - 受診ID
 * @returns {string} 出力ファイルのURL
 */
function exportToExcel(patientId) {
  logInfo(`Excel出力開始: ${patientId}`);

  try {
    // 1. 患者データを収集
    const patientData = collectPatientData(patientId);
    if (!patientData) {
      throw new Error('患者データが見つかりません: ' + patientId);
    }

    // 2. テンプレートシートにデータを転記
    const outputSpreadsheet = prepareOutputSpreadsheet(patientId, patientData);

    // 3. Excelファイルとして出力
    const file = convertToExcel(outputSpreadsheet, patientId, patientData.examDate);

    // 4. 出力日時を記録
    recordExportDate(patientId);

    logInfo(`Excel出力完了: ${file.getUrl()}`);
    return file.getUrl();

  } catch (e) {
    logError('exportToExcel', e);
    throw e;
  }
}

/**
 * 患者データを収集
 * @param {string} patientId - 受診ID
 * @returns {Object|null} 患者データ
 */
function collectPatientData(patientId) {
  // 受診者マスタから基本情報を取得
  const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
  const patientRow = findPatientRow(patientSheet, patientId, patientSheet.getLastRow());

  if (patientRow === 0) return null;

  const patientData = patientSheet.getRange(patientRow, 1, 1, 15).getValues()[0];

  // 身体測定データを取得
  const physicalSheet = getSheet(CONFIG.SHEETS.PHYSICAL);
  let physicalData = null;
  if (physicalSheet) {
    const physicalRow = findPatientRow(physicalSheet, patientId, physicalSheet.getLastRow());
    if (physicalRow > 0) {
      physicalData = physicalSheet.getRange(physicalRow, 1, 1, 23).getValues()[0];
    }
  }

  // 血液検査データを取得
  const bloodSheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
  let bloodData = null;
  if (bloodSheet) {
    const bloodRow = findPatientRow(bloodSheet, patientId, bloodSheet.getLastRow());
    if (bloodRow > 0) {
      bloodData = bloodSheet.getRange(bloodRow, 1, 1, 28).getValues()[0];
    }
  }

  // 所見データを取得
  const findings = getFindings(patientId);

  // 性別から内部コードを取得
  const genderDisplay = patientData[5];  // F列
  const gender = genderDisplay === '女' ? 'F' : 'M';

  return {
    patientId: patientData[0],
    status: patientData[1],
    examDate: patientData[2],
    name: patientData[3],
    kana: patientData[4],
    gender: genderDisplay,
    genderCode: gender,
    birthDate: patientData[6],
    age: patientData[7],
    course: patientData[8],
    company: patientData[9],
    department: patientData[10],
    overallJudgment: patientData[11],
    physical: physicalData,
    blood: bloodData,
    findings: findings
  };
}

/**
 * 出力用スプレッドシートを準備
 * @param {string} patientId - 受診ID
 * @param {Object} patientData - 患者データ
 * @returns {Spreadsheet} 出力用スプレッドシート
 */
function prepareOutputSpreadsheet(patientId, patientData) {
  const ss = getSpreadsheet();

  // 出力用シートにデータを転記
  fillPage1(ss, patientData);
  fillPage5(ss, patientData);

  return ss;
}

/**
 * 1ページ目を転記
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Object} data - 患者データ
 */
function fillPage1(ss, data) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OUTPUT_PAGE1);
  if (!sheet) {
    logInfo('出力用_1ページシートが見つかりません');
    return;
  }

  const mapping = EXCEL_CONFIG.CELL_MAPPING.PAGE1;

  // 基本情報
  if (data.examDate) {
    sheet.getRange(mapping.EXAM_DATE).setValue(formatDate(data.examDate));
  }
  sheet.getRange(mapping.PATIENT_ID).setValue(data.patientId || '');
  sheet.getRange(mapping.GENDER).setValue(data.gender || '');
  sheet.getRange(mapping.NAME).setValue(data.name || '');
  sheet.getRange(mapping.KANA).setValue(data.kana || '');
  if (data.birthDate) {
    sheet.getRange(mapping.BIRTH_DATE).setValue(formatDate(data.birthDate));
  }
  if (data.age) {
    sheet.getRange(mapping.AGE).setValue(data.age);
  }

  // 身体測定
  if (data.physical) {
    const p = data.physical;
    if (p[1]) sheet.getRange(mapping.HEIGHT).setValue(p[1]);  // 身長
    if (p[2]) sheet.getRange(mapping.WEIGHT).setValue(p[2]);  // 体重
    if (p[4]) sheet.getRange(mapping.BMI).setValue(p[4]);     // BMI
    if (p[6]) sheet.getRange(mapping.WAIST).setValue(p[6]);   // 腹囲
    if (p[7]) sheet.getRange(mapping.BP_SYS).setValue(p[7]);  // 血圧収縮期
    if (p[8]) sheet.getRange(mapping.BP_DIA).setValue(p[8]);  // 血圧拡張期
  }

  // 血液検査
  if (data.blood) {
    fillBloodTestCells(sheet, data.blood, data.genderCode, mapping.TEST_CELLS);
  }
}

/**
 * 血液検査のセルを転記
 * @param {Sheet} sheet - シート
 * @param {Array} bloodData - 血液検査データ
 * @param {string} gender - 性別（M/F）
 * @param {Object} cellMapping - セルマッピング
 */
function fillBloodTestCells(sheet, bloodData, gender, cellMapping) {
  // 列インデックス→BMLコード（bloodDataの構造に対応）
  const columnToCode = {
    1: '0000301',   // WBC
    2: '0000302',   // RBC
    3: '0000303',   // Hb
    4: '0000304',   // Ht
    5: '0000308',   // PLT
    6: '0000305',   // MCV
    7: '0000306',   // MCH
    8: '0000307',   // MCHC
    9: '0000401',   // TP
    11: '0000472',  // T-Bil
    12: '0000481',  // AST
    13: '0000482',  // ALT
    14: '0000484',  // γ-GTP
    18: '0000413',  // Cr
    19: '0002696',  // eGFR
    20: '0000407',  // UA
    22: '0000460',  // HDL
    23: '0000410',  // LDL
    24: '0000454',  // TG
    25: '0000503',  // FBS
    26: '0003317',  // HbA1c
    27: '0000658'   // CRP
  };

  for (const [colIdx, code] of Object.entries(columnToCode)) {
    const value = bloodData[parseInt(colIdx)];
    const cells = cellMapping[code];

    if (value !== '' && value !== null && cells) {
      // 値を転記
      sheet.getRange(cells.value).setValue(value);

      // 判定を計算して転記
      const judgment = judgeByCode(code, String(value), '', gender);
      if (judgment && cells.judgment) {
        sheet.getRange(cells.judgment).setValue(judgment);
      }
    }
  }
}

/**
 * 5ページ目（総合所見）を転記
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Object} data - 患者データ
 */
function fillPage5(ss, data) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OUTPUT_PAGE5);
  if (!sheet) {
    logInfo('出力用_5ページシートが見つかりません');
    return;
  }

  const mapping = EXCEL_CONFIG.CELL_MAPPING.PAGE5;

  // 総合所見
  if (data.findings && data.findings.combined) {
    sheet.getRange(mapping.COMBINED_FINDINGS).setValue(data.findings.combined);
  }

  // 総合判定
  if (data.overallJudgment) {
    sheet.getRange(mapping.OVERALL_JUDGMENT).setValue(data.overallJudgment);
  }
}

/**
 * スプレッドシートをExcelに変換
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} patientId - 受診ID
 * @param {Date} examDate - 受診日
 * @returns {File} Excelファイル
 */
function convertToExcel(ss, patientId, examDate) {
  // 一時的なスプレッドシートを作成（出力用シートのみをコピー）
  const tempSs = createTempSpreadsheet(ss, patientId);

  try {
    // ファイル名を生成
    const dateStr = formatDate(examDate || new Date(), 'YYYYMMDD');
    const fileName = EXCEL_CONFIG.FILE_NAME_FORMAT
      .replace('{DATE}', dateStr)
      .replace('{ID}', patientId);

    // Excelとしてエクスポート
    const url = `https://docs.google.com/spreadsheets/d/${tempSs.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();

    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Excelエクスポートに失敗しました: ' + response.getContentText());
    }

    const blob = response.getBlob().setName(fileName);

    // 出力フォルダに保存
    const outputFolder = getOutputFolder();
    const file = outputFolder.createFile(blob);

    logInfo(`Excelファイル作成: ${fileName}`);
    return file;

  } finally {
    // 一時スプレッドシートを削除
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  }
}

/**
 * 一時スプレッドシートを作成
 * @param {Spreadsheet} ss - 元のスプレッドシート
 * @param {string} patientId - 受診ID
 * @returns {Spreadsheet} 一時スプレッドシート
 */
function createTempSpreadsheet(ss, patientId) {
  // 新規スプレッドシートを作成
  const tempSs = SpreadsheetApp.create(`temp_export_${patientId}_${Date.now()}`);

  // 出力用シートをコピー
  const outputSheets = [
    CONFIG.SHEETS.OUTPUT_PAGE1,
    CONFIG.SHEETS.OUTPUT_PAGE2,
    CONFIG.SHEETS.OUTPUT_PAGE3,
    CONFIG.SHEETS.OUTPUT_PAGE4,
    CONFIG.SHEETS.OUTPUT_PAGE5
  ];

  for (const sheetName of outputSheets) {
    const sourceSheet = ss.getSheetByName(sheetName);
    if (sourceSheet) {
      sourceSheet.copyTo(tempSs).setName(sheetName.replace('出力用_', ''));
    }
  }

  // デフォルトの「シート1」を削除
  const defaultSheet = tempSs.getSheetByName('シート1');
  if (defaultSheet && tempSs.getSheets().length > 1) {
    tempSs.deleteSheet(defaultSheet);
  }

  return tempSs;
}

/**
 * 出力日時を記録
 * @param {string} patientId - 受診ID
 */
function recordExportDate(patientId) {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();
  const row = findPatientRow(sheet, patientId, lastRow);

  if (row > 0) {
    sheet.getRange(row, 15).setValue(new Date());  // O列: 出力日時
    sheet.getRange(row, 14).setValue(new Date());  // N列: 最終更新日時
  }
}

/**
 * 複数患者を一括出力
 * @param {Array<string>} patientIds - 受診ID配列
 * @returns {Object} 出力結果
 */
function exportMultipleToExcel(patientIds) {
  const results = {
    success: [],
    failed: []
  };

  for (const patientId of patientIds) {
    try {
      const url = exportToExcel(patientId);
      results.success.push({ patientId, url });
    } catch (e) {
      results.failed.push({ patientId, error: e.message });
    }
  }

  return results;
}

/**
 * 確認待ちステータスの患者を一括出力
 * @returns {Object} 出力結果
 */
function exportPendingPatients() {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { success: [], failed: [] };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const pendingIds = [];

  for (const row of data) {
    if (row[1] === CONFIG.STATUS.PENDING) {
      pendingIds.push(row[0]);
    }
  }

  return exportMultipleToExcel(pendingIds);
}
