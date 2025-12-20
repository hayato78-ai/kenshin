/**
 * templateManager.gs - テンプレート管理機能
 *
 * @description 健診種別ごとのExcelテンプレート管理
 * @version 1.0.0
 * @date 2025-12-18
 */

// ============================================
// 定数定義
// ============================================

const TEMPLATE_CONFIG = {
  // テンプレート管理シート名
  SHEET_NAME: 'テンプレート管理',

  // テンプレート保存フォルダ名（スプレッドシートと同階層に作成）
  FOLDER_NAME: 'templates',

  // 一時ファイル保存フォルダ名
  TEMP_FOLDER_NAME: 'temp_exports',

  // 一時ファイル有効期限（分）
  TEMP_FILE_EXPIRY_MINUTES: 30,

  // シートヘッダー
  HEADERS: ['テンプレートID', '名前', '検診種別', 'ファイルID', 'バージョン', '更新日時', '更新者', '備考'],

  // 検診種別
  EXAM_TYPES: {
    DOCK_STANDARD: '人間ドック（標準）',
    DOCK_PREMIUM: '人間ドック（プレミアム）',
    PERIODIC: '定期健診',
    ROSAI_SECONDARY: '労災二次検診',
    SPECIFIED: '特定健診'
  }
};

// ============================================
// 初期化・セットアップ
// ============================================

/**
 * テンプレート管理の初期セットアップ
 * @returns {Object} セットアップ結果
 */
function setupTemplateManagement() {
  try {
    const ss = getSpreadsheet();

    // テンプレート管理シートを作成
    let sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(TEMPLATE_CONFIG.SHEET_NAME);
      sheet.getRange(1, 1, 1, TEMPLATE_CONFIG.HEADERS.length).setValues([TEMPLATE_CONFIG.HEADERS]);
      sheet.getRange(1, 1, 1, TEMPLATE_CONFIG.HEADERS.length)
        .setBackground('#4a86e8')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // テンプレートフォルダを作成
    const folder = getOrCreateTemplateFolder();

    // 一時フォルダを作成
    const tempFolder = getOrCreateTempFolder();

    return {
      success: true,
      sheetName: TEMPLATE_CONFIG.SHEET_NAME,
      folderId: folder.getId(),
      tempFolderId: tempFolder.getId()
    };

  } catch (e) {
    logError('setupTemplateManagement', e);
    return { success: false, error: e.message };
  }
}

/**
 * テンプレートフォルダを取得または作成
 * @returns {Folder} フォルダ
 */
function getOrCreateTemplateFolder() {
  const ss = getSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolder = ssFile.getParents().next();

  const folders = parentFolder.getFoldersByName(TEMPLATE_CONFIG.FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(TEMPLATE_CONFIG.FOLDER_NAME);
}

/**
 * 一時フォルダを取得または作成
 * @returns {Folder} フォルダ
 */
function getOrCreateTempFolder() {
  const templateFolder = getOrCreateTemplateFolder();

  const folders = templateFolder.getFoldersByName(TEMPLATE_CONFIG.TEMP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }

  return templateFolder.createFolder(TEMPLATE_CONFIG.TEMP_FOLDER_NAME);
}

// ============================================
// テンプレート登録・更新
// ============================================

/**
 * テンプレートを登録（UIから呼び出し）
 * @param {string} name - テンプレート名
 * @param {string} examType - 検診種別
 * @param {string} base64Data - Base64エンコードされたファイルデータ
 * @param {string} fileName - ファイル名
 * @param {string} memo - 備考
 * @returns {Object} 登録結果
 */
function registerTemplate(name, examType, base64Data, fileName, memo) {
  try {
    // Base64デコード
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'application/vnd.ms-excel.sheet.macroEnabled.12',
      fileName
    );

    // テンプレートフォルダに保存
    const folder = getOrCreateTemplateFolder();
    const file = folder.createFile(blob);

    // テンプレートIDを生成
    const templateId = 'TPL_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');

    // 管理シートに登録
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);
    if (!sheet) {
      setupTemplateManagement();
      sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);
    }

    const newRow = [
      templateId,
      name,
      examType,
      file.getId(),
      '1.0',
      new Date(),
      Session.getActiveUser().getEmail(),
      memo || ''
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      templateId: templateId,
      fileId: file.getId(),
      fileName: file.getName()
    };

  } catch (e) {
    logError('registerTemplate', e);
    return { success: false, error: e.message };
  }
}

/**
 * 既存ファイルIDからテンプレートを登録
 * @param {string} name - テンプレート名
 * @param {string} examType - 検診種別
 * @param {string} sourceFileId - 元ファイルID
 * @param {string} memo - 備考
 * @returns {Object} 登録結果
 */
function registerTemplateFromFileId(name, examType, sourceFileId, memo) {
  try {
    // 元ファイルをコピー
    const sourceFile = DriveApp.getFileById(sourceFileId);
    const folder = getOrCreateTemplateFolder();
    const copiedFile = sourceFile.makeCopy(sourceFile.getName(), folder);

    // テンプレートIDを生成
    const templateId = 'TPL_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');

    // 管理シートに登録
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);
    if (!sheet) {
      setupTemplateManagement();
      sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);
    }

    const newRow = [
      templateId,
      name,
      examType,
      copiedFile.getId(),
      '1.0',
      new Date(),
      Session.getActiveUser().getEmail(),
      memo || ''
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      templateId: templateId,
      fileId: copiedFile.getId(),
      fileName: copiedFile.getName()
    };

  } catch (e) {
    logError('registerTemplateFromFileId', e);
    return { success: false, error: e.message };
  }
}

/**
 * テンプレートを更新
 * @param {string} templateId - テンプレートID
 * @param {string} base64Data - Base64エンコードされたファイルデータ
 * @param {string} fileName - ファイル名
 * @returns {Object} 更新結果
 */
function updateTemplate(templateId, base64Data, fileName) {
  try {
    const template = getTemplateById(templateId);
    if (!template) {
      return { success: false, error: 'テンプレートが見つかりません' };
    }

    // 古いファイルを削除
    try {
      DriveApp.getFileById(template.fileId).setTrashed(true);
    } catch (e) {
      // ファイルが既に存在しない場合は無視
    }

    // 新しいファイルを作成
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'application/vnd.ms-excel.sheet.macroEnabled.12',
      fileName
    );

    const folder = getOrCreateTemplateFolder();
    const file = folder.createFile(blob);

    // バージョンをインクリメント
    const currentVersion = parseFloat(template.version) || 1.0;
    const newVersion = (currentVersion + 0.1).toFixed(1);

    // 管理シートを更新
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === templateId) {
        sheet.getRange(i + 1, 4).setValue(file.getId()); // ファイルID
        sheet.getRange(i + 1, 5).setValue(newVersion);    // バージョン
        sheet.getRange(i + 1, 6).setValue(new Date());    // 更新日時
        sheet.getRange(i + 1, 7).setValue(Session.getActiveUser().getEmail()); // 更新者
        break;
      }
    }

    return {
      success: true,
      templateId: templateId,
      fileId: file.getId(),
      version: newVersion
    };

  } catch (e) {
    logError('updateTemplate', e);
    return { success: false, error: e.message };
  }
}

// ============================================
// テンプレート取得
// ============================================

/**
 * テンプレート一覧を取得
 * @param {string} examType - 検診種別でフィルタ（オプション）
 * @returns {Array} テンプレート一覧
 */
function getTemplateList(examType) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const templates = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // 空行スキップ

      // 検診種別フィルタ
      if (examType && row[2] !== examType) continue;

      templates.push({
        templateId: row[0],
        name: row[1],
        examType: row[2],
        fileId: row[3],
        version: row[4],
        updatedAt: row[5],
        updatedBy: row[6],
        memo: row[7]
      });
    }

    return templates;

  } catch (e) {
    logError('getTemplateList', e);
    return [];
  }
}

/**
 * テンプレートIDで取得
 * @param {string} templateId - テンプレートID
 * @returns {Object|null} テンプレート情報
 */
function getTemplateById(templateId) {
  const templates = getTemplateList();
  return templates.find(t => t.templateId === templateId) || null;
}

/**
 * 検診種別のデフォルトテンプレートを取得
 * @param {string} examType - 検診種別
 * @returns {Object|null} テンプレート情報
 */
function getDefaultTemplate(examType) {
  const templates = getTemplateList(examType);
  return templates.length > 0 ? templates[0] : null;
}

/**
 * 検診種別一覧を取得
 * @returns {Array} 検診種別一覧
 */
function getExamTypeList() {
  return Object.entries(TEMPLATE_CONFIG.EXAM_TYPES).map(([key, value]) => ({
    id: key,
    name: value
  }));
}

// ============================================
// Excel出力
// ============================================

/**
 * 検査結果をExcelで出力
 * @param {string} visitId - 受診ID
 * @param {string} templateId - テンプレートID（省略時はデフォルト）
 * @param {string} outputFileName - 出力ファイル名（省略時は自動生成）
 * @returns {Object} 出力結果（ダウンロードURL含む）
 */
function exportToExcelWithTemplate(visitId, templateId, outputFileName) {
  try {
    // 受診情報を取得
    const visit = getVisitRecordById(visitId);
    if (!visit) {
      return { success: false, error: '受診情報が見つかりません' };
    }

    // 患者情報を取得
    const patient = getPatientById(visit.patientId);
    if (!patient) {
      return { success: false, error: '患者情報が見つかりません' };
    }

    // テンプレートを取得
    let template;
    if (templateId) {
      template = getTemplateById(templateId);
    } else {
      template = getDefaultTemplate(visit.examType || TEMPLATE_CONFIG.EXAM_TYPES.DOCK_STANDARD);
    }

    if (!template) {
      return { success: false, error: 'テンプレートが見つかりません。先にテンプレートを登録してください。' };
    }

    // 検査結果を取得
    const testResults = getTestResultsByVisitId(visitId);

    // テンプレートをコピー
    const templateFile = DriveApp.getFileById(template.fileId);
    const tempFolder = getOrCreateTempFolder();

    // 出力ファイル名を生成
    if (!outputFileName) {
      const dateStr = Utilities.formatDate(
        visit.visitDate instanceof Date ? visit.visitDate : new Date(visit.visitDate),
        'Asia/Tokyo',
        'yyyyMMdd'
      );
      outputFileName = patient.name + '_' + dateStr;
    }

    // 拡張子を確保
    if (!outputFileName.endsWith('.xlsx') && !outputFileName.endsWith('.xlsm')) {
      outputFileName += '.xlsx';
    }

    const copiedFile = templateFile.makeCopy(outputFileName, tempFolder);

    // スプレッドシートとして開いてデータを埋め込む
    const ss = SpreadsheetApp.openById(copiedFile.getId());
    fillTemplateData(ss, patient, visit, testResults);

    // Excelとしてエクスポート
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + copiedFile.getId() + '/export?format=xlsx';
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + token }
    });

    // 最終的なExcelファイルを作成
    const excelBlob = response.getBlob().setName(outputFileName);
    const excelFile = tempFolder.createFile(excelBlob);

    // コピーしたスプレッドシートを削除
    copiedFile.setTrashed(true);

    // ダウンロードURLを生成（30分間有効）
    excelFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + excelFile.getId();

    // 一時ファイルに有効期限を設定（プロパティで管理）
    const props = PropertiesService.getScriptProperties();
    const expiry = new Date();
    expiry.setMinutes(expiry.getMinutes() + TEMPLATE_CONFIG.TEMP_FILE_EXPIRY_MINUTES);
    props.setProperty('temp_' + excelFile.getId(), expiry.toISOString());

    return {
      success: true,
      downloadUrl: downloadUrl,
      fileName: outputFileName,
      fileId: excelFile.getId(),
      expiresAt: expiry.toISOString()
    };

  } catch (e) {
    logError('exportToExcelWithTemplate', e);
    return { success: false, error: e.message };
  }
}

/**
 * テンプレートにデータを埋め込む
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Object} patient - 患者情報
 * @param {Object} visit - 受診情報
 * @param {Array} testResults - 検査結果
 */
function fillTemplateData(ss, patient, visit, testResults) {
  // マッピング設定を読み込み
  const mapping = loadMappingConfig();

  // 患者基本情報を埋め込む
  if (mapping.patient_info && mapping.patient_info.fields) {
    const sheet = ss.getSheetByName(mapping.patient_info.sheet || '1ページ');
    if (sheet) {
      const fields = mapping.patient_info.fields;

      if (fields.name && fields.name.cell) {
        sheet.getRange(fields.name.cell).setValue(patient.name || '');
      }
      if (fields.exam_date && fields.exam_date.cell) {
        const dateStr = Utilities.formatDate(
          visit.visitDate instanceof Date ? visit.visitDate : new Date(visit.visitDate),
          'Asia/Tokyo',
          'yyyy/MM/dd'
        );
        sheet.getRange(fields.exam_date.cell).setValue(dateStr);
      }
      if (fields.gender && fields.gender.cell) {
        sheet.getRange(fields.gender.cell).setValue(patient.gender === 'M' ? '男' : '女');
      }
    }
  }

  // 検査結果を埋め込む
  if (mapping.test_items && mapping.test_items.items) {
    const sheet = ss.getSheetByName(mapping.test_items.sheet || '４ページ');
    if (sheet) {
      const items = mapping.test_items.items;

      // 検査結果をBMLコードでマップ
      const resultMap = {};
      testResults.forEach(r => {
        resultMap[r.itemId] = r.value;
      });

      // マッピングに従って値を設定
      for (const [bmlCode, config] of Object.entries(items)) {
        if (config.cell && resultMap[bmlCode] !== undefined) {
          sheet.getRange(config.cell).setValue(resultMap[bmlCode]);
        }
      }
    }
  }

  // 重複項目（複数箇所に同じ値を設定）
  if (mapping.duplicate_items && mapping.duplicate_items.mappings) {
    const sheet = ss.getSheetByName(mapping.test_items?.sheet || '４ページ');
    if (sheet) {
      const resultMap = {};
      testResults.forEach(r => {
        resultMap[r.itemId] = r.value;
      });

      for (const [key, config] of Object.entries(mapping.duplicate_items.mappings)) {
        const bmlCode = key.replace('_alt', '');
        if (config.cell && resultMap[bmlCode] !== undefined) {
          sheet.getRange(config.cell).setValue(resultMap[bmlCode]);
        }
      }
    }
  }
}

/**
 * マッピング設定を読み込む
 * @returns {Object} マッピング設定
 */
function loadMappingConfig() {
  // デフォルトのマッピング設定
  // 実際の運用では設定シートまたはJSONファイルから読み込む
  return {
    patient_info: {
      sheet: '1ページ',
      fields: {
        name: { cell: 'C8' },
        exam_date: { cell: 'I11' },
        gender: { cell: null }
      }
    },
    test_items: {
      sheet: '４ページ',
      items: {
        '0000301': { name: 'WBC', cell: 'K6' },
        '0000302': { name: 'RBC', cell: 'K7' },
        '0000303': { name: 'Hb', cell: 'K8' },
        '0000304': { name: 'Ht', cell: 'K9' },
        '0000308': { name: 'PLT', cell: 'K10' },
        '0000305': { name: 'MCV', cell: 'K11' },
        '0000306': { name: 'MCH', cell: 'K12' },
        '0000307': { name: 'MCHC', cell: 'K13' },
        '0000472': { name: 'T-Bil', cell: 'K20' },
        '0000401': { name: 'TP', cell: 'K21' },
        '0000417': { name: 'ALB', cell: 'K22' },
        '0000491': { name: 'BUN', cell: 'K23' },
        '0000413': { name: 'Cre', cell: 'K24' },
        '0002696': { name: 'eGFR', cell: 'K25' },
        '0000503': { name: 'FBS', cell: 'K26' },
        '0000460': { name: 'HDL', cell: 'K28' },
        '0000410': { name: 'LDL', cell: 'K29' },
        '0000454': { name: 'TG', cell: 'K30' },
        '0000407': { name: 'UA', cell: 'K40' },
        '0000481': { name: 'AST', cell: 'K52' },
        '0000482': { name: 'ALT', cell: 'K53' },
        '0000484': { name: 'GGT', cell: 'K54' },
        '0000497': { name: 'LDH', cell: 'K55' },
        '0013067': { name: 'ALP', cell: 'K56' },
        '0003317': { name: 'HbA1c', cell: 'K70' },
        '0000658': { name: 'CRP', cell: 'K84' },
        '0003550': { name: 'Na', cell: 'K78' },
        '0000421': { name: 'K', cell: 'K79' },
        '0000425': { name: 'Cl', cell: 'K80' }
      }
    },
    duplicate_items: {
      mappings: {
        '0002696_alt': { cell: 'K74' },
        '0000413_alt': { cell: 'K72' },
        '0000491_alt': { cell: 'K73' },
        '0000407_alt': { cell: 'K76' },
        '0000454_alt': { cell: 'K64' },
        '0000460_alt': { cell: 'K65' },
        '0000410_alt': { cell: 'K66' },
        '0000503_alt': { cell: 'K69' }
      }
    }
  };
}

// ============================================
// 一時ファイルクリーンアップ
// ============================================

/**
 * 期限切れの一時ファイルを削除
 * トリガーで定期実行推奨
 */
function cleanupExpiredTempFiles() {
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    const now = new Date();
    let deletedCount = 0;

    for (const [key, value] of Object.entries(allProps)) {
      if (key.startsWith('temp_')) {
        const expiry = new Date(value);
        if (now > expiry) {
          const fileId = key.replace('temp_', '');
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
            deletedCount++;
          } catch (e) {
            // ファイルが既に存在しない場合は無視
          }
          props.deleteProperty(key);
        }
      }
    }

    Logger.log('Cleaned up ' + deletedCount + ' expired temp files');
    return { success: true, deletedCount: deletedCount };

  } catch (e) {
    logError('cleanupExpiredTempFiles', e);
    return { success: false, error: e.message };
  }
}

// ============================================
// UI用API関数
// ============================================

/**
 * テンプレート管理データを取得（UI用）
 * @returns {Object} テンプレート管理データ
 */
function getTemplateManagementData() {
  try {
    // シートが存在しない場合は初期化
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(TEMPLATE_CONFIG.SHEET_NAME);
    if (!sheet) {
      setupTemplateManagement();
    }

    return {
      templates: getTemplateList(),
      examTypes: getExamTypeList()
    };
  } catch (e) {
    console.error('getTemplateManagementData error:', e);
    return {
      templates: [],
      examTypes: getExamTypeList(),
      error: e.message
    };
  }
}

/**
 * Excel出力API（UI用）
 * Cloud Functionsを優先使用し、失敗時はローカル版にフォールバック
 * @param {string} visitId - 受診ID
 * @param {string} templateId - テンプレートID（Cloud Functions版では未使用）
 * @param {string} fileName - ファイル名
 * @returns {Object} 出力結果
 */
function apiExportExcel(visitId, templateId, fileName) {
  console.log('apiExportExcel呼び出し: visitId=' + visitId + ', templateId=' + templateId);

  // visitIdバリデーション
  if (!visitId) {
    return { success: false, error: 'visitIdが指定されていません' };
  }

  // Cloud Functions版を優先使用
  try {
    if (typeof apiExportExcelFromCloud === 'function') {
      console.log('Cloud Functions版でExcel出力を実行: visitId=' + visitId);
      const cloudResult = apiExportExcelFromCloud(visitId);

      if (cloudResult.success) {
        return cloudResult;
      }

      console.warn('Cloud Functions版失敗、ローカル版にフォールバック:', cloudResult.error);
    }
  } catch (e) {
    console.warn('Cloud Functions版エラー、ローカル版にフォールバック:', e.message);
  }

  // ローカル版（従来の方法）
  console.log('ローカル版でExcel出力を実行');
  return exportToExcelWithTemplate(visitId, templateId, fileName);
}

// ============================================
// テスト用関数
// ============================================

/**
 * Excel出力テスト（GASエディタから実行用）
 * 受診記録シートから最新のvisitIdを自動取得してテスト
 */
function testExcelCloudExport() {
  const ss = getSpreadsheet();

  // 受診記録シートから取得（DB_CONFIG.SHEETS.VISIT_RECORD = '受診記録'）
  const sheetName = (typeof DB_CONFIG !== 'undefined' && DB_CONFIG.SHEETS)
    ? DB_CONFIG.SHEETS.VISIT_RECORD
    : '受診記録';

  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.log('シートが見つかりません: ' + sheetName);
    console.log('利用可能なシート: ' + ss.getSheets().map(s => s.getName()).join(', '));
    return;
  }

  if (sheet.getLastRow() <= 1) {
    console.log('受診記録データがありません。テストデータを作成してください。');
    console.log('または testExcelCloudExportWithId("V001") を使用してください。');
    return;
  }

  // 最新の受診IDを取得（2行目、1列目）
  const visitId = sheet.getRange(2, 1).getValue();
  console.log('テスト用visitId: ' + visitId);

  if (!visitId) {
    console.log('有効なvisitIdがありません');
    return;
  }

  // Excel出力テスト
  const result = apiExportExcel(visitId, null, 'テスト出力');
  console.log('結果: ' + JSON.stringify(result));

  if (result.success) {
    console.log('成功！ダウンロードURL: ' + result.downloadUrl);
  } else {
    console.log('失敗: ' + result.error);
  }
}

/**
 * 指定したvisitIdでExcel出力テスト
 * @param {string} visitId - テスト用受診ID
 */
function testExcelCloudExportWithId(visitId) {
  if (!visitId) {
    console.log('visitIdを指定してください');
    console.log('例: testExcelCloudExportWithId("V001")');
    return;
  }

  console.log('テスト用visitId: ' + visitId);

  const result = apiExportExcel(visitId, null, 'テスト出力');
  console.log('結果: ' + JSON.stringify(result));

  if (result.success) {
    console.log('成功！ダウンロードURL: ' + result.downloadUrl);
  } else {
    console.log('失敗: ' + result.error);
  }
}

/**
 * 利用可能なシート一覧を表示
 */
function listAvailableSheets() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();

  console.log('=== 利用可能なシート ===');
  sheets.forEach(function(sheet, index) {
    const lastRow = sheet.getLastRow();
    console.log((index + 1) + '. ' + sheet.getName() + ' (' + lastRow + '行)');
  });
}

/**
 * 患者マスタのデータを確認
 */
function listPatients() {
  const ss = getSpreadsheet();

  // 患者マスタシート名を取得
  const sheetName = (typeof DB_CONFIG !== 'undefined' && DB_CONFIG.SHEETS)
    ? DB_CONFIG.SHEETS.PATIENT_MASTER
    : '患者マスタ';

  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.log('患者マスタシートが見つかりません: ' + sheetName);
    return;
  }

  const lastRow = sheet.getLastRow();
  console.log('=== 患者マスタ (' + lastRow + '行) ===');

  if (lastRow <= 1) {
    console.log('患者データがありません');
    return;
  }

  // ヘッダーと最初の5件を表示
  const data = sheet.getRange(1, 1, Math.min(lastRow, 6), 3).getValues();
  data.forEach(function(row, index) {
    console.log((index === 0 ? 'ヘッダー' : index) + ': ' + row.join(' | '));
  });
}

/**
 * 受診記録のデータを確認
 */
function listVisitRecords() {
  const ss = getSpreadsheet();

  const sheetName = (typeof DB_CONFIG !== 'undefined' && DB_CONFIG.SHEETS)
    ? DB_CONFIG.SHEETS.VISIT_RECORD
    : '受診記録';

  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.log('受診記録シートが見つかりません: ' + sheetName);
    return;
  }

  const lastRow = sheet.getLastRow();
  console.log('=== 受診記録 (' + lastRow + '行) ===');

  if (lastRow <= 1) {
    console.log('受診記録がありません');
    return;
  }

  // ヘッダーと最初の5件を表示
  const data = sheet.getRange(1, 1, Math.min(lastRow, 6), 4).getValues();
  data.forEach(function(row, index) {
    console.log((index === 0 ? 'ヘッダー' : index) + ': ' + row.join(' | '));
  });
}

/**
 * 受診記録のpatientIdを修正
 * 不正なオブジェクト形式を修正し、存在する患者IDに変更
 */
function fixVisitRecordPatientIds() {
  const ss = getSpreadsheet();

  // 患者マスタから有効なIDを取得
  const patientSheetName = (typeof DB_CONFIG !== 'undefined' && DB_CONFIG.SHEETS)
    ? DB_CONFIG.SHEETS.PATIENT_MASTER
    : '患者マスタ';
  const patientSheet = ss.getSheetByName(patientSheetName);

  if (!patientSheet || patientSheet.getLastRow() <= 1) {
    console.log('患者マスタにデータがありません');
    return;
  }

  // 有効な患者ID一覧
  const patientData = patientSheet.getRange(2, 1, patientSheet.getLastRow() - 1, 1).getValues();
  const validPatientIds = patientData.map(row => String(row[0])).filter(id => id);
  console.log('有効な患者ID: ' + validPatientIds.join(', '));

  // 受診記録を修正
  const visitSheetName = (typeof DB_CONFIG !== 'undefined' && DB_CONFIG.SHEETS)
    ? DB_CONFIG.SHEETS.VISIT_RECORD
    : '受診記録';
  const visitSheet = ss.getSheetByName(visitSheetName);

  if (!visitSheet || visitSheet.getLastRow() <= 1) {
    console.log('受診記録がありません');
    return;
  }

  const visitData = visitSheet.getRange(2, 1, visitSheet.getLastRow() - 1, 2).getValues();
  let fixedCount = 0;

  for (let i = 0; i < visitData.length; i++) {
    const visitId = visitData[i][0];
    const patientId = String(visitData[i][1]);

    // 不正なオブジェクト形式をチェック
    if (patientId.includes('patientId=') || patientId.includes('success=')) {
      // 最初の有効な患者IDに置き換え
      const newPatientId = validPatientIds[0];
      visitSheet.getRange(i + 2, 2).setValue(newPatientId);
      console.log('修正: ' + visitId + ' のpatientIdを ' + patientId + ' → ' + newPatientId);
      fixedCount++;
    }
  }

  console.log('修正完了: ' + fixedCount + '件');
}
