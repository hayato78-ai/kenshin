/**
 * Excel出力ブリッジモジュール
 * GAS → Python連携用JSON出力
 *
 * 使い方:
 *   exportToExcelViaPython(patientId)  // 単一患者
 *   exportMultipleToExcelViaPython(patientIds)  // 複数患者
 */

// ============================================
// 設定
// ============================================
const EXCEL_BRIDGE_CONFIG = {
  // Driveフォルダ名（81_結果入力内のサブフォルダ）
  FOLDERS: {
    PENDING: 'pending',       // GAS出力 → Python入力
    COMPLETED: 'completed',   // Excel出力先
    PROCESSED: 'processed',   // 処理済みJSON
    STATUS: 'status'          // ステータス通知
  },

  // 健診種別
  EXAM_TYPES: {
    ROSAI_SECONDARY: 'ROSAI_SECONDARY',
    DOCK_STANDARD: 'DOCK_STANDARD',
    DOCK_PREMIUM: 'DOCK_PREMIUM',
    PERIODIC: 'PERIODIC'
  }
};

// ============================================
// メイン関数
// ============================================

/**
 * Python経由でExcel出力（単一患者）
 * @param {number} rowIndex - 入力シートの行番号
 * @param {string} examType - 健診種別（省略時: ROSAI_SECONDARY）
 * @returns {Object} {success, requestId, error}
 */
function exportToExcelViaPython(rowIndex, examType = 'ROSAI_SECONDARY') {
  try {
    logInfo(`Excel出力開始 (Python連携): row=${rowIndex}, type=${examType}`);

    // 1. 患者データを収集（入力シートの既存データをそのまま使用）
    const patientData = collectPatientDataForExport(rowIndex, examType);
    if (!patientData) {
      throw new Error('患者データの収集に失敗しました');
    }

    // ※ Claude API呼び出しは不要（保健指導は入力時に生成済み）

    // 2. JSONをDriveに保存
    const requestId = generateRequestId();
    const jsonData = buildExportJson(requestId, examType, patientData);
    const savedFile = saveJsonToDrive(requestId, jsonData);

    logInfo(`JSON保存完了: ${savedFile.getName()}`);

    // 3. ステータスを更新（オプション）
    updateExportStatus(rowIndex, 'pending', requestId);

    return {
      success: true,
      requestId: requestId,
      message: 'Excel出力リクエストを送信しました。数秒後に出力フォルダをご確認ください。'
    };

  } catch (e) {
    logError('exportToExcelViaPython', e);
    return {
      success: false,
      requestId: null,
      error: e.message
    };
  }
}

/**
 * Python経由でExcel出力（複数患者）
 * @param {Array<number>} rowIndices - 行番号の配列
 * @param {string} examType - 健診種別
 * @returns {Object} {success, results, error}
 */
function exportMultipleToExcelViaPython(rowIndices, examType = 'ROSAI_SECONDARY') {
  const results = {
    success: [],
    failed: []
  };

  for (const rowIndex of rowIndices) {
    const result = exportToExcelViaPython(rowIndex, examType);
    if (result.success) {
      results.success.push({ rowIndex, requestId: result.requestId });
    } else {
      results.failed.push({ rowIndex, error: result.error });
    }
  }

  return {
    success: results.failed.length === 0,
    results: results,
    message: `${results.success.length}件成功, ${results.failed.length}件失敗`
  };
}

// ============================================
// データ収集関数
// ============================================

/**
 * 患者データを収集（Excel出力用）
 * @param {number} rowIndex - 行番号
 * @param {string} examType - 健診種別
 * @returns {Object|null} 患者データ
 */
function collectPatientDataForExport(rowIndex, examType) {
  const ss = getSpreadsheet();

  // 健診種別に応じた入力シートを取得
  let inputSheet;
  switch (examType) {
    case 'ROSAI_SECONDARY':
      inputSheet = ss.getSheetByName('労災二次検診_入力');
      break;
    case 'DOCK_STANDARD':
    case 'DOCK_PREMIUM':
      inputSheet = ss.getSheetByName('人間ドック_入力');
      break;
    default:
      inputSheet = ss.getSheetByName('労災二次検診_入力');
  }

  if (!inputSheet) {
    logError('collectPatientDataForExport', new Error('入力シートが見つかりません'));
    return null;
  }

  // 案件情報を取得
  const caseInfo = inputSheet.getRange('B1').getValue();
  const doctorName = inputSheet.getRange('B2').getValue();

  // 案件情報をパース
  const caseMatch = caseInfo.match(/(\d+)年(\d+)月(\d+)日\s*(.+)/);
  const examDate = caseMatch ? `${caseMatch[1]}-${caseMatch[2].padStart(2, '0')}-${caseMatch[3].padStart(2, '0')}` : '';
  const companyName = caseMatch ? caseMatch[4].trim() : caseInfo;

  // 患者データを取得
  const rowData = inputSheet.getRange(rowIndex, 1, 1, 20).getValues()[0];

  // 労災二次検診の場合のデータマッピング
  if (examType === 'ROSAI_SECONDARY') {
    return {
      case: {
        case_id: `CASE_${formatDateForId(new Date())}`,
        company_name: companyName,
        exam_date: examDate,
        doctor_name: doctorName
      },
      patient: {
        patient_id: rowData[11] || `PAT_${rowIndex}`,  // カルテ番号
        name: rowData[1],                              // 氏名
        kana: rowData[2] || '',                        // カナ
        gender: rowData[4] === '女性' ? 'F' : 'M',
        age: rowData[3],
        birth_date: ''  // 必要に応じて追加
      },
      blood_tests: {
        hdl: {
          value: toNumberOrNull(rowData[12]),
          judgment: null,  // Python側で計算
          flag: null
        },
        ldl: {
          value: toNumberOrNull(rowData[13]),
          judgment: null,
          flag: null
        },
        tg: {
          value: toNumberOrNull(rowData[14]),
          judgment: null,
          flag: null
        },
        fbs: {
          value: toNumberOrNull(rowData[15]),
          judgment: null,
          flag: null
        },
        hba1c: {
          value: toNumberOrNull(rowData[16]),
          judgment: null,
          flag: null
        }
      },
      ultrasound: {
        cardiac: {
          judgment: rowData[5] || '',
          findings: rowData[6] || ''
        },
        carotid: {
          judgment: rowData[7] || '',
          findings: rowData[8] || ''
        },
        summary: rowData[10] || ''  // 総合所見
      },
      guidance: {
        health_guidance: rowData[9] || '',  // 特定保健指導
        doctor_findings: rowData[10] || ''  // 医師所見
      }
    };
  }

  // 人間ドック等の場合は別途マッピングを追加
  return null;
}

/**
 * 保健指導文を生成（Claude API）
 * @param {Object} patientData - 患者データ
 * @returns {Object} {healthGuidance, doctorFindings}
 */
function generateGuidanceForExport(patientData) {
  // 既存の保健指導があればそのまま使用
  if (patientData.guidance?.health_guidance && patientData.guidance.health_guidance.length > 10) {
    return {
      healthGuidance: patientData.guidance.health_guidance,
      doctorFindings: patientData.guidance.doctor_findings || patientData.ultrasound?.summary || ''
    };
  }

  // Claude APIで生成（既存関数があれば使用）
  try {
    if (typeof generateHealthGuidanceWithClaude === 'function') {
      const guidance = generateHealthGuidanceWithClaude(patientData);
      return {
        healthGuidance: guidance,
        doctorFindings: patientData.ultrasound?.summary || ''
      };
    }
  } catch (e) {
    logInfo('Claude API呼び出しスキップ: ' + e.message);
  }

  // フォールバック: 空文字を返す
  return {
    healthGuidance: '',
    doctorFindings: patientData.ultrasound?.summary || ''
  };
}

// ============================================
// JSON構築・保存関数
// ============================================

/**
 * 出力用JSONを構築
 * @param {string} requestId - リクエストID
 * @param {string} examType - 健診種別
 * @param {Object} patientData - 患者データ
 * @returns {Object} JSON構造
 */
function buildExportJson(requestId, examType, patientData) {
  // 案件フォルダ情報を取得
  const caseFolder = getCaseReportFolder();

  return {
    request_id: requestId,
    created_at: new Date().toISOString(),
    exam_type: examType,

    case: patientData.case,
    patient: patientData.patient,
    blood_tests: patientData.blood_tests,
    ultrasound: patientData.ultrasound,
    guidance: patientData.guidance,

    output: {
      template: getTemplateNameForExamType(examType),
      folder_id: caseFolder?.folderId || getOutputFolderId(),
      folder_path: caseFolder?.folderPath || null,
      case_name: caseFolder?.folderName || null
    }
  };
}

/**
 * JSONをDriveに保存
 * @param {string} requestId - リクエストID
 * @param {Object} jsonData - JSONデータ
 * @returns {File} 保存されたファイル
 */
function saveJsonToDrive(requestId, jsonData) {
  const pendingFolder = getOrCreateSubfolder(EXCEL_BRIDGE_CONFIG.FOLDERS.PENDING);
  const fileName = `${requestId}.json`;
  const content = JSON.stringify(jsonData, null, 2);

  const file = pendingFolder.createFile(fileName, content, MimeType.PLAIN_TEXT);
  return file;
}

// ============================================
// ユーティリティ関数
// ============================================

/**
 * リクエストIDを生成
 * @returns {string} リクエストID
 */
function generateRequestId() {
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const random = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `REQ_${timestamp}_${random}`;
}

/**
 * 日付をID用にフォーマット
 * @param {Date} date - 日付
 * @returns {string} YYYYMMDD形式
 */
function formatDateForId(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
}

/**
 * 数値に変換（nullを許容）
 * @param {*} value - 値
 * @returns {number|null}
 */
function toNumberOrNull(value) {
  if (value === '' || value === null || value === undefined) {
    return null;
  }
  const num = parseFloat(value);
  return isNaN(num) ? null : num;
}

/**
 * 健診種別に対応するテンプレート名を取得
 * @param {string} examType - 健診種別
 * @returns {string} テンプレート名
 */
function getTemplateNameForExamType(examType) {
  const templates = {
    'ROSAI_SECONDARY': 'rosai_secondary',
    'DOCK_STANDARD': 'dock_standard',
    'DOCK_PREMIUM': 'dock_premium',
    'PERIODIC': 'periodic_health'
  };
  return templates[examType] || 'rosai_secondary';
}

/**
 * 出力フォルダIDを取得
 * @returns {string} フォルダID
 */
function getOutputFolderId() {
  const folder = getOrCreateSubfolder(EXCEL_BRIDGE_CONFIG.FOLDERS.COMPLETED);
  return folder.getId();
}

/**
 * 案件の報告書フォルダパスを取得
 * @returns {Object} {folderId, folderPath, folderName}
 */
function getCaseReportFolder() {
  // 一時フォルダID（案件のCSVフォルダ）を取得
  const tempCsvFolderId = PropertiesService.getScriptProperties().getProperty('TEMP_CSV_FOLDER_ID');

  if (!tempCsvFolderId) {
    logInfo('案件フォルダ未設定: デフォルト出力先を使用');
    return null;
  }

  try {
    // CSVフォルダから親フォルダ（案件フォルダ）を取得
    const csvFolder = DriveApp.getFolderById(tempCsvFolderId);
    const caseFolder = csvFolder.getParents().next();
    const caseFolderName = caseFolder.getName();

    // 40_報告書 フォルダを探す
    const reportFolders = caseFolder.getFoldersByName('40_報告書');
    let reportFolder;

    if (reportFolders.hasNext()) {
      reportFolder = reportFolders.next();
    } else {
      // なければ作成
      reportFolder = caseFolder.createFolder('40_報告書');
      logInfo('40_報告書フォルダを作成しました');
    }

    // フォルダパスを構築
    const folderPath = buildFolderPath(reportFolder);

    return {
      folderId: reportFolder.getId(),
      folderPath: folderPath,
      folderName: caseFolderName,
      reportFolderName: reportFolder.getName()
    };

  } catch (e) {
    logError('getCaseReportFolder', e);
    return null;
  }
}

/**
 * フォルダの完全パスを構築
 * @param {Folder} folder - 対象フォルダ
 * @returns {string} フォルダパス
 */
function buildFolderPath(folder) {
  const pathParts = [];
  let current = folder;

  // ルートまで遡る（最大10階層）
  for (let i = 0; i < 10; i++) {
    pathParts.unshift(current.getName());
    const parents = current.getParents();
    if (!parents.hasNext()) break;
    current = parents.next();
    if (current.getName() === 'マイドライブ') break;
  }

  // Google Driveのローカルマウントパスを構築
  const basePath = '/Users/hytenhd_mac/Library/CloudStorage/GoogleDrive-buskenshin@cdmedical.jp/マイドライブ';
  return basePath + '/' + pathParts.join('/');
}

/**
 * ベースフォルダ内のサブフォルダを取得または作成
 * @param {string} subfolderName - サブフォルダ名
 * @returns {Folder} フォルダ
 */
function getOrCreateSubfolder(subfolderName) {
  // 81_結果入力フォルダを取得
  const baseFolderId = getSettingValue('EXCEL_BRIDGE_FOLDER_ID');

  let baseFolder;
  if (baseFolderId && baseFolderId !== 'YOUR_FOLDER_ID') {
    baseFolder = DriveApp.getFolderById(baseFolderId);
  } else {
    // フォールバック: スプレッドシートと同じフォルダに作成
    const ss = getSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    baseFolder = ssFile.getParents().next();

    // excel_output フォルダを作成
    const outputFolders = baseFolder.getFoldersByName('excel_output');
    if (outputFolders.hasNext()) {
      baseFolder = outputFolders.next();
    } else {
      baseFolder = baseFolder.createFolder('excel_output');
    }
  }

  // サブフォルダを取得または作成
  const subFolders = baseFolder.getFoldersByName(subfolderName);
  if (subFolders.hasNext()) {
    return subFolders.next();
  }
  return baseFolder.createFolder(subfolderName);
}

/**
 * 出力ステータスを更新
 * @param {number} rowIndex - 行番号
 * @param {string} status - ステータス
 * @param {string} requestId - リクエストID
 */
function updateExportStatus(rowIndex, status, requestId) {
  // 必要に応じてスプレッドシートにステータスを記録
  // 例: 入力シートの特定列に出力ステータスを記録
  logInfo(`ステータス更新: row=${rowIndex}, status=${status}, requestId=${requestId}`);
}

// ============================================
// ステータス確認関数
// ============================================

/**
 * 出力ステータスを確認
 * @param {string} requestId - リクエストID
 * @returns {Object|null} ステータス情報
 */
function checkExportStatus(requestId) {
  try {
    const statusFolder = getOrCreateSubfolder(EXCEL_BRIDGE_CONFIG.FOLDERS.STATUS);
    const statusFiles = statusFolder.getFilesByName(`${requestId}_status.json`);

    if (statusFiles.hasNext()) {
      const file = statusFiles.next();
      const content = file.getBlob().getDataAsString();
      return JSON.parse(content);
    }

    // ステータスファイルがない場合はpending
    return { status: 'pending', request_id: requestId };

  } catch (e) {
    logError('checkExportStatus', e);
    return null;
  }
}

/**
 * 完了したExcelファイルのURLを取得
 * @param {string} requestId - リクエストID
 * @returns {string|null} ファイルURL
 */
function getCompletedExcelUrl(requestId) {
  const status = checkExportStatus(requestId);
  if (status && status.status === 'completed' && status.output_file) {
    return status.output_file.url;
  }
  return null;
}

// ============================================
// メニュー連携関数
// ============================================

/**
 * 選択行をPython経由でExcel出力（メニューから呼び出し）
 */
function exportSelectedRowViaPython() {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 労災二次検診_入力シートかチェック
  if (sheet.getName() !== '労災二次検診_入力') {
    SpreadsheetApp.getUi().alert('エラー', '労災二次検診_入力シートを選択してください', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const activeRow = ss.getActiveRange().getRow();

  // データ行（6行目以降）かチェック
  if (activeRow < 6) {
    SpreadsheetApp.getUi().alert('エラー', 'データ行（6行目以降）を選択してください', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 行にデータがあるかチェック
  const name = sheet.getRange(activeRow, 2).getValue();
  if (!name) {
    SpreadsheetApp.getUi().alert('エラー', '選択した行にデータがありません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 確認ダイアログ
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Excel出力確認',
    `${name} さんのExcelを出力しますか？\n\n※Python経由で高品質なExcelを生成します`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // 出力実行
  const result = exportToExcelViaPython(activeRow, 'ROSAI_SECONDARY');

  if (result.success) {
    ui.alert('出力リクエスト送信',
      `リクエストID: ${result.requestId}\n\n${result.message}`,
      ui.ButtonSet.OK);
  } else {
    ui.alert('エラー', `出力に失敗しました: ${result.error}`, ui.ButtonSet.OK);
  }
}

/**
 * 全員をPython経由でExcel出力（メニューから呼び出し）
 */
function exportAllRowsViaPython() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('労災二次検診_入力');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', '労災二次検診_入力シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // データ行を検索
  const lastRow = sheet.getLastRow();
  const rowIndices = [];

  for (let row = 6; row <= lastRow; row++) {
    const name = sheet.getRange(row, 2).getValue();
    if (name) {
      rowIndices.push({ row, name });
    }
  }

  if (rowIndices.length === 0) {
    SpreadsheetApp.getUi().alert('エラー', '出力対象のデータがありません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 確認ダイアログ
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '一括Excel出力確認',
    `${rowIndices.length}名分のExcelを出力しますか？\n（1人あたり約1-2秒）`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // 進捗表示付きで出力実行
  const results = { success: [], failed: [] };
  const total = rowIndices.length;

  for (let i = 0; i < total; i++) {
    const { row, name } = rowIndices[i];

    // 進捗表示（トースト通知）
    ss.toast(`処理中: ${name}（${i + 1}/${total}）`, '📊 Excel出力', 3);

    const result = exportToExcelViaPython(row, 'ROSAI_SECONDARY');

    if (result.success) {
      results.success.push({ row, name, requestId: result.requestId });
    } else {
      results.failed.push({ row, name, error: result.error });
    }
  }

  // 完了通知
  ss.toast('', '✅ 出力完了', 1);

  // 結果サマリー
  let message = `✅ 成功: ${results.success.length}件\n`;
  if (results.failed.length > 0) {
    message += `❌ 失敗: ${results.failed.length}件\n\n`;
    message += '失敗した対象:\n';
    results.failed.forEach(f => {
      message += `  - ${f.name}: ${f.error}\n`;
    });
  }

  ui.alert('出力結果', message, ui.ButtonSet.OK);
}
