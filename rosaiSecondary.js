/**
 * 労災二次検診専用モジュール
 *
 * 機能:
 * - 案件フォルダからCSV自動検索
 * - 入力シートへのデータ転記
 * - 超音波検査プルダウン・自動所見入力
 * - 入力状況チェック
 */

// ============================================
// 労災二次検診 セルマッピング
// ============================================
const ROSAI_CELL_MAPPING = {
  // 共通情報
  EXAM_DATE: 'F18',           // 受診日 (結合: F18:H18)
  DOCTOR_NAME: 'M36',         // 担当医師名 (結合: M36:O36)

  // 超音波検査
  CARDIAC_JUDGMENT: 'D19',    // 心臓超音波 判定 (プルダウン)
  CARDIAC_FINDINGS: 'F19',    // 心臓超音波 所見 (結合: F19:G19)
  CAROTID_JUDGMENT: 'D20',    // 頸動脈超音波 判定 (プルダウン)
  CAROTID_FINDINGS: 'F20',    // 頸動脈超音波 所見 (結合: F20:G20)

  // 医師入力項目
  HEALTH_GUIDANCE: 'A31',     // 特定保健指導 (結合: A31:P32)
  DOCTOR_FINDINGS: 'A34',     // 総合所見 (結合: A34:P35)

  // 血液検査結果（参照用）
  HDL_VALUE: 'F21',
  LDL_VALUE: 'F22',
  TG_VALUE: 'F23',
  FBS_VALUE: 'F24',
  HBA1C_VALUE: 'F25',
};

// 超音波検査 判定選択肢
const ULTRASOUND_GRADES = ['A', 'B', 'C', 'D', 'E'];

// 判定A選択時の自動入力テキスト
const AUTO_FINDINGS_TEXT = '異常なし';

// ============================================
// 案件フォルダ設定
// ============================================
const ROSAI_FOLDER_CONFIG = {
  // 案件ベースフォルダID（設定シートから読み込み or 直接指定）
  // /40_労災二次検診/10_案件/
  BASE_FOLDER_ID: null,  // 実行時に設定シートから読み込み

  // CSV格納サブフォルダパス
  CSV_SUBFOLDER: '30_AppSheetデータ',

  // CSVファイル名パターン
  CSV_PATTERN: /^判定結果_.*\.csv$/,
};

// ============================================
// 案件管理
// ============================================

/**
 * 案件フォルダ一覧を取得
 * @returns {Array<Object>} 案件情報の配列
 */
function getRosaiCaseList() {
  const baseFolderId = getRosaiBaseFolderId();
  if (!baseFolderId) {
    throw new Error('労災二次検診の案件フォルダIDが設定されていません。設定シートを確認してください。');
  }

  const baseFolder = DriveApp.getFolderById(baseFolderId);
  const cases = [];

  const subFolders = baseFolder.getFolders();
  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    const folderName = folder.getName();

    // 案件フォルダ名パターン: YYYYMMDD_企業名
    const match = folderName.match(/^(\d{8})_(.+)$/);
    if (match) {
      const dateStr = match[1];
      const companyName = match[2];
      const formattedDate = formatDateFromString(dateStr);

      // CSVファイルの有無をチェック
      const hasCsv = checkCsvExists(folder);

      cases.push({
        folderId: folder.getId(),
        folderName: folderName,
        date: dateStr,
        dateFormatted: formattedDate,
        companyName: companyName,
        hasCsv: hasCsv,
        // 互換性のため追加
        id: folder.getId(),
        name: folderName
      });
    }
  }

  // 日付降順でソート（新しい案件が上）
  cases.sort((a, b) => b.date.localeCompare(a.date));

  return cases;
}

/**
 * 労災二次検診ベースフォルダIDを取得
 * @returns {string|null} フォルダID
 */
function getRosaiBaseFolderId() {
  // 設定シートから読み込み（両方のキー名に対応）
  let folderId = getSettingValue('ROSAI_CASE_FOLDER_ID');
  if (!folderId || folderId === 'YOUR_ROSAI_FOLDER_ID') {
    // 旧キー名もチェック
    folderId = getSettingValue('ROSAI_CSV_FOLDER_ID');
  }
  if (folderId && folderId !== 'YOUR_ROSAI_FOLDER_ID') {
    return folderId;
  }
  return null;
}

/**
 * 日付文字列をフォーマット
 * @param {string} dateStr - YYYYMMDD形式
 * @returns {string} YYYY年MM月DD日形式
 */
function formatDateFromString(dateStr) {
  if (!dateStr || dateStr.length !== 8) return dateStr;
  const year = dateStr.slice(0, 4);
  const month = dateStr.slice(4, 6);
  const day = dateStr.slice(6, 8);
  return `${year}年${parseInt(month)}月${parseInt(day)}日`;
}

/**
 * 案件フォルダ内にCSVが存在するかチェック
 * @param {Folder} caseFolder - 案件フォルダ
 * @returns {boolean}
 */
function checkCsvExists(caseFolder) {
  try {
    const csvFolder = getCsvSubfolder(caseFolder);
    if (!csvFolder) return false;

    const files = csvFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (ROSAI_FOLDER_CONFIG.CSV_PATTERN.test(file.getName())) {
        return true;
      }
    }
  } catch (e) {
    // フォルダアクセスエラー
  }
  return false;
}

/**
 * CSV格納サブフォルダを取得
 * @param {Folder} caseFolder - 案件フォルダ
 * @returns {Folder|null}
 */
function getCsvSubfolder(caseFolder) {
  const subFolders = caseFolder.getFolders();
  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    if (folder.getName() === ROSAI_FOLDER_CONFIG.CSV_SUBFOLDER) {
      return folder;
    }
  }
  return null;
}

// ============================================
// CSV自動検索・読み込み
// ============================================

/**
 * 案件フォルダから最新の判定結果CSVを自動検索
 * @param {string} caseFolderId - 案件フォルダID
 * @returns {Object} {success, file, error}
 */
function findLatestJudgmentCsv(caseFolderId) {
  try {
    const caseFolder = DriveApp.getFolderById(caseFolderId);
    const csvFolder = getCsvSubfolder(caseFolder);

    if (!csvFolder) {
      return {
        success: false,
        file: null,
        error: `CSVフォルダが見つかりません: ${ROSAI_FOLDER_CONFIG.CSV_SUBFOLDER}`
      };
    }

    // 判定結果CSVを検索
    const csvFiles = [];
    const files = csvFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();

      if (ROSAI_FOLDER_CONFIG.CSV_PATTERN.test(name)) {
        csvFiles.push({
          file: file,
          name: name,
          date: file.getLastUpdated()
        });
      }
    }

    if (csvFiles.length === 0) {
      return {
        success: false,
        file: null,
        error: '判定結果CSVが見つかりません。\n30_AppSheetデータフォルダに「判定結果_*.csv」を配置してください。'
      };
    }

    // 最新ファイルを取得（更新日時降順）
    csvFiles.sort((a, b) => b.date - a.date);
    const latestFile = csvFiles[0].file;

    logInfo(`労災二次検診CSV検出: ${latestFile.getName()}`);

    return {
      success: true,
      file: latestFile,
      error: null
    };

  } catch (e) {
    logError('findLatestJudgmentCsv', e);
    return {
      success: false,
      file: null,
      error: `CSVファイル検索エラー: ${e.message}`
    };
  }
}

/**
 * 判定結果CSVを読み込み
 * @param {File} csvFile - CSVファイル
 * @returns {Array<Object>} 受診者データ配列
 */
function loadJudgmentCsv(csvFile) {
  const content = readFileContent(csvFile);
  const lines = content.trim().split('\n');

  if (lines.length < 2) {
    throw new Error('CSVにデータ行がありません');
  }

  // ヘッダー解析
  const headers = lines[0].split(',').map(h => h.trim());
  const results = [];

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const fields = line.split(',');
    const record = {};

    // ヘッダーに基づいてフィールドをマッピング
    for (let j = 0; j < headers.length; j++) {
      record[headers[j]] = fields[j] ? fields[j].trim() : '';
    }

    results.push(record);
  }

  return results;
}

// ============================================
// 入力シート管理
// ============================================

/**
 * 労災二次検診入力シートを作成・更新
 * @param {string} caseFolderId - 案件フォルダID
 * @param {string} doctorName - 担当医師名
 * @returns {Object} {success, sheetUrl, patientCount, error}
 */
function createRosaiInputSheet(caseFolderId, doctorName) {
  try {
    // 1. 案件情報を取得
    const caseFolder = DriveApp.getFolderById(caseFolderId);
    const caseName = caseFolder.getName();
    const match = caseName.match(/^(\d{8})_(.+)$/);

    if (!match) {
      return { success: false, error: '案件フォルダ名の形式が不正です' };
    }

    const examDate = match[1];
    const companyName = match[2];

    // 2. CSVを自動検索
    const csvResult = findLatestJudgmentCsv(caseFolderId);
    if (!csvResult.success) {
      return { success: false, error: csvResult.error };
    }

    // 3. CSVデータを読み込み
    const patients = loadJudgmentCsv(csvResult.file);
    if (patients.length === 0) {
      return { success: false, error: 'CSVにデータがありません' };
    }

    // 4. 入力シートを作成・更新
    const ss = getSpreadsheet();
    let inputSheet = ss.getSheetByName('労災二次検診_入力');

    if (!inputSheet) {
      inputSheet = ss.insertSheet('労災二次検診_入力');
    } else {
      // 既存シートをクリア
      inputSheet.clear();
    }

    // 5. シート構造を設定
    setupRosaiInputSheet(inputSheet, examDate, companyName, doctorName, patients);

    return {
      success: true,
      sheetUrl: ss.getUrl() + '#gid=' + inputSheet.getSheetId(),
      patientCount: patients.length,
      error: null
    };

  } catch (e) {
    logError('createRosaiInputSheet', e);
    return { success: false, error: e.message };
  }
}

/**
 * 入力シートの構造を設定
 * @param {Sheet} sheet - シート
 * @param {string} examDate - 受診日(YYYYMMDD)
 * @param {string} companyName - 企業名
 * @param {string} doctorName - 担当医師名
 * @param {Array<Object>} patients - 受診者データ
 */
function setupRosaiInputSheet(sheet, examDate, companyName, doctorName, patients) {
  // ヘッダー行
  const headers = [
    'No',           // A
    '名前',         // B
    'カナ',         // C
    '生年月日',     // D (H23.2.3形式)
    '年齢',         // E
    '性別',         // F
    '心臓判定',     // G (プルダウン)
    '心臓所見',     // H
    '頸動脈判定',   // I (プルダウン)
    '頸動脈所見',   // J
    '指導',         // K (チェック)
    '所見',         // L (チェック)
    'chart_no',     // M (参照用)
    'HDL',          // N
    'LDL',          // O
    'TG',           // P
    'FBS',          // Q
    'HbA1c',        // R
    'ACR'           // S (尿中アルブミン/Cre比)
  ];

  // 案件情報ヘッダー（1-3行目）
  sheet.getRange('A1').setValue('案件');
  sheet.getRange('B1').setValue(`${formatDateFromString(examDate)} ${companyName}`);
  sheet.getRange('A2').setValue('担当医師');
  sheet.getRange('B2').setValue(doctorName || '');
  sheet.getRange('A3').setValue('受診者数');
  sheet.getRange('B3').setValue(patients.length + '名');

  // データヘッダー（5行目）
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(5, 1, 1, headers.length)
    .setBackground('#4285f4')
    .setFontColor('white')
    .setFontWeight('bold');

  // データ行（6行目以降）
  const dataRows = [];
  for (let i = 0; i < patients.length; i++) {
    const p = patients[i];
    dataRows.push([
      p.No || (i + 1),                              // No
      p.name || '',                                  // 名前
      '',                                            // カナ（CSVにない場合は空）
      '',                                            // 生年月日（手入力: H23.2.3形式）
      p.age || '',                                   // 年齢
      p.gender || '',                                // 性別
      '',                                            // 心臓判定（プルダウン）
      '',                                            // 心臓所見
      '',                                            // 頸動脈判定（プルダウン）
      '',                                            // 頸動脈所見
      '',                                            // 指導チェック
      '',                                            // 所見チェック
      p.chart_no || '',                              // chart_no
      p.hdl_c || p.hdl_c_value || '',                // HDL (両方の形式に対応)
      p.ldl_c || p.ldl_c_value || '',                // LDL
      p.tg || p.tg_value || '',                      // TG
      p.fbs || p.fbs_value || '',                    // FBS
      p.hba1c || p.hba1c_value || '',                // HbA1c
      p.acr || p.acr_value || ''                     // ACR
    ]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(6, 1, dataRows.length, headers.length).setValues(dataRows);
  }

  // プルダウン設定（心臓判定: G列、頸動脈判定: I列）
  const dataStartRow = 6;
  const dataEndRow = 5 + patients.length;

  const gradeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ULTRASOUND_GRADES, true)
    .build();

  // G列（心臓判定）
  sheet.getRange(dataStartRow, 7, patients.length, 1).setDataValidation(gradeRule);
  // I列（頸動脈判定）
  sheet.getRange(dataStartRow, 9, patients.length, 1).setDataValidation(gradeRule);

  // 列幅調整
  sheet.setColumnWidth(1, 40);   // No
  sheet.setColumnWidth(2, 100);  // 名前
  sheet.setColumnWidth(3, 100);  // カナ
  sheet.setColumnWidth(4, 80);   // 生年月日
  sheet.setColumnWidth(5, 40);   // 年齢
  sheet.setColumnWidth(6, 40);   // 性別
  sheet.setColumnWidth(7, 60);   // 心臓判定
  sheet.setColumnWidth(8, 200);  // 心臓所見
  sheet.setColumnWidth(9, 60);   // 頸動脈判定
  sheet.setColumnWidth(10, 200); // 頸動脈所見
  sheet.setColumnWidth(11, 40);  // 指導
  sheet.setColumnWidth(12, 40);  // 所見

  // 入力状況列の条件付き書式（チェックマーク表示）
  setupCheckmarkFormatting(sheet, dataStartRow, dataEndRow);

  logInfo(`労災二次検診入力シート作成: ${patients.length}名`);
}

/**
 * チェックマーク表示用の条件付き書式を設定
 * @param {Sheet} sheet - シート
 * @param {number} startRow - 開始行
 * @param {number} endRow - 終了行
 */
function setupCheckmarkFormatting(sheet, startRow, endRow) {
  // 指導列(J): A31セルに値があれば✅
  // 所見列(K): A34セルに値があれば✅
  // ※実際のチェックはonEdit時またはExcel出力時に行う

  // 入力済みの場合の色設定
  const inputDoneColor = '#d4edda';  // 薄緑

  // G列（心臓所見）が入力済みならJ列を緑に
  // I列（頸動脈所見）が入力済みならK列を緑に
  // これはonEditで動的に更新する
}

// ============================================
// onEdit トリガー処理
// ============================================

/**
 * シート編集時のトリガー（既存onOpenに追加）
 * @param {Object} e - イベントオブジェクト
 */
function onEditRosaiSecondary(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // 労災二次検診入力シートのみ処理
  if (sheetName !== '労災二次検診_入力') {
    return;
  }

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value;

  // データ行のみ処理（6行目以降）
  if (row < 6) {
    return;
  }

  // 心臓判定(F列=6)または頸動脈判定(H列=8)が変更された場合
  if (col === 6 || col === 8) {
    handleUltrasoundJudgmentChange(sheet, row, col, value);
  }

  // 入力状況チェックの更新
  updateInputStatusCheck(sheet, row);
}

/**
 * 超音波判定変更時の処理
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @param {string} value - 選択された判定
 */
function handleUltrasoundJudgmentChange(sheet, row, col, value) {
  // 対応する所見列を特定
  const findingsCol = (col === 6) ? 7 : 9;  // 心臓所見=G(7), 頸動脈所見=I(9)

  if (value === 'A') {
    // Aが選択された場合、「異常なし」を自動入力
    sheet.getRange(row, findingsCol).setValue(AUTO_FINDINGS_TEXT);
    logInfo(`行${row}: 判定A選択 → 所見に「${AUTO_FINDINGS_TEXT}」自動入力`);
  } else if (value === '') {
    // 判定がクリアされた場合、所見もクリア（オプション）
    // sheet.getRange(row, findingsCol).setValue('');
  }
  // B以降の判定では所見は手入力のまま
}

/**
 * 入力状況チェックを更新
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号
 */
function updateInputStatusCheck(sheet, row) {
  // 心臓所見(H列)の入力状況
  const cardiacFindings = sheet.getRange(row, 8).getValue();
  // 頸動脈所見(J列)の入力状況
  const carotidFindings = sheet.getRange(row, 10).getValue();

  // 指導列(K)、所見列(L)のチェック表示を更新
  // ※実際のA31,A34セルは個人票出力時に転記するため、
  //   ここでは入力シート内の進捗管理用

  const cardiacDone = cardiacFindings ? '✅' : '⬜';
  const carotidDone = carotidFindings ? '✅' : '⬜';

  // K列に心臓所見の入力状況、L列に頸動脈所見の入力状況を表示
  // （将来的には指導・所見の入力状況も追加）
  sheet.getRange(row, 11).setValue(cardiacDone);
  sheet.getRange(row, 12).setValue(carotidDone);
}

// ============================================
// 労災二次検診 Excelテンプレート設定
// ============================================
const ROSAI_EXCEL_CONFIG = {
  // テンプレートファイルID（設定シートから読み込み or 直接指定）
  TEMPLATE_FILE_ID: null,  // 実行時に設定シートから読み込み

  // セルマッピング (kenshin_idheart.xlsx 準拠)
  CELL_MAPPING: {
    // 基本情報 (行5)
    COMPANY_NAME: 'B4',      // 事業所名
    PATIENT_NAME: 'B5',      // 受診者名
    GENDER: 'F5',            // 性別
    BIRTH_DATE: 'I5',        // 生年月日
    AGE: 'O5',               // 年齢

    // 受診日 (行18)
    EXAM_DATE: 'E18',        // 今回受診日
    PREV_EXAM_DATE: 'J18',   // 前回受診日

    // 超音波検査 (行19-20)
    CARDIAC_JUDGMENT: 'C19',    // 心臓超音波 判定
    CARDIAC_FINDINGS: 'D19',    // 心臓超音波 所見
    CAROTID_JUDGMENT: 'C20',    // 頸動脈超音波 判定
    CAROTID_FINDINGS: 'D20',    // 頸動脈超音波 所見

    // 血液検査 (行21-28) - 判定:C列, 値:D列, 前回:I列
    HDL_JUDGMENT: 'C21',
    HDL_VALUE: 'D21',
    HDL_PREV: 'I21',

    LDL_JUDGMENT: 'C22',
    LDL_VALUE: 'D22',
    LDL_PREV: 'I22',

    TG_JUDGMENT: 'C23',
    TG_VALUE: 'D23',
    TG_PREV: 'I23',

    FBS_JUDGMENT: 'C24',
    FBS_VALUE: 'D24',
    FBS_PREV: 'I24',

    HBA1C_JUDGMENT: 'C25',
    HBA1C_VALUE: 'D25',
    HBA1C_PREV: 'I25',

    // 腎機能 (行26-28)
    ALB_CRE_JUDGMENT: 'C26',
    ALB_CRE_VALUE: 'D26',

    ALB_JUDGMENT: 'C27',
    ALB_VALUE: 'D27',

    CRE_JUDGMENT: 'C28',
    CRE_VALUE: 'D28',

    // 所見エリア (行30-35)
    HEALTH_GUIDANCE: 'A31',     // 特定保健指導
    DOCTOR_FINDINGS: 'A34',     // 医師所見
    DOCTOR_NAME: 'M36',         // 担当医師名
  }
};

// ============================================
// Excel出力（個人票生成）
// ============================================

/**
 * 入力シートから個人票Excelを生成
 * @param {number} rowIndex - 入力シートの行番号
 * @returns {Object} {success, fileUrl, error}
 */
function exportRosaiPatientToExcel(rowIndex) {
  try {
    const ss = getSpreadsheet();
    const inputSheet = ss.getSheetByName('労災二次検診_入力');

    if (!inputSheet) {
      return { success: false, error: '入力シートが見つかりません' };
    }

    // 案件情報を取得
    const caseInfo = inputSheet.getRange('B1').getValue();
    const doctorName = inputSheet.getRange('B2').getValue();

    // 受診者データを取得
    const rowData = inputSheet.getRange(rowIndex, 1, 1, 17).getValues()[0];

    const patientData = {
      no: rowData[0],
      name: rowData[1],           // カナ名
      kana: rowData[2],           // カナ（別途）
      age: rowData[3],
      gender: rowData[4],
      cardiacJudgment: rowData[5],
      cardiacFindings: rowData[6],
      carotidJudgment: rowData[7],
      carotidFindings: rowData[8],
      guidance: rowData[9],       // 特定保健指導
      findings: rowData[10],      // 医師所見
      chartNo: rowData[11],
      hdl: rowData[12],
      ldl: rowData[13],
      tg: rowData[14],
      fbs: rowData[15],
      hba1c: rowData[16]
    };

    // 案件情報をパース
    const caseMatch = caseInfo.match(/(\d+年\d+月\d+日)\s*(.+)/);
    const examDate = caseMatch ? caseMatch[1] : '';
    const companyName = caseMatch ? caseMatch[2] : caseInfo;

    // テンプレートをコピーしてExcel生成
    const result = generateRosaiExcelFromTemplate(patientData, {
      examDate: examDate,
      companyName: companyName,
      doctorName: doctorName
    });

    return result;

  } catch (e) {
    logError('exportRosaiPatientToExcel', e);
    return { success: false, error: e.message };
  }
}

/**
 * テンプレートから個人票Excelを生成
 * @param {Object} patientData - 患者データ
 * @param {Object} caseInfo - 案件情報
 * @returns {Object} {success, fileUrl, error}
 */
function generateRosaiExcelFromTemplate(patientData, caseInfo) {
  // テンプレートファイルIDを取得
  const templateId = getRosaiTemplateFileId();
  if (!templateId) {
    return { success: false, error: 'テンプレートファイルIDが設定されていません。設定シートのROSAI_TEMPLATE_FILE_IDを確認してください。' };
  }

  try {
    // テンプレートをコピー
    const templateFile = DriveApp.getFileById(templateId);
    const outputFolder = getRosaiOutputFolder();

    // ファイル名: カルテNo_氏名_労災二次.xlsx
    const fileName = `${patientData.chartNo || patientData.no}_${patientData.name}_労災二次`;
    const copiedFile = templateFile.makeCopy(fileName, outputFolder);

    // スプレッドシートとして開く
    const copiedSs = SpreadsheetApp.openById(copiedFile.getId());
    const sheet = copiedSs.getSheetByName('template') || copiedSs.getSheets()[0];

    const mapping = ROSAI_EXCEL_CONFIG.CELL_MAPPING;

    // 基本情報を転記
    sheet.getRange(mapping.COMPANY_NAME).setValue(caseInfo.companyName || '');
    sheet.getRange(mapping.PATIENT_NAME).setValue(patientData.name || '');
    sheet.getRange(mapping.GENDER).setValue(patientData.gender || '');
    sheet.getRange(mapping.AGE).setValue(patientData.age || '');
    sheet.getRange(mapping.EXAM_DATE).setValue(caseInfo.examDate || '');

    // 超音波検査結果を転記
    sheet.getRange(mapping.CARDIAC_JUDGMENT).setValue(patientData.cardiacJudgment || '');
    sheet.getRange(mapping.CARDIAC_FINDINGS).setValue(patientData.cardiacFindings || '');
    sheet.getRange(mapping.CAROTID_JUDGMENT).setValue(patientData.carotidJudgment || '');
    sheet.getRange(mapping.CAROTID_FINDINGS).setValue(patientData.carotidFindings || '');

    // 血液検査結果を転記（値と判定）
    const gender = patientData.gender === '女性' ? 'F' : 'M';

    // HDL
    if (patientData.hdl) {
      sheet.getRange(mapping.HDL_VALUE).setValue(patientData.hdl);
      const hdlJudgment = judge('HDL_CHOLESTEROL', toNumber(patientData.hdl), gender);
      sheet.getRange(mapping.HDL_JUDGMENT).setValue(hdlJudgment);
    }

    // LDL
    if (patientData.ldl) {
      sheet.getRange(mapping.LDL_VALUE).setValue(patientData.ldl);
      const ldlJudgment = judge('LDL_CHOLESTEROL', toNumber(patientData.ldl), gender);
      sheet.getRange(mapping.LDL_JUDGMENT).setValue(ldlJudgment);
    }

    // TG
    if (patientData.tg) {
      sheet.getRange(mapping.TG_VALUE).setValue(patientData.tg);
      const tgJudgment = judge('TRIGLYCERIDES', toNumber(patientData.tg), gender);
      sheet.getRange(mapping.TG_JUDGMENT).setValue(tgJudgment);
    }

    // FBS
    if (patientData.fbs) {
      sheet.getRange(mapping.FBS_VALUE).setValue(patientData.fbs);
      const fbsJudgment = judge('FASTING_GLUCOSE', toNumber(patientData.fbs), gender);
      sheet.getRange(mapping.FBS_JUDGMENT).setValue(fbsJudgment);
    }

    // HbA1c
    if (patientData.hba1c) {
      sheet.getRange(mapping.HBA1C_VALUE).setValue(patientData.hba1c);
      const hba1cJudgment = judge('HBA1C', toNumber(patientData.hba1c), gender);
      sheet.getRange(mapping.HBA1C_JUDGMENT).setValue(hba1cJudgment);
    }

    // 所見を転記
    if (patientData.guidance) {
      sheet.getRange(mapping.HEALTH_GUIDANCE).setValue(patientData.guidance);
    }
    if (patientData.findings) {
      sheet.getRange(mapping.DOCTOR_FINDINGS).setValue(patientData.findings);
    }

    // 担当医師名
    sheet.getRange(mapping.DOCTOR_NAME).setValue(caseInfo.doctorName || '');

    // Excelとしてエクスポート
    SpreadsheetApp.flush();
    const excelBlob = convertSpreadsheetToExcel(copiedSs);
    const excelFile = outputFolder.createFile(excelBlob.setName(fileName + '.xlsx'));

    // 一時スプレッドシートを削除
    DriveApp.getFileById(copiedSs.getId()).setTrashed(true);

    logInfo(`労災二次検診Excel出力完了: ${fileName}`);

    return {
      success: true,
      fileUrl: excelFile.getUrl(),
      fileName: fileName + '.xlsx',
      error: null
    };

  } catch (e) {
    logError('generateRosaiExcelFromTemplate', e);
    return { success: false, error: e.message };
  }
}

/**
 * スプレッドシートをExcel形式に変換
 * @param {Spreadsheet} spreadsheet - スプレッドシート
 * @returns {Blob} Excelファイルのblob
 */
function convertSpreadsheetToExcel(spreadsheet) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Excel変換に失敗: ' + response.getContentText());
  }

  return response.getBlob();
}

/**
 * 労災二次検診テンプレートファイルIDを取得
 * @returns {string|null} ファイルID
 */
function getRosaiTemplateFileId() {
  const templateId = getSettingValue('ROSAI_TEMPLATE_FILE_ID');
  if (templateId && templateId !== 'YOUR_TEMPLATE_FILE_ID') {
    return templateId;
  }
  return null;
}

/**
 * 労災二次検診出力フォルダを取得
 * 案件フォルダ内に「40_結果出力」フォルダを自動作成
 * @param {string} caseFolderId - 案件フォルダID（省略時は入力シートから取得）
 * @returns {Folder} 出力フォルダ
 */
function getRosaiOutputFolder(caseFolderId) {
  const OUTPUT_SUBFOLDER_NAME = '40_結果出力';

  // 案件フォルダIDが指定されていない場合、入力シートから取得を試みる
  if (!caseFolderId) {
    caseFolderId = getCurrentCaseFolderId();
  }

  if (caseFolderId) {
    try {
      const caseFolder = DriveApp.getFolderById(caseFolderId);

      // 40_結果出力フォルダを検索
      const subFolders = caseFolder.getFoldersByName(OUTPUT_SUBFOLDER_NAME);
      if (subFolders.hasNext()) {
        return subFolders.next();
      }

      // なければ作成
      const outputFolder = caseFolder.createFolder(OUTPUT_SUBFOLDER_NAME);
      logInfo(`出力フォルダを作成: ${caseFolder.getName()}/${OUTPUT_SUBFOLDER_NAME}`);
      return outputFolder;

    } catch (e) {
      logError('getRosaiOutputFolder', e);
    }
  }

  // フォールバック: 設定シートのROSAI_OUTPUT_FOLDER_ID
  const folderId = getSettingValue('ROSAI_OUTPUT_FOLDER_ID');
  if (folderId && folderId !== 'YOUR_OUTPUT_FOLDER_ID') {
    return DriveApp.getFolderById(folderId);
  }

  // 最終フォールバック: ルートフォルダ
  logInfo('警告: 出力フォルダが特定できないためルートフォルダを使用');
  return DriveApp.getRootFolder();
}

/**
 * 現在の案件フォルダIDを入力シートから取得
 * @returns {string|null} 案件フォルダID
 */
function getCurrentCaseFolderId() {
  try {
    const ss = getSpreadsheet();
    const inputSheet = ss.getSheetByName('労災二次検診_入力');

    if (!inputSheet) return null;

    // B1セルから案件情報を取得し、案件フォルダを特定
    const caseInfo = inputSheet.getRange('B1').getValue();
    if (!caseInfo) return null;

    // 案件情報から日付と企業名を抽出
    // 形式: "2024年11月19日 社会福祉法人そよかぜの家"
    const match = caseInfo.match(/(\d+)年(\d+)月(\d+)日\s*(.+)/);
    if (!match) return null;

    const year = match[1];
    const month = match[2].padStart(2, '0');
    const day = match[3].padStart(2, '0');
    const companyName = match[4].trim();

    // フォルダ名パターン: YYYYMMDD_企業名
    const expectedFolderName = `${year}${month}${day}_${companyName}`;

    // 案件ベースフォルダから検索
    const baseFolderId = getRosaiBaseFolderId();
    if (!baseFolderId) return null;

    const baseFolder = DriveApp.getFolderById(baseFolderId);
    const folders = baseFolder.getFolders();

    while (folders.hasNext()) {
      const folder = folders.next();
      if (folder.getName() === expectedFolderName) {
        return folder.getId();
      }
    }

    // 部分一致で検索
    const foldersAgain = baseFolder.getFolders();
    while (foldersAgain.hasNext()) {
      const folder = foldersAgain.next();
      const folderName = folder.getName();
      if (folderName.includes(companyName) || folderName.includes(`${year}${month}${day}`)) {
        return folder.getId();
      }
    }

  } catch (e) {
    logError('getCurrentCaseFolderId', e);
  }

  return null;
}

/**
 * 入力シートの全員を一括Excel出力
 * @returns {Object} {success, files, errors}
 */
function exportAllRosaiPatientsToExcel() {
  const ss = getSpreadsheet();
  const inputSheet = ss.getSheetByName('労災二次検診_入力');

  if (!inputSheet) {
    return { success: false, files: [], errors: ['入力シートが見つかりません'] };
  }

  const lastRow = inputSheet.getLastRow();
  if (lastRow < 6) {
    return { success: false, files: [], errors: ['データがありません'] };
  }

  const results = {
    success: true,
    files: [],
    errors: []
  };

  for (let row = 6; row <= lastRow; row++) {
    const result = exportRosaiPatientToExcel(row);
    if (result.success) {
      results.files.push({
        row: row,
        fileName: result.fileName,
        url: result.fileUrl
      });
    } else {
      results.errors.push(`行${row}: ${result.error}`);
    }
  }

  logInfo(`労災二次検診一括出力完了: 成功${results.files.length}件, エラー${results.errors.length}件`);

  return results;
}

// ============================================
// ユーティリティ
// ============================================

/**
 * 入力シートの全受診者の入力状況を一括更新
 */
function refreshAllInputStatus() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('労災二次検診_入力');

  if (!sheet) {
    logInfo('労災二次検診入力シートが見つかりません');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) {
    return;
  }

  for (let row = 6; row <= lastRow; row++) {
    updateInputStatusCheck(sheet, row);
  }

  logInfo('入力状況を更新しました');
}

// ============================================
// 所見テンプレートシート管理
// ============================================

/**
 * 所見テンプレートシートを初期化
 * 超音波所見・総合所見のテンプレートを管理
 */
function initializeFindingsTemplateSheet() {
  const ss = getSpreadsheet();
  const sheetName = '所見テンプレート';

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('確認',
      '所見テンプレートシートが既に存在します。初期化すると既存データが消えます。続行しますか？',
      ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
      return;
    }
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // ヘッダー設定
  const headers = ['種別', '対象', '判定', '所見テキスト', '順序', '有効'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285f4')
    .setFontColor('white')
    .setFontWeight('bold');

  // 初期データ
  const initialData = [
    // 超音波所見 - 心臓
    ['超音波', '心臓', 'A', '異常を認めません', 1, true],
    ['超音波', '心臓', 'B', '軽度の弁膜逆流を認めます', 1, true],
    ['超音波', '心臓', 'B', '軽度の左室肥大を認めます', 2, true],
    ['超音波', '心臓', 'B', '軽度の拡張障害を認めます', 3, true],
    ['超音波', '心臓', 'C', '弁膜症を認めます。経過観察が必要です', 1, true],
    ['超音波', '心臓', 'C', '左室肥大を認めます。経過観察が必要です', 2, true],
    ['超音波', '心臓', 'C', '心肥大を認めます。経過観察が必要です', 3, true],
    ['超音波', '心臓', 'D', '心機能低下を認めます。精密検査をお勧めします', 1, true],
    ['超音波', '心臓', 'D', '重度の弁膜症を認めます。精密検査をお勧めします', 2, true],

    // 超音波所見 - 頸動脈
    ['超音波', '頸動脈', 'A', '異常を認めません', 1, true],
    ['超音波', '頸動脈', 'B', '軽度のIMT肥厚を認めます', 1, true],
    ['超音波', '頸動脈', 'B', '軽度のプラークを認めます', 2, true],
    ['超音波', '頸動脈', 'C', 'プラークを認めます。経過観察が必要です', 1, true],
    ['超音波', '頸動脈', 'C', 'IMT肥厚を認めます。動脈硬化の進行に注意が必要です', 2, true],
    ['超音波', '頸動脈', 'D', '高度狭窄を認めます。精密検査をお勧めします', 1, true],
    ['超音波', '頸動脈', 'D', '不安定プラークの疑いがあります。精密検査をお勧めします', 2, true],

    // 総合所見 - 血液検査項目別
    ['総合所見', 'HDL', 'C', 'HDLコレステロールが低値です。運動習慣の改善をお勧めします。', 1, true],
    ['総合所見', 'HDL', 'D', 'HDLコレステロールが著明低値です。精査をお勧めします。', 1, true],
    ['総合所見', 'LDL', 'C', 'LDLコレステロールが高値です。食事療法をお勧めします。', 1, true],
    ['総合所見', 'LDL', 'D', 'LDLコレステロールが著明高値です。精査・治療をお勧めします。', 1, true],
    ['総合所見', 'TG', 'C', '中性脂肪が高値です。糖質・アルコールの摂取を控えてください。', 1, true],
    ['総合所見', 'TG', 'D', '中性脂肪が著明高値です。精査・治療をお勧めします。', 1, true],
    ['総合所見', 'FBS', 'C', '空腹時血糖が高値です。糖尿病の疑いがあります。', 1, true],
    ['総合所見', 'FBS', 'D', '空腹時血糖が著明高値です。糖尿病の治療が必要です。', 1, true],
    ['総合所見', 'HbA1c', 'C', 'HbA1cが高値です。糖尿病の疑いがあります。', 1, true],
    ['総合所見', 'HbA1c', 'D', 'HbA1cが著明高値です。糖尿病の治療が必要です。', 1, true],
    ['総合所見', 'ACR', 'B', '尿中アルブミンが軽度上昇しています。経過観察をお勧めします。', 1, true],
    ['総合所見', 'ACR', 'C', '尿中アルブミンが上昇しています。腎機能の経過観察が必要です。', 1, true],
    ['総合所見', 'ACR', 'D', '尿中アルブミンが著明上昇しています。腎臓専門医の受診をお勧めします。', 1, true],
  ];

  if (initialData.length > 0) {
    sheet.getRange(2, 1, initialData.length, headers.length).setValues(initialData);
  }

  // 列幅調整
  sheet.setColumnWidth(1, 80);   // 種別
  sheet.setColumnWidth(2, 80);   // 対象
  sheet.setColumnWidth(3, 50);   // 判定
  sheet.setColumnWidth(4, 400);  // 所見テキスト
  sheet.setColumnWidth(5, 50);   // 順序
  sheet.setColumnWidth(6, 50);   // 有効

  // データ検証（判定列）
  const gradeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['A', 'B', 'C', 'D'], true)
    .build();
  sheet.getRange(2, 3, 100, 1).setDataValidation(gradeRule);

  // データ検証（有効列）
  const boolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .build();
  sheet.getRange(2, 6, 100, 1).setDataValidation(boolRule);

  logInfo('所見テンプレートシートを初期化しました');

  const ui = SpreadsheetApp.getUi();
  ui.alert('完了', '所見テンプレートシートを初期化しました。\nテンプレートを編集して使用してください。', ui.ButtonSet.OK);
}

/**
 * 超音波所見テンプレートを取得
 * @param {string} targetOrgan - 対象臓器（心臓/頸動脈）
 * @param {string} judgment - 判定（A/B/C/D）
 * @returns {Array<Object>} テンプレート配列 [{text, order}]
 */
function getUltrasoundTemplates(targetOrgan, judgment) {
  const templates = [];

  try {
    const sheet = getSheet('所見テンプレート');
    if (!sheet) return templates;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return templates;

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    for (const row of data) {
      const type = row[0];      // 種別
      const target = row[1];    // 対象
      const grade = row[2];     // 判定
      const text = row[3];      // 所見テキスト
      const order = row[4];     // 順序
      const enabled = row[5];   // 有効

      if (type === '超音波' && target === targetOrgan && grade === judgment && enabled) {
        templates.push({
          text: text,
          order: order || 1
        });
      }
    }

    // 順序でソート
    templates.sort((a, b) => a.order - b.order);

  } catch (e) {
    logError('getUltrasoundTemplates', e);
  }

  return templates;
}

/**
 * 総合所見テンプレートを取得
 * @param {string} itemCode - 項目コード（HDL/LDL/TG/FBS/HbA1c/ACR）
 * @param {string} judgment - 判定（A/B/C/D）
 * @returns {string|null} テンプレートテキスト
 */
function getSummaryFindingTemplate(itemCode, judgment) {
  try {
    const sheet = getSheet('所見テンプレート');
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    for (const row of data) {
      const type = row[0];      // 種別
      const target = row[1];    // 対象
      const grade = row[2];     // 判定
      const text = row[3];      // 所見テキスト
      const enabled = row[5];   // 有効

      if (type === '総合所見' && target === itemCode && grade === judgment && enabled) {
        return text;
      }
    }
  } catch (e) {
    logError('getSummaryFindingTemplate', e);
  }

  return null;
}

// ============================================
// H/L 判定機能
// ============================================

/**
 * 検査値のH/L（高値/低値）を判定
 * @param {string} itemCode - 項目コード（HDL_CHOLESTEROL等）
 * @param {number} value - 検査値
 * @param {string} gender - 性別（M/F）
 * @returns {string} 'H'（高値）/ 'L'（低値）/ ''（正常範囲）
 */
function getHighLowFlag(itemCode, value, gender) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  const numValue = toNumber(value);
  if (numValue === null) {
    return '';
  }

  // 基準値定義（正常範囲）
  const referenceRanges = {
    'HDL_CHOLESTEROL': { min: 40, max: 100 },
    'LDL_CHOLESTEROL': { min: 60, max: 119 },
    'TRIGLYCERIDES': { min: 30, max: 149 },
    'FASTING_GLUCOSE': { min: 70, max: 99 },
    'HBA1C': { min: 4.6, max: 5.5 },
    'ACR': { min: 0, max: 29.9 },
    'AST_GOT': { min: 0, max: 30 },
    'ALT_GPT': { min: 0, max: 30 },
    'GAMMA_GTP': { min: 0, max: 50 },
    'CREATININE': { min: 0.5, max: 1.0 },
    'EGFR': { min: 60, max: null },
    'URIC_ACID': { min: 2.1, max: 7.0 },
  };

  const range = referenceRanges[itemCode];
  if (!range) {
    return '';
  }

  if (range.max !== null && numValue > range.max) {
    return 'H';
  }
  if (range.min !== null && numValue < range.min) {
    return 'L';
  }

  return '';
}

// ============================================
// 超音波所見ダイアログ
// ============================================

/**
 * 超音波所見入力ダイアログを表示
 * @param {number} rowIndex - 入力シートの行番号
 */
function showUltrasoundFindingsDialog(rowIndex) {
  const ss = getSpreadsheet();
  const inputSheet = ss.getSheetByName('労災二次検診_入力');

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('エラー', '入力シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 現在のデータを取得
  const rowData = inputSheet.getRange(rowIndex, 1, 1, 17).getValues()[0];
  const patientName = rowData[1] || '(氏名不明)';
  const cardiacJudgment = rowData[5] || '';
  const cardiacFindings = rowData[6] || '';
  const carotidJudgment = rowData[7] || '';
  const carotidFindings = rowData[8] || '';

  // テンプレートを取得
  const cardiacTemplates = getUltrasoundTemplates('心臓', cardiacJudgment);
  const carotidTemplates = getUltrasoundTemplates('頸動脈', carotidJudgment);

  // ダイアログHTMLを生成
  const html = createUltrasoundDialogHtml(rowIndex, patientName, {
    cardiacJudgment,
    cardiacFindings,
    cardiacTemplates,
    carotidJudgment,
    carotidFindings,
    carotidTemplates
  });

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `超音波所見入力 - ${patientName}`);
}

/**
 * 超音波所見ダイアログのHTMLを生成
 */
function createUltrasoundDialogHtml(rowIndex, patientName, data) {
  // テンプレート選択肢を生成
  const createOptions = (templates, currentValue) => {
    let options = '<option value="">-- 選択してください --</option>';
    options += '<option value="__FREE__">フリーテキスト入力</option>';

    for (const t of templates) {
      const selected = (currentValue === t.text) ? 'selected' : '';
      options += `<option value="${escapeHtml(t.text)}" ${selected}>${escapeHtml(t.text)}</option>`;
    }
    return options;
  };

  const cardiacOptions = createOptions(data.cardiacTemplates, data.cardiacFindings);
  const carotidOptions = createOptions(data.carotidTemplates, data.carotidFindings);

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
          font-size: 14px;
          padding: 20px;
          margin: 0;
        }
        .section {
          margin-bottom: 25px;
          padding: 15px;
          background: #f8f9fa;
          border-radius: 8px;
        }
        .section-title {
          font-weight: bold;
          font-size: 15px;
          margin-bottom: 10px;
          color: #1a73e8;
        }
        label {
          display: block;
          margin-bottom: 5px;
          font-weight: 500;
          color: #333;
        }
        select, textarea {
          width: 100%;
          padding: 8px;
          margin-bottom: 10px;
          border: 1px solid #ddd;
          border-radius: 4px;
          box-sizing: border-box;
          font-size: 13px;
        }
        textarea {
          height: 80px;
          resize: vertical;
        }
        .judgment-badge {
          display: inline-block;
          padding: 3px 10px;
          border-radius: 4px;
          font-weight: bold;
          margin-left: 10px;
        }
        .judgment-A { background: #d4edda; color: #155724; }
        .judgment-B { background: #fff3cd; color: #856404; }
        .judgment-C { background: #ffe5d0; color: #8a4500; }
        .judgment-D { background: #f8d7da; color: #721c24; }
        .btn-container {
          text-align: right;
          margin-top: 20px;
        }
        .btn {
          padding: 10px 25px;
          margin-left: 10px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
        }
        .btn-primary {
          background: #1a73e8;
          color: white;
        }
        .btn-secondary {
          background: #f1f3f4;
          color: #333;
        }
        .hidden { display: none; }
      </style>
    </head>
    <body>
      <form id="findingsForm">
        <input type="hidden" id="rowIndex" value="${rowIndex}">

        <!-- 心臓超音波 -->
        <div class="section">
          <div class="section-title">
            🫀 心臓超音波
            <span class="judgment-badge judgment-${data.cardiacJudgment || 'A'}">${data.cardiacJudgment || '未選択'}</span>
          </div>
          <label>所見テンプレート:</label>
          <select id="cardiacTemplate" onchange="onTemplateChange('cardiac')">
            ${cardiacOptions}
          </select>
          <label>所見（編集可能）:</label>
          <textarea id="cardiacFindings">${escapeHtml(data.cardiacFindings)}</textarea>
        </div>

        <!-- 頸動脈超音波 -->
        <div class="section">
          <div class="section-title">
            🩺 頸動脈超音波
            <span class="judgment-badge judgment-${data.carotidJudgment || 'A'}">${data.carotidJudgment || '未選択'}</span>
          </div>
          <label>所見テンプレート:</label>
          <select id="carotidTemplate" onchange="onTemplateChange('carotid')">
            ${carotidOptions}
          </select>
          <label>所見（編集可能）:</label>
          <textarea id="carotidFindings">${escapeHtml(data.carotidFindings)}</textarea>
        </div>

        <div class="btn-container">
          <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
          <button type="button" class="btn btn-primary" onclick="saveFindings()">保存</button>
        </div>
      </form>

      <script>
        function onTemplateChange(type) {
          const select = document.getElementById(type + 'Template');
          const textarea = document.getElementById(type + 'Findings');
          const value = select.value;

          if (value === '__FREE__') {
            textarea.value = '';
            textarea.focus();
          } else if (value) {
            textarea.value = value;
          }
        }

        function saveFindings() {
          const rowIndex = parseInt(document.getElementById('rowIndex').value);
          const cardiacFindings = document.getElementById('cardiacFindings').value;
          const carotidFindings = document.getElementById('carotidFindings').value;

          google.script.run
            .withSuccessHandler(function() {
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('保存エラー: ' + error.message);
            })
            .saveUltrasoundFindings(rowIndex, cardiacFindings, carotidFindings);
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * HTMLエスケープ
 */
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * 超音波所見を保存
 * @param {number} rowIndex - 行番号
 * @param {string} cardiacFindings - 心臓所見
 * @param {string} carotidFindings - 頸動脈所見
 */
function saveUltrasoundFindings(rowIndex, cardiacFindings, carotidFindings) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('労災二次検診_入力');

  if (!sheet) {
    throw new Error('入力シートが見つかりません');
  }

  // H列（心臓所見）、J列（頸動脈所見）に保存
  sheet.getRange(rowIndex, 8).setValue(cardiacFindings);
  sheet.getRange(rowIndex, 10).setValue(carotidFindings);

  // 入力状況を更新
  updateInputStatusCheck(sheet, rowIndex);

  logInfo(`超音波所見を保存: 行${rowIndex}`);
}

// ============================================
// 総合所見自動生成（労災二次検診用）
// ============================================

/**
 * 労災二次検診の総合所見を生成
 * @param {Object} patientData - 患者データ
 * @returns {string} 総合所見テキスト
 */
function generateRosaiSummaryFindings(patientData) {
  const findings = [];
  const gender = patientData.gender === '女性' ? 'F' : 'M';

  // 各項目の判定とテンプレートを取得
  const items = [
    { code: 'HDL_CHOLESTEROL', key: 'HDL', value: patientData.hdl },
    { code: 'LDL_CHOLESTEROL', key: 'LDL', value: patientData.ldl },
    { code: 'TRIGLYCERIDES', key: 'TG', value: patientData.tg },
    { code: 'FASTING_GLUCOSE', key: 'FBS', value: patientData.fbs },
    { code: 'HBA1C', key: 'HbA1c', value: patientData.hba1c },
  ];

  // ACRがある場合は追加
  if (patientData.acr) {
    items.push({ code: 'ACR', key: 'ACR', value: patientData.acr });
  }

  for (const item of items) {
    if (!item.value) continue;

    const numValue = toNumber(item.value);
    if (numValue === null) continue;

    // 判定を取得
    const judgment = judge(item.code, numValue, gender);

    // C/D判定の場合、テンプレートを取得
    if (judgment === 'C' || judgment === 'D') {
      const template = getSummaryFindingTemplate(item.key, judgment);
      if (template) {
        findings.push(template);
      }
    }
  }

  // 所見がない場合
  if (findings.length === 0) {
    return '今回の検査では特に問題は認められませんでした。';
  }

  return findings.join('\n');
}

// ============================================
// Excel出力時の統合処理
// ============================================

/**
 * 労災二次検診Excel出力（ダイアログ経由）
 * @param {number} rowIndex - 入力シートの行番号
 */
function exportRosaiPatientWithDialog(rowIndex) {
  // まず超音波所見ダイアログを表示
  showUltrasoundFindingsDialog(rowIndex);
}

/**
 * Excel出力前の確認ダイアログを表示
 * @param {number} rowIndex - 行番号
 */
function showExportConfirmDialog(rowIndex) {
  const ss = getSpreadsheet();
  const inputSheet = ss.getSheetByName('労災二次検診_入力');

  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('エラー', '入力シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rowData = inputSheet.getRange(rowIndex, 1, 1, 17).getValues()[0];

  const patientData = {
    no: rowData[0],
    name: rowData[1],
    age: rowData[3],
    gender: rowData[4],
    cardiacJudgment: rowData[5],
    cardiacFindings: rowData[6],
    carotidJudgment: rowData[7],
    carotidFindings: rowData[8],
    hdl: rowData[12],
    ldl: rowData[13],
    tg: rowData[14],
    fbs: rowData[15],
    hba1c: rowData[16]
  };

  // 総合所見を自動生成
  const autoFindings = generateRosaiSummaryFindings(patientData);

  // 確認ダイアログを表示
  const html = createExportConfirmDialogHtml(rowIndex, patientData, autoFindings);

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Excel出力確認 - ${patientData.name}`);
}

/**
 * 出力確認ダイアログHTMLを生成
 */
function createExportConfirmDialogHtml(rowIndex, patientData, autoFindings) {
  const gender = patientData.gender === '女性' ? 'F' : 'M';

  // H/Lフラグを計算
  const hlData = [
    { label: 'HDL', value: patientData.hdl, code: 'HDL_CHOLESTEROL' },
    { label: 'LDL', value: patientData.ldl, code: 'LDL_CHOLESTEROL' },
    { label: 'TG', value: patientData.tg, code: 'TRIGLYCERIDES' },
    { label: 'FBS', value: patientData.fbs, code: 'FASTING_GLUCOSE' },
    { label: 'HbA1c', value: patientData.hba1c, code: 'HBA1C' },
  ];

  let hlTableRows = '';
  for (const item of hlData) {
    const hl = getHighLowFlag(item.code, item.value, gender);
    const judgment = item.value ? judge(item.code, toNumber(item.value), gender) : '';
    const hlClass = hl === 'H' ? 'hl-high' : (hl === 'L' ? 'hl-low' : '');

    hlTableRows += `
      <tr>
        <td>${item.label}</td>
        <td>${item.value || '-'}</td>
        <td class="${hlClass}">${hl}</td>
        <td>${judgment}</td>
      </tr>
    `;
  }

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
        }
        h3 {
          margin: 0 0 15px 0;
          color: #333;
          font-size: 15px;
        }
        .section {
          margin-bottom: 20px;
          padding: 15px;
          background: #f8f9fa;
          border-radius: 8px;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin-bottom: 10px;
        }
        th, td {
          padding: 8px;
          text-align: left;
          border-bottom: 1px solid #ddd;
        }
        th {
          background: #e9ecef;
          font-weight: 600;
        }
        .hl-high { color: #dc3545; font-weight: bold; }
        .hl-low { color: #0d6efd; font-weight: bold; }
        textarea {
          width: 100%;
          padding: 10px;
          border: 1px solid #ddd;
          border-radius: 4px;
          font-size: 13px;
          height: 100px;
          resize: vertical;
          box-sizing: border-box;
        }
        .ultrasound-findings {
          background: #fff;
          padding: 10px;
          border: 1px solid #ddd;
          border-radius: 4px;
          margin-bottom: 10px;
        }
        .findings-label {
          font-weight: 600;
          margin-bottom: 5px;
        }
        .btn-container {
          text-align: right;
          margin-top: 20px;
        }
        .btn {
          padding: 10px 25px;
          margin-left: 10px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
        }
        .btn-primary {
          background: #1a73e8;
          color: white;
        }
        .btn-secondary {
          background: #f1f3f4;
          color: #333;
        }
        .btn-warning {
          background: #ffc107;
          color: #333;
        }
      </style>
    </head>
    <body>
      <input type="hidden" id="rowIndex" value="${rowIndex}">

      <div class="section">
        <h3>📊 検査結果とH/L判定</h3>
        <table>
          <tr>
            <th>項目</th>
            <th>値</th>
            <th>H/L</th>
            <th>判定</th>
          </tr>
          ${hlTableRows}
        </table>
      </div>

      <div class="section">
        <h3>🔬 超音波検査所見</h3>
        <div class="ultrasound-findings">
          <div class="findings-label">心臓超音波 [${patientData.cardiacJudgment || '-'}]:</div>
          <div>${escapeHtml(patientData.cardiacFindings) || '(未入力)'}</div>
        </div>
        <div class="ultrasound-findings">
          <div class="findings-label">頸動脈超音波 [${patientData.carotidJudgment || '-'}]:</div>
          <div>${escapeHtml(patientData.carotidFindings) || '(未入力)'}</div>
        </div>
        <button type="button" class="btn btn-warning" onclick="editUltrasound()">超音波所見を編集</button>
      </div>

      <div class="section">
        <h3>📝 総合所見（編集可能）</h3>
        <textarea id="summaryFindings">${escapeHtml(autoFindings)}</textarea>
        <div style="font-size: 11px; color: #666; margin-top: 5px;">
          ※ 判定結果に基づいて自動生成されています。必要に応じて編集してください。
        </div>
      </div>

      <div class="btn-container">
        <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
        <button type="button" class="btn btn-primary" onclick="executeExport()">Excel出力</button>
      </div>

      <script>
        function editUltrasound() {
          const rowIndex = parseInt(document.getElementById('rowIndex').value);
          google.script.run.showUltrasoundFindingsDialog(rowIndex);
          google.script.host.close();
        }

        function executeExport() {
          const rowIndex = parseInt(document.getElementById('rowIndex').value);
          const summaryFindings = document.getElementById('summaryFindings').value;

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                alert('Excel出力が完了しました。\\n\\nファイル: ' + result.fileName);
                google.script.host.close();
              } else {
                alert('エラー: ' + result.error);
              }
            })
            .withFailureHandler(function(error) {
              alert('エラー: ' + error.message);
            })
            .executeRosaiExcelExport(rowIndex, summaryFindings);
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * Excel出力を実行
 * @param {number} rowIndex - 行番号
 * @param {string} summaryFindings - 総合所見
 * @returns {Object} {success, fileName, error}
 */
function executeRosaiExcelExport(rowIndex, summaryFindings) {
  try {
    const ss = getSpreadsheet();
    const inputSheet = ss.getSheetByName('労災二次検診_入力');

    if (!inputSheet) {
      return { success: false, error: '入力シートが見つかりません' };
    }

    // 案件情報を取得
    const caseInfo = inputSheet.getRange('B1').getValue();
    const doctorName = inputSheet.getRange('B2').getValue();

    // 受診者データを取得
    const rowData = inputSheet.getRange(rowIndex, 1, 1, 17).getValues()[0];

    const patientData = {
      no: rowData[0],
      name: rowData[1],
      kana: rowData[2],
      age: rowData[3],
      gender: rowData[4],
      cardiacJudgment: rowData[5],
      cardiacFindings: rowData[6],
      carotidJudgment: rowData[7],
      carotidFindings: rowData[8],
      chartNo: rowData[11],
      hdl: rowData[12],
      ldl: rowData[13],
      tg: rowData[14],
      fbs: rowData[15],
      hba1c: rowData[16],
      summaryFindings: summaryFindings  // 編集された総合所見
    };

    // 案件情報をパース
    const caseMatch = caseInfo.match(/(\d+年\d+月\d+日)\s*(.+)/);
    const examDate = caseMatch ? caseMatch[1] : '';
    const companyName = caseMatch ? caseMatch[2] : caseInfo;

    // テンプレートからExcel生成（H/L付き）
    const result = generateRosaiExcelWithHL(patientData, {
      examDate: examDate,
      companyName: companyName,
      doctorName: doctorName
    });

    return result;

  } catch (e) {
    logError('executeRosaiExcelExport', e);
    return { success: false, error: e.message };
  }
}

/**
 * H/L列付きExcelを生成
 * @param {Object} patientData - 患者データ
 * @param {Object} caseInfo - 案件情報
 * @returns {Object} {success, fileUrl, fileName, error}
 */
function generateRosaiExcelWithHL(patientData, caseInfo) {
  const templateId = getRosaiTemplateFileId();
  if (!templateId) {
    return { success: false, error: 'テンプレートファイルIDが設定されていません。' };
  }

  try {
    const templateFile = DriveApp.getFileById(templateId);
    const outputFolder = getRosaiOutputFolder();

    const fileName = `${patientData.chartNo || patientData.no}_${patientData.name}_労災二次`;
    const copiedFile = templateFile.makeCopy(fileName, outputFolder);

    const copiedSs = SpreadsheetApp.openById(copiedFile.getId());
    const sheet = copiedSs.getSheetByName('template') || copiedSs.getSheets()[0];

    const mapping = ROSAI_EXCEL_CONFIG.CELL_MAPPING;
    const gender = patientData.gender === '女性' ? 'F' : 'M';

    // 基本情報を転記
    sheet.getRange(mapping.COMPANY_NAME).setValue(caseInfo.companyName || '');
    sheet.getRange(mapping.PATIENT_NAME).setValue(patientData.name || '');
    sheet.getRange(mapping.GENDER).setValue(patientData.gender || '');
    sheet.getRange(mapping.AGE).setValue(patientData.age || '');
    sheet.getRange(mapping.EXAM_DATE).setValue(caseInfo.examDate || '');

    // 超音波検査結果を転記
    sheet.getRange(mapping.CARDIAC_JUDGMENT).setValue(patientData.cardiacJudgment || '');
    sheet.getRange(mapping.CARDIAC_FINDINGS).setValue(patientData.cardiacFindings || '');
    sheet.getRange(mapping.CAROTID_JUDGMENT).setValue(patientData.carotidJudgment || '');
    sheet.getRange(mapping.CAROTID_FINDINGS).setValue(patientData.carotidFindings || '');

    // 血液検査結果を転記（値・判定・H/L）
    const bloodItems = [
      { key: 'HDL', value: patientData.hdl, code: 'HDL_CHOLESTEROL', valueCell: mapping.HDL_VALUE, judgmentCell: mapping.HDL_JUDGMENT },
      { key: 'LDL', value: patientData.ldl, code: 'LDL_CHOLESTEROL', valueCell: mapping.LDL_VALUE, judgmentCell: mapping.LDL_JUDGMENT },
      { key: 'TG', value: patientData.tg, code: 'TRIGLYCERIDES', valueCell: mapping.TG_VALUE, judgmentCell: mapping.TG_JUDGMENT },
      { key: 'FBS', value: patientData.fbs, code: 'FASTING_GLUCOSE', valueCell: mapping.FBS_VALUE, judgmentCell: mapping.FBS_JUDGMENT },
      { key: 'HBA1C', value: patientData.hba1c, code: 'HBA1C', valueCell: mapping.HBA1C_VALUE, judgmentCell: mapping.HBA1C_JUDGMENT },
    ];

    for (const item of bloodItems) {
      if (item.value) {
        // 値を転記
        sheet.getRange(item.valueCell).setValue(item.value);

        // 判定を計算して転記
        const judgment = judge(item.code, toNumber(item.value), gender);
        sheet.getRange(item.judgmentCell).setValue(judgment);

        // H/Lを計算してH列に転記
        const hl = getHighLowFlag(item.code, item.value, gender);
        if (hl) {
          // H列 = 値セルの2列右（D→F→H の場合は +4）
          // 値がD列の場合、H列はH列（D=4, H=8, 差分4）
          const valueCol = sheet.getRange(item.valueCell).getColumn();
          const hlCol = valueCol + 4;  // H列
          const valueRow = sheet.getRange(item.valueCell).getRow();
          sheet.getRange(valueRow, hlCol).setValue(hl);
        }
      }
    }

    // 総合所見を転記
    if (patientData.summaryFindings) {
      sheet.getRange(mapping.DOCTOR_FINDINGS).setValue(patientData.summaryFindings);
    }

    // 担当医師名
    sheet.getRange(mapping.DOCTOR_NAME).setValue(caseInfo.doctorName || '');

    // Excelとしてエクスポート
    SpreadsheetApp.flush();
    const excelBlob = convertSpreadsheetToExcel(copiedSs);
    const excelFile = outputFolder.createFile(excelBlob.setName(fileName + '.xlsx'));

    // 一時スプレッドシートを削除
    DriveApp.getFileById(copiedSs.getId()).setTrashed(true);

    logInfo(`労災二次検診Excel出力完了（H/L付き）: ${fileName}`);

    return {
      success: true,
      fileUrl: excelFile.getUrl(),
      fileName: fileName + '.xlsx',
      error: null
    };

  } catch (e) {
    logError('generateRosaiExcelWithHL', e);
    return { success: false, error: e.message };
  }
}

// ============================================
// メニューからのExcel出力アクセス
// ============================================

/**
 * 選択中の行のExcel出力確認ダイアログを表示
 * メニューまたはサイドバーから呼び出し
 */
function showExportDialogForSelectedRow() {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();

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

  showExportConfirmDialog(activeRow);
}

/**
 * 労災二次検診の受診者一覧サイドバーを表示
 * 個別Excel出力用
 */
function showRosaiPatientListSidebar() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('労災二次検診_入力');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', '入力シートがありません。先にデータ取込を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const html = HtmlService.createHtmlOutput(getRosaiPatientListHtml(sheet))
    .setTitle('受診者一覧 - Excel出力')
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 受診者一覧サイドバーのHTMLを生成
 * @param {Sheet} sheet - 入力シート
 * @returns {string} HTML
 */
function getRosaiPatientListHtml(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 6) {
    return `
      <html>
      <body style="font-family: 'Hiragino Sans', sans-serif; padding: 15px;">
        <h3>受診者一覧</h3>
        <p style="color: #666;">データがありません</p>
      </body>
      </html>
    `;
  }

  const data = sheet.getRange(6, 1, lastRow - 5, 12).getValues();
  let patientListHtml = '';

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 6;
    const no = row[0];
    const name = row[1];
    const cardiacJudgment = row[6] || '-';  // G列（心臓判定）
    const carotidJudgment = row[8] || '-';  // I列（頸動脈判定）

    if (!name) continue;

    // 入力状況をチェック
    const hasCardiac = row[6] && row[7];    // G列（判定）とH列（所見）
    const hasCarotid = row[8] && row[9];    // I列（判定）とJ列（所見）
    const statusClass = (hasCardiac && hasCarotid) ? 'status-complete' : 'status-pending';
    const statusIcon = (hasCardiac && hasCarotid) ? '✅' : '⏳';

    patientListHtml += `
      <div class="patient-row ${statusClass}" onclick="showExportDialog(${rowIndex})">
        <div class="patient-info">
          <span class="patient-no">${no}</span>
          <span class="patient-name">${name}</span>
        </div>
        <div class="patient-status">
          <span class="judgment-badge">心${cardiacJudgment}/頸${carotidJudgment}</span>
          <span class="status-icon">${statusIcon}</span>
        </div>
      </div>
    `;
  }

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
          font-size: 13px;
          padding: 15px;
          margin: 0;
        }
        h3 {
          margin: 0 0 15px 0;
          color: #1a73e8;
          font-size: 16px;
        }
        .info-text {
          font-size: 11px;
          color: #666;
          margin-bottom: 15px;
        }
        .patient-row {
          padding: 12px;
          border: 1px solid #ddd;
          border-radius: 6px;
          margin-bottom: 8px;
          cursor: pointer;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .patient-row:hover {
          background: #f0f4ff;
          border-color: #1a73e8;
        }
        .status-complete {
          background: #f0fff4;
        }
        .status-pending {
          background: #fff8e6;
        }
        .patient-no {
          font-weight: bold;
          margin-right: 10px;
          color: #555;
        }
        .patient-name {
          font-weight: 500;
        }
        .patient-status {
          display: flex;
          align-items: center;
          gap: 8px;
        }
        .judgment-badge {
          font-size: 11px;
          padding: 2px 6px;
          background: #e9ecef;
          border-radius: 3px;
        }
        .status-icon {
          font-size: 14px;
        }
        .btn {
          display: block;
          width: 100%;
          padding: 12px;
          margin-top: 15px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
          text-align: center;
        }
        .btn-secondary {
          background: #f1f3f4;
          color: #333;
        }
      </style>
    </head>
    <body>
      <h3>📋 受診者一覧</h3>
      <p class="info-text">受診者をクリックしてExcel出力確認画面を開きます</p>

      ${patientListHtml || '<p style="color: #666;">受診者がいません</p>'}

      <button class="btn btn-secondary" onclick="google.script.host.close()">閉じる</button>

      <script>
        function showExportDialog(rowIndex) {
          google.script.run
            .withSuccessHandler(function() {
              // ダイアログが開いたらサイドバーは閉じない
            })
            .withFailureHandler(function(error) {
              alert('エラー: ' + error.message);
            })
            .showExportConfirmDialog(rowIndex);
        }
      </script>
    </body>
    </html>
  `;
}
