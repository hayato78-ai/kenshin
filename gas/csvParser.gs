/**
 * CSVパーサーモジュール
 * BML検査結果CSVの解析処理
 *
 * 対応ファイル形式:
 * - 結果データ: XXXXXXXX_OXXXXXXX.csv（数字_O数字.csv）
 * - 依頼データ: kensa_irai_*.csv（結果なしのためスキップ）
 */

/**
 * CSVファイルを解析（健診種別に応じてパーサーを切替）
 * @param {string} fileId - CSVファイルのID
 * @returns {Array<Object>} 患者データの配列
 */
function parseCSV(fileId) {
  const file = DriveApp.getFileById(fileId);
  const content = readFileContent(file);
  const profile = CONFIG.getProfile();

  // 健診種別に応じてパーサーを切替
  if (profile.csvFormat === 'ROSAI') {
    return parseRosaiCSV(content);
  } else {
    return parseBmlCSV(content);
  }
}

/**
 * BML形式CSVを解析（人間ドック用）
 * @param {string} content - CSVファイル内容
 * @returns {Array<Object>} 患者データの配列
 */
function parseBmlCSV(content) {
  const results = [];
  const lines = content.trim().split('\n');

  for (const line of lines) {
    if (!line.trim()) continue;

    const parsed = parseLine(line);
    if (parsed && parsed.testResults.length > 0) {
      results.push(parsed);
    }
  }

  return results;
}

/**
 * 労災二次検診形式CSVを解析
 * ヘッダー行付きの標準CSVフォーマット
 * @param {string} content - CSVファイル内容
 * @returns {Array<Object>} 患者データの配列
 */
function parseRosaiCSV(content) {
  const results = [];
  const lines = content.trim().split('\n');

  if (lines.length < 2) {
    logInfo('労災CSV: データ行がありません');
    return results;
  }

  // ヘッダー行を解析
  const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
  logInfo(`労災CSV: ヘッダー検出 - ${headers.join(', ')}`);

  // ヘッダーからカラムインデックスを取得
  const columnMap = buildRosaiColumnMap(headers);

  // データ行を処理
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;

    const parsed = parseRosaiLine(line, columnMap);
    if (parsed && (parsed.testResults.length > 0 || parsed.patientInfo.name)) {
      results.push(parsed);
    }
  }

  logInfo(`労災CSV: ${results.length}件のデータを解析`);
  return results;
}

/**
 * 労災二次検診CSVのカラムマッピングを構築
 * @param {Array<string>} headers - ヘッダー配列
 * @returns {Object} カラムマッピング
 */
function buildRosaiColumnMap(headers) {
  // CSVヘッダー名 → 内部キー名のマッピング
  const headerToKey = {
    'chart_no': 'chartNo',
    'name': 'name',
    'age': 'age',
    'gender': 'gender',
    'fbs': 'fbs',
    'hba1c': 'hba1c',
    'tg': 'tg',
    'hdl_c': 'hdl_c',
    'ldl_c': 'ldl_c',
    'got': 'got',
    'gpt': 'gpt',
    'gamma_gpt': 'gamma_gpt',
    'tp': 'tp',
    'ua': 'ua',
    'cre': 'cre',
    'wbc': 'wbc',
    'hb': 'hb',
    'plt': 'plt',
    'bmi': 'bmi',
    'waist': 'waist',
    'sbp': 'sbp',
    'dbp': 'dbp',
    'organization': 'organization',
    '事業所名': 'organization',
    '企業名': 'organization',
    'company': 'organization',
    'company_name': 'organization'
  };

  const columnMap = {};

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const key = headerToKey[header];
    if (key) {
      columnMap[key] = i;
    }
  }

  return columnMap;
}

/**
 * 労災二次検診CSVの1行を解析
 * @param {string} line - CSV行
 * @param {Object} columnMap - カラムマッピング
 * @returns {Object|null} 解析結果
 */
function parseRosaiLine(line, columnMap) {
  const fields = line.split(',').map(f => f.trim());

  // 患者情報を抽出
  const patientInfo = {
    requestId: columnMap.chartNo !== undefined ? fields[columnMap.chartNo] : '',
    name: columnMap.name !== undefined ? fields[columnMap.name] : '',
    age: columnMap.age !== undefined ? fields[columnMap.age] : '',
    gender: columnMap.gender !== undefined ? fields[columnMap.gender] : '',
    examDate: formatDate(new Date(), 'YYYYMMDD'),
    organization: columnMap.organization !== undefined ? fields[columnMap.organization] : ''
  };

  // 性別コード変換（男/女 → M/F）
  if (patientInfo.gender === '男') {
    patientInfo.genderCode = 'M';
  } else if (patientInfo.gender === '女') {
    patientInfo.genderCode = 'F';
  } else {
    patientInfo.genderCode = patientInfo.gender;
  }

  // 検査結果を抽出
  const testResults = [];

  // CSVカラム名 → 判定基準キーのマッピング
  const fieldToCode = {
    'fbs': 'FASTING_GLUCOSE',
    'hba1c': 'HBA1C',
    'tg': 'TRIGLYCERIDES',
    'hdl_c': 'HDL_CHOLESTEROL',
    'ldl_c': 'LDL_CHOLESTEROL',
    'got': 'AST_GOT',
    'gpt': 'ALT_GPT',
    'gamma_gpt': 'GAMMA_GTP',
    'tp': 'TOTAL_PROTEIN',
    'ua': 'URIC_ACID',
    'cre': 'CREATININE',
    'wbc': 'WBC',
    'hb': 'HEMOGLOBIN',
    'plt': 'PLT',
    'bmi': 'BMI',
    'waist': 'WAIST',
    'sbp': 'BLOOD_PRESSURE_SYS',
    'dbp': 'BLOOD_PRESSURE_DIA'
  };

  for (const [fieldKey, codeKey] of Object.entries(fieldToCode)) {
    const colIndex = columnMap[fieldKey];
    if (colIndex !== undefined) {
      const value = fields[colIndex];
      if (value && value !== '') {
        testResults.push({
          code: codeKey,
          value: value,
          flag: '',
          comment: ''
        });
      }
    }
  }

  return {
    patientInfo: patientInfo,
    testResults: testResults
  };
}

/**
 * ファイル内容を読み込み（エンコーディング対応）
 * @param {File} file - ファイル
 * @returns {string} ファイル内容
 */
function readFileContent(file) {
  const blob = file.getBlob();

  // まずShift_JISで試行
  try {
    const content = blob.getDataAsString('Shift_JIS');
    // 文字化けチェック（日本語が含まれるべき）
    if (isValidJapanese(content)) {
      return content;
    }
  } catch (e) {
    // Shift_JIS失敗
  }

  // UTF-8で試行
  try {
    const content = blob.getDataAsString('UTF-8');
    return content;
  } catch (e) {
    throw new Error('ファイルのエンコーディングを検出できません: ' + file.getName());
  }
}

/**
 * 日本語として有効かチェック
 * @param {string} content - 内容
 * @returns {boolean}
 */
function isValidJapanese(content) {
  // 文字化け時に現れる特殊文字がないかチェック
  return !content.includes('�');
}

/**
 * 1行を解析
 * @param {string} line - CSV行
 * @returns {Object|null} 解析結果
 */
function parseLine(line) {
  const fields = line.split(',');

  if (fields.length < CONFIG.CSV.MIN_FIELDS) {
    return null;
  }

  const patientInfo = extractPatientInfo(fields);
  const testResults = extractTestResults(fields);

  return {
    patientInfo: patientInfo,
    testResults: testResults
  };
}

/**
 * 患者情報を抽出
 * BML結果CSVフォーマット:
 * 0: 施設コード, 1: 患者ID, 2: 受診日(YYYYMMDD), 3: 時間, 4: 空,
 * 5: 依頼番号, 6: 性別(1=男,2=女), 7-9: 空, 10: 検体番号
 * 11以降: 検査コード,結果値,フラグ,コメント の繰り返し
 *
 * @param {Array<string>} fields - フィールド配列
 * @returns {Object} 患者情報
 */
function extractPatientInfo(fields) {
  return {
    facilityCode: fields[0] || '',
    requestId: fields[1] || '',
    examDate: fields[2] || '',
    time: fields[3] || '',
    insuranceNo: fields.length > 5 ? fields[5] : '',
    gender: fields.length > 6 ? fields[6] : ''
  };
}

/**
 * 検査結果を抽出
 * @param {Array<string>} fields - フィールド配列
 * @returns {Array<Object>} 検査結果配列
 */
function extractTestResults(fields) {
  const testResults = [];
  let i = CONFIG.CSV.TEST_START_INDEX;

  while (i < fields.length - 1) {
    const code = (fields[i] || '').trim();

    // コードが数字でない場合はスキップ
    if (!code || !/^\d+$/.test(code)) {
      i++;
      continue;
    }

    const result = extractSingleResult(fields, i);
    if (result) {
      testResults.push(result);
    }

    i += CONFIG.CSV.TEST_FIELD_COUNT;
  }

  return testResults;
}

/**
 * 単一の検査結果を抽出
 * @param {Array<string>} fields - フィールド配列
 * @param {number} index - 開始インデックス
 * @returns {Object|null} 検査結果
 */
function extractSingleResult(fields, index) {
  const code = (fields[index] || '').trim();
  const value = (fields[index + 1] || '').trim();
  const flag = fields.length > index + 2 ? (fields[index + 2] || '').trim() : '';
  const comment = fields.length > index + 3 ? (fields[index + 3] || '').trim() : '';

  if (!value) {
    return null;
  }

  return {
    code: code,
    value: value,
    flag: flag,
    comment: comment
  };
}

/**
 * CSVフォルダ内の新規ファイルを検索（サブフォルダ対応）
 * @param {number} maxFiles - 最大取得件数（デフォルト10件）
 * @returns {Array<File>} 未処理CSVファイルの配列
 */
function findNewCsvFiles(maxFiles = 10) {
  const folder = getCsvFolder();
  const newFiles = [];

  // 再帰的にCSVファイルを検索
  findCsvFilesRecursive(folder, newFiles, 0, maxFiles);

  logInfo(`検出CSVファイル数: ${newFiles.length}件（上限: ${maxFiles}件）`);
  return newFiles;
}

/**
 * フォルダ内のCSVファイルを再帰的に検索
 * @param {Folder} folder - 検索対象フォルダ
 * @param {Array<File>} results - 結果を格納する配列
 * @param {number} depth - 現在の深さ（無限ループ防止）
 * @param {number} maxFiles - 最大取得件数
 */
function findCsvFilesRecursive(folder, results, depth = 0, maxFiles = 10) {
  // 深さ制限（5階層まで）または件数上限
  if (depth > 5 || results.length >= maxFiles) {
    return;
  }

  // 現在のフォルダ内のCSVファイルを検索
  const files = folder.getFilesByType(MimeType.CSV);
  while (files.hasNext() && results.length < maxFiles) {
    const file = files.next();
    if (isResultCsvFile(file) && !isFileProcessed(file)) {
      results.push(file);
    }
  }

  // 拡張子が.csvだがMIMEタイプが異なるファイルも検索
  const allFiles = folder.getFiles();
  while (allFiles.hasNext() && results.length < maxFiles) {
    const file = allFiles.next();
    const name = file.getName().toLowerCase();
    if (name.endsWith('.csv') && isResultCsvFile(file) && !isFileProcessed(file)) {
      // 重複チェック
      const exists = results.some(f => f.getId() === file.getId());
      if (!exists) {
        results.push(file);
      }
    }
  }

  // サブフォルダを再帰的に検索
  const subFolders = folder.getFolders();
  while (subFolders.hasNext() && results.length < maxFiles) {
    const subFolder = subFolders.next();
    findCsvFilesRecursive(subFolder, results, depth + 1, maxFiles);
  }
}

/**
 * 結果データCSVファイルかどうかを判定
 * - 結果データ: XXXXXXXX_OXXXXXXX.csv（数字_O数字.csv）→ 処理対象
 * - 依頼データ: kensa_irai_*.csv → スキップ（結果がない）
 * - その他: *.txt, *.log, *.mdb など → スキップ
 *
 * @param {File} file - ファイル
 * @returns {boolean} 結果データCSVファイルならtrue
 */
function isResultCsvFile(file) {
  const name = file.getName();

  // [済]プレフィックスを除去して判定
  const baseName = name.replace(/^\[済\]/, '');

  // 結果データファイルパターン: 数字_O数字.csv (例: 10030937_O6319101.csv)
  if (/^\d+_O\d+\.csv$/i.test(baseName)) {
    return true;
  }

  // 依頼データ（kensa_irai_*）はスキップ
  if (baseName.startsWith('kensa_irai_')) {
    return false;
  }

  // それ以外のCSVファイルは一応処理対象とする
  if (baseName.toLowerCase().endsWith('.csv')) {
    // ただしログファイルなどは除外
    if (baseName.toLowerCase().includes('log') ||
        baseName.toLowerCase().includes('info') ||
        baseName.toLowerCase().includes('report')) {
      return false;
    }
    return true;
  }

  return false;
}

/**
 * CSVデータをスプレッドシートに保存
 * @param {Object} patientData - 患者データ
 * @returns {string} 受診ID
 */
function savePatientData(patientData) {
  const patientInfo = patientData.patientInfo;
  const testResults = patientData.testResults;

  // 受診ID生成
  const examDate = parseExamDate(patientInfo.examDate);
  const patientId = patientInfo.requestId || generatePatientId(examDate);

  // 性別変換
  const genderCode = patientInfo.gender;
  const gender = GENDER_CODE_TO_INTERNAL[genderCode] || 'M';
  const genderDisplay = GENDER_TRANSFORMS[genderCode] || '男';

  // 企業情報の登録・取得（CSVに企業名がある場合）
  let companyInfo = { companyId: '', companyName: '' };
  if (patientInfo.organization) {
    companyInfo = registerOrGetCompany(patientInfo.organization);
  }

  // 受診者マスタに保存
  saveToPatientMaster({
    patientId: patientId,
    status: CONFIG.STATUS.INPUT,
    examDate: examDate,
    gender: genderDisplay,
    csvImportDate: new Date(),
    companyId: companyInfo.companyId,
    companyName: companyInfo.companyName
  });

  // 血液検査に保存
  saveToBloodTest(patientId, testResults, gender);

  return patientId;
}

/**
 * 受診日を解析
 * @param {string} dateStr - 日付文字列（YYYYMMDD）
 * @returns {Date}
 */
function parseExamDate(dateStr) {
  if (!dateStr || dateStr.length !== 8) {
    return new Date();
  }

  const year = parseInt(dateStr.slice(0, 4));
  const month = parseInt(dateStr.slice(4, 6)) - 1;
  const day = parseInt(dateStr.slice(6, 8));

  return new Date(year, month, day);
}

/**
 * 受診者マスタに保存
 * @param {Object} data - 保存データ
 */
function saveToPatientMaster(data) {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);

  // 既存レコードチェック
  const lastRow = sheet.getLastRow();
  const existingRow = findPatientRow(sheet, data.patientId, lastRow);

  if (existingRow > 0) {
    // 既存レコード更新
    updatePatientRow(sheet, existingRow, data);
  } else {
    // 新規レコード追加
    appendPatientRow(sheet, data);
  }
}

/**
 * 患者行を検索
 * @param {Sheet} sheet - シート
 * @param {string} patientId - 受診ID
 * @param {number} lastRow - 最終行
 * @returns {number} 行番号（見つからない場合は0）
 */
function findPatientRow(sheet, patientId, lastRow) {
  if (lastRow < 2) return 0;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const searchId = String(patientId).trim();  // 文字列に変換して比較

  for (let i = 0; i < ids.length; i++) {
    const cellId = String(ids[i][0]).trim();  // セル値も文字列に変換
    if (cellId === searchId) {
      return i + 2;
    }
  }
  return 0;
}

/**
 * 患者行を更新
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号
 * @param {Object} data - データ
 */
function updatePatientRow(sheet, row, data) {
  sheet.getRange(row, 2).setValue(data.status);
  sheet.getRange(row, 13).setValue(data.csvImportDate);
  sheet.getRange(row, 14).setValue(new Date());
}

/**
 * 患者行を追加
 * @param {Sheet} sheet - シート
 * @param {Object} data - データ
 */
function appendPatientRow(sheet, data) {
  const newRow = [
    data.patientId,           // A: 受診ID
    data.status,              // B: ステータス
    data.examDate,            // C: 受診日
    '',                       // D: 氏名
    '',                       // E: カナ
    data.gender,              // F: 性別
    '',                       // G: 生年月日
    '',                       // H: 年齢
    '',                       // I: 受診コース
    data.companyName || '',   // J: 事業所名
    '',                       // K: 所属
    '',                       // L: 総合判定
    data.csvImportDate,       // M: CSV取込日時
    new Date(),               // N: 最終更新日時
    '',                       // O: 出力日時
    data.companyId || ''      // P: 企業ID（新規追加）
  ];

  sheet.appendRow(newRow);
}

/**
 * 血液検査に保存
 * @param {string} patientId - 受診ID
 * @param {Array<Object>} testResults - 検査結果
 * @param {string} gender - 性別（M/F）
 */
function saveToBloodTest(patientId, testResults, gender) {
  const sheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);

  // 既存レコードチェック
  const lastRow = sheet.getLastRow();
  let row = findPatientRow(sheet, patientId, lastRow);

  if (row === 0) {
    // 新規行追加
    row = lastRow + 1;
    sheet.getRange(row, 1).setValue(patientId);
  }

  // コード→列マッピング
  const codeToColumn = {
    '0000301': 2,   // WBC
    '0000302': 3,   // RBC
    '0000303': 4,   // Hb
    '0000304': 5,   // Ht
    '0000308': 6,   // PLT
    '0000305': 7,   // MCV
    '0000306': 8,   // MCH
    '0000307': 9,   // MCHC
    '0000401': 10,  // TP
    '0000402': 11,  // ALB
    '0000472': 12,  // T-Bil
    '0000481': 13,  // AST
    '0000482': 14,  // ALT
    '0000484': 15,  // γ-GTP
    '0000405': 18,  // BUN
    '0000413': 19,  // Cr
    '0002696': 20,  // eGFR
    '0000407': 21,  // UA
    '0000450': 22,  // TC
    '0000460': 23,  // HDL-C
    '0000410': 24,  // LDL-C
    '0000454': 25,  // TG
    '0000503': 26,  // FBS
    '0003317': 27,  // HbA1c
    '0000658': 28   // CRP
  };

  // 検査結果を書き込み
  for (const result of testResults) {
    const col = codeToColumn[result.code];
    if (col) {
      // 数値変換を試みる
      const numValue = toNumber(result.value);
      sheet.getRange(row, col).setValue(numValue !== null ? numValue : result.value);
    }
  }
}

/**
 * CSVファイルを処理
 * @param {File} file - CSVファイル
 * @returns {Object} 処理結果
 */
function processCsvFile(file) {
  const results = {
    success: false,
    patientIds: [],
    errors: []
  };

  try {
    logInfo('CSV処理開始: ' + file.getName());

    const patientDataList = parseCSV(file.getId());

    for (const patientData of patientDataList) {
      try {
        const patientId = savePatientData(patientData);
        results.patientIds.push(patientId);

        // 判定処理を実行
        processJudgments(patientId, patientData);

        // 所見生成
        generatePatientFindings(patientId);

      } catch (e) {
        results.errors.push(`患者処理エラー: ${e.message}`);
        logError('processCsvFile', e);
      }
    }

    // 処理済みマーク
    markFileAsProcessed(file);

    results.success = true;
    logInfo('CSV処理完了: ' + results.patientIds.length + '名');

  } catch (e) {
    results.errors.push(`CSV解析エラー: ${e.message}`);
    logError('processCsvFile', e);
  }

  return results;
}
