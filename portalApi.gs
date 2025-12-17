/**
 * portalApi.gs - Webアプリポータル用 API関数
 * CDmedical 健診結果管理システム
 *
 * フロントエンドから google.script.run 経由で呼び出される関数群
 */

// ============================================================
// スプレッドシート接続
// ============================================================

/**
 * スプレッドシートを取得（Webアプリ対応）
 * @returns {Spreadsheet} スプレッドシート
 */
function getPortalSpreadsheet() {
  // 確実に動作するようにスプレッドシートIDを直接指定
  const SPREADSHEET_ID = '16KtctyT2gd7oJZdcu84kUtuP-D9jB9KtLxxzxXx_wdk';

  // 方法1: 直接IDで取得（最も確実）
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    console.error('直接ID取得失敗:', e);
  }

  // 方法2: スクリプトがバインドされているスプレッドシートを取得
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {
    // Webアプリではエラーになる場合がある
  }

  // 方法3: PropertiesServiceからIDを取得
  try {
    const props = PropertiesService.getScriptProperties();
    const ssId = props.getProperty('SPREADSHEET_ID');
    if (ssId) {
      return SpreadsheetApp.openById(ssId);
    }
  } catch (e) {
    // プロパティが設定されていない場合
  }

  throw new Error('スプレッドシートに接続できません。');
}

/**
 * スプレッドシートIDを設定（初回のみ実行）
 * GASエディタから手動で実行してください
 */
function setupSpreadsheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SPREADSHEET_ID', ss.getId());
    console.log('スプレッドシートID設定完了: ' + ss.getId());
    return { success: true, id: ss.getId() };
  }
  return { success: false, error: 'スプレッドシートが見つかりません' };
}

// ============================================================
// ユーティリティ関数
// ============================================================

/**
 * 文字列を正規化（スペースを統一、大文字小文字を統一）
 * @param {*} str - 入力文字列
 * @returns {string} 正規化された文字列
 */
function normalizeString(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\s　]+/g, ' ')  // 全角・半角スペースを半角スペースに統一
    .trim()
    .toLowerCase();
}

/**
 * カナを正規化（全角カナに統一）
 * @param {*} str - 入力文字列
 * @returns {string} 正規化されたカナ
 */
function normalizeKana(str) {
  if (str === null || str === undefined) return '';
  let result = String(str).replace(/[\s　]+/g, ' ').trim();
  // 半角カナを全角カナに変換
  const kanaMap = {
    'ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ',
    'ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ','ｺ':'コ',
    'ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ',
    'ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ','ﾄ':'ト',
    'ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ',
    'ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ','ﾎ':'ホ',
    'ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ',
    'ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ',
    'ﾗ':'ラ','ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ',
    'ﾜ':'ワ','ｦ':'ヲ','ﾝ':'ン',
    'ｧ':'ァ','ｨ':'ィ','ｩ':'ゥ','ｪ':'ェ','ｫ':'ォ',
    'ｯ':'ッ','ｬ':'ャ','ｭ':'ュ','ｮ':'ョ',
    'ﾞ':'゛','ﾟ':'゜','ｰ':'ー'
  };
  for (const [half, full] of Object.entries(kanaMap)) {
    result = result.replace(new RegExp(half, 'g'), full);
  }
  return result;
}

// ============================================================
// 受診者管理 API
// ============================================================

/**
 * 受診者を検索（修正版 v3 - シリアライズ対応）
 *
 * 【修正ポイント】
 * 1. Dateオブジェクトを文字列に変換してからreturn
 * 2. 値をすべてプリミティブ型に変換
 * 3. undefinedをnullまたは空文字に変換
 *
 * @param {Object} criteria - 検索条件
 * @returns {Object} 検索結果
 */
function portalSearchPatients(criteria) {
  const debugInfo = {
    version: '2025-12-16-v3-serializable',
    criteriaReceived: criteria ? JSON.stringify(criteria) : 'null',
    timestamp: new Date().toISOString()
  };

  try {
    const ss = getPortalSpreadsheet();
    debugInfo.spreadsheetId = ss.getId();

    const sheet = ss.getSheetByName('受診者マスタ');
    if (!sheet) {
      return { success: false, error: '受診者マスタシートが見つかりません', data: [], _debug: debugInfo };
    }
    debugInfo.sheetFound = true;

    const data = sheet.getDataRange().getValues();
    debugInfo.totalRows = data.length;
    debugInfo.firstRowId = data.length > 1 ? String(data[1][0]) : 'NO_DATA';

    const results = [];

    // 検索条件を正規化（nullチェック追加）
    const searchId = criteria && criteria.patientId ? normalizeString(criteria.patientId) : '';
    const searchName = criteria && criteria.name ? normalizeString(criteria.name) : '';
    const searchKana = criteria && criteria.kana ? normalizeKana(criteria.kana) : '';

    debugInfo.normalizedCriteria = { searchId: searchId, searchName: searchName, searchKana: searchKana };

    // 列インデックス:
    // 0:受診ID, 1:ステータス, 2:受診日, 3:氏名, 4:カナ, 5:性別, 6:生年月日, 7:年齢,
    // 8:受診コース, 9:事業所名, 10:所属, 11:総合判定
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 空行スキップ
      if (!row[0]) continue;

      // 検索条件でフィルタ
      let match = true;

      if (searchId) {
        const rowId = normalizeString(row[0]);
        if (!rowId.includes(searchId)) {
          match = false;
        }
      }
      if (searchName && match) {
        const rowName = normalizeString(row[3]);
        if (!rowName.includes(searchName)) {
          match = false;
        }
      }
      if (searchKana && match) {
        const rowKana = normalizeKana(row[4]);
        if (!rowKana.includes(searchKana)) {
          match = false;
        }
      }

      if (match) {
        // ★★★ 重要: すべての値をシリアライズ可能な形式に変換 ★★★
        results.push({
          '受診ID': safeString(row[0]),
          'ステータス': safeString(row[1]),
          '受診日': formatDateToString(row[2]),
          '氏名': safeString(row[3]),
          'カナ': safeString(row[4]),
          '性別': safeString(row[5]),
          '生年月日': formatDateToString(row[6]),
          '年齢': safeNumber(row[7]),
          '受診コース': safeString(row[8]),
          '事業所名': safeString(row[9]),
          '所属': safeString(row[10]),
          '総合判定': safeString(row[11]),
          _rowIndex: i + 1
        });
      }

      // 最大件数制限
      if (results.length >= 100) break;
    }

    debugInfo.resultsCount = results.length;

    return {
      success: true,
      data: results,
      count: results.length,
      _debug: debugInfo
    };

  } catch (error) {
    console.error('portalSearchPatients error:', error);
    debugInfo.error = String(error.message);
    return {
      success: false,
      error: String(error.message),
      data: [],
      _debug: debugInfo
    };
  }
}

// ============================================================
// シリアライズ用ヘルパー関数
// ============================================================

/**
 * 値を安全な文字列に変換
 * @param {*} value - 任意の値
 * @returns {string} 文字列
 */
function safeString(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return formatDateToString(value);
  return String(value);
}

/**
 * 値を安全な数値に変換
 * @param {*} value - 任意の値
 * @returns {number|string} 数値または空文字
 */
function safeNumber(value) {
  if (value === null || value === undefined || value === '') return '';
  const num = Number(value);
  return isNaN(num) ? String(value) : num;
}

/**
 * Dateオブジェクトを文字列に変換
 * @param {*} value - 日付値
 * @returns {string} 日付文字列 (YYYY/MM/DD)
 */
function formatDateToString(value) {
  if (!value) return '';

  // 既に文字列の場合はそのまま返す
  if (typeof value === 'string') return value;

  // Dateオブジェクトの場合
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';

    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const day = String(value.getDate()).padStart(2, '0');
    return year + '/' + month + '/' + day;
  }

  // 数値の場合
  if (typeof value === 'number') {
    return String(value);
  }

  return String(value);
}

/**
 * 受診者を新規登録
 * @param {Object} patientData - 受診者データ
 * @returns {Object} 登録結果
 */
function portalRegisterPatient(patientData) {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('受診者マスタ');
    if (!sheet) {
      return { success: false, error: '受診者マスタシートが見つかりません' };
    }

    // 重複チェック
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (String(existingData[i][0]) === String(patientData.patientId)) {
        return { success: false, error: '受診ID「' + patientData.patientId + '」は既に登録されています' };
      }
    }

    // 列順序（実際のスプレッドシート構造）:
    // 受診ID, ステータス, 受診日, 氏名, カナ, 性別, 生年月日, 年齢, 受診コース, 事業所名, 所属, 総合判定, CSV取込日時, 最終更新日時, 出力日時
    const now = new Date();
    const newRow = [
      patientData.patientId || '',        // 受診ID
      patientData.status || '入力中',      // ステータス
      patientData.examDate || '',          // 受診日
      patientData.name || '',              // 氏名
      patientData.nameKana || '',          // カナ
      patientData.gender || '',            // 性別
      patientData.birthDate || '',         // 生年月日
      patientData.age || '',               // 年齢
      patientData.course || '',            // 受診コース
      patientData.company || '',           // 事業所名
      patientData.department || '',        // 所属
      patientData.overallJudgment || '',   // 総合判定
      '',                                   // CSV取込日時
      now,                                  // 最終更新日時
      ''                                    // 出力日時
    ];

    sheet.appendRow(newRow);

    return { success: true, message: '受診者を登録しました', patientId: patientData.patientId };
  } catch (error) {
    console.error('portalRegisterPatient error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 受診者情報を取得（ID指定）
 * @param {string} patientId - 患者ID
 * @returns {Object} 受診者情報
 */
function portalGetPatient(patientId) {
  try {
    const result = portalSearchPatients({ patientId: patientId });
    if (result.success && result.data.length > 0) {
      return { success: true, data: result.data[0] };
    }
    return { success: false, error: '受診者が見つかりません' };
  } catch (error) {
    console.error('portalGetPatient error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 受診者情報を更新
 * @param {string} patientId - 患者ID
 * @param {Object} updateData - 更新データ
 * @returns {Object} 更新結果
 */
function portalUpdatePatient(patientId, updateData) {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('受診者マスタ');
    if (!sheet) {
      return { success: false, error: '受診者マスタシートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();

    // 列インデックス（実際のスプレッドシート構造）:
    // 0:受診ID, 1:ステータス, 2:受診日, 3:氏名, 4:カナ, 5:性別, 6:生年月日, 7:年齢,
    // 8:受診コース, 9:事業所名, 10:所属, 11:総合判定, 12:CSV取込日時, 13:最終更新日時, 14:出力日時
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(patientId)) {
        // 該当行を更新
        if (updateData['ステータス'] !== undefined) sheet.getRange(i + 1, 2).setValue(updateData['ステータス']);
        if (updateData['受診日'] !== undefined) sheet.getRange(i + 1, 3).setValue(updateData['受診日'] ? new Date(updateData['受診日']) : '');
        if (updateData['氏名'] !== undefined) sheet.getRange(i + 1, 4).setValue(updateData['氏名']);
        if (updateData['カナ'] !== undefined) sheet.getRange(i + 1, 5).setValue(updateData['カナ']);
        if (updateData['性別'] !== undefined) sheet.getRange(i + 1, 6).setValue(updateData['性別']);
        if (updateData['生年月日'] !== undefined) sheet.getRange(i + 1, 7).setValue(updateData['生年月日'] ? new Date(updateData['生年月日']) : '');
        if (updateData['年齢'] !== undefined) sheet.getRange(i + 1, 8).setValue(updateData['年齢']);
        if (updateData['受診コース'] !== undefined) sheet.getRange(i + 1, 9).setValue(updateData['受診コース']);
        if (updateData['事業所名'] !== undefined) sheet.getRange(i + 1, 10).setValue(updateData['事業所名']);
        if (updateData['所属'] !== undefined) sheet.getRange(i + 1, 11).setValue(updateData['所属']);
        if (updateData['総合判定'] !== undefined) sheet.getRange(i + 1, 12).setValue(updateData['総合判定']);

        // 最終更新日時を更新
        sheet.getRange(i + 1, 14).setValue(new Date());

        return { success: true, message: '更新しました' };
      }
    }

    return { success: false, error: '受診者が見つかりません' };
  } catch (error) {
    console.error('portalUpdatePatient error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 受診者詳細を更新（ポータルUI用）
 * js.html から呼び出される
 * @param {string} patientId - 受診ID
 * @param {Object} updateData - 更新データ
 * @returns {Object} 更新結果
 */
function portalUpdatePatientDetail(patientId, updateData) {
  return portalUpdatePatient(patientId, updateData);
}

// ============================================================
// 検査結果入力 API
// ============================================================

/**
 * 血液検査結果を取得
 * @param {string} patientId - 患者ID
 * @returns {Object} 血液検査結果
 */
function portalGetBloodTestResults(patientId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません', data: null };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('患者ID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(patientId)) {
        const result = {};
        headers.forEach((header, idx) => {
          result[header] = data[i][idx];
        });
        result._rowIndex = i + 1;
        return { success: true, data: result };
      }
    }

    return { success: true, data: null, message: '検査結果がありません' };
  } catch (error) {
    console.error('portalGetBloodTestResults error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 身体計測結果を保存
 * @param {string} patientId - 患者ID
 * @param {Object} measurements - 身体計測データ
 * @returns {Object} 保存結果
 */
function portalSaveBodyMeasurements(patientId, measurements) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('患者ID');

    // 身体計測項目のマッピング
    const measurementFields = {
      '身長': measurements.height,
      '体重': measurements.weight,
      'BMI': measurements.bmi,
      '収縮期血圧': measurements.systolicBP,
      '拡張期血圧': measurements.diastolicBP,
      '脈拍': measurements.pulse,
      '視力_右': measurements.visionRight,
      '視力_左': measurements.visionLeft,
      '視力_矯正': measurements.visionCorrected,
      '聴力_右_1000Hz': measurements.hearingRight1000,
      '聴力_右_4000Hz': measurements.hearingRight4000,
      '聴力_左_1000Hz': measurements.hearingLeft1000,
      '聴力_左_4000Hz': measurements.hearingLeft4000,
      '心電図所見': measurements.ecgFindings,
      '心電図コメント': measurements.ecgComment,
      '胸部X線所見': measurements.chestXrayFindings,
      '胸部X線コメント': measurements.chestXrayComment,
      '更新日時': new Date()
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(patientId)) {
        // 該当行を更新
        Object.keys(measurementFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          if (colIdx >= 0 && measurementFields[fieldName] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(measurementFields[fieldName]);
          }
        });
        return { success: true, message: '身体計測結果を保存しました' };
      }
    }

    // 該当行がない場合は新規行を追加
    const newRow = headers.map(header => {
      if (header === '患者ID') return patientId;
      if (measurementFields.hasOwnProperty(header)) return measurementFields[header];
      return '';
    });
    sheet.appendRow(newRow);

    return { success: true, message: '身体計測結果を新規登録しました' };
  } catch (error) {
    console.error('portalSaveBodyMeasurements error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 血液検査結果を保存
 * @param {string} patientId - 患者ID
 * @param {Object} bloodData - 血液検査データ（キー: 項目名, 値: 検査値）
 * @returns {Object} 保存結果
 */
function portalSaveBloodTestResults(patientId, bloodData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('患者ID');

    // 血液検査項目のマッピング（bloodDataのキーはそのまま列名として使用）
    const bloodFields = {
      ...bloodData,
      '血液検査入力日': new Date(),
      '更新日時': new Date()
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(patientId)) {
        // 該当行を更新
        Object.keys(bloodFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          if (colIdx >= 0 && bloodFields[fieldName] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(bloodFields[fieldName]);
          }
        });
        return { success: true, message: '血液検査結果を保存しました' };
      }
    }

    // 該当行がない場合は新規行を追加
    const newRow = headers.map(header => {
      if (header === '患者ID') return patientId;
      if (bloodFields.hasOwnProperty(header)) return bloodFields[header];
      return '';
    });
    sheet.appendRow(newRow);

    return { success: true, message: '血液検査結果を新規登録しました' };
  } catch (error) {
    console.error('portalSaveBloodTestResults error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 検査結果を保存（設計書準拠版）
 * @param {string} patientId - 患者ID
 * @param {Object} inputData - 検査データ（キー: 項目名, 値: 検査値）
 * @returns {Object} 保存結果
 */
function portalSaveInputResults(patientId, inputData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('患者ID');

    // 入力項目のマッピング（inputDataのキーはそのまま列名として使用）
    const inputFields = {
      ...inputData,
      '結果入力日': new Date(),
      '更新日時': new Date()
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(patientId)) {
        // 該当行を更新
        Object.keys(inputFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          if (colIdx >= 0 && inputFields[fieldName] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(inputFields[fieldName]);
          }
        });
        return { success: true, message: '検査結果を保存しました' };
      }
    }

    // 該当行がない場合は新規行を追加
    const newRow = headers.map(header => {
      if (header === '患者ID') return patientId;
      if (inputFields.hasOwnProperty(header)) return inputFields[header];
      return '';
    });
    sheet.appendRow(newRow);

    return { success: true, message: '検査結果を新規登録しました' };
  } catch (error) {
    console.error('portalSaveInputResults error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 所見を保存
 * @param {string} patientId - 患者ID
 * @param {Object} findings - 所見データ
 * @returns {Object} 保存結果
 */
function portalSaveFindings(patientId, findings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('患者ID');

    // 所見項目のマッピング
    const findingFields = {
      '総合判定': findings.overallJudgment,
      '総合所見': findings.overallFindings,
      '医師コメント': findings.doctorComment,
      '生活指導': findings.lifestyleGuidance,
      '要精密検査項目': findings.needsDetailedExam,
      '要治療項目': findings.needsTreatment,
      '所見入力日': new Date(),
      '所見入力者': Session.getActiveUser().getEmail()
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(patientId)) {
        Object.keys(findingFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          if (colIdx >= 0 && findingFields[fieldName] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(findingFields[fieldName]);
          }
        });
        return { success: true, message: '所見を保存しました' };
      }
    }

    return { success: false, error: '検査結果レコードが見つかりません。先にCSV取込を行ってください。' };
  } catch (error) {
    console.error('portalSaveFindings error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// CSV取込 API
// ============================================================

/**
 * CSVデータを取り込み
 * @param {string} csvContent - CSVコンテンツ
 * @param {Object} options - 取込オプション
 * @returns {Object} 取込結果
 */
function portalImportCsv(csvContent, options) {
  try {
    // 既存のprocessCsvContent関数があれば使用
    if (typeof processCsvContent === 'function') {
      return processCsvContent(csvContent, options);
    }

    // CSVをパース
    const lines = csvContent.split(/\r?\n/).filter(line => line.trim());
    if (lines.length < 2) {
      return { success: false, error: 'CSVデータが空または不正です' };
    }

    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    const idColIndex = headers.findIndex(h => h.includes('ID') || h.includes('患者'));

    if (idColIndex < 0) {
      return { success: false, error: 'ID列が見つかりません' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const existingData = sheet.getDataRange().getValues();
    const existingHeaders = existingData[0];
    const existingIdCol = existingHeaders.indexOf('患者ID');
    const existingIds = new Set(existingData.slice(1).map(row => String(row[existingIdCol])));

    let imported = 0;
    let updated = 0;
    let errors = [];

    for (let i = 1; i < lines.length; i++) {
      try {
        const values = lines[i].split(',').map(v => v.trim().replace(/^"|"$/g, ''));
        const patientId = values[idColIndex];

        if (!patientId) {
          errors.push({ row: i + 1, error: 'IDが空です' });
          continue;
        }

        if (existingIds.has(String(patientId))) {
          // 既存レコードを更新
          for (let j = 1; j < existingData.length; j++) {
            if (String(existingData[j][existingIdCol]) === String(patientId)) {
              headers.forEach((header, idx) => {
                const colIdx = existingHeaders.indexOf(header);
                if (colIdx >= 0 && values[idx]) {
                  sheet.getRange(j + 1, colIdx + 1).setValue(values[idx]);
                }
              });
              break;
            }
          }
          updated++;
        } else {
          // 新規行を追加
          const newRow = existingHeaders.map(header => {
            const idx = headers.indexOf(header);
            return idx >= 0 ? values[idx] : '';
          });
          newRow[existingIdCol] = patientId;
          sheet.appendRow(newRow);
          existingIds.add(String(patientId));
          imported++;
        }
      } catch (rowError) {
        errors.push({ row: i + 1, error: rowError.message });
      }
    }

    return {
      success: true,
      imported: imported,
      updated: updated,
      errors: errors,
      total: lines.length - 1,
      message: `取込完了: 新規${imported}件, 更新${updated}件` + (errors.length > 0 ? `, エラー${errors.length}件` : '')
    };
  } catch (error) {
    console.error('portalImportCsv error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 取込済み患者一覧を取得（進捗状況付き）
 * @param {Object} options - フィルタオプション
 * @returns {Object} 患者一覧と進捗状況
 */
function portalGetPatientProgress(options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultSheet = ss.getSheetByName('検査結果');
    const patientSheet = ss.getSheetByName('受診者マスタ');

    if (!resultSheet) {
      return { success: false, error: '検査結果シートが見つかりません', data: [] };
    }

    const resultData = resultSheet.getDataRange().getValues();
    const resultHeaders = resultData[0];

    // 進捗判定用の列インデックス
    const idCol = resultHeaders.indexOf('患者ID');
    const bloodCols = resultHeaders.filter(h => h.includes('血液') || h.includes('HbA1c') || h.includes('コレステロール'));
    const bodyCols = ['身長', '体重', 'BMI', '収縮期血圧'].map(h => resultHeaders.indexOf(h)).filter(i => i >= 0);
    const findingCol = resultHeaders.indexOf('総合判定');

    const patients = [];

    // 受診者マスタから名前を取得するためのマップ
    const patientNames = new Map();
    if (patientSheet) {
      const patientData = patientSheet.getDataRange().getValues();
      const patientHeaders = patientData[0];
      const pIdCol = patientHeaders.indexOf('患者ID');
      const pNameCol = patientHeaders.indexOf('氏名');

      for (let i = 1; i < patientData.length; i++) {
        patientNames.set(String(patientData[i][pIdCol]), patientData[i][pNameCol]);
      }
    }

    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];
      const patientId = row[idCol];

      if (!patientId) continue;

      // 日付フィルタ
      if (options && options.fromDate) {
        const examDate = row[resultHeaders.indexOf('検査日')];
        if (examDate && new Date(examDate) < new Date(options.fromDate)) continue;
      }

      // 進捗状況を計算
      const hasBlood = bloodCols.some(colName => {
        const idx = resultHeaders.indexOf(colName);
        return idx >= 0 && row[idx];
      }) || row[resultHeaders.indexOf('血液検査取込日')];

      const hasBody = bodyCols.some(idx => row[idx]);
      const hasFindings = findingCol >= 0 && row[findingCol];

      patients.push({
        patientId: patientId,
        name: patientNames.get(String(patientId)) || '（未登録）',
        examDate: row[resultHeaders.indexOf('検査日')] || '',
        bloodComplete: !!hasBlood,
        bodyComplete: !!hasBody,
        findingsComplete: !!hasFindings,
        allComplete: hasBlood && hasBody && hasFindings,
        _rowIndex: i + 1
      });
    }

    // 未完了を優先してソート
    patients.sort((a, b) => {
      if (a.allComplete !== b.allComplete) return a.allComplete ? 1 : -1;
      return 0;
    });

    return {
      success: true,
      data: patients,
      count: patients.length,
      stats: {
        total: patients.length,
        complete: patients.filter(p => p.allComplete).length,
        bloodPending: patients.filter(p => !p.bloodComplete).length,
        bodyPending: patients.filter(p => !p.bodyComplete).length,
        findingsPending: patients.filter(p => !p.findingsComplete).length
      }
    };
  } catch (error) {
    console.error('portalGetPatientProgress error:', error);
    return { success: false, error: error.message, data: [] };
  }
}

// ============================================================
// 帳票出力 API
// ============================================================

/**
 * 帳票を生成
 * @param {string} patientId - 患者ID
 * @param {string} reportType - 帳票種類
 * @returns {Object} 生成結果（URL等）
 */
function portalGenerateReport(patientId, reportType) {
  try {
    // 既存の帳票出力関数を呼び出し
    if (typeof generateReport === 'function') {
      return generateReport(patientId, reportType);
    }

    // 基本実装（プレースホルダー）
    return {
      success: true,
      message: '帳票生成は別途実装が必要です',
      patientId: patientId,
      reportType: reportType
    };
  } catch (error) {
    console.error('portalGenerateReport error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 一括帳票出力
 * @param {Array} patientIds - 患者IDリスト
 * @param {string} reportType - 帳票種類
 * @returns {Object} 出力結果
 */
function portalBatchGenerateReports(patientIds, reportType) {
  try {
    const results = [];
    for (const patientId of patientIds) {
      const result = portalGenerateReport(patientId, reportType);
      results.push({ patientId, ...result });
    }

    return {
      success: true,
      results: results,
      successCount: results.filter(r => r.success).length,
      errorCount: results.filter(r => !r.success).length
    };
  } catch (error) {
    console.error('portalBatchGenerateReports error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// 請求管理 API
// ============================================================

/**
 * 請求一覧を取得
 * @param {Object} criteria - 検索条件
 * @returns {Object} 請求一覧
 */
function portalGetBillingList(criteria) {
  try {
    // 既存のgetBillingList関数を呼び出し
    if (typeof getBillingList === 'function') {
      return getBillingList(criteria);
    }

    return { success: false, error: '請求管理機能が未実装です' };
  } catch (error) {
    console.error('portalGetBillingList error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 請求ステータスを更新
 * @param {string} patientId - 患者ID
 * @param {string} newStatus - 新ステータス
 * @returns {Object} 更新結果
 */
function portalUpdateBillingStatus(patientId, newStatus) {
  try {
    // 既存のupdateBillingStatus関数を呼び出し
    if (typeof updateBillingStatus === 'function') {
      return updateBillingStatus(patientId, newStatus);
    }

    return { success: false, error: '請求ステータス更新機能が未実装です' };
  } catch (error) {
    console.error('portalUpdateBillingStatus error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// ユーティリティ API
// ============================================================

/**
 * システム設定を取得
 * @returns {Object} システム設定
 */
function portalGetSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('設定');

    if (!sheet) {
      return {
        success: true,
        data: {
          organizationName: 'CDmedical',
          defaultCourse: '人間ドック',
          taxRate: 0.1
        }
      };
    }

    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        settings[data[i][0]] = data[i][1];
      }
    }

    return { success: true, data: settings };
  } catch (error) {
    console.error('portalGetSettings error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 選択肢マスタを取得
 * @param {string} masterType - マスタ種類
 * @returns {Object} 選択肢リスト
 */
function portalGetMasterOptions(masterType) {
  try {
    const options = {
      judgment: ['A', 'B', 'C', 'D', 'E', 'G'],
      ecgFindings: ['異常なし', '不完全右脚ブロック', '完全右脚ブロック', '左室肥大', '心房細動', 'その他'],
      xrayFindings: ['異常なし', '陳旧性病変', '胸膜肥厚', '側弯', 'その他'],
      courses: ['人間ドック', '定期健診', '労災二次検診', '生活習慣病予防健診'],
      billingStatus: ['未請求', '請求済', '入金済', 'キャンセル']
    };

    if (masterType && options[masterType]) {
      return { success: true, data: options[masterType] };
    }

    return { success: true, data: options };
  } catch (error) {
    console.error('portalGetMasterOptions error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 次の患者IDを取得（連続入力用）
 * @param {string} currentId - 現在の患者ID
 * @param {string} direction - 方向（next/prev）
 * @returns {Object} 次/前の患者ID
 */
function portalGetAdjacentPatient(currentId, direction) {
  try {
    const progressResult = portalGetPatientProgress({});
    if (!progressResult.success || progressResult.data.length === 0) {
      return { success: false, error: '患者データがありません' };
    }

    const patients = progressResult.data;
    const currentIndex = patients.findIndex(p => String(p.patientId) === String(currentId));

    if (currentIndex < 0) {
      return { success: false, error: '現在の患者が見つかりません' };
    }

    let targetIndex;
    if (direction === 'next') {
      targetIndex = currentIndex < patients.length - 1 ? currentIndex + 1 : 0;
    } else {
      targetIndex = currentIndex > 0 ? currentIndex - 1 : patients.length - 1;
    }

    return { success: true, data: patients[targetIndex] };
  } catch (error) {
    console.error('portalGetAdjacentPatient error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// デバッグ用テスト関数
// ============================================================

/**
 * デプロイ確認用テスト関数
 * ブラウザコンソールから google.script.run.testPortalApi() で呼び出し可能
 */
function testPortalApi() {
  return {
    success: true,
    message: 'portalApi.gs is deployed correctly',
    timestamp: new Date().toISOString(),
    version: '2025-12-16'
  };
}

/**
 * 検索機能デバッグ用
 * 固定値を返してAPI呼び出しが正常か確認
 */
function testSearchDebug() {
  return {
    success: true,
    data: [
      {
        '受診ID': 'TEST001',
        'ステータス': 'テスト',
        '受診日': '2025-12-16',
        '氏名': 'テスト 太郎',
        'カナ': 'テスト タロウ',
        '性別': '男',
        '生年月日': '1990-01-01',
        '年齢': 35,
        '受診コース': '人間ドック',
        '事業所名': 'テスト会社',
        '所属': '',
        '総合判定': 'A'
      }
    ],
    count: 1
  };
}

/**
 * 詳細検索デバッグ - GASエディタから実行
 * 11111 を検索して各ステップをログ出力
 */
function debugSearchDetailed() {
  console.log('========== 詳細検索デバッグ開始 ==========');

  const criteria = { patientId: '11111', name: '', kana: '' };
  console.log('検索条件:', JSON.stringify(criteria));

  // スプレッドシート接続
  const ss = getPortalSpreadsheet();
  console.log('スプレッドシートID:', ss.getId());

  const sheet = ss.getSheetByName('受診者マスタ');
  if (!sheet) {
    console.log('ERROR: 受診者マスタシートが見つかりません');
    return;
  }
  console.log('シート名:', sheet.getName());

  const data = sheet.getDataRange().getValues();
  console.log('総行数:', data.length);
  console.log('ヘッダー:', JSON.stringify(data[0]));

  // 検索条件を正規化
  const searchId = criteria.patientId ? normalizeString(criteria.patientId) : '';
  const searchName = criteria.name ? normalizeString(criteria.name) : '';
  const searchKana = criteria.kana ? normalizeKana(criteria.kana) : '';

  console.log('正規化後 - searchId:', searchId, ', searchName:', searchName, ', searchKana:', searchKana);

  console.log('\n========== 各行の検索処理 ==========');

  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = row[0];
    const rowName = row[3];
    const rowKana = row[4];

    // 空行スキップ
    if (!rowId) {
      console.log('行', i + 1, ': 空行 - スキップ');
      continue;
    }

    const normalizedRowId = normalizeString(rowId);
    const normalizedRowName = normalizeString(rowName);
    const normalizedRowKana = normalizeKana(rowKana);

    console.log('行', i + 1, ':');
    console.log('  元データ - ID:', rowId, '(型:', typeof rowId, '), 氏名:', rowName, ', カナ:', rowKana);
    console.log('  正規化後 - ID:', normalizedRowId, ', 氏名:', normalizedRowName, ', カナ:', normalizedRowKana);

    let match = true;

    if (searchId) {
      const idMatch = normalizedRowId.includes(searchId);
      console.log('  ID検索:', normalizedRowId, '.includes(', searchId, ') =', idMatch);
      if (!idMatch) match = false;
    }

    if (searchName && match) {
      const nameMatch = normalizedRowName.includes(searchName);
      console.log('  氏名検索:', normalizedRowName, '.includes(', searchName, ') =', nameMatch);
      if (!nameMatch) match = false;
    }

    if (searchKana && match) {
      const kanaMatch = normalizedRowKana.includes(searchKana);
      console.log('  カナ検索:', normalizedRowKana, '.includes(', searchKana, ') =', kanaMatch);
      if (!kanaMatch) match = false;
    }

    console.log('  → 結果:', match ? 'マッチ ✓' : '不一致 ✗');

    if (match) {
      results.push({
        '受診ID': rowId,
        '氏名': rowName,
        'カナ': rowKana
      });
    }
  }

  console.log('\n========== 検索結果 ==========');
  console.log('ヒット件数:', results.length);
  console.log('結果:', JSON.stringify(results, null, 2));

  return { success: true, data: results, count: results.length };
}

/**
 * Webアプリから呼び出して接続先スプレッドシートを確認
 */
function debugWebAppConnection() {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('受診者マスタ');
    const rowCount = sheet ? sheet.getLastRow() : 0;

    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetExists: !!sheet,
      rowCount: rowCount,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Webアプリコンテキストで検索の全ステップをデバッグ
 * ブラウザコンソールから呼び出し、結果をJSONで返す
 * @param {string} searchId - 検索する受診ID（デフォルト: '11111'）
 * @returns {Object} 各ステップの詳細情報
 */
function debugSearchFromWebApp(searchId) {
  const debug = {
    timestamp: new Date().toISOString(),
    steps: [],
    searchId: searchId || '11111',
    finalResult: null
  };

  try {
    // Step 1: スプレッドシート接続
    debug.steps.push({ step: 1, action: 'getPortalSpreadsheet 呼び出し' });
    const ss = getPortalSpreadsheet();
    debug.steps.push({
      step: 1,
      result: 'SUCCESS',
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName()
    });

    // Step 2: シート取得
    debug.steps.push({ step: 2, action: '受診者マスタシート取得' });
    const sheet = ss.getSheetByName('受診者マスタ');
    if (!sheet) {
      debug.steps.push({ step: 2, result: 'FAILED', error: 'シートが見つからない' });
      debug.finalResult = { success: false, error: 'シートが見つからない' };
      return debug;
    }
    debug.steps.push({ step: 2, result: 'SUCCESS', sheetName: sheet.getName() });

    // Step 3: データ取得
    debug.steps.push({ step: 3, action: 'getDataRange().getValues()' });
    const data = sheet.getDataRange().getValues();
    debug.steps.push({
      step: 3,
      result: 'SUCCESS',
      totalRows: data.length,
      headers: data[0] ? data[0].slice(0, 5).join(', ') + '...' : 'EMPTY'
    });

    // Step 4: データの先頭3行を取得（デバッグ用）
    debug.steps.push({ step: 4, action: 'サンプルデータ取得（先頭3行）' });
    const sampleData = [];
    for (let i = 1; i < Math.min(data.length, 4); i++) {
      sampleData.push({
        rowNum: i + 1,
        col0_受診ID: data[i][0],
        col0_type: typeof data[i][0],
        col3_氏名: data[i][3],
        col4_カナ: data[i][4]
      });
    }
    debug.steps.push({ step: 4, result: 'SUCCESS', sampleData: sampleData });

    // Step 5: 検索実行
    const targetId = searchId || '11111';
    debug.steps.push({ step: 5, action: '検索実行', targetId: targetId });

    const normalizedSearchId = normalizeString(targetId);
    debug.steps.push({ step: 5, normalizedSearchId: normalizedSearchId });

    const results = [];
    const matchLog = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowId = row[0];

      if (!rowId) continue;

      const normalizedRowId = normalizeString(rowId);
      const isMatch = normalizedRowId.includes(normalizedSearchId);

      matchLog.push({
        row: i + 1,
        rawId: rowId,
        normalizedId: normalizedRowId,
        searchId: normalizedSearchId,
        match: isMatch
      });

      if (isMatch) {
        results.push({
          '受診ID': row[0],
          '氏名': row[3],
          'カナ': row[4]
        });
      }

      // 最初の5行分のマッチログのみ保持
      if (matchLog.length > 5) break;
    }

    debug.steps.push({
      step: 5,
      result: 'SUCCESS',
      matchLog: matchLog,
      resultsCount: results.length,
      results: results
    });

    debug.finalResult = {
      success: true,
      data: results,
      count: results.length
    };

  } catch (error) {
    debug.steps.push({
      step: 'ERROR',
      error: error.message,
      stack: error.stack
    });
    debug.finalResult = { success: false, error: error.message };
  }

  return debug;
}
