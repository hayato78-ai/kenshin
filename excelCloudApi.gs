/**
 * Excel Cloud API
 *
 * @description Cloud Functions経由でExcel出力を行うAPI
 * @version 2.0.0
 * @date 2025-12-19
 */

// ============================================
// 設定
// ============================================

/** Cloud Functions URL */
const EXCEL_CLOUD_FUNCTION_URL = 'https://asia-northeast1-cdm-kenshin-system.cloudfunctions.net/generate-excel';

// ============================================
// メイン関数
// ============================================

/**
 * Cloud FunctionsでExcelを生成
 * @param {string} visitId - 受診記録ID
 * @returns {Object} - {success, downloadUrl, filename}
 */
function generateExcelCloud(visitId) {
  try {
    console.log('generateExcelCloud開始: visitId=' + visitId);

    // 1. 受診記録を取得
    const visitRecord = getVisitRecordById(visitId);
    if (!visitRecord) {
      throw new Error('受診記録が見つかりません: ' + visitId);
    }
    console.log('受診記録取得完了:', JSON.stringify(visitRecord));

    // 2. 患者情報を取得
    // patientIdが不正なオブジェクト文字列の場合に対応
    let patientId = visitRecord.patientId;
    if (typeof patientId === 'string' && patientId.includes('patientId=')) {
      // "{success=true, patientId=-0108}" から "-0108" を抽出
      const match = patientId.match(/patientId=([^}]+)/);
      if (match) {
        patientId = match[1].trim();
        console.log('patientId修正: ' + visitRecord.patientId + ' → ' + patientId);
      }
    }

    const patient = getPatientById(patientId);
    if (!patient) {
      throw new Error('患者情報が見つかりません: ' + patientId);
    }
    console.log('患者情報取得完了:', patient.name);

    // 3. 検査結果を取得
    const testResults = getTestResultsByVisitId(visitId);
    console.log('検査結果取得完了: ' + testResults.length + '件');

    // 4. カテゴリ判定を取得
    let categoryJudgments = {};
    try {
      categoryJudgments = getCategoryJudgments(visitRecord.patientId);
    } catch (e) {
      console.warn('カテゴリ判定取得スキップ:', e.message);
    }

    // 5. コース情報を取得
    let courseName = '';
    try {
      courseName = getCourseName(visitRecord.courseId);
    } catch (e) {
      console.warn('コース名取得スキップ:', e.message);
    }

    // 6. API用データ構築
    const payload = {
      patient_id: patient.patientId,  // ファイル名用（ルートレベル）
      patient: {
        id: patient.patientId,
        name: patient.name,
        kana: patient.kana || '',
        birthDate: formatDateForApi(patient.birthdate),
        gender: patient.gender,
        age: visitRecord.age || calculateAge(patient.birthdate),
        company: patient.company || ''
      },
      visit: {
        id: visitRecord.visitId,
        date: formatDateForApi(visitRecord.visitDate),
        course: courseName,
        examTypeId: visitRecord.examTypeId,
        overallJudgment: visitRecord.overallJudgment || '',
        doctorNotes: visitRecord.doctorNotes || ''
      },
      testResults: testResults.map(function(r) {
        return {
          itemId: r.itemId,
          value: r.value,
          numericValue: r.numericValue,
          judgment: r.judgment,
          notes: r.notes || ''
        };
      }),
      categoryJudgments: categoryJudgments
    };

    console.log('APIペイロード構築完了');

    // 7. Cloud Functions呼び出し
    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 60000
    };

    console.log('Cloud Functions呼び出し: ' + EXCEL_CLOUD_FUNCTION_URL);
    const response = UrlFetchApp.fetch(EXCEL_CLOUD_FUNCTION_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log('レスポンスコード: ' + responseCode);

    if (responseCode !== 200) {
      throw new Error('Cloud Functions エラー (HTTP ' + responseCode + '): ' + responseText);
    }

    const result = JSON.parse(responseText);

    if (!result.success) {
      throw new Error('Excel生成失敗: ' + (result.error || '不明なエラー'));
    }

    // 8. Base64デコードしてDriveに保存
    console.log('Excelファイル作成中...');
    console.log('Base64データ長: ' + (result.base64 ? result.base64.length : 0));

    // Base64デコード（result.base64 が正しいキー名）
    const decoded = Utilities.base64Decode(result.base64);
    console.log('デコード完了: ' + decoded.length + ' bytes');

    // Blob作成
    const blob = Utilities.newBlob(decoded);
    blob.setContentType('application/vnd.ms-excel.sheet.macroEnabled.12');
    blob.setName(result.filename);

    const file = DriveApp.createFile(blob);
    const downloadUrl = file.getDownloadUrl();

    console.log('ファイル作成完了: ' + result.filename);

    // 9. 一定時間後にファイル削除をスケジュール（オプション）
    scheduleFileCleanup_(file.getId());

    return {
      success: true,
      downloadUrl: downloadUrl,
      filename: result.filename,
      fileId: file.getId()
    };

  } catch (e) {
    console.error('generateExcelCloud エラー:', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * ポータルから呼び出すエンドポイント
 * @param {string} visitId - 受診ID
 * @returns {Object} 結果
 */
function handleExcelCloudExport(visitId) {
  try {
    const result = generateExcelCloud(visitId);
    return result;
  } catch (e) {
    console.error('handleExcelCloudExport エラー:', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Excel出力（ポータルUI用エントリーポイント）
 * @param {string} visitId - 受診ID
 * @returns {Object} 結果
 */
function apiExportExcelFromCloud(visitId) {
  return handleExcelCloudExport(visitId);
}

// ============================================
// ヘルパー関数
// ============================================

/**
 * 日付をAPI用にフォーマット
 * @param {Date|string} date - 日付
 * @returns {string} YYYY-MM-DD形式
 */
function formatDateForApi(date) {
  if (!date) return '';

  try {
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return '';

    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}

/**
 * 年齢を計算
 * @param {Date|string} birthDate - 生年月日
 * @returns {number} 年齢
 */
function calculateAge(birthDate) {
  if (!birthDate) return 0;

  try {
    const birth = (birthDate instanceof Date) ? birthDate : new Date(birthDate);
    const today = new Date();

    let age = today.getFullYear() - birth.getFullYear();
    const monthDiff = today.getMonth() - birth.getMonth();

    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birth.getDate())) {
      age--;
    }

    return age;
  } catch (e) {
    return 0;
  }
}

/**
 * コース名を取得
 * @param {string} courseId - コースID
 * @returns {string} コース名
 */
function getCourseName(courseId) {
  if (!courseId) return '';

  try {
    // コースマスタから取得を試みる
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('コースマスタ');
    if (!sheet) return courseId;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === courseId) {
        return data[i][1] || courseId;
      }
    }

    return courseId;
  } catch (e) {
    return courseId;
  }
}

/**
 * ファイル削除をスケジュール（1時間後）
 * @param {string} fileId - ファイルID
 */
function scheduleFileCleanup_(fileId) {
  try {
    // 1時間後に削除するトリガーを設定
    const trigger = ScriptApp.newTrigger('cleanupExportedFile_')
      .timeBased()
      .after(60 * 60 * 1000) // 1時間
      .create();

    // トリガーIDとファイルIDを紐付け
    const props = PropertiesService.getScriptProperties();
    const cleanupMap = JSON.parse(props.getProperty('EXCEL_CLEANUP_MAP') || '{}');
    cleanupMap[trigger.getUniqueId()] = fileId;
    props.setProperty('EXCEL_CLEANUP_MAP', JSON.stringify(cleanupMap));

  } catch (e) {
    console.warn('ファイルクリーンアップスケジュール失敗:', e.message);
  }
}

/**
 * エクスポートファイルのクリーンアップ（トリガーから呼び出し）
 * @param {Object} e - トリガーイベント
 */
function cleanupExportedFile_(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const cleanupMap = JSON.parse(props.getProperty('EXCEL_CLEANUP_MAP') || '{}');

    const triggerId = e.triggerUid;
    const fileId = cleanupMap[triggerId];

    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
        console.log('エクスポートファイル削除完了:', fileId);
      } catch (fileError) {
        console.warn('ファイル削除スキップ（既に削除済み）:', fileId);
      }

      // マップから削除
      delete cleanupMap[triggerId];
      props.setProperty('EXCEL_CLEANUP_MAP', JSON.stringify(cleanupMap));
    }

    // トリガー自体を削除
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }

  } catch (e) {
    console.error('cleanupExportedFile_ エラー:', e);
  }
}

// ============================================
// Cloud Functions ヘルスチェック
// ============================================

/**
 * Cloud Functionsの接続テスト
 * @returns {Object} テスト結果
 */
function testCloudFunctionsConnection() {
  try {
    const healthUrl = EXCEL_CLOUD_FUNCTION_URL.replace('/generate-excel', '/health');

    const response = UrlFetchApp.fetch(healthUrl, {
      method: 'GET',
      muteHttpExceptions: true,
      timeout: 10000
    });

    const code = response.getResponseCode();

    return {
      success: code === 200,
      status: code,
      message: code === 200 ? 'Cloud Functions接続OK' : 'Cloud Functions接続エラー'
    };

  } catch (e) {
    return {
      success: false,
      error: e.message,
      message: 'Cloud Functions接続失敗'
    };
  }
}
