/**
 * setupNormalizedStructure.gs
 *
 * 設計書準拠の正規化構造（受診者マスタ + 受診記録）を構築
 * DATA_STRUCTURE_DESIGN.md および Config.gs に基づく
 *
 * @version 1.0.0
 * @date 2025-12-19
 */

// ============================================
// 定数定義
// ============================================

const NORMALIZED_STRUCTURE = {
  SPREADSHEET_ID: '16KtctyT2gd7oJZdcu84kUtuP-D9jB9KtLxxzxXx_wdk',

  // 受診者マスタ（個人情報のみ - Config.gs COLUMN_DEFINITIONS.PATIENT_MASTER 準拠）
  PATIENT_MASTER: {
    sheetName: '受診者マスタ',
    headers: [
      '受診者ID',     // A: PK, P-00001形式
      '氏名',         // B
      'カナ',         // C
      '生年月日',     // D
      '性別',         // E: 男/女
      '郵便番号',     // F
      '住所',         // G
      '電話番号',     // H
      'メール',       // I
      '所属企業',     // J
      '備考',         // K
      '作成日時',     // L
      '更新日時'      // M
    ],
    columnWidths: [120, 100, 120, 100, 50, 80, 200, 120, 150, 150, 200, 150, 150]
  },

  // 受診記録（受診ごとの情報 - Config.gs COLUMN_DEFINITIONS.VISIT_RECORD 準拠 + 拡張）
  VISIT_RECORD: {
    sheetName: '受診記録',
    headers: [
      '受診ID',       // A: PK, YYYYMMDD-NNN形式
      '受診者ID',     // B: FK → 受診者マスタ
      '検診種別',     // C: 人間ドック/定期健診/労災二次/雇入健診
      'コース',       // D: コース名
      '受診日',       // E
      '年齢',         // F: 受診時年齢
      '総合判定',     // G: A/B/C/D/E
      '医師所見',     // H
      'ステータス',   // I: 入力中/確定/報告済
      'CSV取込日時',  // J
      '作成日時',     // K
      '更新日時',     // L
      '出力日時'      // M
    ],
    columnWidths: [130, 100, 100, 150, 100, 50, 80, 300, 80, 150, 150, 150, 150]
  }
};

// ============================================
// メイン関数
// ============================================

/**
 * 正規化構造のシートを構築
 * GASエディタから実行: setupNormalizedStructure()
 */
function setupNormalizedStructure() {
  const ss = SpreadsheetApp.openById(NORMALIZED_STRUCTURE.SPREADSHEET_ID);

  console.log('=== 正規化構造シート構築開始 ===');

  // 1. 受診者マスタを構築
  const patientSheet = createOrRebuildSheet(ss, NORMALIZED_STRUCTURE.PATIENT_MASTER);
  console.log('✅ 受診者マスタシート構築完了');

  // 2. 受診記録を構築
  const visitSheet = createOrRebuildSheet(ss, NORMALIZED_STRUCTURE.VISIT_RECORD);
  console.log('✅ 受診記録シート構築完了');

  // 3. テストデータを追加
  addTestData(patientSheet, visitSheet);
  console.log('✅ テストデータ追加完了');

  // 4. シートを先頭に移動
  ss.setActiveSheet(patientSheet);
  ss.moveActiveSheet(1);
  ss.setActiveSheet(visitSheet);
  ss.moveActiveSheet(2);

  console.log('=== 正規化構造シート構築完了 ===');

  return {
    success: true,
    message: '正規化構造シートを構築しました',
    sheets: {
      patientMaster: patientSheet.getName(),
      visitRecord: visitSheet.getName()
    }
  };
}

/**
 * シートを作成または再構築
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Object} config - シート設定
 * @returns {Sheet} 作成されたシート
 */
function createOrRebuildSheet(ss, config) {
  // 既存シートをバックアップ
  const existingSheet = ss.getSheetByName(config.sheetName);
  if (existingSheet) {
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const backupName = `${config.sheetName}_backup_${timestamp}`;
    existingSheet.setName(backupName);
    console.log(`  既存シートをバックアップ: ${backupName}`);
  }

  // 新しいシートを作成
  const sheet = ss.insertSheet(config.sheetName);

  // ヘッダーを設定
  const headers = config.headers;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅を設定
  config.columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  // ID列をテキスト形式に設定（A列、B列がIDの場合）
  sheet.getRange('A:A').setNumberFormat('@');
  if (config.sheetName === '受診記録') {
    sheet.getRange('B:B').setNumberFormat('@');
  }

  // 日付列の書式設定
  if (config.sheetName === '受診者マスタ') {
    sheet.getRange('D:D').setNumberFormat('yyyy/mm/dd');  // 生年月日
    sheet.getRange('L:L').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // 作成日時
    sheet.getRange('M:M').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // 更新日時
  } else if (config.sheetName === '受診記録') {
    sheet.getRange('E:E').setNumberFormat('yyyy/mm/dd');  // 受診日
    sheet.getRange('J:J').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // CSV取込日時
    sheet.getRange('K:K').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // 作成日時
    sheet.getRange('L:L').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // 更新日時
    sheet.getRange('M:M').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // 出力日時
  }

  // ヘッダー行を固定
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * テストデータを追加（同一患者の複数回受診）
 * @param {Sheet} patientSheet - 受診者マスタシート
 * @param {Sheet} visitSheet - 受診記録シート
 */
function addTestData(patientSheet, visitSheet) {
  const now = new Date();

  // 受診者マスタにテストデータ
  const patientData = [
    [
      'P-00001',                    // 受診者ID
      'テスト 太郎',                // 氏名
      'テスト タロウ',              // カナ
      new Date(1980, 0, 15),        // 生年月日
      '男',                         // 性別
      '150-0001',                   // 郵便番号
      '東京都渋谷区神宮前1-1-1',    // 住所
      '03-1234-5678',               // 電話番号
      'test@example.com',           // メール
      'テスト株式会社',             // 所属企業
      '',                           // 備考
      now,                          // 作成日時
      now                           // 更新日時
    ],
    [
      'P-00002',                    // 受診者ID
      '山田 花子',                  // 氏名
      'ヤマダ ハナコ',              // カナ
      new Date(1985, 5, 20),        // 生年月日
      '女',                         // 性別
      '160-0022',                   // 郵便番号
      '東京都新宿区新宿2-2-2',      // 住所
      '03-9876-5432',               // 電話番号
      'hanako@example.com',         // メール
      'サンプル株式会社',           // 所属企業
      '',                           // 備考
      now,                          // 作成日時
      now                           // 更新日時
    ]
  ];

  patientSheet.getRange(2, 1, patientData.length, patientData[0].length).setValues(patientData);

  // 受診記録にテストデータ（同一患者P-00001の複数回受診）
  const visitData = [
    [
      '20250401-001',               // 受診ID
      'P-00001',                    // 受診者ID (FK)
      '定期健診',                   // 検診種別
      '定期健康診断A',              // コース
      new Date(2025, 3, 1),         // 受診日
      45,                           // 年齢
      'A',                          // 総合判定
      '異常所見なし',               // 医師所見
      '確定',                       // ステータス
      '',                           // CSV取込日時
      now,                          // 作成日時
      now,                          // 更新日時
      ''                            // 出力日時
    ],
    [
      '20251219-001',               // 受診ID
      'P-00001',                    // 受診者ID (FK) - 同一患者の2回目受診
      '人間ドック',                 // 検診種別
      '生活習慣病ドック',           // コース
      new Date(2025, 11, 19),       // 受診日
      45,                           // 年齢
      '',                           // 総合判定（未入力）
      '',                           // 医師所見
      '入力中',                     // ステータス
      '',                           // CSV取込日時
      now,                          // 作成日時
      now,                          // 更新日時
      ''                            // 出力日時
    ],
    [
      '20251220-001',               // 受診ID
      'P-00002',                    // 受診者ID (FK)
      '人間ドック',                 // 検診種別
      '消化器ドック',               // コース
      new Date(2025, 11, 20),       // 受診日
      40,                           // 年齢
      '',                           // 総合判定
      '',                           // 医師所見
      '入力中',                     // ステータス
      '',                           // CSV取込日時
      now,                          // 作成日時
      now,                          // 更新日時
      ''                            // 出力日時
    ]
  ];

  visitSheet.getRange(2, 1, visitData.length, visitData[0].length).setValues(visitData);

  console.log(`  テストデータ: 受診者${patientData.length}件, 受診記録${visitData.length}件`);
}

// ============================================
// 動作確認用関数
// ============================================

/**
 * 正規化構造の確認テスト
 * GASエディタから実行: testNormalizedStructure()
 */
function testNormalizedStructure() {
  const ss = SpreadsheetApp.openById(NORMALIZED_STRUCTURE.SPREADSHEET_ID);

  const patientSheet = ss.getSheetByName('受診者マスタ');
  const visitSheet = ss.getSheetByName('受診記録');

  console.log('=== 正規化構造確認 ===');

  if (!patientSheet) {
    console.log('❌ 受診者マスタシートが見つかりません');
    return { success: false, error: '受診者マスタシートなし' };
  }

  if (!visitSheet) {
    console.log('❌ 受診記録シートが見つかりません');
    return { success: false, error: '受診記録シートなし' };
  }

  console.log('受診者マスタ:');
  console.log('  行数:', patientSheet.getLastRow());
  console.log('  ヘッダー:', patientSheet.getRange(1, 1, 1, 13).getValues()[0].join(', '));

  console.log('受診記録:');
  console.log('  行数:', visitSheet.getLastRow());
  console.log('  ヘッダー:', visitSheet.getRange(1, 1, 1, 13).getValues()[0].join(', '));

  // 受診者と受診記録の紐付け確認
  console.log('\n=== 受診者別受診回数 ===');
  const visitData = visitSheet.getDataRange().getValues();
  const visitsByPatient = {};

  for (let i = 1; i < visitData.length; i++) {
    const patientId = visitData[i][1];
    if (!patientId) continue;

    if (!visitsByPatient[patientId]) {
      visitsByPatient[patientId] = [];
    }
    visitsByPatient[patientId].push({
      visitId: visitData[i][0],
      visitDate: visitData[i][4],
      examType: visitData[i][2]
    });
  }

  for (const [patientId, visits] of Object.entries(visitsByPatient)) {
    console.log(`${patientId}: ${visits.length}回受診`);
    visits.forEach(v => {
      console.log(`  - ${v.visitId} (${v.examType})`);
    });
  }

  return {
    success: true,
    patientCount: patientSheet.getLastRow() - 1,
    visitCount: visitSheet.getLastRow() - 1,
    visitsByPatient: visitsByPatient
  };
}

/**
 * 次の受診者IDを生成
 * @returns {string} P-NNNNN形式のID
 */
function generateNextPatientId() {
  const ss = SpreadsheetApp.openById(NORMALIZED_STRUCTURE.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('受診者マスタ');

  if (!sheet || sheet.getLastRow() < 2) {
    return 'P-00001';
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let maxNum = 0;

  for (const row of data) {
    const id = String(row[0]);
    if (id.startsWith('P-')) {
      const num = parseInt(id.replace('P-', ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  return `P-${String(maxNum + 1).padStart(5, '0')}`;
}

/**
 * 次の受診IDを生成
 * @param {Date} visitDate - 受診日
 * @returns {string} YYYYMMDD-NNN形式のID
 */
function generateNextVisitId(visitDate) {
  const ss = SpreadsheetApp.openById(NORMALIZED_STRUCTURE.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('受診記録');

  const date = visitDate || new Date();
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
  const prefix = `${dateStr}-`;

  if (!sheet || sheet.getLastRow() < 2) {
    return `${prefix}001`;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let maxNum = 0;

  for (const row of data) {
    const id = String(row[0]);
    if (id.startsWith(prefix)) {
      const num = parseInt(id.replace(prefix, ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  return `${prefix}${String(maxNum + 1).padStart(3, '0')}`;
}
