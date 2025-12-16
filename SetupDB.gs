/**
 * 健診結果DB 統合システム - DB初期セットアップ
 *
 * @description 設計書: 健診結果DB_設計書_v1.md Phase 1 基盤構築
 * @version 1.0.0
 * @date 2025-12-14
 */

// ============================================
// メインセットアップ関数
// ============================================

/**
 * Phase 1: DB初期セットアップを実行
 * 設計書4章に基づいて6シートを作成
 */
function setupDatabase() {
  logInfo('===== Phase 1: DB初期セットアップ開始 =====');

  try {
    const ss = getSpreadsheet();

    // 1. 各シートを作成
    createAllSheets(ss);

    // 2. マスタデータを投入
    insertMasterData(ss);

    // 3. 全シートを非表示に設定
    hideAllSheets(ss);

    logInfo('===== Phase 1: DB初期セットアップ完了 =====');

    // 結果をダイアログで表示
    const ui = SpreadsheetApp.getUi();
    ui.alert('セットアップ完了',
      'データベースの初期セットアップが完了しました。\n\n' +
      '作成されたシート:\n' +
      '- 受診者マスタ\n' +
      '- 受診記録\n' +
      '- 検査結果\n' +
      '- 項目マスタ\n' +
      '- 検診種別マスタ\n' +
      '- コースマスタ\n' +
      '- 保健指導記録\n\n' +
      '※ 全シートは非表示に設定されています。',
      ui.ButtonSet.OK);

  } catch (e) {
    logError('setupDatabase', e);
    throw e;
  }
}

/**
 * 全シートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createAllSheets(ss) {
  logInfo('シート作成開始...');

  // シート定義と対応するカラム定義
  const sheetConfigs = [
    { name: DB_CONFIG.SHEETS.PATIENT_MASTER, def: COLUMN_DEFINITIONS.PATIENT_MASTER },
    { name: DB_CONFIG.SHEETS.VISIT_RECORD, def: COLUMN_DEFINITIONS.VISIT_RECORD },
    { name: DB_CONFIG.SHEETS.TEST_RESULT, def: COLUMN_DEFINITIONS.TEST_RESULT },
    { name: DB_CONFIG.SHEETS.ITEM_MASTER, def: COLUMN_DEFINITIONS.ITEM_MASTER },
    { name: DB_CONFIG.SHEETS.EXAM_TYPE_MASTER, def: COLUMN_DEFINITIONS.EXAM_TYPE_MASTER },
    { name: DB_CONFIG.SHEETS.COURSE_MASTER, def: COLUMN_DEFINITIONS.COURSE_MASTER },
    { name: DB_CONFIG.SHEETS.GUIDANCE_RECORD, def: COLUMN_DEFINITIONS.GUIDANCE_RECORD }
  ];

  for (const config of sheetConfigs) {
    createSheet(ss, config.name, config.def);
  }

  // Phase 3: マッピングパターンシート作成
  createMappingPatternSheet(ss);

  logInfo('全シート作成完了');
}

/**
 * マッピングパターンシートを作成（Phase 3: CSVインポート用）
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createMappingPatternSheet(ss) {
  const sheetName = MAPPING_PATTERN_SHEET.NAME;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    return sheet;
  }

  // ヘッダー設定
  const headers = MAPPING_PATTERN_SHEET.HEADERS;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  sheet.setColumnWidth(1, 120);  // パターンID
  sheet.setColumnWidth(2, 150);  // ソース名
  sheet.setColumnWidth(3, 100);  // ヘッダーハッシュ
  sheet.setColumnWidth(4, 100);  // データ種別
  sheet.setColumnWidth(5, 300);  // マッピングJSON
  sheet.setColumnWidth(6, 200);  // 値変換JSON
  sheet.setColumnWidth(7, 80);   // 使用回数
  sheet.setColumnWidth(8, 150);  // 作成日時
  sheet.setColumnWidth(9, 150);  // 更新日時

  // 1行目を固定
  sheet.setFrozenRows(1);

  // シートを非表示
  sheet.hideSheet();

  return sheet;
}

/**
 * シートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @param {Object} definition - カラム定義
 */
function createSheet(ss, sheetName, definition) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    // 既存シートの場合、ヘッダー行のみ更新
  }

  // ヘッダー設定
  const headers = definition.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  if (definition.columnWidths) {
    for (const [col, width] of Object.entries(definition.columnWidths)) {
      const colIndex = columnLetterToIndex(col);
      sheet.setColumnWidth(colIndex, width);
    }
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * 全シートを非表示に設定
 * @param {Spreadsheet} ss - スプレッドシート
 */
function hideAllSheets(ss) {
  logInfo('シート非表示設定...');

  for (const sheetName of Object.values(DB_CONFIG.SHEETS)) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.hideSheet();
      logInfo(`  非表示: ${sheetName}`);
    }
  }

  logInfo('全シート非表示完了');
}

/**
 * 全シートを表示（開発・デバッグ用）
 */
function showAllSheets() {
  const ss = getSpreadsheet();

  for (const sheetName of Object.values(DB_CONFIG.SHEETS)) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.showSheet();
      logInfo(`シート表示: ${sheetName}`);
    }
  }
}

// ============================================
// マスタデータ投入
// ============================================

/**
 * 全マスタデータを投入
 * @param {Spreadsheet} ss - スプレッドシート
 */
function insertMasterData(ss) {
  logInfo('マスタデータ投入開始...');

  // 検診種別マスタ
  insertExamTypeMasterData(ss);

  // コースマスタ
  insertCourseMasterData(ss);

  // 項目マスタ
  insertItemMasterData(ss);

  logInfo('マスタデータ投入完了');
}

/**
 * 検診種別マスタデータを投入
 * @param {Spreadsheet} ss - スプレッドシート
 */
function insertExamTypeMasterData(ss) {
  const sheet = ss.getSheetByName(DB_CONFIG.SHEETS.EXAM_TYPE_MASTER);
  if (!sheet || sheet.getLastRow() > 1) {
    logInfo('  検診種別マスタ: スキップ（既存データあり）');
    return;
  }

  // 設計書4.5準拠
  const data = [
    ['DOCK', '人間ドック', true, 1, true],
    ['REGULAR', '定期検診', false, 2, true],
    ['EMPLOY', '雇入検診', false, 3, true],
    ['ROSAI', '労災二次', false, 4, true],
    ['SPECIFIC', '特定健診', false, 5, true]
  ];

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  logInfo(`  検診種別マスタ: ${data.length}件投入`);
}

/**
 * コースマスタデータを投入
 * @param {Spreadsheet} ss - スプレッドシート
 */
function insertCourseMasterData(ss) {
  const sheet = ss.getSheetByName(DB_CONFIG.SHEETS.COURSE_MASTER);
  if (!sheet || sheet.getLastRow() > 1) {
    logInfo('  コースマスタ: スキップ（既存データあり）');
    return;
  }

  // 人間ドックコース（参考: SYSTEM_DESIGN.md）
  const data = [
    ['DOCK_LIFE', '生活習慣病ドック', 40000, 'HEIGHT,WEIGHT,BMI,BP_SYS,BP_DIA,FBS,HBA1C,TCHO,HDL,LDL,TG,AST,ALT,GGT', 1, true],
    ['DOCK_GI', '消化器ドック', 60000, 'HEIGHT,WEIGHT,BMI,BP_SYS,BP_DIA,FBS,HBA1C,TCHO,HDL,LDL,TG,AST,ALT,GGT,CEA,CA19-9', 2, true],
    ['DOCK_FULL', '全身スクリーニング', 160000, 'HEIGHT,WEIGHT,BMI,BP_SYS,BP_DIA,FBS,HBA1C,TCHO,HDL,LDL,TG,AST,ALT,GGT,UA,CR', 3, true],
    ['DOCK_CANCER', 'がんスクリーニング', 70000, 'HEIGHT,WEIGHT,BMI,CEA,CA19-9,AFP,PSA', 4, true]
  ];

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  logInfo(`  コースマスタ: ${data.length}件投入`);
}

/**
 * 項目マスタデータを投入
 * 設計書5章に基づく判定基準
 * @param {Spreadsheet} ss - スプレッドシート
 */
function insertItemMasterData(ss) {
  const sheet = ss.getSheetByName(DB_CONFIG.SHEETS.ITEM_MASTER);
  if (!sheet || sheet.getLastRow() > 1) {
    logInfo('  項目マスタ: スキップ（既存データあり）');
    return;
  }

  // 設計書5章準拠の項目マスタ初期データ
  // [項目ID, 項目名, カテゴリ, 単位, データ型, 性別差, 判定方法, A下限, A上限, B下限, B上限, C下限, C上限, D条件, A下限_F, A上限_F, 表示順, 有効]
  const data = [
    // 5.1 身体測定
    ['HEIGHT', '身長', '身体測定', 'cm', '数値', false, 'なし', '', '', '', '', '', '', '', '', '', 1, true],
    ['WEIGHT', '体重', '身体測定', 'kg', '数値', false, 'なし', '', '', '', '', '', '', '', '', '', 2, true],
    ['BMI', 'BMI', '身体測定', '', '数値', false, 'ABCD', 18.5, 24.9, 25.0, 27.9, 28.0, 34.9, '>=35,<18.5', '', '', 3, true],
    ['WAIST_M', '腹囲(男)', '身体測定', 'cm', '数値', true, 'ABCD', 0, 84.9, 85.0, 89.9, 90.0, 99.9, '>=100', '', '', 4, true],
    ['WAIST_F', '腹囲(女)', '身体測定', 'cm', '数値', true, 'ABCD', 0, 89.9, 90.0, 94.9, 95.0, 99.9, '>=100', '', '', 5, true],
    ['BP_SYS', '収縮期血圧', '身体測定', 'mmHg', '数値', false, 'ABCD', 90, 129, 130, 139, 140, 159, '>=160,<90', '', '', 6, true],
    ['BP_DIA', '拡張期血圧', '身体測定', 'mmHg', '数値', false, 'ABCD', 50, 84, 85, 89, 90, 99, '>=100,<50', '', '', 7, true],

    // 5.2 血液検査（糖代謝）
    ['FBS', '空腹時血糖', '糖代謝', 'mg/dL', '数値', false, '特殊', 0, 99, 100, 109, 110, 125, '>=126', '', '', 10, true],
    ['HBA1C', 'HbA1c', '糖代謝', '%', '数値', false, '特殊', 0, 5.5, 5.6, 5.9, 6.0, 6.4, '>=6.5', '', '', 11, true],

    // 5.3 血液検査（脂質）
    ['TCHO', '総コレステロール', '脂質', 'mg/dL', '数値', false, 'ABCD', 140, 199, 200, 219, 220, 259, '>=260,<140', '', '', 20, true],
    ['HDL', 'HDLコレステロール', '脂質', 'mg/dL', '数値', false, 'ABCD', 40, 119, 35, 39, 30, 34, '<30', '', '', 21, true],
    ['LDL', 'LDLコレステロール', '脂質', 'mg/dL', '数値', false, 'ABCD', 60, 119, 120, 139, 140, 179, '>=180', '', '', 22, true],
    ['TG', '中性脂肪', '脂質', 'mg/dL', '数値', false, 'ABCD', 30, 149, 150, 299, 300, 499, '>=500', '', '', 23, true],

    // 5.4 血液検査（肝機能）
    ['AST', 'GOT(AST)', '肝機能', 'U/L', '数値', false, 'ABCD', 0, 30, 31, 50, 51, 100, '>100', '', '', 30, true],
    ['ALT', 'GPT(ALT)', '肝機能', 'U/L', '数値', false, 'ABCD', 0, 30, 31, 50, 51, 100, '>100', '', '', 31, true],
    ['GGT', 'γ-GTP', '肝機能', 'U/L', '数値', false, 'ABCD', 0, 50, 51, 100, 101, 200, '>200', '', '', 32, true],

    // 5.5 血液検査（腎機能）
    ['UA', '尿酸', '腎機能', 'mg/dL', '数値', false, 'ABCD', 2.0, 7.0, 7.1, 8.0, 8.1, 9.0, '>9.0', '', '', 40, true],
    ['CR_M', 'クレアチニン(男)', '腎機能', 'mg/dL', '数値', true, 'ABCD', 0.6, 1.1, 1.2, 1.3, 1.4, 1.5, '>1.5', '', '', 41, true],
    ['CR_F', 'クレアチニン(女)', '腎機能', 'mg/dL', '数値', true, 'ABCD', 0.4, 0.8, 0.9, 1.0, 1.1, 1.2, '>1.2', '', '', 42, true],

    // 5.6 尿検査
    ['U_PRO', '尿蛋白', '尿検査', '', '定性', false, '正常値', '', '', '', '', '', '', '(+)以上', '', '', 50, true],
    ['U_GLU', '尿糖', '尿検査', '', '定性', false, '正常値', '', '', '', '', '', '', '(+)以上', '', '', 51, true],
    ['U_OB', '尿潜血', '尿検査', '', '定性', false, '正常値', '', '', '', '', '', '', '(+)以上', '', '', 52, true],

    // 追加の血液検査項目
    ['WBC', '白血球数', '血液一般', '×10³/μL', '数値', false, 'ABCD', 3.2, 8.5, 2.5, 3.1, 2.0, 2.4, '<2.0,>15', '', '', 60, true],
    ['RBC', '赤血球数', '血液一般', '×10⁴/μL', '数値', true, 'ABCD', 400, 539, 380, 399, 350, 379, '<350', 360, 489, 61, true],
    ['HB', 'ヘモグロビン', '血液一般', 'g/dL', '数値', true, 'ABCD', 13.1, 16.6, 12.1, 13.0, 11.0, 12.0, '<11.0', 12.1, 14.5, 62, true],
    ['HT', 'ヘマトクリット', '血液一般', '%', '数値', true, 'ABCD', 38.5, 48.9, 35.0, 38.4, 30.0, 34.9, '<30.0', 34.0, 43.9, 63, true],
    ['PLT', '血小板数', '血液一般', '×10⁴/μL', '数値', false, 'ABCD', 13.0, 34.9, 10.0, 12.9, 8.0, 9.9, '<8.0,>40', '', '', 64, true]
  ];

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  logInfo(`  項目マスタ: ${data.length}件投入`);
}

// ============================================
// データ検証・保守関数
// ============================================

/**
 * DBの整合性をチェック
 */
function validateDatabase() {
  logInfo('===== DB整合性チェック =====');

  const ss = getSpreadsheet();
  const issues = [];

  // 必須シートの存在確認
  for (const [key, sheetName] of Object.entries(DB_CONFIG.SHEETS)) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      issues.push(`シート未作成: ${sheetName}`);
    }
  }

  // マスタデータの存在確認
  const itemMasterSheet = ss.getSheetByName(DB_CONFIG.SHEETS.ITEM_MASTER);
  if (itemMasterSheet && itemMasterSheet.getLastRow() <= 1) {
    issues.push('項目マスタ: データなし');
  }

  const examTypeSheet = ss.getSheetByName(DB_CONFIG.SHEETS.EXAM_TYPE_MASTER);
  if (examTypeSheet && examTypeSheet.getLastRow() <= 1) {
    issues.push('検診種別マスタ: データなし');
  }

  // 結果出力
  if (issues.length === 0) {
    logInfo('整合性チェック: 問題なし');
  } else {
    logInfo(`整合性チェック: ${issues.length}件の問題`);
    for (const issue of issues) {
      logInfo(`  - ${issue}`);
    }
  }

  return issues;
}

/**
 * マスタデータを再投入（リセット）
 */
function resetMasterData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認',
    'マスタデータを初期状態にリセットします。\n' +
    '既存のマスタデータは削除されます。\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  const ss = getSpreadsheet();

  // マスタシートのデータをクリア（ヘッダー以外）
  const masterSheets = [
    DB_CONFIG.SHEETS.ITEM_MASTER,
    DB_CONFIG.SHEETS.EXAM_TYPE_MASTER,
    DB_CONFIG.SHEETS.COURSE_MASTER
  ];

  for (const sheetName of masterSheets) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
  }

  // マスタデータを再投入
  insertMasterData(ss);

  ui.alert('完了', 'マスタデータをリセットしました。', ui.ButtonSet.OK);
}
