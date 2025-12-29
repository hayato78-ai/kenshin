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
  // ※ ITEM_MASTER は EXAM_ITEM_MASTER + JUDGMENT_CRITERIA に統一のため削除
  const sheetConfigs = [
    { name: DB_CONFIG.SHEETS.PATIENT_MASTER, def: COLUMN_DEFINITIONS.PATIENT_MASTER },
    { name: DB_CONFIG.SHEETS.VISIT_RECORD, def: COLUMN_DEFINITIONS.VISIT_RECORD },
    { name: DB_CONFIG.SHEETS.TEST_RESULT, def: COLUMN_DEFINITIONS.TEST_RESULT },
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

  // 人間ドックコース（参考: 1220_new/SYSTEM_DESIGN_SPECIFICATION.md）
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
 * @deprecated 項目マスタは EXAM_ITEM_MASTER + JUDGMENT_CRITERIA に統一済み
 * この関数は後方互換性のため残していますが、使用しないでください。
 * MasterData.js の EXAM_ITEM_MASTER_DATA, JUDGMENT_CRITERIA_DATA を使用してください。
 */
function insertItemMasterData(ss) {
  logInfo('  項目マスタ: 廃止済み（EXAM_ITEM_MASTER + JUDGMENT_CRITERIA を使用）');
  // 旧コードは削除済み
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
  // ※ ITEM_MASTER は廃止 → EXAM_ITEM_MASTER を確認
  const examItemMasterSheet = ss.getSheetByName(DB_CONFIG.SHEETS.EXAM_ITEM_MASTER);
  if (examItemMasterSheet && examItemMasterSheet.getLastRow() <= 1) {
    issues.push('検査項目マスタ: データなし（MasterData.jsのEXAM_ITEM_MASTER_DATAを確認）');
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
  // ※ ITEM_MASTER は廃止 → EXAM_ITEM_MASTER, JUDGMENT_CRITERIA はMasterData.js管理
  const masterSheets = [
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

// ============================================
// Phase 1: 検査項目マスタ拡張シート作成
// ============================================

/**
 * Phase 1: 拡張マスタシートをセットアップ
 * 検査項目マスタ（150項目）、判定基準マスタ、選択肢マスタ、健診コースマスタを作成
 */
function setupExtendedMasterSheets() {
  logInfo('===== Phase 1: 拡張マスタシート作成開始 =====');

  try {
    const ss = getSpreadsheet();

    // 1. 検査項目マスタシート作成
    createExamItemMasterSheet(ss);

    // 2. 判定基準マスタシート作成
    createJudgmentCriteriaSheet(ss);

    // 3. 選択肢マスタシート作成
    createSelectOptionsSheet(ss);

    // 4. 健診コースマスタシート作成
    createExamCourseMasterSheet(ss);

    logInfo('===== Phase 1: 拡張マスタシート作成完了 =====');

    // 結果をダイアログで表示
    const ui = SpreadsheetApp.getUi();
    ui.alert('セットアップ完了',
      '拡張マスタシートの作成が完了しました。\n\n' +
      '作成されたシート:\n' +
      '- 検査項目マスタ（150項目）\n' +
      '- 判定基準マスタ（人間ドック学会2025年度版）\n' +
      '- 選択肢マスタ（定性検査用）\n' +
      '- 健診コースマスタ（6コース）',
      ui.ButtonSet.OK);

  } catch (e) {
    logError('setupExtendedMasterSheets', e);
    throw e;
  }
}

/**
 * 検査項目マスタシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createExamItemMasterSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.EXAM_ITEM_MASTER;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    // 既存データがある場合はスキップ
    if (sheet.getLastRow() > 1) {
      logInfo(`    データ既存、スキップ`);
      return sheet;
    }
  }

  // ヘッダー設定
  const headers = EXAM_ITEM_MASTER_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(EXAM_ITEM_MASTER_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // データ投入
  const data = getExamItemMasterRows();
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    logInfo(`    データ投入: ${data.length}件`);
  }

  return sheet;
}

/**
 * 判定基準マスタシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createJudgmentCriteriaSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.JUDGMENT_CRITERIA;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    if (sheet.getLastRow() > 1) {
      logInfo(`    データ既存、スキップ`);
      return sheet;
    }
  }

  // ヘッダー設定
  const headers = JUDGMENT_CRITERIA_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(JUDGMENT_CRITERIA_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // データ投入
  const data = getJudgmentCriteriaRows();
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    logInfo(`    データ投入: ${data.length}件`);
  }

  return sheet;
}

/**
 * 選択肢マスタシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createSelectOptionsSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.SELECT_OPTIONS;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    if (sheet.getLastRow() > 1) {
      logInfo(`    データ既存、スキップ`);
      return sheet;
    }
  }

  // ヘッダー設定
  const headers = SELECT_OPTIONS_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#fbbc04');
  headerRange.setFontColor('#000000');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(SELECT_OPTIONS_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // データ投入
  const data = getSelectOptionsRows();
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    logInfo(`    データ投入: ${data.length}件`);
  }

  return sheet;
}

/**
 * 健診コースマスタシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createExamCourseMasterSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.EXAM_COURSE_MASTER;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    if (sheet.getLastRow() > 1) {
      logInfo(`    データ既存、スキップ`);
      return sheet;
    }
  }

  // ヘッダー設定
  const headers = EXAM_COURSE_MASTER_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(EXAM_COURSE_MASTER_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // データ投入
  const data = getExamCourseMasterRows();
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    logInfo(`    データ投入: ${data.length}件`);
  }

  return sheet;
}

/**
 * 拡張マスタシートを非表示に設定
 */
function hideExtendedMasterSheets() {
  const ss = getSpreadsheet();
  const sheetNames = [
    DB_CONFIG.SHEETS.EXAM_ITEM_MASTER,
    DB_CONFIG.SHEETS.JUDGMENT_CRITERIA,
    DB_CONFIG.SHEETS.SELECT_OPTIONS,
    DB_CONFIG.SHEETS.EXAM_COURSE_MASTER
  ];

  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.hideSheet();
      logInfo(`シート非表示: ${sheetName}`);
    }
  }
}

/**
 * 拡張マスタシートを表示（開発・デバッグ用）
 */
function showExtendedMasterSheets() {
  const ss = getSpreadsheet();
  const sheetNames = [
    DB_CONFIG.SHEETS.EXAM_ITEM_MASTER,
    DB_CONFIG.SHEETS.JUDGMENT_CRITERIA,
    DB_CONFIG.SHEETS.SELECT_OPTIONS,
    DB_CONFIG.SHEETS.EXAM_COURSE_MASTER
  ];

  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.showSheet();
      logInfo(`シート表示: ${sheetName}`);
    }
  }
}

// ============================================
// Phase 2: 結果入力機能用シート作成（iD-Heart準拠）
// ============================================

/**
 * Phase 2: 結果入力機能用シートをセットアップ
 * M_検査所見マスタ、M_団体マスタ、M_コース項目マスタ、T_判定結果、T_所見を作成
 */
function setupResultInputSheets() {
  logInfo('===== Phase 2: 結果入力機能用シート作成開始 =====');

  try {
    const ss = getSpreadsheet();

    // 1. M_検査所見マスタ（所見テンプレート）
    createFindingTemplateSheet(ss);

    // 2. M_団体マスタ（企業・団体情報）
    createOrganizationMasterSheet(ss);

    // 3. M_コース項目マスタ（コースと検査項目のリレーション）
    createCourseItemSheet(ss);

    // 4. T_判定結果（3レベル判定）
    createJudgmentResultSheet(ss);

    // 5. T_所見（所見記録）
    createFindingsSheet(ss);

    logInfo('===== Phase 2: 結果入力機能用シート作成完了 =====');

    // 結果をダイアログで表示
    const ui = SpreadsheetApp.getUi();
    ui.alert('セットアップ完了',
      '結果入力機能用シートの作成が完了しました。\n\n' +
      '作成されたシート:\n' +
      '- M_検査所見マスタ（所見テンプレート）\n' +
      '- M_団体マスタ（企業・団体情報）\n' +
      '- M_コース項目マスタ（コース-項目関連）\n' +
      '- T_判定結果（3レベル判定結果）\n' +
      '- T_所見（所見記録）',
      ui.ButtonSet.OK);

  } catch (e) {
    logError('setupResultInputSheets', e);
    throw e;
  }
}

/**
 * M_検査所見マスタシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createFindingTemplateSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.FINDING_TEMPLATE;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    if (sheet.getLastRow() > 1) {
      logInfo(`    データ既存、スキップ`);
      return sheet;
    }
  }

  // ヘッダー設定
  const headers = FINDING_TEMPLATE_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#9c27b0');  // 紫
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(FINDING_TEMPLATE_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // サンプルデータ投入
  const sampleData = [
    ['F001', 'BMI', '身体計測', 'B', '体重管理が必要です。', 1, true],
    ['F002', 'BMI', '身体計測', 'C', '肥満傾向です。生活習慣の改善を推奨します。', 1, true],
    ['F003', 'BMI', '身体計測', 'D', '高度肥満です。医療機関での相談を推奨します。', 1, true],
    ['F004', 'BP_SYS', '血圧', 'B', '血圧がやや高めです。塩分控えめの食事を心がけてください。', 1, true],
    ['F005', 'BP_SYS', '血圧', 'C', '高血圧の傾向があります。定期的な血圧測定を推奨します。', 1, true],
    ['F006', 'BP_SYS', '血圧', 'D', '高血圧です。医療機関での精密検査を推奨します。', 1, true],
    ['F007', 'FBS', '糖代謝', 'C', '血糖値がやや高めです。食生活の見直しを推奨します。', 1, true],
    ['F008', 'FBS', '糖代謝', 'D', '糖尿病の疑いがあります。医療機関での精密検査を推奨します。', 1, true],
    ['F009', 'LDL', '脂質', 'C', 'LDLコレステロールが高めです。食事・運動の改善を推奨します。', 1, true],
    ['F010', 'LDL', '脂質', 'D', '脂質異常症の疑いがあります。医療機関での精密検査を推奨します。', 1, true]
  ];

  if (sampleData.length > 0) {
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    logInfo(`    サンプルデータ投入: ${sampleData.length}件`);
  }

  return sheet;
}

/**
 * M_団体マスタシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createOrganizationMasterSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.ORGANIZATION_MASTER;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    if (sheet.getLastRow() > 1) {
      logInfo(`    データ既存、スキップ`);
      return sheet;
    }
  }

  // ヘッダー設定
  const headers = ORGANIZATION_MASTER_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#00bcd4');  // シアン
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(ORGANIZATION_MASTER_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // サンプルデータ投入
  const sampleData = [
    ['ORG001', '個人（一般）', '', '', '', '', 'DOCK_LIFE,DOCK_GI,DOCK_FULL', '都度', '個人受診者用', true],
    ['ORG002', 'サンプル株式会社', '150-0001', '東京都渋谷区神宮前1-1-1', '03-1234-5678', '健康管理部 山田太郎', 'DOCK_LIFE', '一括', '年間契約', true]
  ];

  if (sampleData.length > 0) {
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    logInfo(`    サンプルデータ投入: ${sampleData.length}件`);
  }

  return sheet;
}

/**
 * M_コース項目マスタシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createCourseItemSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.COURSE_ITEM;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    if (sheet.getLastRow() > 1) {
      logInfo(`    データ既存、スキップ`);
      return sheet;
    }
  }

  // ヘッダー設定
  const headers = COURSE_ITEM_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#ff9800');  // オレンジ
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(COURSE_ITEM_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // サンプルデータ投入（生活習慣病ドックの項目例）
  const sampleData = [
    ['DOCK_LIFE', 'HEIGHT', true, 1],
    ['DOCK_LIFE', 'WEIGHT', true, 2],
    ['DOCK_LIFE', 'BMI', true, 3],
    ['DOCK_LIFE', 'BP_SYS', true, 4],
    ['DOCK_LIFE', 'BP_DIA', true, 5],
    ['DOCK_LIFE', 'FBS', true, 10],
    ['DOCK_LIFE', 'HBA1C', true, 11],
    ['DOCK_LIFE', 'TCHO', true, 20],
    ['DOCK_LIFE', 'HDL', true, 21],
    ['DOCK_LIFE', 'LDL', true, 22],
    ['DOCK_LIFE', 'TG', true, 23],
    ['DOCK_LIFE', 'AST', true, 30],
    ['DOCK_LIFE', 'ALT', true, 31],
    ['DOCK_LIFE', 'GGT', true, 32]
  ];

  if (sampleData.length > 0) {
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    logInfo(`    サンプルデータ投入: ${sampleData.length}件`);
  }

  return sheet;
}

/**
 * T_判定結果シートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createJudgmentResultSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.JUDGMENT_RESULT;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}`);
    return sheet;
  }

  // ヘッダー設定
  const headers = JUDGMENT_RESULT_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e91e63');  // ピンク
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(JUDGMENT_RESULT_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // トランザクションシートなのでサンプルデータは投入しない
  logInfo(`    トランザクションシート作成完了`);

  return sheet;
}

/**
 * T_所見シートを作成（縦持ち・検査項目別構造）
 * @param {Spreadsheet} ss - スプレッドシート
 */
function createFindingsSheet(ss) {
  const sheetName = DB_CONFIG.SHEETS.FINDINGS;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`  シート作成: ${sheetName}`);
  } else {
    logInfo(`  シート既存: ${sheetName}（再作成が必要な場合は削除してから実行）`);
    return sheet;
  }

  // ヘッダー設定
  const headers = FINDINGS_DEF.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#607d8b');  // グレー
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  for (const [col, width] of Object.entries(FINDINGS_DEF.columnWidths)) {
    const colIndex = columnLetterToIndex(col);
    sheet.setColumnWidth(colIndex, width);
  }

  // 1行目を固定
  sheet.setFrozenRows(1);

  // 判定列（F列）にデータバリデーション設定
  const judgmentCol = FINDINGS_DEF.columns.JUDGMENT + 1;
  const judgmentRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['A', 'B', 'C', 'D', 'E', 'F', ''], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, judgmentCol, 1000, 1).setDataValidation(judgmentRule);

  // 検査日列（H列）に日付フォーマット設定
  const examDateCol = FINDINGS_DEF.columns.EXAM_DATE + 1;
  sheet.getRange(2, examDateCol, 1000, 1).setNumberFormat('yyyy/mm/dd');

  // 作成日時・更新日時列（J,K列）に日時フォーマット設定
  const createdAtCol = FINDINGS_DEF.columns.CREATED_AT + 1;
  const updatedAtCol = FINDINGS_DEF.columns.UPDATED_AT + 1;
  sheet.getRange(2, createdAtCol, 1000, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sheet.getRange(2, updatedAtCol, 1000, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');

  // トランザクションシートなのでサンプルデータは投入しない
  logInfo(`    T_所見シート作成完了（縦持ち・検査項目別構造）`);

  return sheet;
}

/**
 * 結果入力機能用シートを非表示に設定
 */
function hideResultInputSheets() {
  const ss = getSpreadsheet();
  const sheetNames = [
    DB_CONFIG.SHEETS.FINDING_TEMPLATE,
    DB_CONFIG.SHEETS.ORGANIZATION_MASTER,
    DB_CONFIG.SHEETS.COURSE_ITEM,
    DB_CONFIG.SHEETS.JUDGMENT_RESULT,
    DB_CONFIG.SHEETS.FINDINGS
  ];

  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.hideSheet();
      logInfo(`シート非表示: ${sheetName}`);
    }
  }
}

/**
 * 結果入力機能用シートを表示（開発・デバッグ用）
 */
function showResultInputSheets() {
  const ss = getSpreadsheet();
  const sheetNames = [
    DB_CONFIG.SHEETS.FINDING_TEMPLATE,
    DB_CONFIG.SHEETS.ORGANIZATION_MASTER,
    DB_CONFIG.SHEETS.COURSE_ITEM,
    DB_CONFIG.SHEETS.JUDGMENT_RESULT,
    DB_CONFIG.SHEETS.FINDINGS
  ];

  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.showSheet();
      logInfo(`シート表示: ${sheetName}`);
    }
  }
}
