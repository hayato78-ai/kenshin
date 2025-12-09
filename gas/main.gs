/**
 * メインモジュール
 * エントリーポイントとトリガー管理
 */

// ============================================
// メイン処理関数
// ============================================

/**
 * CSV取込トリガー（フォルダ監視）
 * 毎時実行されるトリガーから呼び出される
 */
function onCsvUploaded() {
  logInfo('===== CSV取込処理開始 =====');

  try {
    // 新規CSVファイルを検索
    const newFiles = findNewCsvFiles();

    if (newFiles.length === 0) {
      logInfo('新規CSVファイルはありません');
      return;
    }

    logInfo(`${newFiles.length}件の新規CSVを検出`);

    let successCount = 0;
    let errorCount = 0;

    for (const file of newFiles) {
      const result = processCsvFile(file);

      if (result.success) {
        successCount += result.patientIds.length;
      } else {
        errorCount++;
      }
    }

    // 処理結果を通知
    if (successCount > 0 || errorCount > 0) {
      const subject = `【健診システム】CSV取込完了: ${successCount}名処理`;
      const body = `CSV取込処理が完了しました。\n\n` +
                   `処理成功: ${successCount}名\n` +
                   `エラー: ${errorCount}件\n\n` +
                   `処理日時: ${new Date().toLocaleString('ja-JP')}`;
      sendNotification(subject, body);
    }

  } catch (e) {
    logError('onCsvUploaded', e);
    sendNotification(
      '【健診システム】CSV取込エラー',
      `CSV取込処理でエラーが発生しました。\n\nエラー: ${e.message}`
    );
  }

  logInfo('===== CSV取込処理完了 =====');
}

/**
 * 患者データを処理
 * @param {string} patientId - 受診ID
 * @returns {Object} 処理結果
 */
function processPatient(patientId) {
  logInfo(`患者処理開始: ${patientId}`);

  try {
    // 所見を再生成
    const findings = regenerateFindings(patientId);

    logInfo(`患者処理完了: ${patientId}`);
    return { success: true, patientId, findings };

  } catch (e) {
    logError('processPatient', e);
    return { success: false, patientId, error: e.message };
  }
}

/**
 * 全患者を処理
 */
function processAll() {
  logInfo('===== 全患者処理開始 =====');

  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    logInfo('処理対象の患者がいません');
    return;
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let count = 0;

  for (const row of ids) {
    const patientId = row[0];
    const status = row[1];

    // 完了以外のステータスを処理
    if (patientId && status !== CONFIG.STATUS.COMPLETE) {
      processPatient(patientId);
      count++;
    }
  }

  logInfo(`===== ${count}名の処理完了 =====`);
}

/**
 * Excel出力（AppSheetから呼び出し用）
 * @param {string} patientId - 受診ID
 * @returns {string} 出力ファイルURL
 */
function exportPatientToExcel(patientId) {
  return exportToExcel(patientId);
}

// ============================================
// トリガー管理
// ============================================

/**
 * トリガーを設定
 */
function setupTriggers() {
  // 既存のトリガーを削除
  removeTriggers();

  // CSV監視トリガー（毎時）
  ScriptApp.newTrigger('onCsvUploaded')
    .timeBased()
    .everyHours(1)
    .create();

  // 日次アラートトリガー（毎日8:00）
  ScriptApp.newTrigger('dailyAlert')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  logInfo('トリガーを設定しました');
}

/**
 * トリガーを削除
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }

  logInfo('トリガーを削除しました');
}

/**
 * 日次アラート
 */
function dailyAlert() {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  let inputCount = 0;
  let pendingCount = 0;

  for (const row of data) {
    switch (row[1]) {
      case CONFIG.STATUS.INPUT:
        inputCount++;
        break;
      case CONFIG.STATUS.PENDING:
        pendingCount++;
        break;
    }
  }

  if (inputCount > 0 || pendingCount > 0) {
    const subject = '【健診システム】未処理データのお知らせ';
    const body = `未処理のデータがあります。\n\n` +
                 `入力中: ${inputCount}件\n` +
                 `確認待ち: ${pendingCount}件\n\n` +
                 `処理日時: ${new Date().toLocaleString('ja-JP')}`;
    sendNotification(subject, body);
  }
}

// ============================================
// 初期設定・メンテナンス
// ============================================

/**
 * 初期セットアップ
 * 初回実行時に呼び出す
 */
function initialSetup() {
  logInfo('===== 初期セットアップ開始 =====');

  // スプレッドシートの構造を確認
  validateSpreadsheetStructure();

  // トリガーを設定
  setupTriggers();

  logInfo('===== 初期セットアップ完了 =====');
}

/**
 * スプレッドシート構造を検証
 */
function validateSpreadsheetStructure() {
  const ss = getSpreadsheet();
  const requiredSheets = [
    CONFIG.SHEETS.PATIENT,
    CONFIG.SHEETS.PHYSICAL,
    CONFIG.SHEETS.BLOOD_TEST,
    CONFIG.SHEETS.FINDINGS,
    CONFIG.SHEETS.JUDGMENT_MASTER,
    CONFIG.SHEETS.FINDINGS_TEMPLATE
  ];

  const missingSheets = [];

  for (const sheetName of requiredSheets) {
    if (!ss.getSheetByName(sheetName)) {
      missingSheets.push(sheetName);
    }
  }

  if (missingSheets.length > 0) {
    throw new Error('必要なシートが見つかりません: ' + missingSheets.join(', '));
  }

  logInfo('スプレッドシート構造: OK');
}

/**
 * 設定を更新（設定シートから読み込み）
 */
function loadSettings() {
  const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  for (const row of data) {
    const key = row[0];
    const value = row[1];

    switch (key) {
      case 'CSV_FOLDER_ID':
        CONFIG.CSV_FOLDER_ID = value;
        break;
      case 'OUTPUT_FOLDER_ID':
        CONFIG.OUTPUT_FOLDER_ID = value;
        break;
    }
  }

  logInfo('設定を読み込みました');
}

// ============================================
// AppSheet連携用関数
// ============================================

/**
 * ステータスを更新
 * @param {string} patientId - 受診ID
 * @param {string} newStatus - 新しいステータス
 */
function updateStatus(patientId, newStatus) {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();
  const row = findPatientRow(sheet, patientId, lastRow);

  if (row > 0) {
    sheet.getRange(row, 2).setValue(newStatus);
    sheet.getRange(row, 14).setValue(new Date());
    logInfo(`ステータス更新: ${patientId} → ${newStatus}`);
  }
}

/**
 * 患者一覧を取得
 * @param {string} status - フィルタするステータス（省略で全件）
 * @returns {Array<Object>} 患者一覧
 */
function getPatientList(status) {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const patients = [];

  for (const row of data) {
    if (!status || row[1] === status) {
      patients.push({
        patientId: row[0],
        status: row[1],
        examDate: row[2],
        name: row[3],
        kana: row[4],
        gender: row[5],
        birthDate: row[6],
        age: row[7],
        overallJudgment: row[11]
      });
    }
  }

  return patients;
}

/**
 * 患者詳細を取得
 * @param {string} patientId - 受診ID
 * @returns {Object|null} 患者詳細
 */
function getPatientDetail(patientId) {
  const data = collectPatientData(patientId);
  return data;
}

// ============================================
// カルテNO指定処理
// ============================================

/**
 * 患者ID選択ダイアログを表示してCSVを取り込む
 * サイドバーでチェックボックス選択式
 */
function importByKarteNo() {
  const html = HtmlService.createHtmlOutput(getKarteNoSelectorHtml())
    .setTitle('患者ID選択')
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 患者ID選択用HTMLを生成
 * @returns {string} HTML文字列
 */
function getKarteNoSelectorHtml() {
  // CSVファイルから患者ID一覧を取得
  const karteList = scanAvailableKarteNos();

  if (karteList.length === 0) {
    return `
      <html>
        <body style="font-family: Arial, sans-serif; padding: 10px;">
          <h3>患者ID選択</h3>
          <p style="color: #666;">CSVフォルダに結果データがありません</p>
          <button onclick="google.script.host.close()">閉じる</button>
        </body>
      </html>
    `;
  }

  let checkboxes = '';
  for (const item of karteList) {
    const processed = item.processed ? ' (処理済)' : '';
    // 処理済みでも選択可能にする（再取込のため）
    const style = item.processed ? 'color: #888;' : '';
    checkboxes += `
      <label style="display: block; margin: 8px 0; ${style}">
        <input type="checkbox" name="karteNo" value="${item.karteNo}">
        ${item.karteNo}${processed}
        <span style="font-size: 11px; color: #888;">(${item.fileCount}ファイル)</span>
      </label>
    `;
  }

  return `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 10px; }
          h3 { margin-top: 0; color: #333; }
          .info { font-size: 12px; color: #666; margin-bottom: 15px; }
          .checkbox-container {
            max-height: 400px;
            overflow-y: auto;
            border: 1px solid #ddd;
            padding: 10px;
            margin-bottom: 15px;
          }
          .btn {
            padding: 10px 20px;
            margin-right: 10px;
            cursor: pointer;
            border: none;
            border-radius: 4px;
          }
          .btn-primary { background: #4285f4; color: white; }
          .btn-secondary { background: #f1f1f1; color: #333; }
          .btn:disabled { background: #ccc; cursor: not-allowed; }
          .select-all { margin-bottom: 10px; }
          #status { margin-top: 15px; padding: 10px; display: none; }
          .processing { background: #fff3cd; }
          .success { background: #d4edda; }
          .error { background: #f8d7da; }
        </style>
      </head>
      <body>
        <h3>患者ID選択</h3>
        <p class="info">取り込む患者IDを選択してください（${karteList.length}件）</p>

        <div class="select-all">
          <button class="btn btn-secondary" onclick="selectAll()">全選択</button>
          <button class="btn btn-secondary" onclick="deselectAll()">全解除</button>
        </div>

        <div class="checkbox-container">
          ${checkboxes}
        </div>

        <button id="importBtn" class="btn btn-primary" onclick="startImport()">取込開始</button>
        <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>

        <div id="status"></div>

        <script>
          function selectAll() {
            document.querySelectorAll('input[name="karteNo"]:not(:disabled)').forEach(cb => cb.checked = true);
          }

          function deselectAll() {
            document.querySelectorAll('input[name="karteNo"]').forEach(cb => cb.checked = false);
          }

          function startImport() {
            const selected = [];
            document.querySelectorAll('input[name="karteNo"]:checked').forEach(cb => {
              selected.push(cb.value);
            });

            if (selected.length === 0) {
              alert('カルテNOを選択してください');
              return;
            }

            document.getElementById('importBtn').disabled = true;
            showStatus('処理中... (' + selected.length + '件)', 'processing');

            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onError)
              .processSelectedKarteNos(selected);
          }

          function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = type;
            status.style.display = 'block';
          }

          function onSuccess(result) {
            showStatus('完了: ' + result.success + '件成功, ' + result.errors.length + '件エラー', 'success');
            document.getElementById('importBtn').disabled = false;

            // チェックボックスを更新（処理済みにする）
            if (result.processedKarteNos) {
              result.processedKarteNos.forEach(karteNo => {
                const cb = document.querySelector('input[value="' + karteNo + '"]');
                if (cb) {
                  cb.disabled = true;
                  cb.checked = false;
                  cb.parentElement.style.color = '#999';
                }
              });
            }
          }

          function onError(error) {
            showStatus('エラー: ' + error.message, 'error');
            document.getElementById('importBtn').disabled = false;
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * 利用可能な患者ID一覧をスキャン
 * 結果データCSVの中身から患者ID（2列目）を抽出
 * @returns {Array<Object>} 患者ID情報の配列
 */
function scanAvailableKarteNos() {
  const folder = getCsvFolder();
  const allFiles = [];
  findAllCsvFilesRecursive(folder, allFiles, 0, 100);

  // 患者IDごとにグループ化
  const karteMap = {};

  for (const file of allFiles) {
    const name = file.getName();
    const processed = name.startsWith('[済]');
    const baseName = name.replace(/^\[済\]/, '');

    // 結果データファイルのみ処理
    if (!isResultCsvFile(file)) {
      continue;
    }

    try {
      // CSVの中身から患者IDを抽出
      const content = readFileContent(file);
      const lines = content.trim().split('\n');

      for (const line of lines) {
        if (!line.trim()) continue;
        const fields = line.split(',');
        if (fields.length >= 2) {
          const patientId = fields[1].trim();  // 2列目が患者ID

          if (patientId && /^\d+$/.test(patientId)) {
            if (!karteMap[patientId]) {
              karteMap[patientId] = {
                karteNo: patientId,
                fileCount: 0,
                processed: true,
                files: [],
                fileNames: []
              };
            }

            // ファイルがまだ追加されていなければ追加
            if (!karteMap[patientId].fileNames.includes(name)) {
              karteMap[patientId].fileCount++;
              karteMap[patientId].files.push(file);
              karteMap[patientId].fileNames.push(name);

              if (!processed) {
                karteMap[patientId].processed = false;
              }
            }
          }
        }
      }
    } catch (e) {
      logError('scanAvailableKarteNos', e);
    }
  }

  // 配列に変換してソート
  const result = Object.values(karteMap);
  result.sort((a, b) => a.karteNo.localeCompare(b.karteNo));

  return result;
}

/**
 * 全CSVファイルを再帰的に検索（フィルタなし）
 */
function findAllCsvFilesRecursive(folder, results, depth = 0, maxFiles = 500) {
  if (depth > 5 || results.length >= maxFiles) {
    return;
  }

  const allFiles = folder.getFiles();
  while (allFiles.hasNext() && results.length < maxFiles) {
    const file = allFiles.next();
    const name = file.getName().toLowerCase();
    if (name.endsWith('.csv')) {
      results.push(file);
    }
  }

  const subFolders = folder.getFolders();
  while (subFolders.hasNext() && results.length < maxFiles) {
    const subFolder = subFolders.next();
    findAllCsvFilesRecursive(subFolder, results, depth + 1, maxFiles);
  }
}

/**
 * 選択された患者IDを処理
 * @param {Array<string>} patientIds - 選択された患者ID配列
 * @returns {Object} 処理結果
 */
function processSelectedKarteNos(patientIds) {
  logInfo(`患者ID指定取込開始: ${patientIds.join(', ')}`);

  const results = {
    success: 0,
    errors: [],
    processedKarteNos: []
  };

  // CSVファイルを再取得
  const folder = getCsvFolder();
  const allFiles = [];
  findAllCsvFilesRecursive(folder, allFiles, 0, 500);

  // 結果データファイルのみフィルタ（処理済みファイルも含める - 再取込対応）
  const resultFiles = allFiles.filter(f => isResultCsvFile(f));

  for (const file of resultFiles) {
    try {
      const content = readFileContent(file);
      const lines = content.trim().split('\n');

      // このファイルに含まれる選択された患者IDがあるか確認
      let hasTargetPatient = false;
      for (const line of lines) {
        const fields = line.split(',');
        if (fields.length >= 2) {
          const csvPatientId = fields[1].trim();
          if (patientIds.includes(csvPatientId)) {
            hasTargetPatient = true;
            break;
          }
        }
      }

      if (hasTargetPatient) {
        const result = processCsvFile(file);
        if (result.success) {
          results.success += result.patientIds.length;
          for (const pid of result.patientIds) {
            if (patientIds.includes(pid) && !results.processedKarteNos.includes(pid)) {
              results.processedKarteNos.push(pid);
            }
          }
          logInfo(`ファイル処理成功: ${file.getName()}`);
        }
      }
    } catch (e) {
      results.errors.push(`${file.getName()}: ${e.message}`);
      logError('processSelectedKarteNos', e);
    }
  }

  logInfo(`患者ID指定取込完了: 成功${results.success}件`);
  return results;
}

/**
 * 利用可能な患者ID一覧を表示（シンプル版）
 */
function showAvailableKarteNos() {
  const ui = SpreadsheetApp.getUi();
  const karteList = scanAvailableKarteNos();

  if (karteList.length === 0) {
    ui.alert('情報', 'CSVフォルダに結果データがありません', ui.ButtonSet.OK);
    return;
  }

  let message = `利用可能な患者ID (${karteList.length}件):\n\n`;

  for (const item of karteList.slice(0, 30)) {
    const status = item.processed ? '[済]' : '';
    message += `${status}${item.karteNo} (${item.fileCount}ファイル)\n`;
  }

  if (karteList.length > 30) {
    message += `\n... 他${karteList.length - 30}件`;
  }

  ui.alert('患者ID一覧', message, ui.ButtonSet.OK);
}

// ============================================
// テスト用関数
// ============================================

/**
 * テスト実行: CSV解析
 */
function testCsvParse() {
  const files = findNewCsvFiles();
  if (files.length > 0) {
    const result = parseCSV(files[0].getId());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    Logger.log('CSVファイルが見つかりません');
  }
}

/**
 * テスト実行: 判定処理
 */
function testJudgment() {
  // AST: 35 → B判定
  Logger.log('AST 35: ' + judge('AST_GOT', 35, 'M'));

  // HbA1c: 6.2 → C判定
  Logger.log('HbA1c 6.2: ' + judge('HBA1C', 6.2, 'M'));

  // Hb(男性): 14.0 → A判定
  Logger.log('Hb(M) 14.0: ' + judge('HEMOGLOBIN_M', 14.0, 'M'));

  // Hb(女性): 14.0 → A判定
  Logger.log('Hb(F) 14.0: ' + judge('HEMOGLOBIN_F', 14.0, 'F'));
}

/**
 * テスト実行: Excel出力
 */
function testExcelExport() {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const patientId = sheet.getRange(2, 1).getValue();
    const url = exportToExcel(patientId);
    Logger.log('出力URL: ' + url);
  } else {
    Logger.log('患者データがありません');
  }
}

// ============================================
// メニュー追加
// ============================================

/**
 * スプレッドシート開いた時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const currentProfile = CONFIG.getProfile();

  ui.createMenu('健診システム')
    // 労災二次検診サブメニュー
    .addSubMenu(ui.createMenu('🏥 労災二次検診')
      .addItem('データ取込（案件選択）', 'showRosaiSecondarySidebar')
      .addItem('受診者一覧（Excel出力）', 'showRosaiPatientListSidebar')
      .addItem('選択行をExcel出力', 'showExportDialogForSelectedRow')
      .addSeparator()
      .addItem('保健指導入力', 'showGuidanceInputForSelectedRow')
      .addSeparator()
      .addItem('入力シートを開く', 'activateRosaiInputSheet'))
    .addSeparator()
    // 健診種別切替メニュー
    .addSubMenu(ui.createMenu('🏥 健診種別')
      .addItem('人間ドックモード', 'setDockMode')
      .addItem('労災二次検診モード', 'setRosaiMode')
      .addSeparator()
      .addItem('現在のモード確認', 'showCurrentExamType'))
    .addSeparator()
    // 案件フォルダ選択メニュー
    .addSubMenu(ui.createMenu('📁 案件フォルダ')
      .addItem('フォルダを選択...', 'selectCsvFolder')
      .addItem('現在のフォルダ確認', 'showCurrentCsvFolder')
      .addSeparator()
      .addItem('デフォルトに戻す', 'resetCsvFolder'))
    .addSeparator()
    .addItem('CSV取込を実行（10件ずつ）', 'onCsvUploaded')
    .addItem('患者ID指定で取込', 'importByKarteNo')
    .addItem('患者ID一覧を表示', 'showAvailableKarteNos')
    .addSeparator()
    .addItem('全患者の所見を再生成', 'regenerateAllFindings')
    .addItem('確認待ち患者を一括出力', 'exportPendingPatients')
    .addSeparator()
    .addSubMenu(ui.createMenu('🤖 Claude AI')
      .addItem('Claudeに質問', 'showClaudeHelpDialog')
      .addItem('選択行の所見を生成', 'generateFindingsForSelectedPatient')
      .addItem('データ検証', 'validateInputSheetWithClaude')
      .addSeparator()
      .addItem('APIキー設定', 'setAnthropicApiKey'))
    .addSeparator()
    .addSubMenu(ui.createMenu('設定')
      .addItem('トリガーを設定', 'setupTriggers')
      .addItem('トリガーを削除', 'removeTriggers')
      .addItem('出力シートレイアウト再設定', 'resetOutputTemplateLayout')
      .addItem('所見テンプレート初期化', 'initializeFindingsTemplateSheet')
      .addItem('初期セットアップ', 'initialSetup'))
    .addToUi();

  // 現在のモードをログ出力
  logInfo(`現在の健診モード: ${currentProfile.name}`);
}

// ============================================
// 健診種別切替関数
// ============================================

/**
 * 人間ドックモードに切り替え
 */
function setDockMode() {
  CONFIG.setExamType('DOCK');
  const ui = SpreadsheetApp.getUi();
  ui.alert('モード切替', '人間ドックモードに切り替えました。\n\nCSV形式: BML\nテンプレート: iD-Heart形式', ui.ButtonSet.OK);
  logInfo('健診種別を人間ドックに変更');
}

/**
 * 労災二次検診モードに切り替え
 */
function setRosaiMode() {
  CONFIG.setExamType('ROSAI_SECONDARY');
  const ui = SpreadsheetApp.getUi();
  ui.alert('モード切替', '労災二次検診モードに切り替えました。\n\nCSV形式: 標準形式（ヘッダー付き）\nテンプレート: 個人票形式', ui.ButtonSet.OK);
  logInfo('健診種別を労災二次検診に変更');
}

/**
 * 現在の健診種別を表示
 */
function showCurrentExamType() {
  const profile = CONFIG.getProfile();
  const ui = SpreadsheetApp.getUi();

  const message = `現在の健診モード: ${profile.name}\n\n` +
                  `コード: ${profile.code}\n` +
                  `CSV形式: ${profile.csvFormat}\n` +
                  `有効項目数: ${profile.enabledItems.length}項目\n\n` +
                  `【有効項目】\n${profile.enabledItems.join(', ')}`;

  ui.alert('健診種別情報', message, ui.ButtonSet.OK);
}

// ============================================
// 案件フォルダ選択機能
// ============================================

/**
 * CSVフォルダを選択（案件ごとに変更可能）
 * フォルダIDまたはURLを入力するダイアログを表示
 */
function selectCsvFolder() {
  const ui = SpreadsheetApp.getUi();

  // 現在のフォルダ情報を取得
  const currentFolderId = getTempCsvFolderId();
  const defaultFolderId = getSettingValue('CSV_FOLDER_ID');

  let prompt = '案件のCSVフォルダIDまたはURLを入力してください。\n\n';
  prompt += '例:\n';
  prompt += '・フォルダID: 1ABC123xyz...\n';
  prompt += '・URL: https://drive.google.com/drive/folders/1ABC123xyz...\n\n';

  if (currentFolderId) {
    prompt += `【現在の設定】\n一時フォルダ: ${currentFolderId}\n`;
  } else {
    prompt += `【現在の設定】\nデフォルト: ${defaultFolderId || '未設定'}\n`;
  }

  const response = ui.prompt('📁 案件フォルダ選択', prompt, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const input = response.getResponseText().trim();

  if (!input) {
    ui.alert('エラー', 'フォルダIDまたはURLを入力してください。', ui.ButtonSet.OK);
    return;
  }

  // URLからフォルダIDを抽出
  let folderId = extractFolderIdFromInput(input);

  if (!folderId) {
    ui.alert('エラー', '有効なフォルダIDまたはURLを入力してください。', ui.ButtonSet.OK);
    return;
  }

  // フォルダの存在確認
  try {
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();

    // 一時フォルダIDをScript Propertiesに保存
    PropertiesService.getScriptProperties().setProperty('TEMP_CSV_FOLDER_ID', folderId);

    ui.alert('設定完了',
      `CSVフォルダを設定しました。\n\n` +
      `フォルダ名: ${folderName}\n` +
      `フォルダID: ${folderId}\n\n` +
      `※この設定は「デフォルトに戻す」で解除できます。`,
      ui.ButtonSet.OK);

    logInfo(`案件フォルダを設定: ${folderName} (${folderId})`);

  } catch (e) {
    ui.alert('エラー',
      `フォルダにアクセスできません。\n\n` +
      `ID: ${folderId}\n` +
      `エラー: ${e.message}\n\n` +
      `フォルダIDが正しいか、アクセス権限があるか確認してください。`,
      ui.ButtonSet.OK);
  }
}

/**
 * 入力値からフォルダIDを抽出
 * @param {string} input - フォルダIDまたはURL
 * @returns {string|null} フォルダID
 */
function extractFolderIdFromInput(input) {
  // 既にIDの場合（英数字とハイフン、アンダースコアのみ）
  if (/^[\w-]+$/.test(input) && input.length > 10) {
    return input;
  }

  // Google DriveのURL形式
  // https://drive.google.com/drive/folders/FOLDER_ID
  // https://drive.google.com/drive/u/0/folders/FOLDER_ID
  const patterns = [
    /\/folders\/([^/?]+)/,
    /id=([^&]+)/
  ];

  for (const pattern of patterns) {
    const match = input.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }

  return null;
}

/**
 * 一時CSVフォルダIDを取得
 * @returns {string|null} 一時フォルダID（未設定の場合はnull）
 */
function getTempCsvFolderId() {
  return PropertiesService.getScriptProperties().getProperty('TEMP_CSV_FOLDER_ID');
}

/**
 * 現在のCSVフォルダ情報を表示
 */
function showCurrentCsvFolder() {
  const ui = SpreadsheetApp.getUi();

  const tempFolderId = getTempCsvFolderId();
  const defaultFolderId = getSettingValue('CSV_FOLDER_ID');

  let message = '';
  let currentFolderId = null;

  if (tempFolderId) {
    currentFolderId = tempFolderId;
    message += '【現在の設定】一時フォルダ（案件指定）\n\n';
  } else if (defaultFolderId && defaultFolderId !== 'YOUR_CSV_FOLDER_ID') {
    currentFolderId = defaultFolderId;
    message += '【現在の設定】デフォルトフォルダ（設定シート）\n\n';
  } else {
    ui.alert('CSVフォルダ情報',
      'CSVフォルダが設定されていません。\n\n' +
      '「フォルダを選択...」から案件フォルダを指定するか、\n' +
      '設定シートにCSV_FOLDER_IDを設定してください。',
      ui.ButtonSet.OK);
    return;
  }

  try {
    const folder = DriveApp.getFolderById(currentFolderId);
    const folderName = folder.getName();
    const folderUrl = folder.getUrl();

    // フォルダ内のCSVファイル数をカウント
    const csvFiles = [];
    findAllCsvFilesRecursive(folder, csvFiles, 0, 100);

    message += `フォルダ名: ${folderName}\n`;
    message += `フォルダID: ${currentFolderId}\n`;
    message += `CSVファイル数: ${csvFiles.length}件\n\n`;
    message += `URL: ${folderUrl}`;

    ui.alert('CSVフォルダ情報', message, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('エラー',
      `フォルダにアクセスできません。\n\n` +
      `ID: ${currentFolderId}\n` +
      `エラー: ${e.message}`,
      ui.ButtonSet.OK);
  }
}

/**
 * CSVフォルダをデフォルトに戻す
 * 一時フォルダ設定を削除
 */
function resetCsvFolder() {
  const ui = SpreadsheetApp.getUi();

  const tempFolderId = getTempCsvFolderId();

  if (!tempFolderId) {
    ui.alert('情報', '一時フォルダは設定されていません。\n既にデフォルト設定を使用しています。', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert('確認',
    '一時フォルダ設定を削除して、デフォルトフォルダに戻しますか？\n\n' +
    `現在の一時フォルダ: ${tempFolderId}`,
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  // 一時フォルダ設定を削除
  PropertiesService.getScriptProperties().deleteProperty('TEMP_CSV_FOLDER_ID');

  const defaultFolderId = getSettingValue('CSV_FOLDER_ID');

  ui.alert('完了',
    `CSVフォルダをデフォルトに戻しました。\n\n` +
    `デフォルトフォルダ: ${defaultFolderId || '未設定'}`,
    ui.ButtonSet.OK);

  logInfo('CSVフォルダをデフォルトに戻しました');
}

// ============================================
// 労災二次検診 サイドバー機能
// ============================================

/**
 * 労災二次検診サイドバーを表示
 */
function showRosaiSecondarySidebar() {
  const html = HtmlService.createHtmlOutput(getRosaiSecondarySidebarHtml())
    .setTitle('労災二次検診')
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 労災二次検診サイドバーのHTMLを生成
 * @returns {string} HTML
 */
function getRosaiSecondarySidebarHtml() {
  // 案件一覧を取得
  let caseListHtml = '';
  let errorMessage = '';

  try {
    const cases = getRosaiCaseList();

    if (cases.length === 0) {
      errorMessage = '案件フォルダが見つかりません。<br>設定シートのROSAI_CASE_FOLDER_IDを確認してください。';
    } else {
      for (const c of cases) {
        const csvStatus = c.hasCsv ? '✅ CSV有' : '⬜ CSV無';
        const csvClass = c.hasCsv ? 'csv-ready' : 'csv-none';

        caseListHtml += `
          <option value="${c.folderId}" data-date="${c.date}" data-company="${c.companyName}">
            ${c.dateFormatted} ${c.companyName}
          </option>
        `;
      }
    }
  } catch (e) {
    errorMessage = `エラー: ${e.message}<br>設定シートのROSAI_CASE_FOLDER_IDを確認してください。`;
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
        .section {
          margin-bottom: 20px;
          padding: 15px;
          background: #f8f9fa;
          border-radius: 8px;
        }
        .section-title {
          font-weight: bold;
          margin-bottom: 10px;
          color: #333;
        }
        label {
          display: block;
          margin-bottom: 5px;
          font-weight: 500;
          color: #555;
        }
        select, input[type="text"] {
          width: 100%;
          padding: 8px;
          margin-bottom: 10px;
          border: 1px solid #ddd;
          border-radius: 4px;
          box-sizing: border-box;
          font-size: 13px;
        }
        select:focus, input:focus {
          border-color: #1a73e8;
          outline: none;
        }
        .btn {
          display: block;
          width: 100%;
          padding: 12px;
          margin-top: 10px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
          font-weight: 500;
        }
        .btn-primary {
          background: #1a73e8;
          color: white;
        }
        .btn-primary:hover {
          background: #1557b0;
        }
        .btn-primary:disabled {
          background: #ccc;
          cursor: not-allowed;
        }
        .btn-secondary {
          background: #f1f3f4;
          color: #333;
        }
        .btn-secondary:hover {
          background: #e8eaed;
        }
        #status {
          margin-top: 15px;
          padding: 10px;
          border-radius: 4px;
          display: none;
        }
        .status-processing {
          background: #fff3cd;
          color: #856404;
        }
        .status-success {
          background: #d4edda;
          color: #155724;
        }
        .status-error {
          background: #f8d7da;
          color: #721c24;
        }
        .info-text {
          font-size: 11px;
          color: #666;
          margin-top: 5px;
        }
        .error-box {
          background: #f8d7da;
          color: #721c24;
          padding: 10px;
          border-radius: 4px;
          margin-bottom: 15px;
        }
        .csv-status {
          font-size: 11px;
          margin-left: 5px;
        }
        .csv-ready { color: #28a745; }
        .csv-none { color: #dc3545; }
      </style>
    </head>
    <body>
      <h3>🏥 労災二次検診</h3>

      ${errorMessage ? `<div class="error-box">${errorMessage}</div>` : ''}

      <div class="section">
        <div class="section-title">1. 案件選択</div>
        <label for="caseSelect">案件（日程）:</label>
        <select id="caseSelect" ${errorMessage ? 'disabled' : ''}>
          <option value="">選択してください</option>
          ${caseListHtml}
        </select>
        <div class="info-text">※ 10_案件フォルダから自動取得</div>
      </div>

      <div class="section">
        <div class="section-title">2. 共通情報</div>
        <label for="doctorName">担当医師名:</label>
        <input type="text" id="doctorName" placeholder="例: 田中 太郎">
        <div class="info-text">※ いつでも入力シートで編集可能</div>
      </div>

      <button id="startBtn" class="btn btn-primary" onclick="startImport()" ${errorMessage ? 'disabled' : ''}>
        データ取込開始
      </button>

      <button class="btn btn-secondary" onclick="openInputSheet()" style="margin-top: 10px;">
        入力シートを開く
      </button>

      <div id="status"></div>

      <script>
        function showStatus(message, type) {
          const status = document.getElementById('status');
          status.innerHTML = message;
          status.className = 'status-' + type;
          status.style.display = 'block';
        }

        function startImport() {
          const caseSelect = document.getElementById('caseSelect');
          const doctorName = document.getElementById('doctorName').value.trim();

          if (!caseSelect.value) {
            alert('案件を選択してください');
            return;
          }

          document.getElementById('startBtn').disabled = true;
          showStatus('処理中... CSVを検索しています', 'processing');

          google.script.run
            .withSuccessHandler(onImportSuccess)
            .withFailureHandler(onImportError)
            .createRosaiInputSheet(caseSelect.value, doctorName);
        }

        function onImportSuccess(result) {
          document.getElementById('startBtn').disabled = false;

          if (result.success) {
            showStatus(
              '✅ 完了: ' + result.patientCount + '名のデータを取込みました<br>' +
              '<a href="' + result.sheetUrl + '" target="_blank">入力シートを開く</a>',
              'success'
            );
          } else {
            showStatus('❌ エラー: ' + result.error, 'error');
          }
        }

        function onImportError(error) {
          document.getElementById('startBtn').disabled = false;
          showStatus('❌ エラー: ' + error.message, 'error');
        }

        function openInputSheet() {
          google.script.run.activateRosaiInputSheet();
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * 労災二次検診入力シートをアクティブにする
 */
function activateRosaiInputSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('労災二次検診_入力');

  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    const ui = SpreadsheetApp.getUi();
    ui.alert('情報', '入力シートがまだ作成されていません。\n先にデータ取込を実行してください。', ui.ButtonSet.OK);
  }
}

/**
 * onEditトリガーに労災二次検診の処理を追加
 * ※既存のonEditに統合する場合はこの関数を呼び出す
 */
function onEdit(e) {
  // 労災二次検診の編集処理
  onEditRosaiSecondary(e);
}

/**
   * デバッグ用テスト関数
   * GASエディタで直接実行して設定を確認
   */
  function testRosaiSetup() {
    console.log('=== 労災二次検診 設定テスト ===');

    // 1. 設定値の確認
    const caseFolderId =
  getSettingValue('ROSAI_CASE_FOLDER_ID');
    const csvFolderId =
  getSettingValue('ROSAI_CSV_FOLDER_ID');
    console.log('ROSAI_CASE_FOLDER_ID:', caseFolderId
  || '未設定');
    console.log('ROSAI_CSV_FOLDER_ID:', csvFolderId ||
   '未設定');

    // 2. getRosaiBaseFolderId()の結果
    const baseFolderId = getRosaiBaseFolderId();
    console.log('getRosaiBaseFolderId() 結果:',
  baseFolderId || 'null');

    if (!baseFolderId) {
      console.log('❌ フォルダIDが取得できません');
      return;
    }

    // 3. フォルダアクセス確認
    try {
      const folder =
  DriveApp.getFolderById(baseFolderId);
      console.log('✅ フォルダ名:', folder.getName());

      // 4. サブフォルダ一覧
      const subFolders = folder.getFolders();
      let count = 0;
      while (subFolders.hasNext() && count < 5) {
        const sub = subFolders.next();
        console.log('  サブフォルダ:', sub.getName());
        count++;
      }
      console.log('=== テスト完了 ===');
    } catch (e) {
      console.log('❌ フォルダアクセスエラー:',
  e.message);
    }
  }