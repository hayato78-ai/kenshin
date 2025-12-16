/**
 * 健診結果DB 統合システム - メイン
 *
 * @description 設計書: 健診結果DB_設計書_v1.md
 * @version 1.0.0
 * @date 2025-12-14
 */

// ============================================
// メニュー設定
// ============================================

/**
 * スプレッドシート起動時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('健診DB')
    .addSubMenu(ui.createMenu('📋 セットアップ')
      .addItem('DB初期セットアップ', 'setupDatabase')
      .addItem('マスタデータリセット', 'resetMasterData')
      .addSeparator()
      .addItem('シート表示（開発用）', 'showAllSheets')
      .addItem('シート非表示', 'hideAllSheetsMenu')
      .addSeparator()
      .addItem('DB整合性チェック', 'validateDatabaseMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('👤 受診者')
      .addItem('受診者検索', 'showPatientSearchDialog')
      .addItem('新規受診者登録', 'showNewPatientDialog'))
    .addSubMenu(ui.createMenu('📝 受診記録')
      .addItem('新規受診登録', 'showNewVisitDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔬 判定')
      .addItem('判定テスト実行', 'testJudgmentEngine')
      .addItem('受診者の判定再計算', 'recalculateJudgmentMenu'))
    .addSeparator()
    .addItem('ℹ️ バージョン情報', 'showVersionInfo')
    .addToUi();
}

// ============================================
// メニューアクション関数
// ============================================

/**
 * DB整合性チェック（メニュー用）
 */
function validateDatabaseMenu() {
  const issues = validateDatabase();
  const ui = SpreadsheetApp.getUi();

  if (issues.length === 0) {
    ui.alert('整合性チェック', 'データベースに問題はありません。', ui.ButtonSet.OK);
  } else {
    ui.alert('整合性チェック',
      `${issues.length}件の問題が見つかりました:\n\n` +
      issues.map(i => '・' + i).join('\n'),
      ui.ButtonSet.OK);
  }
}

/**
 * シート非表示（メニュー用）
 */
function hideAllSheetsMenu() {
  const ss = getSpreadsheet();
  hideAllSheets(ss);

  const ui = SpreadsheetApp.getUi();
  ui.alert('完了', '全シートを非表示にしました。', ui.ButtonSet.OK);
}

/**
 * バージョン情報を表示
 */
function showVersionInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('健診結果DB 統合システム',
    'バージョン: 1.0.0\n' +
    '作成日: 2025-12-14\n\n' +
    '設計書: 健診結果DB_設計書_v1.md\n' +
    'UI設計書: 健診結果DB_UI設計書_v1.md\n\n' +
    'Phase 1: 基盤構築',
    ui.ButtonSet.OK);
}

// ============================================
// 受診者ダイアログ
// ============================================

/**
 * 受診者検索ダイアログを表示
 */
function showPatientSearchDialog() {
  const html = HtmlService.createHtmlOutput(getPatientSearchHtml())
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, '受診者検索');
}

/**
 * 新規受診者登録ダイアログを表示
 */
function showNewPatientDialog() {
  const html = HtmlService.createHtmlOutput(getNewPatientHtml())
    .setWidth(500)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, '新規受診者登録');
}

/**
 * 受診者検索HTML
 */
function getPatientSearchHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; padding: 15px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: 500; }
        input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
        .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-secondary { background: #f1f3f4; color: #333; }
        #results { margin-top: 20px; }
        .result-item { padding: 10px; border-bottom: 1px solid #eee; cursor: pointer; }
        .result-item:hover { background: #f5f5f5; }
      </style>
    </head>
    <body>
      <div class="form-group">
        <label>氏名（部分一致）</label>
        <input type="text" id="name" placeholder="例: 山田">
      </div>
      <div class="form-group">
        <label>カナ（部分一致）</label>
        <input type="text" id="kana" placeholder="例: ヤマダ">
      </div>
      <div class="form-group">
        <label>所属企業（部分一致）</label>
        <input type="text" id="company" placeholder="例: ○○株式会社">
      </div>
      <button class="btn btn-primary" onclick="search()">検索</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">閉じる</button>

      <div id="results"></div>

      <script>
        function search() {
          const criteria = {
            name: document.getElementById('name').value,
            kana: document.getElementById('kana').value,
            company: document.getElementById('company').value
          };

          google.script.run
            .withSuccessHandler(showResults)
            .withFailureHandler(showError)
            .searchPatients(criteria);
        }

        function showResults(patients) {
          const resultsDiv = document.getElementById('results');

          if (patients.length === 0) {
            resultsDiv.innerHTML = '<p>該当する受診者が見つかりませんでした。</p>';
            return;
          }

          let html = '<p>' + patients.length + '件見つかりました</p>';
          for (const p of patients) {
            html += '<div class="result-item" onclick="selectPatient(\\'' + p.patientId + '\\')">';
            html += '<strong>' + p.patientId + '</strong> ' + p.name + ' (' + p.kana + ')';
            if (p.company) html += '<br><small>' + p.company + '</small>';
            html += '</div>';
          }

          resultsDiv.innerHTML = html;
        }

        function showError(error) {
          document.getElementById('results').innerHTML = '<p style="color:red;">エラー: ' + error.message + '</p>';
        }

        function selectPatient(patientId) {
          // 選択した受診者IDをクリップボードにコピー
          navigator.clipboard.writeText(patientId).then(() => {
            alert('受診者ID: ' + patientId + ' をコピーしました');
          });
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * 新規受診者登録HTML
 */
function getNewPatientHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; padding: 15px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: 500; }
        input, select { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
        .required { color: red; }
        .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-secondary { background: #f1f3f4; color: #333; }
        .row { display: flex; gap: 15px; }
        .col { flex: 1; }
        #status { margin-top: 15px; padding: 10px; border-radius: 4px; display: none; }
        .success { background: #d4edda; color: #155724; }
        .error { background: #f8d7da; color: #721c24; }
      </style>
    </head>
    <body>
      <div class="form-group">
        <label>氏名 <span class="required">*</span></label>
        <input type="text" id="name" placeholder="山田 太郎">
      </div>
      <div class="form-group">
        <label>カナ <span class="required">*</span></label>
        <input type="text" id="kana" placeholder="ヤマダ タロウ">
      </div>
      <div class="row">
        <div class="col">
          <div class="form-group">
            <label>生年月日 <span class="required">*</span></label>
            <input type="date" id="birthdate">
          </div>
        </div>
        <div class="col">
          <div class="form-group">
            <label>性別 <span class="required">*</span></label>
            <select id="gender">
              <option value="">選択</option>
              <option value="M">男性</option>
              <option value="F">女性</option>
            </select>
          </div>
        </div>
      </div>
      <div class="form-group">
        <label>電話番号</label>
        <input type="tel" id="phone" placeholder="090-1234-5678">
      </div>
      <div class="form-group">
        <label>所属企業</label>
        <input type="text" id="company" placeholder="○○株式会社">
      </div>

      <button class="btn btn-primary" onclick="register()">登録</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>

      <div id="status"></div>

      <script>
        function register() {
          const name = document.getElementById('name').value;
          const kana = document.getElementById('kana').value;
          const birthdate = document.getElementById('birthdate').value;
          const gender = document.getElementById('gender').value;

          if (!name || !kana || !birthdate || !gender) {
            showStatus('必須項目を入力してください', 'error');
            return;
          }

          const data = {
            name: name,
            kana: kana,
            birthdate: birthdate,
            gender: gender,
            phone: document.getElementById('phone').value,
            company: document.getElementById('company').value
          };

          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onError)
            .createPatient(data);
        }

        function showStatus(message, type) {
          const status = document.getElementById('status');
          status.textContent = message;
          status.className = type;
          status.style.display = 'block';
        }

        function onSuccess(patientId) {
          showStatus('受診者を登録しました: ' + patientId, 'success');
          // 3秒後にダイアログを閉じる
          setTimeout(() => google.script.host.close(), 3000);
        }

        function onError(error) {
          showStatus('エラー: ' + error.message, 'error');
        }
      </script>
    </body>
    </html>
  `;
}

// ============================================
// 受診登録ダイアログ
// ============================================

/**
 * 新規受診登録ダイアログを表示
 */
function showNewVisitDialog() {
  const html = HtmlService.createHtmlOutput(getNewVisitHtml())
    .setWidth(500)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '新規受診登録');
}

/**
 * 新規受診登録HTML
 */
function getNewVisitHtml() {
  // 検診種別とコースをサーバーサイドで取得
  const examTypes = getExamTypeMaster();
  const courses = getCourseMaster();

  let examTypeOptions = '<option value="">選択</option>';
  for (const et of examTypes) {
    examTypeOptions += `<option value="${et.typeId}" data-course-required="${et.courseRequired}">${et.typeName}</option>`;
  }

  let courseOptions = '<option value="">選択（人間ドックの場合）</option>';
  for (const c of courses) {
    courseOptions += `<option value="${c.courseId}">${c.courseName} (${c.price.toLocaleString()}円)</option>`;
  }

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; padding: 15px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: 500; }
        input, select { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
        .required { color: red; }
        .btn { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-right: 10px; }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-secondary { background: #f1f3f4; color: #333; }
        #status { margin-top: 15px; padding: 10px; border-radius: 4px; display: none; }
        .success { background: #d4edda; color: #155724; }
        .error { background: #f8d7da; color: #721c24; }
        .info { font-size: 12px; color: #666; margin-top: 5px; }
      </style>
    </head>
    <body>
      <div class="form-group">
        <label>受診者ID <span class="required">*</span></label>
        <input type="text" id="patientId" placeholder="P00001">
        <div class="info">受診者検索で取得したIDを入力</div>
      </div>
      <div class="form-group">
        <label>受診日 <span class="required">*</span></label>
        <input type="date" id="visitDate" value="${new Date().toISOString().split('T')[0]}">
      </div>
      <div class="form-group">
        <label>検診種別 <span class="required">*</span></label>
        <select id="examTypeId" onchange="onExamTypeChange()">
          ${examTypeOptions}
        </select>
      </div>
      <div class="form-group" id="courseGroup" style="display:none;">
        <label>コース</label>
        <select id="courseId">
          ${courseOptions}
        </select>
      </div>

      <button class="btn btn-primary" onclick="register()">登録</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>

      <div id="status"></div>

      <script>
        function onExamTypeChange() {
          const select = document.getElementById('examTypeId');
          const option = select.options[select.selectedIndex];
          const courseRequired = option.dataset.courseRequired === 'true';

          document.getElementById('courseGroup').style.display = courseRequired ? 'block' : 'none';
        }

        function register() {
          const patientId = document.getElementById('patientId').value;
          const visitDate = document.getElementById('visitDate').value;
          const examTypeId = document.getElementById('examTypeId').value;

          if (!patientId || !visitDate || !examTypeId) {
            showStatus('必須項目を入力してください', 'error');
            return;
          }

          const data = {
            patientId: patientId,
            visitDate: visitDate,
            examTypeId: examTypeId,
            courseId: document.getElementById('courseId').value || ''
          };

          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onError)
            .createVisitRecord(data);
        }

        function showStatus(message, type) {
          const status = document.getElementById('status');
          status.textContent = message;
          status.className = type;
          status.style.display = 'block';
        }

        function onSuccess(visitId) {
          showStatus('受診記録を登録しました: ' + visitId, 'success');
          setTimeout(() => google.script.host.close(), 3000);
        }

        function onError(error) {
          showStatus('エラー: ' + error.message, 'error');
        }
      </script>
    </body>
    </html>
  `;
}

// ============================================
// 判定再計算ダイアログ
// ============================================

/**
 * 判定再計算ダイアログを表示
 */
function recalculateJudgmentMenu() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt('判定再計算',
    '受診IDを入力してください。\n' +
    '例: 20251214-001',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const visitId = response.getResponseText().trim();
  if (!visitId) {
    ui.alert('エラー', '受診IDを入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 受診記録を取得して性別を確認
  const visit = getVisitRecordById(visitId);
  if (!visit) {
    ui.alert('エラー', '受診記録が見つかりません: ' + visitId, ui.ButtonSet.OK);
    return;
  }

  const patient = getPatientById(visit.patientId);
  if (!patient) {
    ui.alert('エラー', '受診者情報が見つかりません', ui.ButtonSet.OK);
    return;
  }

  // 判定を再計算
  const result = recalculateAllJudgments(visitId, patient.gender);

  ui.alert('判定再計算完了',
    `受診ID: ${visitId}\n` +
    `更新件数: ${result.updated}件\n` +
    `総合判定: ${result.overall || 'なし'}`,
    ui.ButtonSet.OK);
}

// ============================================
// トリガー設定
// ============================================

/**
 * トリガーを設定
 */
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onOpen') {
      continue; // onOpenは残す
    }
    ScriptApp.deleteTrigger(trigger);
  }

  logInfo('トリガー設定完了');
}

// ============================================
// テスト関数
// ============================================

/**
 * CRUD操作のテスト
 */
function testCRUD() {
  logInfo('===== CRUD テスト =====');

  // 受診者作成テスト
  const patientId = createPatient({
    name: 'テスト 太郎',
    kana: 'テスト タロウ',
    birthdate: new Date(1980, 0, 15),
    gender: 'M',
    company: 'テスト株式会社'
  });
  logInfo('作成された受診者ID: ' + patientId);

  // 受診者取得テスト
  const patient = getPatientById(patientId);
  logInfo('取得した受診者: ' + JSON.stringify(patient));

  // 受診記録作成テスト
  const visitId = createVisitRecord({
    patientId: patientId,
    visitDate: new Date(),
    examTypeId: 'DOCK',
    courseId: 'DOCK_LIFE'
  });
  logInfo('作成された受診ID: ' + visitId);

  // 検査結果入力テスト
  const result = inputTestResultWithJudgment(visitId, 'BMI', 24.5, 'M');
  logInfo('検査結果: ' + JSON.stringify(result));

  logInfo('===== CRUD テスト完了 =====');
}
