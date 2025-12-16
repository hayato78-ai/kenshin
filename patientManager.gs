/**
 * 受診者管理モジュール
 *
 * 機能:
 * - 受診者検索・一覧表示（SCR-003/004）
 * - 新規受診者登録（SCR-007）
 * - 企業別フィルタ・ソート機能
 *
 * 画面仕様:
 * - SCR-003: 受診者検索画面
 * - SCR-004: 受診者一覧画面
 * - SCR-007: 新規受診登録画面
 */

// ============================================
// 定数定義
// ============================================

const PATIENT_MANAGER_CONFIG = {
  // 検索結果の最大件数
  MAX_SEARCH_RESULTS: 500,

  // ステータス定義
  STATUS: {
    INPUT: '入力中',
    COMPLETE: '完了',
    PENDING: '保留'
  },

  // ID採番プレフィックス
  ID_PREFIX: 'P',

  // 必須フィールド（新規登録時）
  REQUIRED_FIELDS: ['name', 'examDate', 'courseId'],

  // 検索可能フィールド
  SEARCHABLE_FIELDS: ['name', 'nameKana', 'companyName', 'patientId']
};

// ============================================
// 受診者検索機能
// ============================================

/**
 * 受診者を検索
 * @param {Object} criteria - 検索条件
 * @returns {Array} 検索結果
 */
function searchPatients(criteria) {
  logInfo('受診者検索開始: ' + JSON.stringify(criteria));

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const data = patientSheet.getDataRange().getValues();
    const headers = data[0];
    const results = [];

    // ヘッダーのインデックスを取得
    const colIndex = {
      patientId: 0,
      status: 1,
      examDate: 2,
      name: 3,
      nameKana: 4,
      gender: 5,
      birthDate: 6,
      age: 7,
      course: 8,
      company: 9,
      department: 10,
      overallJudgment: 11
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 空行スキップ
      if (!row[colIndex.patientId]) continue;

      // フィルタリング
      if (!matchesCriteria(row, colIndex, criteria)) continue;

      results.push({
        patientId: row[colIndex.patientId],
        status: row[colIndex.status],
        examDate: row[colIndex.examDate] ? formatDate(row[colIndex.examDate]) : '',
        name: row[colIndex.name],
        nameKana: row[colIndex.nameKana],
        gender: row[colIndex.gender],
        birthDate: row[colIndex.birthDate] ? formatDate(row[colIndex.birthDate]) : '',
        age: row[colIndex.age],
        course: row[colIndex.course],
        company: row[colIndex.company],
        department: row[colIndex.department],
        overallJudgment: row[colIndex.overallJudgment],
        rowIndex: i + 1
      });

      // 最大件数制限
      if (results.length >= PATIENT_MANAGER_CONFIG.MAX_SEARCH_RESULTS) {
        break;
      }
    }

    // ソート（受診日降順）
    results.sort((a, b) => {
      if (!a.examDate) return 1;
      if (!b.examDate) return -1;
      return new Date(b.examDate) - new Date(a.examDate);
    });

    logInfo(`受診者検索完了: ${results.length}件`);
    return results;

  } catch (e) {
    logError('searchPatients', e);
    throw e;
  }
}

/**
 * 検索条件にマッチするかチェック
 * @param {Array} row - データ行
 * @param {Object} colIndex - カラムインデックス
 * @param {Object} criteria - 検索条件
 * @returns {boolean}
 */
function matchesCriteria(row, colIndex, criteria) {
  // 名前検索（部分一致）
  if (criteria.name) {
    const searchName = criteria.name.toLowerCase();
    const name = String(row[colIndex.name] || '').toLowerCase();
    const nameKana = String(row[colIndex.nameKana] || '').toLowerCase();
    if (!name.includes(searchName) && !nameKana.includes(searchName)) {
      return false;
    }
  }

  // 企業フィルタ
  if (criteria.companyId || criteria.companyName) {
    const company = String(row[colIndex.company] || '');
    if (criteria.companyId && company !== criteria.companyId) {
      // 企業名でも一致をチェック
      if (criteria.companyName && company !== criteria.companyName) {
        return false;
      }
    }
    if (criteria.companyName && !company.includes(criteria.companyName)) {
      return false;
    }
  }

  // 日付範囲フィルタ
  if (criteria.dateFrom || criteria.dateTo) {
    const examDate = row[colIndex.examDate];
    if (examDate) {
      const examDateObj = new Date(examDate);
      if (criteria.dateFrom && examDateObj < new Date(criteria.dateFrom)) {
        return false;
      }
      if (criteria.dateTo && examDateObj > new Date(criteria.dateTo)) {
        return false;
      }
    } else if (criteria.dateFrom || criteria.dateTo) {
      // 日付が空で日付範囲指定がある場合は除外
      return false;
    }
  }

  // ステータスフィルタ
  if (criteria.status && criteria.status !== 'all') {
    if (row[colIndex.status] !== criteria.status) {
      return false;
    }
  }

  // 受診者IDフィルタ（完全一致）
  if (criteria.patientId) {
    if (row[colIndex.patientId] !== criteria.patientId) {
      return false;
    }
  }

  // コースフィルタ
  if (criteria.courseId) {
    if (row[colIndex.course] !== criteria.courseId) {
      return false;
    }
  }

  return true;
}

/**
 * 受診者詳細を取得
 * @param {string} patientId - 受診者ID
 * @returns {Object} 受診者詳細データ
 */
function getPatientDetail(patientId) {
  logInfo('受診者詳細取得: ' + patientId);

  try {
    // 基本情報を取得
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const data = patientSheet.getDataRange().getValues();
    let patientRow = null;
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === patientId) {
        patientRow = data[i];
        rowIndex = i + 1;
        break;
      }
    }

    if (!patientRow) {
      return null;
    }

    const patient = {
      patientId: patientRow[0],
      status: patientRow[1],
      examDate: patientRow[2] ? formatDate(patientRow[2]) : '',
      name: patientRow[3],
      nameKana: patientRow[4],
      gender: patientRow[5],
      birthDate: patientRow[6] ? formatDate(patientRow[6]) : '',
      age: patientRow[7],
      course: patientRow[8],
      company: patientRow[9],
      department: patientRow[10],
      overallJudgment: patientRow[11],
      rowIndex: rowIndex,
      physical: {},
      blood: {}
    };

    // 身体測定データを取得
    const physicalSheet = getSheet(CONFIG.SHEETS.PHYSICAL);
    if (physicalSheet) {
      const physicalData = physicalSheet.getDataRange().getValues();
      for (let i = 1; i < physicalData.length; i++) {
        if (physicalData[i][0] === patientId) {
          patient.physical = {
            height: physicalData[i][1],
            weight: physicalData[i][2],
            standardWeight: physicalData[i][3],
            BMI: physicalData[i][4],
            bodyFat: physicalData[i][5],
            waist: physicalData[i][6],
            bpSys1: physicalData[i][7],
            bpDia1: physicalData[i][8],
            bpSys2: physicalData[i][9],
            bpDia2: physicalData[i][10]
          };
          break;
        }
      }
    }

    // 血液検査データを取得
    const bloodSheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
    if (bloodSheet) {
      const bloodData = bloodSheet.getDataRange().getValues();
      const bloodHeaders = bloodData[0];
      for (let i = 1; i < bloodData.length; i++) {
        if (bloodData[i][0] === patientId) {
          for (let j = 1; j < bloodHeaders.length; j++) {
            patient.blood[bloodHeaders[j]] = bloodData[i][j];
          }
          break;
        }
      }
    }

    return patient;

  } catch (e) {
    logError('getPatientDetail', e);
    throw e;
  }
}

// ============================================
// 新規受診者登録機能
// ============================================

/**
 * 新規受診者を登録
 * @param {Object} data - 登録データ
 * @returns {Object} 登録結果 {success, patientId, error}
 */
function registerNewPatient(data) {
  logInfo('新規受診者登録開始');

  try {
    // 必須フィールドチェック
    if (!data.name || !data.name.trim()) {
      throw new Error('氏名は必須です');
    }
    if (!data.examDate) {
      throw new Error('受診日は必須です');
    }

    // 受診者IDを生成
    const patientId = generatePatientId();

    // 受診者マスタに登録
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    // 年齢計算
    let age = '';
    if (data.birthDate) {
      age = calculateAge(new Date(data.birthDate), new Date(data.examDate));
    }

    // 新規行を追加
    const newRow = [
      patientId,                                    // 受診ID
      PATIENT_MANAGER_CONFIG.STATUS.INPUT,          // ステータス
      new Date(data.examDate),                      // 受診日
      data.name,                                    // 氏名
      data.nameKana || '',                          // カナ
      data.gender || '',                            // 性別
      data.birthDate ? new Date(data.birthDate) : '', // 生年月日
      age,                                          // 年齢
      data.courseName || data.courseId || '',       // 受診コース
      data.companyName || data.companyId || '',     // 事業所名
      data.department || '',                        // 所属
      '',                                           // 総合判定
      '',                                           // CSV取込日時
      new Date(),                                   // 最終更新日時
      ''                                            // 出力日時
    ];

    patientSheet.appendRow(newRow);

    // 身体測定シートに空行を追加（データ構造を準備）
    const physicalSheet = getSheet(CONFIG.SHEETS.PHYSICAL);
    if (physicalSheet) {
      const physicalRow = [patientId];
      // 残りの列は空
      for (let i = 1; i < 23; i++) {
        physicalRow.push('');
      }
      physicalSheet.appendRow(physicalRow);
    }

    // 血液検査シートに空行を追加
    const bloodSheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
    if (bloodSheet) {
      const bloodRow = [patientId];
      // 残りの列は空
      for (let i = 1; i < 28; i++) {
        bloodRow.push('');
      }
      bloodSheet.appendRow(bloodRow);
    }

    logInfo(`新規受診者登録完了: ${patientId}`);

    return {
      success: true,
      patientId: patientId,
      message: `受診者を登録しました（ID: ${patientId}）`
    };

  } catch (e) {
    logError('registerNewPatient', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 受診者IDを生成
 * @returns {string} 新しい受診者ID（P00001形式）
 */
function generatePatientId() {
  const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
  if (!patientSheet) {
    return 'P00001';
  }

  const data = patientSheet.getDataRange().getValues();
  let maxNum = 0;

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '');
    if (id.startsWith('P')) {
      const num = parseInt(id.substring(1), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  const newNum = maxNum + 1;
  return 'P' + String(newNum).padStart(5, '0');
}

/**
 * 年齢を計算
 * @param {Date} birthDate - 生年月日
 * @param {Date} targetDate - 基準日
 * @returns {number} 年齢
 */
function calculateAge(birthDate, targetDate) {
  let age = targetDate.getFullYear() - birthDate.getFullYear();
  const monthDiff = targetDate.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && targetDate.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
}

/**
 * 受診者情報を更新
 * @param {string} patientId - 受診者ID
 * @param {Object} data - 更新データ
 * @returns {Object} 更新結果
 */
function updatePatient(patientId, data) {
  logInfo('受診者情報更新: ' + patientId);

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const allData = patientSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === patientId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('受診者が見つかりません: ' + patientId);
    }

    // 更新する列
    if (data.name !== undefined) patientSheet.getRange(rowIndex, 4).setValue(data.name);
    if (data.nameKana !== undefined) patientSheet.getRange(rowIndex, 5).setValue(data.nameKana);
    if (data.gender !== undefined) patientSheet.getRange(rowIndex, 6).setValue(data.gender);
    if (data.birthDate !== undefined) patientSheet.getRange(rowIndex, 7).setValue(data.birthDate ? new Date(data.birthDate) : '');
    if (data.course !== undefined) patientSheet.getRange(rowIndex, 9).setValue(data.course);
    if (data.company !== undefined) patientSheet.getRange(rowIndex, 10).setValue(data.company);
    if (data.department !== undefined) patientSheet.getRange(rowIndex, 11).setValue(data.department);
    if (data.status !== undefined) patientSheet.getRange(rowIndex, 2).setValue(data.status);

    // 最終更新日時を更新
    patientSheet.getRange(rowIndex, 14).setValue(new Date());

    // 年齢を再計算
    if (data.birthDate !== undefined) {
      const examDate = patientSheet.getRange(rowIndex, 3).getValue();
      if (examDate && data.birthDate) {
        const age = calculateAge(new Date(data.birthDate), new Date(examDate));
        patientSheet.getRange(rowIndex, 8).setValue(age);
      }
    }

    logInfo('受診者情報更新完了');

    return {
      success: true,
      message: '受診者情報を更新しました'
    };

  } catch (e) {
    logError('updatePatient', e);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================
// UI関連関数
// ============================================

/**
 * 受診者検索ダイアログを表示
 */
function showPatientSearchDialog() {
  const html = HtmlService.createHtmlOutput(getPatientSearchDialogHtml())
    .setWidth(900)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, '受診者検索');
}

/**
 * 受診者検索ダイアログのHTML
 * @returns {string} HTML文字列
 */
function getPatientSearchDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 15px;
      margin: 0;
      line-height: 1.5;
    }
    h3 {
      margin: 0 0 15px 0;
      color: #1a73e8;
      border-bottom: 2px solid #1a73e8;
      padding-bottom: 8px;
    }
    .search-panel {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      margin-bottom: 15px;
    }
    .search-row {
      display: flex;
      gap: 15px;
      flex-wrap: wrap;
      margin-bottom: 10px;
    }
    .search-field {
      flex: 1;
      min-width: 150px;
    }
    .search-field label {
      display: block;
      margin-bottom: 4px;
      font-weight: 500;
      font-size: 12px;
      color: #555;
    }
    .search-field input, .search-field select {
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
    }
    .btn-row {
      display: flex;
      gap: 10px;
      justify-content: flex-end;
      margin-top: 10px;
    }
    .btn {
      padding: 8px 20px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 13px;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    .btn-primary:hover {
      background: #1557b0;
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-success {
      background: #0f9d58;
      color: white;
    }
    .btn-success:hover {
      background: #0b8043;
    }
    .results-panel {
      border: 1px solid #ddd;
      border-radius: 8px;
      overflow: hidden;
    }
    .results-header {
      background: #f1f3f4;
      padding: 10px 15px;
      font-weight: bold;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .results-count {
      font-size: 12px;
      color: #666;
      font-weight: normal;
    }
    .results-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }
    .results-table th {
      background: #f8f9fa;
      padding: 10px 8px;
      text-align: left;
      border-bottom: 2px solid #ddd;
      position: sticky;
      top: 0;
      font-weight: 600;
    }
    .results-table td {
      padding: 10px 8px;
      border-bottom: 1px solid #eee;
    }
    .results-table tr:hover {
      background: #f5f8ff;
    }
    .results-table tr.selected {
      background: #e8f0fe;
    }
    .results-body {
      max-height: 350px;
      overflow-y: auto;
    }
    .status-badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 10px;
      font-size: 11px;
    }
    .status-complete { background: #d4edda; color: #155724; }
    .status-input { background: #fff3cd; color: #856404; }
    .status-pending { background: #f8d7da; color: #721c24; }
    .judgment-badge {
      display: inline-block;
      width: 24px;
      height: 24px;
      line-height: 24px;
      text-align: center;
      border-radius: 50%;
      font-weight: bold;
      font-size: 12px;
    }
    .judgment-A { background: #e8f5e9; color: #2e7d32; }
    .judgment-B { background: #fff8e1; color: #f9a825; }
    .judgment-C { background: #fff3e0; color: #ef6c00; }
    .judgment-D { background: #ffebee; color: #c62828; }
    .action-btn {
      padding: 4px 10px;
      font-size: 11px;
      border: 1px solid #1a73e8;
      background: white;
      color: #1a73e8;
      border-radius: 4px;
      cursor: pointer;
    }
    .action-btn:hover {
      background: #e8f0fe;
    }
    .loading {
      text-align: center;
      padding: 40px;
      color: #666;
    }
    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #1a73e8;
      border-radius: 50%;
      width: 30px;
      height: 30px;
      animation: spin 1s linear infinite;
      margin: 0 auto 10px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    .error { color: #d93025; margin-top: 10px; }
    .no-results {
      text-align: center;
      padding: 40px;
      color: #666;
    }
  </style>
</head>
<body>
  <h3>🔍 受診者検索</h3>

  <div class="search-panel">
    <div class="search-row">
      <div class="search-field">
        <label>氏名・カナ</label>
        <input type="text" id="searchName" placeholder="部分一致検索">
      </div>
      <div class="search-field">
        <label>企業</label>
        <select id="searchCompany">
          <option value="">すべて</option>
        </select>
      </div>
      <div class="search-field">
        <label>ステータス</label>
        <select id="searchStatus">
          <option value="all">すべて</option>
          <option value="入力中">入力中</option>
          <option value="完了">完了</option>
          <option value="保留">保留</option>
        </select>
      </div>
    </div>
    <div class="search-row">
      <div class="search-field">
        <label>受診日（From）</label>
        <input type="date" id="searchDateFrom">
      </div>
      <div class="search-field">
        <label>受診日（To）</label>
        <input type="date" id="searchDateTo">
      </div>
      <div class="search-field">
        <label>受診者ID</label>
        <input type="text" id="searchPatientId" placeholder="P00001">
      </div>
    </div>
    <div class="btn-row">
      <button class="btn btn-secondary" onclick="clearSearch()">クリア</button>
      <button class="btn btn-primary" onclick="executeSearch()">検索</button>
    </div>
  </div>

  <div class="results-panel">
    <div class="results-header">
      <span>検索結果</span>
      <span class="results-count" id="resultsCount">0件</span>
    </div>
    <div class="results-body" id="resultsBody">
      <div class="no-results">検索条件を入力して「検索」ボタンをクリックしてください</div>
    </div>
  </div>

  <div style="margin-top: 15px; text-align: right;">
    <button class="btn btn-secondary" onclick="google.script.host.close()">閉じる</button>
    <button class="btn btn-success" onclick="openNewRegistration()">＋ 新規登録</button>
  </div>

  <script>
    let searchResults = [];

    // 企業リストを読み込み
    google.script.run
      .withSuccessHandler((companies) => {
        const select = document.getElementById('searchCompany');
        companies.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.name;
          opt.textContent = c.name;
          select.appendChild(opt);
        });
      })
      .getCompanyListForDropdown();

    function executeSearch() {
      const criteria = {
        name: document.getElementById('searchName').value,
        companyName: document.getElementById('searchCompany').value,
        status: document.getElementById('searchStatus').value,
        dateFrom: document.getElementById('searchDateFrom').value,
        dateTo: document.getElementById('searchDateTo').value,
        patientId: document.getElementById('searchPatientId').value
      };

      document.getElementById('resultsBody').innerHTML =
        '<div class="loading"><div class="spinner"></div>検索中...</div>';

      google.script.run
        .withSuccessHandler(renderResults)
        .withFailureHandler((e) => {
          document.getElementById('resultsBody').innerHTML =
            '<div class="error">エラー: ' + e.message + '</div>';
        })
        .searchPatients(criteria);
    }

    function renderResults(results) {
      searchResults = results;
      document.getElementById('resultsCount').textContent = results.length + '件';

      if (results.length === 0) {
        document.getElementById('resultsBody').innerHTML =
          '<div class="no-results">該当する受診者が見つかりません</div>';
        return;
      }

      let html = '<table class="results-table"><thead><tr>' +
        '<th>受診ID</th><th>氏名</th><th>企業</th><th>受診日</th>' +
        '<th>コース</th><th>判定</th><th>ステータス</th><th>操作</th>' +
        '</tr></thead><tbody>';

      results.forEach((p, idx) => {
        const statusClass = p.status === '完了' ? 'status-complete' :
                           p.status === '保留' ? 'status-pending' : 'status-input';
        const judgmentClass = p.overallJudgment ? 'judgment-' + p.overallJudgment : '';

        html += '<tr onclick="selectRow(' + idx + ')" data-idx="' + idx + '">' +
          '<td>' + p.patientId + '</td>' +
          '<td><strong>' + p.name + '</strong><br><small style="color:#666">' + (p.nameKana || '') + '</small></td>' +
          '<td>' + (p.company || '-') + '</td>' +
          '<td>' + (p.examDate || '-') + '</td>' +
          '<td>' + (p.course || '-') + '</td>' +
          '<td>' + (p.overallJudgment ? '<span class="judgment-badge ' + judgmentClass + '">' + p.overallJudgment + '</span>' : '-') + '</td>' +
          '<td><span class="status-badge ' + statusClass + '">' + p.status + '</span></td>' +
          '<td><button class="action-btn" onclick="viewDetail(\\'' + p.patientId + '\\')">詳細</button></td>' +
          '</tr>';
      });

      html += '</tbody></table>';
      document.getElementById('resultsBody').innerHTML = html;
    }

    function selectRow(idx) {
      document.querySelectorAll('.results-table tr').forEach(tr => tr.classList.remove('selected'));
      const row = document.querySelector('tr[data-idx="' + idx + '"]');
      if (row) row.classList.add('selected');
    }

    function viewDetail(patientId) {
      // 検索ダイアログを閉じて詳細ダイアログを開く
      google.script.host.close();
      google.script.run.showPatientDetailDialog(patientId);
    }

    function clearSearch() {
      document.getElementById('searchName').value = '';
      document.getElementById('searchCompany').value = '';
      document.getElementById('searchStatus').value = 'all';
      document.getElementById('searchDateFrom').value = '';
      document.getElementById('searchDateTo').value = '';
      document.getElementById('searchPatientId').value = '';
    }

    function openNewRegistration() {
      google.script.host.close();
      google.script.run.showNewPatientDialog();
    }

    // Enterキーで検索実行
    document.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') executeSearch();
    });
  </script>
</body>
</html>
`;
}

/**
 * 新規受診者登録ダイアログを表示
 */
function showNewPatientDialog() {
  const html = HtmlService.createHtmlOutput(getNewPatientDialogHtml())
    .setWidth(600)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, '新規受診者登録');
}

/**
 * 新規受診者登録ダイアログのHTML
 * @returns {string} HTML文字列
 */
function getNewPatientDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 20px;
      margin: 0;
      line-height: 1.6;
    }
    h3 {
      margin: 0 0 20px 0;
      color: #0f9d58;
      border-bottom: 2px solid #0f9d58;
      padding-bottom: 8px;
    }
    .form-section {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 20px;
    }
    .section-title {
      font-weight: bold;
      margin-bottom: 15px;
      color: #333;
      font-size: 14px;
    }
    .form-row {
      display: flex;
      gap: 15px;
      margin-bottom: 15px;
    }
    .form-group {
      flex: 1;
    }
    .form-group.wide {
      flex: 2;
    }
    .form-group label {
      display: block;
      margin-bottom: 5px;
      font-weight: 500;
    }
    .form-group label .required {
      color: #d93025;
      margin-left: 3px;
    }
    .form-group input, .form-group select {
      width: 100%;
      padding: 10px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
    }
    .form-group input:focus, .form-group select:focus {
      outline: none;
      border-color: #0f9d58;
      box-shadow: 0 0 0 2px rgba(15, 157, 88, 0.1);
    }
    .form-group input.error {
      border-color: #d93025;
    }
    .hint {
      font-size: 11px;
      color: #666;
      margin-top: 4px;
    }
    .btn-container {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 20px;
      padding-top: 20px;
      border-top: 1px solid #eee;
    }
    .btn {
      padding: 12px 30px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
    }
    .btn-primary {
      background: #0f9d58;
      color: white;
    }
    .btn-primary:hover {
      background: #0b8043;
    }
    .btn-primary:disabled {
      background: #ccc;
      cursor: not-allowed;
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-link {
      background: none;
      color: #1a73e8;
      text-decoration: underline;
      padding: 5px;
    }
    .message {
      padding: 15px;
      border-radius: 4px;
      margin-bottom: 15px;
      display: none;
    }
    .message.error {
      background: #fce8e6;
      color: #c5221f;
      display: block;
    }
    .message.success {
      background: #e6f4ea;
      color: #137333;
      display: block;
    }
    .loading {
      display: none;
      text-align: center;
      padding: 30px;
    }
    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #0f9d58;
      border-radius: 50%;
      width: 30px;
      height: 30px;
      animation: spin 1s linear infinite;
      margin: 0 auto 10px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <h3>＋ 新規受診者登録</h3>

  <div class="message" id="messageBox"></div>

  <div id="formContainer">
    <div class="form-section">
      <div class="section-title">基本情報</div>

      <div class="form-row">
        <div class="form-group wide">
          <label>氏名<span class="required">*</span></label>
          <input type="text" id="name" placeholder="山田 太郎">
        </div>
        <div class="form-group">
          <label>性別</label>
          <select id="gender">
            <option value="">選択してください</option>
            <option value="男">男</option>
            <option value="女">女</option>
          </select>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group wide">
          <label>フリガナ</label>
          <input type="text" id="nameKana" placeholder="ヤマダ タロウ">
        </div>
        <div class="form-group">
          <label>生年月日</label>
          <input type="date" id="birthDate">
        </div>
      </div>
    </div>

    <div class="form-section">
      <div class="section-title">受診情報</div>

      <div class="form-row">
        <div class="form-group">
          <label>受診日<span class="required">*</span></label>
          <input type="date" id="examDate">
        </div>
        <div class="form-group">
          <label>受診コース</label>
          <select id="courseId">
            <option value="">選択してください</option>
          </select>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>企業・事業所</label>
          <select id="companyId">
            <option value="">選択してください</option>
          </select>
        </div>
        <div class="form-group">
          <label>所属・部署</label>
          <input type="text" id="department" placeholder="営業部">
        </div>
      </div>

      <div class="hint">※ 検査結果は登録後に入力できます</div>
    </div>
  </div>

  <div class="loading" id="loading">
    <div class="spinner"></div>
    <div>登録中...</div>
  </div>

  <div class="btn-container">
    <button class="btn btn-link" onclick="openSearch()">← 検索に戻る</button>
    <div>
      <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
      <button class="btn btn-primary" id="submitBtn" onclick="submitForm()">登録する</button>
    </div>
  </div>

  <script>
    // 本日の日付をデフォルトに設定
    document.getElementById('examDate').valueAsDate = new Date();

    // 企業リストを読み込み
    google.script.run
      .withSuccessHandler((companies) => {
        const select = document.getElementById('companyId');
        companies.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.id;
          opt.textContent = c.name;
          opt.dataset.name = c.name;
          select.appendChild(opt);
        });
      })
      .getCompanyListForDropdown();

    // コースリストを読み込み
    google.script.run
      .withSuccessHandler((courses) => {
        const select = document.getElementById('courseId');
        courses.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.id;
          opt.textContent = c.name;
          opt.dataset.name = c.name;
          select.appendChild(opt);
        });
      })
      .getCourseListForDropdown();

    function submitForm() {
      // バリデーション
      const name = document.getElementById('name').value.trim();
      const examDate = document.getElementById('examDate').value;

      if (!name) {
        showMessage('error', '氏名を入力してください');
        document.getElementById('name').classList.add('error');
        return;
      }
      if (!examDate) {
        showMessage('error', '受診日を入力してください');
        document.getElementById('examDate').classList.add('error');
        return;
      }

      // データ収集
      const companySelect = document.getElementById('companyId');
      const courseSelect = document.getElementById('courseId');

      const data = {
        name: name,
        nameKana: document.getElementById('nameKana').value.trim(),
        gender: document.getElementById('gender').value,
        birthDate: document.getElementById('birthDate').value,
        examDate: examDate,
        courseId: courseSelect.value,
        courseName: courseSelect.selectedOptions[0]?.dataset?.name || courseSelect.value,
        companyId: companySelect.value,
        companyName: companySelect.selectedOptions[0]?.dataset?.name || companySelect.value,
        department: document.getElementById('department').value.trim()
      };

      showLoading(true);
      hideMessage();

      google.script.run
        .withSuccessHandler(handleResult)
        .withFailureHandler(handleError)
        .registerNewPatient(data);
    }

    function handleResult(result) {
      showLoading(false);

      if (result.success) {
        showMessage('success', result.message);
        document.getElementById('submitBtn').disabled = true;

        // 3秒後に閉じる
        setTimeout(() => {
          google.script.host.close();
        }, 2000);
      } else {
        showMessage('error', result.error || '登録に失敗しました');
      }
    }

    function handleError(error) {
      showLoading(false);
      showMessage('error', 'エラー: ' + error.message);
    }

    function showLoading(show) {
      document.getElementById('loading').style.display = show ? 'block' : 'none';
      document.getElementById('formContainer').style.display = show ? 'none' : 'block';
    }

    function showMessage(type, text) {
      const box = document.getElementById('messageBox');
      box.className = 'message ' + type;
      box.textContent = text;
    }

    function hideMessage() {
      document.getElementById('messageBox').className = 'message';
      document.querySelectorAll('input.error').forEach(el => el.classList.remove('error'));
    }

    function openSearch() {
      google.script.host.close();
      google.script.run.showPatientSearchDialog();
    }
  </script>
</body>
</html>
`;
}

// ============================================
// 受診者詳細表示・編集ダイアログ
// ============================================

/**
 * 受診者詳細ダイアログを表示
 * @param {string} patientId - 受診者ID
 */
function showPatientDetailDialog(patientId) {
  const html = HtmlService.createHtmlOutput(getPatientDetailDialogHtml(patientId))
    .setWidth(750)
    .setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(html, '受診者詳細 - ' + patientId);
}

/**
 * 受診者詳細ダイアログのHTML
 * @param {string} patientId - 受診者ID
 * @returns {string} HTML文字列
 */
function getPatientDetailDialogHtml(patientId) {
  // 受診者データを取得
  const patient = getPatientDetail(patientId);

  if (!patient) {
    return `
<!DOCTYPE html>
<html>
<head><base target="_top"></head>
<body style="font-family: 'Hiragino Sans', 'Meiryo', sans-serif; padding: 20px;">
  <h3>エラー</h3>
  <p>受診者が見つかりません: ${patientId}</p>
  <button onclick="google.script.host.close()">閉じる</button>
</body>
</html>`;
  }

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 15px;
      margin: 0;
      line-height: 1.6;
    }
    h3 {
      margin: 0 0 15px 0;
      color: #1a73e8;
      border-bottom: 2px solid #1a73e8;
      padding-bottom: 8px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .header-actions {
      display: flex;
      gap: 8px;
    }
    .tabs {
      display: flex;
      border-bottom: 2px solid #e0e0e0;
      margin-bottom: 15px;
    }
    .tab {
      padding: 10px 20px;
      cursor: pointer;
      border: none;
      background: transparent;
      font-size: 13px;
      color: #666;
      border-bottom: 2px solid transparent;
      margin-bottom: -2px;
    }
    .tab:hover {
      background: #f5f5f5;
    }
    .tab.active {
      color: #1a73e8;
      border-bottom-color: #1a73e8;
      font-weight: 600;
    }
    .tab-content {
      display: none;
    }
    .tab-content.active {
      display: block;
    }
    .form-section {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      margin-bottom: 15px;
    }
    .section-title {
      font-weight: bold;
      margin-bottom: 12px;
      color: #333;
      font-size: 14px;
      border-bottom: 1px solid #ddd;
      padding-bottom: 5px;
    }
    .form-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 15px;
    }
    .form-grid.two-col {
      grid-template-columns: repeat(2, 1fr);
    }
    .form-group {
      display: flex;
      flex-direction: column;
    }
    .form-group.full-width {
      grid-column: 1 / -1;
    }
    .form-group label {
      font-size: 11px;
      color: #666;
      margin-bottom: 4px;
      font-weight: 500;
    }
    .form-group input, .form-group select, .form-group textarea {
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
      background: #fff;
    }
    .form-group input:read-only, .form-group select:disabled, .form-group textarea:read-only {
      background: #f5f5f5;
      color: #333;
    }
    .form-group input:focus, .form-group select:focus, .form-group textarea:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.1);
    }
    .btn {
      padding: 8px 16px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 13px;
    }
    .btn-sm {
      padding: 5px 10px;
      font-size: 12px;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    .btn-primary:hover {
      background: #1557b0;
    }
    .btn-success {
      background: #0f9d58;
      color: white;
    }
    .btn-success:hover {
      background: #0b8043;
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-secondary:hover {
      background: #e8eaed;
    }
    .btn-warning {
      background: #f9ab00;
      color: white;
    }
    .status-badge {
      display: inline-block;
      padding: 3px 10px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: 500;
    }
    .status-complete { background: #d4edda; color: #155724; }
    .status-input { background: #fff3cd; color: #856404; }
    .status-pending { background: #f8d7da; color: #721c24; }
    .judgment-badge {
      display: inline-block;
      width: 28px;
      height: 28px;
      line-height: 28px;
      text-align: center;
      border-radius: 50%;
      font-weight: bold;
      font-size: 14px;
    }
    .judgment-A { background: #e8f5e9; color: #2e7d32; }
    .judgment-B { background: #fff8e1; color: #f9a825; }
    .judgment-C { background: #fff3e0; color: #ef6c00; }
    .judgment-D { background: #ffebee; color: #c62828; }
    .judgment-E { background: #fce4ec; color: #c2185b; }
    .judgment-G { background: #e3f2fd; color: #1565c0; }
    .footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 20px;
      padding-top: 15px;
      border-top: 1px solid #eee;
    }
    .message {
      padding: 10px 15px;
      border-radius: 4px;
      margin-bottom: 15px;
      display: none;
    }
    .message.success { background: #e6f4ea; color: #137333; display: block; }
    .message.error { background: #fce8e6; color: #c5221f; display: block; }
    .message.info { background: #e8f0fe; color: #1a73e8; display: block; }
    .data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }
    .data-table th, .data-table td {
      padding: 8px;
      text-align: left;
      border-bottom: 1px solid #eee;
    }
    .data-table th {
      background: #f8f9fa;
      font-weight: 600;
      color: #555;
    }
    .edit-mode .form-group input:not([readonly]),
    .edit-mode .form-group select:not([disabled]),
    .edit-mode .form-group textarea:not([readonly]) {
      border-color: #1a73e8;
      background: #fff;
    }
    .spinner {
      display: inline-block;
      width: 16px;
      height: 16px;
      border: 2px solid #f3f3f3;
      border-top: 2px solid #1a73e8;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-right: 8px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <h3>
    <span>👤 受診者詳細</span>
    <div class="header-actions">
      <button class="btn btn-sm btn-warning" id="editBtn" onclick="toggleEditMode()">✏️ 編集</button>
      <button class="btn btn-sm btn-success" id="saveBtn" onclick="saveChanges()" style="display:none;">💾 保存</button>
      <button class="btn btn-sm btn-secondary" id="cancelBtn" onclick="cancelEdit()" style="display:none;">キャンセル</button>
    </div>
  </h3>

  <div class="message" id="messageBox"></div>

  <!-- タブナビゲーション -->
  <div class="tabs">
    <button class="tab active" onclick="showTab('basic')">基本情報</button>
    <button class="tab" onclick="showTab('physical')">身体測定</button>
    <button class="tab" onclick="showTab('blood')">血液検査</button>
  </div>

  <!-- 基本情報タブ -->
  <div id="tab-basic" class="tab-content active">
    <div class="form-section">
      <div class="section-title">受診者情報</div>
      <div class="form-grid">
        <div class="form-group">
          <label>受診ID</label>
          <input type="text" id="patientId" value="${patient.patientId || ''}" readonly>
        </div>
        <div class="form-group">
          <label>ステータス</label>
          <select id="status" disabled>
            <option value="入力中" ${patient.status === '入力中' ? 'selected' : ''}>入力中</option>
            <option value="完了" ${patient.status === '完了' ? 'selected' : ''}>完了</option>
            <option value="保留" ${patient.status === '保留' ? 'selected' : ''}>保留</option>
          </select>
        </div>
        <div class="form-group">
          <label>総合判定</label>
          <select id="overallJudgment" disabled>
            <option value="">-</option>
            <option value="A" ${patient.overallJudgment === 'A' ? 'selected' : ''}>A</option>
            <option value="B" ${patient.overallJudgment === 'B' ? 'selected' : ''}>B</option>
            <option value="C" ${patient.overallJudgment === 'C' ? 'selected' : ''}>C</option>
            <option value="D" ${patient.overallJudgment === 'D' ? 'selected' : ''}>D</option>
            <option value="E" ${patient.overallJudgment === 'E' ? 'selected' : ''}>E</option>
            <option value="G" ${patient.overallJudgment === 'G' ? 'selected' : ''}>G</option>
          </select>
        </div>
        <div class="form-group">
          <label>氏名</label>
          <input type="text" id="name" value="${patient.name || ''}" readonly>
        </div>
        <div class="form-group">
          <label>カナ</label>
          <input type="text" id="nameKana" value="${patient.nameKana || ''}" readonly>
        </div>
        <div class="form-group">
          <label>性別</label>
          <select id="gender" disabled>
            <option value="">-</option>
            <option value="男" ${patient.gender === '男' ? 'selected' : ''}>男</option>
            <option value="女" ${patient.gender === '女' ? 'selected' : ''}>女</option>
          </select>
        </div>
        <div class="form-group">
          <label>生年月日</label>
          <input type="date" id="birthDate" value="${patient.birthDate || ''}" disabled>
        </div>
        <div class="form-group">
          <label>年齢</label>
          <input type="text" id="age" value="${patient.age ? patient.age + '歳' : ''}" readonly>
        </div>
        <div class="form-group">
          <label>受診日</label>
          <input type="date" id="examDate" value="${patient.examDate || ''}" disabled>
        </div>
        <div class="form-group">
          <label>受診コース</label>
          <input type="text" id="course" value="${patient.course || ''}" readonly>
        </div>
        <div class="form-group">
          <label>企業・事業所</label>
          <input type="text" id="company" value="${patient.company || ''}" readonly>
        </div>
        <div class="form-group">
          <label>所属・部署</label>
          <input type="text" id="department" value="${patient.department || ''}" readonly>
        </div>
      </div>
    </div>
  </div>

  <!-- 身体測定タブ -->
  <div id="tab-physical" class="tab-content">
    <div class="form-section">
      <div class="section-title">身体測定データ</div>
      <div class="form-grid">
        <div class="form-group">
          <label>身長 (cm)</label>
          <input type="text" value="${patient.physical?.height || '-'}" readonly>
        </div>
        <div class="form-group">
          <label>体重 (kg)</label>
          <input type="text" value="${patient.physical?.weight || '-'}" readonly>
        </div>
        <div class="form-group">
          <label>BMI</label>
          <input type="text" value="${patient.physical?.BMI || '-'}" readonly>
        </div>
        <div class="form-group">
          <label>標準体重 (kg)</label>
          <input type="text" value="${patient.physical?.standardWeight || '-'}" readonly>
        </div>
        <div class="form-group">
          <label>体脂肪率 (%)</label>
          <input type="text" value="${patient.physical?.bodyFat || '-'}" readonly>
        </div>
        <div class="form-group">
          <label>腹囲 (cm)</label>
          <input type="text" value="${patient.physical?.waist || '-'}" readonly>
        </div>
        <div class="form-group">
          <label>血圧（1回目）</label>
          <input type="text" value="${patient.physical?.bpSys1 && patient.physical?.bpDia1 ? patient.physical.bpSys1 + '/' + patient.physical.bpDia1 : '-'}" readonly>
        </div>
        <div class="form-group">
          <label>血圧（2回目）</label>
          <input type="text" value="${patient.physical?.bpSys2 && patient.physical?.bpDia2 ? patient.physical.bpSys2 + '/' + patient.physical.bpDia2 : '-'}" readonly>
        </div>
      </div>
    </div>
  </div>

  <!-- 血液検査タブ -->
  <div id="tab-blood" class="tab-content">
    <div class="form-section">
      <div class="section-title">血液検査データ</div>
      ${getBloodTestTableHtml(patient.blood)}
    </div>
  </div>

  <div class="footer">
    <button class="btn btn-secondary" onclick="backToSearch()">← 検索に戻る</button>
    <div>
      <button class="btn btn-secondary" onclick="google.script.host.close()">閉じる</button>
    </div>
  </div>

  <script>
    let isEditMode = false;
    const originalData = {
      patientId: '${patient.patientId || ''}',
      status: '${patient.status || ''}',
      name: '${patient.name || ''}',
      nameKana: '${patient.nameKana || ''}',
      gender: '${patient.gender || ''}',
      birthDate: '${patient.birthDate || ''}',
      examDate: '${patient.examDate || ''}',
      course: '${patient.course || ''}',
      company: '${patient.company || ''}',
      department: '${patient.department || ''}',
      overallJudgment: '${patient.overallJudgment || ''}'
    };

    function showTab(tabName) {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

      event.target.classList.add('active');
      document.getElementById('tab-' + tabName).classList.add('active');
    }

    function toggleEditMode() {
      isEditMode = true;
      document.body.classList.add('edit-mode');

      // 編集可能にする
      ['name', 'nameKana', 'course', 'company', 'department'].forEach(id => {
        document.getElementById(id).readOnly = false;
      });
      ['status', 'gender', 'birthDate', 'examDate', 'overallJudgment'].forEach(id => {
        document.getElementById(id).disabled = false;
      });

      // ボタン表示切替
      document.getElementById('editBtn').style.display = 'none';
      document.getElementById('saveBtn').style.display = 'inline-block';
      document.getElementById('cancelBtn').style.display = 'inline-block';

      showMessage('info', '編集モードです。変更後「保存」をクリックしてください。');
    }

    function cancelEdit() {
      isEditMode = false;
      document.body.classList.remove('edit-mode');

      // 元のデータに戻す
      document.getElementById('name').value = originalData.name;
      document.getElementById('nameKana').value = originalData.nameKana;
      document.getElementById('gender').value = originalData.gender;
      document.getElementById('status').value = originalData.status;
      document.getElementById('birthDate').value = originalData.birthDate;
      document.getElementById('examDate').value = originalData.examDate;
      document.getElementById('course').value = originalData.course;
      document.getElementById('company').value = originalData.company;
      document.getElementById('department').value = originalData.department;
      document.getElementById('overallJudgment').value = originalData.overallJudgment;

      // 読み取り専用に戻す
      ['name', 'nameKana', 'course', 'company', 'department'].forEach(id => {
        document.getElementById(id).readOnly = true;
      });
      ['status', 'gender', 'birthDate', 'examDate', 'overallJudgment'].forEach(id => {
        document.getElementById(id).disabled = true;
      });

      // ボタン表示切替
      document.getElementById('editBtn').style.display = 'inline-block';
      document.getElementById('saveBtn').style.display = 'none';
      document.getElementById('cancelBtn').style.display = 'none';

      hideMessage();
    }

    function saveChanges() {
      const patientId = document.getElementById('patientId').value;

      const data = {
        name: document.getElementById('name').value,
        nameKana: document.getElementById('nameKana').value,
        gender: document.getElementById('gender').value,
        birthDate: document.getElementById('birthDate').value,
        course: document.getElementById('course').value,
        company: document.getElementById('company').value,
        department: document.getElementById('department').value,
        status: document.getElementById('status').value,
        overallJudgment: document.getElementById('overallJudgment').value
      };

      showMessage('info', '<span class="spinner"></span>保存中...');

      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            showMessage('success', '保存しました。');
            // 編集モードを解除
            isEditMode = false;
            document.body.classList.remove('edit-mode');

            // 読み取り専用に戻す
            ['name', 'nameKana', 'course', 'company', 'department'].forEach(id => {
              document.getElementById(id).readOnly = true;
            });
            ['status', 'gender', 'birthDate', 'examDate', 'overallJudgment'].forEach(id => {
              document.getElementById(id).disabled = true;
            });

            // ボタン表示切替
            document.getElementById('editBtn').style.display = 'inline-block';
            document.getElementById('saveBtn').style.display = 'none';
            document.getElementById('cancelBtn').style.display = 'none';

            // originalDataを更新
            Object.assign(originalData, data);
          } else {
            showMessage('error', 'エラー: ' + result.error);
          }
        })
        .withFailureHandler(function(error) {
          showMessage('error', 'エラー: ' + error.message);
        })
        .updatePatient(patientId, data);
    }

    function backToSearch() {
      google.script.host.close();
      google.script.run.showPatientSearchDialog();
    }

    function showMessage(type, text) {
      const box = document.getElementById('messageBox');
      box.className = 'message ' + type;
      box.innerHTML = text;
    }

    function hideMessage() {
      document.getElementById('messageBox').className = 'message';
    }
  </script>
</body>
</html>`;
}

/**
 * 血液検査データをHTML表形式で出力
 * @param {Object} bloodData - 血液検査データ
 * @returns {string} HTML文字列
 */
function getBloodTestTableHtml(bloodData) {
  if (!bloodData || Object.keys(bloodData).length === 0) {
    return '<p style="color:#666; padding:20px; text-align:center;">血液検査データがありません</p>';
  }

  let html = '<table class="data-table"><thead><tr><th>検査項目</th><th>結果値</th></tr></thead><tbody>';

  const displayOrder = [
    'WBC', 'RBC', 'Hb', 'Ht', 'MCV', 'MCH', 'MCHC', 'PLT',
    'AST', 'ALT', 'γ-GTP', 'ALP', 'LDH', 'T-Bil', 'Alb',
    'BUN', 'CRE', 'UA', 'eGFR',
    'T-CHO', 'TG', 'HDL', 'LDL', 'Non-HDL',
    'GLU', 'HbA1c',
    'CRP', 'RF'
  ];

  // 表示順に従って出力
  for (const key of displayOrder) {
    if (bloodData[key] !== undefined && bloodData[key] !== '') {
      html += '<tr><td>' + key + '</td><td>' + bloodData[key] + '</td></tr>';
    }
  }

  // 表示順にない項目も出力
  for (const [key, value] of Object.entries(bloodData)) {
    if (!displayOrder.includes(key) && value !== undefined && value !== '') {
      html += '<tr><td>' + key + '</td><td>' + value + '</td></tr>';
    }
  }

  html += '</tbody></table>';
  return html;
}

// ============================================
// ドロップダウン用データ取得関数
// ============================================

/**
 * コースリストを取得（ドロップダウン用）
 * @returns {Array} コースリスト [{id, name}]
 */
function getCourseListForDropdown() {
  const result = [];

  try {
    const courseSheet = getSheet(CONFIG.SHEETS.COURSE);
    if (!courseSheet) {
      // コースマスタがない場合はデフォルト値を返す
      return [
        { id: 'CRS001', name: '生活習慣病ドック' },
        { id: 'CRS002', name: '人間ドック標準' },
        { id: 'CRS003', name: 'がんドック' },
        { id: 'CRS004', name: '定期健診A' }
      ];
    }

    const data = courseSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 有効フラグがtrueのコースのみ
      if (row[5] !== false) {
        result.push({
          id: row[0],
          name: row[1]
        });
      }
    }
  } catch (e) {
    logError('getCourseListForDropdown', e);
  }

  return result;
}

// ============================================
// テスト関数
// ============================================

/**
 * 受診者検索のテスト
 */
function testSearchPatients() {
  const results = searchPatients({
    name: '山田',
    status: 'all'
  });
  logInfo('検索結果: ' + results.length + '件');
  if (results.length > 0) {
    logInfo('最初の結果: ' + JSON.stringify(results[0]));
  }
}

/**
 * 新規登録のテスト
 */
function testRegisterNewPatient() {
  const result = registerNewPatient({
    name: 'テスト太郎',
    nameKana: 'テストタロウ',
    gender: '男',
    birthDate: '1980-01-15',
    examDate: '2025-12-16',
    courseName: '生活習慣病ドック',
    companyName: 'テスト株式会社'
  });
  logInfo('登録結果: ' + JSON.stringify(result));
}

/**
 * 検索ダイアログのテスト表示
 */
function testShowPatientSearchDialog() {
  showPatientSearchDialog();
}

/**
 * 新規登録ダイアログのテスト表示
 */
function testShowNewPatientDialog() {
  showNewPatientDialog();
}
