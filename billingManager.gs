/**
 * 請求管理モジュール
 *
 * 機能:
 * - 企業別請求一覧表示（金額ソート・フィルタ）
 * - 請求ステータス管理
 * - 請求額自動計算（コース料金＋オプション）
 * - 請求一覧Excel出力
 */

// ============================================
// 定数定義
// ============================================

const BILLING_CONFIG = {
  // 請求ステータス
  BILLING_STATUS: {
    UNBILLED: '未請求',
    BILLED: '請求済',
    PAID: '入金済',
    CANCELLED: 'キャンセル'
  },

  // ソートオプション
  SORT_OPTIONS: {
    AMOUNT_DESC: 'amount_desc',
    AMOUNT_ASC: 'amount_asc',
    DATE_DESC: 'date_desc',
    DATE_ASC: 'date_asc',
    COMPANY_ASC: 'company_asc'
  },

  // 消費税率
  TAX_RATE: 0.10
};

// ============================================
// 請求データ取得・集計機能
// ============================================

/**
 * 企業別請求一覧を取得
 * @param {Object} criteria - 検索条件
 * @returns {Object} 請求一覧データ
 */
function getBillingList(criteria) {
  logInfo('請求一覧取得開始: ' + JSON.stringify(criteria));

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const patientData = patientSheet.getDataRange().getValues();

    // コース料金マスタを取得
    const coursePrices = getCoursePriceMap();

    // 企業別に集計
    const companyBilling = {};

    for (let i = 1; i < patientData.length; i++) {
      const row = patientData[i];
      const patientId = row[0];
      const status = row[1];
      const examDate = row[2];
      const name = row[3];
      const course = row[8];
      const company = row[9] || '個人';
      const billingStatus = row[15] || BILLING_CONFIG.BILLING_STATUS.UNBILLED;

      // 空行スキップ
      if (!patientId) continue;

      // 日付フィルタ
      if (criteria.dateFrom || criteria.dateTo) {
        if (examDate) {
          const examDateObj = new Date(examDate);
          if (criteria.dateFrom && examDateObj < new Date(criteria.dateFrom)) continue;
          if (criteria.dateTo && examDateObj > new Date(criteria.dateTo)) continue;
        }
      }

      // 請求ステータスフィルタ
      if (criteria.billingStatus && criteria.billingStatus !== 'all') {
        if (billingStatus !== criteria.billingStatus) continue;
      }

      // 企業フィルタ
      if (criteria.companyName && company !== criteria.companyName) continue;

      // 企業データを集計
      if (!companyBilling[company]) {
        companyBilling[company] = {
          companyName: company,
          patients: [],
          totalAmount: 0,
          patientCount: 0,
          billingStatusCount: {
            [BILLING_CONFIG.BILLING_STATUS.UNBILLED]: 0,
            [BILLING_CONFIG.BILLING_STATUS.BILLED]: 0,
            [BILLING_CONFIG.BILLING_STATUS.PAID]: 0,
            [BILLING_CONFIG.BILLING_STATUS.CANCELLED]: 0
          }
        };
      }

      // コース料金を取得
      const coursePrice = coursePrices[course] || 0;

      companyBilling[company].patients.push({
        patientId: patientId,
        name: name,
        examDate: examDate ? formatDate(examDate) : '',
        course: course,
        amount: coursePrice,
        billingStatus: billingStatus,
        rowIndex: i + 1
      });

      companyBilling[company].totalAmount += coursePrice;
      companyBilling[company].patientCount++;
      companyBilling[company].billingStatusCount[billingStatus]++;
    }

    // 配列に変換
    let result = Object.values(companyBilling);

    // ソート
    result = sortBillingList(result, criteria.sortBy || BILLING_CONFIG.SORT_OPTIONS.AMOUNT_DESC);

    // 合計計算
    const summary = {
      totalCompanies: result.length,
      totalPatients: result.reduce((sum, c) => sum + c.patientCount, 0),
      totalAmount: result.reduce((sum, c) => sum + c.totalAmount, 0),
      byStatus: {
        [BILLING_CONFIG.BILLING_STATUS.UNBILLED]: 0,
        [BILLING_CONFIG.BILLING_STATUS.BILLED]: 0,
        [BILLING_CONFIG.BILLING_STATUS.PAID]: 0,
        [BILLING_CONFIG.BILLING_STATUS.CANCELLED]: 0
      }
    };

    result.forEach(c => {
      Object.keys(c.billingStatusCount).forEach(status => {
        summary.byStatus[status] += c.billingStatusCount[status];
      });
    });

    logInfo(`請求一覧取得完了: ${result.length}社, ${summary.totalPatients}名, ¥${summary.totalAmount.toLocaleString()}`);

    return {
      companies: result,
      summary: summary
    };

  } catch (e) {
    logError('getBillingList', e);
    throw e;
  }
}

/**
 * 請求一覧をソート
 * @param {Array} list - 企業リスト
 * @param {string} sortBy - ソートキー
 * @returns {Array} ソート済みリスト
 */
function sortBillingList(list, sortBy) {
  switch (sortBy) {
    case BILLING_CONFIG.SORT_OPTIONS.AMOUNT_DESC:
      return list.sort((a, b) => b.totalAmount - a.totalAmount);
    case BILLING_CONFIG.SORT_OPTIONS.AMOUNT_ASC:
      return list.sort((a, b) => a.totalAmount - b.totalAmount);
    case BILLING_CONFIG.SORT_OPTIONS.DATE_DESC:
      return list.sort((a, b) => {
        const dateA = a.patients.length > 0 ? new Date(a.patients[0].examDate) : new Date(0);
        const dateB = b.patients.length > 0 ? new Date(b.patients[0].examDate) : new Date(0);
        return dateB - dateA;
      });
    case BILLING_CONFIG.SORT_OPTIONS.DATE_ASC:
      return list.sort((a, b) => {
        const dateA = a.patients.length > 0 ? new Date(a.patients[0].examDate) : new Date(0);
        const dateB = b.patients.length > 0 ? new Date(b.patients[0].examDate) : new Date(0);
        return dateA - dateB;
      });
    case BILLING_CONFIG.SORT_OPTIONS.COMPANY_ASC:
      return list.sort((a, b) => a.companyName.localeCompare(b.companyName, 'ja'));
    default:
      return list;
  }
}

/**
 * コース料金マップを取得
 * @returns {Object} コース名 → 料金 のマップ
 */
function getCoursePriceMap() {
  const priceMap = {};

  try {
    const courseSheet = getSheet(CONFIG.SHEETS.COURSE);
    if (!courseSheet) {
      // デフォルト料金
      return {
        '生活習慣病ドック': 25000,
        '人間ドック標準': 45000,
        'がんドック': 80000,
        'レディースドック': 55000,
        '定期健診A': 8000,
        '定期健診B': 12000,
        '雇入時健診': 10000,
        '労災二次健診': 0,
        '特定健康診査': 0
      };
    }

    const data = courseSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const courseName = data[i][2]; // コース名
      const price = data[i][4] || 0;  // 料金
      if (courseName) {
        priceMap[courseName] = price;
      }
    }
  } catch (e) {
    logError('getCoursePriceMap', e);
  }

  return priceMap;
}

/**
 * 企業の請求詳細を取得
 * @param {string} companyName - 企業名
 * @param {Object} criteria - フィルタ条件
 * @returns {Object} 企業の請求詳細
 */
function getCompanyBillingDetail(companyName, criteria) {
  logInfo('企業請求詳細取得: ' + companyName);

  try {
    const billingData = getBillingList({
      ...criteria,
      companyName: companyName
    });

    if (billingData.companies.length === 0) {
      return null;
    }

    return billingData.companies[0];

  } catch (e) {
    logError('getCompanyBillingDetail', e);
    throw e;
  }
}

// ============================================
// 請求ステータス更新機能
// ============================================

/**
 * 請求ステータスを更新
 * @param {string} patientId - 受診者ID
 * @param {string} newStatus - 新しいステータス
 * @returns {Object} 更新結果
 */
function updateBillingStatus(patientId, newStatus) {
  logInfo(`請求ステータス更新: ${patientId} → ${newStatus}`);

  try {
    // ステータス検証
    const validStatuses = Object.values(BILLING_CONFIG.BILLING_STATUS);
    if (!validStatuses.includes(newStatus)) {
      throw new Error('無効な請求ステータス: ' + newStatus);
    }

    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const data = patientSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === patientId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('受診者が見つかりません: ' + patientId);
    }

    // 請求ステータス列が存在するか確認（16列目 = P列）
    const headers = data[0];
    if (headers.length < 16 || headers[15] !== '請求ステータス') {
      // ヘッダーに請求ステータス列を追加
      patientSheet.getRange(1, 16).setValue('請求ステータス');
    }

    // ステータスを更新
    patientSheet.getRange(rowIndex, 16).setValue(newStatus);

    logInfo('請求ステータス更新完了');

    return {
      success: true,
      message: `請求ステータスを「${newStatus}」に更新しました`
    };

  } catch (e) {
    logError('updateBillingStatus', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 複数受診者の請求ステータスを一括更新
 * @param {Array} patientIds - 受診者IDの配列
 * @param {string} newStatus - 新しいステータス
 * @returns {Object} 更新結果
 */
function updateBillingStatusBatch(patientIds, newStatus) {
  logInfo(`請求ステータス一括更新: ${patientIds.length}件 → ${newStatus}`);

  let successCount = 0;
  let failCount = 0;
  const errors = [];

  for (const patientId of patientIds) {
    const result = updateBillingStatus(patientId, newStatus);
    if (result.success) {
      successCount++;
    } else {
      failCount++;
      errors.push(`${patientId}: ${result.error}`);
    }
  }

  return {
    success: failCount === 0,
    successCount: successCount,
    failCount: failCount,
    errors: errors,
    message: `${successCount}件更新完了${failCount > 0 ? `、${failCount}件失敗` : ''}`
  };
}

// ============================================
// 請求ステータス列の初期化
// ============================================

/**
 * 受診者マスタに請求ステータス列を追加
 */
function initializeBillingStatusColumn() {
  logInfo('請求ステータス列初期化開始');

  try {
    const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
    if (!patientSheet) {
      throw new Error('受診者マスタシートが見つかりません');
    }

    const data = patientSheet.getDataRange().getValues();
    const headers = data[0];

    // 既存の列を確認
    if (headers.length >= 16 && headers[15] === '請求ステータス') {
      logInfo('請求ステータス列は既に存在します');
      return { success: true, message: '請求ステータス列は既に存在します' };
    }

    // ヘッダーに列を追加
    patientSheet.getRange(1, 16).setValue('請求ステータス');

    // 既存データにデフォルト値を設定
    const lastRow = patientSheet.getLastRow();
    if (lastRow > 1) {
      const defaultStatus = BILLING_CONFIG.BILLING_STATUS.UNBILLED;
      const values = [];
      for (let i = 2; i <= lastRow; i++) {
        values.push([defaultStatus]);
      }
      patientSheet.getRange(2, 16, values.length, 1).setValues(values);
    }

    logInfo('請求ステータス列初期化完了');

    return {
      success: true,
      message: '請求ステータス列を追加しました'
    };

  } catch (e) {
    logError('initializeBillingStatusColumn', e);
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
 * 請求一覧ダイアログを表示
 */
function showBillingListDialog() {
  const html = HtmlService.createHtmlOutput(getBillingListDialogHtml())
    .setWidth(1000)
    .setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(html, '💰 請求額一覧');
}

/**
 * 請求一覧ダイアログのHTML
 * @returns {string} HTML文字列
 */
function getBillingListDialogHtml() {
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
      color: #ea8600;
      border-bottom: 2px solid #ea8600;
      padding-bottom: 8px;
    }
    .filter-panel {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      margin-bottom: 15px;
    }
    .filter-row {
      display: flex;
      gap: 15px;
      flex-wrap: wrap;
      align-items: flex-end;
    }
    .filter-field {
      flex: 1;
      min-width: 130px;
    }
    .filter-field label {
      display: block;
      margin-bottom: 4px;
      font-weight: 500;
      font-size: 12px;
      color: #555;
    }
    .filter-field input, .filter-field select {
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
    }
    .btn {
      padding: 8px 16px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 13px;
    }
    .btn-primary { background: #ea8600; color: white; }
    .btn-primary:hover { background: #c97200; }
    .btn-secondary { background: #f1f3f4; color: #333; }
    .btn-success { background: #0f9d58; color: white; }
    .btn-success:hover { background: #0b8043; }

    .summary-panel {
      display: flex;
      gap: 20px;
      margin-bottom: 15px;
      background: linear-gradient(135deg, #fff8e1 0%, #ffecb3 100%);
      padding: 15px 20px;
      border-radius: 8px;
      border: 1px solid #ffd54f;
    }
    .summary-item {
      text-align: center;
    }
    .summary-item .label {
      font-size: 11px;
      color: #666;
      margin-bottom: 3px;
    }
    .summary-item .value {
      font-size: 20px;
      font-weight: bold;
      color: #ea8600;
    }
    .summary-item .value.large {
      font-size: 24px;
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
      cursor: pointer;
    }
    .results-table th:hover {
      background: #e8e8e8;
    }
    .results-table th .sort-icon {
      margin-left: 5px;
      font-size: 10px;
    }
    .results-table td {
      padding: 12px 8px;
      border-bottom: 1px solid #eee;
    }
    .results-table tr:hover {
      background: #fff8e1;
    }
    .results-body {
      max-height: 400px;
      overflow-y: auto;
    }
    .amount {
      font-weight: bold;
      color: #ea8600;
      text-align: right;
    }
    .amount.large {
      font-size: 14px;
    }
    .status-badge {
      display: inline-block;
      padding: 3px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 500;
    }
    .status-unbilled { background: #fff3e0; color: #e65100; }
    .status-billed { background: #e3f2fd; color: #1565c0; }
    .status-paid { background: #e8f5e9; color: #2e7d32; }
    .status-cancelled { background: #f5f5f5; color: #757575; }

    .count-badges {
      display: flex;
      gap: 5px;
    }
    .count-badge {
      font-size: 10px;
      padding: 2px 6px;
      border-radius: 8px;
    }

    .action-btn {
      padding: 4px 10px;
      font-size: 11px;
      border: 1px solid #ea8600;
      background: white;
      color: #ea8600;
      border-radius: 4px;
      cursor: pointer;
    }
    .action-btn:hover {
      background: #fff8e1;
    }

    .loading {
      text-align: center;
      padding: 40px;
      color: #666;
    }
    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #ea8600;
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
    .no-results {
      text-align: center;
      padding: 40px;
      color: #666;
    }
    .footer-actions {
      margin-top: 15px;
      display: flex;
      justify-content: space-between;
    }
  </style>
</head>
<body>
  <h3>💰 請求額一覧</h3>

  <div class="filter-panel">
    <div class="filter-row">
      <div class="filter-field">
        <label>企業</label>
        <select id="filterCompany">
          <option value="">すべて</option>
        </select>
      </div>
      <div class="filter-field">
        <label>期間（From）</label>
        <input type="date" id="filterDateFrom">
      </div>
      <div class="filter-field">
        <label>期間（To）</label>
        <input type="date" id="filterDateTo">
      </div>
      <div class="filter-field">
        <label>請求ステータス</label>
        <select id="filterStatus">
          <option value="all">すべて</option>
          <option value="未請求">未請求</option>
          <option value="請求済">請求済</option>
          <option value="入金済">入金済</option>
          <option value="キャンセル">キャンセル</option>
        </select>
      </div>
      <div class="filter-field">
        <label>ソート</label>
        <select id="sortBy">
          <option value="amount_desc">金額（高い順）</option>
          <option value="amount_asc">金額（低い順）</option>
          <option value="date_desc">受診日（新しい順）</option>
          <option value="date_asc">受診日（古い順）</option>
          <option value="company_asc">企業名（あいうえお順）</option>
        </select>
      </div>
      <button class="btn btn-primary" onclick="loadBillingData()">表示</button>
    </div>
  </div>

  <div class="summary-panel">
    <div class="summary-item">
      <div class="label">企業数</div>
      <div class="value" id="summaryCompanies">-</div>
    </div>
    <div class="summary-item">
      <div class="label">受診者数</div>
      <div class="value" id="summaryPatients">-</div>
    </div>
    <div class="summary-item">
      <div class="label">総請求額</div>
      <div class="value large" id="summaryAmount">-</div>
    </div>
    <div class="summary-item">
      <div class="label">未請求</div>
      <div class="value" id="summaryUnbilled" style="color:#e65100">-</div>
    </div>
    <div class="summary-item">
      <div class="label">請求済</div>
      <div class="value" id="summaryBilled" style="color:#1565c0">-</div>
    </div>
    <div class="summary-item">
      <div class="label">入金済</div>
      <div class="value" id="summaryPaid" style="color:#2e7d32">-</div>
    </div>
  </div>

  <div class="results-panel">
    <div class="results-header">
      <span>企業別請求一覧</span>
      <button class="btn btn-success" onclick="exportToExcel()">📊 Excel出力</button>
    </div>
    <div class="results-body" id="resultsBody">
      <div class="loading">
        <div class="spinner"></div>
        データを読み込み中...
      </div>
    </div>
  </div>

  <div class="footer-actions">
    <button class="btn btn-secondary" onclick="google.script.host.close()">閉じる</button>
    <div>
      <button class="btn btn-secondary" onclick="initBillingColumn()">請求列初期化</button>
    </div>
  </div>

  <script>
    let billingData = null;

    // 企業リストを読み込み
    google.script.run
      .withSuccessHandler((companies) => {
        const select = document.getElementById('filterCompany');
        companies.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.name;
          opt.textContent = c.name;
          select.appendChild(opt);
        });
      })
      .getCompanyListForDropdown();

    // 初期表示
    loadBillingData();

    function loadBillingData() {
      const criteria = {
        companyName: document.getElementById('filterCompany').value,
        dateFrom: document.getElementById('filterDateFrom').value,
        dateTo: document.getElementById('filterDateTo').value,
        billingStatus: document.getElementById('filterStatus').value,
        sortBy: document.getElementById('sortBy').value
      };

      document.getElementById('resultsBody').innerHTML =
        '<div class="loading"><div class="spinner"></div>データを読み込み中...</div>';

      google.script.run
        .withSuccessHandler(renderBillingData)
        .withFailureHandler((e) => {
          document.getElementById('resultsBody').innerHTML =
            '<div class="no-results" style="color:red">エラー: ' + e.message + '</div>';
        })
        .getBillingList(criteria);
    }

    function renderBillingData(data) {
      billingData = data;

      // サマリー更新
      document.getElementById('summaryCompanies').textContent = data.summary.totalCompanies + '社';
      document.getElementById('summaryPatients').textContent = data.summary.totalPatients + '名';
      document.getElementById('summaryAmount').textContent = '¥' + data.summary.totalAmount.toLocaleString();
      document.getElementById('summaryUnbilled').textContent = data.summary.byStatus['未請求'] + '件';
      document.getElementById('summaryBilled').textContent = data.summary.byStatus['請求済'] + '件';
      document.getElementById('summaryPaid').textContent = data.summary.byStatus['入金済'] + '件';

      // テーブル描画
      if (data.companies.length === 0) {
        document.getElementById('resultsBody').innerHTML =
          '<div class="no-results">該当するデータがありません</div>';
        return;
      }

      let html = '<table class="results-table"><thead><tr>' +
        '<th>No</th>' +
        '<th>企業名</th>' +
        '<th>受診者数</th>' +
        '<th>請求額</th>' +
        '<th>ステータス内訳</th>' +
        '<th>操作</th>' +
        '</tr></thead><tbody>';

      data.companies.forEach((c, idx) => {
        const statusBadges = [];
        if (c.billingStatusCount['未請求'] > 0) {
          statusBadges.push('<span class="count-badge status-unbilled">未請求: ' + c.billingStatusCount['未請求'] + '</span>');
        }
        if (c.billingStatusCount['請求済'] > 0) {
          statusBadges.push('<span class="count-badge status-billed">請求済: ' + c.billingStatusCount['請求済'] + '</span>');
        }
        if (c.billingStatusCount['入金済'] > 0) {
          statusBadges.push('<span class="count-badge status-paid">入金済: ' + c.billingStatusCount['入金済'] + '</span>');
        }

        html += '<tr>' +
          '<td>' + (idx + 1) + '</td>' +
          '<td><strong>' + c.companyName + '</strong></td>' +
          '<td>' + c.patientCount + '名</td>' +
          '<td class="amount large">¥' + c.totalAmount.toLocaleString() + '</td>' +
          '<td><div class="count-badges">' + statusBadges.join('') + '</div></td>' +
          '<td><button class="action-btn" onclick="viewDetail(\\'' + c.companyName.replace(/'/g, "\\\\'") + '\\')">詳細</button></td>' +
          '</tr>';
      });

      html += '</tbody></table>';
      document.getElementById('resultsBody').innerHTML = html;
    }

    function viewDetail(companyName) {
      // 企業詳細ダイアログを開く（将来実装）
      const company = billingData.companies.find(c => c.companyName === companyName);
      if (company) {
        let details = '【' + companyName + '】\\n\\n';
        details += '受診者数: ' + company.patientCount + '名\\n';
        details += '請求額: ¥' + company.totalAmount.toLocaleString() + '\\n\\n';
        details += '--- 受診者一覧 ---\\n';
        company.patients.forEach(p => {
          details += p.name + ' (' + p.examDate + ') ' + p.course + ' ¥' + p.amount.toLocaleString() + ' [' + p.billingStatus + ']\\n';
        });
        alert(details);
      }
    }

    function exportToExcel() {
      if (!billingData || billingData.companies.length === 0) {
        alert('出力するデータがありません');
        return;
      }

      const criteria = {
        dateFrom: document.getElementById('filterDateFrom').value,
        dateTo: document.getElementById('filterDateTo').value,
        billingStatus: document.getElementById('filterStatus').value
      };

      google.script.run
        .withSuccessHandler((result) => {
          if (result.success) {
            alert('Excel出力完了\\n\\nファイル名: ' + result.fileName);
          } else {
            alert('出力エラー: ' + result.error);
          }
        })
        .withFailureHandler((e) => {
          alert('エラー: ' + e.message);
        })
        .exportBillingToExcel(criteria);
    }

    function initBillingColumn() {
      if (confirm('受診者マスタに「請求ステータス」列を追加します。続行しますか？')) {
        google.script.run
          .withSuccessHandler((result) => {
            alert(result.message || result.error);
            if (result.success) {
              loadBillingData();
            }
          })
          .initializeBillingStatusColumn();
      }
    }
  </script>
</body>
</html>
`;
}

// ============================================
// Excel出力機能
// ============================================

/**
 * 請求一覧をExcel出力
 * @param {Object} criteria - フィルタ条件
 * @returns {Object} 出力結果
 */
function exportBillingToExcel(criteria) {
  logInfo('請求一覧Excel出力開始');

  try {
    const billingData = getBillingList(criteria);

    if (billingData.companies.length === 0) {
      return { success: false, error: '出力するデータがありません' };
    }

    // 出力用スプレッドシートを作成
    const dateStr = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    const fileName = `請求一覧_${dateStr}`;

    const newSs = SpreadsheetApp.create(fileName);
    const sheet = newSs.getActiveSheet();
    sheet.setName('請求一覧');

    // ヘッダー
    const headers = ['No', '企業名', '受診者数', '請求額（税抜）', '消費税', '請求額（税込）',
                     '未請求', '請求済', '入金済'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');

    // データ行
    const rows = [];
    billingData.companies.forEach((c, idx) => {
      const tax = Math.floor(c.totalAmount * BILLING_CONFIG.TAX_RATE);
      rows.push([
        idx + 1,
        c.companyName,
        c.patientCount,
        c.totalAmount,
        tax,
        c.totalAmount + tax,
        c.billingStatusCount['未請求'] || 0,
        c.billingStatusCount['請求済'] || 0,
        c.billingStatusCount['入金済'] || 0
      ]);
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    // 合計行
    const totalRow = rows.length + 2;
    const totalTax = Math.floor(billingData.summary.totalAmount * BILLING_CONFIG.TAX_RATE);
    sheet.getRange(totalRow, 1, 1, headers.length).setValues([[
      '',
      '【合計】',
      billingData.summary.totalPatients,
      billingData.summary.totalAmount,
      totalTax,
      billingData.summary.totalAmount + totalTax,
      billingData.summary.byStatus['未請求'],
      billingData.summary.byStatus['請求済'],
      billingData.summary.byStatus['入金済']
    ]]);
    sheet.getRange(totalRow, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(totalRow, 1, 1, headers.length).setBackground('#f0f0f0');

    // 列幅調整
    sheet.setColumnWidth(1, 50);   // No
    sheet.setColumnWidth(2, 200);  // 企業名
    sheet.setColumnWidth(3, 80);   // 受診者数
    sheet.setColumnWidth(4, 120);  // 税抜
    sheet.setColumnWidth(5, 100);  // 消費税
    sheet.setColumnWidth(6, 120);  // 税込
    sheet.setColumnWidth(7, 80);   // 未請求
    sheet.setColumnWidth(8, 80);   // 請求済
    sheet.setColumnWidth(9, 80);   // 入金済

    // 金額列のフォーマット
    sheet.getRange(2, 4, rows.length + 1, 3).setNumberFormat('¥#,##0');

    // 出力先フォルダに移動
    const outputFolderId = getConfigValue('OUTPUT_FOLDER_ID');
    if (outputFolderId) {
      const file = DriveApp.getFileById(newSs.getId());
      const folder = DriveApp.getFolderById(outputFolderId);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }

    logInfo('請求一覧Excel出力完了: ' + fileName);

    return {
      success: true,
      fileName: fileName,
      url: newSs.getUrl()
    };

  } catch (e) {
    logError('exportBillingToExcel', e);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================
// テスト関数
// ============================================

/**
 * 請求一覧取得のテスト
 */
function testGetBillingList() {
  const result = getBillingList({
    sortBy: 'amount_desc'
  });
  logInfo('請求一覧: ' + result.companies.length + '社');
  logInfo('合計: ¥' + result.summary.totalAmount.toLocaleString());
}

/**
 * 請求一覧ダイアログのテスト表示
 */
function testShowBillingListDialog() {
  showBillingListDialog();
}

/**
 * 請求一覧Excel出力（メニューから直接実行）
 */
function testExportBillingToExcel() {
  const result = exportBillingToExcel({});

  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('Excel出力完了', `ファイル名: ${result.fileName}\n\n出力先フォルダに保存されました。`, ui.ButtonSet.OK);
  } else {
    ui.alert('エラー', result.error, ui.ButtonSet.OK);
  }
}
