/**
 * 組織マスタ管理モジュール
 * 健康保険組合・事業所・企業・コースのCRUD操作
 * Phase 3: 組織階層管理
 */

// ============================================
// ID採番関数
// ============================================

/**
 * 健康保険組合IDを生成
 * @returns {string} HI00001形式のID
 */
function generateHealthInsuranceId() {
  return generateSequentialId(CONFIG.SHEETS.HEALTH_INSURANCE, 'HI', 5);
}

/**
 * 事業所IDを生成
 * @returns {string} BO00001形式のID
 */
function generateBusinessOfficeId() {
  return generateSequentialId(CONFIG.SHEETS.BUSINESS_OFFICE, 'BO', 5);
}

/**
 * 企業IDを生成
 * @returns {string} CO00001形式のID
 */
function generateCompanyId() {
  return generateSequentialId(CONFIG.SHEETS.COMPANY, 'CO', 5);
}

/**
 * コースIDを生成
 * @returns {string} CRS001形式のID
 */
function generateCourseId() {
  return generateSequentialId(CONFIG.SHEETS.COURSE, 'CRS', 3);
}

/**
 * 連番IDを生成
 * @param {string} sheetName - シート名
 * @param {string} prefix - プレフィックス
 * @param {number} digits - 桁数
 * @returns {string} 生成されたID
 */
function generateSequentialId(sheetName, prefix, digits) {
  const sheet = getSheet(sheetName);
  if (!sheet) {
    return `${prefix}${'0'.repeat(digits - 1)}1`;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return `${prefix}${'0'.repeat(digits - 1)}1`;
  }

  // 最終行のIDを取得
  const lastId = sheet.getRange(lastRow, 1).getValue();
  const numPart = parseInt(lastId.replace(prefix, ''), 10);
  const nextNum = numPart + 1;

  return `${prefix}${String(nextNum).padStart(digits, '0')}`;
}

// ============================================
// 健康保険組合 CRUD
// ============================================

/**
 * 健康保険組合を登録
 * @param {Object} data - 組合データ
 * @returns {Object} 結果
 */
function createHealthInsurance(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.HEALTH_INSURANCE);
    if (!sheet) {
      throw new Error('健康保険組合マスタシートが見つかりません');
    }

    const id = generateHealthInsuranceId();
    const now = new Date();

    const row = [
      id,
      data.name || '',
      data.code || '',
      data.phone || '',
      data.address || '',
      data.memo || '',
      data.sortOrder || sheet.getLastRow(),
      data.isActive !== false,
      now,
      now
    ];

    sheet.appendRow(row);
    logInfo(`健康保険組合登録: ${id} - ${data.name}`);

    return { success: true, id: id };
  } catch (e) {
    logError('createHealthInsurance', e);
    return { success: false, error: e.message };
  }
}

/**
 * 健康保険組合一覧を取得
 * @param {boolean} activeOnly - 有効のみ取得
 * @returns {Array<Object>} 組合一覧
 */
function getHealthInsuranceList(activeOnly = true) {
  const sheet = getSheet(CONFIG.SHEETS.HEALTH_INSURANCE);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const result = [];

  for (const row of data) {
    if (activeOnly && row[7] === false) continue;

    result.push({
      insuranceId: row[0],
      name: row[1],
      code: row[2],
      phone: row[3],
      address: row[4],
      memo: row[5],
      sortOrder: row[6],
      isActive: row[7],
      createdAt: row[8],
      updatedAt: row[9]
    });
  }

  return result.sort((a, b) => a.sortOrder - b.sortOrder);
}

/**
 * 健康保険組合を更新
 * @param {string} insuranceId - 組合ID
 * @param {Object} data - 更新データ
 * @returns {Object} 結果
 */
function updateHealthInsurance(insuranceId, data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.HEALTH_INSURANCE);
    if (!sheet) {
      throw new Error('健康保険組合マスタシートが見つかりません');
    }

    const row = findRowById(sheet, insuranceId);
    if (row === -1) {
      throw new Error(`組合ID ${insuranceId} が見つかりません`);
    }

    const now = new Date();

    if (data.name !== undefined) sheet.getRange(row, 2).setValue(data.name);
    if (data.code !== undefined) sheet.getRange(row, 3).setValue(data.code);
    if (data.phone !== undefined) sheet.getRange(row, 4).setValue(data.phone);
    if (data.address !== undefined) sheet.getRange(row, 5).setValue(data.address);
    if (data.memo !== undefined) sheet.getRange(row, 6).setValue(data.memo);
    if (data.sortOrder !== undefined) sheet.getRange(row, 7).setValue(data.sortOrder);
    if (data.isActive !== undefined) sheet.getRange(row, 8).setValue(data.isActive);
    sheet.getRange(row, 10).setValue(now);

    logInfo(`健康保険組合更新: ${insuranceId}`);
    return { success: true };
  } catch (e) {
    logError('updateHealthInsurance', e);
    return { success: false, error: e.message };
  }
}

// ============================================
// 事業所 CRUD
// ============================================

/**
 * 事業所を登録
 * @param {Object} data - 事業所データ
 * @returns {Object} 結果
 */
function createBusinessOffice(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BUSINESS_OFFICE);
    if (!sheet) {
      throw new Error('事業所マスタシートが見つかりません');
    }

    // 親組合の存在確認
    if (!validateHealthInsuranceExists(data.insuranceId)) {
      throw new Error(`健康保険組合 ${data.insuranceId} が存在しません`);
    }

    const id = generateBusinessOfficeId();
    const now = new Date();

    const row = [
      id,
      data.insuranceId,
      data.name || '',
      data.code || '',
      data.phone || '',
      data.address || '',
      data.memo || '',
      data.sortOrder || sheet.getLastRow(),
      data.isActive !== false,
      now,
      now
    ];

    sheet.appendRow(row);
    logInfo(`事業所登録: ${id} - ${data.name}`);

    return { success: true, id: id };
  } catch (e) {
    logError('createBusinessOffice', e);
    return { success: false, error: e.message };
  }
}

/**
 * 事業所一覧を取得
 * @param {string} insuranceId - 組合ID（省略で全件）
 * @param {boolean} activeOnly - 有効のみ取得
 * @returns {Array<Object>} 事業所一覧
 */
function getBusinessOfficeList(insuranceId = null, activeOnly = true) {
  const sheet = getSheet(CONFIG.SHEETS.BUSINESS_OFFICE);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const result = [];

  for (const row of data) {
    if (activeOnly && row[8] === false) continue;
    if (insuranceId && row[1] !== insuranceId) continue;

    result.push({
      officeId: row[0],
      insuranceId: row[1],
      name: row[2],
      code: row[3],
      phone: row[4],
      address: row[5],
      memo: row[6],
      sortOrder: row[7],
      isActive: row[8],
      createdAt: row[9],
      updatedAt: row[10]
    });
  }

  return result.sort((a, b) => a.sortOrder - b.sortOrder);
}

/**
 * 事業所を更新
 * @param {string} officeId - 事業所ID
 * @param {Object} data - 更新データ
 * @returns {Object} 結果
 */
function updateBusinessOffice(officeId, data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BUSINESS_OFFICE);
    if (!sheet) {
      throw new Error('事業所マスタシートが見つかりません');
    }

    const row = findRowById(sheet, officeId);
    if (row === -1) {
      throw new Error(`事業所ID ${officeId} が見つかりません`);
    }

    const now = new Date();

    if (data.insuranceId !== undefined) {
      if (!validateHealthInsuranceExists(data.insuranceId)) {
        throw new Error(`健康保険組合 ${data.insuranceId} が存在しません`);
      }
      sheet.getRange(row, 2).setValue(data.insuranceId);
    }
    if (data.name !== undefined) sheet.getRange(row, 3).setValue(data.name);
    if (data.code !== undefined) sheet.getRange(row, 4).setValue(data.code);
    if (data.phone !== undefined) sheet.getRange(row, 5).setValue(data.phone);
    if (data.address !== undefined) sheet.getRange(row, 6).setValue(data.address);
    if (data.memo !== undefined) sheet.getRange(row, 7).setValue(data.memo);
    if (data.sortOrder !== undefined) sheet.getRange(row, 8).setValue(data.sortOrder);
    if (data.isActive !== undefined) sheet.getRange(row, 9).setValue(data.isActive);
    sheet.getRange(row, 11).setValue(now);

    logInfo(`事業所更新: ${officeId}`);
    return { success: true };
  } catch (e) {
    logError('updateBusinessOffice', e);
    return { success: false, error: e.message };
  }
}

// ============================================
// 企業 CRUD
// ============================================

/**
 * 企業を登録
 * @param {Object} data - 企業データ
 * @returns {Object} 結果
 */
function createCompany(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.COMPANY);
    if (!sheet) {
      throw new Error('企業マスタシートが見つかりません');
    }

    // 親事業所の存在確認
    if (!validateBusinessOfficeExists(data.officeId)) {
      throw new Error(`事業所 ${data.officeId} が存在しません`);
    }

    const id = generateCompanyId();
    const now = new Date();

    const row = [
      id,
      data.officeId,
      data.name || '',
      data.code || '',
      data.phone || '',
      data.address || '',
      data.contactPerson || '',
      data.contactEmail || '',
      data.memo || '',
      data.sortOrder || sheet.getLastRow(),
      data.isActive !== false,
      now,
      now
    ];

    sheet.appendRow(row);
    logInfo(`企業登録: ${id} - ${data.name}`);

    return { success: true, id: id };
  } catch (e) {
    logError('createCompany', e);
    return { success: false, error: e.message };
  }
}

/**
 * 企業一覧を取得
 * @param {string} officeId - 事業所ID（省略で全件）
 * @param {boolean} activeOnly - 有効のみ取得
 * @returns {Array<Object>} 企業一覧
 */
function getCompanyList(officeId = null, activeOnly = true) {
  const sheet = getSheet(CONFIG.SHEETS.COMPANY);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const result = [];

  for (const row of data) {
    if (activeOnly && row[10] === false) continue;
    if (officeId && row[1] !== officeId) continue;

    result.push({
      companyId: row[0],
      officeId: row[1],
      name: row[2],
      code: row[3],
      phone: row[4],
      address: row[5],
      contactPerson: row[6],
      contactEmail: row[7],
      memo: row[8],
      sortOrder: row[9],
      isActive: row[10],
      createdAt: row[11],
      updatedAt: row[12]
    });
  }

  return result.sort((a, b) => a.sortOrder - b.sortOrder);
}

/**
 * 企業を更新
 * @param {string} companyId - 企業ID
 * @param {Object} data - 更新データ
 * @returns {Object} 結果
 */
function updateCompany(companyId, data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.COMPANY);
    if (!sheet) {
      throw new Error('企業マスタシートが見つかりません');
    }

    const row = findRowById(sheet, companyId);
    if (row === -1) {
      throw new Error(`企業ID ${companyId} が見つかりません`);
    }

    const now = new Date();

    if (data.officeId !== undefined) {
      if (!validateBusinessOfficeExists(data.officeId)) {
        throw new Error(`事業所 ${data.officeId} が存在しません`);
      }
      sheet.getRange(row, 2).setValue(data.officeId);
    }
    if (data.name !== undefined) sheet.getRange(row, 3).setValue(data.name);
    if (data.code !== undefined) sheet.getRange(row, 4).setValue(data.code);
    if (data.phone !== undefined) sheet.getRange(row, 5).setValue(data.phone);
    if (data.address !== undefined) sheet.getRange(row, 6).setValue(data.address);
    if (data.contactPerson !== undefined) sheet.getRange(row, 7).setValue(data.contactPerson);
    if (data.contactEmail !== undefined) sheet.getRange(row, 8).setValue(data.contactEmail);
    if (data.memo !== undefined) sheet.getRange(row, 9).setValue(data.memo);
    if (data.sortOrder !== undefined) sheet.getRange(row, 10).setValue(data.sortOrder);
    if (data.isActive !== undefined) sheet.getRange(row, 11).setValue(data.isActive);
    sheet.getRange(row, 13).setValue(now);

    logInfo(`企業更新: ${companyId}`);
    return { success: true };
  } catch (e) {
    logError('updateCompany', e);
    return { success: false, error: e.message };
  }
}

/**
 * 企業IDから企業情報を取得
 * @param {string} companyId - 企業ID
 * @returns {Object|null} 企業情報
 */
function getCompanyById(companyId) {
  const companies = getCompanyList(null, false);
  return companies.find(c => c.companyId === companyId) || null;
}

// ============================================
// 企業-コース紐付け CRUD
// ============================================

/**
 * 企業-コース紐付けを登録
 * @param {Object} data - 紐付けデータ
 * @returns {Object} 結果
 */
function createCompanyCourse(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.COMPANY_COURSE);
    if (!sheet) {
      throw new Error('企業コース紐付けシートが見つかりません');
    }

    // 重複チェック
    if (companyCourseExists(data.companyId, data.courseId)) {
      throw new Error(`企業 ${data.companyId} とコース ${data.courseId} の紐付けは既に存在します`);
    }

    const now = new Date();

    const row = [
      data.companyId,
      data.courseId,
      data.isDefault || false,
      data.burdenType || null,
      data.insuranceBurden || null,
      data.officeBurden || null,
      data.personalBurden || null,
      data.memo || '',
      data.isActive !== false,
      now
    ];

    sheet.appendRow(row);
    logInfo(`企業-コース紐付け登録: ${data.companyId} - ${data.courseId}`);

    return { success: true };
  } catch (e) {
    logError('createCompanyCourse', e);
    return { success: false, error: e.message };
  }
}

/**
 * 企業のコース一覧を取得
 * @param {string} companyId - 企業ID
 * @param {boolean} activeOnly - 有効のみ取得
 * @returns {Array<Object>} コース紐付け一覧
 */
function getCompanyCourseList(companyId, activeOnly = true) {
  const sheet = getSheet(CONFIG.SHEETS.COMPANY_COURSE);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const result = [];

  for (const row of data) {
    if (activeOnly && row[8] === false) continue;
    if (row[0] !== companyId) continue;

    result.push({
      companyId: row[0],
      courseId: row[1],
      isDefault: row[2],
      burdenType: row[3],
      insuranceBurden: row[4],
      officeBurden: row[5],
      personalBurden: row[6],
      memo: row[7],
      isActive: row[8],
      createdAt: row[9]
    });
  }

  return result;
}

/**
 * 企業-コース紐付けが存在するか確認
 * @param {string} companyId - 企業ID
 * @param {string} courseId - コースID
 * @returns {boolean}
 */
function companyCourseExists(companyId, courseId) {
  const courses = getCompanyCourseList(companyId, false);
  return courses.some(c => c.courseId === courseId);
}

/**
 * 企業-コース紐付けを削除（論理削除）
 * @param {string} companyId - 企業ID
 * @param {string} courseId - コースID
 * @returns {Object} 結果
 */
function deleteCompanyCourse(companyId, courseId) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.COMPANY_COURSE);
    if (!sheet) {
      throw new Error('企業コース紐付けシートが見つかりません');
    }

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === companyId && data[i][1] === courseId) {
        sheet.getRange(i + 2, 9).setValue(false); // isActive = false
        logInfo(`企業-コース紐付け削除: ${companyId} - ${courseId}`);
        return { success: true };
      }
    }

    throw new Error('紐付けが見つかりません');
  } catch (e) {
    logError('deleteCompanyCourse', e);
    return { success: false, error: e.message };
  }
}

// ============================================
// 検診種別・コース CRUD
// ============================================

/**
 * 検診種別一覧を取得
 * @param {boolean} activeOnly - 有効のみ取得
 * @returns {Array<Object>} 検診種別一覧
 */
function getExamTypeList(activeOnly = true) {
  const sheet = getSheet(CONFIG.SHEETS.EXAM_TYPE);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const result = [];

  for (const row of data) {
    if (activeOnly && row[4] === false) continue;

    result.push({
      examTypeId: row[0],
      name: row[1],
      code: row[2],
      sortOrder: row[3],
      isActive: row[4],
      createdAt: row[5],
      updatedAt: row[6]
    });
  }

  return result.sort((a, b) => a.sortOrder - b.sortOrder);
}

/**
 * コース一覧を取得
 * @param {string} examTypeId - 検診種別ID（省略で全件）
 * @param {boolean} activeOnly - 有効のみ取得
 * @returns {Array<Object>} コース一覧
 */
function getCourseList(examTypeId = null, activeOnly = true) {
  const sheet = getSheet(CONFIG.SHEETS.COURSE);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const result = [];

  for (const row of data) {
    if (activeOnly && row[6] === false) continue;
    if (examTypeId && row[1] !== examTypeId) continue;

    result.push({
      courseId: row[0],
      examTypeId: row[1],
      name: row[2],
      code: row[3],
      price: row[4],
      sortOrder: row[5],
      isActive: row[6],
      createdAt: row[7],
      updatedAt: row[8]
    });
  }

  return result.sort((a, b) => a.sortOrder - b.sortOrder);
}

/**
 * コースIDからコース情報を取得
 * @param {string} courseId - コースID
 * @returns {Object|null} コース情報
 */
function getCourseById(courseId) {
  const courses = getCourseList(null, false);
  return courses.find(c => c.courseId === courseId) || null;
}

// ============================================
// バリデーション関数
// ============================================

/**
 * 健康保険組合の存在確認
 * @param {string} insuranceId - 組合ID
 * @returns {boolean}
 */
function validateHealthInsuranceExists(insuranceId) {
  const list = getHealthInsuranceList(false);
  return list.some(item => item.insuranceId === insuranceId);
}

/**
 * 事業所の存在確認
 * @param {string} officeId - 事業所ID
 * @returns {boolean}
 */
function validateBusinessOfficeExists(officeId) {
  const list = getBusinessOfficeList(null, false);
  return list.some(item => item.officeId === officeId);
}

/**
 * 企業の存在確認
 * @param {string} companyId - 企業ID
 * @returns {boolean}
 */
function validateCompanyExists(companyId) {
  const list = getCompanyList(null, false);
  return list.some(item => item.companyId === companyId);
}

/**
 * コースの存在確認
 * @param {string} courseId - コースID
 * @returns {boolean}
 */
function validateCourseExists(courseId) {
  const list = getCourseList(null, false);
  return list.some(item => item.courseId === courseId);
}

// ============================================
// ユーティリティ関数
// ============================================

/**
 * シート内でIDを検索して行番号を返す
 * @param {Sheet} sheet - シート
 * @param {string} id - 検索するID
 * @returns {number} 行番号（1始まり）、見つからない場合-1
 */
function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      return i + 2;
    }
  }

  return -1;
}

/**
 * 組織階層を取得（企業IDから健保まで）
 * @param {string} companyId - 企業ID
 * @returns {Object} 階層情報
 */
function getOrganizationHierarchy(companyId) {
  const company = getCompanyById(companyId);
  if (!company) return null;

  const offices = getBusinessOfficeList(null, false);
  const office = offices.find(o => o.officeId === company.officeId);
  if (!office) return { company };

  const insurances = getHealthInsuranceList(false);
  const insurance = insurances.find(i => i.insuranceId === office.insuranceId);

  return {
    insurance: insurance || null,
    office: office,
    company: company
  };
}

// ============================================
// 組織マスタシート作成
// ============================================

/**
 * 組織マスタシートを作成
 * 既存のcreateAllSheets()から呼び出し可能
 */
function createOrganizationMasterSheets() {
  logInfo('===== 組織マスタシート作成開始 =====');

  const ss = getSpreadsheet();
  const orgSheets = [
    CONFIG.SHEETS.HEALTH_INSURANCE,
    CONFIG.SHEETS.BUSINESS_OFFICE,
    CONFIG.SHEETS.COMPANY,
    CONFIG.SHEETS.COMPANY_COURSE,
    CONFIG.SHEETS.EXAM_TYPE,
    CONFIG.SHEETS.COURSE
  ];

  for (const sheetName of orgSheets) {
    if (SHEET_DEFINITIONS[sheetName]) {
      createSheet(ss, sheetName, SHEET_DEFINITIONS[sheetName]);
    }
  }

  logInfo('===== 組織マスタシート作成完了 =====');
}

// ============================================
// メニュー用表示関数
// ============================================

/**
 * 健康保険組合一覧を表示
 */
function showHealthInsuranceList() {
  const ui = SpreadsheetApp.getUi();
  const list = getHealthInsuranceList();

  if (list.length === 0) {
    ui.alert('健康保険組合一覧', '登録されている健康保険組合はありません。', ui.ButtonSet.OK);
    return;
  }

  let message = `健康保険組合一覧 (${list.length}件)\n\n`;
  for (const item of list) {
    message += `${item.insuranceId}: ${item.name}\n`;
  }

  ui.alert('健康保険組合一覧', message, ui.ButtonSet.OK);
}

/**
 * 事業所一覧を表示
 */
function showBusinessOfficeList() {
  const ui = SpreadsheetApp.getUi();
  const list = getBusinessOfficeList();

  if (list.length === 0) {
    ui.alert('事業所一覧', '登録されている事業所はありません。', ui.ButtonSet.OK);
    return;
  }

  let message = `事業所一覧 (${list.length}件)\n\n`;
  for (const item of list) {
    message += `${item.officeId}: ${item.name} (${item.insuranceId})\n`;
  }

  ui.alert('事業所一覧', message, ui.ButtonSet.OK);
}

/**
 * 企業一覧を表示
 */
function showCompanyList() {
  const ui = SpreadsheetApp.getUi();
  const list = getCompanyList();

  if (list.length === 0) {
    ui.alert('企業一覧', '登録されている企業はありません。', ui.ButtonSet.OK);
    return;
  }

  let message = `企業一覧 (${list.length}件)\n\n`;
  for (const item of list) {
    message += `${item.companyId}: ${item.name}\n`;
  }

  ui.alert('企業一覧', message, ui.ButtonSet.OK);
}

/**
 * 検診種別一覧を表示
 */
function showExamTypeList() {
  const ui = SpreadsheetApp.getUi();
  const list = getExamTypeList();

  if (list.length === 0) {
    ui.alert('検診種別一覧', '登録されている検診種別はありません。', ui.ButtonSet.OK);
    return;
  }

  let message = `検診種別一覧 (${list.length}件)\n\n`;
  for (const item of list) {
    message += `${item.examTypeId}: ${item.name}\n`;
  }

  ui.alert('検診種別一覧', message, ui.ButtonSet.OK);
}

/**
 * コース一覧を表示
 */
function showCourseList() {
  const ui = SpreadsheetApp.getUi();
  const list = getCourseList();

  if (list.length === 0) {
    ui.alert('コース一覧', '登録されているコースはありません。', ui.ButtonSet.OK);
    return;
  }

  let message = `コース一覧 (${list.length}件)\n\n`;
  for (const item of list) {
    const price = item.price ? `¥${item.price.toLocaleString()}` : '無料';
    message += `${item.courseId}: ${item.name} (${price})\n`;
  }

  ui.alert('コース一覧', message, ui.ButtonSet.OK);
}

// ============================================
// テスト関数
// ============================================

/**
 * 組織マスタCRUDのテスト
 */
function testOrganizationMasterCrud() {
  logInfo('===== 組織マスタCRUDテスト開始 =====');

  // 1. 健康保険組合登録
  const hiResult = createHealthInsurance({
    name: 'テスト健保組合',
    code: 'TEST001',
    phone: '03-1111-2222'
  });
  logInfo(`健保登録結果: ${JSON.stringify(hiResult)}`);

  if (hiResult.success) {
    // 2. 事業所登録
    const boResult = createBusinessOffice({
      insuranceId: hiResult.id,
      name: 'テスト事業所',
      code: 'TEST-BO001'
    });
    logInfo(`事業所登録結果: ${JSON.stringify(boResult)}`);

    if (boResult.success) {
      // 3. 企業登録
      const coResult = createCompany({
        officeId: boResult.id,
        name: 'テスト株式会社',
        code: 'TEST-CO001'
      });
      logInfo(`企業登録結果: ${JSON.stringify(coResult)}`);

      if (coResult.success) {
        // 4. 企業-コース紐付け
        const ccResult = createCompanyCourse({
          companyId: coResult.id,
          courseId: 'CRS001',
          isDefault: true
        });
        logInfo(`企業-コース紐付け結果: ${JSON.stringify(ccResult)}`);

        // 5. 階層取得
        const hierarchy = getOrganizationHierarchy(coResult.id);
        logInfo(`組織階層: ${JSON.stringify(hierarchy)}`);
      }
    }
  }

  logInfo('===== 組織マスタCRUDテスト完了 =====');
}
