/**
 * 健診結果DB 統合システム - UI用サーバー関数
 *
 * @description UI画面から呼び出されるサーバー関数
 * @version 1.0.0
 * @date 2025-12-14
 */

// ============================================
// Webアプリ エントリーポイント
// ============================================

/**
 * Webアプリのメインエントリーポイント
 * @param {Object} e - イベントオブジェクト
 * @returns {HtmlOutput}
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('ui/Index');
  const html = template.evaluate()
    .setTitle('健診結果管理システム | CDmedical')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  return html;
}

/**
 * 現在のユーザーメールを取得
 * @returns {string}
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

// ============================================
// ダッシュボード用関数
// ============================================

/**
 * ダッシュボードサマリーを取得
 * @returns {Object} {todayVisits, pendingInput, overdue}
 */
function getDashboardSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visitSheet = ss.getSheetByName(DB_CONFIG.SHEETS.VISIT_RECORD);

  if (!visitSheet || visitSheet.getLastRow() < 2) {
    return { todayVisits: 0, pendingInput: 0, overdue: 0 };
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

  const data = visitSheet.getDataRange().getValues();
  const headers = data[0];

  const visitDateIdx = headers.indexOf('受診日');
  const statusIdx = headers.indexOf('ステータス');

  let todayVisits = 0;
  let pendingInput = 0;
  let overdue = 0;

  const weekAgo = new Date(today);
  weekAgo.setDate(weekAgo.getDate() - 7);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const visitDate = row[visitDateIdx];
    const status = row[statusIdx];

    // 日付を比較可能な形式に変換
    let visitDateObj;
    if (visitDate instanceof Date) {
      visitDateObj = visitDate;
    } else if (typeof visitDate === 'string') {
      visitDateObj = new Date(visitDate.replace(/\//g, '-'));
    }

    if (visitDateObj) {
      visitDateObj.setHours(0, 0, 0, 0);

      // 本日の受診
      if (visitDateObj.getTime() === today.getTime()) {
        todayVisits++;
      }

      // 入力待ち
      if (status === 'INPUTTING' || !status) {
        pendingInput++;

        // 1週間以上前で未確定は期限超過
        if (visitDateObj < weekAgo) {
          overdue++;
        }
      }
    }
  }

  return {
    todayVisits: todayVisits,
    pendingInput: pendingInput,
    overdue: overdue
  };
}

/**
 * 最近の受診を取得
 * @param {number} limit - 取得件数
 * @returns {Array}
 */
function getRecentVisits(limit) {
  limit = limit || 5;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visitSheet = ss.getSheetByName(DB_CONFIG.SHEETS.VISIT_RECORD);
  const patientSheet = ss.getSheetByName(DB_CONFIG.SHEETS.PATIENT_MASTER);
  const examTypeSheet = ss.getSheetByName(DB_CONFIG.SHEETS.EXAM_TYPE_MASTER);

  if (!visitSheet || visitSheet.getLastRow() < 2) {
    return [];
  }

  // 患者情報マップ
  const patientMap = {};
  if (patientSheet && patientSheet.getLastRow() >= 2) {
    const patientData = patientSheet.getDataRange().getValues();
    for (let i = 1; i < patientData.length; i++) {
      patientMap[patientData[i][0]] = {
        name: patientData[i][1]
      };
    }
  }

  // 検診種別マップ
  const examTypeMap = {};
  if (examTypeSheet && examTypeSheet.getLastRow() >= 2) {
    const examTypeData = examTypeSheet.getDataRange().getValues();
    for (let i = 1; i < examTypeData.length; i++) {
      examTypeMap[examTypeData[i][0]] = examTypeData[i][1];
    }
  }

  // 受診記録を取得（新しい順）
  const visitData = visitSheet.getDataRange().getValues();
  const headers = visitData[0];

  const visits = [];
  for (let i = 1; i < visitData.length; i++) {
    const row = visitData[i];
    visits.push({
      visitId: row[headers.indexOf('受診ID')],
      patientId: row[headers.indexOf('受診者ID')],
      visitDate: row[headers.indexOf('受診日')],
      examTypeId: row[headers.indexOf('検診種別ID')],
      overallJudgment: row[headers.indexOf('総合判定')],
      status: row[headers.indexOf('ステータス')]
    });
  }

  // 日付でソート（新しい順）
  visits.sort(function(a, b) {
    const dateA = a.visitDate instanceof Date ? a.visitDate : new Date(a.visitDate);
    const dateB = b.visitDate instanceof Date ? b.visitDate : new Date(b.visitDate);
    return dateB - dateA;
  });

  // 上位limit件を返す
  return visits.slice(0, limit).map(function(v) {
    return {
      visitId: v.visitId,
      patientId: v.patientId,
      patientName: patientMap[v.patientId]?.name || '-',
      visitDate: v.visitDate,
      examTypeName: examTypeMap[v.examTypeId] || v.examTypeId,
      overallJudgment: v.overallJudgment,
      status: v.status
    };
  });
}

// ============================================
// 検索用関数
// ============================================

/**
 * 受診者と受診情報を組み合わせて検索
 * @param {Object} criteria - 検索条件
 * @returns {Array}
 */
function searchPatientsWithVisits(criteria) {
  const patients = searchPatients(criteria);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visitSheet = ss.getSheetByName(DB_CONFIG.SHEETS.VISIT_RECORD);
  const examTypeSheet = ss.getSheetByName(DB_CONFIG.SHEETS.EXAM_TYPE_MASTER);

  // 検診種別マップ
  const examTypeMap = {};
  if (examTypeSheet && examTypeSheet.getLastRow() >= 2) {
    const examTypeData = examTypeSheet.getDataRange().getValues();
    for (let i = 1; i < examTypeData.length; i++) {
      examTypeMap[examTypeData[i][0]] = examTypeData[i][1];
    }
  }

  // 受診記録を取得
  const visitMap = {};
  if (visitSheet && visitSheet.getLastRow() >= 2) {
    const visitData = visitSheet.getDataRange().getValues();
    const headers = visitData[0];

    for (let i = 1; i < visitData.length; i++) {
      const row = visitData[i];
      const patientId = row[headers.indexOf('受診者ID')];
      const visitDate = row[headers.indexOf('受診日')];
      const examTypeId = row[headers.indexOf('検診種別ID')];
      const judgment = row[headers.indexOf('総合判定')];
      const status = row[headers.indexOf('ステータス')];

      // 日付条件でフィルタ
      if (criteria.dateFrom || criteria.dateTo) {
        const visitDateObj = visitDate instanceof Date ? visitDate : new Date(visitDate);
        if (criteria.dateFrom && visitDateObj < new Date(criteria.dateFrom)) continue;
        if (criteria.dateTo && visitDateObj > new Date(criteria.dateTo)) continue;
      }

      // 検診種別でフィルタ
      if (criteria.examType && examTypeId !== criteria.examType) continue;

      // ステータスでフィルタ
      if (criteria.status && status !== criteria.status) continue;

      if (!visitMap[patientId] || visitDate > visitMap[patientId].visitDate) {
        visitMap[patientId] = {
          visitDate: visitDate,
          examTypeId: examTypeId,
          judgment: judgment
        };
      }
    }
  }

  // 結果をマージ
  return patients.map(function(p) {
    const lastVisit = visitMap[p.patientId];
    return {
      patientId: p.patientId,
      name: p.name,
      nameKana: p.nameKana,
      lastVisitDate: lastVisit?.visitDate || null,
      lastExamTypeName: lastVisit ? (examTypeMap[lastVisit.examTypeId] || lastVisit.examTypeId) : null,
      lastJudgment: lastVisit?.judgment || null
    };
  }).filter(function(p) {
    // 日付条件がある場合、受診がない患者は除外
    if ((criteria.dateFrom || criteria.dateTo) && !p.lastVisitDate) {
      return false;
    }
    return true;
  });
}

// ============================================
// 登録用関数
// ============================================

/**
 * 受診者と受診を登録
 * @param {string|null} patientId - 既存受診者ID（新規の場合null）
 * @param {Object|null} patientData - 新規受診者データ
 * @param {Object} visitData - 受診データ
 * @returns {Object} {success, patientId, visitId}
 */
function registerPatientAndVisit(patientId, patientData, visitData) {
  try {
    // 新規受診者の場合
    if (!patientId && patientData) {
      patientId = createPatient(patientData);
    }

    if (!patientId) {
      throw new Error('受診者IDが取得できません');
    }

    // 受診記録を作成
    visitData.patientId = patientId;
    const visitId = createVisitRecord(visitData);

    return {
      success: true,
      patientId: patientId,
      visitId: visitId
    };

  } catch (e) {
    logError('registerPatientAndVisit: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================
// 検査結果用関数
// ============================================

/**
 * 前回の検査結果を取得
 * @param {string} patientId - 受診者ID
 * @param {Date|string} currentVisitDate - 現在の受診日
 * @returns {Object} itemId -> value のマップ
 */
function getPreviousTestResults(patientId, currentVisitDate) {
  const visits = getVisitRecordsByPatientId(patientId);

  if (!visits || visits.length < 2) {
    return {};
  }

  // 現在の受診日より前の最新受診を探す
  const currentDate = currentVisitDate instanceof Date ?
    currentVisitDate : new Date(currentVisitDate);

  let previousVisit = null;
  for (const visit of visits) {
    const visitDate = visit.visitDate instanceof Date ?
      visit.visitDate : new Date(visit.visitDate);

    if (visitDate < currentDate) {
      if (!previousVisit || visitDate > new Date(previousVisit.visitDate)) {
        previousVisit = visit;
      }
    }
  }

  if (!previousVisit) {
    return {};
  }

  // 前回の検査結果を取得
  const results = getTestResultsByVisitId(previousVisit.visitId);
  const resultMap = {};

  for (const r of results) {
    resultMap[r.itemId] = r.value;
  }

  return resultMap;
}

/**
 * 検査結果を保存（一括）
 * @param {string} visitId - 受診ID
 * @param {Array} items - 検査項目配列 [{itemId, value}]
 * @param {string} gender - 性別
 * @param {string} status - ステータス (INPUTTING/CONFIRMED)
 * @returns {Object} {success, count, overall}
 */
function saveTestResults(visitId, items, gender, status) {
  try {
    // 既存の結果を削除
    deleteTestResultsByVisitId(visitId);

    // 新しい結果を登録
    const result = inputBatchTestResults(visitId, items, gender);

    // ステータスを更新
    updateVisitRecord(visitId, { status: status });

    return {
      success: true,
      count: result.count,
      overall: result.overall
    };

  } catch (e) {
    logError('saveTestResults: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================
// CSVインポート用API関数（Phase 3）
// ============================================

/**
 * CSVプレビューAPI
 * @param {string} csvContent - CSVコンテンツ
 * @returns {Object} プレビュー結果
 */
function apiPreviewCsv(csvContent) {
  return previewCsv(csvContent, 10);
}

/**
 * AIマッピング推論API
 * @param {string[]} headers - ヘッダー
 * @param {string[][]} sampleRows - サンプル行
 * @param {string} dataType - データ種別
 * @returns {Object} マッピング結果
 */
function apiInferMapping(headers, sampleRows, dataType) {
  // まず既存パターンを検索
  const headersHash = calculateHeadersHash(headers);
  const existingPattern = findMappingPattern(headersHash, dataType);

  if (existingPattern) {
    incrementPatternUseCount(existingPattern.patternId);
    return {
      success: true,
      source: 'pattern',
      patternId: existingPattern.patternId,
      headersHash: headersHash,
      mappings: existingPattern.mappings,
      valueTransforms: existingPattern.valueTransforms,
      overallConfidence: 1.0,
      notes: '既存パターン「' + (existingPattern.sourceName || existingPattern.patternId) + '」を適用（使用回数: ' + (existingPattern.useCount + 1) + '）'
    };
  }

  // AIで推論
  return inferCsvMapping(headers, sampleRows, dataType);
}

/**
 * インポート実行API
 * @param {Object} params - パラメータ
 * @returns {Object} インポート結果
 */
function apiExecuteImport(params) {
  const { dataType, csvContent, mappings, options } = params;

  try {
    // CSVパース
    const parsed = parseCsv(csvContent);
    if (parsed.error) {
      return { success: false, error: parsed.error };
    }

    // マッピング適用
    const mappedData = applyMapping(parsed.headers, parsed.rows, mappings.mappings, mappings.valueTransforms);

    // インポート実行
    let results;
    if (dataType === CSV_CONFIG.DATA_TYPES.PATIENT) {
      results = importPatients(mappedData, options);
    } else {
      // 検査結果の場合：横持ちなら縦持ちに変換
      const verticalData = [];
      for (const row of mappedData) {
        const baseInfo = {
          visitDate: row.visitDate,
          name: row.name,
          kana: row.kana,
          birthdate: row.birthdate,
          gender: row.gender,
          company: row.company
        };

        // 検査項目データを縦持ちに展開
        const testItems = row._testItems || {};
        for (const [itemName, value] of Object.entries(testItems)) {
          verticalData.push({
            ...baseInfo,
            itemId: itemName,
            value: value
          });
        }
      }

      results = importTestResults(verticalData, options);
    }

    // パターン保存（オプション）
    if (options.savePattern && mappings.headersHash) {
      saveMappingPattern({
        headersHash: mappings.headersHash,
        dataType: dataType,
        sourceName: options.sourceName || '',
        mappings: mappings.mappings,
        valueTransforms: mappings.valueTransforms || {}
      });
    }

    // レポート生成
    const report = generateImportReport(results, dataType);

    return {
      success: true,
      results: results,
      report: report
    };

  } catch (e) {
    logError('apiExecuteImport', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * マスタデータ取得（インポート画面用）
 * @returns {Object} マスタデータ
 */
function getImportMasterData() {
  return {
    examTypes: getAllExamTypes(),
    items: getAllItems()
  };
}

// ============================================
// Phase 3: 検査項目マスタ拡張API
// ============================================

/**
 * 拡張検査項目マスタを取得（150項目）
 * MasterData.gsのEXAM_ITEM_MASTER_DATAを使用
 * @param {string} courseId - コースID（オプション）
 * @returns {Array} 検査項目リスト
 */
function getExtendedItemMaster(courseId) {
  // MasterData.gsのデータを使用
  if (typeof EXAM_ITEM_MASTER_DATA === 'undefined') {
    logError('EXAM_ITEM_MASTER_DATA is not defined. Please ensure MasterData.gs is loaded.');
    return [];
  }

  let items = EXAM_ITEM_MASTER_DATA.map(function(item, index) {
    return {
      itemId: item.item_id,
      name: item.item_name,
      category: item.category,
      subcategory: item.subcategory,
      dataType: item.data_type,
      unit: item.unit,
      requiredDock: item.required_dock,
      requiredRegular: item.required_regular,
      requiredSecondary: item.required_secondary,
      displayOrder: index + 1
    };
  });

  // コースでフィルタ
  if (courseId) {
    const courseItems = getRequiredItemsByCourse(courseId);
    if (courseItems && courseItems.length > 0) {
      // コース必須項目にマーク
      items = items.map(function(item) {
        return {
          ...item,
          isRequired: courseItems.includes(item.itemId)
        };
      });
    }
  }

  return items;
}

/**
 * カテゴリ別の項目マスタを取得
 * @param {string} category - カテゴリ名
 * @returns {Array} 検査項目リスト
 */
function getItemsByCategory(category) {
  const allItems = getExtendedItemMaster();
  if (!category) return allItems;

  return allItems.filter(function(item) {
    return item.category === category;
  });
}

/**
 * カテゴリ一覧を取得
 * @returns {Array} カテゴリ名リスト
 */
function getItemCategories() {
  const items = getExtendedItemMaster();
  const categorySet = new Set();

  items.forEach(function(item) {
    if (item.category) {
      categorySet.add(item.category);
    }
  });

  return Array.from(categorySet);
}

/**
 * 選択肢マスタを取得
 * @param {string} itemId - 項目ID（オプション）
 * @returns {Object|Array} 選択肢データ
 */
function getSelectOptions(itemId) {
  if (typeof SELECT_OPTIONS_DATA === 'undefined') {
    return itemId ? null : [];
  }

  if (itemId) {
    const option = SELECT_OPTIONS_DATA.find(function(opt) {
      return opt.item_id === itemId;
    });
    if (option) {
      return {
        itemId: option.item_id,
        options: option.options.split('|'),
        description: option.description
      };
    }
    return null;
  }

  return SELECT_OPTIONS_DATA.map(function(opt) {
    return {
      itemId: opt.item_id,
      options: opt.options.split('|'),
      description: opt.description
    };
  });
}

/**
 * コースマスタを取得（拡張版）
 * @returns {Array} コースリスト
 */
function getExtendedCourseMaster() {
  if (typeof EXAM_COURSE_MASTER_DATA === 'undefined') {
    return [];
  }

  return EXAM_COURSE_MASTER_DATA.map(function(course) {
    return {
      courseId: course.course_id,
      courseName: course.course_name,
      price: course.price,
      description: course.description,
      itemCount: course.item_count,
      requiredItems: course.required_items ? course.required_items.split(',') : []
    };
  });
}

/**
 * コース別必須項目を取得
 * @param {string} courseId - コースID
 * @returns {Array} 必須項目IDリスト
 */
function getCourseRequiredItems(courseId) {
  const courses = getExtendedCourseMaster();
  const course = courses.find(function(c) {
    return c.courseId === courseId;
  });

  return course ? course.requiredItems : [];
}

// ============================================
// Phase 3: 判定計算API
// ============================================

/**
 * 検査項目の判定を計算
 * @param {string} itemId - 項目ID
 * @param {number|string} value - 検査値
 * @param {string} gender - 性別（M/F）
 * @returns {string} 判定（A/B/C/D）
 */
function calculateJudgment(itemId, value, gender) {
  // JudgmentLogic.gsのgetJudgment関数を呼び出し
  if (typeof getJudgment === 'function') {
    return getJudgment(itemId, value, gender);
  }

  logError('getJudgment function not found. Please ensure JudgmentLogic.gs is loaded.');
  return null;
}

/**
 * 糖代謝の組合せ判定を計算
 * @param {number} fbs - 空腹時血糖
 * @param {number} hba1c - HbA1c
 * @returns {string} 判定（A/B/C/D）
 */
function calculateGlucoseJudgment(fbs, hba1c) {
  if (typeof getGlucoseHbA1cJudgment === 'function') {
    return getGlucoseHbA1cJudgment(fbs, hba1c);
  }

  logError('getGlucoseHbA1cJudgment function not found.');
  return null;
}

/**
 * 総合判定を計算
 * @param {string} visitId - 受診ID
 * @returns {Object} 総合判定結果
 */
function calculateOverallJudgmentForVisit(visitId) {
  // 受診情報を取得
  const visit = getVisitRecordById(visitId);
  if (!visit) {
    return { success: false, error: '受診情報が見つかりません' };
  }

  // 受診者情報を取得
  const patient = getPatientById(visit.patientId);
  const gender = patient ? patient.gender : 'M';

  // 検査結果を取得
  const results = getTestResultsByVisitId(visitId);
  if (!results || results.length === 0) {
    return { success: false, error: '検査結果がありません' };
  }

  // 結果をマップに変換
  const resultMap = {};
  results.forEach(function(r) {
    resultMap[r.itemId] = r.value;
  });

  // 総合判定を計算
  if (typeof calculateOverallJudgment === 'function') {
    const overall = calculateOverallJudgment(resultMap, gender);
    return {
      success: true,
      judgment: overall.judgment,
      label: overall.label,
      summary: overall.summary,
      worstItems: overall.worstItems
    };
  }

  return { success: false, error: 'calculateOverallJudgment function not found' };
}

/**
 * 検査結果を一括入力（判定自動計算付き）
 * @param {string} visitId - 受診ID
 * @param {Array} items - 項目配列 [{itemId, value}]
 * @param {string} gender - 性別
 * @returns {Object} {success, count, overall}
 */
function inputBatchTestResults(visitId, items, gender) {
  try {
    let count = 0;
    const resultMap = {};

    for (const item of items) {
      // 判定を計算
      const judgment = calculateJudgment(item.itemId, item.value, gender);

      // 検査結果を作成
      createTestResult({
        visitId: visitId,
        itemId: item.itemId,
        value: item.value,
        judgment: judgment || ''
      });

      resultMap[item.itemId] = item.value;
      count++;
    }

    // 総合判定を計算
    let overall = null;
    if (typeof calculateOverallJudgment === 'function' && count > 0) {
      overall = calculateOverallJudgment(resultMap, gender);

      // 受診記録の総合判定を更新
      updateVisitRecord(visitId, {
        overallJudgment: overall.judgment
      });
    }

    return {
      success: true,
      count: count,
      overall: overall
    };

  } catch (e) {
    logError('inputBatchTestResults: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 判定ラベルを取得
 * @param {string} judgment - 判定（A/B/C/D/E/F）
 * @returns {string} 判定ラベル
 */
function getJudgmentLabelApi(judgment) {
  if (typeof getJudgmentLabel === 'function') {
    return getJudgmentLabel(judgment);
  }

  const labels = {
    'A': '異常なし',
    'B': '軽度異常',
    'C': '要再検査・生活改善',
    'D': '要精密検査・治療',
    'E': '治療中',
    'F': '経過観察中'
  };
  return labels[judgment] || judgment;
}

// ============================================
// Phase 3: 入力画面用データ取得API
// ============================================

/**
 * 入力画面用の初期データを取得
 * @param {string} visitId - 受診ID
 * @returns {Object} 初期データ
 */
function getInputScreenData(visitId) {
  try {
    // 受診情報
    const visit = getVisitRecordById(visitId);
    if (!visit) {
      return { success: false, error: '受診情報が見つかりません' };
    }

    // 受診者情報
    const patient = getPatientById(visit.patientId);

    // 検査項目マスタ（コースでフィルタ）
    const items = getExtendedItemMaster(visit.courseId);

    // 選択肢マスタ
    const selectOptions = getSelectOptions();

    // 既存の検査結果
    const existingResults = getTestResultsByVisitId(visitId);
    const resultMap = {};
    existingResults.forEach(function(r) {
      resultMap[r.itemId] = {
        value: r.value,
        judgment: r.judgment
      };
    });

    // 前回の検査結果
    const previousResults = getPreviousTestResults(visit.patientId, visit.visitDate);

    return {
      success: true,
      visit: visit,
      patient: patient,
      items: items,
      selectOptions: selectOptions,
      existingResults: resultMap,
      previousResults: previousResults
    };

  } catch (e) {
    logError('getInputScreenData: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}
