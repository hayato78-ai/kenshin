/**
 * 健診結果DB 統合システム - 判定エンジン
 *
 * @description 設計書: 健診結果DB_設計書_v1.md 5章 判定基準
 * @version 1.0.0
 * @date 2025-12-14
 */

// ============================================
// 判定処理メイン
// ============================================

/**
 * 検査値の判定を行う
 * @param {string} itemId - 項目ID
 * @param {number|string} value - 検査値
 * @param {string} gender - 性別 (M/F)
 * @returns {string} 判定結果 (A/B/C/D または '')
 */
function calculateJudgment(itemId, value, gender) {
  // 項目マスタを取得
  const itemMaster = getItemMaster(itemId);
  if (!itemMaster) {
    logInfo(`判定スキップ: 項目マスタなし (${itemId})`);
    return '';
  }

  // 判定方法を確認
  const method = itemMaster.judgmentMethod;
  if (method === 'なし') {
    return '';
  }

  // 数値変換
  const numValue = parseFloat(value);
  if (isNaN(numValue)) {
    // 定性検査の場合
    if (itemMaster.dataType === '定性') {
      return judgeQualitative(value, itemMaster);
    }
    return '';
  }

  // 性別差がある項目の場合
  if (itemMaster.genderDiff && gender === 'F' && itemMaster.aMinF !== '') {
    return judgeWithFemaleRange(numValue, itemMaster);
  }

  // 通常の判定
  return judgeNumeric(numValue, itemMaster);
}

/**
 * 数値判定（通常）
 * @param {number} value - 検査値
 * @param {Object} item - 項目マスタ
 * @returns {string} 判定結果
 */
function judgeNumeric(value, item) {
  // D条件チェック（最優先）
  if (item.dCondition && checkDCondition(value, item.dCondition)) {
    return DB_CONFIG.JUDGMENT.D;
  }

  // A範囲チェック
  if (item.aMin !== '' && item.aMax !== '') {
    if (value >= item.aMin && value <= item.aMax) {
      return DB_CONFIG.JUDGMENT.A;
    }
  }

  // B範囲チェック
  if (item.bMin !== '' && item.bMax !== '') {
    if (value >= item.bMin && value <= item.bMax) {
      return DB_CONFIG.JUDGMENT.B;
    }
  }

  // C範囲チェック
  if (item.cMin !== '' && item.cMax !== '') {
    if (value >= item.cMin && value <= item.cMax) {
      return DB_CONFIG.JUDGMENT.C;
    }
  }

  // どの範囲にも該当しない場合はD
  return DB_CONFIG.JUDGMENT.D;
}

/**
 * 数値判定（女性用範囲）
 * @param {number} value - 検査値
 * @param {Object} item - 項目マスタ
 * @returns {string} 判定結果
 */
function judgeWithFemaleRange(value, item) {
  // D条件チェック（最優先）
  if (item.dCondition && checkDCondition(value, item.dCondition)) {
    return DB_CONFIG.JUDGMENT.D;
  }

  // 女性用A範囲チェック
  if (item.aMinF !== '' && item.aMaxF !== '') {
    if (value >= item.aMinF && value <= item.aMaxF) {
      return DB_CONFIG.JUDGMENT.A;
    }
  }

  // 通常の範囲で判定
  return judgeNumeric(value, item);
}

/**
 * D条件をチェック
 * @param {number} value - 検査値
 * @param {string} condition - D条件文字列
 * @returns {boolean} D条件に該当するか
 */
function checkDCondition(value, condition) {
  // 複数条件をカンマで分割
  const conditions = condition.split(',');

  for (const cond of conditions) {
    const trimmed = cond.trim();

    // >=条件
    if (trimmed.startsWith('>=')) {
      const threshold = parseFloat(trimmed.substring(2));
      if (!isNaN(threshold) && value >= threshold) {
        return true;
      }
    }
    // >条件
    else if (trimmed.startsWith('>')) {
      const threshold = parseFloat(trimmed.substring(1));
      if (!isNaN(threshold) && value > threshold) {
        return true;
      }
    }
    // <=条件
    else if (trimmed.startsWith('<=')) {
      const threshold = parseFloat(trimmed.substring(2));
      if (!isNaN(threshold) && value <= threshold) {
        return true;
      }
    }
    // <条件
    else if (trimmed.startsWith('<')) {
      const threshold = parseFloat(trimmed.substring(1));
      if (!isNaN(threshold) && value < threshold) {
        return true;
      }
    }
  }

  return false;
}

/**
 * 定性検査の判定
 * @param {string} value - 検査値
 * @param {Object} item - 項目マスタ
 * @returns {string} 判定結果
 */
function judgeQualitative(value, item) {
  const normalizedValue = value.toString().trim().toLowerCase();

  // 正常値パターン
  const normalPatterns = ['(-)', '-', '陰性', 'negative', '(±)', '±'];

  for (const pattern of normalPatterns) {
    if (normalizedValue === pattern.toLowerCase()) {
      return DB_CONFIG.JUDGMENT.A;
    }
  }

  // 異常値パターン
  const abnormalPatterns = ['(+)', '+', '陽性', 'positive', '(++)', '++', '(+++)', '+++'];

  for (const pattern of abnormalPatterns) {
    if (normalizedValue === pattern.toLowerCase()) {
      // (+)以上は要注意
      if (pattern.includes('++')) {
        return DB_CONFIG.JUDGMENT.C;
      }
      return DB_CONFIG.JUDGMENT.B;
    }
  }

  return '';
}

// ============================================
// 総合判定計算
// ============================================

/**
 * 受診の総合判定を計算
 * @param {string} visitId - 受診ID
 * @returns {string} 総合判定 (A/B/C/D)
 */
function calculateOverallJudgment(visitId) {
  const results = getTestResultsByVisitId(visitId);

  if (results.length === 0) {
    return '';
  }

  // 判定の優先度: D > C > B > A
  const judgmentPriority = {
    'D': 4,
    'C': 3,
    'B': 2,
    'A': 1,
    '': 0
  };

  let maxPriority = 0;
  let overallJudgment = '';

  for (const result of results) {
    const priority = judgmentPriority[result.judgment] || 0;
    if (priority > maxPriority) {
      maxPriority = priority;
      overallJudgment = result.judgment;
    }
  }

  return overallJudgment;
}

/**
 * 受診の全検査結果の判定を再計算
 * @param {string} visitId - 受診ID
 * @param {string} gender - 性別
 * @returns {Object} 処理結果 {updated: number, overall: string}
 */
function recalculateAllJudgments(visitId, gender) {
  const results = getTestResultsByVisitId(visitId);
  let updatedCount = 0;

  for (const result of results) {
    if (result.value !== '' && result.value !== null) {
      const newJudgment = calculateJudgment(result.itemId, result.value, gender);

      if (newJudgment !== result.judgment) {
        updateTestResult(result.resultId, { judgment: newJudgment });
        updatedCount++;
      }
    }
  }

  // 総合判定を更新
  const overallJudgment = calculateOverallJudgment(visitId);
  updateVisitRecord(visitId, { overallJudgment: overallJudgment });

  logInfo(`判定再計算: ${visitId} (${updatedCount}件更新, 総合: ${overallJudgment})`);

  return {
    updated: updatedCount,
    overall: overallJudgment
  };
}

// ============================================
// 検査結果入力＋自動判定
// ============================================

/**
 * 検査結果を入力して自動判定
 * @param {string} visitId - 受診ID
 * @param {string} itemId - 項目ID
 * @param {string|number} value - 検査値
 * @param {string} gender - 性別
 * @returns {Object} 結果 {resultId, judgment}
 */
function inputTestResultWithJudgment(visitId, itemId, value, gender) {
  // 判定を計算
  const judgment = calculateJudgment(itemId, value, gender);

  // 検査結果を作成
  const resultId = createTestResult({
    visitId: visitId,
    itemId: itemId,
    value: value,
    judgment: judgment
  });

  return {
    resultId: resultId,
    judgment: judgment
  };
}

/**
 * 複数の検査結果を一括入力して自動判定
 * @param {string} visitId - 受診ID
 * @param {Array<Object>} items - 検査項目配列 [{itemId, value}]
 * @param {string} gender - 性別
 * @returns {Object} 結果 {count, overall}
 */
function inputBatchTestResults(visitId, items, gender) {
  let count = 0;

  for (const item of items) {
    if (item.value !== '' && item.value !== null && item.value !== undefined) {
      inputTestResultWithJudgment(visitId, item.itemId, item.value, gender);
      count++;
    }
  }

  // 総合判定を計算・更新
  const overallJudgment = calculateOverallJudgment(visitId);
  updateVisitRecord(visitId, { overallJudgment: overallJudgment });

  logInfo(`一括入力完了: ${visitId} (${count}件, 総合: ${overallJudgment})`);

  return {
    count: count,
    overall: overallJudgment
  };
}

// ============================================
// 判定統計
// ============================================

/**
 * 受診の判定サマリを取得
 * @param {string} visitId - 受診ID
 * @returns {Object} 判定サマリ {A: n, B: n, C: n, D: n, total: n}
 */
function getJudgmentSummary(visitId) {
  const results = getTestResultsByVisitId(visitId);

  const summary = {
    A: 0,
    B: 0,
    C: 0,
    D: 0,
    total: 0
  };

  for (const result of results) {
    if (result.judgment) {
      summary[result.judgment] = (summary[result.judgment] || 0) + 1;
      summary.total++;
    }
  }

  return summary;
}

/**
 * カテゴリ別の判定サマリを取得
 * @param {string} visitId - 受診ID
 * @returns {Object} カテゴリ別判定サマリ
 */
function getJudgmentSummaryByCategory(visitId) {
  const results = getTestResultsByVisitId(visitId);
  const itemMasterList = getItemMaster();

  // 項目IDとカテゴリのマッピング
  const itemCategoryMap = {};
  for (const item of itemMasterList) {
    itemCategoryMap[item.itemId] = item.category;
  }

  const summary = {};

  for (const result of results) {
    const category = itemCategoryMap[result.itemId] || 'その他';

    if (!summary[category]) {
      summary[category] = { A: 0, B: 0, C: 0, D: 0, worst: '' };
    }

    if (result.judgment) {
      summary[category][result.judgment]++;

      // 最悪判定を更新
      const priority = { 'D': 4, 'C': 3, 'B': 2, 'A': 1 };
      const currentPriority = priority[summary[category].worst] || 0;
      const newPriority = priority[result.judgment] || 0;

      if (newPriority > currentPriority) {
        summary[category].worst = result.judgment;
      }
    }
  }

  return summary;
}

// ============================================
// 判定テスト関数
// ============================================

/**
 * 判定ロジックのテスト
 */
function testJudgmentEngine() {
  logInfo('===== 判定エンジンテスト =====');

  // BMI判定テスト
  logInfo('BMI 22.5 (正常): ' + calculateJudgment('BMI', 22.5, 'M'));  // A
  logInfo('BMI 26.0 (軽度異常): ' + calculateJudgment('BMI', 26.0, 'M'));  // B
  logInfo('BMI 30.0 (要経過観察): ' + calculateJudgment('BMI', 30.0, 'M'));  // C
  logInfo('BMI 36.0 (要精検): ' + calculateJudgment('BMI', 36.0, 'M'));  // D

  // 血糖判定テスト
  logInfo('FBS 90 (正常): ' + calculateJudgment('FBS', 90, 'M'));  // A
  logInfo('FBS 105 (軽度異常): ' + calculateJudgment('FBS', 105, 'M'));  // B
  logInfo('FBS 120 (要経過観察): ' + calculateJudgment('FBS', 120, 'M'));  // C
  logInfo('FBS 130 (要精検): ' + calculateJudgment('FBS', 130, 'M'));  // D

  // 性別差項目テスト
  logInfo('腹囲(男) 82 (正常): ' + calculateJudgment('WAIST_M', 82, 'M'));  // A
  logInfo('腹囲(女) 82 (正常): ' + calculateJudgment('WAIST_F', 82, 'F'));  // A

  // 定性検査テスト
  logInfo('尿蛋白 (-): ' + calculateJudgment('U_PRO', '(-)', 'M'));  // A
  logInfo('尿蛋白 (+): ' + calculateJudgment('U_PRO', '(+)', 'M'));  // B

  logInfo('===== テスト完了 =====');
}
