/**
 * judgmentEngine.gs - 統合版判定エンジン（最終版）
 *
 * 既存のjudgmentEngine.gsとJudgmentLogic.gsを統合
 * 人間ドック学会2025年度版判定基準に準拠
 *
 * @version 2.1.0
 * @date 2025-12-17
 *
 * 【統合内容】
 * - 既存: BMLコード対応、シート保存機能、judge()関数
 * - 新規: 外部マスタ参照、組合せ判定、API関数、バッチ処理
 *
 * 【外部依存】
 * - CONFIG: システム設定（Config.gsで定義）
 * - getSheet(): シート取得関数（utils.gsで定義）
 * - findPatientRow(): 患者行検索（utils.gsで定義）
 * - EXAM_ITEM_MASTER_DATA: 検査項目マスタ（MasterData.gsで定義、オプション）
 */

// ============================================
// 判定重み定義（JUDGMENT_LABELSはConfig.gsで定義）
// ============================================

const JUDGMENT_WEIGHT = {
  'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 3
};

const JUDGMENT_GRADES = ['D', 'C', 'B', 'A'];

// ============================================
// ヘルパー関数（外部依存のフォールバック）
// ============================================

/**
 * ログ出力（外部定義がない場合のフォールバック）
 */
function logInfo_(message) {
  if (typeof logInfo === 'function') {
    logInfo(message);
  } else {
    Logger.log('[INFO] ' + message);
  }
}

function logError_(context, error) {
  if (typeof logError === 'function') {
    logError(context, error);
  } else {
    Logger.log('[ERROR] ' + context + ': ' + (error.message || error));
  }
}

/**
 * シート取得（外部定義がない場合のフォールバック）
 */
function getSheet_(sheetName) {
  if (typeof getSheet === 'function') {
    return getSheet(sheetName);
  }
  // DB_CONFIGが定義されている場合はそちらを使用
  if (typeof DB_CONFIG !== 'undefined' && DB_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(DB_CONFIG.SPREADSHEET_ID).getSheetByName(sheetName);
  }
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * 患者行検索（外部定義がない場合のフォールバック）
 */
function findPatientRow_(sheet, patientId, lastRow) {
  if (typeof findPatientRow === 'function') {
    return findPatientRow(sheet, patientId, lastRow);
  }
  // シンプルなフォールバック実装
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === patientId) return i + 2;
  }
  return 0;
}

// ============================================
// BMLコード → 判定基準キー マッピング（既存互換）
// ============================================

const CODE_TO_CRITERIA = {
  '0000301': 'WBC',
  '0000302': 'RBC',
  '0000303': 'HEMOGLOBIN',
  '0000304': 'HEMATOCRIT',
  '0000308': 'PLT',
  '0000401': 'TP',
  '0000407': 'URIC_ACID',
  '0000410': 'LDL_C',
  '0000413': 'CREATININE',
  '0000450': 'TC',
  '0000454': 'TG',
  '0000460': 'HDL_C',
  '0000472': 'T_BIL',
  '0000481': 'AST',
  '0000482': 'ALT',
  '0000484': 'GGT',
  '0000503': 'FBS',
  '0000658': 'CRP',
  '0002696': 'EGFR',
  '0003317': 'HBA1C'
};

const CODE_TO_CATEGORY = {
  '0000301': '血液学', '0000302': '血液学', '0000303': '血液学',
  '0000304': '血液学', '0000308': '血液学',
  '0000481': '肝機能', '0000482': '肝機能', '0000484': '肝機能',
  '0000460': '脂質', '0000410': '脂質', '0000454': '脂質',
  '0000503': '糖代謝', '0003317': '糖代謝',
  '0000413': '腎機能', '0002696': '腎機能', '0000407': '腎機能',
  '0000658': '炎症'
};

const GENDER_DEPENDENT_CODES = ['0000303', '0000413'];  // Hb, Cr

// ============================================
// 判定基準データ（マスタから読み込み or フォールバック）
// ============================================

/**
 * 判定基準をマスタシートから取得
 * シートがない場合はフォールバック定数を使用
 */
function getJudgmentCriteriaFromMaster() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('判定基準マスタ');

    if (!sheet) {
      logInfo_('判定基準マスタシートが見つかりません。フォールバック基準を使用します。');
      return FALLBACK_JUDGMENT_CRITERIA;
    }

    const data = sheet.getDataRange().getValues();
    const criteria = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;  // item_idが空の行はスキップ

      criteria.push({
        item_id: row[0],
        gender: row[1] || null,
        a_min: row[2] !== '' ? row[2] : null,
        a_max: row[3] !== '' ? row[3] : null,
        b_min: row[4] !== '' ? row[4] : null,
        b_max: row[5] !== '' ? row[5] : null,
        c_min: row[6] !== '' ? row[6] : null,
        c_max: row[7] !== '' ? row[7] : null,
        d_min: row[8] !== '' ? row[8] : null,
        d_max: row[9] !== '' ? row[9] : null,
        note: row[10] || ''
      });
    }

    return criteria;
  } catch (e) {
    logError_('getJudgmentCriteriaFromMaster', e);
    return FALLBACK_JUDGMENT_CRITERIA;
  }
}

// キャッシュ用変数
let _criteriaCache = null;
let _criteriaCacheTime = null;
const CACHE_DURATION_MS = 5 * 60 * 1000;  // 5分

/**
 * 判定基準を取得（キャッシュ付き）
 */
function getJudgmentCriteria() {
  const now = Date.now();
  if (_criteriaCache && _criteriaCacheTime && (now - _criteriaCacheTime < CACHE_DURATION_MS)) {
    return _criteriaCache;
  }

  _criteriaCache = getJudgmentCriteriaFromMaster();
  _criteriaCacheTime = now;
  return _criteriaCache;
}

/**
 * キャッシュをクリア
 */
function clearJudgmentCriteriaCache() {
  _criteriaCache = null;
  _criteriaCacheTime = null;
}

// ============================================
// フォールバック判定基準（マスタがない場合）
// ============================================

const FALLBACK_JUDGMENT_CRITERIA = [
  // 身体測定
  { item_id: 'BMI', a_min: 18.5, a_max: 24.9, b_min: 25.0, b_max: 27.9, c_min: 28.0, c_max: 29.9, d_min: 30.0, d_max: null },
  { item_id: 'WAIST_M', gender: 'M', a_min: 0, a_max: 84.9, b_min: 85.0, b_max: 89.9, c_min: 90.0, c_max: 99.9, d_min: 100.0, d_max: null },
  { item_id: 'WAIST_F', gender: 'F', a_min: 0, a_max: 89.9, b_min: 90.0, b_max: 94.9, c_min: 95.0, c_max: 99.9, d_min: 100.0, d_max: null },
  { item_id: 'BP_SYSTOLIC', a_min: 0, a_max: 129, b_min: 130, b_max: 139, c_min: 140, c_max: 159, d_min: 160, d_max: null },
  { item_id: 'BP_DIASTOLIC', a_min: 0, a_max: 84, b_min: 85, b_max: 89, c_min: 90, c_max: 99, d_min: 100, d_max: null },

  // 糖代謝
  { item_id: 'FBS', a_min: 0, a_max: 99, b_min: 100, b_max: 109, c_min: 110, c_max: 125, d_min: 126, d_max: null },
  { item_id: 'HBA1C', a_min: 0, a_max: 5.5, b_min: 5.6, b_max: 5.9, c_min: 6.0, c_max: 6.4, d_min: 6.5, d_max: null },

  // 脂質
  { item_id: 'HDL_C', a_min: 40, a_max: 999, b_min: null, b_max: null, c_min: 35, c_max: 39, d_min: 0, d_max: 34, note: '低値が異常' },
  { item_id: 'LDL_C', a_min: 60, a_max: 119, b_min: 120, b_max: 139, c_min: 140, c_max: 179, d_min: 180, d_max: null },
  { item_id: 'TG', a_min: 30, a_max: 149, b_min: 150, b_max: 299, c_min: 300, c_max: 499, d_min: 500, d_max: null },
  { item_id: 'TC', a_min: 140, a_max: 199, b_min: 200, b_max: 219, c_min: 220, c_max: 259, d_min: 260, d_max: null },

  // 肝機能
  { item_id: 'AST', a_min: 0, a_max: 30, b_min: 31, b_max: 35, c_min: 36, c_max: 50, d_min: 51, d_max: null },
  { item_id: 'ALT', a_min: 0, a_max: 30, b_min: 31, b_max: 40, c_min: 41, c_max: 50, d_min: 51, d_max: null },
  { item_id: 'GGT', a_min: 0, a_max: 50, b_min: 51, b_max: 80, c_min: 81, c_max: 100, d_min: 101, d_max: null },
  { item_id: 'T_BIL', a_min: 0.2, a_max: 1.2, b_min: 1.3, b_max: 2.0, c_min: 2.1, c_max: 3.0, d_min: 3.1, d_max: null },

  // 腎機能
  { item_id: 'CREATININE_M', gender: 'M', a_min: 0.1, a_max: 1.00, b_min: 1.01, b_max: 1.09, c_min: 1.10, c_max: 1.29, d_min: 1.30, d_max: null },
  { item_id: 'CREATININE_F', gender: 'F', a_min: 0.1, a_max: 0.70, b_min: 0.71, b_max: 0.79, c_min: 0.80, c_max: 0.99, d_min: 1.00, d_max: null },
  { item_id: 'EGFR', a_min: 60.0, a_max: 999, b_min: 50.0, b_max: 59.9, c_min: 45.0, c_max: 49.9, d_min: 0, d_max: 44.9, note: '低値が異常' },
  { item_id: 'URIC_ACID', a_min: 2.1, a_max: 7.0, b_min: 7.1, b_max: 7.9, c_min: 8.0, c_max: 8.9, d_min: 9.0, d_max: null },

  // 血液学
  { item_id: 'HEMOGLOBIN_M', gender: 'M', a_min: 13.1, a_max: 16.3, b_min: 16.4, b_max: 18.0, c_min: 12.1, c_max: 13.0, d_min: 0, d_max: 12.0, note: '低値・高値両方' },
  { item_id: 'HEMOGLOBIN_F', gender: 'F', a_min: 12.1, a_max: 14.5, b_min: 14.6, b_max: 16.0, c_min: 11.1, c_max: 12.0, d_min: 0, d_max: 11.0, note: '低値・高値両方' },
  { item_id: 'PLT', a_min: 14.5, a_max: 32.9, b_min: 12.3, b_max: 14.4, c_min: 10.0, c_max: 12.2, d_min: 0, d_max: 9.9, note: '低値が異常' },
  { item_id: 'WBC', a_min: 3.5, a_max: 8.9, b_min: 9.0, b_max: 9.9, c_min: 10.0, c_max: 12.0, d_min: 12.1, d_max: null },
  { item_id: 'RBC', a_min: 4.0, a_max: 5.5, b_min: 3.5, b_max: 3.9, c_min: 3.0, c_max: 3.4, d_min: 0, d_max: 2.9, note: '低値が異常' },
  { item_id: 'HEMATOCRIT', a_min: 38.0, a_max: 48.9, b_min: 35.0, b_max: 37.9, c_min: 30.0, c_max: 34.9, d_min: 0, d_max: 29.9, note: '低値が異常' },

  // その他
  { item_id: 'TP', a_min: 6.5, a_max: 7.9, b_min: 8.0, b_max: 8.3, c_min: 6.2, c_max: 6.4, d_min: 0, d_max: 6.1, note: '低値・高値両方' },
  { item_id: 'ALB', a_min: 3.9, a_max: 5.5, b_min: 3.7, b_max: 3.8, c_min: 3.0, c_max: 3.6, d_min: 0, d_max: 2.9, note: '低値が異常' },
  { item_id: 'CRP', a_min: 0, a_max: 0.30, b_min: 0.31, b_max: 0.99, c_min: 1.00, c_max: null, d_min: null, d_max: null }
];

// ============================================
// 単項目判定関数
// ============================================

/**
 * 検査項目IDと値から判定（A/B/C/D）を返す
 * @param {string} itemId - 検査項目ID
 * @param {number|string} value - 検査値
 * @param {string} gender - 性別（M/F）
 * @returns {string|null} 判定（A/B/C/D）
 */
function getJudgment(itemId, value, gender) {
  if (value === null || value === undefined || value === '') return null;

  const numValue = parseFloat(value);
  if (isNaN(numValue)) return null;

  // 性別依存項目の変換
  let lookupId = itemId;
  if (['WAIST', 'CREATININE', 'HEMOGLOBIN'].includes(itemId)) {
    lookupId = itemId + '_' + (gender === 'F' ? 'F' : 'M');
  }

  const criteria = getJudgmentCriteria().find(c => c.item_id === lookupId);
  if (!criteria) return null;

  return evaluateValue(numValue, criteria);
}

/**
 * 既存互換用エイリアス - judge()
 * 既存コードがjudge()を呼び出している場合のための互換関数
 * @param {string} itemKey - 検査項目キー
 * @param {number|string} value - 検査値
 * @param {string} gender - 性別（M/F）
 * @returns {string} 判定（A/B/C/D）または空文字
 */
function judge(itemKey, value, gender) {
  return getJudgment(itemKey, value, gender) || '';
}

/**
 * BMLコードから判定（既存互換）
 * @param {string} code - BMLコード
 * @param {string} valueStr - 検査値（文字列）
 * @param {string} flag - フラグ（H/L）
 * @param {string} gender - 性別（M/F）
 * @returns {string} 判定
 */
function judgeByCode(code, valueStr, flag, gender) {
  const numericValue = toNumber(valueStr);
  if (numericValue === null) return '';

  let criteriaKey = CODE_TO_CRITERIA[code];
  if (!criteriaKey) return judgeByFlag(flag);

  // 性別依存項目
  if (GENDER_DEPENDENT_CODES.includes(code)) {
    criteriaKey = criteriaKey + '_' + (gender === 'F' ? 'F' : 'M');
  }

  const result = getJudgment(criteriaKey, numericValue, gender);
  return result || judgeByFlag(flag);
}

/**
 * フラグから判定を推定（既存互換）
 */
function judgeByFlag(flag) {
  if (flag === 'H' || flag === 'L') return 'C';
  return 'A';
}

/**
 * 値を判定基準で評価
 */
function evaluateValue(value, criteria) {
  const isLowerBad = criteria.note && criteria.note.includes('低値が異常');
  const isBidirectional = criteria.note && criteria.note.includes('低値・高値両方');

  // A判定チェック
  if (criteria.a_min !== null && criteria.a_max !== null) {
    if (value >= criteria.a_min && value <= criteria.a_max) return 'A';
  }

  // 低値が異常のパターン
  if (isLowerBad) {
    if (criteria.d_min !== null && criteria.d_max !== null && value >= criteria.d_min && value <= criteria.d_max) return 'D';
    if (criteria.c_min !== null && criteria.c_max !== null && value >= criteria.c_min && value <= criteria.c_max) return 'C';
    if (criteria.b_min !== null && criteria.b_max !== null && value >= criteria.b_min && value <= criteria.b_max) return 'B';
    return 'A';
  }

  // 双方向パターン
  if (isBidirectional) {
    // 低値側D判定
    if (criteria.d_min !== null && criteria.d_max !== null && value >= criteria.d_min && value <= criteria.d_max) return 'D';
    // 低値側C判定
    if (criteria.c_min !== null && criteria.c_max !== null && value >= criteria.c_min && value <= criteria.c_max) return 'C';
    // 高値側B判定
    if (criteria.b_min !== null && criteria.b_max !== null && value >= criteria.b_min && value <= criteria.b_max) return 'B';
    return 'A';
  }

  // 通常パターン（高値が異常）
  if (criteria.b_min !== null && criteria.b_max !== null && value >= criteria.b_min && value <= criteria.b_max) return 'B';
  if (criteria.c_min !== null && criteria.c_max !== null && value >= criteria.c_min && value <= criteria.c_max) return 'C';
  if (criteria.d_min !== null && value >= criteria.d_min) return 'D';

  return 'A';
}

// ============================================
// 組合せ判定関数（新機能）
// ============================================

/**
 * FBS + HbA1c 組合せ判定
 */
function getGlucoseHbA1cJudgment(fbs, hba1c) {
  if ((fbs === null || fbs === '') && (hba1c === null || hba1c === '')) return null;

  const numFbs = parseFloat(fbs);
  const numHba1c = parseFloat(hba1c);

  if (isNaN(numFbs)) return getJudgment('HBA1C', hba1c, null);
  if (isNaN(numHba1c)) return getJudgment('FBS', fbs, null);

  // FBSランク
  const fbsRank = numFbs < 100 ? 0 : numFbs < 110 ? 1 : numFbs < 126 ? 2 : 3;
  // HbA1cランク
  const hba1cRank = numHba1c < 5.6 ? 0 : numHba1c < 6.0 ? 1 : numHba1c < 6.5 ? 2 : 3;

  const matrix = [
    ['A', 'A', 'B', 'C'],
    ['A', 'B', 'C', 'C'],
    ['B', 'C', 'C', 'D'],
    ['C', 'D', 'D', 'D']
  ];

  return matrix[hba1cRank][fbsRank];
}

/**
 * 血圧組合せ判定
 */
function getBloodPressureJudgment(systolic, diastolic) {
  const sysJ = getJudgment('BP_SYSTOLIC', systolic, null);
  const diaJ = getJudgment('BP_DIASTOLIC', diastolic, null);
  if (!sysJ) return diaJ;
  if (!diaJ) return sysJ;
  return getWorseJudgment(sysJ, diaJ);
}

// ============================================
// 総合判定関数
// ============================================

/**
 * 総合判定を計算（詳細版・新機能）
 * @param {Object} patientResults - 検査結果 { itemId: value, ... }
 * @param {string} gender - 性別
 * @returns {Object} { judgment, label, details, worstItems, summary }
 */
function calculateOverallJudgmentDetailed(patientResults, gender) {
  const results = [];
  let worstJudgment = 'A';

  // FBS+HbA1c組合せ
  if (patientResults.FBS !== undefined && patientResults.HBA1C !== undefined) {
    const gJ = getGlucoseHbA1cJudgment(patientResults.FBS, patientResults.HBA1C);
    if (gJ) {
      results.push({
        itemId: 'GLUCOSE_COMBINED',
        itemName: '糖代謝（組合せ）',
        value: `FBS:${patientResults.FBS}, HbA1c:${patientResults.HBA1C}`,
        judgment: gJ
      });
      if (isWorse(gJ, worstJudgment)) worstJudgment = gJ;
    }
  }

  // 血圧組合せ
  const bpSys = patientResults.BP_SYSTOLIC_1 || patientResults.BP_SYSTOLIC;
  const bpDia = patientResults.BP_DIASTOLIC_1 || patientResults.BP_DIASTOLIC;
  if (bpSys !== undefined && bpDia !== undefined) {
    const bpJ = getBloodPressureJudgment(bpSys, bpDia);
    if (bpJ) {
      results.push({
        itemId: 'BP_COMBINED',
        itemName: '血圧',
        value: `${bpSys}/${bpDia}`,
        judgment: bpJ
      });
      if (isWorse(bpJ, worstJudgment)) worstJudgment = bpJ;
    }
  }

  // その他項目
  const skipItems = ['FBS', 'HBA1C', 'BP_SYSTOLIC', 'BP_DIASTOLIC', 'BP_SYSTOLIC_1', 'BP_DIASTOLIC_1'];
  for (const [itemId, value] of Object.entries(patientResults)) {
    if (skipItems.includes(itemId)) continue;
    const j = getJudgment(itemId, value, gender);
    if (j) {
      const itemInfo = getExamItemById(itemId);
      results.push({
        itemId,
        itemName: itemInfo ? itemInfo.item_name : itemId,
        value,
        judgment: j
      });
      if (isWorse(j, worstJudgment)) worstJudgment = j;
    }
  }

  const worstItems = results.filter(r => r.judgment === worstJudgment);
  const counts = { A: 0, B: 0, C: 0, D: 0 };
  results.forEach(r => { if (counts[r.judgment] !== undefined) counts[r.judgment]++; });

  return {
    judgment: worstJudgment,
    label: JUDGMENT_LABELS[worstJudgment],
    details: results,
    worstItems,
    summary: { totalItems: results.length, counts, worstJudgment }
  };
}

/**
 * 総合判定を計算（シンプル版・既存互換）
 * @param {Array<string>} judgments - 判定配列
 * @returns {string} 総合判定
 */
function calculateOverallJudgment(judgments) {
  const valid = judgments.filter(j => j && JUDGMENT_GRADES.includes(j));
  if (valid.length === 0) return 'A';

  let worst = 'A';
  const order = { 'D': 0, 'C': 1, 'B': 2, 'A': 3 };
  for (const j of valid) {
    if (order[j] < order[worst]) worst = j;
  }
  return worst;
}

// ============================================
// 定性検査判定（新機能）
// ============================================

function getQualitativeJudgment(value) {
  if (!value) return null;
  const v = value.toString().trim();
  const rules = {
    '(-)': 'A', '-': 'A', '陰性': 'A',
    '(±)': 'B', '±': 'B', '疑陽性': 'B',
    '(+)': 'C', '+': 'C', '1+': 'C', '陽性': 'C',
    '(++)': 'C', '++': 'C', '2+': 'C',
    '(+++)': 'D', '+++': 'D', '3+': 'D'
  };
  return rules[v] || null;
}

// ============================================
// ユーティリティ関数
// ============================================

function getWorseJudgment(j1, j2) {
  return JUDGMENT_WEIGHT[j1] >= JUDGMENT_WEIGHT[j2] ? j1 : j2;
}

function isWorse(j1, j2) {
  return (JUDGMENT_WEIGHT[j1] || 0) > (JUDGMENT_WEIGHT[j2] || 0);
}

/**
 * 判定が悪いかどうか比較（既存互換エイリアス）
 */
function isWorseJudgment(j1, j2) {
  return isWorse(j1, j2);
}

function getJudgmentLabel(judgment) {
  return JUDGMENT_LABELS[judgment] || judgment;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  const num = parseFloat(value);
  return isNaN(num) ? null : num;
}

/**
 * 検査項目マスタから項目情報を取得
 * MasterData.gsのEXAM_ITEM_MASTER_DATAを参照
 * @param {string} itemId - 検査項目ID
 * @returns {Object|null} 項目情報 { item_id, item_name, category, ... }
 */
function getExamItemById(itemId) {
  // MasterData.gsで定義されているEXAM_ITEM_MASTER_DATAを参照
  if (typeof EXAM_ITEM_MASTER_DATA !== 'undefined' && Array.isArray(EXAM_ITEM_MASTER_DATA)) {
    return EXAM_ITEM_MASTER_DATA.find(i => i.item_id === itemId) || null;
  }

  // フォールバック: 簡易的な項目名マップ
  const fallbackNames = {
    'BMI': { item_id: 'BMI', item_name: 'BMI', category: '身体測定' },
    'WAIST': { item_id: 'WAIST', item_name: '腹囲', category: '身体測定' },
    'BP_SYSTOLIC': { item_id: 'BP_SYSTOLIC', item_name: '収縮期血圧', category: '循環器' },
    'BP_DIASTOLIC': { item_id: 'BP_DIASTOLIC', item_name: '拡張期血圧', category: '循環器' },
    'FBS': { item_id: 'FBS', item_name: '空腹時血糖', category: '糖代謝' },
    'HBA1C': { item_id: 'HBA1C', item_name: 'HbA1c', category: '糖代謝' },
    'HDL_C': { item_id: 'HDL_C', item_name: 'HDLコレステロール', category: '脂質' },
    'LDL_C': { item_id: 'LDL_C', item_name: 'LDLコレステロール', category: '脂質' },
    'TG': { item_id: 'TG', item_name: '中性脂肪', category: '脂質' },
    'TC': { item_id: 'TC', item_name: '総コレステロール', category: '脂質' },
    'AST': { item_id: 'AST', item_name: 'AST(GOT)', category: '肝機能' },
    'ALT': { item_id: 'ALT', item_name: 'ALT(GPT)', category: '肝機能' },
    'GGT': { item_id: 'GGT', item_name: 'γ-GTP', category: '肝機能' },
    'CREATININE': { item_id: 'CREATININE', item_name: 'クレアチニン', category: '腎機能' },
    'EGFR': { item_id: 'EGFR', item_name: 'eGFR', category: '腎機能' },
    'URIC_ACID': { item_id: 'URIC_ACID', item_name: '尿酸', category: '腎機能' },
    'HEMOGLOBIN': { item_id: 'HEMOGLOBIN', item_name: '血色素量(Hb)', category: '血液学' },
    'PLT': { item_id: 'PLT', item_name: '血小板数', category: '血液学' },
    'WBC': { item_id: 'WBC', item_name: '白血球数', category: '血液学' },
    'RBC': { item_id: 'RBC', item_name: '赤血球数', category: '血液学' },
    'HEMATOCRIT': { item_id: 'HEMATOCRIT', item_name: 'ヘマトクリット', category: '血液学' },
    'CRP': { item_id: 'CRP', item_name: 'CRP', category: '炎症' },
    'TP': { item_id: 'TP', item_name: '総蛋白', category: '蛋白' },
    'ALB': { item_id: 'ALB', item_name: 'アルブミン', category: '蛋白' },
    'T_BIL': { item_id: 'T_BIL', item_name: '総ビリルビン', category: '肝機能' }
  };

  return fallbackNames[itemId] || null;
}

// ============================================
// シート保存機能（既存互換）
// ============================================

/**
 * 患者の判定を処理して保存
 */
function processJudgments(patientId, patientData) {
  logInfo_(`判定処理開始: ${patientId}`);

  const patientInfo = patientData.patientInfo;
  const testResults = patientData.testResults;
  const gender = patientInfo.gender === '2' || patientInfo.gender === 'F' ? 'F' : 'M';

  const judgments = [];
  const judgmentsByCode = {};

  for (const result of testResults) {
    const j = judgeByCode(result.code, result.value, result.flag, gender);
    if (j) {
      judgments.push(j);
      judgmentsByCode[result.code] = j;
    }
  }

  // 判定結果を保存（既存互換）
  saveJudgmentResults(patientId, judgmentsByCode);

  const overall = calculateOverallJudgment(judgments);
  saveOverallJudgment(patientId, overall);

  logInfo_(`判定処理完了: ${patientId}, 総合=${overall}`);
}

/**
 * 判定結果を保存（既存互換）
 */
function saveJudgmentResults(patientId, judgmentsByCode) {
  // 既存の実装を維持（シート構造に依存）
  // 必要に応じて実装を追加
}

/**
 * 総合判定を保存
 */
function saveOverallJudgment(patientId, overall) {
  try {
    const sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS) ? CONFIG.SHEETS.PATIENT : '受診者';
    const sheet = getSheet_(sheetName);
    if (!sheet) {
      logError_('saveOverallJudgment', new Error('患者シートが見つかりません'));
      return;
    }

    const lastRow = sheet.getLastRow();
    const row = findPatientRow_(sheet, patientId, lastRow);

    if (row > 0) {
      sheet.getRange(row, 12).setValue(overall);  // L列: 総合判定
      sheet.getRange(row, 14).setValue(new Date());  // N列: 最終更新日時
      logInfo_(`総合判定を保存: 行${row}, 判定=${overall}`);
    } else {
      logInfo_(`警告: 患者ID ${patientId} の行が見つかりません`);
    }
  } catch (e) {
    logError_('saveOverallJudgment', e);
  }
}

/**
 * カテゴリ別判定を取得（既存互換）
 */
function getCategoryJudgments(patientId) {
  try {
    const bloodSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS) ? CONFIG.SHEETS.BLOOD_TEST : '血液検査';
    const patientSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS) ? CONFIG.SHEETS.PATIENT : '受診者';

    const sheet = getSheet_(bloodSheetName);
    if (!sheet) return {};

    const lastRow = sheet.getLastRow();
    const row = findPatientRow_(sheet, patientId, lastRow);
    if (row === 0) return {};

    // 性別取得
    const patientSheet = getSheet_(patientSheetName);
    let gender = 'M';
    if (patientSheet) {
      const patientRow = findPatientRow_(patientSheet, patientId, patientSheet.getLastRow());
      if (patientRow > 0) {
        const genderDisplay = patientSheet.getRange(patientRow, 6).getValue();
        gender = genderDisplay === '女' ? 'F' : 'M';
      }
    }

    const values = sheet.getRange(row, 2, 1, 27).getValues()[0];
    const columnToCode = {
      0: '0000301', 1: '0000302', 2: '0000303', 3: '0000304', 4: '0000308',
      11: '0000481', 12: '0000482', 13: '0000484',
      17: '0000413', 18: '0002696', 19: '0000407',
      21: '0000460', 22: '0000410', 23: '0000454',
      24: '0000503', 25: '0003317', 26: '0000658'
    };

    const categories = {};
    for (const [colIdx, code] of Object.entries(columnToCode)) {
      const value = values[parseInt(colIdx)];
      if (value === '' || value === null) continue;

      const j = judgeByCode(code, String(value), '', gender);
      const cat = CODE_TO_CATEGORY[code] || 'その他';

      if (!categories[cat]) categories[cat] = { items: [], worst: 'A' };
      categories[cat].items.push({ code, name: getItemNameByCode(code), value, judgment: j });
      if (j && isWorse(j, categories[cat].worst)) categories[cat].worst = j;
    }

    return categories;
  } catch (e) {
    logError_('getCategoryJudgments', e);
    return {};
  }
}

/**
 * BMLコードから項目名を取得（既存互換）
 */
function getItemName(code) {
  return getItemNameByCode(code);
}

/**
 * BMLコードから項目名を取得
 */
function getItemNameByCode(code) {
  const names = {
    '0000301': 'WBC', '0000302': 'RBC', '0000303': 'Hb', '0000304': 'Ht', '0000308': 'PLT',
    '0000401': 'TP', '0000407': 'UA', '0000410': 'LDL-C', '0000413': 'Cr',
    '0000450': 'TC', '0000454': 'TG', '0000460': 'HDL-C', '0000472': 'T-Bil',
    '0000481': 'AST', '0000482': 'ALT', '0000484': 'γ-GTP',
    '0000503': 'FBS', '0000658': 'CRP', '0002696': 'eGFR', '0003317': 'HbA1c'
  };
  return names[code] || code;
}

/**
 * 検査コードから判定基準キーを取得（既存互換）
 */
function getCriteriaKey(code, gender) {
  const baseKey = CODE_TO_CRITERIA[code];
  if (!baseKey) return null;

  // 性別依存項目の場合はサフィックスを付加
  if (GENDER_DEPENDENT_CODES.includes(code)) {
    return `${baseKey}_${gender}`;
  }

  return baseKey;
}

// ============================================
// API関数（新機能）
// ============================================

/**
 * フロントエンドから呼び出す判定API
 */
function apiGetJudgment(params) {
  try {
    if (params.itemId) {
      const j = getJudgment(params.itemId, params.value, params.gender);
      return { success: true, judgment: j, label: j ? getJudgmentLabel(j) : null };
    }
    if (params.results) {
      return { success: true, ...calculateOverallJudgmentDetailed(params.results, params.gender) };
    }
    if (params.fbs !== undefined || params.hba1c !== undefined) {
      const j = getGlucoseHbA1cJudgment(params.fbs, params.hba1c);
      return { success: true, judgment: j, label: j ? getJudgmentLabel(j) : null };
    }
    return { success: false, error: 'Invalid parameters' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * バッチ判定処理（新機能）
 */
function batchCalculateJudgments(patientsData) {
  return patientsData.map((p, i) => {
    try {
      return { index: i, success: true, ...calculateOverallJudgmentDetailed(p.results, p.gender) };
    } catch (e) {
      return { index: i, success: false, error: e.message };
    }
  });
}

// ============================================
// テスト関数
// ============================================

function testJudgmentEngine() {
  Logger.log('=== 統合版判定エンジンテスト ===');

  // 単項目テスト
  const tests = [
    { itemId: 'BMI', value: 22.0, gender: 'M', expected: 'A' },
    { itemId: 'HBA1C', value: 6.2, gender: 'M', expected: 'C' },
    { itemId: 'WAIST', value: 90, gender: 'M', expected: 'C' },
    { itemId: 'WAIST', value: 90, gender: 'F', expected: 'B' },
    { itemId: 'AST', value: 28, gender: 'M', expected: 'A' },
    { itemId: 'ALT', value: 45, gender: 'M', expected: 'C' },
  ];

  let pass = 0;
  for (const t of tests) {
    const r = getJudgment(t.itemId, t.value, t.gender);
    const ok = r === t.expected;
    pass += ok ? 1 : 0;
    Logger.log(`${t.itemId}(${t.gender}): ${t.value} -> ${r} (expected: ${t.expected}) ${ok ? 'PASS' : 'FAIL'}`);
  }

  // judge()エイリアステスト（既存互換）
  Logger.log('\n=== judge()エイリアステスト ===');
  const judgeResult = judge('BMI', 22.0, 'M');
  Logger.log(`judge('BMI', 22.0, 'M') -> ${judgeResult} (expected: A) ${judgeResult === 'A' ? 'PASS' : 'FAIL'}`);

  // 組合せ判定テスト
  Logger.log('\n=== 糖代謝組合せ ===');
  const gTests = [
    { fbs: 90, hba1c: 5.2, expected: 'A' },
    { fbs: 130, hba1c: 6.8, expected: 'D' },
    { fbs: 105, hba1c: 5.8, expected: 'B' },
  ];
  for (const t of gTests) {
    const r = getGlucoseHbA1cJudgment(t.fbs, t.hba1c);
    const ok = r === t.expected;
    Logger.log(`FBS:${t.fbs}, HbA1c:${t.hba1c} -> ${r} (expected: ${t.expected}) ${ok ? 'PASS' : 'FAIL'}`);
  }

  // 血圧組合せテスト
  Logger.log('\n=== 血圧組合せ ===');
  const bpTests = [
    { sys: 120, dia: 75, expected: 'A' },
    { sys: 145, dia: 92, expected: 'C' },
  ];
  for (const t of bpTests) {
    const r = getBloodPressureJudgment(t.sys, t.dia);
    const ok = r === t.expected;
    Logger.log(`BP:${t.sys}/${t.dia} -> ${r} (expected: ${t.expected}) ${ok ? 'PASS' : 'FAIL'}`);
  }

  // BMLコード互換テスト
  Logger.log('\n=== BMLコード互換 ===');
  const bmlTests = [
    { code: '0000481', value: '28', expected: 'A' },  // AST
    { code: '0003317', value: '6.8', expected: 'D' }, // HbA1c
    { code: '0000410', value: '135', expected: 'B' }, // LDL-C
  ];
  for (const t of bmlTests) {
    const r = judgeByCode(t.code, t.value, '', 'M');
    const ok = r === t.expected;
    Logger.log(`${t.code}: ${t.value} -> ${r} (expected: ${t.expected}) ${ok ? 'PASS' : 'FAIL'}`);
  }

  // 定性検査テスト
  Logger.log('\n=== 定性検査 ===');
  const qualTests = [
    { value: '(-)', expected: 'A' },
    { value: '(+)', expected: 'C' },
    { value: '+++', expected: 'D' },
  ];
  for (const t of qualTests) {
    const r = getQualitativeJudgment(t.value);
    const ok = r === t.expected;
    Logger.log(`定性: ${t.value} -> ${r} (expected: ${t.expected}) ${ok ? 'PASS' : 'FAIL'}`);
  }

  // getExamItemByIdテスト
  Logger.log('\n=== getExamItemById ===');
  const itemInfo = getExamItemById('BMI');
  Logger.log(`getExamItemById('BMI') -> ${JSON.stringify(itemInfo)}`);

  // 総合判定テスト
  Logger.log('\n=== 総合判定テスト ===');
  const overallTest = calculateOverallJudgment(['A', 'B', 'C', 'A']);
  Logger.log(`calculateOverallJudgment(['A', 'B', 'C', 'A']) -> ${overallTest} (expected: C)`);

  Logger.log(`\n=== テスト完了: ${pass}/${tests.length} passed ===`);
}

// ============================================
// エクスポート（グローバルスコープで利用可能）
// ============================================
// GASではモジュールシステムがないため、すべての関数はグローバルスコープで定義されています。
// 主要な関数:
// - getJudgment(itemId, value, gender): 単項目判定
// - judge(itemKey, value, gender): 既存互換エイリアス
// - judgeByCode(code, valueStr, flag, gender): BMLコード判定
// - getGlucoseHbA1cJudgment(fbs, hba1c): 糖代謝組合せ判定
// - getBloodPressureJudgment(systolic, diastolic): 血圧組合せ判定
// - calculateOverallJudgment(judgments): 総合判定（シンプル版）
// - calculateOverallJudgmentDetailed(patientResults, gender): 総合判定（詳細版）
// - getQualitativeJudgment(value): 定性検査判定
// - apiGetJudgment(params): API関数
// - batchCalculateJudgments(patientsData): バッチ処理
// - getExamItemById(itemId): 検査項目情報取得
// - processJudgments(patientId, patientData): 判定処理（既存互換）
// - getCategoryJudgments(patientId): カテゴリ別判定（既存互換）
