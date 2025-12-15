/**
 * 判定エンジンモジュール
 * 検査値から判定（A/B/C/D）を決定
 * 人間ドック学会2025年度基準に基づく
 */

// ============================================
// 判定基準マスタ
// ============================================
const JUDGMENT_CRITERIA = {
  // 肝機能
  'AST_GOT': {
    type: 'range',
    A: { min: 0, max: 30 },
    B: { min: 31, max: 35 },
    C: { min: 36, max: 50 },
    D: { min: 51, max: null }
  },
  'ALT_GPT': {
    type: 'range',
    A: { min: 0, max: 30 },
    B: { min: 31, max: 40 },
    C: { min: 41, max: 50 },
    D: { min: 51, max: null }
  },
  'GAMMA_GTP': {
    type: 'range',
    A: { min: 0, max: 50 },
    B: { min: 51, max: 80 },
    C: { min: 81, max: 100 },
    D: { min: 101, max: null }
  },

  // 脂質
  'HDL_CHOLESTEROL': {
    type: 'range',
    A: { min: 40, max: 100 },
    B: null,
    C: { min: 30, max: 39 },
    D: { min: null, max: 29 }
  },
  'LDL_CHOLESTEROL': {
    type: 'range',
    A: { min: 60, max: 119 },
    B: { min: 120, max: 139 },
    C: { min: 140, max: 179 },
    D: { min: 180, max: null }
  },
  'TRIGLYCERIDES': {
    type: 'range',
    A: { min: 30, max: 149 },
    B: { min: 150, max: 299 },
    C: { min: 300, max: 499 },
    D: { min: 500, max: null }
  },

  // 糖代謝
  'FASTING_GLUCOSE': {
    type: 'range',
    A: { min: null, max: 99 },
    B: { min: 100, max: 109 },
    C: { min: 110, max: 125 },
    D: { min: 126, max: null }
  },
  'HBA1C': {
    type: 'range',
    A: { min: null, max: 5.5 },
    B: { min: 5.6, max: 5.9 },
    C: { min: 6.0, max: 6.4 },
    D: { min: 6.5, max: null }
  },

  // 腎機能
  'EGFR': {
    type: 'range',
    A: { min: 60, max: null },
    B: { min: 45, max: 59.9 },
    C: { min: 30, max: 44.9 },
    D: { min: null, max: 29.9 }
  },
  'CREATININE_M': {
    type: 'range',
    gender: 'M',
    A: { min: 0.1, max: 1.0 },
    B: { min: 1.01, max: 1.2 },
    C: { min: 1.21, max: 1.5 },
    D: { min: 1.51, max: null }
  },
  'CREATININE_F': {
    type: 'range',
    gender: 'F',
    A: { min: 0.1, max: 0.7 },
    B: { min: 0.71, max: 0.9 },
    C: { min: 0.91, max: 1.1 },
    D: { min: 1.11, max: null }
  },

  // 血液
  'HEMOGLOBIN_M': {
    type: 'range',
    gender: 'M',
    A: { min: 13.1, max: 16.3 },
    B: { min: 12.1, max: 13.0, or: { min: 16.4, max: 18.0 } },
    C: { min: 11.0, max: 12.0, or: { min: 18.1, max: 19.0 } },
    D: { min: null, max: 10.9, or: { min: 19.1, max: null } }
  },
  'HEMOGLOBIN_F': {
    type: 'range',
    gender: 'F',
    A: { min: 12.1, max: 14.5 },
    B: { min: 11.1, max: 12.0, or: { min: 14.6, max: 16.0 } },
    C: { min: 10.0, max: 11.0, or: { min: 16.1, max: 17.0 } },
    D: { min: null, max: 9.9, or: { min: 17.1, max: null } }
  },
  'WBC': {
    type: 'range',
    A: { min: 3.1, max: 8.4 },
    B: { min: 2.5, max: 3.0, or: { min: 8.5, max: 9.9 } },
    C: { min: 2.0, max: 2.4, or: { min: 10.0, max: 12.0 } },
    D: { min: null, max: 1.9, or: { min: 12.1, max: null } }
  },
  'RBC': {
    type: 'range',
    A: { min: 400, max: 539 },
    B: { min: 360, max: 399, or: { min: 540, max: 579 } },
    C: { min: 320, max: 359, or: { min: 580, max: 619 } },
    D: { min: null, max: 319, or: { min: 620, max: null } }
  },
  'PLT': {
    type: 'range',
    A: { min: 15.0, max: 34.9 },
    B: { min: 13.0, max: 14.9, or: { min: 35.0, max: 39.9 } },
    C: { min: 10.0, max: 12.9, or: { min: 40.0, max: 44.9 } },
    D: { min: null, max: 9.9, or: { min: 45.0, max: null } }
  },

  // その他
  'URIC_ACID': {
    type: 'range',
    A: { min: 2.1, max: 7.0 },
    B: { min: 7.1, max: 8.0 },
    C: { min: 8.1, max: 9.0 },
    D: { min: 9.1, max: null }
  },
  'CRP': {
    type: 'range',
    A: { min: 0, max: 0.3 },
    B: { min: 0.31, max: 0.99 },
    C: { min: 1.0, max: 1.99 },
    D: { min: 2.0, max: null }
  },

  // 身体測定
  'BMI': {
    type: 'range',
    A: { min: 18.5, max: 24.9 },
    B: { min: 25.0, max: 27.9, or: { min: 17.0, max: 18.4 } },
    C: { min: 28.0, max: 29.9, or: { min: 16.0, max: 16.9 } },
    D: { min: 30.0, max: null, or: { min: null, max: 15.9 } }
  },
  'WAIST_M': {
    type: 'range',
    gender: 'M',
    A: { min: null, max: 84.9 },
    B: { min: 85.0, max: 89.9 },
    C: { min: 90.0, max: 99.9 },
    D: { min: 100.0, max: null }
  },
  'WAIST_F': {
    type: 'range',
    gender: 'F',
    A: { min: null, max: 89.9 },
    B: { min: 90.0, max: 94.9 },
    C: { min: 95.0, max: 99.9 },
    D: { min: 100.0, max: null }
  },
  'BP_SYSTOLIC': {
    type: 'range',
    A: { min: null, max: 129 },
    B: { min: 130, max: 139 },
    C: { min: 140, max: 159 },
    D: { min: 160, max: null }
  },
  'BP_DIASTOLIC': {
    type: 'range',
    A: { min: null, max: 84 },
    B: { min: 85, max: 89 },
    C: { min: 90, max: 99 },
    D: { min: 100, max: null }
  }
};

// ============================================
// 判定関数
// ============================================

/**
 * 検査値から判定を返す
 * @param {string} itemKey - 判定基準キー
 * @param {number} value - 検査値
 * @param {string} gender - 性別（M/F）
 * @returns {string} 判定結果（A/B/C/D）または空文字
 */
function judge(itemKey, value, gender) {
  if (!JUDGMENT_CRITERIA[itemKey]) {
    return '';
  }

  const spec = JUDGMENT_CRITERIA[itemKey];

  // 性別フィルタリング
  if (spec.gender && spec.gender !== gender) {
    return '';
  }

  // 各グレードをチェック
  for (const grade of CONFIG.JUDGMENT_GRADES) {
    const rule = spec[grade];
    if (rule === null || rule === undefined) {
      continue;
    }
    if (checkRange(value, rule)) {
      return grade;
    }
  }

  return '';
}

/**
 * BMLコードと値から判定を返す
 * @param {string} code - BML検査コード
 * @param {string} valueStr - 検査値（文字列）
 * @param {string} flag - フラグ（H/L/空）
 * @param {string} gender - 性別（M/F）
 * @returns {string} 判定結果（A/B/C/D）または空文字
 */
function judgeByCode(code, valueStr, flag, gender) {
  // 数値変換
  const numericValue = toNumber(valueStr);
  if (numericValue === null) {
    return '';
  }

  // 判定基準キーを取得
  const criteriaKey = getCriteriaKey(code, gender);

  if (criteriaKey) {
    const result = judge(criteriaKey, numericValue, gender);
    if (result) {
      return result;
    }
  }

  // 判定基準がない場合はフラグから推定
  return judgeByFlag(flag);
}

/**
 * 検査コードから判定基準キーを取得
 * @param {string} code - BML検査コード
 * @param {string} gender - 性別（M/F）
 * @returns {string|null} 判定基準キー
 */
function getCriteriaKey(code, gender) {
  const baseKey = CODE_TO_CRITERIA[code];

  if (!baseKey) {
    return null;
  }

  // 性別依存項目の場合はサフィックスを付加
  if (GENDER_DEPENDENT_CODES.includes(code)) {
    return `${baseKey}_${gender}`;
  }

  return baseKey;
}

/**
 * フラグから判定を推定
 * @param {string} flag - フラグ（H/L）
 * @returns {string} 判定結果
 */
function judgeByFlag(flag) {
  if (flag === 'H' || flag === 'L') {
    return 'C';  // デフォルト異常判定
  }
  return 'A';  // デフォルト正常判定
}

/**
 * 値が範囲内かチェック
 * @param {number} value - 検査値
 * @param {Object} rule - 範囲ルール
 * @returns {boolean} 範囲内ならtrue
 */
function checkRange(value, rule) {
  const minVal = rule.min;
  const maxVal = rule.max;

  let inMainRange = true;

  if (minVal !== null && minVal !== undefined && value < minVal) {
    inMainRange = false;
  }
  if (maxVal !== null && maxVal !== undefined && value > maxVal) {
    inMainRange = false;
  }

  if (inMainRange && (minVal !== null || maxVal !== null)) {
    return true;
  }

  // OR条件がある場合
  if (rule.or) {
    return checkRange(value, rule.or);
  }

  return false;
}

/**
 * 総合判定を計算
 * @param {Array<string>} judgments - 判定配列
 * @returns {string} 総合判定（A/B/C/D）
 */
function calculateOverallJudgment(judgments) {
  const validJudgments = judgments.filter(j => j && CONFIG.JUDGMENT_GRADES.includes(j));

  if (validJudgments.length === 0) {
    return 'A';
  }

  // 最も悪い判定を採用
  const gradeOrder = { 'D': 0, 'C': 1, 'B': 2, 'A': 3 };

  let worst = 'A';
  for (const j of validJudgments) {
    if (gradeOrder[j] < gradeOrder[worst]) {
      worst = j;
    }
  }

  return worst;
}

/**
 * 患者の全判定を処理
 * @param {string} patientId - 受診ID
 * @param {Object} patientData - 患者データ（CSV解析結果）
 */
function processJudgments(patientId, patientData) {
  logInfo(`判定処理開始: ${patientId}`);

  const patientInfo = patientData.patientInfo;
  const testResults = patientData.testResults;

  // 性別取得
  const genderCode = patientInfo.gender;
  const gender = GENDER_CODE_TO_INTERNAL[genderCode] || 'M';

  logInfo(`  検査結果数: ${testResults.length}, 性別: ${gender}`);

  const judgments = [];
  const judgmentsByCode = {};

  // 各検査項目の判定
  for (const result of testResults) {
    const judgment = judgeByCode(result.code, result.value, result.flag, gender);
    if (judgment) {
      judgments.push(judgment);
      judgmentsByCode[result.code] = judgment;
      logInfo(`  ${result.code}: ${result.value} → ${judgment}`);
    }
  }

  // 判定結果を保存
  saveJudgmentResults(patientId, judgmentsByCode);

  // 総合判定を計算・保存
  const overall = calculateOverallJudgment(judgments);
  logInfo(`  総合判定: ${overall} (判定数: ${judgments.length})`);
  saveOverallJudgment(patientId, overall);

  logInfo(`判定処理完了: ${patientId}`);
}

/**
 * 判定結果を保存
 * @param {string} patientId - 受診ID
 * @param {Object} judgmentsByCode - コード別判定
 */
function saveJudgmentResults(patientId, judgmentsByCode) {
  // 判定用列マッピング（判定列は値列の次）
  const sheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
  const lastRow = sheet.getLastRow();
  const row = findPatientRow(sheet, patientId, lastRow);

  if (row === 0) return;

  // 判定列マッピング（値列+1が判定列と仮定）
  // 実際のシート構造に合わせて調整が必要
  // ここでは別シートに判定を保存する方式にすることも可能
}

/**
 * 総合判定を保存
 * @param {string} patientId - 受診ID
 * @param {string} overall - 総合判定
 */
function saveOverallJudgment(patientId, overall) {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();
  logInfo(`  総合判定保存: patientId=${patientId}, lastRow=${lastRow}`);

  const row = findPatientRow(sheet, patientId, lastRow);
  logInfo(`  findPatientRow結果: row=${row}`);

  if (row > 0) {
    sheet.getRange(row, 12).setValue(overall);  // L列: 総合判定
    sheet.getRange(row, 14).setValue(new Date());  // N列: 最終更新日時
    logInfo(`  総合判定を保存: 行${row}, 判定=${overall}`);
  } else {
    logInfo(`  警告: 患者ID ${patientId} の行が見つかりません`);
  }
}

/**
 * カテゴリ別の判定を取得
 * @param {string} patientId - 受診ID
 * @returns {Object} カテゴリ別判定 { カテゴリ: { items: [...], worst: 'A/B/C/D' } }
 */
function getCategoryJudgments(patientId) {
  const sheet = getSheet(CONFIG.SHEETS.BLOOD_TEST);
  const lastRow = sheet.getLastRow();
  const row = findPatientRow(sheet, patientId, lastRow);

  if (row === 0) return {};

  // 性別取得
  const patientSheet = getSheet(CONFIG.SHEETS.PATIENT);
  const patientRow = findPatientRow(patientSheet, patientId, patientSheet.getLastRow());
  let gender = 'M';
  if (patientRow > 0) {
    const genderDisplay = patientSheet.getRange(patientRow, 6).getValue();
    gender = genderDisplay === '女' ? 'F' : 'M';
  }

  // 検査値を取得
  const values = sheet.getRange(row, 2, 1, 27).getValues()[0];

  // コード→列インデックス（0-indexed）
  const columnToCode = {
    0: '0000301',   // WBC
    1: '0000302',   // RBC
    2: '0000303',   // Hb
    3: '0000304',   // Ht
    4: '0000308',   // PLT
    11: '0000481',  // AST
    12: '0000482',  // ALT
    13: '0000484',  // γ-GTP
    17: '0000413',  // Cr
    18: '0002696',  // eGFR
    19: '0000407',  // UA
    21: '0000460',  // HDL
    22: '0000410',  // LDL
    23: '0000454',  // TG
    24: '0000503',  // FBS
    25: '0003317',  // HbA1c
    26: '0000658'   // CRP
  };

  // カテゴリ別に集計
  const categories = {};

  for (const [colIdx, code] of Object.entries(columnToCode)) {
    const value = values[parseInt(colIdx)];
    if (value === '' || value === null) continue;

    const judgment = judgeByCode(code, String(value), '', gender);
    const category = CODE_TO_CATEGORY[code] || 'その他';

    if (!categories[category]) {
      categories[category] = { items: [], worst: 'A' };
    }

    const itemName = getItemName(code);
    categories[category].items.push({
      code: code,
      name: itemName,
      value: value,
      judgment: judgment
    });

    // 最悪の判定を更新
    if (judgment && isWorseJudgment(judgment, categories[category].worst)) {
      categories[category].worst = judgment;
    }
  }

  return categories;
}

/**
 * 判定が悪いかどうか比較
 * @param {string} j1 - 判定1
 * @param {string} j2 - 判定2
 * @returns {boolean} j1がj2より悪ければtrue
 */
function isWorseJudgment(j1, j2) {
  const order = { 'D': 0, 'C': 1, 'B': 2, 'A': 3, '': 4 };
  return (order[j1] || 4) < (order[j2] || 4);
}

/**
 * コードから項目名を取得
 * @param {string} code - BMLコード
 * @returns {string} 項目名
 */
function getItemName(code) {
  const names = {
    '0000301': 'WBC',
    '0000302': 'RBC',
    '0000303': 'Hb',
    '0000304': 'Ht',
    '0000308': 'PLT',
    '0000401': 'TP',
    '0000407': 'UA',
    '0000410': 'LDL-C',
    '0000413': 'Cr',
    '0000450': 'TC',
    '0000454': 'TG',
    '0000460': 'HDL-C',
    '0000472': 'T-Bil',
    '0000481': 'AST',
    '0000482': 'ALT',
    '0000484': 'γ-GTP',
    '0000503': 'FBS',
    '0000658': 'CRP',
    '0002696': 'eGFR',
    '0003317': 'HbA1c'
  };
  return names[code] || code;
}
