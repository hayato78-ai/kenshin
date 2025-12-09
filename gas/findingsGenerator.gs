/**
 * 所見自動生成モジュール
 * 判定結果からカテゴリ別所見を自動生成
 */

// ============================================
// 所見テンプレート（デフォルト）
// ============================================
const FINDINGS_TEMPLATES = {
  '循環器系': {
    'B': '血圧がやや高めです。減塩と適度な運動を心がけてください。',
    'C': '血圧が高めです。生活習慣の見直しと定期的な測定をお勧めします。',
    'D': '血圧が基準値を大きく超えています。医療機関での受診をお勧めします。'
  },
  '消化器系': {
    'B': '肝機能に軽度の異常があります。飲酒量の見直しをお勧めします。',
    'C': '肝機能に異常があります。精密検査をお勧めします。',
    'D': '肝機能に明らかな異常があります。早めの医療機関受診をお勧めします。'
  },
  '代謝系（糖）': {
    'B': '血糖値がやや高めです。糖質の摂取を控えめにしてください。',
    'C': '血糖値が高めです。糖尿病の精密検査をお勧めします。',
    'D': '血糖値が基準値を大きく超えています。糖尿病の治療が必要です。'
  },
  '代謝系（脂質）': {
    'B': '脂質代謝に軽度の異常があります。食事内容の見直しをお勧めします。',
    'C': '脂質代謝に異常があります。動物性脂肪を控え、運動習慣を取り入れてください。',
    'D': '脂質代謝に明らかな異常があります。医療機関での治療をお勧めします。'
  },
  '腎機能': {
    'B': '腎機能に軽度の異常があります。水分を十分に摂取してください。',
    'C': '腎機能に異常があります。定期的な経過観察をお勧めします。',
    'D': '腎機能に明らかな異常があります。専門医の受診をお勧めします。'
  },
  '血液系': {
    'B': '血液検査に軽度の異常があります。経過観察をお勧めします。',
    'C': '血液検査に異常があります。精密検査をお勧めします。',
    'D': '血液検査に明らかな異常があります。早めの受診をお勧めします。'
  },
  'その他': {
    'B_UA': '尿酸値がやや高めです。プリン体を含む食品を控えてください。',
    'C_UA': '尿酸値が高めです。痛風予防のため、食事療法をお勧めします。',
    'D_UA': '尿酸値が基準値を大きく超えています。医療機関での治療をお勧めします。',
    'B_BMI': '体重がやや増加傾向です。適度な運動と食事管理をお勧めします。',
    'C_BMI': '肥満傾向があります。食事と運動による体重管理をお勧めします。',
    'D_BMI': '肥満が認められます。専門家の指導のもと、減量をお勧めします。',
    'B': '軽度の異常が認められます。経過観察をお勧めします。',
    'C': '異常が認められます。精密検査をお勧めします。',
    'D': '明らかな異常が認められます。医療機関の受診をお勧めします。'
  }
};

// ============================================
// 所見生成関数
// ============================================

/**
 * 患者の所見を生成
 * @param {string} patientId - 受診ID
 * @returns {Object} 所見データ
 */
function generatePatientFindings(patientId) {
  // カテゴリ別判定を取得
  const categoryJudgments = getCategoryJudgments(patientId);

  // 各カテゴリの所見を生成
  const findings = {
    patientId: patientId,
    circulatory: '',      // 循環器系
    digestive: '',        // 消化器系
    metabolicSugar: '',   // 代謝系（糖）
    metabolicLipid: '',   // 代謝系（脂質）
    renal: '',            // 腎機能
    blood: '',            // 血液系
    other: '',            // その他
    combined: ''          // 総合所見
  };

  // 各カテゴリの所見を生成
  findings.circulatory = generateCategoryFinding('循環器系', categoryJudgments['循環器系']);
  findings.digestive = generateCategoryFinding('消化器系', categoryJudgments['消化器系']);
  findings.metabolicSugar = generateCategoryFinding('代謝系（糖）', categoryJudgments['代謝系（糖）']);
  findings.metabolicLipid = generateCategoryFinding('代謝系（脂質）', categoryJudgments['代謝系（脂質）']);
  findings.renal = generateCategoryFinding('腎機能', categoryJudgments['腎機能']);
  findings.blood = generateCategoryFinding('血液系', categoryJudgments['血液系']);
  findings.other = generateCategoryFinding('その他', categoryJudgments['その他']);

  // 総合所見を結合
  findings.combined = combineFindings(findings);

  // スプレッドシートに保存
  saveFindings(patientId, findings);

  return findings;
}

/**
 * カテゴリ別の所見を生成
 * @param {string} category - カテゴリ名
 * @param {Object} categoryData - カテゴリデータ { items: [...], worst: 'A/B/C/D' }
 * @returns {string} 所見文
 */
function generateCategoryFinding(category, categoryData) {
  if (!categoryData || !categoryData.worst || categoryData.worst === 'A') {
    return '';
  }

  const worst = categoryData.worst;
  const parts = [];

  // 1. 各項目の判定を出力
  if (categoryData.items && categoryData.items.length > 0) {
    const itemJudgments = [];
    for (const item of categoryData.items) {
      if (item.judgment && item.judgment !== 'A') {
        // 項目名: 値 (判定) の形式
        itemJudgments.push(`${item.name}: ${item.value} (${item.judgment})`);
      }
    }
    if (itemJudgments.length > 0) {
      parts.push(itemJudgments.join(', '));
    }
  }

  // 2. 所見コメントを追加
  let comment = '';

  // テンプレートから取得を試みる
  const template = getTemplateFromSheet(category, worst);
  if (template) {
    comment = template;
  } else {
    // デフォルトテンプレートから取得
    const defaultTemplates = FINDINGS_TEMPLATES[category];
    if (defaultTemplates && defaultTemplates[worst]) {
      comment = defaultTemplates[worst];
    } else if (category === 'その他' && categoryData.items) {
      // その他カテゴリで特定項目のテンプレート
      for (const item of categoryData.items) {
        if (item.judgment && item.judgment !== 'A') {
          const specificKey = `${item.judgment}_${item.name}`;
          if (defaultTemplates && defaultTemplates[specificKey]) {
            comment = defaultTemplates[specificKey];
            break;
          }
        }
      }
    }

    // 汎用テンプレート
    if (!comment) {
      comment = generateGenericFinding(worst);
    }
  }

  if (comment) {
    parts.push(comment);
  }

  return parts.join('\n');
}

/**
 * スプレッドシートからテンプレートを取得
 * @param {string} category - カテゴリ
 * @param {string} judgment - 判定
 * @returns {string|null} テンプレート文
 */
function getTemplateFromSheet(category, judgment) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.FINDINGS_TEMPLATE);
    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    for (const row of data) {
      if (row[0] === category && row[1] === judgment) {
        return row[3];  // D列: コメント
      }
    }
  } catch (e) {
    logError('getTemplateFromSheet', e);
  }

  return null;
}

/**
 * 汎用所見を生成
 * @param {string} judgment - 判定
 * @returns {string} 所見文
 */
function generateGenericFinding(judgment) {
  switch (judgment) {
    case 'B':
      return '軽度の異常が認められます。経過観察をお勧めします。';
    case 'C':
      return '異常が認められます。精密検査をお勧めします。';
    case 'D':
      return '明らかな異常が認められます。医療機関の受診をお勧めします。';
    default:
      return '';
  }
}

/**
 * 所見を結合して総合所見を生成
 * @param {Object} findings - 各カテゴリの所見
 * @returns {string} 総合所見
 */
function combineFindings(findings) {
  const sections = [];

  const categoryOrder = [
    { key: 'circulatory', label: '循環器系' },
    { key: 'digestive', label: '消化器系' },
    { key: 'metabolicSugar', label: '代謝系（糖）' },
    { key: 'metabolicLipid', label: '代謝系（脂質）' },
    { key: 'renal', label: '腎機能' },
    { key: 'blood', label: '血液系' },
    { key: 'other', label: 'その他' }
  ];

  for (const cat of categoryOrder) {
    const text = findings[cat.key];
    if (text && text.trim()) {
      sections.push(`【${cat.label}】\n${text}`);
    }
  }

  if (sections.length === 0) {
    return '今回の検査では特に問題は認められませんでした。';
  }

  return sections.join('\n\n');
}

/**
 * 所見をスプレッドシートに保存
 * @param {string} patientId - 受診ID
 * @param {Object} findings - 所見データ
 */
function saveFindings(patientId, findings) {
  const sheet = getSheet(CONFIG.SHEETS.FINDINGS);
  const lastRow = sheet.getLastRow();

  // 既存レコード検索
  let row = 0;
  const searchId = String(patientId).trim();  // 文字列に変換
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const cellId = String(ids[i][0]).trim();  // セル値も文字列に変換
      if (cellId === searchId) {
        row = i + 2;
        break;
      }
    }
  }

  const rowData = [
    patientId,                // A: 受診ID
    '',                       // B: 既往歴
    '',                       // C: 自覚症状
    '',                       // D: 他覚症状
    findings.circulatory,     // E: 所見_循環器系
    findings.digestive,       // F: 所見_消化器系
    findings.metabolicSugar,  // G: 所見_代謝系_糖
    findings.metabolicLipid,  // H: 所見_代謝系_脂質
    findings.renal,           // I: 所見_腎機能
    findings.blood,           // J: 所見_血液系
    findings.other,           // K: 所見_その他
    findings.combined,        // L: 総合所見
    '',                       // M: 心電図_今回
    '',                       // N: 心電図_前回
    '',                       // O: 腹部超音波_今回
    ''                        // P: 腹部超音波_前回
  ];

  if (row > 0) {
    // 既存レコード更新（所見列のみ）
    sheet.getRange(row, 5, 1, 8).setValues([[
      findings.circulatory,
      findings.digestive,
      findings.metabolicSugar,
      findings.metabolicLipid,
      findings.renal,
      findings.blood,
      findings.other,
      findings.combined
    ]]);
  } else {
    // 新規レコード追加
    sheet.appendRow(rowData);
  }

  logInfo(`所見を保存しました: ${patientId}`);
}

/**
 * 所見を取得
 * @param {string} patientId - 受診ID
 * @returns {Object|null} 所見データ
 */
function getFindings(patientId) {
  const sheet = getSheet(CONFIG.SHEETS.FINDINGS);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const searchId = String(patientId).trim();  // 文字列に変換

  for (const row of data) {
    const cellId = String(row[0]).trim();  // セル値も文字列に変換
    if (cellId === searchId) {
      return {
        patientId: row[0],
        history: row[1],
        subjective: row[2],
        objective: row[3],
        circulatory: row[4],
        digestive: row[5],
        metabolicSugar: row[6],
        metabolicLipid: row[7],
        renal: row[8],
        blood: row[9],
        other: row[10],
        combined: row[11],
        ecgCurrent: row[12],
        ecgPrevious: row[13],
        usCurrent: row[14],
        usPrevious: row[15]
      };
    }
  }

  return null;
}

/**
 * 所見を再生成
 * @param {string} patientId - 受診ID
 * @returns {Object} 所見データ
 */
function regenerateFindings(patientId) {
  return generatePatientFindings(patientId);
}

/**
 * 全患者の所見を再生成
 */
function regenerateAllFindings() {
  const sheet = getSheet(CONFIG.SHEETS.PATIENT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    logInfo('処理対象の患者がいません');
    return;
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let count = 0;
  for (const row of ids) {
    const patientId = row[0];
    if (patientId) {
      try {
        generatePatientFindings(patientId);
        count++;
      } catch (e) {
        logError('regenerateAllFindings', e);
      }
    }
  }

  logInfo(`${count}名の所見を再生成しました`);
}
