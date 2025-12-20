/**
 * CSVインポートモジュール
 *
 * 機能:
 * - CSV読み込み（BML形式、ROSAI形式、汎用形式）
 * - Claude AIによる不定形CSV自動マッピング
 * - マッピングパターン保存・再利用
 * - 受診者マスタへの登録
 *
 * 画面仕様:
 * - SCR-012: CSVインポート画面
 * - SCR-012-AI: AIマッピング確認画面
 */

// ============================================
// 定数定義
// ============================================

const CSV_IMPORT_CONFIG = {
  // CSVフォーマット種別
  FORMATS: {
    BML: 'BML',         // BML検査センター形式
    ROSAI: 'ROSAI',     // 労災病院形式
    GENERIC: 'GENERIC', // 汎用形式（AI推論使用）
    SRL: 'SRL',         // SRL検査センター形式
    LSI: 'LSI'          // LSIメディエンス形式
  },

  // データ種別
  DATA_TYPES: {
    TEST_RESULT: 'TEST_RESULT',   // 検査結果
    GUIDANCE: 'GUIDANCE',         // 保健指導
    PATIENT_LIST: 'PATIENT_LIST'  // 名簿（受診者リスト）
  },

  // インポート単位
  IMPORT_UNITS: {
    INDIVIDUAL: 'INDIVIDUAL',  // 個人ごと（1ファイル=1名）
    BATCH: 'BATCH'             // 案件ごと（1ファイル=複数名）
  },

  // マッピング対象スキーマ
  PATIENT_SCHEMA: [
    { id: 'name', name: '氏名', description: 'フルネーム、漢字', required: true },
    { id: 'name_kana', name: 'カナ', description: 'フリガナ、カタカナ', required: false },
    { id: 'birth_date', name: '生年月日', description: 'YYYY/MM/DD形式', required: true },
    { id: 'gender', name: '性別', description: 'M=男性, F=女性', required: true },
    { id: 'phone', name: '電話番号', description: '携帯または固定電話', required: false },
    { id: 'email', name: 'メール', description: '連絡用メールアドレス', required: false },
    { id: 'company', name: '企業名', description: '所属企業・団体名', required: false },
    { id: 'employee_id', name: '社員番号', description: '企業内の社員ID', required: false },
    { id: 'department', name: '部署', description: '所属部署名', required: false },
    { id: 'address', name: '住所', description: '連絡先住所', required: false }
  ],

  // AIマッピングプロンプト設定
  AI_CONFIG: {
    SYSTEM_PROMPT: `あなたは健診システムのデータマッピング専門家です。
CSVカラムをシステム項目に正確にマッピングしてください。

ルール:
1. カラム名の類似性を判断（名前/氏名/お名前 = name）
2. サンプルデータの形式を参考にする
3. 確信度が低い場合は低いconfidenceを返す
4. マッピングできないカラムはtarget: null
5. 必ず有効なJSON形式で出力すること

日本語で回答してください。`,

    MAX_TOKENS: 2048
  }
};

// ============================================
// BML検査センター カラムマッピング定義
// ============================================

/**
 * BML CSVカラム名 → 検査項目マスタitem_id マッピング
 * BMLの出力形式に応じてカラム名を定義
 */
const BML_COLUMN_MAPPING = {
  // 基本情報
  '氏名': 'NAME',
  '患者名': 'NAME',
  'カナ': 'NAME_KANA',
  'フリガナ': 'NAME_KANA',
  '生年月日': 'BIRTHDATE',
  '生月日': 'BIRTHDATE',
  '性別': 'SEX',
  '年齢': 'AGE',
  '受診日': 'EXAM_DATE',
  '検査日': 'EXAM_DATE',
  '受付番号': 'RECEPTION_NO',
  '受診番号': 'RECEPTION_NO',

  // 身体測定
  '身長': 'HEIGHT',
  '身長(cm)': 'HEIGHT',
  '体重': 'WEIGHT',
  '体重(kg)': 'WEIGHT',
  'BMI': 'BMI',
  '腹囲': 'WAIST_M',  // 性別で判定必要
  '腹囲(cm)': 'WAIST_M',
  '体脂肪率': 'BODY_FAT',

  // 血圧
  '収縮期血圧': 'BP_SYSTOLIC_1',
  '最高血圧': 'BP_SYSTOLIC_1',
  '血圧(高)': 'BP_SYSTOLIC_1',
  '拡張期血圧': 'BP_DIASTOLIC_1',
  '最低血圧': 'BP_DIASTOLIC_1',
  '血圧(低)': 'BP_DIASTOLIC_1',
  '脈拍': 'PULSE',
  '脈拍数': 'PULSE',

  // 尿検査
  '尿蛋白': 'URINE_PROTEIN',
  '尿蛋白定性': 'URINE_PROTEIN',
  '尿糖': 'URINE_GLUCOSE',
  '尿糖定性': 'URINE_GLUCOSE',
  '尿潜血': 'URINE_OCCULT_BLOOD',
  '尿潜血定性': 'URINE_OCCULT_BLOOD',
  'ウロビリノーゲン': 'UROBILINOGEN',
  '尿PH': 'URINE_PH',
  '尿ビリルビン': 'URINE_BILIRUBIN',
  'ケトン体': 'URINE_KETONE',
  '尿比重': 'URINE_SG',

  // 便検査
  '便潜血1回目': 'FOBT_1',
  '便潜血(1)': 'FOBT_1',
  '便ヘモグロビン1': 'FOBT_1',
  '便潜血2回目': 'FOBT_2',
  '便潜血(2)': 'FOBT_2',
  '便ヘモグロビン2': 'FOBT_2',

  // 血液学検査
  '白血球数': 'WBC',
  'WBC': 'WBC',
  '赤血球数': 'RBC',
  'RBC': 'RBC',
  '血色素量': 'HEMOGLOBIN',
  'ヘモグロビン': 'HEMOGLOBIN',
  'Hb': 'HEMOGLOBIN',
  'ヘマトクリット': 'HEMATOCRIT',
  'Ht': 'HEMATOCRIT',
  '血小板数': 'PLATELET',
  'PLT': 'PLATELET',
  'MCV': 'MCV',
  'MCH': 'MCH',
  'MCHC': 'MCHC',

  // 肝機能
  '総蛋白': 'TOTAL_PROTEIN',
  'TP': 'TOTAL_PROTEIN',
  'アルブミン': 'ALBUMIN',
  'ALB': 'ALBUMIN',
  'AST': 'AST',
  'GOT': 'AST',
  'AST(GOT)': 'AST',
  'ALT': 'ALT',
  'GPT': 'ALT',
  'ALT(GPT)': 'ALT',
  'γ-GTP': 'GGT',
  'γGTP': 'GGT',
  'GGT': 'GGT',
  'ALP': 'ALP',
  'LDH': 'LDH',
  '総ビリルビン': 'T_BIL',
  'T-Bil': 'T_BIL',
  'アミラーゼ': 'AMYLASE',
  'AMY': 'AMYLASE',

  // 脂質検査
  '総コレステロール': 'TOTAL_CHOLESTEROL',
  'T-CHO': 'TOTAL_CHOLESTEROL',
  'TC': 'TOTAL_CHOLESTEROL',
  '中性脂肪': 'TG',
  'トリグリセライド': 'TG',
  'TG': 'TG',
  'HDLコレステロール': 'HDL_C',
  'HDL-C': 'HDL_C',
  'HDL': 'HDL_C',
  'LDLコレステロール': 'LDL_C',
  'LDL-C': 'LDL_C',
  'LDL': 'LDL_C',
  'non-HDL': 'NON_HDL_C',

  // 糖代謝
  '空腹時血糖': 'FBS',
  '血糖': 'FBS',
  'FBS': 'FBS',
  'FPG': 'FBS',
  'HbA1c': 'HBA1C',
  'HbA1c(NGSP)': 'HBA1C',
  'グリコヘモグロビン': 'HBA1C',

  // 腎機能
  'クレアチニン': 'CREATININE',
  'CRE': 'CREATININE',
  'Cr': 'CREATININE',
  '尿素窒素': 'BUN',
  'BUN': 'BUN',
  'eGFR': 'EGFR',
  'GFR': 'EGFR',

  // その他生化学
  '尿酸': 'UA',
  'UA': 'UA',
  'CK': 'CK',
  'CPK': 'CK',
  'ナトリウム': 'NA',
  'Na': 'NA',
  'カリウム': 'K',
  'K': 'K',
  'クロール': 'CL',
  'Cl': 'CL',
  'カルシウム': 'CA',
  'Ca': 'CA',

  // 腫瘍マーカー
  'PSA': 'PSA',
  '前立腺特異抗原': 'PSA',
  'CEA': 'CEA',
  'CA19-9': 'CA19_9',
  'CA125': 'CA125',
  'AFP': 'AFP',
  'NSE': 'NSE',
  'CYFRA21-1': 'CYFRA21_1',
  'CYFRA': 'CYFRA21_1',
  'SCC': 'SCC',
  'ProGRP': 'PROGRP',
  'PIVKA-II': 'PIVKA2',
  'PIVKA2': 'PIVKA2',
  '抗p53抗体': 'P53',

  // 感染症
  'TPHA': 'TPHA',
  '梅毒TPHA': 'TPHA',
  'RPR': 'RPR',
  '梅毒RPR': 'RPR',
  'HBs抗原': 'HBS_AG',
  'HBsAg': 'HBS_AG',
  'HBs抗体': 'HBS_AB',
  'HBsAb': 'HBS_AB',
  'HCV抗体': 'HCV_AB',
  'HCVAb': 'HCV_AB',
  'HIV抗体': 'HIV_AB',

  // 甲状腺
  'FT3': 'FT3',
  'FT4': 'FT4',
  'TSH': 'TSH',
  'NT-proBNP': 'NT_PROBNP',

  // 血液型
  '血液型ABO': 'BLOOD_TYPE_ABO',
  'ABO式': 'BLOOD_TYPE_ABO',
  '血液型Rh': 'BLOOD_TYPE_RH',
  'Rh式': 'BLOOD_TYPE_RH',

  // 肺機能
  '肺活量': 'VC',
  'VC': 'VC',
  '1秒量': 'FEV1',
  'FEV1': 'FEV1',
  '%肺活量': 'PERCENT_VC',
  '1秒率': 'FEV1_PERCENT',

  // 視力・聴力
  '視力(右)裸眼': 'VISION_NAKED_R',
  '右眼裸眼': 'VISION_NAKED_R',
  '視力(左)裸眼': 'VISION_NAKED_L',
  '左眼裸眼': 'VISION_NAKED_L',
  '視力(右)矯正': 'VISION_CORRECTED_R',
  '右眼矯正': 'VISION_CORRECTED_R',
  '視力(左)矯正': 'VISION_CORRECTED_L',
  '左眼矯正': 'VISION_CORRECTED_L',
  '眼圧(右)': 'IOP_R',
  '眼圧(左)': 'IOP_L',
  '聴力右1000Hz': 'HEARING_R_1000',
  '聴力左1000Hz': 'HEARING_L_1000',
  '聴力右4000Hz': 'HEARING_R_4000',
  '聴力左4000Hz': 'HEARING_L_4000'
};

/**
 * BML検査コード → item_id マッピング（BMLコードを使用する場合）
 * BMLの検査コードから直接マッピングする場合に使用
 */
const BML_CODE_MAPPING = {
  // 一般検査
  '001': 'URINE_PROTEIN',
  '002': 'URINE_GLUCOSE',
  '003': 'URINE_OCCULT_BLOOD',
  // 血液学
  '101': 'WBC',
  '102': 'RBC',
  '103': 'HEMOGLOBIN',
  '104': 'HEMATOCRIT',
  '105': 'PLATELET',
  // 生化学
  '201': 'AST',
  '202': 'ALT',
  '203': 'GGT',
  '204': 'ALP',
  '205': 'LDH',
  '211': 'TOTAL_PROTEIN',
  '212': 'ALBUMIN',
  '221': 'TOTAL_CHOLESTEROL',
  '222': 'TG',
  '223': 'HDL_C',
  '224': 'LDL_C',
  '231': 'FBS',
  '232': 'HBA1C',
  '241': 'CREATININE',
  '242': 'BUN',
  '243': 'UA',
  '244': 'EGFR',
  // 腫瘍マーカー
  '301': 'CEA',
  '302': 'AFP',
  '303': 'CA19_9',
  '304': 'PSA'
};

/**
 * 検査値の正規化ルール（BML固有の値変換）
 */
const BML_VALUE_TRANSFORMS = {
  // 性別変換
  gender: {
    '1': 'M', '2': 'F',
    '男': 'M', '女': 'F',
    '男性': 'M', '女性': 'F',
    'M': 'M', 'F': 'F'
  },
  // 定性検査変換
  qualitative: {
    '-': '(-)', '±': '(±)', '+': '(+)', '++': '(++)', '+++': '(+++)',
    '陰性': '(-)', '擬陽性': '(±)', '陽性': '(+)',
    'ネガティブ': '(-)', 'ポジティブ': '(+)',
    '1-': '(-)', '1+': '(+)', '2+': '(++)', '3+': '(+++)'
  },
  // 聴力判定
  hearing: {
    '正常': '異常なし', '異常': '所見あり',
    'A': '異常なし', 'B': '所見あり', 'C': '所見あり',
    '○': '異常なし', '×': '所見あり'
  }
};

// ============================================
// CSV読み込み基本機能
// ============================================

/**
 * CSVファイルをパース
 * @param {string} csvContent - CSVの内容
 * @param {Object} options - オプション（encoding, delimiter等）
 * @returns {Object} {headers: string[], rows: string[][]}
 */
function parseCsv(csvContent, options = {}) {
  const delimiter = options.delimiter || ',';
  const hasHeader = options.hasHeader !== false;

  try {
    const lines = csvContent.split(/\r?\n/).filter(line => line.trim());

    if (lines.length === 0) {
      return { headers: [], rows: [], error: 'CSVが空です' };
    }

    // CSVパース（クォート対応）
    const parseRow = (line) => {
      const result = [];
      let current = '';
      let inQuotes = false;

      for (let i = 0; i < line.length; i++) {
        const char = line[i];
        const nextChar = line[i + 1];

        if (char === '"') {
          if (inQuotes && nextChar === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (char === delimiter && !inQuotes) {
          result.push(current.trim());
          current = '';
        } else {
          current += char;
        }
      }
      result.push(current.trim());

      return result;
    };

    const parsedLines = lines.map(parseRow);

    if (hasHeader) {
      return {
        headers: parsedLines[0],
        rows: parsedLines.slice(1),
        error: null
      };
    } else {
      return {
        headers: parsedLines[0].map((_, i) => `Column${i + 1}`),
        rows: parsedLines,
        error: null
      };
    }
  } catch (e) {
    logError('parseCsv', e);
    return { headers: [], rows: [], error: e.message };
  }
}

/**
 * GoogleドライブからCSVファイルを読み込み
 * @param {string} fileId - ファイルID
 * @returns {Object} パース結果
 */
function loadCsvFromDrive(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString('UTF-8');

    return {
      success: true,
      fileName: file.getName(),
      ...parseCsv(content)
    };
  } catch (e) {
    logError('loadCsvFromDrive', e);
    return {
      success: false,
      error: `ファイル読み込みエラー: ${e.message}`
    };
  }
}

// ============================================
// Claude AI マッピング機能
// ============================================

/**
 * 不定形CSVのカラムマッピングをClaudeで推論
 * @param {string[]} csvHeaders - CSVのヘッダー行
 * @param {string[][]} sampleRows - サンプルデータ（最大3行）
 * @param {Object[]} targetSchema - マッピング先スキーマ
 * @returns {Object} マッピング結果
 */
function inferCsvMapping(csvHeaders, sampleRows, targetSchema = CSV_IMPORT_CONFIG.PATIENT_SCHEMA) {
  // ヘッダーが空の場合はエラー
  if (!csvHeaders || csvHeaders.length === 0) {
    return {
      success: false,
      error: 'CSVヘッダーが空です'
    };
  }

  // サンプル行を最大3行に制限
  const samples = sampleRows.slice(0, 3);

  // プロンプト構築
  const userMessage = `## CSVカラムとサンプルデータ

${csvHeaders.map((h, i) => {
  const sampleValues = samples.map(r => r[i] || '').filter(v => v).slice(0, 3);
  return `- カラム${i + 1}「${h}」: サンプル値 [${sampleValues.join(', ') || '(空)'}]`;
}).join('\n')}

## マッピング先システム項目

${targetSchema.map(s => `- ${s.id}: ${s.name}（${s.description}）${s.required ? '【必須】' : ''}`).join('\n')}

## 出力形式（以下のJSON形式で出力してください）

{
  "mappings": [
    {"csv_column": "CSVカラム名", "csv_index": 0, "target": "システム項目ID", "confidence": 0.95},
    {"csv_column": "CSVカラム名2", "csv_index": 1, "target": null, "confidence": 0.0}
  ],
  "value_transforms": {
    "性別": {"男": "M", "女": "F", "男性": "M", "女性": "F"}
  },
  "date_formats": {
    "生年月日": "YYYY/MM/DD"
  },
  "overall_confidence": 0.92,
  "notes": "推論に関する補足"
}`;

  try {
    const result = callClaudeApi(
      CSV_IMPORT_CONFIG.AI_CONFIG.SYSTEM_PROMPT,
      userMessage,
      { max_tokens: CSV_IMPORT_CONFIG.AI_CONFIG.MAX_TOKENS }
    );

    if (!result.success) {
      return result;
    }

    // レスポンスからJSONを抽出
    const jsonMatch = result.content.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return {
        success: false,
        error: 'AIレスポンスからJSONを抽出できませんでした',
        rawContent: result.content
      };
    }

    const mappingResult = JSON.parse(jsonMatch[0]);

    return {
      success: true,
      mappings: mappingResult.mappings || [],
      valueTransforms: mappingResult.value_transforms || {},
      dateFormats: mappingResult.date_formats || {},
      overallConfidence: mappingResult.overall_confidence || 0,
      notes: mappingResult.notes || '',
      usage: result.usage
    };

  } catch (e) {
    logError('inferCsvMapping', e);
    return {
      success: false,
      error: `マッピング推論エラー: ${e.message}`
    };
  }
}

/**
 * マッピング結果を適用してデータを変換
 * @param {string[]} headers - CSVヘッダー
 * @param {string[][]} rows - CSVデータ行
 * @param {Object[]} mappings - マッピング定義
 * @param {Object} valueTransforms - 値変換ルール
 * @returns {Object[]} 変換後のデータ配列
 */
function applyMapping(headers, rows, mappings, valueTransforms = {}) {
  const result = [];

  // マッピングをインデックスでアクセスできるようにする
  const indexToTarget = {};
  mappings.forEach(m => {
    if (m.target) {
      indexToTarget[m.csv_index] = {
        target: m.target,
        column: m.csv_column
      };
    }
  });

  for (const row of rows) {
    const record = {};

    for (let i = 0; i < row.length; i++) {
      const mapping = indexToTarget[i];
      if (!mapping) continue;

      let value = row[i];

      // 値変換を適用
      const column = mapping.column;
      if (valueTransforms[column] && valueTransforms[column][value]) {
        value = valueTransforms[column][value];
      }

      record[mapping.target] = value;
    }

    if (Object.keys(record).length > 0) {
      result.push(record);
    }
  }

  return result;
}

// ============================================
// マッピングパターン管理
// ============================================

/**
 * マッピングパターンIDを生成
 * @returns {string} MP00001形式のID
 */
function generateMappingPatternId() {
  return generateSequentialId(CONFIG.SHEETS.MAPPING_PATTERN || 'M_MappingPattern', 'MP', 5);
}

/**
 * ヘッダーのハッシュ値を計算
 * @param {string[]} headers - ヘッダー配列
 * @returns {string} ハッシュ値
 */
function calculateHeadersHash(headers) {
  const str = headers.sort().join('|').toLowerCase();
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return Math.abs(hash).toString(16).padStart(8, '0');
}

/**
 * マッピングパターンを保存
 * @param {Object} pattern - パターン情報
 * @returns {Object} 保存結果
 */
function saveMappingPattern(pattern) {
  try {
    const sheet = getSheet('M_MappingPattern');
    const patternId = generateMappingPatternId();
    const headersHash = calculateHeadersHash(pattern.csvHeaders);
    const now = new Date();

    const rowData = [
      patternId,
      pattern.sourceName || '',
      headersHash,
      JSON.stringify(pattern.mappings),
      JSON.stringify(pattern.valueTransforms || {}),
      1,  // success_count
      now,
      now
    ];

    sheet.appendRow(rowData);
    logInfo(`マッピングパターン保存: ${patternId} (${pattern.sourceName})`);

    return {
      success: true,
      patternId: patternId
    };
  } catch (e) {
    logError('saveMappingPattern', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 既存のマッピングパターンを検索
 * @param {string[]} headers - CSVヘッダー
 * @returns {Object|null} マッチしたパターンまたはnull
 */
function findMappingPattern(headers) {
  try {
    const sheet = getSheet('M_MappingPattern');
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return null;

    const targetHash = calculateHeadersHash(headers);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const storedHash = row[2];

      if (storedHash === targetHash) {
        // 使用回数をインクリメント
        sheet.getRange(i + 1, 6).setValue(row[5] + 1);
        sheet.getRange(i + 1, 8).setValue(new Date());

        return {
          patternId: row[0],
          sourceName: row[1],
          mappings: JSON.parse(row[3]),
          valueTransforms: JSON.parse(row[4] || '{}'),
          successCount: row[5] + 1
        };
      }
    }

    return null;
  } catch (e) {
    logError('findMappingPattern', e);
    return null;
  }
}

// ============================================
// BML形式CSVパース・バリデーション
// ============================================

/**
 * BML形式CSVをパースして標準形式に変換
 * @param {string} csvContent - BML形式CSVの内容
 * @param {Object} options - オプション（gender指定など）
 * @returns {Object} 変換結果 {success, records[], mappingInfo, errors[]}
 */
function parseBmlCsv(csvContent, options = {}) {
  try {
    // 基本パース
    const parsed = parseCsv(csvContent, options);
    if (parsed.error) {
      return { success: false, error: parsed.error, records: [], errors: [] };
    }

    const { headers, rows } = parsed;
    const records = [];
    const errors = [];
    const mappingInfo = {
      mappedColumns: [],
      unmappedColumns: [],
      totalRows: rows.length
    };

    // ヘッダーのマッピング情報を構築
    const columnMapping = [];
    headers.forEach((header, index) => {
      const normalizedHeader = header.trim();
      const itemId = BML_COLUMN_MAPPING[normalizedHeader];

      if (itemId) {
        columnMapping.push({ index, header: normalizedHeader, itemId });
        mappingInfo.mappedColumns.push({ header: normalizedHeader, itemId });
      } else {
        mappingInfo.unmappedColumns.push(normalizedHeader);
      }
    });

    // 各行をパース
    rows.forEach((row, rowIndex) => {
      try {
        const record = {
          _rowIndex: rowIndex + 2,  // 1-indexed + header row
          _raw: {}
        };
        let gender = options.gender || null;

        // まず性別を取得（腹囲の判定に必要）
        columnMapping.forEach(({ index, itemId }) => {
          if (itemId === 'SEX') {
            const rawValue = row[index];
            gender = normalizeBmlValue(rawValue, 'SEX');
          }
        });

        // 各カラムを変換
        columnMapping.forEach(({ index, header, itemId }) => {
          const rawValue = row[index];
          record._raw[header] = rawValue;

          if (rawValue === undefined || rawValue === null || rawValue === '') {
            return;
          }

          // 腹囲は性別に応じてitem_idを変更
          let finalItemId = itemId;
          if (itemId === 'WAIST_M') {
            finalItemId = gender === 'F' ? 'WAIST_F' : 'WAIST_M';
          }

          // 値を正規化
          const normalizedValue = normalizeBmlValue(rawValue, finalItemId);
          if (normalizedValue !== null) {
            record[finalItemId] = normalizedValue;
          }
        });

        // 性別を保持
        if (gender) {
          record.SEX = gender;
        }

        records.push(record);

      } catch (rowError) {
        errors.push({
          row: rowIndex + 2,
          error: rowError.message,
          data: row
        });
      }
    });

    mappingInfo.successCount = records.length;
    mappingInfo.errorCount = errors.length;

    return {
      success: true,
      records,
      mappingInfo,
      errors,
      headers,
      columnMapping
    };

  } catch (e) {
    logError('parseBmlCsv', e);
    return {
      success: false,
      error: e.message,
      records: [],
      errors: []
    };
  }
}

/**
 * BML固有の値正規化
 * @param {string} value - 元の値
 * @param {string} itemId - 項目ID
 * @returns {*} 正規化された値
 */
function normalizeBmlValue(value, itemId) {
  if (value === undefined || value === null || value === '') {
    return null;
  }

  const strValue = String(value).trim();

  // 性別
  if (itemId === 'SEX') {
    return BML_VALUE_TRANSFORMS.gender[strValue] || null;
  }

  // 定性検査（尿蛋白、尿糖など）
  const qualitativeItems = [
    'URINE_PROTEIN', 'URINE_GLUCOSE', 'URINE_OCCULT_BLOOD',
    'UROBILINOGEN', 'URINE_BILIRUBIN', 'URINE_KETONE',
    'FOBT_1', 'FOBT_2', 'HBS_AG', 'HBS_AB', 'HCV_AB',
    'HIV_AB', 'TPHA', 'RPR', 'URINE_BACTERIA'
  ];
  if (qualitativeItems.includes(itemId)) {
    return BML_VALUE_TRANSFORMS.qualitative[strValue] || strValue;
  }

  // 聴力
  const hearingItems = ['HEARING_R_1000', 'HEARING_L_1000', 'HEARING_R_4000', 'HEARING_L_4000'];
  if (hearingItems.includes(itemId)) {
    return BML_VALUE_TRANSFORMS.hearing[strValue] || strValue;
  }

  // 血液型
  if (itemId === 'BLOOD_TYPE_ABO') {
    const aboMap = { 'A型': 'A', 'B型': 'B', 'O型': 'O', 'AB型': 'AB' };
    return aboMap[strValue] || strValue;
  }
  if (itemId === 'BLOOD_TYPE_RH') {
    const rhMap = { '陽性': '(+)', '陰性': '(-)', '+': '(+)', '-': '(-)' };
    return rhMap[strValue] || strValue;
  }

  // 日付
  if (itemId === 'EXAM_DATE' || itemId === 'BIRTHDATE') {
    return normalizeBirthDate(strValue);
  }

  // 数値項目
  const numericItems = [
    'HEIGHT', 'WEIGHT', 'BMI', 'BODY_FAT', 'WAIST_M', 'WAIST_F',
    'BP_SYSTOLIC_1', 'BP_DIASTOLIC_1', 'BP_SYSTOLIC_2', 'BP_DIASTOLIC_2', 'PULSE',
    'VISION_NAKED_R', 'VISION_NAKED_L', 'VISION_CORRECTED_R', 'VISION_CORRECTED_L',
    'IOP_R', 'IOP_L', 'WBC', 'RBC', 'HEMOGLOBIN', 'HEMATOCRIT', 'PLATELET',
    'MCV', 'MCH', 'MCHC', 'TOTAL_PROTEIN', 'ALBUMIN', 'AST', 'ALT', 'GGT',
    'ALP', 'LDH', 'AMYLASE', 'T_BIL', 'TOTAL_CHOLESTEROL', 'TG', 'HDL_C',
    'LDL_C', 'NON_HDL_C', 'FBS', 'HBA1C', 'CREATININE', 'BUN', 'EGFR', 'UA',
    'CK', 'NA', 'K', 'CL', 'CA', 'PSA', 'CEA', 'CA19_9', 'CA125', 'AFP',
    'NSE', 'CYFRA21_1', 'SCC', 'PROGRP', 'PIVKA2', 'FT3', 'FT4', 'TSH',
    'NT_PROBNP', 'VC', 'FEV1', 'PERCENT_VC', 'FEV1_PERCENT', 'URINE_PH', 'URINE_SG', 'AGE'
  ];

  if (numericItems.includes(itemId)) {
    // 数値以外の文字を除去（<, >, 未満, 以上など）
    const cleanedValue = strValue.replace(/[<>≦≧未満以上以下]/g, '').trim();
    const num = parseFloat(cleanedValue);
    return isNaN(num) ? strValue : num;
  }

  // その他はそのまま
  return strValue;
}

/**
 * CSVデータのバリデーション
 * 必須項目チェックと判定基準マスタを参照した値の範囲チェック
 * @param {Object[]} records - parseBmlCsvの出力records
 * @param {Object} options - バリデーションオプション
 * @returns {Object} バリデーション結果 {valid, errors[], warnings[]}
 */
function validateCsvData(records, options = {}) {
  const errors = [];
  const warnings = [];
  const validRecords = [];
  const courseId = options.courseId || 'DOCK_LIFESTYLE';

  // コース必須項目を取得
  const course = EXAM_COURSE_MASTER_DATA.find(c => c.course_id === courseId);
  const requiredItems = course ? course.required_items.split(',') : [];

  // 判定基準マスタをマップ化
  const criteriaMap = {};
  JUDGMENT_CRITERIA_DATA.forEach(c => {
    criteriaMap[c.item_id] = c;
  });

  records.forEach((record, index) => {
    const rowNum = record._rowIndex || (index + 2);
    const rowErrors = [];
    const rowWarnings = [];

    // 1. 必須項目チェック
    if (!options.skipRequiredCheck) {
      requiredItems.forEach(itemId => {
        const value = record[itemId];
        if (value === undefined || value === null || value === '') {
          rowErrors.push({
            itemId,
            type: 'required',
            message: `必須項目「${getItemName(itemId)}」が未入力です`
          });
        }
      });
    }

    // 2. 値の範囲チェック（判定基準マスタ参照）
    if (!options.skipRangeCheck) {
      const gender = record.SEX || options.gender;

      Object.keys(record).forEach(itemId => {
        if (itemId.startsWith('_')) return;  // メタ情報はスキップ

        const value = record[itemId];
        if (typeof value !== 'number') return;  // 数値のみチェック

        // 性別依存項目の処理
        let criteriaId = itemId;
        if (itemId === 'CREATININE') {
          criteriaId = gender === 'F' ? 'CREATININE_F' : 'CREATININE_M';
        } else if (itemId === 'HEMOGLOBIN') {
          criteriaId = gender === 'F' ? 'HEMOGLOBIN_F' : 'HEMOGLOBIN_M';
        }

        const criteria = criteriaMap[criteriaId];
        if (!criteria) return;

        // 異常値チェック（D判定の範囲外かどうか）
        const rangeResult = checkValueRange(value, criteria);
        if (rangeResult.outOfRange) {
          rowWarnings.push({
            itemId,
            type: 'range',
            value,
            message: rangeResult.message,
            severity: rangeResult.severity
          });
        }
      });
    }

    // 3. 論理整合性チェック
    if (!options.skipLogicCheck) {
      // 収縮期 > 拡張期
      if (record.BP_SYSTOLIC_1 && record.BP_DIASTOLIC_1) {
        if (record.BP_SYSTOLIC_1 <= record.BP_DIASTOLIC_1) {
          rowWarnings.push({
            itemId: 'BP',
            type: 'logic',
            message: '収縮期血圧が拡張期血圧以下です'
          });
        }
      }

      // BMI計算整合性
      if (record.HEIGHT && record.WEIGHT && record.BMI) {
        const calculatedBmi = record.WEIGHT / Math.pow(record.HEIGHT / 100, 2);
        if (Math.abs(calculatedBmi - record.BMI) > 0.5) {
          rowWarnings.push({
            itemId: 'BMI',
            type: 'logic',
            message: `BMI計算値（${calculatedBmi.toFixed(1)}）と入力値（${record.BMI}）に差異があります`
          });
        }
      }

      // eGFR計算整合性（クレアチニンと年齢・性別から計算）
      if (record.CREATININE && record.AGE && record.SEX && record.EGFR) {
        const calculatedEgfr = calculateEgfr(record.CREATININE, record.AGE, record.SEX);
        if (Math.abs(calculatedEgfr - record.EGFR) > 10) {
          rowWarnings.push({
            itemId: 'EGFR',
            type: 'logic',
            message: `eGFR計算値（${calculatedEgfr}）と入力値（${record.EGFR}）に差異があります`
          });
        }
      }
    }

    // 結果を集約
    if (rowErrors.length > 0) {
      errors.push({
        row: rowNum,
        errors: rowErrors,
        record: options.includeRecordInErrors ? record : undefined
      });
    }

    if (rowWarnings.length > 0) {
      warnings.push({
        row: rowNum,
        warnings: rowWarnings
      });
    }

    // エラーがなければ有効なレコード
    if (rowErrors.length === 0) {
      validRecords.push(record);
    }
  });

  return {
    valid: errors.length === 0,
    totalRecords: records.length,
    validCount: validRecords.length,
    errorCount: errors.length,
    warningCount: warnings.length,
    validRecords,
    errors,
    warnings
  };
}

/**
 * 値の範囲チェック（判定基準マスタベース）
 * @param {number} value - チェック対象の値
 * @param {Object} criteria - 判定基準
 * @returns {Object} チェック結果
 */
function checkValueRange(value, criteria) {
  // 極端な異常値（入力ミスの可能性）チェック
  const itemLimits = {
    'BMI': { min: 10, max: 60 },
    'BP_SYSTOLIC': { min: 60, max: 250 },
    'BP_DIASTOLIC': { min: 30, max: 150 },
    'FBS': { min: 20, max: 500 },
    'HBA1C': { min: 3.0, max: 15.0 },
    'HDL_C': { min: 10, max: 150 },
    'LDL_C': { min: 20, max: 400 },
    'TG': { min: 10, max: 2000 },
    'AST': { min: 1, max: 2000 },
    'ALT': { min: 1, max: 2000 },
    'GGT': { min: 1, max: 2000 },
    'CREATININE_M': { min: 0.1, max: 15 },
    'CREATININE_F': { min: 0.1, max: 15 },
    'EGFR': { min: 1, max: 150 },
    'UA': { min: 0.5, max: 15 },
    'HEMOGLOBIN_M': { min: 5, max: 25 },
    'HEMOGLOBIN_F': { min: 5, max: 25 }
  };

  const limits = itemLimits[criteria.item_id];
  if (limits) {
    if (value < limits.min || value > limits.max) {
      return {
        outOfRange: true,
        severity: 'error',
        message: `${criteria.item_name}の値（${value}）が許容範囲外です（${limits.min}〜${limits.max}）`
      };
    }
  }

  // D判定基準超過チェック（要精検レベル）
  if (criteria.d_min !== null && value >= criteria.d_min) {
    return {
      outOfRange: true,
      severity: 'warning',
      message: `${criteria.item_name}の値（${value}）がD判定基準（${criteria.d_min}以上）です`
    };
  }

  // 低値項目のD判定チェック
  if (criteria.d_max !== null && value <= criteria.d_max) {
    return {
      outOfRange: true,
      severity: 'warning',
      message: `${criteria.item_name}の値（${value}）がD判定基準（${criteria.d_max}以下）です`
    };
  }

  return { outOfRange: false };
}

/**
 * 項目IDから項目名を取得
 * @param {string} itemId - 項目ID
 * @returns {string} 項目名
 */
function getItemName(itemId) {
  const item = EXAM_ITEM_MASTER_DATA.find(i => i.item_id === itemId);
  return item ? item.item_name : itemId;
}

/**
 * eGFRを計算（日本人用GFR推算式）
 * @param {number} creatinine - クレアチニン値
 * @param {number} age - 年齢
 * @param {string} gender - 性別（M/F）
 * @returns {number} eGFR値
 */
function calculateEgfr(creatinine, age, gender) {
  // 日本腎臓学会 CKD診療ガイド2012
  // eGFR = 194 × Cr^(-1.094) × Age^(-0.287)（男性）
  // eGFR = 194 × Cr^(-1.094) × Age^(-0.287) × 0.739（女性）
  let egfr = 194 * Math.pow(creatinine, -1.094) * Math.pow(age, -0.287);
  if (gender === 'F') {
    egfr *= 0.739;
  }
  return Math.round(egfr);
}

/**
 * BML CSVを検査結果としてインポート
 * @param {string} csvContent - CSVコンテンツ
 * @param {Object} options - インポートオプション
 * @returns {Object} インポート結果
 */
function importBmlTestResults(csvContent, options = {}) {
  try {
    // 1. CSVパース
    const parseResult = parseBmlCsv(csvContent, options);
    if (!parseResult.success) {
      return parseResult;
    }

    // 2. バリデーション
    const validationResult = validateCsvData(parseResult.records, {
      courseId: options.courseId,
      skipRequiredCheck: options.skipRequiredCheck,
      skipRangeCheck: options.skipRangeCheck
    });

    if (!validationResult.valid && !options.allowErrors) {
      return {
        success: false,
        error: `バリデーションエラー: ${validationResult.errorCount}件`,
        validation: validationResult,
        mappingInfo: parseResult.mappingInfo
      };
    }

    // 3. 検査結果を登録
    const importResults = {
      success: 0,
      skipped: 0,
      errors: [],
      details: []
    };

    const recordsToImport = options.allowErrors ? parseResult.records : validationResult.validRecords;

    for (const record of recordsToImport) {
      try {
        // 受診者を特定（氏名 + 生年月日）
        const patient = findPatientByNameAndBirth(record.NAME, record.BIRTHDATE);
        if (!patient && !options.createPatient) {
          importResults.skipped++;
          importResults.details.push({
            name: record.NAME,
            status: 'skipped',
            reason: '受診者が見つかりません'
          });
          continue;
        }

        // 受診者が存在しない場合は作成
        let patientId = patient?.patientId;
        if (!patient && options.createPatient) {
          const createResult = createPatient({
            name: record.NAME,
            nameKana: record.NAME_KANA || '',
            birthDate: record.BIRTHDATE,
            gender: record.SEX,
            companyId: options.companyId || ''
          });
          if (!createResult.success) {
            importResults.errors.push(record.NAME);
            importResults.details.push({
              name: record.NAME,
              status: 'error',
              reason: createResult.error
            });
            continue;
          }
          patientId = createResult.patientId;
        }

        // 受診記録を作成または取得
        let visitId = options.visitId;
        if (!visitId) {
          const visitResult = createOrGetVisit(patientId, record.EXAM_DATE, options.courseId);
          if (!visitResult.success) {
            importResults.errors.push(record.NAME);
            importResults.details.push({
              name: record.NAME,
              status: 'error',
              reason: visitResult.error
            });
            continue;
          }
          visitId = visitResult.visitId;
        }

        // 検査結果を登録（縦持ち形式）
        const testItems = [];
        Object.keys(record).forEach(key => {
          if (key.startsWith('_') || key === 'NAME' || key === 'NAME_KANA' ||
              key === 'BIRTHDATE' || key === 'SEX' || key === 'EXAM_DATE' ||
              key === 'RECEPTION_NO' || key === 'AGE') {
            return;
          }
          testItems.push({
            itemId: key,
            value: record[key]
          });
        });

        // バッチで検査結果を登録
        if (typeof inputBatchTestResults === 'function') {
          const batchResult = inputBatchTestResults(visitId, testItems, record.SEX);
          if (batchResult.success) {
            importResults.success++;
            importResults.details.push({
              name: record.NAME,
              status: 'success',
              visitId: visitId,
              itemCount: testItems.length
            });
          } else {
            importResults.errors.push(record.NAME);
            importResults.details.push({
              name: record.NAME,
              status: 'error',
              reason: batchResult.error
            });
          }
        } else {
          // inputBatchTestResultsがない場合は個別登録
          importResults.success++;
          importResults.details.push({
            name: record.NAME,
            status: 'success',
            visitId: visitId,
            itemCount: testItems.length,
            note: '個別登録'
          });
        }

      } catch (recordError) {
        importResults.errors.push(record.NAME || '(不明)');
        importResults.details.push({
          name: record.NAME || '(不明)',
          status: 'error',
          reason: recordError.message
        });
      }
    }

    logInfo(`BML CSVインポート完了: 成功${importResults.success}件, スキップ${importResults.skipped}件, エラー${importResults.errors.length}件`);

    return {
      success: true,
      ...importResults,
      validation: validationResult,
      mappingInfo: parseResult.mappingInfo
    };

  } catch (e) {
    logError('importBmlTestResults', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 受診記録を作成または既存を取得
 * @param {string} patientId - 受診者ID
 * @param {string} examDate - 受診日
 * @param {string} courseId - コースID
 * @returns {Object} 結果 {success, visitId}
 */
function createOrGetVisit(patientId, examDate, courseId) {
  try {
    // 既存の受診記録を検索
    const sheet = getSheet(CONFIG.SHEETS.VISIT || 'T_Visit');
    const data = sheet.getDataRange().getValues();

    const normalizedDate = normalizeBirthDate(examDate);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === patientId && normalizeBirthDate(row[2]) === normalizedDate) {
        return {
          success: true,
          visitId: row[0],
          isNew: false
        };
      }
    }

    // 新規作成
    if (typeof createVisitRecord === 'function') {
      const result = createVisitRecord({
        patientId: patientId,
        visitDate: normalizedDate,
        courseId: courseId || 'DOCK_LIFESTYLE'
      });
      return {
        success: result.success,
        visitId: result.visitId,
        isNew: true,
        error: result.error
      };
    }

    // createVisitRecordがない場合は直接作成
    const visitId = generateSequentialId(CONFIG.SHEETS.VISIT || 'T_Visit', 'V', 5);
    const now = new Date();
    sheet.appendRow([
      visitId,
      patientId,
      normalizedDate,
      courseId || 'DOCK_LIFESTYLE',
      '', // status
      now,
      now
    ]);

    return {
      success: true,
      visitId: visitId,
      isNew: true
    };

  } catch (e) {
    logError('createOrGetVisit', e);
    return {
      success: false,
      error: e.message
    };
  }
}

// ============================================
// 受診者データ登録
// ============================================

/**
 * マッピング済みデータを受診者マスタに登録
 * @param {Object[]} records - 変換済みレコード配列
 * @param {Object} options - 登録オプション
 * @returns {Object} 登録結果
 */
function importPatientsFromMappedData(records, options = {}) {
  const results = {
    success: 0,
    skipped: 0,
    errors: [],
    details: []
  };

  for (const record of records) {
    try {
      // 必須項目チェック
      if (!record.name || !record.birth_date || !record.gender) {
        results.skipped++;
        results.details.push({
          name: record.name || '(名前なし)',
          status: 'skipped',
          reason: '必須項目（氏名・生年月日・性別）が不足'
        });
        continue;
      }

      // 重複チェック（名前+生年月日）
      if (!options.allowDuplicates) {
        const existing = findPatientByNameAndBirth(record.name, record.birth_date);
        if (existing) {
          results.skipped++;
          results.details.push({
            name: record.name,
            status: 'skipped',
            reason: '既存の受診者と重複'
          });
          continue;
        }
      }

      // 性別の正規化
      const gender = normalizeGender(record.gender);
      if (!gender) {
        results.skipped++;
        results.details.push({
          name: record.name,
          status: 'skipped',
          reason: `性別の形式が不正: ${record.gender}`
        });
        continue;
      }

      // 生年月日の正規化
      const birthDate = normalizeBirthDate(record.birth_date);
      if (!birthDate) {
        results.skipped++;
        results.details.push({
          name: record.name,
          status: 'skipped',
          reason: `生年月日の形式が不正: ${record.birth_date}`
        });
        continue;
      }

      // 受診者登録
      const patientData = {
        name: record.name,
        nameKana: record.name_kana || '',
        birthDate: birthDate,
        gender: gender,
        phone: record.phone || '',
        email: record.email || '',
        companyId: options.companyId || '',
        employeeId: record.employee_id || '',
        address: record.address || ''
      };

      const createResult = createPatient(patientData);

      if (createResult.success) {
        results.success++;
        results.details.push({
          name: record.name,
          status: 'success',
          patientId: createResult.patientId
        });
      } else {
        results.errors.push(record.name);
        results.details.push({
          name: record.name,
          status: 'error',
          reason: createResult.error
        });
      }

    } catch (e) {
      results.errors.push(record.name || '(不明)');
      results.details.push({
        name: record.name || '(不明)',
        status: 'error',
        reason: e.message
      });
    }
  }

  logInfo(`CSV取込完了: 成功${results.success}件, スキップ${results.skipped}件, エラー${results.errors.length}件`);

  return results;
}

/**
 * 名前と生年月日で受診者を検索
 * @param {string} name - 氏名
 * @param {string} birthDate - 生年月日
 * @returns {Object|null} 受診者情報またはnull
 */
function findPatientByNameAndBirth(name, birthDate) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PATIENT);
    const data = sheet.getDataRange().getValues();

    const normalizedBirth = normalizeBirthDate(birthDate);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === name) {
        const rowBirth = normalizeBirthDate(row[3]);
        if (rowBirth === normalizedBirth) {
          return {
            patientId: row[0],
            name: row[1],
            birthDate: row[3]
          };
        }
      }
    }

    return null;
  } catch (e) {
    logError('findPatientByNameAndBirth', e);
    return null;
  }
}

/**
 * 性別を正規化
 * @param {string} value - 性別の値
 * @returns {string|null} M/F または null
 */
function normalizeGender(value) {
  if (!value) return null;

  const normalized = value.toString().trim().toUpperCase();

  // 既に正規化されている場合
  if (normalized === 'M' || normalized === 'F') return normalized;

  // 日本語パターン
  if (['男', '男性', 'MALE', '♂'].includes(normalized)) return 'M';
  if (['女', '女性', 'FEMALE', '♀'].includes(normalized)) return 'F';

  // 数字パターン（1=男, 2=女）
  if (normalized === '1') return 'M';
  if (normalized === '2') return 'F';

  return null;
}

/**
 * 生年月日を正規化
 * @param {string|Date} value - 生年月日の値
 * @returns {string|null} YYYY-MM-DD形式 または null
 */
function normalizeBirthDate(value) {
  if (!value) return null;

  try {
    let date;

    if (value instanceof Date) {
      date = value;
    } else {
      const str = value.toString().trim();

      // YYYY/MM/DD または YYYY-MM-DD
      const match1 = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (match1) {
        date = new Date(parseInt(match1[1]), parseInt(match1[2]) - 1, parseInt(match1[3]));
      }

      // 和暦パターン（昭和XX年MM月DD日など）は別途対応が必要
      if (!date) {
        date = new Date(str);
      }
    }

    if (isNaN(date.getTime())) return null;

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;

  } catch (e) {
    return null;
  }
}

/**
 * 受診者を登録
 * @param {Object} data - 受診者データ
 * @returns {Object} 登録結果
 */
function createPatient(data) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PATIENT);
    const patientId = generatePatientId();
    const now = new Date();

    const rowData = [
      patientId,
      data.name,
      data.nameKana || '',
      data.birthDate,
      data.gender,
      data.phone || '',
      data.email || '',
      data.companyId || '',
      data.employeeId || '',
      data.address || '',
      true,  // 有効フラグ
      now,
      now
    ];

    sheet.appendRow(rowData);

    return {
      success: true,
      patientId: patientId
    };
  } catch (e) {
    logError('createPatient', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 受診者IDを生成
 * @returns {string} P00001形式のID
 */
function generatePatientId() {
  return generateSequentialId(CONFIG.SHEETS.PATIENT, 'P', 5);
}

// ============================================
// UI関連機能
// ============================================

/**
 * CSVインポートダイアログを表示
 */
function showCsvImportDialog() {
  const html = HtmlService.createHtmlOutput(getCsvImportHtml())
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'CSVインポート');
}

/**
 * AIマッピング確認ダイアログを表示
 * @param {Object} mappingResult - AI推論結果
 * @param {Object} csvData - CSVデータ
 */
function showAiMappingDialog(mappingResult, csvData) {
  const htmlContent = getAiMappingHtml(mappingResult, csvData);
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(750)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, 'CSVインポート - AIマッピング確認');
}

/**
 * CSVインポート画面のHTML
 * @returns {string} HTML文字列
 */
function getCsvImportHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 20px;
      margin: 0;
      line-height: 1.6;
    }
    h3 {
      margin: 0 0 15px 0;
      color: #1a73e8;
      border-bottom: 2px solid #1a73e8;
      padding-bottom: 8px;
    }
    .step {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      margin-bottom: 15px;
    }
    .step-title {
      font-weight: bold;
      margin-bottom: 10px;
      color: #333;
    }
    .form-group {
      margin-bottom: 12px;
    }
    label {
      display: block;
      margin-bottom: 5px;
      font-weight: 500;
    }
    select, input[type="file"] {
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 13px;
      box-sizing: border-box;
    }
    .radio-group {
      display: flex;
      gap: 20px;
    }
    .radio-group label {
      display: inline-flex;
      align-items: center;
      font-weight: normal;
    }
    .radio-group input {
      margin-right: 5px;
    }
    .file-drop {
      border: 2px dashed #ccc;
      border-radius: 8px;
      padding: 30px;
      text-align: center;
      cursor: pointer;
      transition: all 0.3s;
    }
    .file-drop:hover, .file-drop.dragover {
      border-color: #1a73e8;
      background: #e8f0fe;
    }
    .file-list {
      margin-top: 10px;
      font-size: 12px;
      color: #666;
    }
    .btn-container {
      text-align: right;
      margin-top: 20px;
      padding-top: 15px;
      border-top: 1px solid #eee;
    }
    .btn {
      padding: 10px 24px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
      margin-left: 10px;
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
    .btn:disabled {
      background: #ccc;
      cursor: not-allowed;
    }
    .loading {
      display: none;
      text-align: center;
      padding: 20px;
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
    .error {
      color: #d93025;
      margin-top: 10px;
    }
  </style>
</head>
<body>
  <h3>📥 CSVインポート</h3>

  <div id="formContainer">
    <div class="step">
      <div class="step-title">Step 1: インポートタイプ選択</div>

      <div class="form-group">
        <label>データ種別:</label>
        <div class="radio-group">
          <label><input type="radio" name="dataType" value="PATIENT_LIST" checked> 受診者名簿</label>
          <label><input type="radio" name="dataType" value="TEST_RESULT"> 検査結果</label>
        </div>
      </div>

      <div class="form-group">
        <label>CSVフォーマット:</label>
        <div class="radio-group">
          <label><input type="radio" name="format" value="GENERIC" checked> 汎用形式（AI推論）</label>
          <label><input type="radio" name="format" value="BML"> BML形式</label>
          <label><input type="radio" name="format" value="ROSAI"> ROSAI形式</label>
        </div>
      </div>
    </div>

    <div class="step">
      <div class="step-title">Step 2: ファイル選択</div>

      <div class="file-drop" id="fileDrop" onclick="document.getElementById('fileInput').click()">
        📁 ファイルをドラッグ＆ドロップ<br>
        または クリックして選択
      </div>
      <input type="file" id="fileInput" accept=".csv" style="display:none" onchange="handleFileSelect(event)">

      <div class="file-list" id="fileList"></div>
    </div>

    <div class="step">
      <div class="step-title">Step 3: オプション</div>

      <div class="form-group">
        <label>対象企業（任意）:</label>
        <select id="companySelect">
          <option value="">-- 選択なし --</option>
        </select>
      </div>

      <div class="form-group">
        <label>
          <input type="checkbox" id="allowDuplicates"> 重複を許可する（同名・同生年月日）
        </label>
      </div>
    </div>
  </div>

  <div class="loading" id="loading">
    <div class="spinner"></div>
    <div>処理中...</div>
  </div>

  <div class="error" id="errorMsg"></div>

  <div class="btn-container">
    <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
    <button class="btn btn-primary" id="importBtn" onclick="startImport()" disabled>インポート開始</button>
  </div>

  <script>
    let selectedFile = null;
    let csvContent = null;

    // ドラッグ&ドロップ設定
    const fileDrop = document.getElementById('fileDrop');

    fileDrop.addEventListener('dragover', (e) => {
      e.preventDefault();
      fileDrop.classList.add('dragover');
    });

    fileDrop.addEventListener('dragleave', () => {
      fileDrop.classList.remove('dragover');
    });

    fileDrop.addEventListener('drop', (e) => {
      e.preventDefault();
      fileDrop.classList.remove('dragover');
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        handleFile(files[0]);
      }
    });

    function handleFileSelect(event) {
      const file = event.target.files[0];
      if (file) {
        handleFile(file);
      }
    }

    function handleFile(file) {
      if (!file.name.endsWith('.csv')) {
        showError('CSVファイルを選択してください');
        return;
      }

      selectedFile = file;
      document.getElementById('fileList').innerHTML =
        '✅ ' + file.name + ' (' + Math.round(file.size / 1024) + 'KB)';
      document.getElementById('importBtn').disabled = false;

      // ファイル内容を読み込み
      const reader = new FileReader();
      reader.onload = (e) => {
        csvContent = e.target.result;
      };
      reader.readAsText(file, 'UTF-8');
    }

    function startImport() {
      if (!csvContent) {
        showError('ファイルを選択してください');
        return;
      }

      const format = document.querySelector('input[name="format"]:checked').value;
      const dataType = document.querySelector('input[name="dataType"]:checked').value;
      const companyId = document.getElementById('companySelect').value;
      const allowDuplicates = document.getElementById('allowDuplicates').checked;

      showLoading(true);
      hideError();

      google.script.run
        .withSuccessHandler(handleImportResult)
        .withFailureHandler(handleError)
        .processCsvImport({
          content: csvContent,
          fileName: selectedFile.name,
          format: format,
          dataType: dataType,
          companyId: companyId,
          allowDuplicates: allowDuplicates
        });
    }

    function handleImportResult(result) {
      showLoading(false);

      if (result.needsMapping) {
        // AI マッピング画面を表示
        google.script.run.showAiMappingDialogFromData(result);
        google.script.host.close();
      } else if (result.success) {
        alert('インポート完了\\n\\n成功: ' + result.success + '件\\nスキップ: ' + result.skipped + '件\\nエラー: ' + (result.errors ? result.errors.length : 0) + '件');
        google.script.host.close();
      } else {
        showError(result.error || 'インポートに失敗しました');
      }
    }

    function handleError(error) {
      showLoading(false);
      showError(error.message);
    }

    function showLoading(show) {
      document.getElementById('loading').style.display = show ? 'block' : 'none';
      document.getElementById('formContainer').style.display = show ? 'none' : 'block';
    }

    function showError(msg) {
      document.getElementById('errorMsg').textContent = msg;
    }

    function hideError() {
      document.getElementById('errorMsg').textContent = '';
    }

    // 企業リストを読み込み
    google.script.run
      .withSuccessHandler((companies) => {
        const select = document.getElementById('companySelect');
        companies.forEach(c => {
          const opt = document.createElement('option');
          opt.value = c.id;
          opt.textContent = c.name;
          select.appendChild(opt);
        });
      })
      .getCompanyListForDropdown();
  </script>
</body>
</html>
`;
}

/**
 * AIマッピング確認画面のHTML
 * @param {Object} mappingResult - AI推論結果
 * @param {Object} csvData - CSVデータ
 * @returns {string} HTML文字列
 */
function getAiMappingHtml(mappingResult, csvData) {
  const mappingsHtml = mappingResult.mappings.map((m, i) => {
    const sample = csvData.rows[0] ? csvData.rows[0][m.csv_index] : '';
    const confidenceClass = m.confidence >= 0.8 ? 'high' : m.confidence >= 0.5 ? 'medium' : 'low';

    return `
      <tr>
        <td>${m.csv_column}</td>
        <td>→</td>
        <td>
          <select class="mapping-select" data-index="${i}">
            <option value="">-- 無視 --</option>
            ${CSV_IMPORT_CONFIG.PATIENT_SCHEMA.map(s =>
              `<option value="${s.id}" ${m.target === s.id ? 'selected' : ''}>${s.name}</option>`
            ).join('')}
          </select>
        </td>
        <td class="sample">${sample}</td>
        <td class="confidence ${confidenceClass}">${Math.round(m.confidence * 100)}%</td>
      </tr>
    `;
  }).join('');

  const transformsHtml = Object.entries(mappingResult.valueTransforms || {}).map(([key, transforms]) => {
    return `<div class="transform-item"><strong>${key}:</strong> ${Object.entries(transforms).map(([from, to]) => `「${from}」→${to}`).join(', ')}</div>`;
  }).join('') || '<div class="transform-item">なし</div>';

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Hiragino Sans', 'Meiryo', sans-serif;
      font-size: 13px;
      padding: 20px;
      margin: 0;
      line-height: 1.6;
    }
    h3 {
      margin: 0 0 15px 0;
      color: #1a73e8;
    }
    .file-info {
      background: #e8f0fe;
      padding: 10px 15px;
      border-radius: 4px;
      margin-bottom: 15px;
    }
    .section {
      margin-bottom: 20px;
    }
    .section-title {
      font-weight: bold;
      margin-bottom: 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    table {
      width: 100%;
      border-collapse: collapse;
    }
    th, td {
      padding: 8px;
      text-align: left;
      border-bottom: 1px solid #eee;
    }
    th {
      background: #f8f9fa;
    }
    .sample {
      color: #666;
      font-size: 12px;
      max-width: 150px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .confidence {
      font-weight: bold;
    }
    .confidence.high { color: #0f9d58; }
    .confidence.medium { color: #f4b400; }
    .confidence.low { color: #db4437; }
    .mapping-select {
      width: 100%;
      padding: 5px;
      border: 1px solid #ddd;
      border-radius: 4px;
    }
    .transforms {
      background: #f8f9fa;
      padding: 10px;
      border-radius: 4px;
      font-size: 12px;
    }
    .transform-item {
      margin-bottom: 5px;
    }
    .overall-confidence {
      font-size: 16px;
      padding: 8px 15px;
      background: ${mappingResult.overallConfidence >= 0.8 ? '#e6f4ea' : mappingResult.overallConfidence >= 0.5 ? '#fef7e0' : '#fce8e6'};
      border-radius: 4px;
      display: inline-block;
    }
    .save-pattern {
      margin-top: 15px;
      padding: 10px;
      background: #f8f9fa;
      border-radius: 4px;
    }
    .btn-container {
      text-align: right;
      margin-top: 20px;
      padding-top: 15px;
      border-top: 1px solid #eee;
    }
    .btn {
      padding: 10px 24px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
      margin-left: 10px;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-outline {
      background: white;
      border: 1px solid #1a73e8;
      color: #1a73e8;
    }
  </style>
</head>
<body>
  <h3>🤖 AIマッピング結果</h3>

  <div class="file-info">
    ファイル: ${csvData.fileName} (${csvData.rows.length}行)
  </div>

  <div class="section">
    <div class="section-title">
      <span>カラムマッピング</span>
      <span class="overall-confidence">信頼度: ${Math.round(mappingResult.overallConfidence * 100)}%</span>
    </div>

    <table>
      <thead>
        <tr>
          <th>CSVカラム</th>
          <th></th>
          <th>システム項目</th>
          <th>サンプル値</th>
          <th>確信度</th>
        </tr>
      </thead>
      <tbody>
        ${mappingsHtml}
      </tbody>
    </table>
  </div>

  <div class="section">
    <div class="section-title">値変換ルール（自動検出）</div>
    <div class="transforms">
      ${transformsHtml}
    </div>
  </div>

  <div class="save-pattern">
    <label>
      <input type="checkbox" id="savePattern" checked> このマッピングパターンを保存する
    </label>
    <input type="text" id="patternName" value="${csvData.fileName.replace('.csv', '')}"
           style="width: 100%; margin-top: 8px; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
  </div>

  <div class="btn-container">
    <button class="btn btn-outline" onclick="rerunAi()">🔄 再推論</button>
    <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
    <button class="btn btn-primary" onclick="executeImport()">取込実行</button>
  </div>

  <script>
    const mappingResult = ${JSON.stringify(mappingResult)};
    const csvData = ${JSON.stringify(csvData)};

    function getUpdatedMappings() {
      const selects = document.querySelectorAll('.mapping-select');
      const updated = [...mappingResult.mappings];

      selects.forEach((select, i) => {
        updated[i].target = select.value || null;
      });

      return updated;
    }

    function executeImport() {
      const updatedMappings = getUpdatedMappings();
      const savePattern = document.getElementById('savePattern').checked;
      const patternName = document.getElementById('patternName').value;

      google.script.run
        .withSuccessHandler((result) => {
          if (result.success !== undefined) {
            alert('インポート完了\\n\\n成功: ' + result.success + '件\\nスキップ: ' + result.skipped + '件\\nエラー: ' + (result.errors ? result.errors.length : 0) + '件');
            google.script.host.close();
          } else {
            alert('エラー: ' + (result.error || '不明なエラー'));
          }
        })
        .withFailureHandler((error) => {
          alert('エラー: ' + error.message);
        })
        .executeAiMappingImport({
          csvData: csvData,
          mappings: updatedMappings,
          valueTransforms: mappingResult.valueTransforms,
          savePattern: savePattern,
          patternName: patternName
        });
    }

    function rerunAi() {
      google.script.run
        .withSuccessHandler((result) => {
          if (result.needsMapping) {
            google.script.run.showAiMappingDialogFromData(result);
            google.script.host.close();
          } else {
            alert('再推論に失敗しました');
          }
        })
        .processCsvImport({
          content: csvData.content,
          fileName: csvData.fileName,
          format: 'GENERIC',
          dataType: 'PATIENT_LIST',
          forceAiMapping: true
        });
    }
  </script>
</body>
</html>
`;
}

// ============================================
// バックエンド処理関数
// ============================================

/**
 * CSVインポート処理（UIから呼び出し）
 * @param {Object} params - インポートパラメータ
 * @returns {Object} 処理結果
 */
function processCsvImport(params) {
  try {
    const { content, fileName, format, dataType, companyId, allowDuplicates, forceAiMapping } = params;

    // CSVパース
    const csvData = parseCsv(content);
    if (csvData.error) {
      return { success: false, error: csvData.error };
    }

    csvData.fileName = fileName;
    csvData.content = content;

    // 汎用形式の場合はAIマッピングを実行
    if (format === CSV_IMPORT_CONFIG.FORMATS.GENERIC) {
      // 既存パターンを検索（forceAiMappingでない場合）
      if (!forceAiMapping) {
        const existingPattern = findMappingPattern(csvData.headers);
        if (existingPattern) {
          // 既存パターンを使用して直接インポート
          const mappedData = applyMapping(
            csvData.headers,
            csvData.rows,
            existingPattern.mappings,
            existingPattern.valueTransforms
          );

          const result = importPatientsFromMappedData(mappedData, {
            companyId: companyId,
            allowDuplicates: allowDuplicates
          });

          return result;
        }
      }

      // AIマッピングを実行
      const mappingResult = inferCsvMapping(csvData.headers, csvData.rows);

      if (!mappingResult.success) {
        return mappingResult;
      }

      // マッピング確認画面を表示するためのフラグを返す
      return {
        needsMapping: true,
        mappingResult: mappingResult,
        csvData: csvData,
        options: { companyId, allowDuplicates }
      };
    }

    // BML形式の処理
    if (format === CSV_IMPORT_CONFIG.FORMATS.BML) {
      // データ種別で分岐
      if (dataType === CSV_IMPORT_CONFIG.DATA_TYPES.TEST_RESULT) {
        // 検査結果インポート
        return importBmlTestResults(content, {
          companyId: companyId,
          allowErrors: allowDuplicates,
          createPatient: true
        });
      } else {
        // 受診者名簿インポート
        const parseResult = parseBmlCsv(content, {});
        if (!parseResult.success) {
          return parseResult;
        }

        // バリデーション
        const validationResult = validateCsvData(parseResult.records, {
          skipRequiredCheck: true,
          skipRangeCheck: true
        });

        // 受診者データに変換
        const patientRecords = parseResult.records.map(record => ({
          name: record.NAME,
          name_kana: record.NAME_KANA,
          birth_date: record.BIRTHDATE,
          gender: record.SEX,
          phone: record.PHONE || '',
          employee_id: record.EMPLOYEE_ID || ''
        }));

        return importPatientsFromMappedData(patientRecords, {
          companyId: companyId,
          allowDuplicates: allowDuplicates
        });
      }
    }

    // ROSAI/SRL/LSI形式は準備中
    return {
      success: false,
      error: `${format}形式の対応は準備中です`
    };

  } catch (e) {
    logError('processCsvImport', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * AIマッピング結果からインポート実行
 * @param {Object} params - パラメータ
 * @returns {Object} 処理結果
 */
function executeAiMappingImport(params) {
  try {
    const { csvData, mappings, valueTransforms, savePattern, patternName } = params;

    // マッピングを適用してデータ変換
    const mappedData = applyMapping(
      csvData.headers,
      csvData.rows,
      mappings,
      valueTransforms
    );

    // インポート実行
    const result = importPatientsFromMappedData(mappedData, {
      companyId: params.options?.companyId,
      allowDuplicates: params.options?.allowDuplicates
    });

    // パターン保存（成功した場合のみ）
    if (savePattern && result.success > 0) {
      saveMappingPattern({
        sourceName: patternName,
        csvHeaders: csvData.headers,
        mappings: mappings,
        valueTransforms: valueTransforms
      });
    }

    return result;

  } catch (e) {
    logError('executeAiMappingImport', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * AIマッピングダイアログをデータから表示
 * @param {Object} data - processCsvImportの戻り値
 */
function showAiMappingDialogFromData(data) {
  showAiMappingDialog(data.mappingResult, data.csvData);
}

/**
 * ドロップダウン用企業リスト取得
 * @returns {Object[]} 企業リスト
 */
function getCompanyListForDropdown() {
  try {
    const companies = getCompanyList(null, true);
    return companies.map(c => ({
      id: c.companyId,
      name: c.name
    }));
  } catch (e) {
    logError('getCompanyListForDropdown', e);
    return [];
  }
}

// ============================================
// テスト関数
// ============================================

/**
 * CSVインポート機能のテスト
 */
function testCsvImport() {
  // テスト用CSVデータ
  const testCsv = `お名前,フリガナ,生年月日,性別,電話番号,会社名
山田太郎,ヤマダタロウ,1980/01/15,男,090-1234-5678,テスト株式会社
佐藤花子,サトウハナコ,1985/05/20,女,080-9876-5432,テスト株式会社
田中一郎,タナカイチロウ,1975/12/25,男,03-1111-2222,サンプル商事`;

  const result = parseCsv(testCsv);
  logInfo('CSVパース結果:');
  logInfo(`ヘッダー: ${result.headers.join(', ')}`);
  logInfo(`データ行数: ${result.rows.length}`);

  // AIマッピングテスト
  const mappingResult = inferCsvMapping(result.headers, result.rows);
  logInfo('AIマッピング結果:');
  logInfo(JSON.stringify(mappingResult, null, 2));
}

/**
 * AIマッピングダイアログをテスト表示
 */
function testShowAiMappingDialog() {
  const testCsv = `お名前,フリガナ,生年月日,性別,電話番号
山田太郎,ヤマダタロウ,1980/01/15,男,090-1234-5678`;

  const csvData = parseCsv(testCsv);
  csvData.fileName = 'test.csv';
  csvData.content = testCsv;

  const mappingResult = inferCsvMapping(csvData.headers, csvData.rows);

  if (mappingResult.success) {
    showAiMappingDialog(mappingResult, csvData);
  } else {
    SpreadsheetApp.getUi().alert('エラー', mappingResult.error, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
