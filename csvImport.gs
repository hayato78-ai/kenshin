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
    GENERIC: 'GENERIC'  // 汎用形式（AI推論使用）
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

    // BML/ROSAI形式は既存のパーサーを使用
    // TODO: 既存のBML/ROSAIパーサー連携
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
