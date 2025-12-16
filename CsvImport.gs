/**
 * 健診結果DB 統合システム - CSVインポートモジュール
 *
 * 機能:
 * - CSVパース（クォート対応）
 * - データ正規化（性別、日付、検査値）
 * - 横持ち→縦持ち変換
 * - 受診者・検査結果インポート
 * - Claude AIマッピング推論
 * - マッピングパターン保存・再利用
 *
 * @version 1.0.0
 * @date 2025-12-16
 */

// ============================================
// CSVパース機能
// ============================================

/**
 * CSVコンテンツをパース
 * @param {string} csvContent - CSVの内容
 * @param {Object} options - オプション
 * @param {string} options.delimiter - 区切り文字（デフォルト: ,）
 * @param {boolean} options.hasHeader - ヘッダー有無（デフォルト: true）
 * @returns {Object} {headers: string[], rows: string[][], error: string|null}
 */
function parseCsv(csvContent, options = {}) {
  const delimiter = options.delimiter || ',';
  const hasHeader = options.hasHeader !== false;

  try {
    // 改行で分割（空行除去）
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
            // エスケープされたダブルクォート
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
      // ヘッダーなしの場合は自動生成
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
 * CSVプレビュー（最大行数制限付き）
 * @param {string} csvContent - CSVコンテンツ
 * @param {number} maxRows - 最大行数（デフォルト: 10）
 * @returns {Object} プレビュー結果
 */
function previewCsv(csvContent, maxRows = 10) {
  const result = parseCsv(csvContent);

  if (result.error) {
    return result;
  }

  return {
    headers: result.headers,
    rows: result.rows.slice(0, maxRows),
    totalRows: result.rows.length,
    error: null
  };
}

// ============================================
// データ正規化機能
// ============================================

/**
 * 性別を正規化（M/F形式）
 * @param {string} value - 入力値
 * @returns {string} 'M' | 'F' | ''
 */
function normalizeGender(value) {
  if (!value) return '';

  const v = String(value).trim().toUpperCase();

  // 男性パターン
  if (v === 'M' || v === '男' || v === '男性' || v === '1' || v === 'MALE') {
    return 'M';
  }

  // 女性パターン
  if (v === 'F' || v === '女' || v === '女性' || v === '2' || v === 'FEMALE') {
    return 'F';
  }

  return '';
}

/**
 * 生年月日を正規化（Date形式）
 * @param {string|Date} value - 入力値
 * @returns {Date|null} 日付またはnull
 */
function normalizeBirthDate(value) {
  if (!value) return null;

  // すでにDate型
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  const v = String(value).trim();

  // 各種形式を試行
  const patterns = [
    // YYYY/MM/DD
    /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/,
    // MM/DD/YYYY
    /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/,
    // YYYYMMDD
    /^(\d{4})(\d{2})(\d{2})$/,
    // 和暦（昭和・平成・令和）
    /^(明治|大正|昭和|平成|令和)(\d{1,2})年(\d{1,2})月(\d{1,2})日$/
  ];

  // YYYY/MM/DD または YYYY-MM-DD
  let match = v.match(patterns[0]);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }

  // MM/DD/YYYY（米国形式）
  match = v.match(patterns[1]);
  if (match) {
    return new Date(parseInt(match[3]), parseInt(match[1]) - 1, parseInt(match[2]));
  }

  // YYYYMMDD
  match = v.match(patterns[2]);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }

  // 和暦
  match = v.match(patterns[3]);
  if (match) {
    const eraYear = parseInt(match[2]);
    let year;
    switch (match[1]) {
      case '明治': year = eraYear + 1867; break;
      case '大正': year = eraYear + 1911; break;
      case '昭和': year = eraYear + 1925; break;
      case '平成': year = eraYear + 1988; break;
      case '令和': year = eraYear + 2018; break;
      default: return null;
    }
    return new Date(year, parseInt(match[3]) - 1, parseInt(match[4]));
  }

  // JavaScriptの標準パース
  const parsed = new Date(v);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * 受診日を正規化
 * @param {string|Date} value - 入力値
 * @returns {Date|null} 日付またはnull
 */
function normalizeVisitDate(value) {
  return normalizeBirthDate(value);
}

/**
 * 検査値を正規化
 * @param {string} value - 入力値
 * @param {string} dataType - データ型（numeric, text, select）
 * @returns {Object} {value: string, numericValue: number|null}
 */
function normalizeTestValue(value, dataType = 'text') {
  if (value === null || value === undefined || value === '') {
    return { value: '', numericValue: null };
  }

  const v = String(value).trim();

  // 空白のみ
  if (!v) {
    return { value: '', numericValue: null };
  }

  // 数値型の場合
  if (dataType === 'numeric') {
    // 不等号を含む値（<10, >100など）
    const inequalityMatch = v.match(/^([<>≤≥])?\s*(-?\d+\.?\d*)$/);
    if (inequalityMatch) {
      const numValue = parseFloat(inequalityMatch[2]);
      return {
        value: v,
        numericValue: isNaN(numValue) ? null : numValue
      };
    }

    // 純粋な数値
    const numValue = parseFloat(v.replace(/,/g, ''));
    return {
      value: v,
      numericValue: isNaN(numValue) ? null : numValue
    };
  }

  // テキスト型
  return { value: v, numericValue: null };
}

/**
 * 電話番号を正規化
 * @param {string} value - 入力値
 * @returns {string} ハイフン区切りの電話番号
 */
function normalizePhone(value) {
  if (!value) return '';

  // 数字以外を除去
  const digits = String(value).replace(/\D/g, '');

  if (digits.length === 10) {
    // 固定電話（03-1234-5678形式）
    return `${digits.slice(0, 2)}-${digits.slice(2, 6)}-${digits.slice(6)}`;
  } else if (digits.length === 11) {
    // 携帯電話（090-1234-5678形式）
    return `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7)}`;
  }

  return value; // 形式不明はそのまま
}

/**
 * 郵便番号を正規化
 * @param {string} value - 入力値
 * @returns {string} XXX-XXXX形式
 */
function normalizePostalCode(value) {
  if (!value) return '';

  const digits = String(value).replace(/\D/g, '');

  if (digits.length === 7) {
    return `${digits.slice(0, 3)}-${digits.slice(3)}`;
  }

  return value;
}

// ============================================
// 横持ち→縦持ち変換
// ============================================

/**
 * 横持ちデータを縦持ちに変換
 * @param {Object[]} horizontalData - 横持ちデータ（1行=1受診者、列=検査項目）
 * @param {Object} columnMapping - カラムマッピング {columnIndex: itemId, ...}
 * @param {Object} options - オプション
 * @returns {Object[]} 縦持ちデータ（1行=1検査結果）
 */
function convertHorizontalToVertical(horizontalData, columnMapping, options = {}) {
  const verticalData = [];

  // 基本項目のカラムインデックス
  const baseColumns = options.baseColumns || {
    visitDate: 0,
    name: 1,
    kana: 2,
    birthdate: 3,
    gender: 4,
    company: 5
  };

  for (const row of horizontalData) {
    // 基本情報を取得
    const baseInfo = {
      visitDate: row[baseColumns.visitDate],
      name: row[baseColumns.name],
      kana: row[baseColumns.kana],
      birthdate: row[baseColumns.birthdate],
      gender: row[baseColumns.gender],
      company: row[baseColumns.company] || ''
    };

    // 検査項目ごとに縦持ちレコードを作成
    for (const [colIndex, itemId] of Object.entries(columnMapping)) {
      const idx = parseInt(colIndex);
      const value = row[idx];

      // 空値はスキップ（オプションで変更可能）
      if (!options.includeEmpty && (value === '' || value === null || value === undefined)) {
        continue;
      }

      verticalData.push({
        ...baseInfo,
        itemId: itemId,
        value: value
      });
    }
  }

  return verticalData;
}

/**
 * 項目マスタから横持ち用カラムマッピングを生成
 * @param {string[]} csvHeaders - CSVヘッダー
 * @param {Object[]} itemMaster - 項目マスタデータ
 * @returns {Object} {columnIndex: itemId, ...}
 */
function generateItemColumnMapping(csvHeaders, itemMaster) {
  const mapping = {};

  // 項目名→項目IDのマップを作成
  const itemNameToId = {};
  for (const item of itemMaster) {
    itemNameToId[item.itemName] = item.itemId;
    // エイリアスも登録（カッコ内を除去したバージョンなど）
    const simpleName = item.itemName.replace(/[（(].*[）)]/g, '').trim();
    if (simpleName !== item.itemName) {
      itemNameToId[simpleName] = item.itemId;
    }
  }

  // CSVヘッダーとマッチング
  csvHeaders.forEach((header, index) => {
    const normalizedHeader = header.trim();
    if (itemNameToId[normalizedHeader]) {
      mapping[index] = itemNameToId[normalizedHeader];
    }
  });

  return mapping;
}

// ============================================
// バリデーション機能
// ============================================

/**
 * 行データをバリデーション
 * @param {Object} row - データ行
 * @param {Object[]} schema - スキーマ定義
 * @returns {Object} {valid: boolean, errors: string[]}
 */
function validateRow(row, schema) {
  const errors = [];

  for (const field of schema) {
    const value = row[field.id];

    // 必須チェック
    if (field.required && (value === undefined || value === null || value === '')) {
      errors.push(`必須項目「${field.name}」が未入力です`);
      continue;
    }

    // 値が存在する場合のみ形式チェック
    if (value !== undefined && value !== null && value !== '') {
      // 性別チェック
      if (field.id === 'gender') {
        const normalized = normalizeGender(value);
        if (!normalized) {
          errors.push(`「${field.name}」の値が不正です: ${value}`);
        }
      }

      // 生年月日チェック
      if (field.id === 'birthdate' || field.id === 'visitDate') {
        const normalized = normalizeBirthDate(value);
        if (!normalized) {
          errors.push(`「${field.name}」の日付形式が不正です: ${value}`);
        }
      }
    }
  }

  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * 重複チェック（受診者）
 * @param {string} name - 氏名
 * @param {Date} birthdate - 生年月日
 * @returns {Object|null} 既存の受診者データまたはnull
 */
function findDuplicatePatient(name, birthdate) {
  const sheet = getSheet(DB_CONFIG.SHEETS.PATIENT_MASTER);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const cols = COLUMN_DEFINITIONS.PATIENT_MASTER.columns;

  // 生年月日をYYYYMMDD形式に
  const targetBirthdate = formatDateYYYYMMDD(birthdate);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowName = row[cols.NAME];
    const rowBirthdate = formatDateYYYYMMDD(row[cols.BIRTHDATE]);

    if (rowName === name && rowBirthdate === targetBirthdate) {
      return {
        patientId: row[cols.PATIENT_ID],
        name: rowName,
        birthdate: row[cols.BIRTHDATE],
        rowIndex: i
      };
    }
  }

  return null;
}

// ============================================
// インポート処理
// ============================================

/**
 * 受診者データをインポート
 * @param {Object[]} data - インポートデータ配列
 * @param {Object} options - オプション
 * @param {boolean} options.skipErrors - エラー行スキップ
 * @param {boolean} options.allowDuplicates - 重複許可
 * @returns {Object} インポート結果
 */
function importPatients(data, options = {}) {
  const results = {
    success: 0,
    skipped: 0,
    errors: [],
    imported: []
  };

  const skipErrors = options.skipErrors !== false;
  const allowDuplicates = options.allowDuplicates || false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    try {
      // バリデーション
      const validation = validateRow(row, PATIENT_IMPORT_SCHEMA);
      if (!validation.valid) {
        if (skipErrors) {
          results.skipped++;
          results.errors.push({
            row: rowNum,
            type: 'validation',
            message: validation.errors.join('; ')
          });
          continue;
        } else {
          throw new Error(validation.errors.join('; '));
        }
      }

      // 正規化
      const normalizedData = {
        name: row.name,
        kana: row.kana || '',
        birthdate: normalizeBirthDate(row.birthdate),
        gender: normalizeGender(row.gender),
        phone: normalizePhone(row.phone || ''),
        postalCode: normalizePostalCode(row.postalCode || ''),
        address: row.address || '',
        email: row.email || '',
        company: row.company || ''
      };

      // 重複チェック
      if (!allowDuplicates) {
        const duplicate = findDuplicatePatient(normalizedData.name, normalizedData.birthdate);
        if (duplicate) {
          results.skipped++;
          results.errors.push({
            row: rowNum,
            type: 'duplicate',
            message: `既存データと重複: ${duplicate.patientId} (${duplicate.name})`
          });
          continue;
        }
      }

      // 登録
      const createResult = createPatient(normalizedData);
      if (createResult.success) {
        results.success++;
        results.imported.push({
          row: rowNum,
          patientId: createResult.patientId,
          name: normalizedData.name
        });
      } else {
        if (skipErrors) {
          results.skipped++;
          results.errors.push({
            row: rowNum,
            type: 'create',
            message: createResult.error
          });
        } else {
          throw new Error(createResult.error);
        }
      }

    } catch (e) {
      logError('importPatients', e);
      if (skipErrors) {
        results.skipped++;
        results.errors.push({
          row: rowNum,
          type: 'exception',
          message: e.message
        });
      } else {
        throw e;
      }
    }
  }

  return results;
}

/**
 * 検査結果データをインポート
 * @param {Object[]} data - インポートデータ配列（縦持ち形式）
 * @param {Object} options - オプション
 * @returns {Object} インポート結果
 */
function importTestResults(data, options = {}) {
  const results = {
    success: 0,
    skipped: 0,
    errors: [],
    visits: [],       // 作成した受診記録
    testResults: []   // 作成した検査結果
  };

  const skipErrors = options.skipErrors !== false;
  const examTypeId = options.examTypeId || DB_CONFIG.EXAM_TYPE.DOCK;

  // 受診者+受診日でグループ化
  const grouped = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const key = `${row.name}_${formatDateYYYYMMDD(row.visitDate)}`;

    if (!grouped[key]) {
      grouped[key] = {
        baseInfo: {
          name: row.name,
          kana: row.kana || '',
          birthdate: row.birthdate,
          gender: row.gender,
          company: row.company || '',
          visitDate: row.visitDate
        },
        items: [],
        rowIndices: []
      };
    }
    grouped[key].items.push({
      itemId: row.itemId,
      value: row.value,
      judgment: row.judgment || ''
    });
    grouped[key].rowIndices.push(i + 1);
  }

  // グループごとに処理
  for (const [key, group] of Object.entries(grouped)) {
    try {
      const baseInfo = group.baseInfo;

      // 受診者を検索または作成
      let patientId = null;
      const normalizedBirthdate = normalizeBirthDate(baseInfo.birthdate);
      const existing = findDuplicatePatient(baseInfo.name, normalizedBirthdate);

      if (existing) {
        patientId = existing.patientId;
      } else {
        // 新規作成
        const patientResult = createPatient({
          name: baseInfo.name,
          kana: baseInfo.kana,
          birthdate: normalizedBirthdate,
          gender: normalizeGender(baseInfo.gender),
          company: baseInfo.company
        });

        if (!patientResult.success) {
          throw new Error(`受診者作成失敗: ${patientResult.error}`);
        }
        patientId = patientResult.patientId;
      }

      // 受診記録を作成
      const visitDate = normalizeVisitDate(baseInfo.visitDate);
      const visitResult = createVisitRecord({
        patientId: patientId,
        examTypeId: examTypeId,
        visitDate: visitDate
      });

      if (!visitResult.success) {
        throw new Error(`受診記録作成失敗: ${visitResult.error}`);
      }

      results.visits.push({
        visitId: visitResult.visitId,
        patientId: patientId,
        name: baseInfo.name
      });

      // 検査結果を登録
      for (const item of group.items) {
        const normalized = normalizeTestValue(item.value, 'numeric');

        const testResult = createTestResult({
          visitId: visitResult.visitId,
          itemId: item.itemId,
          value: normalized.value,
          numericValue: normalized.numericValue,
          judgment: item.judgment || ''  // 判定は自動計算に任せる場合は空
        });

        if (testResult.success) {
          results.success++;
          results.testResults.push({
            resultId: testResult.resultId,
            itemId: item.itemId
          });
        } else {
          if (skipErrors) {
            results.skipped++;
            results.errors.push({
              rows: group.rowIndices,
              type: 'test_result',
              message: `検査結果登録失敗 (${item.itemId}): ${testResult.error}`
            });
          }
        }
      }

    } catch (e) {
      logError('importTestResults', e);
      if (skipErrors) {
        results.skipped += group.items.length;
        results.errors.push({
          rows: group.rowIndices,
          type: 'exception',
          message: e.message
        });
      } else {
        throw e;
      }
    }
  }

  return results;
}

/**
 * 横持ちCSVから検査結果をインポート（一括処理）
 * @param {string} csvContent - CSVコンテンツ
 * @param {Object} mappings - カラムマッピング
 * @param {Object} options - オプション
 * @returns {Object} インポート結果
 */
function importTestResultsFromHorizontalCsv(csvContent, mappings, options = {}) {
  // CSVパース
  const parsed = parseCsv(csvContent);
  if (parsed.error) {
    return { success: false, error: parsed.error };
  }

  // マッピング適用してオブジェクト配列に変換
  const horizontalData = parsed.rows.map(row => {
    const obj = {};
    for (const [colIndex, field] of Object.entries(mappings.baseColumns || {})) {
      obj[field] = row[parseInt(colIndex)];
    }
    return { ...obj, _raw: row };
  });

  // 横持ち→縦持ち変換
  const verticalData = convertHorizontalToVertical(
    parsed.rows,
    mappings.itemColumns || {},
    {
      baseColumns: mappings.baseColumnIndices || {},
      includeEmpty: false
    }
  );

  // インポート実行
  return importTestResults(verticalData, options);
}

// ============================================
// レポート生成
// ============================================

/**
 * インポート結果レポートを生成
 * @param {Object} results - インポート結果
 * @param {string} dataType - データ種別
 * @returns {Object} レポート
 */
function generateImportReport(results, dataType = 'PATIENT') {
  const report = {
    summary: {
      total: results.success + results.skipped,
      success: results.success,
      skipped: results.skipped,
      successRate: ((results.success / (results.success + results.skipped)) * 100).toFixed(1) + '%'
    },
    details: {
      imported: results.imported || results.visits || [],
      errors: results.errors.map(e => ({
        row: e.row || e.rows,
        type: e.type,
        message: e.message
      }))
    },
    timestamp: new Date().toLocaleString('ja-JP')
  };

  // エラー種別ごとの集計
  const errorSummary = {};
  for (const err of results.errors) {
    errorSummary[err.type] = (errorSummary[err.type] || 0) + 1;
  }
  report.summary.errorSummary = errorSummary;

  return report;
}

/**
 * レポートをCSV形式でエクスポート
 * @param {Object} report - レポート
 * @returns {string} CSV文字列
 */
function exportReportToCsv(report) {
  const lines = [];

  // サマリー
  lines.push('# インポートレポート');
  lines.push(`実行日時,${report.timestamp}`);
  lines.push(`合計,${report.summary.total}`);
  lines.push(`成功,${report.summary.success}`);
  lines.push(`スキップ,${report.summary.skipped}`);
  lines.push(`成功率,${report.summary.successRate}`);
  lines.push('');

  // エラー詳細
  if (report.details.errors.length > 0) {
    lines.push('# エラー詳細');
    lines.push('行,種別,メッセージ');
    for (const err of report.details.errors) {
      const row = Array.isArray(err.row) ? err.row.join('-') : err.row;
      lines.push(`${row},${err.type},"${err.message.replace(/"/g, '""')}"`);
    }
  }

  return lines.join('\n');
}

// ============================================
// Claude AI マッピング機能
// ============================================

/**
 * ヘッダーのハッシュ値を計算
 * @param {string[]} headers - ヘッダー配列
 * @returns {string} ハッシュ文字列
 */
function calculateHeadersHash(headers) {
  const str = headers.map(h => h.trim().toLowerCase()).sort().join('|');
  // 簡易ハッシュ（MD5など使えないため）
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return Math.abs(hash).toString(16);
}

/**
 * 不定形CSVのカラムマッピングをClaudeで推論
 * @param {string[]} csvHeaders - CSVのヘッダー行
 * @param {string[][]} sampleRows - サンプルデータ（最大3行）
 * @param {string} dataType - データ種別（PATIENT / TEST_RESULT）
 * @returns {Object} マッピング結果
 */
function inferCsvMapping(csvHeaders, sampleRows, dataType = 'PATIENT') {
  // ヘッダーが空の場合はエラー
  if (!csvHeaders || csvHeaders.length === 0) {
    return {
      success: false,
      error: 'CSVヘッダーが空です'
    };
  }

  // 対象スキーマを選択
  const targetSchema = dataType === 'PATIENT'
    ? PATIENT_IMPORT_SCHEMA
    : TEST_RESULT_HORIZONTAL_SCHEMA;

  // サンプル行を最大3行に制限
  const samples = sampleRows.slice(0, 3);

  // プロンプト構築
  const systemPrompt = `あなたは健診システムのデータマッピング専門家です。
CSVカラムをシステム項目に正確にマッピングしてください。

ルール:
1. カラム名の類似性を判断（名前/氏名/お名前 = name）
2. サンプルデータの形式を参考にする
3. 確信度が低い場合は低いconfidenceを返す
4. マッピングできないカラムはtarget: null
5. 必ず有効なJSON形式で出力すること
6. 検査項目らしきカラムは検査項目としてマーク

日本語で回答してください。`;

  const userMessage = `## CSVカラムとサンプルデータ

${csvHeaders.map((h, i) => {
  const sampleValues = samples.map(r => r[i] || '').filter(v => v).slice(0, 3);
  return `- カラム${i + 1}「${h}」: サンプル値 [${sampleValues.join(', ') || '(空)'}]`;
}).join('\n')}

## マッピング先システム項目

${targetSchema.map(s => `- ${s.id}: ${s.name}（${s.aliases ? s.aliases.join(', ') : ''}）${s.required ? '【必須】' : ''}`).join('\n')}

## 出力形式（以下のJSON形式で出力してください）

{
  "mappings": [
    {"csv_column": "CSVカラム名", "csv_index": 0, "target": "システム項目ID", "confidence": 0.95, "is_test_item": false},
    {"csv_column": "検査項目名", "csv_index": 5, "target": null, "confidence": 0.0, "is_test_item": true, "test_item_name": "BMI"}
  ],
  "value_transforms": {
    "性別": {"男": "M", "女": "F", "男性": "M", "女性": "F"}
  },
  "overall_confidence": 0.92,
  "notes": "推論に関する補足"
}`;

  try {
    // Claude API呼び出し
    const result = callClaudeApi(systemPrompt, userMessage, { max_tokens: 2048 });

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

    // ヘッダーハッシュを計算
    const headersHash = calculateHeadersHash(csvHeaders);

    return {
      success: true,
      headersHash: headersHash,
      mappings: mappingResult.mappings || [],
      valueTransforms: mappingResult.value_transforms || {},
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
  const testItemColumns = {};  // 検査項目カラム

  mappings.forEach(m => {
    if (m.is_test_item) {
      testItemColumns[m.csv_index] = m.test_item_name || m.csv_column;
    } else if (m.target) {
      indexToTarget[m.csv_index] = {
        target: m.target,
        column: m.csv_column
      };
    }
  });

  for (const row of rows) {
    const record = {};

    // 基本項目のマッピング
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

    // 検査項目カラムを別途格納
    record._testItems = {};
    for (const [idx, itemName] of Object.entries(testItemColumns)) {
      const value = row[parseInt(idx)];
      if (value !== '' && value !== null && value !== undefined) {
        record._testItems[itemName] = value;
      }
    }

    result.push(record);
  }

  return result;
}

// ============================================
// マッピングパターン管理
// ============================================

/**
 * マッピングパターンシートを取得（なければ作成）
 * @returns {Sheet} シート
 */
function getMappingPatternSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(MAPPING_PATTERN_SHEET.NAME);

  if (!sheet) {
    sheet = ss.insertSheet(MAPPING_PATTERN_SHEET.NAME);
    sheet.getRange(1, 1, 1, MAPPING_PATTERN_SHEET.HEADERS.length)
         .setValues([MAPPING_PATTERN_SHEET.HEADERS]);
    sheet.hideSheet();
  }

  return sheet;
}

/**
 * ヘッダーハッシュからマッピングパターンを検索
 * @param {string} headersHash - ヘッダーハッシュ
 * @param {string} dataType - データ種別
 * @returns {Object|null} パターンまたはnull
 */
function findMappingPattern(headersHash, dataType = 'PATIENT') {
  const sheet = getMappingPatternSheet();
  const data = sheet.getDataRange().getValues();
  const cols = MAPPING_PATTERN_SHEET.COLUMNS;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[cols.HEADERS_HASH] === headersHash && row[cols.DATA_TYPE] === dataType) {
      return {
        patternId: row[cols.PATTERN_ID],
        sourceName: row[cols.SOURCE_NAME],
        mappings: JSON.parse(row[cols.MAPPINGS_JSON] || '[]'),
        valueTransforms: JSON.parse(row[cols.VALUE_TRANSFORMS_JSON] || '{}'),
        useCount: row[cols.USE_COUNT],
        rowIndex: i + 1
      };
    }
  }

  return null;
}

/**
 * マッピングパターンを保存
 * @param {Object} pattern - パターン情報
 * @returns {Object} 保存結果
 */
function saveMappingPattern(pattern) {
  try {
    const sheet = getMappingPatternSheet();
    const cols = MAPPING_PATTERN_SHEET.COLUMNS;
    const now = new Date();

    // 既存パターンを検索
    const existing = findMappingPattern(pattern.headersHash, pattern.dataType);

    if (existing) {
      // 更新
      const rowIndex = existing.rowIndex;
      sheet.getRange(rowIndex, cols.MAPPINGS_JSON + 1).setValue(JSON.stringify(pattern.mappings));
      sheet.getRange(rowIndex, cols.VALUE_TRANSFORMS_JSON + 1).setValue(JSON.stringify(pattern.valueTransforms || {}));
      sheet.getRange(rowIndex, cols.USE_COUNT + 1).setValue(existing.useCount + 1);
      sheet.getRange(rowIndex, cols.UPDATED_AT + 1).setValue(now);

      return { success: true, action: 'updated', patternId: existing.patternId };
    } else {
      // 新規作成
      const patternId = `MP${now.getTime()}`;
      const newRow = [
        patternId,
        pattern.sourceName || '',
        pattern.headersHash,
        pattern.dataType,
        JSON.stringify(pattern.mappings),
        JSON.stringify(pattern.valueTransforms || {}),
        1,
        now,
        now
      ];

      sheet.appendRow(newRow);
      return { success: true, action: 'created', patternId: patternId };
    }

  } catch (e) {
    logError('saveMappingPattern', e);
    return { success: false, error: e.message };
  }
}

/**
 * マッピングパターンの使用回数を増加
 * @param {string} patternId - パターンID
 */
function incrementPatternUseCount(patternId) {
  const sheet = getMappingPatternSheet();
  const data = sheet.getDataRange().getValues();
  const cols = MAPPING_PATTERN_SHEET.COLUMNS;

  for (let i = 1; i < data.length; i++) {
    if (data[i][cols.PATTERN_ID] === patternId) {
      const currentCount = data[i][cols.USE_COUNT] || 0;
      sheet.getRange(i + 1, cols.USE_COUNT + 1).setValue(currentCount + 1);
      sheet.getRange(i + 1, cols.UPDATED_AT + 1).setValue(new Date());
      break;
    }
  }
}

// ============================================
// UI連携用API関数
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
      mappings: existingPattern.mappings,
      valueTransforms: existingPattern.valueTransforms,
      overallConfidence: 1.0,
      notes: `既存パターン「${existingPattern.sourceName}」を適用（使用回数: ${existingPattern.useCount + 1}）`
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
    // 横持ちの場合は縦持ちに変換
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

      for (const [itemName, value] of Object.entries(row._testItems || {})) {
        verticalData.push({
          ...baseInfo,
          itemId: itemName,  // 項目IDに変換が必要な場合は別途処理
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
      valueTransforms: mappings.valueTransforms
    });
  }

  // レポート生成
  const report = generateImportReport(results, dataType);

  return {
    success: true,
    results: results,
    report: report
  };
}
