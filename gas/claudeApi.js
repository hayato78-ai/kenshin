/**
 * Claude API 連携モジュール
 *
 * 機能:
 * - エラー診断・トラブルシューティング
 * - 所見文生成支援
 * - データ検証・異常値検出
 *
 * 設定方法:
 * 1. GASエディタ → プロジェクトの設定 → スクリプトプロパティ
 * 2. ANTHROPIC_API_KEY を追加
 */

// ============================================
// 設定・定数
// ============================================

const CLAUDE_CONFIG = {
  API_URL: 'https://api.anthropic.com/v1/messages',
  API_VERSION: '2023-06-01',
  MODEL: 'claude-sonnet-4-20250514',
  MAX_TOKENS: 2048,

  // システムプロンプト
  SYSTEM_PROMPTS: {
    DIAGNOSIS: `あなたは健診システムのトラブルシューティング専門家です。
GAS (Google Apps Script) とスプレッドシートの問題を診断します。
回答は簡潔に、具体的な解決策を提示してください。
日本語で回答してください。`,

    FINDINGS: `あなたは健康診断の所見文作成を支援する医療アシスタントです。
検査値と判定結果から、医師向けの所見文案を作成します。
以下のルールに従ってください：
- 医学的に正確な表現を使用
- 簡潔で分かりやすい文章
- 判定がC/Dの項目に焦点を当てる
- 「〜をお勧めします」「〜が必要です」などの推奨形式
日本語で回答してください。`,

    VALIDATION: `あなたは健診データの検証専門家です。
検査値の妥当性、入力ミスの可能性、異常値を分析します。
問題があれば具体的に指摘してください。
日本語で回答してください。`
  }
};

// ============================================
// API キー管理
// ============================================

/**
 * APIキーを取得
 * @returns {string|null} APIキー
 */
function getAnthropicApiKey() {
  return PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
}

/**
 * APIキーを設定（初回セットアップ用）
 * 実行後はこの関数内のキー文字列を削除すること
 */
function setAnthropicApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Claude API キー設定',
    'Anthropic APIキーを入力してください (sk-ant-...):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const apiKey = response.getResponseText().trim();

  if (!apiKey.startsWith('sk-ant-')) {
    ui.alert('エラー', 'APIキーの形式が正しくありません。sk-ant-で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }

  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', apiKey);
  ui.alert('完了', 'APIキーを保存しました。', ui.ButtonSet.OK);
  logInfo('Claude APIキーを設定しました');
}

/**
 * APIキーが設定されているか確認
 * @returns {boolean}
 */
function hasAnthropicApiKey() {
  const key = getAnthropicApiKey();
  return key && key.startsWith('sk-ant-');
}

// ============================================
// Claude API 呼び出し
// ============================================

/**
 * Claude APIを呼び出す（基本関数）
 * @param {string} systemPrompt - システムプロンプト
 * @param {string} userMessage - ユーザーメッセージ
 * @param {Object} options - オプション（model, max_tokens等）
 * @returns {Object} {success, content, error, usage}
 */
function callClaudeApi(systemPrompt, userMessage, options = {}) {
  const apiKey = getAnthropicApiKey();

  if (!apiKey) {
    return {
      success: false,
      content: null,
      error: 'APIキーが設定されていません。メニュー → 設定 → Claude APIキー設定 から設定してください。'
    };
  }

  const model = options.model || CLAUDE_CONFIG.MODEL;
  const maxTokens = options.max_tokens || CLAUDE_CONFIG.MAX_TOKENS;

  try {
    const response = UrlFetchApp.fetch(CLAUDE_CONFIG.API_URL, {
      method: 'post',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': CLAUDE_CONFIG.API_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: model,
        max_tokens: maxTokens,
        system: systemPrompt,
        messages: [{
          role: 'user',
          content: userMessage
        }]
      }),
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    const responseData = JSON.parse(response.getContentText());

    if (statusCode !== 200) {
      const errorMessage = responseData.error?.message || `HTTP ${statusCode}`;
      logError('callClaudeApi', new Error(errorMessage));
      return {
        success: false,
        content: null,
        error: `API エラー: ${errorMessage}`
      };
    }

    return {
      success: true,
      content: responseData.content[0].text,
      error: null,
      usage: responseData.usage
    };

  } catch (e) {
    logError('callClaudeApi', e);
    return {
      success: false,
      content: null,
      error: `通信エラー: ${e.message}`
    };
  }
}

// ============================================
// エラー診断機能
// ============================================

/**
 * CSV取り込みエラーを診断
 * @param {Object} context - エラーコンテキスト
 * @returns {Object} 診断結果
 */
function diagnoseCsvImportError(context) {
  const message = `
## CSV取り込みエラーの診断をお願いします

### 状況
${context.description || 'CSVを取り込んだが数値が入らない'}

### CSVファイル情報
- ファイル名: ${context.fileName || '不明'}
- ヘッダー: ${context.csvHeaders ? context.csvHeaders.join(', ') : '不明'}
- データ行数: ${context.rowCount || '不明'}

### 期待する動作
- 取り込み先シート: ${context.targetSheet || '労災二次検診_入力'}
- 期待するキー名: ${context.expectedKeys ? context.expectedKeys.join(', ') : '不明'}

### 実際の結果
${context.actualResult || '数値列が空になっている'}

### エラーメッセージ
${context.errorMessage || 'なし'}

原因と解決策を教えてください。
`;

  return callClaudeApi(CLAUDE_CONFIG.SYSTEM_PROMPTS.DIAGNOSIS, message);
}

/**
 * 一般的なエラーを診断
 * @param {string} errorDescription - エラーの説明
 * @param {Object} context - 追加コンテキスト
 * @returns {Object} 診断結果
 */
function diagnoseError(errorDescription, context = {}) {
  let message = `## エラー診断\n\n### 問題\n${errorDescription}\n`;

  if (context.functionName) {
    message += `\n### 発生箇所\n関数: ${context.functionName}\n`;
  }
  if (context.errorStack) {
    message += `\n### スタックトレース\n\`\`\`\n${context.errorStack}\n\`\`\`\n`;
  }
  if (context.inputData) {
    message += `\n### 入力データ\n\`\`\`json\n${JSON.stringify(context.inputData, null, 2)}\n\`\`\`\n`;
  }

  message += '\n原因の特定と解決策を提案してください。';

  return callClaudeApi(CLAUDE_CONFIG.SYSTEM_PROMPTS.DIAGNOSIS, message);
}

// ============================================
// 所見生成支援機能
// ============================================

/**
 * 検査結果から所見文を生成
 * @param {Object} patientData - 患者データ
 * @returns {Object} 生成結果
 */
function generateFindingsWithClaude(patientData) {
  const message = `
## 検査結果から所見文を作成してください

### 患者情報
- 氏名: ${patientData.name || '(非表示)'}
- 年齢: ${patientData.age || '不明'}歳
- 性別: ${patientData.gender || '不明'}

### 検査結果と判定
| 項目 | 値 | 判定 | 基準値 |
|------|-----|------|--------|
| HDLコレステロール | ${patientData.hdl || '-'} mg/dL | ${patientData.hdlJudgment || '-'} | ≥40 |
| LDLコレステロール | ${patientData.ldl || '-'} mg/dL | ${patientData.ldlJudgment || '-'} | <120 |
| 中性脂肪 | ${patientData.tg || '-'} mg/dL | ${patientData.tgJudgment || '-'} | <150 |
| 空腹時血糖 | ${patientData.fbs || '-'} mg/dL | ${patientData.fbsJudgment || '-'} | <100 |
| HbA1c | ${patientData.hba1c || '-'} % | ${patientData.hba1cJudgment || '-'} | <5.6 |
| 尿中アルブミン/Cre比 | ${patientData.acr || '-'} mg/g.cre | ${patientData.acrJudgment || '-'} | <30 |

### 超音波検査
- 心臓: 判定 ${patientData.cardiacJudgment || '-'}、所見: ${patientData.cardiacFindings || '未入力'}
- 頸動脈: 判定 ${patientData.carotidJudgment || '-'}、所見: ${patientData.carotidFindings || '未入力'}

### 依頼事項
1. 総合所見文（200文字以内）を作成してください
2. 特定保健指導の要否と理由
3. 精密検査が必要な項目があれば指摘
`;

  return callClaudeApi(CLAUDE_CONFIG.SYSTEM_PROMPTS.FINDINGS, message);
}

/**
 * 選択した患者の所見をClaudeで生成
 * サイドバーまたはダイアログから呼び出し
 */
function generateFindingsForSelectedPatient() {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== '労災二次検診_入力') {
    SpreadsheetApp.getUi().alert('エラー', '労災二次検診_入力シートを選択してください', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const activeRow = ss.getActiveRange().getRow();

  if (activeRow < 6) {
    SpreadsheetApp.getUi().alert('エラー', 'データ行（6行目以降）を選択してください', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // データ取得
  const rowData = sheet.getRange(activeRow, 1, 1, 18).getValues()[0];
  const gender = rowData[4] === '女性' ? 'F' : 'M';

  const patientData = {
    name: rowData[1],
    age: rowData[3],
    gender: rowData[4],
    cardiacJudgment: rowData[5],
    cardiacFindings: rowData[6],
    carotidJudgment: rowData[7],
    carotidFindings: rowData[8],
    hdl: rowData[12],
    ldl: rowData[13],
    tg: rowData[14],
    fbs: rowData[15],
    hba1c: rowData[16],
    acr: rowData[17]
  };

  // 判定を計算
  patientData.hdlJudgment = patientData.hdl ? judge('HDL_CHOLESTEROL', toNumber(patientData.hdl), gender) : '';
  patientData.ldlJudgment = patientData.ldl ? judge('LDL_CHOLESTEROL', toNumber(patientData.ldl), gender) : '';
  patientData.tgJudgment = patientData.tg ? judge('TRIGLYCERIDES', toNumber(patientData.tg), gender) : '';
  patientData.fbsJudgment = patientData.fbs ? judge('FASTING_GLUCOSE', toNumber(patientData.fbs), gender) : '';
  patientData.hba1cJudgment = patientData.hba1c ? judge('HBA1C', toNumber(patientData.hba1c), gender) : '';
  patientData.acrJudgment = patientData.acr ? judge('ACR', toNumber(patientData.acr), gender) : '';

  // 処理中ダイアログ
  const ui = SpreadsheetApp.getUi();

  // Claude API呼び出し
  const result = generateFindingsWithClaude(patientData);

  if (!result.success) {
    ui.alert('エラー', result.error, ui.ButtonSet.OK);
    return;
  }

  // 結果をダイアログで表示
  showClaudeResultDialog(patientData.name, result.content);
}

/**
 * Claude結果表示ダイアログ
 */
function showClaudeResultDialog(patientName, content) {
  const html = `
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
        .content {
          background: #f8f9fa;
          padding: 15px;
          border-radius: 8px;
          white-space: pre-wrap;
          max-height: 400px;
          overflow-y: auto;
        }
        .btn-container {
          text-align: right;
          margin-top: 20px;
        }
        .btn {
          padding: 10px 20px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
          margin-left: 10px;
        }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-secondary { background: #f1f3f4; color: #333; }
      </style>
    </head>
    <body>
      <h3>🤖 Claude による所見案 - ${patientName}</h3>
      <div class="content">${content.replace(/\n/g, '<br>')}</div>
      <div class="btn-container">
        <button class="btn btn-secondary" onclick="copyToClipboard()">コピー</button>
        <button class="btn btn-primary" onclick="google.script.host.close()">閉じる</button>
      </div>
      <script>
        function copyToClipboard() {
          const text = document.querySelector('.content').innerText;
          navigator.clipboard.writeText(text).then(() => {
            alert('クリップボードにコピーしました');
          });
        }
      </script>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Claude 所見生成');
}

// ============================================
// データ検証機能
// ============================================

/**
 * 検査データの妥当性をClaudeで検証
 * @param {Array<Object>} patients - 患者データ配列
 * @returns {Object} 検証結果
 */
function validateDataWithClaude(patients) {
  // データをテーブル形式に変換
  let dataTable = '| No | 名前 | HDL | LDL | TG | FBS | HbA1c | ACR |\n';
  dataTable += '|-----|------|-----|-----|-----|-----|-------|-----|\n';

  for (const p of patients.slice(0, 20)) { // 最大20件
    dataTable += `| ${p.no || '-'} | ${p.name || '-'} | ${p.hdl || '-'} | ${p.ldl || '-'} | ${p.tg || '-'} | ${p.fbs || '-'} | ${p.hba1c || '-'} | ${p.acr || '-'} |\n`;
  }

  const message = `
## 健診データの検証をお願いします

以下のデータに異常値、入力ミスの可能性、論理的矛盾がないかチェックしてください。

### データ
${dataTable}

### チェック項目
1. 生理的にありえない値（例: HDL > 200, HbA1c > 15）
2. 入力ミスの可能性（桁違い、小数点位置）
3. 項目間の矛盾（例: FBS正常なのにHbA1c高値）
4. 欠損データの影響

問題があれば具体的に指摘し、確認すべき点を教えてください。
`;

  return callClaudeApi(CLAUDE_CONFIG.SYSTEM_PROMPTS.VALIDATION, message);
}

/**
 * 入力シートのデータを検証
 */
function validateInputSheetWithClaude() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('労災二次検診_入力');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', '入力シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert('エラー', 'データがありません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // データ取得
  const data = sheet.getRange(6, 1, lastRow - 5, 18).getValues();
  const patients = data.map(row => ({
    no: row[0],
    name: row[1],
    age: row[3],
    gender: row[4],
    hdl: row[12],
    ldl: row[13],
    tg: row[14],
    fbs: row[15],
    hba1c: row[16],
    acr: row[17]
  })).filter(p => p.name);

  if (patients.length === 0) {
    SpreadsheetApp.getUi().alert('エラー', '検証対象のデータがありません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const result = validateDataWithClaude(patients);

  if (!result.success) {
    SpreadsheetApp.getUi().alert('エラー', result.error, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  showClaudeResultDialog('データ検証', result.content);
}

// ============================================
// ヘルプ・チャット機能
// ============================================

/**
 * Claudeに質問するダイアログを表示
 */
function showClaudeHelpDialog() {
  const html = `
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
        }
        h3 {
          margin: 0 0 15px 0;
          color: #1a73e8;
        }
        textarea {
          width: 100%;
          height: 120px;
          padding: 10px;
          border: 1px solid #ddd;
          border-radius: 4px;
          font-size: 13px;
          resize: vertical;
          box-sizing: border-box;
        }
        .response {
          background: #f8f9fa;
          padding: 15px;
          border-radius: 8px;
          margin-top: 15px;
          min-height: 100px;
          max-height: 250px;
          overflow-y: auto;
          white-space: pre-wrap;
          display: none;
        }
        .btn-container {
          margin-top: 15px;
        }
        .btn {
          padding: 10px 20px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
          margin-right: 10px;
        }
        .btn-primary { background: #1a73e8; color: white; }
        .btn-secondary { background: #f1f3f4; color: #333; }
        .btn:disabled { background: #ccc; cursor: not-allowed; }
        .loading { color: #666; font-style: italic; }
      </style>
    </head>
    <body>
      <h3>🤖 Claudeに質問</h3>
      <p style="color: #666; font-size: 12px;">健診システムについて何でも質問できます</p>

      <textarea id="question" placeholder="例: CSVの取り込みがうまくいかない&#10;例: HDLの判定基準を教えて&#10;例: Excel出力でエラーが出る"></textarea>

      <div class="btn-container">
        <button id="askBtn" class="btn btn-primary" onclick="askClaude()">質問する</button>
        <button class="btn btn-secondary" onclick="google.script.host.close()">閉じる</button>
      </div>

      <div id="response" class="response"></div>

      <script>
        function askClaude() {
          const question = document.getElementById('question').value.trim();
          if (!question) {
            alert('質問を入力してください');
            return;
          }

          document.getElementById('askBtn').disabled = true;
          document.getElementById('response').style.display = 'block';
          document.getElementById('response').innerHTML = '<span class="loading">考え中...</span>';

          google.script.run
            .withSuccessHandler(function(result) {
              document.getElementById('askBtn').disabled = false;
              if (result.success) {
                document.getElementById('response').innerHTML = result.content.replace(/\\n/g, '<br>');
              } else {
                document.getElementById('response').innerHTML = '<span style="color: red;">エラー: ' + result.error + '</span>';
              }
            })
            .withFailureHandler(function(error) {
              document.getElementById('askBtn').disabled = false;
              document.getElementById('response').innerHTML = '<span style="color: red;">エラー: ' + error.message + '</span>';
            })
            .askClaudeQuestion(question);
        }
      </script>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Claude ヘルプ');
}

/**
 * Claudeに質問する
 * @param {string} question - 質問
 * @returns {Object} 回答結果
 */
function askClaudeQuestion(question) {
  const systemPrompt = `あなたは健診結果入力システムのヘルプアシスタントです。
このシステムは以下の機能を持っています：
- CSV取り込み（BML形式、標準形式）
- 判定処理（A/B/C/D判定）
- Excel出力（テンプレートへの転記）
- 労災二次検診対応（超音波検査、所見入力）

ユーザーの質問に対して、簡潔で実用的な回答を提供してください。
日本語で回答してください。`;

  return callClaudeApi(systemPrompt, question);
}
