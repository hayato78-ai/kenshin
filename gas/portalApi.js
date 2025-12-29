/**
 * portalApi.gs - Webアプリポータル用 API関数
 * CDmedical 健診結果管理システム
 *
 * フロントエンドから google.script.run 経由で呼び出される関数群
 */

// ============================================================
// スプレッドシート接続
// ============================================================

/**
 * スプレッドシートを取得（Webアプリ対応）
 * @returns {Spreadsheet} スプレッドシート
 */
function getPortalSpreadsheet() {
  // DB_CONFIGが定義されている場合はそちらを使用（Config.gsで定義）
  if (typeof DB_CONFIG !== 'undefined' && DB_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(DB_CONFIG.SPREADSHEET_ID);
  }

  // フォールバック: 直接IDで取得
  const SPREADSHEET_ID = '16KtctyT2gd7oJZdcu84kUtuP-D9jB9KtLxxzxXx_wdk';
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    console.error('直接ID取得失敗:', e);
  }

  throw new Error('スプレッドシートに接続できません。');
}

/**
 * スプレッドシートIDを設定（初回のみ実行）
 * GASエディタから手動で実行してください
 */
function setupSpreadsheetId() {
  const ss = getPortalSpreadsheet();
  if (ss) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SPREADSHEET_ID', ss.getId());
    console.log('スプレッドシートID設定完了: ' + ss.getId());
    return { success: true, id: ss.getId() };
  }
  return { success: false, error: 'スプレッドシートが見つかりません' };
}

// ============================================================
// ユーティリティ関数
// ============================================================

/**
 * 文字列を正規化（スペースを統一、大文字小文字を統一）
 * @param {*} str - 入力文字列
 * @returns {string} 正規化された文字列
 */
function normalizeString(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\s　]+/g, ' ')  // 全角・半角スペースを半角スペースに統一
    .trim()
    .toLowerCase();
}

/**
 * カナを正規化（全角カナに統一）
 * @param {*} str - 入力文字列
 * @returns {string} 正規化されたカナ
 */
function normalizeKana(str) {
  if (str === null || str === undefined) return '';
  let result = String(str).replace(/[\s　]+/g, ' ').trim();
  // 半角カナを全角カナに変換
  const kanaMap = {
    'ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ',
    'ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ','ｺ':'コ',
    'ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ',
    'ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ','ﾄ':'ト',
    'ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ',
    'ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ','ﾎ':'ホ',
    'ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ',
    'ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ',
    'ﾗ':'ラ','ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ',
    'ﾜ':'ワ','ｦ':'ヲ','ﾝ':'ン',
    'ｧ':'ァ','ｨ':'ィ','ｩ':'ゥ','ｪ':'ェ','ｫ':'ォ',
    'ｯ':'ッ','ｬ':'ャ','ｭ':'ュ','ｮ':'ョ',
    'ﾞ':'゛','ﾟ':'゜','ｰ':'ー'
  };
  for (const [half, full] of Object.entries(kanaMap)) {
    result = result.replace(new RegExp(half, 'g'), full);
  }
  return result;
}

// ============================================================
// 受診者管理 API（v6 - 正規化構造対応）
// ============================================================

// ============================================================
// 検査結果シート列定義（Option B: 1患者1行・非正規化）
// ============================================================

/**
 * 検査結果シートの列定義
 * ※ BMLコードはmapping_template_new.jsonと完全一致
 * 更新日: 2025-12-23
 */
const LAB_RESULT_COLUMNS = {
  // 基本情報（固定列）
  BASE: [
    { col: 0, id: 'PATIENT_ID', name: '受診者ID', type: 'key' },
    { col: 1, id: 'KARTE_NO', name: 'カルテNo', type: 'info' },
    { col: 2, id: 'EXAM_DATE', name: '受診日', type: 'date' },
    { col: 3, id: 'BML_PATIENT_ID', name: 'BML患者ID', type: 'info' },
    { col: 4, id: 'CSV_IMPORT_DATE', name: 'CSV取込日時', type: 'datetime' }
  ],
  // 検査項目（BMLコードはmapping_template_new.json準拠）
  ITEMS: [
    // ===== 血液一般 =====
    { id: 'WBC', name: '白血球数', bmlCode: '0000301', unit: '/μL', category: '血液学検査' },
    { id: 'RBC', name: '赤血球数', bmlCode: '0000302', unit: '万/μL', category: '血液学検査' },
    { id: 'HEMOGLOBIN', name: '血色素量', bmlCode: '0000303', unit: 'g/dL', category: '血液学検査' },
    { id: 'HEMATOCRIT', name: 'ヘマトクリット', bmlCode: '0000304', unit: '%', category: '血液学検査' },
    { id: 'PLATELET', name: '血小板', bmlCode: '0000308', unit: '万/μL', category: '血液学検査' },
    { id: 'MCV', name: 'MCV', bmlCode: '0000305', unit: 'fL', category: '血液学検査' },
    { id: 'MCH', name: 'MCH', bmlCode: '0000306', unit: 'pg', category: '血液学検査' },
    { id: 'MCHC', name: 'MCHC', bmlCode: '0000307', unit: '%', category: '血液学検査' },
    // ===== 白血球像 =====
    { id: 'NEUT', name: '好中球', bmlCode: '0001885', unit: '%', category: '血液学検査' },
    { id: 'BASO', name: '好塩基球', bmlCode: '0001881', unit: '%', category: '血液学検査' },
    { id: 'EOS', name: '好酸球', bmlCode: '0001882', unit: '%', category: '血液学検査' },
    { id: 'LYMPHO', name: 'リンパ球', bmlCode: '0001889', unit: '%', category: '血液学検査' },
    { id: 'MONO', name: '単球', bmlCode: '0001886', unit: '%', category: '血液学検査' },
    // ===== 生化学 =====
    { id: 'TOTAL_PROTEIN', name: '総蛋白', bmlCode: '0000401', unit: 'g/dL', category: '肝胆膵機能' },
    { id: 'ALBUMIN', name: 'アルブミン', bmlCode: '0000417', unit: 'g/dL', category: '肝胆膵機能' },
    { id: 'AST', name: 'AST(GOT)', bmlCode: '0000481', unit: 'U/L', category: '肝胆膵機能' },
    { id: 'ALT', name: 'ALT(GPT)', bmlCode: '0000482', unit: 'U/L', category: '肝胆膵機能' },
    { id: 'GGT', name: 'γ-GTP', bmlCode: '0000484', unit: 'U/L', category: '肝胆膵機能' },
    { id: 'ALP', name: 'ALP(IFCC)', bmlCode: '0013067', unit: 'U/L', category: '肝胆膵機能' },
    { id: 'LDH', name: 'LDH', bmlCode: '0000497', unit: 'U/L', category: '肝胆膵機能' },
    { id: 'T_BIL', name: '総ビリルビン', bmlCode: '0000472', unit: 'mg/dL', category: '肝胆膵機能' },
    // ===== 脂質 =====
    { id: 'TOTAL_CHOLESTEROL', name: '総コレステロール', bmlCode: '0000453', unit: 'mg/dL', category: '脂質検査' },
    { id: 'TG', name: '中性脂肪', bmlCode: '0000454', unit: 'mg/dL', category: '脂質検査' },
    { id: 'HDL_C', name: 'HDLコレステロール', bmlCode: '0000460', unit: 'mg/dL', category: '脂質検査' },
    { id: 'LDL_C', name: 'LDLコレステロール', bmlCode: '0000410', unit: 'mg/dL', category: '脂質検査' },
    // ===== 糖代謝 =====
    { id: 'FBS', name: '空腹時血糖', bmlCode: '0000503', unit: 'mg/dL', category: '糖代謝' },
    { id: 'HBA1C', name: 'HbA1c', bmlCode: '0003317', unit: '%', category: '糖代謝' },
    // ===== 腎機能 =====
    { id: 'CREATININE', name: 'クレアチニン', bmlCode: '0000413', unit: 'mg/dL', category: '腎機能' },
    { id: 'BUN', name: '尿素窒素', bmlCode: '0000491', unit: 'mg/dL', category: '腎機能' },
    { id: 'EGFR', name: 'eGFR', bmlCode: '0002696', unit: 'mL/min', category: '腎機能' },
    // ===== その他 =====
    { id: 'UA', name: '尿酸', bmlCode: '0000407', unit: 'mg/dL', category: '腎機能' },
    { id: 'CK', name: 'CK', bmlCode: '0003845', unit: 'U/L', category: '肝胆膵機能' },
    { id: 'NA', name: 'Na', bmlCode: '0003550', unit: 'mEq/L', category: '腎機能' },
    { id: 'K', name: 'K', bmlCode: '0000421', unit: 'mEq/L', category: '腎機能' },
    { id: 'CL', name: 'Cl', bmlCode: '0000425', unit: 'mEq/L', category: '腎機能' },
    { id: 'CRP', name: 'CRP定量', bmlCode: '0000658', unit: 'mg/dL', category: '血液学検査' },
    // ===== 身体測定（手入力） =====
    { id: 'HEIGHT', name: '身長', bmlCode: null, unit: 'cm', category: '身体測定' },
    { id: 'WEIGHT', name: '体重', bmlCode: null, unit: 'kg', category: '身体測定' },
    { id: 'BMI', name: 'BMI', bmlCode: null, unit: '', category: '身体測定' },
    { id: 'WAIST_M', name: '腹囲', bmlCode: null, unit: 'cm', category: '身体測定' },
    // ===== 血圧（手入力） =====
    { id: 'BP_SYSTOLIC_1', name: '収縮期血圧', bmlCode: null, unit: 'mmHg', category: '血圧' },
    { id: 'BP_DIASTOLIC_1', name: '拡張期血圧', bmlCode: null, unit: 'mmHg', category: '血圧' },
    // ★ 血圧2回目対応: 新規フィールド追加
    { id: 'BP_SYSTOLIC_2', name: '収縮期血圧_2回目', bmlCode: null, unit: 'mmHg', category: '血圧' },
    { id: 'BP_DIASTOLIC_2', name: '拡張期血圧_2回目', bmlCode: null, unit: 'mmHg', category: '血圧' },
    { id: 'PULSE', name: '脈拍', bmlCode: null, unit: 'bpm', category: '血圧' },
    // ===== 尿検査 =====
    { id: 'URINE_PROTEIN', name: '尿蛋白', bmlCode: '0003891', unit: '', category: '尿検査' },
    { id: 'URINE_GLUCOSE', name: '尿糖', bmlCode: '0000055', unit: '', category: '尿検査' },
    { id: 'URINE_OCCULT_BLOOD', name: '尿潜血', bmlCode: '0003893', unit: '', category: '尿検査' },
    // ===== 便潜血 =====
    { id: 'FOBT_1', name: '便潜血1回目', bmlCode: '0000905', unit: '', category: '消化器' },
    { id: 'FOBT_2', name: '便潜血2回目', bmlCode: '0000911', unit: '', category: '消化器' },
    // ===== 感染症 =====
    { id: 'HBS_AG', name: 'HBs抗原', bmlCode: '0004821', unit: '', category: '感染症' },
    { id: 'HCV_AB', name: 'HCV抗体', bmlCode: '0004822', unit: '', category: '感染症' },
    // ===== 追加項目（2025-12-25） =====
    // ===== 尿検査（詳細） =====
    { id: 'URINE_GLUCOSE_QUAL', name: '尿糖(定性)', bmlCode: '0000701', unit: '', category: '尿検査' },
    { id: 'URINE_PROTEIN_QUAL', name: '尿蛋白(定性)', bmlCode: '0000703', unit: '', category: '尿検査' },
    { id: 'URINE_OB_QUAL', name: '尿潜血(定性)', bmlCode: '0000705', unit: '', category: '尿検査' },
    { id: 'UROBILINOGEN', name: 'ウロビリノーゲン', bmlCode: '0000707', unit: '', category: '尿検査' },
    { id: 'URINE_PH', name: '尿PH', bmlCode: '0000709', unit: '', category: '尿検査' },
    { id: 'URINE_BILIRUBIN', name: '尿ビリルビン', bmlCode: '0000711', unit: '', category: '尿検査' },
    { id: 'ACETONE', name: 'アセトン体', bmlCode: '0000713', unit: '', category: '尿検査' },
    // ===== 尿沈渣 =====
    { id: 'URINE_SED_WBC', name: '尿沈渣白血球', bmlCode: '0000721', unit: '', category: '尿沈渣' },
    { id: 'URINE_SED_RBC', name: '尿沈渣赤血球', bmlCode: '0000723', unit: '', category: '尿沈渣' },
    { id: 'URINE_SED_EPI', name: '尿沈渣扁平上皮', bmlCode: '0000725', unit: '', category: '尿沈渣' },
    { id: 'URINE_SED_BACT', name: '尿沈渣細菌', bmlCode: '0000727', unit: '', category: '尿沈渣' },
    // ===== 便検査（詳細） =====
    { id: 'FOBT_HB_1', name: '便ヘモグロビン1回目', bmlCode: '0002151', unit: '', category: '消化器' },
    { id: 'FOBT_HB_2', name: '便ヘモグロビン2回目', bmlCode: '0002152', unit: '', category: '消化器' },
    // ===== 凝固 =====
    { id: 'PT', name: 'プロトロンビン時間(PT)', bmlCode: '0002091', unit: '秒', category: '凝固' },
    { id: 'APTT', name: 'APTT', bmlCode: '0002093', unit: '秒', category: '凝固' },
    // ===== 肝胆膵機能（追加） =====
    { id: 'CHE', name: 'コリンエステラーゼ(Ch-E)', bmlCode: '0000495', unit: 'U/L', category: '肝胆膵機能' },
    { id: 'AMYLASE', name: '血清アミラーゼ', bmlCode: '0000501', unit: 'U/L', category: '肝胆膵機能' },
    // ===== 電解質（追加） =====
    { id: 'CA', name: 'カルシウム(Ca)', bmlCode: '0000429', unit: 'mg/dL', category: '電解質' },
    // ===== 鉄代謝 =====
    { id: 'FE', name: '血清鉄(Fe)', bmlCode: '0000433', unit: 'μg/dL', category: '鉄代謝' },
    { id: 'TIBC', name: '総鉄結合能(TIBC)', bmlCode: '0000435', unit: 'μg/dL', category: '鉄代謝' },
    // ===== 自己抗体 =====
    { id: 'RF', name: 'リウマトイド因子定量', bmlCode: '0002117', unit: 'IU/mL', category: '自己抗体' },
    // ===== 腫瘍マーカー =====
    { id: 'PSA', name: 'PSA', bmlCode: '0005005', unit: 'ng/mL', category: '腫瘍マーカー' },
    { id: 'CEA', name: 'CEA', bmlCode: '0005001', unit: 'ng/mL', category: '腫瘍マーカー' },
    { id: 'CA19_9', name: 'CA19-9', bmlCode: '0005003', unit: 'U/mL', category: '腫瘍マーカー' },
    { id: 'CA125', name: 'CA125', bmlCode: '0005007', unit: 'U/mL', category: '腫瘍マーカー' },
    { id: 'NSE', name: 'NSE', bmlCode: '0005009', unit: 'ng/mL', category: '腫瘍マーカー' },
    { id: 'ELASTASE1', name: 'エラスターゼ1', bmlCode: '0005011', unit: 'ng/dL', category: '腫瘍マーカー' },
    { id: 'ANTI_P53', name: '抗p53抗体', bmlCode: '0005013', unit: 'U/mL', category: '腫瘍マーカー' },
    { id: 'CYFRA21_1', name: 'CYFRA21-1', bmlCode: '0005015', unit: 'ng/mL', category: '腫瘍マーカー' },
    { id: 'SCC', name: 'SCC', bmlCode: '0005017', unit: 'ng/mL', category: '腫瘍マーカー' },
    { id: 'PROGRP', name: 'ProGRP', bmlCode: '0005019', unit: 'pg/mL', category: '腫瘍マーカー' },
    { id: 'AFP', name: 'AFP', bmlCode: '0005021', unit: 'ng/mL', category: '腫瘍マーカー' },
    { id: 'PIVKA2', name: 'PIVKA II', bmlCode: '0005023', unit: 'mAU/mL', category: '腫瘍マーカー' },
    // ===== 感染症（追加） =====
    { id: 'TPHA', name: 'TPHA', bmlCode: '0006001', unit: '', category: '感染症' },
    { id: 'RPR', name: 'RPR定性', bmlCode: '0006003', unit: '', category: '感染症' },
    { id: 'HBS_AG_2', name: 'HBs抗原(定性)', bmlCode: '0006005', unit: '', category: '感染症' },
    { id: 'HBS_AB', name: 'HBs抗体', bmlCode: '0006007', unit: '', category: '感染症' },
    { id: 'HCV_AB_2', name: 'HCV抗体(定性)', bmlCode: '0006009', unit: '', category: '感染症' },
    { id: 'HIV', name: 'HIV-1抗体', bmlCode: '0006011', unit: '', category: '感染症' },
    { id: 'STS', name: 'STS定性', bmlCode: '0006013', unit: '', category: '感染症' },
    { id: 'TP_AB', name: '梅毒トレポネーマ抗体定性', bmlCode: '0006015', unit: '', category: '感染症' },
    // ===== 心臓・甲状腺 =====
    { id: 'NT_PROBNP', name: 'NT-proBNP', bmlCode: '0004001', unit: 'pg/mL', category: '心臓マーカー' },
    { id: 'FT3', name: 'FT3', bmlCode: '0004003', unit: 'pg/mL', category: '甲状腺' },
    { id: 'FT4', name: 'FT4', bmlCode: '0004005', unit: 'ng/dL', category: '甲状腺' },
    { id: 'TSH', name: 'TSH', bmlCode: '0004007', unit: 'μIU/mL', category: '甲状腺' },
    // ===== 血液型 =====
    { id: 'BLOOD_TYPE_ABO', name: '血液型ABO式', bmlCode: '0007001', unit: '', category: '血液型' },
    { id: 'BLOOD_TYPE_RH', name: '血液型Rho・D', bmlCode: '0007003', unit: '', category: '血液型' },
    // ===== ピロリ =====
    { id: 'UREA_BREATH', name: '尿素呼気試験', bmlCode: '0008001', unit: '', category: 'ピロリ菌' },
    { id: 'HP_AB', name: 'ピロリ抗体', bmlCode: '0008003', unit: '', category: 'ピロリ菌' }
  ]
};

/**
 * 検査結果シートのヘッダー配列を生成
 * @returns {Array<string>} ヘッダー名配列
 */
function getLabResultHeaders() {
  const headers = [];
  // 基本情報
  for (const col of LAB_RESULT_COLUMNS.BASE) {
    headers.push(col.name);
  }
  // 検査項目
  for (const item of LAB_RESULT_COLUMNS.ITEMS) {
    headers.push(item.name);
  }
  return headers;
}

/**
 * 検査結果シートのヘッダーを初期化/更新
 * @returns {Object} 結果
 */
function initLabResultSheetHeaders() {
  try {
    const ss = getPortalSpreadsheet();
    let sheet = ss.getSheetByName('検査結果');

    if (!sheet) {
      sheet = ss.insertSheet('検査結果');
    }

    const headers = getLabResultHeaders();
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e8f0fe');

    // 列幅調整
    sheet.setColumnWidth(1, 100); // 受診者ID
    sheet.setColumnWidth(2, 80);  // カルテNo
    sheet.setColumnWidth(3, 100); // 受診日

    return {
      success: true,
      message: `ヘッダー初期化完了（${headers.length}列）`,
      headers: headers
    };
  } catch (e) {
    console.error('initLabResultSheetHeaders error:', e);
    return { success: false, error: e.message };
  }
}

// 列定義（17列構造 - カルテNo追加版）
const PATIENT_COL = {
  PATIENT_ID: 0,      // A: 受診者ID (内部ID: P-00001形式) ※UIで非表示
  KARTE_NO: 1,        // B: カルテNo (クリニック患者ID: 6桁) ★CSV紐付け用
  STATUS: 2,          // C: ステータス
  VISIT_DATE: 3,      // D: 受診日
  NAME: 4,            // E: 氏名
  KANA: 5,            // F: カナ
  GENDER: 6,          // G: 性別
  BIRTHDATE: 7,       // H: 生年月日
  AGE: 8,             // I: 年齢
  COURSE: 9,          // J: 受診コース
  COMPANY: 10,        // K: 事業所名
  DEPARTMENT: 11,     // L: 所属
  OVERALL_JUDGMENT: 12, // M: 総合判定
  CSV_IMPORT_DATE: 13,  // N: CSV取込日時
  UPDATED_AT: 14,     // O: 最終更新日時
  EXPORT_DATE: 15,    // P: 出力日時
  BML_PATIENT_ID: 16  // Q: BML患者ID（BML通信用・トレーサビリティ）
};

const VISIT_COL = {
  VISIT_ID: 0,      // A: 受診ID (YYYYMMDD-NNN形式)
  PATIENT_ID: 1,    // B: 受診者ID (FK)
  EXAM_TYPE: 2,     // C: 検診種別
  COURSE: 3,        // D: コース
  VISIT_DATE: 4,    // E: 受診日
  AGE: 5,           // F: 年齢
  JUDGMENT: 6,      // G: 総合判定
  DOCTOR_NOTES: 7,  // H: 医師所見
  STATUS: 8,        // I: ステータス
  CSV_IMPORTED: 9,  // J: CSV取込日時
  CREATED_AT: 10,   // K: 作成日時
  UPDATED_AT: 11,   // L: 更新日時
  OUTPUT_AT: 12     // M: 出力日時
};

/**
 * 受診者を検索（v6 - 正規化構造対応）
 *
 * 【機能】
 * 1. 受診者マスタと受診記録を結合して検索
 * 2. 受診者ID、氏名、カナで部分一致検索
 * 3. 最新の受診記録情報を付与して返却
 *
 * @param {Object} criteria - 検索条件 {patientId, name, kana}
 * @returns {Object} 検索結果
 */
function portalSearchPatients(criteria) {
  const debugInfo = {
    version: '2025-12-25-v7-optimized',
    criteriaReceived: criteria ? JSON.stringify(criteria) : 'null',
    timestamp: new Date().toISOString()
  };

  try {
    const ss = getPortalSpreadsheet();
    debugInfo.spreadsheetId = ss.getId();

    // 受診者マスタ取得
    const patientSheet = ss.getSheetByName('受診者マスタ');
    if (!patientSheet) {
      return { success: false, error: '受診者マスタシートが見つかりません', data: [], _debug: debugInfo };
    }

    const lastRow = patientSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [], count: 0, _debug: debugInfo };
    }

    // ★ 最適化: 検索用列のみ取得（A〜K列 = 11列）
    // 旧: getDataRange().getValues() → 17列全取得
    // 新: A(受診者ID), B(カルテNo), C(ステータス), D(受診日), E(氏名), F(カナ),
    //     G(性別), H(生年月日), I(年齢), J(コース), K(所属企業)
    const patientData = patientSheet.getRange(1, 1, lastRow, 11).getValues();
    debugInfo.patientRows = patientData.length;
    debugInfo.columnsLoaded = 11;  // 17→11列に削減

    // 検索条件を正規化
    const searchId = criteria && criteria.patientId ? normalizeString(criteria.patientId) : '';
    const searchName = criteria && criteria.name ? normalizeString(criteria.name) : '';
    const searchKana = criteria && criteria.kana ? normalizeKana(criteria.kana) : '';

    debugInfo.normalizedCriteria = { searchId, searchName, searchKana };

    // ★ Phase 1: 検索フィルタリング（受診者マスタのみ）
    const matchedRows = [];  // ヒットした行インデックスを記録

    for (let i = 1; i < patientData.length; i++) {
      const row = patientData[i];

      // 空行スキップ
      if (!row[PATIENT_COL.PATIENT_ID]) continue;

      const patientId = safeString(row[PATIENT_COL.PATIENT_ID]);
      const karteNo = safeString(row[PATIENT_COL.KARTE_NO]);
      const name = safeString(row[PATIENT_COL.NAME]);
      const kana = safeString(row[PATIENT_COL.KANA]);

      // 検索条件でフィルタ（カルテNo/受診者ID, 氏名, カナ）
      let match = true;

      // searchIdはカルテNoまたは受診者IDで検索
      if (searchId) {
        const matchPatientId = normalizeString(patientId).includes(searchId);
        const matchKarteNo = normalizeString(karteNo).includes(searchId);
        if (!matchPatientId && !matchKarteNo) {
          match = false;
        }
      }
      if (searchName && match && !normalizeString(name).includes(searchName)) {
        match = false;
      }
      if (searchKana && match && !normalizeKana(kana).includes(searchKana)) {
        match = false;
      }

      if (match) {
        matchedRows.push({
          rowIndex: i,
          patientId: patientId,
          row: row
        });
      }

      // 最大件数制限
      if (matchedRows.length >= 100) break;
    }

    debugInfo.matchedCount = matchedRows.length;

    // ★ Phase 2: ヒットした患者の受診記録を取得（遅延読み込み）
    let visitMap = {};
    if (matchedRows.length > 0) {
      const visitSheet = ss.getSheetByName('受診記録');
      debugInfo.visitSheetFound = !!visitSheet;

      if (visitSheet && visitSheet.getLastRow() >= 2) {
        // ヒットした患者IDのセットを作成
        const matchedPatientIds = new Set(matchedRows.map(m => m.patientId));

        // ★ 最適化: 受診記録も必要な列のみ取得（A〜H列 = 8列）
        const visitLastRow = visitSheet.getLastRow();
        const visitData = visitSheet.getRange(1, 1, visitLastRow, 8).getValues();
        debugInfo.visitRows = visitData.length;
        debugInfo.visitColumnsLoaded = 8;

        for (let i = 1; i < visitData.length; i++) {
          const row = visitData[i];
          const patientId = safeString(row[VISIT_COL.PATIENT_ID]);

          // ★ ヒットした患者のみ処理
          if (!patientId || !matchedPatientIds.has(patientId)) continue;

          const visitDate = row[VISIT_COL.VISIT_DATE];

          // 最新の受診記録を保持
          if (!visitMap[patientId] ||
              (visitDate && visitMap[patientId].visitDate < visitDate)) {
            visitMap[patientId] = {
              visitId: safeString(row[VISIT_COL.VISIT_ID]),
              visitDate: visitDate,
              examType: safeString(row[VISIT_COL.EXAM_TYPE]),
              course: safeString(row[VISIT_COL.COURSE]),
              age: safeNumber(row[VISIT_COL.AGE]),
              status: safeString(row[VISIT_COL.STATUS]),
              _visitRowIndex: i + 1
            };
          }
        }
      }
    }

    // ★ Phase 3: 結果を構築
    const results = matchedRows.map(m => {
      const row = m.row;
      const visit = visitMap[m.patientId] || {};

      return {
        '受診者ID': m.patientId,
        'カルテNo': safeString(row[PATIENT_COL.KARTE_NO]),
        '氏名': safeString(row[PATIENT_COL.NAME]),
        'カナ': safeString(row[PATIENT_COL.KANA]),
        '生年月日': formatDateToString(row[PATIENT_COL.BIRTHDATE]),
        '性別': safeString(row[PATIENT_COL.GENDER]),
        '所属企業': safeString(row[PATIENT_COL.COMPANY]),
        // 最新の受診記録情報
        '受診ID': visit.visitId || '',
        '受診日': formatDateToString(visit.visitDate),
        '検診種別': visit.examType || '',
        'コース': visit.course || '',
        '年齢': visit.age || '',
        'ステータス': visit.status || '',
        _rowIndex: m.rowIndex + 1,
        _visitRowIndex: visit._visitRowIndex || null
      };
    });

    debugInfo.resultsCount = results.length;

    return {
      success: true,
      data: results,
      count: results.length,
      _debug: debugInfo
    };

  } catch (error) {
    console.error('portalSearchPatients error:', error);
    debugInfo.error = String(error.message);
    return {
      success: false,
      error: String(error.message),
      data: [],
      _debug: debugInfo
    };
  }
}

/**
 * カルテNoで受診者を検索
 * @param {string} karteNo - カルテNo（6桁の数字）
 * @returns {Object} 検索結果 { success, data, error }
 */
function portalSearchByKarteNo(karteNo) {
  try {
    if (!karteNo) {
      return { success: false, error: 'カルテNoが指定されていません', data: null };
    }

    const searchKarteNo = String(karteNo).trim();
    console.log('portalSearchByKarteNo: searching for karteNo=' + searchKarteNo);

    const ss = getPortalSpreadsheet();

    // 1. まず受診者マスタでカルテNo検索
    const patientSheet = ss.getSheetByName('受診者マスタ');
    if (patientSheet && patientSheet.getLastRow() >= 2) {
      const data = patientSheet.getDataRange().getValues();
      const headers = data[0];

      // カルテNo列のインデックスを取得
      let karteNoColIdx = headers.indexOf('カルテNo');
      if (karteNoColIdx < 0) {
        karteNoColIdx = PATIENT_COL.KARTE_NO; // フォールバック
      }

      console.log('受診者マスタ: karteNoColIdx=' + karteNoColIdx + ', totalRows=' + data.length);

      for (let i = 1; i < data.length; i++) {
        const cellKarteNo = String(data[i][karteNoColIdx] || '').trim();
        if (cellKarteNo === searchKarteNo) {
          console.log('受診者マスタで発見: row=' + (i + 1));

          // 受診者情報を構築
          const result = {
            '受診者ID': safeString(data[i][PATIENT_COL.PATIENT_ID]),
            'カルテNo': cellKarteNo,
            '氏名': safeString(data[i][PATIENT_COL.NAME]),
            'カナ': safeString(data[i][PATIENT_COL.KANA]),
            '生年月日': formatDateToString(data[i][PATIENT_COL.BIRTHDATE]),
            '性別': safeString(data[i][PATIENT_COL.GENDER]),
            '年齢': safeString(data[i][PATIENT_COL.AGE]),
            '所属企業': safeString(data[i][PATIENT_COL.COMPANY]),
            '_rowIndex': i + 1,
            '_source': '受診者マスタ'
          };

          return { success: true, data: result };
        }
      }
    }

    // 2. 受診者マスタで見つからなければ、検査結果シートを検索
    const labResultSheet = ss.getSheetByName('検査結果');
    if (labResultSheet && labResultSheet.getLastRow() >= 2) {
      const labData = labResultSheet.getDataRange().getValues();
      const labHeaders = labData[0];

      const labKarteNoColIdx = labHeaders.indexOf('カルテNo');
      const labPatientIdColIdx = labHeaders.indexOf('受診者ID');

      console.log('検査結果シート: karteNoColIdx=' + labKarteNoColIdx + ', totalRows=' + labData.length);

      if (labKarteNoColIdx >= 0) {
        for (let i = 1; i < labData.length; i++) {
          const cellKarteNo = String(labData[i][labKarteNoColIdx] || '').trim();
          if (cellKarteNo === searchKarteNo) {
            console.log('検査結果シートで発見: row=' + (i + 1));

            // 受診者IDがあれば取得
            const patientId = labPatientIdColIdx >= 0 ? safeString(labData[i][labPatientIdColIdx]) : '';

            const result = {
              '受診者ID': patientId,
              'カルテNo': cellKarteNo,
              '_rowIndex': i + 1,
              '_source': '検査結果'
            };

            return { success: true, data: result };
          }
        }
      }
    }

    console.log('カルテNo ' + searchKarteNo + ' は見つかりませんでした');
    return { success: false, error: 'カルテNo ' + searchKarteNo + ' が見つかりません', data: null };

  } catch (error) {
    console.error('portalSearchByKarteNo error:', error);
    return { success: false, error: error.message, data: null };
  }
}

/**
 * 帳票出力用 - 受診者OR検索（カルテNo / 氏名 / カナ）
 * @param {string} searchText - 検索文字列
 * @returns {Object} { success, data: Array, error }
 */
function portalSearchPatientsForReport(searchText) {
  try {
    if (!searchText || String(searchText).trim() === '') {
      return { success: false, error: '検索条件を入力してください', data: [] };
    }

    const searchTerm = String(searchText).trim();
    const normalizedSearch = normalizeString(searchTerm);
    const normalizedKanaSearch = normalizeKana(searchTerm);

    console.log('portalSearchPatientsForReport: searchTerm=' + searchTerm);

    const ss = getPortalSpreadsheet();
    const patientSheet = ss.getSheetByName('受診者マスタ');

    if (!patientSheet || patientSheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }

    const data = patientSheet.getDataRange().getValues();
    const headers = data[0];

    // 列インデックス取得
    const colIdx = {
      patientId: headers.indexOf('受診者ID') >= 0 ? headers.indexOf('受診者ID') : PATIENT_COL.PATIENT_ID,
      karteNo: headers.indexOf('カルテNo') >= 0 ? headers.indexOf('カルテNo') : PATIENT_COL.KARTE_NO,
      name: headers.indexOf('氏名') >= 0 ? headers.indexOf('氏名') : PATIENT_COL.NAME,
      kana: headers.indexOf('カナ') >= 0 ? headers.indexOf('カナ') : PATIENT_COL.KANA,
      birthdate: headers.indexOf('生年月日') >= 0 ? headers.indexOf('生年月日') : PATIENT_COL.BIRTHDATE,
      gender: headers.indexOf('性別') >= 0 ? headers.indexOf('性別') : PATIENT_COL.GENDER,
      age: headers.indexOf('年齢') >= 0 ? headers.indexOf('年齢') : PATIENT_COL.AGE,
      company: headers.indexOf('所属企業') >= 0 ? headers.indexOf('所属企業') : PATIENT_COL.COMPANY
    };

    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const karteNo = String(row[colIdx.karteNo] || '').trim();
      const name = String(row[colIdx.name] || '').trim();
      const kana = String(row[colIdx.kana] || '').trim();

      // OR検索: カルテNo 完全一致 OR 氏名 部分一致 OR カナ 部分一致
      const matchKarteNo = karteNo === searchTerm;
      const matchName = name && normalizeString(name).includes(normalizedSearch);
      const matchKana = kana && normalizeKana(kana).includes(normalizedKanaSearch);

      if (matchKarteNo || matchName || matchKana) {
        results.push({
          '受診者ID': safeString(row[colIdx.patientId]),
          'カルテNo': karteNo,
          '氏名': name,
          'カナ': kana,
          '生年月日': formatDateToString(row[colIdx.birthdate]),
          '性別': safeString(row[colIdx.gender]),
          '年齢': safeString(row[colIdx.age]),
          '所属企業': safeString(row[colIdx.company]),
          '_rowIndex': i + 1
        });
      }
    }

    console.log('portalSearchPatientsForReport: found ' + results.length + ' results');
    return { success: true, data: results };

  } catch (error) {
    console.error('portalSearchPatientsForReport error:', error);
    return { success: false, error: error.message, data: [] };
  }
}

// ============================================================
// 診断用関数
// ============================================================

/**
 * 検索診断 - 受診者マスタの構造と検索ロジックをテスト
 * GASエディタから実行してログを確認
 */
function diagnosePatientSearch() {
  const ss = getPortalSpreadsheet();
  const sheet = ss.getSheetByName('受診者マスタ');

  if (!sheet) {
    console.log('ERROR: 受診者マスタシートが見つかりません');
    return;
  }

  const data = sheet.getDataRange().getValues();
  console.log('=== 受診者マスタ診断 ===');
  console.log('総行数: ' + data.length);

  // ヘッダー行を表示
  console.log('ヘッダー: ' + JSON.stringify(data[0]));

  // 最初の5件のデータを詳細表示
  console.log('\n=== 最初の5件のデータ ===');
  for (let i = 1; i < Math.min(6, data.length); i++) {
    const row = data[i];
    console.log('--- Row ' + i + ' ---');
    console.log('  Col 0 (ID?): [' + row[0] + '] type=' + typeof row[0]);
    console.log('  Col 1 (氏名?): [' + row[1] + '] type=' + typeof row[1]);
    console.log('  Col 2 (カナ?): [' + row[2] + '] type=' + typeof row[2]);
    console.log('  Col 3: [' + row[3] + ']');
    console.log('  Col 4: [' + row[4] + ']');

    // 氏名の文字コードを詳細表示
    if (row[1]) {
      const nameStr = String(row[1]);
      const charCodes = [];
      for (let j = 0; j < nameStr.length; j++) {
        charCodes.push(nameStr.charCodeAt(j).toString(16));
      }
      console.log('  氏名の文字コード: ' + charCodes.join(' '));
    }
  }

  // 検索テスト
  console.log('\n=== 検索テスト ===');
  const testSearchTerm = '遠藤';
  const normalizedSearch = normalizeString(testSearchTerm);
  console.log('検索文字列: [' + testSearchTerm + ']');
  console.log('正規化後: [' + normalizedSearch + ']');

  let foundCount = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = safeString(row[1]);
    const normalizedName = normalizeString(name);

    if (normalizedName.includes(normalizedSearch)) {
      console.log('マッチ！ Row ' + i + ': [' + name + '] → [' + normalizedName + ']');
      foundCount++;
    }
  }
  console.log('マッチ件数: ' + foundCount);

  // 「遠藤」を含む可能性のある行を探す（文字コード比較なし）
  console.log('\n=== 部分一致検索（正規化なし）===');
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[1] || '');
    if (name.indexOf('遠') >= 0 || name.indexOf('藤') >= 0) {
      console.log('部分マッチ Row ' + i + ': [' + name + ']');
    }
  }
}

/**
 * 検査結果シートの診断（Webアプリからも呼び出し可能）
 * @returns {Object} 診断結果
 */
function diagnoseLabResultsSheet() {
  try {
    const ss = getPortalSpreadsheet();
    const result = {
      success: true,
      labResultSheet: null,
      patientSheet: null,
      comparison: null
    };

    // 検査結果シートの診断
    const labSheet = ss.getSheetByName('検査結果');
    if (labSheet) {
      const labData = labSheet.getDataRange().getValues();
      result.labResultSheet = {
        exists: true,
        rowCount: labData.length,
        headers: labData[0] || [],
        sampleData: labData.slice(1, 4).map((row, idx) => ({
          rowIndex: idx + 2,
          col0_受診者ID: row[0],
          col1_カルテNo: row[1],
          col2_受診日: row[2]
        }))
      };
    } else {
      result.labResultSheet = { exists: false };
    }

    // 受診者マスタシートの診断
    const patientSheet = ss.getSheetByName('受診者マスタ');
    if (patientSheet) {
      const patientData = patientSheet.getDataRange().getValues();
      result.patientSheet = {
        exists: true,
        rowCount: patientData.length,
        headers: patientData[0] || [],
        sampleData: patientData.slice(1, 4).map((row, idx) => ({
          rowIndex: idx + 2,
          col0: row[0],
          col1: row[1],
          col2: row[2],
          col3: row[3],
          col4_氏名: row[4]
        }))
      };
    } else {
      result.patientSheet = { exists: false };
    }

    // 期待されるヘッダーとの比較
    result.comparison = {
      expectedLabHeaders: getLabResultHeaders().slice(0, 5),
      expectedPatientCols: {
        PATIENT_ID: PATIENT_COL.PATIENT_ID,
        KARTE_NO: PATIENT_COL.KARTE_NO,
        NAME: PATIENT_COL.NAME
      }
    };

    console.log('=== 検査結果シート診断 ===');
    console.log(JSON.stringify(result, null, 2));

    return result;
  } catch (e) {
    console.error('diagnoseLabResultsSheet error:', e);
    return { success: false, error: e.message };
  }
}

/**
 * 全受診者をリスト（検索なし）
 */
function listAllPatients() {
  const ss = getPortalSpreadsheet();
  const sheet = ss.getSheetByName('受診者マスタ');

  if (!sheet) {
    console.log('ERROR: シートなし');
    return;
  }

  const data = sheet.getDataRange().getValues();
  console.log('ヘッダー: ' + JSON.stringify(data[0]));
  console.log('データ件数: ' + (data.length - 1));

  for (let i = 1; i < Math.min(20, data.length); i++) {
    console.log(i + ': ID=[' + data[i][0] + '] 氏名=[' + data[i][1] + '] カナ=[' + data[i][2] + ']');
  }
}

// ============================================================
// シリアライズ用ヘルパー関数
// ============================================================

/**
 * 値を安全な文字列に変換
 * @param {*} value - 任意の値
 * @returns {string} 文字列
 */
function safeString(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return formatDateToString(value);
  return String(value);
}

/**
 * 値を安全な数値に変換
 * @param {*} value - 任意の値
 * @returns {number|string} 数値または空文字
 */
function safeNumber(value) {
  if (value === null || value === undefined || value === '') return '';
  const num = Number(value);
  return isNaN(num) ? String(value) : num;
}

/**
 * Dateオブジェクトを文字列に変換
 * @param {*} value - 日付値
 * @returns {string} 日付文字列 (YYYY/MM/DD)
 */
function formatDateToString(value) {
  if (!value) return '';

  // 既に文字列の場合はそのまま返す
  if (typeof value === 'string') return value;

  // Dateオブジェクトの場合
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';

    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const day = String(value.getDate()).padStart(2, '0');
    return year + '/' + month + '/' + day;
  }

  // 数値の場合
  if (typeof value === 'number') {
    return String(value);
  }

  return String(value);
}

/**
 * 受診者を新規登録（v6 - 正規化構造対応）
 * 受診者マスタ + 受診記録に登録
 *
 * @param {Object} patientData - 受診者データ
 *   - 個人情報: name, nameKana, birthDate, gender, company, etc.
 *   - 受診情報: examType, course, examDate, status
 *   - existingPatientId: 既存患者の場合（受診記録のみ追加）
 * @returns {Object} 登録結果
 */
function portalRegisterPatient(patientData) {
  try {
    // === PatientService経由で受診者登録 ===
    // 既存患者IDが指定されている場合はそれを使用
    let patientId = patientData.existingPatientId || null;
    let isNewPatient = false;

    if (!patientId) {
      // 氏名＋生年月日で既存患者を検索
      if (patientData.name && patientData.birthDate) {
        const existing = PatientDAO.getByNameAndBirthdate(patientData.name, patientData.birthDate);
        if (existing) {
          patientId = existing.patientId;
        }
      }

      // 新規患者として登録
      if (!patientId) {
        // UIMapping経由でデータ変換してPatientService経由で登録
        const uiData = {
          karteNo: patientData.karteNo || '',
          status: patientData.status || '入力中',
          visitDate: patientData.examDate,
          name: patientData.name || '',
          kana: patientData.nameKana || '',
          gender: patientData.gender || '',
          birthdate: patientData.birthDate,
          course: patientData.course || '',
          company: patientData.company || '',
          department: patientData.department || '',
          bmlPatientId: patientData.bmlPatientId || ''
        };

        const result = PatientService.registerPatient(uiData);
        if (!result.success) {
          return result;
        }
        patientId = result.patientId;
        isNewPatient = true;
      }
    }

    // === 受診記録を追加（従来通り） ===
    const ss = getPortalSpreadsheet();
    const visitSheet = ss.getSheetByName('受診記録');
    if (!visitSheet) {
      return { success: false, error: '受診記録シートが見つかりません' };
    }

    const now = new Date();
    const visitDate = patientData.examDate ? new Date(patientData.examDate) : new Date();
    const visitId = generateNextVisitIdForPortal(visitSheet, visitDate);

    // 年齢を計算
    let age = '';
    if (patientData.birthDate && patientData.examDate) {
      const birth = new Date(patientData.birthDate);
      const exam = new Date(patientData.examDate);
      age = exam.getFullYear() - birth.getFullYear();
      const m = exam.getMonth() - birth.getMonth();
      if (m < 0 || (m === 0 && exam.getDate() < birth.getDate())) {
        age--;
      }
    }

    const newVisitRow = [
      visitId,                              // 受診ID
      patientId,                            // 受診者ID (FK)
      patientData.examType || '人間ドック', // 検診種別
      patientData.course || '',             // コース
      visitDate,                            // 受診日
      age,                                   // 年齢
      patientData.overallJudgment || '',    // 総合判定
      patientData.doctorNotes || '',        // 医師所見
      patientData.status || '入力中',       // ステータス
      '',                                    // CSV取込日時
      now,                                   // 作成日時
      now,                                   // 更新日時
      ''                                     // 出力日時
    ];

    visitSheet.appendRow(newVisitRow);

    return {
      success: true,
      message: isNewPatient ? '新規患者と受診記録を登録しました' : '既存患者に受診記録を追加しました',
      patientId: patientId,
      visitId: visitId,
      isNewPatient: isNewPatient
    };

  } catch (error) {
    console.error('portalRegisterPatient error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 次の受診者IDを生成
 * @param {Sheet} sheet - 受診者マスタシート
 * @returns {string} P-NNNNN形式のID
 */
function generateNextPatientIdForPortal(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return 'P-00001';
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let maxNum = 0;

  for (const row of data) {
    const id = String(row[0]);
    if (id.startsWith('P-')) {
      const num = parseInt(id.replace('P-', ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  return `P-${String(maxNum + 1).padStart(5, '0')}`;
}

/**
 * 次の受診IDを生成
 * @param {Sheet} sheet - 受診記録シート
 * @param {Date} visitDate - 受診日
 * @returns {string} YYYYMMDD-NNN形式のID
 */
function generateNextVisitIdForPortal(sheet, visitDate) {
  const date = visitDate || new Date();
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
  const prefix = `${dateStr}-`;

  if (!sheet || sheet.getLastRow() < 2) {
    return `${prefix}001`;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let maxNum = 0;

  for (const row of data) {
    const id = String(row[0]);
    if (id.startsWith(prefix)) {
      const num = parseInt(id.replace(prefix, ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  return `${prefix}${String(maxNum + 1).padStart(3, '0')}`;
}

/**
 * BML患者IDで受診者を検索
 * CSV取込時の受診者紐付けに使用
 * @param {string} bmlPatientId - BML患者ID（例: 999991）
 * @returns {Object|null} 受診者データまたはnull
 */
function portalFindPatientByBmlId(bmlPatientId) {
  if (!bmlPatientId) return null;

  try {
    const ss = getPortalSpreadsheet();
    const patientSheet = ss.getSheetByName('受診者マスタ');

    if (!patientSheet || patientSheet.getLastRow() < 2) {
      return null;
    }

    const data = patientSheet.getDataRange().getValues();
    const bmlIdStr = String(bmlPatientId).trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[PATIENT_COL.PATIENT_ID]) continue;

      const storedBmlId = String(row[PATIENT_COL.BML_PATIENT_ID] || '').trim();
      if (storedBmlId && storedBmlId === bmlIdStr) {
        return {
          patientId: safeString(row[PATIENT_COL.PATIENT_ID]),
          name: safeString(row[PATIENT_COL.NAME]),
          kana: safeString(row[PATIENT_COL.KANA]),
          birthDate: formatDateToString(row[PATIENT_COL.BIRTHDATE]),
          gender: safeString(row[PATIENT_COL.GENDER]),
          company: safeString(row[PATIENT_COL.COMPANY]),
          bmlPatientId: storedBmlId,
          rowIndex: i + 1
        };
      }
    }

    return null;

  } catch (e) {
    console.error('portalFindPatientByBmlId error:', e);
    return null;
  }
}

/**
 * 受診者情報を取得（ID指定 - v6 正規化構造対応）
 * 受診者情報と全受診履歴を取得
 * @param {string} patientId - 受診者ID (P-00001形式)
 * @returns {Object} 受診者情報と受診履歴
 */
function portalGetPatient(patientId) {
  try {
    const ss = getPortalSpreadsheet();

    // 受診者マスタから検索
    const patientSheet = ss.getSheetByName('受診者マスタ');
    if (!patientSheet) {
      return { success: false, error: '受診者マスタシートが見つかりません' };
    }

    const patientData = patientSheet.getDataRange().getValues();
    let patientInfo = null;
    let patientRowIndex = null;

    for (let i = 1; i < patientData.length; i++) {
      if (safeString(patientData[i][PATIENT_COL.PATIENT_ID]) === String(patientId)) {
        const row = patientData[i];
        patientInfo = {
          '受診者ID': safeString(row[PATIENT_COL.PATIENT_ID]),
          'カルテNo': safeString(row[PATIENT_COL.KARTE_NO]),  // ★追加: CSV紐付け用
          'ステータス': safeString(row[PATIENT_COL.STATUS]),
          '受診日': formatDateToString(row[PATIENT_COL.VISIT_DATE]),
          '氏名': safeString(row[PATIENT_COL.NAME]),
          'カナ': safeString(row[PATIENT_COL.KANA]),
          '性別': safeString(row[PATIENT_COL.GENDER]),
          '生年月日': formatDateToString(row[PATIENT_COL.BIRTHDATE]),
          '年齢': safeString(row[PATIENT_COL.AGE]),
          '受診コース': safeString(row[PATIENT_COL.COURSE]),
          '所属企業': safeString(row[PATIENT_COL.COMPANY]),
          '所属': safeString(row[PATIENT_COL.DEPARTMENT]),
          '総合判定': safeString(row[PATIENT_COL.OVERALL_JUDGMENT]),
          'CSV取込日時': formatDateToString(row[PATIENT_COL.CSV_IMPORT_DATE]),
          '更新日時': formatDateToString(row[PATIENT_COL.UPDATED_AT]),
          '出力日時': formatDateToString(row[PATIENT_COL.EXPORT_DATE]),
          'BML患者ID': safeString(row[PATIENT_COL.BML_PATIENT_ID])
        };
        patientRowIndex = i + 1;
        break;
      }
    }

    if (!patientInfo) {
      return { success: false, error: '受診者が見つかりません' };
    }

    // 受診記録から該当患者の全受診履歴を取得
    const visits = [];
    const visitSheet = ss.getSheetByName('受診記録');

    if (visitSheet && visitSheet.getLastRow() >= 2) {
      const visitData = visitSheet.getDataRange().getValues();

      for (let i = 1; i < visitData.length; i++) {
        if (safeString(visitData[i][VISIT_COL.PATIENT_ID]) === String(patientId)) {
          const row = visitData[i];
          visits.push({
            '受診ID': safeString(row[VISIT_COL.VISIT_ID]),
            '検診種別': safeString(row[VISIT_COL.EXAM_TYPE]),
            'コース': safeString(row[VISIT_COL.COURSE]),
            '受診日': formatDateToString(row[VISIT_COL.VISIT_DATE]),
            '年齢': safeNumber(row[VISIT_COL.AGE]),
            '総合判定': safeString(row[VISIT_COL.JUDGMENT]),
            '医師所見': safeString(row[VISIT_COL.DOCTOR_NOTES]),
            'ステータス': safeString(row[VISIT_COL.STATUS]),
            'CSV取込日時': formatDateToString(row[VISIT_COL.CSV_IMPORTED]),
            '出力日時': formatDateToString(row[VISIT_COL.OUTPUT_AT]),
            _visitRowIndex: i + 1
          });
        }
      }

      // 受診日で降順ソート（新しい順）
      visits.sort((a, b) => {
        const dateA = a['受診日'] ? new Date(a['受診日']) : new Date(0);
        const dateB = b['受診日'] ? new Date(b['受診日']) : new Date(0);
        return dateB - dateA;
      });
    }

    return {
      success: true,
      data: {
        ...patientInfo,
        visits: visits,
        visitCount: visits.length,
        _rowIndex: patientRowIndex
      }
    };

  } catch (error) {
    console.error('portalGetPatient error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 受診者情報を更新（v6 - 正規化構造対応）
 * @param {string} patientId - 受診者ID (P-00001形式)
 * @param {Object} updateData - 更新データ（受診者マスタ用）
 * @returns {Object} 更新結果
 */
function portalUpdatePatient(patientId, updateData) {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('受診者マスタ');
    if (!sheet) {
      return { success: false, error: '受診者マスタシートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (safeString(data[i][PATIENT_COL.PATIENT_ID]) === String(patientId)) {
        // 該当行を更新（旧構造に合わせた列定義）
        if (updateData['ステータス'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.STATUS + 1).setValue(updateData['ステータス']);
        if (updateData['受診日'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.VISIT_DATE + 1).setValue(updateData['受診日'] ? new Date(updateData['受診日']) : '');
        if (updateData['氏名'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.NAME + 1).setValue(updateData['氏名']);
        if (updateData['カナ'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.KANA + 1).setValue(updateData['カナ']);
        if (updateData['性別'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.GENDER + 1).setValue(updateData['性別']);
        if (updateData['生年月日'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.BIRTHDATE + 1).setValue(updateData['生年月日'] ? new Date(updateData['生年月日']) : '');
        if (updateData['年齢'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.AGE + 1).setValue(updateData['年齢']);
        if (updateData['受診コース'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.COURSE + 1).setValue(updateData['受診コース']);
        if (updateData['所属企業'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.COMPANY + 1).setValue(updateData['所属企業']);
        if (updateData['所属'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.DEPARTMENT + 1).setValue(updateData['所属']);
        if (updateData['総合判定'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.OVERALL_JUDGMENT + 1).setValue(updateData['総合判定']);
        if (updateData['BML患者ID'] !== undefined) sheet.getRange(i + 1, PATIENT_COL.BML_PATIENT_ID + 1).setValue(updateData['BML患者ID']);

        // 更新日時を更新
        sheet.getRange(i + 1, PATIENT_COL.UPDATED_AT + 1).setValue(new Date());

        return { success: true, message: '受診者情報を更新しました' };
      }
    }

    return { success: false, error: '受診者が見つかりません' };
  } catch (error) {
    console.error('portalUpdatePatient error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 受診記録を更新（v6 - 正規化構造対応）
 * @param {string} visitId - 受診ID (YYYYMMDD-NNN形式)
 * @param {Object} updateData - 更新データ（受診記録用）
 * @returns {Object} 更新結果
 */
function portalUpdateVisit(visitId, updateData) {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('受診記録');
    if (!sheet) {
      return { success: false, error: '受診記録シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (safeString(data[i][VISIT_COL.VISIT_ID]) === String(visitId)) {
        // 該当行を更新
        if (updateData['検診種別'] !== undefined) sheet.getRange(i + 1, VISIT_COL.EXAM_TYPE + 1).setValue(updateData['検診種別']);
        if (updateData['コース'] !== undefined) sheet.getRange(i + 1, VISIT_COL.COURSE + 1).setValue(updateData['コース']);
        if (updateData['受診日'] !== undefined) sheet.getRange(i + 1, VISIT_COL.VISIT_DATE + 1).setValue(updateData['受診日'] ? new Date(updateData['受診日']) : '');
        if (updateData['年齢'] !== undefined) sheet.getRange(i + 1, VISIT_COL.AGE + 1).setValue(updateData['年齢']);
        if (updateData['総合判定'] !== undefined) sheet.getRange(i + 1, VISIT_COL.JUDGMENT + 1).setValue(updateData['総合判定']);
        if (updateData['医師所見'] !== undefined) sheet.getRange(i + 1, VISIT_COL.DOCTOR_NOTES + 1).setValue(updateData['医師所見']);
        if (updateData['ステータス'] !== undefined) sheet.getRange(i + 1, VISIT_COL.STATUS + 1).setValue(updateData['ステータス']);

        // 更新日時を更新
        sheet.getRange(i + 1, VISIT_COL.UPDATED_AT + 1).setValue(new Date());

        return { success: true, message: '受診記録を更新しました' };
      }
    }

    return { success: false, error: '受診記録が見つかりません' };
  } catch (error) {
    console.error('portalUpdateVisit error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 受診者詳細を更新（ポータルUI用 - 互換性維持）
 * js.html から呼び出される
 * IDの形式で受診者マスタか受診記録かを判別
 * @param {string} id - 受診者ID (P-00001) または 受診ID (YYYYMMDD-NNN)
 * @param {Object} updateData - 更新データ
 * @returns {Object} 更新結果
 */
function portalUpdatePatientDetail(id, updateData) {
  // P-で始まる場合は受診者マスタ
  if (String(id).startsWith('P-')) {
    return portalUpdatePatient(id, updateData);
  }
  // それ以外は受診記録
  return portalUpdateVisit(id, updateData);
}

// ============================================================
// 検査結果入力 API
// ============================================================

/**
 * 血液検査結果を取得（カルテNo検索）
 * @param {string} karteNo - カルテNo
 * @returns {Object} 血液検査結果
 */
function portalGetBloodTestResults(karteNo) {
  try {
    if (!karteNo) {
      return { success: false, error: 'カルテNoが指定されていません', data: null };
    }

    console.log('portalGetBloodTestResults: karteNo=' + karteNo);

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません', data: null };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const karteNoColIndex = headers.indexOf('カルテNo');

    if (karteNoColIndex < 0) {
      return { success: false, error: 'ヘッダーにカルテNoがありません', data: null };
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][karteNoColIndex]).trim() === String(karteNo).trim()) {
        const result = {};
        headers.forEach((header, idx) => {
          result[header] = data[i][idx];
        });
        result._rowIndex = i + 1;
        console.log('Found blood test results at row ' + (i + 1));
        return { success: true, data: result };
      }
    }

    return { success: true, data: null, message: '検査結果がありません' };
  } catch (error) {
    console.error('portalGetBloodTestResults error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 身体計測結果を保存（カルテNo検索）
 * @param {string} karteNo - カルテNo
 * @param {Object} measurements - 身体計測データ
 * @returns {Object} 保存結果
 */
function portalSaveBodyMeasurements(karteNo, measurements) {
  try {
    console.log('=== portalSaveBodyMeasurements開始 ===');
    console.log('karteNo: ' + karteNo + ' (type: ' + typeof karteNo + ')');
    console.log('measurements: ' + JSON.stringify(measurements));

    if (!karteNo) {
      return { success: false, error: 'カルテNoが指定されていません' };
    }

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const karteNoColIndex = headers.indexOf('カルテNo');

    console.log('ヘッダー数: ' + headers.length);
    console.log('カルテNo列idx: ' + karteNoColIndex);
    console.log('体重列idx: ' + headers.indexOf('体重'));

    if (karteNoColIndex < 0) {
      return { success: false, error: 'ヘッダーにカルテNoがありません' };
    }

    // 身体計測項目のマッピング
    const measurementFields = {
      '身長': measurements.height,
      '体重': measurements.weight,
      'BMI': measurements.bmi,
      '収縮期血圧': measurements.systolicBP,
      '拡張期血圧': measurements.diastolicBP,
      // ★ 血圧2回目対応: 2回目フィールド追加
      '収縮期血圧_2回目': measurements.systolicBP2,
      '拡張期血圧_2回目': measurements.diastolicBP2,
      '脈拍': measurements.pulse,
      '視力_右': measurements.visionRight,
      '視力_左': measurements.visionLeft,
      '視力_矯正': measurements.visionCorrected,
      '聴力_右_1000Hz': measurements.hearingRight1000,
      '聴力_右_4000Hz': measurements.hearingRight4000,
      '聴力_左_1000Hz': measurements.hearingLeft1000,
      '聴力_左_4000Hz': measurements.hearingLeft4000,
      '心電図所見': measurements.ecgFindings,
      '心電図コメント': measurements.ecgComment,
      '胸部X線所見': measurements.chestXrayFindings,
      '胸部X線コメント': measurements.chestXrayComment,
      '更新日時': new Date()
    };

    // 各フィールドのマッピング診断
    console.log('=== フィールドマッピング診断 ===');
    Object.keys(measurementFields).forEach(fieldName => {
      const colIdx = headers.indexOf(fieldName);
      const value = measurementFields[fieldName];
      const willSave = colIdx >= 0 && value !== undefined && value !== '';
      console.log(fieldName + ': colIdx=' + colIdx + ', value=[' + value + '], type=' + typeof value + ', willSave=' + willSave);
    });

    for (let i = 1; i < data.length; i++) {
      const storedKarteNo = String(data[i][karteNoColIndex]).trim();
      const searchKarteNo = String(karteNo).trim();

      if (storedKarteNo === searchKarteNo) {
        console.log('★ マッチ: Row ' + (i + 1) + ', karteNo=' + storedKarteNo);

        let savedCount = 0;
        Object.keys(measurementFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          const value = measurementFields[fieldName];
          // 空文字も許可（値をクリアする場合）、undefinedのみスキップ
          if (colIdx >= 0 && value !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(value);
            savedCount++;
            console.log('  保存: ' + fieldName + ' = [' + value + '] → 列' + (colIdx + 1));
          }
        });

        console.log('保存完了: ' + savedCount + '項目');
        return { success: true, message: '身体計測結果を保存しました（' + savedCount + '項目）' };
      }
    }

    console.log('該当行なし - 新規行追加');
    const newRow = headers.map(header => {
      if (header === 'カルテNo') return karteNo;
      if (measurementFields.hasOwnProperty(header)) return measurementFields[header] || '';
      return '';
    });
    sheet.appendRow(newRow);

    return { success: true, message: '身体計測結果を新規登録しました' };
  } catch (error) {
    console.error('portalSaveBodyMeasurements error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 血液検査結果を保存（カルテNo検索）
 * @param {string} karteNo - カルテNo
 * @param {Object} bloodData - 血液検査データ（キー: 項目名, 値: 検査値）
 * @returns {Object} 保存結果
 */
function portalSaveBloodTestResults(karteNo, bloodData) {
  try {
    if (!karteNo) {
      return { success: false, error: 'カルテNoが指定されていません' };
    }

    console.log('portalSaveBloodTestResults: karteNo=' + karteNo);

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const karteNoColIndex = headers.indexOf('カルテNo');

    if (karteNoColIndex < 0) {
      return { success: false, error: 'ヘッダーにカルテNoがありません' };
    }

    // 血液検査項目のマッピング（bloodDataのキーはそのまま列名として使用）
    const bloodFields = {
      ...bloodData,
      '血液検査入力日': new Date(),
      '更新日時': new Date()
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][karteNoColIndex]).trim() === String(karteNo).trim()) {
        // 該当行を更新
        Object.keys(bloodFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          if (colIdx >= 0 && bloodFields[fieldName] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(bloodFields[fieldName]);
          }
        });
        return { success: true, message: '血液検査結果を保存しました' };
      }
    }

    // 該当行がない場合は新規行を追加
    const newRow = headers.map(header => {
      if (header === 'カルテNo') return karteNo;
      if (bloodFields.hasOwnProperty(header)) return bloodFields[header];
      return '';
    });
    sheet.appendRow(newRow);

    return { success: true, message: '血液検査結果を新規登録しました' };
  } catch (error) {
    console.error('portalSaveBloodTestResults error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 検査結果を保存（設計書準拠版）
 * @param {string} patientId - 受診者ID
 * @param {Object} inputData - 検査データ（キー: 項目名, 値: 検査値）
 * @returns {Object} 保存結果
 */
function portalSaveInputResults(patientId, inputData) {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    // 修正: 患者ID → 受診者ID（LAB_RESULT_COLUMNS準拠）
    const idColIndex = headers.indexOf('受診者ID');

    if (idColIndex < 0) {
      console.error('portalSaveInputResults: 受診者ID列が見つかりません。headers=' + JSON.stringify(headers.slice(0, 10)));
      return { success: false, error: '検査結果シートに受診者ID列がありません' };
    }

    // 入力項目のマッピング（inputDataのキーはそのまま列名として使用）
    const inputFields = {
      ...inputData,
      '結果入力日': new Date(),
      '更新日時': new Date()
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]).trim() === String(patientId).trim()) {
        // 該当行を更新
        Object.keys(inputFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          if (colIdx >= 0 && inputFields[fieldName] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(inputFields[fieldName]);
          }
        });
        console.log('portalSaveInputResults: 更新完了 row=' + (i + 1) + ', patientId=' + patientId);
        return { success: true, message: '検査結果を保存しました' };
      }
    }

    // 該当行がない場合は新規行を追加
    const newRow = headers.map(header => {
      if (header === '受診者ID') return patientId;
      if (inputFields.hasOwnProperty(header)) return inputFields[header];
      return '';
    });
    sheet.appendRow(newRow);
    console.log('portalSaveInputResults: 新規追加 patientId=' + patientId);

    return { success: true, message: '検査結果を新規登録しました' };
  } catch (error) {
    console.error('portalSaveInputResults error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 所見を保存
 * @param {string} patientId - 患者ID
 * @param {Object} findings - 所見データ
 * @returns {Object} 保存結果
 */
function portalSaveFindings(patientId, findings) {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf('患者ID');

    // 所見項目のマッピング
    const findingFields = {
      '総合判定': findings.overallJudgment,
      '総合所見': findings.overallFindings,
      '医師コメント': findings.doctorComment,
      '生活指導': findings.lifestyleGuidance,
      '要精密検査項目': findings.needsDetailedExam,
      '要治療項目': findings.needsTreatment,
      '所見入力日': new Date(),
      '所見入力者': Session.getActiveUser().getEmail()
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(patientId)) {
        Object.keys(findingFields).forEach(fieldName => {
          const colIdx = headers.indexOf(fieldName);
          if (colIdx >= 0 && findingFields[fieldName] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(findingFields[fieldName]);
          }
        });
        return { success: true, message: '所見を保存しました' };
      }
    }

    return { success: false, error: '検査結果レコードが見つかりません。先にCSV取込を行ってください。' };
  } catch (error) {
    console.error('portalSaveFindings error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// CSV取込 API
// ============================================================

/**
 * CSVデータを取り込み
 * @param {string} csvContent - CSVコンテンツ
 * @param {Object} options - 取込オプション
 * @returns {Object} 取込結果
 */
function portalImportCsv(csvContent, options) {
  try {
    options = options || {};

    // BMLコード形式を自動検出
    if (typeof isBmlCodeFormat === 'function' && isBmlCodeFormat(csvContent)) {
      return portalImportBmlCodeFormat(csvContent, options);
    }

    // 既存のprocessCsvContent関数があれば使用
    if (typeof processCsvContent === 'function') {
      return processCsvContent(csvContent, options);
    }

    // ヘッダー付きCSVの処理（従来形式）
    const lines = csvContent.split(/\r?\n/).filter(line => line.trim());
    if (lines.length < 2) {
      return { success: false, error: 'CSVデータが空または不正です' };
    }

    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    const idColIndex = headers.findIndex(h => h.includes('ID') || h.includes('患者') || h.includes('カルテ'));

    if (idColIndex < 0) {
      return { success: false, error: 'ID列が見つかりません' };
    }

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');
    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const existingData = sheet.getDataRange().getValues();
    const existingHeaders = existingData[0];
    const existingIdCol = existingHeaders.indexOf('患者ID');
    const existingIds = new Set(existingData.slice(1).map(row => String(row[existingIdCol])));

    let imported = 0;
    let updated = 0;
    let errors = [];

    for (let i = 1; i < lines.length; i++) {
      try {
        const values = lines[i].split(',').map(v => v.trim().replace(/^"|"$/g, ''));
        const patientId = values[idColIndex];

        if (!patientId) {
          errors.push({ row: i + 1, error: 'IDが空です' });
          continue;
        }

        if (existingIds.has(String(patientId))) {
          updated++;
        } else {
          const newRow = existingHeaders.map(header => {
            const idx = headers.indexOf(header);
            return idx >= 0 ? values[idx] : '';
          });
          newRow[existingIdCol] = patientId;
          sheet.appendRow(newRow);
          existingIds.add(String(patientId));
          imported++;
        }
      } catch (rowError) {
        errors.push({ row: i + 1, error: rowError.message });
      }
    }

    return {
      success: true,
      imported: imported,
      updated: updated,
      errors: errors,
      total: lines.length - 1,
      message: `取込完了: 新規${imported}件, 更新${updated}件` + (errors.length > 0 ? `, エラー${errors.length}件` : '')
    };
  } catch (error) {
    console.error('portalImportCsv error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * BMLコード形式CSVの取込処理
 * @param {string} csvContent - BMLコード形式CSV
 * @param {Object} options - オプション
 * @returns {Object} 取込結果
 */
function portalImportBmlCodeFormat(csvContent, options) {
  try {
    // BMLコード形式をパース
    const parseResult = parseBmlCodeFormatCsv(csvContent, options);
    if (!parseResult.success) {
      return { success: false, error: parseResult.error };
    }

    const ss = getPortalSpreadsheet();
    const patientSheet = ss.getSheetByName('受診者マスタ');
    let resultSheet = ss.getSheetByName('検査結果');

    if (!patientSheet) {
      return { success: false, error: '受診者マスタシートが見つかりません' };
    }

    // 検査結果シートがない場合は作成
    if (!resultSheet) {
      resultSheet = ss.insertSheet('検査結果');
    }

    // ヘッダーが空または不正な場合は初期化
    const expectedHeaders = getLabResultHeaders();
    const lastCol = resultSheet.getLastColumn();
    if (lastCol === 0 || lastCol < expectedHeaders.length) {
      const headerRange = resultSheet.getRange(1, 1, 1, expectedHeaders.length);
      headerRange.setValues([expectedHeaders]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e8f0fe');
      console.log('検査結果シートのヘッダーを初期化しました');
    }

    // 既存データのマップを作成（カルテNo → 行番号）★カルテNo統一設計
    const existingData = resultSheet.getDataRange().getValues();
    const resultHeaders = existingData[0];
    const karteNoColIdx = resultHeaders.indexOf('カルテNo');
    const karteNoToRow = new Map();
    for (let i = 1; i < existingData.length; i++) {
      const kno = karteNoColIdx >= 0 ? String(existingData[i][karteNoColIdx]).trim() : '';
      if (kno) {
        karteNoToRow.set(kno, i + 1); // 1-indexed
      }
    }
    console.log('既存カルテNoマップ: ' + karteNoToRow.size + '件');

    let linked = 0;
    let notFound = 0;
    let saved = 0;
    let updated = 0;
    const errors = [];
    const notFoundList = [];

    for (const record of parseResult.records) {
      try {
        // カルテNoで受診者を検索
        const karteNo = record.KARTE_NO;
        if (!karteNo) {
          errors.push({ row: record._rowIndex, error: 'カルテNoが空です' });
          continue;
        }

        // PatientDAOでカルテNo検索
        let patient = null;
        if (typeof PatientDAO !== 'undefined' && PatientDAO.getByKarteNo) {
          patient = PatientDAO.getByKarteNo(karteNo);
        } else {
          // フォールバック: 直接シート検索
          patient = findPatientByKarteNoInternal(patientSheet, karteNo);
        }

        if (!patient) {
          notFound++;
          notFoundList.push({ karteNo: karteNo, examDate: record.EXAM_DATE });
          continue;
        }

        linked++;

        // 検査結果を保存
        const resultRow = buildResultRow(patient.patientId, record, resultSheet);
        if (resultRow) {
          // 既存行があれば更新、なければ追加（★カルテNoで検索）
          const searchKarteNo = String(karteNo).trim();
          const existingRowNum = karteNoToRow.get(searchKarteNo);
          console.log('カルテNo ' + searchKarteNo + ' → 既存行: ' + (existingRowNum || 'なし'));
          if (existingRowNum) {
            // 既存行を更新
            resultSheet.getRange(existingRowNum, 1, 1, resultRow.length).setValues([resultRow]);
            updated++;
            console.log('既存行 ' + existingRowNum + ' を更新');
          } else {
            // 新規行を追加
            resultSheet.appendRow(resultRow);
            karteNoToRow.set(searchKarteNo, resultSheet.getLastRow());
            saved++;
            console.log('新規行追加: ' + resultSheet.getLastRow());
          }
        }

        // 受診者マスタのCSV取込日時を更新
        if (typeof PatientDAO !== 'undefined' && PatientDAO.updateFields) {
          PatientDAO.updateFields(patient.patientId, {
            csvImportDate: new Date(),
            bmlPatientId: record.BML_PATIENT_ID || ''
          });
        }

      } catch (recordError) {
        errors.push({ row: record._rowIndex, error: recordError.message });
      }
    }

    return {
      success: true,
      format: 'BML_CODE_FORMAT',
      total: parseResult.records.length,
      linked: linked,
      notFound: notFound,
      saved: saved,
      updated: updated,
      errors: errors,
      notFoundList: notFoundList.slice(0, 10), // 最大10件
      mappingInfo: parseResult.mappingInfo,
      message: `BML取込完了: 紐付け${linked}件（新規${saved}件, 更新${updated}件）, 未紐付け${notFound}件` +
               (errors.length > 0 ? `, エラー${errors.length}件` : '')
    };

  } catch (error) {
    console.error('portalImportBmlCodeFormat error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * カルテNoで受診者を検索（内部用フォールバック）
 * @deprecated CRUD.js getPatientByKarteNo() に統一済み
 * @param {Sheet} sheet - 受診者マスタシート（互換性のため維持、未使用）
 * @param {string} karteNo - カルテNo
 * @returns {Object|null} 受診者情報またはnull
 */
function findPatientByKarteNoInternal(sheet, karteNo) {
  // CRUD.js の基本関数を使用
  const result = getPatientByKarteNo(karteNo);
  if (result) {
    return {
      patientId: result.patientId,
      karteNo: result.karteNo,
      name: result.name,
      _rowIndex: result.rowIndex
    };
  }
  return null;
}

/**
 * 検査結果行を構築（LAB_RESULT_COLUMNS準拠）
 * @param {string} patientId - 受診者ID
 * @param {Object} record - パース済みCSVレコード
 * @param {Sheet} sheet - 検査結果シート
 * @returns {Array} 行データ配列
 */
function buildResultRow(patientId, record, sheet) {
  try {
    console.log('=== buildResultRow開始 ===');
    console.log('patientId: ' + patientId);
    console.log('record.KARTE_NO: ' + record.KARTE_NO);
    console.log('recordのキー: ' + Object.keys(record).filter(k => !k.startsWith('_')).join(', '));

    // ヘッダーを取得（LAB_RESULT_COLUMNS準拠）
    const headers = getLabResultHeaders();
    const row = new Array(headers.length).fill('');

    // 基本情報をマッピング
    const baseColMap = {};
    LAB_RESULT_COLUMNS.BASE.forEach((col, idx) => {
      baseColMap[col.id] = idx;
    });

    // 基本情報を設定
    if (baseColMap['PATIENT_ID'] !== undefined) row[baseColMap['PATIENT_ID']] = patientId;
    if (baseColMap['KARTE_NO'] !== undefined) row[baseColMap['KARTE_NO']] = record.KARTE_NO || '';
    if (baseColMap['EXAM_DATE'] !== undefined) row[baseColMap['EXAM_DATE']] = record.EXAM_DATE || '';
    if (baseColMap['BML_PATIENT_ID'] !== undefined) row[baseColMap['BML_PATIENT_ID']] = record.BML_PATIENT_ID || '';
    if (baseColMap['CSV_IMPORT_DATE'] !== undefined) row[baseColMap['CSV_IMPORT_DATE']] = new Date();

    // 検査項目をマッピング（BMLコードをキーとして使用）
    const baseLen = LAB_RESULT_COLUMNS.BASE.length;
    let mappedCount = 0;
    LAB_RESULT_COLUMNS.ITEMS.forEach((item, idx) => {
      const colIdx = baseLen + idx;
      // BMLコードをキーとして検索（parseBmlCodeFormatCsvと同期）
      const value = record[item.bmlCode];
      if (value !== undefined && value !== '') {
        row[colIdx] = value;
        mappedCount++;
      }
    });

    console.log('マッピングされた検査項目数: ' + mappedCount);
    console.log('=== buildResultRow完了 ===');

    return row;
  } catch (e) {
    console.error('buildResultRow error:', e);
    return null;
  }
}

/**
 * 受診者IDで検査結果を取得
 * @param {string} patientId - 受診者ID
 * @returns {Object} 検査結果データ
 */
function portalGetLabResults(patientId) {
  try {
    if (!patientId) {
      return { success: false, error: '受診者IDが指定されていません', data: null };
    }

    console.log('portalGetLabResults: searching for patientId=' + patientId);

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');

    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません', data: null };
    }

    const data = sheet.getDataRange().getValues();
    console.log('検査結果 sheet rows: ' + data.length);

    if (data.length < 2) {
      return { success: true, data: null, message: '検査結果がありません' };
    }

    const headers = data[0];
    console.log('検査結果 headers: ' + JSON.stringify(headers.slice(0, 6)));

    const patientIdColIdx = headers.indexOf('受診者ID');
    console.log('受診者ID column index: ' + patientIdColIdx);

    if (patientIdColIdx < 0) {
      return { success: false, error: 'ヘッダーに受診者IDがありません', data: null };
    }

    // デバッグ: 最初の数行のID値を表示
    for (let i = 1; i < Math.min(4, data.length); i++) {
      console.log('Row ' + i + ' 受診者ID: [' + data[i][patientIdColIdx] + ']');
    }

    // 該当患者の行を検索
    for (let i = 1; i < data.length; i++) {
      const storedId = String(data[i][patientIdColIdx]).trim();
      const searchId = String(patientId).trim();
      if (storedId === searchId) {
        console.log('Found match at row ' + i);
        // 検査結果をオブジェクトに変換
        const result = {
          _rowIndex: i + 1,
          items: []
        };

        // 基本情報
        LAB_RESULT_COLUMNS.BASE.forEach((col, idx) => {
          result[col.id] = data[i][idx];
        });

        // 検査項目
        const baseLen = LAB_RESULT_COLUMNS.BASE.length;
        LAB_RESULT_COLUMNS.ITEMS.forEach((item, idx) => {
          const value = data[i][baseLen + idx];
          if (value !== '' && value !== null && value !== undefined) {
            result.items.push({
              id: item.id,
              name: item.name,
              value: value,
              unit: item.unit,
              bmlCode: item.bmlCode
            });
            // フラットアクセス用
            result[item.id] = value;
          }
        });

        return {
          success: true,
          data: result,
          itemCount: result.items.length
        };
      }
    }

    console.log('No match found for patientId: ' + patientId);
    return { success: true, data: null, message: '該当する検査結果がありません', _debug: { searchedId: patientId, rowCount: data.length - 1 } };

  } catch (e) {
    console.error('portalGetLabResults error:', e);
    return { success: false, error: e.message, data: null };
  }
}

/**
 * カルテNoで検査結果を取得
 * @param {string} karteNo - カルテNo
 * @returns {Object} 検査結果データ
 */
function portalGetLabResultsByKarteNo(karteNo) {
  try {
    console.log('=== portalGetLabResultsByKarteNo 開始 ===');
    console.log('引数 karteNo: ' + karteNo + ' (type: ' + typeof karteNo + ')');

    if (!karteNo) {
      console.log('ERROR: カルテNoが空');
      return { success: false, error: 'カルテNoが指定されていません', data: null };
    }

    console.log('portalGetLabResultsByKarteNo: searching for karteNo=' + karteNo);

    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');

    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません', data: null };
    }

    const data = sheet.getDataRange().getValues();
    console.log('検査結果 sheet rows: ' + data.length);

    if (data.length < 2) {
      return { success: true, data: null, message: '検査結果がありません' };
    }

    const headers = data[0];
    console.log('検査結果 headers (first 5): ' + JSON.stringify(headers.slice(0, 5)));

    const karteNoColIdx = headers.indexOf('カルテNo');
    console.log('カルテNo column index: ' + karteNoColIdx);

    if (karteNoColIdx < 0) {
      return { success: false, error: 'ヘッダーにカルテNoがありません', data: null };
    }

    // デバッグ: 最初の数行のカルテNo値を表示
    for (let i = 1; i < Math.min(4, data.length); i++) {
      console.log('Row ' + i + ' カルテNo: [' + data[i][karteNoColIdx] + ']');
    }

    // 該当患者の行を検索
    console.log('検索開始: 対象行数=' + (data.length - 1));
    for (let i = 1; i < data.length; i++) {
      const storedKarteNo = String(data[i][karteNoColIdx]).trim();
      const searchKarteNo = String(karteNo).trim();
      console.log('Row ' + i + ': stored=[' + storedKarteNo + '] vs search=[' + searchKarteNo + '] match=' + (storedKarteNo === searchKarteNo));
      if (storedKarteNo === searchKarteNo) {
        console.log('★ Found match at row ' + i);
        // portalGetLabResultsと同じ形式で返す
        const result = {
          _rowIndex: i + 1,
          items: []
        };

        // 基本情報
        LAB_RESULT_COLUMNS.BASE.forEach((col, idx) => {
          result[col.id] = data[i][idx];
        });

        // 検査項目
        const baseLen = LAB_RESULT_COLUMNS.BASE.length;
        LAB_RESULT_COLUMNS.ITEMS.forEach((item, idx) => {
          const value = data[i][baseLen + idx];
          if (value !== '' && value !== null && value !== undefined) {
            result.items.push({
              id: item.id,
              name: item.name,
              value: value,
              unit: item.unit,
              bmlCode: item.bmlCode
            });
            result[item.id] = value;
          }
        });

        console.log('★ 返却データ: items=' + result.items.length + '件');

        // Dateオブジェクトを文字列に変換（google.script.runでシリアライズ可能にする）
        const safeResult = JSON.parse(JSON.stringify(result, function(key, value) {
          if (value instanceof Date) {
            return value.toISOString();
          }
          return value;
        }));

        const returnObj = {
          success: true,
          data: safeResult,
          itemCount: safeResult.items.length
        };
        console.log('=== portalGetLabResultsByKarteNo 成功終了 ===');
        return returnObj;
      }
    }

    console.log('No match found for karteNo: ' + karteNo);
    return { success: true, data: null, message: '該当する検査結果がありません', _debug: { searchedKarteNo: karteNo, rowCount: data.length - 1 } };

  } catch (e) {
    console.error('portalGetLabResultsByKarteNo error:', e);
    return { success: false, error: e.message, data: null };
  }
}

/**
 * カルテNo検索のテスト（デバッグ用）
 * GASエディタから直接実行してログを確認
 */
function testGetLabResultsByKarteNo() {
  const testKarteNo = '999991';
  console.log('=== テスト開始: カルテNo=' + testKarteNo + ' ===');
  const result = portalGetLabResultsByKarteNo(testKarteNo);
  console.log('=== テスト結果 ===');
  console.log('success: ' + result.success);
  console.log('data: ' + (result.data ? 'あり' : 'なし'));
  if (result.data) {
    console.log('itemCount: ' + (result.data.items ? result.data.items.length : 0));
    if (result.data.items && result.data.items.length > 0) {
      console.log('最初の5項目: ' + JSON.stringify(result.data.items.slice(0, 5)));
    }
  }
  if (result.error) {
    console.log('error: ' + result.error);
  }
  if (result.message) {
    console.log('message: ' + result.message);
  }
  return result;
}

/**
 * 検査結果シートの診断（デバッグ用）
 * シートの構造、ヘッダー、データを確認
 * @returns {Object} 診断結果
 */
function portalDiagnoseLabResultSheet() {
  try {
    console.log('=== 検査結果シート診断開始 ===');
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('検査結果');

    if (!sheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // 期待されるヘッダー
    const expectedHeaders = getLabResultHeaders();

    // ヘッダー比較
    const headerAnalysis = [];
    for (let i = 0; i < Math.min(10, headers.length); i++) {
      const actual = String(headers[i] || '').trim();
      const expected = expectedHeaders[i] || '(未定義)';
      headerAnalysis.push({
        col: i,
        actual: actual,
        expected: expected,
        match: actual === expected
      });
    }

    // カルテNo列の検索（複数の方法で試行）
    const karteNoIdx1 = headers.indexOf('カルテNo');
    const karteNoIdx2 = headers.findIndex(h => String(h).trim().toUpperCase() === 'カルテNO');
    const karteNoIdx3 = headers.findIndex(h => String(h).includes('カルテ'));

    // カルテNo列のデータ収集
    const karteNoValues = [];
    const effectiveIdx = karteNoIdx1 >= 0 ? karteNoIdx1 : (karteNoIdx2 >= 0 ? karteNoIdx2 : karteNoIdx3);
    if (effectiveIdx >= 0) {
      for (let i = 1; i < Math.min(10, data.length); i++) {
        karteNoValues.push({
          row: i + 1,
          value: data[i][effectiveIdx],
          type: typeof data[i][effectiveIdx],
          asString: String(data[i][effectiveIdx])
        });
      }
    }

    // 全データ行の最初の5列を表示
    const sampleData = [];
    for (let i = 1; i < Math.min(5, data.length); i++) {
      sampleData.push({
        row: i + 1,
        values: data[i].slice(0, 5).map(v => String(v))
      });
    }

    const result = {
      success: true,
      sheetName: '検査結果',
      rowCount: data.length - 1,
      colCount: headers.length,
      expectedColCount: expectedHeaders.length,
      headerAnalysis: headerAnalysis,
      karteNoSearch: {
        exactMatch: karteNoIdx1,
        normalizedMatch: karteNoIdx2,
        partialMatch: karteNoIdx3,
        effectiveIndex: effectiveIdx
      },
      karteNoValues: karteNoValues,
      sampleData: sampleData,
      diagnosis: karteNoIdx1 < 0 ? 'ヘッダー不一致の可能性あり' : 'ヘッダー正常'
    };

    // ログ出力
    console.log('=== 診断結果 ===');
    console.log('データ行数: ' + result.rowCount);
    console.log('列数: ' + result.colCount + ' (期待値: ' + result.expectedColCount + ')');
    console.log('カルテNo列インデックス: exactMatch=' + karteNoIdx1 + ', normalized=' + karteNoIdx2 + ', partial=' + karteNoIdx3);
    console.log('診断: ' + result.diagnosis);
    console.log('ヘッダー(先頭10列): ' + JSON.stringify(headerAnalysis));
    console.log('カルテNo値: ' + JSON.stringify(karteNoValues));
    console.log('サンプルデータ: ' + JSON.stringify(sampleData));

    return result;

  } catch (e) {
    console.error('portalDiagnoseLabResultSheet error:', e);
    return { success: false, error: e.message };
  }
}

/**
 * 検査結果の項目定義を取得（UI用）
 * @returns {Object} 項目定義
 */
function portalGetLabItemDefinitions() {
  return {
    success: true,
    base: LAB_RESULT_COLUMNS.BASE,
    items: LAB_RESULT_COLUMNS.ITEMS,
    headers: getLabResultHeaders()
  };
}

/**
 * 取込済み患者一覧を取得（進捗状況付き）
 * @param {Object} options - フィルタオプション
 * @returns {Object} 患者一覧と進捗状況
 */
function portalGetPatientProgress(options) {
  try {
    const ss = getPortalSpreadsheet();
    const resultSheet = ss.getSheetByName('検査結果');
    const patientSheet = ss.getSheetByName('受診者マスタ');

    if (!resultSheet) {
      return { success: false, error: '検査結果シートが見つかりません', data: [] };
    }

    const resultData = resultSheet.getDataRange().getValues();
    const resultHeaders = resultData[0];

    // 進捗判定用の列インデックス
    const idCol = resultHeaders.indexOf('患者ID');
    const bloodCols = resultHeaders.filter(h => h.includes('血液') || h.includes('HbA1c') || h.includes('コレステロール'));
    const bodyCols = ['身長', '体重', 'BMI', '収縮期血圧'].map(h => resultHeaders.indexOf(h)).filter(i => i >= 0);
    const findingCol = resultHeaders.indexOf('総合判定');

    const patients = [];

    // 受診者マスタから名前を取得するためのマップ
    const patientNames = new Map();
    if (patientSheet) {
      const patientData = patientSheet.getDataRange().getValues();
      const patientHeaders = patientData[0];
      const pIdCol = patientHeaders.indexOf('患者ID');
      const pNameCol = patientHeaders.indexOf('氏名');

      for (let i = 1; i < patientData.length; i++) {
        patientNames.set(String(patientData[i][pIdCol]), patientData[i][pNameCol]);
      }
    }

    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];
      const patientId = row[idCol];

      if (!patientId) continue;

      // 日付フィルタ
      if (options && options.fromDate) {
        const examDate = row[resultHeaders.indexOf('検査日')];
        if (examDate && new Date(examDate) < new Date(options.fromDate)) continue;
      }

      // 進捗状況を計算
      const hasBlood = bloodCols.some(colName => {
        const idx = resultHeaders.indexOf(colName);
        return idx >= 0 && row[idx];
      }) || row[resultHeaders.indexOf('血液検査取込日')];

      const hasBody = bodyCols.some(idx => row[idx]);
      const hasFindings = findingCol >= 0 && row[findingCol];

      patients.push({
        patientId: patientId,
        name: patientNames.get(String(patientId)) || '（未登録）',
        examDate: row[resultHeaders.indexOf('検査日')] || '',
        bloodComplete: !!hasBlood,
        bodyComplete: !!hasBody,
        findingsComplete: !!hasFindings,
        allComplete: hasBlood && hasBody && hasFindings,
        _rowIndex: i + 1
      });
    }

    // 未完了を優先してソート
    patients.sort((a, b) => {
      if (a.allComplete !== b.allComplete) return a.allComplete ? 1 : -1;
      return 0;
    });

    return {
      success: true,
      data: patients,
      count: patients.length,
      stats: {
        total: patients.length,
        complete: patients.filter(p => p.allComplete).length,
        bloodPending: patients.filter(p => !p.bloodComplete).length,
        bodyPending: patients.filter(p => !p.bodyComplete).length,
        findingsPending: patients.filter(p => !p.findingsComplete).length
      }
    };
  } catch (error) {
    console.error('portalGetPatientProgress error:', error);
    return { success: false, error: error.message, data: [] };
  }
}

// ============================================================
// 帳票出力 API
// ============================================================

/**
 * 帳票を生成
 * @param {string} patientId - 患者ID
 * @param {string} reportType - 帳票種類
 * @returns {Object} 生成結果（URL等）
 */
function portalGenerateReport(patientId, reportType) {
  try {
    // 既存の帳票出力関数を呼び出し
    if (typeof generateReport === 'function') {
      return generateReport(patientId, reportType);
    }

    // 基本実装（プレースホルダー）
    return {
      success: true,
      message: '帳票生成は別途実装が必要です',
      patientId: patientId,
      reportType: reportType
    };
  } catch (error) {
    console.error('portalGenerateReport error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 一括帳票出力
 * @param {Array} patientIds - 患者IDリスト
 * @param {string} reportType - 帳票種類
 * @returns {Object} 出力結果
 */
function portalBatchGenerateReports(patientIds, reportType) {
  try {
    const results = [];
    for (const patientId of patientIds) {
      const result = portalGenerateReport(patientId, reportType);
      results.push({ patientId, ...result });
    }

    return {
      success: true,
      results: results,
      successCount: results.filter(r => r.success).length,
      errorCount: results.filter(r => !r.success).length
    };
  } catch (error) {
    console.error('portalBatchGenerateReports error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// 請求管理 API
// ============================================================

/**
 * 請求一覧を取得
 * @param {Object} criteria - 検索条件
 * @returns {Object} 請求一覧
 */
function portalGetBillingList(criteria) {
  try {
    // 既存のgetBillingList関数を呼び出し
    if (typeof getBillingList === 'function') {
      return getBillingList(criteria);
    }

    return { success: false, error: '請求管理機能が未実装です' };
  } catch (error) {
    console.error('portalGetBillingList error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 請求ステータスを更新
 * @param {string} patientId - 患者ID
 * @param {string} newStatus - 新ステータス
 * @returns {Object} 更新結果
 */
function portalUpdateBillingStatus(patientId, newStatus) {
  try {
    // 既存のupdateBillingStatus関数を呼び出し
    if (typeof updateBillingStatus === 'function') {
      return updateBillingStatus(patientId, newStatus);
    }

    return { success: false, error: '請求ステータス更新機能が未実装です' };
  } catch (error) {
    console.error('portalUpdateBillingStatus error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// ユーティリティ API
// ============================================================

/**
 * システム設定を取得
 * @returns {Object} システム設定
 */
function portalGetSettings() {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('設定');

    if (!sheet) {
      return {
        success: true,
        data: {
          organizationName: 'CDmedical',
          defaultCourse: '人間ドック',
          taxRate: 0.1
        }
      };
    }

    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        settings[data[i][0]] = data[i][1];
      }
    }

    return { success: true, data: settings };
  } catch (error) {
    console.error('portalGetSettings error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 選択肢マスタを取得
 * @param {string} masterType - マスタ種類
 * @returns {Object} 選択肢リスト
 */
function portalGetMasterOptions(masterType) {
  try {
    const options = {
      judgment: ['A', 'B', 'C', 'D', 'E', 'G'],
      ecgFindings: ['異常なし', '不完全右脚ブロック', '完全右脚ブロック', '左室肥大', '心房細動', 'その他'],
      xrayFindings: ['異常なし', '陳旧性病変', '胸膜肥厚', '側弯', 'その他'],
      courses: ['人間ドック', '定期健診', '労災二次検診', '生活習慣病予防健診'],
      billingStatus: ['未請求', '請求済', '入金済', 'キャンセル']
    };

    if (masterType && options[masterType]) {
      return { success: true, data: options[masterType] };
    }

    return { success: true, data: options };
  } catch (error) {
    console.error('portalGetMasterOptions error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 次の患者IDを取得（連続入力用）
 * @param {string} currentId - 現在の患者ID
 * @param {string} direction - 方向（next/prev）
 * @returns {Object} 次/前の患者ID
 */
function portalGetAdjacentPatient(currentId, direction) {
  try {
    const progressResult = portalGetPatientProgress({});
    if (!progressResult.success || progressResult.data.length === 0) {
      return { success: false, error: '患者データがありません' };
    }

    const patients = progressResult.data;
    const currentIndex = patients.findIndex(p => String(p.patientId) === String(currentId));

    if (currentIndex < 0) {
      return { success: false, error: '現在の患者が見つかりません' };
    }

    let targetIndex;
    if (direction === 'next') {
      targetIndex = currentIndex < patients.length - 1 ? currentIndex + 1 : 0;
    } else {
      targetIndex = currentIndex > 0 ? currentIndex - 1 : patients.length - 1;
    }

    return { success: true, data: patients[targetIndex] };
  } catch (error) {
    console.error('portalGetAdjacentPatient error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// デバッグ用テスト関数
// ============================================================

/**
 * デプロイ確認用テスト関数
 * ブラウザコンソールから google.script.run.testPortalApi() で呼び出し可能
 */
function testPortalApi() {
  return {
    success: true,
    message: 'portalApi.gs is deployed correctly',
    timestamp: new Date().toISOString(),
    version: '2025-12-16'
  };
}

/**
 * 検索機能デバッグ用
 * 固定値を返してAPI呼び出しが正常か確認
 */
function testSearchDebug() {
  return {
    success: true,
    data: [
      {
        '受診ID': 'TEST001',
        'ステータス': 'テスト',
        '受診日': '2025-12-16',
        '氏名': 'テスト 太郎',
        'カナ': 'テスト タロウ',
        '性別': '男',
        '生年月日': '1990-01-01',
        '年齢': 35,
        '受診コース': '人間ドック',
        '事業所名': 'テスト会社',
        '所属': '',
        '総合判定': 'A'
      }
    ],
    count: 1
  };
}

/**
 * 詳細検索デバッグ - GASエディタから実行
 * 11111 を検索して各ステップをログ出力
 */
function debugSearchDetailed() {
  console.log('========== 詳細検索デバッグ開始 ==========');

  const criteria = { patientId: '11111', name: '', kana: '' };
  console.log('検索条件:', JSON.stringify(criteria));

  // スプレッドシート接続
  const ss = getPortalSpreadsheet();
  console.log('スプレッドシートID:', ss.getId());

  const sheet = ss.getSheetByName('受診者マスタ');
  if (!sheet) {
    console.log('ERROR: 受診者マスタシートが見つかりません');
    return;
  }
  console.log('シート名:', sheet.getName());

  const data = sheet.getDataRange().getValues();
  console.log('総行数:', data.length);
  console.log('ヘッダー:', JSON.stringify(data[0]));

  // 検索条件を正規化
  const searchId = criteria.patientId ? normalizeString(criteria.patientId) : '';
  const searchName = criteria.name ? normalizeString(criteria.name) : '';
  const searchKana = criteria.kana ? normalizeKana(criteria.kana) : '';

  console.log('正規化後 - searchId:', searchId, ', searchName:', searchName, ', searchKana:', searchKana);

  console.log('\n========== 各行の検索処理 ==========');

  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = row[0];
    const rowName = row[3];
    const rowKana = row[4];

    // 空行スキップ
    if (!rowId) {
      console.log('行', i + 1, ': 空行 - スキップ');
      continue;
    }

    const normalizedRowId = normalizeString(rowId);
    const normalizedRowName = normalizeString(rowName);
    const normalizedRowKana = normalizeKana(rowKana);

    console.log('行', i + 1, ':');
    console.log('  元データ - ID:', rowId, '(型:', typeof rowId, '), 氏名:', rowName, ', カナ:', rowKana);
    console.log('  正規化後 - ID:', normalizedRowId, ', 氏名:', normalizedRowName, ', カナ:', normalizedRowKana);

    let match = true;

    if (searchId) {
      const idMatch = normalizedRowId.includes(searchId);
      console.log('  ID検索:', normalizedRowId, '.includes(', searchId, ') =', idMatch);
      if (!idMatch) match = false;
    }

    if (searchName && match) {
      const nameMatch = normalizedRowName.includes(searchName);
      console.log('  氏名検索:', normalizedRowName, '.includes(', searchName, ') =', nameMatch);
      if (!nameMatch) match = false;
    }

    if (searchKana && match) {
      const kanaMatch = normalizedRowKana.includes(searchKana);
      console.log('  カナ検索:', normalizedRowKana, '.includes(', searchKana, ') =', kanaMatch);
      if (!kanaMatch) match = false;
    }

    console.log('  → 結果:', match ? 'マッチ ✓' : '不一致 ✗');

    if (match) {
      results.push({
        '受診ID': rowId,
        '氏名': rowName,
        'カナ': rowKana
      });
    }
  }

  console.log('\n========== 検索結果 ==========');
  console.log('ヒット件数:', results.length);
  console.log('結果:', JSON.stringify(results, null, 2));

  return { success: true, data: results, count: results.length };
}

/**
 * Webアプリから呼び出して接続先スプレッドシートを確認
 */
function debugWebAppConnection() {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('受診者マスタ');
    const rowCount = sheet ? sheet.getLastRow() : 0;

    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetExists: !!sheet,
      rowCount: rowCount,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Webアプリコンテキストで検索の全ステップをデバッグ
 * ブラウザコンソールから呼び出し、結果をJSONで返す
 * @param {string} searchId - 検索する受診ID（デフォルト: '11111'）
 * @returns {Object} 各ステップの詳細情報
 */
function debugSearchFromWebApp(searchId) {
  const debug = {
    timestamp: new Date().toISOString(),
    steps: [],
    searchId: searchId || '11111',
    finalResult: null
  };

  try {
    // Step 1: スプレッドシート接続
    debug.steps.push({ step: 1, action: 'getPortalSpreadsheet 呼び出し' });
    const ss = getPortalSpreadsheet();
    debug.steps.push({
      step: 1,
      result: 'SUCCESS',
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName()
    });

    // Step 2: シート取得
    debug.steps.push({ step: 2, action: '受診者マスタシート取得' });
    const sheet = ss.getSheetByName('受診者マスタ');
    if (!sheet) {
      debug.steps.push({ step: 2, result: 'FAILED', error: 'シートが見つからない' });
      debug.finalResult = { success: false, error: 'シートが見つからない' };
      return debug;
    }
    debug.steps.push({ step: 2, result: 'SUCCESS', sheetName: sheet.getName() });

    // Step 3: データ取得
    debug.steps.push({ step: 3, action: 'getDataRange().getValues()' });
    const data = sheet.getDataRange().getValues();
    debug.steps.push({
      step: 3,
      result: 'SUCCESS',
      totalRows: data.length,
      headers: data[0] ? data[0].slice(0, 5).join(', ') + '...' : 'EMPTY'
    });

    // Step 4: データの先頭3行を取得（デバッグ用）
    debug.steps.push({ step: 4, action: 'サンプルデータ取得（先頭3行）' });
    const sampleData = [];
    for (let i = 1; i < Math.min(data.length, 4); i++) {
      sampleData.push({
        rowNum: i + 1,
        col0_受診ID: data[i][0],
        col0_type: typeof data[i][0],
        col3_氏名: data[i][3],
        col4_カナ: data[i][4]
      });
    }
    debug.steps.push({ step: 4, result: 'SUCCESS', sampleData: sampleData });

    // Step 5: 検索実行
    const targetId = searchId || '11111';
    debug.steps.push({ step: 5, action: '検索実行', targetId: targetId });

    const normalizedSearchId = normalizeString(targetId);
    debug.steps.push({ step: 5, normalizedSearchId: normalizedSearchId });

    const results = [];
    const matchLog = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowId = row[0];

      if (!rowId) continue;

      const normalizedRowId = normalizeString(rowId);
      const isMatch = normalizedRowId.includes(normalizedSearchId);

      matchLog.push({
        row: i + 1,
        rawId: rowId,
        normalizedId: normalizedRowId,
        searchId: normalizedSearchId,
        match: isMatch
      });

      if (isMatch) {
        results.push({
          '受診ID': row[0],
          '氏名': row[3],
          'カナ': row[4]
        });
      }

      // 最初の5行分のマッチログのみ保持
      if (matchLog.length > 5) break;
    }

    debug.steps.push({
      step: 5,
      result: 'SUCCESS',
      matchLog: matchLog,
      resultsCount: results.length,
      results: results
    });

    debug.finalResult = {
      success: true,
      data: results,
      count: results.length
    };

  } catch (error) {
    debug.steps.push({
      step: 'ERROR',
      error: error.message,
      stack: error.stack
    });
    debug.finalResult = { success: false, error: error.message };
  }

  return debug;
}

/**
 * 受診者マスタシートを正しい構造で再構築
 * 既存シートはバックアップとして保持
 */
function rebuildPatientMasterSheet() {
  const SPREADSHEET_ID = '16KtctyT2gd7oJZdcu84kUtuP-D9jB9KtLxxzxXx_wdk';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. 既存シートをバックアップ
  const oldSheet = ss.getSheetByName('受診者マスタ');
  if (oldSheet) {
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    oldSheet.setName('受診者マスタ_backup_' + timestamp);
    console.log('既存シートをバックアップしました: 受診者マスタ_backup_' + timestamp);
  }
  
  // 2. 新しいシートを作成
  const newSheet = ss.insertSheet('受診者マスタ');
  
  // 3. 正しいヘッダーを設定
  const headers = [
    '受診ID',        // A
    'ステータス',     // B
    '受診日',        // C
    '氏名',          // D
    'カナ',          // E
    '性別',          // F
    '生年月日',      // G
    '年齢',          // H
    '受診コース',    // I
    '事業所名',      // J
    '所属',          // K
    '総合判定',      // L
    'CSV取込日時',   // M
    '最終更新日時',  // N
    '出力日時'       // O
  ];
  
  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 4. ヘッダー行の書式設定
  const headerRange = newSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 5. A列（受診ID）を書式なしテキストに設定
  const idColumn = newSheet.getRange('A:A');
  idColumn.setNumberFormat('@');  // テキスト形式
  
  // 6. 日付列の書式設定
  newSheet.getRange('C:C').setNumberFormat('yyyy/mm/dd');  // 受診日
  newSheet.getRange('G:G').setNumberFormat('yyyy/mm/dd');  // 生年月日
  newSheet.getRange('M:M').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // CSV取込日時
  newSheet.getRange('N:N').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // 最終更新日時
  newSheet.getRange('O:O').setNumberFormat('yyyy/mm/dd hh:mm:ss');  // 出力日時
  
  // 7. 列幅を調整
  newSheet.setColumnWidth(1, 120);   // 受診ID
  newSheet.setColumnWidth(2, 80);    // ステータス
  newSheet.setColumnWidth(3, 100);   // 受診日
  newSheet.setColumnWidth(4, 120);   // 氏名
  newSheet.setColumnWidth(5, 120);   // カナ
  newSheet.setColumnWidth(6, 50);    // 性別
  newSheet.setColumnWidth(7, 100);   // 生年月日
  newSheet.setColumnWidth(8, 50);    // 年齢
  newSheet.setColumnWidth(9, 150);   // 受診コース
  newSheet.setColumnWidth(10, 150);  // 事業所名
  newSheet.setColumnWidth(11, 100);  // 所属
  newSheet.setColumnWidth(12, 80);   // 総合判定
  
  // 8. テストデータを1件追加
  const testData = [
    '20251219-0001',      // 受診ID（テキスト）
    '入力中',             // ステータス
    new Date(2025, 11, 19), // 受診日
    'テスト 太郎',        // 氏名
    'テスト タロウ',      // カナ
    '男',                 // 性別
    new Date(1980, 0, 15),  // 生年月日
    44,                   // 年齢
    '生活習慣病ドック',   // 受診コース
    'テスト株式会社',     // 事業所名
    '営業部',             // 所属
    '',                   // 総合判定
    '',                   // CSV取込日時
    new Date(),           // 最終更新日時
    ''                    // 出力日時
  ];
  
  newSheet.getRange(2, 1, 1, testData.length).setValues([testData]);
  
  // 9. 行の固定（ヘッダー）
  newSheet.setFrozenRows(1);
  
  // 10. シートを先頭に移動
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);
  
  console.log('✅ 受診者マスタシートを再構築しました');
  console.log('ヘッダー: ' + headers.join(', '));
  console.log('テストデータ1件を追加しました');

  return {
    success: true,
    message: '受診者マスタシートを再構築しました',
    backupSheet: oldSheet ? oldSheet.getName() : null
  };
}

// ============================================================
// Phase 2: 結果入力画面 API（iD-Heart準拠）
// ============================================================

/**
 * 結果入力画面のデータを一括取得
 * @param {string} visitId - 受診ID (YYYYMMDD-NNN形式)
 * @returns {Object} 入力画面に必要な全データ
 */
function getInputScreenData(visitId) {
  try {
    const ss = getPortalSpreadsheet();
    const result = {
      visit: null,
      patient: null,
      items: [],
      existingResults: [],
      previousResults: [],
      selectOptions: {},
      judgmentCriteria: []
    };

    // 1. 受診記録を取得
    const visitSheet = ss.getSheetByName('受診記録');
    if (!visitSheet) {
      return { success: false, error: '受診記録シートが見つかりません' };
    }

    const visitData = visitSheet.getDataRange().getValues();
    let patientId = null;
    let visitDate = null;

    for (let i = 1; i < visitData.length; i++) {
      if (safeString(visitData[i][VISIT_COL.VISIT_ID]) === String(visitId)) {
        const row = visitData[i];
        result.visit = {
          visitId: safeString(row[VISIT_COL.VISIT_ID]),
          patientId: safeString(row[VISIT_COL.PATIENT_ID]),
          examType: safeString(row[VISIT_COL.EXAM_TYPE]),
          course: safeString(row[VISIT_COL.COURSE]),
          visitDate: formatDateToString(row[VISIT_COL.VISIT_DATE]),
          age: safeNumber(row[VISIT_COL.AGE]),
          judgment: safeString(row[VISIT_COL.JUDGMENT]),
          status: safeString(row[VISIT_COL.STATUS])
        };
        patientId = safeString(row[VISIT_COL.PATIENT_ID]);
        visitDate = row[VISIT_COL.VISIT_DATE];
        break;
      }
    }

    if (!result.visit) {
      return { success: false, error: '受診記録が見つかりません: ' + visitId };
    }

    // 2. 受診者情報を取得
    const patientSheet = ss.getSheetByName('受診者マスタ');
    if (patientSheet && patientId) {
      const patientData = patientSheet.getDataRange().getValues();
      for (let i = 1; i < patientData.length; i++) {
        if (safeString(patientData[i][PATIENT_COL.PATIENT_ID]) === patientId) {
          const row = patientData[i];
          result.patient = {
            patientId: safeString(row[PATIENT_COL.PATIENT_ID]),
            name: safeString(row[PATIENT_COL.NAME]),
            kana: safeString(row[PATIENT_COL.KANA]),
            birthDate: formatDateToString(row[PATIENT_COL.BIRTHDATE]),
            gender: safeString(row[PATIENT_COL.GENDER]),
            company: safeString(row[PATIENT_COL.COMPANY])
          };
          break;
        }
      }
    }

    // 3. 検査項目マスタを取得
    const itemSheet = ss.getSheetByName('検査項目マスタ');
    if (itemSheet && itemSheet.getLastRow() > 1) {
      const itemData = itemSheet.getDataRange().getValues();
      const itemHeaders = itemData[0];

      for (let i = 1; i < itemData.length; i++) {
        const row = itemData[i];
        if (!row[0]) continue;  // 空行スキップ

        result.items.push({
          itemId: safeString(row[0]),
          itemCode: safeString(row[1]),
          itemName: safeString(row[2]),
          category: safeString(row[3]),
          subCategory: safeString(row[4]),
          dataType: safeString(row[5]),
          unit: safeString(row[6]),
          displayOrder: safeNumber(row[7]) || i,
          isActive: row[8] !== false
        });
      }

      // 表示順でソート
      result.items.sort((a, b) => a.displayOrder - b.displayOrder);
    }

    // 3.5. BML検査項目をitemsに追加（マスタにない場合のフォールバック）
    const existingItemIds = new Set(result.items.map(i => i.itemId || i.itemCode));
    LAB_RESULT_COLUMNS.ITEMS.forEach((bmlItem, idx) => {
      if (!existingItemIds.has(bmlItem.id)) {
        result.items.push({
          itemId: bmlItem.id,
          itemCode: bmlItem.bmlCode,
          itemName: bmlItem.name,
          category: bmlItem.category,
          subCategory: bmlItem.category,
          dataType: 'numeric',
          unit: bmlItem.unit,
          displayOrder: 1000 + idx,
          isActive: true,
          source: 'BML'
        });
      }
    });

    // 4. 今回の検査結果を取得（縦持ち形式）
    const resultSheet = ss.getSheetByName('検査結果');
    if (resultSheet && resultSheet.getLastRow() > 1) {
      const resultData = resultSheet.getDataRange().getValues();

      for (let i = 1; i < resultData.length; i++) {
        const row = resultData[i];
        if (safeString(row[1]) === String(visitId)) {  // 受診ID列
          result.existingResults.push({
            resultId: safeString(row[0]),
            visitId: safeString(row[1]),
            itemId: safeString(row[2]),
            value: safeString(row[3]),
            numericValue: safeNumber(row[4]),
            judgment: safeString(row[5]),
            notes: safeString(row[6])
          });
        }
      }
    }

    // 4.5. BML横持ちデータから検査結果を取得（LAB_RESULT_COLUMNS形式）
    if (patientId && resultSheet && resultSheet.getLastRow() > 1) {
      const bmlData = resultSheet.getDataRange().getValues();
      const bmlHeaders = bmlData[0];

      // ヘッダーから「受診者ID」列を探す（横持ちフォーマット判定）
      const patientIdColIdx = bmlHeaders.indexOf('受診者ID');

      if (patientIdColIdx >= 0) {
        // 横持ちフォーマットの場合
        for (let i = 1; i < bmlData.length; i++) {
          if (String(bmlData[i][patientIdColIdx]).trim() === String(patientId).trim()) {
            // 各検査項目をexistingResultsに追加
            LAB_RESULT_COLUMNS.ITEMS.forEach((item, idx) => {
              const colIdx = LAB_RESULT_COLUMNS.BASE.length + idx;
              const value = bmlData[i][colIdx];

              if (value !== '' && value !== null && value !== undefined) {
                // 既存のexistingResultsに同じitemIdがあるか確認
                const existingIdx = result.existingResults.findIndex(r => r.itemId === item.id);

                if (existingIdx < 0) {
                  // 新規追加
                  result.existingResults.push({
                    resultId: 'BML_' + item.id,
                    visitId: visitId,
                    itemId: item.id,
                    itemCode: item.bmlCode,
                    itemName: item.name,
                    value: String(value),
                    numericValue: typeof value === 'number' ? value : parseFloat(value) || null,
                    unit: item.unit,
                    judgment: '',
                    notes: '',
                    source: 'BML_CSV'
                  });
                }
              }
            });
            break;
          }
        }
      }
    }

    // 5. 前回の検査結果を取得
    if (patientId && visitDate) {
      // 同一患者の前回受診を検索
      let previousVisitId = null;
      let previousVisitDate = null;

      for (let i = 1; i < visitData.length; i++) {
        const row = visitData[i];
        if (safeString(row[VISIT_COL.PATIENT_ID]) === patientId &&
            safeString(row[VISIT_COL.VISIT_ID]) !== visitId) {
          const vDate = row[VISIT_COL.VISIT_DATE];
          if (vDate && vDate < visitDate) {
            if (!previousVisitDate || vDate > previousVisitDate) {
              previousVisitDate = vDate;
              previousVisitId = safeString(row[VISIT_COL.VISIT_ID]);
            }
          }
        }
      }

      // 前回の検査結果を取得
      if (previousVisitId && resultSheet) {
        const resultData = resultSheet.getDataRange().getValues();
        for (let i = 1; i < resultData.length; i++) {
          const row = resultData[i];
          if (safeString(row[1]) === previousVisitId) {
            result.previousResults.push({
              itemId: safeString(row[2]),
              value: safeString(row[3]),
              judgment: safeString(row[5]),
              notes: safeString(row[6])
            });
          }
        }
      }
    }

    // 6. 選択肢マスタを取得
    const optionSheet = ss.getSheetByName('選択肢マスタ');
    if (optionSheet && optionSheet.getLastRow() > 1) {
      const optionData = optionSheet.getDataRange().getValues();
      for (let i = 1; i < optionData.length; i++) {
        const row = optionData[i];
        const itemId = safeString(row[0]);
        if (itemId) {
          result.selectOptions[itemId] = safeString(row[1]).split(',').map(s => s.trim());
        }
      }
    }

    // 7. 判定基準マスタを取得
    const criteriaSheet = ss.getSheetByName('判定基準マスタ');
    if (criteriaSheet && criteriaSheet.getLastRow() > 1) {
      const criteriaData = criteriaSheet.getDataRange().getValues();
      for (let i = 1; i < criteriaData.length; i++) {
        const row = criteriaData[i];
        if (!row[0]) continue;
        result.judgmentCriteria.push({
          itemId: safeString(row[0]),
          itemName: safeString(row[1]),
          gender: safeString(row[2]),
          unit: safeString(row[3]),
          aMin: safeNumber(row[4]),
          aMax: safeNumber(row[5]),
          bMin: safeNumber(row[6]),
          bMax: safeNumber(row[7]),
          cMin: safeNumber(row[8]),
          cMax: safeNumber(row[9]),
          dMin: safeNumber(row[10]),
          dMax: safeNumber(row[11])
        });
      }
    }

    return {
      success: true,
      data: result,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('getInputScreenData error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 検査結果を一括保存（バッチ処理）
 * @param {string} visitId - 受診ID
 * @param {Array} results - 検査結果配列 [{itemId, value, judgment, notes}, ...]
 * @returns {Object} 保存結果
 */
function saveTestResultsBatch(visitId, results) {
  try {
    if (!visitId || !results || !Array.isArray(results)) {
      return { success: false, error: 'パラメータが不正です' };
    }

    const ss = getPortalSpreadsheet();
    const resultSheet = ss.getSheetByName('検査結果');

    if (!resultSheet) {
      return { success: false, error: '検査結果シートが見つかりません' };
    }

    // 既存の結果を取得（visitIdでフィルタ）
    const existingData = resultSheet.getDataRange().getValues();
    const existingMap = {};  // itemId -> rowIndex

    for (let i = 1; i < existingData.length; i++) {
      if (safeString(existingData[i][1]) === String(visitId)) {
        const itemId = safeString(existingData[i][2]);
        existingMap[itemId] = i + 1;  // 1-indexed row number
      }
    }

    const now = new Date();
    let updatedCount = 0;
    let insertedCount = 0;
    const errors = [];

    // バッチ処理（100件ずつ）
    const batchSize = 100;
    for (let batch = 0; batch < results.length; batch += batchSize) {
      const batchResults = results.slice(batch, batch + batchSize);

      for (const item of batchResults) {
        try {
          const itemId = safeString(item.itemId);
          if (!itemId) continue;

          // 数値変換
          let numericValue = '';
          if (item.value !== '' && item.value !== null && item.value !== undefined) {
            const parsed = parseFloat(item.value);
            if (!isNaN(parsed)) {
              numericValue = parsed;
            }
          }

          if (existingMap[itemId]) {
            // 既存行を更新
            const rowIndex = existingMap[itemId];
            resultSheet.getRange(rowIndex, 4).setValue(item.value || '');  // D: 値
            resultSheet.getRange(rowIndex, 5).setValue(numericValue);       // E: 数値変換
            resultSheet.getRange(rowIndex, 6).setValue(item.judgment || ''); // F: 判定
            resultSheet.getRange(rowIndex, 7).setValue(item.notes || '');    // G: 所見
            resultSheet.getRange(rowIndex, 8).setValue(now);                 // H: 更新日時
            updatedCount++;
          } else {
            // 新規行を追加
            const resultId = generateNextResultId(resultSheet);
            const newRow = [
              resultId,           // A: 結果ID
              visitId,            // B: 受診ID
              itemId,             // C: 項目ID
              item.value || '',   // D: 値
              numericValue,       // E: 数値変換
              item.judgment || '', // F: 判定
              item.notes || '',   // G: 所見
              now                  // H: 作成日時
            ];
            resultSheet.appendRow(newRow);
            existingMap[itemId] = resultSheet.getLastRow();
            insertedCount++;
          }
        } catch (itemError) {
          errors.push({ itemId: item.itemId, error: itemError.message });
        }
      }
    }

    return {
      success: true,
      message: `保存完了: 更新${updatedCount}件, 新規${insertedCount}件`,
      updatedCount: updatedCount,
      insertedCount: insertedCount,
      errors: errors
    };

  } catch (error) {
    console.error('saveTestResultsBatch error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 次の結果IDを生成
 * @param {Sheet} sheet - 検査結果シート
 * @returns {string} R00001形式のID
 */
function generateNextResultId(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return 'R00001';
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let maxNum = 0;

  for (const row of data) {
    const id = String(row[0]);
    if (id.startsWith('R')) {
      const num = parseInt(id.replace('R', ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  return `R${String(maxNum + 1).padStart(5, '0')}`;
}

/**
 * 即時判定計算
 * @param {string} itemId - 項目ID
 * @param {*} value - 検査値
 * @param {string} gender - 性別 (M/F)
 * @returns {Object} 判定結果 {judgment, color, label}
 */
function calculateJudgment(itemId, value, gender) {
  try {
    // 値がない場合
    if (value === '' || value === null || value === undefined) {
      return { judgment: '', color: '', label: '' };
    }

    const numValue = parseFloat(value);

    // 数値でない場合（定性検査など）
    if (isNaN(numValue)) {
      // 定性検査の判定
      const qualitativeJudgments = {
        '(-)': 'A',
        '-': 'A',
        '陰性': 'A',
        '(±)': 'B',
        '±': 'B',
        '(+)': 'C',
        '+': 'C',
        '(2+)': 'D',
        '2+': 'D',
        '(3+)': 'D',
        '3+': 'D',
        '陽性': 'D'
      };
      const judgment = qualitativeJudgments[String(value)] || '';
      return getJudgmentInfo(judgment);
    }

    // 判定基準を取得
    const ss = getPortalSpreadsheet();
    const criteriaSheet = ss.getSheetByName('判定基準マスタ');

    if (!criteriaSheet || criteriaSheet.getLastRow() < 2) {
      return { judgment: '', color: '', label: '判定基準なし' };
    }

    const criteriaData = criteriaSheet.getDataRange().getValues();

    // 該当項目の判定基準を検索
    for (let i = 1; i < criteriaData.length; i++) {
      const row = criteriaData[i];
      if (safeString(row[0]) === itemId) {
        const criteriaGender = safeString(row[2]);

        // 性別チェック（空なら共通、Mなら男性、Fなら女性）
        if (criteriaGender && criteriaGender !== gender) {
          continue;
        }

        const aMin = safeNumber(row[4]);
        const aMax = safeNumber(row[5]);
        const bMin = safeNumber(row[6]);
        const bMax = safeNumber(row[7]);
        const cMin = safeNumber(row[8]);
        const cMax = safeNumber(row[9]);
        const dMin = safeNumber(row[10]);
        const dMax = safeNumber(row[11]);

        // 判定ロジック
        let judgment = '';

        // A判定
        if (aMin !== '' && aMax !== '') {
          if (numValue >= aMin && numValue <= aMax) {
            judgment = 'A';
          }
        }

        // B判定
        if (!judgment && bMin !== '' && bMax !== '') {
          if (numValue >= bMin && numValue <= bMax) {
            judgment = 'B';
          }
        }

        // C判定
        if (!judgment && cMin !== '' && cMax !== '') {
          if (numValue >= cMin && numValue <= cMax) {
            judgment = 'C';
          }
        }

        // D判定（範囲外）
        if (!judgment && (dMin !== '' || dMax !== '')) {
          if ((dMin !== '' && numValue < dMin) || (dMax !== '' && numValue > dMax)) {
            judgment = 'D';
          } else if (!judgment) {
            judgment = 'D';
          }
        }

        // いずれにも該当しない場合
        if (!judgment) {
          judgment = 'D';  // 範囲外はD判定
        }

        return getJudgmentInfo(judgment);
      }
    }

    return { judgment: '', color: '', label: '判定基準なし' };

  } catch (error) {
    console.error('calculateJudgment error:', error);
    return { judgment: '', color: '', label: 'エラー' };
  }
}

/**
 * 判定情報を取得
 * @param {string} judgment - 判定 (A/B/C/D/E/F)
 * @returns {Object} {judgment, color, label}
 */
function getJudgmentInfo(judgment) {
  const info = {
    A: { color: '#e8f5e9', textColor: '#2e7d32', label: '異常なし' },
    B: { color: '#fff8e1', textColor: '#f57f17', label: '軽度異常' },
    C: { color: '#fff3e0', textColor: '#e65100', label: '要経過観察' },
    D: { color: '#ffebee', textColor: '#c62828', label: '要精密検査' },
    E: { color: '#f3e5f5', textColor: '#6a1b9a', label: '治療中' },
    F: { color: '#e3f2fd', textColor: '#1565c0', label: '経過観察中' }
  };

  if (info[judgment]) {
    return {
      judgment: judgment,
      color: info[judgment].color,
      textColor: info[judgment].textColor,
      label: info[judgment].label
    };
  }

  return { judgment: '', color: '', textColor: '', label: '' };
}

/**
 * 3レベル判定を取得
 * @param {string} visitId - 受診ID
 * @returns {Object} 3レベル判定結果
 */
function getThreeLevelJudgment(visitId) {
  try {
    const ss = getPortalSpreadsheet();
    const result = {
      level1: {},  // 項目別判定 {itemId: judgment}
      level2: {},  // カテゴリ別判定 {category: judgment}
      level3: ''   // 総合判定
    };

    // 検査結果を取得
    const resultSheet = ss.getSheetByName('検査結果');
    if (!resultSheet || resultSheet.getLastRow() < 2) {
      return { success: false, error: '検査結果がありません' };
    }

    const resultData = resultSheet.getDataRange().getValues();

    // Level 1: 項目別判定を収集
    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];
      if (safeString(row[1]) === String(visitId)) {
        const itemId = safeString(row[2]);
        const judgment = safeString(row[5]);
        if (itemId && judgment) {
          result.level1[itemId] = judgment;
        }
      }
    }

    // 検査項目マスタからカテゴリを取得
    const itemSheet = ss.getSheetByName('検査項目マスタ');
    const categoryMap = {};  // itemId -> category

    if (itemSheet && itemSheet.getLastRow() > 1) {
      const itemData = itemSheet.getDataRange().getValues();
      for (let i = 1; i < itemData.length; i++) {
        const row = itemData[i];
        categoryMap[safeString(row[0])] = safeString(row[3]);  // category列
      }
    }

    // Level 2: カテゴリ別に最悪の判定を集計
    const categoryJudgments = {};  // category -> [judgments]
    const judgmentOrder = { 'D': 4, 'C': 3, 'B': 2, 'E': 1, 'F': 1, 'A': 0 };

    for (const [itemId, judgment] of Object.entries(result.level1)) {
      const category = categoryMap[itemId] || '未分類';
      if (!categoryJudgments[category]) {
        categoryJudgments[category] = [];
      }
      categoryJudgments[category].push(judgment);
    }

    for (const [category, judgments] of Object.entries(categoryJudgments)) {
      // 最悪の判定を取得
      let worstJudgment = 'A';
      for (const j of judgments) {
        if ((judgmentOrder[j] || 0) > (judgmentOrder[worstJudgment] || 0)) {
          worstJudgment = j;
        }
      }
      result.level2[category] = worstJudgment;
    }

    // Level 3: 総合判定（全カテゴリの最悪判定）
    let overallWorst = 'A';
    for (const judgment of Object.values(result.level2)) {
      if ((judgmentOrder[judgment] || 0) > (judgmentOrder[overallWorst] || 0)) {
        overallWorst = judgment;
      }
    }
    result.level3 = overallWorst;

    return {
      success: true,
      data: result
    };

  } catch (error) {
    console.error('getThreeLevelJudgment error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 当日受診者リストを取得
 * @param {string} dateStr - 日付文字列 (YYYY-MM-DD)
 * @returns {Object} 受診者リスト
 */
function getTodayVisitors(dateStr) {
  try {
    const ss = getPortalSpreadsheet();
    const visitSheet = ss.getSheetByName('受診記録');
    const patientSheet = ss.getSheetByName('受診者マスタ');

    if (!visitSheet) {
      return { success: false, error: '受診記録シートが見つかりません' };
    }

    const targetDate = dateStr ? new Date(dateStr) : new Date();
    const targetDateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyyMMdd');

    const visitData = visitSheet.getDataRange().getValues();
    const visitors = [];

    // 患者情報をマップ化
    const patientMap = {};
    if (patientSheet && patientSheet.getLastRow() > 1) {
      const patientData = patientSheet.getDataRange().getValues();
      for (let i = 1; i < patientData.length; i++) {
        const patientId = safeString(patientData[i][PATIENT_COL.PATIENT_ID]);
        patientMap[patientId] = {
          name: safeString(patientData[i][PATIENT_COL.NAME]),
          kana: safeString(patientData[i][PATIENT_COL.KANA]),
          gender: safeString(patientData[i][PATIENT_COL.GENDER])
        };
      }
    }

    for (let i = 1; i < visitData.length; i++) {
      const row = visitData[i];
      const visitDate = row[VISIT_COL.VISIT_DATE];

      if (visitDate) {
        const visitDateStr = Utilities.formatDate(new Date(visitDate), 'Asia/Tokyo', 'yyyyMMdd');
        if (visitDateStr === targetDateStr) {
          const patientId = safeString(row[VISIT_COL.PATIENT_ID]);
          const patient = patientMap[patientId] || {};

          visitors.push({
            visitId: safeString(row[VISIT_COL.VISIT_ID]),
            patientId: patientId,
            name: patient.name || '',
            kana: patient.kana || '',
            gender: patient.gender || '',
            course: safeString(row[VISIT_COL.COURSE]),
            status: safeString(row[VISIT_COL.STATUS]),
            judgment: safeString(row[VISIT_COL.JUDGMENT]),
            age: safeNumber(row[VISIT_COL.AGE])
          });
        }
      }
    }

    // カナでソート
    visitors.sort((a, b) => (a.kana || '').localeCompare(b.kana || '', 'ja'));

    return {
      success: true,
      data: visitors,
      count: visitors.length,
      date: formatDateToString(targetDate)
    };

  } catch (error) {
    console.error('getTodayVisitors error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// Phase 3: 所見生成・保存 API
// ============================================================

/**
 * 所見を自動生成して保存
 * @param {string} visitId - 受診ID
 * @returns {Object} 生成された所見データ
 */
function generateAndSaveFindings(visitId) {
  try {
    if (!visitId) {
      return { success: false, error: '受診IDが必要です' };
    }

    // 3レベル判定を取得
    const judgmentResult = getThreeLevelJudgment(visitId);
    if (!judgmentResult.success) {
      return { success: false, error: '判定取得失敗: ' + judgmentResult.error };
    }

    const ss = getPortalSpreadsheet();
    const findingsSheet = ss.getSheetByName(DB_CONFIG.SHEETS.FINDINGS);
    if (!findingsSheet) {
      return { success: false, error: 'T_所見シートが見つかりません' };
    }

    // 既存レコード検索
    const lastRow = findingsSheet.getLastRow();
    let existingRow = 0;
    if (lastRow >= 2) {
      const ids = findingsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (safeString(ids[i][0]) === visitId) {
          existingRow = i + 2;
          break;
        }
      }
    }

    // カテゴリ別所見を生成
    const level2 = judgmentResult.level2 || {};
    const findings = {
      circulatory: generateCategoryFindingText('循環器系', level2['血圧'] || level2['循環器'] || 'A'),
      digestive: generateCategoryFindingText('消化器系', level2['肝胆膵機能'] || level2['消化器'] || 'A'),
      metabolicSugar: generateCategoryFindingText('代謝系（糖）', level2['糖代謝'] || 'A'),
      metabolicLipid: generateCategoryFindingText('代謝系（脂質）', level2['脂質検査'] || level2['脂質'] || 'A'),
      renal: generateCategoryFindingText('腎機能', level2['腎機能'] || 'A'),
      blood: generateCategoryFindingText('血液系', level2['血液学検査'] || level2['血液学'] || 'A'),
      other: generateCategoryFindingText('その他', level2['身体測定'] || 'A')
    };

    // 総合所見を生成
    const overallJudgment = judgmentResult.level3 || 'A';
    const combinedFindings = generateCombinedFindings(findings, overallJudgment);

    const now = new Date();
    const rowData = [
      visitId,                    // A: 受診ID
      '',                         // B: 既往歴
      '',                         // C: 自覚症状
      '',                         // D: 他覚症状
      findings.circulatory,       // E: 所見_循環器系
      findings.digestive,         // F: 所見_消化器系
      findings.metabolicSugar,    // G: 所見_代謝系_糖
      findings.metabolicLipid,    // H: 所見_代謝系_脂質
      findings.renal,             // I: 所見_腎機能
      findings.blood,             // J: 所見_血液系
      findings.other,             // K: 所見_その他
      combinedFindings,           // L: 総合所見
      '',                         // M: 心電図_今回
      '',                         // N: 心電図_前回
      '',                         // O: 腹部超音波_今回
      '',                         // P: 腹部超音波_前回
      now                         // Q: 更新日時
    ];

    if (existingRow > 0) {
      // 既存レコードを更新（所見列のみ E-L列）
      findingsSheet.getRange(existingRow, 5, 1, 8).setValues([[
        findings.circulatory,
        findings.digestive,
        findings.metabolicSugar,
        findings.metabolicLipid,
        findings.renal,
        findings.blood,
        findings.other,
        combinedFindings
      ]]);
      findingsSheet.getRange(existingRow, 17).setValue(now);
    } else {
      // 新規レコードを追加
      findingsSheet.appendRow(rowData);
    }

    return {
      success: true,
      message: existingRow > 0 ? '所見を更新しました' : '所見を生成しました',
      findings: findings,
      combined: combinedFindings,
      overallJudgment: overallJudgment
    };

  } catch (error) {
    console.error('generateAndSaveFindings error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * カテゴリ別所見テキストを生成
 * @param {string} category - カテゴリ名
 * @param {string} judgment - 判定 (A/B/C/D/E/F)
 * @returns {string} 所見テキスト
 */
function generateCategoryFindingText(category, judgment) {
  if (!judgment || judgment === 'A') {
    return '';
  }

  // 所見テンプレートマスタから取得を試みる
  try {
    const ss = getPortalSpreadsheet();
    const templateSheet = ss.getSheetByName(DB_CONFIG.SHEETS.FINDING_TEMPLATE);
    if (templateSheet && templateSheet.getLastRow() >= 2) {
      const templateData = templateSheet.getDataRange().getValues();
      for (let i = 1; i < templateData.length; i++) {
        const row = templateData[i];
        if (safeString(row[2]) === category && safeString(row[3]) === judgment) {
          return safeString(row[4]); // E列: 所見テンプレート
        }
      }
    }
  } catch (e) {
    console.warn('テンプレート取得エラー:', e);
  }

  // デフォルトテンプレート
  const defaultTemplates = {
    '循環器系': {
      'B': '血圧がやや高めです。減塩と適度な運動を心がけてください。',
      'C': '血圧が高めです。生活習慣の見直しと定期的な測定をお勧めします。',
      'D': '血圧が基準値を大きく超えています。医療機関での受診をお勧めします。',
      'E': '血圧について治療中です。引き続き治療を継続してください。',
      'F': '血圧について経過観察中です。定期的な測定を継続してください。'
    },
    '消化器系': {
      'B': '肝機能に軽度の異常があります。飲酒量の見直しをお勧めします。',
      'C': '肝機能に異常があります。精密検査をお勧めします。',
      'D': '肝機能に明らかな異常があります。早めの医療機関受診をお勧めします。',
      'E': '肝機能について治療中です。引き続き治療を継続してください。',
      'F': '肝機能について経過観察中です。定期的な検査を継続してください。'
    },
    '代謝系（糖）': {
      'B': '血糖値がやや高めです。糖質の摂取を控えめにしてください。',
      'C': '血糖値が高めです。糖尿病の精密検査をお勧めします。',
      'D': '血糖値が基準値を大きく超えています。糖尿病の治療が必要です。',
      'E': '糖尿病について治療中です。引き続き治療を継続してください。',
      'F': '血糖値について経過観察中です。定期的な検査を継続してください。'
    },
    '代謝系（脂質）': {
      'B': '脂質代謝に軽度の異常があります。食事内容の見直しをお勧めします。',
      'C': '脂質代謝に異常があります。動物性脂肪を控え、運動習慣を取り入れてください。',
      'D': '脂質代謝に明らかな異常があります。医療機関での治療をお勧めします。',
      'E': '脂質異常について治療中です。引き続き治療を継続してください。',
      'F': '脂質について経過観察中です。定期的な検査を継続してください。'
    },
    '腎機能': {
      'B': '腎機能に軽度の異常があります。水分を十分に摂取してください。',
      'C': '腎機能に異常があります。定期的な経過観察をお勧めします。',
      'D': '腎機能に明らかな異常があります。専門医の受診をお勧めします。',
      'E': '腎機能について治療中です。引き続き治療を継続してください。',
      'F': '腎機能について経過観察中です。定期的な検査を継続してください。'
    },
    '血液系': {
      'B': '血液検査に軽度の異常があります。経過観察をお勧めします。',
      'C': '血液検査に異常があります。精密検査をお勧めします。',
      'D': '血液検査に明らかな異常があります。早めの受診をお勧めします。',
      'E': '血液疾患について治療中です。引き続き治療を継続してください。',
      'F': '血液検査について経過観察中です。定期的な検査を継続してください。'
    },
    'その他': {
      'B': '軽度の異常が認められます。経過観察をお勧めします。',
      'C': '異常が認められます。精密検査をお勧めします。',
      'D': '明らかな異常が認められます。医療機関の受診をお勧めします。',
      'E': '治療中の項目があります。引き続き治療を継続してください。',
      'F': '経過観察中の項目があります。定期的な検査を継続してください。'
    }
  };

  const templates = defaultTemplates[category] || defaultTemplates['その他'];
  return templates[judgment] || '';
}

/**
 * 総合所見を生成
 * @param {Object} findings - カテゴリ別所見
 * @param {string} overallJudgment - 総合判定
 * @returns {string} 総合所見
 */
function generateCombinedFindings(findings, overallJudgment) {
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
    return '今回の検査では特に問題は認められませんでした。引き続き健康管理に努めてください。';
  }

  // 総合コメントを追加
  const overallComments = {
    'A': '\n\n【総合】\n今回の検査では特に問題は認められませんでした。',
    'B': '\n\n【総合】\n軽度の異常がありますが、生活習慣の見直しで改善が期待できます。',
    'C': '\n\n【総合】\n要経過観察の項目があります。定期的な検査をお勧めします。',
    'D': '\n\n【総合】\n要精密検査の項目があります。早めに医療機関を受診してください。',
    'E': '\n\n【総合】\n治療中の項目があります。引き続き治療を継続してください。',
    'F': '\n\n【総合】\n経過観察中の項目があります。定期的な検査を継続してください。'
  };

  return sections.join('\n\n') + (overallComments[overallJudgment] || '');
}

/**
 * 所見を取得（旧横持ち構造）
 * @deprecated 新構造は portalGetFindingsByKarteNo() を使用してください
 * @param {string} visitId - 受診ID
 * @returns {Object} 所見データ
 */
function getFindings(visitId) {
  try {
    if (!visitId) {
      return { success: false, error: '受診IDが必要です' };
    }

    const ss = getPortalSpreadsheet();
    const findingsSheet = ss.getSheetByName(DB_CONFIG.SHEETS.FINDINGS);
    if (!findingsSheet) {
      return { success: false, error: 'T_所見シートが見つかりません' };
    }

    const lastRow = findingsSheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: null, message: '所見データがありません' };
    }

    const data = findingsSheet.getRange(2, 1, lastRow - 1, 16).getValues();
    for (const row of data) {
      if (safeString(row[0]) === visitId) {
        return {
          success: true,
          data: {
            visitId: row[0],
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
          }
        };
      }
    }

    return { success: true, data: null, message: '所見データがありません' };

  } catch (error) {
    console.error('getFindings error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 所見を保存（旧横持ち構造・手動編集用）
 * @deprecated 新構造は portalSaveFindingsNew() を使用してください
 * @param {string} visitId - 受診ID
 * @param {Object} findingsData - 所見データ
 * @returns {Object} 保存結果
 */
function saveFindings(visitId, findingsData) {
  try {
    if (!visitId) {
      return { success: false, error: '受診IDが必要です' };
    }

    const ss = getPortalSpreadsheet();
    const findingsSheet = ss.getSheetByName(DB_CONFIG.SHEETS.FINDINGS);
    if (!findingsSheet) {
      return { success: false, error: 'T_所見シートが見つかりません' };
    }

    // 既存レコード検索
    const lastRow = findingsSheet.getLastRow();
    let existingRow = 0;
    if (lastRow >= 2) {
      const ids = findingsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (safeString(ids[i][0]) === visitId) {
          existingRow = i + 2;
          break;
        }
      }
    }

    const now = new Date();

    if (existingRow > 0) {
      // 既存レコードを更新
      if (findingsData.history !== undefined) findingsSheet.getRange(existingRow, 2).setValue(findingsData.history);
      if (findingsData.subjective !== undefined) findingsSheet.getRange(existingRow, 3).setValue(findingsData.subjective);
      if (findingsData.objective !== undefined) findingsSheet.getRange(existingRow, 4).setValue(findingsData.objective);
      if (findingsData.circulatory !== undefined) findingsSheet.getRange(existingRow, 5).setValue(findingsData.circulatory);
      if (findingsData.digestive !== undefined) findingsSheet.getRange(existingRow, 6).setValue(findingsData.digestive);
      if (findingsData.metabolicSugar !== undefined) findingsSheet.getRange(existingRow, 7).setValue(findingsData.metabolicSugar);
      if (findingsData.metabolicLipid !== undefined) findingsSheet.getRange(existingRow, 8).setValue(findingsData.metabolicLipid);
      if (findingsData.renal !== undefined) findingsSheet.getRange(existingRow, 9).setValue(findingsData.renal);
      if (findingsData.blood !== undefined) findingsSheet.getRange(existingRow, 10).setValue(findingsData.blood);
      if (findingsData.other !== undefined) findingsSheet.getRange(existingRow, 11).setValue(findingsData.other);
      if (findingsData.combined !== undefined) findingsSheet.getRange(existingRow, 12).setValue(findingsData.combined);
      if (findingsData.ecgCurrent !== undefined) findingsSheet.getRange(existingRow, 13).setValue(findingsData.ecgCurrent);
      if (findingsData.ecgPrevious !== undefined) findingsSheet.getRange(existingRow, 14).setValue(findingsData.ecgPrevious);
      if (findingsData.usCurrent !== undefined) findingsSheet.getRange(existingRow, 15).setValue(findingsData.usCurrent);
      if (findingsData.usPrevious !== undefined) findingsSheet.getRange(existingRow, 16).setValue(findingsData.usPrevious);
      findingsSheet.getRange(existingRow, 17).setValue(now);
    } else {
      // 新規レコードを追加
      const rowData = [
        visitId,
        findingsData.history || '',
        findingsData.subjective || '',
        findingsData.objective || '',
        findingsData.circulatory || '',
        findingsData.digestive || '',
        findingsData.metabolicSugar || '',
        findingsData.metabolicLipid || '',
        findingsData.renal || '',
        findingsData.blood || '',
        findingsData.other || '',
        findingsData.combined || '',
        findingsData.ecgCurrent || '',
        findingsData.ecgPrevious || '',
        findingsData.usCurrent || '',
        findingsData.usPrevious || '',
        now
      ];
      findingsSheet.appendRow(rowData);
    }

    return {
      success: true,
      message: existingRow > 0 ? '所見を更新しました' : '所見を保存しました'
    };

  } catch (error) {
    console.error('saveFindings error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 判定結果を保存（T_判定結果シート）
 * @param {string} visitId - 受診ID
 * @param {Object} judgmentData - 判定データ
 * @returns {Object} 保存結果
 */
function saveJudgmentResult(visitId, judgmentData) {
  try {
    if (!visitId) {
      return { success: false, error: '受診IDが必要です' };
    }

    const ss = getPortalSpreadsheet();
    const judgmentSheet = ss.getSheetByName(DB_CONFIG.SHEETS.JUDGMENT_RESULT);
    if (!judgmentSheet) {
      return { success: false, error: 'T_判定結果シートが見つかりません' };
    }

    // 既存レコード検索
    const lastRow = judgmentSheet.getLastRow();
    let existingRow = 0;
    if (lastRow >= 2) {
      const ids = judgmentSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (safeString(ids[i][0]) === visitId) {
          existingRow = i + 2;
          break;
        }
      }
    }

    const now = new Date();

    // Level1判定をJSON形式で保存
    const level1Json = JSON.stringify(judgmentData.level1 || {});
    // Level2判定をJSON形式で保存
    const level2Json = JSON.stringify(judgmentData.level2 || {});

    if (existingRow > 0) {
      // 既存レコードを更新
      judgmentSheet.getRange(existingRow, 2).setValue(level1Json);      // B: Lv1判定JSON
      judgmentSheet.getRange(existingRow, 3).setValue(level2Json);      // C: Lv2判定JSON
      judgmentSheet.getRange(existingRow, 4).setValue(judgmentData.level3 || '');  // D: Lv3総合判定
      judgmentSheet.getRange(existingRow, 5).setValue(now);             // E: 判定日時
    } else {
      // 新規レコードを追加
      const rowData = [
        visitId,
        level1Json,
        level2Json,
        judgmentData.level3 || '',
        now
      ];
      judgmentSheet.appendRow(rowData);
    }

    return {
      success: true,
      message: existingRow > 0 ? '判定結果を更新しました' : '判定結果を保存しました'
    };

  } catch (error) {
    console.error('saveJudgmentResult error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 判定と所見を一括で計算・保存
 * @param {string} visitId - 受診ID
 * @returns {Object} 処理結果
 */
function calculateAndSaveAllJudgments(visitId) {
  try {
    if (!visitId) {
      return { success: false, error: '受診IDが必要です' };
    }

    // 3レベル判定を取得
    const judgmentResult = getThreeLevelJudgment(visitId);
    if (!judgmentResult.success) {
      return { success: false, error: '判定計算失敗: ' + judgmentResult.error };
    }

    // 判定結果を保存
    const saveJudgment = saveJudgmentResult(visitId, judgmentResult);
    if (!saveJudgment.success) {
      return { success: false, error: '判定保存失敗: ' + saveJudgment.error };
    }

    // 所見を自動生成して保存
    const findingsResult = generateAndSaveFindings(visitId);
    if (!findingsResult.success) {
      return { success: false, error: '所見生成失敗: ' + findingsResult.error };
    }

    // 受診記録シートの総合判定を更新
    const ss = getPortalSpreadsheet();
    const visitSheet = ss.getSheetByName(DB_CONFIG.SHEETS.VISIT);
    if (visitSheet) {
      const visitData = visitSheet.getDataRange().getValues();
      for (let i = 1; i < visitData.length; i++) {
        if (safeString(visitData[i][VISIT_COL.VISIT_ID]) === visitId) {
          visitSheet.getRange(i + 1, VISIT_COL.JUDGMENT + 1).setValue(judgmentResult.level3);
          break;
        }
      }
    }

    return {
      success: true,
      message: '判定と所見を保存しました',
      judgment: judgmentResult,
      findings: findingsResult
    };

  } catch (error) {
    console.error('calculateAndSaveAllJudgments error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// データ移行ユーティリティ
// ============================================================

/**
 * カルテNo移行スクリプト
 * Q列（BML患者ID）の値をB列（カルテNo）にコピー
 *
 * 使用方法: GASエディタから直接実行
 * @returns {Object} 移行結果
 */
function migrateKarteNoFromBmlPatientId() {
  try {
    const ss = getPortalSpreadsheet();
    const sheet = ss.getSheetByName('受診者マスタ');

    if (!sheet) {
      console.log('ERROR: 受診者マスタシートが見つかりません');
      return { success: false, error: '受診者マスタシートが見つかりません' };
    }

    const lastRow = sheet.getLastRow();
    console.log('受診者マスタ: ' + lastRow + '行');

    let migratedCount = 0;
    let skippedCount = 0;
    const details = [];

    for (let i = 2; i <= lastRow; i++) {
      const patientId = sheet.getRange(i, 1).getValue(); // A列: 受診者ID
      const currentKarteNo = sheet.getRange(i, 2).getValue(); // B列: カルテNo
      const bmlPatientId = sheet.getRange(i, 17).getValue(); // Q列: BML患者ID

      // B列が空でQ列に値がある場合のみコピー
      if (bmlPatientId && !currentKarteNo) {
        sheet.getRange(i, 2).setValue(String(bmlPatientId));
        migratedCount++;
        details.push({ row: i, patientId: patientId, karteNo: bmlPatientId, action: 'migrated' });
        console.log('Row ' + i + ': ' + patientId + ' → カルテNo=' + bmlPatientId);
      } else if (currentKarteNo) {
        skippedCount++;
        details.push({ row: i, patientId: patientId, karteNo: currentKarteNo, action: 'skipped (already set)' });
      } else {
        skippedCount++;
        details.push({ row: i, patientId: patientId, action: 'skipped (no BML ID)' });
      }
    }

    console.log('=== 移行完了 ===');
    console.log('移行件数: ' + migratedCount);
    console.log('スキップ件数: ' + skippedCount);

    return {
      success: true,
      message: '移行完了',
      migratedCount: migratedCount,
      skippedCount: skippedCount,
      details: details
    };

  } catch (error) {
    console.error('migrateKarteNoFromBmlPatientId error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================
// 所見API（縦持ち・検査項目別構造）- 新規追加
// ============================================================

/**
 * カルテNoで所見を取得（縦持ち構造）
 * @param {string} karteNo - カルテNo（6桁）
 * @returns {Object} 取得結果
 */
function portalGetFindingsByKarteNo(karteNo) {
  try {
    if (!karteNo) {
      return { success: false, error: 'カルテNoが必要です' };
    }

    const findings = getFindingsByKarteNo(karteNo);
    return {
      success: true,
      data: findings,
      count: findings.length
    };

  } catch (error) {
    console.error('portalGetFindingsByKarteNo error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 所見を保存（縦持ち構造・upsert）
 * @param {string} karteNo - カルテNo
 * @param {string} patientId - 受診者ID
 * @param {Array<Object>} findingsData - 所見データ配列
 * @returns {Object} 保存結果
 */
function portalSaveFindingsNew(karteNo, patientId, findingsData) {
  try {
    if (!karteNo) {
      return { success: false, error: 'カルテNoが必要です' };
    }
    if (!patientId) {
      return { success: false, error: '受診者IDが必要です' };
    }
    if (!findingsData || !Array.isArray(findingsData)) {
      return { success: false, error: '所見データが必要です' };
    }

    // patientIdを各所見に追加
    const dataWithPatient = findingsData.map(f => ({
      ...f,
      patientId: patientId,
      karteNo: karteNo
    }));

    const result = upsertFindings(karteNo, dataWithPatient);

    return {
      success: true,
      message: `所見を保存しました（作成: ${result.created}, 更新: ${result.updated}）`,
      created: result.created,
      updated: result.updated
    };

  } catch (error) {
    console.error('portalSaveFindingsNew error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 所見テンプレート一覧を取得
 * @param {string} [itemId] - 項目ID（省略時は全件）
 * @param {string} [judgment] - 判定（省略時は全件）
 * @returns {Object} テンプレート一覧
 */
function portalGetFindingTemplates(itemId, judgment) {
  try {
    const templates = getFindingTemplates(itemId, judgment);
    return {
      success: true,
      data: templates,
      count: templates.length
    };

  } catch (error) {
    console.error('portalGetFindingTemplates error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 単一所見を更新
 * @param {string} findingId - 所見ID
 * @param {Object} data - 更新データ
 * @returns {Object} 更新結果
 */
function portalUpdateFinding(findingId, data) {
  try {
    if (!findingId) {
      return { success: false, error: '所見IDが必要です' };
    }

    updateFinding(findingId, data);
    return { success: true, message: '所見を更新しました' };

  } catch (error) {
    console.error('portalUpdateFinding error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 所見を削除
 * @param {string} findingId - 所見ID
 * @returns {Object} 削除結果
 */
function portalDeleteFinding(findingId) {
  try {
    if (!findingId) {
      return { success: false, error: '所見IDが必要です' };
    }

    deleteFinding(findingId);
    return { success: true, message: '所見を削除しました' };

  } catch (error) {
    console.error('portalDeleteFinding error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 所見入力用の検査項目一覧を取得（2ページ目用）
 * @returns {Object} 所見入力対象項目一覧
 */
function portalGetFindingItems() {
  try {
    // 所見入力対象カテゴリ
    const findingCategories = ['画像診断', '心電図', '内視鏡'];
    const findingSubcategories = ['MRI', 'CT', '超音波検査', 'X線'];

    const allItems = EXAM_ITEM_MASTER_DATA;
    const findingItems = allItems.filter(item => {
      return findingCategories.includes(item.category) ||
             findingSubcategories.includes(item.subcategory);
    });

    // カテゴリごとにグループ化
    const grouped = {};
    for (const item of findingItems) {
      const key = item.subcategory || item.category;
      if (!grouped[key]) {
        grouped[key] = [];
      }
      grouped[key].push({
        itemId: item.item_id,
        itemName: item.item_name,
        category: item.category,
        subcategory: item.subcategory
      });
    }

    return {
      success: true,
      data: grouped,
      totalCount: findingItems.length
    };

  } catch (error) {
    console.error('portalGetFindingItems error:', error);
    return { success: false, error: error.message };
  }
}