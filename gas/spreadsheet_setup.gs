/**
 * スプレッドシート初期設定モジュール
 * シート作成とヘッダー設定
 */

// ============================================
// シート定義
// ============================================

const SHEET_DEFINITIONS = {
  // 受診者マスタ
  '受診者マスタ': {
    headers: [
      '受診ID', 'ステータス', '受診日', '氏名', 'カナ', '性別',
      '生年月日', '年齢', '受診コース', '事業所名', '所属',
      '総合判定', 'CSV取込日時', '最終更新日時', '出力日時'
    ],
    columnWidths: {
      A: 150, B: 80, C: 100, D: 100, E: 120, F: 50,
      G: 100, H: 50, I: 100, J: 120, K: 80,
      L: 70, M: 150, N: 150, O: 150
    }
  },

  // 身体測定
  '身体測定': {
    headers: [
      '受診ID', '身長', '体重', '標準体重', 'BMI', '体脂肪率', '腹囲',
      '血圧_収縮期_1', '血圧_拡張期_1', '血圧_収縮期_2', '血圧_拡張期_2',
      '視力_裸眼_右', '視力_裸眼_左', '視力_矯正_右', '視力_矯正_左',
      '聴力_右_1000Hz', '聴力_左_1000Hz', '聴力_右_4000Hz', '聴力_左_4000Hz',
      '眼圧_右', '眼圧_左', '眼底_右', '眼底_左'
    ],
    columnWidths: {
      A: 150, B: 60, C: 60, D: 70, E: 60, F: 70, G: 60,
      H: 90, I: 90, J: 90, K: 90,
      L: 90, M: 90, N: 90, O: 90,
      P: 100, Q: 100, R: 100, S: 100,
      T: 60, U: 60, V: 80, W: 80
    }
  },

  // 血液検査
  '血液検査': {
    headers: [
      '受診ID', 'WBC', 'RBC', 'Hb', 'Ht', 'PLT', 'MCV', 'MCH', 'MCHC',
      'TP', 'ALB', 'T-Bil', 'AST', 'ALT', 'γ-GTP', 'ALP', 'LDH',
      'BUN', 'Cr', 'eGFR', 'UA',
      'TC', 'HDL-C', 'LDL-C', 'TG',
      'FBS', 'HbA1c', 'CRP'
    ],
    columnWidths: {
      A: 150, B: 60, C: 60, D: 60, E: 60, F: 60, G: 60, H: 60, I: 60,
      J: 60, K: 60, L: 60, M: 60, N: 60, O: 60, P: 60, Q: 60,
      R: 60, S: 60, T: 60, U: 60,
      V: 60, W: 60, X: 60, Y: 60,
      Z: 60, AA: 60, AB: 60
    }
  },

  // 所見
  '所見': {
    headers: [
      '受診ID', '既往歴', '自覚症状', '他覚症状',
      '所見_循環器系', '所見_消化器系', '所見_代謝系_糖', '所見_代謝系_脂質',
      '所見_腎機能', '所見_血液系', '所見_その他', '総合所見',
      '心電図_今回', '心電図_前回', '腹部超音波_今回', '腹部超音波_前回'
    ],
    columnWidths: {
      A: 150, B: 150, C: 150, D: 150,
      E: 200, F: 200, G: 200, H: 200,
      I: 200, J: 200, K: 200, L: 300,
      M: 150, N: 150, O: 150, P: 150
    }
  },

  // 判定マスタ
  '判定マスタ': {
    headers: [
      '項目ID', '項目名', 'カテゴリ', '性別依存',
      '基準値下限', '基準値上限',
      'A下限', 'A上限', 'B下限', 'B上限', 'C下限', 'C上限', 'D下限', 'D上限'
    ],
    columnWidths: {
      A: 120, B: 100, C: 100, D: 70,
      E: 80, F: 80,
      G: 60, H: 60, I: 60, J: 60, K: 60, L: 60, M: 60, N: 60
    },
    sampleData: [
      ['AST_GOT', 'AST(GOT)', '消化器系', 'なし', 10, 30, 0, 30, 31, 35, 36, 50, 51, null],
      ['ALT_GPT', 'ALT(GPT)', '消化器系', 'なし', 5, 30, 0, 30, 31, 40, 41, 50, 51, null],
      ['GAMMA_GTP', 'γ-GTP', '消化器系', 'なし', 10, 50, 0, 50, 51, 80, 81, 100, 101, null],
      ['HDL_CHOLESTEROL', 'HDL-C', '代謝系（脂質）', 'なし', 40, 100, 40, 100, null, null, 30, 39, null, 29],
      ['LDL_CHOLESTEROL', 'LDL-C', '代謝系（脂質）', 'なし', 60, 119, 60, 119, 120, 139, 140, 179, 180, null],
      ['TRIGLYCERIDES', 'TG', '代謝系（脂質）', 'なし', 30, 149, 30, 149, 150, 299, 300, 499, 500, null],
      ['FASTING_GLUCOSE', 'FBS', '代謝系（糖）', 'なし', 70, 99, null, 99, 100, 109, 110, 125, 126, null],
      ['HBA1C', 'HbA1c', '代謝系（糖）', 'なし', 4.6, 5.5, null, 5.5, 5.6, 5.9, 6.0, 6.4, 6.5, null],
      ['EGFR', 'eGFR', '腎機能', 'なし', 60, null, 60, null, 45, 59.9, 30, 44.9, null, 29.9],
      ['URIC_ACID', 'UA', 'その他', 'なし', 2.1, 7.0, 2.1, 7.0, 7.1, 8.0, 8.1, 9.0, 9.1, null],
      ['CRP', 'CRP', 'その他', 'なし', 0, 0.3, 0, 0.3, 0.31, 0.99, 1.0, 1.99, 2.0, null],
      ['HEMOGLOBIN_M', 'Hb(男)', '血液系', 'あり', 13.1, 16.3, 13.1, 16.3, 12.1, 13.0, 11.0, 12.0, null, 10.9],
      ['HEMOGLOBIN_F', 'Hb(女)', '血液系', 'あり', 12.1, 14.5, 12.1, 14.5, 11.1, 12.0, 10.0, 11.0, null, 9.9],
      ['CREATININE_M', 'Cr(男)', '腎機能', 'あり', 0.6, 1.0, 0.1, 1.0, 1.01, 1.2, 1.21, 1.5, 1.51, null],
      ['CREATININE_F', 'Cr(女)', '腎機能', 'あり', 0.4, 0.7, 0.1, 0.7, 0.71, 0.9, 0.91, 1.1, 1.11, null]
    ]
  },

  // 所見テンプレート
  '所見テンプレート': {
    headers: ['カテゴリ', '判定', '項目パターン', 'コメント'],
    columnWidths: { A: 120, B: 60, C: 100, D: 400 },
    sampleData: [
      ['循環器系', 'B', '血圧', '血圧がやや高めです。減塩と適度な運動を心がけてください。'],
      ['循環器系', 'C', '血圧', '血圧が高めです。生活習慣の見直しと定期的な測定をお勧めします。'],
      ['循環器系', 'D', '血圧', '血圧が基準値を大きく超えています。医療機関での受診をお勧めします。'],
      ['消化器系', 'B', '肝機能', '肝機能に軽度の異常があります。飲酒量の見直しをお勧めします。'],
      ['消化器系', 'C', '肝機能', '肝機能に異常があります。精密検査をお勧めします。'],
      ['消化器系', 'D', '肝機能', '肝機能に明らかな異常があります。早めの医療機関受診をお勧めします。'],
      ['代謝系（糖）', 'B', '血糖', '血糖値がやや高めです。糖質の摂取を控えめにしてください。'],
      ['代謝系（糖）', 'C', '血糖', '血糖値が高めです。糖尿病の精密検査をお勧めします。'],
      ['代謝系（糖）', 'D', '血糖', '血糖値が基準値を大きく超えています。糖尿病の治療が必要です。'],
      ['代謝系（脂質）', 'B', '脂質', '脂質代謝に軽度の異常があります。食事内容の見直しをお勧めします。'],
      ['代謝系（脂質）', 'C', '脂質', '脂質代謝に異常があります。動物性脂肪を控え、運動習慣を取り入れてください。'],
      ['代謝系（脂質）', 'D', '脂質', '脂質代謝に明らかな異常があります。医療機関での治療をお勧めします。'],
      ['腎機能', 'B', '腎', '腎機能に軽度の異常があります。水分を十分に摂取してください。'],
      ['腎機能', 'C', '腎', '腎機能に異常があります。定期的な経過観察をお勧めします。'],
      ['腎機能', 'D', '腎', '腎機能に明らかな異常があります。専門医の受診をお勧めします。'],
      ['血液系', 'B', '血液', '血液検査に軽度の異常があります。経過観察をお勧めします。'],
      ['血液系', 'C', '血液', '血液検査に異常があります。精密検査をお勧めします。'],
      ['血液系', 'D', '血液', '血液検査に明らかな異常があります。早めの受診をお勧めします。'],
      ['その他', 'B', 'UA', '尿酸値がやや高めです。プリン体を含む食品を控えてください。'],
      ['その他', 'C', 'BMI', '肥満傾向があります。食事と運動による体重管理をお勧めします。']
    ]
  },

  // 設定
  '設定': {
    headers: ['設定キー', '設定値', '説明'],
    columnWidths: { A: 150, B: 300, C: 300 },
    sampleData: [
      ['CSV_FOLDER_ID', 'YOUR_CSV_FOLDER_ID', 'CSVファイル配置フォルダのID'],
      ['OUTPUT_FOLDER_ID', 'YOUR_OUTPUT_FOLDER_ID', 'Excel出力フォルダのID'],
      ['NOTIFICATION_EMAIL', '', '通知先メールアドレス（空欄で現在のユーザー）']
    ]
  }
};

// ============================================
// セットアップ関数
// ============================================

/**
 * すべてのシートを作成
 */
function createAllSheets() {
  logInfo('===== シート作成開始 =====');

  const ss = getSpreadsheet();

  for (const [sheetName, definition] of Object.entries(SHEET_DEFINITIONS)) {
    createSheet(ss, sheetName, definition);
  }

  logInfo('===== シート作成完了 =====');
}

/**
 * シートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @param {Object} definition - シート定義
 */
function createSheet(ss, sheetName, definition) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo(`シート作成: ${sheetName}`);
  } else {
    logInfo(`シート既存: ${sheetName}`);
  }

  // ヘッダー設定
  const headers = definition.headers;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // 列幅設定
  if (definition.columnWidths) {
    for (const [col, width] of Object.entries(definition.columnWidths)) {
      const colIndex = columnLetterToIndex(col);
      sheet.setColumnWidth(colIndex, width);
    }
  }

  // サンプルデータ投入
  if (definition.sampleData && definition.sampleData.length > 0) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      sheet.getRange(2, 1, definition.sampleData.length, definition.sampleData[0].length)
        .setValues(definition.sampleData);
      logInfo(`  サンプルデータ投入: ${definition.sampleData.length}行`);
    }
  }

  // 1行目を固定
  sheet.setFrozenRows(1);
}

/**
 * 列文字をインデックスに変換
 * @param {string} letter - 列文字（A, B, ..., AA, AB, ...）
 * @returns {number} 列インデックス（1始まり）
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

/**
 * 出力用テンプレートシートを作成
 * doc_template_idheart.xlsxのレイアウトを自動設定
 */
function createOutputTemplateSheets() {
  logInfo('===== 出力用テンプレートシート作成 =====');

  const ss = getSpreadsheet();

  // メインの出力用シートを作成
  let sheet = ss.getSheetByName('出力用_1ページ');
  if (!sheet) {
    sheet = ss.insertSheet('出力用_1ページ');
    logInfo('シート作成: 出力用_1ページ');
  }

  // テンプレートレイアウトを設定
  setupOutputTemplateLayout(sheet);

  logInfo('出力用テンプレートシート作成完了');
}

/**
 * 出力用シートのレイアウトを設定（doc_template_idheart.xlsx準拠）
 * @param {Sheet} sheet - シート
 */
function setupOutputTemplateLayout(sheet) {
  // 既存データをクリア
  sheet.clear();

  // ============================================
  // 列幅設定（テンプレートから取得した値）
  // ============================================
  const columnWidths = {
    1: 25,   // A
    2: 60,   // B
    3: 45,   // C
    4: 15,   // D
    5: 105,  // E
    6: 30,   // F
    7: 40,   // G
    8: 25,   // H
    9: 20,   // I
    10: 60,  // J
    11: 20,  // K
    12: 87,  // L
    13: 20,  // M
    14: 5,   // N
    15: 25,  // O
    16: 35,  // P
    17: 60,  // Q
    18: 25,  // R
    19: 28,  // S
    20: 72,  // T
    21: 25,  // U
    22: 12,  // V
    23: 105, // W
    24: 64,  // X
    25: 20,  // Y
    26: 107, // Z
    27: 5,   // AA
    28: 25,  // AB
    29: 40,  // AC
    30: 82,  // AD
    31: 28,  // AE
    32: 90,  // AF
    33: 25,  // AG
    34: 120  // AH
  };

  for (const [col, width] of Object.entries(columnWidths)) {
    sheet.setColumnWidth(parseInt(col), width);
  }

  // ============================================
  // 行高設定
  // ============================================
  const rowHeights = {
    1: 5, 2: 15, 3: 5, 4: 15, 5: 15, 6: 15, 7: 5, 8: 15, 9: 15, 10: 15,
    11: 5, 12: 12, 13: 12, 14: 14, 15: 14, 16: 14, 17: 14, 18: 6, 19: 7,
    20: 14, 21: 14, 22: 14, 23: 14, 24: 14, 25: 14, 26: 14, 27: 14, 28: 7,
    29: 6, 30: 14, 31: 14, 32: 14, 33: 14, 34: 14, 35: 14, 36: 14, 37: 14,
    38: 6, 39: 7, 40: 14, 41: 14, 42: 14, 43: 18, 44: 14, 45: 13, 46: 13,
    47: 13, 48: 6, 49: 8, 50: 8, 51: 13, 52: 18, 53: 13, 54: 13, 55: 13,
    56: 13, 57: 13, 58: 13, 59: 13, 60: 13, 61: 13, 62: 5, 63: 13, 64: 13,
    65: 13, 66: 13, 67: 13
  };

  for (const [row, height] of Object.entries(rowHeights)) {
    sheet.setRowHeight(parseInt(row), height);
  }

  // ============================================
  // ヘッダーと固定ラベル
  // ============================================
  const labels = [
    // タイトル
    {cell: 'O1', value: '健　診　結　果　報　告　書', fontWeight: 'bold', fontSize: 14},

    // 受診者情報セクション
    {cell: 'A2', value: '受診者ｶﾅ'},
    {cell: 'I2', value: '受診ｺｰﾄﾞ'},
    {cell: 'AB2', value: '医療機関名'},
    {cell: 'A3', value: '受診者名'},
    {cell: 'I3', value: '個人ｺｰﾄﾞ'},

    // 判定基準
    {cell: 'O4', value: '判定基準', fontWeight: 'bold'},
    {cell: 'P4', value: 'A', fontWeight: 'bold'},
    {cell: 'Q4', value: '異常なし'},
    {cell: 'U4', value: 'D1', fontWeight: 'bold'},
    {cell: 'X4', value: '要治療'},
    {cell: 'P5', value: 'B', fontWeight: 'bold'},
    {cell: 'Q5', value: '軽度異常あるも日常生活に支障なし'},
    {cell: 'U5', value: 'D2', fontWeight: 'bold'},
    {cell: 'X5', value: '要精密検査'},
    {cell: 'P6', value: 'C', fontWeight: 'bold'},
    {cell: 'Q6', value: '軽度異常あり生活習慣改善を要す'},
    {cell: 'U6', value: 'E', fontWeight: 'bold'},
    {cell: 'X6', value: '現在治療中'},

    // 受診者情報続き
    {cell: 'A5', value: '生年月日'},
    {cell: 'I5', value: '年齢'},
    {cell: 'M5', value: '歳'},
    {cell: 'AB5', value: '〒'},
    {cell: 'A6', value: '保険証記号'},
    {cell: 'I6', value: '保険証番号'},
    {cell: 'A7', value: '団体名'},
    {cell: 'I7', value: '支店名'},
    {cell: 'O8', value: '既往歴'},
    {cell: 'A9', value: '受診日'},
    {cell: 'C9', value: '今　回'},
    {cell: 'J9', value: '前　回'},
    {cell: 'O9', value: '自覚症状'},
    {cell: 'AB9', value: '電話番号：'},
    {cell: 'AF9', value: 'FAX番号：'},
    {cell: 'A10', value: 'ｺｰｽ名'},
    {cell: 'O10', value: '他覚症状'},

    // 検査項目ヘッダー
    {cell: 'A12', value: '検　査　項　目', fontWeight: 'bold'},
    {cell: 'D12', value: '判　定', fontWeight: 'bold'},
    {cell: 'L12', value: '基　準　値', fontWeight: 'bold'},
    {cell: 'O12', value: '検　査　項　目', fontWeight: 'bold'},
    {cell: 'S12', value: '判　定', fontWeight: 'bold'},
    {cell: 'Z12', value: '基　準　値', fontWeight: 'bold'},
    {cell: 'AC12', value: '画　像', fontWeight: 'bold'},
    {cell: 'AE12', value: '判定', fontWeight: 'bold'},
    {cell: 'AF12', value: '検　査　所　見', fontWeight: 'bold'},
    {cell: 'D13', value: '今　回'},
    {cell: 'I13', value: '前　回'},
    {cell: 'S13', value: '今回'},
    {cell: 'V13', value: '前回'},
    {cell: 'AF13', value: '今回'},
    {cell: 'AH13', value: '前回'},

    // 診察・身体測定
    {cell: 'A14', value: '診察・身体測定', fontWeight: 'bold', background: '#d9ead3'},
    {cell: 'B14', value: '身長'},
    {cell: 'B15', value: '体重'},
    {cell: 'B16', value: '腹囲'},
    {cell: 'L16', value: '84.9cm以下'},
    {cell: 'B17', value: '標準体重'},
    {cell: 'B18', value: 'ＢＭＩ'},
    {cell: 'L18', value: '18.5～24.9'},
    {cell: 'B20', value: '診察所見'},

    // 糖代謝
    {cell: 'O14', value: '糖代謝', fontWeight: 'bold', background: '#fff2cc'},
    {cell: 'P14', value: '空腹時血糖'},
    {cell: 'P15', value: 'ＨｂＡ１ｃ(ＪＤＳ)'},
    {cell: 'P16', value: 'ＨｂＡ１ｃ(ＮＧＳＰ)'},
    {cell: 'P17', value: '尿糖(定性)'},
    {cell: 'Z17', value: '(-)'},

    // 脂質代謝
    {cell: 'O18', value: '脂質代謝', fontWeight: 'bold', background: '#fce5cd'},
    {cell: 'P18', value: '総コレステロ－ル'},
    {cell: 'Z18', value: '140～199mg/dl'},
    {cell: 'P20', value: 'ＨＤＬコレステロール'},
    {cell: 'Z20', value: '40～119mg/dl'},
    {cell: 'P21', value: 'ＬＤＬコレステロール'},
    {cell: 'Z21', value: '60～119mg/dl'},
    {cell: 'P22', value: '中性脂肪'},
    {cell: 'Z22', value: '30～149mg/dl'},

    // 眼科
    {cell: 'A22', value: '眼科', fontWeight: 'bold', background: '#c9daf8'},
    {cell: 'B22', value: '視力　裸眼　右'},
    {cell: 'L22', value: '1.0以上'},
    {cell: 'B23', value: '視力　裸眼　左'},
    {cell: 'L23', value: '1.0以上'},
    {cell: 'B24', value: '視力　矯正　右'},
    {cell: 'L24', value: '1.0以上'},
    {cell: 'B25', value: '視力　矯正　左'},
    {cell: 'L25', value: '1.0以上'},

    // 血液一般
    {cell: 'O23', value: '血液一般', fontWeight: 'bold', background: '#f4cccc'},
    {cell: 'P23', value: '白血球数'},
    {cell: 'Z23', value: '3.2～8.5千/μl'},
    {cell: 'P24', value: '赤血球数'},
    {cell: 'Z24', value: '400～539万/μl'},
    {cell: 'P25', value: 'ヘモグロビン'},
    {cell: 'Z25', value: '13.1～16.6g/dl'},
    {cell: 'P26', value: 'ヘマトクリット'},
    {cell: 'Z26', value: '38.5～48.9％'},
    {cell: 'P27', value: 'ＭＣＶ'},
    {cell: 'P28', value: 'ＭＣＨ'},
    {cell: 'P30', value: 'ＭＣＨＣ'},
    {cell: 'P31', value: '血小板数'},
    {cell: 'Z31', value: '13.0～34.9万/μl'},
    {cell: 'P32', value: '好塩基球(BASO)'},
    {cell: 'P33', value: '好酸球(EOS)'},
    {cell: 'P34', value: '好中球(NEUT)'},
    {cell: 'P35', value: 'リンパ球(LYMPO)'},
    {cell: 'P36', value: '単球(MONO)'},

    // 眼底
    {cell: 'A26', value: '眼底', fontWeight: 'bold', background: '#c9daf8'},
    {cell: 'B26', value: '眼底ＫＷ'},
    {cell: 'B27', value: '眼底Ｈ'},
    {cell: 'L27', value: '0'},
    {cell: 'B28', value: '眼底Ｓ'},
    {cell: 'L28', value: '0'},
    {cell: 'B30', value: '眼底所見'},

    // 聴力
    {cell: 'A32', value: '聴力', fontWeight: 'bold', background: '#c9daf8'},
    {cell: 'B32', value: '聴力1000Hz　右'},
    {cell: 'L32', value: '30dB以下'},
    {cell: 'B33', value: '聴力1000Hz　左'},
    {cell: 'L33', value: '30dB以下'},
    {cell: 'B34', value: '聴力4000Hz　右'},
    {cell: 'L34', value: '30dB以下'},
    {cell: 'B35', value: '聴力4000Hz　左'},
    {cell: 'L35', value: '30dB以下'},

    // 血圧
    {cell: 'A36', value: '血圧', fontWeight: 'bold', background: '#d9d2e9'},
    {cell: 'B36', value: '血圧１回目(最高)'},
    {cell: 'L36', value: '129mmHg以下'},
    {cell: 'B37', value: '血圧１回目(最低)'},
    {cell: 'L37', value: '84mmHg以下'},
    {cell: 'B38', value: '血圧２回目(最高)'},
    {cell: 'L38', value: '129mmHg以下'},
    {cell: 'B40', value: '血圧２回目(最低)'},
    {cell: 'L40', value: '84mmHg以下'},

    // 肺機能
    {cell: 'O37', value: '肺機能', fontWeight: 'bold', background: '#cfe2f3'},
    {cell: 'P37', value: '努力性肺活量'},
    {cell: 'P38', value: '１秒量'},
    {cell: 'P40', value: '１秒率'},

    // 肝機能
    {cell: 'A41', value: '肝機能', fontWeight: 'bold', background: '#d9ead3'},
    {cell: 'B41', value: 'ＡＳＴ(ＧＯＴ)'},
    {cell: 'L41', value: '30IU/L以下'},
    {cell: 'B42', value: 'ＡＬＴ(ＧＰＴ)'},
    {cell: 'L42', value: '30IU/L以下'},
    {cell: 'B43', value: 'γ-ＧＴＰ'},
    {cell: 'L43', value: '50IU/L以下'},
    {cell: 'B44', value: 'ＡＬＰ'},
    {cell: 'B45', value: '総ビリルビン'},
    {cell: 'B46', value: '総蛋白'},
    {cell: 'L46', value: '6.5～8.0g/dl'},
    {cell: 'B47', value: 'アルブミン'},
    {cell: 'L47', value: '4.0g/dl以上'},

    // 大腸
    {cell: 'O41', value: '大腸', fontWeight: 'bold', background: '#f4cccc'},
    {cell: 'P41', value: '便潜血１回目'},
    {cell: 'Z41', value: '(-)'},
    {cell: 'P42', value: '便潜血２回目'},
    {cell: 'Z42', value: '(-)'},

    // 子宮
    {cell: 'O43', value: '子宮', fontWeight: 'bold', background: '#ead1dc'},
    {cell: 'P43', value: '子宮頸部細胞診(ベセスダ分類)'},

    // 感染症
    {cell: 'O44', value: '感染症', fontWeight: 'bold', background: '#b6d7a8'},
    {cell: 'P44', value: 'ＨＢｓ抗原'},
    {cell: 'Z44', value: '(-)'},
    {cell: 'P45', value: 'ＨＣＶ抗体'},
    {cell: 'Z45', value: '(-)'},
    {cell: 'P46', value: 'ＨＣＶ-ＲＮＡ'},

    // 膵機能
    {cell: 'A48', value: '膵機能', fontWeight: 'bold', background: '#d9ead3'},
    {cell: 'B48', value: '血清アミラーゼ'},

    // 総合判定
    {cell: 'O48', value: '総合判定', fontWeight: 'bold'},
    {cell: 'V48', value: 'メタボリック\nシンドローム判定'},

    // 尿酸
    {cell: 'A52', value: '尿酸', fontWeight: 'bold', background: '#fff2cc'},
    {cell: 'B52', value: '尿酸'},
    {cell: 'L52', value: '2.1～7.0mg/dl以下'},

    // 総合所見
    {cell: 'O52', value: '■　　総　合　所　見　　■', fontWeight: 'bold'},

    // 尿一般・腎
    {cell: 'A53', value: '尿一般・腎', fontWeight: 'bold', background: '#cfe2f3'},
    {cell: 'B53', value: 'クレアチニン'},
    {cell: 'L53', value: '1.00mg/dl以下'},
    {cell: 'B54', value: '尿素窒素'},
    {cell: 'B55', value: '尿蛋白'},
    {cell: 'L55', value: '(-)'},
    {cell: 'B56', value: '尿潜血'},
    {cell: 'L56', value: '(-)'},
    {cell: 'B57', value: '沈渣　赤血球'},
    {cell: 'B58', value: '沈渣　白血球'},
    {cell: 'B59', value: '沈渣　扁平上皮'},
    {cell: 'B60', value: '沈渣　その他'},

    // 画像検査
    {cell: 'AB14', value: '心電図', fontWeight: 'bold'},
    {cell: 'AB19', value: '胸部X線', fontWeight: 'bold'},
    {cell: 'AB24', value: '胃部X線', fontWeight: 'bold'},
    {cell: 'AB29', value: '胃部内視鏡', fontWeight: 'bold'},
    {cell: 'AB34', value: '腹部超音波', fontWeight: 'bold'},
    {cell: 'AB39', value: '乳房触診', fontWeight: 'bold'},
    {cell: 'AB44', value: 'マンモ\nグラフィ', fontWeight: 'bold'},

    // フッター
    {cell: 'AD61', value: '医師名：'},
    {cell: 'AF63', value: '帳票番号：'},
    {cell: 'AG63', value: 'IH4-2-(2)'}
  ];

  // ラベルを設定
  for (const label of labels) {
    const range = sheet.getRange(label.cell);
    range.setValue(label.value);

    if (label.fontWeight) {
      range.setFontWeight(label.fontWeight);
    }
    if (label.fontSize) {
      range.setFontSize(label.fontSize);
    }
    if (label.background) {
      range.setBackground(label.background);
    }
  }

  // データ入力用セルの参照設定（例）
  setupDataReferences(sheet);

  logInfo('出力用シートのレイアウト設定完了');
}

/**
 * データ参照セルを設定
 * @param {Sheet} sheet - シート
 */
function setupDataReferences(sheet) {
  // 入力データを反映させるセルの定義
  // 実際のデータは別シートから参照するか、直接書き込む

  // データ入力用セルに枠線を設定
  const dataInputCells = [
    // 受診者情報
    'C2:H2', 'K2:N2', 'AD2:AH2',  // カナ、受診コード、医療機関名
    'C3:G3', 'K3:N3',              // 氏名、個人コード
    'C5:F5', 'G5', 'K5:L5',        // 生年月日、性別、年齢
    'C6:F6', 'K6:N6',              // 保険証
    'C7:H7', 'K7:N7',              // 団体名、支店名
    'R8:AA8',                       // 既往歴
    'E9:H9', 'K9:N9',              // 受診日
    'R9:AA9',                       // 自覚症状
    'E10:H10', 'K10:N10',          // コース名
    'R10:AA10',                     // 他覚症状
    'AD5:AH5', 'AD9:AE9', 'AG9:AH9', // 住所、電話

    // 検査値セル（代表的なもの）
    'E14:H14', 'J14:K14',          // 身長
    'E15:H15', 'J15:K15',          // 体重
    'E16:H16', 'J16:K16',          // 腹囲
    'E17:H17',                      // 標準体重
    'E18:H18', 'J18:K18',          // BMI
  ];

  // 枠線のスタイル
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID;

  for (const cellRange of dataInputCells) {
    try {
      sheet.getRange(cellRange).setBorder(true, true, true, true, false, false, '#999999', borderStyle);
    } catch (e) {
      // 無効な範囲はスキップ
    }
  }
}

/**
 * テンプレートレイアウトを再設定（メニューから実行可能）
 */
function resetOutputTemplateLayout() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'レイアウトリセット',
    '出力用_1ページのレイアウトをリセットします。既存のデータは消去されます。続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    const sheet = getSheet(CONFIG.SHEETS.OUTPUT_PAGE1);
    if (sheet) {
      setupOutputTemplateLayout(sheet);
      ui.alert('完了', 'レイアウトをリセットしました。', ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', '出力用_1ページシートが見つかりません。', ui.ButtonSet.OK);
    }
  }
}

/**
 * フォルダIDを設定
 * @param {string} csvFolderId - CSVフォルダID
 * @param {string} outputFolderId - 出力フォルダID
 */
function setFolderIds(csvFolderId, outputFolderId) {
  const sheet = getSheet(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    throw new Error('設定シートが見つかりません');
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'CSV_FOLDER_ID') {
      sheet.getRange(i + 1, 2).setValue(csvFolderId);
    }
    if (data[i][0] === 'OUTPUT_FOLDER_ID') {
      sheet.getRange(i + 1, 2).setValue(outputFolderId);
    }
  }

  // CONFIGを更新
  CONFIG.CSV_FOLDER_ID = csvFolderId;
  CONFIG.OUTPUT_FOLDER_ID = outputFolderId;

  logInfo('フォルダIDを設定しました');
}

/**
 * フォルダIDを取得（URLから）
 * @param {string} folderUrl - フォルダURL
 * @returns {string} フォルダID
 */
function getFolderIdFromUrl(folderUrl) {
  const match = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  throw new Error('フォルダURLからIDを取得できません');
}

/**
 * 完全セットアップを実行
 */
function runFullSetup() {
  logInfo('===== 完全セットアップ開始 =====');

  // 1. データシートを作成
  createAllSheets();

  // 2. 出力用テンプレートシートを作成
  createOutputTemplateSheets();

  // 3. トリガーを設定
  setupTriggers();

  logInfo('===== 完全セットアップ完了 =====');
  logInfo('');
  logInfo('【次のステップ】');
  logInfo('1. 設定シートにフォルダIDを入力してください');
  logInfo('2. template.xlsmのレイアウトを出力用シートにコピーしてください');
  logInfo('3. テストCSVで動作確認してください');
}
