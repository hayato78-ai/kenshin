/**
 * AppSheet テーブル構造セットアップスクリプト
 *
 * 使用方法:
 * 1. GASエディタで setupAppSheetTables() を実行
 * 2. 必要なシートが自動作成される
 * 3. AppSheetでスプレッドシートを接続
 */

// ============================================
// テーブル定義
// ============================================

const APPSHEET_TABLES = {
  // 案件マスタ
  cases: {
    sheetName: 'AS_案件',
    headers: [
      'case_id',           // 案件ID (キー)
      'case_name',         // 案件名
      'exam_type',         // 検診種別: DOCK | ROSAI_SECONDARY | REGULAR
      'client_name',       // 顧客名/事業所名
      'client_address',    // 事業所住所
      'client_contact',    // 担当者名
      'client_phone',      // 電話番号
      'exam_date',         // 検診日
      'exam_location',     // 検診場所
      'start_time',        // 開始時刻
      'end_time',          // 終了時刻
      'slot_interval',     // 枠間隔(分): 15 | 30 | 60
      'csv_file_id',       // CSVファイルID
      'status',            // ステータス: 未着手 | 処理中 | 完了
      'patient_count',     // 受診者数
      'completed_count',   // 完了数
      'current_step',      // 現在のワークフローステップ
      'created_at',        // 作成日時
      'updated_at',        // 更新日時
      'notes'              // 備考
    ],
    initialData: []
  },

  // 受診者マスタ
  patients: {
    sheetName: 'AS_受診者',
    headers: [
      'patient_id',        // 受診者ID (キー)
      'case_id',           // 案件ID (外部キー)
      'patient_no',        // 受診者番号
      'name',              // 氏名
      'name_kana',         // フリガナ
      'gender',            // 性別: M | F
      'birth_date',        // 生年月日
      'age',               // 年齢
      'exam_date',         // 受診日
      // スケジュール関連
      'scheduled_time',    // 予約時刻 (09:00形式)
      'slot_order',        // 順番
      'arrival_status',    // 未来院 | 来院済 | 完了
      // 一次検診結果（スクリーニング用）
      'primary_exam_date', // 一次検診日
      'primary_hdl',       // 一次 HDL-C
      'primary_ldl',       // 一次 LDL-C
      'primary_tg',        // 一次 中性脂肪
      'primary_fbs',       // 一次 空腹時血糖
      'primary_hba1c',     // 一次 HbA1c
      'primary_bp_sys',    // 一次 収縮期血圧
      'primary_bp_dia',    // 一次 拡張期血圧
      'primary_bmi',       // 一次 BMI
      'primary_waist',     // 一次 腹囲
      'screening_result',  // 対象 | 非対象 | 判定中
      // ステータス
      'status',            // ステータス: 未入力 | 入力中 | 確認待ち | 完了
      'current_step',      // 現在のステップ
      'blood_test_status', // 血液検査: 未 | 済
      'ultrasound_status', // 超音波: 未 | 済 | 対象外
      'guidance_status',   // 保健指導: 未 | 済
      'excel_exported',    // Excel出力済み
      'created_at',        // 作成日時
      'updated_at'         // 更新日時
    ],
    initialData: []
  },

  // 血液検査結果
  blood_tests: {
    sheetName: 'AS_血液検査',
    headers: [
      'blood_test_id',     // 血液検査ID (キー)
      'patient_id',        // 受診者ID (外部キー)
      'case_id',           // 案件ID (外部キー)
      // 糖代謝
      'fbs_value',         // 空腹時血糖 値
      'fbs_judgment',      // 空腹時血糖 判定
      'hba1c_value',       // HbA1c 値
      'hba1c_judgment',    // HbA1c 判定
      // 脂質
      'hdl_value',         // HDL-C 値
      'hdl_judgment',      // HDL-C 判定
      'ldl_value',         // LDL-C 値
      'ldl_judgment',      // LDL-C 判定
      'tg_value',          // 中性脂肪 値
      'tg_judgment',       // 中性脂肪 判定
      // 肝機能
      'ast_value',         // AST 値
      'ast_judgment',      // AST 判定
      'alt_value',         // ALT 値
      'alt_judgment',      // ALT 判定
      'ggt_value',         // γ-GTP 値
      'ggt_judgment',      // γ-GTP 判定
      // 腎機能
      'cr_value',          // クレアチニン 値
      'cr_judgment',       // クレアチニン 判定
      'egfr_value',        // eGFR 値
      'egfr_judgment',     // eGFR 判定
      'ua_value',          // 尿酸 値
      'ua_judgment',       // 尿酸 判定
      // 過去値（労災二次用）
      'prev_hdl',          // 前回HDL
      'prev_ldl',          // 前回LDL
      'prev_tg',           // 前回TG
      'prev_fbs',          // 前回FBS
      'prev_hba1c',        // 前回HbA1c
      // メタ
      'data_source',       // データソース: CSV | 手入力 | OCR
      'verified',          // 確認済み
      'created_at',        // 作成日時
      'updated_at'         // 更新日時
    ],
    initialData: []
  },

  // 超音波検査
  ultrasound: {
    sheetName: 'AS_超音波',
    headers: [
      'ultrasound_id',     // 超音波ID (キー)
      'patient_id',        // 受診者ID (外部キー)
      'case_id',           // 案件ID (外部キー)
      // 腹部
      'abd_judgment',      // 腹部判定: A | B | C | D
      'abd_findings',      // 腹部所見（自由記述）
      'abd_liver',         // 肝臓所見
      'abd_gallbladder',   // 胆嚢所見
      'abd_kidney',        // 腎臓所見
      'abd_spleen',        // 脾臓所見
      'abd_pancreas',      // 膵臓所見
      // 頸動脈
      'carotid_judgment',  // 頸動脈判定: A | B | C | D
      'carotid_findings',  // 頸動脈所見
      'carotid_imt_r',     // IMT右
      'carotid_imt_l',     // IMT左
      'carotid_plaque',    // プラーク有無
      // 心臓
      'echo_judgment',     // 心エコー判定
      'echo_findings',     // 心エコー所見
      // メタ
      'verified',          // 確認済み
      'created_at',        // 作成日時
      'updated_at'         // 更新日時
    ],
    initialData: []
  },

  // 保健指導
  guidance: {
    sheetName: 'AS_保健指導',
    headers: [
      'guidance_id',       // 保健指導ID (キー)
      'patient_id',        // 受診者ID (外部キー)
      'case_id',           // 案件ID (外部キー)
      'guidance_type',     // 指導種別: 情報提供 | 動機付け | 積極的
      'ai_generated',      // AI生成テキスト
      'final_text',        // 最終テキスト（編集後）
      'nutrition',         // 栄養指導
      'exercise',          // 運動指導
      'lifestyle',         // 生活習慣指導
      'medical_advice',    // 医療機関受診勧奨
      'verified',          // 確認済み
      'generated_at',      // 生成日時
      'created_at',        // 作成日時
      'updated_at'         // 更新日時
    ],
    initialData: []
  },

  // ワークフローステップ定義
  workflow_steps: {
    sheetName: 'AS_ワークフロー',
    headers: [
      'step_id',           // ステップID (キー)
      'exam_type',         // 検診種別
      'step_order',        // 順序
      'step_name',         // ステップ名
      'step_description',  // 説明
      'action_type',       // アクション種別: upload | input | review | generate | export
      'target_table',      // 対象テーブル
      'required',          // 必須
      'auto_advance',      // 自動進行
      'validation_rule'    // バリデーションルール
    ],
    initialData: [
      // 労災二次検診ワークフロー
      ['ROSAI_STEP_1', 'ROSAI_SECONDARY', 1, 'CSV取込', 'BML/ROSAIフォーマットのCSVをアップロード', 'upload', 'blood_tests', true, true, ''],
      ['ROSAI_STEP_2', 'ROSAI_SECONDARY', 2, '血液検査確認', '検査値と判定を確認・修正', 'review', 'blood_tests', true, false, ''],
      ['ROSAI_STEP_3', 'ROSAI_SECONDARY', 3, '超音波入力', '腹部・頸動脈超音波所見を入力', 'input', 'ultrasound', true, false, ''],
      ['ROSAI_STEP_4', 'ROSAI_SECONDARY', 4, '保健指導生成', 'AIで保健指導文を生成', 'generate', 'guidance', true, false, ''],
      ['ROSAI_STEP_5', 'ROSAI_SECONDARY', 5, '最終確認', '全体を確認してExcel出力', 'export', 'cases', true, false, ''],
      // 人間ドックワークフロー
      ['DOCK_STEP_1', 'DOCK', 1, 'CSV取込', 'BMLフォーマットのCSVをアップロード', 'upload', 'blood_tests', true, true, ''],
      ['DOCK_STEP_2', 'DOCK', 2, '検査結果確認', '全検査結果を確認', 'review', 'blood_tests', true, false, ''],
      ['DOCK_STEP_3', 'DOCK', 3, '所見入力', '各種所見を入力', 'input', 'ultrasound', false, false, ''],
      ['DOCK_STEP_4', 'DOCK', 4, '最終確認', '全体を確認して出力', 'export', 'cases', true, false, ''],
      // 定期検診ワークフロー
      ['REGULAR_STEP_1', 'REGULAR', 1, 'CSV取込', 'CSVをアップロード', 'upload', 'blood_tests', true, true, ''],
      ['REGULAR_STEP_2', 'REGULAR', 2, '確認・出力', '確認して出力', 'export', 'cases', true, false, '']
    ]
  },

  // 所見テンプレート
  findings_templates: {
    sheetName: 'AS_所見テンプレート',
    headers: [
      'template_id',       // テンプレートID
      'category',          // カテゴリ: 腹部 | 頸動脈 | 心臓
      'organ',             // 臓器
      'finding_code',      // 所見コード
      'finding_text',      // 所見テキスト
      'judgment',          // 対応判定: A | B | C | D | E
      'sort_order',        // 表示順
      'active'             // 有効
    ],
    initialData: [
      // ============================================
      // 頸動脈超音波（労災二次検診用）
      // ============================================
      ['CAROTID_A', '頸動脈', '頸動脈', 'NORMAL', '異常なし', 'A', 1, true],
      ['CAROTID_B', '頸動脈', '頸動脈', 'ALMOST_NORMAL', 'ほぼ正常', 'B', 2, true],
      ['CAROTID_C', '頸動脈', '頸動脈', 'FOLLOW_UP', '経過観察', 'C', 3, true],
      ['CAROTID_D', '頸動脈', '頸動脈', 'TREATMENT', '要治療', 'D', 4, true],
      ['CAROTID_E', '頸動脈', '頸動脈', 'FURTHER_EXAM', '要精密検査', 'E', 5, true],
      // 頸動脈 - 詳細所見（追記用）
      ['CAROTID_IMT_NORMAL', '頸動脈', '頸動脈', 'IMT_NORMAL', 'IMT正常範囲', 'A', 10, true],
      ['CAROTID_IMT_MILD', '頸動脈', '頸動脈', 'IMT_MILD', 'IMT軽度肥厚', 'B', 11, true],
      ['CAROTID_IMT_MOD', '頸動脈', '頸動脈', 'IMT_MOD', 'IMT中等度肥厚', 'C', 12, true],
      ['CAROTID_PLAQUE_SMALL', '頸動脈', '頸動脈', 'PLAQUE_SMALL', '小プラーク（狭窄なし）', 'C', 13, true],
      ['CAROTID_PLAQUE_MOD', '頸動脈', '頸動脈', 'PLAQUE_MOD', 'プラーク（軽度狭窄）', 'D', 14, true],
      ['CAROTID_STENOSIS', '頸動脈', '頸動脈', 'STENOSIS', '有意狭窄', 'E', 15, true],

      // ============================================
      // 心臓超音波（労災二次検診用）
      // ============================================
      ['ECHO_A', '心臓', '心臓', 'NORMAL', '異常なし', 'A', 1, true],
      ['ECHO_C', '心臓', '心臓', 'RECHECK_12M', '12ヶ月後再検査', 'C', 3, true],
      // 心臓 - 詳細所見（追記用）
      ['ECHO_VALVE_TR', '心臓', '心臓', 'VALVE_TR', '三尖弁逆流（軽度）', 'A', 10, true],
      ['ECHO_VALVE_MR', '心臓', '心臓', 'VALVE_MR', '僧帽弁逆流（軽度）', 'B', 11, true],
      ['ECHO_VALVE_AR', '心臓', '心臓', 'VALVE_AR', '大動脈弁逆流（軽度）', 'B', 12, true],
      ['ECHO_LVH_MILD', '心臓', '心臓', 'LVH_MILD', '左室肥大（軽度）', 'B', 13, true],
      ['ECHO_LVH_MOD', '心臓', '心臓', 'LVH_MOD', '左室肥大（中等度）', 'C', 14, true],
      ['ECHO_EF_LOW', '心臓', '心臓', 'EF_LOW', '左室駆出率低下', 'C', 15, true],
      ['ECHO_WALL_ABNORMAL', '心臓', '心臓', 'WALL_ABNORMAL', '壁運動異常', 'C', 16, true],
      ['ECHO_DIASTOLIC', '心臓', '心臓', 'DIASTOLIC', '拡張障害', 'C', 17, true],

      // ============================================
      // 腹部超音波（人間ドック用 - 将来対応）
      // ============================================
      // 肝臓
      ['ABD_LIVER_A', '腹部', '肝臓', 'NORMAL', '異常なし', 'A', 1, true],
      ['ABD_LIVER_B1', '腹部', '肝臓', 'FATTY_MILD', '軽度脂肪肝', 'B', 10, true],
      ['ABD_LIVER_B2', '腹部', '肝臓', 'CYST', '肝嚢胞', 'B', 11, true],
      ['ABD_LIVER_C1', '腹部', '肝臓', 'FATTY_MOD', '中等度脂肪肝', 'C', 20, true],
      ['ABD_LIVER_C2', '腹部', '肝臓', 'HEMANGIOMA', '肝血管腫', 'B', 12, true],
      // 胆嚢
      ['ABD_GB_A', '腹部', '胆嚢', 'NORMAL', '異常なし', 'A', 1, true],
      ['ABD_GB_B1', '腹部', '胆嚢', 'POLYP_SMALL', '胆嚢ポリープ(5mm未満)', 'B', 10, true],
      ['ABD_GB_C1', '腹部', '胆嚢', 'POLYP_LARGE', '胆嚢ポリープ(5mm以上)', 'C', 20, true],
      ['ABD_GB_C2', '腹部', '胆嚢', 'STONE', '胆石', 'C', 21, true],
      // 腎臓
      ['ABD_KIDNEY_A', '腹部', '腎臓', 'NORMAL', '異常なし', 'A', 1, true],
      ['ABD_KIDNEY_B1', '腹部', '腎臓', 'CYST', '腎嚢胞', 'B', 10, true],
      ['ABD_KIDNEY_B2', '腹部', '腎臓', 'STONE_SMALL', '腎結石（小）', 'B', 11, true],
      // 脾臓・膵臓
      ['ABD_SPLEEN_A', '腹部', '脾臓', 'NORMAL', '異常なし', 'A', 1, true],
      ['ABD_PANCREAS_A', '腹部', '膵臓', 'NORMAL', '異常なし', 'A', 1, true]
    ]
  },

  // 判定基準マスタ
  judgment_criteria: {
    sheetName: 'AS_判定基準',
    headers: [
      'criteria_id',       // 基準ID
      'item_code',         // 検査項目コード
      'item_name',         // 検査項目名
      'unit',              // 単位
      'gender',            // 性別: ALL | M | F
      'a_lower',           // A判定下限
      'a_upper',           // A判定上限
      'b_lower',           // B判定下限
      'b_upper',           // B判定上限
      'c_lower',           // C判定下限
      'c_upper',           // C判定上限
      'd_threshold',       // D判定閾値
      'exam_type',         // 対象検診種別
      'active'             // 有効
    ],
    initialData: [
      // 糖代謝
      ['FBS_ALL', 'FBS', '空腹時血糖', 'mg/dL', 'ALL', 70, 99, 100, 109, 110, 125, 126, 'ALL', true],
      ['HBA1C_ALL', 'HBA1C', 'HbA1c', '%', 'ALL', 4.6, 5.5, 5.6, 5.9, 6.0, 6.4, 6.5, 'ALL', true],
      // 脂質
      ['HDL_M', 'HDL', 'HDL-C', 'mg/dL', 'M', 40, 999, 35, 39, 30, 34, 29, 'ALL', true],
      ['HDL_F', 'HDL', 'HDL-C', 'mg/dL', 'F', 45, 999, 40, 44, 35, 39, 34, 'ALL', true],
      ['LDL_ALL', 'LDL', 'LDL-C', 'mg/dL', 'ALL', 0, 119, 120, 139, 140, 179, 180, 'ALL', true],
      ['TG_ALL', 'TG', '中性脂肪', 'mg/dL', 'ALL', 0, 149, 150, 199, 200, 399, 400, 'ALL', true],
      // 肝機能
      ['AST_ALL', 'AST', 'AST', 'U/L', 'ALL', 0, 30, 31, 40, 41, 50, 51, 'ALL', true],
      ['ALT_ALL', 'ALT', 'ALT', 'U/L', 'ALL', 0, 30, 31, 40, 41, 50, 51, 'ALL', true],
      ['GGT_ALL', 'GGT', 'γ-GTP', 'U/L', 'ALL', 0, 50, 51, 80, 81, 100, 101, 'ALL', true],
      // 腎機能
      ['CR_M', 'CR', 'クレアチニン', 'mg/dL', 'M', 0.6, 1.0, 1.01, 1.2, 1.21, 1.4, 1.41, 'ALL', true],
      ['CR_F', 'CR', 'クレアチニン', 'mg/dL', 'F', 0.4, 0.8, 0.81, 1.0, 1.01, 1.2, 1.21, 'ALL', true],
      ['EGFR_ALL', 'EGFR', 'eGFR', 'mL/min', 'ALL', 90, 999, 60, 89, 45, 59, 44, 'ALL', true],
      ['UA_M', 'UA', '尿酸', 'mg/dL', 'M', 0, 7.0, 7.1, 8.0, 8.1, 9.0, 9.1, 'ALL', true],
      ['UA_F', 'UA', '尿酸', 'mg/dL', 'F', 0, 6.0, 6.1, 7.0, 7.1, 8.0, 8.1, 'ALL', true]
    ]
  },

  // 設定
  settings: {
    sheetName: 'AS_設定',
    headers: [
      'key',               // 設定キー
      'value',             // 設定値
      'description',       // 説明
      'category'           // カテゴリ
    ],
    initialData: [
      ['CLAUDE_API_KEY', '', 'Claude API キー（環境変数推奨）', 'API'],
      ['WEBHOOK_URL', '', 'AppSheet Webhook URL', 'API'],
      ['DEFAULT_EXAM_TYPE', 'ROSAI_SECONDARY', 'デフォルト検診種別', 'システム'],
      ['AUTO_JUDGMENT', 'true', '自動判定有効', 'システム'],
      ['AI_GUIDANCE_ENABLED', 'true', 'AI保健指導生成有効', 'システム'],
      ['EXCEL_TEMPLATE_ROSAI', '', '労災二次Excelテンプレート ID', 'テンプレート'],
      ['EXCEL_TEMPLATE_DOCK', '', '人間ドックExcelテンプレート ID', 'テンプレート']
    ]
  }
};

// ============================================
// セットアップ関数
// ============================================

/**
 * メインセットアップ関数（UIあり版）
 * スプレッドシートのメニューから実行する場合はこちら
 */
function setupAppSheetTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 確認ダイアログ
  const response = ui.alert(
    'AppSheetテーブル作成',
    'AppSheet用のテーブルを作成します。\n既存のAS_で始まるシートは上書きされます。\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }

  const results = setupAppSheetTablesCore(ss);

  // 結果表示
  ui.alert('セットアップ完了', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * ★推奨★ UIなし版セットアップ関数
 * GASエディタから直接実行する場合はこちらを使用
 * 実行後、「表示」→「ログ」で結果を確認
 */
function setupAppSheetTablesNoUI() {
  Logger.log('=== AppSheet Tables Setup Start ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('❌ エラー: スプレッドシートが取得できません');
    Logger.log('   → スプレッドシートを開いた状態で実行してください');
    return;
  }

  Logger.log(`対象スプレッドシート: ${ss.getName()}`);

  const results = setupAppSheetTablesCore(ss);

  Logger.log('=== Setup Complete ===');
  results.forEach(r => Logger.log(r));

  return results;
}

/**
 * コア処理（UI有無共通）
 */
function setupAppSheetTablesCore(ss) {
  const results = [];
  const tableNames = Object.keys(APPSHEET_TABLES);

  Logger.log(`作成するテーブル数: ${tableNames.length}`);

  // 各テーブルを作成
  for (const [tableName, tableConfig] of Object.entries(APPSHEET_TABLES)) {
    Logger.log(`処理中: ${tableConfig.sheetName}...`);
    try {
      createOrUpdateSheet(ss, tableConfig);
      results.push(`✅ ${tableConfig.sheetName}: 作成完了`);
      Logger.log(`  → 完了`);
    } catch (error) {
      results.push(`❌ ${tableConfig.sheetName}: エラー - ${error.message}`);
      Logger.log(`  → エラー: ${error.message}`);
    }
  }

  return results;
}

/**
 * デバッグ用：1テーブルずつ作成
 * どのテーブルで問題が起きているか特定する場合に使用
 */
function setupSingleTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ★ここで作成したいテーブル名を指定 ★
  const targetTable = 'cases';  // cases, patients, blood_tests, ultrasound, guidance, workflow_steps, findings_templates, judgment_criteria, settings

  const tableConfig = APPSHEET_TABLES[targetTable];
  if (!tableConfig) {
    Logger.log(`テーブルが見つかりません: ${targetTable}`);
    return;
  }

  Logger.log(`Creating single table: ${tableConfig.sheetName}`);

  try {
    createOrUpdateSheet(ss, tableConfig);
    Logger.log(`✅ 完了: ${tableConfig.sheetName}`);
  } catch (error) {
    Logger.log(`❌ エラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * シートを作成または更新
 */
function createOrUpdateSheet(ss, tableConfig) {
  const { sheetName, headers, initialData } = tableConfig;

  // 既存シートを削除
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  // 新規作成
  sheet = ss.insertSheet(sheetName);

  // ヘッダー設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // 初期データ挿入
  if (initialData && initialData.length > 0) {
    const dataRange = sheet.getRange(2, 1, initialData.length, headers.length);
    dataRange.setValues(initialData);
  }

  // 列幅自動調整
  sheet.autoResizeColumns(1, headers.length);

  // ヘッダー行を固定
  sheet.setFrozenRows(1);

  return sheet;
}

// ============================================
// ヘルパー関数
// ============================================

/**
 * 新規案件IDを生成
 */
function generateCaseId() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  const randomStr = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `CASE_${dateStr}_${randomStr}`;
}

/**
 * 新規受診者IDを生成
 */
function generatePatientId(caseId) {
  const randomStr = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `PAT_${caseId.replace('CASE_', '')}_${randomStr}`;
}

/**
 * テーブルデータを取得
 */
function getTableData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

/**
 * テーブルに行を追加
 */
function appendTableRow(sheetName, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => rowData[h] || '');

  sheet.appendRow(row);
  return true;
}

/**
 * テーブル行を更新
 */
function updateTableRow(sheetName, keyColumn, keyValue, updateData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIndex = headers.indexOf(keyColumn);

  if (keyIndex === -1) throw new Error(`Key column not found: ${keyColumn}`);

  for (let i = 1; i < data.length; i++) {
    if (data[i][keyIndex] === keyValue) {
      // 更新
      for (const [key, value] of Object.entries(updateData)) {
        const colIndex = headers.indexOf(key);
        if (colIndex !== -1) {
          sheet.getRange(i + 1, colIndex + 1).setValue(value);
        }
      }
      return true;
    }
  }

  return false; // 該当行なし
}

/**
 * 案件一覧を取得（AppSheet用）
 */
function getCaseList(examType = null) {
  const cases = getTableData('AS_案件');

  if (examType) {
    return cases.filter(c => c.exam_type === examType);
  }

  return cases;
}

/**
 * 受診者一覧を取得（案件ID指定）
 */
function getPatientsByCaseId(caseId) {
  const patients = getTableData('AS_受診者');
  return patients.filter(p => p.case_id === caseId);
}

/**
 * ワークフローステップを取得
 */
function getWorkflowSteps(examType) {
  const steps = getTableData('AS_ワークフロー');
  return steps
    .filter(s => s.exam_type === examType)
    .sort((a, b) => a.step_order - b.step_order);
}

/**
 * 次のステップを取得
 */
function getNextStep(examType, currentStepId) {
  const steps = getWorkflowSteps(examType);
  const currentIndex = steps.findIndex(s => s.step_id === currentStepId);

  if (currentIndex === -1 || currentIndex >= steps.length - 1) {
    return null; // 最後のステップまたは見つからない
  }

  return steps[currentIndex + 1];
}

// ============================================
// カスタムメニュー（AppSheet版 - 必要時に手動で有効化）
// ============================================

// 元のmain.gsのonOpen()を使用するため、この関数はリネーム
// 使いたい場合は関数名を onOpen() に戻す
function onOpen_AppSheet_DISABLED() {
  const ui = SpreadsheetApp.getUi();

  // メイン業務メニュー
  ui.createMenu('🏥 健診システム')
    .addSubMenu(ui.createMenu('📋 案件管理')
      .addItem('新規案件登録', 'showNewCaseDialog')
      .addItem('案件一覧', 'showCaseList')
      .addItem('案件ステータス更新', 'showUpdateCaseStatusDialog'))
    .addSubMenu(ui.createMenu('👥 受診者管理')
      .addItem('受診者一括登録（名簿から）', 'showBulkPatientImportDialog')
      .addItem('受診者個別登録', 'showNewPatientDialog')
      .addItem('スケジュール取込', 'showScheduleImportDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔬 検査・判定')
      .addItem('血液検査CSV取込', 'showCsvImportDialog')
      .addItem('判定一括計算', 'calculateAllJudgments')
      .addItem('二次検診対象判定', 'calculateAllScreening'))
    .addSubMenu(ui.createMenu('📝 超音波・所見')
      .addItem('超音波入力シートへ移動', 'goToUltrasoundSheet')
      .addItem('所見テンプレート確認', 'showFindingsTemplates'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📄 出力')
      .addItem('Excel結果票出力', 'showExcelExportDialog')
      .addItem('AI保健指導生成', 'showGuidanceGenerateDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 設定・管理')
      .addItem('テーブル作成/更新', 'setupAppSheetTables')
      .addItem('テストデータ投入', 'insertTestData_RosaiSecondary')
      .addItem('テストデータ削除', 'clearTestData')
      .addItem('ワークフロー確認', 'showWorkflowSteps'))
    .addToUi();
}

// ============================================
// ダイアログ・UI関数
// ============================================

/**
 * 新規案件登録ダイアログ
 */
function showNewCaseDialog() {
  const ui = SpreadsheetApp.getUi();

  // 案件名
  const nameResult = ui.prompt('新規案件登録 (1/4)', '案件名（会社名など）を入力:', ui.ButtonSet.OK_CANCEL);
  if (nameResult.getSelectedButton() !== ui.Button.OK) return;
  const caseName = nameResult.getResponseText();

  // 検診種別
  const typeResult = ui.prompt('新規案件登録 (2/4)',
    '検診種別を入力:\n1: 労災二次検診\n2: 人間ドック\n3: 定期検診\n\n番号を入力:',
    ui.ButtonSet.OK_CANCEL);
  if (typeResult.getSelectedButton() !== ui.Button.OK) return;
  const typeMap = { '1': 'ROSAI_SECONDARY', '2': 'DOCK', '3': 'REGULAR' };
  const examType = typeMap[typeResult.getResponseText()] || 'ROSAI_SECONDARY';

  // 検診日
  const dateResult = ui.prompt('新規案件登録 (3/4)', '検診日を入力 (例: 2024-12-15):', ui.ButtonSet.OK_CANCEL);
  if (dateResult.getSelectedButton() !== ui.Button.OK) return;
  const examDate = dateResult.getResponseText();

  // 確認
  const confirmResult = ui.alert('確認',
    `以下の内容で案件を登録します:\n\n案件名: ${caseName}\n種別: ${examType}\n検診日: ${examDate}`,
    ui.ButtonSet.YES_NO);

  if (confirmResult !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }

  // 登録実行
  const caseId = createNewCase(caseName, examType, examDate);
  ui.alert('登録完了', `案件を登録しました\n\n案件ID: ${caseId}`, ui.ButtonSet.OK);
}

/**
 * 新規案件を作成
 */
function createNewCase(caseName, examType, examDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AS_案件');

  const now = new Date();
  const caseId = 'CASE_' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');

  sheet.appendRow([
    caseId,           // case_id
    caseName,         // case_name
    examType,         // exam_type
    '',               // client_name
    '',               // client_address
    '',               // client_contact
    '',               // client_phone
    examDate,         // exam_date
    '',               // exam_location
    '',               // start_time
    '',               // end_time
    30,               // slot_interval
    '',               // csv_file_id
    '未着手',         // status
    0,                // patient_count
    0,                // completed_count
    '',               // current_step
    now,              // created_at
    now,              // updated_at
    ''                // notes
  ]);

  return caseId;
}

/**
 * 受診者個別登録ダイアログ
 */
function showNewPatientDialog() {
  const ui = SpreadsheetApp.getUi();

  // 案件選択
  const cases = getCaseList();
  if (cases.length === 0) {
    ui.alert('エラー', '先に案件を登録してください', ui.ButtonSet.OK);
    return;
  }

  let caseOptions = cases.map((c, i) => `${i + 1}: ${c.case_name} (${c.exam_date})`).join('\n');
  const caseResult = ui.prompt('受診者登録 (1/4)', `案件を選択:\n${caseOptions}\n\n番号を入力:`, ui.ButtonSet.OK_CANCEL);
  if (caseResult.getSelectedButton() !== ui.Button.OK) return;
  const selectedCase = cases[parseInt(caseResult.getResponseText()) - 1];

  if (!selectedCase) {
    ui.alert('エラー', '有効な番号を入力してください', ui.ButtonSet.OK);
    return;
  }

  // 氏名
  const nameResult = ui.prompt('受診者登録 (2/4)', '氏名を入力:', ui.ButtonSet.OK_CANCEL);
  if (nameResult.getSelectedButton() !== ui.Button.OK) return;
  const name = nameResult.getResponseText();

  // 性別
  const genderResult = ui.prompt('受診者登録 (3/4)', '性別を入力 (M: 男性 / F: 女性):', ui.ButtonSet.OK_CANCEL);
  if (genderResult.getSelectedButton() !== ui.Button.OK) return;
  const gender = genderResult.getResponseText().toUpperCase();

  // 生年月日
  const birthResult = ui.prompt('受診者登録 (4/4)', '生年月日を入力 (例: 1970-05-15):', ui.ButtonSet.OK_CANCEL);
  if (birthResult.getSelectedButton() !== ui.Button.OK) return;
  const birthDate = birthResult.getResponseText();

  // 登録実行
  const patientId = createNewPatient(selectedCase.case_id, name, gender, birthDate);
  ui.alert('登録完了', `受診者を登録しました\n\n受診者ID: ${patientId}`, ui.ButtonSet.OK);
}

/**
 * 新規受診者を作成
 */
function createNewPatient(caseId, name, gender, birthDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AS_受診者');

  const now = new Date();
  const patientId = 'PAT_' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMddHHmmss');

  // 年齢計算
  let age = '';
  if (birthDate) {
    const birth = new Date(birthDate);
    age = Math.floor((now - birth) / (365.25 * 24 * 60 * 60 * 1000));
  }

  // 受診者数をカウントしてpatient_noを設定
  const existingPatients = getPatientsByCaseId(caseId);
  const patientNo = String(existingPatients.length + 1).padStart(3, '0');

  sheet.appendRow([
    patientId,        // patient_id
    caseId,           // case_id
    patientNo,        // patient_no
    name,             // name
    '',               // name_kana
    gender,           // gender
    birthDate,        // birth_date
    age,              // age
    '',               // exam_date
    '',               // scheduled_time
    '',               // slot_order
    '未来院',         // arrival_status
    // 一次検診結果（空）
    '', '', '', '', '', '', '', '', '', '',
    '',               // screening_result
    '未入力',         // status
    1,                // current_step
    '未',             // blood_test_status
    '未',             // ultrasound_status
    '未',             // guidance_status
    false,            // excel_exported
    now,              // created_at
    now               // updated_at
  ]);

  // 案件の受診者数を更新
  updateCasePatientCount(caseId);

  // 超音波の初期行も作成
  const ultrasoundSheet = ss.getSheetByName('AS_超音波');
  ultrasoundSheet.appendRow([
    `US_${patientId}`, patientId, caseId,
    '', '', '', '', '', '', '',  // 腹部
    '', '', '', '', '',          // 頸動脈
    '', '',                      // 心臓
    false, now, now
  ]);

  return patientId;
}

/**
 * 案件の受診者数を更新
 */
function updateCasePatientCount(caseId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AS_案件');
  const patients = getPatientsByCaseId(caseId);

  updateTableRow('AS_案件', 'case_id', caseId, {
    patient_count: patients.length,
    updated_at: new Date()
  });
}

/**
 * 案件ステータス更新ダイアログ
 */
function showUpdateCaseStatusDialog() {
  const ui = SpreadsheetApp.getUi();

  const cases = getCaseList();
  if (cases.length === 0) {
    ui.alert('案件がありません');
    return;
  }

  let caseOptions = cases.map((c, i) => `${i + 1}: ${c.case_name} [${c.status}]`).join('\n');
  const caseResult = ui.prompt('ステータス更新', `案件を選択:\n${caseOptions}\n\n番号:`, ui.ButtonSet.OK_CANCEL);
  if (caseResult.getSelectedButton() !== ui.Button.OK) return;

  const selectedCase = cases[parseInt(caseResult.getResponseText()) - 1];
  if (!selectedCase) return;

  const statusResult = ui.prompt('新しいステータス',
    '1: 未着手\n2: 処理中\n3: 完了\n\n番号:', ui.ButtonSet.OK_CANCEL);
  if (statusResult.getSelectedButton() !== ui.Button.OK) return;

  const statusMap = { '1': '未着手', '2': '処理中', '3': '完了' };
  const newStatus = statusMap[statusResult.getResponseText()];

  if (newStatus) {
    updateTableRow('AS_案件', 'case_id', selectedCase.case_id, {
      status: newStatus,
      updated_at: new Date()
    });
    ui.alert('更新完了', `ステータスを「${newStatus}」に更新しました`, ui.ButtonSet.OK);
  }
}

/**
 * 超音波入力シートへ移動
 */
function goToUltrasoundSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AS_超音波');
  if (sheet) {
    ss.setActiveSheet(sheet);
    SpreadsheetApp.getUi().alert('AS_超音波シートに移動しました\n\n頸動脈判定(carotid_judgment): A/B/C/D/E\n心臓判定(echo_judgment): A/C');
  }
}

/**
 * 判定一括計算（プレースホルダー）
 */
function calculateAllJudgments() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('判定計算', 'AS_血液検査の全行に対して判定を計算します\n\n（この機能は既存のjudgmentEngine.gsと連携予定）', ui.ButtonSet.OK);
  // TODO: judgmentEngine.gsの関数を呼び出し
}

/**
 * 二次検診対象判定（プレースホルダー）
 */
function calculateAllScreening() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('対象判定', 'AS_受診者の一次検診結果から二次検診対象を判定します\n\n（この機能はappSheetBridge.gsと連携予定）', ui.ButtonSet.OK);
  // TODO: appSheetBridge.gsのcalculateScreeningResult関数を呼び出し
}

/**
 * CSV取込ダイアログ（プレースホルダー）
 */
function showCsvImportDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('CSV取込', '血液検査CSVの取込機能\n\n（既存のcsvParser.gsと連携予定）', ui.ButtonSet.OK);
}

/**
 * Excel出力ダイアログ（プレースホルダー）
 */
function showExcelExportDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Excel出力', '結果票Excel出力機能\n\n（既存のexcelExporter.gsと連携予定）', ui.ButtonSet.OK);
}

/**
 * AI保健指導生成ダイアログ（プレースホルダー）
 */
function showGuidanceGenerateDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('AI保健指導生成', 'Claude APIで保健指導文を生成\n\n（既存のclaudeApi.gsと連携予定）', ui.ButtonSet.OK);
}

/**
 * 名簿一括取込ダイアログ（プレースホルダー）
 */
function showBulkPatientImportDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('名簿一括取込', '別シートまたはCSVから受診者を一括登録\n\n（今後実装予定）', ui.ButtonSet.OK);
}

/**
 * スケジュール取込ダイアログ（プレースホルダー）
 */
function showScheduleImportDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('スケジュール取込', '外部スケジュールシートから受診者と時間枠を取込\n\n（appSheetBridge.gsと連携予定）', ui.ButtonSet.OK);
}

function showCaseList() {
  const cases = getCaseList();
  const ui = SpreadsheetApp.getUi();

  if (cases.length === 0) {
    ui.alert('案件なし', '登録された案件がありません。', ui.ButtonSet.OK);
    return;
  }

  const summary = cases.map(c =>
    `${c.case_id}: ${c.case_name} (${c.exam_type}) - ${c.status}`
  ).join('\n');

  ui.alert('案件一覧', summary, ui.ButtonSet.OK);
}

function showWorkflowSteps() {
  const ui = SpreadsheetApp.getUi();
  const examTypes = ['ROSAI_SECONDARY', 'DOCK', 'REGULAR'];

  let summary = '';
  for (const examType of examTypes) {
    const steps = getWorkflowSteps(examType);
    if (steps.length > 0) {
      summary += `\n【${examType}】\n`;
      steps.forEach(s => {
        summary += `  ${s.step_order}. ${s.step_name}: ${s.step_description}\n`;
      });
    }
  }

  ui.alert('ワークフロー定義', summary || 'ワークフローが未定義です', ui.ButtonSet.OK);
}

// ============================================
// テストデータ投入
// ============================================

/**
 * ★テスト用★ 労災二次検診のサンプルデータを投入
 * GASエディタから直接実行
 */
function insertTestData_RosaiSecondary() {
  Logger.log('=== テストデータ投入開始 ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

  // テスト案件ID
  const caseId = 'TEST_ROSAI_' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');

  // 1. 案件データ投入
  Logger.log('1. 案件データ投入...');
  const casesSheet = ss.getSheetByName('AS_案件');
  if (!casesSheet) {
    Logger.log('❌ AS_案件シートがありません。setupAppSheetTablesNoUI()を先に実行してください。');
    return;
  }

  casesSheet.appendRow([
    caseId,                    // case_id
    'テスト株式会社（労災二次）',  // case_name
    'ROSAI_SECONDARY',         // exam_type
    'テスト株式会社',           // client_name
    '東京都千代田区丸の内1-1-1', // client_address
    '山田太郎',                 // client_contact
    '03-1234-5678',            // client_phone
    today,                     // exam_date
    '当院',                    // exam_location
    '09:00',                   // start_time
    '12:00',                   // end_time
    30,                        // slot_interval
    '',                        // csv_file_id
    '処理中',                  // status
    3,                         // patient_count
    0,                         // completed_count
    'ROSAI_STEP_1',           // current_step
    now,                       // created_at
    now,                       // updated_at
    'テスト用案件'             // notes
  ]);
  Logger.log(`  → 案件作成: ${caseId}`);

  // 2. 受診者データ投入（3名）
  Logger.log('2. 受診者データ投入...');
  const patientsSheet = ss.getSheetByName('AS_受診者');

  const testPatients = [
    {
      id: 'PAT_001',
      name: '検査太郎',
      kana: 'ケンサタロウ',
      gender: 'M',
      birth: '1970-05-15',
      age: 54,
      time: '09:00',
      // 一次検診結果（二次検診対象者パターン: 高血圧+脂質異常）
      primary: {
        date: '2024-10-01',
        hdl: 42, ldl: 158, tg: 210,
        fbs: 102, hba1c: 5.8,
        bp_sys: 148, bp_dia: 92,
        bmi: 26.5, waist: 88
      }
    },
    {
      id: 'PAT_002',
      name: '健診花子',
      kana: 'ケンシンハナコ',
      gender: 'F',
      birth: '1975-08-22',
      age: 49,
      time: '09:30',
      // 一次検診結果（二次検診対象者パターン: 高血圧+糖代謝異常）
      primary: {
        date: '2024-10-01',
        hdl: 58, ldl: 128, tg: 145,
        fbs: 118, hba1c: 6.2,
        bp_sys: 152, bp_dia: 88,
        bmi: 24.2, waist: 82
      }
    },
    {
      id: 'PAT_003',
      name: '診断次郎',
      kana: 'シンダンジロウ',
      gender: 'M',
      birth: '1968-12-03',
      age: 56,
      time: '10:00',
      // 一次検診結果（二次検診対象者パターン: 高血圧+脂質+糖代謝）
      primary: {
        date: '2024-10-01',
        hdl: 35, ldl: 172, tg: 285,
        fbs: 132, hba1c: 6.8,
        bp_sys: 162, bp_dia: 98,
        bmi: 28.1, waist: 95
      }
    }
  ];

  testPatients.forEach((p, idx) => {
    patientsSheet.appendRow([
      p.id,                    // patient_id
      caseId,                  // case_id
      `R${String(idx + 1).padStart(3, '0')}`,  // patient_no
      p.name,                  // name
      p.kana,                  // name_kana
      p.gender,                // gender
      p.birth,                 // birth_date
      p.age,                   // age
      today,                   // exam_date
      p.time,                  // scheduled_time
      idx + 1,                 // slot_order
      '未来院',                // arrival_status
      // 一次検診結果
      p.primary.date,          // primary_exam_date
      p.primary.hdl,           // primary_hdl
      p.primary.ldl,           // primary_ldl
      p.primary.tg,            // primary_tg
      p.primary.fbs,           // primary_fbs
      p.primary.hba1c,         // primary_hba1c
      p.primary.bp_sys,        // primary_bp_sys
      p.primary.bp_dia,        // primary_bp_dia
      p.primary.bmi,           // primary_bmi
      p.primary.waist,         // primary_waist
      '対象',                  // screening_result（手動設定）
      '未入力',                // status
      1,                       // current_step
      '未',                    // blood_test_status
      '未',                    // ultrasound_status
      '未',                    // guidance_status
      false,                   // excel_exported
      now,                     // created_at
      now                      // updated_at
    ]);
    Logger.log(`  → 受診者作成: ${p.id} ${p.name}`);
  });

  // 3. 血液検査データ投入（二次検診当日の検査結果）
  Logger.log('3. 血液検査データ投入...');
  const bloodSheet = ss.getSheetByName('AS_血液検査');

  const bloodTestData = [
    {
      patientId: 'PAT_001',
      fbs: 98, hba1c: 5.6,
      hdl: 45, ldl: 142, tg: 185,
      ast: 28, alt: 32, ggt: 48,
      cr: 0.92, egfr: 72, ua: 6.8
    },
    {
      patientId: 'PAT_002',
      fbs: 108, hba1c: 6.0,
      hdl: 62, ldl: 118, tg: 128,
      ast: 22, alt: 18, ggt: 25,
      cr: 0.68, egfr: 85, ua: 5.2
    },
    {
      patientId: 'PAT_003',
      fbs: 122, hba1c: 6.5,
      hdl: 38, ldl: 165, tg: 248,
      ast: 35, alt: 42, ggt: 68,
      cr: 1.05, egfr: 62, ua: 7.5
    }
  ];

  bloodTestData.forEach(b => {
    bloodSheet.appendRow([
      `BT_${b.patientId}`,     // blood_test_id
      b.patientId,             // patient_id
      caseId,                  // case_id
      b.fbs, '',               // fbs_value, fbs_judgment
      b.hba1c, '',             // hba1c_value, hba1c_judgment
      b.hdl, '',               // hdl_value, hdl_judgment
      b.ldl, '',               // ldl_value, ldl_judgment
      b.tg, '',                // tg_value, tg_judgment
      b.ast, '',               // ast_value, ast_judgment
      b.alt, '',               // alt_value, alt_judgment
      b.ggt, '',               // ggt_value, ggt_judgment
      b.cr, '',                // cr_value, cr_judgment
      b.egfr, '',              // egfr_value, egfr_judgment
      b.ua, '',                // ua_value, ua_judgment
      '', '', '', '', '',      // prev values
      'テスト',                // data_source
      false,                   // verified
      now,                     // created_at
      now                      // updated_at
    ]);
    Logger.log(`  → 血液検査作成: ${b.patientId}`);
  });

  // 4. 超音波データ（空の初期行）
  Logger.log('4. 超音波データ初期化...');
  const ultrasoundSheet = ss.getSheetByName('AS_超音波');

  testPatients.forEach(p => {
    ultrasoundSheet.appendRow([
      `US_${p.id}`,            // ultrasound_id
      p.id,                    // patient_id
      caseId,                  // case_id
      '', '', '', '', '', '', '',  // abd fields (腹部)
      '', '', '', '', '',      // carotid fields (頸動脈)
      '', '',                  // echo fields (心臓)
      false,                   // verified
      now,                     // created_at
      now                      // updated_at
    ]);
    Logger.log(`  → 超音波初期化: ${p.id}`);
  });

  Logger.log('=== テストデータ投入完了 ===');
  Logger.log(`案件ID: ${caseId}`);
  Logger.log(`受診者数: ${testPatients.length}名`);
  Logger.log('');
  Logger.log('【次のステップ】');
  Logger.log('1. AS_超音波シートで頸動脈・心臓の判定と所見を入力');
  Logger.log('2. 判定はテンプレートから選択（A/B/C/D/E）');
  Logger.log('3. AppSheetに接続してUIを確認');
}

/**
 * テストデータを削除
 */
function clearTestData() {
  Logger.log('=== テストデータ削除開始 ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['AS_案件', 'AS_受診者', 'AS_血液検査', 'AS_超音波', 'AS_保健指導'];

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // ヘッダー行以外を削除
        sheet.deleteRows(2, lastRow - 1);
        Logger.log(`✅ ${sheetName}: データ削除完了`);
      } else {
        Logger.log(`⏭️ ${sheetName}: データなし`);
      }
    }
  });

  Logger.log('=== テストデータ削除完了 ===');
}

/**
 * 所見テンプレート一覧を表示（確認用）
 */
function showFindingsTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AS_所見テンプレート');

  if (!sheet) {
    Logger.log('❌ AS_所見テンプレートシートがありません');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  Logger.log('=== 所見テンプレート一覧 ===');

  // カテゴリごとにグループ化
  const categories = {};
  data.slice(1).forEach(row => {
    const category = row[1];  // category列
    if (!categories[category]) {
      categories[category] = [];
    }
    categories[category].push({
      judgment: row[5],       // judgment列
      text: row[4]            // finding_text列
    });
  });

  for (const [category, items] of Object.entries(categories)) {
    Logger.log(`\n【${category}】`);
    items.forEach(item => {
      Logger.log(`  ${item.judgment}: ${item.text}`);
    });
  }
}
