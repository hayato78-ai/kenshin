/**
 * MasterData.gs - 検査項目マスタデータ定義
 * 検査項目マスタ.json / GAS用設定データ.json から変換
 *
 * @version 2.0.0
 * @date 2025-12-24
 * @description 統一ID体系導入 - BMLコード一元化
 *
 * ID体系:
 * - BML検査項目: 7桁BMLコード（例: 0000454 = 中性脂肪）
 * - 医療機関独自: H + カテゴリ2桁 + 連番4桁
 *   - H09xxxx: 基本情報
 *   - H01xxxx: 身体測定
 *   - H02xxxx: 画像診断
 *   - H03xxxx: 聴力検査
 *   - H04xxxx: 眼科検査
 */

// ============================================
// 後方互換用マッピング（英語キー → 統一ID）
// ============================================
const LEGACY_KEY_TO_UNIFIED_ID = {
  // 基本情報
  'COMPANY': 'H090001', 'DEPARTMENT': 'H090002', 'NAME_KANA': 'H090003', 'NAME': 'H090004',
  'BIRTHDATE': 'H090005', 'AGE': 'H090006', 'SEX': 'H090007', 'EXAM_DATE': 'H090010',
  'COURSE': 'H090011', 'RECEPTION_NO': 'H090012', 'OVERALL_JUDGMENT': 'H090020',
  'OVERALL_COMMENT': 'H090021', 'MEDICAL_HISTORY': 'H090030', 'SUBJECTIVE_SYMPTOMS': 'H090031',
  'OBJECTIVE_SYMPTOMS': 'H090032',
  // 身体測定
  'HEIGHT': 'H010001', 'WEIGHT': 'H010002', 'STANDARD_WEIGHT': 'H010003', 'BMI': 'H010004',
  'BODY_FAT': 'H010005', 'WAIST_M': 'H010006', 'WAIST_F': 'H010007',
  'BP_SYSTOLIC_1': 'H010010', 'BP_DIASTOLIC_1': 'H010011', 'BP_SYSTOLIC_2': 'H010012',
  'BP_DIASTOLIC_2': 'H010013', 'PULSE': 'H010020',
  // 画像診断
  'PHYSICAL_EXAM': 'H020001', 'CHEST_XRAY': 'H020010', 'ABDOMINAL_XRAY': 'H020011',
  'ECG': 'H020020', 'UPPER_GI_ENDOSCOPY': 'H020030', 'LOWER_GI_ENDOSCOPY': 'H020031',
  'ABDOMINAL_US': 'H020040', 'CAROTID_US': 'H020041', 'CARDIAC_US': 'H020042',
  'THYROID_US': 'H020043', 'CHEST_CT': 'H020050', 'ABDOMINAL_CT': 'H020051',
  'BRAIN_MRI': 'H020060', 'BRAIN_MRA': 'H020061',
  'ABD_MRI_MRCP': 'H020062', 'DWI': 'H020063', 'CAROTID_MRA': 'H020064', 'UPPER_ABD_MRI': 'H020065',
  'MRCP': 'H020066', 'UTERUS_OVARY_MRI': 'H020067', 'MAMMO_MRI': 'H020068', 'PROSTATE_MRI': 'H020069',
  // 聴力
  'HEARING_R_1000': 'H030001', 'HEARING_L_1000': 'H030002', 'HEARING_R_4000': 'H030003',
  'HEARING_L_4000': 'H030004',
  // 眼科
  'VISION_NAKED_R': 'H040001', 'VISION_NAKED_L': 'H040002', 'VISION_CORRECTED_R': 'H040003',
  'VISION_CORRECTED_L': 'H040004', 'IOP_R': 'H040010', 'IOP_L': 'H040011',
  'FUNDUS_R': 'H040020', 'FUNDUS_L': 'H040021',
  // 尿検査（BMLコード）
  'URINE_PROTEIN': '0000051', 'URINE_GLUCOSE': '0000055', 'URINE_OCCULT_BLOOD': '0000063',
  'UROBILINOGEN': '0000057', 'URINE_PH': '0000061', 'URINE_BILIRUBIN': '0000059',
  'URINE_KETONE': '0000062', 'URINE_SG': '0000060',
  // 尿沈渣（医療機関独自）
  'URINE_WBC': 'H050001', 'URINE_RBC': 'H050002', 'URINE_EPITHELIAL': 'H050003', 'URINE_BACTERIA': 'H050004',
  // 便検査（医療機関独自）
  'FOBT_1': 'H060001', 'FOBT_2': 'H060002',
  // 肺機能（医療機関独自）
  'VC': 'H070001', 'FEV1': 'H070002', 'PERCENT_VC': 'H070003', 'FEV1_PERCENT': 'H070004',
  // 腫瘍マーカー（医療機関独自）
  'PSA': 'H080001', 'CEA': 'H080002', 'CA19_9': 'H080003', 'CA125': 'H080004',
  'AFP': 'H080005', 'NSE': 'H080006', 'CYFRA21_1': 'H080007', 'SCC': 'H080008',
  'PROGRP': 'H080009', 'PIVKA2': 'H080010', 'ELASTASE1': 'H080011', 'P53': 'H080012',
  // 感染症（医療機関独自）
  'HBS_AB': 'H100001', 'HIV_AB': 'H100002',
  // 甲状腺（医療機関独自）
  'FT3': 'H110001', 'FT4': 'H110002', 'TSH': 'H110003',
  // 血液型（医療機関独自）
  'BLOOD_TYPE_ABO': 'H120001', 'BLOOD_TYPE_RH': 'H120002',
  // 血液学検査（BMLコード）
  'WBC': '0000301', 'RBC': '0000302', 'HEMOGLOBIN': '0000303', 'HEMATOCRIT': '0000304',
  'PLATELET': '0000308', 'MCV': '0000305', 'MCH': '0000306', 'MCHC': '0000307',
  'NEUT': '0001889', 'LYMPHO': '0001885', 'MONO': '0001886', 'EOS': '0001882', 'BASO': '0001881',
  // 肝胆膵機能（BMLコード）
  'TOTAL_PROTEIN': '0000401', 'ALBUMIN': '0000417', 'AST': '0000481', 'ALT': '0000482',
  'GGT': '0000484', 'ALP': '0013067', 'LDH': '0013380', 'CHE': '0000491',
  'AMYLASE': '0000501', 'T_BIL': '0000472',
  // 脂質検査（BMLコード）
  'TOTAL_CHOLESTEROL': '0000453', 'TG': '0000454', 'HDL_C': '0000460', 'LDL_C': '0000410',
  'NON_HDL_C': '0003845',
  // 糖代謝（BMLコード）
  'FBS': '0000503', 'HBA1C': '0003317',
  // 腎機能（BMLコード）
  'CREATININE': '0000413', 'BUN': '0000409', 'EGFR': '0002696', 'UA': '0000407',
  // 電解質（BMLコード）
  'CK': '0000497', 'NA': '0000421', 'K': '0000423', 'CL': '0000425', 'CA': '0000427',
  // 感染症（BMLコード）
  'HBS_AG': '0000740', 'HCV_AB': '0003795', 'RPR': '0000905', 'TPHA': '0000911',
  // 心臓マーカー（BMLコード）
  'NT_PROBNP': '0003550', 'CRP': '0000658',
  // 追加検査（BMLコード）
  'FE': '0000435', 'TIBC': '0000437', 'RF': '0002490',
  'PT': '0000331', 'APTT': '0000334',
  'SPUTUM_CYTOLOGY': '0005972', 'UREA_BREATH_TEST': '0003294', 'HP_AB': '0013413'
};

// ============================================
// 検査項目マスタデータ（統一ID版）
// ============================================

const EXAM_ITEM_MASTER_DATA = [
  // 基本情報 - 受診者情報
  { item_id: 'H090001', item_name: '事業所', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'COMPANY' },
  { item_id: 'H090002', item_name: '所属', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'DEPARTMENT' },
  { item_id: 'H090003', item_name: 'カナ氏名', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'NAME_KANA' },
  { item_id: 'H090004', item_name: '氏名', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'NAME' },
  { item_id: 'H090005', item_name: '生年月日', category: '基本情報', subcategory: '受診者情報', data_type: 'date', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'BIRTHDATE' },
  { item_id: 'H090006', item_name: '年齢', category: '基本情報', subcategory: '受診者情報', data_type: 'integer', unit: '才', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'AGE' },
  { item_id: 'H090007', item_name: '性別', category: '基本情報', subcategory: '受診者情報', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'SEX' },

  // 基本情報 - 受診情報
  { item_id: 'H090010', item_name: '受診日', category: '基本情報', subcategory: '受診情報', data_type: 'date', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'EXAM_DATE' },
  { item_id: 'H090011', item_name: '受診コース', category: '基本情報', subcategory: '受診情報', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'COURSE' },
  { item_id: 'H090012', item_name: '受診No', category: '基本情報', subcategory: '受診情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'RECEPTION_NO' },

  // 基本情報 - 判定
  { item_id: 'H090020', item_name: '総合判定', category: '基本情報', subcategory: '判定', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'OVERALL_JUDGMENT' },
  { item_id: 'H090021', item_name: '総合所見', category: '基本情報', subcategory: '判定', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'OVERALL_COMMENT' },

  // 基本情報 - 問診
  { item_id: 'H090030', item_name: '既往歴', category: '基本情報', subcategory: '問診', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'MEDICAL_HISTORY' },
  { item_id: 'H090031', item_name: '自覚症状', category: '基本情報', subcategory: '問診', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'SUBJECTIVE_SYMPTOMS' },
  { item_id: 'H090032', item_name: '他覚症状', category: '基本情報', subcategory: '問診', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'OBJECTIVE_SYMPTOMS' },

  // 画像診断
  { item_id: 'H020001', item_name: '診察所見', category: '画像診断', subcategory: '診察', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'PHYSICAL_EXAM' },
  { item_id: 'H020010', item_name: '胸部X線', category: '画像診断', subcategory: 'X線', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'CHEST_XRAY' },
  { item_id: 'H020011', item_name: '腹部X線', category: '画像診断', subcategory: 'X線', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'ABDOMINAL_XRAY' },
  { item_id: 'H020020', item_name: '心電図', category: '画像診断', subcategory: '心電図', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'ECG' },
  { item_id: 'H020030', item_name: '上部消化管内視鏡検査(胃カメラ)', category: '画像診断', subcategory: '消化器', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'UPPER_GI_ENDOSCOPY' },
  { item_id: 'H020031', item_name: '下部消化管内視鏡検査(大腸カメラ)', category: '画像診断', subcategory: '大腸', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'LOWER_GI_ENDOSCOPY' },
  { item_id: 'H020040', item_name: '腹部超音波', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'ABDOMINAL_US' },
  { item_id: 'H020041', item_name: '頸部超音波', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CAROTID_US' },
  { item_id: 'H020042', item_name: '心臓超音波検査', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CARDIAC_US' },
  { item_id: 'H020043', item_name: '甲状腺超音波検査', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'THYROID_US' },
  { item_id: 'H020050', item_name: '胸部CT', category: '画像診断', subcategory: 'CT', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CHEST_CT' },
  { item_id: 'H020051', item_name: '腹部CT', category: '画像診断', subcategory: 'CT', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'ABDOMINAL_CT' },
  { item_id: 'H020060', item_name: '脳MRI', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'BRAIN_MRI' },
  { item_id: 'H020061', item_name: '脳MRA', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'BRAIN_MRA' },
  { item_id: 'H020062', item_name: '腹部MRI（肝胆膵）＋MRCP', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'ABD_MRI_MRCP' },
  { item_id: 'H020063', item_name: 'DWI', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'DWI' },
  { item_id: 'H020064', item_name: '頸動脈MRA', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CAROTID_MRA' },
  { item_id: 'H020065', item_name: '上腹部MRI', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'UPPER_ABD_MRI' },
  { item_id: 'H020066', item_name: 'MRCP', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'MRCP' },
  { item_id: 'H020067', item_name: '子宮・卵巣MRI', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'UTERUS_OVARY_MRI' },
  { item_id: 'H020068', item_name: 'マンモMRI', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'MAMMO_MRI' },
  { item_id: 'H020069', item_name: '前立腺MRI', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'PROSTATE_MRI' },

  // 身体測定
  { item_id: 'H010001', item_name: '身長', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'cm', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'HEIGHT' },
  { item_id: 'H010002', item_name: '体重', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'kg', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'WEIGHT' },
  { item_id: 'H010003', item_name: '標準体重', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'kg', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'STANDARD_WEIGHT' },
  { item_id: 'H010004', item_name: 'BMI', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'BMI' },
  { item_id: 'H010005', item_name: '体脂肪率', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'BODY_FAT' },
  { item_id: 'H010006', item_name: '腹囲(男性)', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'cm', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'WAIST_M' },
  { item_id: 'H010007', item_name: '腹囲(女性)', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'cm', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'WAIST_F' },

  // 血圧
  { item_id: 'H010010', item_name: '血圧1回目（最高）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'BP_SYSTOLIC_1' },
  { item_id: 'H010011', item_name: '血圧1回目（最低）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'BP_DIASTOLIC_1' },
  { item_id: 'H010012', item_name: '血圧2回目（最高）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'BP_SYSTOLIC_2' },
  { item_id: 'H010013', item_name: '血圧2回目（最低）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'BP_DIASTOLIC_2' },
  { item_id: 'H010020', item_name: '脈拍', category: '血圧', subcategory: '測定', data_type: 'integer', unit: '/分', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'PULSE' },

  // 眼科
  { item_id: 'H040001', item_name: '視力 裸眼 右', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'VISION_NAKED_R' },
  { item_id: 'H040002', item_name: '視力 裸眼 左', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'VISION_NAKED_L' },
  { item_id: 'H040003', item_name: '視力 矯正 右', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'VISION_CORRECTED_R' },
  { item_id: 'H040004', item_name: '視力 矯正 左', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'VISION_CORRECTED_L' },
  { item_id: 'H040010', item_name: '眼圧 右', category: '眼科', subcategory: '眼圧', data_type: 'decimal', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'IOP_R' },
  { item_id: 'H040011', item_name: '眼圧 左', category: '眼科', subcategory: '眼圧', data_type: 'decimal', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'IOP_L' },
  { item_id: 'H040020', item_name: '眼底 右', category: '眼科', subcategory: '眼底', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'FUNDUS_R' },
  { item_id: 'H040021', item_name: '眼底 左', category: '眼科', subcategory: '眼底', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'FUNDUS_L' },

  // 聴力
  { item_id: 'H030001', item_name: '聴力右 1000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'HEARING_R_1000' },
  { item_id: 'H030002', item_name: '聴力左 1000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'HEARING_L_1000' },
  { item_id: 'H030003', item_name: '聴力右 4000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'HEARING_R_4000' },
  { item_id: 'H030004', item_name: '聴力左 4000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false, legacy_key: 'HEARING_L_4000' },

  // 尿検査（BMLコード）
  { item_id: '0000055', item_name: '尿糖(定性)', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'URINE_GLUCOSE' },
  { item_id: '0000051', item_name: '尿蛋白', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'URINE_PROTEIN' },
  { item_id: '0000063', item_name: '尿潜血', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_OCCULT_BLOOD' },
  { item_id: '0000057', item_name: 'ウロビリノーゲン', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'UROBILINOGEN' },
  { item_id: '0000061', item_name: '尿PH', category: '尿検査', subcategory: '尿一般', data_type: 'decimal', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_PH' },
  { item_id: '0000059', item_name: '尿ビリルビン', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_BILIRUBIN' },
  { item_id: '0000062', item_name: 'アセトン体', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_KETONE' },
  { item_id: '0000060', item_name: '尿比重', category: '尿検査', subcategory: '尿一般', data_type: 'decimal', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_SG' },
  { item_id: 'H050001', item_name: '尿沈渣白血球', category: '尿検査', subcategory: '尿沈渣', data_type: 'text', unit: '/HPF', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_WBC' },
  { item_id: 'H050002', item_name: '尿沈渣赤血球', category: '尿検査', subcategory: '尿沈渣', data_type: 'text', unit: '/HPF', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_RBC' },
  { item_id: 'H050003', item_name: '尿沈渣扁平上皮', category: '尿検査', subcategory: '尿沈渣', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_EPITHELIAL' },
  { item_id: 'H050004', item_name: '尿沈渣細菌', category: '尿検査', subcategory: '尿沈渣', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'URINE_BACTERIA' },

  // 便検査
  { item_id: 'H060001', item_name: '便ヘモグロビン1回目', category: '便検査', subcategory: '潜血', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'FOBT_1' },
  { item_id: 'H060002', item_name: '便ヘモグロビン2回目', category: '便検査', subcategory: '潜血', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'FOBT_2' },

  // 血液学検査（BMLコード）
  { item_id: '0000301', item_name: '白血球数', category: '血液学検査', subcategory: '血球', data_type: 'integer', unit: '/μl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'WBC' },
  { item_id: '0000302', item_name: '赤血球数', category: '血液学検査', subcategory: '血球', data_type: 'integer', unit: '万/μl', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'RBC' },
  { item_id: '0000303', item_name: '血色素量(ヘモグロビン)', category: '血液学検査', subcategory: '血球', data_type: 'decimal', unit: 'g/dl', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'HEMOGLOBIN' },
  { item_id: '0000304', item_name: 'ヘマトクリット', category: '血液学検査', subcategory: '血球', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'HEMATOCRIT' },
  { item_id: '0000308', item_name: '血小板(PLT)', category: '血液学検査', subcategory: '血球', data_type: 'decimal', unit: '万/μl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'PLATELET' },
  { item_id: '0000305', item_name: 'MCV', category: '血液学検査', subcategory: '血球指数', data_type: 'decimal', unit: 'fl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'MCV' },
  { item_id: '0000306', item_name: 'MCH', category: '血液学検査', subcategory: '血球指数', data_type: 'decimal', unit: 'pg', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'MCH' },
  { item_id: '0000307', item_name: 'MCHC', category: '血液学検査', subcategory: '血球指数', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'MCHC' },
  { item_id: '0001889', item_name: 'NEUT（好中球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'NEUT' },
  { item_id: '0001885', item_name: 'LYMPHO（リンパ球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'LYMPHO' },
  { item_id: '0001886', item_name: 'MONO（単球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'MONO' },
  { item_id: '0001882', item_name: 'EOS（好酸球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'EOS' },
  { item_id: '0001881', item_name: 'BASO（好塩基球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'BASO' },

  // 肝胆膵機能（BMLコード）
  { item_id: '0000401', item_name: '総蛋白(TP)', category: '肝胆膵機能', subcategory: '蛋白', data_type: 'decimal', unit: 'g/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'TOTAL_PROTEIN' },
  { item_id: '0000417', item_name: 'アルブミン', category: '肝胆膵機能', subcategory: '蛋白', data_type: 'decimal', unit: 'g/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'ALBUMIN' },
  { item_id: '0000481', item_name: 'AST(GOT)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'AST' },
  { item_id: '0000482', item_name: 'ALT(GPT)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'ALT' },
  { item_id: '0000484', item_name: 'γ-GTP', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'GGT' },
  { item_id: '0013067', item_name: 'ALP(IFCC)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'ALP' },
  { item_id: '0013380', item_name: 'LDH(乳酸脱水素酵素)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'LDH' },
  { item_id: '0000491', item_name: 'コリンエステラーゼ（Ch-E）', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CHE' },
  { item_id: '0000501', item_name: '血清アミラーゼ', category: '肝胆膵機能', subcategory: '膵', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'AMYLASE' },
  { item_id: '0000472', item_name: '総ビリルビン（T-Bil）', category: '肝胆膵機能', subcategory: '胆', data_type: 'decimal', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'T_BIL' },

  // 脂質検査（BMLコード）
  { item_id: '0000453', item_name: '総コレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'TOTAL_CHOLESTEROL' },
  { item_id: '0000454', item_name: '中性脂肪（TG）', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'TG' },
  { item_id: '0000460', item_name: 'HDLコレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'HDL_C' },
  { item_id: '0000410', item_name: 'LDLコレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'LDL_C' },
  { item_id: '0003845', item_name: 'non HDLコレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'NON_HDL_C' },

  // 糖代謝（BMLコード）
  { item_id: '0000503', item_name: '空腹時血糖', category: '糖代謝', subcategory: '血糖', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'FBS' },
  { item_id: '0003317', item_name: 'HbA1c', category: '糖代謝', subcategory: '血糖', data_type: 'decimal', unit: '%', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'HBA1C' },

  // 腎機能（BMLコード）
  { item_id: '0000413', item_name: 'クレアチニン', category: '腎機能', subcategory: '腎機能', data_type: 'decimal', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true, legacy_key: 'CREATININE' },
  { item_id: '0000409', item_name: '尿素窒素(BUN)', category: '腎機能', subcategory: '腎機能', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'BUN' },
  { item_id: '0002696', item_name: 'eGFR', category: '腎機能', subcategory: '腎機能', data_type: 'integer', unit: 'ml/分/1.73m²', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'EGFR' },

  // 生化学検査（BMLコード）
  { item_id: '0000407', item_name: '尿酸（UA）', category: '生化学検査', subcategory: '代謝', data_type: 'decimal', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'UA' },
  { item_id: '0000497', item_name: 'クレアチニンキナーゼ(CK)', category: '生化学検査', subcategory: '酵素', data_type: 'integer', unit: 'U/l', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CK' },
  { item_id: '0000421', item_name: 'ナトリウム(Na)', category: '生化学検査', subcategory: '電解質', data_type: 'integer', unit: 'mEq/I', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'NA' },
  { item_id: '0000423', item_name: 'カリウム(K)', category: '生化学検査', subcategory: '電解質', data_type: 'decimal', unit: 'mEq/I', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'K' },
  { item_id: '0000425', item_name: 'クロール(Cl)', category: '生化学検査', subcategory: '電解質', data_type: 'integer', unit: 'mEq/I', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CL' },
  { item_id: '0000427', item_name: 'カルシウム(Ca)', category: '生化学検査', subcategory: '電解質', data_type: 'decimal', unit: 'mg/dl', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CA' },
  { item_id: '0000435', item_name: '血清鉄(Fe)', category: '生化学検査', subcategory: '鉄代謝', data_type: 'integer', unit: 'μg/dl', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'FE' },
  { item_id: '0000437', item_name: '総鉄結合能(TIBC)', category: '生化学検査', subcategory: '鉄代謝', data_type: 'integer', unit: 'μg/dl', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'TIBC' },
  { item_id: '0002490', item_name: 'リウマトイド因子(RF)', category: '生化学検査', subcategory: '免疫', data_type: 'decimal', unit: 'IU/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'RF' },

  // 血清学検査（BMLコード）
  { item_id: '0000658', item_name: 'CRP定量', category: '血清学検査', subcategory: '炎症', data_type: 'decimal', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'CRP' },

  // 凝固検査（BMLコード）
  { item_id: '0000331', item_name: 'プロトロンビン時間(PT)', category: '凝固検査', subcategory: '凝固', data_type: 'decimal', unit: '秒', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'PT' },
  { item_id: '0000334', item_name: 'APTT', category: '凝固検査', subcategory: '凝固', data_type: 'decimal', unit: '秒', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'APTT' },

  // 腫瘍マーカー（医療機関独自）
  { item_id: 'H080001', item_name: 'PSA', category: '腫瘍マーカー', subcategory: '前立腺', data_type: 'decimal', unit: 'ng/ml', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'PSA' },
  { item_id: 'H080002', item_name: 'CEA', category: '腫瘍マーカー', subcategory: '消化器', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CEA' },
  { item_id: 'H080003', item_name: 'CA19-9', category: '腫瘍マーカー', subcategory: '消化器', data_type: 'decimal', unit: 'U/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CA19_9' },
  { item_id: 'H080004', item_name: 'CA125', category: '腫瘍マーカー', subcategory: '婦人科', data_type: 'decimal', unit: 'U/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CA125' },
  { item_id: 'H080005', item_name: 'AFP', category: '腫瘍マーカー', subcategory: '肝', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'AFP' },
  { item_id: 'H080006', item_name: 'NSE', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'NSE' },
  { item_id: 'H080007', item_name: 'CYFRA21-1', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'CYFRA21_1' },
  { item_id: 'H080008', item_name: 'SCC', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'SCC' },
  { item_id: 'H080009', item_name: 'ProGRP', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'pg/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'PROGRP' },
  { item_id: 'H080010', item_name: 'PIVKA II', category: '腫瘍マーカー', subcategory: '肝', data_type: 'decimal', unit: 'mAU/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'PIVKA2' },
  { item_id: 'H080011', item_name: 'エラスターゼ1', category: '腫瘍マーカー', subcategory: '膵', data_type: 'integer', unit: 'ng/dl', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'ELASTASE1' },
  { item_id: 'H080012', item_name: '抗p53抗体', category: '腫瘍マーカー', subcategory: 'その他', data_type: 'decimal', unit: 'U/ml', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'P53' },

  // 感染症検査（BMLコード）
  { item_id: '0000911', item_name: 'TPHA', category: '感染症検査', subcategory: '梅毒', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'TPHA' },
  { item_id: '0000905', item_name: 'RPR定性', category: '感染症検査', subcategory: '梅毒', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'RPR' },
  { item_id: '0000740', item_name: 'HBs抗原', category: '感染症検査', subcategory: 'B型肝炎', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'HBS_AG' },
  { item_id: 'H100001', item_name: 'HBs抗体', category: '感染症検査', subcategory: 'B型肝炎', data_type: 'select', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'HBS_AB' },
  { item_id: '0003795', item_name: 'HCV抗体', category: '感染症検査', subcategory: 'C型肝炎', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'HCV_AB' },
  { item_id: 'H100002', item_name: 'HIV-1抗体', category: '感染症検査', subcategory: 'HIV', data_type: 'select', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'HIV_AB' },
  { item_id: '0013413', item_name: 'ピロリ抗体(H.ピロリ抗体)', category: '感染症検査', subcategory: 'ピロリ', data_type: 'select', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'HP_AB' },
  { item_id: '0003294', item_name: '尿素呼気試験', category: '感染症検査', subcategory: 'ピロリ', data_type: 'select', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'UREA_BREATH_TEST' },

  // 細胞診（BMLコード）
  { item_id: '0005972', item_name: '喀痰細胞診', category: '細胞診', subcategory: '呼吸器', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'SPUTUM_CYTOLOGY' },

  // 内分泌・甲状腺
  { item_id: '0003550', item_name: 'NT-proBNP', category: '内分泌・甲状腺', subcategory: '心臓', data_type: 'decimal', unit: 'pg/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'NT_PROBNP' },
  { item_id: 'H110001', item_name: 'FT3', category: '内分泌・甲状腺', subcategory: '甲状腺', data_type: 'decimal', unit: 'pg/ml', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'FT3' },
  { item_id: 'H110002', item_name: 'FT4', category: '内分泌・甲状腺', subcategory: '甲状腺', data_type: 'decimal', unit: 'ng/dl', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'FT4' },
  { item_id: 'H110003', item_name: 'TSH', category: '内分泌・甲状腺', subcategory: '甲状腺', data_type: 'decimal', unit: 'μIU/mL', required_dock: false, required_regular: false, required_secondary: false, legacy_key: 'TSH' },

  // 血液型（医療機関独自）
  { item_id: 'H120001', item_name: '血液型 ABO式', category: '血液型', subcategory: '血液型', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'BLOOD_TYPE_ABO' },
  { item_id: 'H120002', item_name: '血液型 Rho・D', category: '血液型', subcategory: '血液型', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'BLOOD_TYPE_RH' },

  // 肺機能（医療機関独自）
  { item_id: 'H070001', item_name: '肺活量', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: 'L', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'VC' },
  { item_id: 'H070002', item_name: '１秒量', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: 'L', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'FEV1' },
  { item_id: 'H070003', item_name: '％肺活量', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'PERCENT_VC' },
  { item_id: 'H070004', item_name: '１秒率', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false, legacy_key: 'FEV1_PERCENT' }
];

// ============================================
// 判定基準マスタデータ（人間ドック学会2025年度版）
// ============================================

const JUDGMENT_CRITERIA_DATA = [
  { item_id: 'H010004', item_name: 'BMI', gender: '共通', unit: '', a_min: 18.5, a_max: 24.9, b_min: null, b_max: null, c_min: 25.0, c_max: 29.9, d_min: 30.0, d_max: null, note: '', legacy_key: 'BMI' },
  { item_id: 'H010006', item_name: '腹囲(男)', gender: 'M', unit: 'cm', a_min: 0, a_max: 84.9, b_min: null, b_max: null, c_min: 85.0, c_max: 99.9, d_min: 100.0, d_max: null, note: 'メタボ基準', legacy_key: 'WAIST_M' },
  { item_id: 'H010007', item_name: '腹囲(女)', gender: 'F', unit: 'cm', a_min: 0, a_max: 89.9, b_min: null, b_max: null, c_min: 90.0, c_max: 99.9, d_min: 100.0, d_max: null, note: 'メタボ基準', legacy_key: 'WAIST_F' },
  { item_id: 'H010010', item_name: '血圧(収縮)', gender: '共通', unit: 'mmHg', a_min: 0, a_max: 129, b_min: 130, b_max: 139, c_min: 140, c_max: 159, d_min: 160, d_max: null, note: '', legacy_key: 'BP_SYSTOLIC' },
  { item_id: 'H010011', item_name: '血圧(拡張)', gender: '共通', unit: 'mmHg', a_min: 0, a_max: 84, b_min: 85, b_max: 89, c_min: 90, c_max: 99, d_min: 100, d_max: null, note: '', legacy_key: 'BP_DIASTOLIC' },
  { item_id: '0000503', item_name: '空腹時血糖', gender: '共通', unit: 'mg/dl', a_min: 0, a_max: 99, b_min: 100, b_max: 109, c_min: 110, c_max: 125, d_min: 126, d_max: null, note: 'HbA1cと組合せ判定', legacy_key: 'FBS' },
  { item_id: '0003317', item_name: 'HbA1c', gender: '共通', unit: '%', a_min: 0, a_max: 5.5, b_min: 5.6, b_max: 5.9, c_min: 6.0, c_max: 6.4, d_min: 6.5, d_max: null, note: 'FBSと組合せ判定', legacy_key: 'HBA1C' },
  { item_id: '0000460', item_name: 'HDLコレステロール', gender: '共通', unit: 'mg/dl', a_min: 40, a_max: 999, b_min: null, b_max: null, c_min: 35, c_max: 39, d_min: 0, d_max: 34, note: '低値が異常', legacy_key: 'HDL_C' },
  { item_id: '0000410', item_name: 'LDLコレステロール', gender: '共通', unit: 'mg/dl', a_min: 60, a_max: 119, b_min: 120, b_max: 139, c_min: 140, c_max: 179, d_min: 180, d_max: null, note: '', legacy_key: 'LDL_C' },
  { item_id: '0000454', item_name: '中性脂肪', gender: '共通', unit: 'mg/dl', a_min: 30, a_max: 149, b_min: 150, b_max: 299, c_min: 300, c_max: 499, d_min: 500, d_max: null, note: '空腹時', legacy_key: 'TG' },
  { item_id: '0000481', item_name: 'AST(GOT)', gender: '共通', unit: 'U/I', a_min: 0, a_max: 30, b_min: 31, b_max: 35, c_min: 36, c_max: 50, d_min: 51, d_max: null, note: '', legacy_key: 'AST' },
  { item_id: '0000482', item_name: 'ALT(GPT)', gender: '共通', unit: 'U/I', a_min: 0, a_max: 30, b_min: 31, b_max: 40, c_min: 41, c_max: 50, d_min: 51, d_max: null, note: '', legacy_key: 'ALT' },
  { item_id: '0000484', item_name: 'γ-GTP', gender: '共通', unit: 'U/I', a_min: 0, a_max: 50, b_min: 51, b_max: 80, c_min: 81, c_max: 100, d_min: 101, d_max: null, note: '', legacy_key: 'GGT' },
  { item_id: '0000413_M', item_name: 'クレアチニン(男)', gender: 'M', unit: 'mg/dl', a_min: 0.1, a_max: 1.0, b_min: 1.01, b_max: 1.09, c_min: 1.1, c_max: 1.29, d_min: 1.3, d_max: null, note: 'eGFR優先', legacy_key: 'CREATININE_M' },
  { item_id: '0000413_F', item_name: 'クレアチニン(女)', gender: 'F', unit: 'mg/dl', a_min: 0.1, a_max: 0.7, b_min: 0.71, b_max: 0.79, c_min: 0.8, c_max: 0.99, d_min: 1.0, d_max: null, note: 'eGFR優先', legacy_key: 'CREATININE_F' },
  { item_id: '0002696', item_name: 'eGFR', gender: '共通', unit: 'ml/分/1.73m²', a_min: 60.0, a_max: 999, b_min: 50.0, b_max: 59.9, c_min: 45.0, c_max: 49.9, d_min: 0, d_max: 44.9, note: '低値が異常', legacy_key: 'EGFR' },
  { item_id: '0000407', item_name: '尿酸', gender: '共通', unit: 'mg/dl', a_min: 2.1, a_max: 7.0, b_min: 7.1, b_max: 7.9, c_min: 8.0, c_max: 8.9, d_min: 9.0, d_max: null, note: '', legacy_key: 'UA' },
  { item_id: '0000303_M', item_name: '血色素量(男)', gender: 'M', unit: 'g/dl', a_min: 13.1, a_max: 16.3, b_min: 16.4, b_max: 18.0, c_min: 12.1, c_max: 13.0, d_min: 0, d_max: 12.0, note: '低値・高値両方', legacy_key: 'HEMOGLOBIN_M' },
  { item_id: '0000303_F', item_name: '血色素量(女)', gender: 'F', unit: 'g/dl', a_min: 12.1, a_max: 14.5, b_min: 14.6, b_max: 16.0, c_min: 11.1, c_max: 12.0, d_min: 0, d_max: 11.0, note: '低値・高値両方', legacy_key: 'HEMOGLOBIN_F' },
  { item_id: '0000308', item_name: '血小板', gender: '共通', unit: '万/μl', a_min: 14.5, a_max: 32.9, b_min: 12.3, b_max: 14.4, c_min: 10.0, c_max: 12.2, d_min: 0, d_max: 9.9, note: '', legacy_key: 'PLATELET' },
  { item_id: '0000401', item_name: '総蛋白', gender: '共通', unit: 'g/dl', a_min: 6.5, a_max: 7.9, b_min: 8.0, b_max: 8.3, c_min: 6.2, c_max: 6.4, d_min: 0, d_max: 6.1, note: '', legacy_key: 'TOTAL_PROTEIN' },
  { item_id: '0000417', item_name: 'アルブミン', gender: '共通', unit: 'g/dl', a_min: 3.9, a_max: 5.5, b_min: 3.7, b_max: 3.8, c_min: 3.0, c_max: 3.6, d_min: 0, d_max: 2.9, note: '', legacy_key: 'ALBUMIN' },
  { item_id: '0000658', item_name: 'CRP定量', gender: '共通', unit: 'mg/dl', a_min: 0, a_max: 0.30, b_min: 0.31, b_max: 0.99, c_min: 1.0, c_max: 2.99, d_min: 3.0, d_max: null, note: '炎症マーカー', legacy_key: 'CRP' }
];

// ============================================
// 選択肢マスタデータ（定性検査用）
// ============================================

const SELECT_OPTIONS_DATA = [
  { item_id: '0000051', options: '(-)|(±)|(+)|(++)|(+++)', description: '尿蛋白判定', legacy_key: 'URINE_PROTEIN' },
  { item_id: '0000055', options: '(-)|(±)|(+)|(++)|(+++)', description: '尿糖判定', legacy_key: 'URINE_GLUCOSE' },
  { item_id: '0000063', options: '(-)|(±)|(+)|(++)|(+++)', description: '尿潜血判定', legacy_key: 'URINE_OCCULT_BLOOD' },
  { item_id: '0000057', options: '(-)|(±)|(+)|(++)|(+++)', description: 'ウロビリノーゲン判定', legacy_key: 'UROBILINOGEN' },
  { item_id: '0000059', options: '(-)|(+)|(++)|(+++)', description: '尿ビリルビン判定', legacy_key: 'URINE_BILIRUBIN' },
  { item_id: '0000062', options: '(-)|(±)|(+)|(++)|(+++)', description: 'ケトン体判定', legacy_key: 'URINE_KETONE' },
  { item_id: 'H060001', options: '(-)|(+)', description: '便潜血1回目', legacy_key: 'FOBT_1' },
  { item_id: 'H060002', options: '(-)|(+)', description: '便潜血2回目', legacy_key: 'FOBT_2' },
  { item_id: 'H090007', options: 'M|F', description: '性別', legacy_key: 'SEX' },
  { item_id: 'H120001', options: 'A|B|O|AB', description: 'ABO血液型', legacy_key: 'BLOOD_TYPE_ABO' },
  { item_id: 'H120002', options: '(+)|(-)', description: 'Rh血液型', legacy_key: 'BLOOD_TYPE_RH' },
  { item_id: 'H030001', options: '異常なし|所見あり', description: '聴力判定', legacy_key: 'HEARING_R_1000' },
  { item_id: 'H030002', options: '異常なし|所見あり', description: '聴力判定', legacy_key: 'HEARING_L_1000' },
  { item_id: 'H030003', options: '異常なし|所見あり', description: '聴力判定', legacy_key: 'HEARING_R_4000' },
  { item_id: 'H030004', options: '異常なし|所見あり', description: '聴力判定', legacy_key: 'HEARING_L_4000' },
  { item_id: '0000740', options: '(-)|(+)', description: 'HBs抗原', legacy_key: 'HBS_AG' },
  { item_id: 'H100001', options: '(-)|(+)', description: 'HBs抗体', legacy_key: 'HBS_AB' },
  { item_id: '0003795', options: '(-)|(+)', description: 'HCV抗体', legacy_key: 'HCV_AB' },
  { item_id: 'H100002', options: '(-)|(+)', description: 'HIV抗体', legacy_key: 'HIV_AB' },
  { item_id: '0000911', options: '(-)|(+)', description: '梅毒TPHA', legacy_key: 'TPHA' },
  { item_id: '0000905', options: '(-)|(+)', description: '梅毒RPR', legacy_key: 'RPR' },
  { item_id: 'H090020', options: 'A|B|C|D|E|F', description: '総合判定', legacy_key: 'OVERALL_JUDGMENT' }
];

// ============================================
// 健診コースマスタデータ
// ============================================

const EXAM_COURSE_MASTER_DATA = [
  {
    course_id: 'DOCK_LIFESTYLE',
    course_name: '生活習慣病ドック',
    price: 40000,
    description: '基本的な健康診断',
    item_count: 85,
    required_items: 'H010001,H010002,H010004,H010006,H010007,H010010,H010011,0000503,0003317,0000460,0000410,0000454,0000481,0000482,0000484'
  },
  {
    course_id: 'DOCK_DIGESTIVE',
    course_name: '消化器ドック',
    price: 60000,
    description: 'GS+CS、腫瘍マーカー',
    item_count: 95,
    required_items: 'H010001,H010002,H010004,H010006,H010007,H010010,H010011,0000503,0003317,0000460,0000410,0000454,0000481,0000482,0000484,H080002,H080003,H020030'
  },
  {
    course_id: 'DOCK_PREMIUM',
    course_name: '全身スクリーニング',
    price: 160000,
    description: '全身MRI(DWIBS)含む',
    item_count: 120,
    required_items: 'H010001,H010002,H010004,H010006,H010007,H010010,H010011,0000503,0003317,0000460,0000410,0000454,0000481,0000482,0000484,H080002,H080003,H080005,H080001,H020060,H020061'
  },
  {
    course_id: 'DOCK_CANCER',
    course_name: 'がんスクリーニング',
    price: 70000,
    description: 'CT、腫瘍マーカー多数',
    item_count: 100,
    required_items: 'H010001,H010002,H010004,H080002,H080003,H080005,H080001,H080004,H080006,H020050,H020051'
  },
  {
    course_id: 'REGULAR_CHECKUP',
    course_name: '定期健康診断',
    price: 8000,
    description: '労働安全衛生法準拠',
    item_count: 35,
    required_items: 'H010001,H010002,H010004,H010010,H010011,0000051,0000055,0000302,0000303,0000481,0000482,0000484,0000454,0000460,0000410,0000503,0003317'
  },
  {
    course_id: 'SECONDARY_EXAM',
    course_name: '労災二次健診',
    price: 0,
    description: '脳・心疾患リスク評価',
    item_count: 20,
    required_items: 'H010010,H010011,0000503,0003317,0000460,0000410,0000454,0000413,0000303,H020020'
  }
];

// ============================================
// データ取得用ヘルパー関数
// ============================================

/**
 * 検査項目マスタデータを配列形式で取得
 * @returns {Array} シート挿入用の2次元配列
 */
function getExamItemMasterRows() {
  return EXAM_ITEM_MASTER_DATA.map((item, index) => [
    item.item_id,
    item.item_name,
    item.category,
    item.subcategory,
    item.data_type,
    item.unit,
    item.required_dock,
    item.required_regular,
    item.required_secondary,
    index + 1
  ]);
}

/**
 * 判定基準マスタデータを配列形式で取得
 * @returns {Array} シート挿入用の2次元配列
 */
function getJudgmentCriteriaRows() {
  return JUDGMENT_CRITERIA_DATA.map(item => [
    item.item_id,
    item.item_name,
    item.gender,
    item.unit,
    item.a_min,
    item.a_max,
    item.b_min,
    item.b_max,
    item.c_min,
    item.c_max,
    item.d_min,
    item.d_max,
    item.note
  ]);
}

/**
 * 選択肢マスタデータを配列形式で取得
 * @returns {Array} シート挿入用の2次元配列
 */
function getSelectOptionsRows() {
  return SELECT_OPTIONS_DATA.map(item => [
    item.item_id,
    item.options,
    item.description
  ]);
}

/**
 * 健診コースマスタデータを配列形式で取得
 * @returns {Array} シート挿入用の2次元配列
 */
function getExamCourseMasterRows() {
  return EXAM_COURSE_MASTER_DATA.map(item => [
    item.course_id,
    item.course_name,
    item.price,
    item.description,
    item.item_count,
    item.required_items
  ]);
}

/**
 * 項目IDからマスタデータを取得（legacy_keyにも対応）
 * @param {string} itemId - 項目ID（統一ID または 旧英語キー）
 * @returns {Object|null} マスタデータ
 */
function getExamItemById(itemId) {
  // まず統一IDで検索
  let item = EXAM_ITEM_MASTER_DATA.find(item => item.item_id === itemId);
  if (item) return item;

  // 見つからなければlegacy_keyで検索
  item = EXAM_ITEM_MASTER_DATA.find(item => item.legacy_key === itemId);
  if (item) return item;

  // それでも見つからなければ、LEGACY_KEY_TO_UNIFIED_IDで変換を試みる
  const unifiedId = LEGACY_KEY_TO_UNIFIED_ID[itemId];
  if (unifiedId) {
    return EXAM_ITEM_MASTER_DATA.find(item => item.item_id === unifiedId) || null;
  }

  return null;
}

/**
 * legacyキーを統一IDに変換
 * @param {string} legacyKey - 旧英語キー
 * @returns {string} 統一ID（変換できない場合は元の値を返す）
 */
function convertToUnifiedId(legacyKey) {
  return LEGACY_KEY_TO_UNIFIED_ID[legacyKey] || legacyKey;
}

/**
 * 統一IDからlegacyキーを取得
 * @param {string} unifiedId - 統一ID
 * @returns {string|null} 旧英語キー
 */
function getLegacyKeyFromUnifiedId(unifiedId) {
  const item = EXAM_ITEM_MASTER_DATA.find(item => item.item_id === unifiedId);
  return item ? (item.legacy_key || null) : null;
}

/**
 * 判定基準を取得（性別対応）
 * @param {string} itemId - 項目ID
 * @param {string} gender - 性別（M/F）
 * @returns {Object|null} 判定基準データ
 */
function getJudgmentCriteria(itemId, gender) {
  // まず性別特有の基準を探す
  let criteria = JUDGMENT_CRITERIA_DATA.find(
    c => c.item_id === itemId + '_' + gender
  );

  // なければ共通の基準を探す
  if (!criteria) {
    criteria = JUDGMENT_CRITERIA_DATA.find(
      c => c.item_id === itemId && c.gender === '共通'
    );
  }

  return criteria || null;
}

/**
 * コース別必須項目リストを取得
 * @param {string} courseId - コースID
 * @returns {Array} 必須項目IDの配列
 */
function getRequiredItemsByCourse(courseId) {
  const course = EXAM_COURSE_MASTER_DATA.find(c => c.course_id === courseId);
  if (!course) return [];
  return course.required_items.split(',');
}
