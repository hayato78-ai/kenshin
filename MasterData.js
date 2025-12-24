/**
 * MasterData.gs - 検査項目マスタデータ定義
 * 検査項目マスタ.json / GAS用設定データ.json から変換
 *
 * @version 1.0.0
 * @date 2025-12-17
 */

// ============================================
// 検査項目マスタデータ（150項目）
// ============================================

const EXAM_ITEM_MASTER_DATA = [
  // 基本情報 - 受診者情報
  { item_id: 'COMPANY', item_name: '事業所', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'DEPARTMENT', item_name: '所属', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'NAME_KANA', item_name: 'カナ氏名', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'NAME', item_name: '氏名', category: '基本情報', subcategory: '受診者情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'BIRTHDATE', item_name: '生年月日', category: '基本情報', subcategory: '受診者情報', data_type: 'date', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'AGE', item_name: '年齢', category: '基本情報', subcategory: '受診者情報', data_type: 'integer', unit: '才', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'SEX', item_name: '性別', category: '基本情報', subcategory: '受診者情報', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true },

  // 基本情報 - 受診情報
  { item_id: 'EXAM_DATE', item_name: '受診日', category: '基本情報', subcategory: '受診情報', data_type: 'date', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'COURSE', item_name: '受診コース', category: '基本情報', subcategory: '受診情報', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'RECEPTION_NO', item_name: '受診No', category: '基本情報', subcategory: '受診情報', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true },

  // 基本情報 - 判定
  { item_id: 'OVERALL_JUDGMENT', item_name: '総合判定', category: '基本情報', subcategory: '判定', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'OVERALL_COMMENT', item_name: '総合所見', category: '基本情報', subcategory: '判定', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false },

  // 基本情報 - 問診
  { item_id: 'MEDICAL_HISTORY', item_name: '既往歴', category: '基本情報', subcategory: '問診', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'SUBJECTIVE_SYMPTOMS', item_name: '自覚症状', category: '基本情報', subcategory: '問診', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'OBJECTIVE_SYMPTOMS', item_name: '他覚症状', category: '基本情報', subcategory: '問診', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true },

  // 画像診断
  { item_id: 'PHYSICAL_EXAM', item_name: '診察所見', category: '画像診断', subcategory: '診察', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'CHEST_XRAY', item_name: '胸部X線', category: '画像診断', subcategory: 'X線', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'ABDOMINAL_XRAY', item_name: '腹部X線', category: '画像診断', subcategory: 'X線', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'ECG', item_name: '心電図', category: '画像診断', subcategory: '心電図', data_type: 'text', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'UPPER_GI_ENDOSCOPY', item_name: '上部消化管内視鏡検査(胃カメラ)', category: '画像診断', subcategory: '消化器', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'LOWER_GI_ENDOSCOPY', item_name: '下部消化管内視鏡検査(大腸カメラ)', category: '画像診断', subcategory: '大腸', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'ABDOMINAL_US', item_name: '腹部超音波', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'CAROTID_US', item_name: '頸部超音波', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CARDIAC_US', item_name: '心臓超音波検査', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'THYROID_US', item_name: '甲状腺超音波検査', category: '画像診断', subcategory: '超音波検査', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CHEST_CT', item_name: '胸部CT', category: '画像診断', subcategory: 'CT', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'ABDOMINAL_CT', item_name: '腹部CT', category: '画像診断', subcategory: 'CT', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'BRAIN_MRI', item_name: '脳MRI', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'BRAIN_MRA', item_name: '脳MRA', category: '画像診断', subcategory: 'MRI', data_type: 'text', unit: '', required_dock: false, required_regular: false, required_secondary: false },

  // 身体測定
  { item_id: 'HEIGHT', item_name: '身長', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'cm', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'WEIGHT', item_name: '体重', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'kg', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'STANDARD_WEIGHT', item_name: '標準体重', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'kg', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'BMI', item_name: 'BMI', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'BODY_FAT', item_name: '体脂肪率', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'WAIST_M', item_name: '腹囲(男性)', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'cm', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'WAIST_F', item_name: '腹囲(女性)', category: '身体測定', subcategory: '計測', data_type: 'decimal', unit: 'cm', required_dock: true, required_regular: true, required_secondary: true },

  // 血圧
  { item_id: 'BP_SYSTOLIC_1', item_name: '血圧1回目（最高）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'BP_DIASTOLIC_1', item_name: '血圧1回目（最低）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'BP_SYSTOLIC_2', item_name: '血圧2回目（最高）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'BP_DIASTOLIC_2', item_name: '血圧2回目（最低）', category: '血圧', subcategory: '測定', data_type: 'integer', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'PULSE', item_name: '脈拍', category: '血圧', subcategory: '測定', data_type: 'integer', unit: '/分', required_dock: true, required_regular: false, required_secondary: false },

  // 眼科
  { item_id: 'VISION_NAKED_R', item_name: '視力 裸眼 右', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'VISION_NAKED_L', item_name: '視力 裸眼 左', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'VISION_CORRECTED_R', item_name: '視力 矯正 右', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'VISION_CORRECTED_L', item_name: '視力 矯正 左', category: '眼科', subcategory: '視力', data_type: 'decimal', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'IOP_R', item_name: '眼圧 右', category: '眼科', subcategory: '眼圧', data_type: 'decimal', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'IOP_L', item_name: '眼圧 左', category: '眼科', subcategory: '眼圧', data_type: 'decimal', unit: 'mmHg', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'FUNDUS_R', item_name: '眼底 右', category: '眼科', subcategory: '眼底', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'FUNDUS_L', item_name: '眼底 左', category: '眼科', subcategory: '眼底', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false },

  // 聴力
  { item_id: 'HEARING_R_1000', item_name: '聴力右 1000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'HEARING_L_1000', item_name: '聴力左 1000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'HEARING_R_4000', item_name: '聴力右 4000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false },
  { item_id: 'HEARING_L_4000', item_name: '聴力左 4000Hz', category: '聴力', subcategory: 'オージオ', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: false },

  // 尿検査
  { item_id: 'URINE_GLUCOSE', item_name: '尿糖(定性)', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'URINE_PROTEIN', item_name: '尿蛋白', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'URINE_OCCULT_BLOOD', item_name: '尿潜血', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'UROBILINOGEN', item_name: 'ウロビリノーゲン', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_PH', item_name: '尿PH', category: '尿検査', subcategory: '尿一般', data_type: 'decimal', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_BILIRUBIN', item_name: '尿ビリルビン', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_KETONE', item_name: 'アセトン体', category: '尿検査', subcategory: '尿一般', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_SG', item_name: '尿比重', category: '尿検査', subcategory: '尿一般', data_type: 'decimal', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_WBC', item_name: '尿沈渣白血球', category: '尿検査', subcategory: '尿沈渣', data_type: 'text', unit: '/HPF', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_RBC', item_name: '尿沈渣赤血球', category: '尿検査', subcategory: '尿沈渣', data_type: 'text', unit: '/HPF', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_EPITHELIAL', item_name: '尿沈渣扁平上皮', category: '尿検査', subcategory: '尿沈渣', data_type: 'text', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'URINE_BACTERIA', item_name: '尿沈渣細菌', category: '尿検査', subcategory: '尿沈渣', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },

  // 便検査
  { item_id: 'FOBT_1', item_name: '便ヘモグロビン1回目', category: '便検査', subcategory: '潜血', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'FOBT_2', item_name: '便ヘモグロビン2回目', category: '便検査', subcategory: '潜血', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },

  // 血液学検査
  { item_id: 'WBC', item_name: '白血球数', category: '血液学検査', subcategory: '血球', data_type: 'integer', unit: '/μl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'RBC', item_name: '赤血球数', category: '血液学検査', subcategory: '血球', data_type: 'integer', unit: '万/μl', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'HEMOGLOBIN', item_name: '血色素量(ヘモグロビン)', category: '血液学検査', subcategory: '血球', data_type: 'decimal', unit: 'g/dl', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'HEMATOCRIT', item_name: 'ヘマトクリット', category: '血液学検査', subcategory: '血球', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'PLATELET', item_name: '血小板(PLT)', category: '血液学検査', subcategory: '血球', data_type: 'decimal', unit: '万/μl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'MCV', item_name: 'MCV', category: '血液学検査', subcategory: '血球指数', data_type: 'decimal', unit: 'fl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'MCH', item_name: 'MCH', category: '血液学検査', subcategory: '血球指数', data_type: 'decimal', unit: 'pg', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'MCHC', item_name: 'MCHC', category: '血液学検査', subcategory: '血球指数', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'NEUT', item_name: 'NEUT（好中球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'LYMPHO', item_name: 'LYMPHO（リンパ球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'MONO', item_name: 'MONO（単球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'EOS', item_name: 'EOS（好酸球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'BASO', item_name: 'BASO（好塩基球）', category: '血液学検査', subcategory: '白血球像', data_type: 'decimal', unit: '%', required_dock: false, required_regular: false, required_secondary: false },

  // 肝胆膵機能
  { item_id: 'TOTAL_PROTEIN', item_name: '総蛋白(TP)', category: '肝胆膵機能', subcategory: '蛋白', data_type: 'decimal', unit: 'g/dl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'ALBUMIN', item_name: 'アルブミン', category: '肝胆膵機能', subcategory: '蛋白', data_type: 'decimal', unit: 'g/dl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'AST', item_name: 'AST(GOT)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'ALT', item_name: 'ALT(GPT)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'GGT', item_name: 'γ-GTP', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'ALP', item_name: 'ALP(IFCC)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'LDH', item_name: 'LDH(乳酸脱水素酵素)', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'CHE', item_name: 'コリンエステラーゼ（Ch-E）', category: '肝胆膵機能', subcategory: '肝酵素', data_type: 'integer', unit: 'U/I', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'AMYLASE', item_name: '血清アミラーゼ', category: '肝胆膵機能', subcategory: '膵', data_type: 'integer', unit: 'U/I', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'T_BIL', item_name: '総ビリルビン（T-Bil）', category: '肝胆膵機能', subcategory: '胆', data_type: 'decimal', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false },

  // 脂質検査
  { item_id: 'TOTAL_CHOLESTEROL', item_name: '総コレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'TG', item_name: '中性脂肪（TG）', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'HDL_C', item_name: 'HDLコレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'LDL_C', item_name: 'LDLコレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'NON_HDL_C', item_name: 'non HDLコレステロール', category: '脂質検査', subcategory: 'コレステロール', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false },

  // 糖代謝
  { item_id: 'FBS', item_name: '空腹時血糖', category: '糖代謝', subcategory: '血糖', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'HBA1C', item_name: 'HbA1c', category: '糖代謝', subcategory: '血糖', data_type: 'decimal', unit: '%', required_dock: true, required_regular: true, required_secondary: true },

  // 腎機能
  { item_id: 'CREATININE', item_name: 'クレアチニン', category: '腎機能', subcategory: '腎機能', data_type: 'decimal', unit: 'mg/dl', required_dock: true, required_regular: true, required_secondary: true },
  { item_id: 'BUN', item_name: '尿素窒素(BUN)', category: '腎機能', subcategory: '腎機能', data_type: 'integer', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'EGFR', item_name: 'eGFR', category: '腎機能', subcategory: '腎機能', data_type: 'integer', unit: 'ml/分/1.73m²', required_dock: true, required_regular: false, required_secondary: false },

  // 生化学検査
  { item_id: 'UA', item_name: '尿酸（UA）', category: '生化学検査', subcategory: '代謝', data_type: 'decimal', unit: 'mg/dl', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'CK', item_name: 'クレアチニンキナーゼ(CK)', category: '生化学検査', subcategory: '酵素', data_type: 'integer', unit: 'U/l', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'NA', item_name: 'ナトリウム(Na)', category: '生化学検査', subcategory: '電解質', data_type: 'integer', unit: 'mEq/I', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'K', item_name: 'カリウム(K)', category: '生化学検査', subcategory: '電解質', data_type: 'decimal', unit: 'mEq/I', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CL', item_name: 'クロール(Cl)', category: '生化学検査', subcategory: '電解質', data_type: 'integer', unit: 'mEq/I', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CA', item_name: 'カルシウム(Ca)', category: '生化学検査', subcategory: '電解質', data_type: 'decimal', unit: 'mg/dl', required_dock: false, required_regular: false, required_secondary: false },

  // 腫瘍マーカー
  { item_id: 'PSA', item_name: 'PSA', category: '腫瘍マーカー', subcategory: '前立腺', data_type: 'decimal', unit: 'ng/ml', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CEA', item_name: 'CEA', category: '腫瘍マーカー', subcategory: '消化器', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CA19_9', item_name: 'CA19-9', category: '腫瘍マーカー', subcategory: '消化器', data_type: 'decimal', unit: 'U/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CA125', item_name: 'CA125', category: '腫瘍マーカー', subcategory: '婦人科', data_type: 'decimal', unit: 'U/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'AFP', item_name: 'AFP', category: '腫瘍マーカー', subcategory: '肝', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'NSE', item_name: 'NSE', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'CYFRA21_1', item_name: 'CYFRA21-1', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'SCC', item_name: 'SCC', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'ng/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'PROGRP', item_name: 'ProGRP', category: '腫瘍マーカー', subcategory: '肺', data_type: 'decimal', unit: 'pg/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'PIVKA2', item_name: 'PIVKA II', category: '腫瘍マーカー', subcategory: '肝', data_type: 'decimal', unit: 'mAU/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'ELASTASE1', item_name: 'エラスターゼ1', category: '腫瘍マーカー', subcategory: '膵', data_type: 'integer', unit: 'ng/dl', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'P53', item_name: '抗p53抗体', category: '腫瘍マーカー', subcategory: 'その他', data_type: 'decimal', unit: 'U/ml', required_dock: false, required_regular: false, required_secondary: false },

  // 感染症検査
  { item_id: 'TPHA', item_name: 'TPHA', category: '感染症検査', subcategory: '梅毒', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'RPR', item_name: 'RPR定性', category: '感染症検査', subcategory: '梅毒', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'HBS_AG', item_name: 'HBs抗原', category: '感染症検査', subcategory: 'B型肝炎', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'HBS_AB', item_name: 'HBs抗体', category: '感染症検査', subcategory: 'B型肝炎', data_type: 'select', unit: '', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'HCV_AB', item_name: 'HCV抗体', category: '感染症検査', subcategory: 'C型肝炎', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'HIV_AB', item_name: 'HIV-1抗体', category: '感染症検査', subcategory: 'HIV', data_type: 'select', unit: '', required_dock: false, required_regular: false, required_secondary: false },

  // 内分泌・甲状腺
  { item_id: 'NT_PROBNP', item_name: 'NT-proBNP', category: '内分泌・甲状腺', subcategory: '心臓', data_type: 'decimal', unit: 'pg/mL', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'FT3', item_name: 'FT3', category: '内分泌・甲状腺', subcategory: '甲状腺', data_type: 'decimal', unit: 'pg/ml', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'FT4', item_name: 'FT4', category: '内分泌・甲状腺', subcategory: '甲状腺', data_type: 'decimal', unit: 'ng/dl', required_dock: false, required_regular: false, required_secondary: false },
  { item_id: 'TSH', item_name: 'TSH', category: '内分泌・甲状腺', subcategory: '甲状腺', data_type: 'decimal', unit: 'μIU/mL', required_dock: false, required_regular: false, required_secondary: false },

  // 血液型
  { item_id: 'BLOOD_TYPE_ABO', item_name: '血液型 ABO式', category: '血液型', subcategory: '血液型', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'BLOOD_TYPE_RH', item_name: '血液型 Rho・D', category: '血液型', subcategory: '血液型', data_type: 'select', unit: '', required_dock: true, required_regular: false, required_secondary: false },

  // 肺機能
  { item_id: 'VC', item_name: '肺活量', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: 'L', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'FEV1', item_name: '１秒量', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: 'L', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'PERCENT_VC', item_name: '％肺活量', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false },
  { item_id: 'FEV1_PERCENT', item_name: '１秒率', category: '肺機能', subcategory: 'スパイロ', data_type: 'decimal', unit: '%', required_dock: true, required_regular: false, required_secondary: false }
];

// ============================================
// 判定基準マスタデータ（人間ドック学会2025年度版）
// ============================================

const JUDGMENT_CRITERIA_DATA = [
  { item_id: 'BMI', item_name: 'BMI', gender: '共通', unit: '', a_min: 18.5, a_max: 24.9, b_min: null, b_max: null, c_min: 25.0, c_max: 29.9, d_min: 30.0, d_max: null, note: '' },
  { item_id: 'WAIST_M', item_name: '腹囲(男)', gender: 'M', unit: 'cm', a_min: 0, a_max: 84.9, b_min: null, b_max: null, c_min: 85.0, c_max: 99.9, d_min: 100.0, d_max: null, note: 'メタボ基準' },
  { item_id: 'WAIST_F', item_name: '腹囲(女)', gender: 'F', unit: 'cm', a_min: 0, a_max: 89.9, b_min: null, b_max: null, c_min: 90.0, c_max: 99.9, d_min: 100.0, d_max: null, note: 'メタボ基準' },
  { item_id: 'BP_SYSTOLIC', item_name: '血圧(収縮)', gender: '共通', unit: 'mmHg', a_min: 0, a_max: 129, b_min: 130, b_max: 139, c_min: 140, c_max: 159, d_min: 160, d_max: null, note: '' },
  { item_id: 'BP_DIASTOLIC', item_name: '血圧(拡張)', gender: '共通', unit: 'mmHg', a_min: 0, a_max: 84, b_min: 85, b_max: 89, c_min: 90, c_max: 99, d_min: 100, d_max: null, note: '' },
  { item_id: 'FBS', item_name: '空腹時血糖', gender: '共通', unit: 'mg/dl', a_min: 0, a_max: 99, b_min: 100, b_max: 109, c_min: 110, c_max: 125, d_min: 126, d_max: null, note: 'HbA1cと組合せ判定' },
  { item_id: 'HBA1C', item_name: 'HbA1c', gender: '共通', unit: '%', a_min: 0, a_max: 5.5, b_min: 5.6, b_max: 5.9, c_min: 6.0, c_max: 6.4, d_min: 6.5, d_max: null, note: 'FBSと組合せ判定' },
  { item_id: 'HDL_C', item_name: 'HDLコレステロール', gender: '共通', unit: 'mg/dl', a_min: 40, a_max: 999, b_min: null, b_max: null, c_min: 35, c_max: 39, d_min: 0, d_max: 34, note: '低値が異常' },
  { item_id: 'LDL_C', item_name: 'LDLコレステロール', gender: '共通', unit: 'mg/dl', a_min: 60, a_max: 119, b_min: 120, b_max: 139, c_min: 140, c_max: 179, d_min: 180, d_max: null, note: '' },
  { item_id: 'TG', item_name: '中性脂肪', gender: '共通', unit: 'mg/dl', a_min: 30, a_max: 149, b_min: 150, b_max: 299, c_min: 300, c_max: 499, d_min: 500, d_max: null, note: '空腹時' },
  { item_id: 'AST', item_name: 'AST(GOT)', gender: '共通', unit: 'U/I', a_min: 0, a_max: 30, b_min: 31, b_max: 35, c_min: 36, c_max: 50, d_min: 51, d_max: null, note: '' },
  { item_id: 'ALT', item_name: 'ALT(GPT)', gender: '共通', unit: 'U/I', a_min: 0, a_max: 30, b_min: 31, b_max: 40, c_min: 41, c_max: 50, d_min: 51, d_max: null, note: '' },
  { item_id: 'GGT', item_name: 'γ-GTP', gender: '共通', unit: 'U/I', a_min: 0, a_max: 50, b_min: 51, b_max: 80, c_min: 81, c_max: 100, d_min: 101, d_max: null, note: '' },
  { item_id: 'CREATININE_M', item_name: 'クレアチニン(男)', gender: 'M', unit: 'mg/dl', a_min: 0.1, a_max: 1.0, b_min: 1.01, b_max: 1.09, c_min: 1.1, c_max: 1.29, d_min: 1.3, d_max: null, note: 'eGFR優先' },
  { item_id: 'CREATININE_F', item_name: 'クレアチニン(女)', gender: 'F', unit: 'mg/dl', a_min: 0.1, a_max: 0.7, b_min: 0.71, b_max: 0.79, c_min: 0.8, c_max: 0.99, d_min: 1.0, d_max: null, note: 'eGFR優先' },
  { item_id: 'EGFR', item_name: 'eGFR', gender: '共通', unit: 'ml/分/1.73m²', a_min: 60.0, a_max: 999, b_min: 50.0, b_max: 59.9, c_min: 45.0, c_max: 49.9, d_min: 0, d_max: 44.9, note: '低値が異常' },
  { item_id: 'UA', item_name: '尿酸', gender: '共通', unit: 'mg/dl', a_min: 2.1, a_max: 7.0, b_min: 7.1, b_max: 7.9, c_min: 8.0, c_max: 8.9, d_min: 9.0, d_max: null, note: '' },
  { item_id: 'HEMOGLOBIN_M', item_name: '血色素量(男)', gender: 'M', unit: 'g/dl', a_min: 13.1, a_max: 16.3, b_min: 16.4, b_max: 18.0, c_min: 12.1, c_max: 13.0, d_min: 0, d_max: 12.0, note: '低値・高値両方' },
  { item_id: 'HEMOGLOBIN_F', item_name: '血色素量(女)', gender: 'F', unit: 'g/dl', a_min: 12.1, a_max: 14.5, b_min: 14.6, b_max: 16.0, c_min: 11.1, c_max: 12.0, d_min: 0, d_max: 11.0, note: '低値・高値両方' },
  { item_id: 'PLATELET', item_name: '血小板', gender: '共通', unit: '万/μl', a_min: 14.5, a_max: 32.9, b_min: 12.3, b_max: 14.4, c_min: 10.0, c_max: 12.2, d_min: 0, d_max: 9.9, note: '' },
  { item_id: 'TOTAL_PROTEIN', item_name: '総蛋白', gender: '共通', unit: 'g/dl', a_min: 6.5, a_max: 7.9, b_min: 8.0, b_max: 8.3, c_min: 6.2, c_max: 6.4, d_min: 0, d_max: 6.1, note: '' },
  { item_id: 'ALBUMIN', item_name: 'アルブミン', gender: '共通', unit: 'g/dl', a_min: 3.9, a_max: 5.5, b_min: 3.7, b_max: 3.8, c_min: 3.0, c_max: 3.6, d_min: 0, d_max: 2.9, note: '' }
];

// ============================================
// 選択肢マスタデータ（定性検査用）
// ============================================

const SELECT_OPTIONS_DATA = [
  { item_id: 'URINE_PROTEIN', options: '(-)|(±)|(+)|(++)|(+++)', description: '尿蛋白判定' },
  { item_id: 'URINE_GLUCOSE', options: '(-)|(±)|(+)|(++)|(+++)', description: '尿糖判定' },
  { item_id: 'URINE_OCCULT_BLOOD', options: '(-)|(±)|(+)|(++)|(+++)', description: '尿潜血判定' },
  { item_id: 'UROBILINOGEN', options: '(-)|(±)|(+)|(++)|(+++)', description: 'ウロビリノーゲン判定' },
  { item_id: 'URINE_BILIRUBIN', options: '(-)|(+)|(++)|(+++)', description: '尿ビリルビン判定' },
  { item_id: 'URINE_KETONE', options: '(-)|(±)|(+)|(++)|(+++)', description: 'ケトン体判定' },
  { item_id: 'FOBT_1', options: '(-)|(+)', description: '便潜血1回目' },
  { item_id: 'FOBT_2', options: '(-)|(+)', description: '便潜血2回目' },
  { item_id: 'SEX', options: 'M|F', description: '性別' },
  { item_id: 'BLOOD_TYPE_ABO', options: 'A|B|O|AB', description: 'ABO血液型' },
  { item_id: 'BLOOD_TYPE_RH', options: '(+)|(-)', description: 'Rh血液型' },
  { item_id: 'HEARING_R_1000', options: '異常なし|所見あり', description: '聴力判定' },
  { item_id: 'HEARING_L_1000', options: '異常なし|所見あり', description: '聴力判定' },
  { item_id: 'HEARING_R_4000', options: '異常なし|所見あり', description: '聴力判定' },
  { item_id: 'HEARING_L_4000', options: '異常なし|所見あり', description: '聴力判定' },
  { item_id: 'HBS_AG', options: '(-)|(+)', description: 'HBs抗原' },
  { item_id: 'HBS_AB', options: '(-)|(+)', description: 'HBs抗体' },
  { item_id: 'HCV_AB', options: '(-)|(+)', description: 'HCV抗体' },
  { item_id: 'HIV_AB', options: '(-)|(+)', description: 'HIV抗体' },
  { item_id: 'TPHA', options: '(-)|(+)', description: '梅毒TPHA' },
  { item_id: 'RPR', options: '(-)|(+)', description: '梅毒RPR' },
  { item_id: 'OVERALL_JUDGMENT', options: 'A|B|C|D|E|F', description: '総合判定' }
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
    required_items: 'HEIGHT,WEIGHT,BMI,WAIST_M,WAIST_F,BP_SYSTOLIC_1,BP_DIASTOLIC_1,FBS,HBA1C,HDL_C,LDL_C,TG,AST,ALT,GGT'
  },
  {
    course_id: 'DOCK_DIGESTIVE',
    course_name: '消化器ドック',
    price: 60000,
    description: 'GS+CS、腫瘍マーカー',
    item_count: 95,
    required_items: 'HEIGHT,WEIGHT,BMI,WAIST_M,WAIST_F,BP_SYSTOLIC_1,BP_DIASTOLIC_1,FBS,HBA1C,HDL_C,LDL_C,TG,AST,ALT,GGT,CEA,CA19_9,UPPER_GI_ENDOSCOPY'
  },
  {
    course_id: 'DOCK_PREMIUM',
    course_name: '全身スクリーニング',
    price: 160000,
    description: '全身MRI(DWIBS)含む',
    item_count: 120,
    required_items: 'HEIGHT,WEIGHT,BMI,WAIST_M,WAIST_F,BP_SYSTOLIC_1,BP_DIASTOLIC_1,FBS,HBA1C,HDL_C,LDL_C,TG,AST,ALT,GGT,CEA,CA19_9,AFP,PSA,BRAIN_MRI,BRAIN_MRA'
  },
  {
    course_id: 'DOCK_CANCER',
    course_name: 'がんスクリーニング',
    price: 70000,
    description: 'CT、腫瘍マーカー多数',
    item_count: 100,
    required_items: 'HEIGHT,WEIGHT,BMI,CEA,CA19_9,AFP,PSA,CA125,NSE,CHEST_CT,ABDOMINAL_CT'
  },
  {
    course_id: 'REGULAR_CHECKUP',
    course_name: '定期健康診断',
    price: 8000,
    description: '労働安全衛生法準拠',
    item_count: 35,
    required_items: 'HEIGHT,WEIGHT,BMI,BP_SYSTOLIC_1,BP_DIASTOLIC_1,URINE_PROTEIN,URINE_GLUCOSE,RBC,HEMOGLOBIN,AST,ALT,GGT,TG,HDL_C,LDL_C,FBS,HBA1C'
  },
  {
    course_id: 'SECONDARY_EXAM',
    course_name: '労災二次健診',
    price: 0,
    description: '脳・心疾患リスク評価',
    item_count: 20,
    required_items: 'BP_SYSTOLIC_1,BP_DIASTOLIC_1,FBS,HBA1C,HDL_C,LDL_C,TG,CREATININE,HEMOGLOBIN,ECG'
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
 * 項目IDからマスタデータを取得
 * @param {string} itemId - 項目ID
 * @returns {Object|null} マスタデータ
 */
function getExamItemById(itemId) {
  return EXAM_ITEM_MASTER_DATA.find(item => item.item_id === itemId) || null;
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
