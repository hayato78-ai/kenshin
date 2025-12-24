"""
定数定義モジュール
"""

# CSVパース設定
SUPPORTED_ENCODINGS = ['shift_jis', 'cp932', 'utf-8']
MIN_CSV_FIELDS = 12
TEST_RESULT_START_INDEX = 11
TEST_RESULT_FIELD_COUNT = 4

# 検査コード → 判定基準キーのマッピング
CODE_TO_CRITERIA = {
    "0000481": "AST_GOT",
    "0000482": "ALT_GPT",
    "0000484": "GAMMA_GTP",
    "0000460": "HDL_CHOLESTEROL",
    "0000410": "LDL_CHOLESTEROL",
    "0000454": "TRIGLYCERIDES",
    "0000503": "FASTING_GLUCOSE",
    "0003317": "HBA1C",
    "0000658": "CRP",
    "0002696": "EGFR",
    "0000401": "TOTAL_PROTEIN",
    "0000407": "URIC_ACID",
    # 性別依存項目
    "0000303": "HEMOGLOBIN",  # 血色素量（性別サフィックス付加）
    "0000413": "CREATININE",  # クレアチニン（性別サフィックス付加）
}

# 性別依存の検査コード
GENDER_DEPENDENT_CODES = {"0000303", "0000413"}

# 性別変換
GENDER_TRANSFORMS = {
    "1": "男",
    "2": "女",
    "M": "男",
    "F": "女",
}

# CSV性別コード → 内部表現
GENDER_CODE_TO_INTERNAL = {
    "1": "M",
    "2": "F",
}

# フラグ定義
FLAG_HIGH = "H"
FLAG_LOW = "L"

# 判定グレード
JUDGMENT_GRADES = ["A", "B", "C", "D"]

# デフォルト判定
DEFAULT_JUDGMENT_NORMAL = "A"
DEFAULT_JUDGMENT_ABNORMAL = "C"
