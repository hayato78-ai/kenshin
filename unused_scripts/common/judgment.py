"""
判定ロジックエンジン（人間ドック学会2025年度基準）
"""

from typing import Dict, Any, Optional

from .constants import (
    CODE_TO_CRITERIA,
    GENDER_DEPENDENT_CODES,
    FLAG_HIGH,
    FLAG_LOW,
    JUDGMENT_GRADES,
    DEFAULT_JUDGMENT_NORMAL,
    DEFAULT_JUDGMENT_ABNORMAL,
)


class JudgmentEngine:
    """
    検査値から判定（A/B/C/D）を決定するエンジン

    判定基準は人間ドック学会2025年度基準に基づく
    """

    def __init__(self, criteria: Dict[str, Any]):
        """
        Args:
            criteria: 判定基準辞書 {項目キー: {A: {min, max}, B: {...}, ...}}
        """
        self.criteria = criteria

    def judge(self, item_key: str, value: float, gender: Optional[str] = None) -> str:
        """
        検査値から判定を返す

        Args:
            item_key: 判定基準キー
            value: 検査値（数値）
            gender: 性別 ("M" or "F")

        Returns:
            判定結果 ("A", "B", "C", "D") または空文字列
        """
        if item_key not in self.criteria:
            return ""

        spec = self.criteria[item_key]

        # 性別フィルタリング
        if "gender" in spec and spec["gender"] != gender:
            return ""

        # 各グレードをチェック
        for grade in JUDGMENT_GRADES:
            rule = spec.get(grade)
            if rule is None:
                continue
            if self._check_range(value, rule):
                return grade

        return ""

    def judge_by_code(
        self,
        code: str,
        value_str: str,
        flag: str,
        gender: str
    ) -> str:
        """
        検査コードと値文字列から判定を返す

        Args:
            code: BML検査コード
            value_str: 検査値（文字列）
            flag: フラグ (H/L/空)
            gender: 性別 ("M" or "F")

        Returns:
            判定結果 ("A", "B", "C", "D") または空文字列
        """
        # 数値変換
        try:
            numeric_value = float(value_str)
        except (ValueError, TypeError):
            return ""

        # 判定基準キーを取得
        criteria_key = self._get_criteria_key(code, gender)

        if criteria_key:
            result = self.judge(criteria_key, numeric_value, gender)
            if result:
                return result

        # 判定基準がない場合はフラグから推定
        return self._judge_by_flag(flag)

    def _get_criteria_key(self, code: str, gender: str) -> Optional[str]:
        """検査コードから判定基準キーを取得"""
        base_key = CODE_TO_CRITERIA.get(code)

        if not base_key:
            return None

        # 性別依存項目の場合はサフィックスを付加
        if code in GENDER_DEPENDENT_CODES:
            return f"{base_key}_{gender}"

        return base_key

    def _judge_by_flag(self, flag: str) -> str:
        """フラグから判定を推定"""
        if flag == FLAG_HIGH or flag == FLAG_LOW:
            return DEFAULT_JUDGMENT_ABNORMAL
        return DEFAULT_JUDGMENT_NORMAL

    def _check_range(self, value: float, rule: Dict) -> bool:
        """
        値が範囲内かチェック

        Args:
            value: 検査値
            rule: 範囲ルール {min: float|None, max: float|None, or: {...}}

        Returns:
            範囲内ならTrue
        """
        min_val = rule.get("min")
        max_val = rule.get("max")

        in_main_range = True
        if min_val is not None and value < min_val:
            in_main_range = False
        if max_val is not None and value > max_val:
            in_main_range = False

        if in_main_range and (min_val is not None or max_val is not None):
            return True

        # OR条件がある場合
        if "or" in rule:
            return self._check_range(value, rule["or"])

        return False
