"""
BML検査結果CSVパーサー
"""

from pathlib import Path
from typing import Optional, Dict, List

from .constants import (
    SUPPORTED_ENCODINGS,
    MIN_CSV_FIELDS,
    TEST_RESULT_START_INDEX,
    TEST_RESULT_FIELD_COUNT,
)


class BMLResultParser:
    """
    BML検査結果CSVパーサー

    CSVフォーマット:
        施設コード,依頼ID,受診日,時刻,,保険番号,性別,,,フラグ,区分,
        検査コード,結果値,フラグ,コメント,検査コード,結果値,フラグ,コメント,...
    """

    def parse(self, csv_path: Path) -> List[Dict]:
        """
        BML結果CSVを解析

        Args:
            csv_path: CSVファイルパス

        Returns:
            患者ごとの検査結果リスト

        Raises:
            ValueError: エンコーディング検出失敗時
            FileNotFoundError: ファイルが見つからない場合
        """
        content = self._read_file(csv_path)

        results = []
        for line in content.strip().split('\n'):
            if not line.strip():
                continue

            parsed = self._parse_line(line)
            if parsed:
                results.append(parsed)

        return results

    def _read_file(self, csv_path: Path) -> str:
        """ファイルを読み込み（エンコーディング自動検出）"""
        for enc in SUPPORTED_ENCODINGS:
            try:
                with open(csv_path, 'r', encoding=enc) as f:
                    return f.read()
            except UnicodeDecodeError:
                continue

        raise ValueError(f"ファイルのエンコーディングを検出できません: {csv_path}")

    def _parse_line(self, line: str) -> Optional[Dict]:
        """
        1行を解析

        Args:
            line: CSV行文字列

        Returns:
            解析結果辞書、または無効行の場合None
        """
        fields = line.split(',')

        if len(fields) < MIN_CSV_FIELDS:
            return None

        patient_info = self._extract_patient_info(fields)
        test_results = self._extract_test_results(fields)

        return {
            'patient_info': patient_info,
            'test_results': test_results
        }

    def _extract_patient_info(self, fields: List[str]) -> Dict:
        """患者情報を抽出"""
        return {
            'facility_code': fields[0],
            'request_id': fields[1],
            'exam_date': fields[2],
            'time': fields[3],
            'insurance_no': fields[5] if len(fields) > 5 else '',
            'gender': fields[6] if len(fields) > 6 else '',
        }

    def _extract_test_results(self, fields: List[str]) -> List[Dict]:
        """検査結果を抽出"""
        test_results = []
        i = TEST_RESULT_START_INDEX

        while i < len(fields) - 1:
            code = fields[i].strip() if i < len(fields) else ''

            if not code or not code.isdigit():
                i += 1
                continue

            result = self._extract_single_result(fields, i)
            if result:
                test_results.append(result)

            i += TEST_RESULT_FIELD_COUNT

        return test_results

    def _extract_single_result(self, fields: List[str], index: int) -> Optional[Dict]:
        """単一の検査結果を抽出"""
        code = fields[index].strip()
        value = fields[index + 1].strip() if index + 1 < len(fields) else ''
        flag = fields[index + 2].strip() if index + 2 < len(fields) else ''
        comment = fields[index + 3].strip() if index + 3 < len(fields) else ''

        if not value:
            return None

        return {
            'code': code,
            'value': value,
            'flag': flag,
            'comment': comment
        }
