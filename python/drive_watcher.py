#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Drive監視モジュール（ポーリング方式）
GASからのExcel出力リクエスト（JSON）を検知して処理を実行

Google Drive for Desktop の同期はwatchdogのファイル作成イベントを発火しないため、
ポーリング方式（定期的にpendingフォルダをスキャン）で監視する。

使い方:
    from drive_watcher import DriveWatcher

    watcher = DriveWatcher(settings_path='settings.yaml')
    watcher.start()  # 監視開始（5秒ごとにpendingフォルダをスキャン）
"""

import os
import sys
import json
import time
import logging
import shutil
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Callable, Set

# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('drive_watcher.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class ExportRequestHandler:
    """
    JSONファイルを処理するハンドラ（ポーリング方式用）
    """

    def __init__(
        self,
        pending_folder: Path,
        processed_folder: Path,
        error_folder: Path,
        processor: Callable[[Dict], Dict]
    ):
        """
        初期化

        Args:
            pending_folder: 待機フォルダ（GASがJSONを置く場所）
            processed_folder: 処理済みフォルダ
            error_folder: エラーフォルダ
            processor: 処理関数（JSONデータを受け取り結果を返す）
        """
        self.pending_folder = Path(pending_folder)
        self.processed_folder = Path(processed_folder)
        self.error_folder = Path(error_folder)
        self.processor = processor

        # フォルダ作成
        self.pending_folder.mkdir(parents=True, exist_ok=True)
        self.processed_folder.mkdir(parents=True, exist_ok=True)
        self.error_folder.mkdir(parents=True, exist_ok=True)

    def _process_request(self, json_path: Path):
        """
        リクエストを処理

        Args:
            json_path: リクエストJSONファイルパス
        """
        request_id = json_path.stem

        try:
            # JSON読み込み
            with open(json_path, 'r', encoding='utf-8') as f:
                request_data = json.load(f)

            logger.info(f"📝 処理開始: {request_id}")
            logger.info(f"   検査種別: {request_data.get('exam_type', 'UNKNOWN')}")
            logger.info(f"   患者名: {request_data.get('patient', {}).get('name', 'UNKNOWN')}")

            # 処理実行
            result = self.processor(request_data)

            if result.get('success', False):
                # 成功時: 処理済みフォルダへ移動
                self._move_to_processed(json_path, result)
                logger.info(f"✅ 処理完了: {request_id}")
                logger.info(f"   出力: {result.get('output_path', 'N/A')}")
            else:
                # 失敗時: エラーフォルダへ移動
                self._move_to_error(json_path, result)
                logger.error(f"❌ 処理失敗: {request_id}")
                logger.error(f"   エラー: {result.get('error', 'Unknown error')}")

        except json.JSONDecodeError as e:
            logger.error(f"❌ JSON解析エラー: {json_path.name} - {e}")
            self._move_to_error(json_path, {'error': f'JSON parse error: {e}'})

        except Exception as e:
            logger.error(f"❌ 処理エラー: {json_path.name} - {e}")
            self._move_to_error(json_path, {'error': str(e)})

    def _move_to_processed(self, json_path: Path, result: Dict):
        """処理済みフォルダへ移動"""
        # 結果ファイルを作成
        result_path = self.processed_folder / f"{json_path.stem}_result.json"
        with open(result_path, 'w', encoding='utf-8') as f:
            json.dump({
                'request_id': json_path.stem,
                'status': 'completed',
                'completed_at': datetime.now().isoformat(),
                'result': result
            }, f, ensure_ascii=False, indent=2)

        # 元ファイルを移動
        dest_path = self.processed_folder / json_path.name
        shutil.move(str(json_path), str(dest_path))

    def _move_to_error(self, json_path: Path, result: Dict):
        """エラーフォルダへ移動"""
        # エラー結果ファイルを作成
        error_result_path = self.error_folder / f"{json_path.stem}_error.json"
        with open(error_result_path, 'w', encoding='utf-8') as f:
            json.dump({
                'request_id': json_path.stem,
                'status': 'error',
                'error_at': datetime.now().isoformat(),
                'error': result.get('error', 'Unknown error')
            }, f, ensure_ascii=False, indent=2)

        # 元ファイルを移動
        dest_path = self.error_folder / json_path.name
        if json_path.exists():
            shutil.move(str(json_path), str(dest_path))


class DriveWatcher:
    """
    Google Drive監視クラス（ポーリング方式）
    pendingフォルダを定期的にスキャンし、新規JSONファイルを検知して処理

    Google Drive for Desktopの同期はwatchdogのファイル作成イベントを発火しないため、
    ポーリング方式で監視する。
    """

    def __init__(
        self,
        pending_folder: str,
        processed_folder: str,
        error_folder: str,
        processor: Callable[[Dict], Dict],
        poll_interval: float = 5.0
    ):
        """
        初期化

        Args:
            pending_folder: 待機フォルダパス
            processed_folder: 処理済みフォルダパス
            error_folder: エラーフォルダパス
            processor: 処理関数
            poll_interval: ポーリング間隔（秒）- デフォルト5秒
        """
        self.pending_folder = Path(pending_folder)
        self.processed_folder = Path(processed_folder)
        self.error_folder = Path(error_folder)
        self.processor = processor
        self.poll_interval = poll_interval

        self.handler = None
        self._running = False
        self._processed_files: Set[str] = set()  # 処理済みファイル追跡用

    @classmethod
    def from_settings(cls, settings_path: str, processor: Callable[[Dict], Dict]) -> 'DriveWatcher':
        """
        設定ファイルから初期化

        Args:
            settings_path: settings.yaml のパス
            processor: 処理関数

        Returns:
            DriveWatcher インスタンス
        """
        import yaml

        with open(settings_path, 'r', encoding='utf-8') as f:
            settings = yaml.safe_load(f)

        folders = settings.get('folders', {})

        return cls(
            pending_folder=folders.get('pending', './pending'),
            processed_folder=folders.get('processed', './processed'),
            error_folder=folders.get('error', './error'),
            processor=processor,
            poll_interval=settings.get('poll_interval', 5.0)
        )

    def start(self, process_existing: bool = True):
        """
        監視開始（ポーリング方式）

        Args:
            process_existing: 起動時に既存ファイルを処理するか（常にTrue扱い）
        """
        logger.info("=" * 60)
        logger.info("🚀 Drive Watcher 起動（ポーリング方式）")
        logger.info(f"   監視フォルダ: {self.pending_folder}")
        logger.info(f"   処理済み: {self.processed_folder}")
        logger.info(f"   エラー: {self.error_folder}")
        logger.info(f"   ポーリング間隔: {self.poll_interval}秒")
        logger.info("=" * 60)

        # ハンドラ作成
        self.handler = ExportRequestHandler(
            pending_folder=self.pending_folder,
            processed_folder=self.processed_folder,
            error_folder=self.error_folder,
            processor=self.processor
        )

        self._running = True
        logger.info("👁️ 監視中... (Ctrl+C で終了)")

        try:
            while self._running:
                # ポーリング: pendingフォルダをスキャン
                self._poll_pending_folder()
                time.sleep(self.poll_interval)
        except KeyboardInterrupt:
            self.stop()

    def _poll_pending_folder(self):
        """
        pendingフォルダをスキャンして新規ファイルを処理
        """
        try:
            # JSONファイルを取得
            existing_files = list(self.pending_folder.glob('*.json'))

            # _result.json を除外
            json_files = [f for f in existing_files if not f.stem.endswith('_result')]

            for json_file in json_files:
                # ファイル名で重複処理を防止
                file_key = json_file.name
                if file_key in self._processed_files:
                    continue

                logger.info(f"📥 新規リクエスト検知: {json_file.name}")

                # 処理中としてマーク（処理完了後に移動されるので削除不要）
                self._processed_files.add(file_key)

                # 少し待機（ファイル書き込み完了を待つ）
                time.sleep(0.5)

                # 処理実行
                self.handler._process_request(json_file)

                # 処理済みリストから削除（ファイルは移動されているはず）
                self._processed_files.discard(file_key)

        except Exception as e:
            logger.error(f"❌ ポーリングエラー: {e}")

    def stop(self):
        """監視停止"""
        logger.info("🛑 監視停止中...")
        self._running = False
        logger.info("👋 Drive Watcher 終了")


# テスト用ダミープロセッサ
def dummy_processor(request_data: Dict) -> Dict:
    """
    テスト用ダミープロセッサ
    実際にはunified_transcriber.pyの関数を使用
    """
    logger.info(f"[DUMMY] Processing: {request_data.get('request_id', 'UNKNOWN')}")

    # ダミー処理（実際にはExcel出力を行う）
    time.sleep(2)

    return {
        'success': True,
        'output_path': '/path/to/output.xlsx',
        'message': 'Dummy processing completed'
    }


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Google Drive 監視スクリプト')
    parser.add_argument('--pending', default='./pending', help='待機フォルダパス')
    parser.add_argument('--processed', default='./processed', help='処理済みフォルダパス')
    parser.add_argument('--error', default='./error', help='エラーフォルダパス')
    parser.add_argument('--interval', type=float, default=2.0, help='監視間隔（秒）')
    parser.add_argument('--test', action='store_true', help='テストモード（ダミープロセッサ使用）')

    args = parser.parse_args()

    # プロセッサ選択
    if args.test:
        processor = dummy_processor
        logger.info("⚠️ テストモード: ダミープロセッサを使用")
    else:
        # 本番用: unified_transcriber をインポート
        try:
            from unified_transcriber import process_export_request
            processor = process_export_request
        except ImportError:
            logger.warning("⚠️ unified_transcriber が見つかりません。ダミープロセッサを使用")
            processor = dummy_processor

    # 監視開始
    watcher = DriveWatcher(
        pending_folder=args.pending,
        processed_folder=args.processed,
        error_folder=args.error,
        processor=processor,
        poll_interval=args.interval
    )

    watcher.start()
