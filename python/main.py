#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健診結果Excel出力システム - メインエントリーポイント

GASからのリクエストを監視し、Excelファイルを生成するシステム

使い方:
    # 監視モード（通常運用）
    python main.py

    # テストモード
    python main.py --test

    # 単発処理モード
    python main.py --single request.json

    # 設定確認
    python main.py --check-config
"""

import os
import sys
import argparse
import logging
from pathlib import Path
from datetime import datetime

# ロギング設定
def setup_logging(log_dir: Path = None, debug: bool = False):
    """ロギング設定"""
    log_level = logging.DEBUG if debug else logging.INFO

    handlers = [logging.StreamHandler()]

    if log_dir:
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f"excel_export_{datetime.now().strftime('%Y%m%d')}.log"
        handlers.append(logging.FileHandler(log_file, encoding='utf-8'))

    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )

    return logging.getLogger(__name__)


def print_banner():
    """バナー表示"""
    banner = """
╔════════════════════════════════════════════════════════════════╗
║          健診結果 Excel 出力システム v1.0                       ║
║          GAS → Python Excel 生成                               ║
╠════════════════════════════════════════════════════════════════╣
║  対応検査: 労災二次検診 / 人間ドック / 定期検診                  ║
╚════════════════════════════════════════════════════════════════╝
    """
    print(banner)


def check_dependencies():
    """依存関係チェック"""
    required = ['openpyxl', 'yaml', 'watchdog']
    missing = []

    for package in required:
        try:
            if package == 'yaml':
                import yaml
            elif package == 'openpyxl':
                import openpyxl
            elif package == 'watchdog':
                from watchdog.observers import Observer
        except ImportError:
            missing.append(package)

    if missing:
        print(f"❌ 必要なパッケージがインストールされていません: {', '.join(missing)}")
        print(f"   インストール: pip install {' '.join(missing)}")
        return False

    return True


def check_config(settings_path: Path):
    """設定ファイルチェック"""
    import yaml

    print("\n📋 設定チェック")
    print("=" * 50)

    if not settings_path.exists():
        print(f"❌ 設定ファイルが見つかりません: {settings_path}")
        return False

    with open(settings_path, 'r', encoding='utf-8') as f:
        settings = yaml.safe_load(f)

    # フォルダ存在確認
    folders = settings.get('folders', {})
    all_ok = True

    for folder_type, folder_path in folders.items():
        path = Path(folder_path)
        exists = path.exists()
        status = "✅" if exists else "❌"
        print(f"  {status} {folder_type}: {folder_path}")
        if not exists:
            all_ok = False

    # テンプレート確認
    templates_dir = Path(settings.get('templates_dir', './templates'))
    print(f"\n📁 テンプレートディレクトリ: {templates_dir}")

    if templates_dir.exists():
        templates = list(templates_dir.glob('*.xlsx'))
        for t in templates:
            print(f"  ✅ {t.name}")
    else:
        print(f"  ❌ ディレクトリが見つかりません")
        all_ok = False

    # マッピング確認
    mappings_dir = Path(settings.get('mappings_dir', './config'))
    print(f"\n📁 マッピングディレクトリ: {mappings_dir}")

    if mappings_dir.exists():
        mappings = list(mappings_dir.glob('*.yaml'))
        for m in mappings:
            print(f"  ✅ {m.name}")
    else:
        print(f"  ❌ ディレクトリが見つかりません")
        all_ok = False

    print("\n" + "=" * 50)
    if all_ok:
        print("✅ 設定チェック完了 - 問題なし")
    else:
        print("⚠️ 設定に問題があります。修正してください。")

    return all_ok


def run_watcher(settings_path: Path, test_mode: bool = False):
    """監視モードで実行"""
    from drive_watcher import DriveWatcher
    from unified_transcriber import process_export_request

    import yaml

    # 設定読み込み
    if settings_path.exists():
        with open(settings_path, 'r', encoding='utf-8') as f:
            settings = yaml.safe_load(f)
    else:
        # デフォルト設定
        settings = {
            'folders': {
                'pending': './pending',
                'processed': './processed',
                'error': './error'
            },
            'poll_interval': 2.0
        }

    folders = settings.get('folders', {})

    # プロセッサ選択
    if test_mode:
        from drive_watcher import dummy_processor
        processor = dummy_processor
        print("⚠️ テストモード: ダミープロセッサを使用")
    else:
        processor = process_export_request

    # 監視開始
    watcher = DriveWatcher(
        pending_folder=folders.get('pending', './pending'),
        processed_folder=folders.get('processed', './processed'),
        error_folder=folders.get('error', './error'),
        processor=processor,
        poll_interval=settings.get('poll_interval', 2.0)
    )

    watcher.start()


def run_single(json_path: Path, output_dir: Path = None):
    """単発処理モード"""
    import json
    from unified_transcriber import UnifiedTranscriber

    print(f"\n📝 単発処理モード: {json_path}")

    if not json_path.exists():
        print(f"❌ ファイルが見つかりません: {json_path}")
        return False

    # JSON読み込み
    with open(json_path, 'r', encoding='utf-8') as f:
        request_data = json.load(f)

    # 転記実行
    settings_path = Path(__file__).parent / 'settings.yaml'
    transcriber = UnifiedTranscriber(
        settings_path=str(settings_path) if settings_path.exists() else None
    )

    result = transcriber.transcribe(
        request_data,
        output_dir=str(output_dir) if output_dir else None
    )

    if result['success']:
        print(f"✅ 成功: {result['output_path']}")
        print(f"   転記項目数: {result.get('transcribed_count', 'N/A')}")
        return True
    else:
        print(f"❌ 失敗: {result['error']}")
        return False


def create_sample_settings(settings_path: Path):
    """サンプル設定ファイルを作成"""
    import yaml

    sample_settings = {
        'base_dir': str(Path(__file__).parent),
        'folders': {
            'pending': './pending',
            'processed': './processed',
            'error': './error'
        },
        'output_dir': './output',
        'templates_dir': './templates',
        'mappings_dir': './config',
        'poll_interval': 2.0,
        'log_dir': './logs',
        'exam_types': {
            'ROSAI_SECONDARY': {
                'template': 'kenshin_idheart.xlsx',
                'mapping': 'idheart_cell_mapping.yaml',
                'sheet_name': '入力用'
            },
            'HUMAN_DOCK': {
                'template': 'human_dock_template.xlsx',
                'mapping': 'human_dock_cell_mapping.yaml',
                'sheet_name': '結果入力'
            }
        }
    }

    with open(settings_path, 'w', encoding='utf-8') as f:
        yaml.dump(sample_settings, f, allow_unicode=True, default_flow_style=False)

    print(f"✅ サンプル設定ファイルを作成しました: {settings_path}")


def main():
    parser = argparse.ArgumentParser(
        description='健診結果Excel出力システム',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python main.py                      # 監視モード（通常運用）
  python main.py --test               # テストモード
  python main.py --single req.json    # 単発処理
  python main.py --check-config       # 設定確認
  python main.py --init               # 初期設定ファイル作成
        """
    )

    parser.add_argument('--test', action='store_true', help='テストモード（ダミープロセッサ使用）')
    parser.add_argument('--single', metavar='JSON_FILE', help='単発処理モード')
    parser.add_argument('--output-dir', metavar='DIR', help='出力ディレクトリ（--single時）')
    parser.add_argument('--check-config', action='store_true', help='設定ファイルをチェック')
    parser.add_argument('--init', action='store_true', help='初期設定ファイルを作成')
    parser.add_argument('--settings', default='settings.yaml', help='設定ファイルパス')
    parser.add_argument('--debug', action='store_true', help='デバッグモード')

    args = parser.parse_args()

    # バナー表示
    print_banner()

    # 基準ディレクトリ
    base_dir = Path(__file__).parent
    settings_path = base_dir / args.settings

    # ログ設定
    log_dir = base_dir / 'logs' if not args.debug else None
    logger = setup_logging(log_dir, args.debug)

    # 依存関係チェック
    if not check_dependencies():
        sys.exit(1)

    # 初期設定ファイル作成
    if args.init:
        create_sample_settings(settings_path)
        print("\n次のステップ:")
        print("  1. settings.yaml を編集してパスを設定")
        print("  2. python main.py --check-config で設定確認")
        print("  3. python main.py で監視開始")
        sys.exit(0)

    # 設定確認
    if args.check_config:
        success = check_config(settings_path)
        sys.exit(0 if success else 1)

    # 単発処理
    if args.single:
        json_path = Path(args.single)
        output_dir = Path(args.output_dir) if args.output_dir else None
        success = run_single(json_path, output_dir)
        sys.exit(0 if success else 1)

    # 監視モード
    try:
        run_watcher(settings_path, test_mode=args.test)
    except KeyboardInterrupt:
        print("\n👋 終了")
        sys.exit(0)
    except Exception as e:
        logger.error(f"❌ エラー: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
