#!/bin/bash
echo "============================================"
echo "  人間ドック帳票生成 - 監視プログラム"
echo "============================================"
echo ""

cd "$(dirname "$0")"

# フォルダ作成
mkdir -p pending processed error

echo "監視を開始します..."
echo "（このウィンドウを閉じると監視が停止します）"
echo ""

python3 drive_watcher.py
