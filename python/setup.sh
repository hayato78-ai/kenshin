#!/bin/bash
# =============================================================================
# 健診結果Excel出力システム - macOS セットアップスクリプト
# =============================================================================
# 使用方法: bash setup.sh
# =============================================================================

set -e

echo "🍎 macOS セットアップを開始します..."

# 現在のスクリプトのディレクトリを取得
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Google Drive パスの自動検出
detect_google_drive_path() {
    # CloudStorage形式（新しいGoogle Drive for Desktop）
    local cloud_storage_base="$HOME/Library/CloudStorage"

    if [ -d "$cloud_storage_base" ]; then
        # GoogleDrive-で始まるディレクトリを検索
        local google_drive_dirs=$(find "$cloud_storage_base" -maxdepth 1 -type d -name "GoogleDrive-*" 2>/dev/null)

        if [ -n "$google_drive_dirs" ]; then
            # 最初に見つかったディレクトリを使用
            local first_dir=$(echo "$google_drive_dirs" | head -n 1)
            # マイドライブを追加
            echo "${first_dir}/マイドライブ"
            return 0
        fi
    fi

    # 従来のパス形式
    if [ -d "$HOME/Google Drive/マイドライブ" ]; then
        echo "$HOME/Google Drive/マイドライブ"
        return 0
    fi

    if [ -d "$HOME/Google ドライブ/マイドライブ" ]; then
        echo "$HOME/Google ドライブ/マイドライブ"
        return 0
    fi

    return 1
}

# Google Drive パスの検出
echo "📂 Google Drive パスを検出中..."
GOOGLE_DRIVE_PATH=$(detect_google_drive_path)

if [ -z "$GOOGLE_DRIVE_PATH" ]; then
    echo "❌ Google Drive が見つかりません。"
    echo ""
    echo "手動でパスを入力してください:"
    echo "例: /Users/username/Library/CloudStorage/GoogleDrive-email@example.com/マイドライブ"
    read -p "Google Drive パス: " GOOGLE_DRIVE_PATH
fi

# パスの存在確認
if [ ! -d "$GOOGLE_DRIVE_PATH" ]; then
    echo "❌ 指定されたパスが存在しません: $GOOGLE_DRIVE_PATH"
    exit 1
fi

echo "✅ Google Drive 検出: $GOOGLE_DRIVE_PATH"

# settings.yaml が既に存在するか確認
if [ -f "settings.yaml" ]; then
    echo ""
    echo "⚠️  settings.yaml は既に存在します。"
    read -p "上書きしますか？ (y/N): " confirm
    if [ "$confirm" != "y" ] && [ "$confirm" != "Y" ]; then
        echo "セットアップを中止しました。"
        exit 0
    fi
fi

# テンプレートから settings.yaml を生成
if [ ! -f "settings_template.yaml" ]; then
    echo "❌ settings_template.yaml が見つかりません。"
    exit 1
fi

echo "📝 settings.yaml を生成中..."

# sed でプレースホルダーを置換（macOS の sed は -i '' が必要）
sed "s|\${GOOGLE_DRIVE_BASE}|${GOOGLE_DRIVE_PATH}|g" settings_template.yaml > settings.yaml

echo "✅ settings.yaml を生成しました。"

# Python 依存関係の確認
echo ""
echo "🐍 Python 依存関係を確認中..."

if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 がインストールされていません。"
    echo "   Homebrew: brew install python3"
    echo "   または: https://www.python.org/downloads/"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo "✅ Python $PYTHON_VERSION"

# pip パッケージの確認
echo ""
echo "📦 必要なパッケージを確認中..."

REQUIRED_PACKAGES="openpyxl pyyaml"
MISSING_PACKAGES=""

for pkg in $REQUIRED_PACKAGES; do
    if ! python3 -c "import $pkg" 2>/dev/null; then
        MISSING_PACKAGES="$MISSING_PACKAGES $pkg"
    fi
done

if [ -n "$MISSING_PACKAGES" ]; then
    echo "⚠️  不足パッケージ:$MISSING_PACKAGES"
    read -p "インストールしますか？ (Y/n): " install_confirm
    if [ "$install_confirm" != "n" ] && [ "$install_confirm" != "N" ]; then
        pip3 install $MISSING_PACKAGES
        echo "✅ パッケージをインストールしました。"
    fi
else
    echo "✅ 必要なパッケージは全てインストール済みです。"
fi

# 完了メッセージ
echo ""
echo "=========================================="
echo "🎉 セットアップ完了！"
echo "=========================================="
echo ""
echo "📋 次のステップ:"
echo "   1. settings.yaml の内容を確認"
echo "   2. 監視モードを起動:"
echo "      python3 unified_transcriber.py --watch"
echo ""
echo "📂 フォルダ構成:"
echo "   pending/   - GASからのリクエストJSON"
echo "   processed/ - 処理完了したJSON"
echo "   output/    - 生成されたExcel"
echo ""
