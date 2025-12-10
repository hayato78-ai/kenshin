@echo off
chcp 65001 > nul
echo ============================================
echo   人間ドック帳票生成 - 監視プログラム
echo ============================================
echo.

cd /d "%~dp0"

echo フォルダを確認中...
if not exist "pending" mkdir pending
if not exist "processed" mkdir processed
if not exist "error" mkdir error

echo.
echo 監視を開始します...
echo （このウィンドウを閉じると監視が停止します）
echo.

python drive_watcher.py

if errorlevel 1 (
    echo.
    echo エラーが発生しました。
    echo Pythonがインストールされているか確認してください。
    pause
)
