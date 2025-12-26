# =============================================================================
# 健診結果Excel出力システム - Windows セットアップスクリプト
# =============================================================================
# 使用方法: PowerShell で実行
#   .\setup.ps1
# または
#   powershell -ExecutionPolicy Bypass -File setup.ps1
# =============================================================================

$ErrorActionPreference = "Stop"

Write-Host "🪟 Windows セットアップを開始します..." -ForegroundColor Cyan

# 現在のスクリプトのディレクトリに移動
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# Google Drive パスの自動検出
function Find-GoogleDrivePath {
    # 標準的なGoogle Driveパス
    $possiblePaths = @(
        "$env:USERPROFILE\Google Drive\マイドライブ",
        "$env:USERPROFILE\Google ドライブ\マイドライブ",
        "$env:USERPROFILE\My Drive",
        "G:\マイドライブ",
        "G:\My Drive"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    # Google Drive Stream (ドライブレター形式) の検出
    $driveLetters = @("G", "H", "I", "J", "K")
    foreach ($letter in $driveLetters) {
        $streamPath = "${letter}:\マイドライブ"
        if (Test-Path $streamPath) {
            return $streamPath
        }
        $streamPath = "${letter}:\My Drive"
        if (Test-Path $streamPath) {
            return $streamPath
        }
    }

    return $null
}

# Google Drive パスの検出
Write-Host "📂 Google Drive パスを検出中..." -ForegroundColor Yellow
$GoogleDrivePath = Find-GoogleDrivePath

if (-not $GoogleDrivePath) {
    Write-Host "❌ Google Drive が見つかりません。" -ForegroundColor Red
    Write-Host ""
    Write-Host "手動でパスを入力してください:"
    Write-Host "例: C:\Users\username\Google Drive\マイドライブ"
    Write-Host "例: G:\マイドライブ"
    $GoogleDrivePath = Read-Host "Google Drive パス"
}

# パスの存在確認
if (-not (Test-Path $GoogleDrivePath)) {
    Write-Host "❌ 指定されたパスが存在しません: $GoogleDrivePath" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Google Drive 検出: $GoogleDrivePath" -ForegroundColor Green

# settings.yaml が既に存在するか確認
if (Test-Path "settings.yaml") {
    Write-Host ""
    Write-Host "⚠️  settings.yaml は既に存在します。" -ForegroundColor Yellow
    $confirm = Read-Host "上書きしますか？ (y/N)"
    if ($confirm -ne "y" -and $confirm -ne "Y") {
        Write-Host "セットアップを中止しました。"
        exit 0
    }
}

# テンプレートから settings.yaml を生成
if (-not (Test-Path "settings_template.yaml")) {
    Write-Host "❌ settings_template.yaml が見つかりません。" -ForegroundColor Red
    exit 1
}

Write-Host "📝 settings.yaml を生成中..." -ForegroundColor Yellow

# テンプレートを読み込んでプレースホルダーを置換
$template = Get-Content "settings_template.yaml" -Raw -Encoding UTF8
$settings = $template -replace '\$\{GOOGLE_DRIVE_BASE\}', $GoogleDrivePath
$settings | Out-File -FilePath "settings.yaml" -Encoding UTF8 -NoNewline

Write-Host "✅ settings.yaml を生成しました。" -ForegroundColor Green

# Python 依存関係の確認
Write-Host ""
Write-Host "🐍 Python 依存関係を確認中..." -ForegroundColor Yellow

$pythonCmd = $null
if (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = "python"
} elseif (Get-Command python3 -ErrorAction SilentlyContinue) {
    $pythonCmd = "python3"
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCmd = "py"
}

if (-not $pythonCmd) {
    Write-Host "❌ Python がインストールされていません。" -ForegroundColor Red
    Write-Host "   https://www.python.org/downloads/ からインストールしてください。"
    Write-Host "   インストール時に 'Add Python to PATH' にチェックを入れてください。"
    exit 1
}

$pythonVersion = & $pythonCmd --version 2>&1
Write-Host "✅ $pythonVersion" -ForegroundColor Green

# pip パッケージの確認
Write-Host ""
Write-Host "📦 必要なパッケージを確認中..." -ForegroundColor Yellow

$requiredPackages = @("openpyxl", "pyyaml")
$missingPackages = @()

foreach ($pkg in $requiredPackages) {
    $result = & $pythonCmd -c "import $pkg" 2>&1
    if ($LASTEXITCODE -ne 0) {
        $missingPackages += $pkg
    }
}

if ($missingPackages.Count -gt 0) {
    Write-Host "⚠️  不足パッケージ: $($missingPackages -join ', ')" -ForegroundColor Yellow
    $installConfirm = Read-Host "インストールしますか？ (Y/n)"
    if ($installConfirm -ne "n" -and $installConfirm -ne "N") {
        & $pythonCmd -m pip install $missingPackages
        Write-Host "✅ パッケージをインストールしました。" -ForegroundColor Green
    }
} else {
    Write-Host "✅ 必要なパッケージは全てインストール済みです。" -ForegroundColor Green
}

# 完了メッセージ
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "🎉 セットアップ完了！" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📋 次のステップ:" -ForegroundColor White
Write-Host "   1. settings.yaml の内容を確認"
Write-Host "   2. 監視モードを起動:"
Write-Host "      $pythonCmd unified_transcriber.py --watch"
Write-Host ""
Write-Host "📂 フォルダ構成:" -ForegroundColor White
Write-Host "   pending/   - GASからのリクエストJSON"
Write-Host "   processed/ - 処理完了したJSON"
Write-Host "   output/    - 生成されたExcel"
Write-Host ""
