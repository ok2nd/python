@echo off
chcp 65001 > nul
echo ============================================================
echo   比率指定トリミング  セットアップ＆起動
echo ============================================================
echo.

:: Python確認
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Python が見つかりません。
    echo   https://www.python.org/ からインストールしてください。
    pause
    exit /b 1
)

:: ライブラリインストール
echo [1/2] 必要なライブラリを確認・インストールします...
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo [エラー] ライブラリのインストールに失敗しました。
    pause
    exit /b 1
)

echo [2/2] アプリを起動します...
echo.
python ratio_trim.py
pause
