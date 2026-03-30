@echo off
chcp 65001 > nul
echo 必要なライブラリをインストールします...
pip install PySide6 Pillow piexif
echo.
echo インストール完了。run.bat でアプリを起動してください。
pause
