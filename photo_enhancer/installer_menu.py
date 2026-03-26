#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
install.bat から呼び出されるメニュー。
モデルをダウンロードする場合は sys.exit(1) で通知。
"""
import sys

print("=" * 52)
print("  セットアップ完了！")
print("=" * 52)
print()
print("  モデルをダウンロードしますか？")
print("  (後で download_model.bat でも可)")
print()
ans = input("  Y でダウンロード開始 / それ以外でスキップ: ").strip().lower()
if ans == "y":
    sys.exit(1)   # bat 側で errorlevel 1 → download
else:
    print()
    print("  スキップしました。download_model.bat で後でダウンロードできます。")
    sys.exit(0)
