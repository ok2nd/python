#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高画質化モデルのダウンロードスクリプト

【モデル速度比較】
  Real-ESRGAN     (CNN系)         ← 推奨: 数秒で処理。Upscayl と同等モデル
  DRCT-L          (Transformer系) ← 最高画質だが数分かかる
  HAT             (Transformer系) ← 最高画質だが数分かかる

【ダウンロード方法】
  python download_model.py          -- 対話式メニュー
  python download_model.py esrgan   -- Real-ESRGAN のみ (推奨・高速)
  python download_model.py drct     -- DRCT モデルのみ
  python download_model.py all      -- 全モデル
"""

import sys
from pathlib import Path

MODELS_DIR = Path(__file__).parent / "models"

# ──────────────────────────────────────────────────────────────
# モデル定義
# ──────────────────────────────────────────────────────────────
MODELS = {
    # ════════════════════════════════════════════════
    # CNN 系 (高速 — 数秒)  ← Upscayl と同じ種類
    # ════════════════════════════════════════════════
    "realesrgan_x4": {
        "filename": "RealESRGAN_x4plus.pth",
        "urls": [
            "https://huggingface.co/ai-forever/Real-ESRGAN/resolve/main/RealESRGAN_x4plus.pth",
            "https://github.com/xinntao/Real-ESRGAN/releases/download/v0.1.0/RealESRGAN_x4plus.pth",
        ],
        "label":    "Real-ESRGAN x4 (汎用実写真) ★高速推奨",
        "desc":     "Upscayl と同等モデル。実写真に最適。数秒で処理完了",
        "size":     "約 67 MB",
        "scale":    4,
        "speed":    "★★★ 高速 (数秒)",
    },
    "realesrgan_x4_anime": {
        "filename": "RealESRGAN_x4plus_anime_6B.pth",
        "urls": [
            "https://huggingface.co/ai-forever/Real-ESRGAN/resolve/main/RealESRGAN_x4plus_anime_6B.pth",
            "https://github.com/xinntao/Real-ESRGAN/releases/download/v0.2.2.4/RealESRGAN_x4plus_anime_6B.pth",
        ],
        "label":    "Real-ESRGAN x4 Anime (イラスト・アニメ向け)",
        "desc":     "アニメ・イラスト特化。軽量で高速",
        "size":     "約 18 MB",
        "scale":    4,
        "speed":    "★★★ 高速 (数秒・最軽量)",
    },
    "realesrgan_x2": {
        "filename": "RealESRGAN_x2plus.pth",
        "urls": [
            "https://huggingface.co/ai-forever/Real-ESRGAN/resolve/main/RealESRGAN_x2plus.pth",
            "https://github.com/xinntao/Real-ESRGAN/releases/download/v0.2.1/RealESRGAN_x2plus.pth",
        ],
        "label":    "Real-ESRGAN x2 (軽めのノイズ除去)",
        "desc":     "2倍SRで等倍高画質化。よりナチュラルな仕上がり",
        "size":     "約 67 MB",
        "scale":    2,
        "speed":    "★★★ 高速 (数秒)",
    },
    # ════════════════════════════════════════════════
    # Transformer 系 (高画質・低速 — 数分)
    # ════════════════════════════════════════════════
    "drct_real": {
        "filename": "4xRealWebPhoto_v4_drct-l.pth",
        "urls": [
            "https://huggingface.co/ndkhanh95/DRCT/resolve/main/4xRealWebPhoto_v4_drct-l.pth",
            "https://huggingface.co/Phips/4xRealWebPhoto_v4_drct-l/resolve/main/4xRealWebPhoto_v4_drct-l.pth",
        ],
        "label":    "DRCT-L Real-World (最高画質・低速)",
        "desc":     "CVPR 2024 最高PSNR。処理に数分かかる",
        "size":     "約 243 MB",
        "scale":    4,
        "speed":    "★ 低速 (数分)",
    },
    "drct_l_x4": {
        "filename": "DRCT-L_X4.pth",
        "urls": [
            "https://huggingface.co/ndkhanh95/DRCT/resolve/main/DRCT-L_X4.pth",
        ],
        "label":    "DRCT-L SR x4 (クリーン画像・低速)",
        "desc":     "ノイズのない画像のアップスケール特化",
        "size":     "約 486 MB",
        "scale":    4,
        "speed":    "★ 低速 (数分)",
    },
    "hat_real": {
        "filename": "Real_HAT_GAN_sharper.pth",
        "gdrive_id": "1HqKnm8G_JYlPHt_EkQW6tbEoHBXtFlVK",
        "urls":     [],
        "label":    "Real-HAT GAN (最高画質・低速)",
        "desc":     "HAT オリジナル実写真向け",
        "size":     "約 150 MB",
        "scale":    4,
        "speed":    "★ 低速 (数分)",
    },
}


# ──────────────────────────────────────────────────────────────
# ダウンロード関数
# ──────────────────────────────────────────────────────────────
def download_from_url(url: str, dest: Path) -> bool:
    try:
        import requests
        from tqdm import tqdm
    except ImportError:
        print("  pip install requests tqdm が必要です")
        return False

    print(f"  URL: {url}")
    try:
        resp = requests.get(url, stream=True, timeout=60, allow_redirects=True)
        if resp.status_code in (401, 403):
            print(f"  -> 認証エラー ({resp.status_code})、次のミラーを試みます...")
            return False
        resp.raise_for_status()
        total = int(resp.headers.get("content-length", 0))
        with open(dest, "wb") as f, tqdm(
            total=total, unit="B", unit_scale=True, unit_divisor=1024
        ) as bar:
            for chunk in resp.iter_content(chunk_size=65536):
                f.write(chunk)
                bar.update(len(chunk))
        ok = dest.exists() and dest.stat().st_size > 1_000_000
        if not ok:
            print("  -> ファイルサイズ異常")
            if dest.exists(): dest.unlink()
        return ok
    except Exception as e:
        print(f"  -> エラー: {e}")
        if dest.exists(): dest.unlink()
        return False


def download_from_gdrive(gdrive_id: str, dest: Path) -> bool:
    try:
        import gdown
    except ImportError:
        print("  pip install gdown が必要です")
        return False
    print("  Google Drive からダウンロード中...")
    try:
        url = f"https://drive.google.com/uc?id={gdrive_id}"
        gdown.download(url, str(dest), quiet=False)
        return dest.exists() and dest.stat().st_size > 1_000_000
    except Exception as e:
        print(f"  -> エラー: {e}")
        return False


def download_model(key: str) -> bool:
    MODELS_DIR.mkdir(parents=True, exist_ok=True)
    info = MODELS[key]
    dest = MODELS_DIR / info["filename"]

    if dest.exists():
        size_mb = dest.stat().st_size / 1024**2
        print(f"  既存: {dest.name} ({size_mb:.0f} MB) — スキップ")
        return True

    print(f"\n{'─'*58}")
    print(f"  {info['label']}")
    print(f"  {info['desc']}")
    print(f"  速度: {info.get('speed','')}")
    print(f"  保存先: {dest}")
    print(f"{'─'*58}")

    ok = False
    for i, url in enumerate(info.get("urls", []), 1):
        print(f"\n  [試行 {i}] HTTP ダウンロード...")
        ok = download_from_url(url, dest)
        if ok:
            break

    if not ok and info.get("gdrive_id"):
        print(f"\n  [Google Drive] ダウンロード中...")
        ok = download_from_gdrive(info["gdrive_id"], dest)

    if ok:
        size_mb = dest.stat().st_size / 1024**2
        print(f"\n  完了: {dest.name} ({size_mb:.0f} MB)")
    else:
        print(f"\n  自動ダウンロード失敗。手動でダウンロードしてください。")
        _print_manual_instructions(key, info, dest)

    return ok


def _print_manual_instructions(key: str, info: dict, dest: Path):
    print()
    print("  ━━━ 手動ダウンロード手順 ━━━")
    urls = info.get("urls", [])
    if urls:
        print(f"  1. ブラウザで開く: {urls[0]}")
    elif info.get("gdrive_id"):
        print(f"  1. ブラウザで開く: https://drive.google.com/file/d/{info['gdrive_id']}/view")
    print(f"  2. models/ フォルダに置く: {dest}")
    print()


# ──────────────────────────────────────────────────────────────
# メイン
# ──────────────────────────────────────────────────────────────
def interactive_menu():
    print()
    print("=" * 62)
    print("  高画質化モデルダウンローダー")
    print("=" * 62)
    print()
    print("  ★ 高速モデル (Real-ESRGAN / CNN系) — Upscayl と同じ種類")
    print("  ☆ 高画質モデル (DRCT・HAT / Transformer系) — 数分かかる")
    print()

    items = list(MODELS.items())
    for i, (k, v) in enumerate(items, 1):
        existing = "[済]" if (MODELS_DIR / v["filename"]).exists() else "    "
        print(f"  [{i}] {existing} {v['label']}")
        print(f"          {v['desc']}")
        print(f"          速度: {v.get('speed','')}  /  サイズ: {v['size']}")
        print()

    print("  [0] CNN系3つをまとめてダウンロード (推奨)")
    print()

    choice = input("  番号を入力 (推奨: 1 Real-ESRGAN x4): ").strip()

    if choice == "0":
        for k in ["realesrgan_x4", "realesrgan_x4_anime", "realesrgan_x2"]:
            download_model(k)
    else:
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(items):
                download_model(items[idx][0])
            else:
                print("無効な番号です。")
        except ValueError:
            print("数字を入力してください。")


def main():
    args = sys.argv[1:]
    if not args:
        interactive_menu()
        print()
        input("Enter キーで終了...")
        return

    cmd = args[0].lower()
    if cmd == "all":
        for k in MODELS:
            download_model(k)
    elif cmd == "esrgan":
        for k in ["realesrgan_x4", "realesrgan_x4_anime", "realesrgan_x2"]:
            download_model(k)
    elif cmd in ("drct", "drct_real"):
        download_model("drct_real")
    elif cmd == "hat":
        download_model("hat_real")
    else:
        print(f"不明なコマンド: {cmd}")

    print()
    input("Enter キーで終了...")


if __name__ == "__main__":
    main()
