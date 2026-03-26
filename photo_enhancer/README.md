# 写真等倍高画質化アプリ
**DRCT / HAT (Transformer 系超解像モデル) を使用**

---

## 概要

写真の解像度を変えずに、ノイズ除去・シャープネス強化・細部の復元を行う  
「等倍高画質化」アプリです。

内部では 4× 超解像処理を行い、LANCZOS で元サイズに縮小することで  
最高品質の等倍高画質化を実現します。

### 採用モデルについて

| モデル | 発表 | 特徴 |
|--------|------|------|
| **DRCT-L Real-World** | CVPR 2024 ★推奨 | HAT-L を PSNR +0.14dB 上回り、パラメータ 33% 削減 |
| Real-HAT GAN | CVPR 2023 | 実写真向け HAT オリジナル |
| HAT-L SR×4 | CVPR 2023 | HAT 最高品質 (クリーン画像向け) |

---

## セットアップ

### 前提条件

- **Windows 10/11** (64bit)
- **Python 3.10 以上** — https://www.python.org/downloads/
- **NVIDIA GPU** — RTX 3050 LP 6GB (CUDA 12.1 推奨)

### インストール手順

```
1. install.bat をダブルクリック
   ├─ PyTorch (CUDA 12.1) をインストール
   ├─ 依存ライブラリをインストール
   └─ モデルダウンロードを確認

2. モデルを別途ダウンロードする場合:
   download_model.bat をダブルクリック
   または
   python download_model.py

3. run.bat でアプリ起動
```

---

## 使い方

### 基本操作

1. 左側ウィンドウに画像をドラッグ&ドロップ
2. モデルを選択 (DRCT-L Real-World 推奨)
3. 強度スライダーで調整 (100% = フル効果、50% = 元画像とブレンド)
4. **🚀 高画質化** ボタンをクリック
5. 処理完了後、**💾 保存** で保存

### 表示モード

| モード | 説明 |
|--------|------|
| ウィンドウに合わせる | ウィンドウリサイズに追従 |
| 実寸 (1:1) | ドラッグでスクロール |
| 左右並列 | オリジナルと高画質化後を横並び表示 |
| 重ね合わせ比較 | 黄色いハンドルをドラッグして分割比較 |

### ファイル保存

保存ダイアログのデフォルトファイル名:
- 強度 100%: `元ファイル名-drct.jpg`
- 強度 80%:  `元ファイル名-drct-80.jpg`

---

## トラブルシューティング

### GPU メモリ不足 (Out of Memory)

`main.py` の冒頭にある `TILE_SIZE` を小さくしてください:
```python
TILE_SIZE = 256   # または 128
```

### 処理が遅い (CPU モード)

PyTorch の CUDA 版が正しくインストールされていない可能性があります。
```
python -c "import torch; print(torch.cuda.is_available())"
```
`False` の場合は `install.bat` を再実行してください。

### モデルが認識されない

`models/` フォルダに `.pth` ファイルを置き、  
アプリ右上の **↺ ボタン** でリロードしてください。

---

## ファイル構成

```
photo_enhancer/
├─ main.py            # アプリ本体
├─ download_model.py  # モデルダウンローダー
├─ requirements.txt   # 依存ライブラリ
├─ install.bat        # セットアップ (初回のみ)
├─ run.bat            # アプリ起動
├─ README.md          # このファイル
└─ models/            # モデルファイル置き場
   └─ *.pth
```

---

## 参考リンク

- [DRCT GitHub](https://github.com/ming053l/DRCT) — DRCT 論文・モデル
- [HAT GitHub](https://github.com/XPixelGroup/HAT) — HAT 論文・モデル
- [OpenModelDB](https://openmodeldb.info/) — コミュニティモデル集
- [spandrel](https://github.com/chaiNNer-org/spandrel) — SR モデルローダー
