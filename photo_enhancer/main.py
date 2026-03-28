#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
写真等倍高画質化アプリ
DRCT (Dense Residual Connected Transformer) / HAT (Hybrid Attention Transformer)
CVPR 2024 DRCT は HAT-L を PSNR +0.14dB 上回る最新モデル

使い方:
  1. install.bat でセットアップ
  2. download_model.py でモデルをダウンロード
  3. run.bat またはこのスクリプトを直接実行
"""

import sys
import os
import math
from pathlib import Path
from typing import Optional

import numpy as np
from PIL import Image

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QSlider, QFileDialog, QSplitter,
    QProgressBar, QStatusBar, QGroupBox, QRadioButton, QComboBox,
    QFrame, QSizePolicy, QMessageBox, QStackedWidget, QToolTip
)
from PySide6.QtCore import Qt, QThread, Signal, QRect, QPoint, QTimer, QSize
from PySide6.QtGui import (
    QPixmap, QImage, QPainter, QColor, QFont, QPen,
    QDragEnterEvent, QDropEvent, QLinearGradient, QBrush, QPalette
)

# ============================================================
# 定数・設定
# ============================================================
APP_TITLE   = "写真等倍高画質化"
APP_VERSION = "1.0"
MODELS_DIR  = Path(__file__).parent / "models"
WINDOW_SIZE = 16    # Transformerアテンションウィンドウサイズ (16の倍数にパディング)
# Real-ESRGAN (CNN) : 1024 推奨 → 1200x1200 画像を 4 タイルで処理
# DRCT/HAT (Transformer): 768 推奨 → VRAM 約 4GB
TILE_SIZE   = 1024  # CNN モデルは大きいほど速い (VRAM 約 2GB)
# TILE_SIZE = 768   # DRCT/HAT で VRAM 不足の場合
# TILE_SIZE = 512   # VRAM 不足時
# TILE_SIZE = 256   # VRAM 不足時 (最小)
USE_FP16    = False # fp16 は DRCT/HAT 非対応。CNN も float32 で十分速い


# ============================================================
# 高画質化処理スレッド (DRCT / HAT 共通)
# ============================================================
class EnhancerThread(QThread):
    """
    spandrel ライブラリで DRCT/HAT を自動認識・実行するスレッド
    4xSR → 元サイズにLANCZOS縮小 = 等倍高画質化
    """
    progressChanged = Signal(int)   # 0〜100
    statusChanged   = Signal(str)
    finished        = Signal(object)   # PIL.Image
    errorOccurred   = Signal(str)

    def __init__(self, image: Image.Image, model_path: Path, intensity: float, parent=None):
        super().__init__(parent)
        self.image      = image
        self.model_path = model_path
        self.intensity  = intensity   # 0.0〜1.0

    # ── メイン処理 ──────────────────────────────
    def run(self):
        try:
            import torch
            import torch.nn.functional as F

            # cuDNN オートチューナー有効化 (同サイズ2回目以降が高速化)
            torch.backends.cudnn.benchmark = True

            # ---------- モデルロード ----------
            self.statusChanged.emit("モデルを読み込み中...")
            self.progressChanged.emit(5)

            try:
                import spandrel
            except ImportError:
                self.errorOccurred.emit(
                    "spandrel がインストールされていません。\n"
                    "install.bat を実行してください。"
                )
                return

            device  = torch.device("cuda" if torch.cuda.is_available() else "cpu")
            use_fp16 = USE_FP16 and device.type == "cuda"
            prec_str = "fp16" if use_fp16 else "fp32"
            dev_str  = f"CUDA ({torch.cuda.get_device_name(0)})" if device.type == "cuda" else "CPU"
            self.statusChanged.emit(f"デバイス: {dev_str} [{prec_str}] | モデル読み込み中...")

            descriptor = spandrel.ModelLoader(device=device).load_from_file(str(self.model_path))
            model = descriptor.model
            scale = getattr(descriptor, "scale", 4)
            model.eval()

            # USE_FP16=False のため fp32 のまま使用 (DRCT/HAT は fp16 非対応)
            model = model.float()

            self.progressChanged.emit(15)
            orig_w, orig_h = self.image.size
            self.statusChanged.emit(
                f"前処理中... (入力 {orig_w}×{orig_h} → ×{scale}SR → 等倍)"
            )

            # ---------- 前処理 ----------
            img_rgb = self.image.convert("RGB")
            img_np  = np.array(img_rgb).astype(np.float32) / 255.0

            # [H,W,C] → [1,C,H,W]
            img_t = torch.from_numpy(img_np.transpose(2, 0, 1)).unsqueeze(0).to(device)
            # img_t は float32 のまま

            self.progressChanged.emit(20)
            self.statusChanged.emit(
                f"高画質化処理中 (タイル {TILE_SIZE}px / {prec_str})..."
            )

            # ---------- タイル処理 ----------
            result_t = self._tile_process(model, img_t, scale, device, F)

            self.progressChanged.emit(88)
            self.statusChanged.emit("後処理中 (等倍リサイズ)...")

            # [1,C,Hs,Ws] → PIL  (.float() で fp32 に戻してから numpy 変換)
            result_np = result_t.squeeze(0).cpu().float().numpy().transpose(1, 2, 0).clip(0, 1)
            sr_img    = Image.fromarray((result_np * 255).astype(np.uint8))

            # 等倍: 元解像度に LANCZOS で縮小
            sr_resized = sr_img.resize((orig_w, orig_h), Image.LANCZOS)

            # 強度ブレンド
            if self.intensity < 1.0:
                sr_resized = Image.blend(img_rgb, sr_resized, self.intensity)

            self.progressChanged.emit(100)
            self.statusChanged.emit("完了")
            self.finished.emit(sr_resized)

        except torch.cuda.OutOfMemoryError:
            self.errorOccurred.emit(
                "GPU メモリ不足です。\n\n"
                f"現在の設定: TILE_SIZE={TILE_SIZE}  USE_FP16={USE_FP16}\n\n"
                "【対処方法】main.py 冒頭の値を変更してください:\n"
                "  TILE_SIZE = 512  (または 256)\n"
                "  USE_FP16  = True  (既に True の場合はさらに TILE_SIZE を縮小)"
            )
        except Exception as e:
            import traceback
            self.errorOccurred.emit(f"処理エラー:\n{str(e)}\n\n{traceback.format_exc()}")

    # ── タイル処理 ───────────────────────────────
    def _tile_process(self, model, img_t, scale, device, F):
        import torch

        # 入力を常に float32 に統一 (DRCT/HAT は fp16 非対応)
        img_t = img_t.float()

        _, C, H, W = img_t.shape

        # WINDOW_SIZE の倍数にパディング
        pad_h = (WINDOW_SIZE - H % WINDOW_SIZE) % WINDOW_SIZE
        pad_w = (WINDOW_SIZE - W % WINDOW_SIZE) % WINDOW_SIZE
        if pad_h or pad_w:
            img_t = F.pad(img_t, (0, pad_w, 0, pad_h), mode="reflect")
        _, _, Hp, Wp = img_t.shape

        # 画像が小さければ一括処理
        if Hp <= TILE_SIZE and Wp <= TILE_SIZE:
            with torch.no_grad():
                out = model(img_t)
            return out[:, :, :H * scale, :W * scale].float()

        # ---------- タイル分割処理 ----------
        tiles_h = math.ceil(Hp / TILE_SIZE)
        tiles_w = math.ceil(Wp / TILE_SIZE)
        total   = tiles_h * tiles_w

        # 結果バッファは CPU (システム RAM) に float32 で確保 → VRAM を節約
        output = torch.zeros(1, C, Hp * scale, Wp * scale, dtype=torch.float32)

        count = 0
        for th in range(tiles_h):
            for tw in range(tiles_w):
                y0 = th * TILE_SIZE
                x0 = tw * TILE_SIZE
                y1 = min(y0 + TILE_SIZE, Hp)
                x1 = min(x0 + TILE_SIZE, Wp)
                ph, pw = y1 - y0, x1 - x0

                patch = img_t[:, :, y0:y1, x0:x1]

                # エッジタイルをフルサイズにパディング
                if ph < TILE_SIZE or pw < TILE_SIZE:
                    patch = F.pad(patch, (0, TILE_SIZE - pw, 0, TILE_SIZE - ph), "reflect").float()

                with torch.no_grad():
                    out_patch = model(patch)

                # タイル完了後すぐ CPU (float32) に移動して VRAM を解放
                oy0, ox0 = y0 * scale, x0 * scale
                oy1 = oy0 + ph * scale
                ox1 = ox0 + pw * scale
                output[:, :, oy0:oy1, ox0:ox1] = (
                    out_patch[:, :, :ph * scale, :pw * scale].cpu().float()
                )
                del out_patch, patch
                torch.cuda.empty_cache()

                count += 1
                self.progressChanged.emit(20 + int(68 * count / total))

        return output[:, :, :H * scale, :W * scale]


# ============================================================
# ドラッグ&ドロップ対応 画像表示ウィジェット
# ============================================================
class ImageDropWidget(QWidget):
    """画像ドロップ + フィット/実寸/パン操作"""

    imageDropped = Signal(str)

    ACCEPTED = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp"}

    def __init__(self, placeholder: str = "ここに画像をドロップ", parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self._pixmap:    Optional[QPixmap] = None
        self._fit_mode:  bool = True
        self._placeholder = placeholder
        self._pan_start: Optional[QPoint] = None
        self._pan_offset = QPoint(0, 0)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumSize(200, 150)

    # ── 公開メソッド ─────────────────────────────
    def setPixmap(self, pixmap: QPixmap):
        self._pixmap     = pixmap
        self._pan_offset = QPoint(0, 0)
        self.update()

    def setFitMode(self, fit: bool):
        self._fit_mode   = fit
        self._pan_offset = QPoint(0, 0)
        self.update()

    def clear(self):
        self._pixmap = None
        self.update()

    # ── 描画 ─────────────────────────────────────
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        w, h = self.width(), self.height()

        # 背景グリッド
        p.fillRect(0, 0, w, h, QColor(16, 16, 26))
        grid_color = QColor(26, 26, 40)
        p.setPen(QPen(grid_color, 1))
        step = 24
        for x in range(0, w, step):
            p.drawLine(x, 0, x, h)
        for y in range(0, h, step):
            p.drawLine(0, y, w, y)

        if self._pixmap is None:
            # ドロップゾーン
            p.setPen(QPen(QColor(60, 80, 160, 180), 2, Qt.DashLine))
            p.drawRoundedRect(16, 16, w - 32, h - 32, 10, 10)
            p.setPen(QColor(120, 140, 210))
            p.setFont(QFont("Yu Gothic UI", 11))
            p.drawText(QRect(0, 0, w, h), Qt.AlignCenter, self._placeholder)
        else:
            pw, ph = self._pixmap.width(), self._pixmap.height()
            if self._fit_mode:
                s = min(w / pw, h / ph)
                dw, dh = int(pw * s), int(ph * s)
                dx, dy = (w - dw) // 2, (h - dh) // 2
            else:
                dw, dh = pw, ph
                dx = (w - dw) // 2 + self._pan_offset.x()
                dy = (h - dh) // 2 + self._pan_offset.y()
            p.drawPixmap(QRect(dx, dy, dw, dh), self._pixmap)
        p.end()

    # ── ドラッグ&ドロップ ─────────────────────────
    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls():
            for url in e.mimeData().urls():
                if Path(url.toLocalFile()).suffix.lower() in self.ACCEPTED:
                    e.acceptProposedAction()
                    return

    def dropEvent(self, e: QDropEvent):
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if Path(p).suffix.lower() in self.ACCEPTED:
                self.imageDropped.emit(p)
                break

    # ── パン操作 (実寸モード) ─────────────────────
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton and not self._fit_mode:
            self._pan_start = e.position().toPoint()
            self.setCursor(Qt.ClosedHandCursor)

    def mouseMoveEvent(self, e):
        if self._pan_start is not None:
            self._pan_offset += e.position().toPoint() - self._pan_start
            self._pan_start   = e.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, _):
        self._pan_start = None
        self.setCursor(Qt.ArrowCursor)


# ============================================================
# 重ね合わせ比較ウィジェット (スプリットビュー)
# ============================================================
class OverlayWidget(QWidget):
    """ドラッグで分割線を動かして2枚を比較"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._orig_px:  Optional[QPixmap] = None
        self._enh_px:   Optional[QPixmap] = None
        self._ratio:    float = 0.5
        self._dragging: bool  = False
        self._fit_mode: bool  = True
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumSize(200, 150)
        self.setMouseTracking(True)

    def setImages(self, orig: Optional[QPixmap], enh: Optional[QPixmap]):
        self._orig_px = orig
        self._enh_px  = enh
        self.update()

    def setFitMode(self, fit: bool):
        self._fit_mode = fit
        self.update()

    def _calc_rect(self, px: QPixmap) -> QRect:
        w, h  = self.width(), self.height()
        pw, ph = px.width(), px.height()
        if self._fit_mode:
            s  = min(w / pw, h / ph)
            dw, dh = int(pw * s), int(ph * s)
        else:
            dw, dh = pw, ph
        return QRect((w - dw) // 2, (h - dh) // 2, dw, dh)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        w, h = self.width(), self.height()

        # 背景
        p.fillRect(0, 0, w, h, QColor(16, 16, 26))
        grid_color = QColor(26, 26, 40)
        p.setPen(QPen(grid_color, 1))
        for x in range(0, w, 24): p.drawLine(x, 0, x, h)
        for y in range(0, h, 24): p.drawLine(0, y, w, y)

        if not self._orig_px and not self._enh_px:
            p.setPen(QColor(120, 140, 210))
            p.setFont(QFont("Yu Gothic UI", 11))
            p.drawText(QRect(0, 0, w, h), Qt.AlignCenter,
                       "画像を読み込んで高画質化してください")
            p.end()
            return

        split_x = int(w * self._ratio)

        # 左: オリジナル
        if self._orig_px:
            r = self._calc_rect(self._orig_px)
            p.save()
            p.setClipRect(0, 0, split_x, h)
            p.drawPixmap(r, self._orig_px)
            p.restore()

        # 右: 高画質化後
        if self._enh_px:
            r = self._calc_rect(self._enh_px)
            p.save()
            p.setClipRect(split_x, 0, w - split_x, h)
            p.drawPixmap(r, self._enh_px)
            p.restore()

        # 分割線
        p.setPen(QPen(QColor(255, 220, 40, 230), 2))
        p.drawLine(split_x, 0, split_x, h)

        # ハンドル
        hr = 14
        cx, cy = split_x, h // 2
        p.setBrush(QColor(255, 220, 40, 220))
        p.setPen(QPen(QColor(0, 0, 0, 80), 1))
        p.drawEllipse(cx - hr, cy - hr, hr * 2, hr * 2)
        p.setPen(QPen(QColor(0, 0, 0, 180), 2))
        p.drawLine(cx - 6, cy, cx + 6, cy)
        p.drawLine(cx - 3, cy - 4, cx - 6, cy)
        p.drawLine(cx - 3, cy + 4, cx - 6, cy)
        p.drawLine(cx + 3, cy - 4, cx + 6, cy)
        p.drawLine(cx + 3, cy + 4, cx + 6, cy)

        # ラベル
        p.setFont(QFont("Yu Gothic UI", 9, QFont.Bold))
        bg = QColor(0, 0, 0, 150)
        if split_x > 90:
            p.fillRect(8, 8, 84, 22, bg)
            p.setPen(QColor(180, 200, 255))
            p.drawText(QRect(8, 8, 84, 22), Qt.AlignCenter, "オリジナル")
        if split_x < w - 90:
            p.fillRect(split_x + 8, 8, 84, 22, bg)
            p.setPen(QColor(150, 255, 180))
            p.drawText(QRect(split_x + 8, 8, 84, 22), Qt.AlignCenter, "高画質化後")
        p.end()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._dragging = True

    def mouseMoveEvent(self, e):
        x = e.position().x()
        split_x = int(self.width() * self._ratio)
        self.setCursor(Qt.SplitHCursor if abs(x - split_x) < 24 else Qt.ArrowCursor)
        if self._dragging:
            self._ratio = max(0.02, min(0.98, x / self.width()))
            self.update()

    def mouseReleaseEvent(self, _):
        self._dragging = False


# ============================================================
# メインウィンドウ
# ============================================================
class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE}  v{APP_VERSION}")
        self.resize(1300, 840)

        self._orig_image:  Optional[Image.Image] = None
        self._enh_image:   Optional[Image.Image] = None
        self._orig_path:   Optional[str]         = None
        self._processor:   Optional[EnhancerThread] = None

        self._build_ui()
        self._apply_style()
        self._check_env()

    # ─────────────────────────────────────────────────────────
    # UI 構築
    # ─────────────────────────────────────────────────────────
    def _build_ui(self):
        root_w = QWidget()
        self.setCentralWidget(root_w)
        root = QVBoxLayout(root_w)
        root.setContentsMargins(10, 8, 10, 6)
        root.setSpacing(6)

        root.addWidget(self._build_toolbar())

        # ── 画像エリア (StackedWidget) ──
        self._stack = QStackedWidget()

        # Page 0: 左右並列
        side_w = QWidget()
        side_l = QHBoxLayout(side_w)
        side_l.setContentsMargins(0, 0, 0, 0)
        side_l.setSpacing(4)
        self._splitter = QSplitter(Qt.Horizontal)

        grp_orig = QGroupBox("オリジナル")
        grp_orig.setObjectName("grpOrig")
        ll = QVBoxLayout(grp_orig)
        ll.setContentsMargins(4, 16, 4, 4)
        self._drop_widget = ImageDropWidget(
            "📂  ここに画像をドロップ\n\n"
            "JPEG / PNG / BMP / TIFF / WebP"
        )
        self._drop_widget.imageDropped.connect(self._load_image)
        ll.addWidget(self._drop_widget)

        grp_enh = QGroupBox("高画質化後")
        grp_enh.setObjectName("grpEnh")
        rl = QVBoxLayout(grp_enh)
        rl.setContentsMargins(4, 16, 4, 4)
        self._enh_widget = ImageDropWidget("高画質化後の画像がここに表示されます")
        rl.addWidget(self._enh_widget)

        self._splitter.addWidget(grp_orig)
        self._splitter.addWidget(grp_enh)
        self._splitter.setSizes([650, 650])
        side_l.addWidget(self._splitter)
        self._stack.addWidget(side_w)   # index 0

        # Page 1: 重ね合わせ
        grp_ov = QGroupBox("比較ビュー  —  黄色いハンドルをドラッグして比較")
        grp_ov.setObjectName("grpOverlay")
        ol = QVBoxLayout(grp_ov)
        ol.setContentsMargins(4, 16, 4, 4)
        self._overlay = OverlayWidget()
        ol.addWidget(self._overlay)
        self._stack.addWidget(grp_ov)   # index 1

        root.addWidget(self._stack, 1)

        # プログレスバー
        self._progress = QProgressBar()
        self._progress.setVisible(False)
        self._progress.setFixedHeight(12)
        root.addWidget(self._progress)

        # ステータスバー
        sb = QStatusBar()
        self.setStatusBar(sb)
        self._lbl_status = QLabel("起動中...")
        sb.addWidget(self._lbl_status, 1)
        self._lbl_gpu = QLabel("")
        self._lbl_gpu.setObjectName("gpuLabel")
        sb.addPermanentWidget(self._lbl_gpu)

    def _build_toolbar(self) -> QFrame:
        """上部ツールバー"""
        bar = QFrame()
        bar.setObjectName("toolbar")
        bar.setFixedHeight(106)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(10, 6, 10, 6)
        lay.setSpacing(10)

        # ── モデル選択 ──
        grp_model = QGroupBox("モデル選択")
        ml = QVBoxLayout(grp_model)
        ml.setContentsMargins(8, 4, 8, 6)
        ml.setSpacing(3)
        self._model_combo = QComboBox()
        self._model_combo.setMinimumWidth(260)
        self._model_combo.setObjectName("modelCombo")
        self._refresh_models()
        self._btn_reload = QPushButton("↺")
        self._btn_reload.setFixedSize(26, 26)
        self._btn_reload.setToolTip("models/ フォルダを再スキャン")
        self._btn_reload.clicked.connect(self._refresh_models)
        row_m = QHBoxLayout()
        row_m.addWidget(self._model_combo)
        row_m.addWidget(self._btn_reload)
        ml.addLayout(row_m)
        hint_m = QLabel("models/ フォルダに .pth ファイルを置く")
        hint_m.setObjectName("hintLbl")
        ml.addWidget(hint_m)
        lay.addWidget(grp_model)

        # ── 表示スケール ──
        grp_scale = QGroupBox("表示スケール")
        sl = QVBoxLayout(grp_scale)
        sl.setContentsMargins(8, 4, 8, 4)
        sl.setSpacing(2)
        self._rb_fit    = QRadioButton("ウィンドウに合わせる")
        self._rb_actual = QRadioButton("実寸 (1:1)  ドラッグでスクロール")
        self._rb_fit.setChecked(True)
        self._rb_fit.toggled.connect(self._on_scale_changed)
        sl.addWidget(self._rb_fit)
        sl.addWidget(self._rb_actual)
        lay.addWidget(grp_scale)

        # ── 比較モード ──
        grp_cmp = QGroupBox("比較モード")
        cl = QVBoxLayout(grp_cmp)
        cl.setContentsMargins(8, 4, 8, 4)
        cl.setSpacing(2)
        self._rb_side    = QRadioButton("左右に並べる")
        self._rb_overlay = QRadioButton("重ね合わせ比較")
        self._rb_side.setChecked(True)
        self._rb_side.toggled.connect(self._on_compare_changed)
        cl.addWidget(self._rb_side)
        cl.addWidget(self._rb_overlay)
        lay.addWidget(grp_cmp)

        # ── 強度スライダー ──
        grp_int = QGroupBox("高画質化強度")
        il = QVBoxLayout(grp_int)
        il.setContentsMargins(8, 4, 8, 4)
        il.setSpacing(3)
        row_i = QHBoxLayout()
        self._slider = QSlider(Qt.Horizontal)
        self._slider.setRange(0, 100)
        self._slider.setValue(100)
        self._slider.setMinimumWidth(170)
        self._lbl_int = QLabel("100%")
        self._lbl_int.setFixedWidth(38)
        self._lbl_int.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self._slider.valueChanged.connect(lambda v: self._lbl_int.setText(f"{v}%"))
        row_i.addWidget(self._slider)
        row_i.addWidget(self._lbl_int)
        il.addLayout(row_i)
        h_int = QLabel("← 元画像に近い      最大強度 →")
        h_int.setObjectName("hintLbl")
        il.addWidget(h_int)
        lay.addWidget(grp_int)

        # ── ボタン ──
        grp_btn = QGroupBox("操作")
        bl = QVBoxLayout(grp_btn)
        bl.setContentsMargins(8, 4, 8, 4)
        bl.setSpacing(5)
        self._btn_enhance = QPushButton("🚀  高画質化")
        self._btn_enhance.setObjectName("btnEnhance")
        self._btn_enhance.setEnabled(False)
        self._btn_enhance.setFixedHeight(32)
        self._btn_enhance.clicked.connect(self._start_enhance)

        self._btn_save = QPushButton("💾  保存")
        self._btn_save.setObjectName("btnSave")
        self._btn_save.setEnabled(False)
        self._btn_save.setFixedHeight(32)
        self._btn_save.clicked.connect(self._save)

        bl.addWidget(self._btn_enhance)
        bl.addWidget(self._btn_save)
        lay.addWidget(grp_btn)

        lay.addStretch(1)
        return bar

    # ─────────────────────────────────────────────────────────
    # 初期化・環境チェック
    # ─────────────────────────────────────────────────────────
    def _check_env(self):
        try:
            import torch
            if torch.cuda.is_available():
                name = torch.cuda.get_device_name(0)
                mem  = torch.cuda.get_device_properties(0).total_memory / 1024**3
                self._lbl_gpu.setText(f"GPU: {name}  ({mem:.1f} GB VRAM)")
                self._lbl_gpu.setStyleSheet("color:#55ee88; padding:0 10px;")
            else:
                self._lbl_gpu.setText("GPU 未検出 — CPU モード (低速)")
                self._lbl_gpu.setStyleSheet("color:#ffaa33; padding:0 10px;")
        except ImportError:
            self._lbl_gpu.setText("PyTorch 未インストール")
            self._lbl_gpu.setStyleSheet("color:#ff5555; padding:0 10px;")

        model_files = list(MODELS_DIR.glob("*.pth")) if MODELS_DIR.exists() else []
        if model_files:
            self._lbl_status.setText(
                f"✓ モデル {len(model_files)} 件検出 | 画像をドロップして開始"
            )
        else:
            self._lbl_status.setText(
                "⚠ models/ にモデルなし — download_model.py を実行してください"
            )

    def _refresh_models(self):
        """models/ フォルダをスキャンしてコンボボックスを更新"""
        # モデル名 → 速度ラベル のマッピング
        SPEED_TAGS = {
            "realesrgan": "⚡ 高速",
            "esrgan":     "⚡ 高速",
            "swinir-l":   "🚶 中速",
            "swinir-m":   "🚶 中速",
            "swinir":     "🚶 中速",
            "drct":       "🐢 低速",
            "hat":        "🐢 低速",
        }
        self._model_combo.clear()
        MODELS_DIR.mkdir(parents=True, exist_ok=True)
        pth_files = sorted(MODELS_DIR.glob("*.pth"))
        if pth_files:
            for f in pth_files:
                tag = ""
                name_lower = f.name.lower()
                for key, label in SPEED_TAGS.items():
                    if key in name_lower:
                        tag = f"  [{label}]"
                        break
                self._model_combo.addItem(f"{f.name}{tag}", userData=str(f))
        else:
            self._model_combo.addItem("モデルなし — download_model.py を実行")

    def _get_model_path(self) -> Optional[Path]:
        idx = self._model_combo.currentIndex()
        data = self._model_combo.itemData(idx)
        if data and Path(data).exists():
            return Path(data)
        return None

    # ─────────────────────────────────────────────────────────
    # 画像操作
    # ─────────────────────────────────────────────────────────
    def _load_image(self, path: str):
        try:
            img = Image.open(path)
            self._orig_image = img.copy()
            self._orig_path  = path
            self._enh_image  = None

            px = _pil_to_qpixmap(img)
            self._drop_widget.setPixmap(px)
            self._enh_widget.clear()
            self._overlay.setImages(px, None)

            w, h = img.size
            kb   = os.path.getsize(path) / 1024
            self._lbl_status.setText(
                f"読込完了: {Path(path).name}  [{w}×{h} px  {img.mode}  {kb:.0f} KB]"
            )
            self._btn_enhance.setEnabled(True)
            self._btn_save.setEnabled(False)
        except Exception as e:
            QMessageBox.critical(self, "読み込みエラー", f"画像を開けませんでした:\n{e}")

    # ─────────────────────────────────────────────────────────
    # 高画質化
    # ─────────────────────────────────────────────────────────
    def _start_enhance(self):
        if self._orig_image is None:
            return

        model_path = self._get_model_path()
        if model_path is None:
            QMessageBox.warning(
                self, "モデル未選択",
                "models/ フォルダにモデルファイル (.pth) がありません。\n"
                "download_model.py を実行してください。"
            )
            return

        intensity = self._slider.value() / 100.0
        self._btn_enhance.setEnabled(False)
        self._btn_save.setEnabled(False)
        self._progress.setVisible(True)
        self._progress.setValue(0)

        self._processor = EnhancerThread(self._orig_image, model_path, intensity)
        self._processor.progressChanged.connect(self._progress.setValue)
        self._processor.statusChanged.connect(self._lbl_status.setText)
        self._processor.finished.connect(self._on_done)
        self._processor.errorOccurred.connect(self._on_error)
        self._processor.start()

    def _on_done(self, result: Image.Image):
        self._enh_image = result
        orig_px = _pil_to_qpixmap(self._orig_image)
        enh_px  = _pil_to_qpixmap(result)
        self._enh_widget.setPixmap(enh_px)
        self._overlay.setImages(orig_px, enh_px)
        self._progress.setVisible(False)
        self._btn_enhance.setEnabled(True)
        self._btn_save.setEnabled(True)
        w, h = result.size
        v = self._slider.value()
        self._lbl_status.setText(
            f"✓ 高画質化完了  [{w}×{h} px]  強度: {v}%  |  "
            f"モデル: {self._model_combo.currentText()}"
        )

    def _on_error(self, msg: str):
        self._progress.setVisible(False)
        self._btn_enhance.setEnabled(True)
        QMessageBox.critical(self, "処理エラー", msg)
        self._lbl_status.setText("⚠ 処理エラー")

    # ─────────────────────────────────────────────────────────
    # 保存
    # ─────────────────────────────────────────────────────────
    def _save(self):
        if self._enh_image is None or self._orig_path is None:
            return

        v     = self._slider.value()
        orig  = Path(self._orig_path)
        stem, suf = orig.stem, orig.suffix.lower()

        # モデル名から短縮タグを生成 (例: Real_HAT_GAN_sharper → hat)
        model_name = self._model_combo.currentText().lower()
        if "realesrgan" in model_name and "anime" in model_name:
            tag = "esrgan-anime"
        elif "realesrgan" in model_name and "x2" in model_name:
            tag = "esrgan-x2"
        elif "realesrgan" in model_name:
            tag = "esrgan"
        elif "drct" in model_name:
            tag = "drct"
        elif "hat" in model_name:
            tag = "hat"
        else:
            # .pth ファイル名のステム先頭8文字
            raw = self._model_combo.itemData(self._model_combo.currentIndex()) or model_name
            tag = Path(raw).stem[:12]

        default = f"{stem}-{tag}{suf}" if v == 100 else f"{stem}-{tag}-{v}{suf}"

        path, _ = QFileDialog.getSaveFileName(
            self,
            "高画質化画像を保存",
            str(orig.parent / default),
            "JPEG (*.jpg *.jpeg);;PNG (*.png);;WebP (*.webp);;TIFF (*.tif *.tiff);;BMP (*.bmp)"
        )
        if not path:
            return

        try:
            ext = Path(path).suffix.lower()
            img = self._enh_image.convert("RGB")
            if ext in (".jpg", ".jpeg"):
                img.save(path, "JPEG", quality=95, subsampling=0)
            elif ext == ".png":
                img.save(path, "PNG", compress_level=6)
            elif ext == ".webp":
                img.save(path, "WEBP", quality=95, method=6)
            elif ext in (".tif", ".tiff"):
                img.save(path, "TIFF", compression="lzw")
            else:
                img.save(path)
            self._lbl_status.setText(f"✓ 保存完了: {path}")
        except Exception as e:
            QMessageBox.critical(self, "保存エラー", f"保存に失敗しました:\n{e}")

    # ─────────────────────────────────────────────────────────
    # UI イベント
    # ─────────────────────────────────────────────────────────
    def _on_scale_changed(self):
        fit = self._rb_fit.isChecked()
        self._drop_widget.setFitMode(fit)
        self._enh_widget.setFitMode(fit)
        self._overlay.setFitMode(fit)

    def _on_compare_changed(self):
        self._stack.setCurrentIndex(0 if self._rb_side.isChecked() else 1)

    # ─────────────────────────────────────────────────────────
    # スタイルシート
    # ─────────────────────────────────────────────────────────
    def _apply_style(self):
        self.setStyleSheet("""
        QMainWindow, QWidget {
            background: #0f0f1a;
            color: #c4d4f0;
            font-family: "Yu Gothic UI", "Meiryo UI", "MS Gothic", sans-serif;
            font-size: 10pt;
        }
        QFrame#toolbar {
            background: #181828;
            border: 1px solid #252545;
            border-radius: 6px;
        }
        QGroupBox {
            border: 1px solid #252545;
            border-radius: 6px;
            margin-top: 12px;
            padding: 4px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px;
            color: #7080b0;
            font-size: 9pt;
        }
        QGroupBox#grpOrig::title  { color: #6688ee; font-weight: bold; }
        QGroupBox#grpEnh::title   { color: #44cc88; font-weight: bold; }
        QGroupBox#grpOverlay::title { color: #ddbb44; font-weight: bold; }

        QPushButton#btnEnhance {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                stop:0 #1a3a9a, stop:1 #1a6a9a);
            color: #e8f0ff;
            border: 1px solid #2a5acc;
            border-radius: 5px;
            font-weight: bold;
            font-size: 11pt;
            padding: 2px 10px;
        }
        QPushButton#btnEnhance:hover  {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                stop:0 #2a4acc, stop:1 #2a8acc);
        }
        QPushButton#btnEnhance:pressed { background: #1a2a7a; }
        QPushButton#btnEnhance:disabled { background: #222230; color: #444460; border-color: #333; }

        QPushButton#btnSave {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                stop:0 #0e4420, stop:1 #0e5830);
            color: #ccffdd;
            border: 1px solid #1a8840;
            border-radius: 5px;
            font-weight: bold;
            font-size: 11pt;
            padding: 2px 10px;
        }
        QPushButton#btnSave:hover  { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #186630, stop:1 #188844); }
        QPushButton#btnSave:pressed { background: #0a3018; }
        QPushButton#btnSave:disabled { background: #222230; color: #444460; border-color: #333; }

        QPushButton {
            background: #1e1e30;
            color: #aabbdd;
            border: 1px solid #333355;
            border-radius: 4px;
            padding: 2px 6px;
        }
        QPushButton:hover { background: #282840; border-color: #4455aa; }

        QComboBox#modelCombo {
            background: #1a1a2e;
            color: #aaccff;
            border: 1px solid #334;
            border-radius: 4px;
            padding: 3px 8px;
        }
        QComboBox#modelCombo::drop-down { border: none; }
        QComboBox QAbstractItemView {
            background: #1a1a2e;
            color: #aaccff;
            selection-background-color: #2244aa;
        }

        QSlider::groove:horizontal {
            height: 5px;
            background: #252540;
            border-radius: 3px;
        }
        QSlider::handle:horizontal {
            background: #4488ff;
            border: 2px solid #66aaff;
            width: 14px; height: 14px;
            margin: -5px 0;
            border-radius: 7px;
        }
        QSlider::sub-page:horizontal {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                stop:0 #2244cc, stop:1 #44aacc);
            border-radius: 3px;
        }

        QProgressBar {
            border: 1px solid #252540;
            border-radius: 3px;
            background: #111122;
            color: transparent;
        }
        QProgressBar::chunk {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                stop:0 #1a44cc, stop:1 #22ccaa);
            border-radius: 2px;
        }

        QStatusBar {
            background: #0c0c18;
            color: #7080a0;
            border-top: 1px solid #1a1a2e;
            font-size: 9pt;
        }

        QRadioButton::indicator {
            width: 13px; height: 13px;
        }
        QRadioButton::indicator:unchecked {
            border: 2px solid #334466;
            border-radius: 7px;
            background: #141422;
        }
        QRadioButton::indicator:checked {
            border: 2px solid #3366ff;
            border-radius: 7px;
            background: qradialgradient(cx:0.5,cy:0.5,radius:0.5,
                stop:0 #88aaff, stop:0.5 #3366ff, stop:1 #1a44cc);
        }

        QLabel#hintLbl { color: #44445a; font-size: 8pt; }

        QSplitter::handle:horizontal {
            background: #252540;
            width: 4px;
        }
        QSplitter::handle:horizontal:hover { background: #3355aa; }
        """)


# ============================================================
# ユーティリティ
# ============================================================
def _pil_to_qpixmap(img: Image.Image) -> QPixmap:
    rgb  = img.convert("RGB")
    data = rgb.tobytes("raw", "RGB")
    qi   = QImage(data, rgb.width, rgb.height, rgb.width * 3, QImage.Format_RGB888)
    return QPixmap.fromImage(qi.copy())


# ============================================================
# エントリーポイント
# ============================================================
def main():
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
