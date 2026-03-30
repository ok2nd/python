#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
画像一括縮小アプリ
"""

import os
import io
import sys
import threading
from pathlib import Path

from PIL import Image, ImageCms

try:
    import piexif
    HAS_PIEXIF = True
except ImportError:
    HAS_PIEXIF = False

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QRadioButton, QButtonGroup, QLineEdit,
    QCheckBox, QSpinBox, QGroupBox, QTextEdit,
)
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QFont

# ================================================================
# ★ カスタマイズエリア（ここを自由に変更してください）
# ================================================================
# ラジオボタンに表示するサイズ選択肢（長辺px）
SIZE_OPTIONS = [1600, 1280, 1200, 1024, 800, 640, 400, 320]

# デフォルトで選択されるサイズ（SIZE_OPTIONS のいずれかの値を指定）
DEFAULT_SIZE = 1200

# JPEG画質のデフォルト値（80〜100）
DEFAULT_QUALITY = 90
# ================================================================

# 対応する入力拡張子
SUPPORTED_EXT = {'.jpg', '.jpeg', '.png', '.webp', '.tif', '.tiff', '.bmp'}


# ----------------------------------------------------------------
# 画像処理ユーティリティ
# ----------------------------------------------------------------

def calc_new_size(w, h, long_side, no_upscale):
    """長辺を long_side px にしたときの (幅, 高さ) を返す"""
    long = max(w, h)
    if no_upscale and long <= long_side:
        return w, h
    ratio = long_side / long
    return max(1, round(w * ratio)), max(1, round(h * ratio))


def exif_remove_gps(raw):
    """EXIFバイト列からGPS情報を除去して返す"""
    if not HAS_PIEXIF or not raw:
        return raw
    try:
        d = piexif.load(raw)
        d['GPS'] = {}
        return piexif.dump(d)
    except Exception:
        return raw


def get_srgb_icc():
    """sRGB ICCプロファイルのバイト列を返す"""
    try:
        profile = ImageCms.createProfile('sRGB')
        return ImageCms.ImageCmsProfile(profile).tobytes()
    except Exception:
        return b''


def convert_to_srgb(img):
    """
    ICCプロファイルを使って画像を sRGB に変換する。
    変換後の画像と、新しいICCバイト列のタプルを返す。
    """
    icc = img.info.get('icc_profile', b'')
    if not icc:
        return img, b''
    try:
        src_profile = ImageCms.ImageCmsProfile(io.BytesIO(icc))
        dst_profile = ImageCms.createProfile('sRGB')
        # モードを RGB に揃える
        if img.mode not in ('RGB', 'RGBA'):
            img = img.convert('RGB')
        img = ImageCms.profileToProfile(
            img, src_profile, dst_profile,
            renderingIntent=0,   # PERCEPTUAL
            outputMode='RGB'
        )
        return img, get_srgb_icc()
    except Exception:
        return img, icc


# ----------------------------------------------------------------
# ワーカー（別スレッドで実行）
# ----------------------------------------------------------------

class WorkerSignals(QObject):
    log  = Signal(str)
    done = Signal(str)


def process_images(files, long_side, fmt, quality, exif_mode,
                   no_upscale, adobe_to_srgb, signals):
    """
    画像リストを処理するメイン関数。
    signals.log / signals.done でメインスレッドに通知する。
    """
    ext_map  = {'JPEG': '.jpg', 'PNG': '.png', 'WebP': '.webp'}
    save_ext = ext_map.get(fmt, '.jpg')
    ok = ng = 0

    for fpath in files:
        src = Path(fpath)
        if src.suffix.lower() not in SUPPORTED_EXT:
            continue
        try:
            signals.log.emit(f"▶ {src.name}")

            img = Image.open(src)
            icc  = img.info.get('icc_profile', b'')
            exif = img.info.get('exif', b'')

            # EXIF 正規化（piexif 経由で一度ロード＆ダンプ）
            if HAS_PIEXIF and exif:
                try:
                    exif = piexif.dump(piexif.load(exif))
                except Exception:
                    pass  # 壊れていても継続

            # AdobeRGB → sRGB 変換
            if adobe_to_srgb:
                img, icc = convert_to_srgb(img)

            # 保存形式に合わせてモード変換
            if fmt == 'JPEG':
                if img.mode != 'RGB':
                    img = img.convert('RGB')
            elif img.mode not in ('RGB', 'RGBA', 'L', 'LA', 'P'):
                img = img.convert('RGB')

            # リサイズ
            w, h = img.size
            nw, nh = calc_new_size(w, h, long_side, no_upscale)
            if (nw, nh) != (w, h):
                img = img.resize((nw, nh), Image.LANCZOS)

            # 出力ディレクトリ（元フォルダ / サイズ名）
            out_dir = src.parent / str(long_side)
            out_dir.mkdir(exist_ok=True)
            out_path = out_dir / (src.stem + save_ext)

            # EXIF の取り扱い決定
            save_exif = b''
            if exif_mode == 'inherit':
                save_exif = exif
            elif exif_mode == 'gps_remove':
                save_exif = exif_remove_gps(exif)
            # 'delete' の場合は save_exif = b'' のまま

            # 保存オプション組み立て
            kw = {}
            if icc:                            # ICCプロファイルは必ず維持
                kw['icc_profile'] = icc
            if fmt == 'JPEG':
                kw['quality']     = quality
                kw['subsampling'] = 0          # 高品質クロマサブサンプリング
                if save_exif:
                    kw['exif'] = save_exif
            elif fmt == 'WebP':
                kw['quality'] = quality
                if save_exif:
                    kw['exif'] = save_exif
            elif fmt == 'PNG':
                if save_exif:
                    try:
                        kw['exif'] = save_exif
                    except Exception:
                        pass

            img.save(out_path, fmt, **kw)

            # タイムスタンプを元画像と同じにする
            st = src.stat()
            os.utime(out_path, (st.st_atime, st.st_mtime))

            signals.log.emit(f"  ✓ {out_dir.name}/{out_path.name}  ({nw}×{nh})")
            ok += 1

        except Exception as e:
            signals.log.emit(f"  ✗ エラー: {src.name} → {e}")
            ng += 1

    signals.done.emit(f"✅ 完了： 成功 {ok} 件 / 失敗 {ng} 件")


# ----------------------------------------------------------------
# ドロップエリア
# ----------------------------------------------------------------

class DropArea(QLabel):
    files_dropped = Signal(list)

    _CSS_IDLE  = (
        "QLabel { border: 2px dashed #aaa; border-radius: 8px;"
        " background: #fafafa; color: #555; font-size: 13px; padding: 16px; }"
    )
    _CSS_HOVER = (
        "QLabel { border: 2px dashed #4488ee; border-radius: 8px;"
        " background: #eaf1ff; color: #3366cc; font-size: 13px; padding: 16px; }"
    )
    _CSS_BUSY  = (
        "QLabel { border: 2px dashed #ccc; border-radius: 8px;"
        " background: #f0f0f0; color: #aaa; font-size: 13px; padding: 16px; }"
    )

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumHeight(90)
        self._set_idle()

    def _set_idle(self):
        self.setText("ここに画像をドラッグ＆ドロップ\nJPG / PNG / WebP / TIFF / BMP")
        self.setStyleSheet(self._CSS_IDLE)
        self.setAcceptDrops(True)

    def set_busy(self):
        self.setText("処理中…")
        self.setStyleSheet(self._CSS_BUSY)
        self.setAcceptDrops(False)

    def set_ready(self):
        self._set_idle()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self.setStyleSheet(self._CSS_HOVER)
        else:
            e.ignore()

    def dragLeaveEvent(self, e):
        self.setStyleSheet(self._CSS_IDLE)

    def dropEvent(self, e):
        self.setStyleSheet(self._CSS_IDLE)
        paths = []
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isdir(p):
                for f in Path(p).rglob('*'):
                    if f.suffix.lower() in SUPPORTED_EXT:
                        paths.append(str(f))
            elif Path(p).suffix.lower() in SUPPORTED_EXT:
                paths.append(p)
        if paths:
            self.files_dropped.emit(paths)


# ----------------------------------------------------------------
# メインウィンドウ
# ----------------------------------------------------------------

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("画像一括縮小")
        self._worker_signals = None
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(7)
        root.setContentsMargins(10, 10, 10, 10)

        # ---- ドロップエリア ----
        self.drop_area = DropArea()
        self.drop_area.files_dropped.connect(self._on_drop)
        root.addWidget(self.drop_area)

        # ---- 画像サイズ（長辺） ----
        size_box = QGroupBox("画像サイズ（長辺）")
        sv = QVBoxLayout(size_box)
        sv.setSpacing(3)
        sv.setContentsMargins(8, 6, 8, 8)

        self.size_group = QButtonGroup(self)
        self._size_btns = {}  # {size_int: QRadioButton}

        # SIZE_OPTIONS を 4列で並べる
        row_layout = None
        for i, sz in enumerate(SIZE_OPTIONS):
            if i % 4 == 0:
                row_layout = QHBoxLayout()
                row_layout.setSpacing(4)
                sv.addLayout(row_layout)
            rb = QRadioButton(str(sz))
            self.size_group.addButton(rb)
            self._size_btns[sz] = rb
            row_layout.addWidget(rb)

        # 任意サイズ
        custom_row = QHBoxLayout()
        custom_row.setSpacing(4)
        self._custom_rb   = QRadioButton("任意")
        self._custom_edit = QLineEdit()
        self._custom_edit.setPlaceholderText("数値")
        self._custom_edit.setFixedWidth(58)
        self._custom_edit.setEnabled(False)
        self.size_group.addButton(self._custom_rb)
        self._custom_rb.toggled.connect(
            lambda checked: self._custom_edit.setEnabled(checked)
        )
        custom_row.addWidget(self._custom_rb)
        custom_row.addWidget(self._custom_edit)
        custom_row.addWidget(QLabel("px"))
        custom_row.addStretch()
        sv.addLayout(custom_row)
        root.addWidget(size_box)

        # デフォルトサイズを選択
        if DEFAULT_SIZE in self._size_btns:
            self._size_btns[DEFAULT_SIZE].setChecked(True)
        elif self._size_btns:
            next(iter(self._size_btns.values())).setChecked(True)

        # ---- オプション ----
        opt_box = QGroupBox("オプション")
        ov = QVBoxLayout(opt_box)
        ov.setSpacing(4)
        ov.setContentsMargins(8, 6, 8, 8)

        # EXIF
        exif_row = QHBoxLayout()
        exif_row.addWidget(QLabel("EXIF："))
        self.exif_group = QButtonGroup(self)
        self._exif_btns = {}  # {val_str: QRadioButton}
        for label, val in [("削除", "delete"), ("継承", "inherit"), ("GPS削除", "gps_remove")]:
            rb = QRadioButton(label)
            self.exif_group.addButton(rb)
            self._exif_btns[val] = rb
            exif_row.addWidget(rb)
        self._exif_btns["delete"].setChecked(True)
        exif_row.addStretch()
        ov.addLayout(exif_row)

        # 保存形式
        fmt_row = QHBoxLayout()
        fmt_row.addWidget(QLabel("保存形式："))
        self.fmt_group = QButtonGroup(self)
        self._fmt_btns = {}  # {fmt_str: QRadioButton}
        for fmt in ("JPEG", "PNG", "WebP"):
            rb = QRadioButton(fmt)
            self.fmt_group.addButton(rb)
            self._fmt_btns[fmt] = rb
            fmt_row.addWidget(rb)
        self._fmt_btns["JPEG"].setChecked(True)
        fmt_row.addStretch()
        ov.addLayout(fmt_row)

        # JPEG画質
        qual_row = QHBoxLayout()
        qual_row.addWidget(QLabel("JPEG 画質："))
        self.quality_spin = QSpinBox()
        self.quality_spin.setRange(80, 100)
        self.quality_spin.setValue(DEFAULT_QUALITY)
        qual_row.addWidget(self.quality_spin)
        qual_row.addStretch()
        ov.addLayout(qual_row)

        # チェックボックス
        self.no_upscale_cb = QCheckBox("画像を拡大しない")
        self.no_upscale_cb.setChecked(True)
        ov.addWidget(self.no_upscale_cb)

        self.adobe_srgb_cb = QCheckBox("AdobeRGB → sRGB 変換")
        ov.addWidget(self.adobe_srgb_cb)

        root.addWidget(opt_box)

        # ---- ログ表示 ----
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFixedHeight(76)
        self.log_view.setStyleSheet(
            "font-size: 11px; background: #f5f5f5; border: 1px solid #ddd;"
        )
        self.log_view.setPlaceholderText("処理ログ")
        root.addWidget(self.log_view)

        self.setFixedWidth(375)
        self.adjustSize()

    # ---- ヘルパー ----

    def _get_long_side(self):
        """選択中の長辺サイズを int で返す。不正なら None。"""
        if self._custom_rb.isChecked():
            try:
                v = int(self._custom_edit.text())
                return v if v > 0 else None
            except ValueError:
                return None
        for sz, rb in self._size_btns.items():
            if rb.isChecked():
                return sz
        return None

    def _get_exif_mode(self):
        for val, rb in self._exif_btns.items():
            if rb.isChecked():
                return val
        return "delete"

    def _get_fmt(self):
        for fmt, rb in self._fmt_btns.items():
            if rb.isChecked():
                return fmt
        return "JPEG"

    # ---- イベント ----

    def _on_drop(self, files):
        long_side = self._get_long_side()
        if not long_side:
            self.log_view.append("⚠ 有効なサイズを指定してください。")
            return

        fmt      = self._get_fmt()
        exif_mode = self._get_exif_mode()

        self.drop_area.set_busy()
        self.log_view.clear()
        self.log_view.append(
            f"サイズ: {long_side}px  形式: {fmt}  "
            f"EXIF: {exif_mode}  {len(files)} ファイル"
        )

        # シグナルオブジェクトを保持（GC対策）
        self._worker_signals = WorkerSignals()
        self._worker_signals.log.connect(self.log_view.append)
        self._worker_signals.done.connect(self._on_done)

        threading.Thread(
            target=process_images,
            args=(
                files,
                long_side,
                fmt,
                self.quality_spin.value(),
                exif_mode,
                self.no_upscale_cb.isChecked(),
                self.adobe_srgb_cb.isChecked(),
                self._worker_signals,
            ),
            daemon=True,
        ).start()

    def _on_done(self, msg):
        self.log_view.append(msg)
        self.drop_area.set_ready()


# ----------------------------------------------------------------
# エントリーポイント
# ----------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
