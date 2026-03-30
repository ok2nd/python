#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
比率指定トリミング
Aspect Ratio Trimming Tool for Windows
"""

import tkinter as tk
from tkinter import messagebox, scrolledtext
import os
import sys
import json
from pathlib import Path
from PIL import Image, ImageTk

# ─── パス設定 ───────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

RATIOS_FILE = APP_DIR / "trim_ratios.txt"
CONFIG_FILE = APP_DIR / "trim_config.json"

DEFAULT_RATIOS = """\
# 形式: 名前:横比率:縦比率  (# はコメント行)
正方形:1:1
横 4:3:4:3
縦 3:4:3:4
横 16:9:16:9
縦 9:16:9:16
写真 L判:89:127
写真 2L判:127:178
ハガキ:100:148
A4 横:297:210
A4 縦:210:297
名刺 横:91:55
名刺 縦:55:91
"""

# ─── カラーパレット ──────────────────────────────────────────
C = {
    "bg_dark":    "#1a1a2e",
    "bg_panel":   "#16213e",
    "bg_item":    "#0f3460",
    "bg_canvas":  "#0d0d1a",
    "accent":     "#e94560",
    "accent2":    "#533483",
    "fg":         "#e0e0f0",
    "fg_dim":     "#8888aa",
    "fg_ok":      "#4ecdc4",
    "handle":     "#e94560",
    "white":      "#ffffff",
    "border":     "#2a2a4a",
}

FONT_MAIN = ("Yu Gothic UI", 10)
FONT_BOLD = ("Yu Gothic UI", 10, "bold")
FONT_TITLE = ("Yu Gothic UI", 13, "bold")
FONT_SMALL = ("Yu Gothic UI", 9)
FONT_MONO = ("Consolas", 10)


# ───────────────────────────────────────────────────────────────
class TrimApp:

    def __init__(self, root):
        self.root = root
        self.root.title("比率指定トリミング")
        self.root.geometry("1100x720")
        self.root.minsize(850, 600)
        self.root.configure(bg=C["bg_dark"])

        # 状態
        self.image: Image.Image | None = None
        self.image_path: Path | None = None
        self.icc_profile = None
        self.photo_image = None

        self.scale = 1.0
        self.offset_x = 0
        self.offset_y = 0

        # トリミング矩形（画像座標）
        self.trim_x1 = 0.0
        self.trim_y1 = 0.0
        self.trim_x2 = 0.0
        self.trim_y2 = 0.0
        self.has_trim = False

        # ドラッグ状態
        self.drag_mode = None
        self.create_start_ix = 0.0
        self.create_start_iy = 0.0
        self.drag_start_trim = (0.0, 0.0, 0.0, 0.0)

        # 比率リスト & 設定
        self.ratios: list[tuple[str, float, float]] = []
        self.selected_ratio_idx = 0

        self._load_config()
        self._load_ratios()
        self._build_ui()
        self._setup_dnd()
        self._apply_saved_selection()

    # ══════════════════════════════════════════════
    #  設定ファイル
    # ══════════════════════════════════════════════
    def _load_config(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                self.config = json.load(f)
        except Exception:
            self.config = {}

    def _save_config(self):
        self.config["selected_ratio"] = self.selected_ratio_idx
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False)
        except Exception:
            pass

    def _load_ratios(self):
        if not RATIOS_FILE.exists():
            with open(RATIOS_FILE, "w", encoding="utf-8") as f:
                f.write(DEFAULT_RATIOS)

        self.ratios = []
        try:
            with open(RATIOS_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    parts = line.split(":")
                    if len(parts) == 3:
                        try:
                            name = parts[0].strip()
                            w = float(parts[1].strip())
                            h = float(parts[2].strip())
                            if w > 0 and h > 0:
                                self.ratios.append((name, w, h))
                        except ValueError:
                            pass
        except Exception:
            pass

        if not self.ratios:
            self.ratios = [("正方形", 1, 1), ("横4:3", 4, 3), ("縦3:4", 3, 4)]

    # ══════════════════════════════════════════════
    #  UI 構築
    # ══════════════════════════════════════════════
    def _build_ui(self):
        self.root.columnconfigure(0, weight=0)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        self._build_left_panel()
        self._build_right_panel()

    def _build_left_panel(self):
        frame = tk.Frame(self.root, width=220, bg=C["bg_panel"],
                         relief="flat", bd=0)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_propagate(False)

        # ヘッダー
        header = tk.Frame(frame, bg=C["accent"], height=4)
        header.pack(fill="x")

        tk.Label(frame, text="トリミング比率", bg=C["bg_panel"], fg=C["fg"],
                 font=FONT_TITLE).pack(pady=(18, 8), padx=14, anchor="w")

        # リストボックス
        list_outer = tk.Frame(frame, bg=C["border"], bd=1)
        list_outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        sb = tk.Scrollbar(list_outer, bg=C["bg_dark"], troughcolor=C["bg_dark"])
        sb.pack(side="right", fill="y")

        self.ratio_listbox = tk.Listbox(
            list_outer,
            yscrollcommand=sb.set,
            bg=C["bg_item"],
            fg=C["fg"],
            selectbackground=C["accent"],
            selectforeground=C["white"],
            font=FONT_MAIN,
            activestyle="none",
            borderwidth=0,
            highlightthickness=0,
            relief="flat",
        )
        self.ratio_listbox.pack(side="left", fill="both", expand=True)
        sb.config(command=self.ratio_listbox.yview)
        self.ratio_listbox.bind("<<ListboxSelect>>", self._on_ratio_select)

        self._populate_listbox()

        # ボタン群
        def styled_btn(parent, text, cmd, color=C["bg_dark"]):
            b = tk.Button(
                parent, text=text, command=cmd,
                bg=color, fg=C["fg"], relief="flat",
                font=FONT_SMALL, pady=7, cursor="hand2",
                activebackground=C["accent2"], activeforeground=C["white"],
                borderwidth=0,
            )
            return b

        styled_btn(frame, "⚙  比率設定を編集", self._open_settings).pack(
            fill="x", padx=12, pady=(0, 4))
        styled_btn(frame, "↻  設定を再読み込み", self._reload_ratios).pack(
            fill="x", padx=12, pady=(0, 4))

        # ファイルを開く（DnD 未対応時の代替）
        self.open_btn = styled_btn(frame, "📂  ファイルを開く", self._open_file_dialog)
        self.open_btn.pack(fill="x", padx=12, pady=(0, 16))

        # バージョン表示
        tk.Label(frame, text="比率指定トリミング v1.0", bg=C["bg_panel"],
                 fg=C["fg_dim"], font=("Yu Gothic UI", 8)).pack(side="bottom", pady=8)

    def _populate_listbox(self):
        self.ratio_listbox.delete(0, tk.END)
        for name, w, h in self.ratios:
            wi = int(w) if w == int(w) else w
            hi = int(h) if h == int(h) else h
            self.ratio_listbox.insert(tk.END, f"  {name}  ({wi}:{hi})")

    def _build_right_panel(self):
        frame = tk.Frame(self.root, bg=C["bg_dark"])
        frame.grid(row=0, column=1, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        frame.rowconfigure(1, weight=0)

        # キャンバス
        self.canvas = tk.Canvas(
            frame, bg=C["bg_canvas"], cursor="crosshair",
            highlightthickness=0, relief="flat",
        )
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        # ドロップガイド（初期表示）
        self._draw_drop_guide()

        self.canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.canvas.bind("<Motion>", self._on_canvas_motion)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # ボトムバー
        bot = tk.Frame(frame, bg=C["bg_panel"], height=52)
        bot.grid(row=1, column=0, sticky="ew")
        bot.grid_propagate(False)

        # アクセントライン
        tk.Frame(bot, bg=C["accent"], height=2).pack(fill="x", side="top")

        inner = tk.Frame(bot, bg=C["bg_panel"])
        inner.pack(fill="both", expand=True, padx=10)

        self.info_label = tk.Label(
            inner, text="画像をドラッグ＆ドロップしてください",
            bg=C["bg_panel"], fg=C["fg_dim"], font=FONT_SMALL,
        )
        self.info_label.pack(side="left", pady=12)

        self.trim_info_label = tk.Label(
            inner, text="", bg=C["bg_panel"], fg=C["fg_ok"], font=FONT_SMALL,
        )
        self.trim_info_label.pack(side="left", padx=14, pady=12)

        # 保存ボタン
        self.save_btn = tk.Button(
            inner, text="  💾  保  存  ", command=self._save_image,
            bg=C["accent"], fg=C["white"], relief="flat",
            font=FONT_BOLD, pady=4, state="disabled", cursor="hand2",
            activebackground=C["accent2"], activeforeground=C["white"],
            borderwidth=0,
        )
        self.save_btn.pack(side="right", pady=10, padx=(6, 0))

        self.reset_btn = tk.Button(
            inner, text="リセット", command=self._reset_trim,
            bg=C["bg_dark"], fg=C["fg_dim"], relief="flat",
            font=FONT_SMALL, pady=4, state="disabled", cursor="hand2",
            activebackground=C["border"], activeforeground=C["fg"],
            borderwidth=0,
        )
        self.reset_btn.pack(side="right", pady=10)

    def _draw_drop_guide(self):
        self.canvas.delete("guide")
        cw = self.canvas.winfo_width() or 600
        ch = self.canvas.winfo_height() or 500
        cx, cy = cw // 2, ch // 2

        # 点線の矩形
        self.canvas.create_rectangle(
            cx - 200, cy - 140, cx + 200, cy + 140,
            outline=C["fg_dim"], width=2, dash=(8, 6), tags="guide",
        )
        self.canvas.create_text(
            cx, cy - 20,
            text="🖼",
            font=("Segoe UI Emoji", 36), fill=C["fg_dim"], tags="guide",
        )
        self.canvas.create_text(
            cx, cy + 48,
            text="画像をここにドロップ",
            font=("Yu Gothic UI", 13), fill=C["fg_dim"], tags="guide",
        )
        self.canvas.create_text(
            cx, cy + 76,
            text="または左下のボタンで開く",
            font=("Yu Gothic UI", 10), fill="#555577", tags="guide",
        )

    # ══════════════════════════════════════════════
    #  ドラッグ＆ドロップ
    # ══════════════════════════════════════════════
    def _setup_dnd(self):
        try:
            from tkinterdnd2 import DND_FILES
            self.canvas.drop_target_register(DND_FILES)
            self.canvas.dnd_bind("<<Drop>>", self._on_drop)
            self.dnd_available = True
        except Exception:
            self.dnd_available = False

    def _on_drop(self, event):
        import re
        raw = event.data.strip()
        if raw.startswith("{"):
            files = re.findall(r"\{([^}]+)\}|([^\s{}]+)", raw)
            files = [a or b for a, b in files]
        else:
            files = raw.split()
        if files:
            self._load_image(files[0])

    def _open_file_dialog(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="画像を選択",
            filetypes=[
                ("画像ファイル", "*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp"),
                ("すべて", "*.*"),
            ],
        )
        if path:
            self._load_image(path)

    # ══════════════════════════════════════════════
    #  画像読み込み
    # ══════════════════════════════════════════════
    def _load_image(self, path: str):
        try:
            img = Image.open(path)
            img.load()  # 完全読み込み（遅延読み込みを解消）
            self.image = img
            self.image_path = Path(path)
            self.icc_profile = img.info.get("icc_profile")
            self.has_trim = False
            self._update_display()
            self._auto_trim()
            self.info_label.config(
                text=f"📄 {self.image_path.name}   {img.width} × {img.height} px"
            )
            self.save_btn.config(state="normal")
            self.reset_btn.config(state="normal")
        except Exception as e:
            messagebox.showerror("エラー", f"画像を開けませんでした:\n{e}")

    # ══════════════════════════════════════════════
    #  表示更新
    # ══════════════════════════════════════════════
    def _update_display(self):
        if self.image is None:
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return

        self.scale = min(cw / self.image.width, ch / self.image.height) * 0.94
        dw = max(1, int(self.image.width * self.scale))
        dh = max(1, int(self.image.height * self.scale))
        self.offset_x = (cw - dw) // 2
        self.offset_y = (ch - dh) // 2

        disp = self.image.copy()
        if disp.mode == "P":
            disp = disp.convert("RGBA")
        if disp.mode in ("RGBA", "LA"):
            bg = Image.new("RGB", disp.size, (26, 26, 46))
            bg.paste(disp, mask=disp.split()[-1])
            disp = bg
        elif disp.mode != "RGB":
            disp = disp.convert("RGB")
        disp = disp.resize((dw, dh), Image.LANCZOS)

        self.photo_image = ImageTk.PhotoImage(disp)
        self.canvas.delete("all")
        self.canvas.create_image(self.offset_x, self.offset_y, anchor="nw",
                                  image=self.photo_image, tags="img")
        if self.has_trim:
            self._draw_trim_rect()

    def _draw_trim_rect(self):
        self.canvas.delete("trim_overlay")
        self.canvas.delete("trim_rect")
        if not self.has_trim or self.image is None:
            return

        x1c, y1c = self._img2cv(self.trim_x1, self.trim_y1)
        x2c, y2c = self._img2cv(self.trim_x2, self.trim_y2)

        ix1c = self.offset_x
        iy1c = self.offset_y
        ix2c = self.offset_x + self.image.width * self.scale
        iy2c = self.offset_y + self.image.height * self.scale

        # 暗幕オーバーレイ（4 分割）
        stipple = "gray50"
        ov_cfg = dict(fill="black", stipple=stipple, outline="", tags="trim_overlay")
        self.canvas.create_rectangle(ix1c, iy1c, ix2c, y1c, **ov_cfg)
        self.canvas.create_rectangle(ix1c, y2c, ix2c, iy2c, **ov_cfg)
        self.canvas.create_rectangle(ix1c, y1c, x1c, y2c, **ov_cfg)
        self.canvas.create_rectangle(x2c, y1c, ix2c, y2c, **ov_cfg)

        # 枠線
        self.canvas.create_rectangle(
            x1c, y1c, x2c, y2c,
            outline=C["white"], width=2, tags="trim_rect",
        )



    # ══════════════════════════════════════════════
    #  座標変換
    # ══════════════════════════════════════════════
    def _img2cv(self, ix, iy):
        return ix * self.scale + self.offset_x, iy * self.scale + self.offset_y

    def _cv2img(self, cx, cy):
        return (cx - self.offset_x) / self.scale, (cy - self.offset_y) / self.scale

    # ══════════════════════════════════════════════
    #  比率
    # ══════════════════════════════════════════════
    def _current_ratio(self) -> float:
        if 0 <= self.selected_ratio_idx < len(self.ratios):
            _, w, h = self.ratios[self.selected_ratio_idx]
            return w / h
        return 1.0

    # ══════════════════════════════════════════════
    #  トリミング自動設定
    # ══════════════════════════════════════════════
    def _auto_trim(self):
        if self.image is None:
            return
        ratio = self._current_ratio()
        iw, ih = self.image.width, self.image.height
        if iw / ih > ratio:
            th, tw = ih, ih * ratio
        else:
            tw, th = iw, iw / ratio
        cx, cy = iw / 2, ih / 2
        self.trim_x1 = cx - tw / 2
        self.trim_y1 = cy - th / 2
        self.trim_x2 = cx + tw / 2
        self.trim_y2 = cy + th / 2
        self._clamp_trim()
        self.has_trim = True
        self._draw_trim_rect()
        self._update_trim_info()

    def _clamp_trim(self):
        if self.image is None:
            return
        iw, ih = self.image.width, self.image.height
        self.trim_x1 = max(0.0, min(self.trim_x1, iw))
        self.trim_y1 = max(0.0, min(self.trim_y1, ih))
        self.trim_x2 = max(0.0, min(self.trim_x2, iw))
        self.trim_y2 = max(0.0, min(self.trim_y2, ih))

    def _adjust_trim_to_ratio(self):
        ratio = self._current_ratio()
        cx = (self.trim_x1 + self.trim_x2) / 2
        cy = (self.trim_y1 + self.trim_y2) / 2
        cur_w = abs(self.trim_x2 - self.trim_x1)
        cur_h = abs(self.trim_y2 - self.trim_y1)
        area = cur_w * cur_h
        new_h = (area / ratio) ** 0.5
        new_w = new_h * ratio
        if self.image:
            new_w = min(new_w, self.image.width)
            new_h = min(new_h, self.image.height)
            # 比率を再確認
            if new_w / new_h > ratio:
                new_w = new_h * ratio
            else:
                new_h = new_w / ratio
        self.trim_x1 = cx - new_w / 2
        self.trim_y1 = cy - new_h / 2
        self.trim_x2 = cx + new_w / 2
        self.trim_y2 = cy + new_h / 2
        self._clamp_trim()
        self._draw_trim_rect()
        self._update_trim_info()

    # ══════════════════════════════════════════════
    #  マウスイベント
    # ══════════════════════════════════════════════
    def _get_drag_mode(self, cx, cy) -> str:
        if not self.has_trim or self.image is None:
            return "create"
        x1c, y1c = self._img2cv(self.trim_x1, self.trim_y1)
        x2c, y2c = self._img2cv(self.trim_x2, self.trim_y2)
        hs = 14
        corners = {
            "resize_nw": (x1c, y1c),
            "resize_ne": (x2c, y1c),
            "resize_sw": (x1c, y2c),
            "resize_se": (x2c, y2c),
        }
        for mode, (hx, hy) in corners.items():
            if abs(cx - hx) <= hs and abs(cy - hy) <= hs:
                return mode
        if x1c <= cx <= x2c and y1c <= cy <= y2c:
            return "move"
        return "create"

    def _on_canvas_press(self, event):
        if self.image is None:
            return
        self.drag_mode = self._get_drag_mode(event.x, event.y)
        self.drag_start_trim = (self.trim_x1, self.trim_y1, self.trim_x2, self.trim_y2)
        ix, iy = self._cv2img(event.x, event.y)
        self.create_start_ix = ix
        self.create_start_iy = iy
        if self.drag_mode == "create":
            self.trim_x1 = self.trim_x2 = ix
            self.trim_y1 = self.trim_y2 = iy
            self.has_trim = True

    def _on_canvas_drag(self, event):
        if self.image is None or self.drag_mode is None:
            return
        ix, iy = self._cv2img(event.x, event.y)
        iw, ih = self.image.width, self.image.height
        ratio = self._current_ratio()
        sx1, sy1, sx2, sy2 = self.drag_start_trim

        if self.drag_mode == "move":
            tw = sx2 - sx1
            th = sy2 - sy1
            dx = ix - self.create_start_ix
            dy = iy - self.create_start_iy
            nx1 = max(0.0, min(sx1 + dx, iw - tw))
            ny1 = max(0.0, min(sy1 + dy, ih - th))
            self.trim_x1 = nx1
            self.trim_y1 = ny1
            self.trim_x2 = nx1 + tw
            self.trim_y2 = ny1 + th

        elif self.drag_mode == "create":
            raw_w = ix - self.create_start_ix
            raw_h = iy - self.create_start_iy
            if raw_w == 0 and raw_h == 0:
                return
            sign_h = 1 if raw_h >= 0 else -1
            sign_w = 1 if raw_w >= 0 else -1
            if abs(raw_w) >= abs(raw_h) * ratio:
                new_w = raw_w
                new_h = abs(new_w) / ratio * sign_h
            else:
                new_h = raw_h
                new_w = abs(new_h) * ratio * sign_w
            x1 = min(self.create_start_ix, self.create_start_ix + new_w)
            y1 = min(self.create_start_iy, self.create_start_iy + new_h)
            x2 = max(self.create_start_ix, self.create_start_ix + new_w)
            y2 = max(self.create_start_iy, self.create_start_iy + new_h)
            self.trim_x1 = max(0.0, min(x1, iw))
            self.trim_y1 = max(0.0, min(y1, ih))
            self.trim_x2 = max(0.0, min(x2, iw))
            self.trim_y2 = max(0.0, min(y2, ih))

        else:  # resize_**
            mode = self.drag_mode
            if mode == "resize_se":
                ax, ay = sx1, sy1
            elif mode == "resize_nw":
                ax, ay = sx2, sy2
            elif mode == "resize_ne":
                ax, ay = sx1, sy2
            else:  # sw
                ax, ay = sx2, sy1

            raw_w = ix - ax
            raw_h = iy - ay
            if raw_w == 0 and raw_h == 0:
                return
            sign_h = 1 if raw_h >= 0 else -1
            sign_w = 1 if raw_w >= 0 else -1
            if abs(raw_w) >= abs(raw_h) * ratio:
                new_w = raw_w
                new_h = abs(new_w) / ratio * sign_h
            else:
                new_h = raw_h
                new_w = abs(new_h) * ratio * sign_w
            x1 = min(ax, ax + new_w)
            y1 = min(ay, ay + new_h)
            x2 = max(ax, ax + new_w)
            y2 = max(ay, ay + new_h)
            self.trim_x1 = max(0.0, min(x1, iw))
            self.trim_y1 = max(0.0, min(y1, ih))
            self.trim_x2 = max(0.0, min(x2, iw))
            self.trim_y2 = max(0.0, min(y2, ih))

        self._draw_trim_rect()
        self._update_trim_info()

    def _on_canvas_release(self, event):
        if self.trim_x2 < self.trim_x1:
            self.trim_x1, self.trim_x2 = self.trim_x2, self.trim_x1
        if self.trim_y2 < self.trim_y1:
            self.trim_y1, self.trim_y2 = self.trim_y2, self.trim_y1
        self.drag_mode = None
        self._draw_trim_rect()
        self._update_trim_info()

    def _on_canvas_motion(self, event):
        if self.image is None:
            return
        mode = self._get_drag_mode(event.x, event.y)
        cur_map = {
            "create": "crosshair",
            "move": "fleur",
            "resize_nw": "size_nw_se",
            "resize_se": "size_nw_se",
            "resize_ne": "size_ne_sw",
            "resize_sw": "size_ne_sw",
        }
        self.canvas.config(cursor=cur_map.get(mode, "crosshair"))

    def _on_canvas_resize(self, event):
        if self.image is None:
            self._draw_drop_guide()
        else:
            self._update_display()

    # ══════════════════════════════════════════════
    #  比率選択
    # ══════════════════════════════════════════════
    def _apply_saved_selection(self):
        idx = self.config.get("selected_ratio", 0)
        idx = max(0, min(idx, len(self.ratios) - 1))
        self.selected_ratio_idx = idx
        self.ratio_listbox.selection_set(idx)
        self.ratio_listbox.see(idx)

    def _on_ratio_select(self, event):
        sel = self.ratio_listbox.curselection()
        if not sel:
            return
        self.selected_ratio_idx = sel[0]
        self._save_config()
        if self.image and self.has_trim:
            self._adjust_trim_to_ratio()
        elif self.image:
            self._auto_trim()

    # ══════════════════════════════════════════════
    #  情報ラベル
    # ══════════════════════════════════════════════
    def _update_trim_info(self):
        if self.has_trim and self.image:
            w = int(abs(self.trim_x2 - self.trim_x1))
            h = int(abs(self.trim_y2 - self.trim_y1))
            self.trim_info_label.config(text=f"✂  {w} × {h} px")
        else:
            self.trim_info_label.config(text="")

    # ══════════════════════════════════════════════
    #  リセット
    # ══════════════════════════════════════════════
    def _reset_trim(self):
        if self.image:
            self._auto_trim()

    # ══════════════════════════════════════════════
    #  保存
    # ══════════════════════════════════════════════
    def _save_image(self):
        if self.image is None or not self.has_trim:
            messagebox.showinfo("情報", "画像とトリミング範囲を設定してください")
            return
        x1 = int(round(min(self.trim_x1, self.trim_x2)))
        y1 = int(round(min(self.trim_y1, self.trim_y2)))
        x2 = int(round(max(self.trim_x1, self.trim_x2)))
        y2 = int(round(max(self.trim_y1, self.trim_y2)))
        if x2 - x1 < 1 or y2 - y1 < 1:
            messagebox.showwarning("警告", "トリミング範囲が小さすぎます")
            return

        cropped = self.image.crop((x1, y1, x2, y2))
        suffix = self.image_path.suffix
        # 比率名を取得してファイル名に使用（ファイル名に使えない文字を除去）
        ratio_name = ""
        if 0 <= self.selected_ratio_idx < len(self.ratios):
            ratio_name = self.ratios[self.selected_ratio_idx][0]
        for ch in r'\/:*?"<>|':
            ratio_name = ratio_name.replace(ch, "_")
        ratio_name = ratio_name.strip()
        save_stem = f"{self.image_path.stem}-{ratio_name}" if ratio_name else f"{self.image_path.stem}-trim"
        save_path = self.image_path.parent / f"{save_stem}{suffix}"

        save_kwargs = {}
        if self.icc_profile:
            save_kwargs["icc_profile"] = self.icc_profile

        fmt = suffix.lower()
        if fmt in (".jpg", ".jpeg"):
            save_kwargs.setdefault("quality", 95)
            save_kwargs.setdefault("subsampling", 0)

        try:
            cropped.save(save_path, **save_kwargs)
            messagebox.showinfo(
                "保存完了",
                f"保存しました:\n{save_path.name}\n\n({x2-x1} × {y2-y1} px)",
            )
        except Exception as e:
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{e}")

    # ══════════════════════════════════════════════
    #  設定エディタ
    # ══════════════════════════════════════════════
    def _open_settings(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("比率設定の編集")
        dlg.geometry("520x540")
        dlg.configure(bg=C["bg_panel"])
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(
            dlg,
            text="形式:  名前:横比率:縦比率   （# で始まる行はコメント）",
            bg=C["bg_panel"], fg=C["fg_dim"], font=FONT_SMALL,
        ).pack(pady=(14, 4), padx=14, anchor="w")

        editor = scrolledtext.ScrolledText(
            dlg, font=FONT_MONO, wrap=tk.NONE,
            bg=C["bg_item"], fg=C["fg"], insertbackground=C["accent"],
            selectbackground=C["accent"], selectforeground=C["white"],
            borderwidth=0, highlightthickness=1,
            highlightcolor=C["accent"], highlightbackground=C["border"],
        )
        editor.pack(fill="both", expand=True, padx=14, pady=4)

        try:
            with open(RATIOS_FILE, "r", encoding="utf-8") as f:
                editor.insert("1.0", f.read())
        except Exception:
            editor.insert("1.0", DEFAULT_RATIOS)

        def save_close():
            content = editor.get("1.0", tk.END).rstrip("\n")
            try:
                with open(RATIOS_FILE, "w", encoding="utf-8") as f:
                    f.write(content + "\n")
                self._reload_ratios()
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("エラー", f"保存できませんでした:\n{e}", parent=dlg)

        btn_row = tk.Frame(dlg, bg=C["bg_panel"])
        btn_row.pack(fill="x", padx=14, pady=12)
        tk.Button(
            btn_row, text="キャンセル", command=dlg.destroy,
            bg=C["bg_dark"], fg=C["fg"], relief="flat", font=FONT_SMALL, padx=12, pady=5,
        ).pack(side="right", padx=(6, 0))
        tk.Button(
            btn_row, text="保存して閉じる", command=save_close,
            bg=C["accent"], fg=C["white"], relief="flat", font=FONT_BOLD, padx=12, pady=5,
        ).pack(side="right")

    def _reload_ratios(self):
        old_idx = self.selected_ratio_idx
        self._load_ratios()
        self._populate_listbox()
        new_idx = min(old_idx, len(self.ratios) - 1)
        self.selected_ratio_idx = new_idx
        self.ratio_listbox.selection_set(new_idx)
        self.ratio_listbox.see(new_idx)
        if self.image and self.has_trim:
            self._adjust_trim_to_ratio()

    # ══════════════════════════════════════════════
    #  終了
    # ══════════════════════════════════════════════
    def on_close(self):
        self._save_config()
        self.root.destroy()


# ───────────────────────────────────────────────────────────────
def main():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()

    app = TrimApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
