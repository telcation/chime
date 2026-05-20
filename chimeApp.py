"""
学校チャイムアプリ (Canvas レイヤ版)

レイヤ構成（後ろ→前）:
  Layer 0 : 背景色 (canvas 背景)
  Layer 1 : 画像   (canvas image item)
  Layer 2 : オーバーレイテキスト (スケジュールで指定)
  Layer 3 : 時計・次チャイム・ボタン類 (canvas window item)

スケジュールファイル形式 (タブ or スペース区切り):
  HH:MM  画像ファイル  表示文字  色  大きさ  フォント
  ※ "-" で「なし」を指定。画像ファイル以降は省略可。

日程Excelファイル形式:
  「日程」シートの「月日」列で今日の行を検索し、
  「テキスト」列と「ツール」列の値を準備中画面に表示する。

設定ファイル (config.txt):
  schedule_file=...      スケジュールファイルのパス
  excel_file=...         日程Excelファイルのパス
  fullscreen=True/False
  frontmost=True/False
  bgcolor=#000000
  img_keep_ratio=True/False
  img_width=              空欄=自動
  img_height=             空欄=自動
  standby_font=メイリオ   準備中画面のフォント
  standby_text_size=32    テキスト欄のフォントサイズ
  standby_tool_size=28    ツール欄のフォントサイズ
  standby_text_color=white
  standby_tool_color=white
  standby_text_bg=#2c3e50
  standby_tool_bg=#c0392b
  standby_text_y=80       上部表示のY座標（上端からpx）
  standby_tool_y=80       下部表示のY座標（下端からpx）
"""

import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from datetime import datetime
import threading
import time
import os
import sys
from typing import Optional, List, Dict

import numpy as np
import sounddevice as sd
from scipy.io import wavfile
from PIL import Image, ImageTk
import pandas as pd


# -----------------------------------------------------------------------
# 定数
# -----------------------------------------------------------------------
CONFIG_FILE    = "config.txt"
DEFAULT_IMAGE  = "準備中.png"
CHIME_WAV_PATH = "chime.wav"

# config.txt キー
KEY_SCHEDULE   = "schedule_file"
KEY_EXCEL      = "excel_file"
KEY_FULLSCREEN = "fullscreen"
KEY_FRONTMOST  = "frontmost"
KEY_BGCOLOR    = "bgcolor"
KEY_RATIO      = "img_keep_ratio"
KEY_IMG_W      = "img_width"
KEY_IMG_H      = "img_height"
KEY_SB_FONT    = "standby_font"
KEY_SB_TSIZE   = "standby_text_size"
KEY_SB_LSIZE   = "standby_tool_size"
KEY_SB_TCOLOR  = "standby_text_color"
KEY_SB_LCOLOR  = "standby_tool_color"
KEY_SB_TBG     = "standby_text_bg"
KEY_SB_LBG     = "standby_tool_bg"
KEY_SB_TY      = "standby_text_y"
KEY_SB_LY      = "standby_tool_y"

DEFAULT_BGCOLOR   = "#000000"
DEFAULT_SB_FONT   = "メイリオ"
DEFAULT_SB_TSIZE  = 60
DEFAULT_SB_LSIZE  = 60
DEFAULT_SB_TCOLOR = "white"
DEFAULT_SB_LCOLOR = "white"
DEFAULT_SB_TBG    = "#2c3e50"
DEFAULT_SB_LBG    = "#c0392b"
DEFAULT_SB_TY     = 300
DEFAULT_SB_LY     = 300

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# -----------------------------------------------------------------------
# ユーティリティ
# -----------------------------------------------------------------------
def contrast_color(bg_hex: str) -> str:
    """背景色の輝度に応じて黒か白を返す。"""
    bg_hex = bg_hex.lstrip("#")
    if len(bg_hex) != 6:
        return "white"
    r, g, b = int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16)
    return "black" if (0.299 * r + 0.587 * g + 0.114 * b) > 128 else "white"


def to_wareki(dt: datetime) -> str:
    if dt.year >= 2019:
        return f"令和{dt.year - 2018}"
    return "平成"


def load_today_from_excel(excel_path: str) -> Optional[Dict[str, str]]:
    """
    日程Excelの「日程」シートから今日の行を検索し、
    {"text": str, "tool": str} を返す。
    該当行がなければ None を返す。
    列名「月日」「テキスト」「ツール」で検索する（列番号に依存しない）。
    """
    if not os.path.exists(excel_path):
        messagebox.showwarning(
            "Excelファイルが見つかりません",
            f"以下のパスにファイルが見つかりません。\n{excel_path}")
        return None
    try:
        # header=None で全読み込みし、「月日」を含む行をヘッダとして特定する
        raw = pd.read_excel(excel_path, sheet_name="日程", header=None)

        # ヘッダ行を探す（「月日」というセルがある行）
        header_row = None
        for i, row in raw.iterrows():
            if any(str(v).strip() == "月日" for v in row):
                header_row = i
                break

        if header_row is None:
            messagebox.showwarning(
                "Excel読み込みエラー",
                "「日程」シートに「月日」列が見つかりません。\n"
                "ヘッダ行に「月日」「テキスト」「ツール」が含まれているか確認してください。")
            return None

        # ヘッダ行以降を改めて読み込む
        df = pd.read_excel(excel_path, sheet_name="日程", header=header_row)
        df.columns = [str(c).strip() for c in df.columns]

        missing = [c for c in ("月日", "テキスト", "ツール") if c not in df.columns]
        if missing:
            messagebox.showwarning(
                "Excel列名エラー",
                f"以下の列が見つかりません: {', '.join(missing)}\n"
                f"現在の列名: {list(df.columns)}")
            return None

        today = datetime.now().date()
        for _, row in df.iterrows():
            cell = row["月日"]
            try:
                cell_date = pd.to_datetime(cell).date()
            except Exception:
                continue
            if cell_date == today:
                text_str = str(row["テキスト"]).strip()
                tool_str = str(row["ツール"]).strip()
                text_str = "" if text_str in ("nan", "None", "") else text_str
                tool_str = "" if tool_str in ("nan", "None", "") else tool_str
                if text_str or tool_str:
                    return {"text": text_str, "tool": tool_str}
        return None  # 今日の行なし（授業なし）

    except Exception as e:
        messagebox.showerror(
            "Excel読み込みエラー",
            f"読み込み中に予期しないエラーが発生しました。\n\n"
            f"ファイル: {excel_path}\n"
            f"エラー: {type(e).__name__}: {e}")
        return None


def parse_schedule_line(line: str) -> Optional[Dict]:
    """
    スケジュール行をパースして dict を返す。
    フォーマット: HH:MM  画像  表示文字  色  大きさ  フォント
    省略フィールドは None、"-" も None として扱う。
    """
    line = line.strip()
    if not line or line.startswith("#"):
        return None
    parts = line.split()

    def v(idx):
        val = parts[idx] if idx < len(parts) else None
        return None if (val is None or val == "-") else val

    try:
        return {
            "time" : parts[0],
            "image": v(1),
            "text" : v(2),
            "color": v(3),
            "size" : int(v(4)) if v(4) else None,
            "font" : v(5).replace("_", " ") if v(5) else None,
        }
    except Exception as e:
        print(f"スケジュール行パースエラー: {line!r} → {e}")
        return None


# -----------------------------------------------------------------------
# 設定ダイアログ
# -----------------------------------------------------------------------
class SettingsDialog(tk.Toplevel):
    """設定ダイアログ。"""

    def __init__(self, parent: tk.Tk, app: "ChimeApp"):
        super().__init__(parent)
        self.app = app
        self.title("設定")
        self.resizable(False, False)
        self.grab_set()

        pad = dict(padx=16, pady=5)
        S = ("Helvetica", 13)
        Ss = ("Helvetica", 11)

        # ── ウィンドウ ──────────────────────────
        tk.Label(self, text="■ ウィンドウ", font=("Helvetica", 13, "bold")).pack(anchor="w", padx=16, pady=(10, 2))

        self.frontmost_var = tk.BooleanVar(value=app.frontmost_var.get())
        tk.Checkbutton(self, text="最前面に表示", variable=self.frontmost_var, font=S).pack(anchor="w", **pad)

        self.fullscreen_var = tk.BooleanVar(value=app.fullscreen_var.get())
        tk.Checkbutton(self, text="画面いっぱいに表示", variable=self.fullscreen_var, font=S).pack(anchor="w", **pad)

        # 背景色
        cf = tk.Frame(self); cf.pack(anchor="w", **pad)
        tk.Label(cf, text="背景色:", font=S).pack(side=tk.LEFT)
        self.selected_color = app.bgcolor
        self.color_preview = tk.Label(cf, width=4, bg=app.bgcolor, relief="solid")
        self.color_preview.pack(side=tk.LEFT, padx=6)
        tk.Button(cf, text="選択…", command=self._pick_color).pack(side=tk.LEFT)

        # ── 画像 ──────────────────────────────
        tk.Label(self, text="■ 画像", font=("Helvetica", 13, "bold")).pack(anchor="w", padx=16, pady=(10, 2))

        sf = tk.Frame(self); sf.pack(anchor="w", padx=16, pady=2)
        tk.Label(sf, text="横", font=S).pack(side=tk.LEFT)
        self.img_w_var = tk.StringVar(value=str(app.img_width) if app.img_width else "")
        tk.Entry(sf, textvariable=self.img_w_var, width=6, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(sf, text="px　縦", font=S).pack(side=tk.LEFT)
        self.img_h_var = tk.StringVar(value=str(app.img_height) if app.img_height else "")
        tk.Entry(sf, textvariable=self.img_h_var, width=6, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(sf, text="px　(空欄=自動)", font=Ss, fg="gray").pack(side=tk.LEFT)

        self.ratio_var = tk.BooleanVar(value=app.img_keep_ratio)
        tk.Checkbutton(self, text="縦横比を固定", variable=self.ratio_var, font=S).pack(anchor="w", **pad)

        # ── 準備中画面 ────────────────────────
        tk.Label(self, text="■ 準備中画面（テキスト欄・ツール欄）",
                 font=("Helvetica", 13, "bold")).pack(anchor="w", padx=16, pady=(10, 2))

        sb = app.standby
        # フォント
        ff = tk.Frame(self); ff.pack(anchor="w", padx=16, pady=2)
        tk.Label(ff, text="フォント:", font=S).pack(side=tk.LEFT)
        self.sb_font_var = tk.StringVar(value=sb["font"])
        tk.Entry(ff, textvariable=self.sb_font_var, width=16, font=S).pack(side=tk.LEFT, padx=4)

        # サイズ
        szf = tk.Frame(self); szf.pack(anchor="w", padx=16, pady=2)
        tk.Label(szf, text="テキストサイズ:", font=S).pack(side=tk.LEFT)
        self.sb_tsize_var = tk.StringVar(value=str(sb["text_size"]))
        tk.Entry(szf, textvariable=self.sb_tsize_var, width=5, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(szf, text="pt　ツールサイズ:", font=S).pack(side=tk.LEFT)
        self.sb_lsize_var = tk.StringVar(value=str(sb["tool_size"]))
        tk.Entry(szf, textvariable=self.sb_lsize_var, width=5, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(szf, text="pt", font=S).pack(side=tk.LEFT)

        # 文字色
        clrf = tk.Frame(self); clrf.pack(anchor="w", padx=16, pady=2)
        tk.Label(clrf, text="テキスト文字色:", font=S).pack(side=tk.LEFT)
        self.sb_tcolor_var = tk.StringVar(value=sb["text_color"])
        tk.Entry(clrf, textvariable=self.sb_tcolor_var, width=10, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(clrf, text="ツール文字色:", font=S).pack(side=tk.LEFT)
        self.sb_lcolor_var = tk.StringVar(value=sb["tool_color"])
        tk.Entry(clrf, textvariable=self.sb_lcolor_var, width=10, font=S).pack(side=tk.LEFT, padx=4)

        # 背景色
        bgf = tk.Frame(self); bgf.pack(anchor="w", padx=16, pady=2)
        tk.Label(bgf, text="テキスト背景色:", font=S).pack(side=tk.LEFT)
        self.sb_tbg_var = tk.StringVar(value=sb["text_bg"])
        tk.Entry(bgf, textvariable=self.sb_tbg_var, width=10, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(bgf, text="ツール背景色:", font=S).pack(side=tk.LEFT)
        self.sb_lbg_var = tk.StringVar(value=sb["tool_bg"])
        tk.Entry(bgf, textvariable=self.sb_lbg_var, width=10, font=S).pack(side=tk.LEFT, padx=4)

        # Y座標
        yf = tk.Frame(self); yf.pack(anchor="w", padx=16, pady=2)
        tk.Label(yf, text="テキスト Y（上端から）:", font=S).pack(side=tk.LEFT)
        self.sb_ty_var = tk.StringVar(value=str(sb["text_y"]))
        tk.Entry(yf, textvariable=self.sb_ty_var, width=5, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(yf, text="px　ツール Y（下端から）:", font=S).pack(side=tk.LEFT)
        self.sb_ly_var = tk.StringVar(value=str(sb["tool_y"]))
        tk.Entry(yf, textvariable=self.sb_ly_var, width=5, font=S).pack(side=tk.LEFT, padx=4)
        tk.Label(yf, text="px", font=S).pack(side=tk.LEFT)

        # ── ボタン ────────────────────────────
        bf = tk.Frame(self); bf.pack(pady=12)
        tk.Button(bf, text="OK",      width=8, command=self._ok).pack(side=tk.LEFT, padx=8)
        tk.Button(bf, text="キャンセル", width=8, command=self.destroy).pack(side=tk.LEFT, padx=8)

        self.transient(parent)
        self.wait_window()

    def _pick_color(self):
        color = colorchooser.askcolor(color=self.selected_color, title="背景色を選択")[1]
        if color:
            self.selected_color = color
            self.color_preview.config(bg=color)

    def _ok(self):
        def _i(s, default):
            try: return max(1, int(s.strip())) if s.strip() else default
            except ValueError: return default

        self.app.frontmost_var.set(self.frontmost_var.get())
        self.app.fullscreen_var.set(self.fullscreen_var.get())
        self.app.bgcolor = self.selected_color
        self.app.canvas.config(bg=self.selected_color)
        self.app.img_width      = _i(self.img_w_var.get(), None) if self.img_w_var.get().strip() else None
        self.app.img_height     = _i(self.img_h_var.get(), None) if self.img_h_var.get().strip() else None
        self.app.img_keep_ratio = self.ratio_var.get()

        sb = self.app.standby
        sb["font"]       = self.sb_font_var.get().strip() or DEFAULT_SB_FONT
        sb["text_size"]  = _i(self.sb_tsize_var.get(), DEFAULT_SB_TSIZE)
        sb["tool_size"]  = _i(self.sb_lsize_var.get(), DEFAULT_SB_LSIZE)
        sb["text_color"] = self.sb_tcolor_var.get().strip() or DEFAULT_SB_TCOLOR
        sb["tool_color"] = self.sb_lcolor_var.get().strip() or DEFAULT_SB_LCOLOR
        sb["text_bg"]    = self.sb_tbg_var.get().strip() or DEFAULT_SB_TBG
        sb["tool_bg"]    = self.sb_lbg_var.get().strip() or DEFAULT_SB_LBG
        sb["text_y"]     = _i(self.sb_ty_var.get(), DEFAULT_SB_TY)
        sb["tool_y"]     = _i(self.sb_ly_var.get(), DEFAULT_SB_LY)

        self.app.toggle_fullscreen()
        self.app.save_settings()
        self.app._on_resize()

        # 準備中画面表示中のときだけスロットを再描画する
        if self.app._is_standby:
            self.app._clear_all_slots()
            self.app._show_standby_hints()

        self.destroy()


# -----------------------------------------------------------------------
# 表示設定.txt 読み込み
# -----------------------------------------------------------------------
DISPLAY_FILE = "表示設定.txt"

def load_display_entries() -> List[Dict]:
    """
    表示設定.txt を読み込み、エントリのリストを返す。
    フォーマット: 表示名  画像ファイル  表示文字  色  大きさ  フォント
    """
    path = os.path.join(BASE_DIR, DISPLAY_FILE)
    if not os.path.exists(path):
        return []
    entries = []
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = line.split()
                def v(idx):
                    val = parts[idx] if idx < len(parts) else None
                    return None if (val is None or val == "-") else val
                if len(parts) < 1:
                    continue
                entries.append({
                    "label": v(0),
                    "image": v(1),
                    "text" : v(2),
                    "color": v(3),
                    "size" : int(v(4)) if v(4) else None,
                    "font" : v(5).replace("_", " ") if v(5) else None,
                })
    except Exception as e:
        print(f"表示設定読み込みエラー: {e}")
    return entries


# -----------------------------------------------------------------------
# 手動切替ダイアログ
# -----------------------------------------------------------------------
class ManualSelectDialog(tk.Toplevel):
    """
    手動切替ダイアログ。
    ・表示設定.txt の項目（特殊表示）
    ・スケジュールの項目
    をリストから選択して表示できる。
    任意画像ファイルの選択と、スケジュール復帰も提供。
    """

    # リスト行の種別
    _KIND_DISPLAY  = "display"   # 表示設定.txt
    _KIND_SCHEDULE = "schedule"  # スケジュールTSV
    _KIND_SEP      = "sep"       # セパレータ行（選択不可）

    def __init__(self, parent: tk.Tk, app: "ChimeApp"):
        super().__init__(parent)
        self.app = app
        self.title("手動切替")
        self.resizable(False, False)
        self.grab_set()

        S  = ("Helvetica", 13)
        Sg = ("Helvetica", 11)
        pad = dict(padx=16, pady=5)

        # 画像ファイル選択ボタン
        img_f = tk.Frame(self); img_f.pack(anchor="w", padx=16, pady=(10, 4))
        tk.Button(img_f, text="🖼 画像ファイルを選択…", font=S,
                  command=self._pick_image).pack(side=tk.LEFT)
        self.img_label = tk.Label(img_f, text="", font=Sg, fg="gray")
        self.img_label.pack(side=tk.LEFT, padx=8)

        tk.Label(self, text="表示項目を選択してください（ダブルクリックで即表示）:",
                 font=Sg, fg="gray").pack(anchor="w", padx=16, pady=(4, 2))

        # リストボックス
        list_f = tk.Frame(self); list_f.pack(padx=16, pady=2)
        scrollbar = tk.Scrollbar(list_f, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(list_f, width=44, height=14,
                                  font=("Helvetica", 12),
                                  yscrollcommand=scrollbar.set,
                                  selectmode=tk.SINGLE)
        scrollbar.config(command=self.listbox.yview)
        self.listbox.pack(side=tk.LEFT)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.listbox.bind("<Double-Button-1>", lambda _: self._apply_selected())
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        # エントリ構築（種別と元データを紐付け）
        self._entries: List[Dict] = []   # {"kind": ..., "entry": ...}
        self._build_list()

        # ボタン行
        bf = tk.Frame(self); bf.pack(pady=8)
        tk.Button(bf, text="表示",           width=10, command=self._apply_selected).pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="↩ スケジュールに戻る", width=18, command=self._return_to_schedule).pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="閉じる",         width=8,  command=self.destroy).pack(side=tk.LEFT, padx=6)

        self.transient(parent)
        self.wait_window()

    def _build_list(self) -> None:
        """表示設定.txt → セパレータ → スケジュール の順にリストを構築。"""
        display_entries = load_display_entries()

        if display_entries:
            self.listbox.insert(tk.END, "─── 特殊表示 ───────────────────")
            self._entries.append({"kind": self._KIND_SEP, "entry": None})
            for e in display_entries:
                label = f"  {e['label'] or ''}　{e.get('text') or ''}"
                self.listbox.insert(tk.END, label)
                self._entries.append({"kind": self._KIND_DISPLAY, "entry": e})

        if self.app.schedule_data:
            self.listbox.insert(tk.END, "─── スケジュール ───────────────")
            self._entries.append({"kind": self._KIND_SEP, "entry": None})
            for e in self.app.schedule_data:
                label = f"  {e['time']}　{e.get('text') or '(文字なし)'}"
                self.listbox.insert(tk.END, label)
                self._entries.append({"kind": self._KIND_SCHEDULE, "entry": e})

        # セパレータ行を選択不可・グレー表示にする
        for i, item in enumerate(self._entries):
            if item["kind"] == self._KIND_SEP:
                self.listbox.itemconfig(i, fg="gray", selectbackground="white",
                                        selectforeground="gray")

    def _on_select(self, event=None) -> None:
        """セパレータ行が選択されたら選択解除する。"""
        sel = self.listbox.curselection()
        if sel and self._entries[sel[0]]["kind"] == self._KIND_SEP:
            self.listbox.selection_clear(sel[0])

    def _pick_image(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("画像ファイル", "*.png *.jpg *.jpeg *.gif *.bmp"), ("すべて", "*.*")])
        if path:
            self.img_label.config(text=os.path.basename(path), fg="lightgreen")
            self.app.show_image(path)

    def _apply_selected(self) -> None:
        sel = self.listbox.curselection()
        if not sel:
            return
        item = self._entries[sel[0]]
        if item["kind"] == self._KIND_SEP:
            return
        entry = item["entry"]
        overlay = entry if entry.get("text") else None
        self.app.show_image(entry.get("image"), pending_overlay=overlay)
        self.destroy()

    def _return_to_schedule(self) -> None:
        self.app.return_to_schedule()
        self.destroy()


# -----------------------------------------------------------------------
# メインアプリ
# -----------------------------------------------------------------------
class ChimeApp:

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("学校チャイムアプリ")

        # ── 設定ロード ────────────────────────
        self._cfg: Dict[str, str] = self._load_config()

        self.bgcolor:        str            = self._cfg.get(KEY_BGCOLOR, DEFAULT_BGCOLOR)
        self.img_keep_ratio: bool           = self._cfg.get(KEY_RATIO, "True") != "False"
        self.img_width:      Optional[int]  = self._cfg_int(KEY_IMG_W)
        self.img_height:     Optional[int]  = self._cfg_int(KEY_IMG_H)

        # 準備中画面スタイル設定（まとめて dict で管理）
        self.standby: Dict = {
            "font":       self._cfg.get(KEY_SB_FONT,   DEFAULT_SB_FONT),
            "text_size":  self._cfg_int(KEY_SB_TSIZE)  or DEFAULT_SB_TSIZE,
            "tool_size":  self._cfg_int(KEY_SB_LSIZE)  or DEFAULT_SB_LSIZE,
            "text_color": self._cfg.get(KEY_SB_TCOLOR, DEFAULT_SB_TCOLOR),
            "tool_color": self._cfg.get(KEY_SB_LCOLOR, DEFAULT_SB_LCOLOR),
            "text_bg":    self._cfg.get(KEY_SB_TBG,    DEFAULT_SB_TBG),
            "tool_bg":    self._cfg.get(KEY_SB_LBG,    DEFAULT_SB_LBG),
            "text_y":     self._cfg_int(KEY_SB_TY)     or DEFAULT_SB_TY,
            "tool_y":     self._cfg_int(KEY_SB_LY)     or DEFAULT_SB_LY,
        }

        # ── Excelから今日のテキスト・ツールを取得 ──
        # パス解決: 絶対パスならそのまま、相対パスなら BASE_DIR 基準
        # ローカルで保存した絶対パスが別PCで使えない場合、
        # ファイル名だけで BASE_DIR から再探索する
        raw_excel = self._cfg.get(KEY_EXCEL, "").strip()
        self._excel_path: str = self._resolve_excel_path(raw_excel)
        self._today_info: Optional[Dict] = None  # ウィンドウ初期化後に読み込む

        # ── ウィンドウ初期サイズ ───────────────
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        _fullscreen = self._cfg.get(KEY_FULLSCREEN, "True") != "False"
        if not _fullscreen:
            w, h = 800, 600
            self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2 - 30}")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # ── Canvas ────────────────────────────
        self.canvas = tk.Canvas(root, bg=self.bgcolor, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # ── Canvas 上のウィジェット ───────────
        self.date_label = tk.Label(self.canvas, text="",
                                   font=("Helvetica", 20), bg=self.bgcolor, fg="white")
        self.time_label = tk.Label(self.canvas, text="",
                                   font=("Helvetica", 36, "bold"), bg=self.bgcolor, fg="white")
        self.next_chime_label = tk.Label(self.canvas, text="次のチャイム: なし",
                                         font=("Helvetica", 18), bg=self.bgcolor, fg="#00ff88")
        self.file_label = tk.Label(self.canvas, text="未選択",
                                   font=("Helvetica", 13), bg=self.bgcolor, fg="#aaaaaa")

        btn_cfg = dict(font=("Helvetica", 12), relief="flat", cursor="hand2",
                       bg="#333333", fg="white")
        self.btn_schedule = tk.Button(self.canvas, text="📋 スケジュール",
                                      command=self.select_schedule_file, **btn_cfg)
        self.btn_manual   = tk.Button(self.canvas, text="🖼 手動切替",
                                      command=self.open_manual_select, **btn_cfg)
        self.btn_settings = tk.Button(self.canvas, text="⚙ 設定",
                                      command=self.open_settings, **btn_cfg)

        # オーバーレイ管理
        # center: スケジュール指定テキスト（画面中央）
        # top:    準備中画面テキスト欄（上部）
        # btm:    準備中画面ツール欄（下部）
        self._overlay_center: Optional[Dict] = None  # params
        self._overlay_center_ids: Dict[str, Optional[int]] = {"text": None, "rect": None}
        self._overlay_slots: Dict[str, Dict] = {}    # slot -> {text_id, rect_id, params}
        self._is_standby: bool = False  # 準備中画面表示中か（設定ダイアログの再描画判定に使用）

        # Canvas window アイテム登録（位置は _on_resize で確定）
        self._cw_date  = self.canvas.create_window(0, 0, window=self.date_label,      anchor="nw")
        self._cw_time  = self.canvas.create_window(0, 0, window=self.time_label,      anchor="nw")
        self._cw_chime = self.canvas.create_window(0, 0, window=self.next_chime_label, anchor="nw")
        self._cw_file  = self.canvas.create_window(0, 0, window=self.file_label,      anchor="sw")
        self._cw_btn_s = self.canvas.create_window(0, 0, window=self.btn_schedule,    anchor="se")
        self._cw_btn_m = self.canvas.create_window(0, 0, window=self.btn_manual,      anchor="se")
        self._cw_btn_g = self.canvas.create_window(0, 0, window=self.btn_settings,    anchor="se")

        # 画像アイテム
        self._image_id:    Optional[int]               = None
        self._photo_ref:   Optional[ImageTk.PhotoImage] = None
        self._current_pil: Optional[Image.Image]        = None

        # スケジュール
        self.schedule_data: List[Dict]   = []
        self.chime_played_today: set     = set()
        self.schedule_file_relpath: Optional[str] = self._cfg.get(KEY_SCHEDULE)

        # チェック変数
        # KEY_FRONTMOST が未設定（None）のとき != "False" が True になるのを防ぐ
        _frontmost = self._cfg.get(KEY_FRONTMOST, "True")
        self.frontmost_var = tk.BooleanVar(value=_frontmost != "False")
        self.fullscreen_var = tk.BooleanVar(value=_fullscreen)

        # ── イベント・初期化 ──────────────────
        self.canvas.bind("<Configure>", self._on_resize)

        if not os.path.exists(CHIME_WAV_PATH):
            messagebox.showwarning("音声ファイル警告",
                                   f"チャイムファイル ({CHIME_WAV_PATH}) が見つかりません。")

        self._load_schedule()
        self._update_clock()

        threading.Thread(target=self._monitor_schedule, daemon=True).start()

        # 起動時の初期画面
        entry = self._get_current_entry()
        if entry:
            self.show_image(entry.get("image") or DEFAULT_IMAGE,
                            pending_overlay=entry if entry.get("text") else None)
        else:
            # スケジュール開始前 → 準備中画面
            self.show_image(DEFAULT_IMAGE)

        # Excel読み込みはウィンドウ完全初期化後に実行（messagebox を正常表示するため）
        self.root.after(200, self._load_excel_deferred)

    # ===================================================================
    # 設定ファイル
    # ===================================================================
    def _resolve_excel_path(self, raw: str) -> str:
        """
        config.txt の excel_file 値からファイルの実パスを解決する。
        優先順位:
          1. raw がそのまま存在する（絶対パス or カレントからの相対）
          2. BASE_DIR との結合で存在する
          3. raw のファイル名だけで BASE_DIR から探す（PC移動時のフォールバック）
          4. raw が空の場合は BASE_DIR の xlsx を自動探索（_auto_find_excel で処理）
        """
        if not raw:
            return ""  # 空の場合は _load_excel_deferred で自動探索
        candidates = [
            raw,
            os.path.join(BASE_DIR, raw),
            os.path.join(BASE_DIR, os.path.basename(raw)),
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
        # 見つからなかった場合は BASE_DIR + ファイル名（エラーダイアログはロード時に出す）
        return os.path.join(BASE_DIR, os.path.basename(raw)) if raw else ""

    def _auto_find_excel(self) -> str:
        """
        BASE_DIR 内の .xlsx ファイルを自動探索する。
        1件なら自動採用、複数なら選択ダイアログを表示、0件なら空文字を返す。
        """
        xlsx_files = [
            f for f in os.listdir(BASE_DIR)
            if f.lower().endswith(".xlsx") and not f.startswith("~$")
        ]
        if not xlsx_files:
            return ""
        if len(xlsx_files) == 1:
            found = os.path.join(BASE_DIR, xlsx_files[0])
            messagebox.showinfo(
                "日程Excelを自動検出",
                f"config.txt に excel_file の指定がないため、\n"
                f"以下のファイルを自動的に使用します。\n\n{xlsx_files[0]}\n\n"
                f"次回以降は config.txt に保存されます。")
            return found
        # 複数ある場合は選択ダイアログ
        return self._choose_excel_dialog(xlsx_files)

    def _choose_excel_dialog(self, xlsx_files: List[str]) -> str:
        """複数の xlsx から選択させるダイアログ。選択されたファイルの絶対パスを返す。"""
        dialog = tk.Toplevel(self.root)
        dialog.title("日程Excelファイルの選択")
        dialog.resizable(False, False)
        dialog.grab_set()

        tk.Label(dialog,
                 text="config.txt に excel_file の指定がありません。\n使用する日程Excelファイルを選択してください。",
                 font=("Helvetica", 12), justify="left").pack(padx=20, pady=(14, 6))

        var = tk.StringVar(value=xlsx_files[0])
        for f in xlsx_files:
            tk.Radiobutton(dialog, text=f, variable=var, value=f,
                           font=("Helvetica", 12)).pack(anchor="w", padx=28, pady=2)

        result: List[str] = []  # リストで結果を受け渡す

        def _ok():
            result.append(var.get())
            dialog.destroy()

        def _cancel():
            dialog.destroy()

        bf = tk.Frame(dialog); bf.pack(pady=12)
        tk.Button(bf, text="OK",      width=8, command=_ok).pack(side=tk.LEFT, padx=8)
        tk.Button(bf, text="キャンセル", width=8, command=_cancel).pack(side=tk.LEFT, padx=8)

        dialog.transient(self.root)
        self.root.wait_window(dialog)

        if result:
            return os.path.join(BASE_DIR, result[0])
        return ""

    def _load_excel_deferred(self) -> None:
        """
        ウィンドウ初期化完了後に Excel を読み込み、準備中画面なら再描画する。
        excel_file 未指定の場合は BASE_DIR を自動探索する。
        """
        # excel_file 未指定 → 自動探索
        if not self._excel_path:
            self._excel_path = self._auto_find_excel()
            if self._excel_path:
                # 見つかったパスを config.txt に保存（次回から自動指定）
                self._cfg[KEY_EXCEL] = os.path.relpath(self._excel_path, BASE_DIR)
                self.save_settings()
        if not self._excel_path:
            return
        self._today_info = load_today_from_excel(self._excel_path)
        # 準備中画面を表示中であればヒントを描画する
        if self._is_standby:
            self._clear_all_slots()
            self._show_standby_hints()

    def _load_config(self) -> Dict[str, str]:
        cfg: Dict[str, str] = {}
        path = os.path.join(BASE_DIR, CONFIG_FILE)
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if "=" in line:
                            k, v = line.split("=", 1)
                            cfg[k.strip()] = v.strip()
            except Exception as e:
                messagebox.showwarning("設定ファイルエラー", str(e))
        return cfg

    def _cfg_int(self, key: str) -> Optional[int]:
        v = self._cfg.get(key, "")
        try:
            return max(1, int(v)) if v.strip() else None
        except ValueError:
            return None

    def save_settings(self) -> None:
        sb = self.standby
        data = {
            KEY_SCHEDULE : self.schedule_file_relpath or "",
            KEY_EXCEL    : self._cfg.get(KEY_EXCEL, ""),
            KEY_FULLSCREEN: str(self.fullscreen_var.get()),
            KEY_FRONTMOST : str(self.frontmost_var.get()),
            KEY_BGCOLOR  : self.bgcolor,
            KEY_RATIO    : str(self.img_keep_ratio),
            KEY_IMG_W    : str(self.img_width)  if self.img_width  else "",
            KEY_IMG_H    : str(self.img_height) if self.img_height else "",
            KEY_SB_FONT  : sb["font"],
            KEY_SB_TSIZE : str(sb["text_size"]),
            KEY_SB_LSIZE : str(sb["tool_size"]),
            KEY_SB_TCOLOR: sb["text_color"],
            KEY_SB_LCOLOR: sb["tool_color"],
            KEY_SB_TBG   : sb["text_bg"],
            KEY_SB_LBG   : sb["tool_bg"],
            KEY_SB_TY    : str(sb["text_y"]),
            KEY_SB_LY    : str(sb["tool_y"]),
        }
        with open(os.path.join(BASE_DIR, CONFIG_FILE), "w", encoding="utf-8") as f:
            for k, v in data.items():
                f.write(f"{k}={v}\n")

    # ===================================================================
    # レイアウト
    # ===================================================================
    def _on_resize(self, event=None) -> None:
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw < 2 or ch < 2:
            return

        bg = self.bgcolor
        fg = contrast_color(bg)
        chime_fg = "#007a3d" if fg == "black" else "#00ff88"
        for lbl in (self.date_label, self.time_label, self.file_label):
            lbl.config(bg=bg, fg=fg)
        self.next_chime_label.config(bg=bg, fg=chime_fg)
        for btn in (self.btn_schedule, self.btn_manual, self.btn_settings):
            btn.config(bg="#333333", fg="white")

        # 上部中央: 日付 → 時刻 → 次チャイム
        cx = cw // 2
        for cw_id in (self._cw_date, self._cw_time, self._cw_chime):
            self.canvas.itemconfigure(cw_id, anchor="n")
        self.canvas.coords(self._cw_date, cx, 10)
        dh = self.date_label.winfo_reqheight()
        self.canvas.coords(self._cw_time, cx, 10 + dh + 2)
        th = self.time_label.winfo_reqheight()
        self.canvas.coords(self._cw_chime, cx, 10 + dh + 2 + th + 4)

        # 左下: ファイル名
        self.canvas.coords(self._cw_file, 10, ch - 10)

        # 右下: ボタン3つ
        margin = 8
        x = cw - 10
        for cw_id, btn in ((self._cw_btn_g, self.btn_settings),
                            (self._cw_btn_m, self.btn_manual),
                            (self._cw_btn_s, self.btn_schedule)):
            self.canvas.coords(cw_id, x, ch - 10)
            x -= btn.winfo_reqwidth() + margin

        # 画像リサイズ
        if self._current_pil:
            self._draw_image(cw, ch)

        # オーバーレイ再描画
        self._redraw_all_overlays(cw, ch)

    # ===================================================================
    # 画像表示
    # ===================================================================
    def _get_current_entry(self) -> Optional[Dict]:
        """現在時刻以前の直近スケジュールエントリを返す。"""
        now_min = datetime.now().strftime("%H:%M")
        candidates = [e for e in self.schedule_data if e["time"] <= now_min]
        if not candidates:
            return None
        return max(candidates, key=lambda e: e["time"])

    def show_image(self, image_path: Optional[str],
                   pending_overlay: Optional[Dict] = None) -> None:
        """
        画像を表示する。
        pending_overlay が指定されていれば描画後に中央オーバーレイを表示。
        準備中画像（DEFAULT_IMAGE）の場合は準備中画面のスロット表示を行う。
        """
        # 画像ロード
        resolved = self._resolve_path(image_path)
        if resolved:
            try:
                self._current_pil = Image.open(resolved)
            except Exception as e:
                print(f"画像ロードエラー: {e}")
                self._current_pil = None
        else:
            self._current_pil = None

        is_standby = (
            image_path is not None and
            os.path.basename(str(image_path)) == os.path.basename(DEFAULT_IMAGE)
        )
        self._is_standby = is_standby

        # オーバーレイクリア
        self._clear_center_overlay()
        if not is_standby:
            self._clear_all_slots()

        def _after_draw():
            if pending_overlay:
                self._set_center_overlay(
                    text      = pending_overlay.get("text"),
                    color     = pending_overlay.get("color") or "white",
                    size      = pending_overlay.get("size")  or 28,
                    font_name = pending_overlay.get("font")  or "Helvetica",
                )
            elif is_standby:
                self._show_standby_hints()

        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw > 1 and ch > 1:
            self._draw_image(cw, ch)
            _after_draw()
        else:
            def _deferred():
                cw2 = self.canvas.winfo_width()
                ch2 = self.canvas.winfo_height()
                self._draw_image(cw2, ch2)
                _after_draw()
            self.canvas.after(50, _deferred)

    def _resolve_path(self, image_path: Optional[str]) -> Optional[str]:
        if not image_path:
            return None
        p1 = os.path.join(BASE_DIR, image_path)
        if os.path.exists(p1):
            return p1
        if os.path.exists(image_path):
            return image_path
        return None

    def _draw_image(self, cw: int, ch: int) -> None:
        if self._image_id:
            self.canvas.delete(self._image_id)
            self._image_id = None

        if not self._current_pil or cw < 2 or ch < 2:
            self._redraw_all_overlays(cw, ch)
            return

        img = self._current_pil
        iw, ih = img.size
        tw = min(self.img_width  or cw, cw)
        th = min(self.img_height or ch, ch)
        if self.img_keep_ratio:
            ratio = tw / iw
            nw = max(1, int(iw * ratio))
            nh = max(1, int(ih * ratio))
        else:
            nw, nh = max(1, tw), max(1, th)

        resized = img.resize((nw, nh), Image.Resampling.LANCZOS)
        self._photo_ref = ImageTk.PhotoImage(resized)
        self._image_id = self.canvas.create_image(
            cw // 2, ch // 2, image=self._photo_ref, anchor="center")

        self._raise_widgets()
        self._redraw_all_overlays(cw, ch)

    def _raise_widgets(self) -> None:
        for cw_id in (self._cw_date, self._cw_time, self._cw_chime,
                      self._cw_file, self._cw_btn_s, self._cw_btn_m, self._cw_btn_g):
            self.canvas.tag_raise(cw_id)
        # center overlay
        if self._overlay_center_ids["rect"]:
            self.canvas.tag_raise(self._overlay_center_ids["rect"])
        if self._overlay_center_ids["text"]:
            self.canvas.tag_raise(self._overlay_center_ids["text"])
        # slot overlays
        for slot in self._overlay_slots.values():
            if slot.get("rect_id"): self.canvas.tag_raise(slot["rect_id"])
            if slot.get("text_id"): self.canvas.tag_raise(slot["text_id"])

    # ===================================================================
    # 準備中画面ヒント表示
    # ===================================================================
    def _show_standby_hints(self) -> None:
        """
        _today_info の text / tool を画面の上部・下部に表示する。
        スケジュール開始前（準備中画面）のみ呼ばれる。
        """
        info = self._today_info
        if not info:
            return
        sb = self.standby
        ch = self.canvas.winfo_height()

        text_val = info.get("text", "").strip()
        tool_val = info.get("tool", "").strip()

        if text_val:
            self._set_slot(
                slot      = "top",
                text      = f"テキスト：{text_val}",
                y         = sb["text_y"],
                color     = sb["text_color"],
                size      = sb["text_size"],
                font_name = sb["font"],
                bg_color  = sb["text_bg"],
            )
        if tool_val:
            self._set_slot(
                slot      = "btm",
                text      = f"ツール：{tool_val}",
                y         = ch - sb["tool_y"] if ch > 100 else ch // 2,
                color     = sb["tool_color"],
                size      = sb["tool_size"],
                font_name = sb["font"],
                bg_color  = sb["tool_bg"],
            )

    # ===================================================================
    # オーバーレイ管理
    # ===================================================================
    def _clear_center_overlay(self) -> None:
        for key in ("text", "rect"):
            oid = self._overlay_center_ids.get(key)
            if oid:
                self.canvas.delete(oid)
                self._overlay_center_ids[key] = None
        self._overlay_center = None

    def _set_center_overlay(self, text: Optional[str],
                             color: str = "white", size: int = 28,
                             font_name: str = "Helvetica") -> None:
        self._clear_center_overlay()
        if not text:
            return
        self._overlay_center = dict(text=text, color=color, size=size, font_name=font_name)
        self._draw_center_overlay()

    def _draw_center_overlay(self) -> None:
        p = self._overlay_center
        if not p or not p.get("text"):
            return
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        px, py = cw // 2, ch // 2
        tid = self.canvas.create_text(
            px, py, text=p["text"],
            font=(p["font_name"], p["size"], "bold"),
            fill=p["color"], anchor="center")
        self._overlay_center_ids["text"] = tid
        bbox = self.canvas.bbox(tid)
        if bbox:
            rid = self.canvas.create_rectangle(
                bbox[0] - 8, bbox[1] - 4, bbox[2] + 8, bbox[3] + 4,
                fill="#000000", outline="", stipple="gray50")
            self._overlay_center_ids["rect"] = rid
            self.canvas.tag_lower(rid, tid)
        self._raise_widgets()

    def _clear_slot(self, slot: str) -> None:
        s = self._overlay_slots.pop(slot, {})
        for key in ("rect_id", "text_id"):
            if s.get(key):
                self.canvas.delete(s[key])

    def _clear_all_slots(self) -> None:
        for slot in list(self._overlay_slots.keys()):
            self._clear_slot(slot)

    def _set_slot(self, slot: str, text: str, y: int,
                  color: str = "white", size: int = 28,
                  font_name: str = "Helvetica",
                  bg_color: Optional[str] = None) -> None:
        self._clear_slot(slot)
        cw = self.canvas.winfo_width()
        tid = self.canvas.create_text(
            cw // 2, y, text=text,
            font=(font_name, size, "bold"),
            fill=color, anchor="center")
        rid = None
        if bg_color:
            bbox = self.canvas.bbox(tid)
            if bbox:
                rid = self.canvas.create_rectangle(
                    bbox[0] - 10, bbox[1] - 5, bbox[2] + 10, bbox[3] + 5,
                    fill=bg_color, outline="")
                self.canvas.tag_lower(rid, tid)
        self._overlay_slots[slot] = {
            "text_id": tid, "rect_id": rid,
            "params": dict(text=text, y=y, color=color,
                           size=size, font_name=font_name, bg_color=bg_color),
        }
        self._raise_widgets()

    def _redraw_all_overlays(self, cw: int, ch: int) -> None:
        """リサイズ時にすべてのオーバーレイを再描画。"""
        if self._overlay_center:
            self._clear_center_overlay_ids()
            self._draw_center_overlay()

        sb = self.standby
        for slot, s in list(self._overlay_slots.items()):
            p = s.get("params", {})
            y = (ch - sb["tool_y"] if ch > 100 else ch // 2) if slot == "btm" else p.get("y", ch // 2)
            self._set_slot(
                slot=slot, text=p["text"], y=y,
                color=p.get("color", "white"),
                size=p.get("size", 28),
                font_name=p.get("font_name", "Helvetica"),
                bg_color=p.get("bg_color"),
            )

    def _clear_center_overlay_ids(self) -> None:
        """canvas アイテムだけ削除（params は残す）。"""
        for key in ("text", "rect"):
            oid = self._overlay_center_ids.get(key)
            if oid:
                self.canvas.delete(oid)
                self._overlay_center_ids[key] = None

    # ===================================================================
    # 時計
    # ===================================================================
    def _update_clock(self) -> None:
        now = datetime.now()
        wareki = to_wareki(now)
        self.date_label.config(
            text=f"{now.year}（{wareki}）年{now.month}月{now.day}日")
        self.time_label.config(text=now.strftime("%H:%M:%S"))
        self._update_next_chime(now)
        self.root.after(1000, self._update_clock)

    def _update_next_chime(self, now: datetime) -> None:
        now_str = now.strftime("%H:%M")
        future = [e for e in self.schedule_data if e["time"] > now_str]
        self.next_chime_label.config(
            text=f"次のチャイム: {future[0]['time']}" if future else "次のチャイム: なし")

    # ===================================================================
    # スケジュール監視
    # ===================================================================
    def _monitor_schedule(self) -> None:
        def play_chime():
            if not os.path.exists(CHIME_WAV_PATH):
                return
            try:
                samplerate, data = wavfile.read(CHIME_WAV_PATH)
                sd.play(data, samplerate)
                sd.wait()
            except Exception as e:
                print(f"チャイム再生エラー: {e}")

        last_checked = ""
        while True:
            now = datetime.now()
            now_min = now.strftime("%H:%M")
            if now.second < 10 and now_min != last_checked:
                for entry in self.schedule_data:
                    if entry["time"] == now_min and now_min not in self.chime_played_today:
                        if self.frontmost_var.get():
                            self.root.deiconify()
                            self.root.lift()
                            self.root.attributes("-topmost", True)
                            self.root.after(1000, lambda: self.root.attributes("-topmost", False))
                        self.canvas.after(0, lambda e=entry: self._apply_schedule_entry(e))
                        threading.Thread(target=play_chime, daemon=True).start()
                        self.chime_played_today.add(now_min)
                last_checked = now_min
            time.sleep(2)

    def _apply_schedule_entry(self, entry: Dict) -> None:
        overlay = entry if entry.get("text") else None
        self.show_image(entry.get("image"), pending_overlay=overlay)

    # ===================================================================
    # スケジュール復帰
    # ===================================================================
    def return_to_schedule(self) -> None:
        """現在時刻のスケジュールエントリに戻る。"""
        entry = self._get_current_entry()
        if entry:
            self._apply_schedule_entry(entry)
        else:
            self.show_image(DEFAULT_IMAGE)

    # ===================================================================
    # スケジュール読み込み
    # ===================================================================
    def select_schedule_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if path:
            self.schedule_file_relpath = os.path.relpath(path, BASE_DIR)
            self.file_label.config(text=os.path.basename(path), fg="#aaaaaa")
            self.save_settings()
            self._load_schedule()

    def _load_schedule(self) -> None:
        if not self.schedule_file_relpath:
            self.file_label.config(text="未選択", fg="#aaaaaa")
            return
        abs_path = os.path.join(BASE_DIR, self.schedule_file_relpath)
        if not os.path.exists(abs_path):
            messagebox.showerror("エラー", f"スケジュールファイルが見つかりません:\n{abs_path}")
            self.schedule_file_relpath = None
            self.file_label.config(text="ファイルが見つかりません", fg="red")
            return
        self.schedule_data.clear()
        try:
            with open(abs_path, "r", encoding="utf-8") as f:
                for line in f:
                    entry = parse_schedule_line(line)
                    if entry:
                        self.schedule_data.append(entry)
            self.chime_played_today.clear()
            self.file_label.config(
                text=f"{os.path.basename(abs_path)} (読み込み成功)", fg="lightgreen")
        except Exception as e:
            messagebox.showerror("エラー", f"読み込みエラー:\n{e}")
            self.file_label.config(text="読み込み失敗", fg="red")

    # ===================================================================
    # 手動切替・設定
    # ===================================================================
    def open_manual_select(self) -> None:
        ManualSelectDialog(self.root, self)

    def open_settings(self) -> None:
        SettingsDialog(self.root, self)

    def toggle_fullscreen(self) -> None:
        if self.fullscreen_var.get():
            self.root.update()
            self.root.state("zoomed")
        else:
            self.root.state("normal")
            self.root.geometry("800x600")


# -----------------------------------------------------------------------
# エントリポイント
# -----------------------------------------------------------------------
if __name__ == "__main__":
    if not hasattr(Image, "Resampling"):
        Image.Resampling = Image

    root = tk.Tk()
    if os.path.exists(os.path.join(BASE_DIR, "bell_icon.ico")):
        try:
            root.iconbitmap(os.path.join(BASE_DIR, "bell_icon.ico"))
        except Exception:
            pass

    app = ChimeApp(root)

    if app.fullscreen_var.get():
        root.state("zoomed")
        root.update_idletasks()
        app._on_resize()

    root.mainloop()
