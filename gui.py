"""Hangcha desktop app: universal PDF scan, business review, and approval-gated Excel write.

No unattended write: scheduled tasks run read-only and writes require explicit
in-app confirmation after reviewing a current business plan.
"""
from __future__ import annotations

import json
import hashlib
import os
import queue
import re
import sqlite3
import subprocess
import sys
import threading
import traceback
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog, font as tkfont

from gui_scanner import (HINTS, VERSION, ScanRow, export_csv, flag_duplicate_content,
                         scan_sources, summarize, DEFAULT_EXCLUDED, mark_pending_invoice_status)
import gui_business
from gui_notes import NoteStore
from gui_excel_notes import read_order_notes, brief_note
import gui_schedule
from gui_archive import ArchiveStore, path_key, file_sha256, plan_archive
import automation as legacy
from gui_tree import make_tree, matches_filter
from gui_archive_tree import archive_tree
from gui_scan_cache import ScanCache, incremental_scan, changed_since_scan, CACHE_SCHEMA_VERSION
from gui_archive_repair import plan_historical_delivery_repair

GUI_VERSION = "1.5.34"
BASE = Path(__file__).resolve().parent
# Each Windows login gets independent GUI sources, salesperson, and Excel target.
GUI_CONFIG = Path(os.environ.get("LOCALAPPDATA", str(BASE))) / "HangchaSalesAssistant" / "gui_config.json"

PALETTE = {
    "bg": "#F4F7FB", "white": "#FFFFFF", "ink": "#152238", "muted": "#66758B",
    "navy": "#15375F", "blue": "#225B98", "line": "#DAE2EC",
    "green": "#087B55", "yellow": "#AF6A12", "red": "#B33F3F",
    "striped": "#F7F9FC", "select": "#D9E9FA",
}


def _natural_sort_key(value: str):
    """Compare file names/IDs as humans do: 9 < 10, including H2026... IDs."""
    parts = re.split(r"(\d+)", str(value).casefold())
    return tuple((1, int(p)) if p.isdigit() else (0, p) for p in parts)


def _table_sort_key(column: str, value: str):
    """Return a correctly typed key; None means blank/invalid (always shown last)."""
    value = str(value).strip()
    if not value:
        return None
    if column == "day":
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y", "%d-%m-%y"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
        return None
    if column in {"gross", "net"}:
        num = re.sub(r"[^0-9,.\-]", "", value)
        if not num:
            return None
        # Both decimal conventions occur in European sales documents.
        # For grouped amounts, the last separator is the decimal separator.
        if "," in num and "." in num:
            num = num.replace(".", "").replace(",", ".") if num.rfind(",") > num.rfind(".") else num.replace(",", "")
        elif "," in num:
            num = num.replace(",", ".") if num.count(",") == 1 else num.replace(",", "")
        elif num.count(".") > 1:
            num = num.replace(".", "")
        try:
            return Decimal(num)
        except InvalidOperation:
            return None
    return _natural_sort_key(value)


def _reveal_command(pdf: Path, platform: str | None = None) -> list[str]:
    """Build command-line arguments for non-Windows file managers.

    Windows Explorer uses the Shell API below to avoid fragile /select, parsing,
    which can otherwise open the default Documents directory instead.
    """
    platform = sys.platform if platform is None else platform
    if platform == "darwin":
        return ["open", "-R", str(pdf)]
    return ["xdg-open", str(pdf.parent)]


def _reveal_windows_pdf(pdf: Path) -> None:
    """Ask the Windows Shell to open the *actual* containing folder and select PDF.

    SHOpenFolderAndSelectItems takes PIDLs instead of an explorer.exe command
    line, so spaces, Unicode, parentheses and OneDrive paths are not mistaken
    for command-line switches or routed to the default Documents directory.
    """
    import ctypes
    from ctypes import wintypes

    shell32 = ctypes.WinDLL("shell32", use_last_error=True)
    ole32 = ctypes.WinDLL("ole32", use_last_error=True)

    ole32.CoInitializeEx.argtypes = [ctypes.c_void_p, wintypes.DWORD]
    ole32.CoInitializeEx.restype = ctypes.c_long
    ole32.CoUninitialize.argtypes = []
    ole32.CoUninitialize.restype = None
    ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    ole32.CoTaskMemFree.restype = None

    shell32.SHParseDisplayName.argtypes = [
        wintypes.LPCWSTR, ctypes.c_void_p,
        ctypes.POINTER(ctypes.c_void_p), wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD),
    ]
    shell32.SHParseDisplayName.restype = ctypes.c_long
    shell32.ILFindLastID.argtypes = [ctypes.c_void_p]
    shell32.ILFindLastID.restype = ctypes.c_void_p
    shell32.SHOpenFolderAndSelectItems.argtypes = [
        ctypes.c_void_p, ctypes.c_uint,
        ctypes.POINTER(ctypes.c_void_p), wintypes.DWORD,
    ]
    shell32.SHOpenFolderAndSelectItems.restype = ctypes.c_long

    # GUI callbacks run on the Tk main thread; initialize COM for this call
    # without uninitializing a COM apartment created by another library.
    hr_init = ole32.CoInitializeEx(None, 2)  # COINIT_APARTMENTTHREADED
    initialized_here = hr_init in (0, 1)  # S_OK / S_FALSE
    if hr_init not in (0, 1, -2147417850):  # RPC_E_CHANGED_MODE is harmless
        raise OSError(f"Windows COM initialization failed: 0x{hr_init & 0xffffffff:08X}")

    folder_pidl = ctypes.c_void_p()
    file_pidl = ctypes.c_void_p()
    try:
        for target_path, pidl in ((str(pdf.parent), folder_pidl), (str(pdf), file_pidl)):
            hr = shell32.SHParseDisplayName(target_path, None, ctypes.byref(pidl), 0, None)
            if hr < 0 or not pidl.value:
                raise OSError(f"Windows cannot locate {target_path}: 0x{hr & 0xffffffff:08X}")

        relative_child = shell32.ILFindLastID(file_pidl)
        if not relative_child:
            raise OSError(f"Windows cannot identify the file in its folder: {pdf}")
        child_array = (ctypes.c_void_p * 1)(relative_child)
        hr = shell32.SHOpenFolderAndSelectItems(folder_pidl, 1, child_array, 0)
        if hr < 0:
            raise OSError(f"Explorer could not select {pdf.name}: 0x{hr & 0xffffffff:08X}")
    finally:
        if file_pidl.value:
            ole32.CoTaskMemFree(file_pidl)
        if folder_pidl.value:
            ole32.CoTaskMemFree(folder_pidl)
        if initialized_here:
            ole32.CoUninitialize()


def _row_identifier(row: ScanRow) -> str:
    """Identifier for this document type, never confuse invoice and contract."""
    if row.kind in ('发票', '贷项通知单'):
        return row.invoice or '—'
    if row.kind == '发货单':
        return row.delivery_no or '—'
    return row.contract or '—' 


def _sort_scan_items(items: list[tuple[int, ScanRow]], column: str, descending: bool):
    """Sort display rows only; preserve scan order for equal keys and blank rows."""
    populated, missing = [], []
    for idx, row in items:
        raw = _row_identifier(row) if column == "identifier" else getattr(row, column, '')
        key = _table_sort_key(column, raw)
        (missing if key is None else populated).append((idx, row, key))
    populated.sort(key=lambda item: item[2], reverse=descending)
    return [(idx, row) for idx, row, _ in populated + missing]


class HangchaGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"HCE Oracle | 订单合同管理 v{GUI_VERSION}")
        self.geometry("1460x900")
        self.minsize(1080, 660)
        self.configure(bg=PALETTE["bg"])
        self.sources: list[dict] = []
        self.rows: list[ScanRow] = []
        self.events: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.business_worker: threading.Thread | None = None
        self.business_plan = None
        self._excel_note_cache_key = None
        self._excel_note_cache = {}
        self._excel_note_error = ""
        self._preview_cfg = None
        self.business_busy = False
        self.scanned_complete = False
        self._scan_running = False
        self.scheduled_preview = '--scheduled-preview' in sys.argv
        self.scheduled_run = self.scheduled_preview
        self._shown_rows: dict[str, ScanRow] = {}
        self._tree_members: dict[str, list[ScanRow]] = {}
        self._tree_parent_items: set[str] = set()
        self._tree_orphan_items: set[str] = set()
        self._tree_open_contracts: set[str] = set()
        self._tree_forced_closed: set[str] = set()
        self._sort_column: str | None = None
        self._sort_descending = True
        self._column_titles: dict[str, str] = {}
        self.archive_profile = None
        self.archive_store = None
        self.archived_map = {}
        self._cached_delivery_numbers = {}
        self.checked_paths: set[str] = set()
        self.opened_for_review: dict[str, str] = {}
        self.scan_hashes: dict[str, str] = {}
        self._scan_file_metadata = {}
        self._scan_report = {}
        self.scan_mode_var = tk.StringVar(value='incremental')
        self._last_archive_paths: list[str] = []
        self.font_preset_var = tk.StringVar(value="标准")
        self._init_ui_fonts()
        self._build_style()
        self._build_ui()
        self._load_config()
        self.after(180, self._apply_saved_layout)
        self.archive_profile = self._get_config()["archive_profile"]
        self._current_archive_profile = self._employee_profile_key()
        self.archive_store = ArchiveStore(GUI_CONFIG.parent / "archive_state.sqlite3", self._current_archive_profile)
        self.note_store = NoteStore(GUI_CONFIG.parent / "remark_drafts.sqlite3", self._current_archive_profile)
        self._load_cached_archive_numbers()
        self._last_config = json.loads(json.dumps(self._get_config()))
        # Editing scan inputs immediately invalidates access to the old Teams
        # invoice list. No need to force a new scan for unrelated UI settings.
        for var in (self.salesperson_var, self.ignore_var, self.excel_var):
            var.trace_add('write', lambda *_: self._refresh_teams_access())
        self._refresh_teams_access()
        self.after(100, self._poll_events)
        if self.scheduled_preview:
            self.after(300, self._start_scan)
        else:
            self.after(1000, self._warn_legacy_schedule)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # Keep the font choice within the app; do not override Windows display/DPI
    # settings. Tk points scale using the user's actual monitor scale factor.
    _FONT_PRESETS = {
        "紧凑": {"body": 9, "small": 9, "table": 9, "heading": 9, "section": 11, "title": 12,
                 "stat": 19, "app_title": 19, "mono": 9, "rowheight": 32},
        "标准": {"body": 10, "small": 10, "table": 10, "heading": 10, "section": 12, "title": 12,
                 "stat": 20, "app_title": 20, "mono": 10, "rowheight": 38},
        "舒适": {"body": 11, "small": 11, "table": 11, "heading": 11, "section": 13, "title": 13,
                 "stat": 22, "app_title": 21, "mono": 11, "rowheight": 43},
    }

    def _init_ui_fonts(self):
        available = {name.casefold(): name for name in tkfont.families(self)}
        self._chinese_ui_family = next((available[name.casefold()] for name in
             ("Microsoft YaHei UI", "Microsoft YaHei", "Noto Sans CJK SC", "Noto Sans CJK JP", "PingFang SC", "Segoe UI")
             if name.casefold() in available), "TkDefaultFont")
        self._latin_ui_family = available.get("segoe ui", self._chinese_ui_family)
        self._fonts = {}
        for key, family, weight in (
            ("body", self._chinese_ui_family, "normal"),
            ("small", self._chinese_ui_family, "normal"),
            ("table", self._chinese_ui_family, "normal"),
            ("table_bold", self._chinese_ui_family, "bold"),
            ("heading", self._chinese_ui_family, "bold"),
            ("section", self._chinese_ui_family, "bold"),
            ("title", self._chinese_ui_family, "bold"),
            ("stat", self._latin_ui_family, "bold"),
            ("app_title", self._latin_ui_family, "bold"),
            ("mono", "Consolas" if "consolas" in available else "Courier", "normal"),
        ):
            self._fonts[key] = tkfont.Font(self, family=family, size=10, weight=weight)
        self._apply_font_preset("标准", persist=False)

    def _apply_font_preset(self, preset: str, persist: bool = True):
        if preset not in self._FONT_PRESETS:
            preset = "标准"
        self.font_preset_var.set(preset)
        sizes = self._FONT_PRESETS[preset]
        for name, widget_font in self._fonts.items():
            size = sizes["table"] if name == "table_bold" else sizes[name]
            widget_font.configure(size=size)
        if hasattr(self, "table"):
            s = ttk.Style(self)
            s.configure("Data.Treeview", rowheight=sizes["rowheight"])
            s.configure("Data.Treeview.Heading", padding=(8, 9 if preset != "紧凑" else 7))
            # Existing table items and archive operations remain unchanged.
            self.table.update_idletasks()
        if persist:
            # Visual preferences must NOT invalidate a verified business plan
            # or silently affect archive safety snapshots.
            try:
                GUI_CONFIG.parent.mkdir(parents=True, exist_ok=True)
                GUI_CONFIG.write_text(json.dumps(self._get_config(), ensure_ascii=False, indent=2), encoding="utf-8")
                if hasattr(self, "_last_config"):
                    self._last_config["font_preset"] = preset
            except OSError as exc:
                messagebox.showwarning("界面设置保存失败", str(exc), parent=self)

    def _build_style(self):
        s = ttk.Style(self)
        try: s.theme_use("clam")
        except tk.TclError: pass
        s.configure("TButton", font=self._fonts["body"])
        s.configure("TCheckbutton", font=self._fonts["body"])
        s.configure("TCombobox", font=self._fonts["body"])
        s.configure("TEntry", font=self._fonts["body"])
        s.configure("Main.TFrame", background=PALETTE["bg"])
        s.configure("Card.TFrame", background=PALETTE["white"])
        s.configure("Main.TLabel", background=PALETTE["bg"], foreground=PALETTE["ink"], font=self._fonts["body"])
        s.configure("Card.TLabel", background=PALETTE["white"], foreground=PALETTE["ink"], font=self._fonts["body"])
        s.configure("Hint.TLabel", background=PALETTE["white"], foreground=PALETTE["muted"], font=self._fonts["small"])
        s.configure("Small.TLabel", background=PALETTE["bg"], foreground=PALETTE["muted"], font=self._fonts["small"])
        s.configure("Blue.TButton", font=self._fonts["heading"], foreground="white", background=PALETTE["blue"], padding=(14, 9), borderwidth=0)
        s.map("Blue.TButton", background=[("active", PALETTE["navy"]), ("disabled", "#AEB9C6")])
        s.configure("Soft.TButton", font=self._fonts["small"], padding=(10, 8), borderwidth=1)
        s.configure("Main.TEntry", padding=6)
        s.configure("Main.TCombobox", padding=5)
        s.configure("Data.Treeview", font=self._fonts["table"], rowheight=self._FONT_PRESETS["标准"]["rowheight"], background="white", fieldbackground="white", foreground=PALETTE["ink"], borderwidth=0)
        s.configure("Data.Treeview.Heading", font=self._fonts["table_bold"], background="#E9EFF7", foreground=PALETTE["navy"], relief="flat", padding=(7, 8))
        s.map("Data.Treeview", background=[("selected", PALETTE["select"])], foreground=[("selected", PALETTE["ink"])])
        s.configure("Horizontal.TProgressbar", troughcolor="#E3EAF3", background=PALETTE["blue"], borderwidth=0)

    def _build_ui(self):
        # The source/Excel configuration is secondary: an adjustable sidebar
        # leaves the contract-tree workflow most of the vertical space.
        head = tk.Frame(self, bg=PALETTE["navy"], height=75)
        head.pack(fill="x")
        tk.Label(head, text="HCE Oracle", font=self._fonts["app_title"],
                 bg=PALETTE["navy"], fg="white").pack(anchor="w", padx=20, pady=(9, 0))
        tk.Label(head, text="订单合同树 · PDF 核对 · 业务异常提醒 · 人工确认后写入",
                 font=self._fonts["small"], bg=PALETTE["navy"], fg="#C8D9ED").pack(anchor="w", padx=20, pady=(0, 8))

        frame = ttk.Frame(self, style="Main.TFrame", padding=(10, 8))
        frame.pack(fill="both", expand=True)
        self.workspace = ttk.Panedwindow(frame, orient="horizontal")
        self.workspace.pack(fill="both", expand=True)
        self.sidebar_frame = ttk.Frame(self.workspace, style="Card.TFrame")
        self.right_main = ttk.Frame(self.workspace, style="Main.TFrame", padding=(12, 1, 0, 0))
        self.workspace.add(self.sidebar_frame, weight=0)
        self.workspace.add(self.right_main, weight=1)
        self._sidebar_visible = True
        self._sidebar_width = 320
        self._stats_details_visible = False
        self._detail_visible = False
        self.workspace.bind('<ButtonRelease-1>', self._on_workspace_drag_end, add='+')

        # The sidebar scrolls independently on smaller laptop screens.
        self.sidebar_canvas = tk.Canvas(self.sidebar_frame, background="white",
                                        width=294, borderwidth=0, highlightthickness=0)
        self.sidebar_scrollbar = ttk.Scrollbar(self.sidebar_frame, orient="vertical",
                                              command=self.sidebar_canvas.yview)
        self.sidebar_canvas.configure(yscrollcommand=self.sidebar_scrollbar.set)
        self.sidebar_scrollbar.pack(side="right", fill="y")
        self.sidebar_canvas.pack(side="left", fill="both", expand=True)
        self.sidebar_content = ttk.Frame(self.sidebar_canvas, style="Card.TFrame", padding=(12, 10))
        self._sidebar_canvas_item = self.sidebar_canvas.create_window(
                   (0, 0), window=self.sidebar_content, anchor="nw")
        self.sidebar_content.bind("<Configure>", self._on_sidebar_content_resize)
        self.sidebar_canvas.bind("<Configure>", self._on_sidebar_canvas_resize)

        # LEFT: multi-source configuration. All the existing entry variables,
        # directory selection callbacks, and Excel target handling remain intact.
        sidebar_top = ttk.Frame(self.sidebar_content, style="Card.TFrame")
        sidebar_top.pack(fill="x", pady=(0, 8))
        tk.Label(sidebar_top, text="工作设置", font=self._fonts["section"],
                 fg=PALETTE["navy"], bg="white").pack(side="left")
        ttk.Button(sidebar_top, text="收起 ◂", style="Soft.TButton",
                   command=self._toggle_sidebar).pack(side="right")

        source_header = ttk.Frame(self.sidebar_content, style="Card.TFrame")
        source_header.pack(fill="x", pady=(0, 5))
        tk.Label(source_header, text="PDF 来源目录", font=self._fonts["heading"],
                 fg=PALETTE["navy"], bg="white").pack(side="left")
        ttk.Button(source_header, text="＋ 添加", command=self._add_source,
                   style="Soft.TButton").pack(side="right")
        self.source_tree = ttk.Treeview(self.sidebar_content, columns=("path", "hint"),
                     displaycolumns=("path",), show="headings", height=3,
                     style="Data.Treeview", selectmode="browse")
        self.source_tree.heading("path", text="扫描目录（支持多个目录）")
        self.source_tree.heading("hint", text="依据 PDF 正文判定")
        self.source_tree.column("path", width=260, minwidth=150, stretch=True)
        self.source_tree.column("hint", width=0, stretch=False)
        self.source_tree.pack(fill="x", pady=(0, 4))
        self.source_tree.bind("<<TreeviewSelect>>", self._select_source)
        hintrow = ttk.Frame(self.sidebar_content, style="Card.TFrame")
        hintrow.pack(fill="x", pady=(2, 9))
        self.remove_btn = ttk.Button(hintrow, text="移除选中", style="Soft.TButton",
                                     command=self._remove_source)
        self.remove_btn.pack(side="left")
        self.hint_var = tk.StringVar(value="自动识别")
        ttk.Label(hintrow, text="类型和客户只认 PDF 正文", style="Hint.TLabel").pack(
            side="right", padx=(8, 0))

        ttk.Separator(self.sidebar_content, orient="horizontal").pack(fill="x", pady=(3, 8))
        tk.Label(self.sidebar_content, text="个人 Excel", font=self._fonts["heading"],
                 fg=PALETTE["navy"], bg="white").pack(anchor="w", pady=(0, 4))
        ttk.Label(self.sidebar_content, text="每位员工独立维护", style="Hint.TLabel").pack(anchor="w")
        self.excel_var = tk.StringVar(value="")
        self.rename_date_var = tk.BooleanVar(value=False)
        ttk.Entry(self.sidebar_content, textvariable=self.excel_var,
                  style="Main.TEntry").pack(fill="x", pady=(5, 5))
        excel_buttons = ttk.Frame(self.sidebar_content, style="Card.TFrame")
        excel_buttons.pack(fill="x", pady=(0, 7))
        ttk.Button(excel_buttons, text="选择 Excel", command=self._choose_excel,
                   style="Soft.TButton").pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(excel_buttons, text="新建位置", command=self._choose_new_excel,
                   style="Soft.TButton").pack(side="left", fill="x", expand=True)
        ttk.Checkbutton(self.sidebar_content, text="写入后更新 Excel 文件名日期",
                        variable=self.rename_date_var,
                        command=self._invalidate_plan).pack(anchor="w", pady=(0, 10))

        ttk.Separator(self.sidebar_content, orient="horizontal").pack(fill="x", pady=(0, 8))
        tk.Label(self.sidebar_content, text="扫描参数", font=self._fonts["heading"],
                 fg=PALETTE["navy"], bg="white").pack(anchor="w", pady=(0, 7))
        ttk.Label(self.sidebar_content, text="业务员", style="Card.TLabel").pack(anchor="w")
        self.salesperson_var = tk.StringVar(value="Xie")
        ttk.Entry(self.sidebar_content, textvariable=self.salesperson_var,
                  width=18, style="Main.TEntry").pack(fill="x", pady=(3, 9))
        ttk.Label(self.sidebar_content, text="忽略目录（逗号分隔）", style="Card.TLabel").pack(anchor="w")
        self.ignore_var = tk.StringVar(value="Muster, 模板, 运单, backup, backups, state, .venv")
        ttk.Entry(self.sidebar_content, textvariable=self.ignore_var,
                  style="Main.TEntry").pack(fill="x", pady=(3, 8))
        self.save_settings_btn = ttk.Button(self.sidebar_content, text="保存设置", command=self._save_config,
                   style="Blue.TButton")
        self.save_settings_btn.pack(fill="x", pady=(2, 8))
        ttk.Label(self.sidebar_content,
                  text="设置仅影响当前员工的来源与汇总表。\n本软件不会自行修改 PDF。",
                  style="Hint.TLabel", justify="left").pack(anchor="w", pady=(6, 0))

        # RIGHT: the four primary actions remain directly accessible.
        action = ttk.Frame(self.right_main, style="Main.TFrame")
        action.pack(fill="x", pady=(0, 4))
        self.scan_button = ttk.Button(action, text="开始扫描（只读）",
                       style="Blue.TButton", command=self._start_scan)
        self.scan_button.pack(side="left")
        self.stop_button = ttk.Button(action, text="停止", style="Soft.TButton",
                                      command=self._stop_scan, state="disabled")
        self.stop_button.pack(side="left", padx=(5, 0))
        self.review_button = ttk.Button(action, text="核对 / 预览 Excel",
                            style="Soft.TButton", command=self._preview_business, state="disabled")
        self.review_button.pack(side="left", padx=(7, 0))
        self.write_button = ttk.Button(action, text="确认写入 Excel",
                          style="Soft.TButton", command=self._confirm_write, state="disabled")
        self.write_button.pack(side="left", padx=(5, 0))
        self.other_menu = tk.Menu(self, tearoff=False)
        self.other_menu.add_command(label="导出 CSV 报告", command=self._export, state="disabled")
        self.other_menu.add_command(label="检查历史发货单归档关联（需确认）…",
                                    command=self._repair_historical_deliveries)
        self.other_menu.add_separator()
        self.scan_mode_menu = tk.Menu(self.other_menu, tearoff=False)
        for label, mode in (("快速增量扫描（默认）", "incremental"),
                            ("重新识别未归档 PDF", "full"),
                            ("完整归档核验（SHA-256）", "verify")):
            self.scan_mode_menu.add_radiobutton(label=label, variable=self.scan_mode_var, value=mode)
        self.other_menu.add_cascade(label="扫描模式", menu=self.scan_mode_menu)
        self.other_menu.add_separator()
        self.other_menu.add_command(label="定时检查设置", command=self._schedule_settings)
        self.other_menu.add_command(label="手动更新本周目录名", command=self._confirm_week_folder_rename)
        self.other_menu.add_separator()
        self.other_menu.add_command(label="月度发运统计 / AI补全…", command=self._open_monthly_workspace)
        self.other_menu.add_separator()
        self.font_menu = tk.Menu(self.other_menu, tearoff=False)
        for preset in ("紧凑", "标准", "舒适"):
            self.font_menu.add_radiobutton(label=f"{preset}字体", value=preset,
                variable=self.font_preset_var,
                command=lambda size=preset: self._apply_font_preset(size))
        self.other_menu.add_cascade(label="界面设置 · 字体大小", menu=self.font_menu)
        self.more_button = ttk.Menubutton(action, text="其他功能 ▾", style="Soft.TButton",
                                           direction="below", menu=self.other_menu)
        self.more_button.pack(side="left", padx=(6, 0))
        self.teams_button = ttk.Button(action, text="☁ 上传发票", style="Blue.TButton",
                                       command=self._open_teams_workspace, state="disabled")
        self.teams_button.pack(side="left", padx=(14, 0))
        # Inventory is a separate read-only workspace. Open it even before
        # any order PDF has been scanned (useful for pre-sale availability).
        self.inventory_button = ttk.Button(action, text="▦ 库存搜索", style="Soft.TButton",
                                            command=self._open_inventory_workspace)
        self.inventory_button.pack(side="left", padx=(8, 0))
        self.sidebar_toggle_btn = ttk.Button(action, text="◂ 设置", style="Soft.TButton",
                                               command=self._toggle_sidebar)
        self.sidebar_toggle_btn.pack(side="right", padx=(6, 0))

        self.status_var = tk.StringVar(value="就绪 · 扫描/核对只读 · 写入必须确认")
        ttk.Label(self.right_main, textvariable=self.status_var, style="Small.TLabel").pack(
                  fill="x", pady=(1, 3))
        self.progress = ttk.Progressbar(self.right_main, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 7))

        # Four main counters; optional detailed counters don't consume vertical
        # space while the user is reviewing the contract tree.
        self.stat_widgets = {}
        self.statbar_frame = ttk.Frame(self.right_main, style="Main.TFrame")
        self.statbar_frame.pack(fill="x", pady=(0, 4))
        for key, title in (("contracts", "合同数"), ("pending", "待开票"),
                           ("review", "待复核"), ("issues", "业务异常·核对后")):
            self._make_stat_card(self.statbar_frame, key, title, compact=False)
        self.stats_more_btn = ttk.Button(self.statbar_frame, text="详细统计 ▾",
                            command=self._toggle_stats_details, style="Soft.TButton")
        self.stats_more_btn.pack(side="right", padx=(0, 2))
        self.detail_stats_frame = ttk.Frame(self.right_main, style="Main.TFrame")
        for idx, (key, title) in enumerate((("pdf", "扫描 PDF"), ("orders", "订单识别"),
                             ("invoices", "发票识别"), ("delivery", "发货单"),
                             ("cancelled", "忽略作废"), ("duplicates", "重复文件"),
                             ("archived", "已归档"))):
            self._make_stat_card(self.detail_stats_frame, key, title, compact=True, grid_index=idx)

        results_top = ttk.Frame(self.right_main, style="Main.TFrame")
        results_top.pack(fill="x", pady=(5, 5))
        row_title = ttk.Frame(results_top, style="Main.TFrame")
        row_title.pack(fill="x")
        tk.Label(row_title, text="识别结果 · 合同树", font=self._fonts["title"],
                 fg=PALETTE["navy"], bg=PALETTE["bg"]).pack(side="left")
        self.view_var = tk.StringVar(value="识别结果")
        ttk.Combobox(row_title, textvariable=self.view_var, state="readonly",
                   values=("识别结果", "归档管理"), width=11).pack(side="right", padx=(8, 0))
        self.view_var.trace_add("write", lambda *_: self._switch_view())
        tk.Label(row_title, text="订单组可折叠 · 点击表头排序", font=self._fonts["small"],
                 fg=PALETTE["muted"], bg=PALETTE["bg"]).pack(side="left", padx=(12, 0))
        row_filter = ttk.Frame(results_top, style="Main.TFrame")
        row_filter.pack(fill="x", pady=(5, 0))
        tk.Label(row_filter, text="搜索订单/发货/发票号/客户", bg=PALETTE["bg"],
                 fg=PALETTE["muted"], font=self._fonts["small"]).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._refresh_table())
        ttk.Entry(row_filter, textvariable=self.search_var, width=27).pack(side="left", fill="x", expand=True, padx=(6, 12))
        tk.Label(row_filter, text="状态", bg=PALETTE["bg"], fg=PALETTE["muted"],
                 font=self._fonts["small"]).pack(side="left")
        self.filter_var = tk.StringVar(value="全部")
        self.active_filters = ("全部", "待开票", "已开票", "部分开票", "金额待核",
                               "发货表已登记", "待复核", "重复文件", "忽略作废",
                               "订单", "发票", "发货单", "贷项通知单", "关联订单已归档", "未匹配")
        self.archive_filters = ("全部", "已归档", "订单", "发票", "发货单", "贷项通知单",
                                "关联合同归档", "补充归档发货单", "待开票归档",
                                "作废归档", "已复核归档", "未匹配")
        self.status_filter_combo = ttk.Combobox(row_filter, textvariable=self.filter_var,
                            values=self.active_filters, width=13, state="readonly")
        self.status_filter_combo.pack(side="left", padx=(6, 0))
        self.status_filter_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_table())

        bulk = ttk.Frame(self.right_main, style="Main.TFrame")
        bulk.pack(fill="x", pady=(0, 6))
        self.selected_count_var = tk.StringVar(value="已选择 0 项")
        ttk.Label(bulk, textvariable=self.selected_count_var, style="Small.TLabel").pack(side="left", padx=(0, 12))
        ttk.Button(bulk, text="全选当前筛选结果", command=self._select_all_visible).pack(side="left", padx=(0, 6))
        ttk.Button(bulk, text="取消勾选", command=self._clear_checked).pack(side="left", padx=(0, 6))
        self.bulk_archive_button = ttk.Button(bulk, text="批量归档（需确认）", command=self._archive_checked)
        self.bulk_archive_button.pack(side="left", padx=(0, 6))
        self.bulk_restore_button = ttk.Button(bulk, text="批量恢复（需确认）", command=self._restore_checked)
        self.bulk_restore_button.pack(side="left", padx=(0, 6))
        self.undo_archive_button = ttk.Button(bulk, text="撤销上次归档", command=self._undo_last_archive)
        self.undo_archive_button.pack(side="left", padx=(0, 6))

        table_wrap = ttk.Frame(self.right_main, style="Card.TFrame")
        table_wrap.pack(fill="both", expand=True)
        columns = ("choose", "kind", "customer", "identifier", "day", "gross", "net", "status", "remark")
        self.table = ttk.Treeview(table_wrap, columns=columns, show="tree headings", selectmode="browse", style="Data.Treeview")
        self.table.heading("#0", text="展开全部 ▾", command=self._toggle_all_contracts)
        self.table.column("#0", width=100, minwidth=95, stretch=False, anchor="center")
        cols = {
            "choose": ("☐", 44), "kind": ("类型", 80), "customer": ("客户", 145),
            "identifier": ("订单/发货/发票号", 165), "day": ("PDF 日期", 108), "gross": ("总金额", 100),
            "net": ("Netto", 100), "status": ("综合状态", 185), "remark": ("备注", 285),
        }
        self._column_titles = {col: title for col, (title, _) in cols.items()}
        for col in columns:
            title, width = cols[col]
            if col == 'kind':
                self.table.heading(col, text=title)
            elif col == 'choose':
                self.table.heading(col, text=title, command=self._select_all_visible)
            else:
                self.table.heading(col, text=title, command=lambda c=col: self._sort_by(c))
            self.table.column(col, width=width, minwidth=36 if col == "choose" else 65,
                              stretch=(col in {"customer", "remark"}),
                              anchor="w" if col in {"customer", "remark"} else "center")
        # Group headers get two alternating shades *per business state* so
        # adjacent contracts are visibly separated even when collapsed.
        # Existing Chinese status text stays unchanged for accessibility and compatibility.
        group_colors = {
            "done": (("#E3F5EC", "#D1EEE0"), "#176143"),
            "pending": (("#FFF3D7", "#FFE8C0"), "#8E560C"),
            "alert": (("#FFE8E5", "#FBD9D4"), "#9B3333"),
            "neutral": (("#E9F1FB", "#D7E6F5"), "#204B77"),
        }
        for category, (shades, ink) in group_colors.items():
            for stripe, bg in enumerate(shades):
                self.table.tag_configure(f"contract_{category}_{stripe}", background=bg,
                                         foreground=ink, font=self._fonts["table_bold"])
        # Retained as a backward-compatible tag; new contract rows use the
        # state-aware styles above rather than a single undifferentiated blue.
        self.table.tag_configure("group", background="#EAF2FA", foreground="#15375F", font=self._fonts["table_bold"])
        self.table.tag_configure("orphan", background="#FFF0D9", foreground="#965400", font=self._fonts["table_bold"])
        self.table.tag_configure("review", background="#FFF4E4", foreground="#8A4B16")
        self.table.tag_configure("pending", background="#FFF4DC", foreground="#965400")
        self.table.tag_configure("cancel", background="#EFF3F8", foreground="#68778B")
        for stripe, shade in enumerate(("#E9F1FB", "#D7E6F5")):
            self.table.tag_configure(f"archive_group_{stripe}", background=shade,
                                     foreground="#204B77", font=self._fonts["table_bold"])
        self.table.tag_configure("archive_orphan", background="#FFF0D9", foreground="#965400",
                                 font=self._fonts["table_bold"])
        self.table.tag_configure("even", background=PALETTE["striped"])
        vert = ttk.Scrollbar(table_wrap, orient="vertical", command=self.table.yview)
        horiz = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=vert.set, xscrollcommand=horiz.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        vert.grid(row=0, column=1, sticky="ns")
        horiz.grid(row=1, column=0, sticky="ew")
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)
        self.table.bind("<<TreeviewSelect>>", self._show_detail)
        # A click on the displayed number or Excel note copies it. Intercept it
        # before the existing checkbox/parent expand handlers so copying a
        # contract number does not unexpectedly collapse its children.
        self.table.bind("<Button-1>", self._copy_clicked_cell, add="+")
        self.table.bind("<Button-1>", self._toggle_checkbox, add="+")
        self.table.bind("<Button-1>", self._click_contract_row, add="+")
        self.table.bind("<Double-1>", self._double_click_pdf)
        self.table.bind("<<TreeviewOpen>>", self._track_tree_open, add="+")
        self.table.bind("<<TreeviewClose>>", self._track_tree_close, add="+")
        # Direct right-click menu: no nested "文件操作" submenu.
        # Scope actions to the clicked row, even after sorting/filtering.
        self.result_menu = tk.Menu(self.table, tearoff=False)
        self.result_menu.add_command(label="打开 PDF", command=self._open_selected)
        self.result_menu.add_command(label="在文件夹中显示", command=self._reveal_selected)
        self.result_menu.add_command(label="查询此订单配件库存与库位…", command=self._inventory_selected_order)
        self.result_menu.add_command(label="打印此订单拣货单及发货单…", command=self._print_selected_order)
        self.result_menu.add_separator()
        self.result_menu.add_command(label="标记为已复核", command=self._mark_reviewed_selected)
        self.result_menu.add_command(label="添加/编辑 Excel 备注…", command=self._edit_note_selected)
        self.result_menu.add_command(label="归档此 PDF（需确认）", command=self._archive_selected)
        self.result_menu.add_command(label="恢复此 PDF（需确认）", command=self._restore_selected)
        self._restore_menu_index = self.result_menu.index('end')
        self.result_menu.add_separator()
        self.result_menu.add_command(label="重新定位原始 PDF…", command=self._relink_archive_selected)
        self.result_menu.add_command(label="单独清除失效归档记录…", command=self._clear_invalid_archive_selected)
        self.table.bind("<Button-3>", self._show_result_context_menu)

        self.detail_toggle = ttk.Button(self.right_main, text="▸ 选中记录说明", style="Soft.TButton",
                                        command=self._toggle_detail_panel)
        self.detail_toggle.pack(anchor="w", pady=(6, 0))
        self.detail_wrap = ttk.Frame(self.right_main, style="Card.TFrame", padding=(11, 7))
        self.detail_wrap.pack(fill="x", pady=(10, 0))
        ttk.Label(self.detail_wrap, text="选中记录 / 识别说明", style="Card.TLabel").pack(anchor="w")
        self.detail = tk.Text(self.detail_wrap, height=3, wrap="word", font=self._fonts["mono"], bg="white", fg=PALETTE["ink"], relief="flat")
        self.detail.pack(fill="x", pady=(3, 0))
        self.detail.insert("1.0", "请选择来源目录，然后点击「开始扫描」。所有数据只在本地读取。")
        self.detail.configure(state="disabled")
        self.detail_wrap.pack_forget()
        foot = tk.Label(self.right_main, text="定时只读检查 · 写入 Excel 和改名均需确认 · 保留人工操作", bg=PALETTE["bg"], fg=PALETTE["muted"], font=self._fonts["small"])
        foot.pack(anchor="w", pady=(4, 0))

    def _make_stat_card(self, parent, key: str, title: str, compact: bool, grid_index: int | None = None):
        card = tk.Frame(parent, bg="white", highlightbackground=PALETTE["line"], highlightthickness=1)
        if grid_index is None:
            card.pack(side="left", fill="x", expand=True, padx=(0, 6))
        else:
            card.grid(row=grid_index // 4, column=grid_index % 4, sticky="ew", padx=(0, 6), pady=(0, 6))
            parent.grid_columnconfigure(grid_index % 4, weight=1)
        number = tk.Label(card, text="—" if key == "issues" else "0",
                 font=self._fonts["stat" if not compact else "section"],
                 fg=PALETTE["red"] if key in {"review", "issues"} else PALETTE["navy"], bg="white")
        number.pack(anchor="w", padx=10, pady=(5 if compact else 7, 0))
        tk.Label(card, text=title, font=self._fonts["small"],
                 bg="white", fg=PALETTE["muted"]).pack(anchor="w", padx=10, pady=(0, 5 if compact else 7))
        self.stat_widgets[key] = number

    def _on_sidebar_content_resize(self, _event=None):
        self.sidebar_canvas.configure(scrollregion=self.sidebar_canvas.bbox("all"))

    def _on_sidebar_canvas_resize(self, event):
        self.sidebar_canvas.itemconfigure(self._sidebar_canvas_item, width=event.width)

    def _toggle_detail_panel(self):
        self._detail_visible = not self._detail_visible
        if self._detail_visible:
            self.detail_wrap.pack(fill="x", after=self.detail_toggle, pady=(5, 0))
        else:
            self.detail_wrap.pack_forget()
        self.detail_toggle.configure(text="▾ 收起记录说明" if self._detail_visible else "▸ 选中记录说明")

    def _toggle_stats_details(self):
        self._stats_details_visible = not self._stats_details_visible
        if self._stats_details_visible:
            self.detail_stats_frame.pack(fill="x", after=self.statbar_frame, pady=(0, 5))
        else:
            self.detail_stats_frame.pack_forget()
        self.stats_more_btn.configure(text="收起统计 ▴" if self._stats_details_visible else "详细统计 ▾")
        self._save_layout_preferences()

    def _toggle_sidebar(self):
        if self._sidebar_visible:
            try:
                self._sidebar_width = max(260, min(475, self.workspace.sashpos(0)))
            except tk.TclError:
                pass
            self.workspace.forget(self.sidebar_frame)
            self._sidebar_visible = False
            self.sidebar_toggle_btn.configure(text="▸ 设置")
        else:
            self.workspace.insert(0, self.sidebar_frame, weight=0)
            self._sidebar_visible = True
            self.sidebar_toggle_btn.configure(text="◂ 设置")
            self.after_idle(self._restore_sidebar_sash)
        self._save_layout_preferences()

    def _restore_sidebar_sash(self):
        if not self._sidebar_visible:
            return
        self.workspace.update_idletasks()
        if self.workspace.winfo_width() < 450:
            self.after(100, self._restore_sidebar_sash)
            return
        usable_width = self.workspace.winfo_width()
        desired = max(260, min(int(self._sidebar_width), 475, usable_width - 550))
        try:
            self.workspace.sashpos(0, desired)
        except tk.TclError:
            pass

    def _on_workspace_drag_end(self, _event=None):
        if not self._sidebar_visible:
            return
        try:
            pos = self.workspace.sashpos(0)
        except tk.TclError:
            return
        if pos > 200:
            self._sidebar_width = max(260, min(475, pos))
            self._save_layout_preferences()

    def _save_layout_preferences(self):
        """Visual-only settings: never invalidate or change Excel approval plans."""
        try:
            path = GUI_CONFIG.parent / "ui_layout.json"
            path.parent.mkdir(parents=True, exist_ok=True)
            contents = {"sidebar_open": self._sidebar_visible,
                        "sidebar_width": int(self._sidebar_width),
                        "detailed_stats_open": self._stats_details_visible}
            path.write_text(json.dumps(contents, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            # Layout preferences are optional. A failure must not stop a scan.
            pass

    def _apply_saved_layout(self):
        path = GUI_CONFIG.parent / "ui_layout.json"
        try:
            opts = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
            self._sidebar_width = max(260, min(475, int(opts.get("sidebar_width", 320))))
            if bool(opts.get("detailed_stats_open", False)) != self._stats_details_visible:
                self._toggle_stats_details()
            if not bool(opts.get("sidebar_open", True)):
                self._toggle_sidebar()
            else:
                self._restore_sidebar_sash()
        except (ValueError, OSError, TypeError):
            self._restore_sidebar_sash()

    def _open_monthly_workspace(self):
        """Generate an internal shipment report; no original file is modified."""
        from gui_monthly_ui import MonthlyReportWindow
        win = getattr(self, '_monthly_window', None)
        if win is not None and win.winfo_exists():
            win.lift()
            win.focus_set()
            return
        self._monthly_window = MonthlyReportWindow(self)

    def _open_inventory_workspace(self, order_pdf: str | None = None):
        """Read-only stock lookup; never needs a completed sales scan."""
        from gui_inventory import InventoryWindow
        win = getattr(self, '_inventory_window', None)
        if win is not None and win.winfo_exists():
            win.lift()
            win.focus_set()
            if order_pdf:
                win.open_order(order_pdf)
            return
        self._inventory_window = InventoryWindow(self, order_pdf=order_pdf)

    def _inventory_selected_order(self):
        """Use only the selected sales-order PDF for stock demand."""
        selected = self.table.selection()
        if not selected:
            return
        row = self._shown_rows.get(selected[0])
        if not row or row.kind != '订单':
            messagebox.showwarning('请选择订单', '请选中一条订单行，然后查询其配件库存与库位。')
            return
        if not Path(row.source).is_file():
            messagebox.showerror('源文件不存在', '订单 PDF 已移动或被删除，请重新定位后查询。')
            return
        self._open_inventory_workspace(row.source)

    def _print_selected_order(self):
        """Start the picking/printing workflow from a scanned order row."""
        selected=self.table.selection()
        if not selected:
            return
        row=self._shown_rows.get(selected[0])
        if row is None or row.kind!='订单' or not Path(row.source).is_file():
            messagebox.showwarning('请选择有效订单','请选中仍可读取的订单 PDF，再打开拣货单打印。')
            return
        self._open_inventory_workspace()
        win=getattr(self,'_inventory_window',None)
        if win is not None and win.winfo_exists():
            win.request_order_print(row.source)

    def _teams_scan_ready(self) -> bool:
        """Only a complete, current source scan can populate Teams uploads.

        Excel-only options (date renaming, font) do not invalidate scanned PDFs;
        source lists, source hints, employee identity and Excel/profile do.
        """
        if not self.scanned_complete or self._scan_running or not self.sources:
            return False
        if any(row.kind == '目录异常' for row in self.rows):
            return False
        previous = getattr(self, '_last_config', None)
        if not previous:
            return False
        current = self._get_config()
        for key in ('sources', 'salesperson', 'exclude_names', 'target_excel'):
            if current[key] != previous.get(key):
                return False
        return True

    def _refresh_teams_access(self):
        if hasattr(self, 'teams_button'):
            self.teams_button.configure(state='normal' if self._teams_scan_ready() else 'disabled')

    def _close_teams_for_rescan(self) -> bool:
        """Prevent an already-open Teams list from using pre-rescan PDF rows."""
        win = getattr(self, '_teams_window', None)
        if win is None or not win.winfo_exists():
            return True
        if getattr(win, 'busy', False):
            messagebox.showwarning('Teams 操作尚未完成', '请先完成或关闭 Teams 上传工作区中的操作，再重新扫描。', parent=self)
            return False
        if not messagebox.askyesno('重新扫描 PDF',
                '重新扫描将关闭当前 Teams 发票上传窗口，并清除尚未执行的勾选和拖拽准备。\n'
                '这不会删除任何 PDF 或已上传记录。是否继续？', parent=self):
            return False
        win._close()
        return not win.winfo_exists()

    def _open_teams_workspace(self):
        """Second, isolated workspace; never touches the Excel write plan."""
        if not self._teams_scan_ready():
            messagebox.showwarning('请先完成扫描',
                '请先点击「开始扫描（只读）」并等待完整扫描成功后，再打开上传发票。\n'
                '若更改了 PDF 来源、业务员或 Excel，需重新扫描。', parent=self)
            self._refresh_teams_access()
            return
        win = getattr(self, '_teams_window', None)
        if win is not None and win.winfo_exists():
            win.lift(); win.focus_force(); return
        from gui_teams import TeamsUploadWindow
        self._teams_window = TeamsUploadWindow(self)

    def _add_source(self):
        self._invalidate_plan()
        path = filedialog.askdirectory(title="选择需要递归扫描的 PDF 来源文件夹")
        if not path: return
        path = str(Path(path).resolve())
        if any(os.path.normcase(x["path"]) == os.path.normcase(path) for x in self.sources):
            messagebox.showinfo("已有目录", "该目录已在列表中。")
            return
        self.sources.append({"path": path, "hint": "自动识别"})
        self._refresh_sources()
        self._save_config(silent=True)

    def _refresh_sources(self):
        self._refresh_teams_access()
        self.source_tree.delete(*self.source_tree.get_children())
        for i, src in enumerate(self.sources):
            self.source_tree.insert("", "end", iid=str(i), values=(src["path"], src.get("hint", "自动识别")))

    def _select_source(self, _event=None):
        sel = self.source_tree.selection()
        if sel:
            self.hint_var.set(self.sources[int(sel[0])].get("hint", "自动识别"))

    def _update_source_hint(self, _event=None):
        self._invalidate_plan()
        sel = self.source_tree.selection()
        if sel:
            self.sources[int(sel[0])]["hint"] = self.hint_var.get()
            self._refresh_sources()
            self.source_tree.selection_set(sel[0])
            self._save_config(silent=True)

    def _remove_source(self):
        self._invalidate_plan()
        sel = self.source_tree.selection()
        if not sel: return
        self.sources.pop(int(sel[0]))
        self._refresh_sources()
        self._save_config(silent=True)

    def _choose_excel(self):
        self._invalidate_plan()
        path = filedialog.askopenfilename(title="选择已有汇总 Excel；新文件可在下方路径手动输入", filetypes=[("Excel 工作簿", "*.xlsx"), ("全部文件", "*.*")])
        if path:
            self.excel_var.set(path)
            self._save_config(silent=True)

    def _choose_new_excel(self):
        self._invalidate_plan()
        name = self.salesperson_var.get().strip() or 'Sales'
        suggested = f'业务周汇总表_{name}_更新至{datetime.now():%d.%m.%Y}.xlsx'
        path = filedialog.asksaveasfilename(title='指定新的个人 Excel 文件保存位置（此时不会创建）',
                 defaultextension='.xlsx', initialfile=suggested, filetypes=[('Excel 工作簿', '*.xlsx')])
        if path:
            self.excel_var.set(path)
            self._save_config(silent=True)

    def _get_config(self):
        return {"version": GUI_VERSION, "archive_profile": self.archive_profile or "", "sources": self.sources,
                "target_excel": self.excel_var.get().strip(),
                "salesperson": self.salesperson_var.get().strip() or "Xie",
                "exclude_names": [s.strip() for s in self.ignore_var.get().replace("，", ",").split(",") if s.strip()],
                "rename_excel_date": bool(self.rename_date_var.get()),
                "font_preset": self.font_preset_var.get()}

    def _save_config(self, silent=False):
        self._invalidate_plan()
        self._refresh_teams_access()
        try:
            GUI_CONFIG.parent.mkdir(parents=True, exist_ok=True)
            GUI_CONFIG.write_text(json.dumps(self._get_config(), ensure_ascii=False, indent=2), encoding="utf-8")
            if not silent: messagebox.showinfo("保存完成", f"配置已保存到：\n{GUI_CONFIG}\n\n不会影响原来的 config.json。")
        except Exception as exc:
            if not silent: messagebox.showerror("保存失败", str(exc))

    def _load_config(self):
        if not GUI_CONFIG.exists():
            # Assist migration without assuming files always follow the old layout.
            old = BASE / "config.json"
            if old.exists():
                try:
                    oldcfg = json.loads(old.read_text(encoding="utf-8-sig"))
                    root = oldcfg.get("root_path", "")
                    if root and Path(root).is_dir():
                        self.sources = [{"path": root, "hint": "自动识别"}]
                        self.salesperson_var.set(oldcfg.get("salesperson", "Xie"))
                except Exception:
                    pass
        else:
            try:
                cfg = json.loads(GUI_CONFIG.read_text(encoding="utf-8-sig"))
                self.sources = [s for s in cfg.get("sources", []) if s.get("path")]
                self.archive_profile = cfg.get("archive_profile") or None
                self.salesperson_var.set(cfg.get("salesperson", "Xie"))
                self.ignore_var.set(", ".join(cfg.get("exclude_names", sorted(DEFAULT_EXCLUDED))))
                self.excel_var.set(cfg.get("target_excel", ""))
                self.rename_date_var.set(bool(cfg.get('rename_excel_date', False)))
                self._apply_font_preset(cfg.get("font_preset", "标准"), persist=False)
            except Exception as exc:
                messagebox.showwarning("配置读取失败", f"使用默认配置。\n{exc}")
        if not self.archive_profile:
            import uuid
            self.archive_profile = str(uuid.uuid4())
        self._refresh_sources()
        # Store a persistent employee profile ID, independent of dated Excel names.
        try:
            old_cfg = json.loads(GUI_CONFIG.read_text(encoding='utf-8-sig')) if GUI_CONFIG.exists() else {}
        except (ValueError, OSError):
            old_cfg = {}
        if not old_cfg.get('archive_profile'):
            GUI_CONFIG.parent.mkdir(parents=True, exist_ok=True)
            GUI_CONFIG.write_text(json.dumps(self._get_config(), ensure_ascii=False, indent=2), encoding='utf-8')

    def _employee_profile_key(self):
        """Separate archive DB records for distinct employee/Excel combinations.

        Date-based Excel filename renaming keeps the same profile. With no
        target configured, fall back to a persistent local profile UUID.
        """
        salesperson = self.salesperson_var.get().strip().casefold()
        target = self.excel_var.get().strip()
        if not target:
            return 'setup-' + self.archive_profile
        normalized = path_key(target)
        normalized = re.sub(r'更新至\d{1,2}\.\d{1,2}\.\d{4}(?=\.xlsx$)',
                            '更新至DATE', normalized, flags=re.I)
        identifier = json.dumps([salesperson,normalized], ensure_ascii=False)
        return 'employee-' + hashlib.sha256(identifier.encode('utf-8')).hexdigest()[:32]

    def _sync_employee_archive(self):
        key = self._employee_profile_key()
        if key != self._current_archive_profile:
            self._current_archive_profile = key
            self.archive_store = ArchiveStore(GUI_CONFIG.parent / 'archive_state.sqlite3', key)
            self.note_store = NoteStore(GUI_CONFIG.parent / 'remark_drafts.sqlite3', key)
            self._load_cached_archive_numbers()
            self.archived_map = {}
            self.checked_paths.clear()
            self._last_archive_paths = []

    def _load_cached_archive_numbers(self):
        """Keep historic Delivery No. visible even before the next fast scan.

        The archived database predates Delivery No. storage.  Reuse SHA-bound
        extraction data in the new cache, without reopening archived PDFs.
        Old records without cache will obtain this number after their first
        v1.5.34 scan, never from their Ordernumber fallback.
        """
        self._cached_delivery_numbers = {}
        cache_path = GUI_CONFIG.parent / 'scan_cache.sqlite3'
        if not cache_path.exists() or self.archive_store is None:
            return
        try:
            archived = {item.path:item for item in self.archive_store.all_archived()
                        if item.kind == '发货单'}
            if not archived:
                return
            cache = ScanCache(cache_path, self._current_archive_profile)
            for key, entry in cache.load().items():
                item = archived.get(key)
                if item and entry[0] == CACHE_SCHEMA_VERSION and item.sha == entry[5]:
                    raw = json.loads(entry[6])
                    if raw.get('kind') == '发货单' and raw.get('delivery_no'):
                        self._cached_delivery_numbers[key] = str(raw['delivery_no'])
        except (sqlite3.DatabaseError, OSError, ValueError, KeyError, IndexError):
            # This is a display-only optimisation. Archive operations always
            # use the independent authoritative archive database.
            self._cached_delivery_numbers = {}

    def _start_scan(self):
        if self.worker and self.worker.is_alive(): return
        if self.business_busy: return
        if not self.sources:
            messagebox.showwarning("未选择目录", "请先使用「添加目录」选择至少一个 PDF 来源。")
            return
        if not self._close_teams_for_rescan():
            return
        self._save_config(silent=True)
        self._sync_employee_archive()
        self._invalidate_plan()
        self.review_button.configure(state="disabled")
        self.scanned_complete = False
        self._scan_running = True
        self._refresh_teams_access()
        self.rows = []
        self.archived_map = {}
        self.checked_paths.clear()
        self.opened_for_review.clear()
        self.scan_hashes.clear()
        self._scan_file_metadata = {}
        self._scan_report = {}
        self._shown_rows.clear()
        self._tree_open_contracts.clear()
        self._refresh_table()
        self.cancel_event.clear()
        self.scan_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.other_menu.entryconfigure("导出 CSV 报告", state="disabled")
        self.progress.start(15)
        scan_mode = self.scan_mode_var.get()
        self.status_var.set("扫描中…" + {'incremental':'正在增量检查本地 PDF',
                           'full':'正在重新识别未归档 PDF',
                           'verify':'正在完整核验归档 PDF'}[scan_mode])
        cfg = self._get_config()
        self._last_config = json.loads(json.dumps(cfg))
        archive_snapshot = self.archive_store.all_archived()
        cache_db = GUI_CONFIG.parent / 'scan_cache.sqlite3'
        profile = self._current_archive_profile
        def work():
            try:
                cache = ScanCache(cache_db, profile)
                _rows, report, metadata = incremental_scan(
                    cfg['sources'], cfg['salesperson'], cfg['exclude_names'],
                    self.cancel_event, lambda r: self.events.put(('row', r)),
                    cache, archive_snapshot, scan_mode)
                self.events.put(('done', (self.cancel_event.is_set(), report, metadata)))
            except Exception:
                self.events.put(('error', traceback.format_exc()))
        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _stop_scan(self):
        self.cancel_event.set()
        self.status_var.set("正在停止：处理完当前 PDF 后结束")
        self.stop_button.configure(state="disabled")

    def _poll_events(self):
        dirty = False
        done = None
        processed = 0
        while processed < 180:
            try: event, payload = self.events.get_nowait()
            except queue.Empty: break
            processed += 1
            if event == "row":
                self.rows.append(payload)
                dirty = True
            elif event == "done":
                self._scan_running = False
                cancelled, self._scan_report, self._scan_file_metadata = payload
                self.scanned_complete = not cancelled
                done = "扫描已停止（结果可能不完整）" if cancelled else "只读扫描完成"
            elif event == 'business_plan':
                self.business_busy = False
                if self._preview_cfg is None or self._preview_cfg != self._get_config():
                    self._invalidate_plan()
                    self.review_button.configure(state='normal' if self.scanned_complete else 'disabled')
                    messagebox.showwarning('预览失效', '处理时用户设置发生变化，请重新扫描并核对。')
                    continue
                self.business_plan = payload
                mark_pending_invoice_status(self.rows, payload.invoiced_contracts)
                self._refresh_stats()
                self._refresh_table()
                self.review_button.configure(state='normal')
                self.write_button.configure(state='normal' if payload.has_changes else 'disabled')
                self.status_var.set(f'业务预览完成 · 待新增 {payload.added} 条 · 待写备注 {len(payload.note_changes)} 条 · 异常 {len(payload.warnings)} 条')
                try:
                    self._last_report_path = gui_business.save_preview_report(payload)
                except Exception as exc:
                    self.status_var.set(f'预览完成，但报告保存失败：{exc}')
                self._show_business_plan(payload)
                if self.scheduled_run and payload.warnings:
                    messagebox.showwarning('定时检查异常', f'发现 {len(payload.warnings)} 条需复核问题。请查看业务预览。')
            elif event == 'business_written':
                # This event is emitted ONLY after the Excel write/atomic save
                # succeeds. Never clear drafts on preview, cancellation or errors.
                self.business_busy = False
                self.business_plan = None
                self.write_button.configure(state='disabled')
                self.review_button.configure(state='normal')
                saved, backup, log_error, applied_notes, note_profile = payload
                note_cleanup_warning = ''
                try:
                    # Do not wipe a note that the user edited during the write.
                    # Excel and the DB are different resources: clean the DB only
                    # after receiving the successful write event.
                    removed = NoteStore(GUI_CONFIG.parent / 'remark_drafts.sqlite3', note_profile).discard_if_unchanged(applied_notes)
                    if removed != len(applied_notes):
                        note_cleanup_warning = (f'原拟清理 {len(applied_notes)} 条备注草稿，实际清理 {removed} 条。'
                                                '其余可能已被修改，请重新核对，不会自动删除。')
                except Exception as exc:
                    note_cleanup_warning = f'备注草稿清理失败：{exc}；Excel 已写入，但请重新核对剩余草稿。'
                self.excel_var.set(str(saved))
                self._save_config(silent=True)
                # _refresh_table() reads the current NoteStore contents afresh.
                # Without this, parent rows continue displaying "待写备注1"
                # until the next manual scan even though the draft was deleted.
                self._refresh_stats()
                self._refresh_table()
                remaining_notes = len(self.note_store.list_drafts())
                self.status_var.set(f'Excel 已写入并备份 · 待写备注 {remaining_notes} 条 · 请重新扫描核对')
                messagebox.showinfo('写入成功', f'更新文件：{saved}\n备份：{backup or "首次建表，无旧文件"}\n'
                                    + (f'注意：日志保存失败：{log_error}\n' if log_error else '')
                                    + (f'备注提醒：{note_cleanup_warning}\n' if note_cleanup_warning else '')
                                    + '\n备注状态已刷新；请重新扫描核对业务状态。')
                self.scanned_complete = False
                self._refresh_teams_access()
                self.review_button.configure(state='disabled')
            elif event == 'business_error':
                self.business_busy = False
                self._invalidate_plan()
                self.review_button.configure(state='normal' if self.scanned_complete else 'disabled')
                self.status_var.set('业务检查/写入失败，未确认完成，请查看错误')
                messagebox.showerror('业务检查 / 写入失败', payload[-3600:])
            elif event == "error":
                self._scan_running = False
                self.scanned_complete = False
                done = "扫描出现程序错误"
                messagebox.showerror("扫描失败", payload[-3500:])
        if dirty:
            self._refresh_stats()
            self._refresh_table()
            self.status_var.set(f"扫描中 · 已读取 {len(self.rows)} 个文件/目录错误")
        if done:
            flag_duplicate_content(self.rows)
            self.archived_map = self.archive_store.archived_for(self.rows)
            self._load_cached_archive_numbers()
            # Do not re-hash every unchanged file a second time.  Comparing
            # metadata guards against files changing during the scan; final
            # archive/write actions still re-hash source PDF bytes.
            self.scan_hashes = {path_key(r.source): r.sha for r in self.rows
                                if r.sha and r.kind != '目录异常'}
            changed = changed_since_scan(self._scan_file_metadata)
            if changed:
                self.scanned_complete = False
                done = '扫描期间 PDF 被修改，请重新扫描'
                messagebox.showwarning('文件扫描期间发生变化', '请重新扫描后再进行归档或 Excel 预览：\n' + '\n'.join(changed[:10]))
            if self.scanned_complete:
                mark_pending_invoice_status(self.rows)  # PDF-only initial view
            self._refresh_stats()
            self._refresh_table()
            self.progress.stop()
            self.progress.configure(value=0)
            self.scan_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            if self._scan_report.get('cache_unavailable',0):
                messagebox.showwarning('缓存不可用',
                    '本次已尽可能正常扫描 PDF，但本地增量缓存数据库读取或写入失败。\n'
                    '下次启动可能需要重新解析。请检查本地文件权限或 SQLite 数据库占用。', parent=self)
            if self._scan_report.get('archived_changed',0):
                messagebox.showwarning('归档文件发生变化',
                    f'发现 {self._scan_report["archived_changed"]} 份已归档路径的文件内容或状态需核验。\n'
                    '原归档记录仍然保留；新内容不会自动继承旧归档状态。\n'
                    '请在「归档管理」核对相应历史记录。', parent=self)
            self.other_menu.entryconfigure("导出 CSV 报告", state="normal" if self.rows else "disabled")
            info = self._scan_report
            suffix = (f' · 缓存 {info.get("cached",0)} · 已归档跳过解析 {info.get("archived_cached",0)}'
                      f' · 实际解析 {info.get("reparsed",0)}'
                      f' · 归档校验 {info.get("archive_verified",0)}'
                      f' · 归档疑似变化 {info.get("archived_changed",0)}')
            self.status_var.set(f"{done} · 记录 {len(self.rows)} 条{suffix} · Excel 未修改")
            self.review_button.configure(state="normal" if self.scanned_complete and self.rows else "disabled")
            self._refresh_teams_access()
            if self.scanned_complete and not self._teams_scan_ready():
                self.status_var.set(f"扫描不完整或来源设置已变化 · 记录 {len(self.rows)} 条 · 请排除异常后重新扫描")
            if self.scheduled_preview and self.scanned_complete:
                self.scheduled_preview = False
                self.after(150, self._preview_business)
        self.after(120, self._poll_events)

    def _refresh_stats(self):
        active = [r for r in self.rows if path_key(r.source) not in self.archived_map]
        nums = summarize(active)
        groups, _ = make_tree(active)
        nums['contracts'] = len(groups)
        nums['archived'] = len(self.archived_map)
        for key, num in nums.items():
            if key in self.stat_widgets:
                self.stat_widgets[key].configure(text=str(num))
        if "issues" in self.stat_widgets:
            self.stat_widgets["issues"].configure(
                text=str(len(self.business_plan.warnings)) if self.business_plan is not None else "—")

    def _sort_by(self, column: str):
        # "类型" and checkboxes are structure/actions, never ordering keys.
        if column in {'kind', 'choose'}:
            return
        if self._sort_column == column:
            self._sort_descending = not self._sort_descending
        else:
            self._sort_column = column
            self._sort_descending = True
        for name, title in self._column_titles.items():
            if name in {'choose', 'kind'}: continue
            arrow = (' ▼' if self._sort_descending else ' ▲') if name == self._sort_column else ''
            self.table.heading(name, text=title + arrow)
        self._refresh_table()
        self.table.yview_moveto(0)

    def _update_tree_selection_label(self):
        self.selected_count_var.set(f'已选择 {len(self.checked_paths)} 份 PDF')

    def _toggle_all_contracts(self):
        """Toggle all *visible* contract groups; preserve state across sorts.

        Search can force matching contracts open, so explicit bulk collapse must
        override search expansion until the user opens a group again.
        """
        groups = [iid for iid in self.table.get_children()
                  if iid in self._tree_parent_items or iid in self._tree_orphan_items]
        if not groups:
            return
        should_open = not all(bool(self.table.item(iid, "open")) for iid in groups)
        for iid in groups:
            self.table.item(iid, open=should_open)
            if should_open:
                self._tree_open_contracts.add(iid)
                self._tree_forced_closed.discard(iid)
            else:
                self._tree_open_contracts.discard(iid)
                self._tree_forced_closed.add(iid)
        self._update_expand_heading()

    def _update_expand_heading(self):
        groups = [iid for iid in self.table.get_children()
                  if iid in self._tree_parent_items or iid in self._tree_orphan_items]
        is_all_open = bool(groups) and all(bool(self.table.item(iid, "open")) for iid in groups)
        self.table.heading("#0", text="折叠全部 ▴" if is_all_open else "展开全部 ▾")

    def _remember_tree_openness(self):
        for iid in self._tree_parent_items:
            if self.table.exists(iid):
                if self.table.item(iid, 'open'):
                    self._tree_open_contracts.add(iid)
                else:
                    self._tree_open_contracts.discard(iid)

    def _track_tree_open(self, _event=None):
        iid = self.table.focus()
        if iid in self._tree_parent_items or iid in self._tree_orphan_items:
            self._tree_open_contracts.add(iid)
            self._tree_forced_closed.discard(iid)
        self.after_idle(self._update_expand_heading)

    def _track_tree_close(self, _event=None):
        iid = self.table.focus()
        if iid in self._tree_parent_items or iid in self._tree_orphan_items:
            self._tree_open_contracts.discard(iid)
            self._tree_forced_closed.add(iid)
        self.after_idle(self._update_expand_heading)

    def _click_contract_row(self, event):
        """Single-click ANY part of a contract row to toggle its children.

        The checkbox column keeps its separate bulk-selection behaviour.
        Header, whitespace, and PDF children are not intercepted.
        """
        if self.table.identify_region(event.x, event.y) not in {'cell', 'tree'}:
            return None
        iid = self.table.identify_row(event.y)
        if iid:
            box = self.table.bbox(iid, 'choose')
            if box and box[0] <= event.x < box[0] + box[2]:
                return None  # checkbox has its own selection action
        if iid not in (self._tree_parent_items | self._tree_orphan_items):
            return None
        if not self.table.exists(iid):
            return 'break'
        self.table.selection_set(iid)
        self.table.focus(iid)
        now_open = not bool(self.table.item(iid, 'open'))
        self.table.item(iid, open=now_open)
        if now_open:
            self._tree_open_contracts.add(iid)
            self._tree_forced_closed.discard(iid)
        else:
            self._tree_open_contracts.discard(iid)
            self._tree_forced_closed.add(iid)
        self._show_detail()
        self.after_idle(self._update_expand_heading)
        return 'break'  # prevent Tk's native indicator from toggling a second time

    def _double_click_pdf(self, event):
        # A double click on a heading, separator, scroll bar or empty space
        # must never fall through to the previously selected PDF.
        if self.table.identify_region(event.x, event.y) not in {"tree", "cell"}:
            return "break"
        iid = self.table.identify_row(event.y)
        if not iid:
            return "break"
        if iid in self._tree_parent_items or iid in self._tree_orphan_items:
            # Tk dispatches the second rapid click as <Double-1> rather than
            # <Button-1>. Toggle once for that second click as well, but never
            # open any PDF belonging to the parent group.
            return self._click_contract_row(event)
        if iid not in self._shown_rows:
            return "break"
        # Select exactly the clicked PDF, even if another file was previously
        # selected or the user has sorted/filtered the table.
        self.table.selection_set(iid)
        self.table.focus(iid)
        self._open_selected()
        return "break"

    def _checked_symbol(self, members: list[ScanRow]):
        # A contract selection includes its reliable delivery notes; the archive
        # planner still checks exact contracts and review state before writing.
        eligible = ([r for r in members] if self.view_var.get() == '归档管理'
                    else [r for r in members if r.kind in {'订单', '发票', '发货单', '贷项通知单'}])
        if not eligible:
            return '—'
        count = sum(path_key(r.source) in self.checked_paths for r in eligible)
        return '☑' if count == len(eligible) else ('◩' if count else '☐')

    def _store_row(self, iid: str, row: ScanRow, members: list[ScanRow]):
        self._shown_rows[iid] = row
        self._tree_members[iid] = members

    @staticmethod
    def _contract_visual_category(group) -> str:
        """Visual priority: review/amount problems > pending > clean invoicing.

        This is intentionally display-only and does not alter business states.
        """
        stage = str(group.stage)
        if group.issues or stage in {"金额待核", "票据待核"} or "异常" in group.state or "需核" in group.state:
            return "alert"
        if stage in {"待开票", "部分开票"}:
            return "pending"
        if stage in {"已开票", "已开票（归档）"}:
            return "done"
        return "neutral"

    def _get_excel_order_notes(self) -> dict[str, str]:
        """Read original 订单!G remarks by B contract, without opening Excel COM.

        Cache by file identity/mtime/size; switching employee/workbook, renaming
        the date suffix, or saving Excel invalidates the cache automatically.
        """
        target = Path(self.excel_var.get().strip()).expanduser()
        try:
            stat = target.stat() if target.is_file() else None
            cache_key = (str(target), stat.st_mtime_ns, stat.st_size) if stat else (str(target), None)
        except OSError:
            cache_key = (str(target), 'unavailable')
        if cache_key != self._excel_note_cache_key:
            self._excel_note_cache_key = cache_key
            self._excel_note_cache = {}
            self._excel_note_error = ''
            if cache_key[1] not in (None, 'unavailable'):
                try:
                    self._excel_note_cache = read_order_notes(target)
                except Exception as exc:
                    # Never present notes from an outdated workbook or another employee.
                    self._excel_note_error = str(exc)
                    self._excel_note_cache_key = None  # Retry next refresh after transient lock/sync problems.
        return self._excel_note_cache

    def _order_note_preview(self, contract: str) -> str:
        if self._excel_note_error:
            return '⚠ 备注读取失败'
        return brief_note(self._excel_note_cache.get(contract, ''))

    def _row_values(self, r: ScanRow, status: str | None = None, symbol: str | None = None):
        return (symbol if symbol is not None else ('☑' if path_key(r.source) in self.checked_paths else '☐'),
                r.kind, r.display_customer, _row_identifier(r), r.day, r.gross, r.net,
                status if status is not None else r.status, '')

    def _refresh_table(self):
        old_selection_paths = [path_key(self._shown_rows[i].source) for i in self.table.selection()
                               if i in self._shown_rows]
        # Persist only user-driven expand/collapse events. Search auto-expansion
        # must not become a permanent opened state after the search is cleared.
        self.table.delete(*self.table.get_children())
        self._shown_rows = {}
        self._tree_members = {}
        self._tree_parent_items = set()
        self._tree_orphan_items = set()
        flt = self.filter_var.get()
        q = self.search_var.get().strip().casefold()
        archive_view = self.view_var.get() == '归档管理'
        note_drafts = self.note_store.list_drafts()
        self._get_excel_order_notes()
        if archive_view:
            self._render_archived_contracts(flt, q)
        else:
            groups, unmatched = make_tree(self.rows, set(self.archived_map), path_key)
            def grp_match(g):
                parent_hit = (q and q in ' '.join((g.contract,g.customer,g.stage,g.state,
                                                     self._excel_note_cache.get(g.contract, ''))).casefold())
                matching = [r for r in g.rows if matches_filter(r, flt if flt not in
                            ('已开票','部分开票','金额待核','发货表已登记') else '全部', q)]
                status_hit = flt in (g.stage, '待复核') and (flt != '待复核' or g.issues > 0)
                if flt in ('已开票','部分开票','金额待核','发货表已登记'):
                    return status_hit and (not q or bool(matching) or bool(parent_hit)), matching, bool(q or status_hit)
                if flt == '全部' and not q:
                    return True, g.rows, False
                if parent_hit and flt == '全部':
                    return True, g.rows, True
                if flt == '待开票' and g.stage == '待开票' and (not q or bool(matching) or bool(parent_hit)):
                    return True, g.rows, bool(q)
                if status_hit and not q:
                    return True, g.rows, True
                return bool(matching), matching, bool(matching)

            # Read-only Excel review may find missing invoices, zero amounts or
            # conflicts that PDF-only scanning cannot. Surface those on the
            # collapsed contract row without altering the source PDF statuses.
            if self.business_plan is not None:
                for g in groups:
                    pattern = re.compile(r'(?<!\d)' + re.escape(g.contract) + r'(?!\d)')
                    count = sum(bool(pattern.search(str(w))) for w in self.business_plan.warnings)
                    if count:
                        g.issues += count
                        g.state += f' · 异常{count}'
            display_groups = []
            for g in groups:
                show, filtered_members, expand = grp_match(g)
                if show:
                    display_groups.append((g, filtered_members, expand))
            if self._sort_column:
                def group_key(item):
                    g = item[0]
                    field = {'identifier': g.contract, 'status': g.state,
                             'remark': self._excel_note_cache.get(g.contract, '')}.get(self._sort_column)
                    if field is None:
                        field = getattr(g, self._sort_column, '')
                    return _table_sort_key(self._sort_column, field)
                good = [x for x in display_groups if group_key(x) is not None]
                blank = [x for x in display_groups if group_key(x) is None]
                good.sort(key=group_key, reverse=self._sort_descending)
                display_groups = good + blank
            else:
                display_groups.sort(key=lambda x: (_table_sort_key('day', x[0].day) or datetime.min.date(), _natural_sort_key(x[0].contract)), reverse=True)

            row_indices = {id(r):i for i,r in enumerate(self.rows)}
            for group_number, (g, filtered_members, expand) in enumerate(display_groups):
                iid = 'contract:' + g.contract
                self._store_row(iid,g.primary,g.rows)
                self._tree_parent_items.add(iid)
                # Keep the underlying g.stage / g.state untouched for search,
                # filters, archival rules and Excel reconciliation.
                contract_category = self._contract_visual_category(g)
                self.table.insert('', 'end', iid=iid, text='', open=(iid not in self._tree_forced_closed and (expand or iid in self._tree_open_contracts)),
                    values=(self._checked_symbol(g.rows),f'订单·{len(g.rows)}份',g.primary.display_customer,g.contract,g.day,g.gross,
                            g.net,g.state + (f' · 待写备注{sum(d.contract == g.contract for d in note_drafts)}' if any(d.contract == g.contract for d in note_drafts) else ''),
                            self._order_note_preview(g.contract)),
                    tags=(f'contract_{contract_category}_{group_number % 2}',))
                order_first = {'订单':0,'发票':1,'发货单':2,'贷项通知单':3}
                filtered_members = sorted(filtered_members,key=lambda r:(order_first.get(r.kind,4),r.day,r.invoice,r.source))
                for k,r in enumerate(filtered_members):
                    child = f'file:{row_indices[id(r)]}'  # identity is stable through sorting and filtering
                    self._store_row(child,r,[r])
                    txt = '待复核·已复核' if r.status == '待复核' and self.archive_store.is_reviewed(r) else r.status
                    tag = ('review',) if r.status=='待复核' else (('pending',) if r.status=='待开票' else ())
                    self.table.insert(iid,'end',iid=child,values=self._row_values(r,txt),tags=tag)

            # Group legacy leftover delivery notes whose Ordernumber links to
            # a UNIQUE archived order+invoice group.  They are not "unknown":
            # the active order has simply been archived in an older version.
            historical_groups, _ = plan_historical_delivery_repair(
                self.rows, self.archive_store.all_archived(), set(self.archived_map),
                self.archive_store.is_reviewed)
            historical = [r for _, members in historical_groups for r in members]
            historical_ids = {id(r) for r in historical}
            historical_filtered = [r for r in historical if
                (flt == '关联订单已归档' or matches_filter(r, flt, q)) and
                (not q or matches_filter(r, '全部', q))]
            if historical_filtered:
                hist_iid = 'historical-delivery'
                self._tree_orphan_items.add(hist_iid)
                self.table.insert('', 'end', iid=hist_iid, text='',
                    open=(hist_iid not in self._tree_forced_closed and
                          (bool(q) or flt != '全部' or hist_iid in self._tree_open_contracts)),
                    values=('—', '历史发货单待补归档', '', '', '', '', '',
                            f'{len(historical_filtered)} 份 · 关联订单已归档', ''), tags=('orphan',))
                ids = {id(r) for r in historical_filtered}
                hist_data = [(i,r) for i,r in enumerate(self.rows) if id(r) in ids]
                if self._sort_column:
                    hist_data = _sort_scan_items(hist_data, self._sort_column, self._sort_descending)
                for i,r in hist_data:
                    child = f'file:{i}'
                    self._store_row(child, r, [r])
                    self.table.insert(hist_iid,'end',iid=child,
                        values=self._row_values(r,'关联订单已归档 · 待补归档'),tags=('pending',))
            unmatched = [r for r in unmatched if id(r) not in historical_ids]

            orphan_filtered = [r for r in unmatched if matches_filter(r, flt, q) or
                               (flt == '未匹配' and (not q or matches_filter(r,'全部',q)))]
            if orphan_filtered:
                orphan_iid = 'unmatched'
                self._tree_orphan_items.add(orphan_iid)
                self.table.insert('','end',iid=orphan_iid,text='',open=(orphan_iid not in self._tree_forced_closed and (bool(q) or flt!='全部' or orphan_iid in self._tree_open_contracts)),
                      values=('—','未匹配文件','','','','','',
                              (f'{len(orphan_filtered)} 份 · 需核{sum(r.status == "待复核" for r in orphan_filtered)}' if any(r.status == '待复核' for r in orphan_filtered) else f'{len(orphan_filtered)} 份文件'),''), tags=('orphan',))
                orphan_ids = {id(r) for r in orphan_filtered}
                orphan_data = [(i,r) for i,r in enumerate(self.rows) if id(r) in orphan_ids]
                if self._sort_column:
                    orphan_data = _sort_scan_items(orphan_data, self._sort_column, self._sort_descending)
                for i,r in orphan_data:
                    child = f'file:{i}'
                    self._store_row(child,r,[r])
                    self.table.insert(orphan_iid,'end',iid=child, values=self._row_values(r),tags=('review',) if r.status=='待复核' else ())

        for iid,row in self._shown_rows.items():
            if path_key(row.source) in old_selection_paths and iid not in self._tree_parent_items:
                self.table.selection_add(iid)
                break
        self._update_tree_selection_label()
        self._update_expand_heading()
        self.bulk_archive_button.configure(state='disabled' if archive_view or not self.scanned_complete else 'normal')
        self.bulk_restore_button.configure(state='normal' if archive_view else 'disabled')
        self.undo_archive_button.configure(state='normal' if self._last_archive_paths else 'disabled')

    def _render_archived_contracts(self, flt: str, q: str):
        """The archive view uses the same parent/child contract tree as results.

        Only archived PDFs appear. Never drop records absent from the current
        scan: SQLite history is authoritative for reversible restoration.
        """
        groups, unmatched, archive_items = archive_tree(
            self.rows, self.archived_map, self.archive_store.all_archived(),
            self._cached_delivery_numbers)
        mode_labels = {'pending':'待开票', 'void':'作废', 'reviewed':'已复核',
                       'matched':'关联合同', 'delivery_followup':'补充发货单',
                       'credit':'贷项通知单'}
        archive_mode_filters = {'关联合同归档':'matched', '补充归档发货单':'delivery_followup',
                                '待开票归档':'pending', '作废归档':'void', '已复核归档':'reviewed'}

        def archive_status(r):
            item = archive_items.get(path_key(r.source))
            if item is None:
                return '已归档'
            # Missing files must remain in archive management, but cannot be
            # restored as though their original bytes still existed.
            if not Path(item.path).is_file():
                return '已归档 · 源文件缺失'
            if r.status == '已归档':
                return '已归档 · 来源待核'
            return '已归档 · ' + mode_labels.get(item.mode, item.mode)

        def match_file(r):
            status = archive_status(r)
            item = archive_items.get(path_key(r.source))
            status_match = (flt in ('全部', r.kind, '已归档', r.status, status) or
                            (item is not None and archive_mode_filters.get(flt) == item.mode) or
                            (flt == '未匹配' and r in unmatched))
            query_match = not q or q in ' '.join((r.kind, r.customer, r.contract,
                                                 r.invoice, r.day, r.gross, r.net,
                                                 r.delivery_no, status, r.source, r.note)).casefold()
            return status_match and query_match

        visible_groups = []
        for group in groups:
            parent_hit = q and q in ' '.join((group.contract, group.customer,
                                              group.state, self._excel_note_cache.get(group.contract, ''))).casefold()
            filtered = [r for r in group.rows if match_file(r)]
            if flt == '全部' and not q:
                visible_groups.append((group, group.rows, False))
            elif parent_hit and flt in ('全部', '订单', '已归档'):
                visible_groups.append((group, group.rows, True))
            elif filtered:
                visible_groups.append((group, filtered, bool(q)))
        if self._sort_column:
            def group_key(t):
                g = t[0]
                raw = {'identifier': g.contract, 'status': g.state,
                       'remark': self._excel_note_cache.get(g.contract, '')}.get(
                           self._sort_column, getattr(g, self._sort_column, ''))
                return _table_sort_key(self._sort_column, raw)
            filled = [x for x in visible_groups if group_key(x) is not None]
            blank = [x for x in visible_groups if group_key(x) is None]
            filled.sort(key=group_key, reverse=self._sort_descending)
            visible_groups = filled + blank
        else:
            visible_groups.sort(key=lambda x: (_table_sort_key('day', x[0].day)
                                                     or datetime.min.date(),
                                                     _natural_sort_key(x[0].contract)), reverse=True)
        number = 0
        for stripe, (group, filtered, force_open) in enumerate(visible_groups):
            iid = 'archive-contract:' + group.contract
            self._store_row(iid, group.primary, group.rows)
            self._tree_parent_items.add(iid)
            self.table.insert('', 'end', iid=iid,
                              open=(iid not in self._tree_forced_closed and
                                    (force_open or iid in self._tree_open_contracts)),
                              values=(self._checked_symbol(group.rows), f'订单·{len(group.rows)}份', group.primary.display_customer,
                                      group.contract, group.day, group.gross, group.net,
                                      group.state, self._order_note_preview(group.contract)),
                              tags=(f'archive_group_{stripe % 2}',))
            for r in sorted(filtered, key=lambda x: ({'订单':0,'发票':1,'发货单':2}.get(x.kind,3),
                                                    x.day, x.invoice, x.source)):
                child = f'archive-file:{number}'; number += 1
                self._store_row(child, r, [r])
                self.table.insert(iid, 'end', iid=child,
                                  values=self._row_values(r, archive_status(r)), tags=('cancel',))
        orphan_filtered = [r for r in unmatched if match_file(r)]
        if orphan_filtered:
            orphan_iid = 'archive-unmatched'
            self._tree_orphan_items.add(orphan_iid)
            self.table.insert('', 'end', iid=orphan_iid,
                              open=(orphan_iid not in self._tree_forced_closed and
                                    (bool(q) or orphan_iid in self._tree_open_contracts)),
                              values=('—','未匹配文件','','','','','',
                                      f'{len(orphan_filtered)} 份已归档 PDF',''),
                              tags=('archive_orphan',))
            sorted_orphans = [(i,r) for i,r in enumerate(orphan_filtered)]
            if self._sort_column:
                sorted_orphans = _sort_scan_items(sorted_orphans, self._sort_column, self._sort_descending)
            for _, r in sorted_orphans:
                child = f'archive-file:{number}'; number += 1
                self._store_row(child, r, [r])
                self.table.insert(orphan_iid, 'end', iid=child,
                                  values=self._row_values(r, archive_status(r)), tags=('cancel',))

    def _switch_view(self):
        self.checked_paths.clear()
        self.status_filter_combo.configure(values=(self.archive_filters if self.view_var.get() == '归档管理'
                                                   else self.active_filters))
        self.filter_var.set('全部')
        self._refresh_table()

    def _copy_clicked_cell(self, event):
        """Single-click the number or *full* Excel order note to copy it.

        The contract tree still expands when clicking any other parent cell,
        and the double-click PDF shortcut remains unchanged. Use column bounds
        rather than hardcoded #4/#9 indices so column reordering is harmless.
        """
        if self.table.identify_region(event.x, event.y) != 'cell':
            return None
        iid = self.table.identify_row(event.y)
        if not iid or iid not in self._shown_rows:
            return None
        chosen = None
        for column in ('identifier', 'remark'):
            bounds = self.table.bbox(iid, column)
            if bounds and bounds[0] <= event.x < bounds[0] + bounds[2]:
                chosen = column
                break
        if chosen is None:
            return None
        row = self._shown_rows[iid]
        if chosen == 'identifier':
            value = str(self.table.set(iid, 'identifier')).strip()
            label = '编号'
        else:
            # Preview text may be abbreviated. Copy the entire original
            # 订单!G note, and never mistake a pending-note draft for Excel data.
            if iid not in self._tree_parent_items:
                return None
            value = str(self._excel_note_cache.get(row.contract, '')).strip()
            label = 'Excel 订单备注'
        if not value or value in ('—', '-', '⚠ 备注读取失败'):
            return None
        try:
            self.clipboard_clear()
            self.clipboard_append(value)
        except tk.TclError:
            return None
        self.table.selection_set(iid)
        self.table.focus(iid)
        self._show_detail()
        self.status_var.set(f'已复制{label}：{brief_note(value)}')
        return 'break'

    def _toggle_checkbox(self, event):
        iid = self.table.identify_row(event.y)
        if not iid or self.table.identify_region(event.x,event.y) not in ('cell', 'tree'):
            return
        box = self.table.bbox(iid, 'choose')
        if not box or not (box[0] <= event.x < box[0] + box[2]):
            return
        if iid in self._tree_orphan_items:
            return 'break'
        members = self._tree_members.get(iid, [])
        if not members:
            return 'break'
        if iid in self._tree_parent_items:
            members = ([r for r in members] if self.view_var.get() == '归档管理'
                       else [r for r in members if r.kind in {'订单','发票','发货单','贷项通知单'}])
        paths = {path_key(r.source) for r in members}
        if paths and paths.issubset(self.checked_paths):
            self.checked_paths.difference_update(paths)
        else:
            self.checked_paths.update(paths)
        self._refresh_table()
        return 'break'

    def _select_all_visible(self):
        paths = set()
        for iid,members in self._tree_members.items():
            if iid in self._tree_parent_items or iid in self._tree_orphan_items:
                continue
            for r in members:
                if r.kind in {'订单', '发票', '发货单', '贷项通知单'} or self.view_var.get() == '归档管理':
                    paths.add(path_key(r.source))
        if paths and paths.issubset(self.checked_paths):
            self.checked_paths.difference_update(paths)
        else:
            self.checked_paths.update(paths)
        self._refresh_table()

    def _clear_checked(self):
        self.checked_paths.clear()
        self._refresh_table()

    def _show_detail(self, _event=None):
        selected = self.table.selection()
        if not selected: return
        r = self._shown_rows.get(selected[0])
        if not r: return
        if selected[0] in self._tree_orphan_items:
            return
        archived = path_key(r.source) in self.archived_map or self.view_var.get() == '归档管理'
        reviewed = '；人工已复核' if self.archive_store.is_reviewed(r) else ''
        msg = f"状态：{'已归档' if archived else r.status}{reviewed}    类型：{r.kind}    合同号：{r.contract}    发票号：{r.invoice}"
        if r.kind == '发货单':
            msg += f"    发货单号：{r.delivery_no}"
        msg += f"\n说明：{r.note}\n来源：{r.source}"
        msg += f"\n客户正式名称：{r.customer}（来源：{r.customer_source or '原有识别规则'}）"
        if r.display_customer != r.customer:
            msg += f"；界面简称：{r.display_customer}"
        if selected[0] in self._tree_parent_items:
            group_members = self._tree_members[selected[0]]
            msg = (f'合同：{r.contract}  |  {self.table.set(selected[0], "status")}  | 关联 PDF：{len(group_members)} 份\n'
                   + ('点击箭头查看已归档 PDF；右键恢复此合同将按原归档组恢复其全部文件。\n'
                    if self.view_var.get() == '归档管理' else
                    '点击箭头查看每份 PDF 的详细状态；订单或发票右键归档仍按原有业务规则确认。\n')
                   + '\n'.join(f'· {member.kind}: {Path(member.source).name}  [{member.status}]' for member in group_members))
            msg += f"\n客户正式名称：{r.customer}（来源：{r.customer_source or '原有识别规则'}）"
            if r.display_customer != r.customer:
                msg += f"；界面简称：{r.display_customer}"
            order_excel_note = self._excel_note_cache.get(r.contract, '')
            if order_excel_note:
                msg += f'\n\nExcel 订单备注（G 列）：\n{order_excel_note}'
            elif self._excel_note_error:
                msg += '\n\n⚠ 订单备注读取失败：' + self._excel_note_error
        if archived:
            item = self.archived_map.get(path_key(r.source))
            if item is None:
                item = next((x for x in self.archive_store.all_archived() if x.path == path_key(r.source)), None)
            if item:
                msg += f'\n归档方式：{item.mode}；归档时间：{item.archived_at}；备注：{item.reason or "无"}'
                state = self.archive_store.source_state(item)
                if state != 'ok':
                    msg += ('\n⚠ ' + ('原始 PDF 不在保存路径或无法读取' if state == 'missing' else '原 PDF 内容与归档时不一致')
                            + '；归档历史仍在。恢复前请重新定位相同指纹 PDF，'
                              '或选择「单独清除失效归档记录」。清除不会恢复或删除 PDF。')
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", msg)
        self.detail.configure(state="disabled")

    def _show_result_context_menu(self, event):
        # identify_row() is based on where the user right-clicked, so sorting
        # and filtering cannot accidentally select an unrelated PDF.
        row_id = self.table.identify_row(event.y)
        if not row_id or row_id in self._tree_orphan_items or row_id not in self._shown_rows:
            return  # No context menu on unmatched root, header or blank space.
        self.table.selection_set(row_id)
        self.table.focus(row_id)
        self._show_detail()
        row = self._shown_rows[row_id]
        is_archive = self.view_var.get() == '归档管理'
        self.result_menu.entryconfigure('归档此 PDF（需确认）', state='disabled' if is_archive or not self.scanned_complete or
             (row.status == '待复核' and not self.archive_store.is_reviewed(row)) else 'normal')
        self.result_menu.entryconfigure(self._restore_menu_index,
            label='恢复此合同（需确认）' if is_archive and row_id in self._tree_parent_items
                else '恢复此 PDF（需确认）',
            state='normal' if is_archive else 'disabled')
        can_note = (not is_archive and self.scanned_complete and not self.business_busy
                    and (not self.worker or not self.worker.is_alive())
                    and bool(row.contract) and row.contract.isdigit()
                    and (row_id in self._tree_parent_items or row.kind in ('订单', '发票')))
        self.result_menu.entryconfigure('添加/编辑 Excel 备注…', state='normal' if can_note else 'disabled')
        self.result_menu.entryconfigure('查询此订单配件库存与库位…', state='normal'
            if row.kind == '订单' and Path(row.source).is_file() else 'disabled')
        self.result_menu.entryconfigure('打印此订单拣货单及发货单…', state='normal'
            if not is_archive and row.kind == '订单' and Path(row.source).is_file() else 'disabled')
        opened_ok = (Path(row.source).is_file() and self.opened_for_review.get(path_key(row.source)) == file_sha256(row.source))
        self.result_menu.entryconfigure('标记为已复核', state='normal' if not is_archive and row.status=='待复核'
                and not self.archive_store.is_reviewed(row) and opened_ok else 'disabled')
        item = self._selected_archive_entry() if is_archive else None
        is_stale = bool(item and self.archive_store.source_state(item) != 'ok')
        self.result_menu.entryconfigure('重新定位原始 PDF…', state='normal' if is_stale else 'disabled')
        self.result_menu.entryconfigure('单独清除失效归档记录…', state='normal' if is_stale else 'disabled')
        try:
            self.result_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.result_menu.grab_release()

    def _note_targets_for_selected(self):
        """Use the Hangcha contract and *exact* invoice number for remarks."""
        selected = self.table.selection()
        if not selected:
            return []
        iid = selected[0]
        if iid in self._tree_orphan_items or self.view_var.get() == '归档管理':
            return []
        row = self._shown_rows.get(iid)
        if not row or not row.contract or not row.contract.isdigit():
            return []
        if iid in self._tree_parent_items:
            members = self._tree_members.get(iid, [])
        else:
            members = self._tree_members.get(iid, [])
        has_order = any(r.kind == '订单' and r.contract == row.contract for r in members)
        targets = []
        if has_order:
            targets.append(('订单备注 | 合同 ' + row.contract, 'order', row.contract, ''))
        invoice_ids = sorted({r.invoice for r in members
                              if r.kind == '发票' and r.contract == row.contract and r.invoice
                              and r.status not in ('忽略作废', '待复核', '重复文件')})
        # A directly selected invoice can be edited even without its order PDF.
        if row.kind == '发票' and row.invoice and row.status not in ('忽略作废', '待复核', '重复文件'):
            invoice_ids = sorted(set(invoice_ids) | {row.invoice})
        for invoice in invoice_ids:
            targets.append((f'发货备注 | 合同 {row.contract} / 发票 {invoice}', 'invoice', row.contract, invoice))
        if row.kind == '订单' and not has_order and row.contract:
            targets.insert(0, ('订单备注 | 合同 ' + row.contract, 'order', row.contract, ''))
        return targets

    def _edit_note_selected(self):
        if not self._action_allowed():
            return
        targets = self._note_targets_for_selected()
        if not targets:
            messagebox.showwarning('无法添加备注', '请选中有明确合同号的订单或发票。待复核文件须先解决识别问题。')
            return
        selected_row = self._selected_row()
        win = tk.Toplevel(self)
        win.title('添加 / 编辑 Excel 备注（本地草稿）')
        win.geometry('660x440')
        win.minsize(520, 375)
        win.transient(self)
        win.grab_set()
        shell = ttk.Frame(win, padding=16)
        shell.pack(fill='both', expand=True)
        ttk.Label(shell, text='选择 Excel 备注的目标行', font=self._fonts['section']).pack(anchor='w')
        ttk.Label(shell, text='订单 → 订单表 G 列；发票 → 发货表 L 列。不会修改 PDF。',
                  style='Small.TLabel').pack(anchor='w', pady=(4, 10))
        choices = [item[0] for item in targets]
        initially = next((item for item in targets
                         if selected_row and selected_row.kind == '发票' and item[1] == 'invoice'
                         and item[3] == selected_row.invoice), targets[0])
        choice_var = tk.StringVar(value=initially[0])
        box = ttk.Combobox(shell, textvariable=choice_var, values=choices, state='readonly')
        box.pack(fill='x', pady=(0, 12))
        ttk.Label(shell, text='备注正文（将保留 Excel 原有备注，并在末尾追加）').pack(anchor='w')
        editor = tk.Text(shell, height=8, wrap='word', font=self._fonts['body'], undo=True)
        editor.pack(fill='both', expand=True, pady=(5, 7))
        edited = {}
        last = [initially]
        def fill(item):
            editor.delete('1.0', 'end')
            editor.insert('1.0', edited.get(item[0], self.note_store.get(item[1], item[2], item[3])))
        fill(initially)
        def changed(*_):
            edited[last[0][0]] = editor.get('1.0', 'end-1c')
            current = next(item for item in targets if item[0] == choice_var.get())
            last[0] = current
            fill(current)
        box.bind('<<ComboboxSelected>>', changed)
        ttk.Label(shell, text='这里只保存本地草稿。下次「核对 / 预览 Excel」会显示变更，\n'
                              '只有你再次明确确认写入后才会修改 Excel。',
                  style='Small.TLabel').pack(anchor='w', pady=(2, 9))
        bottom = ttk.Frame(shell)
        bottom.pack(fill='x')
        def save():
            item = last[0]
            text = editor.get('1.0', 'end-1c').strip()
            if not text:
                messagebox.showwarning('备注为空', '请输入备注，或使用「删除当前草稿」。', parent=win)
                return
            try:
                self.note_store.save(item[1], item[2], item[3], text)
            except Exception as exc:
                messagebox.showerror('草稿保存失败', str(exc), parent=win)
                return
            self._invalidate_plan()
            self._refresh_table()
            self.status_var.set('Excel 备注已保存为本地草稿；尚未写入 Excel')
            win.destroy()
        def discard():
            item = last[0]
            if not self.note_store.get(item[1], item[2], item[3]):
                return
            if messagebox.askyesno('删除草稿', '只删除本地尚未写入的备注草稿，不修改 Excel。确定吗？', parent=win):
                self.note_store.discard(item[1], item[2], item[3])
                self._invalidate_plan()
                self._refresh_table()
                win.destroy()
        ttk.Button(bottom, text='删除当前草稿', command=discard).pack(side='left')
        ttk.Button(bottom, text='取消', command=win.destroy).pack(side='right', padx=(8,0))
        ttk.Button(bottom, text='暂存备注', command=save).pack(side='right')
        win.wait_window()

    def _reveal_selected(self):
        sel = self.table.selection()
        if not sel: return
        row = self._shown_rows.get(sel[0])
        if row is None: return
        pdf = Path(row.source)
        if not pdf.is_file():
            messagebox.showwarning("文件不存在", f"找不到这个 PDF，可能已经移动或删除：\n{pdf}")
            return
        if sys.platform == "win32":
            try:
                _reveal_windows_pdf(pdf)
                return
            except Exception as exc:
                # If Shell selection fails (e.g. unavailable network folder),
                # open the *correct* parent directly, never default Documents.
                try:
                    os.startfile(str(pdf.parent))
                except Exception as fallback_exc:
                    messagebox.showerror("在文件夹中显示失败", f"目标：{pdf}\n\n定位失败：{exc}\n打开文件夹也失败：{fallback_exc}")
                else:
                    messagebox.showwarning("已打开 PDF 所在文件夹", f"无法自动选中目标 PDF，已打开它所在的文件夹。\n\n文件：{pdf.name}\n文件夹：{pdf.parent}\n\n原因：{exc}")
                return
        try:
            subprocess.Popen(_reveal_command(pdf))
        except Exception as exc:
            messagebox.showerror("在文件夹中显示失败", str(exc))

    def _open_selected(self):
        sel = self.table.selection()
        if not sel: return
        r = self._shown_rows.get(sel[0])
        if not r or not Path(r.source).is_file(): return
        try:
            if os.name == "nt": os.startfile(r.source)
            elif sys.platform == "darwin": subprocess.Popen(["open", r.source])
            else: subprocess.Popen(["xdg-open", r.source])
            self.opened_for_review[path_key(r.source)] = file_sha256(r.source)
        except Exception as exc: messagebox.showerror("打开失败", str(exc))

    def _action_allowed(self) -> bool:
        if self._get_config() != self._last_config:
            messagebox.showwarning('配置已更改', '来源目录、Excel 或业务员设置有变化，请重新扫描后再归档。')
            return False
        if self.business_busy or (self.worker and self.worker.is_alive()):
            messagebox.showwarning('请先完成任务', '扫描或 Excel 核对正在进行，请稍后再操作归档。')
            return False
        if not self.scanned_complete:
            messagebox.showwarning('需要完整扫描', '请先完成一次完整的 PDF 扫描，然后再操作归档。')
            return False
        return True

    def _mark_reviewed_selected(self):
        row = self._selected_row()
        if row is None or row.status != '待复核' or self.view_var.get() == '归档管理':
            return
        key = path_key(row.source)
        current_hash = file_sha256(row.source) if Path(row.source).is_file() else ''
        if self.opened_for_review.get(key) != current_hash:
            messagebox.showwarning('请先查看 PDF', '必须先通过本程序的「打开 PDF」查看当前文件，再返回选择「标记为已复核」。')
            return
        reason = simpledialog.askstring('人工复核结论',
                f'文件：{Path(row.source).name}\n\n请选择复核结论/原因（例如：已核实、无需登记、其他）：\n'
                '此标记不会修改 PDF，也不会把异常文件自动加入 Excel。', parent=self)
        if not reason or not reason.strip():
            return
        if not messagebox.askyesno('确认人工已复核',
                f'确定已核对此 PDF 的内容？\n\n{Path(row.source).name}\n结论：{reason}\n\n'
                '只允许后续单独归档；不会改变原有「待复核」状态或 Excel 写入校验。', parent=self):
            return
        try:
            self.archive_store.mark_reviewed(row, reason)
            self._show_detail()
            self.status_var.set('人工复核标记已保存。需要归档时仍需单独确认。')
        except Exception as exc:
            messagebox.showerror('复核保存失败', str(exc))

    def _selected_row(self):
        sel = self.table.selection()
        return self._shown_rows.get(sel[0]) if sel else None

    def _archive_selected(self):
        row = self._selected_row()
        if row is not None:
            self._archive_rows([row])

    def _archive_checked(self):
        selected = [r for r in self.rows if path_key(r.source) in self.checked_paths]
        if not selected:
            messagebox.showinfo('未选择', '请先勾选需要归档的订单、发票或发货单。')
        else:
            self._archive_rows(selected)

    def _archive_rows(self, rows):
        if not self._action_allowed() or self.view_var.get() == '归档管理':
            return
        groups, blocked = plan_archive(rows, self.rows, set(self.archived_map),
                                       self.archive_store.is_reviewed,
                                       self.archive_store.all_archived())
        if not groups:
            messagebox.showwarning('无法归档', '\n'.join(blocked[:20]) or '没有符合条件的文件')
            return
        all_files = [r for _,members in groups for r in members]
        try:
            for r in all_files:
                if self.scan_hashes.get(path_key(r.source)) != file_sha256(r.source):
                    raise RuntimeError(f'PDF 在上次扫描后发生变化，请重新扫描：{r.source}')
        except (OSError, RuntimeError) as exc:
            messagebox.showwarning('源文件已变化', str(exc))
            return
        # Make sure the program only persists documents on which scan result is current.
        # The archive store re-hashes every target before an atomic DB transaction.
        modes = Counter(mode for mode,_ in groups)
        lines = [f'本次共选中 {len(rows)} 条；按合同号展开后将归档 {len(all_files)} 份 PDF（{len(groups)} 个归档组）。',
                 f'关联合同归档 {modes["matched"]} 组（包含可靠关联的发货单） / '
                 f'历史遗留发货单补归档 {modes["delivery_followup"]} 组 / '
                 f'作废 {modes["void"]} 份 / 已复核 {modes["reviewed"]} 份 / '
                 f'待开票 {modes["pending"]} 组（含可靠关联发货单） / 贷项通知单 {modes["credit"]} 份。',
                 '原始 PDF 不移动、不删除；只改变 GUI 内的显示与自动写入资格。',
                 '归档不会把未开票订单标记为已开票。', '', '=== 将归档的 PDF ===']
        lines.extend(f'{mode} | {r.contract} | {Path(r.source).name} | {r.source}' for mode, members in groups for r in members)
        if blocked:
            lines += ['', f'=== 不符合条件，以下 {len(blocked)} 项本次不会处理 ==='] + blocked
        if not self._approval_preview('确认归档范围', '\n'.join(lines), '确认归档上述 PDF'):
            return
        if modes['pending']:
            contracts = ', '.join(sorted({r.contract for mode,members in groups if mode=='pending' for r in members}))
            if not messagebox.askyesno('二次确认：待开票订单',
                    f'以下 {modes["pending"]} 组订单尚无有效发票：\n{contracts}\n\n'
                    '归档后它们将从普通待开票列表隐藏，但后台仍保留真实未开票状态。\n'
                    '你是否仍要归档这些未开票订单？', icon='warning', parent=self):
                return
        try:
            affected = self.archive_store.archive(groups)
            self._last_archive_paths = affected
            self.archived_map = self.archive_store.archived_for(self.rows)
            self._load_cached_archive_numbers()
            self.checked_paths.clear()
            self._invalidate_plan()
            self._refresh_stats(); self._refresh_table()
            self.status_var.set(f'已归档 {len(affected)} 个 PDF；Excel 未修改')
            messagebox.showinfo('归档成功', f'归档 {len(affected)} 份 PDF。可切换到「归档管理」随时恢复。\n原始 PDF 和 Excel 均未修改。')
        except Exception as exc:
            messagebox.showerror('归档失败', str(exc))

    def _repair_historical_deliveries(self):
        """Preview and explicitly approve batch repair of leftover delivery PDFs.

        Never silently migrate a prior employee's archive or touch original PDFs.
        """
        if not self._action_allowed():
            return
        if self.view_var.get() == '归档管理':
            messagebox.showinfo('请返回识别结果', '请先切回「识别结果」，再检查未归档发货单。', parent=self)
            return
        groups, unresolved = plan_historical_delivery_repair(
            self.rows, self.archive_store.all_archived(), set(self.archived_map),
            self.archive_store.is_reviewed)
        total = sum(len(files) for _, files in groups)
        if not groups:
            details = '\n'.join(f'【{reason}】{desc}' for reason, desc in unresolved[:35])
            messagebox.showinfo('历史发货单核对结果',
                f'没有可以安全补充归档的发货单。待人工处理 {len(unresolved)} 份。\n\n{details or "未找到未归档发货单。"}',
                parent=self)
            return
        lines = [f'找到 {total} 份可匹配到唯一历史订单归档组的发货单，涉及 {len(groups)} 个合同。',
                 '必须确认后才会写入归档数据库；不会移动或删除原 PDF，也不会改变 Excel。',
                 '以下发货单使用 PDF 正文中唯一的 Ordernumber 关联，不使用发货单号或文件名。', '',
                 '=== 本次可补充归档 ===']
        for mode, files in groups:
            for r in files:
                lines.append(f'Ordernumber={r.contract} | Delivery No.={r.delivery_no or "待核"} | {r.source}')
        lines.extend(['', f'=== 未能安全补归档：{len(unresolved)} 份（本次不操作） ==='])
        lines.extend(f'【{reason}】{desc}' for reason,desc in unresolved)
        if not self._approval_preview('历史发货单归档完整性核对', '\n'.join(lines),
                                      f'确认补充归档 {total} 份发货单'):
            return
        try:
            # The database re-hashes and validates every PDF, and appends the
            # new record to the original unique archive group atomically.
            affected = self.archive_store.archive(groups, reason='v1.5.34 历史发货单人工确认补归档')
            self._last_archive_paths = affected
            self.archived_map = self.archive_store.archived_for(self.rows)
            self._load_cached_archive_numbers()
            self.checked_paths.clear()
            self._invalidate_plan()
            self._refresh_stats()
            self._refresh_table()
            self.status_var.set(f'已补充归档 {len(affected)} 份历史发货单；原 PDF 与 Excel 未修改')
            messagebox.showinfo('历史归档补充完成',
                f'已补充归档 {len(affected)} 份发货单。\n'
                f'仍有 {len(unresolved)} 份未完成唯一匹配或需要人工处理。\n'
                '请到「归档管理」展开对应订单核对。', parent=self)
        except Exception as exc:
            messagebox.showerror('历史发货单补归档失败', str(exc), parent=self)

    def _approval_preview(self, title: str, text: str, approve_label: str) -> bool:
        """Show the full affected-file list before performing a destructive UI-state change."""
        from tkinter.scrolledtext import ScrolledText
        win = tk.Toplevel(self)
        win.title(title); win.geometry('850x570'); win.transient(self); win.grab_set()
        answer = {'yes': False}
        ttk.Label(win, text='请核对全部受影响文件。取消不会更改归档状态。').pack(anchor='w', padx=12, pady=10)
        content = ScrolledText(win, font=self._fonts["mono"], wrap='word')
        content.pack(fill='both', expand=True, padx=12, pady=5)
        content.insert('1.0', text); content.configure(state='disabled')
        bottom = ttk.Frame(win); bottom.pack(fill='x', padx=12, pady=10)
        def submit():
            answer['yes'] = True; win.destroy()
        ttk.Button(bottom, text='取消', command=win.destroy).pack(side='right', padx=(8,0))
        ttk.Button(bottom, text=approve_label, command=submit).pack(side='right')
        win.protocol('WM_DELETE_WINDOW', win.destroy)
        self.wait_window(win)
        return answer['yes']

    def _selected_archive_entry(self):
        """Only real PDF child rows can be re-linked or individually cleared."""
        if self.view_var.get() != '归档管理':
            return None
        selection = self.table.selection()
        if not selection or selection[0] in self._tree_parent_items or selection[0] in self._tree_orphan_items:
            return None
        row = self._shown_rows.get(selection[0])
        if row is None:
            return None
        key = path_key(row.source)
        item = self.archived_map.get(key)
        if item:
            return item
        return next((r for r in self.archive_store.all_archived() if r.path == key), None)

    def _relink_archive_selected(self):
        if self.view_var.get() != '归档管理' or not self._action_allowed():
            return
        item = self._selected_archive_entry()
        if item is None:
            messagebox.showinfo('请选择归档 PDF', '请展开合同或未匹配文件组，选中一份具体的归档 PDF。', parent=self)
            return
        if self.archive_store.source_state(item) == 'ok':
            messagebox.showinfo('无需重新定位', '当前 PDF 文件仍在原位置且内容指纹一致。', parent=self)
            return
        chosen = filedialog.askopenfilename(parent=self, title='重新定位归档时的原始 PDF',
                                            filetypes=[('PDF 文件', '*.pdf')])
        if not chosen:
            return
        try:
            actual = file_sha256(chosen)
        except OSError as exc:
            messagebox.showerror('无法读取所选 PDF', str(exc), parent=self)
            return
        if actual != item.sha:
            messagebox.showerror('文件指纹不匹配',
                '所选 PDF 与归档时的原文件内容不同，不能用同名文件或相同发票号代替。\n\n'
                f'保存指纹：{item.sha}\n所选指纹：{actual}', parent=self)
            return
        if not messagebox.askyesno('确认重新定位',
                f'合同号：{item.contract or "—"}\n发票号：{item.invoice or "—"}\n'
                f'原位置：{item.path}\n新位置：{chosen}\n\n'
                'SHA-256 文件指纹匹配。只更新归档数据库的路径，不复制、移动或修改 PDF/Excel。确认吗？',
                parent=self):
            return
        try:
            self.archive_store.relink_missing(item.path, chosen)
            self.archived_map = self.archive_store.archived_for(self.rows)
            self._last_archive_paths = []  # Do not use stale pre-relink undo roots.
            self.checked_paths.clear()
            self._invalidate_plan(); self._refresh_stats(); self._refresh_table()
            self.status_var.set('归档源文件已重新关联；请重新扫描以刷新文件索引')
            messagebox.showinfo('重新定位成功', '归档组与历史记录已保留，源 PDF 未修改。建议重新扫描确认。', parent=self)
        except Exception as exc:
            messagebox.showerror('重新定位失败', str(exc), parent=self)

    def _clear_invalid_archive_selected(self):
        if self.view_var.get() != '归档管理' or not self._action_allowed():
            return
        item = self._selected_archive_entry()
        if item is None:
            messagebox.showinfo('请选择失效记录', '必须选中具体的归档 PDF 子行，不允许直接清空整个合同组。', parent=self)
            return
        state = self.archive_store.source_state(item)
        if state == 'ok':
            messagebox.showwarning('不能清除有效归档', '原始 PDF 仍存在且指纹相同。若要重新纳入扫描，请使用「恢复」。', parent=self)
            return
        warning = (f'只清除以下 1 条失效归档记录（不会清除整组其他文件）：\n\n'
                   f'类型：{item.kind}\n合同：{item.contract or "—"}\n发票号：{item.invoice or "—"}\n'
                   f'原路径：{item.path}\n异常：{"文件不存在/无法读取" if state == "missing" else "文件内容已被替换"}\n'
                   f'SHA-256：{item.sha}\n\n'
                   '⚠ 清除后，这份 PDF 的归档条目将从归档管理中消失，不能通过「恢复归档」找回。\n'
                   '不会删除任何 PDF/Excel；仅在数据库审计日志中保留历史记录。\n'
                   '如果它是尚未写入 Excel 的发票，关联订单可能重新显示为待开票。\n'
                   '建议优先尝试「重新定位原始 PDF」。')
        if not self._approval_preview('清除失效归档记录 · 操作预览', warning, '继续，进入最终确认'):
            return
        typed = simpledialog.askstring('最终确认：清除归档历史',
                 '这是不可通过普通「恢复」撤销的操作。\n请输入「清除」以确认只清除这一条失效归档记录：',
                 parent=self)
        if typed != '清除':
            return
        try:
            cleared = self.archive_store.clear_invalid([item.path])
            self._last_archive_paths = []
            self.archived_map = self.archive_store.archived_for(self.rows)
            self._load_cached_archive_numbers()
            self.checked_paths.clear()
            self._invalidate_plan(); self._refresh_stats(); self._refresh_table()
            self.status_var.set(f'已清除 {len(cleared)} 条失效归档记录；原始文件和 Excel 未修改')
            messagebox.showinfo('失效归档记录已清理',
                '已从当前归档管理清除记录。操作详情保留在本地 SQLite 审计日志中；'
                '若相关业务状态有变化，请重新扫描并核对 Excel。', parent=self)
        except Exception as exc:
            messagebox.showerror('清除失败', str(exc), parent=self)

    def _restore_selected(self):
        # A parent row represents a contract, not just its first order PDF.
        # Restore all archived PDF members shown under this contract, while
        # preserving each original archive group's linked restore semantics.
        selected = self.table.selection()
        if not selected:
            return
        iid = selected[0]
        if iid in self._tree_parent_items and self.view_var.get() == '归档管理':
            self._restore_rows(list(self._tree_members.get(iid, [])))
        else:
            row = self._selected_row()
            if row is not None:
                self._restore_rows([row])

    def _restore_checked(self):
        selected = [r for r in self._shown_rows.values() if path_key(r.source) in self.checked_paths]
        if not selected:
            messagebox.showinfo('未选择', '请在归档管理中勾选需要恢复的文件。')
        else:
            self._restore_rows(selected)

    def _restore_rows(self, rows):
        if not self._action_allowed() or self.view_var.get() != '归档管理':
            return
        known = {path_key(r.source):self.archived_map[path_key(r.source)].path for r in rows
                 if path_key(r.source) in self.archived_map}
        roots = [known.get(path_key(r.source),r.source) for r in rows]
        expanded, invalid = self.archive_store.check_restore(roots)
        if not expanded:
            messagebox.showinfo('没有待恢复记录', '这几份文件没有归档记录。')
            return
        if invalid:
            messagebox.showwarning('已阻止恢复：存在失效 PDF',
                f'选中的归档组中有 {len(invalid)} 份源文件缺失或内容变化；本次没有恢复任何记录。\n\n'
                + '\n'.join(f'{"缺失" if state=="missing" else "内容变化"}: {item.path}' for item,state in invalid[:12])
                + '\n\n请在归档管理中选中具体 PDF，使用「重新定位原始 PDF」；'
                  '确实不再需要的记录可单独清除。', parent=self)
            return
        if not self._approval_preview('确认恢复归档',
                  f'根据原始归档组将恢复 {len(expanded)} 份 PDF：\n\n' + '\n'.join(expanded)
                  + '\n\n恢复后重新参与扫描与业务核对；不会修改 PDF 或 Excel。',
                  '确认恢复'):
            return
        try:
            self.archive_store.restore(roots)
            self.archived_map = self.archive_store.archived_for(self.rows)
            self._load_cached_archive_numbers()
            self.checked_paths.clear()
            self._invalidate_plan()
            self._refresh_stats(); self._refresh_table()
            self.status_var.set(f'已恢复 {len(expanded)} 个归档记录；请重新扫描以刷新业务状态')
            messagebox.showinfo('恢复成功', f'已恢复 {len(expanded)} 个归档记录。建议重新扫描以更新识别状态。')
        except Exception as exc:
            messagebox.showerror('恢复失败', str(exc))

    def _undo_last_archive(self):
        if not self._last_archive_paths or not self._action_allowed():
            return
        try:
            expanded, invalid = self.archive_store.check_undo_last_archive(self._last_archive_paths)
        except Exception as exc:
            messagebox.showwarning('撤销不可用', str(exc), parent=self)
            return
        if not expanded: return
        if invalid:
            messagebox.showwarning('撤销已暂停',
                '上次归档组中的原始 PDF 已缺失或内容变化，不能直接撤销，避免归档历史消失。\n'
                + '\n'.join(item.path for item,_ in invalid[:8])
                + '\n请先在「归档管理」重新定位或单独清理。', parent=self)
            return
        if not messagebox.askyesno('撤销上次归档', f'恢复刚才归档的 {len(expanded)} 份 PDF？\n原始 PDF 不受影响。', parent=self):
            return
        try:
            self.archive_store.undo_last_archive(self._last_archive_paths)
        except Exception as exc:
            messagebox.showerror('撤销失败', str(exc))
            return
        self._last_archive_paths = []
        self.archived_map = self.archive_store.archived_for(self.rows)
        self._invalidate_plan(); self._refresh_stats(); self._refresh_table()
        self.status_var.set('已撤销上次归档')

    def _export(self):
        if not self.rows:
            messagebox.showinfo("没有数据", "请先运行扫描。")
            return
        dest = filedialog.asksaveasfilename(title="导出只读扫描结果", defaultextension=".csv", filetypes=[("CSV（Excel兼容）", "*.csv")], initialfile=f"Hangcha_PDF_Scan_{datetime.now():%Y%m%d_%H%M%S}.csv")
        if not dest: return
        try:
            active = [r for r in self.rows if path_key(r.source) not in self.archived_map]
            export_csv(active, dest)
            messagebox.showinfo("导出完成", f"导出 {len(active)} 条结果到：\n{dest}\n\nPDF 与 Excel 均未修改。")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _invalidate_plan(self):
        had_excel_preview = self.business_plan is not None
        self.business_plan = None
        self._preview_cfg = None
        if hasattr(self, "write_button"):
            self.write_button.configure(state="disabled")
        # If the employee switches the target workbook, prior Excel-only invoice
        # evidence no longer applies. Revert to explicitly provisional PDF status.
        if had_excel_preview and self.scanned_complete:
            mark_pending_invoice_status(self.rows)
            self._refresh_stats()
            self._refresh_table()

    def _preview_business(self):
        if self.business_busy or (self.worker and self.worker.is_alive()):
            return
        if not self.scanned_complete:
            messagebox.showwarning("扫描不完整", "请先完成一次完整的 PDF 扫描。")
            return
        cfg = self._get_config()
        if not cfg['target_excel']:
            messagebox.showwarning("未设置 Excel", "请填写本员工的 Excel 绝对路径。文件可以尚不存在，例如 C:\\Reports\\Sales_Xie.xlsx。")
            return
        if cfg['sources'] != self._last_config.get('sources') or cfg['exclude_names'] != self._last_config.get('exclude_names') or cfg['salesperson'] != self._last_config.get('salesperson'):
            messagebox.showwarning("配置已修改", "来源目录、排除规则或业务员已变化，请重新扫描 PDF。")
            return
        self._invalidate_plan()
        self._preview_cfg = json.loads(json.dumps(cfg))
        self.business_busy = True
        self.review_button.configure(state='disabled')
        self.status_var.set('业务核对中：正在只读检查 Excel、订单/发票金额及缺票风险…')
        frozen_rows = list(self.rows)
        frozen_archived = frozenset(self.archived_map)
        frozen_notes = tuple(self.note_store.list_drafts())
        def run():
            try:
                plan = gui_business.make_plan(frozen_rows, cfg['target_excel'], cfg['salesperson'], cfg['sources'], cfg['exclude_names'], archived_paths=frozen_archived, pending_notes=frozen_notes,
                                              note_store_path=self.note_store.db_path, note_profile=self.note_store.profile)
                self.events.put(('business_plan', plan))
            except Exception:
                self.events.put(('business_error', traceback.format_exc()))
        self.business_worker = threading.Thread(target=run, daemon=True)
        self.business_worker.start()

    def _show_business_plan(self, plan):
        window = tk.Toplevel(self)
        window.title('业务核对报告 · 只读预览')
        window.geometry('1040x720')
        tk.Label(window, text=f'待新增订单 {len(plan.orders)} 条 | 待新增发票 {len(plan.invoices)} 条 | 待写备注 {len(plan.note_changes)} 条 | 需复核 {len(plan.warnings)} 条',
                 font=self._fonts["section"]).pack(fill='x', padx=12, pady=12)
        tab = ttk.Notebook(window)
        tab.pack(fill='both', expand=True, padx=10, pady=8)
        from tkinter.scrolledtext import ScrolledText
        panels = [('待新增', gui_business.format_preview(plan)),
                  ('异常提醒', '\n\n'.join(plan.warnings) or '本次无异常'),
                  ('普通提示', '\n\n'.join(plan.notices) or '本次无提示'),
                  ('已存在', '\n'.join(plan.skipped) or '无')]
        for label, content in panels:
            frame = ttk.Frame(tab)
            tab.add(frame, text=f'{label}' + (f' ({len(plan.warnings)})' if label=='异常提醒' else ''))
            detail = ScrolledText(frame, wrap='word', font=self._fonts["mono"])
            detail.pack(fill='both', expand=True)
            detail.insert('1.0', content)
            detail.configure(state='disabled')
        ttk.Label(window, text='此窗口只是预览，尚未修改 Excel。关闭此窗口不影响主界面。').pack(pady=(0,10))

    def _confirm_write(self):
        if self.business_busy or not self.business_plan:
            return
        plan = self.business_plan
        rename = self.rename_date_var.get()
        new_name = gui_business.dated_output_target(plan.target) if rename else plan.target
        if self.worker and self.worker.is_alive():
            messagebox.showwarning('不能写入', '扫描仍在运行。')
            return
        if self._preview_cfg != self._get_config():
            messagebox.showwarning('设置已改变', 'PDF 来源、业务员、Excel 路径或选项已改变，请重新扫描并预览。')
            self._invalidate_plan()
            return
        if not plan.has_changes:
            messagebox.showinfo('无需写入', '没有待新增的订单、发票或备注。')
            return
        details = (f'将修改以下文件：\n{plan.target}\n\n'
                   f'写入后文件名：{new_name.name}\n'
                   f'新增订单：{len(plan.orders)} 行\n新增发票：{len(plan.invoices)} 行\n备注写入：{len(plan.note_changes)} 处\n' 
                   f'需要复核的异常：{len(plan.warnings)} 条\n\n'
                   '备注会追加到已有备注末尾，不覆盖原内容；首次新行直接写备注。\n'
                   '将自动备份已有 Excel；不会覆盖人工记录，也不会修改 PDF 或文件夹。\n'
                   + ('文件名更新将保存为新文件，并在成功备份后移除原日期名称的旧文件。\n' if new_name != plan.target else '')
                   + '执行前请保存并关闭此 Excel。\n\n确定执行写入吗？')
        if not messagebox.askyesno('确认修改 Excel', details, icon='warning', parent=self):
            return
        if plan.warnings:
            if not messagebox.askyesno('异常提醒：再次确认',
                    f'当前仍有 {len(plan.warnings)} 条需复核项目。只有已通过安全校验的记录会加入。\n'
                    '不确定的记录会被跳过。仍要写入通过检查的记录吗？', icon='warning', parent=self):
                return
        self.business_busy = True
        self.write_button.configure(state='disabled')
        self.review_button.configure(state='disabled')
        self.status_var.set('正在按确认范围写入 Excel，请不要关闭程序…')
        def run():
            try:
                result = gui_business.apply_confirmed(plan, rename_date=rename)
                log_error = ''
                try:
                    gui_business.save_write_report(plan, *result)
                except Exception as exc:
                    log_error = str(exc)
                self.events.put(('business_written', (result[0], result[1], log_error,
                                                     [c.draft for c in plan.note_changes] + list(plan.note_completed), plan.note_profile)))
            except Exception:
                self.events.put(('business_error', traceback.format_exc()))
        self.business_worker = threading.Thread(target=run, daemon=True)
        self.business_worker.start()

    def _confirm_week_folder_rename(self):
        from datetime import date
        if self.business_busy or (self.worker and self.worker.is_alive()):
            messagebox.showwarning('任务进行中', '请先完成当前任务。')
            return
        if len(self.sources) != 1:
            messagebox.showinfo('不适用', '此操作仅用于一个来源根目录下包含「到DD.MM.YYYY」文件夹的旧版结构。多来源不自动改名。')
            return
        root = Path(self.sources[0]['path'])
        if not root.is_dir():
            messagebox.showerror('目录不存在', str(root))
            return
        try:
            summary = legacy.update_week_dir(root, date.today(), dry_run=True)
        except Exception as exc:
            messagebox.showerror('预览失败', str(exc))
            return
        if not summary.startswith('将重命名'):
            messagebox.showinfo('无需改名 / 需人工检查', summary)
            return
        if not messagebox.askyesno('确认修改文件夹名称', f'将修改此目录中的文件夹名称：\n{root}\n\n{summary}\n\n'
                                   '不会修改里面的 PDF，但 Windows 访问路径可能随之变化。确定执行吗？',
                                   icon='warning', parent=self):
            return
        try:
            result = legacy.update_week_dir(root, date.today(), dry_run=False)
            messagebox.showinfo('完成', result)
            self._invalidate_plan()
            self.scanned_complete = False
            self.review_button.configure(state='disabled')
        except Exception as exc:
            messagebox.showerror('改名失败', str(exc))

    def _warn_legacy_schedule(self):
        if gui_schedule.status() == 'legacy_write':
            if messagebox.askyesno('检测到旧版无人值守写入任务',
                    '旧版定时任务仍会自动写入 Excel，不符合当前必须确认的安全要求。\n\n'
                    '是否立即停用旧版任务？', icon='warning', parent=self):
                try:
                    gui_schedule.disable()
                    messagebox.showinfo('已停用', '旧版无人值守写入任务已停用。')
                except Exception as exc:
                    messagebox.showerror('停用失败', str(exc))

    def _schedule_settings(self):
        state = gui_schedule.status()
        if state == 'legacy_write':
            messagebox.showwarning('检测到旧版自动写入任务',
                '检测到旧版定时任务会自动写入 Excel！建议先停用。新版仅允许定时只读扫描，不会自动写入。')
        if state in ('enabled','legacy_write'):
            if not messagebox.askyesno('确认关闭定时任务',
                   '确定要停用每天 16:45 的定时检查吗？', parent=self):
                return
            try:
                gui_schedule.disable()
                messagebox.showinfo('已停用', 'Windows 定时任务已停用。')
            except Exception as exc:
                messagebox.showerror('停用失败', str(exc))
        else:
            if not self.sources:
                messagebox.showwarning('未配置', '请先设置 PDF 来源目录。')
                return
            if not messagebox.askyesno('确认开启自动定时检查',
                   '是否开启每工作日 16:45 的自动检查？\n\n'
                   '仅自动扫描 PDF 并生成只读业务预览。\n'
                   '绝不自动写入 Excel、修改文件名或启用无人值守写入。\n'
                   '写入仍需你打开 GUI 后逐次确认。\n\n'
                   '此操作会替换同名的旧版定时任务。', icon='warning', parent=self):
                return
            try:
                gui_schedule.enable_preview_only()
                messagebox.showinfo('已启用', '已开启工作日 16:45 的只读检查。')
            except Exception as exc:
                messagebox.showerror('启用失败', str(exc))

    def _on_close(self):
        if self.business_busy:
            messagebox.showwarning('任务执行中', 'Excel 操作尚未完成，请勿直接关闭窗口。')
            return
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("扫描进行中", "扫描尚未结束。确定关闭窗口吗？\n当前扫描将中断，PDF 不会被修改。"):
                return
            self.cancel_event.set()
        self._save_layout_preferences()
        self._save_config(silent=True)
        self.destroy()


if __name__ == "__main__":
    HangchaGUI().mainloop()
