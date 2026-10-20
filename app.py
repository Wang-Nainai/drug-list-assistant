"""药单整理助手 —— Win11 Fluent 风格的药品清单去重工具。"""

from __future__ import annotations

import ctypes
import json
import logging
import os
import queue
import sys
import tempfile
import threading
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path
from tkinter import filedialog, ttk
import tkinter as tk

import customtkinter as ctk
from PIL import Image

import core
import ui_theme as T
from ui_widgets import (
    AppDialog,
    ColumnMapDialog,
    DropZone,
    ModeCard,
    NameReviewDialog,
    StatCard,
)

try:
    from tkinterdnd2 import TkinterDnD

    class _DnDRoot(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *args, **kwargs):  # noqa: ANN002, ANN003
            super().__init__(*args, **kwargs)
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception:  # noqa: BLE001
                self.TkdndVersion = None

except Exception:  # noqa: BLE001 - 没有 tkinterdnd2 时仍可点击选择文件
    _DnDRoot = ctk.CTk  # type: ignore[assignment,misc]


APP_NAME = "药单整理助手"
APP_VERSION = "1.1.1"
APPEARANCE_LABELS = {"light": "浅色", "dark": "深色", "system": "跟随系统"}
LABEL_TO_APPEARANCE = {value: key for key, value in APPEARANCE_LABELS.items()}
PREVIEW_SORT_LABELS = {
    "first": "首次出现",
    "count_desc": "重复次数↓",
    "name": "药名",
}
PREVIEW_LABEL_TO_SORT = {value: key for key, value in PREVIEW_SORT_LABELS.items()}

logger = logging.getLogger("drug_list_assistant")


def _appdata_dir() -> Path:
    base = os.getenv("APPDATA") or str(Path.home())
    return Path(base) / "DrugListAssistant"


def _config_file() -> Path:
    return _appdata_dir() / "settings.json"


def load_config() -> dict:
    defaults = {
        "appearance": "system",
        "mode": core.MODE_SIMPLIFY,
        "open_after": False,
        "last_dir": "",
        "sort_order": "first",
        "batch_all_sheets": False,
    }
    try:
        data = json.loads(_config_file().read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return defaults
        appearance = data.get("appearance")
        if appearance in APPEARANCE_LABELS:
            defaults["appearance"] = appearance
        mode = data.get("mode")
        if mode in (core.MODE_SIMPLIFY, core.MODE_COUNT):
            defaults["mode"] = mode
        defaults["open_after"] = bool(data.get("open_after", False))
        last_dir = data.get("last_dir")
        if isinstance(last_dir, str):
            defaults["last_dir"] = last_dir
        sort_order = data.get("sort_order")
        if sort_order in core.SORT_ORDER_LABELS:
            defaults["sort_order"] = sort_order
        defaults["batch_all_sheets"] = bool(data.get("batch_all_sheets", False))
    except Exception:  # noqa: BLE001
        pass
    return defaults


def save_config(data: dict) -> None:
    try:
        directory = _appdata_dir()
        directory.mkdir(parents=True, exist_ok=True)
        target = _config_file()
        temporary = target.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temporary, target)
    except Exception:  # noqa: BLE001
        logger.debug("无法保存设置", exc_info=True)


def setup_logging() -> None:
    try:
        base = Path(os.getenv("LOCALAPPDATA") or tempfile.gettempdir())
        log_dir = base / "DrugListAssistant"
        log_dir.mkdir(parents=True, exist_ok=True)
        handler = RotatingFileHandler(
            log_dir / "app.log",
            maxBytes=1_000_000,
            backupCount=2,
            encoding="utf-8",
        )
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        )
        logger.setLevel(logging.INFO)
        logger.addHandler(handler)
    except Exception:  # noqa: BLE001
        logger.addHandler(logging.NullHandler())


def _load_image(path: str) -> Image.Image:
    with Image.open(path) as image:
        return image.copy().convert("RGBA")


class App(_DnDRoot):
    def __init__(self, config: dict, initial_path: str | None = None) -> None:
        self._config = config
        self.mode = config.get("mode", core.MODE_SIMPLIFY)
        self.appearance_key = config.get("appearance", "system")
        ctk.set_appearance_mode(self.appearance_key)

        super().__init__(fg_color=T.BG)
        T.detect_fonts(self)
        logger.info("启动 %s v%s，TkDND=%s", APP_NAME, APP_VERSION, getattr(self, "TkdndVersion", None))

        self._closing = False
        self.busy = False
        self.current_file: Path | None = None
        self.workbook: core.SheetData | None = None
        self.result: core.AnalysisResult | None = None
        self.output_path: Path | None = None
        self.output_dir: Path | None = None
        self.output_dir_custom = False
        self.last_dir = config.get("last_dir") or ""
        self.sort_order = config.get("sort_order", "first")
        if self.sort_order not in core.SORT_ORDER_LABELS:
            self.sort_order = "first"
        self.batch_all_sheets = bool(config.get("batch_all_sheets", False))
        self.name_aliases: dict[str, str] = {}
        self._mapping_prompted_for: tuple[str, str] | None = None
        self._mapping_prompt_after: str | None = None
        self._preview_filter_after: str | None = None
        self._preview_mode: str | None = None
        self._dialog: AppDialog | None = None
        self._initial_path = initial_path
        self._job_callbacks: tuple | None = None
        self._job_kind = ""
        self._job_queue: queue.Queue = queue.Queue()
        self._poll_after: str | None = None
        self._watch_after: str | None = None
        self._resize_after: str | None = None
        self._show_after: str | None = None
        self._restore_guard_ready = False
        self._last_resolved_mode = ctk.get_appearance_mode()
        try:
            self._last_scale = ctk.ScalingTracker.get_widget_scaling(self)
        except Exception:  # noqa: BLE001
            self._last_scale = 1.0

        self.title(APP_NAME)
        try:
            self.iconbitmap(T.resource_path("assets", "app.ico"))
        except Exception:  # noqa: BLE001
            pass

        self._set_initial_geometry()
        self._build_ui()
        self._apply_mode(self.mode, reset=False)
        self._set_status("就绪 · 选择 Excel 文件开始", "neutral")
        self._style_table()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Unmap>", self._on_window_unmap, add="+")
        self.bind("<Map>", self._on_window_map, add="+")
        self.bind_all("<Control-o>", lambda _e: self._choose_file())
        self.bind_all("<Control-O>", lambda _e: self._choose_file())
        self.bind_all("<Control-Return>", lambda _e: self._start_processing())
        self.report_callback_exception = self._report_callback_exception

        self._watch_ui()
        self.after(300, self._startup)
        self.after(1200, self._enable_restore_guard)

    # ------------------------------------------------------------- setup ----
    def _set_initial_geometry(self) -> None:
        # Tk 在 DPI 感知模式下返回的屏幕尺寸已经是逻辑像素，无需再除以缩放比。
        screen_w = max(self.winfo_screenwidth(), 800)
        screen_h = max(self.winfo_screenheight(), 600)

        width = min(1120, max(960, int(screen_w - 60)))
        height = min(800, max(700, int(screen_h - 100)))
        width = min(width, max(900, int(screen_w - 20)))
        height = min(height, max(680, int(screen_h - 60)))
        x = max(0, int((screen_w - width) / 2))
        y = max(0, int((screen_h - height) / 2) - 8)

        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(min(1000, width), min(800, height))

    def _enable_restore_guard(self) -> None:
        if not sys.platform.startswith("win"):
            return
        self._restore_guard_ready = True
        try:
            self.attributes("-alpha", 1.0)
        except tk.TclError:
            pass

    def _on_window_unmap(self, event: tk.Event) -> None:
        """最小化后先隐藏窗口，避免恢复首帧露出未重绘的黑色表面。"""
        if (
            event.widget is not self
            or not self._restore_guard_ready
            or not sys.platform.startswith("win")
        ):
            return
        try:
            if self.state() == "iconic":
                self.attributes("-alpha", 0.0)
        except tk.TclError:
            pass

    def _on_window_map(self, event: tk.Event) -> None:
        if (
            event.widget is not self
            or not self._restore_guard_ready
            or not sys.platform.startswith("win")
        ):
            return
        try:
            if float(self.attributes("-alpha")) > 0.01:
                return
            if self._show_after:
                self.after_cancel(self._show_after)
            self.update_idletasks()
            self._show_after = self.after(120, self._show_window_again)
        except (tk.TclError, ValueError):
            pass

    def _show_window_again(self) -> None:
        self._show_after = None
        try:
            if self.winfo_exists():
                self.update_idletasks()
                self.attributes("-alpha", 1.0)
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_body()

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=28, pady=(20, 16))
        header.grid_columnconfigure(1, weight=1)

        logo_file = T.resource_path("assets", "logo.png")
        self._logo_image = ctk.CTkImage(
            light_image=_load_image(logo_file),
            dark_image=_load_image(logo_file),
            size=(44, 44),
        )
        logo = ctk.CTkLabel(header, text="", image=self._logo_image, width=44, height=44)
        logo.grid(row=0, column=0, rowspan=2, padx=(0, 12))

        ctk.CTkLabel(
            header,
            text=APP_NAME,
            font=T.font(20, "bold"),
            text_color=T.TEXT,
            anchor="w",
        ).grid(row=0, column=1, sticky="w")

        subtitle_row = ctk.CTkFrame(header, fg_color="transparent")
        subtitle_row.grid(row=1, column=1, sticky="w", pady=(2, 0))
        ctk.CTkLabel(
            subtitle_row,
            text="Excel 药品清单 · 去重合并与重复统计",
            font=T.font(9.4),
            text_color=T.TEXT_SECONDARY,
        ).pack(side="left")
        ctk.CTkLabel(
            subtitle_row,
            text="●",
            font=T.font(6.5),
            text_color=T.SUCCESS,
        ).pack(side="left", padx=(12, 5))
        ctk.CTkLabel(
            subtitle_row,
            text="本地处理，数据不上传",
            font=T.font(9.4),
            text_color=T.TEXT_SECONDARY,
        ).pack(side="left")

        self.appearance_var = tk.StringVar(value=APPEARANCE_LABELS[self.appearance_key])
        self.appearance_segment = ctk.CTkSegmentedButton(
            header,
            values=list(APPEARANCE_LABELS.values()),
            variable=self.appearance_var,
            command=self._on_appearance_changed,
            width=236,
            height=32,
            corner_radius=8,
            border_width=0,
            fg_color=T.SEGMENT_BG,
            selected_color=T.SEGMENT_SELECTED,
            selected_hover_color=T.SEGMENT_SELECTED,
            unselected_color=T.SEGMENT_BG,
            unselected_hover_color=T.SURFACE_HOVER,
            text_color=T.TEXT,
            font=T.font(9.2),
        )
        self.appearance_segment.grid(row=0, column=2, rowspan=2, sticky="e")
        self.appearance_segment.set(APPEARANCE_LABELS[self.appearance_key])

    def _build_body(self) -> None:
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=28, pady=(0, 18))
        body.grid_columnconfigure(0, minsize=372)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(3, weight=1)

        self._build_mode_card(left)
        self._build_file_card(left)
        self._build_settings_card(left)
        self._build_action_area(left)

        right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        self._build_stats(right)
        self._build_preview(right)
        self._build_status_row(right)

    def _card(self, parent) -> ctk.CTkFrame:
        return ctk.CTkFrame(
            parent,
            fg_color=T.SURFACE,
            corner_radius=12,
            border_width=1,
            border_color=T.BORDER,
        )

    def _section_header(self, parent, title: str, hint: str = "") -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            frame,
            text=title,
            font=T.font(12.2, "bold"),
            text_color=T.TEXT,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        if hint:
            ctk.CTkLabel(
                frame,
                text=hint,
                font=T.font(8.6),
                text_color=T.TEXT_TERTIARY,
                anchor="e",
            ).grid(row=0, column=1, sticky="e")
        return frame

    def _build_mode_card(self, parent) -> None:
        card = self._card(parent)
        card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        card.grid_columnconfigure(0, weight=1)

        header = self._section_header(card, "处理方式", "可随时切换")
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 9))

        self.mode_simplify_card = ModeCard(
            card,
            value=core.MODE_SIMPLIFY,
            title="精简合并",
            description="同名药品合并为一条，规格与序号汇总",
            badge="合",
            command=self._on_mode_clicked,
        )
        self.mode_simplify_card.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))

        self.mode_count_card = ModeCard(
            card,
            value=core.MODE_COUNT,
            title="精确统计",
            description="按「药品名称 + 规格」精确去重统计",
            badge="计",
            command=self._on_mode_clicked,
        )
        self.mode_count_card.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 14))

    def _build_file_card(self, parent) -> None:
        card = self._card(parent)
        card.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        card.grid_columnconfigure(0, weight=1)

        header = self._section_header(card, "Excel 文件", "支持拖拽")
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 9))
        self.batch_all_sheets_var = tk.BooleanVar(value=self.batch_all_sheets)
        self.batch_all_sheets_check = ctk.CTkCheckBox(
            header,
            text="批量全表",
            variable=self.batch_all_sheets_var,
            command=self._on_batch_all_sheets_changed,
            font=T.font(8.2),
            text_color=T.TEXT_TERTIARY,
            fg_color=T.ACCENT,
            hover_color=T.ACCENT_HOVER,
            border_color=T.BORDER_STRONG,
            checkbox_width=15,
            checkbox_height=15,
            corner_radius=4,
        )
        self.batch_all_sheets_check.grid(row=0, column=2, sticky="e", padx=(10, 0))

        self.dropzone = DropZone(
            card,
            on_pick=self._choose_file,
            on_drop=self._handle_drop,
            height=104,
        )
        self.dropzone.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 14))

        self.sheet_row = ctk.CTkFrame(card, fg_color="transparent")
        self.sheet_row.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 14))
        self.sheet_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            self.sheet_row,
            text="工作表",
            font=T.font(9.2),
            text_color=T.TEXT_SECONDARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))

        self.sheet_var = tk.StringVar(value="")
        self.sheet_menu = ctk.CTkOptionMenu(
            self.sheet_row,
            variable=self.sheet_var,
            values=[""],
            command=self._on_sheet_selected,
            width=220,
            height=30,
            corner_radius=8,
            fg_color=T.SURFACE_ALT,
            button_color=T.SURFACE_ALT,
            button_hover_color=T.TRACK,
            text_color=T.TEXT,
            font=T.font(9.2),
            dropdown_fg_color=T.SURFACE,
            dropdown_hover_color=T.ACCENT_SOFT,
            dropdown_text_color=T.TEXT,
            dropdown_font=T.font(9.2),
            dynamic_resizing=False,
        )
        self.sheet_menu.grid(row=0, column=1, sticky="e")
        self.sheet_row.grid_remove()

        tools = ctk.CTkFrame(card, fg_color="transparent")
        tools.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 14))
        tools.grid_columnconfigure((0, 1, 2), weight=1)

        self.mapping_btn = ctk.CTkButton(
            tools,
            text="列映射",
            width=92,
            height=28,
            corner_radius=7,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            text_color_disabled=T.TEXT_TERTIARY,
            font=T.font(8.8),
            command=self._open_column_mapping,
            state="disabled",
        )
        self.mapping_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.name_review_btn = ctk.CTkButton(
            tools,
            text="药名核对",
            width=92,
            height=28,
            corner_radius=7,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            text_color_disabled=T.TEXT_TERTIARY,
            font=T.font(8.8),
            command=self._open_name_review,
            state="disabled",
        )
        self.name_review_btn.grid(row=0, column=1, sticky="ew", padx=4)

        self.batch_btn = ctk.CTkButton(
            tools,
            text="批量文件",
            width=92,
            height=28,
            corner_radius=7,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            text_color_disabled=T.TEXT_TERTIARY,
            font=T.font(8.8),
            command=self._choose_batch_files,
            state="disabled",
        )
        self.batch_btn.grid(row=0, column=2, sticky="ew", padx=(4, 0))

    def _build_settings_card(self, parent) -> None:
        card = self._card(parent)
        card.grid(row=2, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        header = self._section_header(card, "导出设置")
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        header.grid_columnconfigure(1, weight=0)
        ctk.CTkLabel(
            header,
            text="排序",
            font=T.font(8.4),
            text_color=T.TEXT_TERTIARY,
        ).grid(row=0, column=1, sticky="e", padx=(0, 6))
        self.sort_order_var = tk.StringVar(value=core.SORT_ORDER_LABELS[self.sort_order])
        self.sort_menu = ctk.CTkOptionMenu(
            header,
            variable=self.sort_order_var,
            values=list(core.SORT_ORDER_LABELS.values()),
            command=self._on_sort_order_changed,
            width=112,
            height=26,
            corner_radius=7,
            fg_color=T.SURFACE_ALT,
            button_color=T.SURFACE_ALT,
            button_hover_color=T.TRACK,
            text_color=T.TEXT,
            font=T.font(8.2),
            dropdown_fg_color=T.SURFACE,
            dropdown_hover_color=T.ACCENT_SOFT,
            dropdown_text_color=T.TEXT,
            dropdown_font=T.font(8.4),
            dynamic_resizing=False,
        )
        self.sort_menu.grid(row=0, column=2, sticky="e")

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 14))
        row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            row,
            text=T.ICON_FOLDER,
            font=T.icon_font(13),
            text_color=T.ACCENT,
            width=22,
        ).grid(row=0, column=0, padx=(0, 6))

        self.output_label = ctk.CTkLabel(
            row,
            text="与源文件相同的文件夹",
            font=T.font(8.8),
            text_color=T.TEXT_SECONDARY,
            anchor="w",
        )
        self.output_label.grid(row=0, column=1, sticky="ew")

        self.change_output_btn = ctk.CTkButton(
            row,
            text="更改",
            width=58,
            height=28,
            corner_radius=7,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            font=T.font(9.0),
            command=self._choose_output_dir,
        )
        self.change_output_btn.grid(row=0, column=2, padx=(10, 0))

    def _build_action_area(self, parent) -> None:
        divider = ctk.CTkFrame(parent, fg_color=T.BORDER, height=1, corner_radius=0)
        divider.grid(row=4, column=0, sticky="ew", pady=(0, 0))
        divider.grid_propagate(False)

        self.open_after_var = tk.BooleanVar(value=bool(self._config.get("open_after", False)))
        self.open_after_switch = ctk.CTkSwitch(
            parent,
            text="完成后自动打开结果",
            variable=self.open_after_var,
            onvalue=True,
            offvalue=False,
            command=self._save_settings,
            font=T.font(9.6),
            text_color=T.TEXT_SECONDARY,
            progress_color=T.ACCENT,
            fg_color=T.TRACK,
            button_color=("#FFFFFF", "#F2F2F2"),
            button_hover_color=("#FFFFFF", "#FFFFFF"),
            switch_width=36,
            switch_height=18,
            corner_radius=9,
        )
        self.open_after_switch.grid(row=5, column=0, sticky="w", pady=(12, 8))

        self.primary_button = ctk.CTkButton(
            parent,
            text="开始处理",
            height=46,
            corner_radius=10,
            fg_color=T.ACCENT,
            hover_color=T.ACCENT_HOVER,
            text_color=T.ON_ACCENT,
            text_color_disabled=T.DISABLED_TEXT,
            font=T.font(12.2, "bold"),
            command=self._start_processing,
            state="disabled",
        )
        self.primary_button.grid(row=6, column=0, sticky="ew")

        footer = ctk.CTkFrame(parent, fg_color="transparent")
        footer.grid(row=7, column=0, sticky="ew", pady=(10, 0))
        footer.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            footer,
            text="●",
            font=T.font(6.5),
            text_color=T.SUCCESS,
            width=12,
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            footer,
            text="本地处理 · 文件不会上传",
            font=T.font(8.5),
            text_color=T.TEXT_TERTIARY,
            anchor="w",
        ).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(
            footer,
            text=f"v{APP_VERSION}",
            font=T.font(8.3),
            text_color=T.TEXT_TERTIARY,
            anchor="e",
        ).grid(row=0, column=2, sticky="e")

    def _build_stats(self, parent) -> None:
        stats = ctk.CTkFrame(parent, fg_color="transparent")
        stats.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        for index in range(4):
            stats.grid_columnconfigure(index, weight=1, uniform="stats")

        self.stat_input = StatCard(stats, "原始行数")
        self.stat_output = StatCard(stats, "结果条数")
        self.stat_duplicates = StatCard(stats, "重复组")
        self.stat_removed = StatCard(stats, "可减少行")
        for index, card in enumerate(
            (self.stat_input, self.stat_output, self.stat_duplicates, self.stat_removed)
        ):
            card.grid(
                row=0,
                column=index,
                sticky="nsew",
                padx=(0 if index == 0 else 6, 0 if index == 3 else 6),
            )

    def _build_preview(self, parent) -> None:
        self.preview_card = self._card(parent)
        self.preview_card.grid(row=1, column=0, sticky="nsew")
        self.preview_card.grid_columnconfigure(0, weight=1)
        self.preview_card.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(13, 10))
        header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            header,
            text="结果预览",
            font=T.font(12.2, "bold"),
            text_color=T.TEXT,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        self.preview_hint = ctk.CTkLabel(
            header,
            text="等待处理",
            font=T.font(8.6),
            text_color=T.TEXT_TERTIARY,
            anchor="w",
        )
        self.preview_hint.grid(row=0, column=1, sticky="w", padx=(10, 0))

        self.open_file_btn = ctk.CTkButton(
            header,
            text="打开文件",
            width=82,
            height=29,
            corner_radius=7,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            text_color_disabled=T.TEXT_TERTIARY,
            font=T.font(9.0),
            command=self._open_output_file,
            state="disabled",
        )
        self.open_file_btn.grid(row=0, column=2, padx=(8, 6))

        self.open_folder_btn = ctk.CTkButton(
            header,
            text="打开文件夹",
            width=92,
            height=29,
            corner_radius=7,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            text_color_disabled=T.TEXT_TERTIARY,
            font=T.font(9.0),
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_btn.grid(row=0, column=3, padx=(0, 0))

        divider = ctk.CTkFrame(
            self.preview_card,
            fg_color=T.BORDER,
            height=1,
            corner_radius=0,
        )
        divider.grid(row=1, column=0, sticky="ew")
        divider.grid_propagate(False)

        content = ctk.CTkFrame(
            self.preview_card,
            fg_color=T.SURFACE,
            corner_radius=0,
        )
        content.grid(row=2, column=0, sticky="nsew", padx=10, pady=(8, 6))
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        self.preview_toolbar = ctk.CTkFrame(content, fg_color=T.SURFACE, corner_radius=0)
        self.preview_toolbar.grid(row=0, column=0, sticky="ew", padx=6, pady=(0, 6))
        self.preview_toolbar.grid_columnconfigure(1, weight=1)

        self.preview_search_var = tk.StringVar()
        self.preview_search = ctk.CTkEntry(
            self.preview_toolbar,
            textvariable=self.preview_search_var,
            placeholder_text="搜索药名 / 规格",
            width=190,
            height=28,
            corner_radius=7,
            border_width=1,
            border_color=T.BORDER,
            fg_color=T.SURFACE_ALT,
            text_color=T.TEXT,
            placeholder_text_color=T.TEXT_TERTIARY,
            font=T.font(8.8),
        )
        self.preview_search.grid(row=0, column=0, sticky="w")
        self.preview_search_var.trace_add(
            "write", lambda *_args: self._schedule_preview_filter()
        )

        self.preview_only_dup_var = tk.BooleanVar(value=False)
        self.preview_only_dup = ctk.CTkCheckBox(
            self.preview_toolbar,
            text="仅重复",
            variable=self.preview_only_dup_var,
            command=self._refresh_preview,
            font=T.font(8.8),
            text_color=T.TEXT_SECONDARY,
            fg_color=T.ACCENT,
            hover_color=T.ACCENT_HOVER,
            border_color=T.BORDER_STRONG,
            checkbox_width=17,
            checkbox_height=17,
            corner_radius=4,
        )
        self.preview_only_dup.grid(row=0, column=2, sticky="w", padx=(12, 8))

        self.preview_sort_var = tk.StringVar(value=PREVIEW_SORT_LABELS["first"])
        self.preview_sort = ctk.CTkOptionMenu(
            self.preview_toolbar,
            variable=self.preview_sort_var,
            values=list(PREVIEW_LABEL_TO_SORT.keys()),
            command=self._on_preview_sort_changed,
            width=110,
            height=28,
            corner_radius=7,
            fg_color=T.SURFACE_ALT,
            button_color=T.SURFACE_ALT,
            button_hover_color=T.TRACK,
            text_color=T.TEXT,
            font=T.font(8.6),
            dropdown_fg_color=T.SURFACE,
            dropdown_hover_color=T.ACCENT_SOFT,
            dropdown_text_color=T.TEXT,
            dropdown_font=T.font(8.6),
            dynamic_resizing=False,
        )
        self.preview_sort.grid(row=0, column=3, sticky="e")
        self.preview_toolbar.grid_remove()

        self.table_frame = ctk.CTkFrame(content, fg_color=T.SURFACE, corner_radius=0)
        self.table_frame.grid(row=1, column=0, sticky="nsew")
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self.table_frame,
            columns=[],
            show="headings",
            selectmode="browse",
            style="App.Treeview",
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        self.v_scroll = ctk.CTkScrollbar(
            self.table_frame,
            orientation="vertical",
            command=self.tree.yview,
            width=12,
            fg_color="transparent",
            button_color=T.SCROLLBAR,
            button_hover_color=T.SCROLLBAR_HOVER,
        )
        self.v_scroll.grid(row=0, column=1, sticky="ns", padx=(2, 0))
        self.h_scroll = ctk.CTkScrollbar(
            self.table_frame,
            orientation="horizontal",
            command=self.tree.xview,
            height=12,
            fg_color="transparent",
            button_color=T.SCROLLBAR,
            button_hover_color=T.SCROLLBAR_HOVER,
        )
        self.h_scroll.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        self.tree.configure(
            yscrollcommand=self.v_scroll.set,
            xscrollcommand=self.h_scroll.set,
        )
        self.table_frame.bind("<Configure>", self._on_table_configure, add="+")
        self.table_frame.grid_remove()

        self.empty_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.empty_frame.place(relx=0.5, rely=0.48, anchor="center")

        empty_image_path = T.resource_path("assets", "empty_light.png")
        empty_dark_path = T.resource_path("assets", "empty_dark.png")
        self._empty_image = ctk.CTkImage(
            light_image=_load_image(empty_image_path),
            dark_image=_load_image(empty_dark_path),
            size=(208, 136),
        )
        ctk.CTkLabel(
            self.empty_frame,
            text="",
            image=self._empty_image,
        ).pack(pady=(0, 4))

        self.empty_title = ctk.CTkLabel(
            self.empty_frame,
            text="还没有结果",
            font=T.font(13, "bold"),
            text_color=T.TEXT,
        )
        self.empty_title.pack()

        self.empty_desc = ctk.CTkLabel(
            self.empty_frame,
            text="选择文件并点击「开始处理」",
            font=T.font(9.6),
            text_color=T.TEXT_SECONDARY,
            justify="center",
            wraplength=360,
        )
        self.empty_desc.pack(pady=(3, 0))

        self.progress = ctk.CTkProgressBar(
            self.preview_card,
            height=3,
            corner_radius=2,
            mode="indeterminate",
            fg_color=T.TRACK,
            progress_color=T.ACCENT,
        )
        self.progress.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 10))
        self.progress.grid_remove()

    def _build_status_row(self, parent) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.grid(row=2, column=0, sticky="ew", pady=(9, 0))
        row.grid_columnconfigure(1, weight=1)

        self.status_dot = ctk.CTkLabel(
            row,
            text="●",
            font=T.font(7),
            text_color=T.TEXT_TERTIARY,
            width=14,
        )
        self.status_dot.grid(row=0, column=0, sticky="w")

        self.status_label = ctk.CTkLabel(
            row,
            text="就绪",
            font=T.font(9.0),
            text_color=T.TEXT_SECONDARY,
            anchor="w",
        )
        self.status_label.grid(row=0, column=1, sticky="ew")

        ctk.CTkLabel(
            row,
            text="Ctrl+O 选择文件 · Ctrl+Enter 开始",
            font=T.font(8.2),
            text_color=T.TEXT_TERTIARY,
            anchor="e",
        ).grid(row=0, column=2, sticky="e", padx=(10, 0))

    # ------------------------------------------------------------ table ----
    def _column_specs(self, mode: str) -> list[tuple[str, int, int, str, bool]]:
        if mode == core.MODE_SIMPLIFY:
            return [
                ("药品名称", 170, 100, "w", True),
                ("药品规格", 210, 120, "w", True),
                ("重复次数", 72, 60, "center", False),
                ("原始序号", 170, 100, "w", True),
            ]
        return [
            ("序号", 56, 48, "center", False),
            ("药品名称", 150, 100, "w", True),
            ("药品规格", 175, 110, "w", True),
            ("重复次数", 72, 60, "center", False),
            ("原来的序号位置", 145, 100, "w", True),
        ]

    def _style_table(self) -> None:
        try:
            scale = ctk.ScalingTracker.get_widget_scaling(self)
        except Exception:  # noqa: BLE001
            scale = 1.0
        scale = scale or 1.0

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:  # noqa: BLE001
            pass

        body_size = max(10, round(10 * scale))
        heading_size = max(9, round(9.2 * scale))
        row_height = max(28, round(34 * scale))
        pad_x = max(6, round(8 * scale))
        pad_y = max(5, round(7 * scale))

        style.configure(
            "App.Treeview",
            background=T.resolve(T.SURFACE),
            fieldbackground=T.resolve(T.SURFACE),
            foreground=T.resolve(T.TEXT),
            borderwidth=0,
            relief="flat",
            font=(T.FONT_FAMILY, -body_size),
            rowheight=row_height,
        )
        style.configure(
            "App.Treeview.Heading",
            background=T.resolve(T.TABLE_HEADER_BG),
            foreground=T.resolve(T.TABLE_HEADER_TEXT),
            borderwidth=0,
            relief="flat",
            font=(T.FONT_FAMILY, -heading_size, "bold"),
            padding=(pad_x, pad_y),
        )
        style.map(
            "App.Treeview.Heading",
            background=[("active", T.resolve(T.TABLE_HEADER_BG))],
            relief=[("active", "flat")],
        )
        style.map(
            "App.Treeview",
            background=[("selected", T.resolve(T.TABLE_SELECTION))],
            foreground=[("selected", T.resolve(T.TEXT))],
        )
        try:
            style.layout(
                "App.Treeview",
                [("App.Treeview.treearea", {"sticky": "nswe"})],
            )
        except Exception:  # noqa: BLE001
            pass

        try:
            self.tree.tag_configure("odd", background=T.resolve(T.TABLE_ROW_ODD))
            self.tree.tag_configure("even", background=T.resolve(T.TABLE_ROW_EVEN))
            self.tree.configure(style="App.Treeview")
            if self._preview_mode:
                self._configure_table_columns(scale, self._preview_mode)
                self.after_idle(self._update_scrollbars)
        except Exception:  # noqa: BLE001
            logger.debug("表格样式更新失败", exc_info=True)

    def _configure_table_columns(self, scale: float, mode: str) -> None:
        specs = self._column_specs(mode)
        self.tree.configure(columns=[spec[0] for spec in specs])
        for name, width, min_width, anchor, stretch in specs:
            self.tree.heading(name, text=name, anchor=anchor)
            self.tree.column(
                name,
                width=max(30, round(width * scale)),
                minwidth=max(24, round(min_width * scale)),
                anchor=anchor,
                stretch=stretch,
            )

    def _populate_preview(self, result: core.AnalysisResult) -> None:
        self._preview_mode = result.mode
        try:
            scale = ctk.ScalingTracker.get_widget_scaling(self)
        except Exception:  # noqa: BLE001
            scale = 1.0
        self._configure_table_columns(scale, result.mode)
        self.preview_sort_var.set(PREVIEW_SORT_LABELS.get(self.sort_order, "首次出现"))

        self.preview_toolbar.grid()
        self.empty_frame.place_forget()
        self.table_frame.grid()
        self._refresh_preview()

    def _schedule_preview_filter(self) -> None:
        if self._preview_filter_after is not None:
            try:
                self.after_cancel(self._preview_filter_after)
            except Exception:  # noqa: BLE001
                pass
        self._preview_filter_after = self.after(150, self._refresh_preview)

    def _refresh_preview(self) -> None:
        self._preview_filter_after = None
        if self.result is None:
            return
        df = self.result.df.copy(deep=False)
        if self.preview_only_dup_var.get() and "重复次数" in df.columns:
            df = df[df["重复次数"] > 1]

        query = self.preview_search_var.get().strip().casefold()
        if query and len(df):
            mask = df.apply(
                lambda column: column.astype(str).str.casefold().str.contains(
                    query, regex=False, na=False
                )
            ).any(axis=1)
            df = df[mask]

        sort_mode = self.preview_sort_var.get()
        if sort_mode == "重复次数↓" and "重复次数" in df.columns:
            df = df.sort_values(by="重复次数", ascending=False, kind="stable")
        elif sort_mode == "药名" and "药品名称" in df.columns:
            df = df.sort_values(
                by="药品名称",
                key=lambda series: series.astype(str).str.casefold(),
                kind="stable",
            )

        self.tree.delete(*self.tree.get_children())
        for index, row in enumerate(df.head(300).itertuples(index=False, name=None)):
            values = [T.ellipsize_units(str(value), 240, "end") for value in row]
            tag = "even" if index % 2 == 0 else "odd"
            self.tree.insert("", "end", values=values, tags=(tag,))

        total = self.result.output_rows
        visible = len(df)
        if visible != total or query or self.preview_only_dup_var.get():
            suffix = " · 预览前 300 行" if visible > 300 else ""
            self.preview_hint.configure(text=f"显示 {visible:,} / 共 {total:,} 条{suffix}")
        elif total > 300:
            self.preview_hint.configure(text=f"共 {total:,} 条 · 预览前 300 行")
        else:
            self.preview_hint.configure(text=f"共 {total:,} 条")
        self.tree.yview_moveto(0)
        self.tree.xview_moveto(0)
        self.after_idle(self._update_scrollbars)

    def _on_table_configure(self, _event=None) -> None:
        if self._resize_after is not None:
            try:
                self.after_cancel(self._resize_after)
            except Exception:  # noqa: BLE001
                pass
        self._resize_after = self.after(80, self._update_scrollbars)

    def _update_scrollbars(self) -> None:
        self._resize_after = None
        if self._closing:
            return
        try:
            if not self.table_frame.winfo_ismapped():
                return
            self.update_idletasks()

            y_start, y_end = self.tree.yview()
            if y_start <= 0.001 and y_end >= 0.999:
                self.v_scroll.grid_remove()
            else:
                self.v_scroll.grid()

            self.update_idletasks()
            x_start, x_end = self.tree.xview()
            if x_start <= 0.001 and x_end >= 0.999:
                self.h_scroll.grid_remove()
            else:
                self.h_scroll.grid()
        except Exception:  # noqa: BLE001
            logger.debug("更新滚动条失败", exc_info=True)

    def _show_empty_state(self, title: str, description: str) -> None:
        try:
            self.table_frame.grid_remove()
        except Exception:  # noqa: BLE001
            pass
        self.empty_frame.place(relx=0.5, rely=0.48, anchor="center")
        self.empty_title.configure(text=title)
        self.empty_desc.configure(text=description)

    def _reset_result(self) -> None:
        self.result = None
        self.output_path = None
        self._preview_mode = None
        if self._preview_filter_after is not None:
            try:
                self.after_cancel(self._preview_filter_after)
            except Exception:  # noqa: BLE001
                pass
            self._preview_filter_after = None
        try:
            self.preview_search_var.set("")
            self.preview_only_dup_var.set(False)
            self._preview_filter_after = None
            self.preview_toolbar.grid_remove()
            self.tree.delete(*self.tree.get_children())
            self.tree.configure(columns=[])
        except Exception:  # noqa: BLE001
            pass
        self._show_empty_state("还没有结果", "选择文件并点击「开始处理」")
        self.preview_hint.configure(text="等待处理")
        self._set_open_buttons_enabled(False)
        self._set_stats_blank()
        if self.workbook is not None:
            total = self.workbook.total_rows
            self.stat_input.set_value(f"{total:,}" if total else "—", "default" if total else "muted")

    def _set_stats_blank(self) -> None:
        self.stat_input.set_value("—", "muted")
        self.stat_output.set_value("—", "muted")
        self.stat_duplicates.set_value("—", "muted")
        self.stat_removed.set_value("—", "muted")

    def _update_stats(self, result: core.AnalysisResult) -> None:
        self.stat_input.set_value(f"{result.total_rows:,}")
        self.stat_output.set_value(f"{result.output_rows:,}", "accent")
        self.stat_duplicates.set_value(
            f"{result.duplicate_groups:,}",
            "warning" if result.duplicate_groups else "success",
        )
        self.stat_removed.set_value(
            f"{result.removed_rows:,}",
            "success" if result.removed_rows else "default",
        )

    # ----------------------------------------------------------- states ----
    def _set_primary(self, state: str, text: str = "开始处理") -> None:
        try:
            if state == "normal":
                self.primary_button.configure(
                    state="normal",
                    text="开始处理",
                    fg_color=T.ACCENT,
                    hover_color=T.ACCENT_HOVER,
                    text_color=T.ON_ACCENT,
                )
            elif state == "busy":
                self.primary_button.configure(
                    state="disabled",
                    text=text,
                    fg_color=T.ACCENT_SOFT,
                    hover_color=T.ACCENT_SOFT,
                    text_color_disabled=T.ACCENT,
                )
            else:
                self.primary_button.configure(
                    state="disabled",
                    text="开始处理",
                    fg_color=T.DISABLED_FILL,
                    hover_color=T.DISABLED_FILL,
                    text_color_disabled=T.DISABLED_TEXT,
                )
        except Exception:  # noqa: BLE001
            pass

    def _set_open_buttons_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for button in (self.open_file_btn, self.open_folder_btn):
            try:
                button.configure(state=state)
            except Exception:  # noqa: BLE001
                pass

    def _set_busy(self, busy: bool, action: str = "", status_text: str = "") -> None:
        self.busy = busy
        self.dropzone.set_enabled(not busy)
        try:
            self.change_output_btn.configure(state="disabled" if busy else "normal")
            sheet_available = (
                self.workbook is not None and len(self.workbook.sheet_names) > 1
            )
            self.sheet_menu.configure(
                state="disabled" if busy or not sheet_available else "normal"
            )
        except Exception:  # noqa: BLE001
            pass
        self._set_tools_enabled(not busy)

        if busy:
            self.open_after_switch.configure(state="disabled")
            label = {
                "load": "正在读取…",
                "process": "正在处理…",
                "sheet": "正在切换…",
                "names": "正在检查药名…",
                "batch": "正在批量处理…",
            }.get(action, "正在处理…")
            self._set_primary("busy", label)
            try:
                self.progress.grid()
                self.progress.start()
            except Exception:  # noqa: BLE001
                pass
            if status_text:
                self._set_status(status_text, "busy")
        else:
            self.open_after_switch.configure(state="normal")
            usable = self.workbook is not None and self.workbook.usable
            self._set_primary("normal" if usable else "disabled")
            try:
                self.progress.stop()
                self.progress.set(0)
                self.progress.grid_remove()
            except Exception:  # noqa: BLE001
                pass

    def _set_status(self, text: str, tone: str = "neutral") -> None:
        colours = {
            "neutral": T.TEXT_SECONDARY,
            "busy": T.ACCENT,
            "success": T.SUCCESS,
            "error": T.DANGER,
            "warning": T.WARNING,
        }
        dot_colours = {
            "neutral": T.TEXT_TERTIARY,
            "busy": T.ACCENT,
            "success": T.SUCCESS,
            "error": T.DANGER,
            "warning": T.WARNING,
        }
        try:
            self.status_dot.configure(text_color=dot_colours.get(tone, T.TEXT_TERTIARY))
            self.status_label.configure(
                text=T.ellipsize_units(text, 78),
                text_color=colours.get(tone, T.TEXT_SECONDARY),
            )
        except Exception:  # noqa: BLE001
            pass

    # ------------------------------------------------------------- files ----
    def _choose_file(self) -> None:
        if self.busy:
            return
        kwargs = {
            "parent": self,
            "title": "选择药品清单",
            "filetypes": [
                ("Excel 工作簿", "*.xlsx *.xls *.xlsm"),
                ("Excel 2007 及以上", "*.xlsx *.xlsm"),
                ("Excel 97-2003", "*.xls"),
                ("所有文件", "*.*"),
            ],
        }
        if self.last_dir and Path(self.last_dir).is_dir():
            kwargs["initialdir"] = self.last_dir
        path = filedialog.askopenfilename(**kwargs)
        if path:
            self._load_file(path)

    def _handle_drop(self, paths: list[str]) -> None:
        if self.busy:
            return
        excel_paths = [
            path for path in paths if Path(path).suffix.lower() in core.SUPPORTED_EXTENSIONS
        ]
        if not excel_paths:
            self._show_dialog(
                "仅支持 Excel 文件",
                "请拖入 .xlsx、.xls 或 .xlsm 格式的药品清单。",
                "warning",
            )
            return
        self._load_file(excel_paths[0])

    def _load_file(self, file_path: str | os.PathLike[str]) -> None:
        if self.busy:
            return
        try:
            path = core.validate_path(file_path)
        except core.UserError as exc:
            self._show_dialog("无法打开文件", str(exc), "error")
            return

        self.current_file = path
        self.workbook = None
        self.name_aliases = {}
        self._mapping_prompted_for = None
        self._cancel_mapping_prompt()
        self.sheet_var.set("")
        self.sheet_row.grid_remove()
        self._set_tools_enabled(False)
        if not self.output_dir_custom:
            self.output_dir = path.parent
        self._update_output_label()
        self.last_dir = str(path.parent)
        self._save_settings()
        logger.info("读取文件：%s", path)

        self._reset_result()
        self._show_empty_state("正在读取文件", f"正在打开 {path.name}")
        self.dropzone.set_loading(path.name)
        self._set_busy(True, "load", "正在读取文件…")

        self._start_job(
            "load",
            lambda: core.load_workbook(path),
            self._on_workbook_loaded,
            self._on_job_error,
        )

    def _on_workbook_loaded(self, data: core.SheetData) -> None:
        self.workbook = data
        self._set_busy(False)
        self._set_tools_enabled(True)
        self._update_sheet_menu(data)
        self._update_file_card(data)
        self._reset_result()

        if data.usable:
            self._show_empty_state("还没有结果", "选择处理方式后点击「开始处理」")
            self._set_status(
                f"已读取 {data.total_rows:,} 行 · 表头第 {data.header_row + 1} 行 · "
                "选择处理方式后开始",
                "neutral",
            )
            self._set_primary("normal")
        else:
            self._show_empty_state(
                "当前工作表不可用",
                "请检查表头，或点击「列映射」手动指定列",
            )
            self._set_status(data.problem or "当前工作表不可用", "error")
            self._set_primary("disabled")
            self.stat_input.set_value(
                f"{data.total_rows:,}" if data.total_rows else "—",
                "muted",
            )
            self._prompt_column_mapping(data)

    def _set_tools_enabled(self, enabled: bool) -> None:
        if not hasattr(self, "mapping_btn"):
            return
        ready = bool(enabled and self.current_file is not None)
        usable = bool(ready and self.workbook is not None and self.workbook.usable)
        try:
            self.mapping_btn.configure(state="normal" if ready else "disabled")
            self.name_review_btn.configure(state="normal" if usable else "disabled")
            self.batch_btn.configure(state="normal" if ready else "disabled")
        except Exception:  # noqa: BLE001
            pass
        self._update_name_review_btn()

    def _update_name_review_btn(self) -> None:
        if not hasattr(self, "name_review_btn"):
            return
        targets = len(set(self.name_aliases.values()))
        text = f"药名核对 ({targets})" if targets else "药名核对"
        try:
            self.name_review_btn.configure(text=text)
        except Exception:  # noqa: BLE001
            pass

    def _cancel_mapping_prompt(self) -> None:
        if self._mapping_prompt_after is not None:
            try:
                self.after_cancel(self._mapping_prompt_after)
            except Exception:  # noqa: BLE001
                pass
            self._mapping_prompt_after = None

    def _prompt_column_mapping(self, data: core.SheetData) -> None:
        if not data.raw_columns or not data.missing_columns:
            return
        key = (str(data.path), data.sheet_name)
        if self._mapping_prompted_for == key:
            return
        self._mapping_prompted_for = key
        self._cancel_mapping_prompt()
        self._mapping_prompt_after = self.after(220, self._open_column_mapping)

    def _update_sheet_menu(self, data: core.SheetData) -> None:
        if len(data.sheet_names) > 1:
            self.sheet_menu.configure(values=data.sheet_names)
            self.sheet_var.set(data.sheet_name)
            self.sheet_row.grid()
            self.sheet_menu.configure(state="normal")
        else:
            self.sheet_row.grid_remove()
            self.sheet_var.set(data.sheet_name)
            self.sheet_menu.configure(state="disabled")

    def _update_file_card(self, data: core.SheetData) -> None:
        try:
            size = data.path.stat().st_size
        except OSError:
            size = 0
        if data.usable:
            self.dropzone.set_file(
                data.path.name,
                hint=f"点击更换文件 · {T.format_bytes(size)}",
                meta=(
                    f"{data.total_rows:,} 行 · {data.total_columns} 列 · "
                    f"{len(data.sheet_names)} 个工作表"
                ),
            )
        else:
            self.dropzone.set_file(
                data.path.name,
                hint="当前工作表不可用 · 点击更换",
                meta=data.problem or "无法读取当前工作表",
                error=True,
            )

    def _open_column_mapping(self) -> None:
        self._mapping_prompt_after = None
        if self.busy or self.current_file is None or self.workbook is None:
            return
        columns = list(self.workbook.raw_columns)
        if not columns:
            columns = [str(column) for column in self.workbook.df.columns]
        if not columns:
            self._show_dialog(
                "无法设置列映射",
                "当前工作表没有可用的表头，请检查文件。",
                "warning",
            )
            return
        ColumnMapDialog(
            self,
            columns=columns,
            mapping=self.workbook.column_map,
            on_apply=self._apply_column_mapping,
        )

    def _apply_column_mapping(self, mapping: dict[str, int]) -> None:
        if self.busy or self.current_file is None or self.workbook is None:
            return
        workbook = self.workbook
        path = self.current_file
        self._set_tools_enabled(False)
        self._set_busy(True, "load", "正在应用列映射…")
        self._reset_result()
        self._show_empty_state("正在应用列映射", "重新读取当前工作表")
        logger.info("应用列映射：%s %s", path, mapping)
        self._start_job(
            "load",
            lambda: core.load_workbook(
                path,
                workbook.sheet_name,
                header_row=workbook.header_row,
                column_map=mapping,
            ),
            self._on_workbook_loaded,
            self._on_job_error,
        )

    def _open_name_review(self) -> None:
        if self.busy or self.workbook is None:
            return
        if not self.workbook.usable:
            self._show_dialog(
                "当前工作表不可用",
                "请先完成列映射，再检查药名。",
                "warning",
            )
            return
        self._set_tools_enabled(False)
        self._set_busy(True, "names", "正在检查疑似药名…")
        workbook = self.workbook
        aliases = dict(self.name_aliases)
        self._start_job(
            "names",
            lambda: core.find_name_candidates(workbook, aliases),
            self._on_name_candidates,
            self._on_job_error,
        )

    def _on_name_candidates(self, groups: list[core.NameCandidateGroup]) -> None:
        self._set_busy(False)
        self._set_tools_enabled(True)
        if not groups and not self.name_aliases:
            self._set_status("药名检查完成 · 没有发现疑似同名", "success")
            self._show_dialog(
                "药名检查完成",
                "没有发现需要人工确认的疑似同名药名。\n\n"
                "全角/半角、括号和多余空格等格式差异已自动规范化。",
                "success",
            )
            return

        dialog = NameReviewDialog(
            self,
            groups=groups,
            current_aliases=self.name_aliases,
        )
        self.wait_window(dialog)
        result = dialog.result
        if result is None:
            self._set_status("已取消药名核对", "neutral")
            return

        self.name_aliases = dict(result)
        self._update_name_review_btn()
        self._reset_result()
        targets = len(set(self.name_aliases.values()))
        if targets:
            self._show_empty_state(
                "药名合并已更新",
                f"当前有 {targets} 组手工合并，点击「开始处理」重新生成结果",
            )
            self._set_status(f"已设置 {targets} 组手工药名合并", "success")
            logger.info("药名手工合并：%s", self.name_aliases)
        else:
            self._show_empty_state(
                "已清除药名合并",
                "点击「开始处理」重新生成结果",
            )
            self._set_status("已清除全部手工药名合并", "neutral")

    def _choose_batch_files(self) -> None:
        if self.busy:
            return
        kwargs = {
            "parent": self,
            "title": "批量选择药品清单",
            "filetypes": [
                ("Excel 工作簿", "*.xlsx *.xls *.xlsm"),
                ("所有文件", "*.*"),
            ],
        }
        if self.last_dir and Path(self.last_dir).is_dir():
            kwargs["initialdir"] = self.last_dir
        selected = filedialog.askopenfilenames(**kwargs)
        if selected:
            self._start_batch(list(selected))

    def _start_batch(self, paths: list[str]) -> None:
        if self.busy:
            return
        mode = self.mode
        sort_order = self.sort_order
        custom_output = self.output_dir_custom
        configured_output = self.output_dir
        all_sheets = bool(self.batch_all_sheets)
        scope = "全部工作表" if all_sheets else "每个文件首个工作表"
        self._set_tools_enabled(False)
        self._set_busy(True, "batch", f"正在批量处理 {len(paths)} 个文件…")
        self._reset_result()
        self._show_empty_state(
            "正在批量处理",
            f"共 {len(paths)} 个文件 · {scope}",
        )
        logger.info(
            "开始批量处理：%s 个文件，范围=%s，模式=%s",
            len(paths),
            scope,
            mode,
        )

        def work():
            started = time.perf_counter()
            outputs: list[Path] = []
            failures: list[str] = []
            for raw_path in paths:
                try:
                    path = core.validate_path(raw_path)
                    first_sheet = core.load_workbook(path)
                    sheet_names = list(first_sheet.sheet_names)
                    targets = (
                        sheet_names
                        if all_sheets
                        else [first_sheet.sheet_name]
                    )
                    for sheet_name in targets:
                        try:
                            sheet = (
                                first_sheet
                                if sheet_name == first_sheet.sheet_name
                                else core.load_workbook(path, sheet_name)
                            )
                            if not sheet.usable:
                                failures.append(
                                    f"{path.name} / {sheet_name}：{sheet.problem}"
                                )
                                continue
                            result = core.analyze(
                                sheet,
                                mode,
                                sort_order=sort_order,
                            )
                            target = (
                                configured_output
                                if custom_output and configured_output is not None
                                else path.parent
                            )
                            label = (
                                sheet_name
                                if all_sheets and len(sheet_names) > 1
                                else ""
                            )
                            outputs.append(
                                core.export_result(
                                    result,
                                    path,
                                    target,
                                    app_version=APP_VERSION,
                                    output_label=label,
                                )
                            )
                        except Exception as exc:  # noqa: BLE001
                            logger.warning(
                                "批量处理失败：%s / %s",
                                raw_path,
                                sheet_name,
                                exc_info=(type(exc), exc, exc.__traceback__),
                            )
                            failures.append(f"{path.name} / {sheet_name}：{exc}")
                except Exception as exc:  # noqa: BLE001
                    logger.warning(
                        "批量处理失败：%s",
                        raw_path,
                        exc_info=(type(exc), exc, exc.__traceback__),
                    )
                    failures.append(f"{Path(raw_path).name}：{exc}")
            return outputs, failures, time.perf_counter() - started, scope

        self._start_job("batch", work, self._on_batch_processed, self._on_job_error)

    def _on_batch_processed(self, payload) -> None:
        outputs, failures, elapsed, scope = payload
        self._set_busy(False)
        self._set_tools_enabled(True)
        self._set_open_buttons_enabled(bool(outputs))
        if outputs:
            self.output_path = Path(outputs[0])
        duration = "不到 0.1 秒" if elapsed < 0.1 else f"{elapsed:.1f} 秒"
        self._show_empty_state(
            "批量处理完成",
            f"成功 {len(outputs)} 个结果 · 未处理 {len(failures)} 项 · {duration}",
        )
        self._set_status(
            f"批量完成 · 成功 {len(outputs)} 个 · 未处理 {len(failures)} 项",
            "success" if not failures else "warning",
        )
        logger.info(
            "批量处理完成：范围=%s 成功=%s 未处理=%s 用时=%.2f 秒",
            scope,
            len(outputs),
            len(failures),
            elapsed,
        )
        message = (
            f"范围：{scope}。\n"
            f"成功导出 {len(outputs)} 个结果，未处理 {len(failures)} 项。"
        )
        if outputs:
            message += f"\n\n结果保存在各自源文件目录，或指定的导出目录。"
        if failures:
            message += "\n\n未处理明细：\n" + "\n".join(failures[:3])
            if len(failures) > 3:
                message += f"\n……另有 {len(failures) - 3} 项未处理，详见运行日志"
        self._show_dialog(
            "批量处理完成",
            message,
            "success" if not failures else "warning",
        )
        if self.open_after_var.get() and outputs:
            self.after(350, lambda: os.startfile(str(outputs[0])))

    def _on_sheet_selected(self, choice: str) -> None:
        if self.busy or self.workbook is None or not self.current_file:
            return
        if choice == self.workbook.sheet_name:
            return
        self.name_aliases = {}
        self._mapping_prompted_for = None
        self._cancel_mapping_prompt()
        self._set_tools_enabled(False)
        self._set_busy(True, "sheet", "正在切换工作表…")
        self._reset_result()
        self._show_empty_state("正在读取工作表", choice)
        path = self.current_file
        self._start_job(
            "sheet",
            lambda: core.load_sheet(path, choice),
            self._on_sheet_loaded,
            self._on_job_error,
        )

    def _on_sheet_loaded(self, data: core.SheetData) -> None:
        self.workbook = data
        self._set_busy(False)
        self._set_tools_enabled(True)
        self.sheet_var.set(data.sheet_name)
        self._update_file_card(data)
        self._reset_result()
        if data.usable:
            self._show_empty_state("还没有结果", "选择处理方式后点击「开始处理」")
            self._set_status(
                f"已切换工作表 · 共 {data.total_rows:,} 行 · 表头第 {data.header_row + 1} 行",
                "neutral",
            )
            self._set_primary("normal")
        else:
            self._show_empty_state("当前工作表不可用", "请检查表头，或点击「列映射」手动指定列")
            self._set_status(data.problem or "当前工作表不可用", "error")
            self._set_primary("disabled")
            self.stat_input.set_value(
                f"{data.total_rows:,}" if data.total_rows else "—",
                "muted",
            )
            self._prompt_column_mapping(data)

    # --------------------------------------------------------- processing ----
    def _start_processing(self) -> None:
        if self.busy:
            return
        if self.workbook is None or not self.current_file:
            self._show_dialog(
                "还没有选择文件",
                "请先拖入或选择一个 Excel 药品清单。",
                "info",
            )
            return
        if not self.workbook.usable:
            self._show_dialog(
                "当前工作表不可用",
                self.workbook.problem or "请检查表头，或切换其他工作表。",
                "warning",
            )
            return

        self._reset_result()
        self._show_empty_state("正在整理数据", "请稍候，结果马上就好")
        self.stat_input.set_value(f"{self.workbook.total_rows:,}")
        self._set_busy(True, "process", "正在整理数据…")

        workbook = self.workbook
        source = self.current_file
        output_dir = self.output_dir
        mode = self.mode
        aliases = dict(self.name_aliases)
        sort_order = self.sort_order

        def work():
            started = time.perf_counter()
            result = core.analyze(
                workbook,
                mode,
                name_aliases=aliases,
                sort_order=sort_order,
            )
            output = core.export_result(
                result,
                source,
                output_dir,
                app_version=APP_VERSION,
            )
            return result, output, time.perf_counter() - started

        self._start_job("process", work, self._on_processed, self._on_job_error)

    def _on_processed(self, payload) -> None:
        result: core.AnalysisResult
        output: Path
        elapsed: float
        result, output, elapsed = payload
        self.result = result
        self.output_path = Path(output)

        self._set_busy(False)
        self._update_stats(result)
        self._populate_preview(result)
        self._set_open_buttons_enabled(True)

        note = ""
        if result.skipped_rows:
            note = f" · 忽略 {result.skipped_rows:,} 行空名称"
        if result.alias_groups:
            note += f" · 手工合并 {result.alias_groups} 组药名"
        elif result.normalised_name_groups:
            note += f" · 自动规范 {result.normalised_name_groups} 组药名"
        if result.duplicate_groups:
            prefix = "处理完成"
        else:
            prefix = "未发现重复项"
        duration = "不到 0.1 秒" if elapsed < 0.1 else f"{elapsed:.1f} 秒"
        self._set_status(
            f"{prefix} · {duration} · 已保存 {output.name}{note}",
            "success",
        )
        logger.info(
            "处理完成：%s｜%s｜%s｜原始=%s 结果=%s 重复组=%s 明细=%s｜输出=%s",
            result.source_file,
            result.source_sheet,
            core.MODE_LABELS[result.mode],
            result.total_rows,
            result.output_rows,
            result.duplicate_groups,
            len(result.details),
            output,
        )

        if self.open_after_var.get():
            self.after(350, self._open_output_file)

    # -------------------------------------------------------------- jobs ----
    def _start_job(self, kind: str, function, success, error) -> None:
        self._job_kind = kind
        self._job_callbacks = (success, error)

        def worker() -> None:
            try:
                payload = function()
            except BaseException as exc:  # noqa: BLE001
                self._job_queue.put(("error", exc))
            else:
                self._job_queue.put(("ok", payload))

        threading.Thread(target=worker, daemon=True, name=f"job-{kind}").start()
        self._poll_job()

    def _poll_job(self) -> None:
        if self._closing:
            return
        try:
            status, payload = self._job_queue.get_nowait()
        except queue.Empty:
            try:
                self._poll_after = self.after(70, self._poll_job)
            except Exception:  # noqa: BLE001
                self._poll_after = None
            return

        self._poll_after = None
        callbacks = self._job_callbacks
        self._job_callbacks = None
        if callbacks is None:
            return
        success, error = callbacks
        try:
            if status == "ok":
                success(payload)
            else:
                error(payload)
        except Exception as exc:  # noqa: BLE001
            logger.error(
                "任务回调失败：%s",
                self._job_kind,
                exc_info=(type(exc), exc, exc.__traceback__),
            )
            self._set_busy(False)
            self._set_status("处理时发生意外错误", "error")
            self._show_dialog(
                "操作没有完成",
                f"出现了一个意外错误：{exc}",
                "error",
            )

    def _on_job_error(self, exc: BaseException) -> None:
        logger.error(
            "任务失败：%s",
            self._job_kind,
            exc_info=(type(exc), exc, exc.__traceback__),
        )
        kind = self._job_kind
        message = str(exc) if isinstance(exc, core.UserError) else (
            f"处理时遇到了意外问题：{exc}" if str(exc) else "处理时遇到了意外问题。"
        )

        if kind == "load":
            self.current_file = None
            self.workbook = None
            self.output_dir_custom = False
            self.output_dir = None
            self._update_output_label()
            self.dropzone.set_empty()
            self.sheet_var.set("")
            self.sheet_row.grid_remove()
            self._set_busy(False)
            self._set_stats_blank()
            self._show_empty_state("无法读取文件", "请换一个文件后重试")
            self._set_status("文件读取失败", "error")
            self._show_dialog("无法打开文件", message, "error")
            return

        if kind == "sheet":
            self._set_busy(False)
            if self.workbook is not None:
                self.sheet_var.set(self.workbook.sheet_name)
                self._update_file_card(self.workbook)
                if self.workbook.usable:
                    self._show_empty_state("还没有结果", "选择处理方式后点击「开始处理」")
                    self._set_primary("normal")
                else:
                    self._show_empty_state("当前工作表不可用", "请检查表头，或切换其他工作表")
                    self._set_primary("disabled")
            self._set_status("工作表读取失败", "error")
            self._show_dialog("无法读取工作表", message, "error")
            return

        if kind == "names":
            self._set_busy(False)
            self._set_status("药名检查失败", "error")
            self._show_dialog("药名检查失败", message, "error")
            return

        if kind == "batch":
            self._set_busy(False)
            self._show_empty_state("批量处理失败", "请检查文件和导出目录后重试")
            self._set_status("批量处理失败", "error")
            self._show_dialog("批量处理失败", message, "error")
            return

        self._set_busy(False)
        self._show_empty_state("处理没有完成", "请根据提示检查文件后重试")
        self._set_status("处理失败", "error")
        self._show_dialog("处理失败", message, "error")

    # ----------------------------------------------------------- actions ----
    def _on_mode_clicked(self, mode: str) -> None:
        if self.busy or mode == self.mode:
            return
        self._apply_mode(mode, reset=True)

    def _apply_mode(self, mode: str, reset: bool) -> None:
        self.mode = mode
        try:
            self.mode_simplify_card.set_selected(mode == core.MODE_SIMPLIFY)
            self.mode_count_card.set_selected(mode == core.MODE_COUNT)
        except Exception:  # noqa: BLE001
            pass

        if reset:
            self._reset_result()
            if self.workbook is not None and self.workbook.usable:
                label = core.MODE_LABELS.get(mode, mode)
                self._show_empty_state("还没有结果", f"已切换为「{label}」")
                self._set_status(f"已切换为「{label}」，点击开始处理", "neutral")
            elif self.workbook is not None:
                self._set_status(self.workbook.problem or "当前工作表不可用", "error")
            else:
                self._set_status("就绪 · 选择 Excel 文件开始", "neutral")
        self._save_settings()

    def _choose_output_dir(self) -> None:
        if self.busy:
            return
        initial = self.output_dir
        if initial is None and self.current_file is not None:
            initial = self.current_file.parent
        kwargs = {
            "parent": self,
            "title": "选择导出文件夹",
            "mustexist": True,
        }
        if initial is not None and Path(initial).is_dir():
            kwargs["initialdir"] = str(initial)
        selected = filedialog.askdirectory(**kwargs)
        if selected:
            self.output_dir_custom = True
            self.output_dir = Path(selected)
            self._update_output_label()

    def _update_output_label(self) -> None:
        if self.output_dir is None:
            text = "与源文件相同的文件夹"
        else:
            text = str(self.output_dir)
        try:
            self.output_label.configure(
                text=T.ellipsize_units(text, 46),
            )
        except Exception:  # noqa: BLE001
            pass

    def _on_appearance_changed(self, label: str) -> None:
        key = LABEL_TO_APPEARANCE.get(label, "system")
        self.appearance_key = key
        ctk.set_appearance_mode(key)
        self._save_settings()
        self.after(80, self._style_table)

    def _on_sort_order_changed(self, label: str) -> None:
        for key, value in core.SORT_ORDER_LABELS.items():
            if value == label:
                self.sort_order = key
                break
        self._save_settings()
        if hasattr(self, "preview_sort_var"):
            self.preview_sort_var.set(PREVIEW_SORT_LABELS.get(self.sort_order, "首次出现"))
            if self.result is not None:
                self._refresh_preview()
        self._set_status(f"结果排序已设为「{label}」，下次处理生效", "neutral")

    def _on_preview_sort_changed(self, label: str) -> None:
        key = PREVIEW_LABEL_TO_SORT.get(label)
        if key:
            self.sort_order = key
            if hasattr(self, "sort_order_var"):
                self.sort_order_var.set(core.SORT_ORDER_LABELS[key])
            self._save_settings()
        self._refresh_preview()

    def _on_batch_all_sheets_changed(self) -> None:
        self.batch_all_sheets = bool(self.batch_all_sheets_var.get())
        self._save_settings()

    def _save_settings(self) -> None:
        data = {
            "appearance": self.appearance_key,
            "mode": self.mode,
            "open_after": bool(self.open_after_var.get())
            if hasattr(self, "open_after_var")
            else False,
            "last_dir": self.last_dir,
            "sort_order": self.sort_order,
            "batch_all_sheets": self.batch_all_sheets,
        }
        save_config(data)

    def _open_output_file(self) -> None:
        if not self.output_path or not Path(self.output_path).exists():
            self._show_dialog(
                "找不到结果文件",
                "结果文件可能已被移动或删除，请重新处理。",
                "warning",
            )
            return
        try:
            os.startfile(str(self.output_path))  # type: ignore[attr-defined]
        except OSError as exc:
            self._show_dialog("无法打开文件", f"系统无法打开这个文件：{exc}", "error")

    def _open_output_folder(self) -> None:
        if not self.output_path:
            return
        folder = Path(self.output_path).parent
        if not folder.exists():
            self._show_dialog("找不到文件夹", "结果文件夹可能已被移动或删除。", "warning")
            return
        try:
            os.startfile(str(folder))  # type: ignore[attr-defined]
        except OSError as exc:
            self._show_dialog("无法打开文件夹", f"系统无法打开这个文件夹：{exc}", "error")

    # ------------------------------------------------------------ dialogs ----
    def _show_dialog(self, title: str, message: str, kind: str = "info") -> None:
        if self._closing:
            return
        try:
            if self._dialog is not None and self._dialog.winfo_exists():
                self._dialog.destroy()
        except Exception:  # noqa: BLE001
            pass
        try:
            self._dialog = AppDialog(self, title=title, message=message, kind=kind)
        except Exception:  # noqa: BLE001
            logger.error("无法显示对话框", exc_info=True)

    def _report_callback_exception(self, exc, value, tb) -> None:
        logger.error(
            "界面回调异常",
            exc_info=(exc, value, tb),
        )
        try:
            self._show_dialog(
                "界面出现异常",
                f"操作没有完成：{value}",
                "error",
            )
        except Exception:  # noqa: BLE001
            pass

    # ---------------------------------------------------------- lifecycle ----
    def _startup(self) -> None:
        if self._initial_path:
            path = self._initial_path
            self._initial_path = None
            self._load_file(path)

    def _watch_ui(self) -> None:
        if self._closing:
            return
        try:
            resolved = ctk.get_appearance_mode()
            scale = ctk.ScalingTracker.get_widget_scaling(self)
            if resolved != self._last_resolved_mode or abs(scale - self._last_scale) > 0.001:
                self._last_resolved_mode = resolved
                self._last_scale = scale
                self._style_table()
        except Exception:  # noqa: BLE001
            logger.debug("界面缩放检查失败", exc_info=True)
        try:
            self._watch_after = self.after(1500, self._watch_ui)
        except Exception:  # noqa: BLE001
            self._watch_after = None

    def _on_close(self) -> None:
        if self.busy:
            self._show_dialog(
                "正在处理",
                "文件仍在处理中，请等待完成后再关闭窗口。",
                "info",
            )
            return
        self._closing = True
        for after_id in (
            self._poll_after,
            self._watch_after,
            self._resize_after,
            self._preview_filter_after,
            self._mapping_prompt_after,
            self._show_after,
        ):
            if after_id:
                try:
                    self.after_cancel(after_id)
                except Exception:  # noqa: BLE001
                    pass
        self.destroy()


def main() -> None:
    setup_logging()
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "DrugListAssistant.1.1"
            )
        except Exception:  # noqa: BLE001
            pass

    initial_path = None
    for argument in sys.argv[1:]:
        if argument and not argument.startswith("-"):
            initial_path = argument
            break

    config = load_config()
    app = App(config, initial_path=initial_path)
    app.mainloop()


if __name__ == "__main__":
    main()
