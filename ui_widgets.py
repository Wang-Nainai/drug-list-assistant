"""可复用的界面组件：模式卡片、统计卡片、拖拽区域和提示对话框。"""

from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from typing import Callable, Iterable

import customtkinter as ctk

try:
    from tkinterdnd2 import DND_FILES
except Exception:  # noqa: BLE001 - 拖拽不可用时仍可点击选择
    DND_FILES = None

import ui_theme as T


def walk_widgets(widget) -> Iterable:
    yield widget
    for child in widget.winfo_children():
        yield from walk_widgets(child)


def _pointer_inside(widget) -> bool:
    try:
        x, y = widget.winfo_pointerxy()
        target = widget.winfo_containing(x, y)
    except Exception:  # noqa: BLE001
        return False
    while target is not None:
        if target is widget:
            return True
        target = getattr(target, "master", None)
    return False


class ModeCard(ctk.CTkFrame):
    """可点击的单选卡片。"""

    def __init__(
        self,
        master,
        *,
        value: str,
        title: str,
        description: str,
        badge: str,
        command: Callable[[str], None],
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            fg_color=T.SURFACE,
            corner_radius=10,
            border_width=2,
            border_color=T.BORDER,
            **kwargs,
        )
        self.value = value
        self.command = command
        self.selected = False
        self._hover = False
        self._leave_job = None

        self.grid_columnconfigure(1, weight=1)

        self.badge = ctk.CTkLabel(
            self,
            text=badge,
            width=34,
            height=34,
            corner_radius=9,
            fg_color=T.ACCENT_SOFT,
            text_color=T.ACCENT,
            font=T.font(13.5, "bold"),
        )
        self.badge.grid(row=0, column=0, padx=(13, 11), pady=12)

        text_frame = ctk.CTkFrame(self, fg_color="transparent")
        text_frame.grid(row=0, column=1, sticky="w", pady=10)
        text_frame.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(
            text_frame,
            text=title,
            font=T.font(12.5, "bold"),
            text_color=T.TEXT,
            anchor="w",
        )
        self.title_label.grid(row=0, column=0, sticky="w")

        self.description_label = ctk.CTkLabel(
            text_frame,
            text=description,
            font=T.font(9.3),
            text_color=T.TEXT_SECONDARY,
            anchor="w",
        )
        self.description_label.grid(row=1, column=0, sticky="w", pady=(1, 0))

        self.ring = ctk.CTkFrame(
            self,
            width=18,
            height=18,
            corner_radius=9,
            border_width=2,
            border_color=T.BORDER_STRONG,
            fg_color="transparent",
        )
        self.ring.grid(row=0, column=2, padx=(8, 14))
        self.ring.grid_propagate(False)
        self.dot = ctk.CTkFrame(
            self.ring,
            width=8,
            height=8,
            corner_radius=4,
            fg_color="transparent",
        )
        self.dot.place(relx=0.5, rely=0.5, anchor="center")

        for widget in walk_widgets(self):
            widget.bind("<Button-1>", self._clicked, add="+")
            widget.bind("<Enter>", self._entered, add="+")
            widget.bind("<Leave>", self._left, add="+")
            try:
                widget.configure(cursor="hand2")
            except Exception:  # noqa: BLE001
                pass

    def _clicked(self, _event=None) -> None:
        self.command(self.value)

    def _entered(self, _event=None) -> None:
        if self._leave_job is not None:
            try:
                self.after_cancel(self._leave_job)
            except Exception:  # noqa: BLE001
                pass
            self._leave_job = None
        self._hover = True
        self._refresh_visual()

    def _left(self, _event=None) -> None:
        if self._leave_job is not None:
            try:
                self.after_cancel(self._leave_job)
            except Exception:  # noqa: BLE001
                pass
        self._leave_job = self.after(30, self._check_hover)

    def _check_hover(self) -> None:
        self._leave_job = None
        if not _pointer_inside(self):
            self._hover = False
            self._refresh_visual()

    def _refresh_visual(self) -> None:
        if self.selected:
            background = T.SURFACE_TINT
            border = T.ACCENT
        elif self._hover:
            background = T.SURFACE_HOVER
            border = T.BORDER_STRONG
        else:
            background = T.SURFACE
            border = T.BORDER

        try:
            self.configure(fg_color=background, border_color=border)
            self.ring.configure(border_color=T.ACCENT if self.selected else T.BORDER_STRONG)
            self.dot.configure(fg_color=T.ACCENT if self.selected else "transparent")
        except Exception:  # noqa: BLE001
            pass

    def set_selected(self, selected: bool) -> None:
        self.selected = bool(selected)
        self._refresh_visual()


class StatCard(ctk.CTkFrame):
    def __init__(self, master, label: str, value: str = "—", **kwargs) -> None:
        super().__init__(
            master,
            width=130,
            fg_color=T.SURFACE,
            corner_radius=12,
            border_width=1,
            border_color=T.BORDER,
            **kwargs,
        )
        self.grid_columnconfigure(0, weight=1)

        self.name_label = ctk.CTkLabel(
            self,
            text=label,
            font=T.font(9.8),
            text_color=T.TEXT_TERTIARY,
            anchor="w",
        )
        self.name_label.grid(row=0, column=0, sticky="ew", padx=15, pady=(12, 0))

        self.value_label = ctk.CTkLabel(
            self,
            text=value,
            font=T.font(22, "bold"),
            text_color=T.TEXT,
            anchor="w",
        )
        self.value_label.grid(row=1, column=0, sticky="ew", padx=15, pady=(1, 11))

    def set_value(self, value: str, tone: str = "default") -> None:
        colours = {
            "default": T.TEXT,
            "accent": T.ACCENT,
            "success": T.SUCCESS,
            "warning": T.WARNING,
            "muted": T.TEXT_TERTIARY,
        }
        self.value_label.configure(text=value, text_color=colours.get(tone, T.TEXT))


class DropZone(ctk.CTkFrame):
    """支持点击与拖拽的文件选择卡片。"""

    def __init__(
        self,
        master,
        *,
        on_pick: Callable[[], None],
        on_drop: Callable[[list[str]], None],
        height: int = 104,
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            fg_color=T.SURFACE_ALT,
            corner_radius=10,
            border_width=1,
            border_color=T.BORDER,
            height=height,
            cursor="hand2",
            **kwargs,
        )
        self.on_pick = on_pick
        self.on_drop = on_drop
        self._hover = False
        self._drag = False
        self._enabled = True
        self._error = False
        self._leave_job = None
        self._drag_leave_job = None

        self.grid_propagate(False)
        self.grid_columnconfigure(1, weight=1)

        self.icon_badge = ctk.CTkFrame(
            self,
            width=38,
            height=38,
            corner_radius=10,
            fg_color=T.ACCENT_SOFT,
        )
        self.icon_badge.grid(row=0, column=0, rowspan=3, padx=(14, 12))
        self.icon_badge.grid_propagate(False)
        self.icon_label = ctk.CTkLabel(
            self.icon_badge,
            text=T.ICON_DOCUMENT,
            font=T.icon_font(18),
            text_color=T.ACCENT,
        )
        self.icon_label.place(relx=0.5, rely=0.5, anchor="center")

        self.title_label = ctk.CTkLabel(
            self,
            text="拖拽 Excel 文件到这里",
            font=T.font(12, "bold"),
            text_color=T.TEXT,
            anchor="w",
        )
        self.title_label.grid(row=0, column=1, sticky="ew", pady=(14, 0))

        self.hint_label = ctk.CTkLabel(
            self,
            text="或点击选择 · 支持 .xlsx / .xls",
            font=T.font(9.2),
            text_color=T.TEXT_SECONDARY,
            anchor="w",
        )
        self.hint_label.grid(row=1, column=1, sticky="ew", pady=(0, 0))

        self.meta_label = ctk.CTkLabel(
            self,
            text="需含：序号、药品名称、药品规格",
            font=T.font(8.8),
            text_color=T.TEXT_TERTIARY,
            anchor="w",
        )
        self.meta_label.grid(row=2, column=1, sticky="ew", pady=(1, 12))

        self.action_label = ctk.CTkLabel(
            self,
            text="浏览",
            font=T.font(9.8, "bold"),
            text_color=T.ACCENT,
        )
        self.action_label.grid(row=0, column=2, rowspan=3, padx=(8, 14))

        for widget in walk_widgets(self):
            widget.bind("<Button-1>", self._clicked, add="+")
            widget.bind("<Enter>", self._entered, add="+")
            widget.bind("<Leave>", self._left, add="+")

        self._register_drag_and_drop()

    # ------------------------------------------------------------- dnd ----
    def _register_drag_and_drop(self) -> None:
        if DND_FILES is None:
            return
        for widget in walk_widgets(self):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<DropEnter>>", self._drag_entered)
                widget.dnd_bind("<<DropLeave>>", self._drag_left)
                widget.dnd_bind("<<Drop>>", self._dropped)
            except Exception:  # noqa: BLE001
                continue

    def _drag_entered(self, _event=None) -> None:
        if not self._enabled:
            return
        if self._drag_leave_job is not None:
            try:
                self.after_cancel(self._drag_leave_job)
            except Exception:  # noqa: BLE001
                pass
            self._drag_leave_job = None
        self._drag = True
        self._refresh_visual()

    def _drag_left(self, _event=None) -> None:
        if self._drag_leave_job is not None:
            try:
                self.after_cancel(self._drag_leave_job)
            except Exception:  # noqa: BLE001
                pass
        self._drag_leave_job = self.after(60, self._clear_drag)

    def _clear_drag(self) -> None:
        self._drag_leave_job = None
        self._drag = False
        self._refresh_visual()

    def _dropped(self, event) -> None:
        self._drag = False
        if self._drag_leave_job is not None:
            try:
                self.after_cancel(self._drag_leave_job)
            except Exception:  # noqa: BLE001
                pass
            self._drag_leave_job = None
        self._refresh_visual()
        if not self._enabled:
            return
        try:
            items = list(self.tk.splitlist(event.data))
        except Exception:  # noqa: BLE001
            items = [str(event.data)]
        paths = [item for item in items if item and os.path.isfile(item)]
        if paths:
            self.on_drop(paths)

    # ---------------------------------------------------------- hover ----
    def _entered(self, _event=None) -> None:
        if self._leave_job is not None:
            try:
                self.after_cancel(self._leave_job)
            except Exception:  # noqa: BLE001
                pass
            self._leave_job = None
        if not self._enabled:
            return
        self._hover = True
        self._refresh_visual()

    def _left(self, _event=None) -> None:
        if self._leave_job is not None:
            try:
                self.after_cancel(self._leave_job)
            except Exception:  # noqa: BLE001
                pass
        self._leave_job = self.after(30, self._check_hover)

    def _check_hover(self) -> None:
        self._leave_job = None
        if not _pointer_inside(self):
            self._hover = False
            self._refresh_visual()

    def _refresh_visual(self) -> None:
        if not self._enabled:
            background, border = T.SURFACE_ALT, T.BORDER
        elif self._drag:
            background, border = T.ACCENT_SOFT, T.ACCENT
        elif self._hover:
            background, border = T.SURFACE_HOVER, T.BORDER_STRONG
        else:
            background, border = T.SURFACE_ALT, T.BORDER
        try:
            self.configure(fg_color=background, border_color=border)
        except Exception:  # noqa: BLE001
            pass

    def _clicked(self, _event=None) -> None:
        if self._enabled:
            self.on_pick()

    # ----------------------------------------------------------- state ----
    def set_enabled(self, enabled: bool) -> None:
        self._enabled = bool(enabled)
        try:
            self.configure(cursor="hand2" if enabled else "arrow")
        except Exception:  # noqa: BLE001
            pass
        self._refresh_visual()

    def set_empty(self) -> None:
        self._error = False
        self.title_label.configure(text="拖拽 Excel 文件到这里", text_color=T.TEXT)
        self.hint_label.configure(text="或点击选择 · 支持 .xlsx / .xls")
        self.meta_label.configure(
            text="需含：序号、药品名称、药品规格",
            text_color=T.TEXT_TERTIARY,
        )
        self.action_label.configure(text="浏览", text_color=T.ACCENT)
        self.icon_badge.configure(fg_color=T.ACCENT_SOFT)
        self.icon_label.configure(text_color=T.ACCENT)
        self._refresh_visual()

    def set_loading(self, filename: str) -> None:
        self._error = False
        self.title_label.configure(text=T.ellipsize_units(filename, 38), text_color=T.TEXT)
        self.hint_label.configure(text="正在读取文件…", text_color=T.TEXT_SECONDARY)
        self.meta_label.configure(text="请稍候", text_color=T.TEXT_TERTIARY)
        self.action_label.configure(text="更换", text_color=T.ACCENT)
        self.icon_badge.configure(fg_color=T.ACCENT_SOFT)
        self.icon_label.configure(text_color=T.ACCENT)
        self._refresh_visual()

    def set_file(
        self,
        filename: str,
        *,
        hint: str,
        meta: str,
        error: bool = False,
    ) -> None:
        self._error = error
        self.title_label.configure(text=T.ellipsize_units(filename, 38), text_color=T.TEXT)
        self.hint_label.configure(text=T.ellipsize_units(hint, 44), text_color=T.TEXT_SECONDARY)
        self.meta_label.configure(
            text=T.ellipsize_units(meta, 36),
            text_color=T.DANGER if error else T.TEXT_TERTIARY,
        )
        self.action_label.configure(text="更换", text_color=T.ACCENT)
        if error:
            self.icon_badge.configure(fg_color=T.DANGER_SOFT)
            self.icon_label.configure(text_color=T.DANGER)
        else:
            self.icon_badge.configure(fg_color=T.ACCENT_SOFT)
            self.icon_label.configure(text_color=T.ACCENT)
        self._refresh_visual()


class AppDialog(ctk.CTkToplevel):
    """简洁的提示对话框，风格与主窗口一致。"""

    def __init__(
        self,
        parent,
        *,
        title: str,
        message: str,
        kind: str = "info",
    ) -> None:
        super().__init__(parent)
        self.parent = parent
        self.kind = kind
        self.title("提示")
        self.configure(fg_color=T.SURFACE)
        self.resizable(False, False)
        self.transient(parent)

        tone = {
            "error": (T.DANGER, T.DANGER_SOFT, T.ICON_ERROR),
            "warning": (T.WARNING, T.WARNING_SOFT, T.ICON_WARNING),
            "success": (T.SUCCESS, T.SUCCESS_SOFT, T.ICON_CHECK),
            "info": (T.ACCENT, T.ACCENT_SOFT, T.ICON_INFO),
        }.get(kind, (T.ACCENT, T.ACCENT_SOFT, T.ICON_INFO))
        colour, soft, glyph = tone

        estimated_lines = max(2, min(6, (len(message) + 32) // 33))
        height = 190 + estimated_lines * 18
        width = 430
        self.geometry(f"{width}x{height}")

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=22, pady=(20, 18))
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(1, weight=1)

        icon_frame = ctk.CTkFrame(
            content,
            width=40,
            height=40,
            corner_radius=20,
            fg_color=soft,
        )
        icon_frame.grid(row=0, column=0, sticky="n", padx=(0, 14))
        icon_frame.grid_propagate(False)
        ctk.CTkLabel(
            icon_frame,
            text=glyph,
            font=T.icon_font(17),
            text_color=colour,
        ).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            content,
            text=title,
            font=T.font(14, "bold"),
            text_color=T.TEXT,
            anchor="w",
            justify="left",
        ).grid(row=0, column=1, sticky="w", pady=(8, 2))

        ctk.CTkLabel(
            content,
            text=message,
            font=T.font(10.3),
            text_color=T.TEXT_SECONDARY,
            anchor="nw",
            justify="left",
            wraplength=330,
        ).grid(row=1, column=1, sticky="nw", pady=(2, 10))

        button_row = ctk.CTkFrame(content, fg_color="transparent")
        button_row.grid(row=2, column=0, columnspan=2, sticky="e", pady=(4, 0))
        self.ok_button = ctk.CTkButton(
            button_row,
            text="知道了",
            width=96,
            height=34,
            corner_radius=8,
            fg_color=T.ACCENT,
            hover_color=T.ACCENT_HOVER,
            text_color=T.ON_ACCENT,
            font=T.font(10.5, "bold"),
            command=self._close,
        )
        self.ok_button.pack()

        self.protocol("WM_DELETE_WINDOW", self._close)
        self.bind("<Escape>", lambda _e: self._close())
        self.bind("<Return>", lambda _e: self._close())

        try:
            self.iconbitmap(T.resource_path("assets", "app.ico"))
        except Exception:  # noqa: BLE001
            pass

        self.after(40, self._centre_and_focus)

    def _centre_and_focus(self) -> None:
        try:
            self.update_idletasks()
            scale = ctk.ScalingTracker.get_window_scaling(self)
            width = self.winfo_width()
            height = self.winfo_height()
            parent_x = self.parent.winfo_rootx()
            parent_y = self.parent.winfo_rooty()
            parent_w = self.parent.winfo_width()
            parent_h = self.parent.winfo_height()
            x = int((parent_x + (parent_w - width) / 2) / scale)
            y = int((parent_y + (parent_h - height) / 2) / scale)
            self.geometry(f"+{x}+{y}")
            self.grab_set()
            self.lift()
            self.focus_force()
            self.ok_button.focus_set()
        except Exception:  # noqa: BLE001
            pass

    def _close(self) -> None:
        try:
            self.grab_release()
        except Exception:  # noqa: BLE001
            pass
        self.destroy()


class ColumnMapDialog(ctk.CTkToplevel):
    """必要列没有自动识别时，让用户手动指定列位置。"""

    REQUIRED = ("序号", "药品名称", "药品规格")

    def __init__(
        self,
        parent,
        *,
        columns: list[str],
        mapping: dict[str, int],
        on_apply: Callable[[dict[str, int]], None],
    ) -> None:
        super().__init__(parent)
        self.parent = parent
        self.on_apply = on_apply
        self.result: dict[str, int] | None = None
        self.title("设置列对应关系")
        self.configure(fg_color=T.SURFACE)
        self.resizable(False, False)
        self.transient(parent)
        self.geometry("520x390")

        self._options = [
            f"{index + 1}. {label}" for index, label in enumerate(columns)
        ]
        self._option_index = {
            option: index for index, option in enumerate(self._options)
        }
        self._labels = ["请选择"] + self._options
        self._vars: dict[str, tk.StringVar] = {}
        self._menus: dict[str, ctk.CTkOptionMenu] = {}

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=24, pady=(20, 16))
        content.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            content,
            text="设置列对应关系",
            font=T.font(15, "bold"),
            text_color=T.TEXT,
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        ctk.CTkLabel(
            content,
            text="表头没有完全识别，请指定下面三个字段分别对应哪一列。",
            font=T.font(9.6),
            text_color=T.TEXT_SECONDARY,
            anchor="w",
            justify="left",
            wraplength=450,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 16))

        for row, canonical in enumerate(self.REQUIRED, start=2):
            ctk.CTkLabel(
                content,
                text=canonical,
                font=T.font(10.5, "bold"),
                text_color=T.TEXT,
                width=84,
                anchor="w",
            ).grid(row=row, column=0, sticky="w", pady=6)
            current = mapping.get(canonical)
            initial = (
                self._options[current]
                if current is not None and 0 <= current < len(self._options)
                else self._labels[0]
            )
            variable = tk.StringVar(value=initial)
            self._vars[canonical] = variable
            menu = ctk.CTkOptionMenu(
                content,
                variable=variable,
                values=self._labels,
                width=360,
                height=32,
                corner_radius=8,
                fg_color=T.SURFACE_ALT,
                button_color=T.SURFACE_ALT,
                button_hover_color=T.TRACK,
                text_color=T.TEXT,
                font=T.font(9.4),
                dropdown_fg_color=T.SURFACE,
                dropdown_hover_color=T.ACCENT_SOFT,
                dropdown_text_color=T.TEXT,
                dropdown_font=T.font(9.4),
                dynamic_resizing=False,
            )
            menu.grid(row=row, column=1, sticky="ew", pady=6)
            self._menus[canonical] = menu

        self.error_label = ctk.CTkLabel(
            content,
            text="",
            font=T.font(9.0),
            text_color=T.DANGER,
            anchor="w",
        )
        self.error_label.grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

        button_row = ctk.CTkFrame(content, fg_color="transparent")
        button_row.grid(row=6, column=0, columnspan=2, sticky="e", pady=(16, 0))

        ctk.CTkButton(
            button_row,
            text="取消",
            width=82,
            height=34,
            corner_radius=8,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            font=T.font(9.8),
            command=self._close,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            button_row,
            text="应用列映射",
            width=110,
            height=34,
            corner_radius=8,
            fg_color=T.ACCENT,
            hover_color=T.ACCENT_HOVER,
            text_color=T.ON_ACCENT,
            font=T.font(9.8, "bold"),
            command=self._apply,
        ).pack(side="left")

        self.protocol("WM_DELETE_WINDOW", self._close)
        self.bind("<Escape>", lambda _e: self._close())
        try:
            self.iconbitmap(T.resource_path("assets", "app.ico"))
        except Exception:  # noqa: BLE001
            pass
        self.after(40, self._centre_and_focus)

    def _centre_and_focus(self) -> None:
        try:
            self.update_idletasks()
            scale = ctk.ScalingTracker.get_window_scaling(self)
            width = self.winfo_width()
            height = self.winfo_height()
            parent_x = self.parent.winfo_rootx()
            parent_y = self.parent.winfo_rooty()
            parent_w = self.parent.winfo_width()
            parent_h = self.parent.winfo_height()
            x = int((parent_x + (parent_w - width) / 2) / scale)
            y = int((parent_y + (parent_h - height) / 2) / scale)
            self.geometry(f"+{x}+{y}")
            self.grab_set()
            self.lift()
            self.focus_force()
        except Exception:  # noqa: BLE001
            pass

    def _apply(self) -> None:
        result: dict[str, int] = {}
        for canonical, variable in self._vars.items():
            value = variable.get()
            index = self._option_index.get(value)
            if index is None:
                self.error_label.configure(text="请为三个必要字段各选择一列。")
                return
            result[canonical] = index

        if len(set(result.values())) < len(result):
            self.error_label.configure(text="同一个列不能同时对应多个字段。")
            return

        self.result = result
        self._close()
        try:
            self.on_apply(result)
        except Exception:  # noqa: BLE001
            pass

    def _close(self) -> None:
        try:
            self.grab_release()
        except Exception:  # noqa: BLE001
            pass
        self.destroy()


class NameReviewDialog(ctk.CTkToplevel):
    """人工确认疑似同名的药名组；默认不合并。"""

    def __init__(self, parent, *, groups, current_aliases: dict[str, str] | None = None) -> None:
        super().__init__(parent)
        self.parent = parent
        self.groups = list(groups)
        self.current_aliases = dict(current_aliases or {})
        self.result: dict[str, str] | None = None
        self.title("药名核对")
        self.configure(fg_color=T.SURFACE)
        self.transient(parent)
        self.geometry("700x680")
        self.minsize(620, 520)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(20, 8))
        ctk.CTkLabel(
            header,
            text="药名核对",
            font=T.font(15, "bold"),
            text_color=T.TEXT,
            anchor="w",
        ).pack(anchor="w")
        if self.groups:
            summary_text = (
                f"发现 {len(self.groups)} 组写法相近的药名。程序不会自动合并，"
                "请勾选确认后再应用。"
            )
        elif self.current_aliases:
            summary_text = "当前文件已有手工合并的药名。可以清除这些合并后重新处理。"
        else:
            summary_text = "没有发现需要人工确认的疑似同名药名。"
        ctk.CTkLabel(
            header,
            text=summary_text,
            font=T.font(9.6),
            text_color=T.TEXT_SECONDARY,
            anchor="w",
            justify="left",
            wraplength=620,
        ).pack(anchor="w", pady=(4, 0))

        body = ctk.CTkScrollableFrame(
            self,
            fg_color=T.SURFACE_ALT,
            corner_radius=10,
            border_width=1,
            border_color=T.BORDER,
        )
        body.pack(fill="both", expand=True, padx=24, pady=(6, 10))

        self._checks: list[tuple[tk.BooleanVar, object]] = []
        for group in self.groups:
            card = ctk.CTkFrame(
                body,
                fg_color=T.SURFACE,
                corner_radius=9,
                border_width=1,
                border_color=T.BORDER,
            )
            card.pack(fill="x", pady=(0, 8), padx=2)
            card.grid_columnconfigure(0, weight=1)

            canonical = max(
                group.variants,
                key=lambda variant: (variant.count, len(variant.key)),
            )
            variable = tk.BooleanVar(value=False)
            self._checks.append((variable, group))
            ctk.CTkCheckBox(
                card,
                text=f"合并本组（共 {group.total_count} 条）",
                variable=variable,
                font=T.font(10.2, "bold"),
                text_color=T.TEXT,
                fg_color=T.ACCENT,
                hover_color=T.ACCENT_HOVER,
                border_color=T.BORDER_STRONG,
                checkbox_width=20,
                checkbox_height=20,
                corner_radius=5,
            ).grid(row=0, column=0, sticky="w", padx=14, pady=(11, 2))

            ctk.CTkLabel(
                card,
                text=f"合并到：{canonical.display}",
                font=T.font(9.0),
                text_color=T.ACCENT,
                anchor="w",
            ).grid(row=1, column=0, sticky="w", padx=42, pady=(0, 2))

            def _variant_text(variant) -> str:
                spec_text = (
                    "、".join(variant.specs[:4]) if variant.specs else "规格为空"
                )
                more = "…" if len(variant.specs) > 4 else ""
                return (
                    f"• {variant.display}（{variant.count} 条；"
                    f"规格：{spec_text}{more}）"
                )

            names = "\n".join(
                _variant_text(variant)
                for variant in sorted(
                    group.variants, key=lambda item: item.count, reverse=True
                )
            )
            ctk.CTkLabel(
                card,
                text=names,
                font=T.font(9.2),
                text_color=T.TEXT_SECONDARY,
                anchor="w",
                justify="left",
                wraplength=560,
            ).grid(row=2, column=0, sticky="w", padx=42, pady=(0, 10))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=24, pady=(0, 18))
        ctk.CTkLabel(
            footer,
            text="提示：药名相近不等于同一种药，请核对规格、厂家后再合并。",
            font=T.font(8.8),
            text_color=T.TEXT_TERTIARY,
            anchor="w",
        ).pack(side="left")

        ctk.CTkButton(
            footer,
            text="取消",
            width=78,
            height=34,
            corner_radius=8,
            fg_color="transparent",
            hover_color=T.SURFACE_HOVER,
            border_width=1,
            border_color=T.BORDER,
            text_color=T.TEXT_SECONDARY,
            font=T.font(9.6),
            command=self._close,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            footer,
            text="应用选中合并",
            width=118,
            height=34,
            corner_radius=8,
            fg_color=T.ACCENT,
            hover_color=T.ACCENT_HOVER,
            text_color=T.ON_ACCENT,
            font=T.font(9.6, "bold"),
            command=self._apply,
        ).pack(side="right")
        if self.current_aliases:
            ctk.CTkButton(
                footer,
                text="清除已合并",
                width=94,
                height=34,
                corner_radius=8,
                fg_color="transparent",
                hover_color=T.DANGER_SOFT,
                border_width=1,
                border_color=T.BORDER,
                text_color=T.DANGER,
                font=T.font(9.6),
                command=self._clear_aliases,
            ).pack(side="right", padx=(8, 0))

        self.protocol("WM_DELETE_WINDOW", self._close)
        self.bind("<Escape>", lambda _e: self._close())
        try:
            self.iconbitmap(T.resource_path("assets", "app.ico"))
        except Exception:  # noqa: BLE001
            pass
        self.after(40, self._centre_and_focus)

    def _centre_and_focus(self) -> None:
        try:
            self.update_idletasks()
            scale = ctk.ScalingTracker.get_window_scaling(self)
            width = self.winfo_width()
            height = self.winfo_height()
            parent_x = self.parent.winfo_rootx()
            parent_y = self.parent.winfo_rooty()
            parent_w = self.parent.winfo_width()
            parent_h = self.parent.winfo_height()
            x = int((parent_x + (parent_w - width) / 2) / scale)
            y = int((parent_y + (parent_h - height) / 2) / scale)
            self.geometry(f"+{x}+{y}")
            self.grab_set()
            self.lift()
            self.focus_force()
        except Exception:  # noqa: BLE001
            pass

    def _apply(self) -> None:
        aliases: dict[str, str] = dict(self.current_aliases)
        for variable, group in self._checks:
            if not variable.get():
                continue
            canonical = max(
                group.variants,
                key=lambda variant: (variant.count, len(variant.key)),
            )
            for variant in group.variants:
                if variant.key != canonical.key:
                    aliases[variant.key] = canonical.key
        self.result = aliases
        self._close()

    def _clear_aliases(self) -> None:
        self.result = {}
        self._close()

    def _close(self) -> None:
        try:
            self.grab_release()
        except Exception:  # noqa: BLE001
            pass
        self.destroy()
