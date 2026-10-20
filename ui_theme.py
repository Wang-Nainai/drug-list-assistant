"""界面配色、字体与通用小工具。"""

from __future__ import annotations

import os
import sys
import tkinter.font as tkfont

import customtkinter as ctk

# ---------------------------------------------------------------- colours ----
BG = ("#F3F3F3", "#1B1B1B")
SURFACE = ("#FFFFFF", "#272727")
SURFACE_ALT = ("#F9F9F9", "#2E2E2E")
SURFACE_HOVER = ("#F5F5F5", "#313131")
SURFACE_TINT = ("#F5FAFF", "#21313D")
BORDER = ("#E4E4E4", "#3B3B3B")
BORDER_STRONG = ("#C9C9C9", "#4E4E4E")
TEXT = ("#1B1B1B", "#F2F2F2")
TEXT_SECONDARY = ("#5D5D5D", "#BFBFBF")
TEXT_TERTIARY = ("#8C8C8C", "#949494")
ACCENT = ("#0F6CBD", "#4CC2FF")
ACCENT_HOVER = ("#0C5CA8", "#63CBFF")
ON_ACCENT = ("#FFFFFF", "#062A43")
ACCENT_SOFT = ("#E7F1FB", "#1D3344")
ACCENT_SOFT_HOVER = ("#DCEAF8", "#25425A")
SUCCESS = ("#0E7A4E", "#6CCB9B")
SUCCESS_SOFT = ("#E6F5EE", "#1E3A30")
WARNING = ("#B25E00", "#FFB366")
WARNING_SOFT = ("#FFF4E5", "#3D2E1A")
DANGER = ("#C42B1C", "#FF99A4")
DANGER_SOFT = ("#FCEAE8", "#3D2326")
TRACK = ("#E8E8E8", "#3A3A3A")
SEGMENT_BG = ("#ECECEC", "#242424")
SEGMENT_SELECTED = ("#FFFFFF", "#3A3A3A")
DISABLED_FILL = ("#E8E8E8", "#333333")
DISABLED_TEXT = ("#A0A0A0", "#6E6E6E")
SCROLLBAR = ("#CFCFCF", "#555555")
SCROLLBAR_HOVER = ("#B4B4B4", "#6E6E6E")

TABLE_HEADER_BG = ("#F7F7F7", "#2C2C2C")
TABLE_HEADER_TEXT = ("#5D5D5D", "#B5B5B5")
TABLE_ROW_ODD = ("#FAFAFA", "#242424")
TABLE_ROW_EVEN = ("#FFFFFF", "#2A2A2A")
TABLE_SELECTION = ("#E7F1FB", "#26435B")

FONT_FAMILY = "Microsoft YaHei UI"
ICON_FONT_FAMILY = "Segoe Fluent Icons"

# Glyphs from Segoe Fluent Icons / Segoe MDL2 Assets.
ICON_DOCUMENT = "\uE8A5"
ICON_FOLDER = "\uE8B7"
ICON_OPEN = "\uE8E5"
ICON_CHECK = "\uE73E"
ICON_WARNING = "\uE7BA"
ICON_ERROR = "\uE783"
ICON_INFO = "\uE946"
ICON_CLOSE = "\uE711"
ICON_UPLOAD = "\uE898"


def detect_fonts(root) -> None:
    global FONT_FAMILY, ICON_FONT_FAMILY
    try:
        families = set(tkfont.families(root))
    except Exception:  # noqa: BLE001
        families = set()

    for candidate in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI"):
        if candidate in families:
            FONT_FAMILY = candidate
            break

    for candidate in ("Segoe Fluent Icons", "Segoe MDL2 Assets"):
        if candidate in families:
            ICON_FONT_FAMILY = candidate
            break


def font(size: float, weight: str = "normal") -> ctk.CTkFont:
    return ctk.CTkFont(family=FONT_FAMILY, size=max(1, int(round(size))), weight=weight)


def icon_font(size: float) -> ctk.CTkFont:
    return ctk.CTkFont(family=ICON_FONT_FAMILY, size=max(1, int(round(size))))


def resolve(color) -> str:
    """把 (light, dark) 颜色对解析为当前外观下的颜色。"""
    if isinstance(color, (tuple, list)):
        return color[1] if ctk.get_appearance_mode() == "Dark" else color[0]
    return color


def resource_path(*parts: str) -> str:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, *parts)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), *parts)


def display_units(text: str) -> int:
    return sum(2 if ord(char) > 127 else 1 for char in str(text))


def ellipsize(text: str, max_chars: int, mode: str = "middle") -> str:
    text = str(text)
    if max_chars <= 1 or len(text) <= max_chars:
        return text
    if mode == "end":
        return text[: max_chars - 1] + "…"
    keep = max_chars - 1
    left = (keep + 1) // 2
    right = keep - left
    return text[:left] + "…" + (text[-right:] if right else "")


def ellipsize_units(text: str, max_units: int, mode: str = "middle") -> str:
    """按显示宽度截断，中文按 2 个单位计算。"""
    text = str(text)
    if max_units <= 2 or display_units(text) <= max_units:
        return text
    if mode == "end":
        result = ""
        for char in text:
            if display_units(result + char + "…") > max_units:
                break
            result += char
        return result + "…"

    budget = max_units - 1
    left_budget = budget // 2
    right_budget = budget - left_budget
    left = ""
    for char in text:
        if display_units(left + char) > left_budget:
            break
        left += char
    right = ""
    for char in reversed(text):
        if display_units(char + right) > right_budget:
            break
        right = char + right
    return left + "…" + right


def format_bytes(size: float) -> str:
    units = ["B", "KB", "MB", "GB"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{value:.1f} GB"


def format_int(value: int) -> str:
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return str(value)
