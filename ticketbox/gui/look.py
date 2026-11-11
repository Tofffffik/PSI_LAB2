"""Цвета и шрифты окна. Цвет — пара (светлая тема, тёмная тема)."""

import customtkinter as ctk

ACCENT = ("#5B3CC4", "#7C5CFF")
ACCENT_HOVER = ("#4A2FA8", "#6A4AF0")
ON_ACCENT = "#FFFFFF"
ON_ACCENT_SOFT = "#E6E0FF"
MUTED = ("gray38", "gray66")
DANGER = ("#C62828", "#FF8A80")
NOTE_BG = ("#FFF1D6", "#4A3A12")
NOTE_TEXT = ("#7A4B00", "#FFD27A")
FIELD = ("#EEEAFB", "#2E2A3D")
FIELD_TEXT = ("gray10", "gray92")


def option_menu_colors() -> dict:
    """Выпадающие списки в цветах приложения, а не стандартные синие."""
    return dict(fg_color=FIELD, button_color=ACCENT, button_hover_color=ACCENT_HOVER, text_color=FIELD_TEXT,
                dropdown_hover_color=FIELD)


def setup() -> None:
    """Тема по системе; встроенная палитра «blue» для полей и списков."""
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")


def font(size: int = 14, bold: bool = False) -> ctk.CTkFont:
    return ctk.CTkFont(family="Segoe UI", size=size, weight="bold" if bold else "normal")
