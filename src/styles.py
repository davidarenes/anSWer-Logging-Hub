# styles.py
"""
Centralized styling for customtkinter apps.

- Single source of truth for palette, fonts, radii, spacing and sizes.
- Helpers to style common widgets (buttons, labels, entries, cards).
- Titlebar constants for frameless window UIs.
- Dark-mode first; can extend to light mode later if needed.

Usage:
    import customtkinter as ctk
    import styles

    styles.apply_global(appearance="dark", color_theme="dark-blue")
    app = ctk.CTk()
    app.configure(bg=styles.Palette.BG)

    card = styles.card(app)  # styled CTkFrame
    btn  = ctk.CTkButton(card, text="OK")
    styles.style_button(btn, variant="success")
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Literal, Tuple
import customtkinter as ctk

# ---------- Palette ----------
@dataclass(frozen=True)
class Palette:
    # Base background and layered surfaces
    BG: str = "#050b18"
    CARD_DARK: str = "#0f1a2c"
    CARD_DARK_ALT: str = "#0c1524"
    CARD_BORDER: str = "#1e2b40"

    # Typography colors
    TEXT: str = "#f8fafc"
    MUTED: str = "#94a3b8"

    # Accents
    PRIMARY: str = "#38bdf8"
    SUCCESS: str = "#4ade80"
    SUCCESS_HOVER: str = "#22c55e"
    WARNING: str = "#fbbf24"
    DANGER: str = "#fb7185"
    DANGER_HOVER: str = "#f43f5e"
    NEUTRAL: str = "#334155"
    NEUTRAL_HOVER: str = "#475569"

    # Status pill backgrounds
    STATUS_MUTED_BG: str = "#16233b"
    STATUS_SUCCESS_BG: str = "#102d21"
    STATUS_WARNING_BG: str = "#332611"
    STATUS_DANGER_BG: str = "#331420"
    STATUS_INFO_BG: str = "#10253f"

    # Bright status accents (used for sysvar health)
    CHILL_GREEN_BG: str = "#006b3f"
    CHILL_GREEN_TEXT: str = "#7dffb3"
    CHILL_RED_BG: str = "#7a0020"
    CHILL_RED_TEXT: str = "#ff8aa0"

    # Titlebar text (legacy frameless UI hooks)
    TITLEBAR_TEXT: str = "#ffffff"


# ---------- Typography ----------
@dataclass(frozen=True)
class Fonts:
    TITLE: Tuple[str, int, str] = ("Segoe UI", 22, "bold")
    SECTION: Tuple[str, int, str] = ("Segoe UI Semibold", 13, "normal")
    HINT: Tuple[str, int] = ("Segoe UI", 10)
    BODY: Tuple[str, int] = ("Segoe UI", 10)
    BODY_BOLD: Tuple[str, int, str] = ("Segoe UI Semibold", 10, "bold")
    BUTTON: Tuple[str, int, str] = ("Segoe UI Semibold", 10, "normal")
    CAPTION: Tuple[str, int] = ("Segoe UI", 8)


# ---------- Spacing & Radii ----------
@dataclass(frozen=True)
class Metrics:
    RADIUS_SM: int = 8
    RADIUS_MD: int = 12
    RADIUS_LG: int = 18

    PAD_X: int = 10
    PAD_Y: int = 8

    BTN_H: int = 30
    BTN_H_LG: int = 34

    TITLEBAR_H: int = 36


# ---------- Apply global theme ----------
def apply_global(*, appearance: Literal["dark", "light"] = "dark",
                 color_theme: str = "dark-blue") -> None:
    """
    Configure customtkinter global appearance & theme.
    """
    ctk.set_appearance_mode(appearance)
    ctk.set_default_color_theme(color_theme)


# ---------- Cards / Frames ----------
def card(master, *, corner_radius: int | None = None):
    """
    Create a styled CTkFrame representing a 'card' surface.
    """
    cr = corner_radius if corner_radius is not None else Metrics.RADIUS_LG
    return ctk.CTkFrame(
        master,
        corner_radius=cr,
        fg_color=(Palette.CARD_DARK, Palette.CARD_DARK_ALT),
        border_width=1,
        border_color=Palette.CARD_BORDER,
    )


# ---------- Buttons ----------
def style_button(btn: ctk.CTkButton,
                 *,
                 variant: Literal["primary", "success", "danger", "neutral"] = "primary",
                 size: Literal["sm", "md", "lg"] = "md",
                 roundness: Literal["sm", "md", "lg"] = "md") -> None:
    """
    Apply consistent styling to a CTkButton.
    """
    height = { "sm": 30, "md": Metrics.BTN_H, "lg": Metrics.BTN_H_LG }[size]
    radius = {
        "sm": Metrics.RADIUS_SM,
        "md": Metrics.RADIUS_MD,
        "lg": Metrics.RADIUS_LG
    }[roundness]

    if variant == "primary":
        fg = Palette.PRIMARY
        hover = _darken_hex(Palette.PRIMARY, 0.12)
        text = Palette.TEXT
    elif variant == "success":
        fg = Palette.SUCCESS
        hover = Palette.SUCCESS_HOVER
        text = Palette.TEXT
    elif variant == "danger":
        fg = Palette.DANGER
        hover = Palette.DANGER_HOVER
        text = Palette.TEXT
    else:  # neutral
        fg = Palette.NEUTRAL
        hover = Palette.NEUTRAL_HOVER
        text = Palette.TEXT

    btn.configure(
        height=height,
        corner_radius=radius,
        fg_color=fg,
        hover_color=hover,
        text_color=text,
        font=Fonts.BUTTON,
    )


# ---------- Labels ----------
def style_label(label: ctk.CTkLabel,
                *,
                kind: Literal["title", "section", "body", "hint", "caption"] = "body") -> None:
    """
    Apply text style to labels.
    """
    if kind == "title":
        label.configure(font=Fonts.TITLE, text_color=Palette.TEXT)
    elif kind == "section":
        label.configure(font=Fonts.SECTION, text_color=Palette.TEXT)
    elif kind == "hint":
        label.configure(font=Fonts.HINT, text_color=Palette.MUTED)
    elif kind == "caption":
        label.configure(font=Fonts.CAPTION, text_color=Palette.MUTED)
    else:
        label.configure(font=Fonts.BODY, text_color=Palette.TEXT)


# ---------- Entries ----------
def style_entry(entry: ctk.CTkEntry, *,
                roundness: Literal["sm", "md", "lg"] = "md") -> None:
    """
    Apply consistent look to text inputs.
    """
    radius = {
        "sm": Metrics.RADIUS_SM,
        "md": Metrics.RADIUS_MD,
        "lg": Metrics.RADIUS_LG
    }[roundness]

    entry.configure(
        corner_radius=radius,
        height=30,
        font=Fonts.BODY,
        text_color=Palette.TEXT,
        placeholder_text_color=("gray85", "gray80"),
    )


def style_textbox(textbox: ctk.CTkTextbox, *,
                  roundness: Literal["sm", "md", "lg"] = "lg") -> None:
    """
    Give multi-line text areas the same minimal look as the other inputs.
    """
    radius = {
        "sm": Metrics.RADIUS_SM,
        "md": Metrics.RADIUS_MD,
        "lg": Metrics.RADIUS_LG
    }[roundness]

    textbox.configure(
        corner_radius=radius,
        border_width=1,
        border_color=Palette.CARD_BORDER,
        fg_color=(Palette.CARD_DARK, Palette.CARD_DARK),
        text_color=Palette.TEXT,
        font=Fonts.BODY,
        scrollbar_button_color=Palette.NEUTRAL,
        scrollbar_button_hover_color=Palette.NEUTRAL_HOVER,
    )


def style_option_menu(option_menu: ctk.CTkOptionMenu, *,
                      roundness: Literal["sm", "md", "lg"] = "md") -> None:
    """
    Make dropdowns visually match text inputs.
    """
    radius = {
        "sm": Metrics.RADIUS_SM,
        "md": Metrics.RADIUS_MD,
        "lg": Metrics.RADIUS_LG
    }[roundness]

    entry_theme = getattr(ctk.ThemeManager, "theme", {}).get("CTkEntry", {})
    entry_fg = entry_theme.get("fg_color", ("#1f1f1f", "#141414"))
    entry_text = entry_theme.get("text_color", Palette.TEXT)
    dropdown_fg = entry_theme.get("button_color", entry_fg)

    option_menu.configure(
        corner_radius=radius,
        height=34,
        font=Fonts.BODY,
        text_color=entry_text,
        dropdown_text_color=entry_text,
        dropdown_font=Fonts.BODY,
        fg_color=entry_fg,
        button_color=dropdown_fg,
        button_hover_color=Palette.NEUTRAL_HOVER,
    )

# ---------- Titlebar helpers ----------
def titlebar_frame(master):
    """
    Titlebar container with proper height and background color.
    """
    return ctk.CTkFrame(
        master,
        height=Metrics.TITLEBAR_H,
        corner_radius=0,
        fg_color=Palette.BG,
    )


def style_titlebar_button(btn: ctk.CTkButton, *,
                          kind: Literal["minimize", "close", "neutral"] = "neutral") -> None:
    """
    Style window control buttons in the custom titlebar.
    """
    if kind == "close":
        fg = Palette.DANGER
        hover = Palette.DANGER_HOVER
    elif kind == "minimize":
        fg = Palette.NEUTRAL
        hover = Palette.NEUTRAL_HOVER
    else:
        fg = Palette.NEUTRAL
        hover = Palette.NEUTRAL_HOVER

    btn.configure(
        width=32, height=24,
        corner_radius=Metrics.RADIUS_SM,
        fg_color=fg,
        hover_color=hover,
        text_color=Palette.TEXT,
        font=Fonts.BUTTON,
    )


# ---------- Utility: hex color tweak ----------
def _darken_hex(hex_color: str, amount: float) -> str:
    """
    Darken a hex color by `amount` (0..1). Example: 0.12 darkens by 12%.
    """
    hex_color = hex_color.lstrip("#")
    r = max(0, min(255, int(int(hex_color[0:2], 16) * (1 - amount))))
    g = max(0, min(255, int(int(hex_color[2:4], 16) * (1 - amount))))
    b = max(0, min(255, int(int(hex_color[4:6], 16) * (1 - amount))))
    return f"#{r:02x}{g:02x}{b:02x}"
