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
    BG: str = "#081120"
    CARD_DARK: str = "#0f1a2d"
    CARD_DARK_ALT: str = "#0b1526"
    CARD_BORDER: str = "#1f2f47"
    INPUT_BG: str = "#121f34"
    INPUT_BORDER: str = "#253850"

    # Typography colors
    TEXT: str = "#e9f0fb"
    MUTED: str = "#8fa2be"

    # Accents
    PRIMARY: str = "#4cc7ff"
    SUCCESS: str = "#3fdc86"
    SUCCESS_HOVER: str = "#23c76b"
    WARNING: str = "#f6c453"
    DANGER: str = "#ff6b81"
    DANGER_HOVER: str = "#f4506c"
    NEUTRAL: str = "#3b4a63"
    NEUTRAL_HOVER: str = "#4a5c79"

    # Status pill backgrounds
    STATUS_MUTED_BG: str = "#19243a"
    STATUS_SUCCESS_BG: str = "#123428"
    STATUS_WARNING_BG: str = "#3a2c16"
    STATUS_DANGER_BG: str = "#3a1826"
    STATUS_INFO_BG: str = "#13273f"

    # Bright status accents (used for sysvar health)
    CHILL_GREEN_BG: str = "#0a5a3b"
    CHILL_GREEN_TEXT: str = "#7bffb6"
    CHILL_RED_BG: str = "#651225"
    CHILL_RED_TEXT: str = "#ff95a8"

    # Full-app themes (OK / NOK)
    OK_BG: str = "#052c1d"
    OK_CARD: str = "#073c29"
    OK_CARD_ALT: str = "#063322"
    OK_BORDER: str = "#0d6a43"
    OK_INNER: str = "#0f5a3d"

    NOK_BG: str = "#330c17"
    NOK_CARD: str = "#4a1224"
    NOK_CARD_ALT: str = "#40101f"
    NOK_BORDER: str = "#7b1f36"
    NOK_INNER: str = "#63152f"

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
        fg_color=(Palette.INPUT_BG, Palette.INPUT_BG),
        border_color=Palette.INPUT_BORDER,
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
        border_color=Palette.INPUT_BORDER,
        fg_color=(Palette.INPUT_BG, Palette.INPUT_BG),
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

    option_menu.configure(
        corner_radius=radius,
        height=34,
        font=Fonts.BODY,
        text_color=Palette.TEXT,
        dropdown_text_color=Palette.TEXT,
        dropdown_fg_color=Palette.CARD_DARK,
        dropdown_font=Fonts.BODY,
        fg_color=(Palette.INPUT_BG, Palette.INPUT_BG),
        button_color=Palette.INPUT_BG,
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
