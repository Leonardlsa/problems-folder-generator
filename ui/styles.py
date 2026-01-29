"""QSS styles for 错题集生成器 v3.0 - Dark and Light themes"""

import os

# Get the icons directory path
_ICON_DIR = os.path.join(os.path.dirname(__file__), "icons")
_CHEVRON_DARK_PATH = os.path.join(_ICON_DIR, "chevron.svg").replace("\\", "/")


def _base_theme(bg_color: str, fg_color: str, border_color: str,
                accent_color: str, placeholder_color: str, hover_bg: str,
                chevron_path: str) -> str:
    """Generate theme QSS with given colors."""
    return f"""
QMainWindow {{
    background-color: {bg_color};
}}

QWidget {{
    background-color: {bg_color};
    color: {fg_color};
    font-family: "Microsoft YaHei UI", "Noto Sans CJK SC", sans-serif;
}}

QLabel {{
    color: {fg_color};
    font-size: 14px;
    background-color: transparent;
}}

QLabel#titleLabel {{
    font-size: 24px;
    font-weight: bold;
}}

QLabel#authorLabel {{
    font-size: 12px;
    color: {placeholder_color};
}}

QLineEdit {{
    background-color: transparent;
    border: 2px solid {border_color};
    border-radius: 8px;
    color: {fg_color};
    padding: 10px 14px;
    font-size: 14px;
    min-height: 22px;
}}

QLineEdit:focus {{
    border-color: {accent_color};
}}

QLineEdit:read-only {{
    color: {placeholder_color};
}}

QComboBox {{
    background-color: transparent;
    border: 2px solid {border_color};
    border-radius: 8px;
    color: {fg_color};
    padding: 10px 14px;
    font-size: 14px;
    min-height: 22px;
}}

QComboBox:focus {{
    border-color: {accent_color};
}}

QComboBox::drop-down {{
    border: none;
    width: 36px;
    padding-right: 8px;
}}

QComboBox::down-arrow {{
    image: url({chevron_path});
    width: 18px;
    height: 18px;
}}

QComboBox QAbstractItemView {{
    background-color: {bg_color};
    border: 2px solid {border_color};
    border-radius: 4px;
    color: {fg_color};
    selection-background-color: {accent_color};
    selection-color: #ffffff;
    outline: none;
}}

QComboBox QAbstractItemView::item {{
    padding: 10px 14px;
    min-height: 28px;
}}

QComboBox QAbstractItemView::item:hover {{
    background-color: {hover_bg};
}}

QPushButton {{
    background-color: transparent;
    border: 2px solid {border_color};
    border-radius: 8px;
    color: {fg_color};
    padding: 12px 24px;
    font-size: 14px;
    min-width: 80px;
}}

QPushButton:hover {{
    background-color: {hover_bg};
}}

QPushButton:pressed {{
    background-color: {hover_bg};
}}

QPushButton#goButton {{
    font-size: 16px;
    font-weight: bold;
    min-width: 120px;
}}

QPushButton#resetButton {{
    font-size: 16px;
    min-width: 120px;
}}

QPushButton#infoButton, QPushButton#themeButton {{
    border-radius: 17px;
    min-width: 34px;
    max-width: 34px;
    min-height: 34px;
    max-height: 34px;
    padding: 0px;
    font-size: 18px;
    font-weight: bold;
}}

QPushButton#browseButton {{
    border-radius: 8px;
    min-width: 44px;
    max-width: 44px;
    min-height: 44px;
    max-height: 44px;
    padding: 0px;
}}

QMessageBox {{
    background-color: {bg_color};
}}

QMessageBox QLabel {{
    color: {fg_color};
}}

QMessageBox QPushButton {{
    min-width: 80px;
}}
"""


def get_dark_theme() -> str:
    """Return the dark theme QSS."""
    return _base_theme(
        bg_color="#1a1a1a",
        fg_color="#ffffff",
        border_color="#ffffff",
        accent_color="#4a9eff",
        placeholder_color="#888888",
        hover_bg="rgba(255, 255, 255, 0.1)",
        chevron_path=_CHEVRON_DARK_PATH
    )


def get_light_theme() -> str:
    """Return the light theme QSS."""
    # Create light chevron path
    chevron_light = os.path.join(_ICON_DIR, "chevron_light.svg").replace("\\", "/")
    return _base_theme(
        bg_color="#f5f5f5",
        fg_color="#1a1a1a",
        border_color="#1a1a1a",
        accent_color="#0066cc",
        placeholder_color="#666666",
        hover_bg="rgba(0, 0, 0, 0.08)",
        chevron_path=chevron_light
    )


# For backward compatibility
DARK_THEME = get_dark_theme()
LIGHT_THEME = get_light_theme()


# Style for digit input boxes
def get_digit_style(fg_color: str, accent_color: str) -> str:
    return f"""
QLineEdit {{
    background-color: transparent;
    border: none;
    border-bottom: 2px solid {fg_color};
    border-radius: 0px;
    color: {fg_color};
    padding: 4px;
    font-size: 18px;
    min-width: 24px;
    max-width: 24px;
    min-height: 24px;
    text-align: center;
}}

QLineEdit:focus {{
    border-bottom-color: {accent_color};
}}
"""


DIGIT_INPUT_STYLE_DARK = get_digit_style("#ffffff", "#4a9eff")
DIGIT_INPUT_STYLE_LIGHT = get_digit_style("#1a1a1a", "#0066cc")

# Default
DIGIT_INPUT_STYLE = DIGIT_INPUT_STYLE_DARK
