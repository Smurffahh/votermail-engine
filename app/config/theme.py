# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

from dataclasses import dataclass
from pathlib import Path
import sys
from typing import Literal
import customtkinter as ctk

from app.config.settings import FONTS_DIR

FONT_FAMILY_UI = "Inter"
FONT_FAMILY_MONO = "Space Mono"

# Native Dual-Tuple Color Tokens (LIGHT_HEX, DARK_HEX) for CustomTkinter
# Skema Warna: DaisyUI "garden" (Light Mode) & DaisyUI "night" (Dark Mode)
THEME: dict[str, tuple[str, str]] = {
    # Background & Surfaces (DaisyUI Garden | DaisyUI Night)
    "bg_root": ("#D4D2D2", "#090E1A"),          # Base-200 (Garden) | Base-200 (Night)
    "bg_card": ("#E9E7E7", "#0F172A"),          # Base-100 (Garden) | Base-100 (Night)
    "bg_surface": ("#F5F4F4", "#1E293B"),       # Lighter Base-100 / Inner | Slate 800
    "bg_elevated": ("#DFDDDD", "#162033"),      # Kotak Progress & Header Tabel
    "bg_hover": ("#D0CECE", "#26354D"),         # Hover highlight row

    # Borders & Dividers
    "border_card": ("#BFBDBD", "#2E3C56"),      # Hairline border struktural 1px kartu utama
    "border_surface": ("#B2AFAF", "#334460"),   # Border sub-kartu, tombol & input
    "border_subtle": ("#BFBDBD", "#2E3C56"),    # Alias kompatibilitas

    # Typography (WCAG AA Compliant)
    "text_primary": ("#100F0F", "#C8CAD0"),     # Deep near-black (Garden) | Muted soft white (Night)
    "text_secondary": ("#3D3A3A", "#94A3B8"),   # Slate secondary (Garden) | Slate 400 (Night)
    "text_muted": ("#6E6A6A", "#64748B"),       # Slate muted (Garden) | Slate 500 (Night)

    # Secondary Interactive Buttons (DaisyUI Garden Neutral/Rose | DaisyUI Night Indigo)
    "btn_secondary_bg": ("#DCDADA", "#1E293B"),
    "btn_secondary_hover": ("#CECBCB", "#334155"),
    "btn_secondary_border": ("#B2AFAF", "#475569"),
    "btn_secondary_text": ("#100F0F", "#E2E8F0"),

    # Soft Accent Buttons / Secondary CTA ("Pratinjau 1 Halaman" & "Pilih File")
    "btn_accent_soft_bg": ("#E8D6DD", "#1E1F3B"),      # Garden rose soft | Night indigo soft
    "btn_accent_soft_hover": ("#D9C0CB", "#2B2D54"),
    "btn_accent_soft_border": ("#8E4162", "#818CF8"),
    "btn_accent_soft_text": ("#8E4162", "#818CF8"),

    # Neutral / Utility Buttons ("Bersihkan", "Ubah Folder")
    "btn_neutral_bg": ("#291E00", "#1E293B"),
    "btn_neutral_hover": ("#3D2D00", "#334155"),
    "btn_neutral_border": ("#1A1300", "#334155"),
    "btn_neutral_text": ("#E9E7E7", "#CDD0D6"),

    # Primary CTA Button ("Mulai Buat Surat PDF")
    "btn_primary_bg": ("#FE0075", "#38BDF8"),          # Garden primary rose-pink | Night primary sky blue
    "btn_primary_hover": ("#D90064", "#0EA5E9"),
    "btn_primary_text": ("#FFFFFF", "#001019"),

    # Cancel / Danger Button
    "btn_danger_bg": ("#FFE4E6", "#311018"),
    "btn_danger_hover": ("#FECDD3", "#4A1824"),
    "btn_danger_text": ("#C81E2B", "#FB7185"),
    "btn_disabled_bg": ("#DCDADA", "#162033"),
    "btn_disabled_text": ("#8F8B8B", "#4B5563"),

    # Status Badges & Accents
    "status_ready_bg": ("#D1FAE5", "#0D2E2B"),
    "status_ready_text": ("#007A4F", "#2DD4BF"),
    "status_warn_bg": ("#FEF3C7", "#2E2410"),
    "status_warn_text": ("#926B00", "#F4BF50"),
    "status_error_bg": ("#FFE4E6", "#311018"),
    "status_error_text": ("#C81E2B", "#FB7185"),
    "status_info_bg": ("#E0F2FE", "#0C2538"),
    "status_info_text": ("#0077AD", "#38BDF8"),
    "status_neutral_bg": ("#DCDADA", "#1E293B"),
    "status_neutral_text": ("#575353", "#94A3B8"),
    "metric_value": ("#100F0F", "#C8CAD0"),
}

@dataclass(frozen=True)
class ThemeTokens:
    name: str = "Dark"
    corner_radius_card: int = 16
    corner_radius_inner: int = 8
    corner_radius_control: int = 8
    corner_radius_badge: int = 6
    font_family_main: str = FONT_FAMILY_UI
    font_family_mono: str = FONT_FAMILY_MONO

    def __getitem__(self, item: str) -> tuple[str, str]:
        if item in THEME:
            return THEME[item]
        raise KeyError(item)

    @property
    def bg_root(self) -> tuple[str, str]: return THEME["bg_root"]
    @property
    def bg_card(self) -> tuple[str, str]: return THEME["bg_card"]
    @property
    def bg_surface(self) -> tuple[str, str]: return THEME["bg_surface"]
    @property
    def bg_elevated(self) -> tuple[str, str]: return THEME["bg_elevated"]
    @property
    def bg_hover(self) -> tuple[str, str]: return THEME["bg_hover"]
    @property
    def border_card(self) -> tuple[str, str]: return THEME["border_card"]
    @property
    def border_surface(self) -> tuple[str, str]: return THEME["border_surface"]
    @property
    def border_subtle(self) -> tuple[str, str]: return THEME["border_card"]
    @property
    def border_inner(self) -> tuple[str, str]: return THEME["border_surface"]
    @property
    def text_primary(self) -> tuple[str, str]: return THEME["text_primary"]
    @property
    def text_secondary(self) -> tuple[str, str]: return THEME["text_secondary"]
    @property
    def text_muted(self) -> tuple[str, str]: return THEME["text_muted"]
    @property
    def primary_action(self) -> tuple[str, str]: return THEME["btn_primary_bg"]
    @property
    def primary_action_hover(self) -> tuple[str, str]: return THEME["btn_primary_hover"]
    @property
    def danger_action(self) -> tuple[str, str]: return THEME["btn_danger_bg"]
    @property
    def danger_action_hover(self) -> tuple[str, str]: return THEME["btn_danger_hover"]
    @property
    def status_ready(self) -> tuple[str, str]: return THEME["status_ready_text"]
    @property
    def status_ready_bg(self) -> tuple[str, str]: return THEME["status_ready_bg"]
    @property
    def status_warning(self) -> tuple[str, str]: return THEME["status_warn_text"]
    @property
    def status_warn_bg(self) -> tuple[str, str]: return THEME["status_warn_bg"]
    @property
    def status_error(self) -> tuple[str, str]: return THEME["status_error_text"]
    @property
    def status_error_bg(self) -> tuple[str, str]: return THEME["status_error_bg"]
    @property
    def status_neutral(self) -> tuple[str, str]: return THEME["status_neutral_text"]
    @property
    def status_neutral_bg(self) -> tuple[str, str]: return THEME["status_neutral_bg"]
    @property
    def status_info(self) -> tuple[str, str]: return THEME["status_info_text"]
    @property
    def status_info_text(self) -> tuple[str, str]: return THEME["status_info_text"]
    @property
    def status_info_bg(self) -> tuple[str, str]: return THEME["status_info_bg"]
    @property
    def btn_secondary_bg(self) -> tuple[str, str]: return THEME["btn_secondary_bg"]
    @property
    def btn_secondary_hover(self) -> tuple[str, str]: return THEME["btn_secondary_hover"]
    @property
    def btn_secondary_border(self) -> tuple[str, str]: return THEME["btn_secondary_border"]
    @property
    def btn_secondary_text(self) -> tuple[str, str]: return THEME["btn_secondary_text"]
    @property
    def btn_accent_soft_bg(self) -> tuple[str, str]: return THEME["btn_accent_soft_bg"]
    @property
    def btn_accent_soft_hover(self) -> tuple[str, str]: return THEME["btn_accent_soft_hover"]
    @property
    def btn_accent_soft_border(self) -> tuple[str, str]: return THEME["btn_accent_soft_border"]
    @property
    def btn_accent_soft_text(self) -> tuple[str, str]: return THEME["btn_accent_soft_text"]
    @property
    def btn_neutral_bg(self) -> tuple[str, str]: return THEME["btn_neutral_bg"]
    @property
    def btn_neutral_hover(self) -> tuple[str, str]: return THEME["btn_neutral_hover"]
    @property
    def btn_neutral_border(self) -> tuple[str, str]: return THEME["btn_neutral_border"]
    @property
    def btn_neutral_text(self) -> tuple[str, str]: return THEME["btn_neutral_text"]
    @property
    def btn_primary_bg(self) -> tuple[str, str]: return THEME["btn_primary_bg"]
    @property
    def btn_primary_hover(self) -> tuple[str, str]: return THEME["btn_primary_hover"]
    @property
    def btn_primary_text(self) -> tuple[str, str]: return THEME["btn_primary_text"]
    @property
    def btn_danger_bg(self) -> tuple[str, str]: return THEME["btn_danger_bg"]
    @property
    def btn_danger_hover(self) -> tuple[str, str]: return THEME["btn_danger_hover"]
    @property
    def btn_danger_text(self) -> tuple[str, str]: return THEME["btn_danger_text"]
    @property
    def btn_disabled_bg(self) -> tuple[str, str]: return THEME["btn_disabled_bg"]
    @property
    def btn_disabled_text(self) -> tuple[str, str]: return THEME["btn_disabled_text"]
    @property
    def metric_value(self) -> tuple[str, str]: return THEME["metric_value"]
    @property
    def metric_number(self) -> tuple[str, str]: return THEME["metric_value"]
    @property
    def bg_inner(self) -> tuple[str, str]: return THEME["bg_surface"]
    @property
    def disabled_bg(self) -> tuple[str, str]: return THEME["btn_disabled_bg"]
    @property
    def disabled_text(self) -> tuple[str, str]: return THEME["btn_disabled_text"]

DARK_THEME = ThemeTokens(name="Dark")
LIGHT_THEME = ThemeTokens(name="Light")

def get_theme(mode: Literal["Dark", "Light"] = "Dark") -> ThemeTokens:
    return DARK_THEME if str(mode).lower() == "dark" else LIGHT_THEME

_FONTS_LOADED = False

def load_bundled_fonts() -> int:
    """Memuat file font offline (.ttf) dari FONTS_DIR ke CustomTkinter dan Windows GDI secara privat (FR_PRIVATE)."""
    global _FONTS_LOADED
    if _FONTS_LOADED:
        return 0

    loaded_count = 0
    if not FONTS_DIR.exists():
        return 0

    for ttf_path in FONTS_DIR.glob("*.ttf"):
        try:
            if hasattr(ctk, "FontManager") and hasattr(ctk.FontManager, "load_font"):
                ctk.FontManager.load_font(str(ttf_path))

            if sys.platform == "win32":
                try:
                    import ctypes
                    ctypes.windll.gdi32.AddFontResourceExW(str(ttf_path.resolve()), 0x10, 0)
                except Exception:
                    pass
            loaded_count += 1
        except Exception:
            pass

    _FONTS_LOADED = True
    return loaded_count

def get_ui_font(size: int = 12, weight: str = "normal", slant: str = "roman") -> ctk.CTkFont:
    """Mengembalikan CTkFont untuk antarmuka pengguna berbasis font Inter."""
    return ctk.CTkFont(family=FONT_FAMILY_UI, size=size, weight=weight, slant=slant)

def get_mono_font(size: int = 12, weight: str = "normal", slant: str = "roman") -> ctk.CTkFont:
    """Mengembalikan CTkFont untuk log dan metrik tabular berbasis font Space Mono."""
    return ctk.CTkFont(family=FONT_FAMILY_MONO, size=size, weight=weight, slant=slant)
