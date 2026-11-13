# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

from dataclasses import dataclass
from typing import Literal

# Native Dual-Tuple Color Tokens (LIGHT_HEX, DARK_HEX) for CustomTkinter
THEME: dict[str, tuple[str, str]] = {
    # Background & Surfaces
    "bg_root": ("#F1F5F9", "#090D16"),          # Slate 100 (Light) | Deep Night (Dark)
    "bg_card": ("#FFFFFF", "#111726"),          # Pure White (Light) | Slate 900 (Dark)
    "bg_surface": ("#F8FAFC", "#161F30"),       # Slate 50 (Light) | Inner Card Surface (Dark)
    "bg_elevated": ("#F1F5F9", "#1E293B"),      # Kotak Progress & Header Tabel
    "bg_hover": ("#E2E8F0", "#1E293B"),         # Hover highlight row

    # Borders & Dividers
    "border_card": ("#E2E8F0", "#1E293B"),      # Hairline border kartu utama
    "border_surface": ("#CBD5E1", "#26334D"),   # Border sub-kartu & input
    "border_subtle": ("#E2E8F0", "#1E293B"),    # Alias kompatibilitas

    # Typography (100% Terbaca Jelas di Kedua Mode - WCAG AA)
    "text_primary": ("#0F172A", "#F8FAFC"),     # Slate 900 Hitam Pekat (Light) | Putih Bersih (Dark)
    "text_secondary": ("#334155", "#94A3B8"),   # Slate 700 (Light) | Slate 400 (Dark)
    "text_muted": ("#64748B", "#64748B"),       # Slate 500

    # Secondary Interactive Buttons (Jelas Terlihat Sebagai Tombol!)
    "btn_secondary_bg": ("#EEF2F6", "#1E293B"),
    "btn_secondary_hover": ("#E2E8F0", "#334155"),
    "btn_secondary_border": ("#94A3B8", "#3B82F6"),
    "btn_secondary_text": ("#0F172A", "#F8FAFC"),

    # Soft Accent Buttons (Untuk Tombol "Pilih File" & "Pratinjau PDF")
    "btn_accent_soft_bg": ("#EFF6FF", "#172554"),
    "btn_accent_soft_hover": ("#DBEAFE", "#1E3A8A"),
    "btn_accent_soft_border": ("#60A5FA", "#2563EB"),
    "btn_accent_soft_text": ("#1D4ED8", "#93C5FD"),

    # Primary CTA Button ("Buat File PDF Sekarang")
    "btn_primary_bg": ("#10B981", "#10B981"),
    "btn_primary_hover": ("#059669", "#059669"),
    "btn_primary_text": ("#FFFFFF", "#FFFFFF"),

    # Cancel / Danger Button
    "btn_danger_bg": ("#FEE2E2", "#7F1D1D"),
    "btn_danger_hover": ("#FECACA", "#991B1B"),
    "btn_danger_text": ("#B91C1C", "#FCA5A5"),
    "btn_disabled_bg": ("#E2E8F0", "#161F30"),
    "btn_disabled_text": ("#94A3B8", "#475569"),

    # Status Badges & Accents
    "status_ready_bg": ("#ECFDF5", "#064E3B"),
    "status_ready_text": ("#047857", "#34D399"),
    "status_warn_bg": ("#FFFBEB", "#78350F"),
    "status_warn_text": ("#B45309", "#FBBF24"),
    "status_error_bg": ("#FEF2F2", "#7F1D1D"),
    "status_error_text": ("#DC2626", "#F87171"),
    "status_neutral_bg": ("#F1F5F9", "#1E293B"),
    "status_neutral_text": ("#475569", "#94A3B8"),
    "metric_value": ("#0F172A", "#F8FAFC"),
}

@dataclass(frozen=True)
class ThemeTokens:
    name: str = "Dark"
    corner_radius_card: int = 12
    corner_radius_inner: int = 8
    font_family_main: str = "Segoe UI"
    font_family_mono: str = "Consolas"

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
