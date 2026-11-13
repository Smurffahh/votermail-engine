# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

from dataclasses import dataclass
from typing import Literal

@dataclass(frozen=True)
class ThemeTokens:
    name: str
    bg_root: str
    bg_card: str
    bg_surface: str
    bg_hover: str
    border_subtle: str
    text_primary: str
    text_secondary: str
    text_muted: str
    primary_action: str
    primary_action_hover: str
    danger_action: str
    danger_action_hover: str
    status_ready: str
    status_warning: str
    status_error: str
    status_neutral: str
    btn_secondary_bg: str
    btn_secondary_hover: str
    btn_secondary_border: str
    btn_secondary_text: str
    corner_radius_card: int = 12
    corner_radius_inner: int = 8
    font_family_main: str = "Segoe UI"
    font_family_mono: str = "Consolas"

    # Compatibility aliases
    @property
    def bg_inner(self) -> str:
        return self.bg_surface

    @property
    def border_card(self) -> str:
        return self.border_subtle

    @property
    def border_inner(self) -> str:
        return self.border_subtle

    @property
    def metric_number(self) -> str:
        return self.text_primary

    @property
    def execute_action(self) -> str:
        return self.primary_action

    @property
    def execute_action_hover(self) -> str:
        return self.primary_action_hover

    @property
    def disabled_bg(self) -> str:
        return self.bg_surface

    @property
    def disabled_text(self) -> str:
        return self.text_muted

    @property
    def badge_ready_bg(self) -> str:
        return self.bg_surface

    @property
    def badge_ready_text(self) -> str:
        return self.status_ready

    @property
    def badge_process_bg(self) -> str:
        return self.bg_surface

    @property
    def badge_process_text(self) -> str:
        return self.status_warning

    @property
    def badge_amber_bg(self) -> str:
        return self.bg_surface

    @property
    def badge_amber_text(self) -> str:
        return self.status_warning

    @property
    def badge_error_bg(self) -> str:
        return self.bg_surface

    @property
    def badge_error_text(self) -> str:
        return self.status_error

    @property
    def badge_neutral_bg(self) -> str:
        return self.bg_surface

    @property
    def badge_neutral_text(self) -> str:
        return self.status_neutral

DARK_THEME = ThemeTokens(
    name="Dark",
    bg_root="#090D16",
    bg_card="#111726",
    bg_surface="#161F30",
    bg_hover="#1E293B",
    border_subtle="#1E293B",
    text_primary="#F8FAFC",
    text_secondary="#94A3B8",
    text_muted="#64748B",
    primary_action="#10B981",
    primary_action_hover="#059669",
    danger_action="#EF4444",
    danger_action_hover="#DC2626",
    status_ready="#34D399",
    status_warning="#FBBF24",
    status_error="#F87171",
    status_neutral="#94A3B8",
    btn_secondary_bg="#161F30",
    btn_secondary_hover="#1E293B",
    btn_secondary_border="#1E293B",
    btn_secondary_text="#F8FAFC",
    corner_radius_card=12,
    corner_radius_inner=8,
)

LIGHT_THEME = ThemeTokens(
    name="Light",
    bg_root="#F1F5F9",
    bg_card="#FFFFFF",
    bg_surface="#F8FAFC",
    bg_hover="#E2E8F0",
    border_subtle="#E2E8F0",
    text_primary="#0F172A",
    text_secondary="#475569",
    text_muted="#64748B",
    primary_action="#10B981",
    primary_action_hover="#059669",
    danger_action="#DC2626",
    danger_action_hover="#B91C1C",
    status_ready="#059669",
    status_warning="#D97706",
    status_error="#DC2626",
    status_neutral="#64748B",
    btn_secondary_bg="#FFFFFF",
    btn_secondary_hover="#E2E8F0",
    btn_secondary_border="#E2E8F0",
    btn_secondary_text="#0F172A",
    corner_radius_card=12,
    corner_radius_inner=8,
)

def get_theme(mode: Literal["Dark", "Light"] = "Dark") -> ThemeTokens:
    return DARK_THEME if mode.lower() == "dark" else LIGHT_THEME
