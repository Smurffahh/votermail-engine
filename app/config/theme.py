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
    bg_inner: str
    border_card: str
    border_inner: str
    corner_radius_card: int
    corner_radius_inner: int
    text_primary: str
    metric_number: str
    text_secondary: str
    text_muted: str
    primary_action: str
    primary_action_hover: str
    execute_action: str
    execute_action_hover: str
    danger_action: str
    danger_action_hover: str
    disabled_bg: str
    disabled_text: str
    badge_ready_bg: str
    badge_ready_text: str
    badge_process_bg: str
    badge_process_text: str
    badge_error_bg: str
    badge_error_text: str
    badge_amber_bg: str
    badge_amber_text: str
    badge_neutral_bg: str
    badge_neutral_text: str
    font_family_main: str = "Segoe UI"
    font_family_mono: str = "Consolas"

DARK_THEME = ThemeTokens(
    name="Dark",
    bg_root="#0B0F19",
    bg_card="#111827",
    bg_inner="#090D16",
    border_card="#1F2937",
    border_inner="#1E293B",
    corner_radius_card=12,
    corner_radius_inner=8,
    text_primary="#F8FAFC",
    metric_number="#38BDF8",
    text_secondary="#94A3B8",
    text_muted="#64748B",
    primary_action="#0284C7",
    primary_action_hover="#0369A1",
    execute_action="#059669",
    execute_action_hover="#047857",
    danger_action="#DC2626",
    danger_action_hover="#B91C1C",
    disabled_bg="#1E293B",
    disabled_text="#475569",
    badge_ready_bg="#064E3B",
    badge_ready_text="#34D399",
    badge_process_bg="#78350F",
    badge_process_text="#FBBF24",
    badge_error_bg="#7F1D1D",
    badge_error_text="#F87171",
    badge_amber_bg="#F59E0B",
    badge_amber_text="#78350F",
    badge_neutral_bg="#1E293B",
    badge_neutral_text="#94A3B8",
)

LIGHT_THEME = ThemeTokens(
    name="Light",
    bg_root="#F8FAFC",
    bg_card="#FFFFFF",
    bg_inner="#F1F5F9",
    border_card="#E2E8F0",
    border_inner="#CBD5E1",
    corner_radius_card=12,
    corner_radius_inner=8,
    text_primary="#0F172A",
    metric_number="#0284C7",
    text_secondary="#475569",
    text_muted="#94A3B8",
    primary_action="#0284C7",
    primary_action_hover="#0369A1",
    execute_action="#059669",
    execute_action_hover="#047857",
    danger_action="#DC2626",
    danger_action_hover="#B91C1C",
    disabled_bg="#E2E8F0",
    disabled_text="#94A3B8",
    badge_ready_bg="#DCFCE7",
    badge_ready_text="#15803D",
    badge_process_bg="#FEF3C7",
    badge_process_text="#B45309",
    badge_error_bg="#FEE2E2",
    badge_error_text="#B91C1C",
    badge_amber_bg="#F59E0B",
    badge_amber_text="#78350F",
    badge_neutral_bg="#E2E8F0",
    badge_neutral_text="#475569",
)

def get_theme(mode: Literal["Dark", "Light"] = "Dark") -> ThemeTokens:
    return DARK_THEME if mode.lower() == "dark" else LIGHT_THEME
