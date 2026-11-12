# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import importlib.metadata
import sys
from typing import Callable, Literal
import webbrowser
import customtkinter as ctk

from app.config.theme import ThemeTokens
from app.config.version import (
    __app_name__,
    __author__,
    __github_user__,
    __repository__,
    __license__,
    __version__,
    FEATURE_REGISTRY,
)

class MetricCard(ctk.CTkFrame):
    def __init__(
        self,
        master: any,
        title: str,
        value: str,
        theme: ThemeTokens,
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            fg_color=theme.bg_inner,
            border_color=theme.border_inner,
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            **kwargs,
        )
        self.theme = theme

        self.title_label = ctk.CTkLabel(
            self,
            text=title.upper(),
            font=ctk.CTkFont(family=theme.font_family_main, size=10, weight="bold"),
            text_color=theme.text_secondary,
        )
        self.title_label.pack(anchor="w", padx=10, pady=(8, 2))

        self.value_label = ctk.CTkLabel(
            self,
            text=value,
            font=ctk.CTkFont(family=theme.font_family_mono, size=20, weight="bold"),
            text_color=theme.metric_number,
        )
        self.value_label.pack(anchor="w", padx=10, pady=(0, 8))

    def set_value(self, value: str) -> None:
        self.value_label.configure(text=value)

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.configure(fg_color=theme.bg_inner, border_color=theme.border_inner)
        self.title_label.configure(text_color=theme.text_secondary)
        self.value_label.configure(text_color=theme.metric_number)

class StatusBadge(ctk.CTkFrame):
    def __init__(
        self,
        master: any,
        text: str,
        status: Literal["ready", "process", "error", "neutral", "amber"],
        theme: ThemeTokens,
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            corner_radius=6,
            height=26,
            **kwargs,
        )
        self.theme = theme
        self.label = ctk.CTkLabel(
            self,
            text=text,
            font=ctk.CTkFont(family=theme.font_family_mono, size=11, weight="bold"),
        )
        self.label.pack(padx=8, pady=3)
        self.set_status(text, status)

    def set_status(self, text: str, status: Literal["ready", "process", "error", "neutral", "amber"]) -> None:
        if status == "ready":
            bg, fg = self.theme.badge_ready_bg, self.theme.badge_ready_text
        elif status == "process":
            bg, fg = self.theme.badge_process_bg, self.theme.badge_process_text
        elif status == "amber":
            bg, fg = self.theme.badge_amber_bg, self.theme.badge_amber_text
        elif status == "error":
            bg, fg = self.theme.badge_error_bg, self.theme.badge_error_text
        else:
            bg, fg = self.theme.badge_neutral_bg, self.theme.badge_neutral_text

        self.configure(fg_color=bg)
        self.label.configure(text=text, text_color=fg)

    def update_theme(self, theme: ThemeTokens, current_status: str = "neutral") -> None:
        self.theme = theme
        self.set_status(self.label.cget("text"), current_status)

class KadusItemRow(ctk.CTkFrame):
    def __init__(
        self,
        master: any,
        sheet_name: str,
        wilayah_text: str,
        voter_count: int,
        page_count: int,
        is_selected: bool,
        theme: ThemeTokens,
        on_toggle: Callable[[], None] | None = None,
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            fg_color="transparent",
            height=34,
            **kwargs,
        )
        self.theme = theme
        self.sheet_name = sheet_name
        self.voter_count = voter_count
        self.page_count = page_count

        self.checkbox_var = ctk.BooleanVar(value=is_selected)

        clean_sheet = sheet_name.strip()
        clean_wilayah = wilayah_text.strip()
        if clean_sheet.lower() == clean_wilayah.lower():
            display_title = clean_sheet
        elif clean_wilayah.lower().startswith("kadus") and clean_sheet.lower().startswith("kadus"):
            display_title = clean_sheet
        elif clean_wilayah:
            display_title = f"{clean_sheet} — {clean_wilayah}"
        else:
            display_title = clean_sheet

        self.checkbox = ctk.CTkCheckBox(
            self,
            text=display_title,
            variable=self.checkbox_var,
            command=on_toggle,
            font=ctk.CTkFont(family=theme.font_family_main, size=12, weight="bold"),
            fg_color=theme.primary_action,
            hover_color=theme.primary_action_hover,
            border_color=theme.border_inner,
            corner_radius=4,
            width=220,
        )
        self.checkbox.pack(side="left", padx=(8, 4))

        self.info_label = ctk.CTkLabel(
            self,
            text=f"{voter_count} Pemilih • {page_count} Lembar",
            font=ctk.CTkFont(family=theme.font_family_mono, size=11),
            text_color=theme.text_secondary,
        )
        self.info_label.pack(side="left", padx=6)

        self.status_label = ctk.CTkLabel(
            self,
            text="[SIAP]",
            font=ctk.CTkFont(family=theme.font_family_mono, size=11, weight="bold"),
            text_color=theme.badge_ready_text,
        )
        self.status_label.pack(side="right", padx=8)

    def set_printing_active(self) -> None:
        self.status_label.configure(
            text="[SEDANG DICETAK...]",
            text_color=self.theme.badge_process_text,
        )

    def set_completed(self) -> None:
        self.status_label.configure(
            text="[SELESAI ✓]",
            text_color=self.theme.badge_ready_text,
        )

    def set_error(self, err_text: str = "[ERROR]") -> None:
        self.status_label.configure(
            text=err_text,
            text_color=self.theme.badge_error_text,
        )

    def reset_status(self) -> None:
        self.status_label.configure(
            text="[SIAP]",
            text_color=self.theme.badge_ready_text,
        )

    def set_enabled(self, enabled: bool) -> None:
        self.checkbox.configure(state="normal" if enabled else "disabled")

class AboutModal(ctk.CTkToplevel):
    def __init__(self, master: any, theme: ThemeTokens) -> None:
        super().__init__(master)
        self.theme = theme

        self.title("Tentang VoterMail Engine")
        self.geometry("640x660")
        self.minsize(580, 520)
        self.configure(fg_color=theme.bg_root)

        self.transient(master)
        self.grab_set()

        # Center on master
        self.update_idletasks()
        try:
            mx = master.winfo_x()
            my = master.winfo_y()
            mw = master.winfo_width()
            mh = master.winfo_height()
            x = mx + (mw - 640) // 2
            y = my + (mh - 660) // 2
            self.geometry(f"640x660+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        self._build_ui()

    def _build_ui(self) -> None:
        content_frame = ctk.CTkFrame(
            self,
            fg_color=self.theme.bg_card,
            border_color=self.theme.border_card,
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        content_frame.pack(fill="both", expand=True, padx=16, pady=16)

        # Header Section
        header_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=16, pady=(16, 10))

        title_row = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_row.pack(fill="x")

        app_title = ctk.CTkLabel(
            title_row,
            text=__app_name__,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=20, weight="bold"),
            text_color=self.theme.text_primary,
        )
        app_title.pack(side="left")

        version_badge = StatusBadge(
            title_row,
            text=f"v{__version__}",
            status="ready",
            theme=self.theme,
        )
        version_badge.pack(side="left", padx=10)

        app_desc = ctk.CTkLabel(
            header_frame,
            text="High-Volume Word & Excel DPT Mail-Merge Automation for Local Elections",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=self.theme.text_secondary,
        )
        app_desc.pack(anchor="w", pady=(2, 0))

        # Author & Credit Card
        dev_card = ctk.CTkFrame(
            content_frame,
            fg_color=self.theme.bg_inner,
            border_color=self.theme.border_inner,
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
        )
        dev_card.pack(fill="x", padx=16, pady=6)

        dev_row = ctk.CTkFrame(dev_card, fg_color="transparent")
        dev_row.pack(fill="x", padx=12, pady=10)

        dev_info_left = ctk.CTkFrame(dev_row, fg_color="transparent")
        dev_info_left.pack(side="left", fill="both", expand=True)

        dev_title = ctk.CTkLabel(
            dev_info_left,
            text=f"Developer & Creator: {__author__} (GitHub: @{__github_user__})",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=self.theme.text_primary,
        )
        dev_title.pack(anchor="w")

        lic_label = ctk.CTkLabel(
            dev_info_left,
            text=__license__,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
            text_color=self.theme.text_muted,
        )
        lic_label.pack(anchor="w", pady=(2, 0))

        btn_github = ctk.CTkButton(
            dev_row,
            text="Kunjungi GitHub",
            command=lambda: webbrowser.open(__repository__),
            fg_color=self.theme.primary_action,
            hover_color=self.theme.primary_action_hover,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            height=28,
            corner_radius=6,
        )
        btn_github.pack(side="right", padx=(8, 0))

        # System & Dependencies Section
        dep_card = ctk.CTkFrame(
            content_frame,
            fg_color=self.theme.bg_inner,
            border_color=self.theme.border_inner,
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
        )
        dep_card.pack(fill="x", padx=16, pady=6)

        dep_title = ctk.CTkLabel(
            dep_card,
            text="STATUS KOMPONEN & DEPENDENSI SISTEM",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=self.theme.text_secondary,
        )
        dep_title.pack(anchor="w", padx=12, pady=(8, 4))

        dep_grid = ctk.CTkFrame(dep_card, fg_color="transparent")
        dep_grid.pack(fill="x", padx=12, pady=(0, 8))

        packages_to_check = [
            ("Python Runtime", sys.version.split()[0]),
            ("CustomTkinter", self._get_pkg_version("customtkinter")),
            ("Pandas", self._get_pkg_version("pandas")),
            ("OpenPyXL", self._get_pkg_version("openpyxl")),
            ("Python-Docx", self._get_pkg_version("python-docx")),
            ("PyPDF", self._get_pkg_version("pypdf")),
            ("Pydantic", self._get_pkg_version("pydantic")),
            ("PyWin32 (COM)", self._get_pkg_version("pywin32")),
        ]

        for idx, (pkg_name, pkg_ver) in enumerate(packages_to_check):
            row = idx // 2
            col = idx % 2
            item_lbl = ctk.CTkLabel(
                dep_grid,
                text=f"• {pkg_name}: {pkg_ver}",
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=10),
                text_color=self.theme.text_primary,
                anchor="w",
            )
            item_lbl.grid(row=row, column=col, sticky="w", padx=8, pady=1)

        # Changelog / Feature Registry
        changelog_label = ctk.CTkLabel(
            content_frame,
            text="RIWAYAT FITUR & PEMBARUAN (CHANGELOG)",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=self.theme.text_secondary,
        )
        changelog_label.pack(anchor="w", padx=16, pady=(8, 4))

        changelog_scroll = ctk.CTkScrollableFrame(
            content_frame,
            fg_color=self.theme.bg_inner,
            border_color=self.theme.border_inner,
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            height=160,
        )
        changelog_scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        for item in reversed(FEATURE_REGISTRY):
            entry_frame = ctk.CTkFrame(changelog_scroll, fg_color="transparent")
            entry_frame.pack(fill="x", pady=4, padx=4)

            entry_header = ctk.CTkFrame(entry_frame, fg_color="transparent")
            entry_header.pack(fill="x")

            v_badge = ctk.CTkLabel(
                entry_header,
                text=f"v{item['version']}",
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
                text_color=self.theme.metric_number,
            )
            v_badge.pack(side="left")

            t_label = ctk.CTkLabel(
                entry_header,
                text=f" — {item['title']}",
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                text_color=self.theme.text_primary,
            )
            t_label.pack(side="left")

            d_badge = ctk.CTkLabel(
                entry_header,
                text=item["date"],
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=9),
                text_color=self.theme.text_muted,
            )
            d_badge.pack(side="right")

            desc_label = ctk.CTkLabel(
                entry_frame,
                text=item["description"],
                font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
                text_color=self.theme.text_secondary,
                wraplength=520,
                justify="left",
            )
            desc_label.pack(anchor="w", pady=(1, 0))

        # Bottom Close Button
        btn_close = ctk.CTkButton(
            content_frame,
            text="Tutup",
            command=self.destroy,
            fg_color=self.theme.disabled_bg,
            hover_color=self.theme.border_card,
            text_color=self.theme.text_primary,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            height=30,
            corner_radius=6,
        )
        btn_close.pack(anchor="e", padx=16, pady=(0, 12))

    def _get_pkg_version(self, pkg_name: str) -> str:
        try:
            return importlib.metadata.version(pkg_name)
        except Exception:
            return "Aktif"
