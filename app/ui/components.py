# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import os
from pathlib import Path
import re
import sys
from typing import Any, Callable, Literal
import customtkinter as ctk

from app.config.settings import DEFAULT_ADDRESS_FORMAT
from app.config.theme import THEME, ThemeTokens
from app.config.version import (
    __app_name__,
    __author__,
    __version__,
)
from app.core.excel_scanner import VoterData, is_blocked_pii, build_voter_address
from app.ui.modals import (
    apply_window_icon,
    get_col_letter,
    ModernNotificationModal,
    AboutModal,
    UpdateNotificationModal,
    ColumnMapperModal,
)

class InAppToastBanner(ctk.CTkFrame):
    def __init__(self, parent: Any, theme: ThemeTokens | None = None, **kwargs: Any) -> None:
        self.theme = theme or ThemeTokens()
        super().__init__(
            parent,
            fg_color=THEME["status_ready_bg"],
            border_color=THEME["status_ready_text"],
            border_width=1,
            corner_radius=self.theme.corner_radius_control,
            **kwargs,
        )
        self.parent = parent
        self._timer_id: str | None = None

        self.label_icon = ctk.CTkLabel(
            self,
            text="✓",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=12, weight="bold"),
            text_color=THEME["status_ready_text"],
            width=24,
        )
        self.label_icon.pack(side="left", padx=(10, 4), pady=6)

        self.label_msg = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_primary"],
            anchor="w",
            justify="left",
        )
        self.label_msg.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=6)

        self.btn_close = ctk.CTkButton(
            self,
            text="✕",
            width=22,
            height=22,
            fg_color="transparent",
            hover_color=THEME["bg_hover"],
            text_color=THEME["text_secondary"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            corner_radius=4,
            command=self.hide,
        )
        self.btn_close.pack(side="right", padx=(0, 8), pady=6)

    def show(
        self,
        message: str,
        variant: Literal["success", "info", "warning"] = "success",
        duration_ms: int = 3500,
        before_widget: Any = None,
    ) -> None:
        if self._timer_id:
            try:
                self.after_cancel(self._timer_id)
            except Exception:
                pass
            self._timer_id = None

        if variant == "success":
            icon = "✓"
            bg = THEME["status_ready_bg"]
            border = THEME["status_ready_text"]
            fg = THEME["status_ready_text"]
        elif variant == "warning":
            icon = "⚠"
            bg = THEME["status_warn_bg"]
            border = THEME["status_warn_text"]
            fg = THEME["status_warn_text"]
        else:
            icon = "ℹ"
            bg = THEME["btn_accent_soft_bg"]
            border = THEME["btn_accent_soft_border"]
            fg = THEME["btn_accent_soft_text"]

        self.configure(fg_color=bg, border_color=border)
        self.label_icon.configure(text=icon, text_color=fg)
        self.label_msg.configure(text=message, text_color=THEME["text_primary"])

        if before_widget is not None:
            self.pack(fill="x", padx=16, pady=(4, 6), before=before_widget)
        else:
            self.pack(fill="x", padx=16, pady=(4, 6))

        self._timer_id = self.after(duration_ms, self.hide)

    def hide(self) -> None:
        if self._timer_id:
            try:
                self.after_cancel(self._timer_id)
            except Exception:
                pass
            self._timer_id = None
        self.pack_forget()

class FilterStatusBanner(ctk.CTkFrame):
    def __init__(self, parent: Any, theme: ThemeTokens | None = None, **kwargs: Any) -> None:
        self.theme = theme or ThemeTokens()
        super().__init__(
            parent,
            fg_color=THEME["status_info_bg"],
            border_color=THEME["status_info_text"],
            border_width=1,
            corner_radius=self.theme.corner_radius_badge,
            height=30,
            **kwargs,
        )
        self.current_state: Literal["all", "empty", "active", "no_match"] = "all"
        self.current_message = ""

        self.label_icon = ctk.CTkLabel(
            self,
            text="ℹ",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=12, weight="bold"),
            text_color=THEME["status_info_text"],
            width=22,
        )
        self.label_icon.pack(side="left", padx=(10, 4), pady=4)

        self.label_msg = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_primary"],
            anchor="w",
            justify="left",
        )
        self.label_msg.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=4)

    def set_status(
        self,
        state: Literal["all", "empty", "active", "no_match"],
        message: str,
    ) -> None:
        self.current_state = state
        self.current_message = message

        if state == "all":
            icon = "ℹ"
            bg = THEME["status_info_bg"]
            border = THEME["status_info_text"]
            fg = THEME["status_info_text"]
            text_col = THEME["text_primary"]
        elif state == "empty":
            icon = "⚠"
            bg = THEME["status_warn_bg"]
            border = THEME["status_warn_text"]
            fg = THEME["status_warn_text"]
            text_col = THEME["text_primary"]
        elif state == "active":
            icon = "✓"
            bg = THEME["status_ready_bg"]
            border = THEME["status_ready_text"]
            fg = THEME["status_ready_text"]
            text_col = THEME["text_primary"]
        else:
            icon = "✕"
            bg = THEME["status_error_bg"]
            border = THEME["status_error_text"]
            fg = THEME["status_error_text"]
            text_col = THEME["status_error_text"]

        self.configure(fg_color=bg, border_color=border)
        self.label_icon.configure(text=icon, text_color=fg)
        self.label_msg.configure(text=message, text_color=text_col)

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.set_status(self.current_state, self.current_message)

class DocumentSourceRow(ctk.CTkFrame):
    """Komponen 2 baris rapi untuk sumber dokumen Word dan Excel dengan toolbar anti-clipping."""

    def __init__(
        self,
        master: any,
        doc_type: str,
        initial_path: str,
        theme: ThemeTokens,
        on_browse: Callable[[], None],
        on_recheck: Callable[[], None],
        on_preview: Callable[[], None] | None = None,
        recheck_label: str = "Cek Ulang",
        on_map_columns: Callable[[], None] | None = None,
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            **kwargs,
        )
        self.theme = theme
        self.doc_type = doc_type
        self.current_path = initial_path
        self.on_browse = on_browse
        self.on_recheck = on_recheck
        self.on_preview = on_preview
        self.on_map_columns = on_map_columns

        self._build_ui(recheck_label)

    def _build_ui(self, recheck_label: str) -> None:
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=14, pady=10)

        top_row = ctk.CTkFrame(inner, fg_color="transparent")
        top_row.pack(fill="x")

        self.type_label = ctk.CTkLabel(
            top_row,
            text=self.doc_type,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_secondary"],
        )
        self.type_label.pack(side="left")

        self.status_label = ctk.CTkLabel(
            top_row,
            text="• Menunggu pemeriksaan...",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_muted"],
        )
        self.status_label.pack(side="right")

        filename = os.path.basename(self.current_path) if self.current_path else "Belum ada berkas dipilih"
        self.filename_label = ctk.CTkLabel(
            inner,
            text=filename,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=THEME["text_primary"],
            anchor="w",
        )
        self.filename_label.pack(fill="x", pady=(2, 4))

        self.btn_row = ctk.CTkFrame(inner, fg_color="transparent")
        self.btn_row.pack(fill="x", pady=(6, 0))

        is_word = "word" in self.doc_type.lower()
        browse_text = "Pilih Template Word" if is_word else "Pilih Berkas DPT"

        self.btn_browse = ctk.CTkButton(
            self.btn_row,
            text=browse_text,
            command=self.on_browse,
            fg_color=THEME["btn_accent_soft_bg"],
            hover_color=THEME["btn_accent_soft_hover"],
            border_width=1,
            border_color=THEME["btn_accent_soft_border"],
            text_color=THEME["btn_accent_soft_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            width=125 if is_word else 115,
            height=30,
            corner_radius=self.theme.corner_radius_control,
        )
        self.btn_browse.pack(side="left", padx=(0, 6))

        if self.on_map_columns:
            self.btn_map_col = ctk.CTkButton(
                self.btn_row,
                text="Atur Kolom",
                command=self.on_map_columns,
                fg_color=THEME["btn_secondary_bg"],
                hover_color=THEME["btn_secondary_hover"],
                border_width=1,
                border_color=THEME["btn_secondary_border"],
                text_color=THEME["btn_secondary_text"],
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                width=90,
                height=30,
                corner_radius=self.theme.corner_radius_control,
            )
            self.btn_map_col.pack(side="left", padx=(0, 6))

        recheck_width = 95 if is_word else 100
        self.btn_recheck = ctk.CTkButton(
            self.btn_row,
            text=recheck_label,
            command=self.on_recheck,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_width=1,
            border_color=THEME["btn_secondary_border"],
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            width=recheck_width,
            height=30,
            corner_radius=self.theme.corner_radius_control,
        )
        self.btn_recheck.pack(side="left", padx=(0, 6))

        if self.on_preview:
            self.btn_preview = ctk.CTkButton(
                self.btn_row,
                text="Pratinjau PDF",
                command=self.on_preview,
                fg_color=THEME["btn_secondary_bg"],
                hover_color=THEME["btn_secondary_hover"],
                border_width=1,
                border_color=THEME["btn_secondary_border"],
                text_color=THEME["btn_secondary_text"],
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                width=110,
                height=30,
                corner_radius=self.theme.corner_radius_control,
            )
            self.btn_preview.pack(side="left")

    def set_file(self, file_path: str) -> None:
        self.current_path = file_path
        filename = os.path.basename(file_path) if file_path else "Belum ada berkas dipilih"
        self.filename_label.configure(text=filename)

    def set_status(
        self,
        text: str,
        state: Literal["ready", "warning", "error", "process", "neutral"] = "neutral",
    ) -> None:
        if state == "ready":
            color = THEME["status_ready_text"]
        elif state in ("warning", "process", "amber"):
            color = THEME["status_warn_text"]
        elif state == "error":
            color = THEME["status_error_text"]
        else:
            color = THEME["text_muted"]

        if text.startswith("•") or text.startswith("✓") or text.startswith("["):
            display_text = text
        else:
            display_text = f"• {text}"
        self.status_label.configure(text=display_text, text_color=color)

    def update_theme(self, theme: ThemeTokens) -> None:
        pass

class CleanMetricBar(ctk.CTkFrame):
    """Bar horizontal 3 metrik yang bersih dengan font monospace tabular-nums."""

    def __init__(self, master: any, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(
            master,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            height=62,
            **kwargs,
        )
        self.theme = theme

        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure((0, 2, 4), weight=1)
        self.grid_columnconfigure((1, 3), weight=0)

        m1_box = ctk.CTkFrame(self, fg_color="transparent")
        m1_box.grid(row=0, column=0, padx=12, pady=8, sticky="nsew")
        self.lbl_t1 = ctk.CTkLabel(
            m1_box,
            text="Wilayah terpilih",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=THEME["text_secondary"],
        )
        self.lbl_t1.pack(anchor="center")
        self.lbl_v1 = ctk.CTkLabel(
            m1_box,
            text="0 Kadus",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=16, weight="bold"),
            text_color=THEME["metric_value"],
        )
        self.lbl_v1.pack(anchor="center")

        div1 = ctk.CTkFrame(self, fg_color=THEME["border_surface"], width=1)
        div1.grid(row=0, column=1, sticky="ns", pady=8)

        m2_box = ctk.CTkFrame(self, fg_color="transparent")
        m2_box.grid(row=0, column=2, padx=12, pady=8, sticky="nsew")
        self.lbl_t2 = ctk.CTkLabel(
            m2_box,
            text="Total pemilih",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=THEME["text_secondary"],
        )
        self.lbl_t2.pack(anchor="center")
        self.lbl_v2 = ctk.CTkLabel(
            m2_box,
            text="0 Orang",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=16, weight="bold"),
            text_color=THEME["metric_value"],
        )
        self.lbl_v2.pack(anchor="center")

        div2 = ctk.CTkFrame(self, fg_color=THEME["border_surface"], width=1)
        div2.grid(row=0, column=3, sticky="ns", pady=8)

        m3_box = ctk.CTkFrame(self, fg_color="transparent")
        m3_box.grid(row=0, column=4, padx=12, pady=8, sticky="nsew")
        self.lbl_t3 = ctk.CTkLabel(
            m3_box,
            text="Kebutuhan kertas",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=THEME["text_secondary"],
        )
        self.lbl_t3.pack(anchor="center")
        self.lbl_v3 = ctk.CTkLabel(
            m3_box,
            text="0 Lembar F4",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=16, weight="bold"),
            text_color=THEME["metric_value"],
        )
        self.lbl_v3.pack(anchor="center")

    def update_metrics(self, kadus_count: int, total_voters: int, total_pages: int) -> None:
        self.lbl_v1.configure(text=f"{kadus_count:,} Kadus")
        self.lbl_v2.configure(text=f"{total_voters:,} Orang")
        self.lbl_v3.configure(text=f"{total_pages:,} Lembar F4")

    def update_theme(self, theme: ThemeTokens) -> None:
        pass

class KadusRowItem(ctk.CTkFrame):
    """Baris Kadus interaktif dengan tombol buka PDF per wilayah."""

    def __init__(
        self,
        master: any,
        sheet_name: str,
        voter_count: int,
        page_count: int,
        is_selected: bool,
        theme: ThemeTokens,
        on_toggle: Callable[[], None] | None = None,
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            height=40,
            **kwargs,
        )
        self.theme = theme
        self.sheet_name = sheet_name
        self.voter_count = voter_count
        self.page_count = page_count
        self.generated_pdf_path: str = ""

        self.checkbox_var = ctk.BooleanVar(value=is_selected)

        self.checkbox = ctk.CTkCheckBox(
            self,
            text=sheet_name,
            variable=self.checkbox_var,
            command=on_toggle,
            font=ctk.CTkFont(family=theme.font_family_main, size=12, weight="bold"),
            text_color=THEME["text_primary"],
            fg_color=THEME["btn_primary_bg"],
            hover_color=THEME["btn_primary_hover"],
            border_color=THEME["border_surface"],
            corner_radius=4,
            width=140,
        )
        self.checkbox.pack(side="left", padx=(12, 6), pady=6)

        self.info_label = ctk.CTkLabel(
            self,
            text=f"{voter_count:,} Pemilih  •  {page_count:,} Lembar",
            font=ctk.CTkFont(family=theme.font_family_mono, size=11),
            text_color=THEME["text_secondary"],
        )
        self.info_label.pack(side="left", padx=8)

        self.right_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.right_frame.pack(side="right", padx=(4, 12))

        self.btn_open_pdf = ctk.CTkButton(
            self.right_frame,
            text="Buka PDF",
            command=self._open_pdf,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=theme.font_family_main, size=10, weight="bold"),
            width=70,
            height=26,
            corner_radius=6,
        )

        self.status_label = ctk.CTkLabel(
            self.right_frame,
            text="• Siap",
            font=ctk.CTkFont(family=theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["status_ready_text"],
        )
        self.status_label.pack(side="right")

    def _open_pdf(self) -> None:
        if self.generated_pdf_path and os.path.exists(self.generated_pdf_path):
            try:
                os.startfile(self.generated_pdf_path)
            except Exception:
                pass

    def set_printing_active(self) -> None:
        self.status_label.configure(
            text="• Mencetak...",
            text_color=THEME["status_warn_text"],
        )

    def set_completed(self, pdf_path: str = "") -> None:
        self.generated_pdf_path = pdf_path
        self.status_label.configure(
            text="• Selesai ✓",
            text_color=THEME["status_ready_text"],
        )
        if pdf_path and os.path.exists(pdf_path):
            self.btn_open_pdf.pack(side="right", padx=(0, 8))

    def set_error(self, err_text: str = "• Gagal") -> None:
        self.status_label.configure(
            text=err_text,
            text_color=THEME["status_error_text"],
        )

    def reset_status(self) -> None:
        self.generated_pdf_path = ""
        self.btn_open_pdf.pack_forget()
        self.status_label.configure(
            text="• Siap",
            text_color=THEME["status_ready_text"],
        )

    def set_enabled(self, enabled: bool) -> None:
        self.checkbox.configure(state="normal" if enabled else "disabled")

    def update_filter_stats(self, matched_count: int, matched_pages: int, is_filter_active: bool) -> None:
        if not is_filter_active:
            self.info_label.configure(text=f"{self.voter_count:,} Pemilih  •  {self.page_count:,} Lembar")
            self.status_label.configure(text="• Siap", text_color=THEME["status_ready_text"])
            self.checkbox.configure(state="normal")
        else:
            if matched_count > 0:
                self.info_label.configure(text=f"{matched_count:,} / {self.voter_count:,} Pemilih  •  {matched_pages:,} Lembar")
                self.status_label.configure(text=f"• {matched_count:,} Cocok", text_color=THEME["status_ready_text"])
                self.checkbox.configure(state="normal")
            else:
                self.info_label.configure(text=f"0 / {self.voter_count:,} Pemilih  •  0 Lembar")
                self.status_label.configure(text="• Dilewati", text_color=THEME["text_muted"])

    def update_theme(self, theme: ThemeTokens) -> None:
        pass

# Alias for backwards compatibility
KadusItemRow = KadusRowItem

class VoterPreviewTable(ctk.CTkFrame):
    """Tabel pratinjau data pemilih yang telah ditransformasikan dengan tata letak monospace terstruktur."""

    def __init__(self, master: any, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(
            master,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            **kwargs,
        )
        self.theme = theme
        self.cached_voters: list[VoterData] = []
        self.cached_address_format: str = DEFAULT_ADDRESS_FORMAT
        self.is_filtered: bool = False
        self.custom_empty_message: str | None = None
        self.total_count: int | None = None
        self.matched_count: int | None = None

        self._build_ui()

    def _configure_grid_columns(self, frame: ctk.CTkFrame) -> None:
        frame.grid_columnconfigure(0, weight=0, minsize=48)
        frame.grid_columnconfigure(1, weight=2, minsize=190)
        frame.grid_columnconfigure(2, weight=0, minsize=95)
        frame.grid_columnconfigure(3, weight=0, minsize=90)
        frame.grid_columnconfigure(4, weight=4)
        frame.grid_columnconfigure(5, weight=0, minsize=100)

    def _build_ui(self) -> None:
        self.status_bar = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_elevated"],
            corner_radius=6,
            height=32,
        )
        self.status_bar.pack(fill="x", padx=10, pady=(8, 4))

        self.status_left = ctk.CTkFrame(self.status_bar, fg_color="transparent")
        self.status_left.pack(side="left", fill="y", padx=(10, 4))

        self.lbl_status_indicator = ctk.CTkLabel(
            self.status_left,
            text="•",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=14, weight="bold"),
            text_color=THEME["text_muted"],
        )
        self.lbl_status_indicator.pack(side="left", padx=(0, 4))

        self.lbl_status_summary = ctk.CTkLabel(
            self.status_left,
            text="Pratinjau Sampel: Menunggu data Excel DPT",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_primary"],
            anchor="w",
        )
        self.lbl_status_summary.pack(side="left")

        self.lbl_format_info = ctk.CTkLabel(
            self.status_bar,
            text=f"Format: {self.cached_address_format}",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_secondary"],
            anchor="e",
        )
        self.lbl_format_info.pack(side="right", padx=(0, 10))

        self.header_row = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_surface"],
            border_width=1,
            border_color=THEME["border_surface"],
            corner_radius=6,
            height=30,
        )
        self.header_row.pack(fill="x", padx=10, pady=(0, 4))

        self._configure_grid_columns(self.header_row)

        ctk.CTkLabel(
            self.header_row,
            text="NO",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
            text_color=THEME["text_muted"],
            anchor="w",
        ).grid(row=0, column=0, padx=(10, 4), pady=5, sticky="w")

        ctk.CTkLabel(
            self.header_row,
            text="NAMA PEMILIH",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_muted"],
            anchor="w",
        ).grid(row=0, column=1, padx=4, pady=5, sticky="w")

        ctk.CTkLabel(
            self.header_row,
            text="L/P",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_muted"],
            anchor="w",
        ).grid(row=0, column=2, padx=4, pady=5, sticky="w")

        ctk.CTkLabel(
            self.header_row,
            text="USIA",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_muted"],
            anchor="w",
        ).grid(row=0, column=3, padx=4, pady=5, sticky="w")

        ctk.CTkLabel(
            self.header_row,
            text="ALAMAT PEMILIH",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_muted"],
            anchor="w",
        ).grid(row=0, column=4, padx=4, pady=5, sticky="ew")

        ctk.CTkLabel(
            self.header_row,
            text="WILAYAH",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_muted"],
            anchor="e",
        ).grid(row=0, column=5, padx=(4, 10), pady=5, sticky="e")

        self.rows_container = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            scrollbar_fg_color="transparent",
            scrollbar_button_color=THEME["bg_elevated"],
            scrollbar_button_hover_color=THEME["border_surface"],
        )
        self.rows_container.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        self._render_rows()

    def set_voters(
        self,
        voters: list[VoterData],
        address_format: str | None = None,
        is_filtered: bool = False,
        empty_message: str | None = None,
        total_count: int | None = None,
        matched_count: int | None = None,
    ) -> None:
        self.cached_voters = list(voters)
        self.is_filtered = is_filtered
        self.custom_empty_message = empty_message
        if total_count is not None:
            self.total_count = total_count
        if matched_count is not None:
            self.matched_count = matched_count
        if address_format:
            self.cached_address_format = address_format
        self._render_rows()

    def update_address_format(self, address_format: str) -> None:
        self.cached_address_format = address_format
        self.lbl_format_info.configure(text=f"Format: {address_format}")
        self._render_rows()

    def _render_rows(self) -> None:
        for child in self.rows_container.winfo_children():
            child.destroy()

        limit = 25 if self.is_filtered else 5
        displayed_voters = self.cached_voters[:limit]
        shown_count = len(displayed_voters)

        format_display = self.cached_address_format or DEFAULT_ADDRESS_FORMAT
        self.lbl_format_info.configure(text=f"Format: {format_display}")

        if self.is_filtered:
            matched = self.matched_count if self.matched_count is not None else len(self.cached_voters)
            if matched > 0:
                self.lbl_status_indicator.configure(text="•", text_color=THEME["status_info_text"])
                self.lbl_status_summary.configure(
                    text=f"Pratinjau Filter: Menampilkan {shown_count} dari {matched:,} Pemilih Cocok".replace(",", ".")
                )
            else:
                self.lbl_status_indicator.configure(text="•", text_color=THEME["status_error_text"])
                self.lbl_status_summary.configure(text="Pratinjau Filter: 0 Pemilih Cocok")
        else:
            if self.cached_voters:
                tot = self.total_count if (self.total_count is not None and self.total_count > 0) else len(self.cached_voters)
                self.lbl_status_indicator.configure(text="•", text_color=THEME["status_ready_text"])
                self.lbl_status_summary.configure(
                    text=f"Pratinjau Sampel: Menampilkan {shown_count} dari {tot:,} Pemilih".replace(",", ".")
                )
            else:
                self.lbl_status_indicator.configure(text="•", text_color=THEME["text_muted"])
                self.lbl_status_summary.configure(text="Pratinjau Sampel: Menunggu data Excel DPT")

        if len(displayed_voters) <= 6:
            try:
                self.rows_container._scrollbar.grid_remove()
            except Exception:
                pass
        else:
            try:
                self.rows_container._scrollbar.grid()
                self.rows_container.configure(
                    scrollbar_fg_color="transparent",
                    scrollbar_button_color=THEME["bg_elevated"],
                    scrollbar_button_hover_color=THEME["border_surface"],
                )
            except Exception:
                pass

        if not displayed_voters:
            msg = self.custom_empty_message or (
                "Tidak ada pemilih yang cocok dengan filter. Periksa kembali nomor, RT, atau nama yang diketik."
                if self.is_filtered
                else "Belum ada data pemilih. Pilih dan scan berkas Excel DPT di kartu 1 untuk melihat pratinjau 5 pemilih pertama."
            )
            empty_box = ctk.CTkFrame(
                self.rows_container,
                fg_color=THEME["bg_elevated"],
                border_color=THEME["border_surface"],
                border_width=1,
                corner_radius=6,
            )
            empty_box.pack(fill="x", padx=16, pady=32)

            ctk.CTkLabel(
                empty_box,
                text="ℹ  " + msg,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
                text_color=THEME["text_muted"],
                justify="center",
                wraplength=650,
            ).pack(padx=16, pady=20)
            return

        for idx, voter in enumerate(displayed_voters):
            is_odd_item = (idx % 2 == 0)
            bg_col = THEME["bg_surface"] if is_odd_item else "transparent"
            border_col = THEME["border_card"]
            border_w = 0 if is_odd_item else 1

            row_frame = ctk.CTkFrame(
                self.rows_container,
                fg_color=bg_col,
                border_color=border_col,
                border_width=border_w,
                corner_radius=6,
                height=34,
            )
            row_frame.pack(fill="x", padx=2, pady=2)
            self._configure_grid_columns(row_frame)

            def _bind_hover(target: Any, target_row: ctk.CTkFrame, default_bg: Any) -> None:
                def _on_enter(e: Any) -> None:
                    try:
                        target_row.configure(fg_color=THEME["bg_hover"])
                    except Exception:
                        pass
                def _on_leave(e: Any) -> None:
                    try:
                        target_row.configure(fg_color=default_bg)
                    except Exception:
                        pass
                target.bind("<Enter>", _on_enter, add="+")
                target.bind("<Leave>", _on_leave, add="+")

            _bind_hover(row_frame, row_frame, bg_col)

            lbl_no = ctk.CTkLabel(
                row_frame,
                text=f"{idx + 1:02d}",
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=12),
                text_color=THEME["text_muted"],
                anchor="w",
            )
            lbl_no.grid(row=0, column=0, padx=(10, 4), pady=5, sticky="w")
            _bind_hover(lbl_no, row_frame, bg_col)

            lbl_nama = ctk.CTkLabel(
                row_frame,
                text=voter.nama,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=12, weight="bold"),
                text_color=THEME["text_primary"],
                anchor="w",
            )
            lbl_nama.grid(row=0, column=1, padx=4, pady=5, sticky="w")
            _bind_hover(lbl_nama, row_frame, bg_col)

            jk_raw = voter.jenis_kelamin.strip().upper()
            if jk_raw in ("L", "LAKI-LAKI"):
                jk_clean = "Laki-laki"
            elif jk_raw in ("P", "PEREMPUAN"):
                jk_clean = "Perempuan"
            else:
                jk_clean = voter.jenis_kelamin

            lbl_jk = ctk.CTkLabel(
                row_frame,
                text=jk_clean,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=12),
                text_color=THEME["text_secondary"],
                anchor="w",
            )
            lbl_jk.grid(row=0, column=2, padx=4, pady=5, sticky="w")
            _bind_hover(lbl_jk, row_frame, bg_col)

            umur_val = str(voter.umur).strip()
            if not umur_val or umur_val == "-":
                umur_clean = "-"
            else:
                raw_num = re.sub(r"(?i)\s*(tahun|thn).*$", "", umur_val).strip()
                umur_clean = f"{raw_num} Tahun" if raw_num else umur_val
            lbl_umur = ctk.CTkLabel(
                row_frame,
                text=umur_clean,
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=12),
                text_color=THEME["text_secondary"],
                anchor="w",
            )
            lbl_umur.grid(row=0, column=3, padx=4, pady=5, sticky="w")
            _bind_hover(lbl_umur, row_frame, bg_col)

            display_addr = build_voter_address(voter, self.cached_address_format) if self.cached_address_format else voter.alamat
            lbl_addr = ctk.CTkLabel(
                row_frame,
                text=display_addr,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=12),
                text_color=THEME["text_secondary"],
                anchor="w",
            )
            lbl_addr.grid(row=0, column=4, padx=4, pady=5, sticky="ew")
            _bind_hover(lbl_addr, row_frame, bg_col)

            wilayah_name = voter.sheet_name.strip() if voter.sheet_name else "KADUS 1"
            badge = ctk.CTkFrame(
                row_frame,
                fg_color=THEME["bg_elevated"],
                border_color=THEME["border_surface"],
                border_width=1,
                corner_radius=4,
                height=22,
            )
            badge.grid(row=0, column=5, padx=(4, 10), pady=5, sticky="e")
            _bind_hover(badge, row_frame, bg_col)

            badge_lbl = ctk.CTkLabel(
                badge,
                text=wilayah_name.upper(),
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=10, weight="bold"),
                text_color=THEME["btn_accent_soft_text"],
            )
            badge_lbl.pack(padx=8, pady=2)
            _bind_hover(badge_lbl, row_frame, bg_col)

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.configure(fg_color=THEME["bg_surface"], border_color=THEME["border_surface"])
        if hasattr(self, "status_bar"):
            self.status_bar.configure(fg_color=THEME["bg_elevated"])
        if hasattr(self, "header_row"):
            self.header_row.configure(fg_color=THEME["bg_surface"], border_color=THEME["border_surface"])
        self._render_rows()

class MetricCard(ctk.CTkFrame):
    """Kartu metrik mandiri (kompatibilitas)."""

    def __init__(self, master: any, title: str, value: str, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(
            master,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            **kwargs,
        )
        self.theme = theme
        self.title_label = ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(family=theme.font_family_main, size=10, weight="bold"),
            text_color=THEME["text_secondary"],
        )
        self.title_label.pack(anchor="w", padx=12, pady=(8, 2))
        self.value_label = ctk.CTkLabel(
            self,
            text=value,
            font=ctk.CTkFont(family=theme.font_family_mono, size=18, weight="bold"),
            text_color=THEME["metric_value"],
        )
        self.value_label.pack(anchor="w", padx=12, pady=(0, 8))

    def set_value(self, value: str) -> None:
        self.value_label.configure(text=value)

    def update_theme(self, theme: ThemeTokens) -> None:
        pass

class StatusBadge(ctk.CTkFrame):
    """Badge status minimalis."""

    def __init__(self, master: any, text: str, status: str, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(master, corner_radius=6, height=26, fg_color=THEME["bg_surface"], border_color=THEME["border_surface"], border_width=1, **kwargs)
        self.theme = theme
        self.label = ctk.CTkLabel(
            self,
            text=text,
            font=ctk.CTkFont(family=theme.font_family_mono, size=11, weight="bold"),
            text_color=THEME["text_primary"],
        )
        self.label.pack(padx=8, pady=3)
        self.set_status(text, status)

    def set_status(self, text: str, status: str) -> None:
        if status == "ready":
            color = THEME["status_ready_text"]
        elif status in ("process", "amber", "warning"):
            color = THEME["status_warn_text"]
        elif status == "error":
            color = THEME["status_error_text"]
        else:
            color = THEME["text_secondary"]

        self.label.configure(text=text, text_color=color)

    def update_theme(self, theme: ThemeTokens, current_status: str = "neutral") -> None:
        pass


__all__ = [
    "InAppToastBanner",
    "FilterStatusBanner",
    "DocumentSourceRow",
    "CleanMetricBar",
    "KadusRowItem",
    "VoterPreviewTable",
    "MetricCard",
    "StatusBadge",
    "apply_window_icon",
    "get_col_letter",
    "ModernNotificationModal",
    "AboutModal",
    "UpdateNotificationModal",
    "ColumnMapperModal",
]
