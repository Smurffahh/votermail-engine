# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import importlib.metadata
import os
from pathlib import Path
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
from app.core.excel_scanner import VoterData

class DocumentSourceRow(ctk.CTkFrame):
    """Komponen baris ringkas dan rapi untuk sumber dokumen Word dan Excel."""

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
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            fg_color=theme.bg_surface,
            border_color=theme.border_subtle,
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

        self._build_ui(recheck_label)

    def _build_ui(self, recheck_label: str) -> None:
        # Left container: Document type, filename, and friendly status
        self.left_box = ctk.CTkFrame(self, fg_color="transparent")
        self.left_box.pack(side="left", fill="both", expand=True, padx=14, pady=10)

        self.type_label = ctk.CTkLabel(
            self.left_box,
            text=self.doc_type,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=self.theme.text_secondary,
        )
        self.type_label.pack(anchor="w")

        filename = os.path.basename(self.current_path) if self.current_path else "Belum ada berkas dipilih"
        self.filename_label = ctk.CTkLabel(
            self.left_box,
            text=filename,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=self.theme.text_primary,
        )
        self.filename_label.pack(anchor="w", pady=(1, 2))

        self.status_label = ctk.CTkLabel(
            self.left_box,
            text="Menunggu pemeriksaan...",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=self.theme.text_muted,
        )
        self.status_label.pack(anchor="w")

        # Right container: Action buttons
        self.right_box = ctk.CTkFrame(self, fg_color="transparent")
        self.right_box.pack(side="right", padx=12, pady=10)

        if self.on_preview:
            self.btn_preview = ctk.CTkButton(
                self.right_box,
                text="Pratinjau PDF",
                command=self.on_preview,
                fg_color=self.theme.btn_secondary_bg,
                hover_color=self.theme.btn_secondary_hover,
                border_color=self.theme.btn_secondary_border,
                border_width=1,
                text_color=self.theme.btn_secondary_text,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                height=28,
                corner_radius=6,
            )
            self.btn_preview.pack(side="left", padx=4)

        self.btn_recheck = ctk.CTkButton(
            self.right_box,
            text=recheck_label,
            command=self.on_recheck,
            fg_color=self.theme.btn_secondary_bg,
            hover_color=self.theme.btn_secondary_hover,
            border_color=self.theme.btn_secondary_border,
            border_width=1,
            text_color=self.theme.btn_secondary_text,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            height=28,
            corner_radius=6,
        )
        self.btn_recheck.pack(side="left", padx=4)

        self.btn_browse = ctk.CTkButton(
            self.right_box,
            text="Ganti File",
            command=self.on_browse,
            fg_color=self.theme.btn_secondary_bg,
            hover_color=self.theme.btn_secondary_hover,
            border_color=self.theme.btn_secondary_border,
            border_width=1,
            text_color=self.theme.btn_secondary_text,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            height=28,
            corner_radius=6,
        )
        self.btn_browse.pack(side="left", padx=(4, 0))

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
            color = self.theme.status_ready
        elif state == "warning":
            color = self.theme.status_warning
        elif state == "error":
            color = self.theme.status_error
        else:
            color = self.theme.text_secondary

        self.status_label.configure(text=text, text_color=color)

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.configure(fg_color=theme.bg_surface, border_color=theme.border_subtle)
        self.type_label.configure(text_color=theme.text_secondary)
        self.filename_label.configure(text_color=theme.text_primary)
        for btn in [getattr(self, "btn_preview", None), self.btn_recheck, self.btn_browse]:
            if btn:
                btn.configure(
                    fg_color=theme.btn_secondary_bg,
                    hover_color=theme.btn_secondary_hover,
                    border_color=theme.btn_secondary_border,
                    text_color=theme.btn_secondary_text,
                )

class CleanMetricBar(ctk.CTkFrame):
    """Bar horizontal 3 metrik yang bersih dengan font monospace tabular-nums."""

    def __init__(self, master: any, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(
            master,
            fg_color=theme.bg_surface,
            border_color=theme.border_subtle,
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

        # Metric 1: Wilayah Terpilih
        m1_box = ctk.CTkFrame(self, fg_color="transparent")
        m1_box.grid(row=0, column=0, padx=12, pady=8, sticky="nsew")
        self.lbl_t1 = ctk.CTkLabel(
            m1_box,
            text="WILAYAH TERPILIH",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=self.theme.text_secondary,
        )
        self.lbl_t1.pack(anchor="center")
        self.lbl_v1 = ctk.CTkLabel(
            m1_box,
            text="0 Kadus",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=16, weight="bold"),
            text_color=self.theme.text_primary,
        )
        self.lbl_v1.pack(anchor="center")

        # Divider 1
        div1 = ctk.CTkFrame(self, fg_color=self.theme.border_subtle, width=1)
        div1.grid(row=0, column=1, sticky="ns", pady=8)

        # Metric 2: Total Pemilih
        m2_box = ctk.CTkFrame(self, fg_color="transparent")
        m2_box.grid(row=0, column=2, padx=12, pady=8, sticky="nsew")
        self.lbl_t2 = ctk.CTkLabel(
            m2_box,
            text="TOTAL PEMILIH",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=self.theme.text_secondary,
        )
        self.lbl_t2.pack(anchor="center")
        self.lbl_v2 = ctk.CTkLabel(
            m2_box,
            text="0 Orang",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=16, weight="bold"),
            text_color=self.theme.text_primary,
        )
        self.lbl_v2.pack(anchor="center")

        # Divider 2
        div2 = ctk.CTkFrame(self, fg_color=self.theme.border_subtle, width=1)
        div2.grid(row=0, column=3, sticky="ns", pady=8)

        # Metric 3: Estimasi Kertas
        m3_box = ctk.CTkFrame(self, fg_color="transparent")
        m3_box.grid(row=0, column=4, padx=12, pady=8, sticky="nsew")
        self.lbl_t3 = ctk.CTkLabel(
            m3_box,
            text="KEBUTUHAN KERTAS",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=self.theme.text_secondary,
        )
        self.lbl_t3.pack(anchor="center")
        self.lbl_v3 = ctk.CTkLabel(
            m3_box,
            text="0 Lembar F4",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=16, weight="bold"),
            text_color=self.theme.text_primary,
        )
        self.lbl_v3.pack(anchor="center")

    def update_metrics(self, kadus_count: int, total_voters: int, total_pages: int) -> None:
        self.lbl_v1.configure(text=f"{kadus_count:,} Kadus")
        self.lbl_v2.configure(text=f"{total_voters:,} Orang")
        self.lbl_v3.configure(text=f"{total_pages:,} Lembar F4")

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.configure(fg_color=theme.bg_surface, border_color=theme.border_subtle)
        for lbl in [self.lbl_t1, self.lbl_t2, self.lbl_t3]:
            lbl.configure(text_color=theme.text_secondary)
        for lbl in [self.lbl_v1, self.lbl_v2, self.lbl_v3]:
            lbl.configure(text_color=theme.text_primary)

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
            fg_color=theme.bg_surface,
            border_color=theme.border_subtle,
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
            fg_color=theme.primary_action,
            hover_color=theme.primary_action_hover,
            border_color=theme.border_subtle,
            corner_radius=4,
            width=140,
        )
        self.checkbox.pack(side="left", padx=(12, 6), pady=6)

        self.info_label = ctk.CTkLabel(
            self,
            text=f"{voter_count:,} Pemilih  •  {page_count:,} Lembar",
            font=ctk.CTkFont(family=theme.font_family_mono, size=11),
            text_color=theme.text_secondary,
        )
        self.info_label.pack(side="left", padx=8)

        # Right container for status & optional open PDF button
        self.right_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.right_frame.pack(side="right", padx=(4, 12))

        self.btn_open_pdf = ctk.CTkButton(
            self.right_frame,
            text="Buka PDF",
            command=self._open_pdf,
            fg_color=theme.btn_secondary_bg,
            hover_color=theme.btn_secondary_hover,
            border_color=theme.btn_secondary_border,
            border_width=1,
            text_color=theme.btn_secondary_text,
            font=ctk.CTkFont(family=theme.font_family_main, size=10, weight="bold"),
            width=65,
            height=24,
            corner_radius=4,
        )

        self.status_label = ctk.CTkLabel(
            self.right_frame,
            text="• Siap",
            font=ctk.CTkFont(family=theme.font_family_main, size=11, weight="bold"),
            text_color=theme.status_ready,
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
            text_color=self.theme.status_warning,
        )

    def set_completed(self, pdf_path: str = "") -> None:
        self.generated_pdf_path = pdf_path
        self.status_label.configure(
            text="• Selesai ✓",
            text_color=self.theme.status_ready,
        )
        if pdf_path and os.path.exists(pdf_path):
            self.btn_open_pdf.pack(side="right", padx=(0, 8))

    def set_error(self, err_text: str = "• Gagal") -> None:
        self.status_label.configure(
            text=err_text,
            text_color=self.theme.status_error,
        )

    def reset_status(self) -> None:
        self.generated_pdf_path = ""
        self.btn_open_pdf.pack_forget()
        self.status_label.configure(
            text="• Siap",
            text_color=self.theme.status_ready,
        )

    def set_enabled(self, enabled: bool) -> None:
        self.checkbox.configure(state="normal" if enabled else "disabled")

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.configure(fg_color=theme.bg_surface, border_color=theme.border_subtle)
        self.checkbox.configure(
            fg_color=theme.primary_action,
            hover_color=theme.primary_action_hover,
            border_color=theme.border_subtle,
        )
        self.info_label.configure(text_color=theme.text_secondary)
        self.btn_open_pdf.configure(
            fg_color=theme.btn_secondary_bg,
            hover_color=theme.btn_secondary_hover,
            border_color=theme.btn_secondary_border,
            text_color=theme.btn_secondary_text,
        )

# Alias for backwards compatibility
KadusItemRow = KadusRowItem

class VoterPreviewTable(ctk.CTkFrame):
    """Tabel pratinjau 5 contoh data pemilih yang telah ditransformasikan."""

    def __init__(self, master: any, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(
            master,
            fg_color=theme.bg_surface,
            border_color=theme.border_subtle,
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            **kwargs,
        )
        self.theme = theme

        self._build_ui()

    def _build_ui(self) -> None:
        # Table Header
        self.header_row = ctk.CTkFrame(self, fg_color="transparent", height=28)
        self.header_row.pack(fill="x", padx=10, pady=(8, 4))

        cols = [("NO", 40), ("NAMA PEMILIH", 170), ("L/P", 80), ("UMUR", 70), ("ALAMAT & WILAYAH", 240)]
        for title, width in cols:
            lbl = ctk.CTkLabel(
                self.header_row,
                text=title,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
                text_color=self.theme.text_muted,
                width=width,
                anchor="w",
            )
            lbl.pack(side="left", padx=4)

        # Hairline separator
        self.divider = ctk.CTkFrame(self, fg_color=self.theme.border_subtle, height=1)
        self.divider.pack(fill="x", padx=6, pady=2)

        # Rows Container
        self.rows_container = ctk.CTkFrame(self, fg_color="transparent")
        self.rows_container.pack(fill="both", expand=True, padx=6, pady=(2, 8))

        self.empty_label = ctk.CTkLabel(
            self.rows_container,
            text="Belum ada data pemilih. Silakan scan berkas Excel DPT terlebih dahulu.",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=self.theme.text_muted,
        )
        self.empty_label.pack(pady=24)

    def set_voters(self, voters: list[VoterData]) -> None:
        for child in self.rows_container.winfo_children():
            child.destroy()

        if not voters:
            self.empty_label = ctk.CTkLabel(
                self.rows_container,
                text="Tidak ada pemilih untuk ditampilkan.",
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
                text_color=self.theme.text_muted,
            )
            self.empty_label.pack(pady=24)
            return

        for voter in voters[:5]:
            row = ctk.CTkFrame(self.rows_container, fg_color="transparent", height=24)
            row.pack(fill="x", padx=4, pady=2)

            c1 = ctk.CTkLabel(
                row,
                text=voter.nomor,
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
                text_color=self.theme.text_secondary,
                width=40,
                anchor="w",
            )
            c1.pack(side="left", padx=4)

            c2 = ctk.CTkLabel(
                row,
                text=voter.nama,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                text_color=self.theme.text_primary,
                width=170,
                anchor="w",
            )
            c2.pack(side="left", padx=4)

            c3 = ctk.CTkLabel(
                row,
                text=voter.jenis_kelamin,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
                text_color=self.theme.text_secondary,
                width=80,
                anchor="w",
            )
            c3.pack(side="left", padx=4)

            c4 = ctk.CTkLabel(
                row,
                text=voter.umur,
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
                text_color=self.theme.text_secondary,
                width=70,
                anchor="w",
            )
            c4.pack(side="left", padx=4)

            alamat_wilayah = f"{voter.alamat} ({voter.sheet_name})" if voter.sheet_name else voter.alamat
            c5 = ctk.CTkLabel(
                row,
                text=alamat_wilayah,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
                text_color=self.theme.text_secondary,
                width=240,
                anchor="w",
            )
            c5.pack(side="left", padx=4)

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.configure(fg_color=theme.bg_surface, border_color=theme.border_subtle)
        self.divider.configure(fg_color=theme.border_subtle)

class MetricCard(ctk.CTkFrame):
    """Kartu metrik mandiri (kompatibilitas)."""

    def __init__(self, master: any, title: str, value: str, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(
            master,
            fg_color=theme.bg_surface,
            border_color=theme.border_subtle,
            border_width=1,
            corner_radius=theme.corner_radius_inner,
            **kwargs,
        )
        self.theme = theme
        self.title_label = ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(family=theme.font_family_main, size=10, weight="bold"),
            text_color=theme.text_secondary,
        )
        self.title_label.pack(anchor="w", padx=12, pady=(8, 2))
        self.value_label = ctk.CTkLabel(
            self,
            text=value,
            font=ctk.CTkFont(family=theme.font_family_mono, size=18, weight="bold"),
            text_color=theme.text_primary,
        )
        self.value_label.pack(anchor="w", padx=12, pady=(0, 8))

    def set_value(self, value: str) -> None:
        self.value_label.configure(text=value)

    def update_theme(self, theme: ThemeTokens) -> None:
        self.theme = theme
        self.configure(fg_color=theme.bg_surface, border_color=theme.border_subtle)
        self.title_label.configure(text_color=theme.text_secondary)
        self.value_label.configure(text_color=theme.text_primary)

class StatusBadge(ctk.CTkFrame):
    """Badge status minimalis."""

    def __init__(self, master: any, text: str, status: str, theme: ThemeTokens, **kwargs) -> None:
        super().__init__(master, corner_radius=6, height=26, fg_color=theme.bg_surface, **kwargs)
        self.theme = theme
        self.label = ctk.CTkLabel(
            self,
            text=text,
            font=ctk.CTkFont(family=theme.font_family_mono, size=11, weight="bold"),
            text_color=theme.text_primary,
        )
        self.label.pack(padx=8, pady=3)
        self.set_status(text, status)

    def set_status(self, text: str, status: str) -> None:
        if status == "ready":
            color = self.theme.status_ready
        elif status in ("process", "amber"):
            color = self.theme.status_warning
        elif status == "error":
            color = self.theme.status_error
        else:
            color = self.theme.text_secondary

        self.label.configure(text=text, text_color=color)

    def update_theme(self, theme: ThemeTokens, current_status: str = "neutral") -> None:
        self.theme = theme
        self.configure(fg_color=theme.bg_surface)
        self.set_status(self.label.cget("text"), current_status)

class AboutModal(ctk.CTkToplevel):
    def __init__(self, master: any, theme: ThemeTokens) -> None:
        super().__init__(master)
        self.theme = theme

        self.title("Tentang VoterMail Engine")
        self.geometry("620x640")
        self.minsize(560, 500)
        self.configure(fg_color=theme.bg_root)

        self.transient(master)
        self.grab_set()

        self.update_idletasks()
        try:
            mx = master.winfo_x()
            my = master.winfo_y()
            mw = master.winfo_width()
            mh = master.winfo_height()
            x = mx + (mw - 620) // 2
            y = my + (mh - 640) // 2
            self.geometry(f"620x640+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        self._build_ui()

    def _build_ui(self) -> None:
        content_frame = ctk.CTkFrame(
            self,
            fg_color=self.theme.bg_card,
            border_color=self.theme.border_subtle,
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        content_frame.pack(fill="both", expand=True, padx=16, pady=16)

        # Header
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

        v_pill = ctk.CTkFrame(
            title_row,
            fg_color=self.theme.bg_surface,
            border_color=self.theme.border_subtle,
            border_width=1,
            corner_radius=6,
        )
        v_pill.pack(side="left", padx=10)
        ctk.CTkLabel(
            v_pill,
            text=f"v{__version__}",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
            text_color=self.theme.status_ready,
        ).pack(padx=8, pady=2)

        app_desc = ctk.CTkLabel(
            header_frame,
            text="Otomasi Cetak Surat Pemberitahuan Pemilih (Model C6) Pilkades",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=self.theme.text_secondary,
        )
        app_desc.pack(anchor="w", pady=(2, 0))

        # Author Card
        dev_card = ctk.CTkFrame(
            content_frame,
            fg_color=self.theme.bg_surface,
            border_color=self.theme.border_subtle,
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
            text=f"Pengembang: {__author__} (@{__github_user__})",
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
            text="GitHub",
            command=lambda: webbrowser.open(__repository__),
            fg_color=self.theme.btn_secondary_bg,
            hover_color=self.theme.btn_secondary_hover,
            border_color=self.theme.btn_secondary_border,
            border_width=1,
            text_color=self.theme.btn_secondary_text,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            height=28,
            corner_radius=6,
        )
        btn_github.pack(side="right")

        # Feature History
        hist_title = ctk.CTkLabel(
            content_frame,
            text="Riwayat Pembaruan Fitur",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=self.theme.text_primary,
        )
        hist_title.pack(anchor="w", padx=16, pady=(10, 4))

        hist_scroll = ctk.CTkScrollableFrame(
            content_frame,
            fg_color=self.theme.bg_surface,
            border_color=self.theme.border_subtle,
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            height=260,
        )
        hist_scroll.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        for feat in reversed(FEATURE_REGISTRY):
            item = ctk.CTkFrame(hist_scroll, fg_color="transparent")
            item.pack(fill="x", pady=4)

            top = ctk.CTkFrame(item, fg_color="transparent")
            top.pack(fill="x")

            v_lbl = ctk.CTkLabel(
                top,
                text=f"v{feat.get('version', '')}",
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
                text_color=self.theme.status_ready,
            )
            v_lbl.pack(side="left")

            t_lbl = ctk.CTkLabel(
                top,
                text=f"— {feat.get('title', '')}",
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                text_color=self.theme.text_primary,
            )
            t_lbl.pack(side="left", padx=4)

            d_lbl = ctk.CTkLabel(
                top,
                text=feat.get("date", ""),
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=10),
                text_color=self.theme.text_muted,
            )
            d_lbl.pack(side="right")

            desc = feat.get("description", "")
            if desc:
                ctk.CTkLabel(
                    item,
                    text=desc,
                    font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
                    text_color=self.theme.text_secondary,
                    wraplength=480,
                    justify="left",
                ).pack(anchor="w", pady=(1, 0))
