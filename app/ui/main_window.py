# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import math
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
import time
from typing import Any
import webbrowser
import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.config.settings import (
    AppConfig,
    load_app_config,
    save_app_config,
    clean_temp_dir,
    OUTPUT_DIR,
    TEMP_DIR,
)
from app.config.theme import DARK_THEME, LIGHT_THEME, ThemeTokens, get_theme, THEME
from app.config.version import (
    __app_name__,
    __author__,
    __repository__,
    __license__,
    __version__,
)
from app.core.excel_scanner import (
    PreFlightScanResult,
    SheetScanSummary,
    VoterData,
    ExcelSnapshot,
    scan_and_validate_excel,
    compute_excel_snapshot,
    analyze_excel_diff,
    get_excel_file_fingerprint,
)
from app.core.mail_merger import (
    generate_preview_via_word,
    merge_voters_to_pdf_via_word_com,
)
from app.core.pdf_verifier import (
    VerificationResult,
    export_audit_excel,
    verify_kadus_pdf,
)
from app.core.security import sanitize_filename, to_win_abs_path, validate_safe_output_path
from app.core.word_preparer import (
    WordCheckResult,
    WordSnapshot,
    check_and_prepare_word_template,
    compute_word_snapshot,
    analyze_word_diff,
    get_file_fingerprint,
)
from app.ui.components import (
    DocumentSourceRow,
    CleanMetricBar,
    KadusRowItem,
    VoterPreviewTable,
    StatusBadge,
    AboutModal,
    ColumnMapperModal,
)

class MainWindow(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.cfg: AppConfig = load_app_config()
        self.theme: ThemeTokens = get_theme(self.cfg.theme_mode)

        ctk.set_appearance_mode(self.theme.name)
        ctk.set_default_color_theme("blue")

        self.title("VoterMail Engine — Otomasi Surat Pemberitahuan Pemilih")
        self.geometry("1340x940")
        self.minsize(1180, 820)
        self.configure(fg_color=THEME["bg_root"])

        self.worker_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.is_processing = False

        self.word_check_result: WordCheckResult | None = None
        self.excel_scan_result: PreFlightScanResult | None = None
        self.word_ready: bool = False
        self.excel_ready: bool = False

        self.kadus_rows: dict[str, KadusRowItem] = {}
        self.log_history_all: list[str] = []
        self.log_history_preview: list[str] = []
        self.log_history_audit: list[str] = []
        self.current_log_tab = "preview"

        self.start_process_time: float = 0.0

        self.word_fingerprint: tuple[str, float, int] | None = None
        self.excel_fingerprint: tuple[str, float, int] | None = None
        self.word_snapshot: WordSnapshot | None = None
        self.excel_snapshot: ExcelSnapshot | None = None
        self.word_file_dirty: bool = False
        self.excel_file_dirty: bool = False
        self.custom_column_map: dict[str, int] | None = None

        self._build_ui()
        self._poll_worker_queue()

        if self.cfg.word_path and os.path.exists(self.cfg.word_path):
            self.after(300, self._run_word_check)
        if self.cfg.excel_path and os.path.exists(self.cfg.excel_path):
            self.after(600, self._run_excel_scan)

        self.after(3000, self._check_external_file_changes)
        self.protocol("WM_DELETE_WINDOW", self._on_close_window)
        self.bind("<F5>", lambda e: self._on_f5_refresh())

    def _on_f5_refresh(self) -> None:
        if self.is_processing:
            return
        self._run_word_check()
        self._run_excel_scan()

    def _on_close_window(self) -> None:
        if self.is_processing:
            if not messagebox.askyesno(
                "Konfirmasi Keluar",
                "Proses pembuatan dokumen sedang berjalan di latar belakang.\n\nApakah Anda yakin ingin membatalkan proses dan menutup aplikasi?",
            ):
                return
            self.cancel_event.set()

        try:
            clean_temp_dir()
        except Exception:
            pass

        self.destroy()

    def _build_ui(self) -> None:
        # Header Container
        self.header_frame = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=0,
            height=64,
        )
        self.header_frame.pack(side="top", fill="x")

        title_container = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        title_container.pack(side="left", padx=20, pady=8)

        title_top = ctk.CTkFrame(title_container, fg_color="transparent")
        title_top.pack(anchor="w")

        app_title = ctk.CTkLabel(
            title_top,
            text=f"{__app_name__}",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=16, weight="bold"),
            text_color=THEME["text_primary"],
        )
        app_title.pack(side="left")

        self.badge_version = StatusBadge(
            title_top,
            text=f"v{__version__}",
            status="ready",
            theme=self.theme,
        )
        self.badge_version.pack(side="left", padx=10)

        app_sub = ctk.CTkLabel(
            title_container,
            text="Otomasi Surat Pemberitahuan Pemilih Pilkades Karangjengkol",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_secondary"],
        )
        app_sub.pack(anchor="w")

        header_right = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        header_right.pack(side="right", padx=20, pady=10)

        current_val = "☾ Gelap" if self.cfg.theme_mode.lower() == "dark" else "☀ Terang"
        self.theme_segment = ctk.CTkSegmentedButton(
            header_right,
            values=["☀ Terang", "☾ Gelap"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            command=self._on_theme_toggled,
            height=32,
            corner_radius=8,
            fg_color=THEME["bg_surface"],
            selected_color="#2563EB",
            selected_hover_color="#1D4ED8",
            unselected_color=THEME["bg_surface"],
            unselected_hover_color=THEME["btn_secondary_hover"],
            text_color=THEME["text_primary"],
        )
        self.theme_segment.set(current_val)
        self.theme_segment.pack(side="left", padx=(0, 10))

        self.btn_about = ctk.CTkButton(
            header_right,
            text="Tentang & Credit",
            command=self._open_about_modal,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            height=32,
            corner_radius=8,
        )
        self.btn_about.pack(side="left", padx=(0, 10))

        sec_pill = ctk.CTkFrame(
            header_right,
            fg_color=THEME["status_ready_bg"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=8,
            height=32,
        )
        sec_pill.pack(side="left")
        ctk.CTkLabel(
            sec_pill,
            text="Zero PII Leak • 100% Offline",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
            text_color=THEME["status_ready_text"],
        ).pack(padx=10, pady=5)

        # Footer Bar: Author Credits & Quick Links
        self.footer_frame = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=0,
            height=36,
        )
        self.footer_frame.pack(side="bottom", fill="x")

        footer_left = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        footer_left.pack(side="left", padx=16, pady=4)

        lbl_footer_info = ctk.CTkLabel(
            footer_left,
            text=f"{__app_name__} v{__version__} • Dikembangkan oleh {__author__}",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_secondary"],
        )
        lbl_footer_info.pack(side="left")

        footer_right = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        footer_right.pack(side="right", padx=16, pady=4)

        lbl_footer_sec = ctk.CTkLabel(
            footer_right,
            text="100% Pemrosesan Lokal & Aman • GitHub (@Smurffahh)",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_secondary"],
            cursor="hand2",
        )
        lbl_footer_sec.pack(side="left")
        lbl_footer_sec.bind("<Button-1>", lambda _: webbrowser.open(__repository__))

        # Main Workspace Container: 2-Column Responsive Layout without nested outer scrollbars
        main_body = ctk.CTkFrame(self, fg_color=THEME["bg_root"])
        main_body.pack(fill="both", expand=True, padx=14, pady=(10, 6))

        # Left Column (approx 46% width): Source Preparation & Kadus Selection
        self.left_col = ctk.CTkFrame(main_body, fg_color="transparent", width=590)
        self.left_col.pack(side="left", fill="both", expand=False, padx=(0, 10))
        self.left_col.pack_propagate(False)

        # Right Column (approx 54% width): Output Settings, Mailing Engine, Live Logs & Preview Table
        self.right_col = ctk.CTkFrame(main_body, fg_color="transparent")
        self.right_col.pack(side="right", fill="both", expand=True)

        self._build_card_sources(self.left_col)
        self._build_card_kadus(self.left_col)
        self._build_card_exec(self.right_col)
        self._build_card_activity(self.right_col)

    def _build_card_sources(self, parent: ctk.CTkFrame) -> None:
        self.card_sources = ctk.CTkFrame(
            parent,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        self.card_sources.pack(fill="x", pady=(0, 10))

        header = ctk.CTkFrame(self.card_sources, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(12, 8))

        ctk.CTkLabel(
            header,
            text="1. Sumber Dokumen",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=THEME["text_primary"],
        ).pack(side="left")

        self.doc_row_word = DocumentSourceRow(
            self.card_sources,
            doc_type="Dokumen Master Word",
            initial_path=self.cfg.word_path,
            theme=self.theme,
            on_browse=self._browse_word_file,
            on_recheck=self._run_word_check,
            on_preview=self._generate_preview,
            recheck_label="Cek Ulang",
        )
        self.doc_row_word.pack(fill="x", padx=12, pady=(0, 8))

        self.doc_row_excel = DocumentSourceRow(
            self.card_sources,
            doc_type="Data Pemilih Excel DPT",
            initial_path=self.cfg.excel_path,
            theme=self.theme,
            on_browse=self._browse_excel_file,
            on_recheck=self._run_excel_scan,
            on_map_columns=self._open_column_mapper,
            recheck_label="Scan Ulang",
        )
        self.doc_row_excel.pack(fill="x", padx=12, pady=(0, 12))

    def _build_card_kadus(self, parent: ctk.CTkFrame) -> None:
        self.card_kadus = ctk.CTkFrame(
            parent,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        self.card_kadus.pack(fill="both", expand=True)

        header = ctk.CTkFrame(self.card_kadus, fg_color="transparent")
        header.pack(fill="x", padx=14, pady=(12, 8))

        ctk.CTkLabel(
            header,
            text="2. Pilih Wilayah Cetak (Kadus / RW)",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=THEME["text_primary"],
        ).pack(side="left")

        btn_box = ctk.CTkFrame(header, fg_color="transparent")
        btn_box.pack(side="right")

        self.btn_select_all = ctk.CTkButton(
            btn_box,
            text="Pilih Semua",
            width=76,
            height=28,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["btn_secondary_text"],
            command=self._select_all_sheets,
        )
        self.btn_select_all.pack(side="left", padx=(0, 6))

        self.btn_deselect_all = ctk.CTkButton(
            btn_box,
            text="Reset",
            width=64,
            height=28,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["btn_secondary_text"],
            command=self._deselect_all_sheets,
        )
        self.btn_deselect_all.pack(side="left")

        self.metric_bar = CleanMetricBar(self.card_kadus, self.theme)
        self.metric_bar.pack(fill="x", padx=12, pady=(0, 8))

        self.sheets_scroll_frame = ctk.CTkScrollableFrame(
            self.card_kadus,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
        )
        self.sheets_scroll_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        self.excel_status_bar = ctk.CTkLabel(
            self.card_kadus,
            text="Belum ada file Excel yang dipindai.",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_muted"],
            anchor="w",
        )
        self.excel_status_bar.pack(fill="x", padx=14, pady=(0, 10))

    def _build_card_exec(self, parent: ctk.CTkFrame) -> None:
        self.card_exec = ctk.CTkFrame(
            parent,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        self.card_exec.pack(fill="x", pady=(0, 10))

        header = ctk.CTkFrame(self.card_exec, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(12, 8))

        ctk.CTkLabel(
            header,
            text="3. Pengaturan & Cetak PDF",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=THEME["text_primary"],
        ).pack(side="left")

        self.badge_mailing = StatusBadge(header, "Menunggu Word & Excel", "neutral", self.theme)
        self.badge_mailing.pack(side="right")

        out_row = ctk.CTkFrame(self.card_exec, fg_color="transparent")
        out_row.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkLabel(
            out_row,
            text="Folder Hasil PDF:",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_secondary"],
        ).pack(side="left", padx=(0, 8))

        self.output_entry = ctk.CTkEntry(
            out_row,
            placeholder_text="Pilih direktori output...",
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_primary"],
            height=30,
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.output_entry.insert(0, self.cfg.output_dir)

        self.btn_browse_output = ctk.CTkButton(
            out_row,
            text="Ubah Folder",
            width=86,
            height=30,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["btn_secondary_text"],
            command=self._browse_output_dir,
        )
        self.btn_browse_output.pack(side="left", padx=(0, 6))

        self.btn_open_output = ctk.CTkButton(
            out_row,
            text="Buka Folder",
            width=86,
            height=30,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["btn_secondary_text"],
            command=self._open_output_dir,
        )
        self.btn_open_output.pack(side="left")

        ctrl_row = ctk.CTkFrame(self.card_exec, fg_color="transparent")
        ctrl_row.pack(fill="x", padx=16, pady=(0, 10))

        addr_box = ctk.CTkFrame(ctrl_row, fg_color="transparent")
        addr_box.pack(side="left")

        ctk.CTkLabel(
            addr_box,
            text="Format Alamat:",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_secondary"],
        ).pack(side="left", padx=(0, 8))

        self.address_mode_menu = ctk.CTkOptionMenu(
            addr_box,
            values=["Lengkap (Alamat, RT / RW)", "Hanya Alamat"],
            width=210,
            height=28,
            fg_color=THEME["btn_secondary_bg"],
            button_color=THEME["btn_secondary_hover"],
            button_hover_color=THEME["btn_secondary_hover"],
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_primary"],
            dropdown_fg_color=THEME["bg_card"],
            dropdown_text_color=THEME["text_primary"],
            dropdown_hover_color=THEME["bg_hover"],
            command=lambda _: self._save_active_config(),
        )
        self.address_mode_menu.set(self.cfg.address_mode)
        self.address_mode_menu.pack(side="left")

        verify_box = ctk.CTkFrame(ctrl_row, fg_color="transparent")
        verify_box.pack(side="right")

        self.verify_switch = ctk.CTkSwitch(
            verify_box,
            text="Periksa Keakuratan PDF Otomatis",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_primary"],
            progress_color=THEME["btn_primary_bg"],
            command=self._save_active_config,
        )
        if self.cfg.enable_pdf_verification:
            self.verify_switch.select()
        else:
            self.verify_switch.deselect()
        self.verify_switch.pack(side="right")

        btn_action_row = ctk.CTkFrame(self.card_exec, fg_color="transparent")
        btn_action_row.pack(fill="x", padx=16, pady=(0, 8))

        self.btn_start = ctk.CTkButton(
            btn_action_row,
            text="Buat File PDF Sekarang",
            height=38,
            fg_color=THEME["btn_disabled_bg"],
            hover_color=THEME["btn_primary_hover"],
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=THEME["btn_disabled_text"],
            state="disabled",
            command=self._start_mailing_process,
        )
        self.btn_start.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_cancel = ctk.CTkButton(
            btn_action_row,
            text="Batalkan",
            width=100,
            height=38,
            fg_color=THEME["btn_disabled_bg"],
            hover_color=THEME["btn_danger_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=12, weight="bold"),
            text_color=THEME["btn_disabled_text"],
            state="disabled",
            command=self._cancel_mailing,
        )
        self.btn_cancel.pack(side="left")

        prog_box = ctk.CTkFrame(
            self.card_exec,
            fg_color=THEME["bg_elevated"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
        )
        prog_box.pack(fill="x", padx=16, pady=(0, 12))

        prog_inner = ctk.CTkFrame(prog_box, fg_color="transparent")
        prog_inner.pack(fill="x", padx=10, pady=8)

        self.progress_bar = ctk.CTkProgressBar(
            prog_inner,
            height=10,
            corner_radius=5,
            progress_color=THEME["btn_primary_bg"],
            fg_color=THEME["border_card"],
        )
        self.progress_bar.set(0.0)
        self.progress_bar.pack(fill="x", pady=(0, 6))

        stats_row = ctk.CTkFrame(prog_inner, fg_color="transparent")
        stats_row.pack(fill="x")

        self.status_detail_label = ctk.CTkLabel(
            stats_row,
            text="Siap Memproses: 0 Pemilih",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
            text_color=THEME["text_primary"],
        )
        self.status_detail_label.pack(side="left")

        self.eta_label = ctk.CTkLabel(
            stats_row,
            text="Estimasi Sisa Waktu: -",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_secondary"],
        )
        self.eta_label.pack(side="right")

    def _build_card_activity(self, parent: ctk.CTkFrame) -> None:
        self.card_activity = ctk.CTkFrame(
            parent,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        self.card_activity.pack(fill="both", expand=True)

        header = ctk.CTkFrame(self.card_activity, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(12, 8))

        ctk.CTkLabel(
            header,
            text="Ringkasan Data & Aktivitas",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=THEME["text_primary"],
        ).pack(side="left")

        ctrl_right = ctk.CTkFrame(header, fg_color="transparent")
        ctrl_right.pack(side="right")

        self.activity_tabs = ctk.CTkSegmentedButton(
            ctrl_right,
            values=["Pratinjau Pemilih", "Hasil Cek PDF", "Log Sistem"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            command=self._on_activity_tab_selected,
            height=28,
            corner_radius=6,
            fg_color=THEME["bg_surface"],
            selected_color="#2563EB",
            selected_hover_color="#1D4ED8",
            unselected_color=THEME["bg_surface"],
            unselected_hover_color=THEME["btn_secondary_hover"],
            text_color=THEME["text_primary"],
        )
        self.activity_tabs.set("Pratinjau Pemilih")
        self.activity_tabs.pack(side="left", padx=(0, 8))

        self.btn_clear_log = ctk.CTkButton(
            ctrl_right,
            text="Bersihkan",
            width=70,
            height=28,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
            text_color=THEME["btn_secondary_text"],
            command=self._clear_logs,
        )
        self.btn_clear_log.pack(side="left")

        self.activity_content_frame = ctk.CTkFrame(self.card_activity, fg_color="transparent")
        self.activity_content_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        self.voter_table = VoterPreviewTable(self.activity_content_frame, self.theme)

        self.audit_textbox = ctk.CTkTextbox(
            self.activity_content_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_primary"],
            wrap="char",
        )

        self.log_textbox = ctk.CTkTextbox(
            self.activity_content_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_primary"],
            wrap="char",
        )

        self.voter_table.pack(fill="both", expand=True)

    def _browse_word_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Pilih File Template Word",
            filetypes=[("Dokumen Word", "*.doc;*.docx;*.rtf"), ("Semua File", "*.*")],
        )
        if path:
            self.cfg.word_path = path
            self.doc_row_word.set_file(path)
            self._save_active_config()
            self._run_word_check()

    def _browse_excel_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Pilih File Excel DPT",
            filetypes=[("File Excel", "*.xlsx;*.xls"), ("Semua File", "*.*")],
        )
        if path:
            self.cfg.excel_path = path
            self.doc_row_excel.set_file(path)
            self._save_active_config()
            self._run_excel_scan()

    def _browse_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Pilih Folder Output PDF")
        if path:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)
            self._save_active_config()

    def _open_output_dir(self) -> None:
        folder = self.output_entry.get().strip() or str(OUTPUT_DIR.resolve())
        if os.path.exists(folder):
            try:
                os.startfile(folder)
            except Exception as exc:
                messagebox.showerror("Error", f"Gagal membuka folder: {exc}")
        else:
            messagebox.showinfo("Informasi", "Folder output belum dibuat di disk.")

    def _save_active_config(self) -> None:
        self.cfg.word_path = self.doc_row_word.current_path
        self.cfg.excel_path = self.doc_row_excel.current_path
        self.cfg.output_dir = self.output_entry.get().strip()
        self.cfg.address_mode = self.address_mode_menu.get().strip()
        self.cfg.enable_pdf_verification = self.verify_switch.get() == 1
        self.cfg.theme_mode = self.theme.name
        save_app_config(self.cfg)

    def _on_theme_toggled(self, val: str) -> None:
        mode = "Dark" if "gelap" in val.lower() or val.lower() == "dark" else "Light"
        ctk.set_appearance_mode(mode)
        self.theme = get_theme(mode)
        self.cfg.theme_mode = mode
        save_app_config(self.cfg)
        self._check_ready_state()
        self._log(f"[TEMA] Mengganti tema visual ke: {mode}")

    def _log(self, text: str, category: str = "all") -> None:
        timestamp = time.strftime("%H:%M:%S")
        entry = f"[{timestamp}] {text}"
        self.log_history_all.append(entry)

        if category == "preview":
            self.log_history_preview.append(entry)
        elif category == "audit":
            self.log_history_audit.append(entry)

        if self.current_log_tab == "all":
            self.log_textbox.insert("end", entry + "\n")
            self.log_textbox.see("end")
        elif self.current_log_tab == "audit" and category == "audit":
            self.audit_textbox.insert("end", entry + "\n")
            self.audit_textbox.see("end")

    def _clear_logs(self) -> None:
        if self.current_log_tab == "audit":
            self.log_history_audit.clear()
            self.audit_textbox.delete("1.0", "end")
        else:
            self.log_history_all.clear()
            self.log_textbox.delete("1.0", "end")

    def _on_activity_tab_selected(self, val: str) -> None:
        self.voter_table.pack_forget()
        self.audit_textbox.pack_forget()
        self.log_textbox.pack_forget()

        if val == "Pratinjau Pemilih":
            self.current_log_tab = "preview"
            self.voter_table.pack(fill="both", expand=True)
        elif val == "Hasil Cek PDF":
            self.current_log_tab = "audit"
            self.audit_textbox.pack(fill="both", expand=True)
            self.audit_textbox.delete("1.0", "end")
            content = "\n".join(self.log_history_audit) or "Belum ada laporan audit PDF. Jalankan proses mailing terlebih dahulu."
            self.audit_textbox.insert("end", content + "\n")
            self.audit_textbox.see("end")
        else:
            self.current_log_tab = "all"
            self.log_textbox.pack(fill="both", expand=True)
            self.log_textbox.delete("1.0", "end")
            content = "\n".join(self.log_history_all) or "Belum ada riwayat log aktivitas."
            self.log_textbox.insert("end", content + "\n")
            self.log_textbox.see("end")

    def _run_word_check(self) -> None:
        if self.is_processing:
            return
        path = self.cfg.word_path
        if not path or not os.path.exists(path):
            self.doc_row_word.set_status("Berkas Word tidak ditemukan di path yang ditentukan.", "error")
            return

        self.doc_row_word.set_status("Memeriksa struktur dokumen...", "process")
        self._log(f"[WORD] Memeriksa struktur dokumen: {os.path.basename(path)}")

        def worker() -> None:
            try:
                res = check_and_prepare_word_template(path, work_dir=TEMP_DIR)
                self.worker_queue.put(("word_done", res))
            except Exception as exc:
                self.worker_queue.put(("word_error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_word_results(self, res: WordCheckResult) -> None:
        self.word_check_result = res
        path = self.cfg.word_path
        self.doc_row_word.set_file(path)

        new_snap = compute_word_snapshot(
            res.prepared_docx_path,
            path,
            is_valid=res.is_valid,
            detected_slot1=res.detected_fields_slot1,
            detected_slot2=res.detected_fields_slot2,
        )
        if self.word_snapshot is not None and self.word_file_dirty:
            diff_lines = analyze_word_diff(self.word_snapshot, new_snap)
            for line in diff_lines:
                self._log(line)

        self.word_snapshot = new_snap
        self.word_fingerprint = (new_snap.sha256, new_snap.mtime, new_snap.size)
        self.word_file_dirty = False

        if res.is_valid:
            self.word_ready = True
            mode_desc = "2 Surat/Lembar" if res.layout_mode == "2-UP" else "1 Surat/Lembar"
            self.doc_row_word.set_status(f"Siap Mailing • Mode {mode_desc}", "ready")
            self._log(f"[WORD SIAP] {res.status_badge}: {res.prepared_docx_path}")
        else:
            self.word_ready = False
            self.doc_row_word.set_status(f"Perlu Perbaikan: {res.message}", "error")
            self._log(f"[WORD ERROR] {res.message}")

        self._update_selected_metrics()
        self._check_ready_state()

    def _open_column_mapper(self) -> None:
        if not self.excel_scan_result or not self.excel_scan_result.available_columns:
            messagebox.showinfo(
                "Informasi",
                "Pilih berkas Excel DPT terlebih dahulu agar aplikasi dapat membaca daftar kolom yang tersedia.",
            )
            return

        def on_applied(new_mapping: dict[str, int]) -> None:
            self.custom_column_map = new_mapping
            self._log(f"[KOLOM DPT] Pemetaan kolom manual diterapkan: {new_mapping}")
            self._run_excel_scan()

        hdr_idx = self.excel_scan_result.sheet_summaries[0].header_row_index if self.excel_scan_result.sheet_summaries else 0
        ColumnMapperModal(
            parent=self,
            theme=self.theme,
            available_columns=self.excel_scan_result.available_columns,
            current_mapping=self.custom_column_map or self.excel_scan_result.column_mapping,
            on_apply=on_applied,
            header_row_index=hdr_idx,
        )

    def _run_excel_scan(self) -> None:
        if self.is_processing:
            return
        path = self.cfg.excel_path
        if not path or not os.path.exists(path):
            self.doc_row_excel.set_status("Berkas Excel tidak ditemukan.", "error")
            self.excel_status_bar.configure(text="File Excel tidak ditemukan.", text_color=self.theme.status_error)
            return

        addr_mode = "lengkap" if "Lengkap" in self.address_mode_menu.get() else "hanya_alamat"
        custom_map = self.custom_column_map
        self.doc_row_excel.set_status("Memindai struktur DPT...", "process")
        self._log(f"[EXCEL] Memindai struktur DPT: {os.path.basename(path)}")

        def worker() -> None:
            try:
                res = scan_and_validate_excel(path, address_mode=addr_mode, custom_column_map=custom_map)
                self.worker_queue.put(("excel_done", res))
            except Exception as exc:
                self.worker_queue.put(("excel_error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_excel_results(self, res: PreFlightScanResult) -> None:
        self.excel_scan_result = res
        path = self.cfg.excel_path
        self.doc_row_excel.set_file(path)

        new_snap = compute_excel_snapshot(res, path)
        if self.excel_snapshot is not None and self.excel_file_dirty:
            diff_lines = analyze_excel_diff(self.excel_snapshot, new_snap)
            for line in diff_lines:
                self._log(line)

        self.excel_snapshot = new_snap
        self.excel_fingerprint = (new_snap.sha256, new_snap.mtime, new_snap.size)
        self.excel_file_dirty = False

        if res.schema_message:
            self._log(f"[EXCEL FORMAT] {res.schema_message}")

        for child in self.sheets_scroll_frame.winfo_children():
            child.destroy()
        self.kadus_rows.clear()

        voter_sheets = [s for s in res.sheet_summaries if s.is_voter_sheet]

        if not voter_sheets:
            self.excel_ready = False
            self.doc_row_excel.set_status("Tidak ditemukan Sheet DPT yang valid.", "error")
            self.excel_status_bar.configure(text="Tidak ditemukan Sheet DPT yang valid.", text_color=self.theme.status_error)
            self._check_ready_state()
            return

        layout_1up = bool(self.word_check_result and self.word_check_result.layout_mode == "1-UP")

        for s in voter_sheets:
            page_count = s.total_voters if layout_1up else s.expected_pages
            row = KadusRowItem(
                self.sheets_scroll_frame,
                sheet_name=s.sheet_name,
                voter_count=s.total_voters,
                page_count=page_count,
                is_selected=True,
                theme=self.theme,
                on_toggle=self._update_selected_metrics,
            )
            row.pack(fill="x", pady=2, padx=4)
            self.kadus_rows[s.sheet_name] = row

        sample_voters: list[VoterData] = []
        for s in voter_sheets:
            if s.voters:
                sample_voters.extend(s.voters)
                if len(sample_voters) >= 5:
                    break
        self.voter_table.set_voters(sample_voters[:5])

        self.excel_ready = True
        if res.is_valid:
            self.doc_row_excel.set_status(f"Siap Mailing • {res.total_sheets} Kadus ({res.total_voters:,} pemilih)", "ready")
            self.excel_status_bar.configure(
                text=f"✓ Ditemukan {res.total_sheets} Sheet DPT valid ({res.total_voters:,} pemilih). Zero Error.",
                text_color=self.theme.status_ready,
            )
        elif res.schema_type == "PERLU_PEMETAAN_MANUAL":
            self.doc_row_excel.set_status("Perlu Pemetaan Kolom • Klik 'Atur Kolom'", "warning")
            self.excel_status_bar.configure(
                text="Perhatian: Kolom Nama Pemilih belum terdeteksi. Silakan klik 'Atur Kolom'.",
                text_color=self.theme.status_warning,
            )
        else:
            self.doc_row_excel.set_status(f"Perlu Perhatian • Ditemukan {len(res.all_issues)} anomali", "warning")
            self.excel_status_bar.configure(
                text=f"Perhatian: Ditemukan {len(res.all_issues)} anomali pada data pemilih.",
                text_color=self.theme.status_warning,
            )

        if res.preview_samples:
            self._log("=== PRATINJAU DATA PEMILIH (5 BARIS PERTAMA) ===", category="preview")
            for sample in res.preview_samples:
                self._log(sample, category="preview")

        self._update_selected_metrics()
        self._check_ready_state()

    def _check_external_file_changes(self) -> None:
        if not self.is_processing:
            word_path = self.cfg.word_path
            if word_path and os.path.exists(word_path) and self.word_fingerprint:
                curr_fp = get_file_fingerprint(word_path)
                if curr_fp != self.word_fingerprint and not self.word_file_dirty:
                    self.word_file_dirty = True
                    self.doc_row_word.set_status("[File Diperbarui - Klik Cek Ulang]", "warning")
                    self.word_ready = False
                    self.word_check_result = None
                    ready_docx = TEMP_DIR / "template_ready.docx"
                    if ready_docx.exists():
                        try:
                            ready_docx.unlink(missing_ok=True)
                        except Exception:
                            pass
                    self._check_ready_state()

            excel_path = self.cfg.excel_path
            if excel_path and os.path.exists(excel_path) and self.excel_fingerprint:
                curr_fp = get_excel_file_fingerprint(excel_path)
                if curr_fp != self.excel_fingerprint and not self.excel_file_dirty:
                    self.excel_file_dirty = True
                    self.doc_row_excel.set_status("[File Diperbarui - Klik Scan Ulang]", "warning")
                    self.excel_ready = False
                    self.excel_scan_result = None
                    self._check_ready_state()

        self.after(3000, self._check_external_file_changes)

    def _select_all_sheets(self) -> None:
        for row in self.kadus_rows.values():
            row.checkbox_var.set(True)
        self._update_selected_metrics()

    def _deselect_all_sheets(self) -> None:
        for row in self.kadus_rows.values():
            row.checkbox_var.set(False)
        self._update_selected_metrics()

    def _update_selected_metrics(self) -> None:
        if not self.excel_scan_result:
            return

        selected_sheets = [
            s for s in self.excel_scan_result.sheet_summaries
            if s.sheet_name in self.kadus_rows and self.kadus_rows[s.sheet_name].checkbox_var.get()
        ]

        layout = self.word_check_result.layout_mode if self.word_check_result else "2-UP"
        kadus_count = len(selected_sheets)
        total_voters = sum(s.total_voters for s in selected_sheets)
        total_pages = total_voters if layout == "1-UP" else sum(s.expected_pages for s in selected_sheets)

        self.metric_bar.update_metrics(kadus_count, total_voters, total_pages)

        self.status_detail_label.configure(
            text=f"Siap Memproses: 0 / {total_voters:,} Pemilih ({kadus_count} Kadus)".replace(",", ".")
        )

        self._check_ready_state()

    def _check_ready_state(self) -> None:
        selected_count = sum(1 for row in self.kadus_rows.values() if row.checkbox_var.get())

        if self.is_processing:
            return

        if self.word_ready and self.excel_ready and selected_count > 0:
            self.btn_start.configure(
                state="normal",
                text="Buat File PDF Sekarang",
                fg_color=THEME["btn_primary_bg"],
                hover_color=THEME["btn_primary_hover"],
                text_color=THEME["btn_primary_text"],
            )
            self.badge_mailing.set_status("Siap Mailing", "ready")
        elif selected_count == 0 and self.word_ready and self.excel_ready:
            self.btn_start.configure(
                state="disabled",
                text="Pilih Minimal 1 Sheet Kadus",
                fg_color=THEME["btn_disabled_bg"],
                text_color=THEME["btn_disabled_text"],
            )
            self.badge_mailing.set_status("Pilih Kadus", "neutral")
        else:
            self.btn_start.configure(
                state="disabled",
                text="Buat File PDF Sekarang",
                fg_color=THEME["btn_disabled_bg"],
                text_color=THEME["btn_disabled_text"],
            )
            self.badge_mailing.set_status("Menunggu Word & Excel", "neutral")

    def _generate_preview(self) -> None:
        if not self.word_check_result or not self.word_check_result.is_valid:
            messagebox.showwarning("Peringatan", "Jalankan 'Cek Ulang' pada berkas Word terlebih dahulu.")
            return

        out_dir = self.output_entry.get().strip() or str(OUTPUT_DIR.resolve())
        os.makedirs(out_dir, exist_ok=True)
        preview_pdf = to_win_abs_path(os.path.join(out_dir, "Pratinjau_Surat_Word_1_Halaman.pdf"))

        sample_voters: list[VoterData] = []
        if self.excel_scan_result:
            for s in self.excel_scan_result.sheet_summaries:
                if s.voters:
                    sample_voters = s.voters[:2]
                    v2_nama = sample_voters[1].nama if len(sample_voters) > 1 else "-"
                    self._log(f"[PREVIEW] Mengambil 2 data pemilih riil dari '{s.sheet_name}': {sample_voters[0].nama} & {v2_nama}")
                    break

        if not sample_voters:
            self._log("[PREVIEW] Excel belum di-scan, menggunakan data contoh lengkap (2 pemilih)...")

        template_doc = to_win_abs_path(self.word_check_result.prepared_docx_path)
        layout = self.word_check_result.layout_mode if self.word_check_result else "2-UP"
        self._log(f"[PREVIEW] Membuat pratinjau 1 halaman ({layout}) menggunakan MS Word COM...")
        ok, msg, final_pdf = generate_preview_via_word(template_doc, sample_voters, preview_pdf, layout_mode=layout)

        if ok:
            self._log(f"[PREVIEW SELESAI] Pratinjau tersimpan di: {final_pdf}")
            try:
                os.startfile(final_pdf)
            except Exception:
                pass
            messagebox.showinfo("Pratinjau Selesai", f"File pratinjau berhasil dibuat:\n{final_pdf}")
        else:
            messagebox.showerror("Gagal Pratinjau", f"Gagal membuat pratinjau Word: {msg}")

    def _start_mailing_process(self) -> None:
        if self.is_processing or not self.word_check_result or not self.excel_scan_result:
            return

        selected_sheets = [
            s for s in self.excel_scan_result.sheet_summaries
            if s.sheet_name in self.kadus_rows and self.kadus_rows[s.sheet_name].checkbox_var.get()
        ]
        if not selected_sheets:
            messagebox.showwarning("Peringatan", "Pilih minimal 1 Sheet Kadus untuk dicetak.")
            return

        output_dir = self.output_entry.get().strip() or str(OUTPUT_DIR.resolve())
        os.makedirs(output_dir, exist_ok=True)

        self.is_processing = True
        self.cancel_event.clear()
        self.start_process_time = time.time()

        self.btn_start.configure(
            state="disabled",
            text="Sedang Memproses Mailing...",
            fg_color=THEME["btn_disabled_bg"],
            text_color=THEME["btn_disabled_text"],
        )
        self.btn_cancel.configure(
            state="normal",
            fg_color=THEME["btn_danger_bg"],
            hover_color=THEME["btn_danger_hover"],
            text_color=THEME["btn_danger_text"],
        )
        for row in self.kadus_rows.values():
            row.set_enabled(False)
            row.reset_status()

        self.badge_mailing.set_status("Memproses...", "process")
        self.progress_bar.set(0.0)

        total_voters_all = sum(s.total_voters for s in selected_sheets)
        template_doc = self.word_check_result.prepared_docx_path
        verify_pdf_on = self.verify_switch.get() == 1
        layout_mode = self.word_check_result.layout_mode if self.word_check_result else "2-UP"

        self._log(f"[MAILING MULAI] Memulai proses mailing {len(selected_sheets)} Kadus ({total_voters_all:,} pemilih, Mode: {layout_mode})...")

        def worker() -> None:
            processed_voters_acc = 0
            audit_results: list[VerificationResult] = []

            for s in selected_sheets:
                if self.cancel_event.is_set():
                    self.worker_queue.put(("process_cancelled", "Dibatalkan oleh pengguna."))
                    return

                self.worker_queue.put(("kadus_started", s.sheet_name))
                safe_name = sanitize_filename(f"SURAT_PEMBERITAHUAN_{s.sheet_name}")
                out_pdf = os.path.join(output_dir, f"{safe_name}.pdf")

                def prog_cb(cur: int, tot: int, sheet: str) -> None:
                    overall = processed_voters_acc + cur
                    self.worker_queue.put(("progress_update", (overall, total_voters_all, sheet)))

                ok, pages, msg = merge_voters_to_pdf_via_word_com(
                    template_docx_path=template_doc,
                    voters=s.voters,
                    sheet_name=s.sheet_name,
                    wilayah_name=s.wilayah_text,
                    output_pdf_path=out_pdf,
                    progress_callback=prog_cb,
                    cancel_event=self.cancel_event,
                    layout_mode=layout_mode,
                )

                if not ok:
                    self.worker_queue.put(("kadus_failed", (s.sheet_name, msg)))
                    return

                processed_voters_acc += s.total_voters

                if verify_pdf_on and not self.cancel_event.is_set():
                    self.worker_queue.put(("verify_started", s.sheet_name))
                    ver = verify_kadus_pdf(out_pdf, s.sheet_name, s.voters, layout_mode=layout_mode)
                    audit_results.append(ver)
                    self.worker_queue.put(("verify_done", (ver, out_pdf)))
                else:
                    self.worker_queue.put(("kadus_completed", (s.sheet_name, out_pdf)))

            if audit_results and verify_pdf_on:
                audit_file = export_audit_excel(audit_results, output_dir)
                self.worker_queue.put(("audit_exported", audit_file))

            self.worker_queue.put(("all_completed", (len(selected_sheets), total_voters_all)))

        threading.Thread(target=worker, daemon=True).start()

    def _cancel_mailing(self) -> None:
        if self.is_processing:
            self.cancel_event.set()
            self._log("[BATAL] Mengirim sinyal pembatalan ke proses latar belakang...")
            self.btn_cancel.configure(state="disabled", text="Membatalkan...")

    def _poll_worker_queue(self) -> None:
        try:
            while True:
                msg_type, payload = self.worker_queue.get_nowait()

                if msg_type == "word_done":
                    self._apply_word_results(payload)
                elif msg_type == "word_error":
                    self.word_ready = False
                    self.doc_row_word.set_status(f"Gagal: {payload}", "error")
                    self._log(f"[WORD ERROR] {payload}")
                    messagebox.showerror("Error Word", f"Gagal memeriksa file Word:\n{payload}")
                elif msg_type == "excel_done":
                    self._apply_excel_results(payload)
                elif msg_type == "excel_error":
                    self.excel_ready = False
                    self.doc_row_excel.set_status(f"Gagal: {payload}", "error")
                    self._log(f"[EXCEL ERROR] {payload}")
                    messagebox.showerror("Error Excel", f"Gagal membaca file Excel:\n{payload}")

                elif msg_type == "kadus_started":
                    sheet_name = payload
                    if sheet_name in self.kadus_rows:
                        self.kadus_rows[sheet_name].set_printing_active()
                    self._log(f"[CETAK] Sedang memproses PDF untuk {sheet_name}...")

                elif msg_type == "progress_update":
                    cur, tot, sheet = payload
                    pct = cur / tot if tot > 0 else 0.0
                    self.progress_bar.set(pct)
                    self.status_detail_label.configure(
                        text=f"Selesai: {cur:,} / {tot:,} Pemilih ({pct * 100:.1f}%) • {sheet}".replace(",", ".")
                    )
                    elapsed = time.time() - self.start_process_time
                    if cur > 0 and elapsed > 1:
                        speed = cur / elapsed
                        rem = tot - cur
                        eta_sec = rem / speed if speed > 0 else 0
                        mins, secs = divmod(int(eta_sec), 60)
                        self.eta_label.configure(text=f"Estimasi Sisa: {mins:02d}:{secs:02d} ({speed:.1f} pemilih/dtk)")

                elif msg_type == "kadus_completed":
                    sheet_name, out_pdf = payload
                    if sheet_name in self.kadus_rows:
                        self.kadus_rows[sheet_name].set_completed(out_pdf)
                    self._log(f"[SUKSES] PDF {sheet_name} selesai dibuat: {os.path.basename(out_pdf)}")

                elif msg_type == "kadus_failed":
                    sheet_name, err = payload
                    if sheet_name in self.kadus_rows:
                        self.kadus_rows[sheet_name].set_error("• Gagal")
                    self._on_process_finished()
                    self._log(f"[ERROR] Gagal memproses {sheet_name}: {err}")
                    messagebox.showerror("Gagal Mailing", f"Terjadi kesalahan saat memproses {sheet_name}:\n{err}")

                elif msg_type == "verify_started":
                    self._log(f"[AUDIT] Memverifikasi integritas PDF: {payload}...")

                elif msg_type == "verify_done":
                    ver, out_pdf = payload
                    if ver.sheet_name in self.kadus_rows:
                        if ver.is_valid:
                            self.kadus_rows[ver.sheet_name].set_completed(out_pdf)
                        else:
                            self.kadus_rows[ver.sheet_name].set_error("• Anomali")
                    self._log(ver.audit_summary, category="audit")

                elif msg_type == "audit_exported":
                    self._log(f"[AUDIT SELESAI] Laporan audit tersimpan di: {payload}", category="audit")

                elif msg_type == "process_cancelled":
                    self._on_process_finished()
                    self._log(f"[BATAL] {payload}")
                    messagebox.showinfo("Dibatalkan", "Proses mailing telah dihentikan pengguna.")

                elif msg_type == "all_completed":
                    kadus_count, total_v = payload
                    self._on_process_finished()
                    elapsed = time.time() - self.start_process_time
                    mins, secs = divmod(int(elapsed), 60)
                    self.progress_bar.set(1.0)
                    self.eta_label.configure(text=f"Total Waktu: {mins:02d}:{secs:02d}")
                    self.status_detail_label.configure(
                        text=f"Selesai 100%: {total_v:,} Pemilih ({kadus_count} Kadus)".replace(",", ".")
                    )
                    self._log(f"[SELESAI 100%] Berhasil membuat {kadus_count} file PDF ({total_v:,} pemilih) dalam {mins:02d}:{secs:02d}.")
                    clean_temp_dir()
                    messagebox.showinfo(
                        "Mailing Selesai",
                        f"Semua surat panggilan ({kadus_count} Kadus, {total_v:,} pemilih) berhasil dibuat di folder:\n{self.output_entry.get()}"
                    )

        except queue.Empty:
            pass

        self.after(100, self._poll_worker_queue)

    def _on_process_finished(self) -> None:
        self.is_processing = False
        self.btn_cancel.configure(
            state="disabled",
            fg_color=THEME["btn_disabled_bg"],
            text_color=THEME["btn_disabled_text"],
            text="Batalkan",
        )
        for row in self.kadus_rows.values():
            row.set_enabled(True)
        self._check_ready_state()

    def _open_about_modal(self) -> None:
        AboutModal(self, self.theme)
