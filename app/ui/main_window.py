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
from typing import Any, Literal
import webbrowser
import customtkinter as ctk
from tkinter import filedialog

from app.config.settings import (
    AppConfig,
    load_app_config,
    save_app_config,
    clean_temp_dir,
    OUTPUT_DIR,
    TEMP_DIR,
    ADDRESS_FORMAT_OPTIONS,
    DEFAULT_ADDRESS_FORMAT,
    FILTER_MODE_ALL,
    FILTER_MODE_NUMBER,
    FILTER_MODE_RT,
    FILTER_MODE_NAME,
    FILTER_MODE_OPTIONS,
    DEFAULT_FILTER_MODE,
)
from app.config.theme import (
    DARK_THEME,
    LIGHT_THEME,
    ThemeTokens,
    get_theme,
    THEME,
    load_bundled_fonts,
    FONT_FAMILY_UI,
    FONT_FAMILY_MONO,
)
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
    apply_voter_filter,
    build_filter_filename_suffix,
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
from app.core.security import (
    sanitize_filename,
    to_win_abs_path,
    validate_safe_output_path,
    translate_exception_to_user_message,
)
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
    FilterStatusBanner,
    AboutModal,
    ColumnMapperModal,
    UpdateNotificationModal,
    ModernNotificationModal,
    InAppToastBanner,
    apply_window_icon,
)

def format_elapsed_time(seconds: float) -> str:
    total_secs = max(0, int(seconds))
    hours, remainder = divmod(total_secs, 3600)
    mins, secs = divmod(remainder, 60)
    if hours > 0:
        return f"{hours:02d}:{mins:02d}:{secs:02d}"
    return f"{mins:02d}:{secs:02d}"

def shorten_path(path: str | Path | None, max_len: int = 42) -> str:
    if not path:
        return ""
    p = Path(path)
    try:
        rel = p.relative_to(Path.cwd())
        res = str(rel)
    except Exception:
        parts = p.parts
        if len(parts) > 3:
            res = f"...\\{parts[-2]}\\{parts[-1]}"
        else:
            res = str(p)
    if len(res) > max_len:
        res = f"...\\{p.name}"
    return res

def format_voter_preview_table(samples: list[str]) -> list[str]:
    w_no, w_nama, w_jk, w_umur, w_alamat = 4, 24, 4, 5, 24
    def clip(t: Any, w: int) -> str:
        s = str(t).strip()
        return (s[:w-3] + "...") if len(s) > w else s.ljust(w)

    top = "┌─" + "─"*w_no + "─┬─" + "─"*w_nama + "─┬─" + "─"*w_jk + "─┬─" + "─"*w_umur + "─┬─" + "─"*w_alamat + "─┐"
    hdr = "│ " + clip("No", w_no) + " │ " + clip("Nama", w_nama) + " │ " + clip("L/P", w_jk) + " │ " + clip("Usia", w_umur) + " │ " + clip("Alamat", w_alamat) + " │"
    mid = "├─" + "─"*w_no + "─┼─" + "─"*w_nama + "─┼─" + "─"*w_jk + "─┼─" + "─"*w_umur + "─┼─" + "─"*w_alamat + "─┤"
    bot = "└─" + "─"*w_no + "─┴─" + "─"*w_nama + "─┴─" + "─"*w_jk + "─┴─" + "─"*w_umur + "─┴─" + "─"*w_alamat + "─┘"

    table_lines = [top, hdr, mid]
    for s in samples:
        parts: dict[str, str] = {}
        for seg in s.split(" | "):
            if ":" in seg:
                k, v = seg.split(":", 1)
                parts[k.strip().lower()] = v.strip()
        no = clip(parts.get("no", "-"), w_no)
        nama = clip(parts.get("nama", "-"), w_nama)
        jk = clip(parts.get("jk", "-"), w_jk)
        umur = clip(parts.get("umur", "-"), w_umur)
        alamat = clip(parts.get("alamat", "-"), w_alamat)
        table_lines.append(f"│ {no} │ {nama} │ {jk} │ {umur} │ {alamat} │")
    table_lines.append(bot)
    return table_lines

class MainWindow(ctk.CTk):
    def __init__(self) -> None:
        load_bundled_fonts()
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    "danangsatria.votermail.engine"
                )
            except Exception:
                pass

        super().__init__()
        apply_window_icon(self)

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

        self.latest_update_info: dict[str, Any] | None = None
        self.has_shown_startup_update_modal = False
        self.after(1200, self._start_startup_update_check)

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
            def do_exit() -> None:
                self.cancel_event.set()
                try:
                    clean_temp_dir()
                except Exception:
                    pass
                self.destroy()

            ModernNotificationModal(
                parent=self,
                variant="confirm",
                title="Hentikan Proses & Keluar?",
                summary="Proses pembuatan PDF sedang berjalan di latar belakang. Jika Anda keluar sekarang, berkas PDF yang sedang diproses tidak akan disimpan secara utuh.",
                solution="Klik 'Lanjutkan Mencetak' untuk membiarkan proses selesai, atau klik 'Ya, Hentikan Proses' untuk membatalkan dan menutup aplikasi.",
                primary_action_label="Ya, Hentikan Proses",
                primary_action_callback=do_exit,
                secondary_action_label="Lanjutkan Mencetak",
                theme=self.theme,
            )
            return

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

        self.btn_update_header = ctk.CTkButton(
            header_right,
            text="⬆ Update Tersedia",
            command=self._open_update_modal,
            fg_color="#10B981",
            hover_color="#059669",
            text_color="#FFFFFF",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            height=32,
            corner_radius=8,
        )

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

        # In-App Non-Intrusive Toast Banner
        self.toast_banner = InAppToastBanner(self, theme=self.theme)

        # Main Workspace Container: 2-Column Responsive Layout without nested outer scrollbars
        self.main_body = ctk.CTkFrame(self, fg_color=THEME["bg_root"])
        self.main_body.pack(fill="both", expand=True, padx=14, pady=(10, 6))

        # Left Column (approx 46% width): Source Preparation & Kadus Selection
        self.left_col = ctk.CTkFrame(self.main_body, fg_color="transparent", width=590)
        self.left_col.pack(side="left", fill="both", expand=False, padx=(0, 10))
        self.left_col.pack_propagate(False)

        # Right Column (approx 54% width): Output Settings, Mailing Engine, Live Logs & Preview Table
        self.right_col = ctk.CTkFrame(self.main_body, fg_color="transparent")
        self.right_col.pack(side="right", fill="both", expand=True)

        self._build_card_sources(self.left_col)
        self._build_card_kadus(self.left_col)
        self._build_card_exec(self.right_col)
        self._build_card_activity(self.right_col)

    def show_toast(
        self,
        message: str,
        variant: Literal["success", "info", "warning"] = "success",
        duration_ms: int = 3500,
    ) -> None:
        if hasattr(self, "toast_banner"):
            self.toast_banner.show(message, variant=variant, duration_ms=duration_ms, before_widget=self.main_body)

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
            text="Belum ada berkas Excel DPT yang dipindai. Pilih berkas Excel di kartu 1 untuk menampilkan daftar wilayah.",
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

        hdr_right = ctk.CTkFrame(header, fg_color="transparent")
        hdr_right.pack(side="right")

        self.btn_preview_header = ctk.CTkButton(
            hdr_right,
            text="Pratinjau 1 Halaman",
            width=140,
            height=28,
            fg_color=THEME["btn_accent_soft_bg"],
            hover_color=THEME["btn_accent_soft_hover"],
            border_color=THEME["btn_accent_soft_border"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["btn_accent_soft_text"],
            command=self._generate_preview,
        )
        self.btn_preview_header.pack(side="left", padx=(0, 8))

        self.badge_mailing = StatusBadge(hdr_right, "Menunggu Word & Excel", "neutral", self.theme)
        self.badge_mailing.pack(side="left")

        out_row = ctk.CTkFrame(self.card_exec, fg_color="transparent")
        out_row.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkLabel(
            out_row,
            text="Folder Hasil PDF:",
            width=135,
            anchor="w",
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
            width=105,
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
            width=105,
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
            width=135,
            anchor="w",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_secondary"],
        ).pack(side="left", padx=(0, 8))

        self.address_mode_menu = ctk.CTkOptionMenu(
            addr_box,
            values=ADDRESS_FORMAT_OPTIONS,
            width=280,
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
            command=self._on_address_format_changed,
        )
        curr_addr = self.cfg.address_mode if self.cfg.address_mode in ADDRESS_FORMAT_OPTIONS else DEFAULT_ADDRESS_FORMAT
        self.address_mode_menu.set(curr_addr)
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

        filter_row = ctk.CTkFrame(self.card_exec, fg_color="transparent")
        filter_row.pack(fill="x", padx=16, pady=(0, 8))

        ctk.CTkLabel(
            filter_row,
            text="Filter Cetak:",
            width=135,
            anchor="w",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_secondary"],
        ).pack(side="left", padx=(0, 8))

        self.filter_mode_menu = ctk.CTkOptionMenu(
            filter_row,
            values=FILTER_MODE_OPTIONS,
            width=230,
            height=32,
            fg_color=THEME["btn_secondary_bg"],
            button_color=THEME["btn_secondary_hover"],
            button_hover_color=THEME["btn_secondary_hover"],
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_primary"],
            dropdown_fg_color=THEME["bg_card"],
            dropdown_text_color=THEME["text_primary"],
            dropdown_hover_color=THEME["bg_hover"],
            command=self._on_filter_mode_changed,
        )
        self.filter_mode_menu.set(DEFAULT_FILTER_MODE)
        self.filter_mode_menu.pack(side="left", padx=(0, 8))

        self.btn_filter_reset = ctk.CTkButton(
            filter_row,
            text="✕",
            width=32,
            height=32,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=12, weight="bold"),
            text_color=THEME["btn_secondary_text"],
            command=self._on_filter_reset,
            state="disabled",
        )
        self.btn_filter_reset.pack(side="right", padx=(8, 0))

        self.filter_query_entry = ctk.CTkEntry(
            filter_row,
            placeholder_text="Semua pemilih dalam Kadus terpilih akan dicetak",
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_primary"],
            height=32,
            state="disabled",
        )
        self.filter_query_entry.pack(side="left", fill="x", expand=True)
        self.filter_query_entry.bind("<KeyRelease>", self._on_filter_key_release)

        self.filter_status_banner = FilterStatusBanner(self.card_exec, theme=self.theme)
        self.filter_status_banner.pack(fill="x", padx=16, pady=(0, 10))
        self.filter_status_banner.set_status(
            "all",
            "STATUS CETAK: Semua Pemilih • Menunggu berkas DPT Excel",
        )

        btn_action_row = ctk.CTkFrame(self.card_exec, fg_color="transparent")
        btn_action_row.pack(fill="x", padx=16, pady=(0, 8))

        self.btn_start = ctk.CTkButton(
            btn_action_row,
            text="Mulai Buat Surat PDF",
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

        self.time_label = ctk.CTkLabel(
            stats_row,
            text="Waktu Berjalan: 00:00",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
            text_color=THEME["text_secondary"],
        )
        self.time_label.pack(side="right")
        self.eta_label = self.time_label

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
            wrap="none",
        )

        self.log_textbox = ctk.CTkTextbox(
            self.activity_content_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_primary"],
            wrap="none",
        )

        self.voter_table.pack(fill="both", expand=True)
        try:
            self.audit_textbox._textbox.configure(spacing1=2, spacing3=2)
            self.log_textbox._textbox.configure(spacing1=2, spacing3=2)
        except Exception:
            pass
        self._apply_log_tag_styles()

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
                err_dict = translate_exception_to_user_message(exc, context_label="Folder Hasil PDF")
                ModernNotificationModal(
                    parent=self,
                    variant="error",
                    title="Gagal Membuka Folder",
                    summary=err_dict["summary"],
                    solution=err_dict["solution"],
                    technical_detail=err_dict["technical_detail"],
                    primary_action_label="Mengerti & Tutup",
                    theme=self.theme,
                )
        else:
            ModernNotificationModal(
                parent=self,
                variant="warning",
                title="Folder Belum Tersedia",
                summary="Folder penyimpanan PDF belum dibuat di komputer Anda.",
                solution="Folder akan dibuat secara otomatis saat proses pembuatan surat PDF dimulai.",
                primary_action_label="Mengerti & Tutup",
                theme=self.theme,
            )

    def _save_active_config(self) -> None:
        self.cfg.word_path = self.doc_row_word.current_path
        self.cfg.excel_path = self.doc_row_excel.current_path
        self.cfg.output_dir = self.output_entry.get().strip()
        self.cfg.address_mode = self.address_mode_menu.get().strip()
        self.cfg.enable_pdf_verification = self.verify_switch.get() == 1
        self.cfg.theme_mode = self.theme.name
        save_app_config(self.cfg)

    def _on_address_format_changed(self, val: str) -> None:
        self.cfg.address_mode = val
        self._save_active_config()
        self.voter_table.update_address_format(val)
        self._on_filter_changed()
        self.show_toast(f"Format alamat diperbarui: '{val}'", variant="success")
        self._log(f"[PENGATURAN] Format alamat diubah: {val}")

    def _on_filter_mode_changed(self, mode: str) -> None:
        if mode == FILTER_MODE_ALL:
            self.filter_query_entry.delete(0, "end")
            self.filter_query_entry.configure(
                state="disabled",
                placeholder_text="Semua pemilih dalam Kadus terpilih akan dicetak",
            )
            self.btn_filter_reset.configure(state="disabled")
        else:
            self.filter_query_entry.configure(state="normal")
            if mode == FILTER_MODE_NUMBER:
                self.filter_query_entry.configure(placeholder_text="Ketik nomor/rentang, misal: 15, 42, 100-115")
            elif mode == FILTER_MODE_RT:
                self.filter_query_entry.configure(placeholder_text="Ketik nomor RT, misal: 003 atau 1, 2")
            elif mode == FILTER_MODE_NAME:
                self.filter_query_entry.configure(placeholder_text="Ketik nama pemilih, misal: Ade Setiawan, Adnan")
            self.btn_filter_reset.configure(state="normal")
            self.filter_query_entry.focus_set()

        self._on_filter_changed()

    def _on_filter_reset(self) -> None:
        self.filter_mode_menu.set(FILTER_MODE_ALL)
        self.filter_query_entry.delete(0, "end")
        self.filter_query_entry.configure(
            state="disabled",
            placeholder_text="Semua pemilih dalam Kadus terpilih akan dicetak",
        )
        self.btn_filter_reset.configure(state="disabled")
        self._on_filter_changed()
        self.show_toast("Filter pemilih di-reset ke Semua Pemilih", variant="info")

    def _on_filter_key_release(self, event: Any = None) -> None:
        if hasattr(self, "_filter_debounce_id") and self._filter_debounce_id:
            try:
                self.after_cancel(self._filter_debounce_id)
            except Exception:
                pass
        self._filter_debounce_id = self.after(150, self._on_filter_changed)

    def _on_filter_changed(self) -> None:
        if not self.excel_scan_result:
            if hasattr(self, "filter_status_banner"):
                self.filter_status_banner.set_status("all", "STATUS CETAK: Semua Pemilih • Menunggu berkas DPT Excel")
            return

        filter_mode = self.filter_mode_menu.get().strip()
        filter_query = self.filter_query_entry.get().strip()
        is_mode_active = (filter_mode != FILTER_MODE_ALL)
        is_query_filled = bool(filter_query)
        is_filtered = (is_mode_active and is_query_filled)

        layout = self.word_check_result.layout_mode if self.word_check_result else "2-UP"

        total_matched_all = 0
        total_matched_pages = 0
        active_kadus_count = 0
        preview_voters: list[VoterData] = []

        total_unfiltered = sum(
            s.total_voters
            for s in self.excel_scan_result.sheet_summaries
            if s.is_voter_sheet and (s.sheet_name in self.kadus_rows and self.kadus_rows[s.sheet_name].checkbox_var.get())
        )

        for s in self.excel_scan_result.sheet_summaries:
            if not s.is_voter_sheet:
                continue

            if is_filtered:
                matched_sheet_voters = apply_voter_filter(s.voters, filter_mode, filter_query)
            else:
                matched_sheet_voters = s.voters

            m_cnt = len(matched_sheet_voters)
            m_pages = m_cnt if layout == "1-UP" else math.ceil(m_cnt / 2)

            if s.sheet_name in self.kadus_rows:
                self.kadus_rows[s.sheet_name].update_filter_stats(m_cnt, m_pages, is_filtered)

            is_checked = (s.sheet_name in self.kadus_rows and self.kadus_rows[s.sheet_name].checkbox_var.get())
            if is_checked:
                if is_filtered:
                    total_matched_all += m_cnt
                    total_matched_pages += m_pages
                    if m_cnt > 0:
                        active_kadus_count += 1
                    if len(preview_voters) < 25:
                        preview_voters.extend(matched_sheet_voters[:25 - len(preview_voters)])
                else:
                    total_matched_all += s.total_voters
                    total_matched_pages += (s.total_voters if layout == "1-UP" else s.expected_pages)
                    active_kadus_count += 1

        selected_total_checked = sum(1 for row in self.kadus_rows.values() if row.checkbox_var.get())

        self.metric_bar.update_metrics(active_kadus_count, total_matched_all, total_matched_pages)

        curr_addr_mode = self.address_mode_menu.get().strip()
        if is_filtered:
            self.voter_table.set_voters(
                preview_voters,
                address_format=curr_addr_mode,
                is_filtered=True,
                total_count=total_unfiltered,
                matched_count=total_matched_all,
                empty_message="Tidak ada pemilih yang cocok dengan filter tersebut. Periksa kembali nomor/RT/nama yang diketik.",
            )
        else:
            sample_voters: list[VoterData] = []
            for s in self.excel_scan_result.sheet_summaries:
                if s.voters and (s.sheet_name in self.kadus_rows and self.kadus_rows[s.sheet_name].checkbox_var.get()):
                    sample_voters.extend(s.voters)
                    if len(sample_voters) >= 5:
                        break
            if not sample_voters:
                for s in self.excel_scan_result.sheet_summaries:
                    if s.voters:
                        sample_voters.extend(s.voters)
                        if len(sample_voters) >= 5:
                            break
            self.voter_table.set_voters(
                sample_voters[:5],
                address_format=curr_addr_mode,
                is_filtered=False,
                total_count=self.excel_scan_result.total_voters if self.excel_scan_result else 0,
            )

        # Update FilterStatusBanner states
        if hasattr(self, "filter_status_banner"):
            if not is_mode_active:
                self.btn_filter_reset.configure(state="disabled")
                self.filter_status_banner.set_status(
                    "all",
                    f"STATUS CETAK: Semua Pemilih • Total {total_matched_all:,} Pemilih ({active_kadus_count} Kadus)".replace(",", "."),
                )
            elif not is_query_filled:
                self.btn_filter_reset.configure(state="normal")
                self.filter_status_banner.set_status(
                    "empty",
                    "FILTER BELUM DIISI: Ketik kriteria filter pada kolom di atas untuk menyaring pemilih",
                )
            elif is_filtered and total_matched_all > 0:
                self.btn_filter_reset.configure(state="normal")
                self.filter_status_banner.set_status(
                    "active",
                    f"FILTER DITERAPKAN: [{filter_mode}: {filter_query}] • Menyaring {total_matched_all:,} dari {total_unfiltered:,} Pemilih ({active_kadus_count} Kadus)".replace(",", "."),
                )
            else:
                self.btn_filter_reset.configure(state="normal")
                self.filter_status_banner.set_status(
                    "no_match",
                    "TIDAK ADA DATA COCOK: 0 pemilih ditemukan untuk filter tersebut. Periksa kembali ketikan Anda",
                )

        if self.is_processing:
            return

        if is_filtered:
            if selected_total_checked == 0:
                self.btn_start.configure(
                    state="disabled",
                    text="Pilih Minimal 1 Wilayah Kadus",
                    fg_color=THEME["btn_disabled_bg"],
                    text_color=THEME["btn_disabled_text"],
                )
                self.status_detail_label.configure(
                    text="Pilih minimal 1 wilayah Kadus yang ingin dicetak."
                )
            elif total_matched_all > 0:
                self.btn_start.configure(
                    state="normal",
                    text=f"Cetak PDF Terfilter ({total_matched_all:,} Pemilih)".replace(",", "."),
                    fg_color=THEME["btn_primary_bg"],
                    hover_color=THEME["btn_primary_hover"],
                    text_color=THEME["btn_primary_text"],
                )
                self.status_detail_label.configure(
                    text=f"Filter Aktif: {total_matched_all:,} Pemilih Cocok ({active_kadus_count} Kadus)".replace(",", ".")
                )
            else:
                self.btn_start.configure(
                    state="disabled",
                    text="Tidak Ada Pemilih Cocok",
                    fg_color=THEME["btn_disabled_bg"],
                    text_color=THEME["btn_disabled_text"],
                )
                self.status_detail_label.configure(
                    text="Tidak ada pemilih yang cocok dengan filter tersebut. Periksa kembali nomor/RT/nama yang diketik."
                )
        else:
            self._check_ready_state()
            self.status_detail_label.configure(
                text=f"Siap Memproses: 0 / {total_matched_all:,} Pemilih ({active_kadus_count} Kadus)".replace(",", ".")
            )

    def _on_theme_toggled(self, val: str) -> None:
        mode = "Dark" if "gelap" in val.lower() or val.lower() == "dark" else "Light"
        ctk.set_appearance_mode(mode)
        self.theme = get_theme(mode)
        self.cfg.theme_mode = mode
        save_app_config(self.cfg)
        self._check_ready_state()
        self._apply_log_tag_styles()
        if hasattr(self, "filter_status_banner"):
            self.filter_status_banner.update_theme(self.theme)
        if hasattr(self, "voter_table"):
            self.voter_table.update_theme(self.theme)
        if hasattr(self, "btn_preview_header"):
            self.btn_preview_header.configure(
                fg_color=THEME["btn_accent_soft_bg"],
                hover_color=THEME["btn_accent_soft_hover"],
                border_color=THEME["btn_accent_soft_border"],
                text_color=THEME["btn_accent_soft_text"],
            )
        self._log(f"[TEMA] Mengganti tema visual ke: {mode}")

    def _apply_log_tag_styles(self) -> None:
        if not hasattr(self, "log_textbox") or not hasattr(self, "audit_textbox"):
            return
        is_dark = (ctk.get_appearance_mode().lower() == "dark")
        idx = 1 if is_dark else 0

        ts_color = THEME["text_muted"][idx]
        div_color = THEME["border_card"][idx]
        msg_color = THEME["text_primary"][idx]
        msg_dim = THEME["text_secondary"][idx]

        cat_info = THEME["status_info_text"][idx]
        cat_success = THEME["status_ready_text"][idx]
        cat_warn = THEME["status_warn_text"][idx]
        cat_error = THEME["status_error_text"][idx]
        cat_accent = THEME["btn_accent_soft_text"][idx]

        for tb in (self.log_textbox, self.audit_textbox):
            raw = tb._textbox
            raw.tag_config("ts", foreground=ts_color)
            raw.tag_config("div", foreground=div_color)
            raw.tag_config("msg", foreground=msg_color)
            raw.tag_config("msg_dim", foreground=msg_dim)

            raw.tag_config("cat_info", foreground=cat_info)
            raw.tag_config("cat_success", foreground=cat_success)
            raw.tag_config("cat_warning", foreground=cat_warn)
            raw.tag_config("cat_error", foreground=cat_error)
            raw.tag_config("cat_process", foreground=cat_accent)
            raw.tag_config("cat_config", foreground=cat_info)
            raw.tag_config("cat_theme", foreground=cat_accent)
            raw.tag_config("cat_preview", foreground=cat_accent)
            raw.tag_config("cat_audit", foreground=cat_info)

            raw.tag_config("table_border", foreground=div_color)
            raw.tag_config("table_hdr", foreground=cat_accent)
            raw.tag_config("table_row", foreground=msg_color)

    def _render_single_log_line(self, raw: Any, line: str) -> None:
        if line.startswith(("┌", "├", "└")):
            raw.insert("end", line + "\n", "table_border")
        elif line.startswith("│ No   │"):
            raw.insert("end", line + "\n", "table_hdr")
        elif line.startswith("│ ") and line.endswith(" │") and " │ " in line and not line.startswith(("1", "2", "0")):
            raw.insert("end", line + "\n", "table_row")
        elif " │ " in line:
            parts = line.split(" │ ", 2)
            if len(parts) == 3:
                ts, cat_pad, msg = parts
                cat = cat_pad.strip().upper()
                tag = "cat_info"
                if cat == "SUKSES": tag = "cat_success"
                elif cat == "PERINGATAN": tag = "cat_warning"
                elif cat == "ERROR": tag = "cat_error"
                elif cat in ("MEMPROSES", "PRATINJAU", "TEMA"): tag = "cat_process"

                raw.insert("end", ts, "ts")
                raw.insert("end", " │ ", "div")
                raw.insert("end", f"{cat:<10}", tag)
                raw.insert("end", " │ ", "div")
                raw.insert("end", msg + "\n", "msg")
            else:
                raw.insert("end", line + "\n", "msg")
        elif line.startswith("[") and "]" in line:
            raw.insert("end", line + "\n", "cat_preview")
        else:
            raw.insert("end", line + "\n", "msg")

    def _render_history_to_textbox(self, textbox: ctk.CTkTextbox, history: list[str]) -> None:
        raw = textbox._textbox
        if not history:
            placeholder = "Belum ada laporan audit PDF. Jalankan proses mailing terlebih dahulu." if textbox == self.audit_textbox else "Belum ada riwayat log aktivitas."
            raw.insert("end", placeholder + "\n", "msg_dim")
            return

        for line in history:
            self._render_single_log_line(raw, line)
        textbox.see("end")

    def _log(self, text: str, category: str = "all") -> None:
        timestamp = time.strftime("%H:%M:%S")

        is_table_border = text.startswith(("┌", "├", "└"))
        is_table_hdr = text.startswith("│ No   │")
        is_table_row = text.startswith("│ ") and text.endswith(" │") and " │ " in text

        cat_name = "INFO"
        tag_cat = "cat_info"
        clean_msg = text

        if is_table_border or is_table_hdr or is_table_row:
            entry = text
        elif text.startswith("==="):
            entry = f"[{timestamp}] {text}"
        else:
            if text.startswith("[") and "]" in text:
                bracket_tag = text[1:text.index("]")].strip().upper()
                clean_msg = text[text.index("]") + 1:].strip()
                if any(k in bracket_tag for k in ("SELESAI", "SUKSES", "BERHASIL", "READY", "VALID")):
                    cat_name = "SUKSES"
                    tag_cat = "cat_success"
                elif any(k in bracket_tag for k in ("PERINGATAN", "WARN", "PERHATIAN", "ANOMALI")):
                    cat_name = "PERINGATAN"
                    tag_cat = "cat_warning"
                elif any(k in bracket_tag for k in ("ERROR", "GAGAL", "ERR")):
                    cat_name = "ERROR"
                    tag_cat = "cat_error"
                elif "PROSES" in bracket_tag or "CETAK" in bracket_tag:
                    cat_name = "MEMPROSES"
                    tag_cat = "cat_process"
                elif "PENGATURAN" in bracket_tag or "CONFIG" in bracket_tag:
                    cat_name = "PENGATURAN"
                    tag_cat = "cat_config"
                elif "TEMA" in bracket_tag:
                    cat_name = "TEMA"
                    tag_cat = "cat_theme"
                elif "PREVIEW" in bracket_tag:
                    cat_name = "PRATINJAU"
                    tag_cat = "cat_preview"
                elif "AUDIT" in bracket_tag:
                    cat_name = "AUDIT"
                    tag_cat = "cat_audit"
                else:
                    cat_name = bracket_tag[:10]
                    tag_cat = "cat_info"
            elif any(k in text.lower() for k in ("error", "gagal", "terjadi kesalahan")):
                cat_name = "ERROR"
                tag_cat = "cat_error"
            elif any(k in text.lower() for k in ("sukses", "berhasil", "selesai dibuat", "100%")):
                cat_name = "SUKSES"
                tag_cat = "cat_success"
            elif any(k in text.lower() for k in ("perhatian", "peringatan", "anomali")):
                cat_name = "PERINGATAN"
                tag_cat = "cat_warning"
            elif any(k in text.lower() for k in ("memproses", "menggabungkan", "sedang")):
                cat_name = "MEMPROSES"
                tag_cat = "cat_process"

            cat_padded = f"{cat_name:<10}"
            entry = f"{timestamp} │ {cat_padded} │ {clean_msg}"

        self.log_history_all.append(entry)
        if category == "preview":
            self.log_history_preview.append(entry)
        elif category == "audit":
            self.log_history_audit.append(entry)

        def _insert_styled(raw_tb: Any) -> None:
            if is_table_border:
                raw_tb.insert("end", text + "\n", "table_border")
            elif is_table_hdr:
                raw_tb.insert("end", text + "\n", "table_hdr")
            elif is_table_row:
                raw_tb.insert("end", text + "\n", "table_row")
            elif text.startswith("==="):
                raw_tb.insert("end", f"[{timestamp}] ", "ts")
                raw_tb.insert("end", text + "\n", "cat_preview")
            else:
                raw_tb.insert("end", timestamp, "ts")
                raw_tb.insert("end", " │ ", "div")
                raw_tb.insert("end", f"{cat_name:<10}", tag_cat)
                raw_tb.insert("end", " │ ", "div")
                raw_tb.insert("end", clean_msg + "\n", "msg")

        if self.current_log_tab == "all":
            _insert_styled(self.log_textbox._textbox)
            self.log_textbox.see("end")
        elif self.current_log_tab == "audit" and category == "audit":
            _insert_styled(self.audit_textbox._textbox)
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

        if hasattr(self, "btn_clear_log"):
            if val == "Log Sistem":
                self.btn_clear_log.pack(side="left")
            else:
                self.btn_clear_log.pack_forget()

        if val == "Pratinjau Pemilih":
            self.current_log_tab = "preview"
            self.voter_table.pack(fill="both", expand=True)
        elif val == "Hasil Cek PDF":
            self.current_log_tab = "audit"
            self.audit_textbox.pack(fill="both", expand=True)
            self.audit_textbox.delete("1.0", "end")
            self._render_history_to_textbox(self.audit_textbox, self.log_history_audit)
        else:
            self.current_log_tab = "all"
            self.log_textbox.pack(fill="both", expand=True)
            self.log_textbox.delete("1.0", "end")
            self._render_history_to_textbox(self.log_textbox, self.log_history_all)

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
            status_txt = res.format_status_text if res.format_status_text else f"• Mode {mode_desc} Siap Mailing"
            self.doc_row_word.set_status(status_txt, "ready")
            self._log(f"[WORD SIAP] {res.status_badge}: {shorten_path(res.prepared_docx_path)}")
        else:
            self.word_ready = False
            self.doc_row_word.set_status(f"Perlu Perbaikan: {res.message}", "error")
            self._log(f"[WORD ERROR] {res.message}")

        self._update_selected_metrics()
        self._check_ready_state()

    def _open_column_mapper(self) -> None:
        if not self.excel_scan_result or not self.excel_scan_result.available_columns:
            ModernNotificationModal(
                parent=self,
                variant="warning",
                title="Pilih Berkas Excel DPT",
                summary="Aplikasi belum memindai daftar kolom dari berkas Excel DPT.",
                solution="Pilih berkas Excel DPT yang valid pada kartu '1. Sumber Dokumen' terlebih dahulu.",
                primary_action_label="Mengerti & Tutup",
                theme=self.theme,
            )
            return

        def on_applied(new_mapping: dict[str, int]) -> None:
            self.custom_column_map = new_mapping
            self._log(f"[KOLOM DPT] Pemetaan kolom manual diterapkan: {new_mapping}")
            self.show_toast("✓ Pemetaan kolom manual berhasil disimpan dan diterapkan.", variant="success")
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

        addr_mode = self.address_mode_menu.get().strip()
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
        self.voter_table.set_voters(
            sample_voters[:5],
            address_format=self.address_mode_menu.get().strip(),
            total_count=res.total_voters,
        )

        self.excel_ready = True
        if res.is_valid:
            status_txt = res.format_status_text if res.format_status_text else f"✓ Siap Mailing • {res.total_sheets} Kadus ({res.total_voters:,} pemilih)"
            if res.is_standard_dpt:
                status_bar_txt = f"✓ Sesuai Format Standar DPT ({res.total_sheets} Sheet Kadus Seragam • {res.total_voters:,} Pemilih). Zero Error."
            else:
                status_bar_txt = f"✓ Format Kustom Terpetakan ({res.total_sheets} Wilayah • {res.total_voters:,} Pemilih). Siap Mailing."
            self.doc_row_excel.set_status(status_txt, "ready")
            self.excel_status_bar.configure(
                text=status_bar_txt,
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
            for table_line in format_voter_preview_table(res.preview_samples[:5]):
                self._log(table_line, category="preview")

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
        self._on_filter_changed()

    def _check_ready_state(self) -> None:
        selected_count = sum(1 for row in self.kadus_rows.values() if row.checkbox_var.get())

        if self.is_processing:
            return

        if self.word_ready and self.excel_ready and selected_count > 0:
            self.btn_start.configure(
                state="normal",
                text="Mulai Buat Surat PDF",
                fg_color=THEME["btn_primary_bg"],
                hover_color=THEME["btn_primary_hover"],
                text_color=THEME["btn_primary_text"],
            )
            self.badge_mailing.set_status("Siap Mailing", "ready")
        elif selected_count == 0 and self.word_ready and self.excel_ready:
            self.btn_start.configure(
                state="disabled",
                text="Pilih Minimal 1 Wilayah Kadus",
                fg_color=THEME["btn_disabled_bg"],
                text_color=THEME["btn_disabled_text"],
            )
            self.badge_mailing.set_status("Pilih Kadus", "neutral")
        else:
            self.btn_start.configure(
                state="disabled",
                text="Pilih Dokumen Word & Excel Dahulu",
                fg_color=THEME["btn_disabled_bg"],
                text_color=THEME["btn_disabled_text"],
            )
            self.badge_mailing.set_status("Menunggu Word & Excel", "neutral")

    def _generate_preview(self) -> None:
        if not self.word_check_result or not self.word_check_result.is_valid:
            ModernNotificationModal(
                parent=self,
                variant="warning",
                title="Template Word Belum Siap",
                summary="Template Word belum diperiksa atau strukturnya belum valid.",
                solution="Tekan tombol 'Cek Ulang' di kartu '1. Sumber Dokumen' terlebih dahulu sebelum membuat pratinjau.",
                primary_action_label="Mengerti & Tutup",
                theme=self.theme,
            )
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
        curr_addr_fmt = self.address_mode_menu.get().strip()
        ok, msg, final_pdf = generate_preview_via_word(
            template_doc, sample_voters, preview_pdf, layout_mode=layout, address_format=curr_addr_fmt
        )

        if ok:
            self._log(f"[PREVIEW SELESAI] Pratinjau tersimpan di: {shorten_path(final_pdf)}")
            try:
                os.startfile(final_pdf)
            except Exception:
                pass
            self.show_toast("✓ Berkas pratinjau 1 lembar berhasil dibuat dan dibuka.", variant="success")
        else:
            err_dict = translate_exception_to_user_message(Exception(msg), context_label="Pratinjau Surat")
            ModernNotificationModal(
                parent=self,
                variant="error",
                title="Gagal Membuat Pratinjau",
                summary=err_dict["summary"],
                solution=err_dict["solution"],
                technical_detail=err_dict["technical_detail"],
                primary_action_label="Mengerti & Tutup",
                theme=self.theme,
            )

    def _start_mailing_process(self) -> None:
        if self.is_processing or not self.word_check_result or not self.excel_scan_result:
            return

        selected_sheets = [
            s for s in self.excel_scan_result.sheet_summaries
            if s.sheet_name in self.kadus_rows and self.kadus_rows[s.sheet_name].checkbox_var.get()
        ]
        if not selected_sheets:
            ModernNotificationModal(
                parent=self,
                variant="warning",
                title="Belum Ada Wilayah Dipilih",
                summary="Tidak ada Kadus / RW yang dicentang untuk dicetak.",
                solution="Centang minimal satu wilayah pada kartu '2. Pilih Wilayah (Kadus/RW)' sebelum memulai pencetakan.",
                primary_action_label="Pilih Semua Wilayah",
                primary_action_callback=self._select_all_sheets,
                secondary_action_label="Pilih Nanti",
                theme=self.theme,
            )
            return

        filter_mode = self.filter_mode_menu.get().strip()
        filter_query = self.filter_query_entry.get().strip()
        is_filtered = (filter_mode != FILTER_MODE_ALL and bool(filter_query))

        sheets_data: list[tuple[SheetScanSummary, list[VoterData]]] = []
        for s in selected_sheets:
            if is_filtered:
                matched_v = apply_voter_filter(s.voters, filter_mode, filter_query)
                if not matched_v:
                    self._log(f"[INFO] {s.sheet_name} dilewati karena tidak ada pemilih yang cocok dengan filter.")
                    continue
                sheets_data.append((s, matched_v))
            else:
                sheets_data.append((s, s.voters))

        if not sheets_data:
            ModernNotificationModal(
                parent=self,
                variant="warning",
                title="Tidak Ada Pemilih Ditemukan",
                summary="Kriteria filter cetak yang dimasukkan tidak cocok dengan satupun data pemilih pada wilayah yang dipilih.",
                solution="Periksa kembali nomor urut, nomor RT, atau ejaan nama pemilih, atau klik 'Reset Filter' untuk mencetak seluruh pemilih.",
                primary_action_label="Reset Filter",
                primary_action_callback=self._on_filter_reset,
                secondary_action_label="Tutup",
                theme=self.theme,
            )
            return

        output_dir = self.output_entry.get().strip() or str(OUTPUT_DIR.resolve())
        os.makedirs(output_dir, exist_ok=True)

        self.is_processing = True
        self.cancel_event.clear()
        self.start_process_time = time.monotonic()
        self.time_label.configure(text="Waktu Berjalan: 00:00")

        self.btn_start.configure(
            state="disabled",
            text="Sedang Memproses Surat PDF...",
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

        total_voters_all = sum(len(v_list) for _, v_list in sheets_data)
        template_doc = self.word_check_result.prepared_docx_path
        verify_pdf_on = self.verify_switch.get() == 1
        layout_mode = self.word_check_result.layout_mode if self.word_check_result else "2-UP"
        active_address_format = self.address_mode_menu.get().strip()
        filter_suffix = build_filter_filename_suffix(filter_mode, filter_query) if is_filtered else ""

        filter_info = f" | Filter: {filter_mode} ({filter_query})" if is_filtered else ""
        self._log(f"[MAILING MULAI] Memulai proses mailing {len(sheets_data)} Kadus ({total_voters_all:,} pemilih, Mode: {layout_mode}{filter_info})...")

        def worker() -> None:
            processed_voters_acc = 0
            audit_results: list[VerificationResult] = []

            for s, voters_list in sheets_data:
                if self.cancel_event.is_set():
                    self.worker_queue.put(("process_cancelled", "Dibatalkan oleh pengguna."))
                    return

                self.worker_queue.put(("kadus_started", s.sheet_name))
                safe_name = sanitize_filename(f"SURAT_PEMBERITAHUAN_{s.sheet_name}{filter_suffix}")
                out_pdf = os.path.join(output_dir, f"{safe_name}.pdf")

                def prog_cb(cur: int, tot: int, sheet: str) -> None:
                    overall = processed_voters_acc + cur
                    self.worker_queue.put(("progress_update", (overall, total_voters_all, sheet)))

                def phase_cb(phase_text: str, pct: float) -> None:
                    self.worker_queue.put(("phase_update", (phase_text, pct)))

                ok, pages, msg = merge_voters_to_pdf_via_word_com(
                    template_docx_path=template_doc,
                    voters=voters_list,
                    sheet_name=s.sheet_name,
                    wilayah_name=s.wilayah_text,
                    output_pdf_path=out_pdf,
                    progress_callback=prog_cb,
                    phase_callback=phase_cb,
                    cancel_event=self.cancel_event,
                    layout_mode=layout_mode,
                    address_format=active_address_format,
                )

                if not ok:
                    self.worker_queue.put(("kadus_failed", (s.sheet_name, msg)))
                    return

                processed_voters_acc += len(voters_list)

                if verify_pdf_on and not self.cancel_event.is_set():
                    self.worker_queue.put(("verify_started", s.sheet_name))
                    ver = verify_kadus_pdf(out_pdf, s.sheet_name, voters_list, layout_mode=layout_mode)
                    audit_results.append(ver)
                    self.worker_queue.put(("verify_done", (ver, out_pdf)))
                else:
                    self.worker_queue.put(("kadus_completed", (s.sheet_name, out_pdf)))

            if audit_results and verify_pdf_on:
                audit_file = export_audit_excel(audit_results, output_dir)
                self.worker_queue.put(("audit_exported", audit_file))

            self.worker_queue.put(("all_completed", (len(sheets_data), total_voters_all)))

        threading.Thread(target=worker, daemon=True).start()

    def _cancel_mailing(self) -> None:
        if self.is_processing:
            def do_cancel() -> None:
                self.cancel_event.set()
                self._log("[BATAL] Mengirim sinyal pembatalan ke proses latar belakang...")
                self.btn_cancel.configure(state="disabled", text="Membatalkan...")

            ModernNotificationModal(
                parent=self,
                variant="confirm",
                title="Hentikan Proses Pembuatan PDF?",
                summary="Proses pembuatan surat panggilan sedang berjalan. Jika dibatalkan sekarang, berkas PDF yang sedang diproses tidak akan disimpan secara utuh.",
                solution="Klik 'Lanjutkan Mencetak' untuk membiarkan proses selesai, atau klik 'Ya, Hentikan Proses' untuk menghentikan pembuatan dokumen.",
                primary_action_label="Ya, Hentikan Proses",
                primary_action_callback=do_cancel,
                secondary_action_label="Lanjutkan Mencetak",
                theme=self.theme,
            )

    def _poll_worker_queue(self) -> None:
        if self.is_processing and self.start_process_time > 0:
            elapsed = time.monotonic() - self.start_process_time
            self.time_label.configure(text=f"Waktu Berjalan: {format_elapsed_time(elapsed)}")

        try:
            while True:
                msg_type, payload = self.worker_queue.get_nowait()

                if msg_type == "word_done":
                    self._apply_word_results(payload)
                elif msg_type == "word_error":
                    self.word_ready = False
                    self.doc_row_word.set_status(f"Gagal: {payload}", "error")
                    self._log(f"[WORD ERROR] {payload}")
                    err_dict = translate_exception_to_user_message(Exception(payload), context_label="Template Word")
                    ModernNotificationModal(
                        parent=self,
                        variant="error",
                        title=err_dict["title"],
                        summary=err_dict["summary"],
                        solution=err_dict["solution"],
                        technical_detail=err_dict["technical_detail"],
                        primary_action_label="Mengerti & Tutup",
                        theme=self.theme,
                    )
                elif msg_type == "excel_done":
                    self._apply_excel_results(payload)
                elif msg_type == "excel_error":
                    self.excel_ready = False
                    self.doc_row_excel.set_status(f"Gagal: {payload}", "error")
                    self._log(f"[EXCEL ERROR] {payload}")
                    err_dict = translate_exception_to_user_message(Exception(payload), context_label="Excel DPT")
                    ModernNotificationModal(
                        parent=self,
                        variant="error",
                        title=err_dict["title"],
                        summary=err_dict["summary"],
                        solution=err_dict["solution"],
                        technical_detail=err_dict["technical_detail"],
                        primary_action_label="Mengerti & Tutup",
                        theme=self.theme,
                    )

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

                elif msg_type == "phase_update":
                    phase_text, _ = payload
                    self.badge_mailing.set_status("Memproses...", "process")
                    self.status_detail_label.configure(text=phase_text)

                elif msg_type == "kadus_completed":
                    sheet_name, out_pdf = payload
                    if sheet_name in self.kadus_rows:
                        self.kadus_rows[sheet_name].set_completed(out_pdf)
                    self._log(f"[SUKSES] PDF {sheet_name} selesai dibuat: {os.path.basename(out_pdf)}")

                elif msg_type == "kadus_failed":
                    sheet_name, err = payload
                    elapsed = time.monotonic() - self.start_process_time if self.start_process_time > 0 else 0.0
                    if sheet_name in self.kadus_rows:
                        self.kadus_rows[sheet_name].set_error("• Gagal")
                    self._on_process_finished()
                    self.time_label.configure(text=f"Waktu Berjalan: {format_elapsed_time(elapsed)} (Dihentikan)")
                    self._log(f"[ERROR] Gagal memproses {sheet_name}: {err}")
                    err_dict = translate_exception_to_user_message(Exception(err), context_label=f"Wilayah {sheet_name}")
                    ModernNotificationModal(
                        parent=self,
                        variant="error",
                        title=err_dict["title"],
                        summary=err_dict["summary"],
                        solution=err_dict["solution"],
                        technical_detail=err_dict["technical_detail"],
                        primary_action_label="Mengerti & Tutup",
                        theme=self.theme,
                    )

                elif msg_type == "verify_started":
                    self._log(f"[AUDIT] Memverifikasi integritas PDF: {shorten_path(payload)}...")

                elif msg_type == "verify_done":
                    ver, out_pdf = payload
                    if ver.sheet_name in self.kadus_rows:
                        if ver.is_valid:
                            self.kadus_rows[ver.sheet_name].set_completed(out_pdf)
                        else:
                            self.kadus_rows[ver.sheet_name].set_error("• Anomali")
                    self._log(ver.audit_summary, category="audit")

                elif msg_type == "audit_exported":
                    self._log(f"[AUDIT SELESAI] Laporan audit tersimpan di: {shorten_path(payload)}", category="audit")

                elif msg_type == "process_cancelled":
                    elapsed = time.monotonic() - self.start_process_time if self.start_process_time > 0 else 0.0
                    self._on_process_finished()
                    self.time_label.configure(text=f"Waktu Berjalan: {format_elapsed_time(elapsed)} (Dihentikan)")
                    self._log(f"[BATAL] {payload}")
                    ModernNotificationModal(
                        parent=self,
                        variant="warning",
                        title="Proses Dibatalkan",
                        summary="Proses pembuatan surat panggilan telah dihentikan atas permintaan pengguna. Berkas sementara telah dibersihkan secara aman.",
                        solution="Anda dapat memulai proses baru kapan saja dengan menekan tombol 'Mulai Buat Surat PDF'.",
                        primary_action_label="Mengerti & Tutup",
                        theme=self.theme,
                    )

                elif msg_type == "all_completed":
                    kadus_count, total_v = payload
                    elapsed = time.monotonic() - self.start_process_time if self.start_process_time > 0 else 0.0
                    self._on_process_finished()
                    self.progress_bar.set(1.0)
                    self.time_label.configure(text=f"Waktu Berjalan: {format_elapsed_time(elapsed)} (Selesai)")
                    self.status_detail_label.configure(
                        text=f"Selesai 100%: {total_v:,} Pemilih ({kadus_count} Kadus)".replace(",", ".")
                    )
                    self._log(f"[SELESAI 100%] Berhasil membuat {kadus_count} file PDF ({total_v:,} pemilih) dalam {format_elapsed_time(elapsed)}.")
                    clean_temp_dir()

                    layout_mode = self.word_check_result.layout_mode if self.word_check_result else "2-UP"
                    has_two_slots = (layout_mode == "2-UP")
                    total_lembar = math.ceil(total_v / 2) if has_two_slots else total_v

                    has_anomaly = any(
                        "anomali" in row.status_label.cget("text").lower()
                        for row in self.kadus_rows.values()
                    )

                    if has_anomaly:
                        ModernNotificationModal(
                            parent=self,
                            variant="warning",
                            title="Surat Selesai dengan Catatan Verifikasi",
                            summary=f"Sebanyak {kadus_count} file PDF berhasil dibuat, namun ditemukan selisih halaman atau anomali pada hasil pemeriksaan PDF.",
                            solution="Periksa tab 'Hasil Cek PDF' untuk melihat rincian ketidaksesuaian jumlah halaman sebelum mencetak surat fisik.",
                            info_box={
                                "Wilayah Selesai": f"{kadus_count} Kadus",
                                "Total Tercetak": f"{total_v:,} Surat ({total_lembar:,} Lembar F4)".replace(",", "."),
                                "Waktu Berjalan": format_elapsed_time(elapsed),
                                "Status Cek PDF": "Perlu Pemeriksaan Ulang (Ada Selisih)",
                            },
                            primary_action_label="Lihat Hasil Cek PDF",
                            primary_action_callback=lambda: self.tabview.set("Hasil Cek PDF"),
                            secondary_action_label="Buka Folder PDF",
                            secondary_action_callback=self._open_output_dir,
                            theme=self.theme,
                        )
                    else:
                        ModernNotificationModal(
                            parent=self,
                            variant="success",
                            title="Pembuatan Surat Selesai",
                            summary=f"Seluruh berkas surat panggilan ({kadus_count} Kadus, {total_v:,} pemilih) berhasil dibuat dan diverifikasi siap cetak.".replace(",", "."),
                            info_box={
                                "Wilayah Selesai": f"{kadus_count} Kadus",
                                "Total Tercetak": f"{total_v:,} Surat ({total_lembar:,} Lembar F4)".replace(",", "."),
                                "Waktu Berjalan": format_elapsed_time(elapsed),
                                "Status Cek PDF": "Lolos Verifikasi (100% Sesuai)",
                            },
                            primary_action_label="Buka Folder Hasil PDF",
                            primary_action_callback=self._open_output_dir,
                            secondary_action_label="Selesai",
                            theme=self.theme,
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

    def _start_startup_update_check(self) -> None:
        def check_worker() -> None:
            from app.core.updater import check_for_updates
            info = check_for_updates(__version__, timeout=5.0)
            self.after(0, self._on_update_check_finished, info)

        threading.Thread(target=check_worker, daemon=True).start()

    def _on_update_check_finished(self, info: dict[str, Any]) -> None:
        if info.get("update_available"):
            self.latest_update_info = info
            latest_v = info.get("latest_version", "")
            self.btn_update_header.configure(text=f"⬆ Update v{latest_v}")
            try:
                self.btn_update_header.pack(side="left", padx=(0, 10), before=self.btn_about)
            except Exception:
                self.btn_update_header.pack(side="left", padx=(0, 10))

            if not self.has_shown_startup_update_modal:
                self.has_shown_startup_update_modal = True
                UpdateNotificationModal(self, self.theme, info)

    def _open_update_modal(self) -> None:
        if self.latest_update_info:
            UpdateNotificationModal(self, self.theme, self.latest_update_info)
        else:
            self._start_startup_update_check()
