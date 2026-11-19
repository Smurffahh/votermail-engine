# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import importlib.metadata
import os
from pathlib import Path
import re
import sys
import threading
from typing import Any, Callable, Literal
import webbrowser
import customtkinter as ctk

from app.config.settings import APP_ICON_ICO, APP_ICON_PNG, DEFAULT_ADDRESS_FORMAT
from app.config.theme import THEME, ThemeTokens
from app.config.version import (
    __app_name__,
    __author__,
    __github_user__,
    __repository__,
    __license__,
    __version__,
    FEATURE_REGISTRY,
)
from app.core.excel_scanner import VoterData, is_blocked_pii, build_voter_address

def apply_window_icon(window: Any) -> None:
    if not APP_ICON_ICO.exists():
        return

    icon_str = str(APP_ICON_ICO.resolve())

    def _set_icon() -> None:
        try:
            if hasattr(window, "winfo_exists") and window.winfo_exists():
                window.iconbitmap(icon_str)
        except Exception:
            pass

    _set_icon()

    if hasattr(window, "after"):
        try:
            window.after(220, _set_icon)
            window.after(300, _set_icon)
        except Exception:
            pass

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

class ModernNotificationModal(ctk.CTkToplevel):
    def __init__(
        self,
        parent: Any,
        variant: Literal["success", "warning", "error", "confirm"] = "success",
        title: str = "",
        summary: str = "",
        solution: str = "",
        info_box: dict[str, str] | str | None = None,
        technical_detail: str = "",
        primary_action_label: str = "",
        primary_action_callback: Callable[[], None] | None = None,
        secondary_action_label: str = "",
        secondary_action_callback: Callable[[], None] | None = None,
        theme: ThemeTokens | None = None,
    ) -> None:
        super().__init__(parent)
        self.parent = parent
        self.variant = variant
        self.summary_text = summary
        self.solution_text = solution
        self.info_box = info_box
        self.technical_detail = technical_detail
        self.primary_callback = primary_action_callback
        self.secondary_callback = secondary_action_callback
        self.theme = theme or ThemeTokens()
        self.result = False

        if not title:
            if variant == "success":
                title = "Proses Berhasil"
            elif variant == "warning":
                title = "Perhatian"
            elif variant == "error":
                title = "Terjadi Kesalahan"
            elif variant == "confirm":
                title = "Konfirmasi Tindakan"
        self.title(title)
        self.resizable(False, False)

        apply_window_icon(self)
        self.transient(parent)
        self.grab_set()

        if variant == "success":
            badge_icon = "✓"
            badge_bg = THEME["status_ready_bg"]
            badge_fg = THEME["status_ready_text"]
            accent_bar_color = ("#10B981", "#10B981")
            default_primary_lbl = "Buka Folder Hasil PDF"
            default_secondary_lbl = "Selesai"
        elif variant == "warning":
            badge_icon = "!"
            badge_bg = THEME["status_warn_bg"]
            badge_fg = THEME["status_warn_text"]
            accent_bar_color = ("#F59E0B", "#F59E0B")
            default_primary_lbl = "Mengerti"
            default_secondary_lbl = "Tutup"
        elif variant == "error":
            badge_icon = "✕"
            badge_bg = THEME["status_error_bg"]
            badge_fg = THEME["status_error_text"]
            accent_bar_color = ("#EF4444", "#EF4444")
            default_primary_lbl = "Mengerti & Tutup"
            default_secondary_lbl = "Coba Lagi"
        else:
            badge_icon = "?"
            badge_bg = THEME["btn_accent_soft_bg"]
            badge_fg = THEME["btn_accent_soft_text"]
            accent_bar_color = ("#2563EB", "#2563EB")
            default_primary_lbl = "Ya, Hentikan Proses"
            default_secondary_lbl = "Lanjutkan Mencetak"

        self.primary_label = primary_action_label or default_primary_lbl
        self.secondary_label = secondary_action_label

        card = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        card.pack(fill="both", expand=True, padx=2, pady=2)

        ctk.CTkFrame(
            card,
            height=4,
            corner_radius=2,
            fg_color=accent_bar_color,
        ).pack(fill="x", side="top")

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=(16, 16))

        hdr_frame = ctk.CTkFrame(content, fg_color="transparent")
        hdr_frame.pack(fill="x", pady=(0, 10))

        badge = ctk.CTkLabel(
            hdr_frame,
            text=badge_icon,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=15, weight="bold"),
            text_color=badge_fg,
            fg_color=badge_bg,
            corner_radius=16,
            width=32,
            height=32,
        )
        badge.pack(side="left", padx=(0, 12))

        title_lbl = ctk.CTkLabel(
            hdr_frame,
            text=title,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=15, weight="bold"),
            text_color=THEME["text_primary"],
            anchor="w",
        )
        title_lbl.pack(side="left", fill="x", expand=True)

        if self.summary_text:
            summary_lbl = ctk.CTkLabel(
                content,
                text=self.summary_text,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
                text_color=THEME["text_secondary"],
                anchor="w",
                justify="left",
                wraplength=480,
            )
            summary_lbl.pack(fill="x", pady=(0, 10))

        if self.info_box:
            info_frame = ctk.CTkFrame(
                content,
                fg_color=THEME["bg_elevated"],
                border_color=THEME["border_surface"],
                border_width=1,
                corner_radius=self.theme.corner_radius_inner,
            )
            info_frame.pack(fill="x", pady=(0, 10))

            if isinstance(self.info_box, dict):
                for k, v in self.info_box.items():
                    row_f = ctk.CTkFrame(info_frame, fg_color="transparent")
                    row_f.pack(fill="x", padx=12, pady=3)
                    ctk.CTkLabel(
                        row_f,
                        text=f"{k}:",
                        font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
                        text_color=THEME["text_secondary"],
                        anchor="w",
                    ).pack(side="left")
                    ctk.CTkLabel(
                        row_f,
                        text=str(v),
                        font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
                        text_color=THEME["text_primary"],
                        anchor="e",
                    ).pack(side="right")
            else:
                ctk.CTkLabel(
                    info_frame,
                    text=str(self.info_box),
                    font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
                    text_color=THEME["text_primary"],
                    justify="left",
                    anchor="w",
                ).pack(padx=12, pady=8, fill="x")

        if self.solution_text:
            sol_frame = ctk.CTkFrame(
                content,
                fg_color=THEME["bg_surface"],
                border_color=THEME["border_surface"],
                border_width=1,
                corner_radius=6,
            )
            sol_frame.pack(fill="x", pady=(0, 10))

            ctk.CTkLabel(
                sol_frame,
                text="Langkah Perbaikan:",
                font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
                text_color=THEME["text_primary"],
                anchor="w",
            ).pack(anchor="w", padx=10, pady=(6, 2))

            ctk.CTkLabel(
                sol_frame,
                text=self.solution_text,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
                text_color=THEME["text_secondary"],
                justify="left",
                anchor="w",
                wraplength=460,
            ).pack(anchor="w", padx=10, pady=(0, 6))

        self.tech_container = None
        if self.technical_detail:
            tech_toggle_row = ctk.CTkFrame(content, fg_color="transparent")
            tech_toggle_row.pack(fill="x", pady=(0, 4))

            self.btn_toggle_tech = ctk.CTkButton(
                tech_toggle_row,
                text="Lihat Detail Teknis ▾",
                font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
                text_color=THEME["text_muted"],
                fg_color="transparent",
                hover_color=THEME["bg_hover"],
                height=22,
                corner_radius=4,
                command=self._toggle_tech_details,
            )
            self.btn_toggle_tech.pack(side="left")

            self.tech_container = ctk.CTkFrame(
                content,
                fg_color=THEME["bg_surface"],
                border_color=THEME["border_surface"],
                border_width=1,
                corner_radius=6,
            )

            tech_box = ctk.CTkTextbox(
                self.tech_container,
                height=70,
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=10),
                text_color=THEME["text_secondary"],
                fg_color="transparent",
                wrap="word",
            )
            tech_box.pack(fill="both", expand=True, padx=8, pady=(6, 4))
            tech_box.insert("1.0", self.technical_detail)
            tech_box.configure(state="disabled")

            self.btn_copy_log = ctk.CTkButton(
                self.tech_container,
                text="Salin Log",
                width=80,
                height=24,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
                fg_color=THEME["btn_secondary_bg"],
                hover_color=THEME["btn_secondary_hover"],
                text_color=THEME["btn_secondary_text"],
                corner_radius=4,
                command=self._copy_tech_log,
            )
            self.btn_copy_log.pack(side="right", padx=8, pady=(0, 6))

        btn_row = ctk.CTkFrame(content, fg_color="transparent")
        btn_row.pack(fill="x", pady=(8, 0))

        if self.secondary_label:
            ctk.CTkButton(
                btn_row,
                text=self.secondary_label,
                height=34,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                fg_color=THEME["btn_secondary_bg"],
                hover_color=THEME["btn_secondary_hover"],
                text_color=THEME["btn_secondary_text"],
                border_color=THEME["border_surface"],
                border_width=1,
                corner_radius=self.theme.corner_radius_inner,
                command=self._on_secondary_action,
            ).pack(side="left")

        if variant == "confirm":
            pri_bg = THEME["btn_danger_bg"]
            pri_hover = THEME["btn_danger_hover"]
            pri_text = THEME["btn_danger_text"]
        elif variant == "success":
            pri_bg = THEME["btn_primary_bg"]
            pri_hover = THEME["btn_primary_hover"]
            pri_text = THEME["btn_primary_text"]
        else:
            pri_bg = THEME["btn_secondary_bg"]
            pri_hover = THEME["btn_secondary_hover"]
            pri_text = THEME["btn_secondary_text"]

        ctk.CTkButton(
            btn_row,
            text=self.primary_label,
            height=34,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            fg_color=pri_bg,
            hover_color=pri_hover,
            text_color=pri_text,
            corner_radius=self.theme.corner_radius_inner,
            command=self._on_primary_action,
        ).pack(side="right")

        self.bind("<Escape>", lambda e: self._on_escape())
        self.bind("<Return>", lambda e: self._on_return())

        self._adjust_geometry()

    def _adjust_geometry(self) -> None:
        self.update_idletasks()
        req_w = 540
        req_h = self.winfo_reqheight() + 20
        req_h = max(240, min(req_h, 640))

        pw = self.parent.winfo_width() if hasattr(self.parent, "winfo_width") else 1000
        ph = self.parent.winfo_height() if hasattr(self.parent, "winfo_height") else 700
        px = self.parent.winfo_rootx() if hasattr(self.parent, "winfo_rootx") else 100
        py = self.parent.winfo_rooty() if hasattr(self.parent, "winfo_rooty") else 100

        x = max(10, px + (pw - req_w) // 2)
        y = max(10, py + (ph - req_h) // 2)
        self.geometry(f"{req_w}x{req_h}+{x}+{y}")
        self.lift()
        self.focus_force()

    def _toggle_tech_details(self) -> None:
        if not self.tech_container:
            return
        if self.tech_container.winfo_ismapped():
            self.tech_container.pack_forget()
            self.btn_toggle_tech.configure(text="Lihat Detail Teknis ▾")
        else:
            self.tech_container.pack(fill="x", pady=(0, 8))
            self.btn_toggle_tech.configure(text="Sembunyikan Detail Teknis ▴")
        self._adjust_geometry()

    def _copy_tech_log(self) -> None:
        try:
            self.clipboard_clear()
            self.clipboard_append(self.technical_detail)
            self.btn_copy_log.configure(text="Tersalin ✓")
            self.after(2000, lambda: self.btn_copy_log.configure(text="Salin Log"))
        except Exception:
            pass

    def _on_primary_action(self) -> None:
        self.result = True
        self.destroy()
        if self.primary_callback:
            try:
                self.primary_callback()
            except Exception:
                pass

    def _on_secondary_action(self) -> None:
        self.result = False
        self.destroy()
        if self.secondary_callback:
            try:
                self.secondary_callback()
            except Exception:
                pass

    def _on_escape(self) -> None:
        if self.secondary_label and self.secondary_callback:
            self._on_secondary_action()
        else:
            self.destroy()

    def _on_return(self) -> None:
        self._on_primary_action()

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
        # Inner padding container
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=14, pady=10)

        # Baris 1: Header Atas (Kiri: Kategori dokumen, Kanan: Badge Status)
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

        # Baris 2: Nama File (Lebar Penuh 100%, 13px bold)
        filename = os.path.basename(self.current_path) if self.current_path else "Belum ada berkas dipilih"
        self.filename_label = ctk.CTkLabel(
            inner,
            text=filename,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=13, weight="bold"),
            text_color=THEME["text_primary"],
            anchor="w",
        )
        self.filename_label.pack(fill="x", pady=(2, 4))

        # Baris 3: Toolbar Tombol Aksi Proporsional (Anti-Clipping)
        self.btn_row = ctk.CTkFrame(inner, fg_color="transparent")
        self.btn_row.pack(fill="x", pady=(6, 0))

        is_word = "word" in self.doc_type.lower()
        browse_text = "Pilih Template Word" if is_word else "Pilih Berkas DPT"

        # 1. Tombol Pilih File (Aksen Soft Biru)
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

        # 2. Tombol Atur Kolom (Hanya jika on_map_columns disediakan)
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

        # 3. Tombol Cek Ulang / Scan Ulang (Tombol Sekunder Tegas)
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

        # 4. Tombol Pratinjau PDF (Hanya untuk Word)
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

        # Metric 1: Wilayah Terpilih
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

        # Divider 1
        div1 = ctk.CTkFrame(self, fg_color=THEME["border_surface"], width=1)
        div1.grid(row=0, column=1, sticky="ns", pady=8)

        # Metric 2: Total Pemilih
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

        # Divider 2
        div2 = ctk.CTkFrame(self, fg_color=THEME["border_surface"], width=1)
        div2.grid(row=0, column=3, sticky="ns", pady=8)

        # Metric 3: Estimasi Kertas
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

        # Right container for status & optional open PDF button
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
        # 1. Sub-Header Status Bar
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

        # 2. Column Header Bar
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

        # 3. Rows Container (CTkScrollableFrame)
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

        # Scrollbar visibility control
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
            # Alternating row colors
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

            # Col 0: NO
            lbl_no = ctk.CTkLabel(
                row_frame,
                text=f"{idx + 1:02d}",
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=12),
                text_color=THEME["text_muted"],
                anchor="w",
            )
            lbl_no.grid(row=0, column=0, padx=(10, 4), pady=5, sticky="w")
            _bind_hover(lbl_no, row_frame, bg_col)

            # Col 1: NAMA PEMILIH
            lbl_nama = ctk.CTkLabel(
                row_frame,
                text=voter.nama,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=12, weight="bold"),
                text_color=THEME["text_primary"],
                anchor="w",
            )
            lbl_nama.grid(row=0, column=1, padx=4, pady=5, sticky="w")
            _bind_hover(lbl_nama, row_frame, bg_col)

            # Col 2: L/P
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

            # Col 3: USIA
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

            # Col 4: ALAMAT PEMILIH
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

            # Col 5: WILAYAH (Pill badge)
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

class AboutModal(ctk.CTkToplevel):
    def __init__(self, master: any, theme: ThemeTokens) -> None:
        super().__init__(master)
        self.theme = theme

        self.title("Tentang VoterMail Engine")
        self.geometry("620x640")
        self.minsize(560, 500)
        self.configure(fg_color=THEME["bg_root"])
        apply_window_icon(self)

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

        self.bind("<Escape>", lambda e: self.destroy())
        self._build_ui()

    def _build_ui(self) -> None:
        content_frame = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        content_frame.pack(fill="both", expand=True, padx=16, pady=16)

        # Header
        header_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=16, pady=(16, 10))

        title_row = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_row.pack(fill="x")

        if APP_ICON_PNG.exists():
            try:
                from PIL import Image
                logo_pil = Image.open(APP_ICON_PNG)
                self.about_logo_img = ctk.CTkImage(light_image=logo_pil, dark_image=logo_pil, size=(36, 36))
                icon_lbl = ctk.CTkLabel(title_row, image=self.about_logo_img, text="")
                icon_lbl.pack(side="left", padx=(0, 10))
            except Exception:
                pass

        app_title = ctk.CTkLabel(
            title_row,
            text=__app_name__,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=20, weight="bold"),
            text_color=THEME["text_primary"],
        )
        app_title.pack(side="left")

        v_pill = ctk.CTkFrame(
            title_row,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
        )
        v_pill.pack(side="left", padx=10)
        ctk.CTkLabel(
            v_pill,
            text=f"v{__version__}",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
            text_color=THEME["status_ready_text"],
        ).pack(padx=8, pady=2)

        app_desc = ctk.CTkLabel(
            header_frame,
            text="Otomasi Cetak Surat Pemberitahuan Pemilih (Model C6) Pilkades",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_secondary"],
        )
        app_desc.pack(anchor="w", pady=(2, 0))

        # Author Card
        dev_card = ctk.CTkFrame(
            content_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
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
            text_color=THEME["text_primary"],
        )
        dev_title.pack(anchor="w")

        lic_label = ctk.CTkLabel(
            dev_info_left,
            text=__license__,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
            text_color=THEME["text_muted"],
        )
        lic_label.pack(anchor="w", pady=(2, 0))

        btn_github = ctk.CTkButton(
            dev_row,
            text="GitHub",
            command=lambda: webbrowser.open(__repository__),
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            text_color=THEME["btn_secondary_text"],
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
            text_color=THEME["text_primary"],
        )
        hist_title.pack(anchor="w", padx=16, pady=(10, 4))

        hist_scroll = ctk.CTkScrollableFrame(
            content_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            height=260,
        )
        hist_scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        footer_row = ctk.CTkFrame(content_frame, fg_color="transparent")
        footer_row.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_check_update = ctk.CTkButton(
            footer_row,
            text="Cek Pembaruan",
            height=32,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            corner_radius=self.theme.corner_radius_inner,
            command=self._manual_check_updates,
        )
        self.btn_check_update.pack(side="left")

        btn_close = ctk.CTkButton(
            footer_row,
            text="Tutup",
            height=32,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            corner_radius=self.theme.corner_radius_inner,
            command=self.destroy,
        )
        btn_close.pack(side="right")

        for feat in reversed(FEATURE_REGISTRY):
            item = ctk.CTkFrame(hist_scroll, fg_color="transparent")
            item.pack(fill="x", pady=4)

            top = ctk.CTkFrame(item, fg_color="transparent")
            top.pack(fill="x")

            v_lbl = ctk.CTkLabel(
                top,
                text=f"v{feat.get('version', '')}",
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
                text_color=THEME["status_ready_text"],
            )
            v_lbl.pack(side="left")

            t_lbl = ctk.CTkLabel(
                top,
                text=f"— {feat.get('title', '')}",
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
                text_color=THEME["text_primary"],
            )
            t_lbl.pack(side="left", padx=4)

            d_lbl = ctk.CTkLabel(
                top,
                text=feat.get("date", ""),
                font=ctk.CTkFont(family=self.theme.font_family_mono, size=10),
                text_color=THEME["text_muted"],
            )
            d_lbl.pack(side="right")

            desc = feat.get("description", "")
            if desc:
                ctk.CTkLabel(
                    item,
                    text=desc,
                    font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
                    text_color=THEME["text_secondary"],
                    wraplength=480,
                    justify="left",
                ).pack(anchor="w", pady=(1, 0))

    def _manual_check_updates(self) -> None:
        self.btn_check_update.configure(state="disabled", text="Memeriksa...")

        def run_check() -> None:
            from app.core.updater import check_for_updates
            info = check_for_updates(__version__)

            def finish() -> None:
                self.btn_check_update.configure(state="normal", text="Cek Pembaruan")
                if info.get("update_available"):
                    master = self.master
                    self.destroy()
                    UpdateNotificationModal(master, self.theme, info)
                elif info.get("error"):
                    ModernNotificationModal(
                        parent=self,
                        variant="warning",
                        title="Gagal Memeriksa Pembaruan",
                        summary=f"Tidak dapat terhubung ke server pembaruan: {info['error']}",
                        solution="Pastikan perangkat terhubung ke internet dan domain GitHub tidak diblokir.",
                        primary_action_label="Mengerti & Tutup",
                        theme=self.theme,
                    )
                else:
                    ModernNotificationModal(
                        parent=self,
                        variant="success",
                        title="Sudah Versi Terbaru",
                        summary=f"Anda saat ini sudah menggunakan versi resmi terbaru (v{__version__}).",
                        primary_action_label="Mengerti",
                        theme=self.theme,
                    )

            self.after(0, finish)

        threading.Thread(target=run_check, daemon=True).start()

class UpdateNotificationModal(ctk.CTkToplevel):
    def __init__(self, master: any, theme: ThemeTokens, update_info: dict[str, Any]) -> None:
        super().__init__(master)
        self.theme = theme
        self.update_info = update_info
        latest_ver = update_info.get("latest_version", "Baru")

        self.title(f"Pembaruan Tersedia: VoterMail Engine v{latest_ver}")
        self.geometry("580x530")
        self.minsize(520, 460)
        self.configure(fg_color=THEME["bg_root"])
        apply_window_icon(self)

        self.transient(master)
        self.grab_set()

        self.update_idletasks()
        try:
            mx = master.winfo_x()
            my = master.winfo_y()
            mw = master.winfo_width()
            mh = master.winfo_height()
            x = mx + (mw - 580) // 2
            y = my + (mh - 530) // 2
            self.geometry(f"580x530+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        self.bind("<Escape>", lambda e: self.destroy())
        self._build_ui()

    def _build_ui(self) -> None:
        content_frame = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        content_frame.pack(fill="both", expand=True, padx=16, pady=16)

        # Header Frame
        header_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=16, pady=(16, 10))

        title_row = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_row.pack(fill="x")

        if APP_ICON_PNG.exists():
            try:
                from PIL import Image
                logo_pil = Image.open(APP_ICON_PNG)
                self.modal_logo_img = ctk.CTkImage(light_image=logo_pil, dark_image=logo_pil, size=(36, 36))
                icon_lbl = ctk.CTkLabel(title_row, image=self.modal_logo_img, text="")
                icon_lbl.pack(side="left", padx=(0, 10))
            except Exception:
                pass

        latest_ver = self.update_info.get("latest_version", "Baru")
        curr_ver = self.update_info.get("current_version", __version__)

        app_title = ctk.CTkLabel(
            title_row,
            text=f"Pembaruan Tersedia: v{latest_ver}",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=18, weight="bold"),
            text_color=THEME["text_primary"],
        )
        app_title.pack(side="left")

        # Subtitle Version Pill
        ver_pill = ctk.CTkFrame(
            header_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
        )
        ver_pill.pack(anchor="w", pady=(8, 0))

        ver_text = f"Versi Anda saat ini: v{curr_ver}  →  Versi Terbaru: v{latest_ver}"
        ctk.CTkLabel(
            ver_pill,
            text=ver_text,
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11, weight="bold"),
            text_color=THEME["status_ready_text"],
        ).pack(padx=10, pady=4)

        # Release Notes Section
        notes_label = ctk.CTkLabel(
            content_frame,
            text=f"Apa yang Baru di Versi v{latest_ver}:",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            text_color=THEME["text_primary"],
        )
        notes_label.pack(anchor="w", padx=16, pady=(12, 4))

        self.notes_box = ctk.CTkTextbox(
            content_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=self.theme.corner_radius_inner,
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_secondary"],
            wrap="word",
            height=140,
        )
        self.notes_box.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        raw_notes = self.update_info.get("release_notes", "").strip()
        if not raw_notes:
            raw_notes = "Pembaruan kestabilan, peningkatan performa, dan penyempurnaan antarmuka VoterMail Engine."
        self.notes_box.insert("1.0", raw_notes)
        self.notes_box.configure(state="disabled")

        # Safety Reassurance Banner
        safe_frame = ctk.CTkFrame(
            content_frame,
            fg_color=THEME["status_ready_bg"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
        )
        safe_frame.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(
            safe_frame,
            text="✓ Seluruh file Excel DPT, template Word, dan hasil PDF Anda di folder data/ tetap aman 100% dan tidak akan terhapus.",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=10, weight="bold"),
            text_color=THEME["status_ready_text"],
            justify="left",
            wraplength=520,
        ).pack(anchor="w", padx=10, pady=6)

        # Progress Area (Initially Unpacked)
        self.progress_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_frame,
            height=10,
            corner_radius=5,
            fg_color=THEME["bg_surface"],
            progress_color=THEME["btn_primary_bg"],
        )
        self.progress_bar.pack(fill="x", pady=(2, 4))
        self.progress_bar.set(0.0)

        self.progress_label = ctk.CTkLabel(
            self.progress_frame,
            text="Menyiapkan unduhan...",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=11),
            text_color=THEME["text_secondary"],
        )
        self.progress_label.pack(anchor="w")

        # Action Buttons Row
        self.btn_row = ctk.CTkFrame(content_frame, fg_color="transparent")
        self.btn_row.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_later = ctk.CTkButton(
            self.btn_row,
            text="Nanti Saja",
            width=100,
            height=34,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            corner_radius=self.theme.corner_radius_inner,
            command=self.destroy,
        )
        self.btn_later.pack(side="left", padx=(0, 8))

        self.btn_github = ctk.CTkButton(
            self.btn_row,
            text="Lihat di GitHub",
            width=120,
            height=34,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_color=THEME["border_surface"],
            border_width=1,
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            corner_radius=self.theme.corner_radius_inner,
            command=lambda: webbrowser.open(self.update_info.get("html_url", f"{__repository__}/releases/latest")),
        )
        self.btn_github.pack(side="left")

        self.btn_update = ctk.CTkButton(
            self.btn_row,
            text=f"Update Sekarang (v{latest_ver})",
            height=34,
            fg_color=THEME["btn_primary_bg"],
            hover_color=THEME["btn_primary_hover"],
            text_color=THEME["btn_primary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            corner_radius=self.theme.corner_radius_inner,
            command=self._start_update,
        )
        self.btn_update.pack(side="right")

    def _start_update(self) -> None:
        self.btn_later.configure(state="disabled")
        self.btn_github.configure(state="disabled")
        self.btn_update.configure(state="disabled", text="Memperbarui...")

        self.progress_frame.pack(fill="x", padx=16, pady=(0, 10), before=self.btn_row)
        self.progress_bar.set(0.0)

        def worker() -> None:
            from app.core.updater import download_and_apply_update

            def prog_cb(cur: int, tot: int, msg: str) -> None:
                pct = cur / tot if tot > 0 else 0.0
                self.after(0, lambda: self._update_progress(pct, msg))

            dl_url = self.update_info.get("download_url", "")
            h_url = self.update_info.get("html_url", "")
            ok, msg = download_and_apply_update(dl_url, html_url=h_url, progress_callback=prog_cb)
            self.after(0, lambda: self._on_update_done(ok, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, pct: float, msg: str) -> None:
        self.progress_bar.set(pct)
        self.progress_label.configure(text=msg)

    def _on_update_done(self, ok: bool, msg: str) -> None:
        if ok:
            ModernNotificationModal(
                parent=self.master,
                variant="success",
                title="Pembaruan Berhasil",
                summary=msg,
                primary_action_label="Mengerti & Tutup",
                theme=self.theme,
            )
            self.destroy()
        else:
            ModernNotificationModal(
                parent=self,
                variant="error",
                title="Gagal Memperbarui",
                summary=msg,
                solution="Periksa koneksi internet Anda atau unduh paket rilis terbaru secara manual dari GitHub Releases.",
                primary_action_label="Mengerti & Tutup",
                theme=self.theme,
            )
            self.btn_later.configure(state="normal")
            self.btn_github.configure(state="normal")
            self.btn_update.configure(state="normal", text="Coba Update Lagi")
            self.progress_frame.pack_forget()

def get_col_letter(idx: int) -> str:
    result = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result

class ColumnMapperModal(ctk.CTkToplevel):
    """Modal pemetaan kolom Excel DPT kustom dengan proteksi Zero PII Leak."""

    def __init__(
        self,
        parent: Any,
        theme: ThemeTokens,
        available_columns: list[tuple[int, str]],
        current_mapping: dict[str, int],
        on_apply: Callable[[dict[str, int]], None],
        header_row_index: int = 0,
        **kwargs: Any,
    ) -> None:
        super().__init__(parent, **kwargs)
        self.theme = theme
        self.available_columns = available_columns
        self.current_mapping = dict(current_mapping)
        self.on_apply = on_apply
        self.header_row_index = header_row_index

        self.title("Pemetaan Kolom Excel DPT — VoterMail Engine")
        self.geometry("540x590")
        self.resizable(False, False)
        self.configure(fg_color=THEME["bg_root"])
        apply_window_icon(self)
        self.transient(parent)
        self.grab_set()

        self.field_vars: dict[str, ctk.StringVar] = {}
        self.label_to_col_idx: dict[str, int | None] = {}

        self._build_ui()

    def _build_ui(self) -> None:
        container = ctk.CTkFrame(
            self,
            fg_color=THEME["bg_card"],
            border_color=THEME["border_card"],
            border_width=1,
            corner_radius=self.theme.corner_radius_card,
        )
        container.pack(fill="both", expand=True, padx=16, pady=16)

        header_box = ctk.CTkFrame(container, fg_color="transparent")
        header_box.pack(fill="x", padx=16, pady=(14, 8))

        ctk.CTkLabel(
            header_box,
            text="Pemetaan Kolom Data Excel",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=15, weight="bold"),
            text_color=THEME["text_primary"],
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_box,
            text="Pilih kolom yang sesuai dari berkas Excel Anda. Kolom NIK diblokir permanen demi keamanan data warga.",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            text_color=THEME["text_secondary"],
            wraplength=480,
            justify="left",
        ).pack(anchor="w", pady=(2, 4))

        info_pill = ctk.CTkFrame(
            header_box,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
        )
        info_pill.pack(anchor="w", pady=(4, 0))

        ctk.CTkLabel(
            info_pill,
            text=f"Baris Header Terdeteksi: Baris {self.header_row_index + 1} • {len(self.available_columns)} Kolom Terbaca",
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=10, weight="bold"),
            text_color=THEME["status_ready_text"],
        ).pack(padx=8, pady=3)

        empty_label = "(Tidak Digunakan / Kosongkan)"
        self.label_to_col_idx[empty_label] = None
        col_options = [empty_label]
        col_idx_to_label: dict[int, str] = {}

        for col_idx, col_name in self.available_columns:
            if not is_blocked_pii(col_name):
                col_letter = get_col_letter(col_idx)
                opt_str = f"{col_name} (Kolom {col_letter})"
                col_options.append(opt_str)
                self.label_to_col_idx[opt_str] = col_idx
                col_idx_to_label[col_idx] = opt_str

        form_scroll = ctk.CTkScrollableFrame(
            container,
            fg_color="transparent",
            height=320,
        )
        form_scroll.pack(fill="both", expand=True, padx=16, pady=(4, 8))

        fields = [
            ("nama", "Nama Pemilih *", True, "Kolom utama nama lengkap pemilih pada surat."),
            ("no", "Nomor Urut", False, "Bawaan penomoran otomatis jika dikosongkan."),
            ("jk", "Jenis Kelamin (L/P)", False, "Diubah otomatis menjadi Laki-laki / Perempuan."),
            ("umur", "Usia / Umur", False, "Otomatis diformat dengan akhiran 'Tahun'."),
            ("alamat", "Alamat Domisili", False, "Alamat tempat tinggal pemilih."),
            ("rt", "RT (Opsional)", False, "Nomor RT pemilih."),
            ("rw", "RW (Opsional)", False, "Nomor RW pemilih."),
        ]

        for key, label_text, is_required, hint in fields:
            row_frame = ctk.CTkFrame(form_scroll, fg_color="transparent")
            row_frame.pack(fill="x", pady=5)

            left_box = ctk.CTkFrame(row_frame, fg_color="transparent")
            left_box.pack(side="left", fill="x", expand=True)

            lbl_color = THEME["text_primary"] if not is_required else THEME["btn_accent_soft_text"]
            ctk.CTkLabel(
                left_box,
                text=label_text,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=12, weight="bold"),
                text_color=lbl_color,
            ).pack(anchor="w")

            ctk.CTkLabel(
                left_box,
                text=hint,
                font=ctk.CTkFont(family=self.theme.font_family_main, size=10),
                text_color=THEME["text_muted"],
            ).pack(anchor="w")

            curr_idx = self.current_mapping.get(key)
            default_val = col_idx_to_label.get(curr_idx, empty_label) if curr_idx is not None else empty_label

            var = ctk.StringVar(value=default_val)
            self.field_vars[key] = var

            opt_menu = ctk.CTkOptionMenu(
                row_frame,
                values=col_options,
                variable=var,
                width=230,
                height=30,
                corner_radius=self.theme.corner_radius_control,
                fg_color=THEME["bg_surface"],
                button_color=THEME["btn_secondary_bg"],
                button_hover_color=THEME["btn_secondary_hover"],
                text_color=THEME["text_primary"],
                font=ctk.CTkFont(family=self.theme.font_family_main, size=11),
            )
            opt_menu.pack(side="right", padx=(8, 0))

        btn_row = ctk.CTkFrame(container, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(10, 14))

        ctk.CTkButton(
            btn_row,
            text="Batal",
            width=90,
            height=34,
            fg_color=THEME["btn_secondary_bg"],
            hover_color=THEME["btn_secondary_hover"],
            border_width=1,
            border_color=THEME["border_surface"],
            text_color=THEME["btn_secondary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            corner_radius=self.theme.corner_radius_control,
            command=self.destroy,
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            btn_row,
            text="Terapkan & Scan Ulang",
            height=34,
            fg_color=THEME["btn_primary_bg"],
            hover_color=THEME["btn_primary_hover"],
            text_color=THEME["btn_primary_text"],
            font=ctk.CTkFont(family=self.theme.font_family_main, size=11, weight="bold"),
            corner_radius=self.theme.corner_radius_control,
            command=self._apply_and_close,
        ).pack(side="right")

    def _apply_and_close(self) -> None:
        nama_choice = self.field_vars["nama"].get()
        nama_idx = self.label_to_col_idx.get(nama_choice)
        if nama_idx is None:
            ModernNotificationModal(
                parent=self,
                variant="warning",
                title="Kolom Nama Diperlukan",
                summary="Kolom 'Nama Pemilih' wajib ditentukan agar aplikasi dapat mencetak surat pemanggilan.",
                solution="Pilih kolom yang memuat nama warga dari daftar, lalu klik 'Terapkan & Scan Ulang'.",
                primary_action_label="Mengerti & Lengkapi",
                theme=self.theme,
            )
            return

        new_map: dict[str, int] = {}
        for k, var in self.field_vars.items():
            choice = var.get()
            idx = self.label_to_col_idx.get(choice)
            if idx is not None:
                new_map[k] = idx

        self.on_apply(new_map)
        self.destroy()
