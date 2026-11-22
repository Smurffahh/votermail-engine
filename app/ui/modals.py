# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import os
from pathlib import Path
import sys
import threading
from typing import Any, Callable, Literal
import webbrowser
import customtkinter as ctk

from app.config.settings import APP_ICON_ICO, APP_ICON_PNG
from app.config.theme import THEME, ThemeTokens
from app.config.version import (
    __app_name__,
    __author__,
    __github_user__,
    __license__,
    __repository__,
    __version__,
    FEATURE_REGISTRY,
)
from app.core.excel_scanner import is_blocked_pii

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

    # Re-apply via timer to ensure Windows DWM respects the icon during window realization
    if hasattr(window, "after"):
        try:
            window.after(220, _set_icon)
            window.after(300, _set_icon)
        except Exception:
            pass

def get_col_letter(idx: int) -> str:
    result = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result

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

class AboutModal(ctk.CTkToplevel):
    def __init__(self, master: Any, theme: ThemeTokens) -> None:
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
    def __init__(self, master: Any, theme: ThemeTokens, update_info: dict[str, Any]) -> None:
        super().__init__(master)
        self.theme = theme
        self.update_info = update_info
        latest_ver = update_info.get("latest_version", "Baru")

        self.title(f"Pembaruan Tersedia: VoterMail Engine v{latest_ver}")
        self.geometry("580x570")
        self.minsize(520, 480)
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
            y = my + (mh - 570) // 2
            self.geometry(f"580x570+{max(0, x)}+{max(0, y)}")
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
        update_mode = str(self.update_info.get("update_mode", "FULL")).upper()
        dl_size = self.update_info.get("download_size_bytes", 0)

        from app.core.updater import format_file_size
        size_str = format_file_size(dl_size) if dl_size > 0 else ""
        mode_badge_text = "Patch Cepat" if update_mode == "PATCH" else "Pembaruan Penuh"
        badge_parts = [mode_badge_text]
        if size_str:
            badge_parts.append(size_str)
        badge_parts.append("Verifikasi SHA-256 Otomatis")
        full_badge_text = " • ".join(badge_parts)

        app_title = ctk.CTkLabel(
            title_row,
            text=f"Pembaruan Tersedia: v{latest_ver}",
            font=ctk.CTkFont(family=self.theme.font_family_main, size=18, weight="bold"),
            text_color=THEME["text_primary"],
        )
        app_title.pack(side="left")

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

        mode_pill = ctk.CTkFrame(
            header_frame,
            fg_color=THEME["bg_surface"],
            border_color=THEME["border_surface"],
            border_width=1,
            corner_radius=6,
        )
        mode_pill.pack(anchor="w", pady=(6, 0))

        ctk.CTkLabel(
            mode_pill,
            text=full_badge_text,
            font=ctk.CTkFont(family=self.theme.font_family_mono, size=10, weight="bold"),
            text_color=THEME["status_ready_text"] if update_mode == "PATCH" else THEME["btn_primary_bg"],
        ).pack(padx=10, pady=3)

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
            height=130,
        )
        self.notes_box.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        raw_notes = self.update_info.get("release_notes", "").strip()
        if not raw_notes:
            raw_notes = "Pembaruan kestabilan, peningkatan performa, dan penyempurnaan antarmuka VoterMail Engine."
        self.notes_box.insert("1.0", raw_notes)
        self.notes_box.configure(state="disabled")

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

        if size_str:
            btn_update_label = f"Update Sekarang ({mode_badge_text} • {size_str})"
        else:
            btn_update_label = f"Update Sekarang ({mode_badge_text})"

        self.btn_update = ctk.CTkButton(
            self.btn_row,
            text=btn_update_label,
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
        if hasattr(self.master, "_prepare_shutdown_for_update"):
            if not self.master._prepare_shutdown_for_update():
                return

        self.btn_later.configure(state="disabled")
        self.btn_github.configure(state="disabled")
        self.btn_update.configure(state="disabled", text="Memperbarui...")

        self.progress_frame.pack(fill="x", padx=16, pady=(0, 10), before=self.btn_row)
        self.progress_bar.set(0.0)

        def worker() -> None:
            is_frozen = getattr(sys, "frozen", False)
            dl_url = self.update_info.get("download_url", "")
            h_url = self.update_info.get("html_url", "")
            expected_sha = self.update_info.get("expected_sha256", "")
            update_mode = self.update_info.get("update_mode", "FULL")
            target_ver = self.update_info.get("latest_version", "") or "latest"

            def prog_cb(cur: int, tot: int, msg: str) -> None:
                pct = cur / tot if tot > 0 else 0.0
                self.after(0, lambda: self._update_progress(pct, msg))

            if not is_frozen:
                from app.core.updater import download_and_apply_update
                ok, msg = download_and_apply_update(
                    download_url=dl_url,
                    html_url=h_url,
                    progress_callback=prog_cb,
                    expected_sha256=expected_sha,
                    update_mode=update_mode,
                    target_version=target_ver,
                )
                self.after(0, lambda: self._on_update_done(ok, msg))
                return

            from app.core.updater import stage_update_package, launch_detached_patcher

            ok, msg, staging_dir, source_dir = stage_update_package(
                download_url=dl_url,
                target_version=target_ver,
                expected_sha256=expected_sha,
                update_mode=update_mode,
                progress_callback=prog_cb,
            )

            if not ok or not staging_dir or not source_dir:
                self.after(0, lambda: self._on_update_done(False, msg))
                return

            self.after(0, lambda: self._update_progress(1.0, "Meluncurkan proses pemasangan..."))

            app_dir = Path(sys.executable).resolve().parent
            exe_path = app_dir / "VoterMail_Engine.exe"
            backup_dir = staging_dir.parent / "backup"

            launch_ok, launch_msg = launch_detached_patcher(
                parent_pid=os.getpid(),
                app_dir=app_dir,
                exe_path=exe_path,
                staging_dir=source_dir,
                backup_dir=backup_dir,
                update_mode=update_mode,
                target_version=target_ver,
                icon_path=staging_dir.parent / "votermail.ico",
                theme_mode=self.theme.name,
                is_git=False,
            )

            if not launch_ok:
                self.after(0, lambda: self._on_update_done(False, launch_msg))
                return

            if hasattr(self.master, "_shutdown_and_exit_for_update"):
                self.after(200, self.master._shutdown_and_exit_for_update)
            else:
                self.after(200, lambda: os._exit(0))

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

class ColumnMapperModal(ctk.CTkToplevel):
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
