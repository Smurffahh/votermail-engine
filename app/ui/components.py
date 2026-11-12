from typing import Callable, Literal
import customtkinter as ctk
from app.config.theme import ThemeTokens

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

        # Smart clean label without redundant naming
        clean_sheet = sheet_name.strip()
        clean_wilayah = wilayah_text.strip()
        if clean_sheet.lower() == clean_wilayah.lower():
            display_title = clean_sheet
        elif clean_wilayah:
            display_title = f"{clean_sheet} ({clean_wilayah})"
        else:
            display_title = clean_sheet

        self.checkbox = ctk.CTkCheckBox(
            self,
            text=display_title,
            variable=self.checkbox_var,
            font=ctk.CTkFont(family=theme.font_family_main, size=12, weight="bold"),
            command=on_toggle,
            checkbox_width=18,
            checkbox_height=18,
            corner_radius=4,
            border_width=2,
        )
        self.checkbox.pack(side="left", padx=(6, 8), pady=4)

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
