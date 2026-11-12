# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import importlib.util
import multiprocessing
import os
import subprocess
import sys
import threading

REQUIRED_PACKAGES = {
    "customtkinter": "customtkinter>=5.2.0",
    "pandas": "pandas>=2.0.0",
    "openpyxl": "openpyxl>=3.1.0",
    "docx": "python-docx>=1.1.0",
    "docxtpl": "docxtpl>=0.16.0",
    "docxcompose": "docxcompose>=1.4.0",
    "win32com": "pywin32>=306",
    "pypdf": "pypdf>=4.0.0",
    "pydantic": "pydantic>=2.0.0",
}

def get_missing_packages() -> dict[str, str]:
    missing = {}
    for mod_name, spec in REQUIRED_PACKAGES.items():
        if importlib.util.find_spec(mod_name) is None:
            missing[mod_name] = spec
    return missing

def install_missing_packages_with_gui(missing: dict[str, str]) -> bool:
    import tkinter as tk
    from tkinter import ttk, messagebox

    root = tk.Tk()
    root.title("VoterMail Engine — Persiapan Sistem")
    root.geometry("480x200")
    root.resizable(False, False)
    root.configure(bg="#0B0F19")

    root.update_idletasks()
    x = (root.winfo_screenwidth() - 480) // 2
    y = (root.winfo_screenheight() - 200) // 2
    root.geometry(f"480x200+{x}+{y}")

    lbl_title = tk.Label(
        root,
        text="Menyiapkan VoterMail Engine (by Danang Satria) untuk Penggunaan Pertama Kali...",
        font=("Segoe UI", 11, "bold"),
        fg="#F8FAFC",
        bg="#0B0F19",
        wraplength=440,
        justify="left",
    )
    lbl_title.pack(anchor="w", padx=20, pady=(20, 8))

    pkg_list_str = ", ".join(missing.keys())
    lbl_desc = tk.Label(
        root,
        text=f"Menginstal komponen yang dibutuhkan ({pkg_list_str}).\nMohon tunggu sebentar...",
        font=("Segoe UI", 9),
        fg="#94A3B8",
        bg="#0B0F19",
        wraplength=440,
        justify="left",
    )
    lbl_desc.pack(anchor="w", padx=20, pady=(0, 15))

    progress = ttk.Progressbar(root, mode="indeterminate", length=440)
    progress.pack(padx=20, pady=(0, 15))
    progress.start(10)

    lbl_status = tk.Label(
        root,
        text="Sedang mengunduh dan memasang paket resmi via pip...",
        font=("Segoe UI", 8),
        fg="#38BDF8",
        bg="#0B0F19",
    )
    lbl_status.pack(anchor="w", padx=20)

    success = [False]
    error_msg = [""]

    def run_install() -> None:
        specs_to_install = list(missing.values())
        cmd = [sys.executable, "-m", "pip", "install", "--upgrade", *specs_to_install]
        try:
            subprocess.check_call(cmd)
            importlib.invalidate_caches()
            success[0] = True
        except Exception as exc:
            error_msg[0] = str(exc)
        finally:
            root.after(0, root.destroy)

    thread = threading.Thread(target=run_install, daemon=True)
    thread.start()

    root.mainloop()

    if not success[0]:
        fallback_root = tk.Tk()
        fallback_root.withdraw()
        messagebox.showerror(
            "Gagal Memasang Dependensi",
            f"VoterMail Engine tidak dapat memasang paket secara otomatis:\n\n{error_msg[0]}\n\n"
            "Pastikan komputer terhubung ke internet, atau jalankan perintah berikut secara manual di terminal:\n"
            f"pip install {' '.join(missing.values())}",
        )
        fallback_root.destroy()
        return False

    return True

def ensure_dependencies() -> None:
    if getattr(sys, "frozen", False):
        return
    missing = get_missing_packages()
    if not missing:
        return
    ok = install_missing_packages_with_gui(missing)
    if not ok:
        sys.exit(1)

def main() -> None:
    ensure_dependencies()
    from app.ui.main_window import MainWindow

    app = MainWindow()
    app.mainloop()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
