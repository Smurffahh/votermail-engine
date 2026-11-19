# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
from typing import Any, Callable
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zipfile

from app.config.settings import BASE_DIR, TEMP_DIR
from app.config.version import (
    __version__,
    __repository__,
    parse_version_tuple,
)

GITHUB_API_LATEST_RELEASE_URL = "https://api.github.com/repos/Smurffahh/votermail-engine/releases/latest"

ALLOWED_UPDATE_HOSTS = {
    "github.com",
    "api.github.com",
    "objects.githubusercontent.com",
    "raw.githubusercontent.com",
    "github-releases.githubusercontent.com",
}

def validate_update_url(url: str) -> bool:
    if not url or not isinstance(url, str):
        return False
    try:
        parsed = urllib.parse.urlparse(url.strip())
        if parsed.scheme.lower() != "https":
            return False
        host = (parsed.hostname or "").lower()
        if host not in ALLOWED_UPDATE_HOSTS:
            return False
        if host in ("github.com", "api.github.com"):
            path = parsed.path.lower()
            if not ("smurffahh/votermail-engine" in path):
                return False
        return True
    except Exception:
        return False

def safe_extract_zip(zip_path: str | Path, target_dir: str | Path) -> list[Path]:
    z_path = Path(zip_path).resolve()
    t_dir = Path(target_dir).resolve()
    t_dir.mkdir(parents=True, exist_ok=True)

    extracted_files: list[Path] = []
    with zipfile.ZipFile(z_path, "r") as zf:
        for member in zf.infolist():
            member_path = Path(member.filename)
            if member_path.is_absolute() or ".." in member_path.parts:
                raise PermissionError(f"Potensi eksploitasi Zip Slip terdeteksi pada file: {member.filename}")

            dest = (t_dir / member_path).resolve()
            if not dest.is_relative_to(t_dir):
                raise PermissionError(f"Target ekstraksi di luar direktori aman: {member.filename}")

            zf.extract(member, t_dir)
            extracted_files.append(dest)
    return extracted_files

def check_for_updates(current_version: str = __version__, timeout: float = 5.0) -> dict[str, Any]:
    req = urllib.request.Request(
        GITHUB_API_LATEST_RELEASE_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "VoterMail-Engine",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            if response.status != 200:
                return {
                    "update_available": False,
                    "current_version": current_version,
                    "latest_version": current_version,
                    "release_title": "",
                    "release_notes": "",
                    "download_url": "",
                    "html_url": f"{__repository__}/releases/latest",
                    "published_at": "",
                    "error": f"HTTP {response.status}",
                }
            raw_data = response.read().decode("utf-8")
            data = json.loads(raw_data)

            tag_name = data.get("tag_name", "").strip()
            latest_clean = re.sub(r"^[^\d]*", "", tag_name) or current_version

            is_newer = parse_version_tuple(tag_name) > parse_version_tuple(current_version)

            assets = data.get("assets", [])
            download_url = ""
            for a in assets:
                name = a.get("name", "").lower()
                if name.endswith(".zip") and "votermail" in name:
                    download_url = a.get("browser_download_url", "")
                    break
            if not download_url:
                for a in assets:
                    if a.get("name", "").lower().endswith(".zip"):
                        download_url = a.get("browser_download_url", "")
                        break

            return {
                "update_available": is_newer,
                "current_version": current_version,
                "latest_version": latest_clean,
                "release_title": data.get("name") or f"VoterMail Engine v{latest_clean}",
                "release_notes": data.get("body", ""),
                "download_url": download_url,
                "html_url": data.get("html_url") or f"{__repository__}/releases/latest",
                "published_at": data.get("published_at", ""),
                "error": None,
            }
    except Exception as exc:
        return {
            "update_available": False,
            "current_version": current_version,
            "latest_version": current_version,
            "release_title": "",
            "release_notes": "",
            "download_url": "",
            "html_url": f"{__repository__}/releases/latest",
            "published_at": "",
            "error": str(exc),
        }

def download_and_apply_update(
    download_url: str,
    html_url: str = "",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[bool, str]:
    is_frozen = getattr(sys, "frozen", False)

    if not is_frozen:
        git_dir = BASE_DIR / ".git"
        if git_dir.exists():
            if progress_callback:
                progress_callback(1, 2, "Mengambil pembaruan dari repositori Git...")
            res = subprocess.run(
                ["git", "pull", "--ff-only", "origin", "main"],
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                check=False,
            )
            if res.returncode == 0:
                if progress_callback:
                    progress_callback(2, 2, "Pembaruan Git selesai.")
                return True, "Berhasil memperbarui kode sumber dari Git. Silakan restart aplikasi."
            return False, f"Gagal git pull: {res.stderr.strip() or res.stdout.strip()}"

        target_url = html_url or f"{__repository__}/releases/latest"
        webbrowser.open(target_url)
        return True, "Halaman rilis terbaru telah dibuka di peramban Anda untuk pengunduhan manual."

    if not validate_update_url(download_url):
        return False, "URL pembaruan tidak valid atau tidak berasal dari domain resmi GitHub."

    temp_zip = Path(tempfile.mktemp(suffix=".zip", dir=str(TEMP_DIR)))
    req = urllib.request.Request(
        download_url,
        headers={"User-Agent": "VoterMail-Engine"},
    )

    try:
        with urllib.request.urlopen(req) as resp, open(temp_zip, "wb") as f_out:
            total_size = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            chunk_size = 65536

            while True:
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                f_out.write(chunk)
                downloaded += len(chunk)

                if progress_callback:
                    if total_size > 0:
                        pct = (downloaded / total_size) * 100
                        mb_dl = downloaded / (1024 * 1024)
                        mb_tot = total_size / (1024 * 1024)
                        msg = f"Mengunduh pembaruan... {pct:.0f}% ({mb_dl:.1f} / {mb_tot:.1f} MB)"
                    else:
                        mb_dl = downloaded / (1024 * 1024)
                        msg = f"Mengunduh pembaruan... ({mb_dl:.1f} MB)"
                    progress_callback(downloaded, total_size, msg)

        staging_dir = TEMP_DIR / "_update_staging"
        if staging_dir.exists():
            shutil.rmtree(staging_dir, ignore_errors=True)

        safe_extract_zip(temp_zip, staging_dir)

        source_dir = None
        if (staging_dir / "VoterMail_Engine" / "VoterMail_Engine.exe").exists():
            source_dir = staging_dir / "VoterMail_Engine"
        elif (staging_dir / "VoterMail_Engine.exe").exists():
            source_dir = staging_dir
        else:
            matches = list(staging_dir.rglob("VoterMail_Engine.exe"))
            if matches:
                source_dir = matches[0].parent

        if not source_dir:
            return False, "File VoterMail_Engine.exe tidak ditemukan di dalam paket rilis yang diunduh."

        app_dir = Path(sys.executable).resolve().parent
        app_exe = app_dir / "VoterMail_Engine.exe"

        bat_path = TEMP_DIR / "_apply_update.bat"
        bat_lines = [
            "@echo off",
            "setlocal",
            "echo Memasang pembaruan VoterMail Engine...",
            "ping 127.0.0.1 -n 3 >nul",
            f'if exist "{source_dir}\\_internal" (',
            f'    xcopy /E /Y /I "{source_dir}\\_internal" "{app_dir}\\_internal" >nul',
            ")",
            f'if exist "{source_dir}\\VoterMail_Engine.exe" (',
            f'    copy /Y "{source_dir}\\VoterMail_Engine.exe" "{app_dir}\\VoterMail_Engine.exe" >nul',
            ")",
            f'if exist "{source_dir}\\PANDUAN_PANITIA.txt" (',
            f'    copy /Y "{source_dir}\\PANDUAN_PANITIA.txt" "{app_dir}\\PANDUAN_PANITIA.txt" >nul',
            ")",
            f'rmdir /S /Q "{staging_dir}" >nul',
            f'start "" "{app_exe}"',
            'del "%~f0"',
            "",
        ]
        bat_content = "\r\n".join(bat_lines)
        bat_path.write_bytes(bat_content.encode("ascii"))

        if progress_callback:
            progress_callback(total_size, total_size, "Pembaruan siap dipasang. Menjalankan installer...")

        CREATE_NO_WINDOW = 0x08000000
        DETACHED_PROCESS = 0x00000008
        subprocess.Popen(
            ["cmd.exe", "/c", str(bat_path)],
            creationflags=CREATE_NO_WINDOW | DETACHED_PROCESS,
            close_fds=True,
        )
        sys.exit(0)

    except Exception as exc:
        return False, f"Gagal mengunduh atau mengekstrak pembaruan: {exc}"
    finally:
        if temp_zip.exists():
            try:
                temp_zip.unlink(missing_ok=True)
            except Exception:
                pass
