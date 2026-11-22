# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import hashlib
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
    "release-assets.githubusercontent.com",
}

def compute_deps_hash(requirements_content_or_path: str | Path | None = None) -> str:
    """Menghitung SHA-256 dari requirements.txt yang dinormalisasi."""
    if requirements_content_or_path is None:
        req_file = BASE_DIR / "requirements.txt"
        if req_file.exists():
            text = req_file.read_text(encoding="utf-8")
        else:
            text = ""
    elif isinstance(requirements_content_or_path, Path) or (
        isinstance(requirements_content_or_path, str)
        and "\n" not in requirements_content_or_path
        and Path(requirements_content_or_path).exists()
    ):
        text = Path(requirements_content_or_path).read_text(encoding="utf-8")
    else:
        text = str(requirements_content_or_path)

    lines: list[str] = []
    for raw_line in text.splitlines():
        cleaned = raw_line.split("#")[0].strip().lower()
        if cleaned:
            lines.append(cleaned)
    lines.sort()
    normalized = "\n".join(lines)
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()

def compute_file_sha256(file_path: str | Path) -> str:
    """Menghitung SHA-256 berkas biner secara streaming."""
    p = Path(file_path)
    if not p.exists():
        return ""
    h = hashlib.sha256()
    with open(p, "rb") as f:
        while True:
            chunk = f.read(65536)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()

def get_local_build_meta() -> dict[str, Any]:
    """Membaca metadata build lokal (version dan deps_hash) dari aset terdistribusi."""
    candidates = [
        BASE_DIR / "app" / "assets" / "build_meta.json",
        Path(sys.executable).resolve().parent / "app" / "assets" / "build_meta.json",
    ]
    for c in candidates:
        if c.exists():
            try:
                data = json.loads(c.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
    return {
        "version": __version__,
        "deps_hash": compute_deps_hash(),
    }

def format_file_size(size_bytes: int | float | None) -> str:
    """Format ukuran berkas ke notasi desimal koma Indonesia (KB/MB)."""
    if not size_bytes or size_bytes <= 0:
        return "0 MB"
    size = float(size_bytes)
    if size >= 1024 * 1024:
        mb = size / (1024 * 1024)
        return f"{mb:.1f} MB".replace(".", ",")
    elif size >= 1024:
        kb = size / 1024
        return f"{kb:.0f} KB"
    return f"{int(size)} B"

def validate_update_url(url: str) -> bool:
    """Validasi ketat skema HTTPS dan domain resmi GitHub."""
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
    """Ekstraksi arsip ZIP dengan pengaman anti-Zip-Slip."""
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
    """Memeriksa rilis terbaru di GitHub dan menentukan mode PATCH vs FULL."""
    req = urllib.request.Request(
        GITHUB_API_LATEST_RELEASE_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "VoterMail-Engine",
        },
    )
    fallback_res = {
        "update_available": False,
        "current_version": current_version,
        "latest_version": current_version,
        "release_title": "",
        "release_notes": "",
        "download_url": "",
        "html_url": f"{__repository__}/releases/latest",
        "published_at": "",
        "update_mode": "FULL",
        "download_size_bytes": 0,
        "expected_sha256": "",
        "error": None,
    }

    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            if response.status != 200:
                res = dict(fallback_res)
                res["error"] = f"HTTP {response.status}"
                return res
            raw_data = response.read().decode("utf-8")
            data = json.loads(raw_data)

            tag_name = data.get("tag_name", "").strip()
            latest_clean = re.sub(r"^[^\d]*", "", tag_name) or current_version
            is_newer = parse_version_tuple(tag_name) > parse_version_tuple(current_version)

            assets = data.get("assets", [])
            manifest_asset = None
            patch_asset = None
            full_asset = None

            for a in assets:
                name = a.get("name", "").lower()
                if name == "update_manifest.json":
                    manifest_asset = a
                elif name.endswith("_patch.zip") or ("patch" in name and name.endswith(".zip")):
                    patch_asset = a
                elif name.endswith(".zip") and "votermail" in name:
                    full_asset = a

            if not full_asset:
                for a in assets:
                    if a.get("name", "").lower().endswith(".zip"):
                        full_asset = a
                        break

            manifest_data: dict[str, Any] | None = None
            if manifest_asset and validate_update_url(manifest_asset.get("browser_download_url", "")):
                try:
                    m_req = urllib.request.Request(
                        manifest_asset["browser_download_url"],
                        headers={"User-Agent": "VoterMail-Engine"},
                    )
                    with urllib.request.urlopen(m_req, timeout=timeout) as m_resp:
                        if m_resp.status == 200:
                            manifest_data = json.loads(m_resp.read().decode("utf-8"))
                except Exception:
                    manifest_data = None

            local_meta = get_local_build_meta()
            local_deps_hash = str(local_meta.get("deps_hash", "")).strip().lower()

            update_mode = "FULL"
            download_url = full_asset.get("browser_download_url", "") if full_asset else ""
            download_size_bytes = int(full_asset.get("size", 0)) if full_asset else 0
            expected_sha256 = ""

            if manifest_data:
                remote_deps_hash = str(manifest_data.get("deps_hash", "")).strip().lower()
                patch_info = manifest_data.get("patch") or manifest_data.get("assets", {}).get("patch_zip", {})
                full_info = manifest_data.get("full") or manifest_data.get("assets", {}).get("full_zip", {})

                if local_deps_hash and local_deps_hash == remote_deps_hash and (patch_asset or patch_info):
                    update_mode = "PATCH"
                    if patch_asset:
                        download_url = patch_asset.get("browser_download_url", "")
                    expected_sha256 = patch_info.get("sha256", "")
                    raw_size = patch_info.get("size_bytes")
                    if raw_size is None:
                        raw_size = patch_info.get("size")
                    if raw_size is None and patch_asset:
                        raw_size = patch_asset.get("size", 0)
                    download_size_bytes = int(raw_size or 0)
                else:
                    update_mode = "FULL"
                    if full_asset:
                        download_url = full_asset.get("browser_download_url", "")
                    expected_sha256 = full_info.get("sha256", "")
                    raw_size = full_info.get("size_bytes")
                    if raw_size is None:
                        raw_size = full_info.get("size")
                    if raw_size is None and full_asset:
                        raw_size = full_asset.get("size", 0)
                    download_size_bytes = int(raw_size or 0)

            return {
                "update_available": is_newer,
                "current_version": current_version,
                "latest_version": latest_clean,
                "release_title": data.get("name") or f"VoterMail Engine v{latest_clean}",
                "release_notes": data.get("body", ""),
                "download_url": download_url,
                "html_url": data.get("html_url") or f"{__repository__}/releases/latest",
                "published_at": data.get("published_at", ""),
                "update_mode": update_mode,
                "download_size_bytes": download_size_bytes,
                "expected_sha256": expected_sha256,
                "error": None,
            }
    except Exception as exc:
        res = dict(fallback_res)
        res["error"] = str(exc)
        return res

def stage_update_package(
    download_url: str,
    target_version: str,
    expected_sha256: str = "",
    update_mode: str = "FULL",
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> tuple[bool, str, Path | None, Path | None]:
    """Mengunduh berkas ke %TEMP%, memverifikasi SHA-256, dan mengekstrak ke folder staging."""
    if not validate_update_url(download_url):
        return False, "URL pembaruan tidak valid atau tidak berasal dari domain resmi GitHub.", None, None

    target_dir = Path(tempfile.gettempdir()) / f"VoterMail_Update_{target_version}"
    if target_dir.exists():
        shutil.rmtree(target_dir, ignore_errors=True)
    target_dir.mkdir(parents=True, exist_ok=True)

    package_zip = target_dir / "package.zip"
    staging_dir = target_dir / "staging"

    req = urllib.request.Request(
        download_url,
        headers={"User-Agent": "VoterMail-Engine"},
    )

    h = hashlib.sha256()
    try:
        with urllib.request.urlopen(req) as resp, open(package_zip, "wb") as f_out:
            total_size = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            chunk_size = 65536
            mode_label = "Patch Cepat" if update_mode == "PATCH" else "Paket Penuh"

            while True:
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                f_out.write(chunk)
                h.update(chunk)
                downloaded += len(chunk)

                if progress_callback:
                    if total_size > 0:
                        pct = (downloaded / total_size) * 100
                        mb_dl = downloaded / (1024 * 1024)
                        mb_tot = total_size / (1024 * 1024)
                        msg = f"Mengunduh {mode_label}... {pct:.0f}% ({mb_dl:.1f} / {mb_tot:.1f} MB)".replace(".", ",")
                    else:
                        mb_dl = downloaded / (1024 * 1024)
                        msg = f"Mengunduh {mode_label}... ({mb_dl:.1f} MB)".replace(".", ",")
                    progress_callback(downloaded, total_size, msg)

        # Verifikasi integritas checksum SHA-256
        actual_sha256 = h.hexdigest().lower()
        if expected_sha256 and expected_sha256.strip().lower() != actual_sha256:
            try:
                package_zip.unlink(missing_ok=True)
            except Exception:
                pass
            return False, "File pembaruan tidak lolos pemeriksaan integritas (SHA-256). Pembaruan dibatalkan demi keamanan.", None, None

        if progress_callback:
            progress_callback(total_size, total_size, "Memeriksa dan mengekstrak berkas aman...")

        safe_extract_zip(package_zip, staging_dir)

        # Cari direktori sumber payload
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
            return False, "File VoterMail_Engine.exe tidak ditemukan di dalam paket yang diunduh.", None, None

        # Cadangkan ikon untuk notifikasi floating PowerShell
        src_ico = BASE_DIR / "app" / "assets" / "votermail.ico"
        if src_ico.exists():
            shutil.copy2(src_ico, target_dir / "votermail.ico")

        return True, "Paket pembaruan siap dipasang.", staging_dir, source_dir

    except Exception as exc:
        try:
            if package_zip.exists():
                package_zip.unlink(missing_ok=True)
        except Exception:
            pass
        return False, f"Gagal mengunduh paket pembaruan: {exc}", None, None

def generate_patch_script(target_dir: str | Path) -> Path:
    """Menghasilkan skrip PowerShell mandiri apply_patch.ps1 dalam format ASCII murni dan CRLF."""
    t_dir = Path(target_dir)
    ps_lines = [
        "# VoterMail Engine - Detached Patch Updater & Floating Toast Card",
        "# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)",
        "",
        "param(",
        '    [string]$JobFile = ""',
        ")",
        "",
        '$ErrorActionPreference = "Stop"',
        "",
        "if (-not $JobFile) {",
        '    $JobFile = Join-Path $PSScriptRoot "patch_job.json"',
        "}",
        "",
        "if (-not (Test-Path -LiteralPath $JobFile)) {",
        "    exit 1",
        "}",
        "",
        "$rawJson = [System.IO.File]::ReadAllText($JobFile, [System.Text.Encoding]::UTF8)",
        "$job = $rawJson | ConvertFrom-Json",
        "",
        "Add-Type -AssemblyName System.Windows.Forms",
        "Add-Type -AssemblyName System.Drawing",
        "[System.Windows.Forms.Application]::EnableVisualStyles()",
        "",
        '$isLight = ($job.theme_mode -eq "light")',
        "if ($isLight) {",
        '    $cBg = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")',
        '    $cBorder = [System.Drawing.ColorTranslator]::FromHtml("#CBD5E1")',
        '    $cTextPrimary = [System.Drawing.ColorTranslator]::FromHtml("#0F172A")',
        '    $cTextSecondary = [System.Drawing.ColorTranslator]::FromHtml("#64748B")',
        '    $cBtnSecBg = [System.Drawing.ColorTranslator]::FromHtml("#E2E8F0")',
        '    $cBtnSecText = [System.Drawing.ColorTranslator]::FromHtml("#1E293B")',
        "} else {",
        '    $cBg = [System.Drawing.ColorTranslator]::FromHtml("#0F172A")',
        '    $cBorder = [System.Drawing.ColorTranslator]::FromHtml("#334155")',
        '    $cTextPrimary = [System.Drawing.ColorTranslator]::FromHtml("#F8FAFC")',
        '    $cTextSecondary = [System.Drawing.ColorTranslator]::FromHtml("#94A3B8")',
        '    $cBtnSecBg = [System.Drawing.ColorTranslator]::FromHtml("#334155")',
        '    $cBtnSecText = [System.Drawing.ColorTranslator]::FromHtml("#F8FAFC")',
        "}",
        '$cAccentBlue = [System.Drawing.ColorTranslator]::FromHtml("#2563EB")',
        '$cAccentEmerald = [System.Drawing.ColorTranslator]::FromHtml("#10B981")',
        '$cAccentRed = [System.Drawing.ColorTranslator]::FromHtml("#EF4444")',
        "",
        "$form = New-Object System.Windows.Forms.Form",
        "$form.Width = 370",
        "$form.Height = 146",
        "$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None",
        "$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual",
        "$form.TopMost = $true",
        "$form.ShowInTaskbar = $false",
        "$form.BackColor = $cBg",
        "",
        "$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea",
        "$form.Location = New-Object System.Drawing.Point(($wa.Right - 370 - 16), ($wa.Bottom - 146 - 16))",
        "",
        "$form.add_Paint({",
        "    param($sender, $e)",
        "    $pen = New-Object System.Drawing.Pen($cBorder, 1)",
        "    $e.Graphics.DrawRectangle($pen, 0, 0, ($sender.Width - 1), ($sender.Height - 1))",
        "    $pen.Dispose()",
        "})",
        "",
        "$accentBar = New-Object System.Windows.Forms.Panel",
        "$accentBar.Location = New-Object System.Drawing.Point(0, 0)",
        "$accentBar.Size = New-Object System.Drawing.Size(370, 3)",
        "$accentBar.BackColor = $cAccentBlue",
        "$form.Controls.Add($accentBar)",
        "",
        "$iconBox = New-Object System.Windows.Forms.PictureBox",
        "$iconBox.Location = New-Object System.Drawing.Point(14, 14)",
        "$iconBox.Size = New-Object System.Drawing.Size(20, 20)",
        "$iconBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage",
        "if ($job.icon_path -and (Test-Path -LiteralPath $job.icon_path)) {",
        "    try {",
        "        $ico = New-Object System.Drawing.Icon($job.icon_path)",
        "        $form.Icon = $ico",
        "        $iconBox.Image = $ico.ToBitmap()",
        "    } catch {}",
        "}",
        "$form.Controls.Add($iconBox)",
        "",
        "$lblTitle = New-Object System.Windows.Forms.Label",
        "$lblTitle.Location = New-Object System.Drawing.Point(40, 14)",
        "$lblTitle.Size = New-Object System.Drawing.Size(316, 20)",
        '$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)',
        "$lblTitle.ForeColor = $cTextPrimary",
        '$lblTitle.Text = "Memasang Pembaruan v$($job.target_version)..."',
        "$form.Controls.Add($lblTitle)",
        "",
        "$lblStatus = New-Object System.Windows.Forms.Label",
        "$lblStatus.Location = New-Object System.Drawing.Point(14, 40)",
        "$lblStatus.Size = New-Object System.Drawing.Size(342, 38)",
        '$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)',
        "$lblStatus.ForeColor = $cTextSecondary",
        '$lblStatus.Text = "Menutup VoterMail Engine..."',
        "$form.Controls.Add($lblStatus)",
        "",
        "$progBar = New-Object System.Windows.Forms.ProgressBar",
        "$progBar.Location = New-Object System.Drawing.Point(14, 86)",
        "$progBar.Size = New-Object System.Drawing.Size(342, 8)",
        "$progBar.Value = 20",
        "$progBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous",
        "$form.Controls.Add($progBar)",
        "",
        "$btnPanel = New-Object System.Windows.Forms.Panel",
        "$btnPanel.Location = New-Object System.Drawing.Point(14, 98)",
        "$btnPanel.Size = New-Object System.Drawing.Size(342, 36)",
        "$btnPanel.Visible = $false",
        "",
        "$btnSec = New-Object System.Windows.Forms.Button",
        "$btnSec.Location = New-Object System.Drawing.Point(120, 2)",
        "$btnSec.Size = New-Object System.Drawing.Size(95, 30)",
        '$btnSec.Text = "Nanti Saja"',
        "$btnSec.BackColor = $cBtnSecBg",
        "$btnSec.ForeColor = $cBtnSecText",
        "$btnSec.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat",
        "$btnSec.FlatAppearance.BorderSize = 0",
        '$btnSec.Font = New-Object System.Drawing.Font("Segoe UI", 9)',
        "$btnSec.add_Click({ $form.Close() })",
        "$btnPanel.Controls.Add($btnSec)",
        "",
        "$btnPri = New-Object System.Windows.Forms.Button",
        "$btnPri.Location = New-Object System.Drawing.Point(222, 2)",
        "$btnPri.Size = New-Object System.Drawing.Size(120, 30)",
        '$btnPri.Text = "Buka Aplikasi"',
        "$btnPri.BackColor = $cAccentEmerald",
        "$btnPri.ForeColor = [System.Drawing.Color]::White",
        "$btnPri.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat",
        "$btnPri.FlatAppearance.BorderSize = 0",
        '$btnPri.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)',
        "$btnPri.add_Click({",
        "    try {",
        "        if ($job.is_git) {",
        "            if ($job.python_exe -and $job.main_py) {",
        '                Start-Process -FilePath $job.python_exe -ArgumentList "`"$($job.main_py)`"" -WorkingDirectory $job.app_dir',
        "            } else {",
        '                Start-Process -FilePath "python" -ArgumentList "main.py" -WorkingDirectory $job.app_dir',
        "            }",
        "        } else {",
        "            Start-Process -FilePath $job.exe_path -WorkingDirectory $job.app_dir",
        "        }",
        "    } catch {}",
        "    $form.Close()",
        "})",
        "$btnPanel.Controls.Add($btnPri)",
        "$form.Controls.Add($btnPanel)",
        "",
        "$script:step = 0",
        "$script:waited = 0",
        "$script:backedUpFiles = New-Object System.Collections.Generic.List[string]",
        "",
        "$timer = New-Object System.Windows.Forms.Timer",
        "$timer.Interval = 250",
        "$timer.add_Tick({",
        "    try {",
        "        if ($script:step -eq 0) {",
        "            if ($job.parent_pid -gt 0) {",
        "                $proc = Get-Process -Id $job.parent_pid -ErrorAction SilentlyContinue",
        "                if ($proc -and (-not $proc.HasExited) -and ($script:waited -lt 15000)) {",
        "                    $script:waited += 250",
        "                    return",
        "                }",
        "            }",
        "            $script:step = 1",
        '            $lblStatus.Text = "Mencadangkan dan memasang pembaruan..."',
        "            $progBar.Value = 45",
        "            return",
        "        }",
        "",
        "        if ($script:step -eq 1) {",
        "            $timer.Stop()",
        "            $hasError = $false",
        '            $errorMsg = ""',
        "            try {",
        "                if ($job.is_git) {",
        "                    Set-Location -LiteralPath $job.app_dir",
        "                    $gitOut = & git pull --ff-only origin main 2>&1",
        "                    if ($LASTEXITCODE -ne 0) {",
        '                        throw "Gagal git pull: $gitOut"',
        "                    }",
        "                } else {",
        "                    $srcDir = $job.staging_dir",
        "                    if (-not (Test-Path -LiteralPath $srcDir)) {",
        '                        throw "Folder staging tidak ditemukan: $srcDir"',
        "                    }",
        '                    if (Test-Path -LiteralPath (Join-Path $srcDir "VoterMail_Engine")) {',
        '                        $srcDir = Join-Path $srcDir "VoterMail_Engine"',
        "                    }",
        "                    $bDir = $job.backup_dir",
        "                    if (-not (Test-Path -LiteralPath $bDir)) {",
        "                        New-Item -ItemType Directory -LiteralPath $bDir -Force | Out-Null",
        "                    }",
        "                    $appDir = $job.app_dir",
        "                    $allFiles = Get-ChildItem -LiteralPath $srcDir -Recurse -File",
        "                    foreach ($f in $allFiles) {",
        r'                        $rel = $f.FullName.Substring($srcDir.Length).TrimStart("\", "/")',
        "                        $relLower = $rel.ToLower()",
        r'                        if ($relLower.StartsWith("data\output") -or $relLower.StartsWith("data/output")) { continue }',
        r'                        if ($relLower.StartsWith("logs\") -or $relLower.StartsWith("logs/")) { continue }',
        '                        if ($relLower -eq "config.json") { continue }',
        r'                        if (($relLower.StartsWith("data\templates") -or $relLower.StartsWith("data/templates")) -and (-not $relLower.Contains("sample_"))) { continue }',
        "                        $targetFile = Join-Path $appDir $rel",
        "                        $targetParent = [System.IO.Path]::GetDirectoryName($targetFile)",
        "                        if (Test-Path -LiteralPath $targetFile) {",
        "                            $backupTarget = Join-Path $bDir $rel",
        "                            $bParent = [System.IO.Path]::GetDirectoryName($backupTarget)",
        "                            if (-not (Test-Path -LiteralPath $bParent)) {",
        "                                New-Item -ItemType Directory -LiteralPath $bParent -Force | Out-Null",
        "                            }",
        "                            Copy-Item -LiteralPath $targetFile -Destination $backupTarget -Force",
        "                            $script:backedUpFiles.Add($rel)",
        "                        }",
        "                        if (-not (Test-Path -LiteralPath $targetParent)) {",
        "                            New-Item -ItemType Directory -LiteralPath $targetParent -Force | Out-Null",
        "                        }",
        "                        Copy-Item -LiteralPath $f.FullName -Destination $targetFile -Force",
        "                    }",
        "                }",
        "            } catch {",
        "                $hasError = $true",
        "                $errorMsg = $_.Exception.Message",
        "            }",
        "",
        "            if ($hasError) {",
        "                try {",
        "                    $bDir = $job.backup_dir",
        "                    foreach ($rel in $script:backedUpFiles) {",
        "                        $bSrc = Join-Path $bDir $rel",
        "                        $tDst = Join-Path $job.app_dir $rel",
        "                        if (Test-Path -LiteralPath $bSrc) {",
        "                            Copy-Item -LiteralPath $bSrc -Destination $tDst -Force",
        "                        }",
        "                    }",
        "                } catch {}",
        "                $accentBar.BackColor = $cAccentRed",
        '                $lblTitle.Text = "Pembaruan Gagal"',
        '                $lblStatus.Text = "Terjadi kendala saat memasang. Berkas lama telah dipulihkan secara aman."',
        "                $progBar.Visible = $false",
        "                $btnPanel.Visible = $true",
        "                $btnSec.Visible = $false",
        '                $btnPri.Text = "Tutup"',
        "                $btnPri.BackColor = $cBtnSecBg",
        "                $btnPri.ForeColor = $cBtnSecText",
        "                $btnPri.Location = New-Object System.Drawing.Point(222, 2)",
        "                return",
        "            }",
        "",
        "            try {",
        "                if (Test-Path -LiteralPath $job.backup_dir) {",
        "                    Remove-Item -LiteralPath $job.backup_dir -Recurse -Force -ErrorAction SilentlyContinue",
        "                }",
        "                if (Test-Path -LiteralPath $job.staging_dir) {",
        "                    Remove-Item -LiteralPath $job.staging_dir -Recurse -Force -ErrorAction SilentlyContinue",
        "                }",
        '                $pkgZip = Join-Path $PSScriptRoot "package.zip"',
        "                if (Test-Path -LiteralPath $pkgZip) {",
        "                    Remove-Item -LiteralPath $pkgZip -Force -ErrorAction SilentlyContinue",
        "                }",
        "            } catch {}",
        "",
        "            try {",
        "                [System.Media.SystemSounds]::Asterisk.Play()",
        "            } catch {}",
        "",
        "            $accentBar.BackColor = $cAccentEmerald",
        "            $bullet = [char]0x2022",
        '            $lblTitle.Text = "Pembaruan Selesai $bullet v$($job.target_version)"',
        '            $lblStatus.Text = "VoterMail Engine siap digunakan kembali. Semua file data pemilih Anda tetap aman."',
        "            $progBar.Visible = $false",
        "            $btnPanel.Visible = $true",
        "            return",
        "        }",
        "    } catch {",
        "        $timer.Stop()",
        "        $accentBar.BackColor = $cAccentRed",
        '        $lblTitle.Text = "Pembaruan Gagal"',
        '        $lblStatus.Text = "Pemasangan dihentikan: $($_.Exception.Message)"',
        "        $progBar.Visible = $false",
        "        $btnPanel.Visible = $true",
        "    }",
        "})",
        "",
        "$form.add_Load({",
        "    $timer.Start()",
        "})",
        "",
        "[System.Windows.Forms.Application]::Run($form)",
        "",
    ]
    ps1_path = t_dir / "apply_patch.ps1"
    content = "\r\n".join(ps_lines)
    ps1_path.write_bytes(content.encode("ascii"))
    return ps1_path

def launch_detached_patcher(
    parent_pid: int,
    app_dir: str | Path,
    exe_path: str | Path,
    staging_dir: str | Path,
    backup_dir: str | Path,
    update_mode: str,
    target_version: str,
    icon_path: str | Path | None = None,
    theme_mode: str = "dark",
    is_git: bool = False,
    python_exe: str = "",
    main_py: str = "",
) -> tuple[bool, str]:
    """Menulis patch_job.json dan meluncurkan PowerShell floating toast card secara terpisah."""
    t_dir = Path(staging_dir).parent
    job_path = t_dir / "patch_job.json"

    ico_str = str(icon_path) if icon_path and Path(icon_path).exists() else ""
    if not ico_str:
        cand_ico = t_dir / "votermail.ico"
        if cand_ico.exists():
            ico_str = str(cand_ico)

    job_data = {
        "parent_pid": int(parent_pid),
        "app_dir": str(Path(app_dir).resolve()),
        "exe_path": str(Path(exe_path).resolve()),
        "staging_dir": str(Path(staging_dir).resolve()),
        "backup_dir": str(Path(backup_dir).resolve()),
        "update_mode": str(update_mode).upper(),
        "target_version": str(target_version),
        "icon_path": ico_str,
        "theme_mode": str(theme_mode).lower(),
        "is_git": bool(is_git),
        "python_exe": str(python_exe or sys.executable),
        "main_py": str(main_py or (BASE_DIR / "main.py")),
    }

    job_path.write_text(json.dumps(job_data, indent=2), encoding="utf-8")
    ps1_path = generate_patch_script(t_dir)

    CREATE_NO_WINDOW = 0x08000000
    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200

    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        str(ps1_path),
        "-JobFile",
        str(job_path),
    ]

    try:
        subprocess.Popen(
            cmd,
            creationflags=CREATE_NO_WINDOW | DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
            close_fds=True,
        )
        return True, "Proses pemasangan pembaruan berhasil diluncurkan."
    except Exception as exc:
        return False, f"Gagal meluncurkan patcher PowerShell: {exc}"

def download_and_apply_update(
    download_url: str,
    html_url: str = "",
    progress_callback: Callable[[int, int, str], None] | None = None,
    expected_sha256: str = "",
    update_mode: str = "FULL",
    target_version: str = "",
) -> tuple[bool, str]:
    """Mengunduh dan menerapkan pembaruan (kompatibilitas terpadu)."""
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

    t_ver = target_version or "latest"
    ok, msg, staging_dir, source_dir = stage_update_package(
        download_url=download_url,
        target_version=t_ver,
        expected_sha256=expected_sha256,
        update_mode=update_mode,
        progress_callback=progress_callback,
    )
    if not ok or not staging_dir or not source_dir:
        return False, msg

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
        target_version=t_ver,
        icon_path=staging_dir.parent / "votermail.ico",
        theme_mode="dark",
        is_git=False,
    )
    if not launch_ok:
        return False, launch_msg

    sys.exit(0)
