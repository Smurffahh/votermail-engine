# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import argparse
from datetime import date
import importlib
from pathlib import Path
import re
import sys

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.config.version import (
    __version__ as current_version,
    parse_version_tuple,
    validate_version_registry,
)

BASE_DIR = PROJECT_ROOT
VERSION_FILE = BASE_DIR / "app" / "config" / "version.py"
README_FILE = BASE_DIR / "README.md"

def calculate_next_version(current: str, bump_type: str) -> str:
    parts = parse_version_tuple(current)
    major = parts[0] if len(parts) > 0 else 0
    minor = parts[1] if len(parts) > 1 else 0
    patch = parts[2] if len(parts) > 2 else 0

    if bump_type == "major":
        return f"{major + 1}.0.0"
    if bump_type == "minor":
        return f"{major}.{minor + 1}.0"
    if bump_type == "patch":
        return f"{major}.{minor}.{patch + 1}"
    raise ValueError(f"Tipe bump tidak valid: '{bump_type}'. Pilih 'major', 'minor', atau 'patch'.")

def bump_version(
    new_version: str,
    title: str,
    description: str,
    release_date: str | None = None,
    highlights: list[str] | None = None,
) -> None:
    validate_version_registry()

    if not re.match(r"^\d+\.\d+\.\d+$", new_version):
        raise ValueError(
            f"Format versi tidak valid: '{new_version}'. Wajib mengikuti format SemVer mayor.minor.patch (contoh: 1.11.0)."
        )

    curr_tuple = parse_version_tuple(current_version)
    new_tuple = parse_version_tuple(new_version)

    if new_tuple <= curr_tuple:
        raise ValueError(
            f"Regresi versi ditolak: Versi baru 'v{new_version}' harus lebih besar secara monotonik dari versi saat ini 'v{current_version}'."
        )

    today_str = release_date or date.today().isoformat()
    content = VERSION_FILE.read_text(encoding="utf-8")

    registry_pattern = r"(FEATURE_REGISTRY:\s*list\[dict\[str,\s*(?:str|Any)\]\]\s*=\s*\[)([\s\S]*?)(\n\])"
    match = re.search(registry_pattern, content)
    if not match:
        raise RuntimeError("FEATURE_REGISTRY tidak ditemukan di app/config/version.py")

    entry_lines = [
        "    {",
        f'        "version": "{new_version}",',
        f'        "date": "{today_str}",',
        f'        "title": "{title}",',
        f'        "description": "{description}",',
    ]
    if highlights:
        entry_lines.append('        "highlights": [')
        for h in highlights:
            escaped_h = h.replace('"', '\\"')
            entry_lines.append(f'            "{escaped_h}",')
        entry_lines.append("        ],")
    entry_lines.append("    },")
    new_entry = "\n".join(entry_lines) + "\n"

    updated_registry = match.group(1) + match.group(2) + new_entry + "]"
    updated_content = content[:match.start()] + updated_registry + content[match.end():]
    VERSION_FILE.write_text(updated_content, encoding="utf-8")
    print(f"✓ Berhasil menambahkan v{new_version} ke {VERSION_FILE.name}")

    if README_FILE.exists():
        readme_txt = README_FILE.read_text(encoding="utf-8")
        badge_pat = r"(badge/version-)([\d\.]+)(-[^)]+)"
        if re.search(badge_pat, readme_txt):
            readme_txt = re.sub(badge_pat, rf"\g<1>{new_version}\g<3>", readme_txt)
        zip_pat = r"(VoterMail_Engine_v)[\d\.]*(_Windows_x64_Portable\.zip)"
        if re.search(zip_pat, readme_txt):
            readme_txt = re.sub(zip_pat, rf"\g<1>{new_version}\g<2>", readme_txt)
        README_FILE.write_text(readme_txt, encoding="utf-8")
        print(f"✓ Berhasil menyinkronkan versi di {README_FILE.name} menjadi v{new_version}")

    import app.config.version as ver_mod
    importlib.reload(ver_mod)
    ver_mod.validate_version_registry()
    print(f"✓ Validasi integritas Monotonic Version Guardrail berhasil untuk v{new_version}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="VoterMail Engine — Semantic Feature Version Bumper")
    parser.add_argument("--version", type=str, default=None, help="Nomor versi baru (contoh: 1.11.0)")
    parser.add_argument("--bump", type=str, choices=["patch", "minor", "major"], default=None, help="Tipe kenaikan versi otomatis")
    parser.add_argument("--title", type=str, required=True, help="Judul singkat rilis fitur")
    parser.add_argument("--desc", type=str, required=True, help="Deskripsi penambahan fitur")
    parser.add_argument("--date", type=str, default=None, help="Tanggal rilis YYYY-MM-DD (opsional)")
    parser.add_argument("--highlight", action="append", default=[], help="Poin sorotan rilis (dapat dipanggil berulang)")

    args = parser.parse_args()

    target_version = args.version
    if not target_version:
        if args.bump:
            target_version = calculate_next_version(current_version, args.bump)
        else:
            print("Error: Salah satu dari --version atau --bump wajib disertakan.", file=sys.stderr)
            sys.exit(1)

    try:
        bump_version(target_version, args.title, args.desc, args.date, args.highlight)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
