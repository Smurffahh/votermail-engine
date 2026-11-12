# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

__author__ = "Danang Satria"

import argparse
from datetime import date
from pathlib import Path
import re
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
VERSION_FILE = BASE_DIR / "app" / "config" / "version.py"
README_FILE = BASE_DIR / "README.md"

def bump_version(new_version: str, title: str, description: str, release_date: str | None = None) -> None:
    if not re.match(r"^\d+\.\d+\.\d+$", new_version):
        raise ValueError(f"Format versi tidak valid: '{new_version}'. Gunakan format SemVer (misal: 1.5.0).")

    today_str = release_date or date.today().isoformat()

    content = VERSION_FILE.read_text(encoding="utf-8")

    registry_pattern = r"(FEATURE_REGISTRY:\s*list\[dict\[str,\s*str\]\]\s*=\s*\[)([\s\S]*?)(\n\])"
    match = re.search(registry_pattern, content)
    if not match:
        raise RuntimeError("FEATURE_REGISTRY tidak ditemukan di app/config/version.py")

    new_entry = (
        f"    {{\n"
        f'        "version": "{new_version}",\n'
        f'        "date": "{today_str}",\n'
        f'        "title": "{title}",\n'
        f'        "description": "{description}",\n'
        f"    }},\n"
    )

    updated_registry = match.group(1) + match.group(2) + new_entry + "]"
    updated_content = content[:match.start()] + updated_registry + content[match.end():]
    VERSION_FILE.write_text(updated_content, encoding="utf-8")
    print(f"✓ Berhasil menambahkan v{new_version} ke {VERSION_FILE.name}")

    if README_FILE.exists():
        readme_txt = README_FILE.read_text(encoding="utf-8")
        badge_pat = r"(badge/version-)([\d\.]+)(-blue)"
        if re.search(badge_pat, readme_txt):
            readme_txt = re.sub(badge_pat, rf"\g<1>{new_version}\g<3>", readme_txt)
        header_pat = r"(# VoterMail Engine)(\s+v[\d\.]+)?"
        readme_txt = re.sub(header_pat, rf"\g<1> v{new_version}", readme_txt, count=1)
        README_FILE.write_text(readme_txt, encoding="utf-8")
        print(f"✓ Berhasil memperbarui versi di {README_FILE.name} menjadi v{new_version}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="VoterMail Engine — Semantic Feature Version Bumper")
    parser.add_argument("--version", type=str, required=True, help="Nomor versi baru (contoh: 1.5.0)")
    parser.add_argument("--title", type=str, required=True, help="Judul singkat rilis fitur")
    parser.add_argument("--desc", type=str, required=True, help="Deskripsi penambahan fitur")
    parser.add_argument("--date", type=str, default=None, help="Tanggal rilis YYYY-MM-DD (opsional)")

    args = parser.parse_args()
    try:
        bump_version(args.version, args.title, args.desc, args.date)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
