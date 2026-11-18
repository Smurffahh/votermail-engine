# VoterMail Engine — High-Volume Word & Excel DPT Mail-Merge Automation
# Created by Danang Satria (https://github.com/Smurffahh/votermail-engine)

from pathlib import Path
import sys
from PIL import Image, ImageDraw

def generate_votermail_icon(output_dir: Path | None = None) -> tuple[Path, Path]:
    if output_dir is None:
        base_dir = Path(__file__).resolve().parent.parent
        output_dir = base_dir / "app" / "assets"
    
    output_dir.mkdir(parents=True, exist_ok=True)
    ico_path = output_dir / "votermail.ico"
    png_path = output_dir / "votermail.png"

    # 1. Canvas 1024x1024 Supersampling
    size = 1024
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 2. Base Squircle / Rounded Rectangle
    pad = 64
    r_base = 200
    base_box = [pad, pad, size - pad, size - pad]

    # Create vertical gradient for base squircle (#0F172A to #1E293B)
    grad = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    g_draw = ImageDraw.Draw(grad)
    c_top = (15, 23, 42, 255)     # #0F172A
    c_bot = (30, 41, 59, 255)     # #1E293B
    for y in range(pad, size - pad):
        factor = (y - pad) / (size - 2 * pad)
        r = int(c_top[0] + factor * (c_bot[0] - c_top[0]))
        g = int(c_top[1] + factor * (c_bot[1] - c_top[1]))
        b = int(c_top[2] + factor * (c_bot[2] - c_top[2]))
        g_draw.line([(pad, y), (size - pad, y)], fill=(r, g, b, 255))

    mask = Image.new("L", (size, size), 0)
    m_draw = ImageDraw.Draw(mask)
    m_draw.rounded_rectangle(base_box, radius=r_base, fill=255)
    img.paste(grad, (0, 0), mask)

    # Subtle inner border (#334155, 12px)
    draw.rounded_rectangle(base_box, radius=r_base, outline=(51, 65, 85, 255), width=12)

    # 3. Lembar Dokumen Putih Bersih (#F8FAFC)
    doc_x0, doc_y0 = 300, 190
    doc_x1, doc_y1 = 724, 620
    doc_r = 32
    draw.rounded_rectangle([doc_x0, doc_y0, doc_x1, doc_y1], radius=doc_r, fill=(248, 250, 252, 255))
    draw.rounded_rectangle([doc_x0, doc_y0, doc_x1, doc_y1], radius=doc_r, outline=(226, 232, 240, 255), width=6)

    # Garis Teks Simbolis Dokumen
    # Header line (#3B82F6)
    draw.rounded_rectangle([350, 260, 560, 288], radius=14, fill=(59, 130, 246, 255))
    # Line 1 (#94A3B8)
    draw.rounded_rectangle([350, 320, 674, 344], radius=12, fill=(148, 163, 184, 255))
    # Line 2 (#CBD5E1)
    draw.rounded_rectangle([350, 376, 614, 400], radius=12, fill=(203, 213, 225, 255))

    # 4. Badan Amplop Surat Modern (Geometric Folds)
    env_x0, env_y0 = 220, 440
    env_x1, env_y1 = 804, 780
    env_r = 40
    center_x = (env_x0 + env_x1) // 2
    fold_y = 630

    # Amplop base frame
    env_mask = Image.new("L", (size, size), 0)
    em_draw = ImageDraw.Draw(env_mask)
    em_draw.rounded_rectangle([env_x0, env_y0, env_x1, env_y1], radius=env_r, fill=255)

    env_layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    el_draw = ImageDraw.Draw(env_layer)
    el_draw.rectangle([env_x0, env_y0, env_x1, env_y1], fill=(29, 78, 216, 255)) # #1D4ED8

    # Left flap (#2563EB)
    el_draw.polygon([(env_x0, env_y0), (center_x, fold_y), (env_x0, env_y1)], fill=(37, 99, 235, 255))
    # Right flap (#1E40AF)
    el_draw.polygon([(env_x1, env_y0), (center_x, fold_y), (env_x1, env_y1)], fill=(30, 64, 175, 255))
    # Bottom flap (#1D4ED8)
    el_draw.polygon([(env_x0, env_y1), (center_x, fold_y - 20), (env_x1, env_y1)], fill=(29, 78, 216, 255))
    # Top rim highlight (#60A5FA, 8px)
    el_draw.line([(env_x0, env_y0), (center_x, fold_y)], fill=(96, 165, 250, 255), width=8)
    el_draw.line([(center_x, fold_y), (env_x1, env_y0)], fill=(96, 165, 250, 255), width=8)

    img.paste(env_layer, (0, 0), env_mask)
    draw.rounded_rectangle([env_x0, env_y0, env_x1, env_y1], radius=env_r, outline=(30, 58, 138, 255), width=6)

    # 5. Badge Verifikasi Pemilih (Emerald Green Circle + White Checkmark)
    b_cx, b_cy = 740, 720
    b_rad = 115
    badge_box = [b_cx - b_rad, b_cy - b_rad, b_cx + b_rad, b_cy + b_rad]

    # Cutout border (#0F172A, width 20px) to separate from envelope
    draw.ellipse([b_cx - b_rad - 10, b_cy - b_rad - 10, b_cx + b_rad + 10, b_cy + b_rad + 10], fill=(15, 23, 42, 255))

    # Badge Gradient Fill (#34D399 to #059669)
    b_grad = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    bg_draw = ImageDraw.Draw(b_grad)
    c_bg_top = (52, 211, 153, 255) # #34D399
    c_bg_bot = (5, 150, 105, 255)  # #059669
    for y in range(badge_box[1], badge_box[3]):
        f = (y - badge_box[1]) / (2 * b_rad)
        r = int(c_bg_top[0] + f * (c_bg_bot[0] - c_bg_top[0]))
        g = int(c_bg_top[1] + f * (c_bg_bot[1] - c_bg_top[1]))
        b = int(c_bg_top[2] + f * (c_bg_bot[2] - c_bg_top[2]))
        bg_draw.line([(badge_box[0], y), (badge_box[2], y)], fill=(r, g, b, 255))

    b_mask = Image.new("L", (size, size), 0)
    bm_draw = ImageDraw.Draw(b_mask)
    bm_draw.ellipse(badge_box, fill=255)
    img.paste(b_grad, (0, 0), b_mask)
    draw.ellipse(badge_box, outline=(16, 185, 129, 255), width=6)

    # White Checkmark
    chk_start = (b_cx - 50, b_cy + 5)
    chk_corner = (b_cx - 15, b_cy + 42)
    chk_end = (b_cx + 52, b_cy - 40)
    draw.line([chk_start, chk_corner], fill=(255, 255, 255, 255), width=28)
    draw.line([chk_corner, chk_end], fill=(255, 255, 255, 255), width=28)
    # Smooth line caps
    draw.ellipse([chk_start[0] - 14, chk_start[1] - 14, chk_start[0] + 14, chk_start[1] + 14], fill=(255, 255, 255, 255))
    draw.ellipse([chk_corner[0] - 14, chk_corner[1] - 14, chk_corner[0] + 14, chk_corner[1] + 14], fill=(255, 255, 255, 255))
    draw.ellipse([chk_end[0] - 14, chk_end[1] - 14, chk_end[0] + 14, chk_end[1] + 14], fill=(255, 255, 255, 255))

    # 6. Export PNG (256x256)
    png_img = img.resize((256, 256), Image.Resampling.LANCZOS)
    png_img.save(png_path, format="PNG")

    # 7. Export Multi-Resolution ICO
    ico_sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(ico_path, format="ICO", sizes=ico_sizes)

    return ico_path, png_path

if __name__ == "__main__":
    ico, png = generate_votermail_icon()
    print(f"[✓] Ikon berhasil dibuat:")
    print(f"    - ICO: {ico} ({ico.stat().st_size:,} bytes)")
    print(f"    - PNG: {png} ({png.stat().st_size:,} bytes)")
