"""Photo handling with Pillow — replaces v1's sharp.

- watermark_and_save: stamps the serial (O12 / I7) top-right on a real photo.
- generate_receipt_card: renders the black & gold "Entered manually" card for
  photo-less saves, so every row still links to a consistent image.
"""
from __future__ import annotations

import io
import os
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps

GOLD = "#c9a24b"
GOLD_LT = "#e7c873"
INK = "#f4efe6"
DIM = "#8c8478"
BG = "#0c0c0d"
LINE = "#3a342a"

_FONT_CANDIDATES = [
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]


def _font(size: int, bold: bool = True) -> ImageFont.FreeTypeFont:
    for p in _FONT_CANDIDATES:
        if ("bd" in p or "Bold" in p) != bold:
            continue
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    for p in _FONT_CANDIDATES:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def watermark_and_save(image_bytes: bytes, label: str, out_path: Path) -> None:
    img = Image.open(io.BytesIO(image_bytes))
    img = ImageOps.exif_transpose(img)  # same as sharp's .rotate()
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    font_size = max(36, round(min(w, h) * 0.06))
    font = _font(font_size, bold=True)
    pad_x = round(font_size * 0.45)
    pad_y = round(font_size * 0.25)
    bbox = font.getbbox(label)
    text_w = bbox[2] - bbox[0]
    box_w = text_w + pad_x * 2
    box_h = font_size + pad_y * 2
    offset = round(min(w, h) * 0.025)

    overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    dr = ImageDraw.Draw(overlay)
    x0 = w - offset - box_w
    y0 = offset
    dr.rounded_rectangle(
        [x0, y0, x0 + box_w, y0 + box_h], radius=pad_y, fill=(0, 0, 0, 199)
    )
    dr.text((x0 + pad_x, y0 + pad_y - bbox[1] * 0), label, font=font, fill="white")
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    img.save(out_path, "JPEG", quality=90)


def generate_receipt_card(
    label: str, fields: list[tuple[str, str]], out_path: Path
) -> None:
    W, H = 900, 1200
    img = Image.new("RGB", (W, H), BG)
    dr = ImageDraw.Draw(img)
    dr.rounded_rectangle([24, 24, W - 24, H - 24], radius=36, outline=LINE, width=2)

    def centered(y, text, font, fill, tracking=0):
        # crude letter-spacing: insert hair spaces
        if tracking:
            text = (" " * tracking).join(list(text))
        bbox = dr.textbbox((0, 0), text, font=font)
        dr.text(((W - (bbox[2] - bbox[0])) / 2, y), text, font=font, fill=fill)

    centered(118, "SHEILA McGAVIGAN", _font(32, True), GOLD_LT, tracking=1)
    centered(172, "PERMANENT MAKEUP", _font(20, False), DIM, tracking=1)
    dr.line([70, 250, W - 70, 250], fill=GOLD, width=2)
    dr.text((70, 288), "ENTERED MANUALLY", font=_font(30, True), fill=GOLD)

    y = 360
    row_h = 92
    f_key = _font(22, False)
    f_val = _font(38, True)
    for k, v in fields:
        if v in (None, ""):
            continue
        dr.text((70, y - 26), str(k).upper(), font=f_key, fill=DIM)
        dr.text((70, y + 4), str(v), font=f_val, fill=INK)
        dr.line([70, y + 58, W - 70, y + 58], fill=LINE, width=1)
        y += row_h
        if y > H - 200:
            break

    f_label = _font(64, True)
    bbox = dr.textbbox((0, 0), label, font=f_label)
    dr.text((W - 70 - (bbox[2] - bbox[0]), H - 70 - (bbox[3] - bbox[1])), label,
            font=f_label, fill=GOLD)
    img.save(out_path, "JPEG", quality=92)


def finalize_pending_photo(photos_dir: Path, photo_name: str) -> None:
    """Idempotently promote a transaction's pending photo into its final name."""
    final = photos_dir / photo_name
    pending = photos_dir / f"{photo_name}.pending"
    if final.is_file():
        pending.unlink(missing_ok=True)
        return
    try:
        os.replace(pending, final)
    except FileNotFoundError:
        # Another retry may have completed the same promotion concurrently.
        if not final.is_file():
            raise


def recover_pending_photos(db_conn, photos_dir: Path, staging_dir: Path) -> None:
    """Complete committed photos and remove files from rolled-back/crashed saves."""
    photos_dir.mkdir(parents=True, exist_ok=True)
    for pending in photos_dir.glob("*.jpg.pending"):
        photo_name = pending.name[:-len(".pending")]
        row = db_conn.execute(
            "SELECT 1 FROM receipts WHERE photo = ? LIMIT 1", (photo_name,)
        ).fetchone()
        if row:
            finalize_pending_photo(photos_dir, photo_name)
        else:
            pending.unlink(missing_ok=True)
    if staging_dir.is_dir():
        for staged in staging_dir.iterdir():
            if staged.is_file():
                staged.unlink(missing_ok=True)
