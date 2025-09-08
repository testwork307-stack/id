# -*- coding: utf-8 -*-
"""Streamlit – HR ID Generator (Arabic-aware, ZIP/RAR, robust)
- Name is bold and nudged left/up
- "ID:" label changed to "الرقم الوظيفي:"
- Spacing: job under name (+10), employee number under job (+15)
"""

import os, io, shutil, zipfile, tempfile
from pathlib import Path
import cv2
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter

# Arabic text handling
import arabic_reshaper
from bidi.algorithm import get_display

# ===================== CONFIG =====================
PHOTO_POS = (111, 168)
PHOTO_SIZE = (300, 300)
BARCODE_POS = (570, 465)
BARCODE_SIZE = (390, 120)

NAME_OFFSET_X = -40   # negative = left
NAME_OFFSET_Y = -20   # negative = up

# ===================== UI =========================
st.set_page_config(page_title="HR ID Card Generator", page_icon="🎫", layout="wide")
st.title("🎫 HR ID Card Generator")

with st.sidebar:
    st.markdown("**Tips**")
    st.markdown("- Excel columns: **الاسم**, **الوظيفة**, **الرقم**, **الرقم القومي**, **الصورة**.")
    st.markdown("- Archive can have nested folders; the app searches recursively.")
    st.markdown("- For best Arabic rendering, upload proper TTF fonts.")

# Optional custom fonts
font_ar_file = st.sidebar.file_uploader("Arabic font (TTF/OTF, e.g., Amiri)", type=["ttf", "otf"], key="ar_font")
font_en_file = st.sidebar.file_uploader("English font (TTF/OTF)", type=["ttf", "otf"], key="en_font")

# Optional override for unrar tool on Windows
custom_unrar = st.sidebar.text_input("Path to unrar.exe (if needed)")

excel_file = st.file_uploader("📂 Upload Excel (.xlsx)", type=["xlsx"], key="xlsx")
photos_archive = st.file_uploader("📦 Upload Photos (ZIP or RAR)", type=["zip", "rar"], key="archive")
template_file = st.file_uploader("🖼 Upload Card Template (PNG/JPG)", type=["png", "jpg", "jpeg"], key="tpl")

# ================== Helpers =======================
def load_font_from_upload(upload, fallback_name: str, size: int):
    """Load a font from an uploaded file; otherwise try common local fonts; otherwise PIL default."""
    if upload is not None:
        try:
            return ImageFont.truetype(io.BytesIO(upload.read()), size)
        except Exception:
            st.warning(f"⚠️ Failed to load uploaded font for {fallback_name}. Falling back to default.")
    # Fallbacks
    for candidate in [
        "Amiri-Regular.ttf", "Amiri.ttf", "NotoNaskhArabic-Regular.ttf",
        "HacenMaghreb.ttf", "HacenMaghreb (1).ttf",
        "Arial.ttf", "Tahoma.ttf",
    ]:
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            continue
    return ImageFont.load_default()

def prepare_text(text: str) -> str:
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)

def draw_aligned_text(draw, xy, text, font, fill="black", anchor="rt"):
    if not text:
        return
    draw.text(xy, str(text), font=font, fill=fill, anchor=anchor)

def draw_bold_text(draw, xy, text, font, fill="black", anchor="rt"):
    for dx, dy in [(0,0), (1,0), (0,1), (1,1)]:
        draw_aligned_text(draw, (xy[0]+dx, xy[1]+dy), text, font, fill=fill, anchor=anchor)

def find_photo_path(root_dir: str, requested: str):
    """Search recursively for requested photo (ignores folder structure, case-insensitive)."""
    if not requested:
        return None
    requested = str(requested).strip().lower()
    req_stem = Path(requested).stem.lower()

    for dirpath, _, filenames in os.walk(root_dir):
        for fn in filenames:
            if Path(fn).stem.lower() == req_stem:  # match name without extension
                return os.path.join(dirpath, fn)
    return None

def crop_face_and_shoulders(image_path: str):
    img = cv2.imread(image_path)
    if img is None:
        return None
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    faces = face_cascade.detectMultiScale(gray, 1.1, 5)
    if len(faces) == 0:
        return None
    x, y, w, h = faces[0]
    y_start = max(0, y - int(0.3 * h))
    y_end   = min(img.shape[0], y + int(2.0 * h))
    x_start = max(0, x - int(0.3 * w))
    x_end   = min(img.shape[1], x + int(1.3 * w))
    cropped = img[y_start:y_end, x_start:x_end]
    return Image.fromarray(cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB))

def ensure_rar_support(custom_path: str | None = None):
    import rarfile
    if custom_path and Path(custom_path).exists():
        rarfile.UNRAR_TOOL = custom_path
        return True
    return True

# ================== Main logic ====================
if excel_file and photos_archive and template_file:
    font_ar = load_font_from_upload(font_ar_file, "Arabic", 36)
    font_en = load_font_from_upload(font_en_file, "English", 30)

    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        st.stop()

    try:
        template = Image.open(template_file).convert("RGB")
    except Exception as e:
        st.error(f"❌ Failed to read template: {e}")
        st.stop()

    tmpdir = tempfile.mkdtemp(prefix="idcards_")
    archive_path = os.path.join(tmpdir, photos_archive.name)
    with open(archive_path, "wb") as f:
        f.write(photos_archive.getbuffer())

    # Extract photos
    try:
        if archive_path.lower().endswith(".zip"):
            with zipfile.ZipFile(archive_path, "r") as zf:
                zf.extractall(tmpdir)
        elif archive_path.lower().endswith(".rar"):
            import rarfile
            if ensure_rar_support(custom_unrar):
                with rarfile.RarFile(archive_path, "r") as rf:
                    rf.extractall(tmpdir)
            else:
                st.error("❌ RAR not supported here.")
                shutil.rmtree(tmpdir, ignore_errors=True)
                st.stop()
        else:
            st.error("❌ Unsupported archive type.")
            shutil.rmtree(tmpdir, ignore_errors=True)
            st.stop()
    except Exception as e:
        st.error(f"❌ Failed to extract archive: {e}")
        shutil.rmtree(tmpdir, ignore_errors=True)
        st.stop()

    output_cards = []
    for _, row in df.iterrows():
        card = template.copy()
        draw = ImageDraw.Draw(card)

        name = prepare_text(str(row.get("الاسم", "")).strip())
        job  = prepare_text(str(row.get("الوظيفة", "")).strip())
        num  = str(row.get("الرقم", "")).strip()
        national_id = str(row.get("الرقم القومي", "")).strip()
        photo_filename = str(row.get("الصورة", "")).strip()

        # Text placement
        base_name_xy = (915, 240)
        name_xy = (base_name_xy[0] + NAME_OFFSET_X, base_name_xy[1] + NAME_OFFSET_Y)
        draw_bold_text(draw, name_xy, name, font_ar, fill="black", anchor="rt")

        name_bbox   = draw.textbbox((0, 0), name, font=font_ar)
        name_height = (name_bbox[3] - name_bbox[1]) + 20
        job_xy = (name_xy[0], name_xy[1] + name_height)
        draw_aligned_text(draw, job_xy, job, font=font_ar, fill="black", anchor="rt")

        job_bbox   = draw.textbbox((0, 0), job, font=font_ar)
        job_height = (job_bbox[3] - job_bbox[1]) + 25
        id_xy = (name_xy[0], job_xy[1] + job_height)
        job_id_label = prepare_text(f"الرقم الوظيفي: {num}")
        draw_aligned_text(draw, id_xy, job_id_label, font=font_ar, fill="black", anchor="rt")

        # Photo
        photo_path = find_photo_path(tmpdir, photo_filename)
        if photo_path:
            try:
                cropped = crop_face_and_shoulders(photo_path)
                img = cropped if cropped else Image.open(photo_path)
                img = img.convert("RGB").resize(PHOTO_SIZE)
                card.paste(img, PHOTO_POS)
            except Exception as e:
                st.warning(f"⚠️ Photo issue for {row.get('الاسم', '')}: {e}")
        else:
            st.warning(f"📷 Photo not found for {row.get('الاسم', '')} ({photo_filename})")

        # Barcode
        if national_id:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_barcode:
                out_noext = tmp_barcode.name[:-4]
            barcode = Code128(national_id, writer=ImageWriter())
            barcode_path = barcode.save(out_noext, {"write_text": False})
            with Image.open(barcode_path) as bimg:
                bimg = bimg.convert("RGB").resize(BARCODE_SIZE)
                card.paste(bimg, BARCODE_POS)
            for p in [out_noext + ".png", out_noext + ".svg"]:
                try: os.remove(p)
                except: pass

        output_cards.append(card)

    # Export PDF
    if output_cards:
        pdf_path = os.path.join(tmpdir, "All_ID_Cards.pdf")
        output_cards[0].save(pdf_path, save_all=True, append_images=output_cards[1:])
        with open(pdf_path, "rb") as f:
            st.download_button("⬇️ Download All ID Cards (PDF)", f, file_name="All_ID_Cards.pdf")
        st.success(f"✅ Generated {len(output_cards)} cards")
        st.image(output_cards[0], caption="Preview", width=320)
    else:
        st.warning("No cards generated.")

else:
    st.info("👆 Upload Excel, Photos archive, and Template to start.")
