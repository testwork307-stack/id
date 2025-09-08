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

# Fine-tune these two to move the *name* relative to original point:
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
    # Fallbacks – try common installed fonts; finally PIL default
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
    """Arabic reshape + bidi for correct display."""
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)

def draw_aligned_text(draw: ImageDraw.ImageDraw, xy, text, font, fill="black", anchor="rt"):
    """Anchored text; multi-line supported line-by-line."""
    if not text:
        return
    lines = str(text).split("\n")
    x, y = xy
    for i, line in enumerate(lines):
        if i > 0:
            bbox = draw.textbbox((0, 0), line, font=font)
            y += (bbox[3] - bbox[1])
        draw.text((x, y), line, font=font, fill=fill, anchor=anchor)

def draw_bold_text(draw, xy, text, font, fill="black", anchor="rt"):
    """Fake-bold by layering 1px offsets (PIL-friendly)."""
    for dx, dy in [(0,0), (1,0), (0,1), (1,1)]:
        draw_aligned_text(draw, (xy[0]+dx, xy[1]+dy), text, font, fill=fill, anchor=anchor)

def find_photo_path(root_dir: str, requested: str):
    """Find photo by stem match (case/ext-insensitive), search recursively."""
    if not requested:
        return None
    requested = str(requested).strip().lower()
    req_stem = Path(requested).stem.lower()

    for dirpath, _, filenames in os.walk(root_dir):
        for fn in filenames:
            fn_stem = Path(fn).stem.lower()
            if fn_stem == req_stem:  # match name without caring about folders
                return os.path.join(dirpath, fn)
    return None



def crop_face_and_shoulders(image_path: str):
    """Optional: crop around the first detected face area."""
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
    """Enable rarfile support if a RAR is uploaded (Windows path or PATH)."""
    import rarfile
    if custom_path:
        if Path(custom_path).exists():
            rarfile.UNRAR_TOOL = custom_path
            return True
        else:
            st.warning("⚠️ Provided unrar.exe path does not exist. Falling back to autodetect.")
    if os.name == "nt":
        for p in [
            r"C:\Program Files\WinRAR\UnRAR.exe",
            r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
            r"C:\Windows\unrar.exe",
            r"C:\Windows\System32\unrar.exe",
        ]:
            if Path(p).exists():
                rarfile.UNRAR_TOOL = p
                return True
    return True  # assume in PATH or rarfile can handle

# ================== Main logic ====================
if excel_file and photos_archive and template_file:
    # Fonts: Arabic size 36 (smaller name to avoid photo overlap)
    font_ar = load_font_from_upload(font_ar_file, "Arabic", 36)
    font_en = load_font_from_upload(font_en_file, "English", 30)

    # Read inputs
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        st.stop()

    try:
        template = Image.open(template_file).convert("RGB")
    except Exception as e:
        st.error(f"❌ Failed to read template image: {e}")
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
                st.error("❌ RAR support not available. Provide a valid unrar.exe path or upload ZIP.")
                shutil.rmtree(tmpdir, ignore_errors=True)
                st.stop()
        else:
            st.error("❌ Unsupported archive type. Upload ZIP or RAR.")
            shutil.rmtree(tmpdir, ignore_errors=True)
            st.stop()
    except Exception as e:
        st.error(f"❌ Failed to extract archive: {e}")
        shutil.rmtree(tmpdir, ignore_errors=True)
        st.stop()

    output_cards: list[Image.Image] = []
    progress = st.progress(0)
    status = st.empty()

    for idx, row in df.iterrows():
        status.info(f"Processing {idx+1}/{len(df)} – {row.get('الاسم', '')}")
        card = template.copy()
        draw = ImageDraw.Draw(card)

        # Prepare texts (Arabic shaping + bidi)
        name = prepare_text(str(row.get("الاسم", "")).strip())
        job  = prepare_text(str(row.get("الوظيفة", "")).strip())
        num  = str(row.get("الرقم", "")).strip()               # الموظف/الرقم الوظيفي
        national_id = str(row.get("الرقم القومي", "")).strip()  # للباركود
        photo_filename = str(row.get("الصورة", "")).strip()

        # ---- TEXT PLACEMENT ----
        base_name_xy = (915, 240)  # anchor reference from the design

        # 1) Draw NAME (bold, nudged left/up)
        name_xy = (base_name_xy[0] + NAME_OFFSET_X, base_name_xy[1] + NAME_OFFSET_Y)
        draw_bold_text(draw, name_xy, name, font_ar, fill="black", anchor="rt")

        # 2) Measure NAME height and add extra spacing (+10) so JOB sits lower
        name_bbox   = draw.textbbox((0, 0), name, font=font_ar)
        name_height = (name_bbox[3] - name_bbox[1]) + 20   # ↑ increased spacing from 5 -> 10

        # 3) Draw JOB directly under NAME with that spacing
        job_xy = (name_xy[0], name_xy[1] + name_height)
        draw_aligned_text(draw, job_xy, job, font=font_ar, fill="black", anchor="rt")

        # 4) Measure JOB height and add larger spacing (+15) so EMPLOYEE NUMBER sits even lower
        job_bbox   = draw.textbbox((0, 0), job, font=font_ar)
        job_height = (job_bbox[3] - job_bbox[1]) + 25      # ↑ increased spacing from 5 -> 15

        # 5) Draw EMPLOYEE NUMBER under JOB (aligned to the right)
        job_id_label = prepare_text(f"الرقم الوظيفي: {num}")
        id_xy = (name_xy[0], job_xy[1] + job_height)
        draw_aligned_text(draw, id_xy, job_id_label, font=font_ar, fill="black", anchor="rt")

        # ---- PHOTO ----
        photo_path = find_photo_path(tmpdir, photo_filename)
        if photo_path and os.path.exists(photo_path):
            try:
                cropped = crop_face_and_shoulders(photo_path)
                img = cropped if cropped is not None else Image.open(photo_path)
                img = img.convert("RGB").resize(PHOTO_SIZE)
                card.paste(img, PHOTO_POS)
            except Exception as e:
                st.warning(f"⚠️ Failed to place photo for '{row.get('الاسم', '')}': {e}")
        else:
            st.warning(f"📷 Photo not found for '{row.get('الاسم', '')}'. Requested: {photo_filename}")

        # ---- BARCODE (using national ID) ----
        try:
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
                    except Exception: pass
            else:
                st.warning(f"🧾 National ID missing for '{row.get('الاسم', '')}'. Skipped barcode.")
        except Exception as e:
            st.warning(f"⚠️ Failed to generate barcode for '{row.get('الاسم', '')}': {e}")

        output_cards.append(card)
        progress.progress(int(((idx + 1) / max(len(df), 1)) * 100))

    status.empty()

    # ---- EXPORT PDF ----
    if output_cards:
        try:
            pdf_path = os.path.join(tmpdir, "All_ID_Cards.pdf")
            output_cards[0].save(pdf_path, save_all=True, append_images=output_cards[1:])
            with open(pdf_path, "rb") as f:
                st.download_button("⬇️ Download All ID Cards (PDF)", f, file_name="All_ID_Cards.pdf")
            st.success(f"✅ Generated {len(output_cards)} cards")
            st.image(output_cards[0], caption="Preview", width=320)
        except Exception as e:
            st.error(f"❌ Failed to write PDF: {e}")
    else:
        st.warning("No cards generated.")

else:
    st.info("👆 Upload the three inputs to start: Excel, Photos archive, Template image.")
