import os
import copy
import time
from PyPDF2 import PdfReader, PdfWriter, PageObject, Transformation
import streamlit as st
from PIL import Image
from io import BytesIO
import zipfile

# --- LOAD LOGO ---
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")  # Make sure logo.png is in the same folder
if os.path.exists(logo_path):
    logo = Image.open(logo_path)
    st.image(logo, width=500)  # Adjust width as needed

# --- PAGE TITLE ---
st.title("CV Branding with Letterhead")

# --- SAFETY / TUNING ---
PAGE_PADDING = 20        # points of padding between CV and letterhead edges
MAX_WRITE_RETRIES = 5    # handle file locks
SCALE_DOWN_IF_LARGER = True  # scale CV down if it doesn't fit letterhead
SCALE_MARGIN = 2.0       # extra shrink (percent)

# --- FOLDER SETUP ---
output_folder = "branded_cvs"
# Clear folder before each run
if os.path.exists(output_folder):
    for f in os.listdir(output_folder):
        file_path = os.path.join(output_folder, f)
        if os.path.isfile(file_path):
            os.remove(file_path)
else:
    os.makedirs(output_folder, exist_ok=True)


# --- DEFAULT LETTERHEAD ---
st.subheader("Letterhead")
uploaded_letterhead = st.file_uploader(
    "Upload a custom letterhead (PDF only)",
    type="pdf",
    key="letterhead"
)

# Path to default letterhead included in the repo
default_letterhead_path = st.secrets["default_letterhead"]

if uploaded_letterhead:
    st.success("Using uploaded letterhead")
    letterhead_stream = BytesIO(uploaded_letterhead.read())
else:
    if os.path.exists(default_letterhead_path):
        st.info("Using default letterhead")
        with open(default_letterhead_path, "rb") as f:
            letterhead_stream = BytesIO(f.read())
    else:
        st.error("⚠️ Default letterhead not found. Upload a letterhead to proceed.")
        st.stop()


# --- UPLOAD CV FILES ---
st.subheader("Upload CVs to Brand")
uploaded_files = st.file_uploader("Upload PDF resumes", type="pdf", accept_multiple_files=True)

# --- PROCESS FILES ---
if uploaded_files:
    # Read letterhead once
    brand_reader = PdfReader(letterhead_stream)
    brand_page_orig = brand_reader.pages[0]
    brand_w = float(brand_page_orig.mediabox.width)
    brand_h = float(brand_page_orig.mediabox.height)
    st.info(f"Letterhead size: {brand_w} x {brand_h} pts")

    for file in uploaded_files:
        fname = file.name
        try:
            cv_reader = PdfReader(file)
        except Exception as e:
            st.warning(f"Failed to open {fname}: {e}")
            continue

        writer = PdfWriter()

        for page_index, cv_page in enumerate(cv_reader.pages):
            brand_page = copy.deepcopy(brand_page_orig)
            cv_copy = copy.deepcopy(cv_page)

            cv_w = float(cv_copy.mediabox.width)
            cv_h = float(cv_copy.mediabox.height)

            avail_w = brand_w - 2 * PAGE_PADDING
            avail_h = brand_h - 2 * PAGE_PADDING

            scale = 1.0
            if SCALE_DOWN_IF_LARGER and (cv_w > avail_w or cv_h > avail_h):
                sx = avail_w / cv_w
                sy = avail_h / cv_h
                scale = min(sx, sy) * (1.0 - SCALE_MARGIN / 100.0)
                if scale <= 0:
                    scale = min(sx, sy)

            scaled_w = cv_w * scale
            scaled_h = cv_h * scale
            x_offset = (brand_w - scaled_w) / 2.0
            y_offset = (brand_h - scaled_h) / 2.0

            transf = Transformation().scale(scale, scale).translate(x_offset, y_offset)

            final_page = PageObject.create_blank_page(width=brand_w, height=brand_h)
            final_page.merge_page(brand_page)
            cv_copy.add_transformation(transf)
            final_page.merge_page(cv_copy)
            writer.add_page(final_page)

        out_path = os.path.join(output_folder, fname)
        for attempt in range(1, MAX_WRITE_RETRIES + 1):
            try:
                with open(out_path, "wb") as f_out:
                    writer.write(f_out)
                st.success(f"✅ {fname} -> written to {output_folder}")
                break
            except PermissionError:
                wait = 1.0 * attempt
                st.warning(f"PermissionError writing {out_path}, retrying in {wait}s (attempt {attempt})...")
                time.sleep(wait)
        else:
            st.error(f"❌ Failed to write {out_path} after {MAX_WRITE_RETRIES} attempts.")

    st.info(f"All processed CVs are saved in '{output_folder}'")

    # --- ZIP DOWNLOAD ---
    st.subheader("Download Branded CVs as ZIP")
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for fname in os.listdir(output_folder):
            if fname.lower().endswith(".pdf"):
                file_path = os.path.join(output_folder, fname)
                zip_file.write(file_path, arcname=fname)
    zip_buffer.seek(0)

    st.download_button(
        label="Download All Branded CVs",
        data=zip_buffer,
        file_name="branded_cvs.zip",
        mime="application/zip"
    )

