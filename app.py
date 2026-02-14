"""
app.py - Simple SRT to DOCX Converter
Upload SRT files → Save DOCX files directly to your folder
"""

import os
import io
import zipfile
import tempfile
import streamlit as st

from srt_parser import SRTParser
from docx_writer import DOCXWriter


# ─── Page Config ──────────────────────────────────────────────
st.set_page_config(
    page_title="SRT to DOCX Converter",
    page_icon="📝",
    layout="centered"
)

# ─── Simple CSS ───────────────────────────────────────────────
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ─── Initialize ───────────────────────────────────────────────
parser = SRTParser()
writer = DOCXWriter()

MAX_FILE_SIZE_MB = 500


# ─── Helper Functions ─────────────────────────────────────────
def parse_uploaded_srt(uploaded_file):
    """Parse an uploaded SRT file."""
    content_bytes = uploaded_file.getvalue()

    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    content = None

    for enc in encodings:
        try:
            content = content_bytes.decode(enc)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if content is None:
        return None

    content = content.replace('\r\n', '\n').replace('\r', '\n')
    content = content.lstrip('\ufeff').strip()

    subtitles = parser._regex_parse(content)
    if not subtitles:
        subtitles = parser._block_parse(content)

    return subtitles


def convert_and_save(subtitles, filename, output_folder):
    """Convert subtitles and save DOCX directly to folder."""
    base_name = os.path.splitext(filename)[0]
    docx_filename = f"{base_name}.docx"
    output_path = os.path.join(output_folder, docx_filename)

    # Handle duplicate filenames
    counter = 1
    while os.path.exists(output_path):
        output_path = os.path.join(output_folder, f"{base_name} ({counter}).docx")
        counter += 1

    writer.create_document(
        subtitles=subtitles,
        source_filename=filename,
        output_path=output_path
    )

    return output_path


def format_size(size_bytes):
    """Format bytes to readable string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


# ─── App UI ───────────────────────────────────────────────────

# Title
st.title("📝 SRT to DOCX Converter")
st.write("Upload subtitle files → Save Word documents to your folder")
st.divider()

# ─── Output Folder ────────────────────────────────────────────
st.subheader("📁 Output Folder")

# Default folder = Desktop
default_folder = os.path.join(os.path.expanduser("~"), "Desktop", "SRT_Converted")

output_folder = st.text_input(
    "Enter folder path where DOCX files will be saved:",
    value=default_folder,
    help="Type the full folder path. Folder will be created if it doesn't exist."
)

# ─── File Upload ──────────────────────────────────────────────
st.divider()
st.subheader("📄 Upload SRT Files")

uploaded_files = st.file_uploader(
    "Choose SRT files",
    type=["srt"],
    accept_multiple_files=True
)

# ─── Main Logic ───────────────────────────────────────────────
if uploaded_files:

    # Validate file sizes
    oversized = []
    for f in uploaded_files:
        size_mb = f.size / (1024 * 1024)
        if size_mb > MAX_FILE_SIZE_MB:
            oversized.append(f"{f.name} ({size_mb:.1f} MB)")

    if oversized:
        st.error(
            f"Files exceed {MAX_FILE_SIZE_MB} MB limit:\n\n"
            + "\n".join(oversized)
        )
        st.stop()

    # Show file count
    total_size = sum(f.size for f in uploaded_files)
    st.info(
        f"📁 **{len(uploaded_files)}** file(s) uploaded "
        f"({format_size(total_size)})"
    )

    # Show files list
    with st.expander("View uploaded files"):
        for f in uploaded_files:
            st.write(f"📄 {f.name} — {format_size(f.size)}")

    st.divider()

    # ─── Convert Button ───────────────────────────────────────
    if st.button(
        f"🚀 Convert & Save {len(uploaded_files)} file(s) to folder",
        type="primary",
        use_container_width=True
    ):
        # Validate folder path
        if not output_folder.strip():
            st.error("❌ Please enter an output folder path!")
            st.stop()

        # Create folder
        try:
            os.makedirs(output_folder.strip(), exist_ok=True)
        except Exception as e:
            st.error(f"❌ Cannot create folder: {e}")
            st.stop()

        saved_files = []
        errors = []

        progress = st.progress(0, text="Converting...")

        for i, uploaded_file in enumerate(uploaded_files):
            fname = uploaded_file.name
            progress.progress(
                (i + 1) / len(uploaded_files),
                text=f"Converting: {fname}"
            )

            try:
                uploaded_file.seek(0)
                subs = parse_uploaded_srt(uploaded_file)

                if not subs:
                    errors.append(f"❌ **{fname}** — No subtitles found")
                    continue

                # Save directly to folder
                saved_path = convert_and_save(
                    subs, fname, output_folder.strip()
                )
                saved_files.append((fname, saved_path, len(subs)))

            except Exception as e:
                errors.append(f"❌ **{fname}** — {str(e)}")

        progress.progress(1.0, text="✅ Done!")

        # ─── Show Results ─────────────────────────────────────
        if saved_files:
            st.success(
                f"🎉 **{len(saved_files)} file(s) saved** to:\n\n"
                f"`{output_folder.strip()}`"
            )

            st.divider()
            st.subheader("✅ Saved Files")

            for original_name, saved_path, sub_count in saved_files:
                docx_name = os.path.basename(saved_path)
                st.write(
                    f"✅ **{docx_name}** — "
                    f"{sub_count} subtitles"
                )

        if errors:
            st.divider()
            st.warning(f"⚠️ {len(errors)} file(s) failed")
            for err in errors:
                st.write(err)

else:
    st.info("👆 Upload `.srt` files to get started")

# ─── Footer ───────────────────────────────────────────────────
st.divider()
st.caption("SRT to DOCX Converter — Files save directly to your folder")
