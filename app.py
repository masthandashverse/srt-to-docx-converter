"""
app.py - Simple SRT to DOCX Converter
Upload SRT files → Download DOCX files
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
    .stDownloadButton > button {
        width: 100%;
        background-color: #1a478a !important;
        color: white !important;
    }
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


def convert_to_docx(subtitles, filename):
    """Convert subtitles to DOCX bytes."""
    base_name = os.path.splitext(filename)[0]
    docx_filename = f"{base_name}.docx"

    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        tmp_path = tmp.name

    try:
        writer.create_document(
            subtitles=subtitles,
            source_filename=filename,
            output_path=tmp_path
        )

        with open(tmp_path, 'rb') as f:
            docx_bytes = f.read()
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

    return docx_filename, docx_bytes


def make_zip(files_list):
    """Create ZIP from list of (filename, bytes) tuples."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in files_list:
            zf.writestr(name, data)
    buf.seek(0)
    return buf.getvalue()


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
st.write("Upload subtitle files → Get Word documents")
st.divider()

# ─── File Upload ──────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload SRT files",
    type=["srt"],
    accept_multiple_files=True
)

# ─── Validate File Sizes ─────────────────────────────────────
if uploaded_files:

    # Check file sizes
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

    # Show uploaded files count
    total_size = sum(f.size for f in uploaded_files)
    st.info(
        f"📁 **{len(uploaded_files)}** file(s) uploaded "
        f"({format_size(total_size)})"
    )

    # ─── Convert Button ───────────────────────────────────────
    if st.button(
        f"🚀 Convert {len(uploaded_files)} file(s)",
        type="primary",
        use_container_width=True
    ):
        results = []
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

                docx_name, docx_bytes = convert_to_docx(subs, fname)
                results.append((docx_name, docx_bytes, len(subs)))

            except Exception as e:
                errors.append(f"❌ **{fname}** — {str(e)}")

        progress.progress(1.0, text="✅ Done!")

        # Store results
        st.session_state["results"] = results
        st.session_state["errors"] = errors

    # ─── Show Results ─────────────────────────────────────────
    if "results" in st.session_state:
        results = st.session_state["results"]
        errors = st.session_state["errors"]

        if results:
            st.success(f"✅ {len(results)} file(s) converted!")

        if errors:
            st.warning(f"⚠️ {len(errors)} file(s) failed")
            for err in errors:
                st.write(err)

        if results:
            st.divider()
            st.subheader("💾 Download")

            # ZIP download for multiple files
            if len(results) > 1:
                zip_data = make_zip([(n, b) for n, b, _ in results])
                st.download_button(
                    label=f"📦 Download ALL ({len(results)} files) as ZIP",
                    data=zip_data,
                    file_name="converted_subtitles.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.divider()

            # Individual downloads
            for docx_name, docx_bytes, sub_count in results:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"📄 **{docx_name}** — {sub_count} subtitles")
                with col2:
                    st.download_button(
                        label="⬇️ Download",
                        data=docx_bytes,
                        file_name=docx_name,
                        mime="application/vnd.openxmlformats-officedocument"
                             ".wordprocessingml.document",
                        key=f"dl_{docx_name}"
                    )

else:
    st.info("👆 Upload `.srt` files to get started")

# ─── Footer ───────────────────────────────────────────────────
st.divider()
st.caption("SRT to DOCX Converter")