"""
app.py - Streamlit Web App for SRT to DOCX Converter
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

# ─── Initialize Parser and Writer ─────────────────────────────
parser = SRTParser()
writer = DOCXWriter()


# ─── Helper Functions ─────────────────────────────────────────
def parse_uploaded_srt(uploaded_file):
    """Parse an uploaded SRT file and return subtitle entries."""
    content_bytes = uploaded_file.getvalue()

    # Try different encodings
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

    # Clean content
    content = content.replace('\r\n', '\n').replace('\r', '\n')
    content = content.lstrip('\ufeff').strip()

    # Parse using the existing parser methods
    subtitles = parser._regex_parse(content)
    if not subtitles:
        subtitles = parser._block_parse(content)

    return subtitles


def convert_to_docx(subtitles, filename, style):
    """Convert subtitles to DOCX and return bytes."""
    base_name = os.path.splitext(filename)[0]
    docx_filename = f"{base_name}.docx"

    # Create temp file for writing
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        tmp_path = tmp.name

    try:
        writer.create_document(
            subtitles=subtitles,
            source_filename=filename,
            output_path=tmp_path,
            style=style
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


# ─── App UI ───────────────────────────────────────────────────

# Title
st.title("📝 SRT to DOCX Converter")
st.write("Convert subtitle files (.srt) to Word documents (.docx)")
st.divider()

# ─── Sidebar: Settings ────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    style = st.radio(
        "Output Format:",
        options=["table", "plain", "formatted", "text_only", "script"],
        index=0,
        format_func=lambda x: {
            "table": "📊 Table - Columns with timestamps",
            "plain": "📄 Plain - Numbered with separators",
            "formatted": "🎨 Formatted - Inline timestamps",
            "text_only": "📝 Text Only - No timestamps",
            "script": "🎬 Script - Screenplay style",
        }[x]
    )

    st.divider()
    st.subheader("📖 How to Use")
    st.markdown("""
    1. Upload `.srt` files above
    2. Pick a format style
    3. Click **Convert**
    4. Download your `.docx` files
    """)

# ─── File Upload ──────────────────────────────────────────────
st.subheader("📁 Upload SRT Files")

uploaded_files = st.file_uploader(
    "Choose SRT files",
    type=["srt"],
    accept_multiple_files=True
)

# ─── Show Uploaded Files ──────────────────────────────────────
if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

    # Show file list
    with st.expander("View uploaded files", expanded=False):
        for f in uploaded_files:
            size_kb = f.size / 1024
            st.write(f"📄 **{f.name}** — {size_kb:.1f} KB")

    st.divider()

    # ─── Preview (optional) ───────────────────────────────────
    st.subheader("👁️ Preview")

    preview_file = st.selectbox(
        "Select file to preview:",
        uploaded_files,
        format_func=lambda f: f.name
    )

    if preview_file:
        try:
            subs = parse_uploaded_srt(preview_file)
            preview_file.seek(0)  # Reset for later use

            if subs:
                st.info(f"Found **{len(subs)}** subtitles")

                # Show first few as table
                show_count = min(10, len(subs))
                preview_data = []
                for sub in subs[:show_count]:
                    d = sub.to_dict() if hasattr(sub, 'to_dict') else sub
                    preview_data.append({
                        "#": d["index"],
                        "Start": d["start_time"],
                        "End": d["end_time"],
                        "Text": d["text"][:80]
                    })

                st.dataframe(preview_data, use_container_width=True, hide_index=True)

                if len(subs) > show_count:
                    st.caption(f"Showing first {show_count} of {len(subs)} entries")
            else:
                st.warning("Could not parse subtitles from this file")
        except Exception as e:
            st.error(f"Preview error: {e}")

    st.divider()

    # ─── Convert Button ───────────────────────────────────────
    st.subheader("🚀 Convert")

    if st.button(
        f"Convert {len(uploaded_files)} file(s) to DOCX",
        type="primary",
        use_container_width=True
    ):
        results = []
        errors = []

        # Progress bar
        progress = st.progress(0, text="Starting...")

        for i, uploaded_file in enumerate(uploaded_files):
            fname = uploaded_file.name
            progress.progress(
                (i + 1) / len(uploaded_files),
                text=f"Converting: {fname}"
            )

            try:
                # Parse
                uploaded_file.seek(0)
                subs = parse_uploaded_srt(uploaded_file)

                if not subs:
                    errors.append(f"❌ **{fname}** — No subtitles found")
                    continue

                # Convert
                docx_name, docx_bytes = convert_to_docx(subs, fname, style)
                results.append((docx_name, docx_bytes, len(subs)))

            except Exception as e:
                errors.append(f"❌ **{fname}** — {str(e)}")

        progress.progress(1.0, text="Done!")

        # ─── Store Results in Session ─────────────────────────
        st.session_state["results"] = results
        st.session_state["errors"] = errors

    # ─── Show Results ─────────────────────────────────────────
    if "results" in st.session_state:
        results = st.session_state["results"]
        errors = st.session_state["errors"]

        if results:
            st.success(f"🎉 {len(results)} file(s) converted successfully!")
        if errors:
            st.warning(f"⚠️ {len(errors)} file(s) had errors")
            for err in errors:
                st.write(err)

        st.divider()

        if results:
            st.subheader("💾 Download")

            # Download ALL as ZIP (if multiple files)
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
                st.write("**Or download individually:**")

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

# ─── Empty State ──────────────────────────────────────────────
else:
    st.info("👆 Upload `.srt` files to get started")

# ─── Footer ───────────────────────────────────────────────────
st.divider()
st.caption("📝 SRT to DOCX Converter • Built with Streamlit & python-docx")