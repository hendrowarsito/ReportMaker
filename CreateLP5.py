import streamlit as st
from docx import Document
from io import BytesIO
import pandas as pd
import zipfile
import locale
import os

# ====================== CONFIGURATION ======================
st.set_page_config(
    page_title="SRR Kalibata Report Maker",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk tampilan profesional
st.markdown("""
    <style>
    .main-header {
        font-size: 2.3rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 0.3rem;
        font-weight: 700;
    }
    .sub-header {
        font-size: 1.15rem;
        color: #475569;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #1E40AF;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.65rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #1E3A8A;
    }
    .upload-section {
        background-color: #F8FAFC;
        padding: 1.8rem;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        margin-bottom: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# ====================== SIDEBAR ======================
with st.sidebar:
    st.markdown("### 📋 SRR Kalibata")
    st.markdown("**Report Maker**")
    st.markdown("---")
    
    st.info("""
    Aplikasi ini digunakan untuk membuat laporan secara otomatis 
    menggunakan template Microsoft Word dan data Excel.
    """)
    
    st.markdown("### 📌 Panduan Singkat")
    st.caption("1. Upload template DOCX\n2. Upload file Excel\n3. Klik Generate Reports")

# ====================== MAIN CONTENT ======================
st.markdown('<h1 class="main-header">SRR KALIBATA REPORT MAKER</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Sistem Pembuatan Laporan Otomatis • Professional Edition</p>', unsafe_allow_html=True)

st.markdown("---")

# Upload Section
col1, col2 = st.columns([3, 1])

with col1:
    st.subheader("📤 Upload Dokumen")
    
    uploaded_templates = st.file_uploader(
        "Template Laporan (.docx)", 
        type="docx", 
        accept_multiple_files=True,
        help="Bisa upload lebih dari satu template"
    )
    
    uploaded_excel = st.file_uploader(
        "Data Lembar Kerja (.xlsx)", 
        type="xlsx",
        help="File Excel yang berisi data pengganti placeholder"
    )

with col2:
    st.subheader("ℹ️ Status")
    if uploaded_templates:
        st.success(f"✅ {len(uploaded_templates)} Template terupload")
    if uploaded_excel:
        st.success("✅ Data Excel terupload")

# ====================== CORE LOGIC ======================
if uploaded_templates and uploaded_excel:
    try:
        data = pd.read_excel(uploaded_excel)
        
        st.subheader("📊 Preview Data")
        st.dataframe(data, use_container_width=True, height=320)

        templates = {}
        all_placeholders = set()
        
        for file in uploaded_templates:
            document = Document(file)
            placeholders = extract_placeholders(document)
            templates[file.name] = {"document": document, "placeholders": placeholders}
            all_placeholders.update(placeholders)

        unmatched = [ph for ph in all_placeholders if ph not in data.columns]
        if unmatched:
            st.warning(f"⚠️ Placeholder berikut tidak ditemukan di Excel: {', '.join(unmatched)}")

        if st.button("🚀 Generate Reports", type="primary", use_container_width=True):
            with st.spinner("Sedang memproses laporan..."):
                generated_files = []
                for index, row in data.iterrows():
                    for template_name, template_data in templates.items():
                        document = replace_placeholders(template_data["document"], row.to_dict())
                        file_name = f"{index+1:03d}_{template_name}"
                        buffer = save_docx(document)
                        generated_files.append((file_name, buffer))

                zip_buffer = generate_zip(generated_files)
                
                st.success(f"✅ Berhasil membuat {len(generated_files)} dokumen laporan!")
                st.download_button(
                    label="📥 Download Semua Laporan sebagai ZIP",
                    data=zip_buffer,
                    file_name="Laporan_SRR_Kalibata.zip",
                    mime="application/zip",
                    use_container_width=True
                )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #64748B; font-size: 0.95rem;'>"
    "© 2026 SRR Kalibata • Report Maker Professional</p>", 
    unsafe_allow_html=True
)

# ====================== HELPER FUNCTIONS ======================
def format_number_indonesia(value):
    """Format number to Indonesian format (e.g., 12.000,00)."""
    try:
        locale.setlocale(locale.LC_NUMERIC, "id_ID.UTF-8")
        return locale.format_string("%.2f", value, grouping=True)
    except:
        return str(value)

def replace_placeholders(document, replacements):
    """Replace placeholders in a DOCX document with provided values."""
    for paragraph in document.paragraphs:
        for key, value in replacements.items():
            formatted_value = format_number_indonesia(value) if isinstance(value, (int, float)) else value
            paragraph.text = paragraph.text.replace(f"{{{key}}}", str(formatted_value))
    return document

def extract_placeholders(document):
    """Extract placeholders from a DOCX document."""
    placeholders = set()
    for paragraph in document.paragraphs:
        if "{" in paragraph.text and "}" in paragraph.text:
            placeholders.update(
                part.strip("{}") for part in paragraph.text.split() if part.startswith("{") and part.endswith("}")
            )
    return sorted(placeholders)

def save_docx(document):
    """Save changes to a new DOCX file."""
    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer

def generate_zip(files):
    """Generate a ZIP file from a list of files."""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_name, file_buffer in files:
            zf.writestr(file_name, file_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer
