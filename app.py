import streamlit as st
import pandas as pd
import io
import json
import os
from PIL import Image
from pypdf import PdfReader
import google.generativeai as genai
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page Configuration
st.set_page_config(
    page_title="SheetGen AI | Document to Excel Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Fetch API Key silently from Streamlit Secrets or Environment
api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY", "")

# Modern Dark Theme, Static Sidebar Lock & Motion UI CSS
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Space+Grotesk:wght@600;700&display=swap');

.stApp {
    background: radial-gradient(circle at 50% 0%, #111827 0%, #080b11 75%, #05070a 100%) !important;
    color: #e2e8f0;
    font-family: 'Plus Jakarta Sans', sans-serif;
}

section[data-testid="stSidebar"] {
    min-width: 300px !important;
    max-width: 300px !important;
    width: 300px !important;
    background-color: #0b0f19 !important;
    border-right: 1px solid rgba(255, 255, 255, 0.08) !important;
}

[data-testid="stSidebarResizer"] {
    display: none !important;
    pointer-events: none !important;
}

h1, h2, h3 {
    font-family: 'Space Grotesk', sans-serif !important;
    letter-spacing: -0.02em;
    font-weight: 700;
}

.header-wrapper {
    display: flex;
    align-items: center;
    gap: 16px;
    margin-top: 0.5rem;
    margin-bottom: 0.25rem;
}

.excel-title-text {
    background: linear-gradient(135deg, #22c55e 0%, #10b981 50%, #38bdf8 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 2.8rem;
    font-weight: 800;
    line-height: 1.15;
    font-family: 'Space Grotesk', sans-serif;
}

.sub-heading {
    color: #94a3b8;
    font-size: 1.05rem;
    font-weight: 400;
    margin-bottom: 1.75rem;
}

@keyframes floatCard1 {
    0% { transform: translateY(0px); }
    50% { transform: translateY(-7px); }
    100% { transform: translateY(0px); }
}
@keyframes floatCard2 {
    0% { transform: translateY(0px); }
    50% { transform: translateY(-5px); }
    100% { transform: translateY(0px); }
}
@keyframes floatCard3 {
    0% { transform: translateY(0px); }
    50% { transform: translateY(-8px); }
    100% { transform: translateY(0px); }
}

.feature-card {
    background: rgba(17, 24, 39, 0.7);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 14px;
    padding: 1.3rem;
    backdrop-filter: blur(12px);
    box-shadow: 0 8px 30px rgba(0, 0, 0, 0.35);
    transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
}

.card-anim-1 { animation: floatCard1 5.5s ease-in-out infinite; }
.card-anim-2 { animation: floatCard2 6.2s ease-in-out infinite 0.75s; }
.card-anim-3 { animation: floatCard3 5.8s ease-in-out infinite 1.5s; }

.feature-card:hover {
    transform: translateY(-9px) scale(1.02) !important;
    border-color: rgba(34, 197, 94, 0.45) !important;
    box-shadow: 0 16px 36px rgba(0, 0, 0, 0.5), 0 0 22px rgba(34, 197, 94, 0.25) !important;
}

.sidebar-item {
    background: rgba(17, 24, 39, 0.6);
    border: 1px solid rgba(255, 255, 255, 0.05);
    border-radius: 8px;
    padding: 10px 12px;
    margin-bottom: 8px;
}
.sidebar-title {
    font-size: 0.85rem;
    font-weight: 600;
    color: #e2e8f0;
    margin-bottom: 2px;
}
.sidebar-desc {
    font-size: 0.75rem;
    color: #94a3b8;
    line-height: 1.25;
}

.status-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 4px 12px;
    border-radius: 9999px;
    font-size: 0.75rem;
    font-weight: 600;
    background: rgba(34, 197, 94, 0.12);
    color: #4ade80;
    border: 1px solid rgba(34, 197, 94, 0.3);
    margin-bottom: 1rem;
}

/* ENLARGED & PERMANENTLY VISIBLE DATA EDITOR ACTION TOOLBAR */
[data-testid="stElementToolbar"] {
    opacity: 1 !important;
    visibility: visible !important;
    display: flex !important;
    gap: 6px !important;
    background: rgba(15, 23, 42, 0.95) !important;
    border: 1px solid rgba(34, 197, 94, 0.35) !important;
    border-radius: 10px !important;
    padding: 4px 8px !important;
    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.5), 0 0 12px rgba(34, 197, 94, 0.2) !important;
    top: -46px !important;
    right: 0px !important;
}

[data-testid="stElementToolbar"] button {
    width: 34px !important;
    height: 34px !important;
    border-radius: 8px !important;
    background: rgba(255, 255, 255, 0.08) !important;
    color: #4ade80 !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    transition: all 0.2s ease-in-out !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
}

[data-testid="stElementToolbar"] button:hover {
    background: rgba(34, 197, 94, 0.25) !important;
    border-color: #22c55e !important;
    color: #ffffff !important;
    transform: scale(1.12) !important;
}

[data-testid="stElementToolbar"] svg {
    width: 18px !important;
    height: 18px !important;
    fill: currentColor !important;
}

/* Single Master Excel Download Button (Emerald Green) */
.stDownloadButton > button {
    background: linear-gradient(135deg, #107C41 0%, #15803d 100%) !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    font-size: 1.05rem !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.85rem 2rem !important;
    box-shadow: 0 4px 25px rgba(16, 124, 65, 0.45) !important;
    transition: all 0.25s ease-in-out !important;
}
.stDownloadButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 30px rgba(16, 124, 65, 0.75) !important;
}

.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #16a34a 0%, #15803d 100%) !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.65rem 1.75rem !important;
    box-shadow: 0 0 20px rgba(22, 163, 74, 0.4) !important;
}

div[data-testid="stFileUploader"] {
    background: rgba(17, 24, 39, 0.5) !important;
    border: 1px dashed rgba(255, 255, 255, 0.15) !important;
    border-radius: 14px !important;
    padding: 1.5rem !important;
}
</style>
""", unsafe_allow_html=True)

# Microsoft Excel Vector Badges
EXCEL_ICON_MAIN = '<svg width="48" height="48" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="14" y="6" width="28" height="36" rx="4" fill="#107C41"/><rect x="23" y="13" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="13" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="23" y="19" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="19" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="23" y="25" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="25" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="23" y="31" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="31" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="6" y="9" width="22" height="30" rx="3" fill="#185C37" style="filter: drop-shadow(2px 4px 6px rgba(0,0,0,0.5));"/><path d="M11.5 30L15.3 24L11.8 18H15.2L17 21.5L18.8 18H22.2L18.7 24L22.5 30H19.1L17 26.2L14.9 30H11.5Z" fill="white"/></svg>'

EXCEL_ICON_SIDEBAR = '<svg width="28" height="28" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle;"><rect x="14" y="6" width="28" height="36" rx="4" fill="#107C41"/><rect x="23" y="13" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="13" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="23" y="19" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="19" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="23" y="25" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="25" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="23" y="31" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="31" y="31" width="6" height="4" fill="#ffffff" fill-opacity="0.85"/><rect x="6" y="9" width="22" height="30" rx="3" fill="#185C37"/><path d="M11.5 30L15.3 24L11.8 18H15.2L17 21.5L18.8 18H22.2L18.7 24L22.5 30H19.1L17 26.2L14.9 30H11.5Z" fill="white"/></svg>'

# Static Sidebar Configuration
with st.sidebar:
    st.markdown(f'<div style="display:flex; align-items:center; gap:10px; margin-bottom: 6px;">{EXCEL_ICON_SIDEBAR}<span style="font-size:1.35rem; font-weight:700; color:#ffffff; font-family:\'Space Grotesk\',sans-serif;">SheetGen AI</span></div>', unsafe_allow_html=True)
    st.caption("Automated Tabular Data Extraction Engine")
    st.markdown("<hr style='border:none; border-top:1px solid rgba(255,255,255,0.08); margin:12px 0;'>", unsafe_allow_html=True)
    
    st.markdown("<div style='font-size:0.8rem; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; color:#94a3b8; margin-bottom:10px;'>Core Capabilities</div>", unsafe_allow_html=True)
    
    sidebar_items_html = (
        '<div class="sidebar-item"><div class="sidebar-title">⚡ Instant Image Optimizer</div>'
        '<div class="sidebar-desc">Rapid OCR for handwritten registers and multi-page tables.</div></div>'
        '<div class="sidebar-item"><div class="sidebar-title">🗃️ Multi-Table Isolation</div>'
        '<div class="sidebar-desc">Guaranteed distinct tabs for each uploaded page/table.</div></div>'
        '<div class="sidebar-item"><div class="sidebar-title">🧹 Auto-Sanitization & Styling</div>'
        '<div class="sidebar-desc">Removes artifacts, auto-fits columns & formats Excel sheets.</div></div>'
        '<div class="sidebar-item"><div class="sidebar-title">✏️ Visual In-Browser Grid</div>'
        '<div class="sidebar-desc">Enlarged toolbar to add rows, search, hide columns & edit data.</div></div>'
        '<div class="sidebar-item"><div class="sidebar-title">📥 Native .xlsx Generator</div>'
        '<div class="sidebar-desc">Multi-sheet Excel workbook with formatted master tab.</div></div>'
    )
    st.markdown(sidebar_items_html, unsafe_allow_html=True)
    st.markdown("<hr style='border:none; border-top:1px solid rgba(255,255,255,0.08); margin:12px 0;'>", unsafe_allow_html=True)
    st.markdown("🟢 **System Status:** Ready")

# Hero Header
st.markdown('<div class="status-badge">⚡ Instant Document to Spreadsheet</div>', unsafe_allow_html=True)
main_header_html = f'<div class="header-wrapper">{EXCEL_ICON_MAIN}<div class="excel-title-text">SheetGen AI</div></div>'
st.markdown(main_header_html, unsafe_allow_html=True)
st.markdown('<div class="sub-heading">Upload single or batch images & PDFs → Automatically convert them to clean, editable Excel workbooks.</div>', unsafe_allow_html=True)

# 3 Floating Motion Feature Cards
col_card1, col_card2, col_card3 = st.columns(3)
with col_card1:
    card1_html = (
        '<div class="feature-card card-anim-1">'
        '<div style="font-size: 1.15rem; font-weight:700; color:#4ade80; margin-bottom:4px;">1. Upload File(s)</div>'
        '<div style="font-size:0.85rem; color:#94a3b8; line-height:1.4;">Drop photos or PDFs containing tabular registers.</div>'
        '</div>'
    )
    st.markdown(card1_html, unsafe_allow_html=True)

with col_card2:
    card2_html = (
        '<div class="feature-card card-anim-2">'
        '<div style="font-size: 1.15rem; font-weight:700; color:#38bdf8; margin-bottom:4px;">2. AI Formats & Cleans</div>'
        '<div style="font-size:0.85rem; color:#94a3b8; line-height:1.4;">Extracts rows, columns, and transcribes values rapidly.</div>'
        '</div>'
    )
    st.markdown(card2_html, unsafe_allow_html=True)

with col_card3:
    card3_html = (
        '<div class="feature-card card-anim-3">'
        '<div style="font-size: 1.15rem; font-weight:700; color:#a78bfa; margin-bottom:4px;">3. Download Excel</div>'
        '<div style="font-size:0.85rem; color:#94a3b8; line-height:1.4;">Get a styled multi-sheet Excel file (.xlsx) ready for use.</div>'
        '</div>'
    )
    st.markdown(card3_html, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Fast In-Memory Image Resizer
def prepare_image(raw_bytes):
    img = Image.open(io.BytesIO(raw_bytes))
    if img.mode != "RGB":
        img = img.convert("RGB")
    max_dim = 1600
    if max(img.size) > max_dim:
        img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=85, optimize=True)
    out.seek(0)
    return Image.open(out)

# Enterprise OpenPyXL Workbook Formatter
def format_excel_workbook(writer):
    workbook = writer.book
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="107C41", end_color="107C41", fill_type="solid")
    regular_font = Font(name="Calibri", size=10, color="1F2937")
    
    thin_border = Border(
        left=Side(style='thin', color='E5E7EB'),
        right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),
        bottom=Side(style='thin', color='E5E7EB')
    )
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        ws.views.sheetView[0].showGridLines = True
        
        # Style Header Row
        ws.row_dimensions[1].height = 26
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        
        # Style Data Rows
        for row_num in range(2, ws.max_row + 1):
            is_even = (row_num % 2 == 0)
            ws.row_dimensions[row_num].height = 20
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_num, column=col_num)
                if cell.value is None or str(cell.value).strip().lower() in ["none", "nan"]:
                    cell.value = ""
                cell.font = regular_font
                cell.border = thin_border
                
                val_str = str(cell.value or "").strip()
                if val_str.isdigit() and len(val_str) < 5:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
                if is_even:
                    cell.fill = zebra_fill
                    
        # Auto-adjust column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

# Multi-Model Resilient Cascade
def execute_extraction_cascade(files_data, key_str):
    genai.configure(api_key=key_str)
    
    prompt = """
    You are an expert Enterprise Data Engineer and Senior Operations Analyst.
    Extract all tabular data from the uploaded file(s) with maximum accuracy, cleaning, and structure (including handwritten registers and field rosters):
    
    1. EXTRACT ALL DISTINCT TABLES:
       - Transcribe every row, column header, serial number, employee/contact name, reporting manager, phone number(s), state, and remarks.
       - Clean values: If a cell is empty or has noise, make it an empty string "". Do NOT write literal "None" or "NaN".
       - Standardize headers across pages (e.g. use "NO.", "EMPLOYEE NAME", "REPORTING MANAGER", "PHONE", "STATE", "REMARKS").
       - Separate distinct pages or sheets into distinct tables with descriptive titles (e.g. "Employee Register - Page 1", "Employee Register - Page 2").
       
    2. COMPREHENSIVE EXECUTIVE BUSINESS SUMMARY:
       - Write a clear, structured executive summary in clean Markdown.
       - Include:
         * 📌 **Document Scope**: Brief explanation of what this data represents.
         * 📊 **Operational Metrics**: Total records parsed, managers count, geographic regions covered.
         * 🔍 **Key Observations & Action Items**: Connectivity issues, remarks summary (e.g., unreachable numbers, switch-offs).

    Return output strictly as valid JSON matching this schema:
    {
      "analysis": "Structured summary text.",
      "tables": [
        {
          "table_name": "Employee Register - Page 1",
          "headers": ["NO.", "EMPLOYEE NAME", "REPORTING MANAGER", "PHONE", "STATE", "REMARKS"],
          "rows": [
            ["1", "John Doe", "Jane Smith", "9876543210", "ASSAM", "Active"]
          ]
        }
      ]
    }
    """
    
    contents = []
    for file_bytes, mime_type in files_data:
        if "image" in mime_type:
            contents.append(prepare_image(file_bytes))
        elif "pdf" in mime_type:
            reader = PdfReader(io.BytesIO(file_bytes))
            pdf_text = "\n".join([p.extract_text() or "" for p in reader.pages])
            contents.append(f"PDF Content:\n{pdf_text}")
            
    contents.append(prompt)

    model_cascade = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash-latest"]
    
    last_err = None
    for model_name in model_cascade:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(
                contents,
                generation_config={"response_mime_type": "application/json"}
            )
            if response and response.text:
                raw_text = response.text.strip()
                if "```json" in raw_text:
                    raw_text = raw_text.split("```json")[1].split("```")[0].strip()
                elif "```" in raw_text:
                    raw_text = raw_text.split("```")[1].split("```")[0].strip()
                return raw_text, model_name
        except Exception as err:
            last_err = err
            continue

    raise Exception(f"Extraction failed across models. Last Error: {last_err}")

# Unique Sheet Name Generator
def create_unique_sheet_name(raw_name, index, seen_set):
    clean = "".join(c for c in raw_name if c.isalnum() or c in (' ', '_', '-')).strip()
    clean = clean.replace('_', ' ')
    if not clean:
        clean = f"Page {index+1}"
    base_name = f"Sheet {index+1} - {clean[:16]}".strip()
    candidate = base_name[:31]
    count = 1
    while candidate in seen_set:
        suffix = f" ({count})"
        candidate = base_name[:31 - len(suffix)] + suffix
        count += 1
    seen_set.add(candidate)
    return candidate

# Document Upload Section with Multiple File Support
uploaded_files = st.file_uploader(
    "Drop your PDF document(s) or images here", 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if uploaded_files:
    if not api_key:
        st.error("⚠️ System Error: GEMINI_API_KEY not found in Streamlit Secrets. Please add it to your app settings.")
    else:
        files_data = []
        for file in uploaded_files:
            file_bytes = file.read()
            files_data.append((file_bytes, file.type))
        
        col_prev, col_action = st.columns([1, 2])
        with col_prev:
            st.subheader(f"📄 Uploaded Files ({len(uploaded_files)})")
            for idx, file in enumerate(uploaded_files):
                if "image" in file.type:
                    img = Image.open(io.BytesIO(files_data[idx][0]))
                    st.image(img, caption=file.name, use_container_width=True)
                elif file.type == "application/pdf":
                    st.info(f"📑 PDF: **{file.name}** ({len(files_data[idx][0])/1024:.1f} KB)")
                
        with col_action:
            st.subheader("⚡ Convert to Excel")
            st.caption(f"Extract and compile all {len(uploaded_files)} file(s) into a unified multi-sheet Excel spreadsheet.")
            if st.button("🚀 Extract Tables & Convert to Excel", type="primary", use_container_width=True):
                with st.spinner("Extracting tabular data and formatting Excel sheets..."):
                    try:
                        raw_json_str, used_model = execute_extraction_cascade(files_data, api_key)
                        data = json.loads(raw_json_str)
                        st.session_state["extracted_data"] = data
                        st.session_state["model_used"] = used_model
                        st.toast(f"Extracted successfully via {used_model}!", icon="⚡")
                    except Exception as e:
                        st.error(f"Processing Error: {e}")

# Results Display & Editable Table Grid
if "extracted_data" in st.session_state:
    data = st.session_state["extracted_data"]
    
    st.markdown("---")
    st.subheader("💡 Executive Summary & Business Insights")
    
    summary_text = data.get("analysis", "No summary provided.")
    with st.container():
        st.markdown(summary_text)
    
    tables = data.get("tables", [])
    if not tables:
        st.warning("No tables found in the uploaded file(s).")
    else:
        st.markdown("---")
        st.subheader("✏️ Review & Edit Tables")
        st.caption("Use the top-right toolbar on each table to Add Rows (+), Hide/Show Columns (👁), Search (🔍), or go Fullscreen (⛶). Double-click any cell to edit values directly.")
        
        edited_dfs = {}
        for idx, tbl in enumerate(tables):
            table_name = tbl.get("table_name", f"Table_{idx+1}")
            headers = tbl.get("headers", [])
            rows = tbl.get("rows", [])
            
            cleaned_rows = []
            for r in rows:
                cleaned_rows.append([("" if val is None or str(val).strip() in ["None", "NaN", "nan"] else str(val).strip()) for val in r])
                
            df = pd.DataFrame(cleaned_rows, columns=headers if headers else None)
            df.fillna("", inplace=True)
            
            # Clean Table Title with HTML badge
            badge_html = f'''
            <div style="display:flex; align-items:center; gap:12px; margin-top:1.8rem; margin-bottom:0.75rem;">
                <span style="font-size:1.25rem; font-weight:700; color:#ffffff; font-family:'Space Grotesk', sans-serif;">📊 {table_name}</span>
                <span style="font-size:0.75rem; font-weight:600; background:rgba(34,197,94,0.15); color:#4ade80; border:1px solid rgba(34,197,94,0.3); padding:3px 10px; border-radius:9999px;">{len(df)} Records</span>
            </div>
            '''
            st.markdown(badge_html, unsafe_allow_html=True)
            
            # Interactive Data Editor with styled persistent toolbar
            edited_df = st.data_editor(
                df, 
                key=f"editor_{idx}", 
                num_rows="dynamic", 
                use_container_width=True,
                height=min(450, 45 + len(df) * 35)
            )
            edited_dfs[table_name] = edited_df

        # Multi-Tab Styled Excel Workbook Generator (Single Master Output)
        excel_buffer = io.BytesIO()
        seen_sheets = set()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            # If multiple pages exist, create a Master Consolidated Tab
            if len(edited_dfs) > 1:
                try:
                    combined_df = pd.concat(list(edited_dfs.values()), ignore_index=True)
                    combined_df.fillna("", inplace=True)
                    combined_df.to_excel(writer, sheet_name="Master Combined Records", index=False)
                    seen_sheets.add("Master Combined Records")
                except Exception:
                    pass
            
            for idx, (name, df) in enumerate(edited_dfs.items()):
                sheet_title = create_unique_sheet_name(name, idx, seen_sheets)
                df.to_excel(writer, sheet_name=sheet_title, index=False)
                
            format_excel_workbook(writer)
                
        excel_data = excel_buffer.getvalue()

        # Single Master Download Action Bar
        st.markdown("---")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        
        button_label = "📥 Download Excel Workbook (.xlsx)" if len(edited_dfs) == 1 else f"📥 Download Complete Multi-Sheet Excel Workbook ({len(edited_dfs)} Sheets + Master Tab)"
        
        st.download_button(
            label=button_label,
            data=excel_data,
            file_name="sheetgen_extracted_workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
