import streamlit as st
import pandas as pd
import io
import json
import os
from PIL import Image
from google import genai
from google.genai import types

# Page Config
st.set_page_config(
    page_title="SheetGen AI | Document to Excel Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Fetch API Key silently from Streamlit Secrets or Environment
api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

# Modern Dark Theme Styling & SVG Badges
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Space+Grotesk:wght@600;700&display=swap');

    .stApp {
        background: radial-gradient(circle at 50% 0%, #111827 0%, #080b11 75%, #05070a 100%) !important;
        color: #e2e8f0;
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    h1, h2, h3 {
        font-family: 'Space Grotesk', sans-serif !important;
        letter-spacing: -0.02em;
        font-weight: 700;
    }
    
    .header-container {
        display: flex;
        align-items: center;
        gap: 14px;
        margin-top: 0.5rem;
        margin-bottom: 0.25rem;
    }

    .excel-title {
        background: linear-gradient(135deg, #22c55e 0%, #10b981 50%, #38bdf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.75rem !important;
        font-weight: 800 !important;
        line-height: 1.15;
    }

    .sub-heading {
        color: #94a3b8;
        font-size: 1.05rem;
        font-weight: 400;
        margin-bottom: 1.5rem;
    }

    .feature-card {
        background: rgba(17, 24, 39, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 1.25rem;
        backdrop-filter: blur(12px);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.35);
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

    section[data-testid="stSidebar"] {
        background-color: #0b0f19 !important;
        border-right: 1px solid rgba(255, 255, 255, 0.06);
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

# Vector Excel Logo SVG
EXCEL_LOGO_SVG = """
<svg width="44" height="44" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect width="48" height="48" rx="10" fill="#107C41"/>
    <path d="M15.5 33L21.8 24L16.2 15H21.2L24.3 20.8L27.4 15H32.4L26.7 24L33 33H28L24.2 26.8L20.4 33H15.5Z" fill="white"/>
</svg>
"""

SIDEBAR_EXCEL_SVG = """
<svg width="28" height="28" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align: middle; margin-right: 8px;">
    <rect width="48" height="48" rx="8" fill="#107C41"/>
    <path d="M15.5 33L21.8 24L16.2 15H21.2L24.3 20.8L27.4 15H32.4L26.7 24L33 33H28L24.2 26.8L20.4 33H15.5Z" fill="white"/>
</svg>
"""

# Clean Sidebar
with st.sidebar:
    st.markdown(f'<div style="display:flex; align-items:center; gap:8px;">{SIDEBAR_EXCEL_SVG} <h3 style="margin:0; font-size:1.4rem;">SheetGen AI</h3></div>', unsafe_allow_html=True)
    st.caption("Automated Tabular Data Extraction Engine")
    st.markdown("---")
    st.markdown("""
    **Core Capabilities:**
    * 📑 **Reads PDFs & Images:** Scans reports, invoices, and statements.
    * 🗃️ **Finds Multiple Tables:** Separates different tables cleanly.
    * 🧹 **Cleans Up Numbers:** Removes symbols so formulas work.
    * ✏️ **Lets You Edit:** Double-click cells to fix typos in-browser.
    * 📥 **Exports to Excel:** Downloads clean `.xlsx` spreadsheets.
    """)
    st.markdown("---")
    st.markdown("🟢 **System Status:** Ready")

# Hero Header with Vector Excel Icon
st.markdown('<div class="status-badge">⚡ Instant Document to Spreadsheet</div>', unsafe_allow_html=True)
st.markdown(f'''
<div class="header-container">
    {EXCEL_LOGO_SVG}
    <div class="excel-title">SheetGen AI</div>
</div>
''', unsafe_allow_html=True)
st.markdown('<div class="sub-heading">Upload any PDF or image table → Automatically convert it to a clean, editable Excel file.</div>', unsafe_allow_html=True)

# 3 Step Visual Workflow Cards
col_card1, col_card2, col_card3 = st.columns(3)
with col_card1:
    st.markdown("""
    <div class="feature-card">
        <div style="font-size: 1.15rem; font-weight:700; color:#4ade80; margin-bottom:4px;">1. Upload File</div>
        <div style="font-size:0.85rem; color:#94a3b8;">Drop scanned photos or multi-page PDFs with data tables.</div>
    </div>
    """, unsafe_allow_html=True)
with col_card2:
    st.markdown("""
    <div class="feature-card">
        <div style="font-size: 1.15rem; font-weight:700; color:#38bdf8; margin-bottom:4px;">2. AI Cleans & Formats</div>
        <div style="font-size:0.85rem; color:#94a3b8;">AI extracts rows, columns, and fixes messy numbers automatically.</div>
    </div>
    """, unsafe_allow_html=True)
with col_card3:
    st.markdown("""
    <div class="feature-card">
        <div style="font-size: 1.15rem; font-weight:700; color:#a78bfa; margin-bottom:4px;">3. Download Excel (.xlsx)</div>
        <div style="font-size:0.85rem; color:#94a3b8;">Get a multi-sheet Excel file ready for formulas and charts.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Helper Function: Process via Gemini Vision
def process_document(file_bytes, mime_type, key):
    client = genai.Client(api_key=key)
    prompt = """
    You are an expert Data Specialist.
    Analyze the uploaded document or image carefully:
    
    1. EXTRACT ALL DISTINCT TABLES:
       - Identify each separate table clearly (e.g. Table 1, Table 2).
       - Clean numbers: remove currency symbols and footnote markers so Excel can calculate sums directly.
       - Provide standard column headers.
       
    2. SUMMARY & PATTERNS:
       - What kind of data is this (Financials, Sales, Inventory, Invoice)?
       - Write a brief, easy-to-read summary with key insights and trends.

    Return your response strictly as valid JSON:
    {
      "analysis": "Short Markdown summary of key business insights.",
      "tables": [
        {
          "table_name": "Table Name",
          "headers": ["Header 1", "Header 2", "Header 3"],
          "rows": [
            ["Row1 Col1", "Row1 Col2", "Row1 Col3"],
            ["Row2 Col1", "Row2 Col2", "Row2 Col3"]
          ]
        }
      ]
    }
    """
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=[
            types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
            prompt
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json"
        )
    )
    return response.text

# Document Upload Section
uploaded_file = st.file_uploader("Drop your PDF document or image here", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    if not api_key:
        st.error("⚠️ System Error: API Key not configured in Streamlit Secrets. Please add GEMINI_API_KEY to your app secrets.")
    else:
        file_type = uploaded_file.type
        file_bytes = uploaded_file.read()
        
        col_prev, col_action = st.columns([1, 2])
        with col_prev:
            st.subheader("📄 Uploaded File Preview")
            if "image" in file_type:
                img = Image.open(io.BytesIO(file_bytes))
                st.image(img, use_container_width=True)
            elif file_type == "application/pdf":
                st.info(f"Loaded PDF: **{uploaded_file.name}** ({len(file_bytes)/1024:.1f} KB)")
                
        with col_action:
            st.subheader("⚡ Convert to Excel")
            st.caption("Click below to extract all tables and generate your spreadsheet.")
            if st.button("🚀 Extract Tables & Convert to Excel", type="primary", use_container_width=True):
                with st.spinner("Reading document and generating Excel sheets..."):
                    try:
                        raw_response = process_document(file_bytes, file_type, api_key)
                        data = json.loads(raw_response)
                        st.session_state["extracted_data"] = data
                        st.toast("Extraction complete!", icon="✅")
                    except Exception as e:
                        st.error(f"Processing Error: {e}")

# Results Display & Editable Table Grid
if "extracted_data" in st.session_state:
    data = st.session_state["extracted_data"]
    
    st.markdown("---")
    st.subheader("💡 Key Summary & Insights")
    st.markdown(data.get("analysis", "No summary provided."))
    
    tables = data.get("tables", [])
    if not tables:
        st.warning("No tables found in this file.")
    else:
        st.markdown("---")
        st.subheader("✏️ Review & Edit Tables")
        st.caption("You can double-click any cell below to fix any typos before downloading.")
        
        edited_dfs = {}
        for idx, tbl in enumerate(tables):
            table_name = tbl.get("table_name", f"Table_{idx+1}")
            headers = tbl.get("headers", [])
            rows = tbl.get("rows", [])
            
            df = pd.DataFrame(rows, columns=headers if headers else None)
            
            st.markdown(f"#### 📊 {table_name}")
            edited_df = st.data_editor(df, key=f"editor_{idx}", num_rows="dynamic", use_container_width=True)
            edited_dfs[table_name] = edited_df

        # Multi-Tab Excel Generator
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for name, df in edited_dfs.items():
                clean_sheet_name = "".join(c for c in name if c.isalnum() or c in (' ', '_'))[:30]
                df.to_excel(writer, sheet_name=clean_sheet_name or f"Sheet_{idx+1}", index=False)
        excel_data = excel_buffer.getvalue()

        st.markdown("---")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="📥 Download Excel File (.xlsx)",
                data=excel_data,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with c2:
            if len(edited_dfs) == 1:
                first_df = list(edited_dfs.values())[0]
                csv_data = first_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download CSV File (.csv)",
                    data=csv_data,
                    file_name="extracted_table.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.info("ℹ️ Multiple tables detected: Use Excel download to keep them in separate sheets.")
