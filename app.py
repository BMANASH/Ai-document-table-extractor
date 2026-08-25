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
    page_title="DataLens AI | Enterprise Document Intelligence",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enterprise Dark Tech CSS Scaffolding
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;700&family=JetBrains+Mono:wght@400;500&display=swap');

    /* Global Theme Overrides */
    .stApp {
        background: radial-gradient(circle at 50% 0%, #111827 0%, #080b11 75%, #05070a 100%) !important;
        color: #e2e8f0;
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    /* Typography */
    h1, h2, h3 {
        font-family: 'Space Grotesk', sans-serif !important;
        letter-spacing: -0.02em;
        font-weight: 700;
    }
    
    .gradient-title {
        background: linear-gradient(135deg, #60a5fa 0%, #a855f7 50%, #ec4899 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.75rem !important;
        font-weight: 800 !important;
        line-height: 1.15;
        margin-bottom: 0.25rem;
    }

    /* Subheadings */
    .sub-heading {
        color: #94a3b8;
        font-size: 1.05rem;
        font-weight: 400;
        margin-bottom: 1.5rem;
    }

    /* Glassmorphism Feature Cards */
    .glass-card {
        background: rgba(17, 24, 39, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 1.25rem;
        backdrop-filter: blur(12px);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.35);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .glass-card:hover {
        border-color: rgba(96, 165, 250, 0.35);
        transform: translateY(-2px);
    }

    /* Badge Tags */
    .status-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.03em;
        text-transform: uppercase;
        background: rgba(59, 130, 246, 0.12);
        color: #60a5fa;
        border: 1px solid rgba(59, 130, 246, 0.3);
        margin-bottom: 1rem;
    }

    /* Sidebar Customization */
    section[data-testid="stSidebar"] {
        background-color: #0b0f19 !important;
        border-right: 1px solid rgba(255, 255, 255, 0.06);
    }

    /* Primary Action Buttons */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.65rem 1.75rem !important;
        font-size: 0.95rem !important;
        box-shadow: 0 0 20px rgba(37, 99, 235, 0.45) !important;
        transition: all 0.25s ease-in-out !important;
    }
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 0 28px rgba(37, 99, 235, 0.7) !important;
    }

    /* File Uploader Container */
    div[data-testid="stFileUploader"] {
        background: rgba(17, 24, 39, 0.5) !important;
        border: 1px dashed rgba(255, 255, 255, 0.15) !important;
        border-radius: 14px !important;
        padding: 1.5rem !important;
        transition: border-color 0.2s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #60a5fa !important;
    }

    /* Code & Metrics Blocks */
    code {
        font-family: 'JetBrains Mono', monospace !important;
        background: rgba(255, 255, 255, 0.06) !important;
        color: #38bdf8 !important;
        border-radius: 4px !important;
        padding: 2px 6px !important;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar Configuration
with st.sidebar:
    st.markdown("### ⚡ Neural Core Config")
    
    # Check if API Key is configured via Streamlit Secrets
    default_key = os.environ.get("GEMINI_API_KEY", "")
    if not default_key and "GEMINI_API_KEY" in st.secrets:
        default_key = st.secrets["GEMINI_API_KEY"]

    if default_key:
        api_key = default_key
        st.success("🔒 System API Key Active", icon="✅")
    else:
        api_key = st.text_input("Gemini API Key", type="password", placeholder="Enter AI Studio Key...", help="Obtain a free API key at aistudio.google.com")
        if not api_key:
            st.caption("💡 Enter your key above to enable vision synthesis.")

    st.markdown("---")
    st.markdown("""
    **Core Capabilities:**
    * 📑 **Multi-Page Vector Parsing:** Ingests long corporate statements & PDFs.
    * 🗃️ **Multi-Table Detection:** Isolates distinct tables into separate tabs.
    * 🧹 **Automated Sanitization:** Cleans currencies, footnotes, and whitespace.
    * 📊 **Live Editable Grid:** Modify numbers in-browser before export.
    * 📥 **Native Excel Generator:** Download `.xlsx` with formula-ready floats.
    """)

# Hero Header
st.markdown('<div class="status-badge">⚡ Powered by Gemini 2.5 Flash Vision</div>', unsafe_allow_html=True)
st.markdown('<div class="gradient-title">DataLens AI Studio</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-heading">Autonomous Document-to-Data Pipeline & Executive Tabular Synthesizer</div>', unsafe_allow_html=True)

# 3 Feature Spec Cards
col_card1, col_card2, col_card3 = st.columns(3)
with col_card1:
    st.markdown("""
    <div class="glass-card">
        <div style="font-size: 1.25rem; font-weight:700; color:#60a5fa; margin-bottom:4px;">01. Ingestion</div>
        <div style="font-size:0.85rem; color:#94a3b8;">High-resolution scan & vector PDF parsing with table boundary detection.</div>
    </div>
    """, unsafe_allow_html=True)
with col_card2:
    st.markdown("""
    <div class="glass-card">
        <div style="font-size: 1.25rem; font-weight:700; color:#a855f7; margin-bottom:4px;">02. Sanitization</div>
        <div style="font-size:0.85rem; color:#94a3b8;">Removes noisy symbols, fixes broken layouts, and aligns column types.</div>
    </div>
    """, unsafe_allow_html=True)
with col_card3:
    st.markdown("""
    <div class="glass-card">
        <div style="font-size: 1.25rem; font-weight:700; color:#ec4899; margin-bottom:4px;">03. Multi-Tab Export</div>
        <div style="font-size:0.85rem; color:#94a3b8;">Outputs clean, multi-sheet Excel (.xlsx) workbooks or flat CSV files.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Extraction Logic
def process_document(file_bytes, mime_type, key):
    client = genai.Client(api_key=key)
    prompt = """
    You are an expert Data Engineer and Senior Financial Analyst.
    Analyze the uploaded document or image carefully:
    
    1. EXTRACT ALL DISTINCT TABLES:
       - Identify each separate table clearly (e.g. Income Statement, Segment Revenue, Operational Breakdown).
       - Sanitize numbers: remove stray footnote symbols, convert string representations to clean floats or integers where appropriate.
       - Provide standard, concise column headers.
       
    2. EXECUTIVE BUSINESS ANALYSIS:
       - Classify the document domain (e.g., Financial Earnings, Retail Ledger, Inventory, Invoice).
       - Provide key patterns, significant variance, anomalies, and summary metrics.

    Return your response strictly as raw valid JSON without markdown fences:
    {
      "analysis": "Markdown formatted executive summary and pattern breakdown.",
      "tables": [
        {
          "table_name": "Table Name",
          "headers": ["Col 1", "Col 2", "Col 3"],
          "rows": [
            ["Val 1", "Val 2", "Val 3"],
            ["Val 4", "Val 5", "Val 6"]
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
uploaded_file = st.file_uploader("Drop PDF document or scanned image here", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file and not api_key:
    st.warning("⚠️ Please provide a Gemini API Key in the left sidebar to initialize the AI engine.")

if uploaded_file and api_key:
    file_type = uploaded_file.type
    file_bytes = uploaded_file.read()
    
    col_prev, col_action = st.columns([1, 2])
    with col_prev:
        st.subheader("📄 Document Preview")
        if "image" in file_type:
            img = Image.open(io.BytesIO(file_bytes))
            st.image(img, use_container_width=True)
        elif file_type == "application/pdf":
            st.info(f"Loaded PDF: **{uploaded_file.name}** ({len(file_bytes)/1024:.1f} KB)")
            
    with col_action:
        st.subheader("⚡ Execute Pipeline")
        st.caption("Trigger optical document analysis, table segmentation, and value normalization.")
        if st.button("🚀 Extract & Synthesize Data", type="primary", use_container_width=True):
            with st.spinner("Processing document matrix and synthesizing tables..."):
                try:
                    raw_response = process_document(file_bytes, file_type, api_key)
                    data = json.loads(raw_response)
                    st.session_state["extracted_data"] = data
                    st.toast("Data successfully synthesized!", icon="✨")
                except Exception as e:
                    st.error(f"Processing Error: {e}")

# Display Results & Interactive Data Editor
if "extracted_data" in st.session_state:
    data = st.session_state["extracted_data"]
    
    st.markdown("---")
    st.subheader("📊 Executive Business Analysis")
    st.markdown(data.get("analysis", "No commentary generated."))
    
    tables = data.get("tables", [])
    if not tables:
        st.warning("No structured tabular data detected in the provided file.")
    else:
        st.markdown("---")
        st.subheader("📝 Live Interactive Data Grid")
        st.caption("Double-click any cell to adjust values or correct typos prior to downloading.")
        
        edited_dfs = {}
        for idx, tbl in enumerate(tables):
            table_name = tbl.get("table_name", f"Table_{idx+1}")
            headers = tbl.get("headers", [])
            rows = tbl.get("rows", [])
            
            df = pd.DataFrame(rows, columns=headers if headers else None)
            
            st.markdown(f"#### 📑 {table_name}")
            edited_df = st.data_editor(df, key=f"editor_{idx}", num_rows="dynamic", use_container_width=True)
            edited_dfs[table_name] = edited_df

        # Generate Multi-Tab Excel Workbook
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
                label="📥 Download Excel Workbook (.xlsx)",
                data=excel_data,
                file_name="datalens_extracted_tables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with c2:
            if len(edited_dfs) == 1:
                first_df = list(edited_dfs.values())[0]
                csv_data = first_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download CSV (.csv)",
                    data=csv_data,
                    file_name="datalens_extracted_table.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.info("ℹ️ Multiple tables detected: Download as Excel to preserve multi-sheet organization.")
