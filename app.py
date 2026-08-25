import streamlit as st
import pandas as pd
import io
import json
from PIL import Image
from google import genai
from google.genai import types

# Page Configuration
st.set_page_config(
    page_title="DataLens AI | Enterprise Table Extractor",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling: Modern Black Tech Theme & Font Stack
st.markdown("""
<style>
    @import url('https://fonts.cdnfonts.com/css/nexa-pro');
    
    /* Main Background & Base Styling */
    .stApp {
        background-color: #0b0f17;
        color: #e2e8f0;
        font-family: 'Chopin Regular', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }
    
    /* Headers Styling */
    h1, h2, h3 {
        font-family: 'Nexa Pro', sans-serif !important;
        font-weight: 700 !important;
        letter-spacing: 0.5px;
        color: #ffffff !important;
    }
    
    /* Subheadings & Section Titles */
    h4, h5, h6, .stSubheader {
        font-family: 'Fabric Grotesk Variable Regular', 'Inter', sans-serif !important;
        font-weight: 600 !important;
        color: #94a3b8 !important;
    }
    
    /* Cards & Containers */
    div[data-testid="stVerticalBlock"] > div[style*="background-color"] {
        background: #111827 !important;
        border: 1px solid #1f2937 !important;
        border-radius: 12px;
        padding: 18px;
    }
    
    /* Primary Buttons */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb, #1d4ed8);
        color: #ffffff;
        border: 1px solid #3b82f6;
        border-radius: 8px;
        font-family: 'Nexa Pro', sans-serif;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #1d4ed8, #1e40af);
        border-color: #60a5fa;
        box-shadow: 0 0 15px rgba(59, 130, 246, 0.4);
    }
    
    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #0d121d;
        border-right: 1px solid #1e293b;
    }
    
    /* File Uploader Box */
    div[data-testid="stFileUploader"] {
        background: #111827;
        border: 1px dashed #374151;
        border-radius: 10px;
        padding: 12px;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar Configuration
with st.sidebar:
    st.markdown("### ⚙️ System Config")
    api_key = st.text_input("Gemini API Key", type="password", help="Enter your Gemini API key from Google AI Studio.")
    st.markdown("---")
    st.markdown("""
    **Supported Formats:**
    * 📑 Multi-page PDFs (`.pdf`)
    * 🖼️ High-Res Images (`.png`, `.jpg`, `.jpeg`)
    
    **Features:**
    * ⚡ Multi-table detection per page
    * 🧹 Auto-sanitization & number parsing
    * 📊 Interactive cell-level data editor
    * 📥 Multi-tab Excel (`.xlsx`) export
    """)

# Main UI Header
st.title("⚡ DataLens AI")
st.caption("Next-Gen Document Intelligence & Tabular Data Synthesizer")

# Helper function to call Gemini Multimodal
def process_document(file_bytes, mime_type, user_key):
    client = genai.Client(api_key=user_key)
    
    prompt = """
    You are an expert Data Engineer and Senior Business Analyst.
    Analyze the uploaded document/image carefully:
    
    1. EXTRACT ALL DISTINCT TABLES:
       - Detect all separate tables (e.g. Table 1, Table 2).
       - Sanitize and clean raw data: format numbers cleanly (e.g., standard floats/integers without stray symbols or erratic footnotes) so Excel formulas compute immediately.
       - Provide standard, concise column headers.
       
    2. BUSINESS PATTERN & INSIGHT ANALYSIS:
       - Identify the exact document domain (e.g., Corporate Earnings, Retail P&L, Inventory Ledger, Balance Sheet).
       - Provide executive-level commentary: key performance trends, notable ratios, and anomalies.

    Return your output STRICTLY as valid JSON without extra markdown fences:
    {
      "analysis": "Executive Markdown summary detailing domain classification, trends, and analytical observations.",
      "tables": [
        {
          "table_name": "Concise Table Name",
          "headers": ["Header1", "Header2", "Header3"],
          "rows": [
            ["Val1", "Val2", "Val3"],
            ["Val4", "Val5", "Val6"]
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

# Upload Section
uploaded_file = st.file_uploader("Upload document or scan for automated extraction", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file and not api_key:
    st.warning("🔑 Please enter your Gemini API Key in the left sidebar to initialize extraction.")

if uploaded_file and api_key:
    file_type = uploaded_file.type
    file_bytes = uploaded_file.read()
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("📄 Source Preview")
        if "image" in file_type:
            img = Image.open(io.BytesIO(file_bytes))
            st.image(img, use_container_width=True)
        elif file_type == "application/pdf":
            st.info(f"Loaded PDF Document: **{uploaded_file.name}** ({len(file_bytes)/1024:.1f} KB)")
            
    if st.button("⚡ Synthesize & Extract Tables", type="primary"):
        with st.spinner("Processing document topology and parsing data matrix..."):
            try:
                raw_response = process_document(file_bytes, file_type, api_key)
                data = json.loads(raw_response)
                st.session_state["extracted_data"] = data
                st.success("Extraction and normalization complete.")
            except Exception as e:
                st.error(f"Processing error: {e}")

# Results Display
if "extracted_data" in st.session_state:
    data = st.session_state["extracted_data"]
    
    st.markdown("---")
    st.subheader("📊 Executive Analysis & Intelligence")
    st.markdown(data.get("analysis", "No commentary generated."))
    
    tables = data.get("tables", [])
    
    if not tables:
        st.warning("No structured tabular data detected.")
    else:
        st.markdown("---")
        st.subheader("📝 Live Interactive Data Grid")
        st.caption("Double-click any cell to modify values prior to exporting.")
        
        edited_dfs = {}
        for idx, tbl in enumerate(tables):
            table_name = tbl.get("table_name", f"Table_{idx+1}")
            headers = tbl.get("headers", [])
            rows = tbl.get("rows", [])
            
            df = pd.DataFrame(rows, columns=headers if headers else None)
            
            st.markdown(f"#### 📑 {table_name}")
            edited_df = st.data_editor(df, key=f"editor_{idx}", num_rows="dynamic", use_container_width=True)
            edited_dfs[table_name] = edited_df

        # Multi-Tab Excel Generation
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for name, df in edited_dfs.items():
                clean_sheet_name = "".join(c for c in name if c.isalnum() or c in (' ', '_'))[:30]
                df.to_excel(writer, sheet_name=clean_sheet_name or f"Sheet{idx+1}", index=False)
        
        excel_data = excel_buffer.getvalue()

        st.markdown("---")
        col_down1, col_down2 = st.columns(2)
        with col_down1:
            st.download_button(
                label="📥 Export Excel Workbook (.xlsx)",
                data=excel_data,
                file_name="datalens_extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with col_down2:
            if len(edited_dfs) == 1:
                first_df = list(edited_dfs.values())[0]
                csv_data = first_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Export Flat File (.csv)",
                    data=csv_data,
                    file_name="datalens_extracted_data.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.info("ℹ️ Multiple tables detected: Exporting to Excel will preserve separate sheets.")
