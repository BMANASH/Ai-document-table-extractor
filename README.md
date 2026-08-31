# SheetGen AI — Universal Document to Excel & AI Dashboard Suite

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://sheetgen-analytics-suite.streamlit.app/)

🔗 **Live Web Application**: [https://sheetgen-analytics-suite.streamlit.app/](https://sheetgen-analytics-suite.streamlit.app/)

---

SheetGen AI is an intelligent data extraction and Business Intelligence (BI) tool that instantly transforms single or multi-page documents (Images & PDFs) into structured, editable Excel spreadsheets (`.xlsx`) and interactive visual dashboards.

Powered by multimodal vision AI, SheetGen AI handles printed, digital, and handwritten tabular documents—such as invoices, ledgers, rosters, attendance sheets, and receipts—with high accuracy.

---

## Core Features

### 1. Universal OCR Vision Extraction
* **Multi-Format Ingestion**: Supports PNG, JPG, JPEG, and PDF documents.
* **Handwriting & Scan Recognition**: Accurately isolates tabular rows, sanitizes noisy text, and detects authentic column headers.
* **Domain Agnostic**: Works seamlessly across receipts, payroll sheets, inventory registers, and financial invoices.

### 2. Mode 1: In-Browser Data Grid & Editor
* **Interactive Editing**: Double-click any table cell directly in the browser to modify or correct values.
* **Toolbar Controls**: Search within records , hide/view specific columns , and dynamically add new rows (`+`).
* **Clean Base Export**: Download the cleaned, formatted data as a standalone multi-sheet `.xlsx` file.

### 3. Mode 2: Executive Dashboard Engine (Live Formulas)
* **Instant KPI Metrics**: Automatically computes key summaries such as Total Record Count, Unique Classifications, and Dominant Categories.
* **Interactive Web Visualizations**: Renders dynamic dark-glass Plotly bar charts and doughnut charts inside the app.
* **Formula-Linked Excel Export**: Downloads a ready-to-present executive Excel workbook where summary tables and charts are powered by native live Excel formulas (`=COUNTA()`, `=COUNTIF()`, `=SUM()`).

### 4. Mode 3: Talk with AI (Custom BI Copilot)
* **Natural Language Chart Generation**: Ask the AI Copilot to build custom charts and metrics on demand (e.g., *"Show a breakdown of status flags by region"*).
* **Pre-Built Quick Prompts**: One-click action chips to automatically audit exceptions, identify top contributors, or summarize categories.
* **Custom Excel Synthesis**: Generates a tailored Excel tab reflecting your custom copilot charts and calculations.

### 5. Live Diagnostics & Transparency
* **Active Model Tracking**: Displays the exact Gemini vision engine processing your data.
* **Real-Time Latency Stopwatch**: Measures the precise end-to-end extraction duration per batch.

---

## Supported Batch Upload Formats & Limits

To ensure optimal extraction accuracy and processing speed:

| Document Type | Recommended Batch Limit | File Size Constraint |
| :--- | :--- | :--- |
| **Images / Photos** (`PNG`, `JPG`, `JPEG`) | Max **6 pages / photos** per batch | Total batch under **20 MB** |
| **PDF Documents** (`PDF`) | Max **2 files** per batch | Total batch under **20 MB** |

> **Note**: For best results, please upload either Images or PDFs in a single batch, rather than mixing both formats together.

---

## Tech Stack

* **Frontend / Framework**: [Streamlit](https://streamlit.io/)
* **AI Engine**: [Google Generative AI](https://ai.google.dev/) (`gemini-3.5-flash`)
* **Data Processing**: [Pandas](https://pandas.pydata.org/), [Pillow (PIL)](https://python-pillow.org/), [pypdf](https://pypdf.readthedocs.io/)
* **Excel Generation & Formulas**: [openpyxl](https://openpyxl.readthedocs.io/)
* **Interactive Charts**: [Plotly Express](https://plotly.com/python/)

---

## Quickstart & Local Installation

### 1. Clone the Repository
```bash
git clone https://github.com/BMANASH/SheetGen-Analytics-Suite.git
cd SheetGen-Analytics-Suite
