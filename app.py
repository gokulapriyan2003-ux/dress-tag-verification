# Dress Tag & Master Sheet Verifier Web App (v3.4)
import streamlit as st
import pandas as pd
import openpyxl
import os
import sys
import io
import re
import tempfile
import importlib
import urllib.request
from openpyxl.styles import PatternFill

# Import the core logic from compare_tags.py with forced module reload
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import compare_tags
importlib.reload(compare_tags)
from compare_tags import (
    extract_pdf_tags,
    extract_excel_master,
    compare,
    get_updated_mrp
)


@st.cache_data(ttl=600, show_spinner="Fetching latest MRP Google Sheet...")
def load_cached_gsheet(url: str, local_path: str) -> dict:
    dfs = {}
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as response:
            data = response.read()
            xls = pd.ExcelFile(io.BytesIO(data))
            for name in xls.sheet_names:
                dfs[name] = pd.read_excel(xls, sheet_name=name)
            try:
                with open(local_path, "wb") as f:
                    f.write(data)
            except Exception:
                pass
    except Exception:
        if os.path.exists(local_path):
            try:
                xls = pd.ExcelFile(local_path)
                for name in xls.sheet_names:
                    dfs[name] = pd.read_excel(xls, sheet_name=name)
            except Exception:
                pass
    return dfs

st.set_page_config(
    page_title="Dress Tag Verifier",
    layout="wide"
)

# Custom Styling for premium aesthetics
st.markdown("""
    <style>
    .main-title {
        font-size: 2.8rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.1rem;
    }
    .subtitle {
        font-size: 1.2rem;
        color: #4B5563;
        margin-bottom: 2rem;
    }
    .metric-box {
        background-color: #F3F4F6;
        padding: 1.5rem;
        border-radius: 0.75rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #2563EB;
    }
    .metric-label {
        font-size: 0.9rem;
        color: #6B7280;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Dress Tag & Master Sheet Verifier</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Extract SKU fields from multi-tag PDF and validate them against Excel & Google Sheet references</div>', unsafe_allow_html=True)
st.caption("⚡ Engine v3.4: Multi-User Portal Active (Auto Tag-Detection & In-Memory Isolation)")

script_dir = os.path.dirname(os.path.abspath(__file__))

# Layout: Sidebar configuration
st.sidebar.header("Portal Configuration")
st.sidebar.markdown("""
**Verification Workflow:**
1. Upload the **Tag PDF**
2. Upload the **Master Excel Sheet**
3. Verify or adjust the **Tag Mode**
4. Click **Run Verification**
""")
sheet_name = st.sidebar.text_input("Excel Sheet Name (Optional, uses first sheet if blank)", value="")
st.sidebar.markdown("---")
if st.sidebar.button("🔄 Clear Cache"):
    st.cache_data.clear()
    st.sidebar.success("Cache cleared! Next run will fetch fresh Google Sheet data.")

# Step 1: Reference File Uploaders
st.subheader("1. Upload Reference Files")
col1, col2 = st.columns(2)

with col1:
    pdf_file = st.file_uploader("Upload Tag PDF", type=["pdf"])
with col2:
    xlsx_file = st.file_uploader("Upload Master Excel", type=["xlsx"])

# Auto-detect Tag Verification Mode from PDF content
auto_tag_type = "D2C Dress tag file"
if pdf_file is not None:
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_file.getvalue())) as p:
            if len(p.pages) > 0:
                p_text = (p.pages[0].extract_text() or "").upper()
                if any(k in p_text for k in ["OUTER BOX", "BOX STICKER", "SERIALISED"]):
                    auto_tag_type = "B2B Box Sticker tag file"
                elif "BUNDLE" in p_text:
                    auto_tag_type = "B2B Bundle Sticker tag file"
    except Exception:
        pass

# Step 2: Tag Verification Mode Selection
st.subheader("2. Select Tag Verification Mode")
tag_options = ["D2C Dress tag file", "B2B Box Sticker tag file", "B2B Bundle Sticker tag file"]
default_idx = tag_options.index(auto_tag_type) if auto_tag_type in tag_options else 0
tag_type = st.selectbox(
    "Tag Verification Type",
    options=tag_options,
    index=default_idx,
    help="Auto-detected from your PDF. You can change this if needed."
)

# Run Verification Button
if st.button("Run Verification", type="primary"):
    if not pdf_file or len(pdf_file.getvalue()) == 0:
        st.warning("⚠️ Please upload a valid Tag PDF file to proceed.")
    elif not xlsx_file or len(xlsx_file.getvalue()) == 0:
        st.warning("⚠️ Please upload a valid Master Excel file to proceed.")
    else:
        with st.spinner("Processing tags and master sheet..."):
            temp_pdf = None
            temp_xlsx = None
            try:
                # Use unique per-session temporary files to prevent Windows file locking and cross-user collision
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f_pdf:
                    f_pdf.write(pdf_file.getvalue())
                    temp_pdf = f_pdf.name
                # Master Sheet session-state caching (prevents re-parsing 50,000 rows on repeated verifications)
                master_cache_key = f"{xlsx_file.name}_{xlsx_file.size}_{sheet_name or 'auto'}"
                if st.session_state.get("cached_master_key") == master_cache_key and "cached_master_df" in st.session_state:
                    excel_df = st.session_state["cached_master_df"]
                else:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f_xlsx:
                        f_xlsx.write(xlsx_file.getvalue())
                        temp_xlsx = f_xlsx.name
                    excel_df = extract_excel_master(temp_xlsx, sheet_name=sheet_name if sheet_name else None)
                    st.session_state["cached_master_key"] = master_cache_key
                    st.session_state["cached_master_df"] = excel_df

                gsheet_url = "https://docs.google.com/spreadsheets/d/1Q7nboN_Rezl807J0naA0QczTyoAQ6WM-KNmp_F26n5M/export?format=xlsx"
                gsheet_path = os.path.join(script_dir, "google_sheet_mrp.xlsx")
                gsheet_dfs = load_cached_gsheet(gsheet_url, gsheet_path)

                # Load tags from PDF
                pdf_df = extract_pdf_tags(temp_pdf)
                
                if len(pdf_df) == 0:
                    st.error("❌ No tags were extracted from the uploaded PDF. Please verify that this is a valid tag sheet containing 'SKU Code:' text.")
                    st.stop()

                # Perform comparison
                report_df = compare(pdf_df, excel_df, gsheet_dfs, tag_type=tag_type)
                
                n_mismatch = (report_df["Status"] != "✅ Match").sum()
                n_total = len(report_df)
                success_rate = round(((n_total - n_mismatch) / n_total) * 100, 1) if n_total > 0 else 0
                
                # Display Metrics
                st.subheader("Verification Summary")
                m_col1, m_col2, m_col3 = st.columns(3)
                with m_col1:
                    st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-value">{len(pdf_df)}</div>
                            <div class="metric-label">Tags Extracted</div>
                        </div>
                    """, unsafe_allow_html=True)
                with m_col2:
                    st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-value">{n_total}</div>
                            <div class="metric-label">Field Checks Run</div>
                        </div>
                    """, unsafe_allow_html=True)
                with m_col3:
                    color = "#10B981" if success_rate == 100 else "#EF4444"
                    st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-value" style="color: {color}">{success_rate}%</div>
                            <div class="metric-label">Success Rate</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                # Build Comparison Report in-memory (prevents PermissionError and cross-user file locking)
                report_buf = io.BytesIO()
                pdf_skus = set(pdf_df["SKU"].dropna().astype(str).str.strip().str.upper())
                sku_cols = [c for c in excel_df.columns if any(k in str(c).lower() for k in ["sku", "item code", "gtin"])]
                if sku_cols:
                    matched_mask = excel_df[sku_cols[0]].astype(str).str.strip().str.upper().isin(pdf_skus)
                    matched_excel = excel_df[matched_mask]
                    if len(matched_excel) == 0:
                        matched_excel = excel_df.head(100)
                else:
                    matched_excel = excel_df.head(100)

                with pd.ExcelWriter(report_buf, engine="openpyxl") as writer:
                    report_df.to_excel(writer, sheet_name="Comparison_Report", index=False)
                    pdf_df.to_excel(writer, sheet_name="PDF_Extracted", index=False)
                    matched_excel.to_excel(writer, sheet_name="Excel_Matched_Master", index=False)

                    # Color the report sheet directly in-memory
                    workbook = writer.book
                    worksheet = writer.sheets["Comparison_Report"]
                    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    
                    status_col_idx = report_df.columns.get_loc("Status") + 1
                    for row_idx in range(2, len(report_df) + 2):
                        status_val = str(report_df.iloc[row_idx - 2].get("Status", ""))
                        fill = green_fill if "Match" in status_val and "Mis" not in status_val and "Not found" not in status_val else red_fill
                        for col_idx in range(1, len(report_df.columns) + 1):
                            worksheet.cell(row=row_idx, column=col_idx).fill = fill

                excel_bytes = report_buf.getvalue()

                # Optionally save a background copy to disk if not locked
                try:
                    out_path = os.path.join(script_dir, "tag_comparison_report.xlsx")
                    with open(out_path, "wb") as f_out:
                        f_out.write(excel_bytes)
                except Exception:
                    pass

                # Show results
                if n_mismatch == 0:
                    st.balloons()
                    st.success("All field checks passed successfully! EAN Barcode, Sizes, and MRP values are 100% accurate.")
                else:
                    st.error(f"Found {n_mismatch} mismatches. Please check the report or view details below.")
                    mismatch_df = report_df[report_df["Status"] != "✅ Match"]
                    st.dataframe(mismatch_df, use_container_width=True)

                # Provide Download Button from in-memory bytes
                style_codes = []
                if "Style" in pdf_df.columns:
                    style_codes = [str(x).strip() for x in pdf_df["Style"].dropna().unique() if str(x).strip()]
                elif "Lot No (Google Sheet)" in report_df.columns:
                    style_codes = [str(x).strip() for x in report_df["Lot No (Google Sheet)"].dropna().unique() if str(x).strip()]
                
                cleaned_styles = []
                for code in style_codes:
                    cleaned = re.sub(r"[\\/*?:\"<>|]", "_", code)
                    if cleaned:
                        cleaned_styles.append(cleaned)
                        
                if cleaned_styles:
                    dl_filename = f"{'_'.join(cleaned_styles)}_comparison_report.xlsx"
                else:
                    dl_filename = "tag_comparison_report.xlsx"

                st.download_button(
                    label=f"Download {dl_filename}",
                    data=excel_bytes,
                    file_name=dl_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
                # Show full comparison table
                with st.expander("View Full Comparison Report Details"):
                    st.dataframe(report_df, use_container_width=True)
                    
            except Exception as ex:
                err_str = str(ex)
                if "No /Root object" in err_str or "PdfminerException" in type(ex).__name__:
                    st.error("❌ Could not read the uploaded PDF: The file is empty (0.0B) or corrupted. Please check the PDF file on your computer, re-download it if necessary, and re-upload.")
                else:
                    st.error(f"Error during processing: {ex}")
                    st.exception(ex)
            finally:
                if temp_pdf and os.path.exists(temp_pdf):
                    try: os.unlink(temp_pdf)
                    except Exception: pass
                if temp_xlsx and os.path.exists(temp_xlsx):
                    try: os.unlink(temp_xlsx)
                    except Exception: pass
