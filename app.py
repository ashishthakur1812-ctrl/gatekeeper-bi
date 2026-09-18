import streamlit as st
import pandas as pd
import subprocess
import os
import glob
import time
import tempfile
import io
import sys

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Gatekeeper BI | Enterprise Suite",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Gatekeeper BI: Executive Dashboard Engine")
st.caption("Universal 2FA Autonomous Reporting Pipeline (v71.0 Enterprise Master)")

# --- SESSION STATE INITIALIZATION ---
if "uploaded_data" not in st.session_state:
    st.session_state["uploaded_data"] = None

# --- 3-WAY UNIVERSAL INPUT ENGINE ---
st.write("### 📂 Input Business Data")
tab_upload, tab_paste, tab_demo = st.tabs([
    "📁 File Upload (Desktop/iOS)",
    "📋 Paste CSV (Mobile Safe)",
    "🧪 1-Click Demo Data"
])

up_file = None

with tab_upload:
    up_file = st.file_uploader("Upload raw business dataset (.csv or .xlsx)", key="main_file_uploader")
    if up_file is not None:
        if up_file.name.lower().endswith(('.csv', '.xlsx')):
            st.session_state["uploaded_data"] = up_file
        else:
            st.error("⚠️ Invalid format! Please upload only .csv or .xlsx")

with tab_paste:
    st.info("💡 Note: Direct pasting only supports plain text (.csv). For Excel (.xlsx), use the Upload tab.")
    pasted_text = st.text_area("Paste raw CSV text directly here (Recommended for Android users)", height=150)
    if st.button("📥 Load Pasted Data"):
        if pasted_text.strip():
            f = io.BytesIO(pasted_text.encode('utf-8'))
            f.name = "pasted_data.csv"
            st.session_state["uploaded_data"] = f
            st.success("Data successfully loaded!")
        else:
            st.warning("Please paste valid CSV data first.")

with tab_demo:
    if st.button("🚀 Load Enterprise Sample (Instant Mobile Test)"):
        demo_csv = (
            "Transaction_ID,Date,Customer,Amount,Status,Region\n"
            "TXN1001,2026-09-01,Aarav Sharma,15500,Completed,North\n"
            "TXN1002,2026-09-02,Priya Patel,8200,Pending,West\n"
            "TXN1003,2026-09-03,Rohan Verma,45000,Completed,South\n"
            "TXN1004,2026-09-04,Sneha Rao,-1200,Failed,East\n"
            "TXN1005,2026-09-05,Vikas Gupta,23400,Completed,North\n"
        )
        f = io.BytesIO(demo_csv.encode('utf-8'))
        f.name = "demo_dataset.csv"
        st.session_state["uploaded_data"] = f
        st.success("Enterprise demo data ready!")

uploaded_file = up_file if up_file is not None else st.session_state.get("uploaded_data")

if uploaded_file is None:
    st.info("👆 Please upload a file, paste CSV data, or load demo data above to proceed.")
    st.stop()

# --- PREVIEW & VALIDATION PIPELINE ---
uploaded_file.seek(0)
st.write("---")

try:
    if uploaded_file.name.lower().endswith('.csv'):
        df_preview = pd.read_csv(uploaded_file)
    else:
        xl = pd.ExcelFile(uploaded_file)
        target_sheet = 'Cleaned_Data' if 'Cleaned_Data' in xl.sheet_names else xl.sheet_names[0]
        df_preview = xl.parse(target_sheet)
except Exception as e:
    st.error(f"File read error: {e}")
    st.stop()

sample_cols = [str(c).lower() for c in df_preview.columns]
if any("filter" in c or "dashboard" in c for c in sample_cols) or (df_preview.isnull().sum().sum() / (df_preview.size or 1)) > 0.8:
    st.error("⚠️ Invalid Raw Data: Raw transactional dataset upload karein.")
    st.stop()

# --- 3-CARD ENTERPRISE HEALTH METRICS ---
st.subheader("📊 Dataset Overview & Preview")
total_cells = df_preview.size
null_cells = df_preview.isnull().sum().sum()
clean_ratio = ((total_cells - null_cells) / total_cells * 100) if total_cells > 0 else 0

col1, col2, col3 = st.columns(3)
col1.metric("Total Records", f"{len(df_preview):,}")
col2.metric("Total Features / Columns", f"{len(df_preview.columns):,}")
col3.metric("Data Health / Integrity", f"{clean_ratio:.1f}%", delta=f"{clean_ratio:.1f}% Clean")

st.dataframe(df_preview.head(10), use_container_width=True)
st.write("---")

# --- COMPILATION & EXECUTION ENGINE ---
if st.button("🚀 Compile Full Enterprise Audit Suite", type="primary"):
    with st.spinner("Executing Autonomous Pipeline (v71.0 Enterprise Master)..."):
        try:
            file_ext = ".csv" if uploaded_file.name.lower().endswith(".csv") else ".xlsx"
            temp_input = f"input_run{file_ext}"
            
            uploaded_file.seek(0)
            with open(temp_input, "wb") as f_out:
                f_out.write(uploaded_file.read())
                f_out.flush()

            cmd = [sys.executable, "pipeline_v71_DYNAMIC.py", temp_input]
            process_res = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd(), timeout=40)

            if process_res.returncode != 0:
                st.error("❌ Compilation Pipeline Failed:")
                st.code(f"--- STDOUT ---\n{process_res.stdout}\n\n--- STDERR ---\n{process_res.stderr}")
                st.stop()

            # Retrieve rich dashboard from reports/ directory
            generated_excel = glob.glob(os.path.join("reports", "*Dashboard*.xlsx")) or glob.glob("*.xlsx")
            generated_csv = glob.glob(os.path.join("quarantine", "*.csv")) or glob.glob("*.csv")
            generated_parquet = glob.glob(os.path.join("clean_data", "*.parquet")) or glob.glob("*.parquet")

            if not generated_excel:
                st.error("⚠️ Dashboard output not found.")
                st.code(process_res.stdout)
                st.stop()

            st.success("✅ Complete Executive Dashboard Suite compiled successfully!")

            d_col1, d_col2, d_col3 = st.columns(3)

            with open(generated_excel[0], "rb") as ef:
                d_col1.download_button(
                    label="📥 Download Executive Suite (.xlsx)",
                    data=ef.read(),
                    file_name=os.path.basename(generated_excel[0]),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if generated_csv:
                q_xlsx_file = generated_csv[0].replace('.csv', '.xlsx')
                target_file = q_xlsx_file if os.path.exists(q_xlsx_file) else generated_csv[0]
                is_excel = target_file.endswith('.xlsx')
        
                with open(target_file, "rb") as cf:
                    d_col2.download_button(
                        label=f"🛡️ Download Quarantine Audit ({'.xlsx' if is_excel else '.csv'})",
                        data=cf.read(),
                        file_name=os.path.basename(target_file),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if is_excel else "text/csv"
                )

            if generated_parquet:
                with open(generated_parquet[0], "rb") as pf:
                    d_col3.download_button(
                        label="⚡ Download Parquet Mirror (.parquet)",
                        data=pf.read(),
                        file_name=os.path.basename(generated_parquet[0]),
                        mime="application/octet-stream"
                    )
        except subprocess.TimeoutExpired:
            st.error("⏱️ Pipeline execution timed out!")
        except Exception as ex:
            st.error(f"⚠️ Execution Error: {ex}")
