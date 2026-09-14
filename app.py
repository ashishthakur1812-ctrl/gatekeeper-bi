import streamlit as st
import pandas as pd
import subprocess
import os
import glob
import time
import tempfile
import io
st.set_page_config(page_title="Gatekeeper BI | Enterprise Suite", page_icon="🛡️", layout="wide")

st.title("🛡️ Gatekeeper BI: Executive Dashboard Engine")
st.caption("Universal 2FA Autonomous Reporting Pipeline (v71.0 Enterprise Master)")

# --- 3-WAY UNIVERSAL INPUT ENGINE ---
if "uploaded_data" not in st.session_state:
    st.session_state["uploaded_data"] = None

st.write("### 📂 Input Business Data")
tab_upload, tab_paste, tab_demo = st.tabs(["📁 File Upload (Desktop/iOS)", "📋 Paste CSV (Mobile Safe)", "🧪 1-Click Demo Data"])

with tab_upload:
    up_file = st.file_uploader("Upload raw business dataset (.csv or .xlsx)", key="main_file_uploader")
    if up_file is not None:
        if up_file.name.lower().endswith(('.csv', '.xlsx')):
            st.session_state["uploaded_data"] = up_file
        else:
            st.error("⚠️ Invalid format! Please upload only .csv or .xlsx")

with tab_paste:
    pasted_text = st.text_area("Paste CSV text directly here (Recommended for Android users)", height=150)
    if st.button("📥 Load Pasted Data"):
        if pasted_text.strip():
            f = io.BytesIO(pasted_text.encode('utf-8'))
            f.name = "pasted_data.csv"
            st.session_state["uploaded_data"] = f
            st.success("Data successfully loaded!")
        else:
            st.warning("Please paste some CSV text first.")

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

uploaded_file = st.session_state["uploaded_data"]

if uploaded_file is not None:
    uploaded_file.seek(0)
    st.write("---")
    try:
        if uploaded_file.name.endswith('.csv'):
            df_preview = pd.read_csv(uploaded_file)
        else:
            xl = pd.ExcelFile(uploaded_file)
            target_sheet = 'Cleaned_Data' if 'Cleaned_Data' in xl.sheet_names else xl.sheet_names[0]
            df_preview = xl.parse(target_sheet)
    except Exception as e:
        st.error(f"File read error: {e}")
        st.stop()
else:
    st.stop()
    # Guard: Detect pre-generated dashboards
    sample_cols = [str(c).lower() for c in df_preview.columns]
    if any("filter" in c or "dashboard" in c for c in sample_cols) or (df_preview.isnull().sum().sum() / (df_preview.size or 1)) > 0.8:
        st.error("⚠️ Invalid Raw Data: Raw transactional dataset upload karein.")
        st.stop()

    c1, c2, c3 = st.columns(3)
    c1.metric("Raw Ingested Records", f"{len(df_preview):,}")
    c2.metric("Detected Columns", len(df_preview.columns))
    c3.metric("Data Completeness", f"{(1 - df_preview.isnull().sum().sum() / (df_preview.size or 1)):.1%}")

    with st.expander("🔍 Ingested Schema Preview", expanded=False):
        st.dataframe(df_preview.head(10), use_container_width=True)

    if st.button("🚀 Compile Full Enterprise Audit Suite", type="primary"):
        run_timestamp = time.time()
        with st.spinner("Processing Data Sanitization, Dual-Slot Routing & Compiling Artifacts..."):
            with tempfile.TemporaryDirectory() as tmp_dir:
                input_path = os.path.join(tmp_dir, uploaded_file.name)
                with open(input_path, "wb") as f:
                    uploaded_file.seek(0)
                    f.write(uploaded_file.read())

                proc = subprocess.run(
                    ["python", "pipeline_v71_DYNAMIC.py"],
                    input=f"{input_path}\n",
                    text=True,
                    capture_output=True
                )

                if proc.returncode != 0:
                    st.error("Pipeline crashed during processing. Terminal error:")
                    st.code(proc.stderr or proc.stdout)
                    st.stop()

                # Locate Fresh Reports
                fresh_excels = [f for f in glob.glob("reports/*.xlsx") if not os.path.basename(f).startswith("~$") and os.path.getmtime(f) >= run_timestamp - 2]
                quarantine_files = [f for f in glob.glob("quarantine/*") if os.path.getmtime(f) >= run_timestamp - 2]
                parquet_files = [f for f in glob.glob("reports/*.parquet") if os.path.getmtime(f) >= run_timestamp - 2]

                if not fresh_excels:
                    st.error("Dashboard workbook generate nahi ho payi.")
                    st.code(proc.stdout)
                else:
                    st.success("✅ Audit Suite compiled! All enterprise delivery artifacts are ready.")
                    st.write("### 📦 Enterprise Delivery Artifacts")
                    col_a, col_b, col_c = st.columns(3)

                    # 1. Master Excel Dashboard
                    latest_excel = max(fresh_excels, key=os.path.getmtime)
                    with open(latest_excel, "rb") as f:
                        excel_bytes = f.read()
                    with col_a:
                        st.download_button(
                            label="📊 1. Executive Dashboard (.xlsx)",
                            data=excel_bytes,
                            file_name=os.path.basename(latest_excel),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                    # 2. Quarantine Audit Log
                    with col_b:
                        if quarantine_files:
                            latest_quar = max(quarantine_files, key=os.path.getmtime)
                            with open(latest_quar, "rb") as f:
                                quar_bytes = f.read()
                            st.download_button(
                                label="🛡️ 2. Quarantine Audit Log (.csv)",
                                data=quar_bytes,
                                file_name=os.path.basename(latest_quar),
                                mime="text/csv",
                                use_container_width=True
                            )
                        else:
                            st.info("🛡️ Quarantine: 0 dirty rows dropped (100% clean data).")

                    # 3. Optimized Parquet File
                    with col_c:
                        stem_name = os.path.splitext(uploaded_file.name)[0]
                        if parquet_files:
                            latest_parquet = max(parquet_files, key=os.path.getmtime)
                            with open(latest_parquet, "rb") as f:
                                parq_bytes = f.read()
                            st.download_button(
                                label="⚡ 3. Analytical Parquet (.parquet)",
                                data=parq_bytes,
                                file_name=os.path.basename(latest_parquet),
                                mime="application/octet-stream",
                                use_container_width=True
                            )
                        else:
                            xl = pd.ExcelFile(latest_excel)
                            sheet_to_use = 'Cleaned_Data' if 'Cleaned_Data' in xl.sheet_names else xl.sheet_names[0]
                            df_clean = xl.parse(sheet_to_use)
                            parq_path = os.path.join(tmp_dir, "cleaned_dataset.parquet")
                            df_clean.to_parquet(parq_path, index=False)
                            with open(parq_path, "rb") as f:
                                parq_bytes = f.read()
                            st.download_button(
                                label="⚡ 3. Analytical Parquet (.parquet)",
                                data=parq_bytes,
                                file_name=f"{stem_name}_Cleaned.parquet",
                                mime="application/octet-stream",
                                use_container_width=True
                            )
