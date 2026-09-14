import streamlit as st
import pandas as pd
import subprocess
import os
import glob
import time
import tempfile

st.set_page_config(page_title="Gatekeeper BI | Enterprise Suite", page_icon="🛡️", layout="wide")

st.title("🛡️ Gatekeeper BI: Executive Dashboard Engine")
st.caption("Universal 2FA Autonomous Reporting Pipeline (v71.0 Enterprise Master)")

uploaded_file = st.file_uploader("Upload raw business dataset (.csv or .xlsx)")

if uploaded_file is not None:
    if not (uploaded_file.name.lower().endswith('.csv') or uploaded_file.name.lower().endswith('.xlsx')):
        st.error("⚠️ Invalid format! Please upload only .csv or .xlsx file.")
        st.stop()
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
