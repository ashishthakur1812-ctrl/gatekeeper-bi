import json
import os
import time
from pathlib import Path
import pandas as pd

try:
    import streamlit as st
    if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
        os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

from google import genai
from google.genai import types

MODEL_CANDIDATES = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.1-flash-lite"]

def profile_metadata(df: pd.DataFrame, dataset_name: str) -> dict:
    total_rows = len(df)
    cols_meta = {}
    for col in df.columns:
        c_clean = str(col).strip()
        c_low = c_clean.lower()
        n_unique = int(df[col].nunique())
        is_num = bool(pd.api.types.is_numeric_dtype(df[col]))

        # Quarantine IDs, Keys, and Codes
        is_id = False
        id_exact = {"id", "invoice_no", "order_no", "customer_id", "product_id", "patient_id", "trip_id", "trans_id", "row_id", "uuid", "account_no", "loan_id", "emp_id", "pin", "zip", "pincode", "zipcode", "hospital_zip_code", "origin_hub_pin"}
        id_suffixes = ("_id", "_code", "_no", "_num", "_uuid", "_ref", "_pin", "_zip", "_ssn")
        is_metric_word = any(m in c_low for m in ["amount", "count", "revenue", "cost", "salary", "price", "total", "margin", "rate", "distance", "km", "hours", "days", "score", "balance", "interest", "bonus"])

        if not is_metric_word:
            if c_low in id_exact or c_low.endswith(id_suffixes):
                is_id = True
            elif (not is_num) and total_rows > 30 and (n_unique / total_rows) > 0.80:
                is_id = True

        stats = {}
        sample_labels = []
        if is_num and not is_id:
            clean_num = pd.to_numeric(df[col], errors="coerce").dropna()
            if not clean_num.empty:
                stats = {
                    "min": float(round(clean_num.min(), 2)),
                    "max": float(round(clean_num.max(), 2)),
                    "mean": float(round(clean_num.mean(), 2))
                }
        elif not is_num and not is_id and n_unique <= 15:
            # Safe structural metadata: capture top 3 classification labels
            sample_labels = [str(x) for x in df[col].dropna().unique()[:3]]

        cols_meta[c_clean] = {
            "dtype": str(df[col].dtype),
            "nunique": n_unique,
            "is_numeric": is_num,
            "is_quarantined_id": is_id,
            "stats": stats,
            "sample_labels": sample_labels
        }
    return {
        "dataset_name": dataset_name,
        "total_records": total_rows,
        "columns": cols_meta
    }

def inspect_contract(contract: dict, meta: dict) -> list:
    violations = []
    cols_meta = meta.get("columns", {})
    available_cols = set(cols_meta.keys())
    
    # 1. Existence and Quarantine Check
    used_cols = []
    for slot in contract.get("kpi_slots", []):
        col = slot.get("column")
        slot_num = slot.get("slot")
        fmt = slot.get("format", "")

        if col not in available_cols:
            violations.append(f"Slot {slot_num}: Column '{col}' does not exist in dataset.")
            continue

        c_info = cols_meta.get(col, {})
        if c_info.get("is_quarantined_id"):
            violations.append(f"Slot {slot_num}: Column '{col}' is a Quarantined Tracking ID. FORBIDDEN.")

        if not c_info.get("is_numeric"):
            violations.append(f"Slot {slot_num}: Column '{col}' is non-numeric string. Cannot aggregate.")

        stats = c_info.get("stats", {})
        max_val = stats.get("max", 0)
        if fmt == "0.0%" and max_val > 1.5:
            violations.append(f"Slot {slot_num}: Column '{col}' has values up to {max_val}. Percentage format '0.0%' invalid.")

        used_cols.append(col)

    # 2. Duplicate Metric Check
    if len(used_cols) >= 2 and used_cols[0] == used_cols[1]:
        violations.append(f"Slot 2 and Slot 3 both use duplicate column '{used_cols[0]}'. Distinct metrics required.")

    # 3. Chart Dimension & Macro Hierarchy Checks
    dims = contract.get("dimensions", {})
    macro = dims.get("macro_dimension")
    sec = dims.get("secondary_dimension")
    
    if macro and macro not in available_cols:
        violations.append(f"Bar Chart Dimension '{macro}' does not exist in dataset.")
    if sec and sec not in available_cols:
        violations.append(f"Donut Chart Dimension '{sec}' does not exist in dataset.")
    if macro and sec and macro == sec:
        violations.append(f"Bar and Donut charts cannot use exact same dimension '{macro}'.")

    # Priority check: If a clear 'category'/'department' exists, it must be macro
    for cand in available_cols:
        cand_low = cand.lower()
        if any(k in cand_low for k in ["category", "department", "service", "product_category"]):
            if macro != cand and sec == cand:
                violations.append(f"Hierarchy violation: Descriptive core taxonomy '{cand}' MUST be 'macro_dimension' for Bar Chart, not secondary.")

    return violations

def run_agent(dataset_path: str) -> dict:
    stem = Path(dataset_path).stem
    cache_path = Path(f"{stem}_contract.json")
    if cache_path.exists():
        try:
            return json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    ext = Path(dataset_path).suffix.lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(dataset_path)
    else:
        df = pd.read_csv(dataset_path)

    meta = profile_metadata(df, stem)
    base_prompt = f"""You are an Enterprise BI Semantic Architect.
Extract the optimal executive dashboard contract from this dataset metadata:

{json.dumps(meta, indent=2)}

CRITICAL GOVERNANCE RULES:
1. STRICT ISOLATION: NEVER select columns where "is_quarantined_id" is true, nor tracking IDs, codes, or strings.
2. NO DUPLICATE MEASURES: Slot 2 and Slot 3 MUST NOT use the same column.
3. MATHEMATICAL FORMAT ACCURACY:
   - Currency/Spend: format "$#,##0".
   - Counts/Hours/Scores/Days: format "#,##0" or "#,##0.0".
   - ONLY use "0.0%" if column is an actual ratio/percentage. Never format scores or days as %.
4. KPI ALLOCATION:
   - Slot 2 (Primary Volume): Dominant volume / throughput metric (SUM).
   - Slot 3 (Operational / Intensive Metric): Expense/Volume (SUM) OR Duration/Tenure/Experience/Wait-time (AVG). NEVER use SUM for durations or ages.
   - Slot 4 (Efficiency / Intensity): Average performance rate, ratio, score or stay length (AVG).
5. VISUAL HIERARCHY RULES:
   - "macro_dimension": Primary qualitative business entity for the Bar Chart. ALWAYS choose the descriptive, human-readable core taxonomy (e.g., category, department, product). Inspect "sample_labels" to avoid cryptic 3-4 letter codes.
   - "secondary_dimension": Distinct qualitative dimension for Donut Chart (e.g., segment, tier, region).

Respond ONLY with valid JSON strictly matching:
{{
  "domain": "<DETECTED_DOMAIN>",
  "primary_measure": {{"column": "<slot 2 col>", "aggregation": "SUM"}},
  "kpi_slots": [
    {{"slot": 2, "title": "<TITLE>", "column": "<slot 2 col>", "aggregation": "SUM", "format": "<FORMAT>"}},
    {{"slot": 3, "title": "<TITLE>", "column": "<slot 3 col>", "aggregation": "SUM", "format": "<FORMAT>"}},
    {{"slot": 4, "title": "<TITLE>", "column": "<slot 4 col>", "aggregation": "AVG", "format": "<FORMAT>"}}
  ],
  "dimensions": {{
    "macro_dimension": "<best bar chart col>",
    "secondary_dimension": "<best donut chart col>"
  }}
}}"""

    client = genai.Client()
    current_prompt = base_prompt
    contract = None

    for iteration in range(1, 4):
        print(f"[AGENT ATTEMPT {iteration}] Requesting contract from Gemini...")
        for model_name in MODEL_CANDIDATES:
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=current_prompt,
                    config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0, seed=42)
                )
                contract = json.loads(response.text.strip())
                break
            except Exception as e:
                time.sleep(1)
                continue

        if not contract:
            raise RuntimeError("Gemini endpoints unreachable.")

        violations = inspect_contract(contract, meta)
        if not violations:
            print(f"[AGENT APPROVED] Contract 100% compliant on iteration {iteration}!")
            break
        else:
            print(f"[AGENT CRITIQUE] Violations detected on attempt {iteration}:")
            for v in violations:
                print(f"  * {v}")
            critique_message = "\n".join(f"- {v}" for v in violations)
            current_prompt = f"""{base_prompt}

CRITICAL REJECTION NOTICE ON YOUR PREVIOUS OUTPUT:
Your previous JSON was rejected due to strict constitutional violations:
{critique_message}

Fix these specific violations immediately. Pick alternative valid columns from metadata and regenerate the JSON strictly conforming to rules."""

    if not contract:
        raise RuntimeError("Failed to obtain valid contract after critique loops.")

    cache_path.write_text(json.dumps(contract, indent=2), encoding="utf-8")
    return contract
