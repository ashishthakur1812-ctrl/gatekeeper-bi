import time
import json
import os
try:
    import streamlit as st
    if "GEMINI_API_KEY" in st.secrets:
        os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass
from pathlib import Path
import sys
from google import genai
from google.genai import types
import numpy as np
import pandas as pd


def get_working_llm_contract(client, prompt: str) -> dict:
    candidates = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.1-flash-lite"]
    delays = [1, 2, 4]
    for model_name in candidates:
        for attempt, delay in enumerate(delays):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.0,
                        seed=42
                    )
                )
                data = json.loads(response.text)
                print(f"[LIVE SUCCESS] LLM Contract established via: {model_name}")
                return data
            except Exception as err:
                err_str = str(err)
                if ("503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str) and attempt < len(delays) - 1:
                    print(f"[COOLDOWN RETRY] {model_name} busy/rate-limited. Waiting {delay}s before next attempt ({attempt+1}/{len(delays)})...")
                    time.sleep(delay)
                    continue
                print(f"[MODEL RETRY] {model_name} failed: {err}")
                break
    raise RuntimeError("All configured Gemini endpoints unreachable.")

def run_agent(dataset_path: str):
    stem = Path(dataset_path).stem

    # Universal File Ingestion
    if dataset_path.endswith((".xlsx", ".xls")):
        df = pd.read_excel(dataset_path)
    else:
        df = pd.read_csv(dataset_path)

    total_rows = len(df)
    cols = list(df.columns)

    # -------------------------------------------------------------------------
    # 1. ENRICHED METADATA PROFILER (Distribution, Skewness & Types)
    # -------------------------------------------------------------------------
    col_profiles = {}
    detected_dates = []

    for c in cols:
        n_unique = int(df[c].nunique())
        is_num = bool(pd.api.types.is_numeric_dtype(df[c]))

        # Date Detection
        is_date = False
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            is_date = True
        elif df[c].dtype == "object":
            sample_dates = df[c].dropna().head(10).astype(str)
            parsed = pd.to_datetime(sample_dates, errors="coerce")
            if parsed.notnull().sum() >= 7:
                is_date = True

        if is_date or any(k in str(c).lower() for k in ["date", "time", "year", "month", "day"]):
            is_date = True
            detected_dates.append(c)
        # Skewness Calculation (Dominance share of top value)
        top_share = 0.0
        if total_rows > 0 and n_unique > 0:
            top_share = round(
                float(df[c].value_counts(normalize=True).iloc[0]) * 100, 1
            )

        # Identifier / Noise Flag (IDs, Codes, Phone, UUID)
        is_id = False
        c_low = c.lower()
        if any(
            bad in c_low
            for bad in [
                "id",
                "code",
                "zip",
                "phone",
                "pincode",
                "uuid",
                "row_num",
            ]
        ):
            is_id = True
        elif total_rows > 50 and (n_unique / total_rows) > 0.6:
            is_id = True

        col_profiles[c] = {
            "dtype": str(df[c].dtype),
            "nunique": n_unique,
            "is_numeric": is_num,
            "is_date": is_date,
            "is_id_like": is_id,
            "top_value_dominance_pct": top_share,
        }

    meta_payload = {
        "columns": cols,
        "sample_rows": df.head(2).to_dict(orient="records"),
        "profiles": col_profiles,
        "detected_dates": detected_dates,
    }

    # -------------------------------------------------------------------------
    # 2. STRICT ENTERPRISE PROMPT (Semantic Context with Hard Guardrails)
    # -------------------------------------------------------------------------
    prompt = f"""
You are an enterprise Business Intelligence semantic architect.
Analyze this dataset metadata and determine optimal KPI bindings, dimensions, and styling for an executive dashboard.

CRITICAL ARCHITECTURE RULES:
1. DOMAIN IDENTIFICATION: Detect dataset domain (e.g., Financial, Healthcare, Human Resources, Logistics, E-commerce, SaaS).
2. CURRENCY RULE: Set "currency_symbol" to "$" ONLY if monetary/financial metrics exist. Otherwise set to "".
3. 4 UNIVERSAL KPI SLOTS:
   - Slot 1 (Scale/Volume): Count of primary entity/transactions (agg: "COUNT", format: "#,##0").
   - Slot 2 (Primary Throughput): Dominant volume metric (agg: "SUM", format: "$#,##0" if currency else "#,##0").
   - Slot 3 (Operational Friction/Secondary): Cost, delay, incident, or secondary volume (agg: "SUM" or "AVG", format: "$#,##0" or "#,##0").
   - Slot 4 (Efficiency/Intensity): Rate, margin, ratio, or quality score (agg: "AVG", format: "0.0%" or "0.00").
4. DIMENSIONS:
   - "macro_dimension": Categorical column with 3 to 25 unique values, top_value_dominance_pct < 80 (Used for Primary Ranking Chart).
   - "secondary_dimension": Distinct categorical column for drilldown distribution (MUST NOT be equal to macro_dimension).
   - "temporal_dimension": Primary date column from detected_dates, or null.
5. PALETTE: Choose from "SLATE_CORPORATE", "EMERALD_GROWTH", "OCEAN_BLUE", "AMBER_EXECUTIVE".

METADATA:
{json.dumps(meta_payload, indent=2)}

Respond ONLY with valid JSON strictly conforming to:
{{
  "dataset_signature": "{stem}",
  "domain": "<DETECTED_DOMAIN>",
  "theme_palette": "<CHOSEN_PALETTE>",
  "currency_symbol": "<$ or empty string>",
  "kpi_slots": [
    {{"slot": 1, "title": "<KPI_TITLE>", "column": "<col_name>", "aggregation": "COUNT", "format": "#,##0"}},
    {{"slot": 2, "title": "<KPI_TITLE>", "column": "<col_name>", "aggregation": "SUM", "format": "<FORMAT>"}},
    {{"slot": 3, "title": "<KPI_TITLE>", "column": "<col_name>", "aggregation": "<SUM or AVG>", "format": "<FORMAT>"}},
    {{"slot": 4, "title": "<KPI_TITLE>", "column": "<col_name>", "aggregation": "AVG", "format": "<FORMAT>"}}
  ],
  "primary_measure": {{
    "column": "<same as slot 2 column>",
    "aggregation": "SUM",
    "optimization_goal": "MAX"
  }},
  "dimensions": {{
    "temporal_dimension": "<date col or null>",
    "macro_dimension": "<best categorical column>",
    "secondary_dimension": "<distinct secondary categorical column>"
  }},
  "kpi_measures": [
    ["<slot 2 column>", "SUM"],
    ["<slot 3 column>", "SUM"]
  ]
}}
"""

    # -------------------------------------------------------------------------
    # 3. LLM EXECUTION WITH STATISTICAL FALLBACK (ZERO KEYWORD BIAS)
    # -------------------------------------------------------------------------
    try:
        api_key = os.environ.get("GEMINI_API_KEY")
        client = genai.Client(api_key=api_key)
        contract = get_working_llm_contract(client, prompt)
    except Exception as e:
        print(f"[FALLBACK ACTIVATED] Pure Statistical Engine engaged. Reason: {e}")
        
        # Keyword-Agnostic Mathematical Profiling
        valid_nums = [c for c, p in col_profiles.items() if p["is_numeric"] and not p["is_id_like"]]
        valid_cats = [c for c, p in col_profiles.items() if not p["is_numeric"] and not p["is_id_like"] and not p["is_date"] and p.get("top_value_dominance_pct", 0) < 80]
        
        if not valid_cats:
            valid_cats = [c for c, p in col_profiles.items() if not p["is_numeric"] and not p["is_date"]]
        if not valid_cats:
            valid_cats = list(cols[:1])
            
        macro_d = valid_cats[0]
        sec_d = valid_cats[1] if len(valid_cats) > 1 else (cols[1] if len(cols) > 1 else macro_d)
        time_d = detected_dates[0] if detected_dates else None
        
        # Identify Ratios vs Additive Measures mathematically
        ratios = [c for c in valid_nums if any(k in c.lower() for k in ["pct", "rate", "margin", "ratio", "percent"]) or (col_profiles[c].get("max", 1) <= 1.0)]
        additives = [c for c in valid_nums if c not in ratios]
        
        p_measure = additives[0] if additives else (valid_nums[0] if valid_nums else None)
        s3_measure = additives[1] if len(additives) > 1 else (valid_nums[1] if len(valid_nums) > 1 else p_measure)
        s4_measure = ratios[0] if ratios else (valid_nums[-1] if valid_nums else None)
        
        # Domain currency inference (Financial vs Units)
        has_curr = any(k in c.lower() for c in valid_nums for k in ["revenue", "sales", "cost", "price", "amount", "profit"])
        c_sym = "$" if has_curr else ""
        
        contract = {
            "dataset_signature": stem,
            "domain": "Enterprise General",
            "theme_palette": "SLATE_CORPORATE",
            "currency_symbol": c_sym,
            "kpi_slots": [
                {"slot": 1, "title": "TOTAL RECORDS", "column": cols[0], "aggregation": "COUNT", "format": "#,##0"},
                {"slot": 2, "title": f"TOTAL {p_measure.upper()}" if p_measure else "TOTAL VOLUME", "column": p_measure, "aggregation": "SUM", "format": f"{c_sym}#,##0" if c_sym else "#,##0"},
                {"slot": 3, "title": f"SUM {s3_measure.upper()}" if s3_measure else "TOTAL METRIC", "column": s3_measure, "aggregation": "SUM", "format": f"{c_sym}#,##0" if c_sym else "#,##0"},
                {"slot": 4, "title": f"AVG {s4_measure.upper()}" if s4_measure else "EFFICIENCY INDEX", "column": s4_measure, "aggregation": "AVG", "format": "0.0%" if ratios else "0.00"}
            ],
            "primary_measure": {
                "column": p_measure,
                "aggregation": "SUM" if p_measure else "COUNT",
                "optimization_goal": "MAX"
            },
            "dimensions": {
                "temporal_dimension": time_d,
                "macro_dimension": macro_d,
                "secondary_dimension": sec_d
            },
            "kpi_measures": [[p_measure, "SUM"]] if p_measure else []
        }

    # -------------------------------------------------------------------------
    # 4. MATHEMATICAL POST-AUDIT & HARMONIZATION (ZERO CRASH ASSURANCE)
    # -------------------------------------------------------------------------
    col_map = {c.lower(): c for c in cols}
    valid_nums = [c for c, p in col_profiles.items() if p["is_numeric"] and not p["is_id_like"]]
    
    # Audit Dimensions
    dims = contract.get("dimensions", {})
    t_dim = dims.get("temporal_dimension")
    if t_dim and str(t_dim).lower() in col_map:
        contract["dimensions"]["temporal_dimension"] = col_map[str(t_dim).lower()]
    elif detected_dates:
        contract["dimensions"]["temporal_dimension"] = detected_dates[0]
    else:
        contract["dimensions"]["temporal_dimension"] = None
        
    m_dim = dims.get("macro_dimension", "")
    if m_dim and str(m_dim).lower() in col_map:
        contract["dimensions"]["macro_dimension"] = col_map[str(m_dim).lower()]
    else:
        contract["dimensions"]["macro_dimension"] = cols[0]
        
    s_dim = dims.get("secondary_dimension", "")
    if s_dim and str(s_dim).lower() in col_map and col_map[str(s_dim).lower()] != contract["dimensions"]["macro_dimension"]:
        contract["dimensions"]["secondary_dimension"] = col_map[str(s_dim).lower()]
    else:
        alt_cats = [c for c in cols if c != contract["dimensions"]["macro_dimension"]]
        contract["dimensions"]["secondary_dimension"] = alt_cats[0] if alt_cats else contract["dimensions"]["macro_dimension"]
        
    # Audit KPI Slots
    raw_slots = contract.get("kpi_slots", [])
    audited_slots = []
    default_num = valid_nums[0] if valid_nums else cols[0]
    
    for i, s in enumerate(raw_slots):
        slot_num = s.get("slot", i + 1)
        raw_c = str(s.get("column", "")).lower()
        real_c = col_map.get(raw_c, default_num)
        agg = str(s.get("aggregation", "SUM")).upper()
        if agg not in ["SUM", "AVG", "COUNT"]:
            agg = "SUM"
        audited_slots.append({
            "slot": slot_num,
            "title": s.get("title", f"KPI {slot_num}").upper(),
            "column": real_c,
            "aggregation": agg,
            "format": s.get("format", "#,##0")
        })
    
    # If LLM omitted slots, generate all 4 deterministically
    while len(audited_slots) < 4:
        idx = len(audited_slots) + 1
        audited_slots.append({
            "slot": idx,
            "title": f"METRIC {idx}",
            "column": default_num,
            "aggregation": "SUM",
            "format": "#,##0"
        })
    contract["kpi_slots"] = audited_slots
    
    # Harmonize legacy primary_measure & kpi_measures
    contract["primary_measure"] = {
        "column": audited_slots[1]["column"] if len(audited_slots) > 1 else default_num,
        "aggregation": audited_slots[1]["aggregation"] if len(audited_slots) > 1 else "SUM",
        "optimization_goal": "MAX"
    }
    contract["kpi_measures"] = [[s["column"], s["aggregation"]] for s in audited_slots if s["column"] in valid_nums]
    
    # 5. PERSIST CONTRACT FILE
    # -------------------------------------------------------------------------
    Path("contracts").mkdir(exist_ok=True)
    out_file = Path("contracts") / f"{stem}_contract.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(contract, f, indent=2)
    print(f"[CONTRACT LOCKED & AUDITED] -> {out_file}")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "test_data.csv"
    run_agent(target)
