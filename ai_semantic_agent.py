import os
import sys
import json
import pandas as pd
from pathlib import Path
from google import genai
from google.genai import types

def get_working_llm_contract(client, prompt: str) -> dict:
    candidates = ["gemini-3.6-flash", "gemini-2.5-pro"]
    try:
        for m in client.models.list():
            name = getattr(m, 'name', '') or str(m)
            clean = name.split('models/')[-1] if 'models/' in name else name
            if clean not in candidates and "preview-tts" not in clean:
                candidates.append(clean)
    except Exception:
        pass

    for model_name in candidates:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            data = json.loads(response.text)
            print(f"[LIVE SUCCESS] LLM Contract established via: {model_name}")
            return data
        except Exception:
            continue
    raise RuntimeError("All candidate endpoints unreachable.")

def run_agent(dataset_path: str):
    stem = Path(dataset_path).stem
    df = pd.read_csv(dataset_path)
    cols = list(df.columns)
    
    meta_payload = {
        "columns": cols,
        "sample_rows": df.head(2).to_dict(orient="records"),
        "unique_counts": {c: int(df[c].nunique()) for c in cols}
    }
    
    prompt = f"""
    You are an enterprise Business Intelligence semantic architect.
    Analyze this dataset metadata and determine optimal KPI bindings.
    METADATA:
    {json.dumps(meta_payload, indent=2)}
    Respond ONLY with valid JSON conforming strictly to:
    {{
        "dataset_signature": "{stem}",
        "theme_palette": "HEALTHCARE_EMERALD",
        "currency_symbol": "",
        "primary_measure": {{
            "column": "<must exist in columns>",
            "aggregation": "SUM",
            "optimization_goal": "MAX"
        }},
        "dimensions": {{
            "macro_dimension": "<categorical dimension for bar chart>",
            "secondary_dimension": "<categorical dimension DIFFERENT from macro_dimension>"
        }},
        "kpi_measures": [
            ["<measure_col_1>", "AVG"],
            ["<measure_col_2>", "SUM"]
        ]
    }}
    Rules: Intensive (Age/Rate/Pct/Latency) must be AVG. Financial/Volume must be SUM. Raw JSON only.
    """
    
    try:
        api_key = os.environ.get("GEMINI_API_KEY")
        client = genai.Client(api_key=api_key)
        contract = get_working_llm_contract(client, prompt)
    except Exception as e:
        print(f"[FALLBACK ACTIVATED] Reason: {e}")
        contract = {
            "dataset_signature": stem,
            "theme_palette": "HEALTHCARE_EMERALD",
            "currency_symbol": "",
            "primary_measure": {"column": cols[1], "aggregation": "SUM", "optimization_goal": "MAX"},
            "dimensions": {"macro_dimension": cols[0], "secondary_dimension": cols[0]},
            "kpi_measures": [[cols[1], "SUM"]]
        }

    # Invariant Grounding
    col_map = {c.lower(): c for c in cols}
    if contract.get("primary_measure", {}).get("column", "").lower() in col_map:
        contract["primary_measure"]["column"] = col_map[contract["primary_measure"]["column"].lower()]
    if contract.get("dimensions", {}).get("macro_dimension", "").lower() in col_map:
        contract["dimensions"]["macro_dimension"] = col_map[contract["dimensions"]["macro_dimension"].lower()]
    if contract.get("dimensions", {}).get("secondary_dimension", "").lower() in col_map:
        contract["dimensions"]["secondary_dimension"] = col_map[contract["dimensions"]["secondary_dimension"].lower()]

    for item in contract.get("kpi_measures", []):
        if any(k in str(item[0]).lower() for k in ['age', 'rate', 'pct', 'ratio', 'margin', 'latency']):
            item[1] = "AVG"

    Path("contracts").mkdir(exist_ok=True)
    out_file = Path("contracts") / f"{stem}_contract.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(contract, f, indent=2)
    print(f"[CONTRACT LOCKED] -> {out_file}")

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "test_data.csv"
    run_agent(target)
