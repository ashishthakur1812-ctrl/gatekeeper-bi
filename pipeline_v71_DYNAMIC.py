import json
from pathlib import Path
import os
import re
import sys
import csv
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, Optional, Sequence
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrData, StrRef, StrVal, NumData, NumRef, NumVal, NumDataSource
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend


# Validation policy is intentionally explicit so a deployment can change these
# values without changing the circuit-breaker implementation.
GRACEFUL_DEFAULTS: Dict[str, object] = {
    'city': 'Unknown',
    'category': 'Uncategorized',
    'notes': '',
}
FATAL_CORRUPTION_THRESHOLD_PERCENT = 5.0
PRIMARY_KEY_PATTERN = r'(?:^|[_\s-])(?:id|key|uuid|code|sku|account|record)(?:[_\s-]|$)'
NON_NEGATIVE_METRIC_PATTERN = r'(?:amount|amt|metric|measure|value|revenue|price|cost|total|quantity|qty|count|score|salary)'


@dataclass(frozen=True)
class ValidationResult:
    dataframe: pd.DataFrame
    soft_imputations: int
    fatal_corrupt_rows: int
    status: str


def _matching_columns(columns: pd.Index, configured: Sequence[str]) -> pd.Index:
    configured_index = pd.Index(configured, dtype='object').astype(str).str.casefold()
    normalized_columns = pd.Series(columns, index=columns, dtype='object').astype(str).str.casefold()
    return normalized_columns[normalized_columns.isin(configured_index)].index


def _resolve_primary_key(df: pd.DataFrame, primary_key: Optional[str]) -> Optional[str]:
    if primary_key in df.columns:
        return primary_key
    normalized = pd.Series(df.columns, index=df.columns, dtype='object').astype(str)
    candidates = normalized[normalized.str.contains(PRIMARY_KEY_PATTERN, case=False, regex=True)].index
    return str(candidates[0]) if len(candidates) else None


def _resolve_non_negative_metrics(df: pd.DataFrame, primary_key: Optional[str]) -> pd.Index:
    numeric_columns = df.select_dtypes(include=np.number).columns
    normalized = pd.Series(numeric_columns, index=numeric_columns, dtype='object').astype(str)
    metric_columns = normalized[normalized.str.contains(NON_NEGATIVE_METRIC_PATTERN, case=False, regex=True)].index
    if primary_key is not None:
        metric_columns = metric_columns[metric_columns != primary_key]
    return metric_columns


def _write_validation_log(*args, **kwargs):
    import os
    from pathlib import Path
    if len(args) >= 8:
        output_dir, base_stem, raw_count, clean_count, soft_imp, fatal_corr, proc_time, status = args[:8]
    elif len(args) == 7:
        output_dir, raw_count, clean_count, soft_imp, fatal_corr, proc_time, status = args
        cur_file = globals().get('CURRENT_INPUT_FILE', 'Dataset')
        base_stem = Path(cur_file).stem.replace('_Cleaned', '').replace('_Gatekeeper_Dashboard', '')
    else:
        output_dir = args[0] if len(args) > 0 else 'reports'
        cur_file = globals().get('CURRENT_INPUT_FILE', 'Dataset')
        base_stem = Path(cur_file).stem.replace('_Cleaned', '').replace('_Gatekeeper_Dashboard', '')
        raw_count = args[1] if len(args) > 1 else 0
        clean_count = args[2] if len(args) > 2 else 0
        soft_imp = args[3] if len(args) > 3 else 0
        fatal_corr = args[4] if len(args) > 4 else 0
        proc_time = args[5] if len(args) > 5 else 0.0
        status = args[6] if len(args) > 6 else 'SUCCESS'
    os.makedirs(output_dir, exist_ok=True)
    log_file_name = f"{base_stem}_Run_Summary_Log.txt"
    log_file_path = os.path.join(output_dir, log_file_name)
    with open(log_file_path, 'w', encoding='utf-8') as f:
        f.write(f"Dataset Signature: {base_stem}\n")
        f.write(f"Total Rows Ingested: {raw_count}\n")
        f.write(f"Total Rows Exported: {clean_count}\n")
        f.write(f"Soft Imputations Applied: {soft_imp}\n")
        f.write(f"Fatal Corrupt Rows Blocked: {fatal_corr}\n")
        f.write(f"Total Processing Time (seconds): {float(proc_time):.2f}\n")
        f.write(f"Status: {status}\n")


def validate_with_circuit_breaker(
    df: pd.DataFrame,
    output_dir: str,
    total_rows_ingested: Optional[int] = None,
    critical_columns: Optional[Sequence[str]] = None,
    primary_key: Optional[str] = None,
    numeric_bound_columns: Optional[Sequence[str]] = None,
    started_at: Optional[float] = None,
) -> ValidationResult:
    """Apply vectorized soft defaults and enforce fatal data-quality invariants."""
    validation_df = df.copy()
    os.makedirs(output_dir, exist_ok=True)
    started = time.perf_counter() if started_at is None else started_at
    ingested_rows = len(validation_df) if total_rows_ingested is None else total_rows_ingested

    default_columns = _matching_columns(validation_df.columns, GRACEFUL_DEFAULTS.keys())
    default_values = pd.Series(GRACEFUL_DEFAULTS, dtype='object')
    defaults_to_apply = default_values.reindex(default_columns.astype(str).str.casefold()).set_axis(default_columns).to_dict()
    soft_imputations = int(validation_df.loc[:, default_columns].isna().sum().sum()) if len(default_columns) else 0
    if defaults_to_apply:
        validation_df = validation_df.fillna(defaults_to_apply)

    resolved_key = _resolve_primary_key(validation_df, primary_key)
    inferred_critical = _resolve_non_negative_metrics(validation_df, resolved_key)
    configured_critical = _matching_columns(
        validation_df.columns,
        inferred_critical if critical_columns is None else critical_columns,
    )
    critical = configured_critical.union(pd.Index([resolved_key])) if resolved_key else configured_critical
    critical_null_mask = validation_df.loc[:, critical].isna().any(axis=1) if len(critical) else pd.Series(False, index=validation_df.index)

    duplicate_mask = validation_df[resolved_key].duplicated(keep=False) if resolved_key else pd.Series(False, index=validation_df.index)
    bound_columns = _matching_columns(validation_df.columns, numeric_bound_columns or _resolve_non_negative_metrics(validation_df, resolved_key))
    bound_mask = validation_df.loc[:, bound_columns].lt(0).any(axis=1) if len(bound_columns) else pd.Series(False, index=validation_df.index)
    fatal_mask = critical_null_mask | duplicate_mask | bound_mask

    fatal_corrupt_rows = int(fatal_mask.sum())
    
    # Audit trail: Tag reason for quarantine before export
    if fatal_corrupt_rows > 0:
        _reasons = []
        for idx in validation_df[fatal_mask].index:
            r = []
            if critical_null_mask.loc[idx]:
                r.append("CRITICAL_NULL")
            if duplicate_mask.loc[idx]:
                r.append("DUPLICATE_KEY")
            if bound_mask.loc[idx]:
                r.append("NEGATIVE_VALUE_VIOLATION")
            _reasons.append("; ".join(r) if r else "ANOMALY")
        
        quarantine_export_df = validation_df.loc[fatal_mask].copy()
        quarantine_export_df.insert(0, "Quarantine_Reason", _reasons)
    else:
        quarantine_export_df = validation_df.iloc[0:0].copy()
    total_rows = len(validation_df)
    fatal_percentage = (fatal_corrupt_rows / total_rows * 100.0) if total_rows else 0.0
    os.makedirs('quarantine', exist_ok=True); _q_stem = os.path.splitext(os.path.basename(globals().get('CURRENT_INPUT_FILE') or globals().get('file_input') or 'records'))[0]; quarantine_path = os.path.join('quarantine', f'{_q_stem}_quarantine.csv')
    elapsed_seconds = max(0.0, time.perf_counter() - started)

    # Production-Grade Adaptive Circuit Breaker
    is_small_batch = total_rows <= 50
    is_breached = (fatal_percentage > 50.0 or (total_rows - fatal_corrupt_rows) < 2) if is_small_batch else (fatal_percentage > FATAL_CORRUPTION_THRESHOLD_PERCENT)

    if is_breached:
        quarantine_export_df.to_csv(quarantine_path, index=False)
        _write_validation_log(output_dir, ingested_rows, 0, soft_imputations, fatal_corrupt_rows, elapsed_seconds, 'CRITICAL HALT')
        raise SystemExit(1)

    clean_df = validation_df.loc[~fatal_mask].copy()
    status = 'SUCCESS'
    if fatal_corrupt_rows:
        quarantine_export_df.to_csv(quarantine_path, index=False)
        status = 'PARTIAL SUCCESS (QUARANTINED)'
    _write_validation_log(output_dir, ingested_rows, len(clean_df), soft_imputations, fatal_corrupt_rows, elapsed_seconds, status)
    return ValidationResult(clean_df, soft_imputations, fatal_corrupt_rows, status)

# ==============================================================================
# MODULE 1: COMPOSITE DE-COUPLER & INGESTION FIREWALL
# ==============================================================================

def clean_file_path(path_str):
    if not path_str: return ""
    return str(path_str).strip().replace('"', '').replace("'", "").strip('& ')

def format_compact_num(val, is_ratio=False, is_avg=False):
    if is_ratio: return f"{val * 100:.1f}%"
    if is_avg: return f"{val:.1f}"
    abs_v = abs(val)
    if abs_v >= 1000000: return f"{val/1000000:.2f}M"
    elif abs_v >= 1000: return f"{val/1000:.2f}K"
    return f"{val:,.2f}"

def get_math_format(df, col_name, agg_type):
    series = pd.to_numeric(df[col_name], errors='coerce').dropna()
    if series.empty: return '#,##0.0'
    c_max = series.max()
    if agg_type == 'AVG' and c_max <= 1.0: return '0.0%'
    if agg_type == 'AVG': return '#,##0.1'
    return '#,##0.00'

def heal_and_ingest_csv(clean_path):
    encodings = ['utf-8-sig', 'utf-8', 'ISO-8859-1', 'cp1252', 'latin1']
    raw_lines = []
    for enc in encodings:
        try:
            with open(clean_path, 'r', encoding=enc, errors='replace') as f:
                raw_lines = [line.strip() for line in f if line.strip()]
            if raw_lines: break
        except Exception:
            continue

    if not raw_lines: return None
    reader = list(csv.reader(raw_lines))
    if not reader: return None

    headers = [str(c).strip().replace('\ufeff', '').replace('"', '') for c in reader[0]]
    expected_len = len(headers)
    sanitized_rows = []

    billing_idx = 6
    for idx_h, h in enumerate(headers):
        if any(k in h.lower() for k in ['bill', 'amt', 'amount', 'revenue', 'price', 'inr']):
            billing_idx = idx_h
            break

    n_tail = expected_len - 1 - billing_idx

    for row in reader[1:]:
        if not row or not any(str(x).strip() for x in row):
            continue
        row_tokens = [str(tok).strip() for tok in row]
        L = len(row_tokens)

        if L == expected_len:
            sanitized_rows.append(row_tokens)
        elif L > expected_len:
            head = row_tokens[:billing_idx]
            tail = row_tokens[L - n_tail : L] if n_tail > 0 else []
            mid = " ".join(row_tokens[billing_idx : L - n_tail])
            sanitized_rows.append(head + [mid] + tail)
        else:
            sanitized_rows.append(row_tokens + [''] * (expected_len - L))

    return pd.DataFrame(sanitized_rows, columns=headers)

def ingest_file(clean_path):
    if not os.path.exists(clean_path): return None
    ext = os.path.splitext(clean_path)[-1].lower()
    try:
        if ext == '.csv':
            return heal_and_ingest_csv(clean_path)
        elif ext in ['.xlsx', '.xls']:
            return pd.read_excel(clean_path)
    except:
        return None
    return None

def clean_dataframe(df):
    cleaned = df.copy()
    cleaned.columns = [str(c).strip().replace('\n', ' ') for c in cleaned.columns]
    initial_count = len(cleaned)
    cols = list(cleaned.columns)

    # De-couple composite text values (e.g. '28450 Active' -> Billing=28450, Status=Active)
    for r_i in range(len(cleaned)):
        for c_i in range(len(cols) - 1):
            val_str = str(cleaned.iat[r_i, c_i]).strip()
            # Match number + string pattern
            m = re.match(r'^([₹$â‚¹\s\d,.\-]+)\s+([A-Za-z].*)$', val_str)
            if m:
                cleaned.iat[r_i, c_i] = m.group(1).strip()
                if pd.isna(cleaned.iat[r_i, c_i + 1]) or str(cleaned.iat[r_i, c_i + 1]).strip() == '':
                    cleaned.iat[r_i, c_i + 1] = m.group(2).strip()

    for col in cleaned.columns:
        col_low = str(col).lower()
        is_code = any(k in col_low for k in ['id', 'code', 'pin', 'zip', 'key', 'sku', 'inv', 'uuid', 'sl_no', 'vin', 'phone', 'account', 'no.', 'unit'])
        cleaned[col] = cleaned[col].replace(['nan', 'NaN', 'None', 'null', 'INVALID_DATE', '<NA>', ''], np.nan)
        
        if is_code:
            cleaned[col] = cleaned[col].dropna().astype(str).str.upper().str.replace(r'\.0$', '', regex=True).str.replace(r'\s+', ' ', regex=True)
            continue
            
        if cleaned[col].dtype == object or pd.api.types.is_string_dtype(cleaned[col]):
            s = cleaned[col]
            
            if any(k in col_low for k in ['date', 'time', 'day', 'period', 'ts', 'timestamp']):
                parsed = pd.to_datetime(s.astype(str).str.replace(r'[_/]', '-', regex=True), format='mixed', errors='coerce')
                if parsed.notna().sum() >= (0.3 * len(cleaned)):
                    cleaned[col] = parsed.dt.strftime('%Y-%m-%d')
                    continue
                    
            s_num = s.astype(str).apply(lambda x: re.sub(r'[^\d.\-]', '', str(x)) if pd.notna(x) and str(x).strip() != '' else np.nan)
            converted = pd.to_numeric(s_num, errors='coerce')
            if converted.notna().sum() >= (0.5 * len(cleaned)):
                cleaned[col] = converted
                continue
                
            cleaned[col] = s.dropna().astype(str).apply(lambda x: ' '.join(str(x).replace('_', ' ').replace('-', ' ').split()).title() if str(x).strip() not in ['', 'Nan'] else np.nan)
            
    cleaned.dropna(how='all', inplace=True)
    cleaned.drop_duplicates(inplace=True)
    return cleaned, initial_count - len(cleaned)

# ==============================================================================
# MODULE 2: MATHEMATICAL GATE & ALGEBRAIC PROFILER (NUMERIC FIRST)
# ==============================================================================

def profile_algebraic_types(df):
    schema = {
        'Additive_Measures': [],
        'Intensive_Measures': [],
        'Categorical_Dims': [],
        'Temporal_Dims': [],
        'Identifier_Keys': [],
        'Metric_Aggregations': {}
    }
    n_rows = len(df)
    if n_rows == 0:
        return schema

    for col in df.columns:
        col_str = str(col).strip()
        col_low = col_str.lower()
        series = df[col].dropna()
        series = series[series != '']
        n_unique = series.nunique()
        if n_unique == 0:
            continue

        uniqueness_ratio = n_unique / n_rows

        # Rule 1: Temporal Gate
        if any(k in col_low for k in ['date', 'time', 'ts', 'timestamp', 'period', 'month', 'year']) or pd.api.types.is_datetime64_any_dtype(series):
            schema['Temporal_Dims'].append(col_str)
            continue

        # Rule 2: Explicit ID / Pincode Keyword Gate
        if any(k in col_low for k in ['id', 'code', 'pin', 'zip', 'key', 'sku', 'phone', 'account', 'unit']):
            if 2 <= n_unique <= 15 and uniqueness_ratio < 0.40:
                schema['Categorical_Dims'].append(col_str)
            else:
                schema['Identifier_Keys'].append(col_str)
            continue

        # Rule 3: Pure Numeric Classification & Mathematical Traps Guard
        if pd.api.types.is_numeric_dtype(series):
            # Trap A: Unlabeled Serial / Key / Pincode
            if (uniqueness_ratio > 0.85 and len(series) > 50) and (series.dtype in ['int64', 'int32', 'int16', 'int8']) and not any(k in col_low for k in ['amount', 'bill', 'sales', 'revenue', 'cost', 'spend', 'price', 'fee', 'charge', 'total', 'age']):
                schema['Identifier_Keys'].append(col_str)
                continue

            # Trap B: Discrete State Code
            if (series.dtype in ['int64', 'int32']) and (2 <= n_unique <= 6) and (len(series) > 50) and not any(k in col_low for k in ['amount', 'bill', 'sales', 'revenue', 'cost', 'spend', 'price', 'fee', 'charge', 'total']):
                schema['Categorical_Dims'].append(col_str)
                continue

            c_min = float(series.min())
            c_max = float(series.max())
            c_mean = float(series.mean()) if len(series) > 0 else 0.0
            c_std = float(series.std()) if len(series) > 1 else 0.0
            cv = (c_std / abs(c_mean)) if c_mean != 0 else 1.0
            skew = float(series.skew()) if len(series) > 2 else 0.0

            # Extensive Additive (Volume, Currency, Quantities) -> Always SUM
            is_extensive = any(k in col_low for k in ['sales', 'revenue', 'cost', 'spend', 'expense', 'profit', 'volume', 'qty', 'quantity', 'units', 'amount', 'total', 'gmv', 'loss', 'count'])

            # Bounded Intensive (Scores, Ratings, Percentages) -> Always AVERAGE
            is_ratio = (c_min >= -1.0) and (c_max <= 1.0) and (series.dtype in ['float64', 'float32'])
            is_pct_rate = (c_min >= 0.0) and (c_max <= 100.0) and any(k in col_low for k in ['pct', 'percent', 'rate', 'ratio', 'margin', 'efficiency'])
            is_rating_score = (c_min >= 0.0) and (c_max <= 100.0) and any(k in col_low for k in ['score', 'rating', 'stars', 'grade', 'index', 'nps', 'csat'])

            # Latency/Duration -> MEDIAN
            is_latency = any(k in col_low for k in ['delay', 'duration', 'latency', 'tat', 'stay', 'wait', 'ping', 'transit', 'lead_time'])

            # Steady sensor / low CV
            is_steady = (cv < 0.15) and (c_min > 0) and not is_extensive

            if is_extensive:
                schema['Additive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'SUM'
            elif is_ratio or is_pct_rate or is_rating_score:
                schema['Intensive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'AVERAGE'
            elif is_latency:
                schema['Intensive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'MEDIAN'
            elif is_steady:
                schema['Intensive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'AVERAGE'
            elif (abs(skew) > 1.2) and (series.dtype in ['float64', 'float32']) and not is_extensive:
                schema['Intensive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'AVERAGE'
            else:
                schema['Additive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'SUM'
            continue

        # Rule 4: Categorical Text Dimensions
        if 2 <= n_unique <= 30 and (uniqueness_ratio < 0.90 or n_rows <= 50):
            schema['Categorical_Dims'].append(col_str)
        else:
            schema['Identifier_Keys'].append(col_str)

    return schema

    for col in df.columns:
        col_str = str(col).strip()
        col_low = col_str.lower()
        series = df[col].dropna()
        series = series[series != '']
        n_unique = series.nunique()
        if n_unique == 0:
            continue

        uniqueness_ratio = n_unique / n_rows

        # Rule 1: Temporal Gate
        if any(k in col_low for k in ['date', 'time', 'ts', 'timestamp', 'period', 'month', 'year']) or pd.api.types.is_datetime64_any_dtype(series):
            schema['Temporal_Dims'].append(col_str)
            continue

        # Rule 2: Explicit ID / Pincode Keyword Gate
        if any(k in col_low for k in ['id', 'code', 'pin', 'zip', 'key', 'sku', 'phone', 'account', 'unit']):
            if 2 <= n_unique <= 15 and uniqueness_ratio < 0.40:
                schema['Categorical_Dims'].append(col_str)
            else:
                schema['Identifier_Keys'].append(col_str)
            continue

        # Rule 3: Pure Numeric Classification & Mathematical Traps Guard
        if pd.api.types.is_numeric_dtype(series):
            # Trap A: Unlabeled Serial / Key / Pincode (High uniqueness integer)
            if (uniqueness_ratio > 0.85 and len(series) > 50) and (series.dtype in ['int64', 'int32', 'int16', 'int8']) and not any(k in col_low for k in ['amount', 'bill', 'sales', 'revenue', 'cost', 'spend', 'price', 'fee', 'charge', 'total', 'age']):
                schema['Identifier_Keys'].append(col_str)
                continue

            # Trap B: Discrete State Code (Status 200, 404, 500 or State 1, 2, 3)
            if (series.dtype in ['int64', 'int32']) and (2 <= n_unique <= 6) and (len(series) > 50) and not any(k in col_low for k in ['amount', 'bill', 'sales', 'revenue', 'cost', 'spend', 'price', 'fee', 'charge', 'total']):
                schema['Categorical_Dims'].append(col_str)
                continue

            c_min = float(series.min())
            c_max = float(series.max())
            c_mean = float(series.mean()) if len(series) > 0 else 0.0
            c_std = float(series.std()) if len(series) > 1 else 0.0
            cv = (c_std / abs(c_mean)) if c_mean != 0 else 1.0
            skew = float(series.skew()) if len(series) > 2 else 0.0

            # Rule A: Explicit Extensive Measures (Additive Volumes / Amounts always SUM)
        is_extensive = any(k in col_low for k in ['sales', 'revenue', 'cost', 'spend', 'expense', 'profit', 'volume', 'qty', 'quantity', 'units', 'amount', 'total', 'gmv', 'loss', 'count'])
        
        # Rule B: Bounded Ratings, Scores & Percentages (Intensive always AVERAGE)
        is_ratio = (c_min >= -1.0) and (c_max <= 1.0) and (series.dtype in ['float64', 'float32'])
        is_pct_rate = (c_min >= 0.0) and (c_max <= 100.0) and any(k in col_low for k in ['pct', 'percent', 'rate', 'ratio', 'margin', 'efficiency'])
        is_rating_score = (c_min >= 0.0) and (c_max <= 100.0) and any(k in col_low for k in ['score', 'rating', 'stars', 'grade', 'index', 'nps', 'csat'])
        
        # Rule C: Time/Latency Metrics (Intensive MEDIAN)
        is_latency = any(k in col_low for k in ['delay', 'duration', 'latency', 'tat', 'stay', 'wait', 'ping', 'transit', 'lead_time'])
        
        # Rule D: Steady-State Sensor (Low Dispersion CV < 0.15)
        is_steady = (cv < 0.15) and (c_min > 0) and not is_extensive

        if is_extensive:
            schema['Additive_Measures'].append(col_str)
            schema['Metric_Aggregations'][col_str] = 'SUM'
        elif is_ratio or is_pct_rate or is_rating_score:
            schema['Intensive_Measures'].append(col_str)
            schema['Metric_Aggregations'][col_str] = 'AVERAGE'
        elif is_latency:
            schema['Intensive_Measures'].append(col_str)
            schema['Metric_Aggregations'][col_str] = 'MEDIAN'
        elif is_steady:
            schema['Intensive_Measures'].append(col_str)
            schema['Metric_Aggregations'][col_str] = 'AVERAGE'
        elif (abs(skew) > 1.2) and (series.dtype in ['float64', 'float32']) and not is_extensive:
            schema['Intensive_Measures'].append(col_str)
            schema['Metric_Aggregations'][col_str] = 'AVERAGE'
        else:
            schema['Additive_Measures'].append(col_str)
            schema['Metric_Aggregations'][col_str] = 'SUM'
        continue

        # Rule 4: Categorical Text Dimensions
        if 2 <= n_unique <= 30 and (uniqueness_ratio < 0.90 or n_rows <= 50):
            schema['Categorical_Dims'].append(col_str)
        else:
            schema['Identifier_Keys'].append(col_str)

    return schema

# MODULE 3: SECTOR ONTOLOGY & EXECUTIVE DIAGNOSTIC
# ==============================================================================

SECTOR_THEMES = {
    'HEALTHCARE_CLINICAL': {'header_fill': '134E4A', 'sub_fill': '115E59', 'accent_fill': '0D9488', 'card_bg': 'F0FDFA', 'filter_bg': 'CCFBF1', 'badge_top': 'D1FAE5', 'badge_lag': 'FFE4E6', 'title_color': '134E4A'},
    'AUTOMOTIVE_EV': {'header_fill': '0F172A', 'sub_fill': '1E293B', 'accent_fill': '0284C7', 'card_bg': 'F0F9FF', 'filter_bg': 'E0F2FE', 'badge_top': 'DCFCE7', 'badge_lag': 'FFE4E6', 'title_color': '0F172A'},
    'ECOMMERCE_RETAIL': {'header_fill': '1E1B4B', 'sub_fill': '312E81', 'accent_fill': '4F46E5', 'card_bg': 'EEF2FF', 'filter_bg': 'E0E7FF', 'badge_top': 'DCFCE7', 'badge_lag': 'FFE4E6', 'title_color': '1E1B4B'},
    'LOGISTICS_SUPPLY': {'header_fill': '1E293B', 'sub_fill': '334155', 'accent_fill': '2563EB', 'card_bg': 'F8FAFC', 'filter_bg': 'E2E8F0', 'badge_top': 'DCFCE7', 'badge_lag': 'FFE4E6', 'title_color': '1E293B'},
    'GENERAL_ENTERPRISE': {'header_fill': '1F4E79', 'sub_fill': '2F5597', 'accent_fill': '41719C', 'card_bg': 'F2F5F9', 'filter_bg': 'D9E1F2', 'badge_top': 'E2EFDA', 'badge_lag': 'FCE4D6', 'title_color': '1F4E79'}
}

def build_mathematical_profile(df):
    math_schema = profile_algebraic_types(df)
    def calc_range(col):
        try: return df[col].max() - df[col].min()
        except: return 0
    math_schema['Additive_Measures'].sort(key=lambda x: calc_range(x), reverse=True)
    
    filtered_monetary = math_schema['Additive_Measures'][:3]
    intensive_cands = math_schema['Intensive_Measures'][:3]
    cat_pool = math_schema['Categorical_Dims']
    if not cat_pool:
        df['__Global_Cohort__'] = 'All Data'; cat_pool = ['__Global_Cohort__']
        
    macro_dim = cat_pool[0]
    sec_dim = cat_pool[1] if len(cat_pool) > 1 else macro_dim
    primary_measure = filtered_monetary[0] if filtered_monetary else (intensive_cands[0] if intensive_cands else df.columns[0])
    primary_agg = 'AVG' if primary_measure in intensive_cands else 'SUM'

    s2_unq = df[sec_dim].nunique() if sec_dim in df.columns else 0
    if 2 <= s2_unq <= 6: chart2_config = {'mode': 'STATUS_DOUGHNUT', 'dim': sec_dim, 'title': f'Distribution Segment ({sec_dim})'}
    elif intensive_cands: chart2_config = {'mode': 'INTENSIVE_AVG', 'dim': sec_dim, 'measure': intensive_cands[0], 'title': f'Average Distribution by {sec_dim}'}
    else: chart2_config = {'mode': 'CATEGORY_BAR', 'dim': sec_dim, 'title': f'Volume Distribution by {sec_dim}'}

    kpi_measures = []
    for m in filtered_monetary:
        if len(kpi_measures) < 2: kpi_measures.append((m, math_schema.get('Metric_Aggregations', {}).get(m, 'SUM')))
    for m in intensive_cands:
        if len(kpi_measures) < 3: kpi_measures.append((m, math_schema.get('Metric_Aggregations', {}).get(m, 'AVG')))
    if len(kpi_measures) < 3 and filtered_monetary:
        for m in filtered_monetary:
            if m not in [k[0] for k in kpi_measures] and len(kpi_measures) < 3: kpi_measures.append((m, math_schema.get('Metric_Aggregations', {}).get(m, 'SUM')))

    cols_text = ' '.join(df.columns).lower()
    if any(k in cols_text for k in ['patient', 'doctor', 'clinic', 'hospital']): theme = 'HEALTHCARE_CLINICAL'
    elif any(k in cols_text for k in ['vehicle', 'ev', 'mileage', 'battery', 'fleet']): theme = 'AUTOMOTIVE_EV'
    elif any(k in cols_text for k in ['order', 'sku', 'discount', 'sales']): theme = 'ECOMMERCE_RETAIL'
    elif any(k in cols_text for k in ['transit', 'delivery', 'distance', 'carrier', 'depot', 'run_km']): theme = 'LOGISTICS_SUPPLY'
    else: theme = 'GENERAL_ENTERPRISE'

    return {
        'sector': theme, 'title': 'AUTONOMOUS EXECUTIVE DASHBOARD', 'vol_label': 'TOTAL RECORDS / UNITS',
        'macro_dim': macro_dim, 'sec_dim': sec_dim, 'primary_measure': primary_measure, 'primary_agg': primary_agg,
        'chart1_type': 'col', 'chart2_config': chart2_config, 'kpi_measures': kpi_measures[:3],
        'palette': SECTOR_THEMES[theme],
        'temporal_dim': math_schema['Temporal_Dims'][0] if math_schema['Temporal_Dims'] else None
    }

def execute_math_agg(df, dim, metric, agg_type):
    valid_df = df[df[dim].notna() & (df[dim] != '')].copy()
    if valid_df.empty: return valid_df, pd.Series(dtype=float)
    valid_df[metric] = pd.to_numeric(valid_df[metric], errors='coerce')
    agg_func = 'mean' if agg_type == 'AVG' else 'sum'
    return valid_df, valid_df.groupby(dim)[metric].agg(agg_func).sort_values(ascending=False)

def diagnose_root_cause(df, dim1, metric, agg_type, opt_goal="MAX"):
    try:
        _res = execute_math_agg(df, dim1, metric, agg_type)
        agg_d1 = _res[1] if isinstance(_res, tuple) else _res
        if len(agg_d1) > 1:
            if opt_goal == "MIN":
                # For MIN metrics, highest cost cohort is the laggard / cost-drag
                lag_dim1, lag_val, tot_val = str(agg_d1.index[0]), float(agg_d1.iloc[0]), float(agg_d1.sum())
                lag_share = (lag_val / tot_val) * 100.0 if tot_val > 0 else 0
                return f"Primary cost concentration localized in '{lag_dim1}', driving {lag_share:.1f}% of total expenditure.", lag_dim1
            else:
                # For MAX metrics, lowest volume cohort is the underperformer
                lag_dim1, lag_val, tot_val = str(agg_d1.index[-1]), float(agg_d1.iloc[-1]), float(agg_d1.sum())
                lag_share = (lag_val / tot_val) * 100.0 if tot_val > 0 else 0
                return f"Negative variance heavily localized in '{lag_dim1}', capturing only {lag_share:.1f}% of throughput.", lag_dim1
        return 'Operational metrics within nominal statistical tolerance.', None
    except:
        return f"Isolated structural deviation in '{dim1}' baseline cohorts.", None

def generate_nlg_executive_summary(df, profile):
    dim1, metric, agg_type = profile['macro_dim'], profile['primary_measure'], profile['primary_agg']
    opt_goal = profile.get('optimization_goal', 'MAX')
    if not dim1 or not metric or df.empty or dim1 not in df.columns:
        return ["• Pipeline processed securely."]
    try:
        valid_df, agg_d1 = execute_math_agg(df, dim1, metric, agg_type)
        if agg_d1.empty:
            return ["• Zero net measure variance across dimensions."]

        total_val = float(valid_df[metric].mean() if agg_type == 'AVG' else valid_df[metric].sum())
        is_ratio = (agg_type == 'AVG' and valid_df[metric].max() <= 1.0)
        diag_str, lag_dim = diagnose_root_cause(valid_df, dim1, metric, agg_type, opt_goal)

        if opt_goal == "MIN":
            top_d1_name = str(agg_d1.index[-1]) # Lowest expenditure is top efficiency leader
            action_str = f"Action Directive: Operational efficiency benchmark governed by '{top_d1_name}'. Recommended budget containment audit for high-cost cohort '{lag_dim}'." if lag_dim else "Action Directive: Cost structures operating within target parameters."
            lead_line = f"• Core Performance: Aggregate {metric.replace('_', ' ')} index reaches {format_compact_num(total_val, is_ratio, agg_type=='AVG')}, anchored by efficiency leader '{top_d1_name}'."
        else:
            top_d1_name = str(agg_d1.index[0]) # Highest volume is growth leader
            action_str = f"Action Directive: Immediate operational audit recommended for '{lag_dim}' to optimize resource allocation and prevent further lag." if lag_dim else "Action Directive: Maintain current operational bandwidth and scale successful cohorts."
            lead_line = f"• Core Performance: Aggregate {metric.replace('_', ' ')} index reaches {format_compact_num(total_val, is_ratio, agg_type=='AVG')}, spearheaded by '{top_d1_name}'."

        lines = [
            lead_line,
            f"• Root-Cause & Strategic Diagnostic: {diag_str}",
            f"• {action_str}"
        ]
        return lines
    except:
        return ["• Executive Overview: Pipeline processed with aggregate metrics."]

# ==============================================================================
# MODULE 4: PREDICTIVE FORECAST ENGINE (DYNAMIC RESOLUTION)
# ==============================================================================

def generate_predictive_forecast_sheet(wb, df, profile):
    metric_col, date_col, m_agg = profile.get('primary_measure'), profile.get('temporal_dim'), profile.get('primary_agg')
    if not metric_col or metric_col not in df.columns: return
    
    df_temp, has_date = df.copy(), False
    if date_col:
        df_temp['__dt'] = pd.to_datetime(df_temp[date_col], errors='coerce')
        valid_dt = df_temp.dropna(subset=['__dt', metric_col]).copy()
        if len(valid_dt) >= 2:
            unique_months = valid_dt['__dt'].dt.to_period('M').nunique()
            freq = 'M' if unique_months >= 3 else 'D'
            
            valid_dt['__period'] = valid_dt['__dt'].dt.to_period(freq)
            valid_dt[metric_col] = pd.to_numeric(valid_dt[metric_col], errors='coerce')
            agg_f = 'mean' if m_agg == 'AVG' else 'sum'
            periodic = valid_dt.groupby('__period')[metric_col].agg(agg_f).reset_index()
            periodic.rename(columns={'__period': 'Period', metric_col: 'Actual'}, inplace=True)
            periodic['Period'] = periodic['Period'].astype(str)
            has_date = True
            
    if not has_date:
        n = len(df_temp)
        if n < 2: return
        df_temp['__cohort'] = pd.qcut(range(n), q=min(6, n), labels=[f"T-{min(6, n) - i}" for i in range(min(6, n))])
        df_temp[metric_col] = pd.to_numeric(df_temp[metric_col], errors='coerce')
        periodic = df_temp.groupby('__cohort', observed=False)[metric_col].agg('mean' if m_agg == 'AVG' else 'sum').reset_index()
        periodic.rename(columns={'__cohort': 'Period', metric_col: 'Actual'}, inplace=True)
        periodic['Period'] = periodic['Period'].astype(str)
        
    periodic = periodic.tail(18) if len(periodic) > 18 else periodic
    y_all = periodic['Actual'].values.astype(float)
    fit_y = y_all[-6:] if len(y_all) >= 6 else y_all
    slope, _ = np.polyfit(np.arange(len(fit_y)), fit_y, 1) if len(fit_y) > 1 else (0.0, 0)
    
    last_act = float(fit_y[-1])
    forecast_y = [last_act + (slope * k * (0.85**k)) for k in range(1, 4)]
    
    final_proj = float(forecast_y[-1])
    growth_pct = ((final_proj - last_act) / abs(last_act)) * 100 if last_act != 0 else 0.0
    
    ws_fc = wb.create_sheet(title="Executive_Forecast")
    ws_fc.sheet_view.showGridLines = False
    ws_fc.merge_cells('A1:N1')
    ws_fc['A1'] = f"  PREDICTIVE FORECAST ({metric_col})"
    ws_fc['A1'].font = Font(size=11, bold=True, color="FFFFFF")
    ws_fc['A1'].fill = PatternFill(start_color=profile['palette']['header_fill'], fill_type="solid")
    ws_fc.row_dimensions[1].height = 26
    
    ws_fc.merge_cells('E2:K2')
    opt_goal = profile.get('optimization_goal', 'MAX')
    is_favorable = (growth_pct <= 0) if opt_goal == 'MIN' else (growth_pct >= 0)
    if opt_goal == 'MIN':
        tag = 'Cost Savings / Efficiency Gain' if growth_pct <= 0 else 'Cost Inflation / Overrun'
    else:
        tag = 'Growth / Volume Gain' if growth_pct >= 0 else 'Contraction / Decline'
    arrow = '▼' if growth_pct < 0 else ('▲' if growth_pct > 0 else '•')
    trend_txt = f"  PROJECTED TRAJECTORY: {arrow} {abs(growth_pct):.1f}% {tag} expected over next 3 cycles."
    ws_fc['E2'] = trend_txt
    ws_fc['E2'].font = Font(bold=True, size=10, color='065F46' if is_favorable else '991B1B')
    ws_fc['E2'].fill = PatternFill(start_color='ECFDF5' if is_favorable else 'FEF2F2', fill_type='solid')
    
    for c_letter in ['A','B','C','D','E','F','G']: ws_fc.column_dimensions[c_letter].width = 15.5
    
    t_head_font = Font(bold=True)
    t_fill = PatternFill(start_color="F1F5F9", fill_type="solid")
    ws_fc['A3'], ws_fc['B3'], ws_fc['C3'] = "Timeline Phase", "Status", "Value Metric"
    for cell in ['A3', 'B3', 'C3']:
        ws_fc[cell].font = t_head_font
        ws_fc[cell].fill = t_fill
    
    curr_r = 4
    for p, act in zip(periodic['Period'], periodic['Actual']):
        ws_fc[f'A{curr_r}'], ws_fc[f'B{curr_r}'], ws_fc[f'C{curr_r}'] = p, "Historical", act
        ws_fc[f'C{curr_r}'].number_format = get_math_format(df, metric_col, m_agg)
        curr_r += 1
    for idx, fv in enumerate(forecast_y, 1):
        ws_fc[f'A{curr_r}'], ws_fc[f'B{curr_r}'], ws_fc[f'C{curr_r}'] = f"Proj +{idx}", "Projected", fv
        ws_fc[f'A{curr_r}'].font = Font(bold=True, color="0284C7")
        ws_fc[f'C{curr_r}'].number_format = get_math_format(df, metric_col, m_agg)
        curr_r += 1
        
    c_fc = LineChart()
    c_fc.title, c_fc.style, c_fc.height, c_fc.width = f"Projected Trend Analysis ({metric_col})", 10, 8.5, 14.5
    c_fc.y_axis.majorGridlines = None
    c_fc.legend = None
    c_fc.y_axis.scaling.min = 0
    c_fc.add_data(Reference(ws_fc, min_col=3, min_row=3, max_row=curr_r-1), titles_from_data=True)
    c_fc.set_categories(Reference(ws_fc, min_col=1, min_row=4, max_row=curr_r-1))
    s1 = c_fc.series[0]
    s1.graphicalProperties.line.solidFill = "0284C7"
    ws_fc.add_chart(c_fc, "E4")

# ==============================================================================
# MODULE 5: PRODUCTION DASHBOARD ENGINE (EXECUTIVE RENDERER)
# ==============================================================================

def get_agg_form(func, col, d1, d2, r_start, r_end):
    core_f = f'IF(AND(Executive_Dashboard!$J$1="All", Executive_Dashboard!$M$1="All"), {func}(Cleaned_Data!{col}{r_start}:{col}{r_end}), IF(Executive_Dashboard!$J$1="All", {func}IF(Cleaned_Data!{d2}{r_start}:{d2}{r_end}, Executive_Dashboard!$M$1, Cleaned_Data!{col}{r_start}:{col}{r_end}), IF(Executive_Dashboard!$M$1="All", {func}IF(Cleaned_Data!{d1}{r_start}:{d1}{r_end}, Executive_Dashboard!$J$1, Cleaned_Data!{col}{r_start}:{col}{r_end}), {func}IFS(Cleaned_Data!{col}{r_start}:{col}{r_end}, Cleaned_Data!{d1}{r_start}:{d1}{r_end}, Executive_Dashboard!$J$1, Cleaned_Data!{d2}{r_start}:{d2}{r_end}, Executive_Dashboard!$M$1))))'
    return f'=IFERROR({core_f}, 0)'


def resolve_json_contract(file_path: str, df: pd.DataFrame, math_profile: dict) -> dict:
    os.makedirs("contracts", exist_ok=True)
    stem = Path(file_path).stem.replace("_Gatekeeper_Dashboard", "").replace("_Cleaned", "")
    contract_path = Path("contracts") / f"{stem}_contract.json"
    
    # Return existing contract if present (Allows zero-code manual overrides)
    if contract_path.exists():
        try:
            with open(contract_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    # --- FULL AUTONOMOUS ZERO-TOUCH GEMINI AI TRIGGER ---
    try:
        import ai_semantic_agent
        print(f"[AI AGENT] Auto-generating semantic contract via Gemini for: {file_path}")
        ai_semantic_agent.run_agent(file_path)
        if contract_path.exists():
            with open(contract_path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as ai_err:
        print(f"[AI FALLBACK] Proceeding with heuristic contract. Reason: {ai_err}")
            
    # Heuristic determination of currency & polarity
    p_meas = math_profile.get('primary_measure', '')
    p_meas_lower = p_meas.lower()
    
    # Currency determination
    if any(k in p_meas_lower for k in ["inr", "rupee", "rs"]):
        curr = ""
    elif any(k in p_meas_lower for k in ["usd", "dollar", "cost", "price", "revenue", "spend"]):
        curr = "$"
    elif any(k in p_meas_lower for k in ["eur", "euro"]):
        curr = "€"
    else:
        curr = ""
        
    # Optimization Goal (Min vs Max)
    if any(k in p_meas_lower for k in ["cost", "expense", "defect", "loss", "delay", "error", "churn", "maintenance"]):
        opt_goal = "MIN"
    else:
        opt_goal = "MAX"
        
    contract = {
        "dataset_signature": stem,
        "theme_palette": math_profile.get("sector", "GENERAL_ENTERPRISE"),
        "currency_symbol": curr,
        "primary_measure": {
            "column": p_meas,
            "aggregation": math_profile.get("primary_agg", "SUM"),
            "optimization_goal": opt_goal
        },
        "dimensions": {
            "macro_dimension": math_profile.get("macro_dim"),
            "secondary_dimension": math_profile.get("sec_dim")
        },
        "kpi_measures": [m[0] if isinstance(m, (list, tuple)) else m for m in math_profile.get("kpi_measures", [])]
    }
    
    with open(contract_path, "w", encoding="utf-8") as f:
        json.dump(contract, f, indent=2)
    print(f"[CONTRACT] Generated JSON Contract: {contract_path}")
    return contract

def build_universal_dashboard(df, profile, output_path, dropped_count=0):
    contract = resolve_json_contract(output_path, df, profile)
    curr_sym = contract.get("currency_symbol", "")
    opt_goal = contract.get("primary_measure", {}).get("optimization_goal", "MAX")
    profile["optimization_goal"] = opt_goal
    m_meas = contract.get("primary_measure", {}).get("column", profile.get("primary_measure", ""))
    wb = openpyxl.Workbook()
    pal = profile['palette']
    
    ws_data = wb.active
    ws_data.title = "Cleaned_Data"
    headers = list(df.columns)
    ws_data.append(headers)
    for row in df.itertuples(index=False, name=None): ws_data.append(list(row))
    num_rows = len(df) + 1

    # FACT 1: FIXED PROFESSIONAL WIDTH (Zero Processing Delay)
    for i in range(len(headers)):
        ws_data.column_dimensions[get_column_letter(i+1)].width = 18

    # FACT 2: NATIVE PREMIUM STYLING (No XML Corruption)
    from openpyxl.styles import PatternFill, Font
    for cx in ws_data[1]:
        cx.fill = PatternFill(start_color="1F4E78", fill_type="solid")
        cx.font = Font(color="FFFFFF", bold=True)
    ws_data.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{num_rows}"
    ws_data.freeze_panes = "A2"

    ws_calc = wb.create_sheet(title="Calculations")
    raw_d1, raw_d2 = profile['macro_dim'], profile['sec_dim']
    m1_col, m1_agg = profile['primary_measure'], profile['primary_agg']

    # Auto-balance cardinality: Parent (Zone, 2-6 unique) vs Child (Branch/Entity, 6-15 unique)
    u1_cnt = df[raw_d1].dropna().nunique() if raw_d1 in df.columns else 0
    u2_cnt = df[raw_d2].dropna().nunique() if raw_d2 in df.columns else 0

    if u1_cnt < u2_cnt:
        parent_dim, child_dim = raw_d1, raw_d2
    else:
        parent_dim, child_dim = raw_d2, raw_d1

    dim1_col, dim2_col = child_dim, parent_dim
    d1_let = get_column_letter(headers.index(dim1_col) + 1)  # Child (Branch / Entity)
    d2_let = get_column_letter(headers.index(dim2_col) + 1)  # Parent (Zone / Category)
    m1_let = get_column_letter(headers.index(m1_col) + 1)    # Primary Measure Metric

    unique_dim1 = [str(x) for x in df[dim1_col].dropna().unique()][:12]
    if dim1_col:
        ws_calc['A1'], ws_calc['B1'] = str(dim1_col), str(m1_col)
        agg_str = 'AVERAGE' if m1_agg == 'AVG' else 'SUM'
        for i, val in enumerate(unique_dim1, start=2):
            ws_calc[f'A{i}'] = str(val)
            ws_calc[f'B{i}'] = f'=IFERROR(IF(Executive_Dashboard!$J$1="All", {agg_str}IFS(Cleaned_Data!{m1_let}2:{m1_let}{num_rows}, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Calculations!A{i}), {agg_str}IFS(Cleaned_Data!{m1_let}2:{m1_let}{num_rows}, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Calculations!A{i}, Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Executive_Dashboard!$J$1)), 0)'

    unique_dim2 = [str(x) for x in df[dim2_col].dropna().unique()][:15]
    if unique_dim2:
        ws_calc['D1'], ws_calc['E1'] = str(dim2_col), "Volume"
        for i, val in enumerate(unique_dim2, start=2):
            ws_calc[f'D{i}'] = str(val)
            ws_calc[f'E{i}'] = f'=IFERROR(IF(Executive_Dashboard!$M$1="All", COUNTIF(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Calculations!D{i}), COUNTIFS(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Calculations!D{i}, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Executive_Dashboard!$M$1)), 0)'

    ws_dash = wb.create_sheet(title="Executive_Dashboard", index=0)
    ws_dash.sheet_view.showGridLines = False
    for c_letter in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N']:
        ws_dash.column_dimensions[c_letter].width = 12.5
    for r in range(1, 45):
        for c in range(1, 16):
            ws_dash.cell(row=r, column=c).fill = PatternFill(start_color="FFFFFF", fill_type="solid")
    ws_calc.sheet_state = 'hidden'

    f_head = PatternFill(start_color=pal['header_fill'], fill_type="solid")
    f_sub = PatternFill(start_color=pal['sub_fill'], fill_type="solid")
    f_card = PatternFill(start_color=pal['card_bg'], fill_type="solid")
    t_border = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'), top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

    ws_dash.merge_cells('A1:G1')
    ws_dash['A1'] = f"  {profile['title']}"
    ws_dash['A1'].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws_dash['A1'].fill = f_head
    ws_dash['A1'].alignment = Alignment(vertical="center")

    # Dropdown 1 ($J$1) = Parent (Zone), Dropdown 2 ($M$1) = Child (Branch)
    for sc_start, sc_end, s_name, s_col, dv_list in [('H', 'I', dim2_col, 'J', unique_dim2), ('K', 'L', dim1_col, 'M', unique_dim1)]:
        ws_dash.merge_cells(f'{sc_start}1:{sc_end}1')
        ws_dash[f'{sc_start}1'] = f"Filter {s_name}:"
        ws_dash[f'{sc_start}1'].font = Font(name="Calibri", size=8.5, bold=True, color=pal['title_color'])
        ws_dash[f'{sc_start}1'].fill = f_card
        ws_dash[f'{s_col}1'] = "All"
        ws_dash[f'{s_col}1'].fill = PatternFill(start_color=pal['filter_bg'], fill_type="solid")
        ws_dash[f'{s_col}1'].border = t_border
        if dv_list:
            dv = DataValidation(type="list", formula1=f'"{",".join(["All"] + dv_list[:15])}"', allow_blank=True)
            ws_dash.add_data_validation(dv)
            dv.add(f'{s_col}1')

    ws_dash.merge_cells('A2:F2')
    ws_dash['A2'] = f"  DATA GOVERNANCE: v71.0 Final Enterprise Master | Stamp: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws_dash['A2'].font = Font(size=7.5, bold=True, color="475569")
    ws_dash['A2'].fill = PatternFill(start_color="F1F5F9", fill_type="solid")

    cards_data = [(profile['vol_label'], f'=IFERROR(IF(AND($J$1="All", $M$1="All"), COUNTA(Cleaned_Data!A2:A{num_rows}), IF($J$1="All", COUNTIF(Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, $M$1), IF($M$1="All", COUNTIF(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, $J$1), COUNTIFS(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, $J$1, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, $M$1)))), 0)', '#,##0')]
    for metric, agg_type in profile['kpi_measures']:
        c_let = get_column_letter(headers.index(metric) + 1)
        lbl = f"{agg_type} {str(metric).upper().replace('_', ' ')}"
        func = 'AVERAGE' if agg_type in ['AVG', 'AVERAGE', 'MEDIAN'] else 'SUM'
        fmt = get_math_format(df, metric, agg_type)
        form = get_agg_form(func, c_let, d2_let, d1_let, 2, num_rows)
        cards_data.append((lbl, form, fmt))

    mid_r = max(2, (num_rows - 2) // 2 + 2)
    v_f1 = f'IF(AND(Executive_Dashboard!$J$1="All", Executive_Dashboard!$M$1="All"), COUNTA(Cleaned_Data!A2:A{mid_r-1}), IF(Executive_Dashboard!$J$1="All", COUNTIF(Cleaned_Data!{d1_let}2:{d1_let}{mid_r-1}, Executive_Dashboard!$M$1), IF(Executive_Dashboard!$M$1="All", COUNTIF(Cleaned_Data!{d2_let}2:{d2_let}{mid_r-1}, Executive_Dashboard!$J$1), COUNTIFS(Cleaned_Data!{d2_let}2:{d2_let}{mid_r-1}, Executive_Dashboard!$J$1, Cleaned_Data!{d1_let}2:{d1_let}{mid_r-1}, Executive_Dashboard!$M$1))))'
    v_f2 = f'IF(AND(Executive_Dashboard!$J$1="All", Executive_Dashboard!$M$1="All"), COUNTA(Cleaned_Data!A{mid_r}:A{num_rows}), IF(Executive_Dashboard!$J$1="All", COUNTIF(Cleaned_Data!{d1_let}{mid_r}:{d1_let}{num_rows}, Executive_Dashboard!$M$1), IF(Executive_Dashboard!$M$1="All", COUNTIF(Cleaned_Data!{d2_let}{mid_r}:{d2_let}{num_rows}, Executive_Dashboard!$J$1), COUNTIFS(Cleaned_Data!{d2_let}{mid_r}:{d2_let}{num_rows}, Executive_Dashboard!$J$1, Cleaned_Data!{d1_let}{mid_r}:{d1_let}{num_rows}, Executive_Dashboard!$M$1))))'
    ws_calc['G2'] = f'=IFERROR({v_f1}, 0)'
    ws_calc['H2'] = f'=IFERROR({v_f2}, 0)'
    ws_calc['I2'] = '=IFERROR((H2 - G2) / ABS(G2), 0)'

    for m_idx, (m_col_k, m_agg_k) in enumerate(profile['kpi_measures'][:3], start=3):
        cl = get_column_letter(headers.index(m_col_k) + 1)
        func = 'AVERAGE' if m_agg_k in ['AVG', 'AVERAGE', 'MEDIAN'] else 'SUM'
        ws_calc[f'G{m_idx}'] = get_agg_form(func, cl, d2_let, d1_let, 2, mid_r-1)
        ws_calc[f'H{m_idx}'] = get_agg_form(func, cl, d2_let, d1_let, mid_r, num_rows)
        ws_calc[f'I{m_idx}'] = f'=IFERROR((H{m_idx} - G{m_idx}) / ABS(G{m_idx}), 0)'

    ws_calc['I6'] = '=IFERROR(AVERAGE(I2:I5), 0)'
    ws_dash.merge_cells('G2:N2')
    ws_dash['G2'] = f'= "📈 BUSINESS HEALTH INDEX: " & ROUND(MIN(100, MAX(0, 50 + (Calculations!I6*100))), 0) & "/100   |   " & IF(Calculations!I6>0, "Expanding ▲", IF(Calculations!I6<0, "Contracting ▼", "Stagnant ◂▸"))'
    ws_dash['G2'].font = Font(size=8.5, bold=True, color='065F46')
    ws_dash['G2'].fill = PatternFill(start_color='ECFDF5', fill_type="solid")

    c_slots = [('A','C'), ('D','F'), ('H','J'), ('L','N')]
    for idx, (title, formula, num_fmt) in enumerate(cards_data[:4]):
        cs, ce = c_slots[idx]
        ws_dash.merge_cells(f'{cs}3:{ce}3')
        ws_dash[f'{cs}3'] = title
        ws_dash[f'{cs}3'].font = Font(name="Calibri", size=8.5, bold=True, color="FFFFFF")
        ws_dash[f'{cs}3'].fill = f_sub
        ws_dash[f'{cs}3'].alignment = Alignment(horizontal="center")
        
        ws_dash.merge_cells(f'{cs}4:{ce}4')
        ws_dash[f'{cs}4'] = formula
        ws_dash[f'{cs}4'].font = Font(size=12, bold=True, color=pal['title_color'])
        ws_dash[f'{cs}4'].fill = f_card
        ws_dash[f'{cs}4'].number_format = num_fmt
        ws_dash[f'{cs}4'].alignment = Alignment(horizontal="center")
        
        ws_dash.merge_cells(f'{cs}5:{ce}5')
        ws_dash[f'{cs}5'] = f'=IF(Calculations!I{idx+2}=0, "▶ Steady", IF(Calculations!I{idx+2}>0, "▲ +" & TEXT(Calculations!I{idx+2}, "0.0%") & " Mom", "▼ " & TEXT(Calculations!I{idx+2}, "0.0%") & " Drag"))'
        ws_dash[f'{cs}5'].font = Font(size=7.5, bold=True, color="334155")
        ws_dash[f'{cs}5'].fill = f_card
        ws_dash[f'{cs}5'].alignment = Alignment(horizontal="center")

    if unique_dim1:
        c1 = BarChart()
        c1.style, c1.height, c1.width = 10, 6.6, 13.5
        c1.legend = None
        c1.dataLabels = None
        c1.y_axis.delete = False
        c1.x_axis.delete = False
        c1.y_axis.majorGridlines = None
        c1.x_axis.majorGridlines = None
        # Safe Signal Unpacking
        agg_raw = execute_math_agg(df, dim1_col, m1_col, m1_agg)
        agg_preview = agg_raw[0] if isinstance(agg_raw, tuple) else agg_raw
        try:
            has_c1_neg = bool((agg_preview.select_dtypes(include=['number']) < 0).any().any()) if isinstance(agg_preview, pd.DataFrame) else False
            agg_max = float(agg_preview.max())
            agg_min = float(agg_preview.min())
        except Exception:
            has_c1_neg = bool((pd.to_numeric(df[m1_col], errors='coerce') < 0).any()) if m1_col in df.columns else False
            agg_max = float(df[m1_col].max()) if m1_col in df.columns else 0.0
            agg_min = float(df[m1_col].min()) if m1_col in df.columns else 0.0

        cardinality_c1 = len(unique_dim1)
        sec_type = profile.get('sector', 'GENERAL_ENTERPRISE')

        # Slot 1 Domain-Aware Decision
        if has_c1_neg:
            c1.type = 'col'
            c1.title = f"Net Variance: {m1_col.replace('_', ' ')} by {dim1_col.replace('_', ' ')}"
            c1.y_axis.scaling.min = agg_min * 1.2
            c1.y_axis.scaling.max = agg_max * 1.2 if agg_max > 0 else 0
        elif cardinality_c1 > 6 or sec_type in ['ECOMMERCE_RETAIL', 'LOGISTICS_SUPPLY']:
            c1.type = 'bar'  # Horizontal prevents label overlap
            c1.title = f"Performance Ranking: {m1_col.replace('_', ' ')} by {dim1_col.replace('_', ' ')}"
            c1.y_axis.scaling.min = 0
            if agg_max > 0: c1.y_axis.scaling.max = agg_max * 1.25
        else:
            c1.type = 'col'
            c1.title = f"Global Baseline: {m1_col.replace('_', ' ')} by {dim1_col.replace('_', ' ')}"
            c1.y_axis.scaling.min = 0
            if agg_max > 0: c1.y_axis.scaling.max = agg_max * 1.25

        if agg_max >= 1_000_000:
            c1.y_axis.number_format = '$#,##0,, "M"'
            c1.y_axis.sourceLinked = False
        c1.add_data(Reference(ws_calc, min_col=2, min_row=1, max_row=len(unique_dim1)+1), titles_from_data=True)
        c1.set_categories(Reference(ws_calc, min_col=1, min_row=2, max_row=len(unique_dim1)+1))
        ws_dash.add_chart(c1, 'A6')

    if unique_dim2:
        c2 = DoughnutChart() if profile['chart2_config']['mode'] == 'STATUS_DOUGHNUT' else BarChart()
        c2.title, c2.style, c2.height, c2.width = profile['chart2_config']['title'], 10, 6.6, 12.8
        c2.dataLabels = DataLabelList()
        if isinstance(c2, DoughnutChart):
            c2.dataLabels.showPercent = True
            c2.dataLabels.showVal = False
            c2.holeSize, c2.legend = 65, Legend()
            c2.legend.legendPos = "r"
        else:
            sec_type = profile.get('sector', 'GENERAL_ENTERPRISE')
            if sec_type in ['ECOMMERCE_RETAIL', 'LOGISTICS_SUPPLY'] and len(unique_dim2) > 5:
                c2.type = "bar"
            else:
                c2.type = "col"
            c2.y_axis.delete = False
            c2.y_axis.majorGridlines = None
            c2.x_axis.majorGridlines = None
            c2.legend = Legend()
            c2.legend.legendPos = "r"
            c2.y_axis.title = f"{m_meas} ({curr_sym})" if curr_sym else f"{m_meas}"
            c2.dataLabels.showVal = True
            c2.dataLabels.showVal = False
            c2.dataLabels.showPercent = False
            c2.legend = Legend()
            c2.legend.legendPos = "b"
        c2.dataLabels.showCatName = False
        c2.dataLabels.showSerName = False
        c2.dataLabels.showLegendKey = False
            
        c2.add_data(Reference(ws_calc, min_col=5, min_row=1, max_row=len(unique_dim2)+1), titles_from_data=True)
        c2.series[0].cat = AxDataSource(strRef=StrRef(f"'Calculations'!$D$2:$D{len(unique_dim2)+1}", strCache=StrData(pt=[StrVal(idx=i, v=str(x)) for i, x in enumerate(unique_dim2)])))
        dim2_name = profile['chart2_config'].get('dim', dim2_col)
        v_counts = df[dim2_name].value_counts()
        c2_vals = [int(v_counts.get(x, 0)) for x in unique_dim2]
        c2.series[0].val = NumDataSource(numRef=NumRef(f=f"'Calculations'!$E$2:$E${len(unique_dim2)+1}", numCache=NumData(pt=[NumVal(idx=i, v=val) for i, val in enumerate(c2_vals)])))
        ws_dash.add_chart(c2, "H6")

    z3_start = 21
    ws_dash.merge_cells(f'A{z3_start}:F{z3_start}')
    ws_dash[f'A{z3_start}'] = "EXECUTIVE BENCHMARK MATRIX"
    ws_dash[f'A{z3_start}'].font = Font(size=8.5, bold=True, color="FFFFFF")
    ws_dash[f'A{z3_start}'].fill = f_sub
    
    fmt_m1 = get_math_format(df, m1_col, m1_agg)
    calc_end_r = 1 + len(unique_dim1)
    top_func = 'SMALL' if opt_goal == 'MIN' else 'LARGE'
    lag_func = 'LARGE' if opt_goal == 'MIN' else 'SMALL'
    matrix_specs = [
        ('Top', top_func, 1, pal['badge_top'], '\u2605'),
        ('Top', top_func, 2, pal['badge_top'], '\u2605'),
        ('Lag', lag_func, 2 if len(unique_dim1) > 2 else 1, pal['badge_lag'], '\u25bc'),
        ('Lag', lag_func, 1, pal['badge_lag'], '\u25bc'),
    ]
    r_idx = z3_start + 1
    for lbl, func, k_rank, bg_fill, sym in matrix_specs[:min(len(unique_dim1), 4)]:
        ws_dash.merge_cells(f'A{r_idx}:D{r_idx}')
        ws_dash.merge_cells(f'E{r_idx}:F{r_idx}')
        ws_dash[f'A{r_idx}'] = f'="{sym} {lbl}: " & IFERROR(INDEX(Calculations!$A$2:$A${calc_end_r}, MATCH({func}(Calculations!$B$2:$B${calc_end_r}, {k_rank}), Calculations!$B$2:$B${calc_end_r}, 0)), "N/A")'
        ws_dash[f'E{r_idx}'] = f'=IFERROR({func}(Calculations!$B$2:$B${calc_end_r}, {k_rank}), 0)'
        ws_dash[f'A{r_idx}'].fill = PatternFill(start_color=bg_fill, fill_type="solid")
        ws_dash[f'E{r_idx}'].number_format = fmt_m1
        r_idx += 1

    ws_dash.merge_cells('A27:N27')
    ws_dash['A27'] = "  GLOBAL BASELINE EXECUTIVE AUDIT (PORTFOLIO BENCHMARK)"
    ws_dash['A27'].font = Font(size=8.5, bold=True, color="FFFFFF")
    ws_dash['A27'].fill = f_head
    
    for s_idx, line in enumerate(generate_nlg_executive_summary(df, profile), start=28):
        ws_dash.merge_cells(f'A{s_idx}:N{s_idx}')
        ws_dash[f'A{s_idx}'] = f"  {line}"
        ws_dash[f'A{s_idx}'].font = Font(size=8.5)
        ws_dash[f'A{s_idx}'].fill = PatternFill(start_color="F8FAFC", fill_type="solid")

    generate_predictive_forecast_sheet(wb, df, profile)

    # Dynamically apply currency/number formats from contract (No hardcoded dollar overwrite)
    pass

    try:
        wb.save(output_path)
        print(f"\n[SUCCESS] Universal Gatekeeper Dashboard generated: {output_path}")
    except Exception as e:
        print(f"\n[Error saving file]: {e}")

def process_pipeline(raw_input_path):
    input_path = clean_file_path(raw_input_path)
    started_at = time.perf_counter()
    print("\n" + "="*68 + "\n   UNIVERSAL 2FA AUTONOMOUS BI ENGINE v71.0 (ENTERPRISE MASTER) \n" + "="*68)
    
    raw_df = ingest_file(input_path)
    if raw_df is None or raw_df.empty: return print("[Error] Invalid data or incorrect file path.")
    clean_df, dropped = clean_dataframe(raw_df)
    global CURRENT_INPUT_FILE; CURRENT_INPUT_FILE = input_path; validation = validate_with_circuit_breaker(
        clean_df,
        output_dir='reports',
        total_rows_ingested=len(raw_df),
        started_at=started_at,
    )
    clean_df = validation.dataframe
    
    schema = profile_algebraic_types(clean_df)
    print("\n--- ZERO-GUESSWORK SCHEMA AUDIT ---")
    print(f" -> Additive Measures (SUM) : {schema['Additive_Measures']}")
    print(f" -> Intensive Rates (AVG)   : {schema['Intensive_Measures']}")
    print(f" -> Categorical Dims        : {schema['Categorical_Dims'][:5]}")
    print(f" -> Temporal Dims (Locked)  : {schema['Temporal_Dims'][:5]}")
    print("---------------------------------\n")
    
    profile = build_mathematical_profile(clean_df)
    os.makedirs('reports', exist_ok=True); base_stem = os.path.splitext(os.path.basename(input_path))[0]; output_name = os.path.join('reports', f'{base_stem}_Gatekeeper_Dashboard.xlsx')
    build_universal_dashboard(clean_df, profile, output_name, dropped)
    _write_validation_log(
        'reports',
        base_stem,
        len(raw_df),
        len(clean_df),
        validation.soft_imputations,
        validation.fatal_corrupt_rows,
        max(0.0, time.perf_counter() - started_at),
        validation.status,
    )
    # --- AUTOMATED CLEAN PARQUET STORAGE ---
    try:
        in_p = raw_input_path if 'raw_input_path' in locals() else (input_path if 'input_path' in locals() else 'dataset')
        os.makedirs('clean_data', exist_ok=True); pq_path = os.path.join('clean_data', f'{base_stem}_Cleaned.parquet')
        clean_df.to_parquet(pq_path, index=False)
        print(f"[STORAGE] Clean Parquet Exported: {pq_path}")
    except Exception as _err:
        print(f"[STORAGE ERROR] Parquet save failed: {_err}")

if __name__ == "__main__":
    try:
        file_input = sys.argv[1] if len(sys.argv) > 1 else input("Enter CSV/Excel file path: ")
        process_pipeline(file_input)
    except Exception as e:
        import traceback
        print("\n" + "!"*60)
        print("   SYSTEM CRASH DETECTED (MATHEMATICAL ENGINE HALTED)   ")
        print("!"*60)
        traceback.print_exc()
        print("!"*60)
        
  
