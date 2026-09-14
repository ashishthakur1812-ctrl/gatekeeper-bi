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

GRACEFUL_DEFAULTS: Dict[str, object] = {
    'city': 'Unknown',
    'category': 'Uncategorized',
    'notes': '',
}
FATAL_CORRUPTION_THRESHOLD_PERCENT = 5.0
PRIMARY_KEY_PATTERN = r'(?:^|[_\s-])(?:id|key|uuid|code|sku|account|record)(?:[_\s-]|$)'
NON_NEGATIVE_METRIC_PATTERN = r'(?:amount|amt|metric|measure|value|revenue|price|cost|total|quantity|qty|count|score|salary|ctc)'

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
    fin_pat = r'(?:revenue|profit|margin|loss|variance|pnl|ebitda|delta|net|earnings|income)'
    metric_columns = metric_columns[~metric_columns.str.contains(fin_pat, case=False, regex=True)]
    return metric_columns

def _write_validation_log(*args, **kwargs):
    if len(args) >= 8:
        output_dir, base_stem, raw_count, clean_count, soft_imp, fatal_corr, proc_time, status = args[:8]
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
    with open(os.path.join(output_dir, log_file_name), 'w', encoding='utf-8') as f:
        f.write(f"Dataset Signature: {base_stem}\nTotal Rows Ingested: {raw_count}\nTotal Rows Exported: {clean_count}\nSoft Imputations Applied: {soft_imp}\nFatal Corrupt Rows Blocked: {fatal_corr}\nTotal Processing Time (seconds): {float(proc_time):.2f}\nStatus: {status}\n")

def validate_with_circuit_breaker(df: pd.DataFrame, output_dir: str, total_rows_ingested: Optional[int] = None, critical_columns: Optional[Sequence[str]] = None, primary_key: Optional[str] = None, numeric_bound_columns: Optional[Sequence[str]] = None, started_at: Optional[float] = None) -> ValidationResult:
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
    configured_critical = _matching_columns(validation_df.columns, inferred_critical if critical_columns is None else critical_columns)
    critical = configured_critical.union(pd.Index([resolved_key])) if resolved_key else configured_critical
    critical_null_mask = validation_df.loc[:, critical].isna().any(axis=1) if len(critical) else pd.Series(False, index=validation_df.index)

    duplicate_mask = validation_df[resolved_key].duplicated(keep=False) if resolved_key else pd.Series(False, index=validation_df.index)
    bound_columns = _matching_columns(validation_df.columns, numeric_bound_columns or _resolve_non_negative_metrics(validation_df, resolved_key))
    bound_mask = validation_df.loc[:, bound_columns].lt(0).any(axis=1) if len(bound_columns) else pd.Series(False, index=validation_df.index)
    fatal_mask = critical_null_mask | duplicate_mask | bound_mask

    fatal_corrupt_rows = int(fatal_mask.sum())
    quarantine_export_df = validation_df.loc[fatal_mask].copy() if fatal_corrupt_rows > 0 else validation_df.iloc[0:0].copy()
    total_rows = len(validation_df)
    fatal_percentage = (fatal_corrupt_rows / total_rows * 100.0) if total_rows else 0.0
    os.makedirs('quarantine', exist_ok=True)
    _q_stem = os.path.splitext(os.path.basename(globals().get('CURRENT_INPUT_FILE') or 'records'))[0]
    quarantine_path = os.path.join('quarantine', f'{_q_stem}_quarantine.csv')
    elapsed_seconds = max(0.0, time.perf_counter() - started)

    is_small_batch = total_rows <= 50
    is_breached = (fatal_percentage > 50.0 or (total_rows - fatal_corrupt_rows) < 2) if is_small_batch else (fatal_percentage > FATAL_CORRUPTION_THRESHOLD_PERCENT)

    if is_breached:
        quarantine_export_df.to_csv(quarantine_path, index=False)
        _write_validation_log(output_dir, ingested_rows, 0, soft_imputations, fatal_corrupt_rows, elapsed_seconds, 'CRITICAL HALT')
        raise SystemExit(1)

    clean_df = validation_df.loc[~fatal_mask].copy()
    status = 'SUCCESS' if not fatal_corrupt_rows else 'PARTIAL SUCCESS (QUARANTINED)'
    if fatal_corrupt_rows:
        quarantine_export_df.to_csv(quarantine_path, index=False)
    _write_validation_log(output_dir, ingested_rows, len(clean_df), soft_imputations, fatal_corrupt_rows, elapsed_seconds, status)
    return ValidationResult(clean_df, soft_imputations, fatal_corrupt_rows, status)

def clean_file_path(path_str):
    return str(path_str).strip().replace('"', '').replace("'", "").strip('& ') if path_str else ""

def format_compact_num(val, is_ratio=False, is_avg=False):
    if is_ratio: return f"{val * 100:.1f}%"
    if is_avg: return f"{val:.1f}"
    abs_v = abs(val)
    if abs_v >= 10000000: return f"{val/10000000:.2f}Cr"
    elif abs_v >= 100000: return f"{val/100000:.2f}L"
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
        if any(k in h.lower() for k in ['bill', 'amt', 'amount', 'revenue', 'price', 'inr', 'ctc', 'salary']):
            billing_idx = idx_h
            break

    n_tail = expected_len - 1 - billing_idx
    for row in reader[1:]:
        if not row or not any(str(x).strip() for x in row): continue
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
        if ext == '.csv': return heal_and_ingest_csv(clean_path)
        elif ext in ['.xlsx', '.xls']: return pd.read_excel(clean_path)
    except: return None
    return None

def clean_dataframe(df):
    cleaned = df.copy()
    cleaned.columns = [str(c).strip().replace('\n', ' ') for c in cleaned.columns]
    initial_count = len(cleaned)

    # Fast vectorized cleaning (No slow nested row iteration)
    for col in cleaned.columns:
        col_low = str(col).lower()
        is_code = any(k in col_low for k in ['id', 'code', 'pin', 'zip', 'key', 'sku', 'inv', 'sl_no', 'account', 'no.'])
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
                    
            s_num = s.astype(str).str.replace(r'[^\d.\-]', '', regex=True)
            converted = pd.to_numeric(s_num, errors='coerce')
            if converted.notna().sum() >= (0.5 * len(cleaned)):
                cleaned[col] = converted
                continue
                
            cleaned[col] = s.dropna().astype(str).str.replace('_', ' ').str.replace('-', ' ').str.title()
            
    cleaned.dropna(how='all', inplace=True)
    cleaned.drop_duplicates(inplace=True)
    return cleaned, initial_count - len(cleaned)

def profile_algebraic_types(df):
    schema = {'Additive_Measures': [], 'Intensive_Measures': [], 'Categorical_Dims': [], 'Temporal_Dims': [], 'Identifier_Keys': [], 'Metric_Aggregations': {}}
    n_rows = len(df)
    if n_rows == 0: return schema

    for col in df.columns:
        col_str = str(col).strip()
        col_low = col_str.lower()
        series = df[col].dropna()
        series = series[series != '']
        n_unique = series.nunique()
        if n_unique == 0: continue
        uniqueness_ratio = n_unique / n_rows

        if any(k in col_low for k in ['date', 'time', 'ts', 'timestamp', 'period', 'month', 'year']) or pd.api.types.is_datetime64_any_dtype(series):
            schema['Temporal_Dims'].append(col_str)
            continue

        if any(k in col_low for k in ['id', 'code', 'pin', 'zip', 'key', 'sku', 'phone', 'account']):
            if 2 <= n_unique <= 15 and uniqueness_ratio < 0.40:
                schema['Categorical_Dims'].append(col_str)
            else:
                schema['Identifier_Keys'].append(col_str)
            continue

        if pd.api.types.is_numeric_dtype(series):
            if (uniqueness_ratio > 0.85 and len(series) > 50) and (series.dtype in ['int64', 'int32', 'int16', 'int8']) and not any(k in col_low for k in ['amount', 'bill', 'sales', 'revenue', 'cost', 'spend', 'price', 'total', 'salary', 'ctc']):
                schema['Identifier_Keys'].append(col_str)
                continue

            if (series.dtype in ['int64', 'int32']) and (2 <= n_unique <= 6) and (len(series) > 50) and not any(k in col_low for k in ['amount', 'bill', 'sales', 'revenue', 'cost', 'price', 'total', 'salary', 'ctc']):
                schema['Categorical_Dims'].append(col_str)
                continue

            c_min, c_max = float(series.min()), float(series.max())
            is_extensive = any(k in col_low for k in ['sales', 'revenue', 'cost', 'spend', 'expense', 'profit', 'volume', 'qty', 'amount', 'total', 'ctc', 'salary', 'bonus'])
            is_ratio = (c_min >= -1.0) and (c_max <= 1.0) and (series.dtype in ['float64', 'float32'])
            is_rating = (c_min >= 0.0) and (c_max <= 10.0) and any(k in col_low for k in ['rating', 'score', 'stars'])

            if is_extensive:
                schema['Additive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'SUM'
            elif is_ratio or is_rating:
                schema['Intensive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'AVERAGE'
            else:
                schema['Additive_Measures'].append(col_str)
                schema['Metric_Aggregations'][col_str] = 'SUM'
            continue

        if 2 <= n_unique <= 30 and (uniqueness_ratio < 0.90 or n_rows <= 50):
            schema['Categorical_Dims'].append(col_str)
        else:
            schema['Identifier_Keys'].append(col_str)

    return schema

SECTOR_THEMES = {
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

    chart2_config = {'mode': 'COLUMN_VOLUME', 'dim': sec_dim, 'title': f'Volume Breakdown by {sec_dim}'}

    kpi_measures = []
    for m in filtered_monetary:
        if len(kpi_measures) < 2: kpi_measures.append((m, math_schema.get('Metric_Aggregations', {}).get(m, 'SUM')))
    for m in intensive_cands:
        if len(kpi_measures) < 3: kpi_measures.append((m, math_schema.get('Metric_Aggregations', {}).get(m, 'AVG')))
    if len(kpi_measures) < 3 and filtered_monetary:
        for m in filtered_monetary:
            if m not in [k[0] for k in kpi_measures] and len(kpi_measures) < 3: kpi_measures.append((m, math_schema.get('Metric_Aggregations', {}).get(m, 'SUM')))

    return {
        'sector': 'GENERAL_ENTERPRISE', 'title': 'AUTONOMOUS EXECUTIVE DASHBOARD', 'vol_label': 'TOTAL RECORDS / UNITS',
        'macro_dim': macro_dim, 'sec_dim': sec_dim, 'primary_measure': primary_measure, 'primary_agg': primary_agg,
        'chart1_type': 'col', 'chart2_config': chart2_config, 'kpi_measures': kpi_measures[:3],
        'palette': SECTOR_THEMES['GENERAL_ENTERPRISE'],
        'temporal_dim': math_schema['Temporal_Dims'][0] if math_schema['Temporal_Dims'] else None
    }

def execute_math_agg(df, dim, metric, agg_type):
    valid_df = df[df[dim].notna() & (df[dim] != '')].copy()
    if valid_df.empty: return valid_df, pd.Series(dtype=float)
    valid_df[metric] = pd.to_numeric(valid_df[metric], errors='coerce')
    agg_func = 'mean' if agg_type == 'AVG' else 'sum'
    return valid_df, valid_df.groupby(dim)[metric].agg(agg_func).sort_values(ascending=False)

def generate_nlg_executive_summary(df, profile):
    dim1, metric, agg_type = profile['macro_dim'], profile['primary_measure'], profile['primary_agg']
    if not dim1 or not metric or df.empty or dim1 not in df.columns:
        return ["• Pipeline processed securely."]
    try:
        valid_df, agg_d1 = execute_math_agg(df, dim1, metric, agg_type)
        if agg_d1.empty: return ["• Zero net variance across dimensions."]
        total_val = float(valid_df[metric].mean() if agg_type == 'AVG' else valid_df[metric].sum())
        top_leader = str(agg_d1.index[0])
        lag_leader = str(agg_d1.index[-1]) if len(agg_d1) > 1 else None
        
        lines = [
            f"• Core Performance: Aggregate {metric.replace('_', ' ')} reaches {format_compact_num(total_val)}, steered by '{top_leader}'.",
            f"• Portfolio Analysis: Lower cohort localized in '{lag_leader}' with targeted audit recommended." if lag_leader else "• Portfolio metrics balanced evenly across cohorts.",
            "• Action Directive: Scale core growth verticals and maintain audit compliance across regional tiers."
        ]
        return lines
    except:
        return ["• Executive Overview: Pipeline processed with aggregate metrics."]

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
    arrow = '▼' if growth_pct < 0 else '▲'
    ws_fc['E2'] = f"  PROJECTED TRAJECTORY: {arrow} {abs(growth_pct):.1f}% Dynamic Shift expected over next 3 cycles."
    ws_fc['E2'].font = Font(bold=True, size=10, color='065F46' if growth_pct >= 0 else '991B1B')
    ws_fc['E2'].fill = PatternFill(start_color='ECFDF5' if growth_pct >= 0 else 'FEF2F2', fill_type='solid')
    
    for c_letter in ['A','B','C','D','E','F','G']: ws_fc.column_dimensions[c_letter].width = 16.0
    
    ws_fc['A3'], ws_fc['B3'], ws_fc['C3'] = "Timeline Phase", "Status", "Value Metric"
    for cell in ['A3', 'B3', 'C3']:
        ws_fc[cell].font = Font(bold=True)
        ws_fc[cell].fill = PatternFill(start_color="F1F5F9", fill_type="solid")
    
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
    c_fc.legend = None
    c_fc.y_axis.scaling.min = 0
    c_fc.add_data(Reference(ws_fc, min_col=3, min_row=3, max_row=curr_r-1), titles_from_data=True)
    c_fc.set_categories(Reference(ws_fc, min_col=1, min_row=4, max_row=curr_r-1))
    c_fc.series[0].graphicalProperties.line.solidFill = "0284C7"
    ws_fc.add_chart(c_fc, "E4")

def build_universal_dashboard(df, profile, output_path, dropped_count=0):
    wb = openpyxl.Workbook()
    pal = profile['palette']
    
    # 1. Cleaned Data Sheet
    ws_data = wb.active
    ws_data.title = "Cleaned_Data"
    headers = list(df.columns)
    ws_data.append(headers)
    for row in df.itertuples(index=False, name=None): ws_data.append(list(row))
    num_rows = len(df) + 1

    for i, col in enumerate(headers):
        col_letter = get_column_letter(i+1)
        max_len = max(len(str(col)), 14)
        ws_data.column_dimensions[col_letter].width = min(max_len + 4, 30)

    for cx in ws_data[1]:
        cx.fill = PatternFill(start_color="1F4E78", fill_type="solid")
        cx.font = Font(color="FFFFFF", bold=True)
    ws_data.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{num_rows}"
    ws_data.freeze_panes = "A2"

    # 2. Calculations Sheet
    ws_calc = wb.create_sheet(title="Calculations")
    dim1_col, dim2_col, m1_col, m1_agg = profile['macro_dim'], profile['sec_dim'], profile['primary_measure'], profile['primary_agg']
    d1_let, d2_let, m1_let = get_column_letter(headers.index(dim1_col)+1), get_column_letter(headers.index(dim2_col)+1), get_column_letter(headers.index(m1_col)+1)
    
    unique_dim1 = [str(x) for x in df[dim1_col].dropna().unique()][:10]
    if dim1_col:
        ws_calc['A1'], ws_calc['B1'] = str(dim1_col), str(m1_col)
        agg_str = 'AVERAGE' if m1_agg == 'AVG' else 'SUM'
        for i, val in enumerate(unique_dim1, start=2):
            ws_calc[f'A{i}'] = str(val)
            ws_calc[f'B{i}'] = f'=IFERROR(IF(Executive_Dashboard!$M$1="All", {agg_str}IFS(Cleaned_Data!{m1_let}2:{m1_let}{num_rows}, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Calculations!A{i}), {agg_str}IFS(Cleaned_Data!{m1_let}2:{m1_let}{num_rows}, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Calculations!A{i}, Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Executive_Dashboard!$M$1)), 0)'
            
    # Chart 2 Categories (Limit to top 6 to prevent overcrowding)
    sec_counts = df[dim2_col].value_counts()
    unique_dim2 = [str(x) for x in sec_counts.index[:6]]
    if unique_dim2:
        ws_calc['D1'], ws_calc['E1'] = str(dim2_col), "Volume"
        for i, val in enumerate(unique_dim2, start=2):
            ws_calc[f'D{i}'] = str(val)
            ws_calc[f'E{i}'] = f'=IFERROR(IF(Executive_Dashboard!$J$1="All", COUNTIF(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Calculations!D{i}), COUNTIFS(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Calculations!D{i}, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Executive_Dashboard!$J$1)), 0)'

    # 3. Executive Dashboard Sheet
    ws_dash = wb.create_sheet(title="Executive_Dashboard", index=0)
    ws_dash.sheet_view.showGridLines = False
    for c_letter in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N']: ws_dash.column_dimensions[c_letter].width = 13.0
    for r in range(1, 45):
        for c in range(1, 16): ws_dash.cell(row=r, column=c).fill = PatternFill(start_color="FFFFFF", fill_type="solid")
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
    
    for sc_start, sc_end, s_name, s_col, dv_list in [('H', 'I', dim1_col, 'J', unique_dim1), ('K', 'L', dim2_col, 'M', unique_dim2)]:
        ws_dash.merge_cells(f'{sc_start}1:{sc_end}1')
        ws_dash[f'{sc_start}1'] = f"Filter {s_name}:"
        ws_dash[f'{sc_start}1'].font = Font(name="Calibri", size=8.5, bold=True, color=pal['title_color'])
        ws_dash[f'{sc_start}1'].fill = f_card
        ws_dash[f'{s_col}1'] = "All"
        ws_dash[f'{s_col}1'].fill = PatternFill(start_color=pal['filter_bg'], fill_type="solid")
        ws_dash[f'{s_col}1'].border = t_border
        if dv_list:
            dv = DataValidation(type="list", formula1=f'"{",".join(["All"] + dv_list[:12])}"', allow_blank=True)
            ws_dash.add_data_validation(dv)
            dv.add(f'{s_col}1')
            
    ws_dash.merge_cells('A2:F2')
    ws_dash['A2'] = f"  DATA GOVERNANCE: v71.0 Final Enterprise Master | Active Run"
    ws_dash['A2'].font = Font(size=7.5, bold=True, color="475569")
    ws_dash['A2'].fill = PatternFill(start_color="F1F5F9", fill_type="solid")

    cards_data = [(profile['vol_label'], f'=IFERROR(IF(AND($J$1="All", $M$1="All"), COUNTA(Cleaned_Data!A2:A{num_rows}), IF($J$1="All", COUNTIF(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, $M$1), IF($M$1="All", COUNTIF(Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, $J$1), COUNTIFS(Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, $J$1, Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, $M$1)))), 0)', '#,##0')]
    for metric, agg_type in profile['kpi_measures']:
        c_let = get_column_letter(headers.index(metric) + 1)
        lbl = f"{agg_type} {str(metric).upper().replace('_', ' ')}"
        func = 'AVERAGE' if agg_type in ['AVG', 'AVERAGE', 'MEDIAN'] else 'SUM'
        fmt = get_math_format(df, metric, agg_type)
        form = f'=IFERROR(IF(AND(Executive_Dashboard!$J$1="All", Executive_Dashboard!$M$1="All"), {func}(Cleaned_Data!{c_let}2:{c_let}{num_rows}), IF(Executive_Dashboard!$J$1="All", {func}IF(Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Executive_Dashboard!$M$1, Cleaned_Data!{c_let}2:{c_let}{num_rows}), IF(Executive_Dashboard!$M$1="All", {func}IF(Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Executive_Dashboard!$J$1, Cleaned_Data!{c_let}2:{c_let}{num_rows}), {func}IFS(Cleaned_Data!{c_let}2:{c_let}{num_rows}, Cleaned_Data!{d1_let}2:{d1_let}{num_rows}, Executive_Dashboard!$J$1, Cleaned_Data!{d2_let}2:{d2_let}{num_rows}, Executive_Dashboard!$M$1)))), 0)'
        cards_data.append((lbl, form, fmt))

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
        ws_dash[f'{cs}5'] = "▶ Verified Nominal"
        ws_dash[f'{cs}5'].font = Font(size=7.5, bold=True, color="334155")
        ws_dash[f'{cs}5'].fill = f_card
        ws_dash[f'{cs}5'].alignment = Alignment(horizontal="center")

    # Chart 1: Original Clean Column Chart (NO Data Labels Overlapping)
    if unique_dim1:
        c1 = BarChart()
        c1.type = 'col'
        c1.style = 10
        c1.height = 6.8
        c1.width = 13.8
        c1.legend = None
        c1.dataLabels = None
        c1.title = f"Performance Ranking: {m1_col.replace('_', ' ')} by {dim1_col.replace('_', ' ')}"
        
        c1.y_axis.delete = False
        c1.x_axis.delete = False
        c1.y_axis.majorGridlines = None
        c1.x_axis.majorGridlines = None
        c1.y_axis.number_format = '#,##0'

        c1.add_data(Reference(ws_calc, min_col=2, min_row=1, max_row=len(unique_dim1)+1), titles_from_data=True)
        c1.set_categories(Reference(ws_calc, min_col=1, min_row=2, max_row=len(unique_dim1)+1))
        ws_dash.add_chart(c1, 'A6')

    # Chart 2: Original Column Chart (Restored from Donut, Zero Clutter, Clean Right Legend)
    if unique_dim2:
        c2 = BarChart()
        c2.type = 'col'
        c2.style = 10
        c2.height = 6.8
        c2.width = 13.8
        c2.dataLabels = None
        c2.title = f"Volume Breakdown by {dim2_col}"
        
        c2.legend = Legend()
        c2.legend.legendPos = "r"
        
        c2.y_axis.delete = False
        c2.x_axis.delete = False
        c2.y_axis.majorGridlines = None
        c2.x_axis.majorGridlines = None

        c2.add_data(Reference(ws_calc, min_col=5, min_row=1, max_row=len(unique_dim2)+1), titles_from_data=True)
        c2.set_categories(Reference(ws_calc, min_col=4, min_row=2, max_row=len(unique_dim2)+1))
        ws_dash.add_chart(c2, "H6")

    # Benchmark Matrix
    z3_start = 21
    ws_dash.merge_cells(f'A{z3_start}:F{z3_start}')
    ws_dash[f'A{z3_start}'] = "EXECUTIVE BENCHMARK MATRIX"
    ws_dash[f'A{z3_start}'].font = Font(size=8.5, bold=True, color="FFFFFF")
    ws_dash[f'A{z3_start}'].fill = f_sub
    
    fmt_m1 = get_math_format(df, m1_col, m1_agg)
    calc_end_r = 1 + len(unique_dim1)
    matrix_specs = [
        ('Top Leader', 'LARGE', 1, 'DCFCE7', '★'),
        ('Secondary', 'LARGE', 2, 'DCFCE7', '★'),
        ('Review Needed', 'SMALL', 1, 'FFE4E6', '▼')
    ]
    r_idx = z3_start + 1
    for lbl, func, k_rank, bg_fill, sym in matrix_specs[:min(len(unique_dim1), 3)]:
        ws_dash.merge_cells(f'A{r_idx}:D{r_idx}')
        ws_dash.merge_cells(f'E{r_idx}:F{r_idx}')
        ws_dash[f'A{r_idx}'] = f'="{sym} {lbl}: " & IFERROR(INDEX(Calculations!$A$2:$A${calc_end_r}, MATCH({func}(Calculations!$B$2:$B${calc_end_r}, {k_rank}), Calculations!$B$2:$B${calc_end_r}, 0)), "N/A")'
        ws_dash[f'E{r_idx}'] = f'=IFERROR({func}(Calculations!$B$2:$B${calc_end_r}, {k_rank}), 0)'
        ws_dash[f'A{r_idx}'].fill = PatternFill(start_color=bg_fill, fill_type="solid")
        ws_dash[f'E{r_idx}'].number_format = fmt_m1
        r_idx += 1

    # Narrative Audit Section
    ws_dash.merge_cells('A26:N26')
    ws_dash['A26'] = "  GLOBAL BASELINE EXECUTIVE AUDIT (PORTFOLIO BENCHMARK)"
    ws_dash['A26'].font = Font(size=8.5, bold=True, color="FFFFFF")
    ws_dash['A26'].fill = f_head
    
    for s_idx, line in enumerate(generate_nlg_executive_summary(df, profile), start=27):
        ws_dash.merge_cells(f'A{s_idx}:N{s_idx}')
        ws_dash[f'A{s_idx}'] = f"  {line}"
        ws_dash[f'A{s_idx}'].font = Font(size=8.5)
        ws_dash[f'A{s_idx}'].fill = PatternFill(start_color="F8FAFC", fill_type="solid")

    generate_predictive_forecast_sheet(wb, df, profile)
    wb.save(output_path)
    print(f"\n[SUCCESS] Universal Gatekeeper Dashboard generated: {output_path}")

def process_pipeline(raw_input_path):
    input_path = clean_file_path(raw_input_path)
    started_at = time.perf_counter()
    raw_df = ingest_file(input_path)
    if raw_df is None or raw_df.empty: return
    clean_df, dropped = clean_dataframe(raw_df)
    global CURRENT_INPUT_FILE; CURRENT_INPUT_FILE = input_path
    validation = validate_with_circuit_breaker(
        clean_df,
        output_dir='reports',
        total_rows_ingested=len(raw_df),
        started_at=started_at,
    )
    clean_df = validation.dataframe
    profile = build_mathematical_profile(clean_df)
    
    os.makedirs('reports', exist_ok=True)
    base_stem = os.path.splitext(os.path.basename(input_path))[0]
    output_name = os.path.join('reports', f'{base_stem}_Gatekeeper_Dashboard.xlsx')
    build_universal_dashboard(clean_df, profile, output_name, dropped)
    _write_validation_log('reports', base_stem, len(raw_df), len(clean_df), validation.soft_imputations, validation.fatal_corrupt_rows, max(0.0, time.perf_counter() - started_at), validation.status)

    try:
        os.makedirs('clean_data', exist_ok=True)
        clean_df.to_parquet(os.path.join('clean_data', f'{base_stem}_Cleaned.parquet'), index=False)
    except Exception:
        pass

if __name__ == "__main__":
    if len(sys.argv) > 1:
        process_pipeline(sys.argv[1])
