"""
DataCleanse Pro - Intelligent Data Cleaning Tool
Author of this Project | Kajola Gbenga Adewale
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import re
import warnings
from datetime import datetime
from collections import defaultdict
import openpyxl

warnings.filterwarnings("ignore")

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DataCleanse Pro",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

* { font-family: 'Inter', sans-serif; }

/* ── Root palette ── */
:root {
    --bg-deep:   #0B0F1A;
    --bg-card:   #111827;
    --bg-panel:  #1C2333;
    --bg-input:  #1E2A3A;
    --border:    #2D3A4F;
    --accent:    #3B82F6;
    --accent2:   #10B981;
    --warn:      #F59E0B;
    --danger:    #EF4444;
    --text-1:    #F1F5F9;
    --text-2:    #94A3B8;
    --text-3:    #64748B;
    --mono:      'JetBrains Mono', monospace;
}

/* ── Global background ── */
.stApp { background: var(--bg-deep); }
section[data-testid="stSidebar"] { background: var(--bg-card); }

/* ── Hide streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

/* ── Main header ── */
.dc-header {
    background: linear-gradient(135deg, #0F2027, #203A43, #2C5364);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 36px 40px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.dc-header::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(59,130,246,0.15) 0%, transparent 70%);
}
.dc-header h1 {
    font-size: 2.4rem; font-weight: 700;
    background: linear-gradient(90deg, #60A5FA, #34D399);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    margin: 0 0 8px 0;
}
.dc-header p { color: var(--text-2); margin: 0; font-size: 1.05rem; }

/* ── Step indicator ── */
.step-bar {
    display: flex; gap: 0;
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 6px;
    margin-bottom: 28px;
    overflow: hidden;
}
.step-item {
    flex: 1; text-align: center;
    padding: 10px 8px;
    border-radius: 8px;
    font-size: 0.8rem; font-weight: 500;
    color: var(--text-3);
    transition: all 0.3s;
}
.step-item.active {
    background: linear-gradient(135deg, #1D4ED8, #0F766E);
    color: white;
}
.step-item.done {
    background: rgba(16, 185, 129, 0.15);
    color: var(--accent2);
}

/* ── Cards ── */
.dc-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 24px;
    margin-bottom: 18px;
}
.dc-card-header {
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 18px;
    padding-bottom: 14px;
    border-bottom: 1px solid var(--border);
}
.dc-card-header h3 {
    margin: 0; font-size: 1.05rem; font-weight: 600;
    color: var(--text-1);
}

/* ── Issue severity badges ── */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.72rem; font-weight: 600;
    letter-spacing: 0.05em; text-transform: uppercase;
}
.badge-critical { background: rgba(239,68,68,0.15); color: #F87171; border: 1px solid rgba(239,68,68,0.3); }
.badge-high     { background: rgba(245,158,11,0.15); color: #FCD34D; border: 1px solid rgba(245,158,11,0.3); }
.badge-medium   { background: rgba(59,130,246,0.15);  color: #93C5FD; border: 1px solid rgba(59,130,246,0.3); }
.badge-low      { background: rgba(16,185,129,0.15);  color: #6EE7B7; border: 1px solid rgba(16,185,129,0.3); }
.badge-info     { background: rgba(139,92,246,0.15);  color: #C4B5FD; border: 1px solid rgba(139,92,246,0.3); }

/* ── Issue item ── */
.issue-item {
    background: var(--bg-panel);
    border: 1px solid var(--border);
    border-left: 4px solid;
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 10px;
}
.issue-item.critical { border-left-color: #EF4444; }
.issue-item.high     { border-left-color: #F59E0B; }
.issue-item.medium   { border-left-color: #3B82F6; }
.issue-item.low      { border-left-color: #10B981; }

.issue-title { font-weight: 600; color: var(--text-1); font-size: 0.9rem; margin-bottom: 4px; }
.issue-desc  { color: var(--text-2); font-size: 0.82rem; line-height: 1.5; }
.issue-meta  { color: var(--text-3); font-size: 0.75rem; margin-top: 6px; font-family: var(--mono); }

/* ── Quality score ring ── */
.score-display {
    text-align: center; padding: 32px 24px;
    background: var(--bg-panel);
    border-radius: 14px;
    border: 1px solid var(--border);
}
.score-number {
    font-size: 5rem; font-weight: 800;
    line-height: 1;
    background: linear-gradient(135deg, #60A5FA, #34D399);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.score-grade {
    font-size: 1.6rem; font-weight: 700;
    margin-top: 4px; color: var(--text-1);
}
.score-label { color: var(--text-2); font-size: 0.9rem; margin-top: 8px; }

/* ── Metric tiles ── */
.metric-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
.metric-tile {
    background: var(--bg-panel);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 16px;
    text-align: center;
}
.metric-val { font-size: 1.8rem; font-weight: 700; color: var(--text-1); }
.metric-key { font-size: 0.75rem; color: var(--text-2); margin-top: 4px; text-transform: uppercase; letter-spacing: 0.06em; }

/* ── Progress bar ── */
.prog-bar {
    height: 8px; background: var(--bg-panel);
    border-radius: 4px; overflow: hidden;
    margin-top: 6px;
}
.prog-fill {
    height: 100%; border-radius: 4px;
    transition: width 0.6s ease;
}
.prog-green  { background: linear-gradient(90deg, #10B981, #34D399); }
.prog-blue   { background: linear-gradient(90deg, #3B82F6, #60A5FA); }
.prog-amber  { background: linear-gradient(90deg, #F59E0B, #FCD34D); }
.prog-red    { background: linear-gradient(90deg, #EF4444, #F87171); }

/* ── Method selection cards ── */
.method-card {
    background: var(--bg-panel);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 8px;
    cursor: pointer;
    transition: border-color 0.2s;
}
.method-card:hover { border-color: var(--accent); }
.method-card.selected { border-color: var(--accent); background: rgba(59,130,246,0.08); }
.method-name { font-weight: 600; color: var(--text-1); font-size: 0.88rem; }
.method-desc { color: var(--text-2); font-size: 0.78rem; margin-top: 3px; }

/* ── Diff highlight ── */
.changed-cell { background: rgba(16,185,129,0.12); color: #34D399; }

/* ── Upload zone ── */
.upload-zone {
    background: var(--bg-panel);
    border: 2px dashed var(--border);
    border-radius: 16px;
    padding: 48px 32px;
    text-align: center;
    transition: border-color 0.3s;
}
.upload-zone:hover { border-color: var(--accent); }
.upload-icon { font-size: 3rem; margin-bottom: 16px; }
.upload-text { color: var(--text-2); font-size: 0.95rem; }

/* ── Streamlit overrides ── */
.stButton > button {
    background: linear-gradient(135deg, #1D4ED8, #0F766E) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 10px 24px !important;
    font-size: 0.9rem !important;
    transition: opacity 0.2s !important;
}
.stButton > button:hover { opacity: 0.88 !important; }
.stButton > button:disabled { opacity: 0.4 !important; }

div[data-testid="stCheckbox"] { color: var(--text-1); }
div[data-testid="stSelectbox"] label { color: var(--text-2); font-size: 0.85rem; }
.stDataFrame { border-radius: 10px; overflow: hidden; }

.stTabs [data-baseweb="tab-list"] {
    background: var(--bg-card) !important;
    border-radius: 10px !important;
    padding: 4px !important;
    gap: 4px !important;
    border: 1px solid var(--border) !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--text-2) !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #1D4ED8, #0F766E) !important;
    color: white !important;
}

div[data-testid="stAlert"] {
    border-radius: 10px !important;
    border: 1px solid var(--border) !important;
}

.stExpander { border: 1px solid var(--border) !important; border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CORE ENGINE FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def load_file(uploaded_file):
    """Load CSV or Excel file, return DataFrame + metadata."""
    name = uploaded_file.name
    ext  = name.rsplit(".", 1)[-1].lower()
    try:
        if ext == "csv":
            # try utf-8 first, fallback to latin-1
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8")
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin-1")
        elif ext in ("xlsx", "xls"):
            df = pd.read_excel(uploaded_file, engine="openpyxl" if ext == "xlsx" else "xlrd")
        else:
            return None, f"Unsupported file type: .{ext}"
        return df, None
    except Exception as e:
        return None, str(e)


def _is_string_dtype(dtype):
    """Pandas 3.x compatible string dtype check (object OR StringDtype)."""
    if dtype == object:
        return True
    dtype_str = str(dtype)
    if dtype_str in ("str", "string", "String"):
        return True
    if hasattr(dtype, "name") and dtype.name in ("str", "string", "String"):
        return True
    # pandas StringDtype check
    try:
        import pandas as pd
        if isinstance(dtype, pd.StringDtype):
            return True
    except Exception:
        pass
    return False


def infer_column_types(df):
    """Return dict mapping column -> inferred semantic type."""
    type_map = {}
    for col in df.columns:
        s = df[col].dropna()
        if len(s) == 0:
            type_map[col] = "empty"
            continue
        dtype = df[col].dtype
        # Numeric
        if pd.api.types.is_numeric_dtype(dtype):
            if set(s.unique()).issubset({0, 1, True, False}):
                type_map[col] = "binary"
            elif s.nunique() <= 20 and s.nunique() / len(s) < 0.05:
                type_map[col] = "categorical_numeric"
            else:
                type_map[col] = "numeric"
            continue
        # Try parse as datetime
        if _is_string_dtype(dtype):
            sample = s.head(50).astype(str)
            try:
                parsed = pd.to_datetime(sample, infer_datetime_format=True, errors="coerce")
                if parsed.notna().mean() > 0.7:
                    type_map[col] = "datetime"
                    continue
            except Exception:
                pass
            # Check cardinality for categorical
            uniq_rate = s.nunique() / len(s)
            if s.nunique() <= 50 or uniq_rate < 0.05:
                type_map[col] = "categorical"
            else:
                type_map[col] = "text"
    return type_map


def detect_issues(df):
    """
    Comprehensive issue detection. Returns list of issue dicts.
    Each dict has: id, severity, category, title, description,
                   affected_columns, affected_rows, recommendation, fix_key
    """
    issues = []
    n_rows, n_cols = df.shape
    col_types = infer_column_types(df)

    # ── 1. Missing Values ─────────────────────────────────────────────────────
    missing = df.isnull().sum()
    miss_pct = (missing / n_rows * 100).round(2)
    cols_with_missing = missing[missing > 0]

    if len(cols_with_missing) > 0:
        # per-column issues for columns with >0% missing
        for col, cnt in cols_with_missing.items():
            pct = miss_pct[col]
            if pct >= 50:
                sev = "critical"
                rec = f"Column '{col}' is >50% empty. Consider dropping or imputing with domain knowledge."
            elif pct >= 20:
                sev = "high"
                rec = f"Fill with median/mode for '{col}' or apply model-based imputation."
            elif pct >= 5:
                sev = "medium"
                rec = f"Fill '{col}' with mean/median/mode or flag with indicator column."
            else:
                sev = "low"
                rec = f"'{col}' has minor missingness. Safe to fill with mean/median/mode."

            ctype = col_types.get(col, "unknown")
            if ctype in ("numeric", "categorical_numeric", "binary"):
                method_hint = "mean, median, or zero-fill"
            elif ctype == "categorical":
                method_hint = "mode (most frequent value) or 'Unknown' label"
            elif ctype == "datetime":
                method_hint = "forward/backward fill or drop"
            else:
                method_hint = "mode or 'Unknown' label"

            issues.append({
                "id": f"MISS_{col}",
                "severity": sev,
                "category": "Missing Values",
                "title": f"Missing values in '{col}'",
                "description": (
                    f"{cnt:,} out of {n_rows:,} rows ({pct}%) are missing in this column. "
                    f"Detected type: {ctype}. Recommended fill method: {method_hint}."
                ),
                "affected_columns": [col],
                "affected_rows": int(cnt),
                "recommendation": rec,
                "fix_key": "missing_values",
                "meta": {"pct": pct, "col_type": ctype},
            })

    # ── 2. Duplicate Rows ─────────────────────────────────────────────────────
    dup_mask = df.duplicated()
    n_dups   = dup_mask.sum()
    if n_dups > 0:
        pct = round(n_dups / n_rows * 100, 2)
        issues.append({
            "id": "DUPS_FULL",
            "severity": "high" if pct > 5 else "medium",
            "category": "Duplicates",
            "title": f"Duplicate rows detected ({n_dups:,} rows = {pct}%)",
            "description": (
                f"{n_dups:,} rows are exact duplicates of other rows. "
                f"This inflates counts, distorts averages, and corrupts groupby aggregations. "
                f"Removing duplicates preserves only the first occurrence of each unique row."
            ),
            "affected_columns": list(df.columns),
            "affected_rows": int(n_dups),
            "recommendation": "Drop duplicate rows, keeping the first occurrence.",
            "fix_key": "duplicates",
            "meta": {"n_dups": int(n_dups), "pct": pct},
        })

    # Duplicate in single id-like column
    for col in df.columns:
        if "id" in col.lower() or col.lower() in ("key", "code", "ref", "reference"):
            n_dup_col = df[col].dropna().duplicated().sum()
            if n_dup_col > 0:
                issues.append({
                    "id": f"DUPS_COL_{col}",
                    "severity": "high",
                    "category": "Duplicates",
                    "title": f"Duplicate values in ID column '{col}'",
                    "description": (
                        f"'{col}' appears to be an identifier column but has {n_dup_col:,} duplicate values. "
                        f"ID columns must be unique. Duplicates suggest data entry errors or merge issues."
                    ),
                    "affected_columns": [col],
                    "affected_rows": int(n_dup_col),
                    "recommendation": "Investigate source of duplicate IDs. Remove or flag them.",
                    "fix_key": "duplicates",
                    "meta": {},
                })

    # ── 3. Data Type Issues ───────────────────────────────────────────────────
    for col in df.columns:
        s = df[col].dropna()
        if len(s) == 0:
            continue
        dtype = df[col].dtype

        # Object column that should be numeric
        if _is_string_dtype(dtype):
            # Try numeric conversion
            numeric_attempt = pd.to_numeric(s.astype(str).str.replace(",", "").str.strip(), errors="coerce")
            pct_convertible = numeric_attempt.notna().mean()
            if pct_convertible > 0.7 and col_types.get(col) not in ("datetime", "categorical"):
                n_non_numeric = int((1 - pct_convertible) * len(s))
                issues.append({
                    "id": f"TYPE_NUM_{col}",
                    "severity": "high",
                    "category": "Data Types",
                    "title": f"'{col}' stored as text but contains numeric data",
                    "description": (
                        f"{pct_convertible*100:.0f}% of values in '{col}' are numeric but the column is stored as text. "
                        f"{n_non_numeric} non-numeric entries (possibly commas, currency symbols, or typos) prevent automatic parsing."
                    ),
                    "affected_columns": [col],
                    "affected_rows": n_non_numeric,
                    "recommendation": f"Strip special characters and convert '{col}' to numeric type.",
                    "fix_key": "type_conversion",
                    "meta": {"target_type": "numeric", "pct_ok": pct_convertible},
                })

            # Try datetime
            elif col_types.get(col) == "datetime":
                issues.append({
                    "id": f"TYPE_DATE_{col}",
                    "severity": "medium",
                    "category": "Data Types",
                    "title": f"'{col}' should be datetime but is stored as text",
                    "description": (
                        f"'{col}' contains date-like values but is stored as a text string. "
                        f"Date arithmetic, sorting, and time-series analysis will fail until this is converted."
                    ),
                    "affected_columns": [col],
                    "affected_rows": 0,
                    "recommendation": f"Convert '{col}' to datetime64 using pd.to_datetime().",
                    "fix_key": "type_conversion",
                    "meta": {"target_type": "datetime"},
                })

    # ── 4. Outliers ───────────────────────────────────────────────────────────
    for col in df.columns:
        if col_types.get(col) not in ("numeric", "categorical_numeric"):
            continue
        s = df[col].dropna()
        if len(s) < 10:
            continue
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            continue
        lower = q1 - 3 * iqr
        upper = q3 + 3 * iqr
        outlier_mask = (s < lower) | (s > upper)
        n_out = outlier_mask.sum()
        if n_out > 0:
            pct = round(n_out / len(s) * 100, 2)
            issues.append({
                "id": f"OUT_{col}",
                "severity": "medium" if pct < 5 else "high",
                "category": "Outliers",
                "title": f"Outliers detected in '{col}'",
                "description": (
                    f"{n_out:,} values ({pct}%) in '{col}' are beyond 3x IQR from Q1/Q3. "
                    f"Range of outliers: [{s[outlier_mask].min():,.2f}, {s[outlier_mask].max():,.2f}]. "
                    f"Normal range: [{lower:,.2f}, {upper:,.2f}]."
                ),
                "affected_columns": [col],
                "affected_rows": int(n_out),
                "recommendation": "Cap outliers to 1.5xIQR fence, replace with median, or investigate individually.",
                "fix_key": "outliers",
                "meta": {"q1": q1, "q3": q3, "iqr": iqr, "pct": pct},
            })

    # ── 5. Whitespace and String Issues ───────────────────────────────────────
    str_issues_found = []
    for col in df.columns:
        if not _is_string_dtype(df[col].dtype):
            continue
        s = df[col].dropna().astype(str)
        n_leading  = (s != s.str.strip()).sum()
        n_extra_sp = s.str.contains(r"  +", regex=True).sum()
        n_mixed    = 0
        unique_vals = s.unique()
        if len(unique_vals) <= 200:
            lower_set  = set(v.lower() for v in unique_vals)
            title_set  = set(v.title() for v in unique_vals)
            upper_set  = set(v.upper() for v in unique_vals)
            if len(lower_set) < len(set(unique_vals)):
                n_mixed = len(set(unique_vals)) - len(lower_set)

        total_str_issues = n_leading + n_extra_sp + n_mixed
        if total_str_issues > 0:
            str_issues_found.append(col)
            desc_parts = []
            if n_leading > 0:   desc_parts.append(f"{n_leading} rows have leading/trailing whitespace")
            if n_extra_sp > 0:  desc_parts.append(f"{n_extra_sp} rows have multiple consecutive spaces")
            if n_mixed > 0:     desc_parts.append(f"{n_mixed} apparent case variants (e.g. 'Lagos' vs 'LAGOS')")
            issues.append({
                "id": f"STR_{col}",
                "severity": "medium" if total_str_issues > 50 else "low",
                "category": "String Quality",
                "title": f"String inconsistencies in '{col}'",
                "description": "; ".join(desc_parts) + f". Total affected entries: {total_str_issues:,}.",
                "affected_columns": [col],
                "affected_rows": int(total_str_issues),
                "recommendation": "Strip whitespace and standardise case to title case or upper case.",
                "fix_key": "string_cleaning",
                "meta": {"n_leading": int(n_leading), "n_mixed": int(n_mixed)},
            })

    # ── 6. Constant / Near-Constant Columns ──────────────────────────────────
    for col in df.columns:
        n_unique = df[col].nunique(dropna=True)
        if n_unique <= 1:
            issues.append({
                "id": f"CONST_{col}",
                "severity": "medium",
                "category": "Redundant Data",
                "title": f"'{col}' is constant (only 1 unique value)",
                "description": (
                    f"'{col}' has only one unique value across all {n_rows:,} rows. "
                    f"Constant columns carry zero information and add noise to analysis and machine learning models."
                ),
                "affected_columns": [col],
                "affected_rows": n_rows,
                "recommendation": "Drop this column - it provides no analytical value.",
                "fix_key": "drop_constant",
                "meta": {},
            })

    # ── 7. High Cardinality Categorical ───────────────────────────────────────
    for col in df.columns:
        if col_types.get(col) != "categorical":
            continue
        n_unique = df[col].nunique()
        uniq_rate = n_unique / n_rows
        if n_unique > 100 and uniq_rate > 0.5:
            issues.append({
                "id": f"HCARD_{col}",
                "severity": "low",
                "category": "Cardinality",
                "title": f"Very high cardinality in '{col}' ({n_unique:,} unique values)",
                "description": (
                    f"'{col}' has {n_unique:,} unique values ({uniq_rate*100:.1f}% of rows). "
                    f"This may be a free-text field treated as categorical, making groupby and visualisation impractical."
                ),
                "affected_columns": [col],
                "affected_rows": 0,
                "recommendation": "Consider treating as free-text, extracting keywords, or grouping rare values as 'Other'.",
                "fix_key": "cardinality",
                "meta": {"n_unique": n_unique},
            })

    # ── 8. Negative Values in Likely-Positive Columns ────────────────────────
    for col in df.columns:
        if col_types.get(col) not in ("numeric",):
            continue
        s = df[col].dropna()
        # Heuristics: column likely should be positive
        pos_hints = any(kw in col.lower() for kw in
                        ("price","cost","amount","qty","quantity","revenue","profit",
                         "age","count","total","sum","fee","salary","weight","height",
                         "sales","volume","rate","score"))
        if pos_hints:
            n_neg = (s < 0).sum()
            if n_neg > 0:
                issues.append({
                    "id": f"NEG_{col}",
                    "severity": "high",
                    "category": "Invalid Values",
                    "title": f"Negative values in '{col}' (expected non-negative)",
                    "description": (
                        f"{n_neg:,} rows have negative values in '{col}' which should logically be >=0. "
                        f"Examples: {sorted(s[s<0].unique()[:3].tolist())}. These may be data entry errors."
                    ),
                    "affected_columns": [col],
                    "affected_rows": int(n_neg),
                    "recommendation": f"Replace negative values in '{col}' with absolute value, zero, or NaN and re-impute.",
                    "fix_key": "negative_values",
                    "meta": {"n_neg": int(n_neg)},
                })

    # ── 9. All-Null Columns ───────────────────────────────────────────────────
    for col in df.columns:
        if df[col].isnull().all():
            issues.append({
                "id": f"NULL_{col}",
                "severity": "critical",
                "category": "Empty Data",
                "title": f"'{col}' is entirely empty (100% null)",
                "description": (
                    f"'{col}' contains no data whatsoever - all {n_rows:,} values are null. "
                    f"This column cannot be used for any analysis."
                ),
                "affected_columns": [col],
                "affected_rows": n_rows,
                "recommendation": "Drop this column entirely.",
                "fix_key": "drop_empty_cols",
                "meta": {},
            })

    # ── 10. Date Range Anomalies ──────────────────────────────────────────────
    for col in df.columns:
        if col_types.get(col) != "datetime":
            continue
        try:
            dates = pd.to_datetime(df[col], errors="coerce", infer_datetime_format=True)
            valid_dates = dates.dropna()
            if len(valid_dates) == 0:
                continue
            min_d, max_d = valid_dates.min(), valid_dates.max()
            # Flag future dates beyond today
            today = pd.Timestamp.now().normalize()
            n_future = (valid_dates > today).sum()
            n_very_old = (valid_dates.dt.year < 1900).sum()
            if n_future > 0 or n_very_old > 0:
                parts = []
                if n_future > 0:   parts.append(f"{n_future} dates are in the future")
                if n_very_old > 0: parts.append(f"{n_very_old} dates are before year 1900")
                issues.append({
                    "id": f"DATE_{col}",
                    "severity": "medium",
                    "category": "Date Anomalies",
                    "title": f"Suspicious dates in '{col}'",
                    "description": (
                        f"Column '{col}' has date range {min_d.date()} to {max_d.date()}. "
                        + "; ".join(parts) + "."
                    ),
                    "affected_columns": [col],
                    "affected_rows": int(n_future + n_very_old),
                    "recommendation": "Investigate and correct anomalous dates. Consider capping to a valid range.",
                    "fix_key": "date_anomalies",
                    "meta": {},
                })
        except Exception:
            pass

    # ── 11. Mixed Data Types in Column ───────────────────────────────────────
    for col in df.columns:
        if not _is_string_dtype(df[col].dtype):
            continue
        s = df[col].dropna()
        if len(s) == 0:
            continue
        type_counts = s.apply(type).value_counts()
        if len(type_counts) > 1:
            issues.append({
                "id": f"MIX_{col}",
                "severity": "medium",
                "category": "Mixed Types",
                "title": f"Mixed data types in '{col}'",
                "description": (
                    f"'{col}' contains values of multiple Python types: "
                    + ", ".join(f"{t.__name__} ({c})" for t, c in type_counts.items()) + ". "
                    f"Mixed types cause unpredictable behaviour in aggregation and ML pipelines."
                ),
                "affected_columns": [col],
                "affected_rows": 0,
                "recommendation": "Standardise to a single consistent type.",
                "fix_key": "string_cleaning",
                "meta": {},
            })

    # Sort by severity priority
    sev_order = {"critical": 0, "high": 1, "medium": 2, "low": 3}
    issues.sort(key=lambda x: sev_order.get(x["severity"], 4))

    return issues, col_types


def compute_quality_score(df, issues):
    """
    Compute a 0-100 data quality score with breakdown.
    Returns: overall_score, breakdown_dict
    """
    n_rows, n_cols = df.shape
    total_cells = n_rows * n_cols

    # Completeness (30 pts)
    missing_cells = df.isnull().sum().sum()
    completeness = max(0, 1 - missing_cells / total_cells) if total_cells > 0 else 1
    completeness_score = round(completeness * 30, 1)

    # Uniqueness (20 pts)
    n_dups = df.duplicated().sum()
    uniqueness = max(0, 1 - n_dups / n_rows) if n_rows > 0 else 1
    uniqueness_score = round(uniqueness * 20, 1)

    # Validity (25 pts) - based on critical/high issues
    critical_issues = [i for i in issues if i["severity"] == "critical"]
    high_issues     = [i for i in issues if i["severity"] == "high"]
    validity_penalty = min(25, len(critical_issues) * 8 + len(high_issues) * 4)
    validity_score   = max(0, 25 - validity_penalty)

    # Consistency (15 pts) - string quality, mixed types
    str_issues = [i for i in issues if i["category"] in ("String Quality", "Mixed Types")]
    type_issues = [i for i in issues if i["category"] == "Data Types"]
    consistency_penalty = min(15, len(str_issues) * 3 + len(type_issues) * 4)
    consistency_score   = max(0, 15 - consistency_penalty)

    # Timeliness / Structure (10 pts)
    date_issues   = [i for i in issues if i["category"] == "Date Anomalies"]
    struct_issues = [i for i in issues if i["category"] in ("Redundant Data", "Empty Data")]
    structure_penalty = min(10, len(date_issues) * 3 + len(struct_issues) * 5)
    structure_score   = max(0, 10 - structure_penalty)

    overall = round(
        completeness_score + uniqueness_score + validity_score +
        consistency_score + structure_score, 1
    )

    return overall, {
        "Completeness":  (completeness_score, 30),
        "Uniqueness":    (uniqueness_score, 20),
        "Validity":      (validity_score, 25),
        "Consistency":   (consistency_score, 15),
        "Structure":     (structure_score, 10),
    }


def grade_score(score):
    if score >= 95: return "A+", "#34D399"
    if score >= 88: return "A",  "#10B981"
    if score >= 80: return "B+", "#60A5FA"
    if score >= 72: return "B",  "#3B82F6"
    if score >= 62: return "C+", "#FCD34D"
    if score >= 50: return "C",  "#F59E0B"
    if score >= 35: return "D",  "#F97316"
    return "F", "#EF4444"


# ══════════════════════════════════════════════════════════════════════════════
# CLEANING ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def apply_cleaning(df_raw, selected_fixes, options, col_types, issues):
    """
    Apply selected cleaning operations.
    Returns: cleaned_df, change_log (list of dicts)
    """
    df = df_raw.copy()
    log = []
    n_rows_start = len(df)

    # ── Drop fully empty columns ──────────────────────────────────────────────
    if "drop_empty_cols" in selected_fixes:
        empty_cols = [c for c in df.columns if df[c].isnull().all()]
        if empty_cols:
            df = df.drop(columns=empty_cols)
            log.append({
                "operation": "Dropped empty columns",
                "detail": f"Removed {len(empty_cols)} fully-null columns: {empty_cols}",
                "rows_affected": n_rows_start,
            })

    # ── Drop constant columns ─────────────────────────────────────────────────
    if "drop_constant" in selected_fixes:
        const_cols = [c for c in df.columns if df[c].nunique(dropna=True) <= 1]
        if const_cols:
            df = df.drop(columns=const_cols)
            log.append({
                "operation": "Dropped constant columns",
                "detail": f"Removed {len(const_cols)} constant columns: {const_cols}",
                "rows_affected": n_rows_start,
            })

    # ── Duplicates ────────────────────────────────────────────────────────────
    if "duplicates" in selected_fixes:
        n_before = len(df)
        df = df.drop_duplicates(keep="first")
        df = df.reset_index(drop=True)
        removed = n_before - len(df)
        if removed > 0:
            log.append({
                "operation": "Removed duplicate rows",
                "detail": f"Dropped {removed:,} duplicate rows. Kept first occurrence of each unique row.",
                "rows_affected": removed,
            })

    # ── Missing values ────────────────────────────────────────────────────────
    if "missing_values" in selected_fixes:
        strategy = options.get("missing_strategy", "smart")
        total_filled = 0
        for col in df.columns:
            if df[col].isnull().sum() == 0:
                continue
            n_missing = df[col].isnull().sum()
            ctype = col_types.get(col, "unknown")
            pct_missing = n_missing / len(df) * 100

            # Drop columns >80% missing if drop_high_missing option set
            if pct_missing > 80 and options.get("drop_high_missing", True):
                df = df.drop(columns=[col])
                log.append({
                    "operation": f"Dropped high-missing column '{col}'",
                    "detail": f"'{col}' was {pct_missing:.1f}% missing. Dropped as unreliable.",
                    "rows_affected": n_missing,
                })
                continue

            if strategy == "drop_rows":
                n_before_dr = len(df)
                df = df.dropna(subset=[col]).reset_index(drop=True)
                log.append({
                    "operation": f"Dropped rows with missing '{col}'",
                    "detail": f"Removed {n_before_dr - len(df):,} rows where '{col}' was null.",
                    "rows_affected": n_before_dr - len(df),
                })
            else:
                # Smart strategy
                if ctype in ("numeric", "categorical_numeric", "binary"):
                    fill_strategy = options.get("numeric_fill", "median")
                    if fill_strategy == "mean":
                        fill_val = df[col].mean()
                    elif fill_strategy == "zero":
                        fill_val = 0
                    else:
                        fill_val = df[col].median()
                    df[col] = df[col].fillna(fill_val)
                    total_filled += n_missing
                    log.append({
                        "operation": f"Filled missing values in '{col}'",
                        "detail": f"Filled {n_missing:,} nulls with {fill_strategy} = {fill_val:.4f}.",
                        "rows_affected": n_missing,
                    })
                elif ctype in ("categorical", "text"):
                    fill_val = df[col].mode()[0] if not df[col].mode().empty else "Unknown"
                    fill_label = options.get("categorical_fill", "mode")
                    if fill_label == "unknown":
                        fill_val = "Unknown"
                    elif fill_label == "drop":
                        n_bd = len(df)
                        df = df.dropna(subset=[col])
                        log.append({
                            "operation": f"Dropped rows with missing '{col}'",
                            "detail": f"Removed {n_bd - len(df):,} rows.",
                            "rows_affected": n_bd - len(df),
                        })
                        continue
                    df[col] = df[col].fillna(fill_val)
                    total_filled += n_missing
                    log.append({
                        "operation": f"Filled missing values in '{col}'",
                        "detail": f"Filled {n_missing:,} nulls in categorical column with '{fill_val}'.",
                        "rows_affected": n_missing,
                    })
                elif ctype == "datetime":
                    fill_method = options.get("date_fill", "ffill")
                    df[col] = df[col].ffill() if fill_method == "ffill" else df[col].bfill()
                    log.append({
                        "operation": f"Filled missing dates in '{col}'",
                        "detail": f"Used {fill_method} to fill {n_missing:,} missing dates.",
                        "rows_affected": n_missing,
                    })
                elif _is_string_dtype(df[col].dtype):
                    fill_val = df[col].mode()[0] if not df[col].mode().empty else "Unknown"
                    if options.get("categorical_fill") == "unknown":
                        fill_val = "Unknown"
                    df[col] = df[col].fillna(fill_val)
                    total_filled += n_missing
                    log.append({
                        "operation": f"Filled missing values in '{col}'",
                        "detail": f"Filled {n_missing:,} nulls in string column with '{fill_val}'.",
                        "rows_affected": n_missing,
                    })
                else:
                    fill_val = df[col].mode()[0] if not df[col].mode().empty else "Unknown"
                    df[col] = df[col].fillna(fill_val)
                    total_filled += n_missing
                    log.append({
                        "operation": f"Filled missing values in '{col}'",
                        "detail": f"Filled {n_missing:,} nulls with mode/Unknown.",
                        "rows_affected": n_missing,
                    })
        df = df.reset_index(drop=True)

    # ── Type conversion ───────────────────────────────────────────────────────
    if "type_conversion" in selected_fixes:
        for col in list(df.columns):
            if not _is_string_dtype(df[col].dtype):
                continue
            ctype = col_types.get(col, "unknown")
            s = df[col].dropna().astype(str)

            # Try numeric
            if ctype not in ("datetime", "categorical"):
                cleaned_s = df[col].astype(str).str.replace(r"[,\$£€₦\s]", "", regex=True).str.strip()
                numeric = pd.to_numeric(cleaned_s, errors="coerce")
                pct_ok = numeric.notna().mean()
                if pct_ok > 0.7:
                    df[col] = numeric
                    log.append({
                        "operation": f"Converted '{col}' to numeric",
                        "detail": f"Stripped currency/comma characters and converted to float64. {int(pct_ok*len(df)):,} values converted.",
                        "rows_affected": int(pct_ok * len(df)),
                    })
                    continue

            # Try datetime
            if ctype == "datetime":
                try:
                    df[col] = pd.to_datetime(df[col], infer_datetime_format=True, errors="coerce")
                    log.append({
                        "operation": f"Converted '{col}' to datetime",
                        "detail": f"Parsed date strings to datetime64. Unparseable values become NaT.",
                        "rows_affected": len(df),
                    })
                except Exception:
                    pass

    # ── String cleaning ───────────────────────────────────────────────────────
    if "string_cleaning" in selected_fixes:
        case_mode = options.get("case_mode", "title")
        str_cols_processed = []
        for col in df.columns:
            if not _is_string_dtype(df[col].dtype):
                continue
            changed = False
            # Strip whitespace - work on non-null values only
            mask_notna = df[col].notna()
            s_work = df[col][mask_notna].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
            if case_mode == "title":
                s_work = s_work.str.title()
            elif case_mode == "upper":
                s_work = s_work.str.upper()
            elif case_mode == "lower":
                s_work = s_work.str.lower()
            # Write back only non-null positions
            result = df[col].copy()
            result[mask_notna] = s_work.values
            changed = not result.equals(df[col])
            if changed:
                df[col] = result
                str_cols_processed.append(col)
        if str_cols_processed:
            log.append({
                "operation": "Cleaned string columns",
                "detail": f"Stripped whitespace and applied {case_mode} case to {len(str_cols_processed)} columns: {str_cols_processed[:8]}{'...' if len(str_cols_processed)>8 else ''}.",
                "rows_affected": len(df) * len(str_cols_processed),
            })

    # ── Outliers ──────────────────────────────────────────────────────────────
    if "outliers" in selected_fixes:
        outlier_method = options.get("outlier_method", "cap")
        outlier_cols_fixed = []
        for col in df.columns:
            if col_types.get(col) not in ("numeric", "categorical_numeric"):
                continue
            s = df[col].dropna()
            if len(s) < 10:
                continue
            q1, q3 = s.quantile(0.25), s.quantile(0.75)
            iqr = q3 - q1
            if iqr == 0:
                continue
            lower_fence = q1 - 1.5 * iqr
            upper_fence = q3 + 1.5 * iqr
            outlier_mask = (df[col] < lower_fence) | (df[col] > upper_fence)
            n_out = outlier_mask.sum()
            if n_out == 0:
                continue
            if outlier_method == "cap":
                df[col] = df[col].clip(lower=lower_fence, upper=upper_fence)
            elif outlier_method == "median":
                df.loc[outlier_mask, col] = df[col].median()
            elif outlier_method == "drop":
                df = df[~outlier_mask]
            outlier_cols_fixed.append(f"{col} ({n_out} rows)")
        if outlier_cols_fixed:
            df = df.reset_index(drop=True)
            log.append({
                "operation": f"Handled outliers ({outlier_method})",
                "detail": f"Applied '{outlier_method}' to outliers in: {'; '.join(outlier_cols_fixed[:6])}{'...' if len(outlier_cols_fixed)>6 else ''}.",
                "rows_affected": sum(int(s.split("(")[1].split(" ")[0]) for s in outlier_cols_fixed if "(" in s),
            })

    # ── Negative values ───────────────────────────────────────────────────────
    if "negative_values" in selected_fixes:
        neg_strategy = options.get("neg_strategy", "abs")
        neg_cols_fixed = []
        for col in df.columns:
            if col_types.get(col) not in ("numeric",):
                continue
            pos_hints = any(kw in col.lower() for kw in
                            ("price","cost","amount","qty","quantity","revenue","profit",
                             "age","count","total","sum","fee","salary","weight","height","sales","volume"))
            if not pos_hints:
                continue
            n_neg = (df[col] < 0).sum()
            if n_neg == 0:
                continue
            if neg_strategy == "abs":
                df[col] = df[col].abs()
            elif neg_strategy == "zero":
                df.loc[df[col] < 0, col] = 0
            elif neg_strategy == "nan":
                df.loc[df[col] < 0, col] = np.nan
            neg_cols_fixed.append(f"{col} ({n_neg})")
        if neg_cols_fixed:
            log.append({
                "operation": f"Fixed negative values ({neg_strategy})",
                "detail": f"Columns fixed: {'; '.join(neg_cols_fixed)}.",
                "rows_affected": sum(int(s.split("(")[1].rstrip(")")) for s in neg_cols_fixed if "(" in s),
            })

    return df, log


def to_download_bytes(df, fmt="csv"):
    """Convert DataFrame to bytes for download."""
    if fmt == "csv":
        return df.to_csv(index=False).encode("utf-8")
    elif fmt == "xlsx":
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Cleaned Data")
        return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE INIT
# ══════════════════════════════════════════════════════════════════════════════

def init_state():
    defaults = {
        "step": 1,           # 1=Upload, 2=Diagnose, 3=Configure, 4=Clean, 5=Verify
        "df_raw": None,
        "df_clean": None,
        "issues": [],
        "col_types": {},
        "selected_fixes": set(),
        "options": {},
        "change_log": [],
        "initial_score": None,
        "final_score": None,
        "file_name": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()
S = st.session_state


# ══════════════════════════════════════════════════════════════════════════════
# UI HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def render_header():
    st.markdown("""
    <div class="dc-header">
        <h1>🧹 DataCleanse Pro</h1>
        <p>Intelligent, automated data quality analysis and cleaning - upload, diagnose, clean, verify, and download.</p>
    </div>
    """, unsafe_allow_html=True)


def render_step_bar(current):
    steps = ["Upload", "Diagnose", "Configure", "Clean", "Verify & Download"]
    html = '<div class="step-bar">'
    for i, s in enumerate(steps, 1):
        cls = "active" if i == current else ("done" if i < current else "step-item")
        icon = "✓" if i < current else str(i)
        html += f'<div class="step-item {cls}">{icon} {s}</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_issue(issue):
    sev = issue["severity"]
    badge_cls = f"badge-{sev}"
    aff_rows = issue["affected_rows"]
    meta_str = f"Affected rows: {aff_rows:,}" if aff_rows else ""
    if issue.get("affected_columns"):
        meta_str += f" | Columns: {', '.join(issue['affected_columns'][:4])}"

    st.markdown(f"""
    <div class="issue-item {sev}">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px;">
            <div class="issue-title">{issue['title']}</div>
            <span class="badge {badge_cls}">{sev}</span>
        </div>
        <div class="issue-desc">{issue['description']}</div>
        <div class="issue-meta">
            📌 {issue['recommendation']}<br>
            {f'• {meta_str}' if meta_str else ''}
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_score_panel(score, breakdown, label="Quality Score"):
    grade, grade_col = grade_score(score)
    st.markdown(f"""
    <div class="score-display">
        <div class="score-number" style="background:linear-gradient(135deg, {grade_col}, #60A5FA);
             -webkit-background-clip:text; -webkit-text-fill-color:transparent;">
            {score:.0f}
        </div>
        <div class="score-grade" style="color:{grade_col};">{grade}</div>
        <div class="score-label">{label} / 100</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    for dim, (val, max_val) in breakdown.items():
        pct = val / max_val * 100
        color_cls = "prog-green" if pct >= 80 else "prog-blue" if pct >= 60 else "prog-amber" if pct >= 40 else "prog-red"
        st.markdown(f"""
        <div style="margin-bottom:12px;">
            <div style="display:flex; justify-content:space-between; font-size:0.8rem; color:var(--text-2); margin-bottom:4px;">
                <span>{dim}</span>
                <span>{val:.1f} / {max_val}</span>
            </div>
            <div class="prog-bar">
                <div class="prog-fill {color_cls}" style="width:{pct:.1f}%;"></div>
            </div>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 - UPLOAD
# ══════════════════════════════════════════════════════════════════════════════

def step_upload():
    render_header()
    render_step_bar(1)

    col_left, col_right = st.columns([3, 2], gap="large")

    with col_left:
        st.markdown("""
        <div class="dc-card">
            <div class="dc-card-header">
                <span style="font-size:1.3rem;">📂</span>
                <h3>Upload Your Dataset</h3>
            </div>
        """, unsafe_allow_html=True)

        uploaded = st.file_uploader(
            "Drag and drop your CSV or Excel file",
            type=["csv", "xlsx", "xls"],
            label_visibility="collapsed",
        )

        if uploaded:
            with st.spinner("Loading file..."):
                df, err = load_file(uploaded)
            if err:
                st.error(f"❌ Failed to load file: {err}")
            else:
                S.df_raw   = df
                S.file_name = uploaded.name
                S.step     = 1  # stay here until they click proceed

                st.success(f"✅ File loaded successfully: **{uploaded.name}**")

                # Quick preview
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("**Preview (first 5 rows):**")
                st.dataframe(df.head(), use_container_width=True, height=200)

        st.markdown("</div>", unsafe_allow_html=True)

        if S.df_raw is not None:
            if st.button("▶ Proceed to Diagnosis", use_container_width=True):
                with st.spinner("Analysing data quality..."):
                    issues, col_types = detect_issues(S.df_raw)
                    score, breakdown = compute_quality_score(S.df_raw, issues)
                S.issues       = issues
                S.col_types    = col_types
                S.initial_score = score
                S.step = 2
                st.rerun()

    with col_right:
        st.markdown("""
        <div class="dc-card">
            <div class="dc-card-header">
                <span style="font-size:1.3rem;">💡</span>
                <h3>What DataCleanse Pro Does</h3>
            </div>
        """, unsafe_allow_html=True)

        features = [
            ("🔍", "Diagnoses", "Detects 11 categories of data issues including missing values, duplicates, outliers, type mismatches, string inconsistencies, and more."),
            ("🎛️", "Configures", "You choose exactly which cleaning methods to apply. Smart defaults provided, but full control stays with you."),
            ("⚙️", "Cleans", "Applies intelligent cleaning with full audit trail showing every change made to every column."),
            ("📊", "Grades", "Re-scores your dataset after cleaning across 5 quality dimensions: Completeness, Uniqueness, Validity, Consistency, Structure."),
            ("⬇️", "Exports", "Download your cleaned data as CSV or Excel, ready for analysis."),
        ]

        for icon, title, desc in features:
            st.markdown(f"""
            <div style="display:flex; gap:12px; margin-bottom:14px; padding:12px;
                        background:var(--bg-panel); border-radius:8px; border:1px solid var(--border);">
                <div style="font-size:1.4rem; min-width:28px;">{icon}</div>
                <div>
                    <div style="font-weight:600; color:var(--text-1); font-size:0.88rem;">{title}</div>
                    <div style="color:var(--text-2); font-size:0.79rem; margin-top:3px;">{desc}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 - DIAGNOSE
# ══════════════════════════════════════════════════════════════════════════════

def step_diagnose():
    render_step_bar(2)

    df = S.df_raw
    issues = S.issues
    n_rows, n_cols = df.shape

    # ── Header metrics ────────────────────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    metric_style = """
    background:var(--bg-card); border:1px solid var(--border);
    border-radius:10px; padding:16px; text-align:center;
    """
    sev_counts = {s: sum(1 for i in issues if i["severity"] == s)
                  for s in ("critical","high","medium","low")}
    total_issues = len(issues)

    with c1:
        st.markdown(f'<div style="{metric_style}"><div style="font-size:1.8rem;font-weight:700;color:#F1F5F9;">{n_rows:,}</div><div style="color:#94A3B8;font-size:0.75rem;margin-top:4px;">ROWS</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div style="{metric_style}"><div style="font-size:1.8rem;font-weight:700;color:#F1F5F9;">{n_cols}</div><div style="color:#94A3B8;font-size:0.75rem;margin-top:4px;">COLUMNS</div></div>', unsafe_allow_html=True)
    with c3:
        col_c = "#EF4444" if sev_counts["critical"] > 0 else "#F1F5F9"
        st.markdown(f'<div style="{metric_style}"><div style="font-size:1.8rem;font-weight:700;color:{col_c};">{sev_counts["critical"]+sev_counts["high"]}</div><div style="color:#94A3B8;font-size:0.75rem;margin-top:4px;">CRITICAL / HIGH</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div style="{metric_style}"><div style="font-size:1.8rem;font-weight:700;color:#F59E0B;">{sev_counts["medium"]}</div><div style="color:#94A3B8;font-size:0.75rem;margin-top:4px;">MEDIUM</div></div>', unsafe_allow_html=True)
    with c5:
        miss_pct = round(df.isnull().sum().sum() / (n_rows*n_cols) * 100, 1)
        st.markdown(f'<div style="{metric_style}"><div style="font-size:1.8rem;font-weight:700;color:#60A5FA;">{miss_pct}%</div><div style="color:#94A3B8;font-size:0.75rem;margin-top:4px;">MISSING CELLS</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Main layout ───────────────────────────────────────────────────────────
    left_col, right_col = st.columns([3, 2], gap="large")

    with left_col:
        st.markdown("""
        <div class="dc-card">
            <div class="dc-card-header">
                <span style="font-size:1.3rem;">🔍</span>
                <h3>Detected Issues</h3>
            </div>
        """, unsafe_allow_html=True)

        if not issues:
            st.success("🎉 No issues detected! Your dataset appears clean.")
        else:
            # Group by category
            categories = {}
            for issue in issues:
                cat = issue["category"]
                if cat not in categories:
                    categories[cat] = []
                categories[cat].append(issue)

            for cat, cat_issues in categories.items():
                with st.expander(f"📁 {cat}  ({len(cat_issues)} issue{'s' if len(cat_issues)>1 else ''})", expanded=True):
                    for issue in cat_issues:
                        render_issue(issue)

        st.markdown("</div>", unsafe_allow_html=True)

    with right_col:
        # Quality score
        st.markdown('<div class="dc-card"><div class="dc-card-header"><span style="font-size:1.3rem;">📊</span><h3>Initial Quality Score</h3></div>', unsafe_allow_html=True)
        score, breakdown = compute_quality_score(df, issues)
        render_score_panel(score, breakdown, "Initial Quality Score")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Column type map
        st.markdown('<div class="dc-card"><div class="dc-card-header"><span style="font-size:1.3rem;">🗂️</span><h3>Column Type Analysis</h3></div>', unsafe_allow_html=True)
        type_df = pd.DataFrame([
            {"Column": col, "Dtype": str(df[col].dtype), "Inferred": S.col_types.get(col, "unknown"), "Nulls": f"{df[col].isnull().sum()} ({df[col].isnull().mean()*100:.1f}%)"}
            for col in df.columns
        ])
        st.dataframe(type_df, use_container_width=True, height=300,
                     hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_back, col_fwd = st.columns([1, 1])
    with col_back:
        if st.button("← Back to Upload", use_container_width=True):
            S.step = 1; st.rerun()
    with col_fwd:
        if st.button("▶ Configure Cleaning →", use_container_width=True):
            S.step = 3; st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 - CONFIGURE
# ══════════════════════════════════════════════════════════════════════════════

def step_configure():
    render_step_bar(3)

    issues = S.issues
    if not issues:
        st.info("No issues found. Your data is already clean.")
        if st.button("← Back"): S.step = 2; st.rerun()
        return

    # Build available fix categories from issues
    available_fixes = defaultdict(list)
    for issue in issues:
        available_fixes[issue["fix_key"]].append(issue)

    st.markdown("""
    <div class="dc-card">
        <div class="dc-card-header">
            <span style="font-size:1.3rem;">🎛️</span>
            <h3>Select Cleaning Operations</h3>
        </div>
        <p style="color:var(--text-2); font-size:0.85rem; margin-bottom:18px;">
            Review each detected issue and choose which operations to apply.
            Smart default options are pre-selected. Adjust the strategy for each operation below.
        </p>
    </div>
    """, unsafe_allow_html=True)

    selected = set()
    opts = {}

    # ── Fix group definitions ──────────────────────────────────────────────────
    FIX_GROUPS = {
        "drop_empty_cols": {
            "label": "🗑️ Drop fully-empty columns",
            "desc": "Remove columns where 100% of values are null. These columns carry zero information.",
            "default": True,
            "has_options": False,
        },
        "drop_constant": {
            "label": "🗑️ Drop constant columns",
            "desc": "Remove columns with only 1 unique value. They add noise to analysis and machine learning.",
            "default": True,
            "has_options": False,
        },
        "duplicates": {
            "label": "🔁 Remove duplicate rows",
            "desc": "Drop rows that are exact copies of another row. Keeps the first occurrence.",
            "default": True,
            "has_options": False,
        },
        "missing_values": {
            "label": "❓ Handle missing values",
            "desc": "Fill or remove null values using context-appropriate strategies per column type.",
            "default": True,
            "has_options": True,
        },
        "type_conversion": {
            "label": "🔄 Fix data type mismatches",
            "desc": "Convert text columns containing numbers/dates to their correct data type.",
            "default": True,
            "has_options": False,
        },
        "string_cleaning": {
            "label": "✂️ Clean string inconsistencies",
            "desc": "Strip whitespace, remove extra spaces, standardise letter case in text columns.",
            "default": True,
            "has_options": True,
        },
        "outliers": {
            "label": "📈 Handle outliers",
            "desc": "Detect and treat statistical outliers using IQR method on numeric columns.",
            "default": False,
            "has_options": True,
        },
        "negative_values": {
            "label": "➖ Fix negative values",
            "desc": "Correct negative values in columns that should logically be non-negative (price, quantity, age, etc.).",
            "default": True,
            "has_options": True,
        },
    }

    # Only show fix groups that have detected issues
    shown_groups = {k: v for k, v in FIX_GROUPS.items() if k in available_fixes}

    for fix_key, fix_def in shown_groups.items():
        n_issues = len(available_fixes[fix_key])
        n_rows_affected = sum(i["affected_rows"] for i in available_fixes[fix_key])
        sev_list = [i["severity"] for i in available_fixes[fix_key]]
        worst_sev = min(sev_list, key=lambda s: {"critical":0,"high":1,"medium":2,"low":3}.get(s,4))

        badge_cls = f"badge-{worst_sev}"

        col_check, col_detail = st.columns([1, 4])

        with col_check:
            checked = st.checkbox(
                fix_def["label"],
                value=fix_def["default"],
                key=f"chk_{fix_key}",
            )
        with col_detail:
            st.markdown(f"""
            <div style="padding:8px 0; color:var(--text-2); font-size:0.82rem; line-height:1.5;">
                {fix_def['desc']}<br>
                <span class="badge {badge_cls}">{worst_sev}</span>
                &nbsp; <span style="color:var(--text-3);">{n_issues} issue(s) | {n_rows_affected:,} rows affected</span>
            </div>
            """, unsafe_allow_html=True)

        if checked:
            selected.add(fix_key)

        # Options for this fix
        if checked and fix_def["has_options"]:
            with st.expander(f"⚙️ Configure: {fix_def['label'].replace('🗑️','').replace('❓','').replace('🔄','').replace('✂️','').replace('📈','').replace('➖','').strip()}", expanded=False):
                if fix_key == "missing_values":
                    opts["missing_strategy"] = st.radio(
                        "Strategy for missing values",
                        ["smart", "drop_rows"],
                        index=0,
                        format_func=lambda x: {
                            "smart": "Smart fill (recommended) - fills with mean/median/mode per column type",
                            "drop_rows": "Drop rows - remove any row with a missing value (may lose data)",
                        }[x],
                        key="missing_strat",
                    )
                    if opts["missing_strategy"] == "smart":
                        opts["numeric_fill"] = st.selectbox(
                            "Fill numeric columns with:",
                            ["median", "mean", "zero"],
                            index=0,
                            key="num_fill",
                        )
                        opts["categorical_fill"] = st.selectbox(
                            "Fill categorical columns with:",
                            ["mode", "unknown", "drop"],
                            index=0,
                            format_func=lambda x: {
                                "mode": "Mode (most frequent value)",
                                "unknown": '"Unknown" label',
                                "drop": "Drop those rows",
                            }[x],
                            key="cat_fill",
                        )
                        opts["drop_high_missing"] = st.checkbox(
                            "Drop columns that are >80% missing",
                            value=True,
                            key="drop_hi_miss",
                        )

                elif fix_key == "string_cleaning":
                    opts["case_mode"] = st.selectbox(
                        "Standardise letter case to:",
                        ["title", "upper", "lower", "none"],
                        index=0,
                        format_func=lambda x: {
                            "title": "Title Case (e.g. Lagos State)",
                            "upper": "UPPER CASE",
                            "lower": "lower case",
                            "none": "No case change - whitespace only",
                        }[x],
                        key="case_mode",
                    )

                elif fix_key == "outliers":
                    opts["outlier_method"] = st.radio(
                        "Outlier treatment method",
                        ["cap", "median", "drop"],
                        index=0,
                        format_func=lambda x: {
                            "cap": "Cap to 1.5xIQR fence (winsorize) - safest, preserves row count",
                            "median": "Replace with column median - conservative imputation",
                            "drop": "Drop outlier rows - reduces dataset size",
                        }[x],
                        key="outlier_meth",
                    )

                elif fix_key == "negative_values":
                    opts["neg_strategy"] = st.radio(
                        "How to fix negative values",
                        ["abs", "zero", "nan"],
                        index=0,
                        format_func=lambda x: {
                            "abs": "Take absolute value (convert -50 to 50)",
                            "zero": "Set to zero",
                            "nan": "Replace with NaN then re-impute",
                        }[x],
                        key="neg_strat",
                    )

        st.markdown("<hr style='border-color:var(--border); margin:4px 0;'>", unsafe_allow_html=True)

    # Save and proceed
    S.selected_fixes = selected
    S.options = opts

    st.markdown("<br>", unsafe_allow_html=True)
    col_back, col_fwd = st.columns([1, 1])
    with col_back:
        if st.button("← Back to Diagnosis", use_container_width=True):
            S.step = 2; st.rerun()
    with col_fwd:
        if st.button(f"▶ Apply {len(selected)} Cleaning Operations →", use_container_width=True, disabled=len(selected)==0):
            S.step = 4; st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4 - CLEAN
# ══════════════════════════════════════════════════════════════════════════════

def step_clean():
    render_step_bar(4)

    if S.df_clean is None:
        with st.spinner("🧹 Cleaning your dataset..."):
            df_cleaned, change_log = apply_cleaning(
                S.df_raw,
                S.selected_fixes,
                S.options,
                S.col_types,
                S.issues,
            )
        S.df_clean   = df_cleaned
        S.change_log = change_log

    df_raw   = S.df_raw
    df_clean = S.df_clean
    log      = S.change_log

    # ── Stats comparison ──────────────────────────────────────────────────────
    r_orig, c_orig = df_raw.shape
    r_clean, c_clean = df_clean.shape
    rows_removed = r_orig - r_clean
    cols_removed = c_orig - c_clean
    miss_before  = df_raw.isnull().sum().sum()
    miss_after   = df_clean.isnull().sum().sum()
    dups_before  = df_raw.duplicated().sum()
    dups_after   = df_clean.duplicated().sum()

    st.markdown("""
    <div class="dc-card">
        <div class="dc-card-header">
            <span style="font-size:1.3rem;">✅</span>
            <h3>Cleaning Complete</h3>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    def metric_tile(val, label, sub="", color="#34D399"):
        return f"""
        <div style="background:var(--bg-panel);border:1px solid var(--border);border-radius:10px;
                    padding:18px;text-align:center;">
            <div style="font-size:2rem;font-weight:700;color:{color};">{val}</div>
            <div style="font-size:0.75rem;color:#94A3B8;text-transform:uppercase;letter-spacing:0.06em;margin-top:4px;">{label}</div>
            {f'<div style="font-size:0.72rem;color:#64748B;margin-top:3px;">{sub}</div>' if sub else ''}
        </div>"""

    with c1:
        st.markdown(metric_tile(f"{len(log)}", "Operations Applied", f"{len(S.selected_fixes)} types"), unsafe_allow_html=True)
    with c2:
        color = "#34D399" if miss_after == 0 else "#F59E0B"
        st.markdown(metric_tile(f"{miss_after:,}", "Missing Cells Left", f"was {miss_before:,}", color), unsafe_allow_html=True)
    with c3:
        color = "#34D399" if dups_after == 0 else "#F59E0B"
        st.markdown(metric_tile(f"{dups_after}", "Duplicates Left", f"was {dups_before:,}", color), unsafe_allow_html=True)
    with c4:
        color = "#F59E0B" if rows_removed > 0 else "#34D399"
        st.markdown(metric_tile(f"{rows_removed:,}", "Rows Removed", f"{r_clean:,} remain", color), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Tabs: Log | Preview ───────────────────────────────────────────────────
    tab_log, tab_before, tab_after = st.tabs(["📋 Change Log", "🗂️ Original Data", "✅ Cleaned Data"])

    with tab_log:
        if log:
            for i, entry in enumerate(log, 1):
                st.markdown(f"""
                <div class="issue-item medium" style="border-left-color:#34D399;">
                    <div class="issue-title">#{i} — {entry['operation']}</div>
                    <div class="issue-desc">{entry['detail']}</div>
                    <div class="issue-meta">Rows affected: {entry['rows_affected']:,}</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("No changes were made.")

    with tab_before:
        st.caption(f"Original: {r_orig:,} rows × {c_orig} columns")
        st.dataframe(df_raw.head(50), use_container_width=True, height=350)

    with tab_after:
        st.caption(f"Cleaned: {r_clean:,} rows × {c_clean} columns")
        st.dataframe(df_clean.head(50), use_container_width=True, height=350)

    st.markdown("<br>", unsafe_allow_html=True)
    col_back, col_fwd = st.columns([1, 1])
    with col_back:
        if st.button("← Back to Configure", use_container_width=True):
            S.df_clean = None   # reset so it re-runs cleaning
            S.step = 3; st.rerun()
    with col_fwd:
        if st.button("▶ Verify Quality & Download →", use_container_width=True):
            S.step = 5; st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5 - VERIFY & DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════

def step_verify():
    render_step_bar(5)

    df_raw   = S.df_raw
    df_clean = S.df_clean

    # Re-run full quality analysis on cleaned data
    with st.spinner("Running post-cleaning quality audit..."):
        issues_after, _  = detect_issues(df_clean)
        score_after, breakdown_after = compute_quality_score(df_clean, issues_after)
        S.final_score = score_after

    score_before = S.initial_score
    score_after_val = score_after

    improvement = round(score_after_val - score_before, 1)
    grade_b, _ = grade_score(score_before)
    grade_a, col_a = grade_score(score_after_val)

    # ── Score comparison ──────────────────────────────────────────────────────
    st.markdown("""
    <div class="dc-card">
        <div class="dc-card-header">
            <span style="font-size:1.3rem;">📊</span>
            <h3>Quality Verification Report</h3>
        </div>
    </div>
    """, unsafe_allow_html=True)

    left_col, mid_col, right_col = st.columns([2, 1, 2], gap="large")

    with left_col:
        st.markdown("#### Before Cleaning")
        score_b, breakdown_b = compute_quality_score(df_raw, S.issues)
        render_score_panel(score_b, breakdown_b, "Before")

    with mid_col:
        arrow = "⬆️" if improvement > 0 else ("⬇️" if improvement < 0 else "➡️")
        arrow_color = "#34D399" if improvement > 0 else ("#EF4444" if improvement < 0 else "#94A3B8")
        st.markdown(f"""
        <div style="display:flex; flex-direction:column; align-items:center; justify-content:center;
                    height:100%; padding-top:60px; gap:12px;">
            <div style="font-size:2.5rem;">{arrow}</div>
            <div style="font-size:2rem; font-weight:800; color:{arrow_color};">
                {'+' if improvement > 0 else ''}{improvement:.1f}
            </div>
            <div style="font-size:0.78rem; color:var(--text-2); text-align:center;">
                Score improvement
            </div>
        </div>
        """, unsafe_allow_html=True)

    with right_col:
        st.markdown("#### After Cleaning")
        render_score_panel(score_after_val, breakdown_after, "After")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Remaining issues ──────────────────────────────────────────────────────
    if issues_after:
        st.markdown("""
        <div class="dc-card">
            <div class="dc-card-header">
                <span style="font-size:1.3rem;">⚠️</span>
                <h3>Remaining Issues After Cleaning</h3>
            </div>
            <p style="color:var(--text-2); font-size:0.84rem; margin-bottom:12px;">
                The following issues were not addressed by the selected cleaning operations,
                or require manual intervention.
            </p>
        """, unsafe_allow_html=True)
        for issue in issues_after[:10]:
            render_issue(issue)
        if len(issues_after) > 10:
            st.caption(f"... and {len(issues_after)-10} more issues.")
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.success("🎉 **No remaining issues detected!** Your dataset is fully clean.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Satisfaction check ────────────────────────────────────────────────────
    st.markdown("""
    <div class="dc-card">
        <div class="dc-card-header">
            <span style="font-size:1.3rem;">✅</span>
            <h3>Verify Your Satisfaction</h3>
        </div>
    """, unsafe_allow_html=True)

    satisfied = st.radio(
        "Is the cleaned data ready for your use case?",
        ["Yes - download the cleaned file", "No - I want to go back and adjust the cleaning options"],
        index=0,
        key="satisfaction",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if "Yes" in satisfied:
        # ── Download section ──────────────────────────────────────────────────
        st.markdown("""
        <div class="dc-card">
            <div class="dc-card-header">
                <span style="font-size:1.3rem;">⬇️</span>
                <h3>Download Cleaned Dataset</h3>
            </div>
        """, unsafe_allow_html=True)

        dl_col1, dl_col2, dl_col3 = st.columns([1, 1, 2])

        base_name = S.file_name.rsplit(".", 1)[0] if "." in S.file_name else S.file_name
        ts = datetime.now().strftime("%Y%m%d_%H%M")

        with dl_col1:
            csv_bytes = to_download_bytes(df_clean, "csv")
            st.download_button(
                label="⬇️  Download as CSV",
                data=csv_bytes,
                file_name=f"{base_name}_cleaned_{ts}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with dl_col2:
            xlsx_bytes = to_download_bytes(df_clean, "xlsx")
            st.download_button(
                label="⬇️  Download as Excel",
                data=xlsx_bytes,
                file_name=f"{base_name}_cleaned_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl_col3:
            r_clean, c_clean = df_clean.shape
            st.markdown(f"""
            <div style="background:var(--bg-panel); border:1px solid var(--border); border-radius:8px;
                        padding:14px; font-size:0.82rem; color:var(--text-2);">
                <b style="color:var(--text-1);">File details:</b><br>
                {r_clean:,} rows × {c_clean} columns<br>
                Quality score: <b style="color:{col_a};">{score_after_val:.0f}/100 ({grade_a})</b><br>
                Cleaned on: {datetime.now().strftime("%d %B %Y, %H:%M")}
            </div>
            """, unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

        # ── Download change report ────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        report_lines = [
            "DataCleanse Pro - Cleaning Report",
            f"File: {S.file_name}",
            f"Cleaned: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            f"Original: {df_raw.shape[0]:,} rows x {df_raw.shape[1]} cols",
            f"Cleaned : {df_clean.shape[0]:,} rows x {df_clean.shape[1]} cols",
            f"Quality Before: {score_before:.1f}/100 ({grade_b})",
            f"Quality After : {score_after_val:.1f}/100 ({grade_a})",
            f"Improvement   : +{improvement:.1f} points",
            "",
            "CHANGE LOG:",
        ]
        for i, entry in enumerate(S.change_log, 1):
            report_lines.append(f"  {i}. {entry['operation']}")
            report_lines.append(f"     {entry['detail']}")
            report_lines.append(f"     Rows affected: {entry['rows_affected']:,}")
            report_lines.append("")

        report_bytes = "\n".join(report_lines).encode("utf-8")
        st.download_button(
            label="📄 Download Cleaning Report (.txt)",
            data=report_bytes,
            file_name=f"{base_name}_cleaning_report_{ts}.txt",
            mime="text/plain",
        )

        # ── Start over ────────────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Start Over with a New File"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    else:
        st.markdown("<br>", unsafe_allow_html=True)
        col_back, _ = st.columns([1, 2])
        with col_back:
            if st.button("← Back to Configure Cleaning", use_container_width=True):
                S.df_clean = None
                S.step = 3; st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════════════════

def main():
    step = S.step

    if step == 1:
        step_upload()
    elif step == 2:
        step_diagnose()
    elif step == 3:
        step_configure()
    elif step == 4:
        step_clean()
    elif step == 5:
        step_verify()
    else:
        step_upload()

main()
