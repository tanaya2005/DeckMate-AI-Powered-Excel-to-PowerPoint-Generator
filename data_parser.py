"""
data_parser.py — Excel → aggregated summary (Day 1)

Takes an uploaded .xlsx file (multi-tab) and returns a JSON-safe summary dict,
one entry per sheet. Only aggregated stats are returned — NEVER raw row-level data.

This is the first stage of the DeckMate pipeline:
    Excel → [data_parser] → summary dict → LLM planner → PPTX builder
"""

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from io import BytesIO

import pandas as pd
import numpy as np


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _infer_column_type(series: pd.Series) -> str:
    """Return a human-readable type label for a pandas Series."""
    if pd.api.types.is_bool_dtype(series):
        return "boolean"
    if pd.api.types.is_numeric_dtype(series):
        if pd.api.types.is_float_dtype(series):
            return "numeric (float)"
        return "numeric (integer)"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "datetime"
    # Try to parse as dates if the column is object-type
    if series.dtype == object:
        try:
            import warnings
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                parsed = pd.to_datetime(series, errors="coerce")
            non_null = parsed.dropna()
            if len(non_null) >= 0.5 * len(series.dropna()) and len(non_null) > 0:
                return "datetime (inferred)"
        except Exception:
            pass
    return "categorical / text"


def _safe_value(val: Any) -> Any:
    """Convert numpy/pandas scalars to JSON-safe Python types."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        return round(float(val), 4)
    if isinstance(val, (np.bool_,)):
        return bool(val)
    if isinstance(val, pd.Timestamp):
        return val.isoformat()
    return val


def _detect_header_row(df_raw: pd.DataFrame, max_scan: int = 10) -> int:
    """
    Heuristic: scan the first `max_scan` rows and pick the one that looks most
    like a header (most non-null, non-numeric, unique string values).
    Returns 0 if the existing first row already looks fine.
    """
    best_row = 0
    best_score = 0

    scan_limit = min(max_scan, len(df_raw))
    for i in range(scan_limit):
        row = df_raw.iloc[i]
        non_null = row.dropna()
        if len(non_null) == 0:
            continue
        # Score: count of string values that are unique
        str_vals = [v for v in non_null if isinstance(v, str) and v.strip()]
        score = len(set(str_vals))
        # Bonus: penalise rows where most values are numbers (likely data, not header)
        num_count = sum(1 for v in non_null if isinstance(v, (int, float)) and not isinstance(v, bool))
        score -= num_count
        if score > best_score:
            best_score = score
            best_row = i

    return best_row


# ---------------------------------------------------------------------------
# Core parsing
# ---------------------------------------------------------------------------

def _summarize_numeric(series: pd.Series) -> Dict[str, Any]:
    """Aggregate stats for a numeric column."""
    clean = series.dropna()
    if clean.empty:
        return {"count": 0, "non_null": 0, "note": "all values missing"}
    return {
        "count": int(len(series)),
        "non_null": int(len(clean)),
        "missing_pct": round((series.isna().sum() / len(series)) * 100, 1),
        "sum": _safe_value(clean.sum()),
        "mean": _safe_value(clean.mean()),
        "median": _safe_value(clean.median()),
        "min": _safe_value(clean.min()),
        "max": _safe_value(clean.max()),
        "std_dev": _safe_value(clean.std()),
    }


def _summarize_categorical(series: pd.Series, top_n: int = 5) -> Dict[str, Any]:
    """Aggregate stats for a categorical / text column."""
    clean = series.dropna().astype(str)
    if clean.empty:
        return {"count": 0, "non_null": 0, "note": "all values missing"}

    n_unique = int(clean.nunique())
    n_non_null = int(len(clean))

    # Flag high-cardinality columns that look like identifiers
    # (e.g. client names, order IDs, email addresses)
    is_identifier_like = (
        n_non_null >= 10 and (n_unique / n_non_null) >= 0.9
    )

    value_counts = clean.value_counts().head(top_n)
    result = {
        "count": int(len(series)),
        "non_null": n_non_null,
        "missing_pct": round((series.isna().sum() / len(series)) * 100, 1),
        "unique_values": n_unique,
    }

    if is_identifier_like:
        result["identifier_like"] = True
        result["note"] = (
            "High-cardinality column — likely an identifier (e.g. names, IDs). "
            "Top values suppressed to avoid data leakage."
        )
    else:
        result["identifier_like"] = False
        result["top_values"] = {str(k): int(v) for k, v in value_counts.items()}

    return result


def _summarize_datetime(series: pd.Series) -> Dict[str, Any]:
    """Aggregate stats for a datetime column."""
    # Try converting if not already datetime
    if not pd.api.types.is_datetime64_any_dtype(series):
        series = pd.to_datetime(series, errors="coerce")
    clean = series.dropna()
    if clean.empty:
        return {"count": 0, "non_null": 0, "note": "all values missing"}
    return {
        "count": int(len(series)),
        "non_null": int(len(clean)),
        "missing_pct": round((series.isna().sum() / len(series)) * 100, 1),
        "min_date": _safe_value(clean.min()),
        "max_date": _safe_value(clean.max()),
        "date_range_days": int((clean.max() - clean.min()).days),
    }


def _summarize_boolean(series: pd.Series) -> Dict[str, Any]:
    """Aggregate stats for a boolean column."""
    clean = series.dropna()
    if clean.empty:
        return {"count": 0, "non_null": 0, "note": "all values missing"}
    return {
        "count": int(len(series)),
        "non_null": int(len(clean)),
        "missing_pct": round((series.isna().sum() / len(series)) * 100, 1),
        "true_count": int(clean.sum()),
        "false_count": int((~clean).sum()),
        "true_pct": round(float(clean.mean()) * 100, 1),
    }


def summarize_sheet(df: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
    """
    Produce an aggregate-only summary for a single dataframe (sheet).
    No raw rows are ever included.
    """
    if df.empty or df.shape[0] == 0:
        return {
            "sheet_name": sheet_name,
            "row_count": 0,
            "column_count": 0,
            "note": "Sheet is empty or has no data rows.",
            "columns": {},
        }

    columns_summary: Dict[str, Dict[str, Any]] = {}

    for col in df.columns:
        col_type = _infer_column_type(df[col])

        if col_type == "boolean":
            stats = _summarize_boolean(df[col])
        elif col_type.startswith("numeric"):
            stats = _summarize_numeric(df[col])
        elif col_type.startswith("datetime"):
            stats = _summarize_datetime(df[col])
        else:
            stats = _summarize_categorical(df[col])

        columns_summary[str(col)] = {
            "inferred_type": col_type,
            **stats,
        }

    return {
        "sheet_name": sheet_name,
        "row_count": int(len(df)),
        "column_count": int(len(df.columns)),
        "columns": columns_summary,
    }


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_excel(
    source: Union[str, Path, BytesIO],
    header_detection: bool = True,
) -> tuple[Dict[str, Any], Dict[str, pd.DataFrame]]:
    """
    Parse a multi-tab .xlsx file and return:
      1. summary  — JSON-safe dict with aggregate stats per sheet
      2. dataframes — dict[sheet_name → cleaned DataFrame] (kept in memory for
                      the PPTX builder to pull real chart data from later)

    Parameters
    ----------
    source : str | Path | BytesIO
        File path or in-memory bytes of the Excel workbook.
    header_detection : bool
        If True, attempt to auto-detect the real header row when headers
        are not in row 0.

    Returns
    -------
    (summary_dict, dataframes_dict)
    """
    # Read all sheets with no header assumption first (for header detection)
    try:
        all_sheets_raw = pd.read_excel(source, sheet_name=None, header=None, engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"Could not read the Excel file: {exc}") from exc

    summary: Dict[str, Any] = {}
    dataframes: Dict[str, pd.DataFrame] = {}

    for sheet_name, df_raw in all_sheets_raw.items():
        # Skip fully empty sheets
        if df_raw.dropna(how="all").empty:
            summary[sheet_name] = {
                "sheet_name": sheet_name,
                "row_count": 0,
                "column_count": 0,
                "note": "Sheet is empty.",
                "columns": {},
            }
            dataframes[sheet_name] = pd.DataFrame()
            continue

        # ---- Header detection ----
        if header_detection:
            header_row = _detect_header_row(df_raw)
        else:
            header_row = 0

        # Re-read with the detected header row
        if isinstance(source, BytesIO):
            source.seek(0)
        df = pd.read_excel(source, sheet_name=sheet_name, header=header_row, engine="openpyxl")
        if isinstance(source, BytesIO):
            source.seek(0)

        # Drop fully-empty rows and columns
        df = df.dropna(how="all").dropna(axis=1, how="all")

        # Clean column names — strip whitespace, deduplicate
        df.columns = [
            str(c).strip() if not str(c).startswith("Unnamed") else f"Column_{i+1}"
            for i, c in enumerate(df.columns)
        ]

        dataframes[sheet_name] = df
        summary[sheet_name] = summarize_sheet(df, sheet_name)

    return summary, dataframes


# ---------------------------------------------------------------------------
# CLI test block
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # If a file path is given as argument, parse it.
    # Otherwise, generate a small sample multi-tab workbook for testing.

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        print(f"Parsing: {file_path}\n")
        summary, dfs = parse_excel(file_path)
    else:
        # ---------- Create a sample multi-tab Excel for self-testing ----------
        print("No file provided — generating a sample multi-tab Excel for testing.\n")
        from io import BytesIO as _BytesIO

        buf = _BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            # Sheet 1: Regional Sales
            sales = pd.DataFrame({
                "Region": ["North", "South", "East", "West", "North", "South",
                            "East", "West", "North", "South"] * 3,
                "Product": ["Widget A", "Widget B", "Widget C", "Widget A",
                            "Widget B", "Widget C", "Widget A", "Widget B",
                            "Widget C", "Widget A"] * 3,
                "Revenue": np.random.randint(10_000, 500_000, size=30),
                "Units Sold": np.random.randint(50, 5000, size=30),
                "Date": pd.date_range("2024-01-01", periods=30, freq="M"),
            })
            sales.to_excel(writer, sheet_name="Regional Sales", index=False)

            # Sheet 2: Customer Feedback (mostly text)
            feedback = pd.DataFrame({
                "Customer Segment": np.random.choice(
                    ["Enterprise", "SMB", "Startup", "Government"], size=20
                ),
                "Satisfaction": np.random.choice(
                    ["Very Satisfied", "Satisfied", "Neutral", "Dissatisfied"], size=20
                ),
                "NPS Score": np.random.randint(1, 11, size=20),
                "Response Date": pd.date_range("2024-06-01", periods=20, freq="W"),
            })
            feedback.to_excel(writer, sheet_name="Customer Feedback", index=False)

            # Sheet 3: Nearly empty sheet (edge case — headers but all NaN data)
            empty_ish = pd.DataFrame({
                "ColA": [np.nan, np.nan],
                "ColB": [np.nan, np.nan],
            })
            empty_ish.to_excel(writer, sheet_name="Empty Sheet", index=False)

            # Sheet 4: Messy header (header on row 3)
            messy = pd.DataFrame([
                ["", "", "", ""],
                ["", "", "", ""],
                ["Quarter", "Channel", "Spend", "ROI %"],
                ["Q1", "Online", 50000, 12.5],
                ["Q1", "Retail", 35000, 8.3],
                ["Q2", "Online", 62000, 15.1],
                ["Q2", "Retail", 41000, 9.7],
                ["Q3", "Online", 71000, 18.2],
                ["Q3", "Retail", 45000, 11.0],
                ["Q4", "Online", 80000, 20.5],
                ["Q4", "Retail", 52000, 13.4],
            ])
            messy.to_excel(writer, sheet_name="Marketing Spend", index=False, header=False)

        buf.seek(0)
        summary, dfs = parse_excel(buf)

    # Pretty-print the aggregate summary
    print("=" * 60)
    print("AGGREGATE SUMMARY (this is what gets sent to the LLM)")
    print("=" * 60)
    print(json.dumps(summary, indent=2, default=str))

    # Quick sanity check: no raw rows should appear
    print("\n" + "=" * 60)
    print("DATAFRAMES (kept in-memory for chart building, NOT sent to LLM)")
    print("=" * 60)
    for name, df in dfs.items():
        print(f"\n--- {name} ---")
        print(f"  Shape: {df.shape}")
        if not df.empty:
            print(f"  Columns: {list(df.columns)}")
            print(f"  First 3 rows (for dev eyes only):")
            print(df.head(3).to_string(index=False))
