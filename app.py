# Copyright (c) 2026 George Gilman
# SPDX-License-Identifier: MIT
"""
Vine ETV Analyzer
-----------------
A lightweight Streamlit GUI application that dynamically parses Amazon Vine
Itemized Reports (.xlsx), tracks your floating window of Estimated Tax Value
(ETV) against a rolling cap (e.g., $6,000 over 30 days), and projects when
older orders roll off so you can acquire items again.

Run instructions:
    pip install streamlit pandas openpyxl plotly
    streamlit run app.py
"""

from __future__ import annotations

import io
import math
import os
import re
import threading
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Constants & Canonical Column Mapping
# ---------------------------------------------------------------------------

REQUIRED_COLUMNS: dict[str, list[str]] = {
    "order_date": ["order date", "orderdate", "transaction date", "date"],
    "etv": [
        "estimated tax value",
        "estimated tax value (etv)",
        "etv",
        "tax value",
        "fair market value",
        "fmv",
    ],
}

OPTIONAL_COLUMNS: dict[str, list[str]] = {
    "order_type": ["order type", "ordertype", "type", "transaction type"],
    "cancelled_date": [
        "cancelled date",
        "canceled date",
        "cancellation date",
        "cancel date",
    ],
    "order_number": ["order number", "order id", "orderid", "order #"],
    "asin": ["asin", "item id"],
    "product_name": [
        "product name",
        "item name",
        "title",
        "description",
        "product description",
    ],
}


# ---------------------------------------------------------------------------
# Dynamic Excel Parsing & Cleaning
# ---------------------------------------------------------------------------


def _normalize_header(val: Any) -> str:
    """Convert a cell value into a normalized lowercase header string."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    text = str(val).strip().lower()
    return re.sub(r"\s+", " ", text)


def _find_matching_col(
    headers: list[str],
    aliases: list[str],
    used_indices: set[int],
    exclude_terms: tuple[str, ...] = (),
) -> int | None:
    """
    Find the best matching column index for `aliases` in `headers`.
    Pass 1: Exact header match against aliases (in alias priority order).
    Pass 2: Word-boundary substring match against aliases (skipping excluded terms).
    """
    # Pass 1: Exact match in alias priority order
    for alias in aliases:
        for idx, h in enumerate(headers):
            if idx in used_indices or not h:
                continue
            if h == alias:
                return idx

    # Pass 2: Substring match on word boundaries in alias priority order
    for alias in aliases:
        pattern = re.compile(rf"(?:^|\b){re.escape(alias)}(?:\b|$)")
        for idx, h in enumerate(headers):
            if idx in used_indices or not h:
                continue
            if exclude_terms and any(term in h for term in exclude_terms):
                continue
            if pattern.search(h):
                return idx

    return None


def _match_column_map(headers: list[str]) -> dict[str, int] | None:
    """
    Attempt to map canonical column keys to column indices in `headers`.
    Returns a dictionary of canonical_name -> col_idx if all REQUIRED_COLUMNS
    are present, otherwise None.
    """
    col_map: dict[str, int] = {}
    used_indices: set[int] = set()

    for canon_key, aliases in REQUIRED_COLUMNS.items():
        exclude = ("cancel", "ship", "deliver") if canon_key == "order_date" else ()
        matched_idx = _find_matching_col(headers, aliases, used_indices, exclude_terms=exclude)
        if matched_idx is None:
            return None
        col_map[canon_key] = matched_idx
        used_indices.add(matched_idx)

    for canon_key, aliases in OPTIONAL_COLUMNS.items():
        matched_idx = _find_matching_col(headers, aliases, used_indices)
        if matched_idx is not None:
            col_map[canon_key] = matched_idx
            used_indices.add(matched_idx)

    return col_map


def _parse_currency(val: Any) -> float:
    """
    Parse a currency value into a finite float.
    Handles numeric types, '$1,234.56', '-$45.00', '($45.00)', etc.
    Raises ValueError if unparseable, boolean, NaN, or infinite.
    """
    if val is None or isinstance(val, bool) or (isinstance(val, float) and pd.isna(val)):
        raise ValueError(f"Empty or boolean currency value: {val!r}")
    if isinstance(val, (int, float)):
        num = float(val)
        if not math.isfinite(num):
            raise ValueError(f"Non-finite currency value: {val!r}")
        return num

    s = str(val).strip()
    if not s or s.lower() in {"nan", "none", "n/a", "-", "--", "inf", "-inf"}:
        raise ValueError(f"Non-numeric currency string: {val!r}")

    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()
    if "-" in s:
        is_negative = True
        s = s.replace("-", "").strip()

    s = re.sub(r"[$€£,\s]", "", s)
    num = float(s)
    if not math.isfinite(num):
        raise ValueError(f"Non-finite currency string: {val!r}")
    return -abs(num) if is_negative else num


def _parse_datetime(val: Any) -> pd.Timestamp | None:
    """
    Parse a cell value into a normalized (midnight) pandas Timestamp.
    Returns None if the value is not a valid calendar date (e.g., plain numbers,
    bare 4-digit years, booleans, or FAQ text).
    """
    if val is None or isinstance(val, (bool, int, float)):
        return None
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return None
        ts = val.tz_localize(None) if val.tzinfo is not None else val
        return ts.normalize()
    if isinstance(val, datetime):
        ts = pd.Timestamp(val)
        if ts.tzinfo is not None:
            ts = ts.tz_localize(None)
        return ts.normalize()
    if isinstance(val, date):
        return pd.Timestamp(val).normalize()

    s = str(val).strip()
    if not s or s.lower() in {"nan", "none", "nat", "n/a", ""}:
        return None

    # Ignore long informational / FAQ strings or bare numbers/years (e.g., "2026")
    if len(s) > 50 or re.fullmatch(r"[-+]?\d+(?:\.\d+)?", s):
        return None

    try:
        ts = pd.to_datetime(s, errors="raise")
        if pd.isna(ts):
            return None
        if ts.tzinfo is not None:
            ts = ts.tz_localize(None)
        return ts.normalize()
    except (ValueError, TypeError, pd.errors.ParserError, pd.errors.OutOfBoundsDatetime):
        return None


@st.cache_data(show_spinner=False)
def _load_vine_excel_bytes(raw_bytes: bytes) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Cached wrapper around `load_vine_excel` for in-memory workbook bytes."""
    return load_vine_excel(io.BytesIO(raw_bytes))


def load_vine_excel(file_source: Any) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Dynamically locate and parse the Amazon Vine itemized table from an Excel
    workbook (.xlsx). Automatically skips title/metadata header rows and strips
    out non-data footer/FAQ rows.

    Returns:
        (cleaned_dataframe, metadata_dict)
    Raises:
        ValueError with descriptive explanation on invalid/unreadable data.
    """
    if hasattr(file_source, "seek"):
        try:
            file_source.seek(0)
        except (OSError, io.UnsupportedOperation):
            pass

    try:
        excel_file = pd.ExcelFile(file_source, engine="openpyxl")
    except (ValueError, OSError, KeyError, RuntimeError) as exc:
        raise ValueError(
            f"Unable to open Excel file. Please ensure the file is a valid '.xlsx' workbook. Details: {exc}"
        ) from exc

    if not excel_file.sheet_names:
        raise ValueError("The uploaded Excel workbook contains no worksheets.")

    selected_sheet = None
    header_row_idx = None
    col_map: dict[str, int] | None = None
    raw_df: pd.DataFrame | None = None

    # Scan each worksheet to find the table header row containing Order Date & ETV
    for sheet_name in excel_file.sheet_names:
        df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        if df_sheet.empty:
            continue

        max_scan = min(len(df_sheet), 100)
        for r_idx in range(max_scan):
            row_values = [_normalize_header(v) for v in df_sheet.iloc[r_idx].tolist()]
            matched = _match_column_map(row_values)
            if matched is not None:
                selected_sheet = sheet_name
                header_row_idx = r_idx
                col_map = matched
                raw_df = df_sheet
                break
        if col_map is not None:
            break

    if raw_df is None or header_row_idx is None or col_map is None:
        raise ValueError(
            "Could not locate the itemized data table. Ensure the worksheet includes "
            "'Order Date' and 'Estimated Tax Value' column headers."
        )

    # Extract original column names from the identified header row
    raw_headers = [
        str(v).strip() if not pd.isna(v) else f"Column_{i}"
        for i, v in enumerate(raw_df.iloc[header_row_idx].tolist())
    ]

    data_slice = raw_df.iloc[header_row_idx + 1 :].copy()
    if data_slice.empty:
        raise ValueError("The data table header was found, but there are no rows below it.")

    records: list[dict[str, Any]] = []
    skipped_footer_rows = 0
    malformed_rows = 0

    order_date_idx = col_map["order_date"]
    etv_idx = col_map["etv"]
    order_type_idx = col_map.get("order_type")
    cancelled_date_idx = col_map.get("cancelled_date")
    order_num_idx = col_map.get("order_number")
    asin_idx = col_map.get("asin")
    product_name_idx = col_map.get("product_name")

    for _, row in data_slice.iterrows():
        # Check if entire row is empty
        if row.isna().all():
            continue

        raw_order_date = row.iloc[order_date_idx]
        raw_etv = row.iloc[etv_idx]

        parsed_order_date = _parse_datetime(raw_order_date)

        # If Order Date is not a valid date, treat as a footer/FAQ note row
        if parsed_order_date is None:
            skipped_footer_rows += 1
            continue

        try:
            parsed_etv = _parse_currency(raw_etv)
        except ValueError:
            malformed_rows += 1
            continue

        # Order Type handling
        order_type_val = "ORDER"
        if order_type_idx is not None:
            raw_ot = row.iloc[order_type_idx]
            if raw_ot is not None and not pd.isna(raw_ot) and str(raw_ot).strip():
                order_type_val = str(raw_ot).strip().upper()
            elif parsed_etv < 0:
                order_type_val = "CANCELLATION"
        elif parsed_etv < 0:
            order_type_val = "CANCELLATION"

        # If Order Type explicitly indicates CANCELLATION and ETV was recorded as positive,
        # ensure ETV is negative per specification ("negative for cancellations").
        if "CANCEL" in order_type_val and parsed_etv > 0:
            parsed_etv = -abs(parsed_etv)

        # Cancelled Date handling
        parsed_cancelled_date = None
        if cancelled_date_idx is not None:
            parsed_cancelled_date = _parse_datetime(row.iloc[cancelled_date_idx])

        record: dict[str, Any] = {
            "Order Date": parsed_order_date,
            "Estimated Tax Value": round(parsed_etv, 2),
            "Order Type": order_type_val,
            "Cancelled Date": parsed_cancelled_date,
        }

        if order_num_idx is not None:
            val = row.iloc[order_num_idx]
            record["Order Number"] = "" if pd.isna(val) else str(val).strip()
        if asin_idx is not None:
            val = row.iloc[asin_idx]
            record["ASIN"] = "" if pd.isna(val) else str(val).strip()
        if product_name_idx is not None:
            val = row.iloc[product_name_idx]
            record["Product Name"] = "" if pd.isna(val) else str(val).strip()

        records.append(record)

    if not records:
        raise ValueError(
            "No valid order rows with both a valid 'Order Date' and numeric 'Estimated Tax Value' were found."
        )

    cleaned_df = (
        pd.DataFrame(records)
        .sort_values(by="Order Date", ascending=True)
        .reset_index(drop=True)
    )

    metadata = {
        "sheet_name": selected_sheet,
        "header_row_number": header_row_idx + 1,
        "detected_headers": raw_headers,
        "valid_rows": len(cleaned_df),
        "skipped_footer_rows": skipped_footer_rows,
        "malformed_rows": malformed_rows,
        "min_order_date": cleaned_df["Order Date"].min(),
        "max_order_date": cleaned_df["Order Date"].max(),
    }

    return cleaned_df, metadata


# ---------------------------------------------------------------------------
# Calculation Engine
# ---------------------------------------------------------------------------


def compute_rolling_window_total(
    df: pd.DataFrame,
    target_date: pd.Timestamp,
    window_days: int,
    inclusive_start: bool = False,
) -> float:
    """
    Calculate the sum of `Estimated Tax Value` for orders within the trailing
    window ending on `target_date`.

    By specification:
      - For Day D, items ordered in `(D - Window Length, D]` are summed.
      - When `inclusive_start` is False (default), interval is `(D - window_days, D]`,
        which spans `window_days` calendar days ending on D.
      - When `inclusive_start` is True, interval is `[D - window_days, D]`.
    """
    target_ts = pd.Timestamp(target_date).normalize()
    cutoff_ts = target_ts - pd.Timedelta(days=int(window_days))

    if inclusive_start:
        mask = (df["Order Date"] >= cutoff_ts) & (df["Order Date"] <= target_ts)
    else:
        mask = (df["Order Date"] > cutoff_ts) & (df["Order Date"] <= target_ts)

    total = float(df.loc[mask, "Estimated Tax Value"].sum())
    return round(total, 2)


def generate_decay_projection(
    df: pd.DataFrame,
    base_date: pd.Timestamp,
    window_days: int,
    threshold: float,
    horizon_days: int,
    inclusive_start: bool = False,
) -> tuple[pd.DataFrame, dict[str, Any] | None]:
    """
    Project the daily rolling window total for each day D from Day 0 (`base_date`)
    through `base_date + horizon_days` assuming no new additions after `base_date`.

    Returns:
        (projection_df, crossing_info)
        where `crossing_info` is a dict describing when the total drops below
        `threshold` (or None if it never drops below within the horizon).
    """
    base_ts = pd.Timestamp(base_date).normalize()
    rows: list[dict[str, Any]] = []
    crossing_info: dict[str, Any] | None = None

    for day_offset in range(int(horizon_days) + 1):
        proj_date = base_ts + pd.Timedelta(days=day_offset)
        window_start = proj_date - pd.Timedelta(days=int(window_days))

        projected_total = compute_rolling_window_total(
            df=df,
            target_date=proj_date,
            window_days=window_days,
            inclusive_start=inclusive_start,
        )

        variance = round(projected_total - float(threshold), 2)
        is_below = projected_total < float(threshold)
        status = "Below Threshold" if is_below else "Above Threshold"

        if is_below and crossing_info is None:
            crossing_info = {
                "date": proj_date,
                "day_offset": day_offset,
                "projected_total": projected_total,
                "variance": variance,
                "already_below": (day_offset == 0),
            }

        rows.append(
            {
                "Day Offset": day_offset,
                "Date": proj_date,
                "Window Start": window_start
                + (pd.Timedelta(days=0) if inclusive_start else pd.Timedelta(days=1)),
                "Window End": proj_date,
                "Projected Total": projected_total,
                "Variance to Threshold": variance,
                "Status": status,
            }
        )

    proj_df = pd.DataFrame(rows)
    proj_df["Daily Roll-Off ($)"] = (
        proj_df["Projected Total"].diff().fillna(0.0).round(2)
    )
    return proj_df, crossing_info


def compute_timeframe_summary(
    df: pd.DataFrame,
    base_date: pd.Timestamp,
    inclusive_start: bool = False,
) -> pd.DataFrame:
    """
    Build the Timeframe Summary Table for standard rolling periods:
    7 Days, 30 Days, 60 Days, 90 Days, and Full Year Total.
    """
    base_ts = pd.Timestamp(base_date).normalize()
    periods = [
        ("7 Days", 7),
        ("30 Days", 30),
        ("60 Days", 60),
        ("90 Days", 90),
    ]

    summary_rows: list[dict[str, Any]] = []
    for label, days in periods:
        cutoff = base_ts - pd.Timedelta(days=days)
        if inclusive_start:
            sub = df[(df["Order Date"] >= cutoff) & (df["Order Date"] <= base_ts)]
            start_display = cutoff.strftime("%Y-%m-%d")
        else:
            sub = df[(df["Order Date"] > cutoff) & (df["Order Date"] <= base_ts)]
            start_display = (cutoff + pd.Timedelta(days=1)).strftime("%Y-%m-%d")

        orders_mask = sub["Estimated Tax Value"] > 0
        cancels_mask = sub["Estimated Tax Value"] < 0

        summary_rows.append(
            {
                "Timeframe": f"Rolling {label}",
                "Date Range": f"{start_display} to {base_ts.strftime('%Y-%m-%d')}",
                "Total Records": len(sub),
                "Gross Orders ($)": round(
                    float(sub.loc[orders_mask, "Estimated Tax Value"].sum()), 2
                ),
                "Cancellations ($)": round(
                    float(sub.loc[cancels_mask, "Estimated Tax Value"].sum()), 2
                ),
                "Net Tax Value ($)": round(
                    float(sub["Estimated Tax Value"].sum()), 2
                ),
            }
        )

    # Full Year / Cumulative Total across all valid rows in the report
    min_d = df["Order Date"].min().strftime("%Y-%m-%d")
    max_d = df["Order Date"].max().strftime("%Y-%m-%d")
    orders_mask_all = df["Estimated Tax Value"] > 0
    cancels_mask_all = df["Estimated Tax Value"] < 0

    summary_rows.append(
        {
            "Timeframe": "Full Year Total",
            "Date Range": f"{min_d} to {max_d}",
            "Total Records": len(df),
            "Gross Orders ($)": round(
                float(df.loc[orders_mask_all, "Estimated Tax Value"].sum()), 2
            ),
            "Cancellations ($)": round(
                float(df.loc[cancels_mask_all, "Estimated Tax Value"].sum()), 2
            ),
            "Net Tax Value ($)": round(float(df["Estimated Tax Value"].sum()), 2),
        }
    )

    return pd.DataFrame(summary_rows)


# ---------------------------------------------------------------------------
# Chart Builder
# ---------------------------------------------------------------------------


def build_decay_chart(
    proj_df: pd.DataFrame,
    threshold: float,
    crossing_info: dict[str, Any] | None,
    window_days: int,
) -> go.Figure:
    """Create an interactive Plotly line chart of the projected decay curve."""
    fig = go.Figure()

    # Main Projected Decay Curve (solid line)
    fig.add_trace(
        go.Scatter(
            x=proj_df["Date"],
            y=proj_df["Projected Total"],
            mode="lines+markers",
            name=f"Projected {window_days}-Day Rolling Total",
            line={"color": "#2563EB", "width": 3},
            marker={"size": 6, "color": "#1D4ED8"},
            customdata=proj_df[
                ["Day Offset", "Variance to Threshold", "Daily Roll-Off ($)", "Status"]
            ].values,
            hovertemplate=(
                "<b>%{x|%B %d, %Y}</b> (Day +%{customdata[0]})<br>"
                "Projected Total: <b>$%{y:,.2f}</b><br>"
                "Daily Roll-Off: $%{customdata[2]:,.2f}<br>"
                "Variance to Threshold: $%{customdata[1]:+,.2f}<br>"
                "Status: %{customdata[3]}"
                "<extra></extra>"
            ),
        )
    )

    # Horizontal Threshold Line (dashed red line)
    fig.add_trace(
        go.Scatter(
            x=[proj_df["Date"].min(), proj_df["Date"].max()],
            y=[threshold, threshold],
            mode="lines",
            name=f"Threshold (${threshold:,.2f})",
            line={"color": "#DC2626", "width": 2.5, "dash": "dash"},
            hovertemplate=f"Threshold Target: <b>${threshold:,.2f}</b><extra></extra>",
        )
    )

    # Highlight Crossing Date if detected within the projection horizon
    if crossing_info is not None:
        c_date = crossing_info["date"]
        c_val = crossing_info["projected_total"]
        label_prefix = (
            "Already Under Threshold"
            if crossing_info["already_below"]
            else "Threshold Crossing"
        )

        fig.add_trace(
            go.Scatter(
                x=[c_date],
                y=[c_val],
                mode="markers+text",
                name=label_prefix,
                marker={
                    "size": 14,
                    "color": "#16A34A",
                    "symbol": "star" if not crossing_info["already_below"] else "circle",
                    "line": {"color": "#FFFFFF", "width": 2},
                },
                text=[f"  {c_date.strftime('%b %d, %Y')} (${c_val:,.2f})"],
                textposition="bottom right",
                textfont={"color": "#15803D", "size": 12, "family": "Arial Black"},
                hovertemplate=(
                    f"<b>{label_prefix}</b><br>"
                    "Date: <b>%{x|%B %d, %Y}</b><br>"
                    "Projected Total: <b>$%{y:,.2f}</b>"
                    "<extra></extra>"
                ),
            )
        )

        fig.add_vline(
            x=c_date.to_pydatetime(),
            line_width=1.5,
            line_dash="dot",
            line_color="#16A34A",
        )

    fig.update_layout(
        title={
            "text": f"Projected {window_days}-Day Rolling ETV Decay Curve",
            "font": {"size": 18},
        },
        xaxis={
            "title": "Projection Date",
            "tickformat": "%b %d, %Y",
            "showgrid": True,
            "gridcolor": "rgba(148, 163, 184, 0.2)",
        },
        yaxis={
            "title": "Rolling ETV ($)",
            "tickprefix": "$",
            "tickformat": ",.2f",
            "rangemode": "tozero",
            "showgrid": True,
            "gridcolor": "rgba(148, 163, 184, 0.2)",
        },
        hovermode="x unified",
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "right",
            "x": 1.0,
        },
        margin={"l": 40, "r": 30, "t": 70, "b": 40},
        height=460,
    )

    return fig


# ---------------------------------------------------------------------------
# Sample Workbook Generator (for instant demo / testing if needed)
# ---------------------------------------------------------------------------


def build_demo_excel_bytes() -> bytes:
    """Create a realistic sample Amazon Vine Itemized Tax Report in memory."""
    buf = io.BytesIO()
    rows = [
        ["Amazon Vine Itemized Tax Report - Tax Year 2026", None, None, None, None, None, None],
        ["Account Holder: Sample Vine Voice", None, None, None, None, None, None],
        [None, None, None, None, None, None, None],
        [
            "Order Number",
            "ASIN",
            "Product Name",
            "Order Type",
            "Order Date",
            "Cancelled Date",
            "Estimated Tax Value",
        ],
        ["112-0000001-00001", "B00DEMO001", "4K OLED Portable Monitor", "ORDER", "2026-08-14", None, 429.99],
        ["112-0000001-00002", "B00DEMO002", "Commercial Espresso Machine", "ORDER", "2026-09-05", None, 899.00],
        ["112-0000001-00003", "B00DEMO003", "Ergonomic Mesh Office Chair", "ORDER", "2026-09-22", None, 540.00],
        ["112-0000001-00004", "B00DEMO004", "Smart Robot Vacuum & Mop", "ORDER", "2026-10-08", None, 780.00],
        ["112-0000001-00005", "B00DEMO005", "Wireless Mechanical Keyboard", "ORDER", "2026-10-21", None, 199.50],
        ["112-0000001-00006", "B00DEMO006", "Laser Home Theater Projector", "ORDER", "2026-10-25", None, 2199.99],
        ["112-0000001-00007", "B00DEMO007", "Portable Power Station 2000W", "ORDER", "2026-10-28", None, 1499.00],
        ["112-0000001-00008", "B00DEMO008", "Titanium Cookware 12pc Set", "ORDER", "2026-10-31", None, 766.09],
        ["112-0000001-00009", "B00DEMO009", "Studio Condenser Microphone", "ORDER", "2026-11-02", None, 349.99],
        ["112-0000001-00009", "B00DEMO009", "Studio Condenser Microphone", "CANCELLATION", "2026-11-02", "2026-11-03", -349.99],
        ["112-0000001-00010", "B00DEMO010", "Carbon Road Bike Wheelset", "ORDER", "2026-11-05", None, 1850.00],
        ["112-0000001-00011", "B00DEMO011", "Dual-Zone Wine Cooler Fridge", "ORDER", "2026-11-09", None, 1299.00],
        ["112-0000001-00012", "B00DEMO012", "Full-Frame Mirrorless Lens 35mm", "ORDER", "2026-11-12", None, 1498.00],
        ["112-0000001-00013", "B00DEMO013", "Smart Home Gym Cable System", "ORDER", "2026-11-16", None, 880.09],
        [None, None, None, None, None, None, None],
        ["Frequently Asked Questions (FAQ)", None, None, None, None, None, None],
        ["Q: Why does a cancelled order appear twice?", "A: Cancelled orders are offset with a negative Estimated Tax Value.", None, None, None, None, None],
        ["Note: This report is provided for informational purposes only.", None, None, None, None, None, None],
    ]
    demo_df = pd.DataFrame(rows)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        demo_df.to_excel(writer, index=False, header=False, sheet_name="Itemized Report")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Streamlit GUI Application
# ---------------------------------------------------------------------------


def main() -> None:
    icon_file = Path(__file__).resolve().parent / "icon.png"
    st.set_page_config(
        page_title="Vine ETV Analyzer",
        page_icon=str(icon_file) if icon_file.is_file() else "📉",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # Custom CSS for clean KPI scorecard styling
    st.markdown(
        """
        <style>
        div[data-testid="stMetric"] {
            background-color: rgba(248, 250, 252, 0.85);
            border: 1px solid rgba(203, 213, 225, 0.7);
            border-radius: 10px;
            padding: 14px 18px;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("📉 Vine ETV Analyzer")
    st.caption(
        "See your floating window of Amazon Vine items and project when older orders roll off "
        "below the rolling cap (e.g., $6,000 over 30 days) so you can acquire items again."
    )

    # -----------------------------------------------------------------------
    # Sidebar Controls
    # -----------------------------------------------------------------------
    with st.sidebar:
        st.header("1. Data Source")
        uploaded_file = st.file_uploader(
            "Upload Amazon Vine Itemized Report (.xlsx)",
            type=["xlsx"],
            help="Select your Amazon Vine Itemized Report Excel file (.xlsx).",
        )

        # Also check if there are any .xlsx files in the local directory for convenience
        local_xlsx_files = sorted(
            [str(p.name) for p in Path.cwd().glob("*.xlsx") if not p.name.startswith("~$")]
        )
        selected_local_file: str | None = None
        if uploaded_file is None and local_xlsx_files:
            use_local = st.checkbox("Or select an Excel file from workspace", value=False)
            if use_local:
                selected_local_file = st.selectbox("Workspace .xlsx file", local_xlsx_files)

        use_demo = st.checkbox(
            "Load built-in sample report",
            value=False,
            disabled=(uploaded_file is not None or selected_local_file is not None),
            help="Preview the dashboard using realistic sample Amazon Vine report data.",
        )

        st.divider()
        st.header("2. Projection Parameters")

        window_length = st.number_input(
            "Rolling Window Length (Days)",
            min_value=1,
            max_value=365,
            value=30,
            step=1,
            help="Number of trailing days included in the rolling window (default: 30 days).",
        )

        tax_threshold = st.number_input(
            "Tax Threshold ($)",
            min_value=0.0,
            max_value=1_000_000.0,
            value=6000.00,
            step=100.0,
            format="%.2f",
            help="Target dollar threshold to monitor (default: $6,000.00).",
        )

        sync_horizon = st.checkbox(
            "Match Projection Horizon to Window Length",
            value=True,
            help="When enabled, projects forward by the same number of days as the Rolling Window Length.",
        )

        if sync_horizon:
            projection_horizon = int(window_length)
            st.number_input(
                "Projection Horizon (Days)",
                min_value=1,
                max_value=365,
                value=int(window_length),
                step=1,
                disabled=True,
            )
        else:
                projection_horizon = int(
                    st.number_input(
                        "Projection Horizon (Days)",
                        min_value=1,
                        max_value=365,
                        value=int(window_length),
                        step=1,
                        help="Number of days into the future to project.",
                    )
                )

        base_date_container = st.container()
        st.divider()
        if st.button(
            "🚪 Quit Application",
            width="stretch",
            help="Shut down the local server and exit the application cleanly.",
        ):
            st.success("Shutting down server... You can now close this browser tab.")
            threading.Timer(0.4, lambda: os._exit(0)).start()
            st.stop()

        st.caption("© 2026 George Gilman · [MIT License](https://github.com/ggilman/vineprojection/blob/main/LICENSE)")

    # -----------------------------------------------------------------------
    # Determine Active File Input
    # -----------------------------------------------------------------------
    file_source: Any = None
    source_label = ""

    if uploaded_file is not None:
        file_source = uploaded_file
        source_label = uploaded_file.name
    elif selected_local_file is not None:
        file_source = str(Path.cwd() / selected_local_file)
        source_label = selected_local_file
    elif use_demo:
        file_source = io.BytesIO(build_demo_excel_bytes())
        source_label = "Sample_Vine_Tax_Report_2026.xlsx (Demo)"

    if file_source is None:
        st.info(
            "👈 **Upload your Amazon Vine Itemized Report (`.xlsx`)** in the sidebar "
            "(or toggle **Load built-in sample report**) to view your rolling ETV window and projection."
        )
        with st.expander("Expected Excel Columns & Format Details", expanded=True):
            st.markdown(
                """
                The parser automatically scans your `.xlsx` file for the itemized data table
                regardless of how many summary or header rows appear at the top of the sheet,
                and strips out FAQ/footer notes at the bottom.

                - **`Order Date`**: Date the Vine order was placed.
                - **`Estimated Tax Value`**: Numeric dollar value (negative for cancellations).
                - **`Order Type`**: e.g., `ORDER`, `CANCELLATION` (automatically inferred if omitted).
                - **`Cancelled Date`**: Date of cancellation (if present).
                """
            )
        return

    # -----------------------------------------------------------------------
    # Parse Workbook with Graceful Error Handling
    # -----------------------------------------------------------------------
    try:
        if hasattr(file_source, "getvalue"):
            df, meta = _load_vine_excel_bytes(file_source.getvalue())
        elif isinstance(file_source, (str, Path)):
            df, meta = _load_vine_excel_bytes(Path(file_source).read_bytes())
        else:
            df, meta = load_vine_excel(file_source)
    except ValueError as err:
        st.error(f"**Unable to parse Excel report:** {err}")
        return
    except (OSError, RuntimeError, KeyError) as exc:
        st.error(f"**Unexpected error while reading file:** {exc}")
        return

    # -----------------------------------------------------------------------
    # Base Date Selection (Sidebar extension once file is loaded)
    # -----------------------------------------------------------------------
    latest_order_date: pd.Timestamp = meta["max_order_date"]
    today_ts = pd.Timestamp(date.today()).normalize()

    with base_date_container:
        st.divider()
        st.header("3. Base Date Settings")
        base_date_mode = st.radio(
            "Base Date (Day 0)",
            options=[
                f"Today's Date ({today_ts.strftime('%Y-%m-%d')})",
                f"Latest Order Date ({latest_order_date.strftime('%Y-%m-%d')})",
                "Custom Date",
            ],
            index=0,
            help="Starting reference date (Day 0) for the current rolling window and future decay projection.",
        )

        if base_date_mode.startswith("Today's Date"):
            base_date = today_ts
        elif base_date_mode.startswith("Latest Order Date"):
            base_date = latest_order_date
        else:
            picked_date = st.date_input(
                "Select Custom Base Date",
                value=today_ts.date(),
            )
            base_date = pd.Timestamp(picked_date).normalize()

    # -----------------------------------------------------------------------
    # Run Core Calculations
    # -----------------------------------------------------------------------
    proj_df, crossing_info = generate_decay_projection(
        df=df,
        base_date=base_date,
        window_days=int(window_length),
        threshold=float(tax_threshold),
        horizon_days=int(projection_horizon),
        inclusive_start=False,
    )

    current_window_total = float(proj_df.iloc[0]["Projected Total"])
    total_annual_tax_value = round(float(df["Estimated Tax Value"].sum()), 2)
    timeframe_summary_df = compute_timeframe_summary(
        df=df,
        base_date=base_date,
        inclusive_start=False,
    )

    # Format Threshold Crossing KPI display
    if crossing_info is not None:
        c_date_str = crossing_info["date"].strftime("%B %d, %Y")
        c_val_str = f"${crossing_info['projected_total']:,.2f}"
        if crossing_info["already_below"]:
            crossing_headline = f"Currently Under ({c_val_str})"
            crossing_sub = f"As of {c_date_str} (Day 0)"
        else:
            crossing_headline = f"{c_date_str} — {c_val_str}"
            crossing_sub = f"In {crossing_info['day_offset']} day(s) from Base Date"
    else:
        # Check when it eventually crosses even beyond the chosen horizon
        _extended_df, extended_crossing = generate_decay_projection(
            df=df,
            base_date=base_date,
            window_days=int(window_length),
            threshold=float(tax_threshold),
            horizon_days=max(int(window_length), int(projection_horizon)),
            inclusive_start=False,
        )
        if extended_crossing is not None:
            c_date_str = extended_crossing["date"].strftime("%B %d, %Y")
            c_val_str = f"${extended_crossing['projected_total']:,.2f}"
            crossing_headline = f"{c_date_str} — {c_val_str}"
            crossing_sub = f"Outside {projection_horizon}d horizon (Day +{extended_crossing['day_offset']})"
        else:
            crossing_headline = "Not Reached in Horizon"
            crossing_sub = f"Remains >= ${tax_threshold:,.2f} through +{projection_horizon}d"

    # -----------------------------------------------------------------------
    # 1. Top Header & KPI Scorecards
    # -----------------------------------------------------------------------
    st.subheader("Key Performance Indicators")
    kpi1, kpi2, kpi3, kpi4 = st.columns([1, 1, 1.45, 1])

    variance_now = round(current_window_total - float(tax_threshold), 2)
    with kpi1:
        st.metric(
            label=f"Current {int(window_length)}-Day Window Total",
            value=f"${current_window_total:,.2f}",
            delta=f"{'+' if variance_now >= 0 else ''}${variance_now:,.2f} vs. Threshold",
            delta_color="inverse",
        )

    with kpi2:
        st.metric(
            label="Threshold Target",
            value=f"${float(tax_threshold):,.2f}",
            delta=f"Base Date: {base_date.strftime('%b %d, %Y')}",
            delta_color="off",
        )

    with kpi3:
        st.metric(
            label="Threshold Crossing Date",
            value=crossing_headline,
            delta=crossing_sub,
            delta_color="normal" if crossing_info is not None else "off",
        )

    with kpi4:
        st.metric(
            label="Total Annual Tax Value",
            value=f"${total_annual_tax_value:,.2f}",
            delta=f"{len(df):,} net line items",
            delta_color="off",
        )

    # Contextual Status Banner
    if crossing_info is not None:
        if crossing_info["already_below"]:
            st.success(
                f"✅ **Currently Under Threshold:** Your {int(window_length)}-day rolling total "
                f"as of **{base_date.strftime('%B %d, %Y')}** is **${current_window_total:,.2f}**, "
                f"which is already **${abs(variance_now):,.2f}** below the **${tax_threshold:,.2f}** threshold."
            )
        else:
            st.warning(
                f"⏳ **Threshold Crossing Detected:** Assuming no new orders, your {int(window_length)}-day "
                f"rolling total will drop below **${tax_threshold:,.2f}** on "
                f"**{crossing_info['date'].strftime('%B %d, %Y')}** (in **{crossing_info['day_offset']} days**) "
                f"at **${crossing_info['projected_total']:,.2f}**."
            )
    else:
        st.info(
            f"ℹ️ **Above Threshold Across Horizon:** The projected {int(window_length)}-day rolling total "
            f"does not drop below **${tax_threshold:,.2f}** within the **{projection_horizon}-day** projection window."
        )

    # -----------------------------------------------------------------------
    # 2. Main Chart Area
    # -----------------------------------------------------------------------
    st.subheader("Projected Decay Curve")
    chart_fig = build_decay_chart(
        proj_df=proj_df,
        threshold=float(tax_threshold),
        crossing_info=crossing_info,
        window_days=int(window_length),
    )
    st.plotly_chart(chart_fig, width="stretch")

    # -----------------------------------------------------------------------
    # 3. Summary & Daily Breakdown Tables
    # -----------------------------------------------------------------------
    col_left, col_right = st.columns([1, 1.15])

    with col_left:
        st.subheader("Timeframe Summary Table")
        st.caption(
            f"Rolling period totals calculated relative to Base Date ({base_date.strftime('%Y-%m-%d')})."
        )
        styled_summary = timeframe_summary_df.style.format(
            {
                "Total Records": "{:,}",
                "Gross Orders ($)": "${:,.2f}",
                "Cancellations ($)": "${:,.2f}",
                "Net Tax Value ($)": "${:,.2f}",
            }
        )
        st.dataframe(
            styled_summary,
            width="stretch",
            hide_index=True,
        )

    with col_right:
        st.subheader("Daily Projection Schedule")
        st.caption(
            f"Day-by-day decay from {base_date.strftime('%Y-%m-%d')} through "
            f"+{projection_horizon} days assuming $0.00 in new orders."
        )

        schedule_display_df = proj_df[
            ["Date", "Projected Total", "Variance to Threshold", "Status"]
        ].copy()
        schedule_display_df["Date"] = schedule_display_df["Date"].dt.strftime("%Y-%m-%d")

        def _highlight_status(row: pd.Series) -> list[str]:
            if row["Status"] == "Below Threshold":
                return [
                    "background-color: rgba(22, 163, 74, 0.12); color: #15803D; font-weight: 600;"
                ] * len(row)
            return [""] * len(row)

        styled_schedule = (
            schedule_display_df.style.format(
                {
                    "Projected Total": "${:,.2f}",
                    "Variance to Threshold": "${:+,.2f}",
                }
            )
            .apply(_highlight_status, axis=1)
        )

        st.dataframe(
            styled_schedule,
            width="stretch",
            hide_index=True,
            height=380,
        )

    # -----------------------------------------------------------------------
    # Expandable Parsed Report Data & Metadata Inspector
    # -----------------------------------------------------------------------
    with st.expander(
        f"📋 View Cleaned Itemized Transactions ({len(df):,} records from '{source_label}')",
        expanded=False,
    ):
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.markdown(f"**Worksheet:** `{meta['sheet_name']}`")
        m_col2.markdown(f"**Header Row:** Row `{meta['header_row_number']}`")
        m_col3.markdown(f"**Valid Order Rows:** `{meta['valid_rows']:,}`")
        m_col4.markdown(
            f"**Stripped Footer/Non-Data Rows:** `{meta['skipped_footer_rows'] + meta['malformed_rows']:,}`"
        )

        display_tx_df = df.copy()
        display_tx_df["Order Date"] = display_tx_df["Order Date"].dt.strftime("%Y-%m-%d")
        if "Cancelled Date" in display_tx_df.columns:
            display_tx_df["Cancelled Date"] = display_tx_df["Cancelled Date"].apply(
                lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) and x is not None else ""
            )

        st.dataframe(
            display_tx_df.style.format({"Estimated Tax Value": "${:,.2f}"}),
            width="stretch",
            hide_index=True,
        )

    st.divider()
    st.caption(
        "Vine ETV Analyzer · "
        "Copyright © 2026 George Gilman · "
        "[MIT License](https://github.com/ggilman/vineprojection/blob/main/LICENSE)"
    )


if __name__ == "__main__":
    main()
