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

Optional live Amazon Vine connection (cap status + auto-downloaded reports):
    pip install playwright && python -m playwright install chromium
"""

from __future__ import annotations

import io
import json
import math
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import warnings
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# Amazon's itemized reports are written without a default cell style; openpyxl
# warns about it on every load and then uses its own, which is harmless here.
warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style",
    category=UserWarning,
    module="openpyxl",
)

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
    "shipped_date": [
        "shipped date",
        "ship date",
        "shipping date",
        "shipment date",
        "date shipped",
    ],
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


# Which date places a row in the rolling window. "Ship" uses Shipped Date, falling
# back to Cancelled Date for cancellations (they never ship) and then Order Date.
DATE_BASIS_SHIP = "Ship Date"
DATE_BASIS_ORDER = "Order Date"
DATE_BASIS_OPTIONS = [DATE_BASIS_SHIP, DATE_BASIS_ORDER]
BASIS_COL = "Basis Date"


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

    s = re.sub(r"[$€£,\s]", "", s)
    # Only a leading or trailing minus ("-$45.00", "$-45.00", "45.00-") means negative;
    # a hyphen anywhere else leaves the value unparseable (malformed row).
    if s.startswith("-") or s.endswith("-"):
        is_negative = True
        s = s[1:] if s.startswith("-") else s[:-1]
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
    excel_file.close()

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
    shipped_date_idx = col_map.get("shipped_date")
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

        parsed_shipped_date = None
        if shipped_date_idx is not None:
            parsed_shipped_date = _parse_datetime(row.iloc[shipped_date_idx])

        record: dict[str, Any] = {
            "Order Date": parsed_order_date,
            "Shipped Date": parsed_shipped_date,
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
        "has_shipped_date": shipped_date_idx is not None,
    }

    return cleaned_df, metadata


def apply_date_basis(df: pd.DataFrame, basis: str) -> pd.DataFrame:
    """
    Return a copy of `df` with a `Basis Date` column, the date every calculation
    uses, sorted by it. Ship basis: Shipped Date, else Cancelled Date (cancellations
    never ship), else Order Date. Order basis: Order Date.
    """
    out = df.copy()
    if basis == DATE_BASIS_SHIP:
        basis_dates = pd.to_datetime(out["Shipped Date"], errors="coerce")
        if "Cancelled Date" in out.columns:
            basis_dates = basis_dates.fillna(pd.to_datetime(out["Cancelled Date"], errors="coerce"))
        basis_dates = basis_dates.fillna(out["Order Date"])
    else:
        basis_dates = out["Order Date"]
    out[BASIS_COL] = pd.to_datetime(basis_dates).dt.normalize()
    return out.sort_values(by=BASIS_COL, ascending=True, kind="stable").reset_index(drop=True)


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
    window ending on `target_date`, placing each row by its `Basis Date`.

    By specification:
      - For Day D, items dated in `(D - Window Length, D]` are summed.
      - When `inclusive_start` is False (default), interval is `(D - window_days, D]`,
        which spans `window_days` calendar days ending on D.
      - When `inclusive_start` is True, interval is `[D - window_days, D]`.
    """
    target_ts = pd.Timestamp(target_date).normalize()
    cutoff_ts = target_ts - pd.Timedelta(days=int(window_days))

    if inclusive_start:
        mask = (df[BASIS_COL] >= cutoff_ts) & (df[BASIS_COL] <= target_ts)
    else:
        mask = (df[BASIS_COL] > cutoff_ts) & (df[BASIS_COL] <= target_ts)

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
    year_df: pd.DataFrame | None = None,
    max_days: int | None = None,
) -> pd.DataFrame:
    """
    Build the Timeframe Summary Table for standard rolling periods:
    7 Days, 30 Days, 60 Days, 90 Days, and Full Year Total.
    `year_df` (default: `df`) holds the rows the Full Year Total covers, so a
    prior year's report loaded for the rolling window doesn't count toward it.
    With `max_days` (data covering only that many days back), longer periods and
    the Full Year row are left out rather than shown as partial totals.
    """
    base_ts = pd.Timestamp(base_date).normalize()
    periods = [
        ("7 Days", 7),
        ("30 Days", 30),
        ("60 Days", 60),
        ("90 Days", 90),
    ]
    if max_days is not None:
        periods = [p for p in periods if p[1] <= max_days]

    summary_rows: list[dict[str, Any]] = []
    for label, days in periods:
        cutoff = base_ts - pd.Timedelta(days=days)
        if inclusive_start:
            sub = df[(df[BASIS_COL] >= cutoff) & (df[BASIS_COL] <= base_ts)]
            start_display = cutoff.strftime("%Y-%m-%d")
        else:
            sub = df[(df[BASIS_COL] > cutoff) & (df[BASIS_COL] <= base_ts)]
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

    if max_days is not None:
        return pd.DataFrame(summary_rows)

    # Full Year / Cumulative Total across all valid rows in the report
    df = df if year_df is None else year_df
    min_d = df[BASIS_COL].min().strftime("%Y-%m-%d")
    max_d = df[BASIS_COL].max().strftime("%Y-%m-%d")
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
# Amazon Vine Live Connection (Opt-In Feature)
# ---------------------------------------------------------------------------

VINE_ITEMS_URL = "https://www.amazon.com/vine/vine-items"
VINE_ACCOUNT_URL = "https://www.amazon.com/vine/account"
VINE_ORDERS_URL = "https://www.amazon.com/vine/orders"
VINE_ORDERS_FILENAME = "vine_orders.json"
# The rolling window the orders pages are read back over, a page-count safety
# cap (10 orders per page), and the pause between page reads.
ORDERS_LOOKBACK_DAYS = 30
ORDERS_MAX_PAGES = 80
ORDERS_PAGE_DELAY_MS = 700
VINE_REPORTS_DIRNAME = "vine_reports"
VINE_CONNECTION_FILENAME = ".vine_connection.json"

CAP_OPEN = "OPEN"
CAP_CAPPED = "CAPPED"
CAP_NEEDS_LOGIN = "NEEDS_LOGIN"
CAP_BLOCKED = "BLOCKED"
CAP_UNKNOWN = "UNKNOWN"

CAP_MESSAGE_PATTERNS: tuple[str, ...] = (
    r"claiming\s+will\s+resume\s+soon",
    r"vine\s+limits\s+claiming\s+activity",
)
# NOTE: do NOT add "to discuss automated access to amazon.com" here — that string
# appears in the standard footer of every regular Amazon page and would make
# logged-in pages misclassify as BLOCKED.
BLOCKED_MESSAGE_PATTERNS: tuple[str, ...] = (
    r"enter the characters you see in this image",
    r"api-services-support@amazon\.com",
)
LOGIN_WALL_URL_MARKERS: tuple[str, ...] = (
    "/ap/signin",
    "/ap/mfa",
    "/ap/cvh",
    "/ap/authentication",
    "/signin",
)
CAPTCHA_URL_MARKERS: tuple[str, ...] = ("/errors/validateCaptcha",)
# at-main = actual authenticated-session token. Do NOT trust lc-main (it only
# marks a "remembered account" and survives on the sign-in page, which would
# make the interactive login window close before the user actually signs in).
# Also do NOT trust session-id / ubid-main: Amazon issues those during the
# /ap/signin handshake itself, even when sign-in never completes.
SIGNED_IN_COOKIE_NAMES: frozenset[str] = frozenset({"at-main"})


def _reports_dir() -> Path:
    # Beside the app, not the launch directory, so reports are found however it's started.
    return Path(__file__).resolve().parent / VINE_REPORTS_DIRNAME


def _app_data_root() -> Path:
    home = Path.home()
    if os.name == "nt":
        local = os.getenv("LOCALAPPDATA")
        base = Path(local) if local else home / "AppData" / "Local"
    elif sys.platform == "darwin":
        base = home / "Library" / "Application Support"
    else:
        base = Path(os.getenv("XDG_DATA_HOME", str(home / ".local" / "share")))
    return base / "VineProjection"


def _session_root() -> Path:
    return _app_data_root() / "vine-session"


def _connection_marker_path() -> Path:
    return _session_root() / VINE_CONNECTION_FILENAME


def session_ready() -> bool:
    return _connection_marker_path().is_file()


def connection_meta() -> dict[str, Any]:
    try:
        data = json.loads(_connection_marker_path().read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_connection_meta(channel: str | None, source: str) -> None:
    payload = {
        "channel": channel,
        "source": source,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        _connection_marker_path().write_text(
            json.dumps(payload, indent=2) + "\n", encoding="utf-8"
        )
    except OSError:
        pass


def disconnect_session() -> None:
    root = _session_root()
    if root.exists():
        shutil.rmtree(root, ignore_errors=True)


def _browser_user_data_roots() -> list[dict[str, str]]:
    home = Path.home()
    candidates: list[dict[str, str]] = []
    if os.name == "nt":
        local = Path(os.getenv("LOCALAPPDATA", str(home / "AppData" / "Local")))
        candidates.append(
            {
                "browser": "Chrome",
                "channel": "chrome",
                "root": str(local / "Google" / "Chrome" / "User Data"),
            }
        )
        candidates.append(
            {
                "browser": "Edge",
                "channel": "msedge",
                "root": str(local / "Microsoft" / "Edge" / "User Data"),
            }
        )
    elif sys.platform == "darwin":
        support = home / "Library" / "Application Support"
        candidates.append(
            {"browser": "Chrome", "channel": "chrome", "root": str(support / "Google" / "Chrome")}
        )
        candidates.append(
            {"browser": "Edge", "channel": "msedge", "root": str(support / "Microsoft" / "Edge")}
        )
    else:
        cfg = Path(os.getenv("XDG_CONFIG_HOME", str(home / ".config")))
        candidates.append(
            {"browser": "Chrome", "channel": "chrome", "root": str(cfg / "google-chrome")}
        )
        candidates.append(
            {"browser": "Edge", "channel": "msedge", "root": str(cfg / "microsoft-edge")}
        )
    return [c for c in candidates if (Path(c["root"]) / "Local State").is_file()]


def _cookies_have_amazon(cookies_path: Path) -> bool | None:
    try:
        uri = f"file:{cookies_path.as_posix()}?mode=ro"
        conn = sqlite3.connect(uri, uri=True)
        try:
            # Chrome/Edge renamed the cookies host column to host_key.
            for column in ("host_key", "host"):
                try:
                    row = conn.execute(
                        f"SELECT 1 FROM cookies WHERE {column} LIKE '%amazon%' LIMIT 1"
                    ).fetchone()
                    return row is not None
                except sqlite3.OperationalError:
                    continue
            return None
        finally:
            conn.close()
    except sqlite3.Error:
        try:
            return b".amazon." in cookies_path.read_bytes()
        except OSError:
            return None


def find_browser_profiles() -> list[dict[str, Any]]:
    profiles: list[dict[str, Any]] = []
    for entry in _browser_user_data_roots():
        root = Path(entry["root"])
        dirs: list[tuple[str, str]] = []
        try:
            state = json.loads((root / "Local State").read_text(encoding="utf-8"))
            info_cache = state.get("info_cache") or {}
            if isinstance(info_cache, dict):
                for dir_key, info in info_cache.items():
                    if not isinstance(info, dict):
                        continue
                    rel_dir = str(info.get("dir") or dir_key)
                    display_name = str(info.get("name") or dir_key)
                    dirs.append((rel_dir, display_name))
        except (OSError, ValueError):
            pass
        if not dirs:
            dirs = [("Default", "Default")]
        seen: set[str] = set()
        for rel_dir, display_name in dirs:
            if not rel_dir or rel_dir in seen:
                continue
            seen.add(rel_dir)
            cookies_rel: str | None = None
            for candidate_rel in (f"{rel_dir}/Network/Cookies", f"{rel_dir}/Cookies"):
                if (root / candidate_rel).is_file():
                    cookies_rel = candidate_rel
                    break
            if cookies_rel is None:
                continue
            profiles.append(
                {
                    "browser": entry["browser"],
                    "channel": entry["channel"],
                    "root": str(root),
                    "profile_dir": rel_dir,
                    "profile_name": display_name,
                    "cookies_rel": cookies_rel,
                    "has_amazon": _cookies_have_amazon(root / cookies_rel),
                }
            )
    rank = {True: 0, None: 1, False: 2}
    profiles.sort(
        key=lambda p: (rank.get(p["has_amazon"], 1), p["browser"], p["profile_name"])
    )
    return profiles


def _copy_browser_profile(profile: dict[str, Any]) -> None:
    root = Path(profile["root"])
    session = _session_root()
    tmp = session.with_name(session.name + ".tmp")
    try:
        if tmp.exists():
            shutil.rmtree(tmp, ignore_errors=True)
        default_dir = tmp / "Default"
        default_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(root / "Local State", tmp / "Local State")
        cookies_rel = Path(str(profile.get("cookies_rel") or "Default/Cookies"))
        # cookies_rel starts with the source profile's folder ("Default", "Profile 1");
        # in the session that profile becomes "Default", so keep only the part inside it.
        inside_profile = Path(*cookies_rel.parts[1:]) if len(cookies_rel.parts) > 1 else cookies_rel
        _copy_cookie_trio(root / cookies_rel, (default_dir / inside_profile).parent)
        # Only discard the existing session once the fresh copy fully succeeded.
        if session.exists():
            shutil.rmtree(session, ignore_errors=True)
        tmp.replace(session)
    except OSError:
        shutil.rmtree(tmp, ignore_errors=True)
        raise
    _save_connection_meta(
        str(profile.get("channel") or ""),
        f"{profile.get('browser', 'browser')}:{profile.get('profile_name', 'profile')}",
    )


def _kill_session_browser() -> None:
    """Kill only browser processes that were started against OUR vine-session
    profile (e.g. a sign-in window that refused to close), so a subsequent
    launch of the same profile is not blocked by the singleton lock."""
    if sys.platform != "win32":
        return
    ps_script = (
        "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe' OR Name='msedge.exe'\" "
        "| Where-Object { $_.CommandLine -like '*vine-session*' } "
        "| ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            timeout=20,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        pass


def _launch_vine_context(pw: Any, headless: bool = False) -> tuple[Any, str | None]:
    from playwright.sync_api import Error as PlaywrightError

    session = _session_root()
    session.mkdir(parents=True, exist_ok=True)
    meta_channel = connection_meta().get("channel") or None
    # Prefer a real installed browser (Chrome, then Edge) over the bundled
    # Chromium build — real browsers have far better standing with Amazon.
    channels: list[str | None] = []
    for ch in (meta_channel, "chrome", "msedge"):
        if ch is not None and ch not in channels:
            channels.append(ch)
    channels.append(None)
    errors: list[str] = []
    for ch in channels:
        # Playwright disables the Chromium sandbox by default, which makes Chrome
        # show the "unsupported command-line flag: --no-sandbox" warning infobar in
        # our visible sign-in window. Enable the sandbox first; fall back only if
        # this machine cannot launch sandboxed.
        launch_args = [
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-blink-features=AutomationControlled",
            "--hide-crash-restore-bubble",
        ]
        # Amazon rejects both headless sessions and off-screen/minimized
        # windows as bot behaviour (verified empirically: a valid, fully
        # signed-in session still bounces to /ap/signin unless the window is
        # normal and visible). Recurring checks therefore run headful —
        # expect a Chrome window to briefly appear.
        for sandbox in (True, False):
            for attempt in (1, 2):
                try:
                    context = pw.chromium.launch_persistent_context(
                        user_data_dir=str(session),
                        channel=ch,
                        chromium_sandbox=sandbox,
                        headless=headless,
                        timeout=25000,
                        locale="en-US",
                        args=launch_args,
                    )
                    return context, ch
                except PlaywrightError as exc:
                    message = str(exc)
                    errors.append(f"{ch or 'chromium'}: {message.splitlines()[0]}")
                    lowered = message.lower()
                    locky = any(
                        s in lowered
                        for s in ("already running", "singleton", "user data dir", "in use")
                    )
                    if attempt == 1 and locky:
                        _kill_session_browser()
                        time.sleep(4.0)
                        continue
                    break
    raise ConnectionError(
        "Unable to launch a browser engine for the Vine live check. " + " | ".join(errors)
    )


def classify_vine_page(html: str, url: str, visible_text: str | None = None) -> tuple[str, str]:
    """`visible_text` is the page's rendered text. The cap messages are matched only
    there when given: Vine pages carry hidden alert templates with the same wording,
    so matching the raw HTML would report CAPPED while claiming is open."""
    if any(marker in url for marker in LOGIN_WALL_URL_MARKERS) or "/ap/" in url:
        return CAP_NEEDS_LOGIN, url
    if any(marker in url for marker in CAPTCHA_URL_MARKERS):
        return CAP_BLOCKED, "Amazon captcha challenge"
    lowered = html.casefold()
    for pattern in BLOCKED_MESSAGE_PATTERNS:
        match = re.search(pattern, lowered)
        if match:
            return CAP_BLOCKED, match.group(0)
    cap_source = html if visible_text is None else visible_text
    for pattern in CAP_MESSAGE_PATTERNS:
        match = re.search(pattern, cap_source, re.IGNORECASE)
        if match:
            return CAP_CAPPED, match.group(0)
    if url.startswith(VINE_ITEMS_URL) or "amazon.com/vine" in url:
        return CAP_OPEN, ""
    return CAP_UNKNOWN, url


def _orders_cache_path() -> Path:
    return _reports_dir() / VINE_ORDERS_FILENAME


# Runs inside the Vine orders page: fetch one page of /vine/orders with the
# page's own session and read each row's order timestamp (epoch ms; the visible
# date is filled in by script, so it is empty in fetched HTML) and Fair market value.
_ORDERS_PAGE_JS = """
async (pageNo) => {
  const resp = await fetch('/vine/orders?page=' + pageNo, {credentials: 'include'});
  const finalUrl = resp.url || '';
  const html = await resp.text();
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const rows = [];
  for (const tr of doc.querySelectorAll('table.vvp-orders-table tr.vvp-orders-table--row')) {
    const tsCell = tr.querySelector('[data-order-timestamp]');
    const cells = tr.querySelectorAll('td');
    const link = tr.querySelector('a[href*="/dp/"]');
    const asin = link ? ((link.getAttribute('href') || '').match(/\\/dp\\/([A-Z0-9]{10})/) || [])[1] || '' : '';
    const name = tr.querySelector('.a-truncate-full');
    rows.push({
      ts: tsCell ? Number(tsCell.getAttribute('data-order-timestamp')) : null,
      fmv: cells.length > 3 ? cells[3].textContent.trim() : '',
      asin: asin,
      name: name ? name.textContent.trim() : (cells.length > 1 ? cells[1].textContent.trim() : ''),
    });
  }
  return {status: resp.status, url: finalUrl, rows: rows};
}
"""


def fetch_vine_orders(page: Any, lookback_days: int = ORDERS_LOOKBACK_DAYS) -> dict[str, Any]:
    """
    Walk the Vine "Orders" pages (newest first) and collect every order placed in
    the last `lookback_days` days with its Fair market value. Stops at the first
    page whose oldest order is older than that. Saves the result to the orders
    cache file and returns a summary of what was fetched.
    """
    cutoff = pd.Timestamp(date.today()) - pd.Timedelta(days=int(lookback_days))
    page.goto(VINE_ORDERS_URL + "?page=1", wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(1500)
    status, _ = classify_live_page(page)
    if status == CAP_NEEDS_LOGIN:
        raise ConnectionError("Amazon redirected the Vine orders page back to sign-in.")
    if status == CAP_BLOCKED:
        raise ConnectionError("Amazon showed a bot check on the Vine orders page.")

    orders: list[dict[str, Any]] = []
    pages_read = 0
    reached_cutoff = False
    for page_no in range(1, ORDERS_MAX_PAGES + 1):
        if page_no > 1:
            page.wait_for_timeout(ORDERS_PAGE_DELAY_MS)  # go easy on Amazon
        data = page.evaluate(_ORDERS_PAGE_JS, page_no)
        final_url = str(data.get("url") or "")
        if any(marker in final_url for marker in LOGIN_WALL_URL_MARKERS) or "/ap/" in final_url:
            raise ConnectionError("Amazon asked for sign-in while reading the Vine orders pages.")
        rows = data.get("rows") or []
        pages_read += 1
        if not rows:
            reached_cutoff = True  # past the last page
            break
        oldest: pd.Timestamp | None = None
        for row in rows:
            if not row.get("ts"):
                continue
            # Local calendar date of the order, as the page itself shows it.
            order_dt = pd.Timestamp(datetime.fromtimestamp(float(row["ts"]) / 1000.0))
            oldest = order_dt if oldest is None else min(oldest, order_dt)
            if order_dt.normalize() <= cutoff:
                continue
            try:
                fmv = _parse_currency(row.get("fmv"))
            except ValueError:
                continue
            orders.append(
                {
                    "order_date": order_dt.isoformat(),
                    "fmv": round(fmv, 2),
                    "asin": str(row.get("asin") or ""),
                    "name": str(row.get("name") or ""),
                }
            )
        if oldest is not None and oldest.normalize() <= cutoff:
            reached_cutoff = True
            break

    payload = {
        "fetched_at": datetime.now().isoformat(timespec="seconds"),
        "lookback_days": int(lookback_days),
        "pages_read": pages_read,
        "complete": reached_cutoff,
        "orders": orders,
    }
    cache = _orders_cache_path()
    cache.parent.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps(payload, indent=1), encoding="utf-8")
    return {
        "path": str(cache),
        "fetched_at": payload["fetched_at"],
        "count": len(orders),
        "total_fmv": round(sum(o["fmv"] for o in orders), 2),
        "pages_read": pages_read,
        "complete": reached_cutoff,
    }


def load_vine_orders_cache() -> tuple[pd.DataFrame, dict[str, Any]] | None:
    """
    Read the orders saved by `fetch_vine_orders` into the same shape that
    `load_vine_excel` returns (Fair market value goes in the value column).
    Returns None when nothing has been fetched yet.
    """
    cache = _orders_cache_path()
    if not cache.is_file():
        return None
    try:
        payload = json.loads(cache.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    records = [
        {
            "Order Date": pd.Timestamp(o["order_date"]),
            "Shipped Date": None,
            "Estimated Tax Value": float(o["fmv"]),
            "Order Type": "ORDER",
            "Cancelled Date": None,
            "ASIN": o.get("asin", ""),
            "Product Name": o.get("name", ""),
        }
        for o in payload.get("orders", [])
        if o.get("order_date") is not None and o.get("fmv") is not None
    ]
    columns = [
        "Order Date", "Shipped Date", "Estimated Tax Value", "Order Type",
        "Cancelled Date", "ASIN", "Product Name",
    ]
    df = (
        pd.DataFrame(records, columns=columns)
        .sort_values(by="Order Date", ascending=True)
        .reset_index(drop=True)
    )
    meta = {
        "sheet_name": "Vine Orders page",
        "header_row_number": 0,
        "detected_headers": ["Product", "Order date", "Fair market value"],
        "valid_rows": len(df),
        "skipped_footer_rows": 0,
        "malformed_rows": 0,
        "min_order_date": df["Order Date"].min() if len(df) else None,
        "max_order_date": df["Order Date"].max() if len(df) else None,
        "has_shipped_date": False,
        "fetched_at": payload.get("fetched_at"),
        "lookback_days": int(payload.get("lookback_days") or ORDERS_LOOKBACK_DAYS),
        "complete": bool(payload.get("complete")),
        "pages_read": int(payload.get("pages_read") or 0),
    }
    return df, meta


def _cookies_db_path() -> Path:
    session = _session_root()
    primary = session / "Default" / "Network" / "Cookies"
    if primary.is_file():
        return primary
    return session / "Default" / "Cookies"


def _copy_cookie_trio(src: Path, dest_parent: Path) -> None:
    dest_parent.mkdir(parents=True, exist_ok=True)
    for suffix in ("", "-wal", "-shm"):
        candidate = src.with_name(src.name + suffix)
        if candidate.is_file():
            shutil.copy2(candidate, dest_parent / (src.name + suffix))


def _cookie_db_has_signed_in(db_path: Path) -> bool:
    if not db_path.is_file():
        return False
    tmp = _app_data_root() / f"_cookieprobe_{os.getpid()}"
    shutil.rmtree(tmp, ignore_errors=True)
    try:
        _copy_cookie_trio(db_path, tmp)
        con = sqlite3.connect(str(tmp / db_path.name))
        try:
            rows = con.execute("select name from cookies").fetchall()
        finally:
            con.close()
    except (OSError, sqlite3.Error):
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    names = {str(row[0]) for row in rows}
    return bool(names & SIGNED_IN_COOKIE_NAMES)


def _cookie_backup_dir() -> Path:
    return _session_root() / "cookies_backup"


def _save_cookie_backup() -> None:
    """Snapshot the (authenticated) cookie DB after a successful live check.
    Amazon occasionally wipes the Vine session after flagging an automated
    visit; the backup lets the next run self-heal without another sign-in."""
    src = _cookies_db_path()
    if not src.is_file():
        return
    try:
        _copy_cookie_trio(src, _cookie_backup_dir())
    except OSError:
        pass


def _restore_cookie_backup() -> bool:
    """Put the backed-up cookie DB back if the live one lost the authenticated
    cookies but the backup still has them. Only runs while no browser holds
    the profile."""
    backup_main = _cookie_backup_dir() / "Cookies"
    live = _cookies_db_path()
    if not backup_main.is_file():
        return False
    if _cookie_db_has_signed_in(live):
        return False
    if not _cookie_db_has_signed_in(backup_main):
        return False
    try:
        for suffix in ("", "-wal", "-shm"):
            stale = live.with_name(live.name + suffix)
            if stale.is_file():
                stale.unlink()
        _copy_cookie_trio(backup_main, live.parent)
    except OSError:
        return False
    return True


def _session_has_cookies() -> bool:
    default_dir = _session_root() / "Default"
    return (default_dir / "Network" / "Cookies").is_file() or (default_dir / "Cookies").is_file()


@st.cache_resource(show_spinner=False)
def _browser_lock() -> threading.Lock:
    """One lock for the whole Streamlit server process. cache_resource survives
    script reruns and is shared by every browser tab, so only one flow at a time
    drives (or copies over) the shared vine-session profile. Without it, a second
    tab's automatic check would hit the profile lock and kill the first tab's
    open sign-in window via _kill_session_browser()."""
    return threading.Lock()


BROWSER_BUSY_MESSAGE = (
    "Another Amazon Vine check or sign-in window is already running. "
    "Finish or close it, then try again."
)
SIGN_IN_TIMEOUT_SECONDS = 900.0
SIGN_IN_POLL_SECONDS = 1.5

LIVE_STATUS_ERRORS: dict[str, str] = {
    CAP_NEEDS_LOGIN: "Amazon does not recognize the saved session — sign-in required.",
    CAP_BLOCKED: "Amazon presented a security/bot check — use 'Sign in' to retry interactively.",
}


def _new_live_result() -> dict[str, Any]:
    return {
        "connected": False,
        "cap_status": None,
        "orders": None,
        "orders_error": None,
        "error": None,
        "channel": None,
        "checked_at": datetime.now().isoformat(timespec="seconds"),
    }


def _ensure_session_profile() -> str | None:
    """
    Make sure a reusable browser session profile exists. Returns None when ready,
    otherwise a human-readable explanation of why silent reuse is unavailable.

    An existing session that holds an authenticated cookie is never overwritten:
    it may be a fresh interactive sign-in whose connection marker was not written
    yet (for example the user closed the window themselves).
    """
    if session_ready():
        return None
    live_db = _cookies_db_path()
    if _cookie_db_has_signed_in(live_db):
        return None
    last_error = ""
    for profile in find_browser_profiles():
        if profile.get("has_amazon") is False:
            continue  # nothing worth copying from this profile
        try:
            _copy_browser_profile(profile)
            return None
        except OSError as exc:
            last_error = str(exc)
    if _session_has_cookies():
        # A partial session (e.g. from an unfinished sign-in) is still usable;
        # the live check will report NEEDS_LOGIN if it is not signed in.
        return None
    if not last_error:
        return (
            "No Chrome/Edge profile with Amazon cookies was found for silent session reuse. "
            "Use the sign-in option once to connect."
        )
    lowered = last_error.lower()
    if "permission" in lowered or "being used by another process" in lowered or "errno 13" in lowered:
        return (
            "Chrome/Edge is running and protects its cookie file while open. "
            "Fully close the browser and press Refresh (or Connect) for silent reuse, "
            "or use the sign-in option once."
        )
    return f"Could not copy the browser profile for silent reuse ({last_error})."


def _visible_text(page: Any) -> str | None:
    from playwright.sync_api import Error as PlaywrightError

    try:
        # innerText leaves out elements hidden with CSS (display:none, etc.).
        return page.inner_text("body", timeout=10000)
    except PlaywrightError:
        return None


def classify_live_page(page: Any) -> tuple[str, str]:
    return classify_vine_page(page.content(), page.url, _visible_text(page))


def _visit_and_classify_items(page: Any) -> tuple[str, str]:
    page.goto(VINE_ITEMS_URL, wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(2500)
    status, detail = classify_live_page(page)
    if status == CAP_NEEDS_LOGIN:
        # Amazon's silent SSO re-auth can briefly flash the /ap/signin URL
        # before landing back on the Vine page without any user action.
        # Re-read the same page a couple of times; do not hammer Amazon with
        # extra navigations (that itself looks like bot traffic).
        for _ in range(2):
            page.wait_for_timeout(2500)
            status, detail = classify_live_page(page)
            if status != CAP_NEEDS_LOGIN:
                break
    return status, detail


def _record_vine_status(
    context: Any,
    page: Any,
    status: str,
    detail: str,
    result: dict[str, Any],
    used_channel: str | None,
    save_report: bool,
) -> bool:
    """Store a classified Vine page in `result`; on a real Vine page also read the
    recent orders (Fair market value) and persist the connection marker. Returns connected."""
    from playwright.sync_api import Error as PlaywrightError

    result["cap_status"] = status
    result["channel"] = used_channel
    if status not in (CAP_OPEN, CAP_CAPPED):
        result["error"] = LIVE_STATUS_ERRORS.get(
            status, f"Unexpected Vine page state (URL: {detail})."
        )
        return False
    result["connected"] = True
    result["error"] = None
    if save_report:
        try:
            result["orders"] = fetch_vine_orders(page)
        except (ConnectionError, PlaywrightError, OSError, ValueError) as exc:
            result["orders_error"] = str(exc)[:300]
    _save_connection_meta(used_channel or "", connection_meta().get("source", "manual-login"))
    return True


def _run_live_check_locked(save_report: bool, bootstrap: bool = True) -> dict[str, Any]:
    """Body of run_live_check; the caller must hold _browser_lock(). With
    `bootstrap` False the existing session profile is used as it is and never
    replaced by a copy of the user's browser profile."""
    result = _new_live_result()
    bootstrap_note = _ensure_session_profile() if bootstrap else None
    if bootstrap_note is not None:
        result["cap_status"] = CAP_NEEDS_LOGIN
        result["error"] = bootstrap_note
        return result
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import sync_playwright

    connected = False
    with sync_playwright() as pw:
        context = None
        try:
            context, used_channel = _launch_vine_context(pw)
            page = context.pages[0] if context.pages else context.new_page()
            status, detail = _visit_and_classify_items(page)
            if status == CAP_NEEDS_LOGIN and _cookie_db_has_signed_in(
                _cookie_backup_dir() / "Cookies"
            ):
                # Amazon may have wiped the live cookies since the last good
                # check. Close the browser FIRST (the cookie DB can only be
                # swapped while nothing holds the profile), restore the
                # authenticated snapshot, and try once more.
                try:
                    context.close()
                except PlaywrightError:
                    pass
                context = None
                if _restore_cookie_backup():
                    context, used_channel = _launch_vine_context(pw)
                    page = context.pages[0] if context.pages else context.new_page()
                    status, detail = _visit_and_classify_items(page)
            connected = _record_vine_status(
                context, page, status, detail, result, used_channel, save_report
            )
        except (ConnectionError, PlaywrightError) as exc:
            result["error"] = str(exc)[:400]
        finally:
            if context is not None:
                try:
                    context.close()
                except PlaywrightError:
                    pass
    # Only back up cookies that were just proven to work.
    if connected:
        _save_cookie_backup()
    return result


def _playwright_missing_result() -> dict[str, Any] | None:
    try:
        import playwright.sync_api  # noqa: F401
    except ImportError:
        result = _new_live_result()
        result["error"] = (
            "The 'playwright' package is not installed. Re-run the launcher installer, or run: "
            "pip install playwright && python -m playwright install chromium"
        )
        return result
    return None


def run_live_check(save_report: bool = True) -> dict[str, Any]:
    missing = _playwright_missing_result()
    if missing is not None:
        return missing
    lock = _browser_lock()
    if not lock.acquire(blocking=False):
        result = _new_live_result()
        result["error"] = BROWSER_BUSY_MESSAGE
        return result
    try:
        return _run_live_check_locked(save_report)
    finally:
        lock.release()


def _auth_cookie_value(context: Any) -> str | None:
    from playwright.sync_api import Error as PlaywrightError

    try:
        cookies = context.cookies("https://www.amazon.com")
    except PlaywrightError:
        return None
    for cookie in cookies:
        if str(cookie.get("name", "")) in SIGNED_IN_COOKIE_NAMES:
            return str(cookie.get("value", ""))
    return None


def _on_login_wall(url: str) -> bool:
    return "/ap/" in url or any(marker in url for marker in LOGIN_WALL_URL_MARKERS)


def _is_vine_url(url: str) -> bool:
    return "amazon.com/vine" in url and not _on_login_wall(url)


def _save_login_attempt_snapshot(pages: list[Any]) -> None:
    from playwright.sync_api import Error as PlaywrightError

    debug_dir = _reports_dir() / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    for index, page in enumerate(pages):
        try:
            url = page.url
            content = page.content()
        except PlaywrightError:
            continue
        header = (
            f"<!-- url: {url}\n"
            f"     saved: {datetime.now().isoformat(timespec='seconds')} -->\n"
        )
        try:
            (debug_dir / f"login_attempt_{stamp}_{index}.html").write_text(
                header + content, encoding="utf-8"
            )
        except OSError:
            continue


def _wait_for_sign_in(context: Any, say: Any) -> tuple[str, Any | None, str]:
    """
    Let the user sign in inside the open window. Returns (outcome, vine_page,
    last_url) where outcome is "ok", "closed" or "timeout".

    Success means a tab is actually showing a Vine page (not a sign-in URL).
    The mere presence of the at-main cookie is NOT trusted: a stale one from a
    copied profile or an expired session is already in the jar when the window
    opens, which used to close the window instantly. A *newly issued* at-main
    only triggers a probe in a separate tab, so the user's own tab is never
    navigated away mid-sign-in (e.g. during MFA).
    """
    from playwright.sync_api import Error as PlaywrightError

    page = context.pages[0] if context.pages else context.new_page()
    seen_token = _auth_cookie_value(context)
    try:
        page.goto(VINE_ITEMS_URL, wait_until="domcontentloaded", timeout=60000)
    except PlaywrightError:
        pass  # sign-in redirects can hang; keep polling either way
    say(
        "Browser window opened — sign in to Amazon there. It closes by itself once "
        "the Vine page loads (up to 15 minutes)."
    )
    deadline = time.monotonic() + SIGN_IN_TIMEOUT_SECONDS
    next_notice = time.monotonic() + 30.0
    last_url = ""
    while time.monotonic() < deadline:
        try:
            pages = [p for p in context.pages if not p.is_closed()]
        except PlaywrightError:
            return "closed", None, last_url
        if not pages:
            return "closed", None, last_url
        urls: list[tuple[Any, str]] = []
        for candidate in pages:
            try:
                urls.append((candidate, candidate.url))
            except PlaywrightError:
                continue
        for candidate, url in urls:
            if _is_vine_url(url):
                return "ok", candidate, url
        if urls:
            last_url = urls[0][1]

        token = _auth_cookie_value(context)
        if token and token != seen_token:
            seen_token = token
            probe = None
            try:
                probe = context.new_page()
                probe.goto(VINE_ITEMS_URL, wait_until="domcontentloaded", timeout=45000)
                probe.wait_for_timeout(1500)
                if _is_vine_url(probe.url):
                    return "ok", probe, probe.url
            except PlaywrightError:
                pass
            if probe is not None:
                try:
                    probe.close()
                except PlaywrightError:
                    pass

        if time.monotonic() >= next_notice:
            next_notice = time.monotonic() + 30.0
            minutes_left = max(0, int((deadline - time.monotonic()) // 60))
            say(
                f"Still waiting for sign-in… about {minutes_left} min left. "
                f"Current page: {last_url[:120] or '(blank)'}"
            )
        time.sleep(SIGN_IN_POLL_SECONDS)
    return "timeout", None, last_url


def perform_sign_in(progress: Any = None, save_report: bool = True) -> dict[str, Any]:
    """
    Open a visible browser on the session profile, let the user sign in, then
    check the Vine status and read the recent orders in the SAME window. Reusing
    the window avoids the old close-then-relaunch race on the profile lock.
    Returns a result dict shaped like run_live_check().
    """

    def say(message: str) -> None:
        if progress is not None:
            progress(message)

    missing = _playwright_missing_result()
    if missing is not None:
        return missing
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import sync_playwright

    lock = _browser_lock()
    if not lock.acquire(blocking=False):
        result = _new_live_result()
        result["error"] = BROWSER_BUSY_MESSAGE
        return result
    try:
        result = _new_live_result()
        connected = False
        outcome = ""
        with sync_playwright() as pw:
            context = None
            try:
                context, used_channel = _launch_vine_context(pw, headless=False)
                result["channel"] = used_channel
                outcome, vine_page, last_url = _wait_for_sign_in(context, say)
                if outcome == "ok":
                    say("Signed in — checking claim status and reading your recent Vine orders...")
                    vine_page.wait_for_timeout(1500)
                    status, detail = classify_live_page(vine_page)
                    connected = _record_vine_status(
                        context, vine_page, status, detail, result, used_channel, save_report
                    )
                elif outcome == "timeout":
                    _save_login_attempt_snapshot(list(context.pages))
                    result["cap_status"] = CAP_NEEDS_LOGIN
                    result["error"] = (
                        "Timed out waiting for sign-in to finish "
                        f"(window's last page: {last_url[:160] or '(unknown)'}). "
                        "If your account is not a Vine member the Vine page never loads. "
                        "A page copy was saved under vine_reports/debug/."
                    )
            except (ConnectionError, PlaywrightError) as exc:
                result["error"] = str(exc)[:400]
            finally:
                if context is not None:
                    try:
                        context.close()
                    except PlaywrightError:
                        pass
        if connected:
            _save_cookie_backup()
        elif outcome == "closed":
            # The browser writes cookies to disk as it exits, so a sign-in the
            # user finished right before closing the window shows up only now.
            say("The window was closed — checking whether the sign-in went through...")
            # Use the session as the window left it: copying a browser profile
            # over it would throw away a sign-in that is partly done.
            result = _run_live_check_locked(save_report, bootstrap=False)
            if not result.get("connected"):
                result["error"] = "The sign-in window was closed before sign-in finished."
        return result
    finally:
        lock.release()


def perform_connect(progress: Any = None) -> dict[str, Any]:
    def say(message: str) -> None:
        if progress is not None:
            progress(message)

    say("Trying your saved Amazon session first (silent re-check)...")
    result = run_live_check()
    error_text = str(result.get("error") or "")
    if (
        result.get("connected")
        or "not installed" in error_text
        or error_text == BROWSER_BUSY_MESSAGE
    ):
        return result
    say("No usable saved session — opening a one-time Amazon sign-in window.")
    return perform_sign_in(progress=progress)


def save_debug_snapshot() -> tuple[list[str], str | None]:
    try:
        from playwright.sync_api import Error as PlaywrightError
        from playwright.sync_api import sync_playwright
    except ImportError:
        return [], "The 'playwright' package is not installed."
    lock = _browser_lock()
    if not lock.acquire(blocking=False):
        return [], BROWSER_BUSY_MESSAGE
    try:
        return _save_debug_snapshot_locked(PlaywrightError, sync_playwright)
    finally:
        lock.release()


def _save_debug_snapshot_locked(
    PlaywrightError: Any, sync_playwright: Any
) -> tuple[list[str], str | None]:
    debug_dir = _reports_dir() / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    saved: list[str] = []
    with sync_playwright() as pw:
        context = None
        try:
            try:
                context, _ = _launch_vine_context(pw)
            except ConnectionError as exc:
                return [], str(exc)
            page = context.pages[0] if context.pages else context.new_page()
            for label, url in (("items", VINE_ITEMS_URL), ("account", VINE_ACCOUNT_URL)):
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=45000)
                    page.wait_for_timeout(2500)
                    html_path = debug_dir / f"{label}_{stamp}.html"
                    png_path = debug_dir / f"{label}_{stamp}.png"
                    meta_path = debug_dir / f"{label}_{stamp}.txt"
                    content = page.content()
                    html_path.write_text(content, encoding="utf-8")
                    meta_path.write_text(
                        f"url: {page.url}\n"
                        f"classification: {classify_vine_page(content, page.url, _visible_text(page))}\n",
                        encoding="utf-8",
                    )
                    page.screenshot(path=str(png_path), full_page=True)
                    saved.extend([str(html_path), str(png_path), str(meta_path)])
                except PlaywrightError as exc:
                    return saved, f"{label} page: {str(exc).splitlines()[0]}"
        except PlaywrightError as exc:
            return saved, str(exc).splitlines()[0]
        finally:
            if context is not None:
                try:
                    context.close()
                except PlaywrightError:
                    pass
    return saved, None


# ---------------------------------------------------------------------------
# Sample Workbook Generator (for instant demo / testing if needed)
# ---------------------------------------------------------------------------


def build_demo_excel_bytes() -> bytes:
    """Create a realistic sample Amazon Vine Itemized Tax Report in memory."""
    buf = io.BytesIO()
    rows = [
        ["Amazon Vine Itemized Tax Report - Tax Year 2026", None, None, None, None, None, None],
        ["Account Holder: Sample Vine Voice", None, None, None, None, None, None],
        [None, None, None, None, None, None, None, None],
        [
            "Order Number",
            "ASIN",
            "Product Name",
            "Order Type",
            "Order Date",
            "Shipped Date",
            "Cancelled Date",
            "Estimated Tax Value",
        ],
        ["112-0000001-00001", "B00DEMO001", "4K OLED Portable Monitor", "ORDER", "2026-08-14", "2026-08-16", None, 429.99],
        ["112-0000001-00002", "B00DEMO002", "Commercial Espresso Machine", "ORDER", "2026-09-05", "2026-09-07", None, 899.00],
        ["112-0000001-00003", "B00DEMO003", "Ergonomic Mesh Office Chair", "ORDER", "2026-09-22", "2026-09-24", None, 540.00],
        ["112-0000001-00004", "B00DEMO004", "Smart Robot Vacuum & Mop", "ORDER", "2026-10-08", "2026-10-10", None, 780.00],
        ["112-0000001-00005", "B00DEMO005", "Wireless Mechanical Keyboard", "ORDER", "2026-10-21", "2026-10-23", None, 199.50],
        ["112-0000001-00006", "B00DEMO006", "Laser Home Theater Projector", "ORDER", "2026-10-25", "2026-10-27", None, 2199.99],
        ["112-0000001-00007", "B00DEMO007", "Portable Power Station 2000W", "ORDER", "2026-10-28", "2026-10-30", None, 1499.00],
        ["112-0000001-00008", "B00DEMO008", "Titanium Cookware 12pc Set", "ORDER", "2026-10-31", "2026-11-02", None, 766.09],
        ["112-0000001-00009", "B00DEMO009", "Studio Condenser Microphone", "ORDER", "2026-11-02", "2026-11-04", None, 349.99],
        ["112-0000001-00009", "B00DEMO009", "Studio Condenser Microphone", "CANCELLATION", "2026-11-02", None, "2026-11-03", -349.99],
        ["112-0000001-00010", "B00DEMO010", "Carbon Road Bike Wheelset", "ORDER", "2026-11-05", "2026-11-07", None, 1850.00],
        ["112-0000001-00011", "B00DEMO011", "Dual-Zone Wine Cooler Fridge", "ORDER", "2026-11-09", "2026-11-11", None, 1299.00],
        ["112-0000001-00012", "B00DEMO012", "Full-Frame Mirrorless Lens 35mm", "ORDER", "2026-11-12", "2026-11-14", None, 1498.00],
        ["112-0000001-00013", "B00DEMO013", "Smart Home Gym Cable System", "ORDER", "2026-11-16", "2026-11-18", None, 880.09],
        [None, None, None, None, None, None, None, None],
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
        div[data-testid="stMetricValue"] {
            font-size: clamp(1.35rem, 2.2vw, 2rem);
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

    if session_ready() and not st.session_state.get("vine_live_attempted"):
        st.session_state["vine_live_attempted"] = True
        with st.spinner(
            "Checking your saved Amazon Vine session (a Chrome window flashes briefly — "
            "Amazon blocks headless browsers)..."
        ):
            st.session_state["vine_live"] = run_live_check()

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

        live_orders = load_vine_orders_cache()
        use_live_orders = False
        if live_orders is not None:
            fetched_text = str(live_orders[1].get("fetched_at") or "")[:16].replace("T", " ")
            use_live_orders = st.checkbox(
                "Use Vine Orders page (Fair market value)",
                value=(uploaded_file is None and selected_local_file is None),
                disabled=(uploaded_file is not None or selected_local_file is not None),
                help=(
                    f"{live_orders[1]['valid_rows']:,} orders from the last "
                    f"{live_orders[1]['lookback_days']} days, read from amazon.com/vine/orders "
                    f"at {fetched_text}. Refresh in the Amazon Vine Live section re-reads them."
                ),
            )

        use_demo = st.checkbox(
            "Load built-in sample report",
            value=False,
            disabled=(
                uploaded_file is not None
                or selected_local_file is not None
                or use_live_orders
            ),
            help="Preview the dashboard using realistic sample Amazon Vine report data.",
        )

        st.divider()
        st.header("🔌 Amazon Vine Live")
        live_state = st.session_state.get("vine_live")
        if session_ready():
            if live_state:
                cap_status = str(live_state.get("cap_status"))
                cap_badge = {
                    CAP_CAPPED: "🔴 Capped",
                    CAP_OPEN: "🟢 Claiming open",
                    CAP_NEEDS_LOGIN: "🟠 Sign-in needed",
                    CAP_BLOCKED: "🟠 Blocked",
                    CAP_UNKNOWN: "⚪ Unknown",
                }.get(cap_status, "⚪ Unknown")
                checked_time = str(live_state.get("checked_at", ""))[11:16]
                st.markdown(f"**Connection: Active** · {cap_badge} · checked {checked_time}")
                if live_state.get("error"):
                    st.caption(f"⚠️ {live_state['error']}")
                if live_state.get("orders_error"):
                    st.caption(f"⚠️ Reading Vine orders: {live_state['orders_error']}")
                orders_info = live_state.get("orders")
                if orders_info:
                    st.caption(
                        f"📦 {orders_info['count']:,} orders in the last {ORDERS_LOOKBACK_DAYS} days · "
                        f"${orders_info['total_fmv']:,.2f} fair market value "
                        f"({orders_info['pages_read']} page(s) read)"
                    )
            else:
                st.markdown("**Connection: Active** — use Refresh to check status.")
            live_col1, live_col2 = st.columns(2)
            if live_col1.button("🔄 Refresh", width="stretch"):
                with st.spinner("Refreshing Amazon Vine status..."):
                    st.session_state["vine_live"] = run_live_check()
                st.rerun()
            if live_col2.button("🔑 Sign in", width="stretch"):
                with st.status("Opening the Amazon sign-in window...", expanded=True) as status:
                    def _sign_in_progress(message: str) -> None:
                        st.write(message)

                    result = perform_sign_in(progress=_sign_in_progress)
                    if not result.get("connected") and result.get("error"):
                        st.write(f"⚠️ {result['error']}")
                    st.session_state["vine_live"] = result
                    if result.get("connected"):
                        status.update(label="Connected to Amazon Vine", state="complete")
                    else:
                        status.update(
                            label=(
                                f"Still not connected ({result.get('cap_status') or 'UNKNOWN'}) "
                                "— see the message above and try again"
                            ),
                            state="error",
                        )
                st.rerun()
            if st.button("⛔ Disconnect", width="stretch"):
                disconnect_session()
                st.session_state.pop("vine_live", None)
                st.session_state["vine_live_attempted"] = False
                st.rerun()
        else:
            st.caption(
                "Optional: reuse your Chrome/Edge Amazon sign-in (close the browser entirely "
                "for the cookie snapshot to copy). If that isn't possible, a one-time sign-in "
                "window opens. Later checks re-open a Chrome window briefly — Amazon blocks "
                "headless browsers. Amazon passwords are never read or stored."
            )
            if st.button("🔌 Connect to Amazon Vine", width="stretch"):
                with st.status("Connecting to Amazon Vine...", expanded=True) as status:
                    def _connect_progress(message: str) -> None:
                        st.write(message)

                    result = perform_connect(progress=_connect_progress)
                    st.session_state["vine_live"] = result
                    if result.get("connected"):
                        status.update(label="Connected to Amazon Vine", state="complete")
                    else:
                        status.update(
                            label=(
                                f"Could not connect ({result.get('cap_status') or 'UNKNOWN'}) "
                                "— see the message above"
                            ),
                            state="error",
                        )
                st.session_state["vine_live_attempted"] = True
                st.rerun()
        with st.expander("Troubleshooting"):
            st.caption(
                "If the live check misbehaves (unexpected dropdown or bot check), save page "
                "snapshots and share the files in vine_reports/debug/ for selector tuning."
            )
            if st.button("Save page snapshot (debug)"):
                with st.spinner("Saving Vine page snapshots..."):
                    saved_files, snapshot_error = save_debug_snapshot()
                if saved_files:
                    st.success("Saved: " + ", ".join(Path(p).name for p in saved_files))
                if snapshot_error:
                    st.error(f"Snapshot issue: {snapshot_error}")
                if not saved_files and not snapshot_error:
                    st.error("No browser engine available to save snapshots.")

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

        date_basis = st.radio(
            "Count Items By",
            options=DATE_BASIS_OPTIONS,
            index=0,
            horizontal=True,
            help=(
                "Which date places an item in the rolling window. Ship Date uses the "
                "Shipped Date (cancellations use their Cancelled Date, since they never "
                "ship); Order Date uses the date the item was ordered."
            ),
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
    elif use_live_orders:
        file_source = "live-orders"
        source_label = "amazon.com/vine/orders (Fair market value)"
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
                - **`Shipped Date`**: Date the item shipped (used by default to place items in the window).
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
        if use_live_orders and live_orders is not None:
            df, meta = live_orders
            if df.empty:
                st.info(
                    f"No Vine orders were found in the last {meta['lookback_days']} days on "
                    "amazon.com/vine/orders, so the rolling window is $0.00."
                )
                return
        elif hasattr(file_source, "getvalue"):
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

    # The current year's rows, for the annual figures, before any prior year is added.
    year_df = df
    # Place every row by the chosen date (Ship or Order) for all calculations.
    df = apply_date_basis(df, date_basis)
    year_df = apply_date_basis(year_df, date_basis)
    if use_live_orders:
        if date_basis == DATE_BASIS_SHIP:
            st.info(
                "The Vine Orders page lists order dates only, so items are counted by Order Date."
            )
        if not meta.get("complete", True):
            st.warning(
                f"Stopped after {meta['pages_read']} Vine orders pages before reaching "
                f"{meta['lookback_days']} days back; the window total may be low."
            )
        if int(window_length) > int(meta["lookback_days"]):
            st.warning(
                f"Only the last {meta['lookback_days']} days of orders were read, so a "
                f"{int(window_length)}-day window will miss older orders."
            )
    elif date_basis == DATE_BASIS_SHIP and not meta.get("has_shipped_date", False):
        st.warning(
            "This report has no Shipped Date column, so items are counted by Order Date "
            "(cancellations by Cancelled Date)."
        )

    # -----------------------------------------------------------------------
    # Base Date Selection (Sidebar extension once file is loaded)
    # -----------------------------------------------------------------------
    latest_basis_date: pd.Timestamp = df[BASIS_COL].max()
    latest_basis_label = "Latest Ship Date" if date_basis == DATE_BASIS_SHIP else "Latest Order Date"
    today_ts = pd.Timestamp(date.today()).normalize()

    with base_date_container:
        st.divider()
        st.header("3. Base Date Settings")
        base_date_mode = st.radio(
            "Base Date (Day 0)",
            options=[
                f"Today's Date ({today_ts.strftime('%Y-%m-%d')})",
                f"{latest_basis_label} ({latest_basis_date.strftime('%Y-%m-%d')})",
                "Custom Date",
            ],
            index=0,
            help="Starting reference date (Day 0) for the current rolling window and future decay projection.",
        )

        if base_date_mode.startswith("Today's Date"):
            base_date = today_ts
        elif base_date_mode.startswith(latest_basis_label):
            base_date = latest_basis_date
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
    total_annual_tax_value = round(float(year_df["Estimated Tax Value"].sum()), 2)
    timeframe_summary_df = compute_timeframe_summary(
        df=df,
        base_date=base_date,
        inclusive_start=False,
        year_df=year_df,
        max_days=int(meta["lookback_days"]) if use_live_orders else None,
    )

    # Format Threshold Crossing KPI display. The value is just the short date (or a
    # two-word status) so it fits the card; the amount and timing go underneath.
    def _short_date(ts: pd.Timestamp) -> str:
        # "Dec 4"; the year is added only when it differs from the base date's.
        text = f"{ts:%b} {ts.day}"
        return text if ts.year == base_date.year else f"{text}, {ts.year}"

    def _days(n: int) -> str:
        return f"{n} day" if n == 1 else f"{n} days"

    if crossing_info is not None:
        c_val_str = f"${crossing_info['projected_total']:,.2f}"
        if crossing_info["already_below"]:
            crossing_headline = "Under now"
            crossing_sub = f"{c_val_str} as of {_short_date(crossing_info['date'])}"
        else:
            crossing_headline = _short_date(crossing_info["date"])
            crossing_sub = f"{c_val_str} · in {_days(crossing_info['day_offset'])}"
    else:
        # A shorter horizon than the window may stop before the crossing: look as far
        # as the window. A longer one already covered it, so there's nothing to redo.
        extended_crossing = None
        if int(projection_horizon) < int(window_length):
            _extended_df, extended_crossing = generate_decay_projection(
                df=df,
                base_date=base_date,
                window_days=int(window_length),
                threshold=float(tax_threshold),
                horizon_days=int(window_length),
                inclusive_start=False,
            )
        if extended_crossing is not None:
            c_val_str = f"${extended_crossing['projected_total']:,.2f}"
            crossing_headline = _short_date(extended_crossing["date"])
            crossing_sub = (
                f"{c_val_str} · in {_days(extended_crossing['day_offset'])} (past horizon)"
            )
        else:
            crossing_headline = "Not reached"
            crossing_sub = f"Stays ≥ ${tax_threshold:,.2f} for {_days(int(projection_horizon))}"

    # -----------------------------------------------------------------------
    # 1. Top Header & KPI Scorecards
    # -----------------------------------------------------------------------
    st.subheader("Key Performance Indicators")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

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
        if use_live_orders:
            st.metric(
                label=f"Fair Market Value, Last {meta['lookback_days']} Days",
                value=f"${total_annual_tax_value:,.2f}",
                delta=f"{len(year_df):,} orders from the Vine Orders page",
                delta_color="off",
            )
        else:
            st.metric(
                label="Total Annual Tax Value",
                value=f"${total_annual_tax_value:,.2f}",
                delta=f"{len(year_df):,} net line items",
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

    live_result = st.session_state.get("vine_live")
    if live_result:
        live_cap_status = str(live_result.get("cap_status"))
        if live_cap_status == CAP_CAPPED:
            mismatch_note = ""
            if current_window_total < float(tax_threshold):
                mismatch_note = (
                    " Your report-based projection disagrees — Amazon also counts items that "
                    "have not shipped yet or are otherwise missing from the itemized report."
                )
            st.error(
                "🚫 **Claiming is paused.** Amazon has paused your Vine claiming because "
                f"your ETV cap was reached.{mismatch_note}"
            )
        elif live_cap_status == CAP_OPEN:
            st.success("✅ **Claiming is online**")
        elif live_cap_status in (CAP_NEEDS_LOGIN, CAP_BLOCKED):
            st.warning(
                f"🟠 **Live Amazon check unavailable:** {live_result.get('error') or 'Sign-in required.'} "
                "Use **🔑 Sign in** in the sidebar to refresh the session."
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
            f"Rolling period totals calculated relative to Base Date ({base_date.strftime('%Y-%m-%d')}), "
            f"counting items by {date_basis}."
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
        for date_col in ("Order Date", "Shipped Date", "Cancelled Date", BASIS_COL):
            if date_col in display_tx_df.columns:
                display_tx_df[date_col] = display_tx_df[date_col].apply(
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
