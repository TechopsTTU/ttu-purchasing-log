# app.py


# Standard library imports
import time
from io import BytesIO
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

# Third-party imports
import streamlit as st
import pandas as pd
import plotly.express as px
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image as ReportLabImage,
    Table,
    TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle

# Configure Streamlit page
st.set_page_config(
    page_title="TTU Purchase Orders Log",
    layout="wide",  # 'wide' layout to accommodate sidebar elements
    initial_sidebar_state="expanded",
)

# Apply Seaborn style
# Apply Seaborn style
# sns.set_style("whitegrid")
# Inject CSS for styling
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    :root {
        --ttu-navy: #1f4e79;
        --ttu-azure: #2c5aa0;
        --ttu-sky: #90caf9;
        --ttu-gold: #ffd54f;
        --ttu-sage: #aed581;
        --ttu-surface: rgba(255, 255, 255, 0.75);
        --ttu-border: rgba(0, 0, 0, 0.08);
    }

    .stApp {
        background: var(--background-color, #f7f9fb);
        font-family: 'Inter', sans-serif;
        color: inherit;
    }

    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }

    .title {
        font-size: clamp(2rem, 4vw, 2.75rem);
        font-weight: 700;
        color: var(--ttu-navy);
        text-align: left;
        text-transform: uppercase;
        margin: 0;
        line-height: 1.1;
    }

    .dashboard-subtitle {
        color: rgba(0, 0, 0, 0.6);
        font-size: 1rem;
        margin-top: 0.35rem;
    }

    .card {
        background: var(--ttu-surface);
        border-radius: 16px;
        border: 1px solid var(--ttu-border);
        padding: 18px;
        box-shadow: 0 16px 40px rgba(31, 78, 121, 0.08);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        min-height: 110px;
    }

    .card:hover {
        transform: translateY(-3px);
        box-shadow: 0 24px 50px rgba(31, 78, 121, 0.12);
    }

    .card h3 {
        margin-bottom: 0.6rem;
        font-size: 0.95rem;
        font-weight: 600;
        color: rgba(0, 0, 0, 0.65);
        letter-spacing: 0.4px;
        text-transform: uppercase;
    }

    .card p {
        margin: 0;
        font-size: 1.5rem;
        font-weight: 700;
        color: var(--ttu-navy);
    }

    .percentage-value {
        font-size: clamp(2.4rem, 5vw, 3rem);
        color: #1b8a5a;
        font-weight: 800;
        margin: 0;
    }

    .progress-wrapper {
        width: 100%;
        margin-top: 10px;
    }

    .progress-track {
        width: 100%;
        height: 12px;
        border-radius: 999px;
        background: rgba(31, 78, 121, 0.12);
        overflow: hidden;
        box-shadow: inset 0 1px 3px rgba(0, 0, 0, 0.15);
    }

    .progress-fill {
        height: 100%;
        border-radius: 999px;
        background: linear-gradient(90deg, #1b8a5a 0%, #4caf50 100%);
        transition: width 0.6s ease;
    }

    .insight-pill {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: rgba(31, 78, 121, 0.12);
        color: var(--ttu-navy);
        padding: 8px 14px;
        border-radius: 999px;
        font-weight: 500;
        margin: 6px 6px 6px 0;
        box-shadow: inset 0 0 0 1px rgba(31, 78, 121, 0.15);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(31, 78, 121, 0.95), rgba(44, 90, 160, 0.92));
        color: white;
    }

    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] .stMarkdown p {
        color: rgba(255, 255, 255, 0.92) !important;
    }

    /* Enhanced button styling */
    .stButton > button {
        background: linear-gradient(145deg, var(--ttu-navy) 0%, var(--ttu-azure) 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.2rem;
        font-weight: 600;
        font-size: 0.9rem;
        transition: all 0.3s ease;
        box-shadow: 0 6px 18px rgba(31, 78, 121, 0.25);
    }

    .stButton > button:hover {
        background: linear-gradient(145deg, var(--ttu-azure) 0%, var(--ttu-navy) 100%);
        transform: translateY(-2px);
        box-shadow: 0 10px 24px rgba(31, 78, 121, 0.35);
    }

    /* Section headers */
    .section-header {
        background: linear-gradient(90deg, rgba(31, 78, 121, 0.95), rgba(44, 90, 160, 0.85));
        color: white;
        padding: 12px 18px;
        border-radius: 12px;
        margin: 24px 0 16px 0;
        font-weight: 600;
        font-size: 1.1rem;
        box-shadow: 0 10px 24px rgba(31, 78, 121, 0.25);
    }

    /* Enhanced table styling */
    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 20px 50px rgba(15, 34, 58, 0.15);
        background: rgba(255, 255, 255, 0.9);
        border: 1px solid rgba(31, 78, 121, 0.12);
        margin-bottom: 1.25rem;
    }

    div[data-testid="stDataFrame"]::-webkit-scrollbar {
        height: 10px;
    }

    div[data-testid="stDataFrame"]::-webkit-scrollbar-thumb {
        background: rgba(31, 78, 121, 0.35);
        border-radius: 8px;
    }

    /* File uploader styling */
    .stFileUploader {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 20px;
        border: 2px dashed rgba(255, 255, 255, 0.35);
    }

    /* Processing time styling */
    .processing-time {
        text-align: center;
        color: rgba(0, 0, 0, 0.55);
        font-size: 0.8rem;
        font-style: italic;
        margin-top: 30px;
        padding: 10px 14px;
        background: rgba(255, 255, 255, 0.8);
        border-radius: 20px;
        border: 1px solid rgba(31, 78, 121, 0.1);
        backdrop-filter: blur(10px);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# Title and Logo
def display_logo_title():
    col1, col2 = st.columns([1, 4])
    with col1:
        if os.path.exists("TTU_LOGO.jpg"):
            st.image("TTU_LOGO.jpg", width=120)
    with col2:
        st.markdown(
            """
            <div style='margin-top: 20px;'>
                <h1 class="title">TTU Purchase Orders Log</h1>
                <p class='dashboard-subtitle' style='margin-top: -6px;'>
                    📊 Purchase Order Analytics & Reporting Dashboard
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )


# Callback to reset all sidebar filters
def reset_filters(defaults: Dict[str, Any]) -> None:
    for key, value in defaults.items():
        st.session_state[key] = value
    st.experimental_rerun()


def _read_source_frame(name: str, payload: bytes) -> pd.DataFrame:
    buffer = BytesIO(payload)
    try:
        return pd.read_excel(buffer, engine="openpyxl")
    except Exception:
        buffer.seek(0)
        return pd.read_csv(buffer)


# Caching the data loading function
@st.cache_data(show_spinner=False)
def load_and_process_data(
    file_payloads: Tuple[Tuple[str, bytes], ...], use_demo: bool
) -> Tuple[pd.DataFrame, Dict[str, Any], List[str]]:
    quality: Dict[str, Any] = {
        "sources": [],
        "rows_loaded": 0,
        "rows_retained": 0,
        "drops": {},
    }

    frames: List[pd.DataFrame] = []

    if use_demo:
        demo_path = Path("data/demo_purchase_orders.csv")
        if demo_path.exists():
            df_demo = pd.read_csv(demo_path)
            df_demo["__source__"] = demo_path.name
            frames.append(df_demo)
            quality["sources"].append(demo_path.name)
        else:
            return pd.DataFrame(), quality, []
    else:
        for name, payload in file_payloads:
            if payload is None:
                continue
            try:
                df_source = _read_source_frame(name, payload)
            except Exception:
                continue
            df_source["__source__"] = name
            frames.append(df_source)
            quality["sources"].append(name)

    if not frames:
        return pd.DataFrame(), quality, []

    raw_df = pd.concat(frames, ignore_index=True)
    quality["rows_loaded"] = len(raw_df)

    raw_df.rename(columns={"Acct": "Purchase Account"}, inplace=True)

    date_columns = [col for col in raw_df.columns if "date" in col.lower()]
    for col in date_columns:
        raw_df[col] = pd.to_datetime(raw_df[col], errors="coerce").dt.normalize()

    if "OrderDate" not in raw_df.columns:
        quality["drops"]["missing_order_date_column"] = quality["rows_loaded"]
        return pd.DataFrame(), quality, date_columns

    missing_order_dates = raw_df["OrderDate"].isna().sum()
    df_filtered = raw_df.dropna(subset=["OrderDate"]).copy()
    quality["drops"]["missing_order_date"] = int(missing_order_dates)

    cutoff = pd.to_datetime("2022-01-01")
    prior_to_cutoff = (df_filtered["OrderDate"] < cutoff).sum()
    df_filtered = df_filtered[df_filtered["OrderDate"] >= cutoff].copy()
    quality["drops"]["prior_to_2022"] = int(prior_to_cutoff)

    critical_columns = ["OrderDate", "PONumber", "Total"]
    missing_critical = len(df_filtered) - len(df_filtered.dropna(subset=critical_columns))
    df_filtered.dropna(subset=critical_columns, inplace=True)
    quality["drops"]["critical_missing"] = int(missing_critical)

    before_duplicates = len(df_filtered)
    df_filtered.drop_duplicates(inplace=True)
    quality["drops"]["duplicates_removed"] = int(before_duplicates - len(df_filtered))

    numerical_columns = ["Total", "Amt", "QtyOrdered", "QtyRemaining"]
    for col in numerical_columns:
        if col in df_filtered.columns:
            df_filtered[col] = pd.to_numeric(df_filtered[col], errors="coerce").fillna(0.0)

    if "Purchase Account" in df_filtered.columns:
        before_accounts = len(df_filtered)
        df_filtered["Purchase Account"] = (
            df_filtered["Purchase Account"]
            .astype(str)
            .str.replace(r"[^0-9]", "", regex=True)
        )
        df_filtered["Purchase Account"] = df_filtered["Purchase Account"].str.zfill(8)
        df_filtered = df_filtered[df_filtered["Purchase Account"].str.strip() != ""]
        df_filtered["Purchase Account"] = df_filtered["Purchase Account"].str.replace(
            r"(\d{4})(\d{4})", r"\1-\2", regex=True
        )
        quality["drops"]["invalid_purchase_account"] = int(
            before_accounts - len(df_filtered)
        )
    else:
        quality["drops"]["missing_purchase_account_column"] = len(df_filtered)

    df_filtered.rename(
        columns={"QtyRemaining": "qty on order/backordered"}, inplace=True
    )

    df_filtered.sort_values("OrderDate", inplace=True)
    quality["rows_retained"] = len(df_filtered)

    return df_filtered, quality, date_columns


# Map POStatus codes
def map_po_status(df):
    po_status_mapping = {
        "NN": "NEW",
        "AN": "OPEN",
        "F": "RECEIVED",
        "BN": "BACKORDERED",
    }
    if "POStatus" in df.columns:
        df["POStatus"] = df["POStatus"].map(po_status_mapping).fillna(df["POStatus"])
    else:
        st.write("'POStatus' column is missing.")
    return df


# Sidebar filter builder
def build_filter_sidebar(df: pd.DataFrame) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    order_min = df["OrderDate"].min().date()
    order_max = df["OrderDate"].max().date()
    defaults: Dict[str, Any] = {
        "order_date_range": (order_min, order_max),
        "purchase_account_filter": "All",
        "requisitioner_filter": "All",
    }

    if "RequestDate" in df.columns and df["RequestDate"].notna().any():
        request_min = df["RequestDate"].min().date()
        request_max = df["RequestDate"].max().date()
        defaults["request_date_range"] = (request_min, request_max)

    vendor_options = sorted(df.get("VendorName", pd.Series(dtype=str)).dropna().unique())
    defaults["vendor_filter"] = vendor_options

    po_status_options = sorted(df.get("POStatus", pd.Series(dtype=str)).dropna().unique())
    defaults["po_status_filter"] = po_status_options

    total_series = df.get("Total", pd.Series(dtype=float)).dropna()
    if total_series.empty:
        total_min, total_max = 0.0, 0.0
    else:
        total_min = float(total_series.min())
        total_max = float(total_series.max())
    if total_min == total_max:
        total_max = total_min + 1
    defaults["total_range"] = (total_min, total_max)

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    with st.sidebar:
        st.markdown("### 🎯 Filters")

        order_range = st.date_input(
            "Order date range",
            value=st.session_state["order_date_range"],
            key="order_date_range",
        )

        request_range = None
        if "request_date_range" in defaults:
            request_range = st.date_input(
                "Request date range",
                value=st.session_state.get("request_date_range", defaults["request_date_range"]),
                key="request_date_range",
            )

        purchase_accounts = ["All"] + sorted(df["Purchase Account"].dropna().unique().tolist())
        selected_account = st.selectbox(
            "Purchase account",
            options=purchase_accounts,
            index=purchase_accounts.index(st.session_state["purchase_account_filter"])
            if st.session_state["purchase_account_filter"] in purchase_accounts
            else 0,
            key="purchase_account_filter",
        )

        requisitioners = ["All"] + sorted(df.get("Requisitioner", pd.Series(dtype=str)).dropna().unique().tolist())
        selected_requisitioner = st.selectbox(
            "Requisitioner",
            options=requisitioners,
            index=requisitioners.index(st.session_state["requisitioner_filter"])
            if st.session_state["requisitioner_filter"] in requisitioners
            else 0,
            key="requisitioner_filter",
        )

        vendor_default = [
            value
            for value in st.session_state.get("vendor_filter", vendor_options)
            if value in vendor_options
        ]
        if not vendor_default and vendor_options:
            vendor_default = vendor_options
        st.session_state["vendor_filter"] = vendor_default
        selected_vendors = st.multiselect(
            "Vendors",
            options=vendor_options,
            default=vendor_default,
            key="vendor_filter",
        )

        status_default = [
            value
            for value in st.session_state.get("po_status_filter", po_status_options)
            if value in po_status_options
        ]
        if not status_default and po_status_options:
            status_default = po_status_options
        st.session_state["po_status_filter"] = status_default
        selected_statuses = st.multiselect(
            "PO status",
            options=po_status_options,
            default=status_default,
            key="po_status_filter",
        )

        selected_total_range = st.slider(
            "Total amount range",
            min_value=float(total_min),
            max_value=float(total_max),
            value=st.session_state.get("total_range", (total_min, total_max)),
            step=100.0 if total_max - total_min > 1000 else 10.0,
            key="total_range",
        )

        st.button("Reset filters", key="reset_filters_button", on_click=reset_filters, args=(defaults,))

    filters = {
        "order_date_range": order_range,
        "request_date_range": request_range,
        "purchase_account": selected_account,
        "requisitioner": selected_requisitioner,
        "vendors": selected_vendors,
        "statuses": selected_statuses,
        "total_range": selected_total_range,
    }

    return filters, defaults


def apply_filters(df: pd.DataFrame, filters: Dict[str, Any]) -> pd.DataFrame:
    filtered = df.copy()

    order_start, order_end = filters["order_date_range"]
    filtered = filtered[
        (filtered["OrderDate"] >= pd.to_datetime(order_start))
        & (filtered["OrderDate"] <= pd.to_datetime(order_end))
    ]

    request_range = filters.get("request_date_range")
    if request_range and "RequestDate" in filtered.columns:
        req_start, req_end = request_range
        filtered = filtered[
            (filtered["RequestDate"] >= pd.to_datetime(req_start))
            & (filtered["RequestDate"] <= pd.to_datetime(req_end))
        ]

    if filters.get("purchase_account") and filters["purchase_account"] != "All":
        filtered = filtered[filtered["Purchase Account"] == filters["purchase_account"]]

    if filters.get("requisitioner") and filters["requisitioner"] != "All":
        filtered = filtered[filtered["Requisitioner"] == filters["requisitioner"]]

    if filters.get("vendors"):
        filtered = filtered[filtered["VendorName"].isin(filters["vendors"])]

    if filters.get("statuses"):
        filtered = filtered[filtered["POStatus"].isin(filters["statuses"])]

    total_min, total_max = filters.get("total_range", (None, None))
    if total_min is not None and total_max is not None and "Total" in filtered.columns:
        filtered = filtered[(filtered["Total"] >= total_min) & (filtered["Total"] <= total_max)]

    return filtered


def render_data_quality(quality: Dict[str, Any]) -> None:
    drops = quality.get("drops", {})
    total_dropped = sum(drops.values())
    if total_dropped == 0:
        return

    reason_labels = {
        "missing_order_date_column": "Missing order date column",
        "missing_order_date": "Rows without order date",
        "prior_to_2022": "Orders before 2022",
        "critical_missing": "Rows missing critical values",
        "duplicates_removed": "Duplicate rows",
        "invalid_purchase_account": "Rows without a valid purchase account",
        "missing_purchase_account_column": "Missing purchase account column",
    }

    retained = quality.get("rows_retained", 0)
    loaded = quality.get("rows_loaded", 0)

    with st.sidebar.expander("🧹 Data quality checks", expanded=False):
        st.markdown(
            f"**{loaded - total_dropped:,} of {loaded:,} rows** remain after cleaning." if loaded else "No rows processed yet."
        )
        for reason, count in drops.items():
            if count:
                label = reason_labels.get(reason, reason.replace("_", " ").title())
                st.markdown(f"- {label}: **{int(count):,}**")
        if retained:
            st.caption(f"Rows available for analysis: {retained:,}")


# Create index cards
def display_index_cards(metrics):
    if not metrics:
        return

    column_count = min(3, len(metrics))
    columns = st.columns(column_count)

    for idx, (metric_name, metric_info) in enumerate(metrics.items()):
        bg_color = metric_info.get("bg_color", "var(--ttu-surface)")
        value = metric_info.get(metric_name, "N/A")
        value_formatted = value.replace("<br>", "<br/>")
        with columns[idx % column_count]:
            st.markdown(
                f"""
                <div class="card" style='background-color: {bg_color}; width: 100%;'>
                    <h3>{metric_name}</h3>
                    <p>{value_formatted}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
def plotly_to_image(fig, format="png", **kwargs):
    """Safely convert a Plotly figure to an in-memory image.

    Falls back gracefully if Kaleido/Chrome is not available.

    Parameters
    ----------
    fig : plotly.graph_objs.Figure
        The Plotly figure to convert.
    format : str, optional
        Image format understood by Kaleido (default ``"png"``).
    kwargs : Any
        Additional keyword arguments passed to ``fig.to_image``.

    Returns
    -------
    BytesIO or None
        A buffer containing the image data if conversion succeeds,
        otherwise ``None``.
    """
    try:
        image_bytes = fig.to_image(format=format, **kwargs)
        return BytesIO(image_bytes)
    except Exception as exc:  # pragma: no cover - runtime dependency on Kaleido
        st.warning(
            "Unable to export the Plotly figure to an image. "
            "Install Kaleido for image export support."
        )
        st.write(str(exc))
        return None


def format_currency(value) -> str:
    """Return the value formatted as a currency string."""
    try:
        if pd.isna(value):
            return "$0.00"
        return f"${float(value):,.2f}"
    except (TypeError, ValueError):
        return "$0.00"


def format_percentage(value) -> str:
    try:
        return f"{value:.2f}%"
    except (TypeError, ValueError):
        return "0.00%"
# Generate matrix of on-time delivery metrics by GL account (Purchase Account)
def otd_matrix_by_account(df: pd.DataFrame) -> pd.DataFrame:
    """Return on-time and late counts with percentage by Purchase Account."""
    required_cols = {"RecDate", "RequestDate", "Purchase Account"}
    if not required_cols.issubset(df.columns):
        return pd.DataFrame()

    temp = df.copy()
    temp["RecDate"] = pd.to_datetime(temp["RecDate"]).dt.normalize()
    temp["RequestDate"] = pd.to_datetime(temp["RequestDate"]).dt.normalize()
    temp["On_Time"] = temp["RecDate"] <= temp["RequestDate"]

    summary = (
        temp.groupby("Purchase Account")["On_Time"]
        .agg(On_Time="sum", Late=lambda x: (~x).sum())
        .reset_index()
    )
    summary["On-Time %"] = (
        summary["On_Time"] / (summary["On_Time"] + summary["Late"]) * 100
    ).round(2)
    summary.rename(columns={"On_Time": "On-Time"}, inplace=True)
    return summary


# Main application logic


def main():
    display_logo_title()
    st.divider()

    with st.sidebar:
        if os.path.exists("TTU_LOGO.jpg"):
            st.image("TTU_LOGO.jpg", use_column_width=True)
            st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📁 Data source")
        uploaded_files = st.file_uploader(
            "Upload Excel workbooks",
            type=["xlsx"],
            accept_multiple_files=True,
            help="Upload one or more monthly exports to merge them automatically.",
        )
        st.caption("Uploaded files are merged in the order you select.")
        use_demo = st.toggle(
            "Use demo dataset",
            value=not uploaded_files,
            help="Load a bundled sample workbook to explore the dashboard.",
        )
        if uploaded_files:
            st.success(
                f"✅ {len(uploaded_files)} file{'s' if len(uploaded_files) > 1 else ''} ready for analysis."
            )
        elif use_demo:
            st.info("Demo dataset is active.")
        st.markdown("---")

    processing_start_time = time.time()
    file_payloads = (
        tuple((uploaded_file.name, uploaded_file.getvalue()) for uploaded_file in uploaded_files)
        if uploaded_files
        else tuple()
    )

    df_processed, quality, date_columns = load_and_process_data(file_payloads, use_demo)
    if df_processed.empty:
        render_data_quality(quality)
        st.info("Upload at least one Excel workbook or enable the demo dataset from the sidebar to begin.")
        return

    df_processed = map_po_status(df_processed)

    filters, defaults = build_filter_sidebar(df_processed)
    render_data_quality(quality)

    df_filtered = apply_filters(df_processed, filters)

    if df_filtered.empty:
        st.warning("No results found for the selected filters. Try adjusting them in the sidebar.")
        return

    order_start_date = pd.to_datetime(filters["order_date_range"][0]).date()
    order_end_date = pd.to_datetime(filters["order_date_range"][1]).date()
    request_range = filters.get("request_date_range")
    request_start_date = pd.to_datetime(request_range[0]).date() if request_range else None
    request_end_date = pd.to_datetime(request_range[1]).date() if request_range else None
    selected_requisitioner = filters["requisitioner"]
    selected_purchase_account = filters["purchase_account"]

    total_line_items = len(df_filtered)
    total_unique_pos = df_filtered["PONumber"].nunique() if "PONumber" in df_filtered.columns else total_line_items

    summary_col1, summary_col2 = st.columns([3, 2])
    with summary_col1:
        st.markdown(
            f"**Analyzing {total_line_items:,} line items across {total_unique_pos:,} purchase orders.**"
        )
        if quality.get("sources"):
            st.caption("Sources merged: " + ", ".join(quality["sources"]))
        total_removed = sum(quality.get("drops", {}).values())
        if total_removed:
            st.caption(f"Data cleaning removed {total_removed:,} rows — see the sidebar for details.")
    with summary_col2:
        st.download_button(
            label="Download filtered data (CSV)",
            data=df_filtered.to_csv(index=False).encode("utf-8"),
            file_name="ttu_purchase_orders_filtered.csv",
            mime="text/csv",
        )

    metrics = {}
    if "Amt" in df_filtered.columns and "POStatus" in df_filtered.columns:
        total_open_orders_amt = df_filtered[df_filtered["POStatus"] == "OPEN"]["Amt"].sum()
        metrics["Total Open Orders Amt"] = {
            "Total Open Orders Amt": f"${total_open_orders_amt:,.2f}",
            "bg_color": "rgba(144, 202, 249, 0.45)",
        }
    else:
        metrics["Total Open Orders Amt"] = {
            "Total Open Orders Amt": "$0.00",
            "bg_color": "rgba(144, 202, 249, 0.45)",
        }

    if "PONumber" in df_filtered.columns:
        total_orders_placed = df_filtered["PONumber"].nunique()
        metrics["Total Orders Placed"] = {
            "Total Orders Placed": f"{total_orders_placed}",
            "bg_color": "rgba(255, 205, 210, 0.45)",
        }

    metrics["Total Lines Ordered"] = {
        "Total Lines Ordered": f"{total_line_items}",
        "bg_color": "rgba(200, 230, 201, 0.45)",
    }

    if {"Total", "PONumber", "VendorName", "Requisitioner"}.issubset(df_filtered.columns):
        max_total_row = df_filtered.loc[df_filtered["Total"].idxmax()]
        max_total_formatted = f"${max_total_row['Total']:,.2f}"
        most_expensive_order_info = (
            f"PO Number: {max_total_row['PONumber']}<br/>"
            f"Vendor: {max_total_row['VendorName']}<br/>"
            f"Requisitioner: {max_total_row['Requisitioner']}<br/>"
            f"Total: {max_total_formatted}"
        )
        metrics["Most Expensive Order"] = {
            "Most Expensive Order": most_expensive_order_info,
            "bg_color": "rgba(255, 213, 79, 0.45)",
        }
    else:
        metrics["Most Expensive Order"] = {
            "Most Expensive Order": "N/A",
            "bg_color": "rgba(255, 213, 79, 0.45)",
        }

    display_index_cards(metrics)

    if selected_requisitioner != "All":
        last_order = df_filtered.sort_values(by="OrderDate", ascending=False).head(1)
        if not last_order.empty:
            last_order_row = last_order.iloc[0]
            last_order_info = (
                f"Order Date: {last_order_row['OrderDate'].date()}<br/>"
                f"PO Number: {last_order_row['PONumber']}<br/>"
                f"Vendor: {last_order_row['VendorName']}<br/>"
                f"Total: ${last_order_row['Total']:,.2f}<br/>"
                f"Status: {last_order_row['POStatus']}"
            )
            st.markdown(
                f"""
                <div class="card" style='background-color: rgba(187, 222, 251, 0.55); width: 100%;'>
                    <h3>Last Order for {selected_requisitioner}</h3>
                    <p>{last_order_info}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("### 📊 Trends & Insights")
    trend_summary = pd.DataFrame()
    vendor_summary_pdf = pd.DataFrame()
    pdf_sections = []

    if {"OrderDate", "Total"}.issubset(df_filtered.columns):
        trend_df = df_filtered.copy()
        trend_df["OrderDate"] = pd.to_datetime(trend_df["OrderDate"], errors="coerce")
        trend_df.dropna(subset=["OrderDate"], inplace=True)
        if not trend_df.empty:
            trend_df["Order Month"] = trend_df["OrderDate"].dt.to_period("M").dt.to_timestamp()
            agg_dict = {"total_spend": ("Total", "sum")}
            if "PONumber" in trend_df.columns:
                agg_dict["unique_pos"] = ("PONumber", "nunique")
            else:
                agg_dict["unique_pos"] = ("Total", "size")
            trend_summary = trend_df.groupby("Order Month").agg(**agg_dict).reset_index()
            trend_summary.rename(
                columns={"total_spend": "Total Spend", "unique_pos": "Unique POs"}, inplace=True
            )

    if {"VendorName", "Total"}.issubset(df_filtered.columns):
        vendor_summary_pdf = (
            df_filtered.groupby("VendorName")["Total"].sum().reset_index().sort_values("Total", ascending=False)
        )
        vendor_summary_pdf = vendor_summary_pdf.head(10)

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        if not trend_summary.empty:
            spend_fig = px.line(
                trend_summary,
                x="Order Month",
                y="Total Spend",
                markers=True,
                title="Spend Over Time",
            )
            spend_fig.update_layout(
                title_font=dict(color="#1f4e79", size=16),
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                hovermode="x unified",
                yaxis_tickprefix="$",
            )
            st.plotly_chart(spend_fig, use_container_width=True)
    with chart_col2:
        if not vendor_summary_pdf.empty:
            vendor_fig = px.bar(
                vendor_summary_pdf,
                x="VendorName",
                y="Total",
                title="Top Vendors by Spend",
                text_auto=".2s",
            )
            vendor_fig.update_layout(
                title_font=dict(color="#1f4e79", size=16),
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                yaxis_tickprefix="$",
                xaxis_tickangle=-35,
            )
            st.plotly_chart(vendor_fig, use_container_width=True)

    if not trend_summary.empty:
        pdf_sections.append(("Spend Trend by Month", trend_summary.copy(), None))
    if not vendor_summary_pdf.empty:
        pdf_sections.append(("Top Vendors by Spend", vendor_summary_pdf.copy(), None))

    delivery_insights = []
    delivery_ready = False
    delivery_message = ""
    delivery_summary_pdf = pd.DataFrame()
    delivery_summary_display = pd.DataFrame()
    late_df = pd.DataFrame()
    late_account_summary_pdf = pd.DataFrame()
    late_requisitioner_summary_pdf = pd.DataFrame()
    late_pos_display = pd.DataFrame()
    late_pos_pdf = pd.DataFrame()
    on_time_percentage = 0.0
    late_percentage = 0.0

    if {"RecDate", "RequestDate"}.issubset(df_filtered.columns):
        df_delivery = df_filtered.dropna(subset=["RecDate", "RequestDate"]).copy()
        df_delivery["RecDate"] = pd.to_datetime(df_delivery["RecDate"], errors="coerce").dt.normalize()
        df_delivery["RequestDate"] = pd.to_datetime(df_delivery["RequestDate"], errors="coerce").dt.normalize()
        df_delivery.dropna(subset=["RecDate", "RequestDate"], inplace=True)

        if not df_delivery.empty:
            on_time_mask = df_delivery["RecDate"] <= df_delivery["RequestDate"]
            on_time_pos = df_delivery[on_time_mask]
            late_df = df_delivery[~on_time_mask]
            on_time_count = (
                on_time_pos["PONumber"].nunique() if "PONumber" in df_delivery.columns else len(on_time_pos)
            )
            late_count = (
                late_df["PONumber"].nunique() if "PONumber" in df_delivery.columns else len(late_df)
            )
            total_pos = on_time_count + late_count

            delivery_summary_pdf = pd.DataFrame(
                {
                    "Metric": [
                        "On-Time Orders",
                        "Late Orders",
                        "Total Orders",
                        "On-Time %",
                        "Late %",
                    ],
                    "Value": [
                        on_time_count,
                        late_count,
                        total_pos,
                        round((on_time_count / total_pos) * 100, 2) if total_pos else 0.0,
                        round((late_count / total_pos) * 100, 2) if total_pos else 0.0,
                    ],
                }
            )

            if total_pos > 0:
                on_time_percentage = (on_time_count / total_pos) * 100
                late_percentage = 100 - on_time_percentage
                delivery_ready = True
            else:
                delivery_message = "No purchase orders have both request and receive dates within the selected filters."

            delivery_summary_display = delivery_summary_pdf.copy()
            delivery_summary_display.loc[0:2, "Value"] = delivery_summary_display.loc[0:2, "Value"].map(
                lambda x: f"{int(x):,}"
            )
            delivery_summary_display.loc[3:4, "Value"] = delivery_summary_display.loc[3:4, "Value"].map(
                format_percentage
            )

            if not late_df.empty:
                late_df["Days Late"] = (late_df["RecDate"] - late_df["RequestDate"]).dt.days
                late_account_summary_pdf = (
                    late_df.groupby("Purchase Account")
                    .agg(
                        Late_Orders=("PONumber", "nunique"),
                        Late_Lines=("PONumber", "size"),
                        Avg_Days_Late=("Days Late", "mean"),
                        Max_Days_Late=("Days Late", "max"),
                        Late_Order_Value=("Total", "sum"),
                    )
                    .reset_index()
                )
                late_account_summary_pdf.sort_values(
                    by=["Late_Orders", "Late_Order_Value"], ascending=False, inplace=True
                )
                late_account_summary_pdf["Avg_Days_Late"] = late_account_summary_pdf["Avg_Days_Late"].round(1)
                late_account_summary_pdf.rename(
                    columns={
                        "Late_Orders": "Late Orders",
                        "Late_Lines": "Late Lines",
                        "Avg_Days_Late": "Avg Days Late",
                        "Max_Days_Late": "Max Days Late",
                        "Late_Order_Value": "Late Order Value",
                    },
                    inplace=True,
                )
                late_account_summary_pdf["Max Days Late"] = late_account_summary_pdf["Max Days Late"].fillna(0).astype(int)

                late_requisitioner_summary_pdf = (
                    late_df.groupby("Requisitioner")
                    .agg(
                        Late_Orders=("PONumber", "nunique"),
                        Late_Lines=("PONumber", "size"),
                        Avg_Days_Late=("Days Late", "mean"),
                        Max_Days_Late=("Days Late", "max"),
                        Late_Order_Value=("Total", "sum"),
                    )
                    .reset_index()
                )
                late_requisitioner_summary_pdf.sort_values(
                    by=["Late_Orders", "Late_Order_Value"], ascending=False, inplace=True
                )
                late_requisitioner_summary_pdf["Avg_Days_Late"] = late_requisitioner_summary_pdf["Avg_Days_Late"].round(1)
                late_requisitioner_summary_pdf.rename(
                    columns={
                        "Late_Orders": "Late Orders",
                        "Late_Lines": "Late Lines",
                        "Avg_Days_Late": "Avg Days Late",
                        "Max_Days_Late": "Max Days Late",
                        "Late_Order_Value": "Late Order Value",
                    },
                    inplace=True,
                )
                late_requisitioner_summary_pdf["Max Days Late"] = (
                    late_requisitioner_summary_pdf["Max Days Late"].fillna(0).astype(int)
                )

                detail_columns = [
                    "OrderDate",
                    "RequestDate",
                    "RecDate",
                    "PONumber",
                    "Purchase Account",
                    "Requisitioner",
                    "VendorName",
                    "Total",
                    "Days Late",
                ]
                available_detail_columns = [
                    col for col in detail_columns if col in late_df.columns
                ]
                late_pos_display = late_df[available_detail_columns].copy()
                for col in ["OrderDate", "RequestDate", "RecDate"]:
                    if col in late_pos_display.columns:
                        late_pos_display[col] = pd.to_datetime(
                            late_pos_display[col], errors="coerce"
                        ).dt.date
                if "Total" in late_pos_display.columns:
                    late_pos_display.rename(columns={"Total": "Total Amount"}, inplace=True)
                    late_pos_display["Total Amount"] = late_pos_display["Total Amount"].apply(
                        format_currency
                    )
                late_pos_display.sort_values(by="Days Late", ascending=False, inplace=True)
                late_pos_display.reset_index(drop=True, inplace=True)
                late_pos_pdf = late_pos_display.copy()

                if not late_account_summary_pdf.empty:
                    top_account = late_account_summary_pdf.iloc[0]
                    delivery_insights.append(
                        f"Purchase Account {top_account['Purchase Account']} has {int(top_account['Late Orders'])} late orders averaging {top_account['Avg Days Late']:.1f} days late."
                    )
                if not late_requisitioner_summary_pdf.empty:
                    top_req = late_requisitioner_summary_pdf.iloc[0]
                    delivery_insights.append(
                        f"{top_req['Requisitioner']} has {int(top_req['Late Orders'])} late orders with up to {int(top_req['Max Days Late'])} days delay."
                    )
        else:
            delivery_message = "No delivery performance data is available after removing rows with missing dates."
    else:
        delivery_message = "'RecDate' and/or 'RequestDate' columns are missing."

    if selected_requisitioner != "All":
        delivery_insights.append(
            f"Delivery metrics are scoped to requisitioner {selected_requisitioner}."
        )

    matrix_df = otd_matrix_by_account(df_filtered)

    account_value_summary_pdf = pd.DataFrame()
    if "Purchase Account" in df_filtered.columns:
        account_group = df_filtered.groupby("Purchase Account")
        account_value_summary_pdf = account_group["PONumber"].nunique().rename("Unique POs").to_frame()
        account_value_summary_pdf["Order Lines"] = account_group.size()
        account_value_summary_pdf["Total Value"] = account_group["Total"].sum()
        if "Amt" in df_filtered.columns:
            account_value_summary_pdf["Open Amount"] = account_group["Amt"].sum()
        else:
            account_value_summary_pdf["Open Amount"] = 0.0
        account_value_summary_pdf["Avg Order Value"] = (
            account_value_summary_pdf["Total Value"]
            / account_value_summary_pdf["Unique POs"].replace(0, pd.NA)
        )
        account_value_summary_pdf["Avg Order Value"] = account_value_summary_pdf["Avg Order Value"].fillna(0.0)
        account_value_summary_pdf.reset_index(inplace=True)
        account_value_summary_pdf.sort_values(by="Total Value", ascending=False, inplace=True)
        account_value_summary_pdf["Unique POs"] = account_value_summary_pdf["Unique POs"].astype(int)
        account_value_summary_pdf["Order Lines"] = account_value_summary_pdf["Order Lines"].astype(int)

    requisitioner_summary_pdf = pd.DataFrame()
    if "Requisitioner" in df_filtered.columns:
        req_group = df_filtered.groupby("Requisitioner")
        requisitioner_summary_pdf = req_group["PONumber"].nunique().rename("Unique POs").to_frame()
        requisitioner_summary_pdf["Order Lines"] = req_group.size()
        requisitioner_summary_pdf["Total Value"] = req_group["Total"].sum()
        if "Amt" in df_filtered.columns:
            requisitioner_summary_pdf["Open Amount"] = req_group["Amt"].sum()
        else:
            requisitioner_summary_pdf["Open Amount"] = 0.0
        requisitioner_summary_pdf["Avg Order Value"] = (
            requisitioner_summary_pdf["Total Value"]
            / requisitioner_summary_pdf["Unique POs"].replace(0, pd.NA)
        )
        requisitioner_summary_pdf["Avg Order Value"] = requisitioner_summary_pdf["Avg Order Value"].fillna(0.0)

        if not late_requisitioner_summary_pdf.empty:
            late_req_join = late_requisitioner_summary_pdf.set_index("Requisitioner")[
                ["Late Orders", "Late Lines", "Avg Days Late", "Max Days Late", "Late Order Value"]
            ]
            requisitioner_summary_pdf = requisitioner_summary_pdf.join(late_req_join, how="left")
        else:
            requisitioner_summary_pdf["Late Orders"] = 0
            requisitioner_summary_pdf["Late Lines"] = 0
            requisitioner_summary_pdf["Avg Days Late"] = 0.0
            requisitioner_summary_pdf["Max Days Late"] = 0.0
            requisitioner_summary_pdf["Late Order Value"] = 0.0

        requisitioner_summary_pdf.fillna(
            {
                "Late Orders": 0,
                "Late Lines": 0,
                "Avg Days Late": 0.0,
                "Max Days Late": 0.0,
                "Late Order Value": 0.0,
            },
            inplace=True,
        )
        requisitioner_summary_pdf.reset_index(inplace=True)
        requisitioner_summary_pdf.sort_values(by="Total Value", ascending=False, inplace=True)
        requisitioner_summary_pdf["Late Orders"] = requisitioner_summary_pdf["Late Orders"].astype(int)
        requisitioner_summary_pdf["Late Lines"] = requisitioner_summary_pdf["Late Lines"].astype(int)
        requisitioner_summary_pdf["Avg Days Late"] = requisitioner_summary_pdf["Avg Days Late"].round(1)
        requisitioner_summary_pdf["Max Days Late"] = requisitioner_summary_pdf["Max Days Late"].round(0)
        requisitioner_summary_pdf["Unique POs"] = requisitioner_summary_pdf["Unique POs"].astype(int)
        requisitioner_summary_pdf["Order Lines"] = requisitioner_summary_pdf["Order Lines"].astype(int)

    if delivery_ready and not delivery_summary_pdf.empty:
        pdf_sections.append(("On-Time Delivery Summary", delivery_summary_pdf.copy(), None))
    if not matrix_df.empty:
        pdf_sections.append(("On-Time Delivery by Purchase Account", matrix_df.copy(), None))
    if not account_value_summary_pdf.empty:
        pdf_sections.append(("Purchase Account Value Summary", account_value_summary_pdf.copy(), None))
    if not requisitioner_summary_pdf.empty:
        pdf_sections.append(("Requisitioner Value Summary", requisitioner_summary_pdf.copy(), None))
    if not late_account_summary_pdf.empty:
        pdf_sections.append(("Late Orders by Purchase Account", late_account_summary_pdf.copy(), None))
    if not late_requisitioner_summary_pdf.empty:
        pdf_sections.append(("Late Orders by Requisitioner", late_requisitioner_summary_pdf.copy(), None))
    if not late_pos_pdf.empty:
        pdf_sections.append(("Detailed Late Orders", late_pos_pdf.copy(), None))

    delivery_tab, accounts_tab, requisitioner_tab = st.tabs(
        ["Delivery Health", "Purchase Accounts", "Requisitioners"]
    )

    with delivery_tab:
        st.subheader("Delivery Health Overview")
        if delivery_ready:
            progress_html = f"""
                <div class="card" style='background-color: {"rgba(174, 213, 129, 0.55)"}; width: 100%;'>
                    <h3>On-Time Delivery</h3>
                    <p class="percentage-value">{format_percentage(on_time_percentage)}</p>
                    <div class="progress-wrapper">
                        <div class="progress-track">
                            <div class="progress-fill" style="width: {on_time_percentage:.2f}%;"></div>
                        </div>
                    </div>
                    <p>On-Time: {format_percentage(on_time_percentage)} | Late: {format_percentage(late_percentage)}</p>
                </div>
            """
            st.markdown(progress_html, unsafe_allow_html=True)
            if not delivery_summary_display.empty:
                st.dataframe(delivery_summary_display.set_index("Metric"), use_container_width=True)
        else:
            st.info(delivery_message)

        if not late_pos_display.empty:
            st.markdown("#### Late orders detail")
            st.dataframe(late_pos_display, use_container_width=True)

        if delivery_insights:
            st.markdown("#### Quick Insights")
            insights_html = " ".join(
                [f"<span class='insight-pill'>💡 {insight}</span>" for insight in delivery_insights]
            )
            st.markdown(insights_html, unsafe_allow_html=True)

    with accounts_tab:
        st.subheader("Purchase Accounts Overview")
        if not matrix_df.empty:
            matrix_display = matrix_df.copy()
            matrix_display["On-Time"] = matrix_display["On-Time"].astype(int)
            matrix_display["Late"] = matrix_display["Late"].astype(int)
            st.dataframe(matrix_display, use_container_width=True)
        if not account_value_summary_pdf.empty:
            account_display = account_value_summary_pdf.copy()
            for col in ["Total Value", "Open Amount", "Avg Order Value"]:
                if col in account_display.columns:
                    account_display[col] = account_display[col].apply(format_currency)
            st.markdown("#### Spend by purchase account")
            st.dataframe(account_display, use_container_width=True)
        if not late_account_summary_pdf.empty:
            late_account_display = late_account_summary_pdf.copy()
            if "Late Order Value" in late_account_display.columns:
                late_account_display["Late Order Value"] = late_account_display["Late Order Value"].apply(
                    format_currency
                )
            st.markdown("#### Late orders by purchase account")
            st.dataframe(late_account_display, use_container_width=True)

    with requisitioner_tab:
        st.subheader("Requisitioner Overview")
        if not requisitioner_summary_pdf.empty:
            requisitioner_display = requisitioner_summary_pdf.copy()
            for col in ["Total Value", "Open Amount", "Avg Order Value", "Late Order Value"]:
                if col in requisitioner_display.columns:
                    requisitioner_display[col] = requisitioner_display[col].apply(format_currency)
            st.dataframe(requisitioner_display, use_container_width=True)
        if not late_requisitioner_summary_pdf.empty:
            late_req_display = late_requisitioner_summary_pdf.copy()
            if "Late Order Value" in late_req_display.columns:
                late_req_display["Late Order Value"] = late_req_display["Late Order Value"].apply(format_currency)
            st.markdown("#### Late orders by requisitioner")
            st.dataframe(late_req_display, use_container_width=True)

    processing_end_time = time.time()
    total_processing_time = processing_end_time - processing_start_time
    st.markdown(
        f"<div class='processing-time'>Processed in {total_processing_time:.2f} seconds.</div>",
        unsafe_allow_html=True,
    )

    if st.button("Generate PDF Report"):
        with st.expander("ℹ️ PDF Report Information", expanded=False):
            st.info(
                """
                **PDF Report Contents:**
                - Key performance indicators and metrics
                - Data tables with current filter context
                - Late order summaries by account and requisitioner

                **Note:** Interactive charts are best explored in the dashboard.
                """
            )

        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            leftMargin=36,
            rightMargin=36,
            topMargin=36,
            bottomMargin=36,
        )
        elements = []
        styles = getSampleStyleSheet()
        heading_style = ParagraphStyle(
            name="Heading1",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=18,
            textColor=colors.HexColor("#1f4e79"),
            spaceAfter=12,
            alignment=TA_CENTER,
        )
        subheading_style = ParagraphStyle(
            name="Heading2",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=14,
            textColor=colors.HexColor("#2c5aa0"),
            spaceAfter=8,
        )
        bullet_style = ParagraphStyle(
            name="Bullet",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=11,
            leftIndent=18,
            bulletIndent=9,
            spaceAfter=4,
        )
        normal_style = ParagraphStyle(
            name="Normal",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=11,
            spaceAfter=6,
        )

        try:
            try:
                elements.append(ReportLabImage("TTU_LOGO.jpg", width=1.5 * inch, height=1.5 * inch))
                elements.append(Spacer(1, 12))
            except Exception:
                pass

            subtitle = f"Filters: Order dates {order_start_date} to {order_end_date}"
            if request_start_date and request_end_date:
                subtitle += f" | Request dates {request_start_date} to {request_end_date}"
            if selected_requisitioner != "All":
                subtitle += f" | Requisitioner: {selected_requisitioner}"
            if selected_purchase_account != "All":
                subtitle += f" | Account: {selected_purchase_account}"

            title = "TTU Purchase Orders Log Report"
            elements.append(Paragraph(title, heading_style))
            elements.append(Spacer(1, 6))
            elements.append(Paragraph(subtitle, normal_style))
            elements.append(Spacer(1, 18))

            toc_items = ["Key Performance Indicators"] + [title for title, _, _ in pdf_sections]
            elements.append(Paragraph("Table of Contents", subheading_style))
            for idx, item in enumerate(toc_items, 1):
                elements.append(Paragraph(f"{idx}. {item}", bullet_style))
            elements.append(Spacer(1, 18))

            elements.append(Paragraph("Key Performance Indicators", subheading_style))
            for metric_name, metric_info in metrics.items():
                text_content = (
                    metric_info.get(metric_name, "N/A")
                    .replace("<br/>", "<br />")
                    .replace("<br>", "<br />")
                )
                lines = text_content.split("<br />")
                if len(lines) > 1:
                    elements.append(Paragraph(f"<b>{metric_name}:</b>", normal_style))
                    for line in lines:
                        elements.append(Paragraph(line, bullet_style, bulletText="•"))
                else:
                    elements.append(Paragraph(f"<b>{metric_name}:</b> {text_content}", normal_style))
                elements.append(Spacer(1, 6))
            elements.append(Spacer(1, 12))

            for title_text, data, _ in pdf_sections:
                data_for_pdf = data.copy() if isinstance(data, pd.DataFrame) else data
                elements.append(Paragraph(title_text, subheading_style))
                elements.append(Spacer(1, 8))
                if isinstance(data_for_pdf, pd.DataFrame) and not data_for_pdf.empty:
                    date_cols = data_for_pdf.select_dtypes(
                        include=["datetime64[ns]", "datetime64[ns, UTC]"]
                    ).columns
                    for col in date_cols:
                        data_for_pdf[col] = data_for_pdf[col].astype(str)

                    table_data = [list(data_for_pdf.columns)] + data_for_pdf.values.tolist()
                    table = Table(table_data, repeatRows=1, hAlign="LEFT")
                    table.setStyle(
                        TableStyle(
                            [
                                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
                                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d0dae6")),
                                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                                ("FONTSIZE", (0, 0), (-1, -1), 9),
                                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#eef3f9")]),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                                ("TOPPADDING", (0, 0), (-1, -1), 6),
                            ]
                        )
                    )
                    elements.append(table)
                    elements.append(Spacer(1, 12))
                elif isinstance(data, pd.DataFrame) and data.empty:
                    elements.append(
                        Paragraph(
                            "No data available for this analysis.",
                            normal_style,
                        )
                    )
                    elements.append(Spacer(1, 12))

            doc.build(elements)
            pdf = buffer.getvalue()
            buffer.close()

            st.download_button(
                label="Download PDF Report",
                data=pdf,
                file_name="TTU_Purchase_Orders_Log_Report.pdf",
                mime="application/pdf",
            )
        except Exception as e:
            st.error(f"An error occurred while generating the PDF: {e}")


if __name__ == "__main__":
    main()
