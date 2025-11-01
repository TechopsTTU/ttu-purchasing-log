# app.py


# Standard library imports
import time
from io import BytesIO
import os

# Third-party imports
import streamlit as st
import pandas as pd
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
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Overall app styling */
    .stApp {
        background: white;
        font-family: 'Inter', sans-serif;
    }
    
    /* Main container */
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        max-width: 1200px;
    }
    
    /* Title styling */
    .title {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f4e79;
        text-align: left;
        text-transform: uppercase;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        margin: 0;
        line-height: 1.2;
    }
    
    /* Enhanced card styling */
    .card {
        background: linear-gradient(145deg, #ffffff 0%, #f8f9fa 100%);
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
        border: 1px solid rgba(255,255,255,0.2);
        min-height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin: 10px 0;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        backdrop-filter: blur(10px);
    }
    
    .card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 40px rgba(0,0,0,0.15);
    }
    
    .card h3 {
        margin-bottom: 10px;
        text-align: center;
        font-size: 0.9rem;
        font-weight: 600;
        color: #495057;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .card p {
        margin: 0;
        font-size: 1.4rem;
        text-align: center;
        font-weight: 700;
        white-space: pre-wrap;
        color: #1f4e79;
    }

    /* Enhanced percentage styling */
    .percentage-value {
        font-size: 3rem;
        color: #28a745;
        font-weight: 800;
        margin: 0;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }

    .progress-wrapper {
        width: 100%;
        margin-top: 10px;
    }

    .progress-track {
        width: 100%;
        height: 12px;
        border-radius: 999px;
        background: #e9ecef;
        overflow: hidden;
        box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
    }

    .progress-fill {
        height: 100%;
        border-radius: 999px;
        background: linear-gradient(90deg, #28a745 0%, #8bc34a 100%);
        transition: width 0.6s ease;
    }

    .insight-pill {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: rgba(31, 78, 121, 0.08);
        color: #1f4e79;
        padding: 8px 14px;
        border-radius: 999px;
        font-weight: 500;
        margin: 6px 6px 6px 0;
    }

    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, #1f4e79 0%, #2c5aa0 100%);
    }
    
    .css-1d391kg .stSelectbox label,
    .css-1d391kg .stDateInput label,
    .css-1d391kg .stCheckbox label,
    .css-1d391kg h1, .css-1d391kg h2, .css-1d391kg h3 {
        color: white !important;
        font-weight: 500;
    }
    
    /* Enhanced button styling */
    .stButton > button {
        background: linear-gradient(145deg, #1f4e79 0%, #2c5aa0 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.2rem;
        font-weight: 600;
        font-size: 0.9rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(31, 78, 121, 0.3);
    }
    
    .stButton > button:hover {
        background: linear-gradient(145deg, #2c5aa0 0%, #1f4e79 100%);
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(31, 78, 121, 0.4);
    }
    
    /* Section headers */
    .section-header {
        background: linear-gradient(90deg, #1f4e79 0%, #2c5aa0 100%);
        color: white;
        padding: 15px 20px;
        border-radius: 10px;
        margin: 20px 0 15px 0;
        font-weight: 600;
        font-size: 1.2rem;
        box-shadow: 0 4px 15px rgba(31, 78, 121, 0.2);
    }
    
    /* Enhanced table styling */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    /* File uploader styling */
    .stFileUploader {
        background: rgba(255,255,255,0.1);
        border-radius: 10px;
        padding: 20px;
        border: 2px dashed rgba(255,255,255,0.3);
    }
    
    /* Processing time styling */
    .processing-time {
        text-align: center;
        color: #6c757d;
        font-size: 0.8rem;
        font-style: italic;
        margin-top: 30px;
        padding: 10px;
        background: rgba(255,255,255,0.7);
        border-radius: 20px;
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
                <p style='color: #666; font-size: 16px; margin-top: -10px;'>
                    📊 Purchase Order Analytics & Reporting Dashboard
                </p>
            </div>
            """, 
            unsafe_allow_html=True
        )


# Callback function to reset date filters
def reset_dates():
    # Ensure that initial dates are stored in session state
    if (
        "initial_order_min_date" in st.session_state
        and "initial_order_max_date" in st.session_state
    ):
        st.session_state.order_start_date = st.session_state.initial_order_min_date
        st.session_state.order_end_date = st.session_state.initial_order_max_date
    if (
        "initial_request_min_date" in st.session_state
        and "initial_request_max_date" in st.session_state
    ):
        st.session_state.request_start_date = st.session_state.initial_request_min_date
        st.session_state.request_end_date = st.session_state.initial_request_max_date


# Callback function to reset selectbox filters
def reset_filters():
    st.session_state.selected_purchase_account = "All"
    st.session_state.selected_requisitioner = "All"


# Caching the data loading function
@st.cache_data
def load_and_process_data(uploaded_file):
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        # Rename 'Acct' to 'Purchase Account'
        df.rename(columns={"Acct": "Purchase Account"}, inplace=True)

        # Convert date columns to datetime and normalize time
        date_columns = [col for col in df.columns if "date" in col.lower()]
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.normalize()

        # Verify 'OrderDate' column
        if "OrderDate" in df.columns:
            df["OrderDate"] = pd.to_datetime(
                df["OrderDate"], errors="coerce"
            ).dt.normalize()
            df = df.dropna(subset=["OrderDate"])
        else:
            st.error("'OrderDate' column is missing.")
            return pd.DataFrame(), [], []

        # Filter rows where 'OrderDate' >= 2022-01-01
        df_filtered = df[df["OrderDate"] >= pd.to_datetime("2022-01-01")].copy()

        # Handle missing values in critical columns
        critical_columns = ["OrderDate", "PONumber", "Total"]
        df_filtered.dropna(subset=critical_columns, inplace=True)

        # Remove duplicates
        df_filtered.drop_duplicates(inplace=True)

        # Ensure numerical columns are correct
        numerical_columns = ["Total", "Amt", "QtyOrdered", "QtyRemaining"]
        for col in numerical_columns:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(
                    df_filtered[col], errors="coerce"
                ).fillna(0)

        # Format 'Purchase Account'
        if "Purchase Account" in df_filtered.columns:
            df_filtered = df_filtered.dropna(subset=["Purchase Account"])
            df_filtered["Purchase Account"] = (
                df_filtered["Purchase Account"].astype(int).astype(str).str.zfill(8)
            )
            df_filtered["Purchase Account"] = df_filtered[
                "Purchase Account"
            ].str.replace(r"(\d{4})(\d{4})", r"\1-\2", regex=True)
        else:
            st.error("'Purchase Account' column is missing.")

        # Rename 'QtyRemaining'
        df_filtered.rename(
            columns={"QtyRemaining": "qty on order/backordered"}, inplace=True
        )

        return df_filtered, [], date_columns
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame(), [], []


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


# Create index cards
def display_index_cards(metrics):
    for metric_name, metric_info in metrics.items():
        bg_color = metric_info.get("bg_color", "#FFFFFF")
        value = metric_info.get(metric_name, "N/A")
        value_formatted = value.replace("<br>", "<br/>")
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
    # Display title only (without logo) in main area
    st.markdown(
        """
        <div style='margin-top: 20px;'>
            <h1 class="title">TTU Purchase Orders Log</h1>
            <p style='color: #666; font-size: 16px; margin-top: -10px;'>
                📊 Purchase Order Analytics & Reporting Dashboard
            </p>
        </div>
        """, 
        unsafe_allow_html=True
    )

    with st.sidebar:
        # Clean logo display at the top
        if os.path.exists("TTU_LOGO.jpg"):
            # Display logo at full sidebar width
            st.image("TTU_LOGO.jpg", use_column_width=True)
            st.markdown("<br>", unsafe_allow_html=True)
        
        # Data Upload Section
        st.markdown("### 📁 Data Upload")
        uploaded_file = st.file_uploader(
            "Choose an Excel file here", 
            type=["xlsx"],
            help="Upload your purchase orders Excel file to begin analysis"
        )
        
        if uploaded_file:
            st.success(f"✅ File uploaded: {uploaded_file.name}")
        
        st.markdown("---")

    if uploaded_file:
        processing_start_time = time.time()
        df_filtered, _, date_columns = load_and_process_data(uploaded_file)

        if not df_filtered.empty:
            df_filtered = map_po_status(df_filtered)

            # Store the initial min and max dates
            initial_order_min_date = df_filtered["OrderDate"].min().date()
            initial_order_max_date = df_filtered["OrderDate"].max().date()
            if "RequestDate" in df_filtered.columns:
                initial_request_min_date = df_filtered["RequestDate"].min().date()
                initial_request_max_date = df_filtered["RequestDate"].max().date()
            else:
                initial_request_min_date = None
                initial_request_max_date = None

            # Store initial dates in session state
            st.session_state.initial_order_min_date = initial_order_min_date
            st.session_state.initial_order_max_date = initial_order_max_date
            if initial_request_min_date and initial_request_max_date:
                st.session_state.initial_request_min_date = initial_request_min_date
                st.session_state.initial_request_max_date = initial_request_max_date

            # Initialize session state for date inputs
            if "order_start_date" not in st.session_state:
                st.session_state.order_start_date = initial_order_min_date
            if "order_end_date" not in st.session_state:
                st.session_state.order_end_date = initial_order_max_date
            if (
                "request_start_date" not in st.session_state
                and initial_request_min_date is not None
            ):
                st.session_state.request_start_date = initial_request_min_date
            if (
                "request_end_date" not in st.session_state
                and initial_request_max_date is not None
            ):
                st.session_state.request_end_date = initial_request_max_date

            with st.sidebar:
                # Date Filters Section
                st.markdown("### 📅 Date Filters")
                
                with st.expander("📋 Order Date Range", expanded=True):
                    # Get current min and max dates from df_filtered
                    order_min_date = df_filtered["OrderDate"].min().date()
                    order_max_date = df_filtered["OrderDate"].max().date()

                    # Ensure default values are within min and max
                    order_start_default = st.session_state.order_start_date
                    order_end_default = st.session_state.order_end_date

                    if (
                        order_start_default < order_min_date
                        or order_start_default > order_max_date
                    ):
                        order_start_default = order_min_date
                    if (
                        order_end_default < order_min_date
                        or order_end_default > order_max_date
                    ):
                        order_end_default = order_max_date

                    order_start_date = st.date_input(
                        "Start Date",
                        order_start_default,
                        min_value=order_min_date,
                        max_value=order_max_date,
                        key="order_start_date_input",
                    )
                    order_end_date = st.date_input(
                        "End Date",
                        order_end_default,
                        min_value=order_min_date,
                        max_value=order_max_date,
                        key="order_end_date_input",
                    )

                    # Update session state
                    st.session_state.order_start_date = order_start_date
                    st.session_state.order_end_date = order_end_date

                    if order_start_date > order_end_date:
                        st.error("⚠️ End Date must be after Start Date.")

                # Filter df_filtered based on OrderDate range
                df_filtered = df_filtered[
                    (df_filtered["OrderDate"] >= pd.to_datetime(order_start_date))
                    & (df_filtered["OrderDate"] <= pd.to_datetime(order_end_date))
                ]

                # Request Date Filter
                if "RequestDate" in df_filtered.columns and not df_filtered.empty:
                    request_min_date = df_filtered["RequestDate"].min().date()
                    request_max_date = df_filtered["RequestDate"].max().date()

                    # Add a checkbox for enabling the RequestDate filter
                    filter_request_date = st.checkbox(
                        "🗓️ Enable Request Date Filter", value=False
                    )

                    # Show date pickers only if checkbox is selected
                    if filter_request_date:
                        with st.expander("📋 Request Date Range", expanded=True):
                            # Ensure default values are within min and max
                            request_start_default = st.session_state.request_start_date
                            request_end_default = st.session_state.request_end_date

                            if (
                                request_start_default < request_min_date
                                or request_start_default > request_max_date
                            ):
                                request_start_default = request_min_date
                            if (
                                request_end_default < request_min_date
                                or request_end_default > request_max_date
                            ):
                                request_end_default = request_max_date

                            request_start_date = st.date_input(
                                "Request Start Date",
                                request_start_default,
                                min_value=request_min_date,
                                max_value=request_max_date,
                                key="request_start_date_input",
                            )
                            request_end_date = st.date_input(
                                "Request End Date",
                                request_end_default,
                                min_value=request_min_date,
                                max_value=request_max_date,
                                key="request_end_date_input",
                            )

                            # Update session state
                            st.session_state.request_start_date = request_start_date
                            st.session_state.request_end_date = request_end_date

                            if request_start_date > request_end_date:
                                st.error("⚠️ Request End Date must be after Request Start Date.")

                            # Filter df_filtered based on RequestDate range
                            df_filtered = df_filtered[
                                (
                                    df_filtered["RequestDate"]
                                    >= pd.to_datetime(request_start_date)
                                )
                                & (
                                    df_filtered["RequestDate"]
                                    <= pd.to_datetime(request_end_date)
                                )
                            ]
                elif "RequestDate" in df_filtered.columns:
                    st.error("❌ No data available for the selected OrderDate range.")
                
                st.button("🔄 Reset Dates", on_click=reset_dates, type="secondary")
                
                st.markdown("---")
                
                # Category Filters Section
                st.markdown("### 🔍 Category Filters")

                # Filter by Purchase Account
                if "Purchase Account" in df_filtered.columns:
                    purchase_accounts = sorted(
                        df_filtered["Purchase Account"].dropna().unique().tolist()
                    )
                    options_purchase = ["All"] + purchase_accounts

                    # Initialize session state for purchase account
                    if "selected_purchase_account" not in st.session_state:
                        st.session_state.selected_purchase_account = "All"

                    # Determine the index for the selectbox
                    if st.session_state.selected_purchase_account in options_purchase:
                        index_purchase = options_purchase.index(
                            st.session_state.selected_purchase_account
                        )
                    else:
                        index_purchase = 0  # Default to 'All' if not found

                    selected_purchase_account = st.selectbox(
                        "Select Purchase Account",
                        options=options_purchase,
                        index=index_purchase,
                        key="selected_purchase_account",
                    )
                else:
                    selected_purchase_account = "All"
                    st.error("'Purchase Account' column is missing.")

                # Filter by Requisitioner
                if "Requisitioner" in df_filtered.columns:
                    requisitioners = sorted(
                        df_filtered["Requisitioner"].dropna().unique().tolist()
                    )
                    options_requisitioner = ["All"] + requisitioners

                    # Initialize session state for requisitioner
                    if "selected_requisitioner" not in st.session_state:
                        st.session_state.selected_requisitioner = "All"

                    # Determine the index for the selectbox
                    if st.session_state.selected_requisitioner in options_requisitioner:
                        index_requisitioner = options_requisitioner.index(
                            st.session_state.selected_requisitioner
                        )
                    else:
                        index_requisitioner = 0  # Default to 'All' if not found

                    selected_requisitioner = st.selectbox(
                        "Select Requisitioner",
                        options=options_requisitioner,
                        index=index_requisitioner,
                        key="selected_requisitioner",
                    )
                else:
                    selected_requisitioner = "All"
                    st.error("'Requisitioner' column is missing.")

                # Reset All Filters Button
                st.button("Reset All", on_click=reset_filters)

                # Apply filters
                if selected_requisitioner != "All":
                    df_filtered = df_filtered[
                        df_filtered["Requisitioner"] == selected_requisitioner
                    ]

                if selected_purchase_account != "All":
                    df_filtered = df_filtered[
                        df_filtered["Purchase Account"] == selected_purchase_account
                    ]

                # Handle case when no results are found after filtering
                if df_filtered.empty:
                    st.warning(
                        "No results found for the selected filters. Please adjust the filters and try again."
                    )
                    return  # Stop further processing

                # KPI Title
                kpi_title = f"Key Performance Indicators"

                st.header(kpi_title)

                # Calculate metrics
                metrics = {}

                # Total Open Orders Amt
                if "Amt" in df_filtered.columns and "POStatus" in df_filtered.columns:
                    total_open_orders_amt = df_filtered[
                        df_filtered["POStatus"] == "OPEN"
                    ]["Amt"].sum()
                    total_open_orders_amt_formatted = f"${total_open_orders_amt:,.2f}"
                    metrics["Total Open Orders Amt"] = {
                        "Total Open Orders Amt": total_open_orders_amt_formatted,
                        "bg_color": "#90CAF9",
                    }
                else:
                    metrics["Total Open Orders Amt"] = {
                        "Total Open Orders Amt": "$0.00",
                        "bg_color": "#90CAF9",
                    }

                # Total Orders Placed
                if "PONumber" in df_filtered.columns:
                    total_orders_placed = df_filtered["PONumber"].nunique()
                    metrics["Total Orders Placed"] = {
                        "Total Orders Placed": f"{total_orders_placed}",
                        "bg_color": "#FFCDD2",
                    }

                # Total Lines Ordered
                total_lines_ordered = df_filtered.shape[0]
                metrics["Total Lines Ordered"] = {
                    "Total Lines Ordered": f"{total_lines_ordered}",
                    "bg_color": "#C8E6C9",
                }

                # Most Expensive Order
                if {"Total", "PONumber", "VendorName", "Requisitioner"}.issubset(
                    df_filtered.columns
                ):
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
                        "bg_color": "#FFD54F",
                    }
                else:
                    metrics["Most Expensive Order"] = {
                        "Most Expensive Order": "N/A",
                        "bg_color": "#FFD54F",
                    }

                # Display Key Metrics
                display_index_cards(metrics)

            # Main Content Area

            # Under the title, if a requisitioner is selected, display last order details index card
            if selected_requisitioner != "All":
                # Get last order details for the selected requisitioner
                last_order = df_filtered.sort_values(
                    by="OrderDate", ascending=False
                ).head(1)
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
                        <div class="card" style='background-color: #BBDEFB; width: 100%;'>
                            <h3>Last Order for {selected_requisitioner}</h3>
                            <p>{last_order_info}</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            # Removed 'Detailed Analyses' subheader as per instruction

            pdf_elements = []

            # Only show On-Time Delivery Performance if no requisitioner is selected
            if selected_requisitioner == "All":
                # On-Time Delivery Performance
                if (
                    "RecDate" in df_filtered.columns
                    and "RequestDate" in df_filtered.columns
                ):
                    # Remove time from 'RecDate' and 'RequestDate' columns
                    df_filtered["RecDate"] = pd.to_datetime(
                        df_filtered["RecDate"]
                    ).dt.normalize()
                    df_filtered["RequestDate"] = pd.to_datetime(
                        df_filtered["RequestDate"]
                    ).dt.normalize()

                    on_time_pos = df_filtered[
                        df_filtered["RecDate"] <= df_filtered["RequestDate"]
                    ]
                    late_pos = df_filtered[
                        df_filtered["RecDate"] > df_filtered["RequestDate"]
                    ]
                    # Ensure the late purchase orders dataframe is always defined
                    late_df = late_pos.copy()
                    on_time_count = on_time_pos["PONumber"].nunique()
                    late_count = late_pos["PONumber"].nunique()
                    total_pos = on_time_count + late_count

                    if total_pos > 0:
                        on_time_percentage = (on_time_count / total_pos) * 100
                        late_percentage = 100 - on_time_percentage
                        delivery_ready = True
                    else:
                        delivery_message = (
                            "No purchase orders have both request and receive dates within the selected filters."
                        )
                        late_df = pd.DataFrame()

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
                                on_time_count if total_pos > 0 else 0,
                                late_count if total_pos > 0 else 0,
                                total_pos,
                                round(on_time_percentage, 2) if total_pos > 0 else 0.0,
                                round(late_percentage, 2) if total_pos > 0 else 0.0,
                            ],
                        }
                    )

                    delivery_summary_display = delivery_summary_pdf.copy()
                    delivery_summary_display.loc[0:2, "Value"] = delivery_summary_display.loc[
                        0:2, "Value"
                    ].map(lambda x: f"{int(x):,}")
                    delivery_summary_display.loc[3:4, "Value"] = delivery_summary_display.loc[
                        3:4, "Value"
                    ].map(format_percentage)

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
                        late_account_summary_pdf["Avg_Days_Late"] = late_account_summary_pdf[
                            "Avg_Days_Late"
                        ].round(1)
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
                        late_account_summary_pdf["Max Days Late"] = late_account_summary_pdf[
                            "Max Days Late"
                        ].fillna(0).astype(int)

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
                        late_requisitioner_summary_pdf["Avg_Days_Late"] = (
                            late_requisitioner_summary_pdf["Avg_Days_Late"].round(1)
                        )
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
                        late_requisitioner_summary_pdf["Max Days Late"] = late_requisitioner_summary_pdf[
                            "Max Days Late"
                        ].fillna(0).astype(int)

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
                        late_pos_display.sort_values(by="Days Late", ascending=False, inplace=True)
                        late_pos_display.reset_index(drop=True, inplace=True)
                        late_pos_pdf = late_pos_display.copy()
                        if "Total Amount" in late_pos_pdf.columns:
                            late_pos_pdf["Total Amount"] = late_pos_pdf["Total Amount"].apply(
                                format_currency
                            )
                        if "Total Amount" in late_pos_display.columns:
                            late_pos_display["Total Amount"] = late_pos_display["Total Amount"].apply(
                                format_currency
                            )

                        if not late_account_summary_pdf.empty:
                            top_account = late_account_summary_pdf.iloc[0]
                            delivery_insights.append(
                                f"Purchase Account {top_account['Purchase Account']} drives {int(top_account['Late Orders'])} late orders averaging {top_account['Avg Days Late']:.1f} days late."
                            )
                        if not late_requisitioner_summary_pdf.empty:
                            top_req = late_requisitioner_summary_pdf.iloc[0]
                            delivery_insights.append(
                                f"{top_req['Requisitioner']} has {int(top_req['Late Orders'])} late orders with up to {int(top_req['Max Days Late'])} days delay."
                            )
                else:
                    delivery_message = (
                        "No delivery performance data is available after removing rows with missing dates."
                    )
            else:
                delivery_message = "'RecDate' and/or 'RequestDate' columns are missing."

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
                account_value_summary_pdf["Avg Order Value"] = account_value_summary_pdf[
                    "Avg Order Value"
                ].fillna(0.0)
                account_value_summary_pdf.reset_index(inplace=True)
                account_value_summary_pdf.sort_values(
                    by="Total Value", ascending=False, inplace=True
                )
                account_value_summary_pdf["Unique POs"] = account_value_summary_pdf[
                    "Unique POs"
                ].astype(int)
                account_value_summary_pdf["Order Lines"] = account_value_summary_pdf[
                    "Order Lines"
                ].astype(int)

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
                requisitioner_summary_pdf["Avg Order Value"] = requisitioner_summary_pdf[
                    "Avg Order Value"
                ].fillna(0.0)

                if not late_requisitioner_summary_pdf.empty:
                    late_req_join = late_requisitioner_summary_pdf.set_index("Requisitioner")[
                        ["Late Orders", "Late Lines", "Avg Days Late", "Max Days Late", "Late Order Value"]
                    ]
                    requisitioner_summary_pdf = requisitioner_summary_pdf.join(
                        late_req_join, how="left"
                    )
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
                requisitioner_summary_pdf.sort_values(
                    by="Total Value", ascending=False, inplace=True
                )
                requisitioner_summary_pdf["Late Orders"] = requisitioner_summary_pdf[
                    "Late Orders"
                ].astype(int)
                requisitioner_summary_pdf["Late Lines"] = requisitioner_summary_pdf[
                    "Late Lines"
                ].astype(int)
                requisitioner_summary_pdf["Avg Days Late"] = requisitioner_summary_pdf[
                    "Avg Days Late"
                ].round(1)
                requisitioner_summary_pdf["Max Days Late"] = requisitioner_summary_pdf[
                    "Max Days Late"
                ].round(0)
                requisitioner_summary_pdf["Unique POs"] = requisitioner_summary_pdf[
                    "Unique POs"
                ].astype(int)
                requisitioner_summary_pdf["Order Lines"] = requisitioner_summary_pdf[
                    "Order Lines"
                ].astype(int)

            delivery_tab, accounts_tab, requisitioner_tab = st.tabs(
                ["Delivery Health", "Purchase Accounts", "Requisitioners"]
            )

            with delivery_tab:
                st.subheader("Delivery Health Overview")
                if delivery_ready:
                    progress_html = f"""
                        <div class="card" style='background-color: #AED581; width: 100%;'>
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
                        st.dataframe(
                            delivery_summary_display.set_index("Metric"),
                            use_container_width=True,
                        )
                        pdf_sections.append(
                            ("On-Time Delivery Summary", delivery_summary_pdf.copy(), None)
                        )

                    if delivery_insights:
                        st.markdown("#### Quick Insights")
                        insights_html = " ".join(
                            [f"<span class='insight-pill'>💡 {insight}</span>" for insight in delivery_insights]
                        )
                        st.markdown(insights_html, unsafe_allow_html=True)

            # Processing time
            processing_end_time = time.time()
            total_processing_time = processing_end_time - processing_start_time
            st.markdown(
                f"<p style='text-align: center; color: lightgray; font-size: 10px;'>Total processing time: {total_processing_time:.2f} seconds</p>",
                unsafe_allow_html=True,
            )

            # Generate PDF Report
            if st.button("Generate PDF Report"):
                # Show info about chart limitations
                with st.expander("ℹ️ PDF Report Information", expanded=False):
                    st.info("""
                    **PDF Report Contents:**
                    - Key performance indicators and metrics
                    - Data tables with all analysis results
                    - Charts may not be included if browser compatibility issues exist
                    
                    **Note:** For best chart viewing experience, use the interactive dashboard above.
                    """)
                
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
                elements = []
                styles = getSampleStyleSheet()
                # Custom styles for headings and bullet lists
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
                    textColor=colors.HexColor("#1976d2"),
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
                    # Add logo if available
                    try:
                        elements.append(ReportLabImage("TTU_LOGO.jpg", width=1.5*inch, height=1.5*inch))
                        elements.append(Spacer(1, 12))
                    except Exception:
                        pass

                    title = f"TTU Purchase Orders Log Report<br/>({order_start_date} to {order_end_date})"
                    elements.append(Paragraph(title, heading_style))
                    elements.append(Spacer(1, 18))

                    # Table of Contents
                    toc_items = [
                        "Key Performance Indicators",
                        "On-Time Delivery by Purchase Account",
                        "PO Count per Requisitioner by Order Date",
                    ]
                    elements.append(Paragraph("Table of Contents", subheading_style))
                    for idx, item in enumerate(toc_items, 1):
                        elements.append(Paragraph(f"{idx}. {item}", bullet_style))
                    elements.append(Spacer(1, 18))

                    # Key Performance Indicators section
                    elements.append(Paragraph("Key Performance Indicators", subheading_style))
                    # Use bullet points for metrics
                    for metric_name, metric_info in metrics.items():
                        text_content = (
                            metric_info.get(metric_name, "N/A")
                            .replace("<br/>", "<br />")
                            .replace("<br>", "<br />")
                        )
                        # Split into lines for bullet points
                        lines = text_content.split("<br />")
                        if len(lines) > 1:
                            elements.append(Paragraph(f"<b>{metric_name}:</b>", normal_style))
                            for line in lines:
                                elements.append(Paragraph(line, bullet_style, bulletText="•"))
                        else:
                            elements.append(Paragraph(f"<b>{metric_name}:</b> {text_content}", normal_style))
                        elements.append(Spacer(1, 6))
                    elements.append(Spacer(1, 12))

                    # Add analyses with headings and consistent table style
                    for title_text, data, _ in pdf_sections:
                        data_for_pdf = data.copy() if isinstance(data, pd.DataFrame) else data
                        elements.append(Paragraph(title_text, subheading_style))
                        elements.append(Spacer(1, 8))
                        if isinstance(data_for_pdf, pd.DataFrame) and not data_for_pdf.empty:
                            # Convert date columns to strings to avoid issues in PDF table
                            date_cols = data_for_pdf.select_dtypes(
                                include=["datetime64[ns]", "datetime64[ns, UTC]"]
                            ).columns
                            for col in date_cols:
                                data_for_pdf[col] = data_for_pdf[col].astype(str)

                            table_data = [list(data_for_pdf.columns)] + data_for_pdf.values.tolist()
                            t = Table(table_data, repeatRows=1, hAlign="LEFT")
                            t.setStyle(
                                TableStyle(
                                    [
                                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
                                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                                        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                                        ("FONTSIZE", (0, 0), (-1, -1), 9),
                                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
                                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                                    ]
                                )
                            )
                            elements.append(t)
                            elements.append(Spacer(1, 12))
                        elif isinstance(data, pd.DataFrame) and data.empty:
                            elements.append(
                                Paragraph(
                                    "No data available for this analysis.",
                                    normal_style,
                                )
                            )
                            elements.append(Spacer(1, 12))

                    # Build the PDF
                    doc.build(elements)
                    pdf = buffer.getvalue()
                    buffer.close()

                    # Download button
                    st.download_button(
                        label="Download PDF Report",
                        data=pdf,
                        file_name="TTU_Purchase_Orders_Log_Report.pdf",
                        mime="application/pdf",
                    )
                except Exception as e:
                    st.error(f"An error occurred while generating the PDF: {e}")
        else:
            st.write("Please upload an Excel file to proceed.")
    else:
        st.write("Please upload an Excel file to proceed.")


if __name__ == "__main__":
    main()
