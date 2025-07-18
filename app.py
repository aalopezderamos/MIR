import streamlit as st 
import pandas as pd
from datetime import datetime, timedelta
from datetime import datetime, date
import numpy as np
import io
from streamlit import column_config
import openai
from docx import Document
from docx.shared import Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import math
import re
from io import BytesIO
from xlsxwriter.utility import xl_col_to_name
import requests
from PIL import Image
import requests


# ==================== CONSTANTS ====================
CONFIG = {
    "export": {
        "ship_to": "SILVER EAGLE BEVERAGES LLC-TX (SAN ANTON)",
        "ship_to_num": "2225590",
        "load_type": "Truck",
        "status": "Open",
        "cols": [
            "#", "Ship To Location", "Ship To Location Number", "Delivery #", "Delivery PO",
            "Load Type", "Shipment Status", "Item Description", "Item SKU", "Order Qty",
            "MSO #", "Req Delivery Date", "Promised Ship Date", "Actual Ship Date"
        ]
    },
    "defaults": {
        "target_doh": 22,
        "order_qty": 0,
        "po_number": "",
        "days_to_add": 14
    },
    "colors": {
        "overview":     "#E2EFDA",  # table header
        "ros":          "#63BE7B",  # ROS column highlight
        "to_order":     "#FFFF00",  # “To Order” column highlight
        "po":           "#FFF2CC",
        "order_builder":"#E1CCF0"
    },
    "gpt_model": "gpt-4"
}

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="SEB Supplier Overview", 
    layout="wide"
)
st.markdown(
    f"""
    <style>
      /* 0) base background override */
      .stApp {{
        background-color: #ffffff !important;
      }}

      @media only screen and (orientation: portrait) {{
        /* stack columns */
        .stColumns {{
          display: flex !important;
          flex-direction: column !important;
        }}
        .stColumns > div {{
          width: 100% !important;
        }}

        /* shrink stDataFrame tables */
        .stDataFrame table {{
          font-size: 0.8em !important;
          transform-origin: top left;
        }}
        .stDataFrame th, .stDataFrame td {{
          padding: 4px 6px !important;
        }}

        /* shrink markdown tables */
        .stMarkdown table {{
          font-size: 0.8em !important;
          transform-origin: top left;
        }}
        .stMarkdown th, .stMarkdown td {{
          padding: 4px 6px !important;
        }}

        /* allow horizontal scroll */
        .stDataFrame > div, .stMarkdown table {{
          overflow-x: auto !important;
        }}
      }}

      /* ─── GLOBAL EXPANDER HEADER STYLES ───────────────────────────────── */
      div[data-testid="stExpander"] > button[data-testid="stExpanderHeader"][aria-expanded] {{
        background-color: {CONFIG['colors']['overview']} !important;
        padding: 10px !important;
        border-radius: 8px !important;
        margin-bottom: 0 !important;
        font-size: 1.25rem !important;
        font-weight: 600 !important;
        cursor: pointer !important;
      }}

      div[data-testid="stExpander"] > button[data-testid="stExpanderHeader"][aria-expanded] {{
        background-color: {CONFIG['colors']['order_builder']} !important;
      }}

      div[data-testid="stExpander"] > button[data-testid="stExpanderHeader"][aria-expanded] {{
        background-color: {CONFIG['colors']['po']} !important;
      }}

      div[data-testid="stExpander"] > button[data-testid="stExpanderHeader"][aria-expanded="true"] {{
        background-color: inherit !important;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)
# ─── Remote file URL ───────────────────────────────────────────────
GITHUB_RAW_URL = (
    "https://raw.githubusercontent.com/"
    "aalopezderamos/MIR/main/Master%20Incoming%20Report%20NEW.xlsm"
)

# ==================== HELPER FUNCTIONS ====================
def find_supplier_col(df: pd.DataFrame) -> str | None:
    """Find supplier column case-insensitively."""
    return next((c for c in df.columns if isinstance(c, str) and "supplier" in c.lower()), None)

def recompute_fields(df: pd.DataFrame) -> pd.DataFrame:
    """Recalculate all derived fields with proper formatting."""
    today = datetime.today().date()
    df = df.copy()
    # Ensure numeric fields
    numeric_cols = ["COH", "ROS", "Order Qty", "Projected Inventory", "Target Inventory", "To Order"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).round(2)
    # Core calculations
    df["COH"] = pd.to_numeric(df["On Floor"], errors="coerce").fillna(0)
    df["ROS"] = pd.to_numeric(df["Avg Weekday Depletion"], errors="coerce").fillna(0)
    if "Order Qty" not in df.columns and "Ordered" in df.columns:
        df["Order Qty"] = df["Ordered"]
    df["Order Qty"] = pd.to_numeric(df["Order Qty"], errors="coerce").fillna(0)
    df["Days Until Next Delivery"] = (
        (df["Next Delivery Date"] - pd.Timestamp(today)).dt.days
    ).clip(lower=0)
    df["Projected Inventory"] = (
        (df["COH"] + df["Order Qty"]) - df["ROS"] * df["Days Until Next Delivery"]
    ).clip(lower=0)
    df["Target Inventory"] = df["ROS"] * df["Target DOH"]
    df["To Order"] = (df["Target Inventory"] - df["Projected Inventory"]).round(0).astype(int)
    return df.round(2)

def build_export_df(data: dict) -> pd.DataFrame:
    """Prepare export-ready DataFrame from session state builder inputs,
       only exporting rows where Order Qty, PO Number and Delivery Date are all filled."""
    rows = []
    for key, df in data.items():
        if not key.endswith("_overview"):
            continue
        df = pd.DataFrame(df)
        # ensure the builder cols exist
        required = ["PO Number", "Order Qty", "Delivery Date"]
        if not set(required).issubset(df.columns):
            continue

        # only keep rows where none of the three is null or blank
        mask = (
            df[required]
            .notna()                               # not NaN
            .all(axis=1)
            & df["PO Number"].astype(str).str.strip().astype(bool)  # not blank string
        )

        for _, r in df.loc[mask].iterrows():
            rows.append({
                "#": len(rows) + 1,
                "Ship To Location": CONFIG["export"]["ship_to"],
                "Ship To Location Number": CONFIG["export"]["ship_to_num"],
                "Delivery #": r["PO Number"],
                "Delivery PO": r["PO Number"],
                "Load Type": CONFIG["export"]["load_type"],
                "Shipment Status": CONFIG["export"]["status"],
                "Item Description": r.get("Product Name", ""),
                "Item SKU": r.get("SPID", r.get("Product Num", "")),
                "Order Qty": round(float(r["Order Qty"]), 2),
                "MSO #": "",
                "Req Delivery Date": r["Delivery Date"],
                "Promised Ship Date": "",
                "Actual Ship Date": "",
            })

    return pd.DataFrame(rows, columns=CONFIG["export"]["cols"])

@st.cache_data(hash_funcs={io.BytesIO: lambda _: None})
def fetch_remote_file(url: str) -> io.BytesIO:
    """
    Download the Excel from GitHub and wrap in a BytesIO.
    Cached so we don’t re-download on every rerun.
    """
    resp = requests.get(url, timeout=10)
    resp.raise_for_status()
    return io.BytesIO(resp.content)

@st.cache_data(hash_funcs={io.BytesIO: lambda _: None})
def load_data(file: io.BytesIO) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Load and parse the three sheets from the given Excel file.
    Cached on the BytesIO contents so repeated runs are fast.
    """
    try:
        supplier_df = pd.read_excel(file, sheet_name="Supplier Info")
        po_df       = pd.read_excel(file, sheet_name="PO Info")
        overview_df = pd.read_excel(
            file,
            sheet_name="Overview",
            skiprows=4,
            header=0
        )

        # strip whitespace from all column names
        for df in (supplier_df, po_df, overview_df):
            df.columns = df.columns.str.strip()

        # ensure “Product Num” is an integer column
        if "Product Num" in overview_df.columns:
            overview_df["Product Num"] = (
                pd.to_numeric(overview_df["Product Num"], errors="coerce")
                  .round(0)
                  .astype("Int64")
            )

        # if there’s a 12th column, treat it as OOS Risk
        if overview_df.shape[1] >= 12:
            overview_df["OOS Risk"] = overview_df.iloc[:, 11]

        return supplier_df, po_df, overview_df

    except Exception as e:
        st.error(f"Error loading file: {e}")
        # on error, return three empty DataFrames so your app won’t crash
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

@st.cache_data(hash_funcs={io.BytesIO: lambda _: None})
def load_shortcode_data(file: io.BytesIO) -> pd.DataFrame:
    """Load the 'Short Code Data' sheet once (cached)."""
    df = pd.read_excel(file, sheet_name="Short Code Data")
    df.columns = df.columns.str.strip()
    return df

def display_min_order_progress(supplier_df, supplier_col, supplier, font_size=18):
    """Display the minimum order progress using a truck graphic as the bar, filling the trailer
    area (from X=1 to X=160 and Y=5 to Y=35) with the beer image, plus a marker and centered percentage text.
    font_size: size in pixels for the percentage text"""
    try:
        # Compute percentage
        min_order_series = supplier_df[supplier_df[supplier_col] == supplier]["Minimum Order Met?"]
        pct = 0
        if not min_order_series.empty:
            pct_str = str(min_order_series.iloc[0])
            pct = float(pct_str.strip('%'))/100 if '%' in pct_str else float(pct_str)/100

        # Weight text
        weight_series = supplier_df[supplier_df[supplier_col] == supplier][["Total Order Weight", "Minimum Order Weight"]]
        weight_str = None
        if not weight_series.empty:
            total = weight_series.iloc[0]["Total Order Weight"]
            minimum = weight_series.iloc[0]["Minimum Order Weight"]
            weight_str = f"{int(total):,}lbs/{int(minimum):,}lbs"

        # Image URLs
        truck_url = "https://raw.githubusercontent.com/aalopezderamos/Egg/main/Truck.png"
        beer_url  = "https://raw.githubusercontent.com/aalopezderamos/Egg/main/Beer.png"

        # Trailer fill parameters
        trailer_start = 1       # px offset of trailer region start
        trailer_width = 160     # px width corresponding to 100%
        fill_top = 1            # px from top of truck image where trailer begins
        fill_bottom = 35        # px from top of truck image where trailer ends
        fill_height = fill_bottom - fill_top  # computed trailer height in px

        # Compute fill width in px
        fill_width = int(pct * trailer_width)

        # Build HTML with layering and z-index
        html = f"""
        <div style='position:relative; display:inline-block;'>
          <!-- Base truck image on top layer -->
          <img src='{truck_url}' alt='truck' style='display:block; position:relative; z-index:2;' />
          <!-- Trailer fill container behind truck -->
          <div style='position:absolute; top:{fill_top}px; left:{trailer_start}px; width:{fill_width}px; height:{fill_height}px; overflow:hidden; z-index:1;'>
            <img src='{beer_url}' alt='fill' style='width:{trailer_width}px; height:{fill_height}px; object-fit:cover; display:block;' />
          </div>
          <!-- Marker line above fill -->
          <div style='position:absolute; top:{fill_top}px; left:{trailer_start + fill_width}px; width:2px; height:{fill_height}px; background-color:black; z-index:3;'></div>
          <!-- Centered percentage text on top -->
          <div style='position:absolute; top:-10px; left:-10px; width:100%; height:100%; display:flex; align-items:center; justify-content:center; font-weight:bold; font-size:{font_size}px; color:black; text-shadow:1px 1px 2px rgba(255,255,255,0.7); z-index:4;'>
            {pct:.0%}
          </div>
        </div>"""
        # Append weight text to side
        if weight_str:
            html += f"<span style='margin-left:12px; font-weight:bold;'>{weight_str}</span>"

        st.markdown(html, unsafe_allow_html=True)
        return pct
    except Exception:
        st.markdown("<div style='color:gray;'>Progress unavailable</div>", unsafe_allow_html=True)
        return 0

def display_overview_and_builder(supplier, overview_df, overview_col):
    """Display the overview and order builder sections in a single table, with foldable headers and styled expanders."""
    key = f"{supplier}_overview"

    # 1. Load or initialize session-state DataFrame
    if key in st.session_state:
        df = st.session_state[key].copy()
    else:
        df = overview_df[overview_df[overview_col] == supplier].copy()
        if df.empty:
            st.info("No data for this supplier.")
            return

        today = datetime.today().date()
        defaults = CONFIG["defaults"]
        next_delivery = today + timedelta(days=defaults["days_to_add"])
        next_str = next_delivery.strftime("%m/%d/%Y")

        # initialize builder defaults
        for col, default in [
            ("Target DOH", defaults["target_doh"]),
            ("Order Qty", defaults["order_qty"]),
            ("PO Number", defaults["po_number"]),
            ("Delivery Date", next_str),
            ("Next Delivery Date", next_str),
        ]:
            if col in df.columns:
                df[col] = df[col].replace("", default)
            else:
                df[col] = default

        # convert & compute fields
        df["Next Delivery Date"] = pd.to_datetime(df["Next Delivery Date"], errors="coerce")
        df["Delivery Date"] = pd.to_datetime(df["Delivery Date"], errors="coerce")
        df["Target DOH"] = pd.to_numeric(df["Target DOH"], errors="coerce").fillna(defaults["target_doh"])
        df["Order Qty"] = pd.to_numeric(df["Order Qty"], errors="coerce").fillna(defaults["order_qty"])
        df = recompute_fields(df)

        st.session_state[key] = df.copy()

    # 3. Overview expander
    with st.expander("Overview", expanded=False):
        disp_df = st.session_state[key].reset_index(drop=True)
        disp_df.index = disp_df.index + 1

        # hyperlink Product Num
        if "Product Num" in disp_df.columns:
            disp_df["Product Num"] = (
                pd.to_numeric(disp_df["Product Num"], errors="coerce")
                  .round(0)
                  .fillna(0)
                  .astype(int)
                  .astype(str)
            ).apply(lambda x: (
                f'<a href=\"https://sbsabs.encompass8.com/Home?DashboardID=100018&ProductID={x}\" '
                f'target=\"_blank\">{x}</a>'
            ))

        # style rules for ROS, Product Name, OOS Risk
        disp_df["ROS"] = disp_df["ROS"].round(1)
        ros_min, ros_max = disp_df["ROS"].min(), disp_df["ROS"].max()
        ros_range = ros_max - ros_min or 1
        green_rgb = (99, 190, 123)
        ros_colors = disp_df["ROS"].apply(
            lambda v: (
                f"background-color: rgb({int(255 - (255 - green_rgb[0]) * (v - ros_min) / ros_range)},"
                f"{int(255 - (255 - green_rgb[1]) * (v - ros_min) / ros_range)},"
                f"{int(255 - (255 - green_rgb[2]) * (v - ros_min) / ros_range)})"
            ) if v > 0 else ""
        )

        def product_name_color(row):
            doh, coh, ros = row["Days of Inventory"], row["COH"], row["ROS"]
            if (pd.isna(ros) or ros == 0) and (pd.isna(coh) or coh == 0):
                return "color: red"
            if doh >= 30:
                return "background-color: #9BC2E5"
            if doh > 16:
                return "background-color: #C6EFCE"
            if doh > 10:
                return "background-color: #FFEB9C"
            return "background-color: #FFC7CE"

        def oos_risk_color(val):
            return "color: red" if val > 0 else "color: green"

        overview_cols = ["Product Num", "Product Name", "COH", "ROS",
                         "Days of Inventory", "OOS Risk"]
        styled = (
            disp_df[overview_cols]
              .style
              .apply(lambda _: ros_colors, subset=["ROS"])
              .apply(lambda row: [product_name_color(row) if col == "Product Name" else "" for col in overview_cols], axis=1)
              .applymap(oos_risk_color, subset=["OOS Risk"])
              .format({"ROS": "{:.1f}", "COH": "{:.0f}",
                       "Days of Inventory": "{:.1f}", "OOS Risk": "{:.0f}"})
              .set_table_styles([
                  {"selector": "th.row_heading, td.row_heading", "props": [("font-weight", "normal")]}
              ])
        )
        html = styled.to_html(index=True, index_names=False, escape=False)
        st.markdown(html, unsafe_allow_html=True)

    # 5. Order Builder expander
    with st.expander("Order Builder", expanded=False):
        st.markdown(
            """
            <style>
              .ag-cell { white-space: normal !important; line-height: 1.3 !important; }
              .ag-cell[col-id=\"Product Name\"], .ag-header-cell-label[col-id=\"Product Name\"] {
                white-space: nowrap !important;
              }
              .ag-root-wrapper { width: 100% !important; }
              .ag-cell[col-id=\"To Order\"], .ag-header-cell[col-id=\"To Order\"] {
                background-color: yellow !important;
              }
            </style>
            """,
            unsafe_allow_html=True
        )

        editor_df = disp_df.copy()
        round_cols = ["Target DOH", "Days Until Next Delivery",
                      "Projected Inventory", "Target Inventory",
                      "To Order", "Order Qty"]
        editor_df[round_cols] = editor_df[round_cols].round(0)

        cols = ["Product Name", "Target DOH", "Next Delivery Date",
                "Days Until Next Delivery", "Projected Inventory",
                "Target Inventory", "To Order", "Order Qty",
                "PO Number", "Delivery Date"]
        edited = st.data_editor(
            editor_df[cols],
            key=f"{supplier}_builder",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Product Name": column_config.Column(disabled=True),
                "Target DOH": column_config.NumberColumn(min_value=0, max_value=100,
                                                         step=1, format="%.0f"),
                "Next Delivery Date": column_config.DateColumn(format="MM/DD/YYYY"),
                "Days Until Next Delivery": column_config.NumberColumn(disabled=True, format="%.0f"),
                "Projected Inventory": column_config.NumberColumn(disabled=True, format="%.0f"),
                "Target Inventory": column_config.NumberColumn(disabled=True, format="%.0f"),
                "To Order": column_config.NumberColumn(disabled=True, format="%.0f"),
                "Order Qty": column_config.NumberColumn(format="%.0f"),
                "PO Number": column_config.TextColumn(),
                "Delivery Date": column_config.DateColumn(format="MM/DD/YYYY"),
            },
            num_rows="dynamic"
        )

        if st.button("Refresh Calculations", key=f"{supplier}_refresh"):
            df.update(edited)
            df = recompute_fields(df)
            st.session_state[key] = df.copy()
            st.rerun()

        # Store for export
        st.session_state.setdefault("export_data", {})[key] = df.copy()

def display_po_and_shipments(supplier, po_df, po_col, overview_df, overview_col):
    """Display POs, Shipments, and Notes for a given supplier in separate tabs, with an expander."""
    # initialize defaults
    po_count, po_numbers = 0, []

    # Foldable section using Streamlit expander
    with st.expander("POs, Shipments & Notes", expanded=False):
        tab_po, tab_ship, tab_notes = st.tabs(["🖨️ POs", "🚚 Shipments", "🗒️ Notes"])

        # ——— POs Tab —————————————————————————————————————————————————
        with tab_po:
            mask = (
                (po_df[po_col] == supplier)
                & po_df.get("PO Num", pd.Series()).notna()
                & po_df["PO Num"].astype(str).str.strip().astype(bool)
            )
            supplier_pos = po_df.loc[mask]

            if supplier_pos.empty:
                st.info("No open POs for this supplier.")
            else:
                po_count = supplier_pos["PO Num"].nunique()
                po_numbers = sorted(
                    supplier_pos["PO Num"].dropna().astype(str).unique()
                )
                st.markdown(f"**{po_count} Open POs**")
                for po_num in po_numbers:
                    # build link to PO dashboard
                    purchase_series = supplier_pos.loc[
                        supplier_pos["PO Num"] == po_num, "Purchase ID"
                    ].dropna()
                    if not purchase_series.empty:
                        raw_pid = purchase_series.iloc[0]
                        pid = str(int(raw_pid)) if isinstance(raw_pid, (float, np.floating)) else str(raw_pid).strip()
                    else:
                        pid = ""
                    link = (
                        f"https://sbsabs.encompass8.com/Home?DashboardID=168160&KeyValue={pid}"
                    )
                    st.markdown(f"[🧾 PO #{po_num}]({link})", unsafe_allow_html=True)

                    # detail table per PO
                    det_cols = [c for c in ["Purchase ID", "Receive Date", "Product", "Ordered"]
                                if c in supplier_pos.columns]
                    det = supplier_pos[supplier_pos["PO Num"] == po_num][det_cols].copy()
                    if "Receive Date" in det.columns:
                        det["Receive Date"] = (
                            pd.to_datetime(det["Receive Date"], errors="coerce")
                              .dt.strftime("%m/%d/%Y").fillna("")
                        )
                    det = det.reset_index(drop=True)
                    st.dataframe(det, use_container_width=True)

        # ——— Shipments Tab ——————————————————————————————————————————————————
        with tab_ship:
            desired = ["Product Num", "Product Name", "First Shipment", "Second Shipment", "Third Shipment"]
            cols = [c for c in desired if c in overview_df.columns]

            if overview_df.empty or overview_col is None or not {"Product Num", "Product Name"}.issubset(cols):
                st.info("No shipment information available for this supplier.")
            else:
                ship_df = overview_df.loc[
                    overview_df[overview_col] == supplier,
                    cols
                ].copy()

                date_cols = [c for c in ["First Shipment", "Second Shipment", "Third Shipment"] if c in ship_df.columns]
                if date_cols:
                    ship_df = ship_df.dropna(how="all", subset=date_cols)

                if ship_df.empty:
                    st.info("No upcoming shipments recorded.")
                else:
                    det = det.reset_index(drop=True)
                    st.dataframe(det, use_container_width=True)

        # ——— Notes Tab ——————————————————————————————————————————————————
        with tab_notes:
            st.markdown(
                """
                <div style='background-color:#FFF8DC; padding:10px; border-radius:8px;'>
                  <strong>Notes & Next Steps:</strong>
                </div>
                """,
                unsafe_allow_html=True
            )
            note_key = f"{supplier}_notes"
            default_note = "Notes written here will appear on export"

            # pull existing or fall back to default
            current = st.session_state.get(note_key, "")
            text = st.text_area(
                label="",
                value=current if current else default_note,
                height=150,
                key=note_key
            )

    return po_count, po_numbers

def display_shortcode(supplier: str,
                      shortcode_df: pd.DataFrame):
    # 1) figure out which column holds “supplier” in the shortcode sheet
    code_sup_col = find_supplier_col(shortcode_df)
    if not code_sup_col:
        st.error("❌ Could not locate a 'Supplier' column in Short Code Data.")
        return

    # 2) now filter by that real column
    df = shortcode_df.loc[
        shortcode_df[code_sup_col] == supplier,
        [
            "Product Name", "Product ID", "Supplier Family",
            "Code Date", "Inventory", "Daily Rate of Sales",
            "Days on Hand", "Shelf Life Remaining",
            "Shelf Life Days", "Expiration Date", "Receive Date",
        ]
    ].copy()

    if df.empty:
        st.info("No short‐code data for this supplier.")
        return

    # 3) format your dates
    for dt in ("Code Date", "Expiration Date", "Receive Date"):
        if dt in df:
            df[dt] = (
                pd.to_datetime(df[dt], errors="coerce")
                  .dt.strftime("%m/%d/%Y")
            )
    st.subheader("Short Code Data")
    st.dataframe(df, use_container_width=True)

def display_supplier(
    supplier: str,
    supplier_df: pd.DataFrame,
    po_df: pd.DataFrame,
    overview_df: pd.DataFrame,
    supplier_col: str,
    po_col: str,
    overview_col: str,
    shortcode_df: pd.DataFrame):
    
    """Display all information for a single supplier, with four full-width sections under Details."""
    # ─── Header with logo + select checkbox ─────────────────────────
    col1, col2 = st.columns([1, 10])
    with col1:
        selected = st.checkbox("", key=f"select_{supplier}", value=False)
    with col2:
        logo_url = None
        if "Logos" in supplier_df.columns:
            logos = (
                supplier_df.loc[supplier_df[supplier_col] == supplier, "Logos"]
                .dropna().astype(str)
            )
            if not logos.empty:
                logo_url = logos.iloc[0].strip()
        if logo_url:
            st.markdown(
                f"""
                <div style="display: flex; align-items: center;">
                  <img src="{logo_url}" alt="{supplier} logo"
                       style="width:72px; height:auto; margin-right:8px;" />
                  <span style="font-size:26px; font-weight:600;">{supplier}</span>
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.markdown(f"<h3 style='font-size:26px;'>🏬 {supplier}</h3>",
                        unsafe_allow_html=True)

    # track selection
    st.session_state.setdefault("selected_suppliers", {})[supplier] = selected
    if not selected:
        return None, None, None, None

    # ─── Minimum order progress ────────────────────────────────────
    min_order_pct = display_min_order_progress(supplier_df, supplier_col, supplier)

    # prepare defaults for PO info
    po_count, po_numbers = 0, []

    # ─── Details expander ──────────────────────────────────────────
    with st.expander("Details", expanded=False):
        # Overview & Order Builder (full-width)
        display_overview_and_builder(supplier, overview_df, overview_col)

        # POs, Shipments & Notes (full-width) and capture outputs
        po_count, po_numbers = display_po_and_shipments(
            supplier, po_df, po_col, overview_df, overview_col
        )

        # Short Code Data (full-width)
        with st.expander("Short Code Data", expanded=False):
            display_shortcode(supplier, shortcode_df)

    # ─── Build metrics for summary/export ─────────────────────────
    overview_data = st.session_state.get(f"{supplier}_overview", pd.DataFrame())
    items_under_10 = (
        (overview_data["Days of Inventory"] < 10).sum()
        if not overview_data.empty else 0
    )
    oos_risks = (
        overview_data[overview_data["OOS Risk"] > 0][
            ["Product Num", "Product Name", "OOS Risk"]
        ].values.tolist()
        if not overview_data.empty else []
    )

    # return tuple for summary and export
    return min_order_pct, items_under_10, oos_risks, (po_count, po_numbers)

def _export_report_to_excel_bytes(
    supplier_data: dict[str, tuple],
    overview_df: pd.DataFrame,
    overview_col: str,
    supplier_logo_urls: dict[str, str],
    supplier_manager_map: dict[str, str],
    supplier_order_day_map: dict[str, str],  # <-- new mapping for Order Day
    po_df: pd.DataFrame,
    po_col: str,
) -> BytesIO:
    """
    Build an Excel workbook in memory containing:
      - One sheet per supplier in `supplier_data`
      - Logos, summary, overview tables, and formatting
      - A POs table inserted two rows below each overview, with columns reordered
      - Tab color set based on Brand Manager via supplier_manager_map
      - Gridlines turned off and zoom set to 80% on each sheet
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        # ─── Tab‐color palette by Brand Manager ────────────────────────────
        manager_colors = {
            "Tanya Marthes":       "#40E0D0",  # Turquoise
            "Mandiee Franco-Neff": "#FFFF00",  # Yellow
            "Mark Navarro":        "#6FD66F",  # Green
            "Dennis Diem":         "#F2C063",  # Orange
            "Ryan Mulle":          "#F47272",  # Red
        }

        # ─── Excel formats ─────────────────────────────────────────────────
        sup_fmt           = wb.add_format({"bold": True, "font_size": 18})
        sec_hdr_fmt       = wb.add_format({"bold": True})
        blue_fmt          = wb.add_format({"font_color": "#0070C0"})
        red_fmt           = wb.add_format({"font_color": "#FF0000"})
        two_dec_fmt       = wb.add_format({"num_format": "0.00", "border": 1})
        int_fmt           = wb.add_format({"num_format": "0",    "border": 1})
        date_fmt          = wb.add_format({"num_format": "mm/dd/yyyy", "border": 1})
        centered_wrap_fmt = wb.add_format({"align": "center", "valign": "vcenter", "text_wrap": True})
        due_fmt           = wb.add_format({"bold": True, "font_size": 18, "font_color": "#00B050"})
        hyperlink_fmt     = wb.add_format({"align": "center","font_color": "#0563C0", "underline": 1, "border": 1})
        to_order_num_fmt  = wb.add_format({"align": "center","num_format": "0", "bg_color": "#FFFF00", "border": 1})

        header_fmt1 = wb.add_format({"bold": True, "bg_color": "#002060", "font_color": "#FFFFFF", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        header_fmt2 = wb.add_format({"bold": True, "bg_color": "#7030A0", "font_color": "#FFFFFF", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        header_fmt3 = wb.add_format({"bold": True, "bg_color": "#C5D9F1", "font_color": "#000000", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})

        thin_border_fmt    = wb.add_format({"border": 1})
        thick_top_fmt      = wb.add_format({"top": 5})
        thick_bottom_fmt   = wb.add_format({"bottom": 5})
        thick_left_fmt     = wb.add_format({"left": 5})
        thick_right_fmt    = wb.add_format({"right": 5})
        to_order_date_fmt  = wb.add_format({"num_format": "mm/dd/yyyy", "bg_color": "#FFFF00", "border": 1})
        oos_bad_fmt        = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1})
        oos_good_fmt       = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "border": 1})
        neutral_fmt        = wb.add_format({"bg_color": "#FFEB9C", "border": 1})
        days_hi_fmt        = wb.add_format({"bg_color": "#9BC2E5", "border": 1})
        depletion_zero_fmt = wb.add_format({"font_color": "#FF0000", "border": 1})

        pixel_widths      = [223,91,97,70,93,70,64,69,57,77,77,68,68,68,225,225,225]
        two_dec_cols      = {4,5}
        int_cols          = {11,12,13}
        next_delivery_idx = 9
        to_order_idx      = 13
        projected_oos_idx = 7

        # ─── Build one sheet per supplier ───────────────────────────────────
        for supplier, data in supplier_data.items():
            sheet = supplier[:31]
            ws = wb.add_worksheet(sheet)
            writer.sheets[sheet] = ws

            # hide gridlines & set zoom
            ws.hide_gridlines(2)
            ws.set_zoom(80)

            # tab color by manager
            manager = supplier_manager_map.get(supplier)
            if manager_colors.get(manager):
                ws.set_tab_color(manager_colors[manager])

            # 1) Supplier title
            ws.write(0, 0, supplier, sup_fmt)

            # 2) Logo at P1 (with error handling)
            logo_url = supplier_logo_urls.get(supplier, "")
            if logo_url:
                try:
                    resp = requests.get(logo_url, timeout=5)
                    resp.raise_for_status()
                    if not resp.headers.get("Content-Type", "").startswith("image/"):
                        raise ValueError("URL did not return an image")
                    img = Image.open(BytesIO(resp.content))
                    w, _ = img.size
                    scale = 220 / w
                    bio = BytesIO(resp.content)
                    ws.insert_image("P1", "logo.png", {"image_data": bio, "x_scale": scale, "y_scale": scale})
                except Exception as e:
                    print(f"Warning: could not load logo for {supplier}: {e}")

            # 3) Due date as Order Day
            order_day = supplier_order_day_map.get(supplier, "").upper()
            if order_day:
                ws.write(0, 16, f"DUE {order_day}", due_fmt)

            # 4) Summary text
            summary_text = generate_chatgpt_summary({supplier: data})
            lines = [ln for ln in summary_text.splitlines() if ln.strip()]
            in_oos = next_note = next_po = False
            for idx, raw in enumerate(lines[1:], start=1):
                txt = raw.replace("**", "").strip()
                if txt.startswith("Notes & Next Steps:"):
                    ws.write(idx, 0, txt, sec_hdr_fmt); next_note=True; continue
                if next_note:
                    ws.write(idx, 0, txt, blue_fmt); next_note=False; continue
                if txt.startswith("PO Recommendation:"):
                    ws.write(idx, 0, txt, sec_hdr_fmt); next_po=True; continue
                if next_po:
                    ws.write(idx, 0, txt, blue_fmt); next_po=False; continue
                if txt.startswith("OOS Risk:"):
                    ws.write(idx, 0, txt, sec_hdr_fmt); in_oos=True; continue
                if in_oos:
                    if txt.startswith("Order Builder Table:"):
                        ws.write(idx, 0, txt, sec_hdr_fmt); in_oos=False
                    else:
                        ws.write(idx, 0, txt, red_fmt)
                    continue
                if txt.endswith(":"):
                    ws.write(idx, 0, txt, sec_hdr_fmt)
                else:
                    ws.write(idx, 0, txt)

            # 5) Overview table
            start_row = len(lines) + 2
            desired_cols = [
                "Product Name", "Product Num", "SPID", "On Floor",
                "Avg Weekday Depletion", "Days of Inventory",
                "Ordered", "Projected OOS Risk", "Target DOH",
                "Next Delivery", "Days Until Next Delivery",
                "Projected Inventory", "Target Inventory",
                "To Order", "First Shipment", "Second Shipment", "Third Shipment"
            ]
            tbl = overview_df.loc[overview_df[overview_col] == supplier, desired_cols].copy()
            tbl.replace([np.inf, -np.inf], np.nan, inplace=True)
            tbl.fillna("", inplace=True)

            # headers
            for c, col_name in enumerate(tbl.columns):
                fmt = header_fmt1 if c <= 7 else header_fmt2 if c <= 13 else header_fmt3
                ws.write(start_row, c, col_name, fmt)

            # data rows
            for r, row in enumerate(tbl.itertuples(index=False), start=start_row+1):
                awd, dni = row[4], row[5]
                for c, val in enumerate(row):
                    if c == 0:
                        style = depletion_zero_fmt if awd==0 else (
                            days_hi_fmt if dni>30 else
                            oos_good_fmt if dni>16 else
                            neutral_fmt if dni>10 else
                            oos_bad_fmt
                        )
                        ws.write(r, c, val, style)
                    elif c == 1:
                        pid = int(val) if isinstance(val, (int, float, np.integer)) else val
                        url = f"https://sbsabs.encompass8.com/Home?DashboardID=100018&ProductID={pid}"
                        ws.write_formula(r, c, f'=HYPERLINK("{url}",{pid})', hyperlink_fmt)
                    elif c == next_delivery_idx:
                        try:
                            dt = pd.to_datetime(val)
                            ws.write_datetime(r, c, dt, date_fmt)
                        except:
                            ws.write(r, c, val)
                    elif c == to_order_idx:
                        num = float(val) if val not in (None,"",np.nan) else 0
                        ws.write_number(r, c, num, to_order_num_fmt)
                    else:
                        cell = "" if isinstance(val,(int,float)) and not math.isfinite(val) else val
                        if c in two_dec_cols:
                            ws.write(r, c, cell, two_dec_fmt)
                        elif c in int_cols:
                            ws.write(r, c, cell, int_fmt)
                        elif c in {14,15,16} and cell:
                            try:
                                ship_dt = pd.to_datetime(cell)
                                ws.write_datetime(r, c, ship_dt, date_fmt)
                            except:
                                ws.write(r, c, cell)
                        else:
                            ws.write(r, c, cell)

            # ─── 6) Conditional formatting & borders ───────────────────────
            n = tbl.shape[0]
            row0, rowN = start_row, start_row + n
            last_col = min(len(tbl.columns) - 1, 16)

            # a) OOS Risk
            ws.conditional_format(
                row0 + 1, projected_oos_idx, rowN, projected_oos_idx,
                {"type": "cell", "criteria": ">", "value": 0, "format": oos_bad_fmt}
            )
            ws.conditional_format(
                row0 + 1, projected_oos_idx, rowN, projected_oos_idx,
                {"type": "cell", "criteria": "==", "value": 0, "format": oos_good_fmt}
            )

            # b) Thin grid inside table
            ws.conditional_format(
                row0, 0, rowN, last_col,
                {"type": "no_errors", "format": thin_border_fmt}
            )
            # c) Thick top border
            ws.conditional_format(
                row0, 0, row0, last_col,
                {"type": "no_errors", "format": thick_top_fmt}
            )
            # d) Thick bottom border
            ws.conditional_format(
                rowN, 0, rowN, last_col,
                {"type": "no_errors", "format": thick_bottom_fmt}
            )
            # e) Thick left border
            ws.conditional_format(
                row0, 0, rowN, 0,
                {"type": "no_errors", "format": thick_left_fmt}
            )
            # f) Thick right border at Q
            ws.conditional_format(
                row0, last_col, rowN, last_col,
                {"type": "no_errors", "format": thick_right_fmt}
            )

            # 7) Layout tweaks
            ws.set_row(start_row, 68 * 0.75, centered_wrap_fmt)
            for c, px in enumerate(pixel_widths):
                ws.set_column(c, c, px / 7.0)

            # 8) Insert POs table two rows below overview
            overview_end = start_row + n
            po_start = overview_end + 2

            # Section title
            ws.write(po_start, 0, f"POs for {supplier}", sec_hdr_fmt)

            # Reordered headers
            po_cols = ["Product", "Purchase ID", "Receive Date", "Ordered", "PO Num"]
            for c, col_name in enumerate(po_cols):
                ws.write(po_start+1, c, col_name, header_fmt2)

            # Data rows
            mask = (
                (po_df[po_col] == supplier)
                & po_df["PO Num"].notna()
                & po_df["PO Num"].astype(str).str.strip().astype(bool)
            )
            supplier_pos = po_df.loc[mask, po_cols]
            for r_idx, row in enumerate(supplier_pos.itertuples(index=False), start=po_start+2):
                product, purchase_id, recv_date, ordered, po_num = row
                ws.write(r_idx, 0, product, thin_border_fmt)

                # hyperlink for Purchase ID
                pid = int(purchase_id) if isinstance(purchase_id, (int, float)) else str(purchase_id)
                url = f"https://sbsabs.encompass8.com/Home?DashboardID=168160&KeyValue={pid}"
                ws.write_formula(
                    r_idx, 1,
                    f'=HYPERLINK("{url}", "{pid}")',
                    hyperlink_fmt
                )

                # Receive Date
                try:
                    dt = pd.to_datetime(recv_date)
                    ws.write_datetime(r_idx, 2, dt, date_fmt)
                except:
                    ws.write(r_idx, 2, recv_date or "", thin_border_fmt)

                ws.write_number(r_idx, 3, float(ordered), two_dec_fmt)
                ws.write(r_idx, 4, po_num, int_fmt)

        output.seek(0)
        return output

def generate_chatgpt_summary(supplier_data):
    """Generate a procurement summary with bolded section headers for readability,
       using the user-entered Notes & Next Steps if available,
       with first-person PO language including current open PO count."""
    import random
    import streamlit as st

    # ─── Phrase Pools ─────────────────────────────────────────────────────────
    inventory_good_phrases = [
        "stock levels are strong across the board",
        "we’ve got plenty of coverage on all SKUs",
        "inventory health is spot-on with no weak points",
        "supply levels look solid and well balanced",
        "Stock levels are healthy and stable across the portfolio."
        "We’re in a strong position from a supply standpoint."
        "Inventory coverage is solid across all products within the portfolio.",
        "We have ample inventory to support current and projected demand.",
        "Inventory availability is excellent with no current areas of concern.",
        "We're well positioned to meet any upcoming demand spikes.",
        "All SKUs are in good shape with comfortable/solid buffer levels.",
        "Our inventory profile is well balanced with minimal excess or shortage risk."

    ]
    inventory_mid_phrases = [
        "Inventory position remains strong overall, with a few isolated weak spots.",
        "Stock levels are healthy, though a couple of items are trending light on inventory.",
        "We’re in a good place overall, with a few SKUs needing extra attention.",
        "Overall coverage is solid, but with a few soft spots that need attention.",
        "The inventory for the portfolio looks stable, with minor exceptions.",
        "A few products are flirting with low coverage, but nothing critical yet.",
        "No major concerns, but a handful of items are running a little lean.",
        "Inventory is in good shape, though we’ve flagged some skus in the portfolio that need monitoring.",
        "General health is good, with a couple of areas that could benefit from an order or pickup request from another wholesaler.",
        "Supply position is strong overall, but we plan on keeping an eye on some outliers.",
        "Coverage remains sufficient, but we’re tracking a few early-warning SKUs.",
        "Most categories are covered, but there’s room for tightening in select items for this supplier.",
        "No immediate risks, but a few segments are edging toward caution zones.",
        "We're largely well-positioned, with only a couple of spots approaching the low mark.",

    ]
    inventory_risk_phrases = [
        "inventory levels are dipping in some areas",
        "we’re seeing tighter coverage on a few SKUs",
        "supply could tighten soon if not supplemented",
        "certain items are running lean and need watching",
        "we’re showing early signs of strain in some lines",
        "Inventory levels are trending downward in key areas.",
        "We’re starting to see low coverage emerge across a few items within the portfolio.",
        "Stock is thinning on select SKUs and may require further attention.",
        "Some items are approaching threshold levels and need close monitoring.",
        "Early indicators suggest potential shortages if current demand continues.",
        "Certain SKUs are at risk of stockouts if replenishment doesn’t land soon.",
        "Supply constraints are beginning to put pressure on inventory levels.",
        "Lean inventory is limiting our ability to respond to fluctuations in demand.",
        "Some SKUs might need expedited replenishment to avoid disruption.",
        "We’re flagging a few areas for urgent restock consideration. Short-term risk exists if supply timelines shift or delay."
       
    ]

    po_low_phrases = [
        "I will hold off on placing a PO since the MOQ hasn’t been reached.",
        "I’ll wait to place any PO until we hit the minimum order threshold.",
        "I'm not placing a PO right now because MOQ isn’t met with current ROS/Demand.",
        "I’ll defer the purchase order until our quantities reach the required MOQ. No PO needed this week",
        "At the moment our volume is below the supplier’s MOQ, so I’ll pause on releasing an order and re-evaluate next week.",
        "I’ll revisit this once our accumulated demand satisfies the minimum‑order requirement. No PO needed at this time",
        "Since current needs fall short of the MOQ, I won’t generate a PO just yet. We will take a look again next week",
        "We don’t hit the supplier’s minimum with today’s quantities, so I’ll wait on placing this order.",
        "I’m holding the order until we can consolidate more units and clear the MOQ. We'll take a look at it again this week.",
        "Let’s postpone issuing a PO until our forecasted volume meets the minimum threshold.",
        "Because we’re under the minimum lot size, I’ll push this order to the next cycle. No PO needed this week.",

    ]
    po_mid_phrases = [
        "I’ll bump up the order slightly so we hit the MOQ.",
        "I can increase the order size to satisfy the MOQ & get a PO placed.",
        "I’m adjusting the order to just clear the MOQ and we can place the PO's.",
        "I’ll increase the line‑item quantity so we comfortably reach the MOQ. We can place a PO for the below",
        "Let me round the order up to the minimum threshold to keep it moving so we can get a PO placed.",
        "I can pad the order a bit to satisfy the supplier’s MOQ requirement. We should be able to stretch a PO",
        "We’re just shy of the minimum Ill work on adding a few cases will get us there.",
        "I’m boosting the order to align with the supplier’s MOQ and avoid delays.",
        "An adjustment to the DOH Target will allow this purchase to meet the MOQ.",
        "I’ll expand the DOH Order Target so we achieve the required MOQ.",
        
    ]
    po_high_phrases = [
        "I will go ahead and place the recommended PO now that MOQ is met.",
        "I’m ready to submit the PO as planned.",
        "Quantities look solid—I'll release the PO today.",
        "Since we’ve cleared the MOQ, I’ll process this weeks order right away.",
        "All order thresholds met; I’m issuing a PO for the below this afternoon.",
        "The minimum order qty is satisfied, so I’ll finalize and submit the purchase order today.",
        "MOQ achieved; I’ll move forward with the order & place the PO into the system now.",
        "We’re at the required quantity, and I’ll dispatch the PO accordingly this afternoon for the below.",
        "I’ll go ahead and greenlight the purchase order for the qty below.",    
    ]

    blocks = []
    for supplier, data in supplier_data.items():
        min_order_pct, items_under_10, oos_risks, po_info = data
        po_count, po_numbers = po_info or (0, [])

        # Pull in user notes (or empty if none)
        notes = st.session_state.get(f"{supplier}_notes", "").strip()

        # ─── Compute total items and ratio ───────────────────────────────────
        overview_df = st.session_state.get(f"{supplier}_overview", pd.DataFrame())
        total_items = overview_df.shape[0]
        ratio = (items_under_10 / total_items) if total_items > 0 else 0

        # ─── Inventory health phrasing with mid-case ────────────────────────
        if items_under_10 == 0:
            health_phrase = random.choice(inventory_good_phrases)
        elif ratio < 0.25:
            health_phrase = random.choice(inventory_mid_phrases)
        else:
            health_phrase = random.choice(inventory_risk_phrases)

        # ─── Determine PO Recommendation ─────────────────────────────────────
        if min_order_pct < 0.75:
            po_choice = random.choice(po_low_phrases)
        elif min_order_pct < 1.0:
            po_choice = random.choice(po_mid_phrases)
        else:
            po_choice = random.choice(po_high_phrases)

        # ─── Build the text block ─────────────────────────────────────────────
        lines = [
            f"**{supplier}:**",
            "",
            "**Inventory Summary:**",
            f"Overall, {health_phrase}, with {items_under_10} item(s) under 10 days on hand (DOH).",
            "",
            "**Notes & Next Steps:**"
        ]
        if notes:
            lines.append(notes)
        lines.append("")

        # PO Recommendation with count + first-person phrasing
        po_header = f"**PO Recommendation:**"
        po_line = (
            f"I currently have {po_count} open PO{'s' if po_count != 1 else ''}. {po_choice}"
        )
        lines += [po_header, po_line, ""]

        # OOS Risk bullets
        lines.append("**OOS Risk:**")
        if oos_risks:
            if len(oos_risks) > 6:
                lines.append("• Multiple items with OOS Risk; see highlighted below.")
            else:
                for sku, name, risk in oos_risks:
                    lines.append(f"• {sku} {name} – {int(risk)} days")
        else:
            lines.append("• NO OOS RISK CURRENTLY")

        lines += [
            "",
            "**Order Builder Table:**",
            # left blank for user editing
        ]

        blocks.append("\n".join(lines))

    return "\n\n".join(blocks)

def display_export_section():
    """Display the export controls."""
    st.divider()
    if "export_data" not in st.session_state:
        return
    if st.button("Export to PO CSV", key="export_btn"):
        export_df = build_export_df(st.session_state["export_data"])
        if export_df.empty:
            st.warning("Nothing to export - check required fields.")
        else:
            csv = export_df.to_csv(index=False)
            st.download_button(
                label="Download CSV",
                data=csv,
                file_name=f"PO_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )

TITLE_LOGO_URL = "https://media.glassdoor.com/sqll/6024123/silver-eagle-beverages-squarelogo-1646829838016.png"

def main():
    # ─── Top-of-page header ──────────────────────────────────────
    header_html = f"""
    <div style="display: flex; align-items: center; margin-bottom: 16px;">
      <img
        src="{TITLE_LOGO_URL}"
        alt="Silver Eagle Beverages Logo"
        style="width:64px; height:auto; margin-right:12px;"
      />
      <h1 style="margin:0; font-size:2rem;">SEB Supplier Overview</h1>
    </div>
    """
    st.markdown(header_html, unsafe_allow_html=True)

    # ─── Choose data source ────────────────────────────────────────
    use_remote = st.checkbox("Load Master Incoming Report from GitHub", value=True)
    if use_remote:
        st.info("Fetching 'Master Incoming Report NEW.xlsm' from GitHub…")
        file_stream = fetch_remote_file(GITHUB_RAW_URL)
    else:
        file_stream = st.file_uploader(
            "Upload 'Master Incoming Report.xlsm'",
            type=["xlsm", "xlsx"]
        )
        if not file_stream:
            return st.info("⬆️ Upload the Excel file to begin.")

    # ─── Load sheets ────────────────────────────────────────────────
    supplier_df, po_df, overview_df = load_data(file_stream)
    shortcode_df = load_shortcode_data(file_stream)

    # find supplier column in main sheets
    supplier_col = find_supplier_col(supplier_df)
    po_col       = find_supplier_col(po_df)
    overview_col = find_supplier_col(overview_df) if not overview_df.empty else None

    if not supplier_col or not po_col:
        return st.error("❌ Could not locate required 'Supplier' columns.")
    if "Order Day" not in supplier_df.columns:
        return st.error("❌ 'Order Day' column not found in Supplier Info.")

    # ─── Filters ────────────────────────────────────────────────────
    supplier_df["Order Day"] = (
        supplier_df["Order Day"]
        .astype(str)
        .str.strip()
        .str.capitalize()
    )
    col1, col2 = st.columns(2)
    with col1:
        selected_day = st.selectbox(
            "Select Order Day",
            ["Any Day","Monday","Tuesday","Wednesday","Thursday","Friday"],
            index=0
        )
    with col2:
        managers = ["All"] + sorted(supplier_df["Brand Manager"].dropna().unique())
        selected_manager = st.selectbox("Filter by Brand Manager", managers, index=0)

    # apply filters
    filtered = supplier_df[supplier_df[supplier_col] != "Anheuser Busch"]
    if selected_day != "Any Day":
        filtered = filtered[filtered["Order Day"] == selected_day]
    if selected_manager != "All":
        filtered = filtered[filtered["Brand Manager"] == selected_manager]
    if filtered.empty:
        return st.warning("No suppliers match current filters.")

    # build order-day map
    supplier_order_day_map = dict(
        zip(
            supplier_df[supplier_col],
            supplier_df["Order Day"]
        )
    )

    st.subheader("Suppliers")
    if "report_data" not in st.session_state:
        st.session_state.report_data = {}

    # display each supplier
    for supplier in sorted(filtered[supplier_col].unique()):
        result = display_supplier(
            supplier,
            supplier_df,
            po_df,
            overview_df,
            supplier_col,
            po_col,
            overview_col,
            shortcode_df
        )
        if st.session_state.get("selected_suppliers", {}).get(supplier):
            # store for exports and summaries
            st.session_state.report_data[supplier] = result

    # ─── Export controls ───────────────────────────────────────────
    display_export_section()

    # ─── Procurement Assistant ─────────────────────────────────────
    if st.button("🙋🏻 Procurement Assistant", use_container_width=True):
        if not st.session_state.report_data:
            st.warning("Please select at least one supplier to generate a summary.")
        else:
            summary = generate_chatgpt_summary(st.session_state.report_data)
            with st.expander("View ChatGPT Summary", expanded=True):
                st.markdown(summary)
            st.download_button(
                label="🙋🏻 Procurement Assistant",
                data=summary,
                file_name=f"Procurement_Summary_{datetime.now():%Y%m%d}.txt",
                mime="text/plain"
            )

    # ─── DSR Excel Export ──────────────────────────────────────────
    if st.button("💽 Export DSR to Excel", use_container_width=True):
        if not st.session_state.report_data:
            st.warning("Please select at least one supplier to export.")
        else:
            supplier_logo_urls = dict(
                zip(
                    supplier_df[supplier_col],
                    supplier_df.get("Logos", pd.Series()).astype(str).fillna("").tolist()
                )
            )
            supplier_manager_map = dict(
                zip(
                    supplier_df[supplier_col],
                    supplier_df["Brand Manager"]
                )
            )
            excel_bytes = _export_report_to_excel_bytes(
                st.session_state.report_data,
                overview_df,
                overview_col,
                supplier_logo_urls,
                supplier_manager_map,
                supplier_order_day_map,
                po_df,
                po_col
            )
            st.download_button(
                label="💽 Download DSR Excel Report",
                data=excel_bytes,
                file_name=f"Procurement_Report_{datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
