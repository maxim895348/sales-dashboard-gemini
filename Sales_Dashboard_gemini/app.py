import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import traceback

# --- CONFIGURATION ---
st.set_page_config(
    page_title="S&OP Control Tower",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CACHED DATA LOADER ---
@st.cache_data
def load_data(uploaded_file):
    """
    Ingests the complex Excel file, handles multi-sheet logic and cleaning.
    """
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # 1. Load S&OP (Master Data)
        sop_sheet = next((s for s in sheet_names if "S&OP" in s or "Meeting" in s), None)
        
        if not sop_sheet:
            st.error("❌ Critical: S&OP sheet not found in Excel.")
            return None, None

        # --- SMART HEADER DETECTION (Safe Mode) ---
        # Read first 15 rows to find where the actual header is
        df_preview = pd.read_excel(xls, sheet_name=sop_sheet, header=None, nrows=15)
        header_idx = 0
        
        for idx, row in df_preview.iterrows():
            row_str = row.astype(str).str.lower().tolist()
            # Look for row containing both 'status' and 'country'
            if any('status' in x for x in row_str) and any('country' in x for x in row_str):
                header_idx = idx
                break
        
        # Load with identified header
        df_sop = pd.read_excel(xls, sheet_name=sop_sheet, header=header_idx)
        
        # Force column names to string to avoid Int/Object errors
        df_sop.columns = df_sop.columns.astype(str).str.strip()
        
        # Rename Order ID column
        id_col = next((c for c in df_sop.columns if 'Proforma' in c), df_sop.columns[0])
        df_sop.rename(columns={id_col: 'Order_ID'}, inplace=True)

        # Filter out garbage rows
        if 'Order_ID' in df_sop.columns:
            df_sop = df_sop[~df_sop['Order_ID'].astype(str).str.lower().isin(['input', 'formula', 'nan', 'order number'])]
            df_sop = df_sop.dropna(subset=['Order_ID'])
        
        # Convert Dates
        date_cols = [col for col in df_sop.columns if 'date' in str(col).lower()]
        for col in date_cols:
            df_sop[col] = pd.to_datetime(df_sop[col], errors='coerce')

        # 2. Load Orders (Detail Data)
        orders_sheet = next((s for s in sheet_names if s.strip() == "Orders"), None)
        qty_merged = False # Flag to track if we successfully merged quantities
        
        if orders_sheet:
            # Smart Header for Orders sheet
            df_ord_preview = pd.read_excel(xls, sheet_name=orders_sheet, header=None, nrows=30)
            ord_header_idx = 20 # Fallback
            
            for idx, row in df_ord_preview.iterrows():
                row_str = row.astype(str).str.lower().tolist()
                # Look for 'order' AND 'number'
                if any('order' in x and 'number' in x for x in row_str):
                    ord_header_idx = idx
                    break
            
            df_orders = pd.read_excel(xls, sheet_name=orders_sheet, header=ord_header_idx)
            df_orders.columns = df_orders.columns.astype(str).str.strip()
            
            # Identify ID column in Orders
            oid_col = next((c for c in df_orders.columns if 'order' in str(c).lower() and 'number' in str(c).lower()), df_orders.columns[0])
            df_orders.rename(columns={oid_col: 'Order_ID'}, inplace=True)
            
            # Clean IDs for merging
            df_orders['Order_ID'] = df_orders['Order_ID'].astype(str).str.replace(r'\.0$', '', regex=True)
            
            # Aggregate Quantity
            qty_col = next((c for c in df_orders.columns if 'quantity' in str(c).lower()), None)
            
            if qty_col:
                # Ensure Quantity is numeric
                df_orders[qty_col] = pd.to_numeric(df_orders[qty_col], errors='coerce').fillna(0)
                df_agg = df_orders.groupby('Order_ID')[qty_col].sum().reset_index()
                df_agg.rename(columns={qty_col: 'Total_Qty'}, inplace=True)
                
                # CLEAN MERGE: Ensure Order_ID types match
                df_sop['Order_ID'] = df_sop['Order_ID'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_agg['Order_ID'] = df_agg['Order_ID'].astype(str)
                
                # Drop Total_Qty if it somehow already exists in SOP to avoid 'Total_Qty_x' suffix
                if 'Total_Qty' in df_sop.columns:
                    df_sop.drop(columns=['Total_Qty'], inplace=True)
                
                df_sop = pd.merge(df_sop, df_agg, on='Order_ID', how='left')
                qty_merged = True

        # Fallback if merge didn't happen
        if not qty_merged:
            df_sop['Total_Qty'] = 0

        # Fill NaNs resulting from merge
        df_sop['Total_Qty'] = df_sop['Total_Qty'].fillna(0)

        # 3. Final Polish
        # Normalize Status
        if 'Status' in df_sop.columns:
            df_sop['Status'] = df_sop['Status'].fillna('Unknown').astype(str).str.upper().str.strip()
        else:
            df_sop['Status'] = 'UNKNOWN'
        
        # Normalize Pallets
        pallet_col = next((c for c in df_sop.columns if 'pallet' in str(c).lower()), None)
        if pallet_col:
            df_sop['Pallets'] = pd.to_numeric(df_sop[pallet_col], errors='coerce').fillna(0)
        else:
            df_sop['Pallets'] = 0
            
        return df_sop, sheet_names
        
    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
        with st.expander("See technical details (Traceback)"):
            st.code(traceback.format_exc())
        return None, None

# --- UI LAYOUT ---
def main():
    st.sidebar.title("🎛️ Controls")
    
    uploaded_file = st.sidebar.file_uploader("Upload S&OP Excel", type=['xlsx', 'xls'])
    
    if not uploaded_file:
        st.title("🚀 Board Room S&OP Dashboard")
        st.info("👆 Please upload the master `ORDERS.xlsx` file to begin.")
        return

    # Load Data
    with st.spinner('Parsing ERP Data...'):
        df, sheets = load_data(uploaded_file)
    
    if df is None:
        return

    # --- SIDEBAR FILTERS ---
    st.sidebar.divider()
    
    # Year/Month Filter
    date_col = next((c for c in df.columns if 'entry date' in str(c).lower()), None)
    if date_col:
        min_date = df[date_col].min()
        max_date = df[date_col].max()
        if pd.notnull(min_date) and pd.notnull(max_date):
            start_date, end_date = st.sidebar.date_input(
                "Date Range",
                [min_date, max_date]
            )
            df = df[(df[date_col] >= pd.to_datetime(start_date)) & (df[date_col] <= pd.to_datetime(end_date))]

    # Status Filter
    statuses = ['All'] + sorted(df['Status'].unique().tolist())
    selected_status = st.sidebar.selectbox("Order Status", statuses)
    if selected_status != 'All':
        df = df[df['Status'] == selected_status]

    # Country Filter
    if 'Country' in df.columns:
        countries = ['All'] + sorted(df['Country'].astype(str).unique().tolist())
        selected_country = st.sidebar.selectbox("Market / Country", countries)
        if selected_country != 'All':
            df = df[df['Country'] == selected_country]

    # --- KPI CALCULATIONS ---
    total_orders = len(df)
    total_pallets = df['Pallets'].sum()
    total_qty = df['Total_Qty'].sum()
    
    # SAFE STATUS BUCKETS
    hold_orders = df[df['Status'].str.contains('HOLD', na=False)]
    
    # Payment Block
    payment_col = next((c for c in df.columns if 'payment' in str(c).lower() and 'status' in str(c).lower()), None)
    blocked_payment = pd.DataFrame()
    if payment_col:
         blocked_payment = df[df[payment_col].astype(str).str.contains('PAYMENT', case=False, na=False)]

    # --- MAIN DASHBOARD ---
    st.title(f"📊 S&OP Control Tower")
    st.markdown("Executive overview of Supply Chain & Financial Pipeline.")

    # Top Level KPI
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total Pipeline Volume", f"{total_pallets:,.0f} plts", f"{total_orders} orders")
    kpi2.metric("Orders on HOLD", f"{len(hold_orders)}", f"{hold_orders['Pallets'].sum():,.0f} plts", delta_color="inverse")
    kpi3.metric("Blocked by Payment", f"{len(blocked_payment)}", "Requires Finance Action", delta_color="inverse")
    kpi4.metric("Total Items (Qty)", f"{total_qty:,.0f}", "Units")

    # TABS
    tab1, tab2, tab3 = st.tabs(["📈 Board View", "💰 Financial Blocks", "🚚 Operations & Log"])

    with tab1:
        c1, c2 = st.columns([2, 1])
        
        with c1:
            st.subheader("Pipeline Flow (Status)")
            status_counts = df['Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            fig_funnel = px.funnel(status_counts, x='Count', y='Status', title="Order Lifecycle Funnel")
            st.plotly_chart(fig_funnel, use_container_width=True)
            
        with c2:
            st.subheader("Market Distribution")
            if 'Country' in df.columns:
                country_counts = df.groupby('Country')['Pallets'].sum().reset_index()
                fig_pie = px.pie(country_counts, values='Pallets', names='Country', title="Volume by Country")
                st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader("⚠️ Critical Attention Required (Hold > 30 Days)")
        if date_col:
            now = pd.Timestamp.now()
            df['Days_Open'] = (now - df[date_col]).dt.days
            aging_holds = df[(df['Status'].str.contains('HOLD')) & (df['Days_Open'] > 30)].sort_values('Days_Open', ascending=False)
            if not aging_holds.empty:
                cols_to_show = ['Order_ID', 'Country', 'Status', 'Days_Open']
                if payment_col: cols_to_show.append(payment_col)
                st.dataframe(aging_holds[cols_to_show].head(10), use_container_width=True)
            else:
                st.success("No long-term holds detected.")

    with tab2:
        st.subheader("Financial Bottlenecks")
        col_pay1, col_pay2 = st.columns(2)
        
        with col_pay1:
            if payment_col:
                pay_summary = df[payment_col].fillna('Unknown').value_counts().reset_index()
                pay_summary.columns = ['Payment Status', 'Count']
                fig_pay = px.bar(pay_summary, x='Payment Status', y='Count', color='Payment Status', title="Orders by Payment Status")
                st.plotly_chart(fig_pay, use_container_width=True)
            else:
                st.info("Payment Status column not found.")
        
        with col_pay2:
            st.info("💡 Insight: Orders marked 'PAYMENT' cannot ship until cleared.")
            st.metric("Value Locked in Pre-Payment", f"{blocked_payment['Pallets'].sum():,.0f} Pallets")

        st.subheader("Action List: Payment Pending")
        st.dataframe(blocked_payment, use_container_width=True)

    with tab3:
        st.subheader("Operations Schedule")
        ready_col = next((c for c in df.columns if 'ready' in str(c).lower() and 'date' in str(c).lower()), None)
        ship_col = next((c for c in df.columns if 'shipment' in str(c).lower() and 'date' in str(c).lower()), None)
        
        if ready_col and ship_col:
            gantt_df = df.dropna(subset=[ready_col, ship_col]).copy()
            gantt_df = gantt_df.head(50) 
            
            fig_gantt = px.timeline(
                gantt_df, 
                x_start=ready_col, 
                x_end=ship_col, 
                y='Order_ID', 
                color='Status',
                title="Production to Shipment Timeline (Top 50)"
            )
            st.plotly_chart(fig_gantt, use_container_width=True)
        else:
            st.warning("Date columns for Gantt chart not auto-detected.")

        st.subheader("Full Data Drill-down")
        st.dataframe(df, use_container_width=True)

if __name__ == "__main__":
    main()
