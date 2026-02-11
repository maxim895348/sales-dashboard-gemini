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

# --- HELPER: ROBUST HEADER FINDER ---
def find_header_idx(df_preview, keywords):
    """Scans first rows to find the header index based on keywords."""
    for idx, row in df_preview.iterrows():
        row_str = row.astype(str).str.lower().tolist()
        # Check if ALL keywords are present in the row (fuzzy match)
        if all(any(k in x for x in row_str) for k in keywords):
            return idx
    return None

# --- CACHED DATA LOADER ---
@st.cache_data
def load_data(uploaded_file):
    """
    Master ETL function. Returns:
    1. df_sop_merged: Main S&OP data merged with Order totals.
    2. df_orders_detail: Raw detail data for Inventory analysis.
    3. df_inventory: Aggregated stock data from NL/EE.
    """
    logs = []
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # ==========================================
        # 1. LOAD S&OP (MASTER DATA)
        # ==========================================
        sop_sheet = next((s for s in sheet_names if "S&OP" in s or "Meeting" in s), None)
        if not sop_sheet:
            return None, None, None, ["Critical: S&OP sheet not found."]

        # Smart Header: Look for 'Status' and 'Country'
        df_prev_sop = pd.read_excel(xls, sheet_name=sop_sheet, header=None, nrows=15)
        sop_header = find_header_idx(df_prev_sop, ['status', 'country'])
        if sop_header is None: sop_header = 1 # Fallback
        
        df_sop = pd.read_excel(xls, sheet_name=sop_sheet, header=sop_header)
        df_sop.columns = df_sop.columns.astype(str).str.strip()
        
        # Identify Order ID
        sop_id_col = next((c for c in df_sop.columns if 'Proforma' in c), df_sop.columns[0])
        df_sop.rename(columns={sop_id_col: 'Order_ID'}, inplace=True)

        # Cleanup Rows
        if 'Order_ID' in df_sop.columns:
            df_sop = df_sop[~df_sop['Order_ID'].astype(str).str.lower().isin(['input', 'formula', 'nan', 'order number'])]
            df_sop = df_sop.dropna(subset=['Order_ID'])
            df_sop['Order_ID'] = df_sop['Order_ID'].astype(str).str.replace(r'\.0$', '', regex=True)

        # Dates
        date_cols = [col for col in df_sop.columns if 'date' in str(col).lower()]
        for col in date_cols:
            df_sop[col] = pd.to_datetime(df_sop[col], errors='coerce')

        # Status & Pallets Normalization
        if 'Status' not in df_sop.columns: df_sop['Status'] = 'UNKNOWN'
        df_sop['Status'] = df_sop['Status'].fillna('Unknown').astype(str).str.upper().str.strip()
        
        pallet_col = next((c for c in df_sop.columns if 'pallet' in str(c).lower()), None)
        df_sop['Pallets'] = pd.to_numeric(df_sop[pallet_col], errors='coerce').fillna(0) if pallet_col else 0

        # ==========================================
        # 2. LOAD ORDERS (DETAILS)
        # ==========================================
        orders_sheet = next((s for s in sheet_names if s.strip() == "Orders"), None)
        df_orders_detail = pd.DataFrame()
        
        if orders_sheet:
            # Smart Header: Look for 'Order' and 'Number'
            df_prev_ord = pd.read_excel(xls, sheet_name=orders_sheet, header=None, nrows=30)
            ord_header = find_header_idx(df_prev_ord, ['order', 'number'])
            if ord_header is None: ord_header = 20 # Fallback
            
            df_orders_detail = pd.read_excel(xls, sheet_name=orders_sheet, header=ord_header)
            df_orders_detail.columns = df_orders_detail.columns.astype(str).str.strip()
            
            # Identify Order ID
            ord_id_col = next((c for c in df_orders_detail.columns if 'order' in str(c).lower() and 'number' in str(c).lower()), df_orders_detail.columns[0])
            df_orders_detail.rename(columns={ord_id_col: 'Order_ID'}, inplace=True)
            df_orders_detail['Order_ID'] = df_orders_detail['Order_ID'].astype(str).str.replace(r'\.0$', '', regex=True)
            
            # Identify Product Code (Crucial for Inventory)
            # Look for 'Article', 'Product', 'Code' - avoid descriptions
            prod_col = next((c for c in df_orders_detail.columns if any(x in str(c).lower() for x in ['finished product', 'article no', 'sku'])), None)
            # Fallback: look for generic 'Product' but try to avoid descriptions
            if not prod_col:
                 prod_col = next((c for c in df_orders_detail.columns if 'product' in str(c).lower() and 'description' not in str(c).lower()), None)
            
            if prod_col:
                df_orders_detail.rename(columns={prod_col: 'Product_Code'}, inplace=True)
            else:
                df_orders_detail['Product_Code'] = 'Unknown_Product'

            # Identify Quantity
            qty_col = next((c for c in df_orders_detail.columns if 'quantity' in str(c).lower()), None)
            if qty_col:
                df_orders_detail['Quantity'] = pd.to_numeric(df_orders_detail[qty_col], errors='coerce').fillna(0)
            else:
                df_orders_detail['Quantity'] = 0

            # --- MERGE QTY TO S&OP ---
            df_agg = df_orders_detail.groupby('Order_ID')['Quantity'].sum().reset_index()
            df_agg.rename(columns={'Quantity': 'Total_Qty'}, inplace=True)
            
            if 'Total_Qty' in df_sop.columns: df_sop.drop(columns=['Total_Qty'], inplace=True)
            df_sop = pd.merge(df_sop, df_agg, on='Order_ID', how='left')
            df_sop['Total_Qty'] = df_sop['Total_Qty'].fillna(0)

        else:
            df_sop['Total_Qty'] = 0

        # ==========================================
        # 3. LOAD INVENTORY (STOCKLISTS)
        # ==========================================
        df_stock_total = pd.DataFrame(columns=['Product_Code', 'Stock_Qty', 'Location'])
        
        # --- NL Stock ---
        sheet_nl = next((s for s in sheet_names if 'stocklist' in s.lower() and 'nl' in s.lower()), None)
        if sheet_nl:
            try:
                # Based on snippets, header likely row 1
                df_nl = pd.read_excel(xls, sheet_name=sheet_nl, header=1)
                df_nl.columns = df_nl.columns.astype(str).str.strip()
                # Find columns
                prod_nl = next((c for c in df_nl.columns if 'prod' in c.lower() and 'code' in c.lower()), None) # 'Prod.code'
                qty_nl = next((c for c in df_nl.columns if 'quantity' in c.lower()), None)
                
                if prod_nl and qty_nl:
                    temp_nl = df_nl[[prod_nl, qty_nl]].copy()
                    temp_nl.columns = ['Product_Code', 'Stock_Qty']
                    temp_nl['Location'] = 'NL'
                    df_stock_total = pd.concat([df_stock_total, temp_nl])
            except:
                logs.append("Warning: Failed to parse Stocklist NL")

        # --- EE Stock ---
        sheet_ee = next((s for s in sheet_names if 'stocklist' in s.lower() and 'ee' in s.lower()), None)
        if sheet_ee:
            try:
                # Based on snippets, header likely row 0
                df_ee = pd.read_excel(xls, sheet_name=sheet_ee, header=0)
                df_ee.columns = df_ee.columns.astype(str).str.strip()
                # Find columns
                prod_ee = next((c for c in df_ee.columns if 'article' in c.lower()), None) # 'Article No.'
                qty_ee = next((c for c in df_ee.columns if 'quantity' in c.lower()), None)
                
                if prod_ee and qty_ee:
                    temp_ee = df_ee[[prod_ee, qty_ee]].copy()
                    temp_ee.columns = ['Product_Code', 'Stock_Qty']
                    temp_ee['Location'] = 'EE'
                    df_stock_total = pd.concat([df_stock_total, temp_ee])
            except:
                logs.append("Warning: Failed to parse Stocklist EE")

        # Clean Stock Data
        if not df_stock_total.empty:
            df_stock_total['Stock_Qty'] = pd.to_numeric(df_stock_total['Stock_Qty'], errors='coerce').fillna(0)
            df_stock_total['Product_Code'] = df_stock_total['Product_Code'].astype(str).str.strip()
            # Aggregate by Product
            df_inventory = df_stock_total.groupby('Product_Code')['Stock_Qty'].sum().reset_index()
        else:
            df_inventory = pd.DataFrame(columns=['Product_Code', 'Stock_Qty'])

        return df_sop, df_orders_detail, df_inventory, logs
        
    except Exception as e:
        return None, None, None, [f"Fatal Error: {str(e)}", traceback.format_exc()]

# --- UI LAYOUT ---
def main():
    st.sidebar.title("🎛️ Controls")
    
    uploaded_file = st.sidebar.file_uploader("Upload Master Excel", type=['xlsx', 'xls'])
    
    if not uploaded_file:
        st.title("🚀 Board Room S&OP Dashboard")
        st.info("👆 Please upload the `ORDERS.xlsx` file.")
        return

    # Load Data
    with st.spinner('Processing S&OP + Inventory Layers...'):
        df_sop, df_details, df_inv, logs = load_data(uploaded_file)
    
    # Error Handling
    if df_sop is None:
        st.error("❌ Failed to load data.")
        with st.expander("Error Logs"):
            for log in logs: st.code(log)
        return
    elif logs:
        with st.expander("⚠️ Loading Warnings (Non-Critical)"):
            for log in logs: st.write(log)

    # --- SIDEBAR FILTERS ---
    st.sidebar.divider()
    
    # Date Filter
    date_col = next((c for c in df_sop.columns if 'entry date' in str(c).lower()), None)
    if date_col:
        min_date = df_sop[date_col].min()
        max_date = df_sop[date_col].max()
        if pd.notnull(min_date) and pd.notnull(max_date):
            start_date, end_date = st.sidebar.date_input("Date Range", [min_date, max_date])
            df_sop = df_sop[(df_sop[date_col] >= pd.to_datetime(start_date)) & (df_sop[date_col] <= pd.to_datetime(end_date))]

    # Status Filter
    statuses = ['All'] + sorted(df_sop['Status'].unique().tolist())
    selected_status = st.sidebar.selectbox("Order Status", statuses)
    if selected_status != 'All':
        df_sop = df_sop[df_sop['Status'] == selected_status]

    # Country Filter
    if 'Country' in df_sop.columns:
        countries = ['All'] + sorted(df_sop['Country'].astype(str).unique().tolist())
        selected_country = st.sidebar.selectbox("Market / Country", countries)
        if selected_country != 'All':
            df_sop = df_sop[df_sop['Country'] == selected_country]

    # --- CALCULATIONS ---
    total_orders = len(df_sop)
    total_pallets = df_sop['Pallets'].sum()
    
    # Status Buckets
    hold_orders = df_sop[df_sop['Status'].str.contains('HOLD', na=False)]
    
    # Payment Block
    payment_col = next((c for c in df_sop.columns if 'payment' in str(c).lower() and 'status' in str(c).lower()), None)
    blocked_payment = pd.DataFrame()
    if payment_col:
         blocked_payment = df_sop[df_sop[payment_col].astype(str).str.contains('PAYMENT', case=False, na=False)]

    # --- MAIN DASHBOARD ---
    st.title(f"📊 S&OP Control Tower")
    st.markdown("Executive overview of Supply Chain, Finance & Inventory.")

    # Top Level KPI
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Pipeline Volume", f"{total_pallets:,.0f} plts", f"{total_orders} orders")
    kpi2.metric("Orders on HOLD", f"{len(hold_orders)}", f"{hold_orders['Pallets'].sum():,.0f} plts", delta_color="inverse")
    kpi3.metric("Finance Blocks", f"{len(blocked_payment)}", "Action Required", delta_color="inverse")
    
    # Inventory KPI (if available)
    if not df_inv.empty and not df_details.empty:
        total_stock = df_inv['Stock_Qty'].sum()
        kpi4.metric("Global Stock Level", f"{total_stock:,.0f}", "Units")
    else:
        kpi4.metric("Inventory Data", "N/A", "Check Sheets")

    # TABS
    tab1, tab2, tab3, tab4 = st.tabs(["📈 Board View", "💰 Financial", "🚚 Operations", "📦 Inventory & Risks"])

    # --- TAB 1: BOARD ---
    with tab1:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.subheader("Pipeline Flow (Status)")
            status_counts = df_sop['Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            fig_funnel = px.funnel(status_counts, x='Count', y='Status', title="Order Lifecycle")
            st.plotly_chart(fig_funnel, use_container_width=True)
        with c2:
            st.subheader("Market Volume")
            if 'Country' in df_sop.columns:
                country_counts = df_sop.groupby('Country')['Pallets'].sum().reset_index()
                fig_pie = px.pie(country_counts, values='Pallets', names='Country')
                st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader("⚠️ Long-Term Holds (>30 Days)")
        if date_col:
            now = pd.Timestamp.now()
            df_sop['Days_Open'] = (now - df_sop[date_col]).dt.days
            aging = df_sop[(df_sop['Status'].str.contains('HOLD')) & (df_sop['Days_Open'] > 30)].sort_values('Days_Open', ascending=False)
            if not aging.empty:
                cols = ['Order_ID', 'Country', 'Status', 'Days_Open']
                if payment_col: cols.append(payment_col)
                st.dataframe(aging[cols].head(10), use_container_width=True)
            else:
                st.success("No critical aging holds.")

    # --- TAB 2: FINANCE ---
    with tab2:
        st.subheader("Financial Bottlenecks")
        col_pay1, col_pay2 = st.columns(2)
        with col_pay1:
            if payment_col:
                pay_summary = df_sop[payment_col].fillna('Unknown').value_counts().reset_index()
                pay_summary.columns = ['Payment Status', 'Count']
                fig_pay = px.bar(pay_summary, x='Payment Status', y='Count', color='Payment Status')
                st.plotly_chart(fig_pay, use_container_width=True)
            else:
                st.info("Payment Status column missing.")
        with col_pay2:
            st.metric("Locked Value (Pre-Payment)", f"{blocked_payment['Pallets'].sum():,.0f} Pallets")
            st.dataframe(blocked_payment[['Order_ID', 'Country', payment_col] if payment_col else []], use_container_width=True)

    # --- TAB 3: OPS ---
    with tab3:
        st.subheader("Shipment Schedule")
        ready_col = next((c for c in df_sop.columns if 'ready' in str(c).lower() and 'date' in str(c).lower()), None)
        ship_col = next((c for c in df_sop.columns if 'shipment' in str(c).lower() and 'date' in str(c).lower()), None)
        
        if ready_col and ship_col:
            gantt_df = df_sop.dropna(subset=[ready_col, ship_col]).head(50)
            fig_gantt = px.timeline(gantt_df, x_start=ready_col, x_end=ship_col, y='Order_ID', color='Status')
            st.plotly_chart(fig_gantt, use_container_width=True)
        else:
            st.warning("Dates for Gantt not found.")
        st.dataframe(df_sop, use_container_width=True)

    # --- TAB 4: INVENTORY (NEW) ---
    with tab4:
        st.subheader("📦 Demand vs Supply Analysis")
        
        if df_details.empty or df_inv.empty:
            st.warning("Insufficient data for inventory analysis. Need 'Orders' details and 'Stocklist' sheets.")
        else:
            # 1. Calculate Demand per Product
            # Group details by Product Code
            if 'Product_Code' in df_details.columns:
                df_demand = df_details.groupby('Product_Code')['Quantity'].sum().reset_index()
                df_demand.rename(columns={'Quantity': 'Demand_Qty'}, inplace=True)
                
                # 2. Merge with Inventory
                df_risk = pd.merge(df_demand, df_inv, on='Product_Code', how='outer').fillna(0)
                
                # 3. Calculate Gap
                df_risk['Balance'] = df_risk['Stock_Qty'] - df_risk['Demand_Qty']
                df_risk['Status'] = np.where(df_risk['Balance'] >= 0, '✅ OK', '❌ Shortage')
                
                # Metrics
                shortage_items = df_risk[df_risk['Balance'] < 0]
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Unique Products Ordered", len(df_demand))
                c2.metric("Items in Shortage", len(shortage_items), delta=-len(shortage_items))
                c3.metric("Total Deficit (Units)", f"{abs(shortage_items['Balance'].sum()):,.0f}")
                
                st.divider()
                
                # Visualization of Shortages
                if not shortage_items.empty:
                    st.subheader("Critical Shortages (Stock < Demand)")
                    fig_risk = px.bar(
                        shortage_items.sort_values('Balance').head(15),
                        x='Balance', y='Product_Code',
                        color='Balance',
                        title="Top 15 Shortages",
                        orientation='h',
                        color_continuous_scale='RdYlGn'
                    )
                    st.plotly_chart(fig_risk, use_container_width=True)
                
                st.subheader("Detailed Inventory Status")
                st.dataframe(
                    df_risk.sort_values('Balance'),
                    column_config={
                        "Balance": st.column_config.ProgressColumn(
                            "Net Balance",
                            format="%d",
                            min_value=int(df_risk['Balance'].min()),
                            max_value=int(df_risk['Balance'].max()),
                        )
                    },
                    use_container_width=True
                )
            else:
                st.error("Could not identify 'Product Code' column in Orders sheet.")

if __name__ == "__main__":
    main()
