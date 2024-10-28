import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import base64
import time

# Schneider Electric Brand Colors
SCHNEIDER_COLORS = {
    'primary_green': '#3DCD58',
    'dark_green': '#004F3B',
    'light_gray': '#F5F5F5',
    'dark_gray': '#333333',
    'white': '#FFFFFF',
    'accent_blue': '#007ACC',
    'chart_colors': ['#3DCD58', '#004F3B', '#007ACC', '#00A6A0', '#676767', '#CCCCCC']
}

def load_css():
    """Load custom CSS styles with Schneider Electric branding"""
    st.markdown(f"""
        <style>
        .main {{
            background-color: #1E1E1E;
            color: {SCHNEIDER_COLORS['white']};
        }}
        .stApp {{
            background-color: #1E1E1E;
        }}
        # ... (rest of the CSS styles)
        </style>
    """, unsafe_allow_html=True)

def process_dataframe_wwp(df):
    try:
        # Rename columns
        column_mapping = {
            'Part Number (Standardized)': 'Part Number',
            'Supplier DUNS Elementary Code': 'DUNS Elementary Code',
            'Next 12m Projection Quantity (Normalized UoM)': '12m Projection Quantity',
            'Line Price (EUR/NUoM) (Includes SQL FX)': 'Unit Price in Euros',
            'CPR:Best Line Price (including Logistics Simulation Delta if any) (EUR/NUoM) (Global)': 'Best Price in Euros',
            'CPR:Quantity of Best Price Line (NUoM) (Global)': 'Best Price Quantity',
            'CPR:Site Name of Best Price Line (Global)': 'Best Price Site',
            'CPR:Site Region of Best Price Line (Global)': 'Best Price Region',
            'CPR:Supplier Name of Best Price Line (Global)': 'Best Price Supplier',
            'CPR:Total Opportunity (EUR), including Logistics Simulation (Global)': 'Total Opportunity'
        }
        df = df.rename(columns=column_mapping)

        # Convert numeric columns
        for col in df.select_dtypes(include=['object']).columns:
            try:
                df[col] = pd.to_numeric(df[col].str.replace(',', ''))
            except:
                pass

        # Apply filters
        india_sites = ['IN Bangalore ITB', 'IN Chennai', 'IN Hyderabad', 'IN Bangalore SEPFC']
        category_codes = ('A', 'B', 'C', 'D', 'H', 'K', 'G', 'E', 'P1', 'P2', 'M1', 'M2')
        
        df_filtered = df[
            (df['Site Name'].isin(india_sites)) & 
            (df['Category Code'].str.startswith(category_codes))
        ]

        # Apply spend and region filters
        df_filtered = df_filtered[
            (df_filtered['Spend (EUR)'] > 50000) & 
            (df_filtered['Best Price Region'] != 'India / MEA') & 
            (df_filtered['Total Opportunity'] <= -5000)
        ]

        # Calculate ratio
        df_filtered['Qty/projection'] = ((df_filtered['Best Price Quantity'] / df_filtered['12m Projection Quantity']) * 100)
        
        # Filter one-time buys
        df_filtered = df_filtered[df_filtered['Qty/projection'] > 5]

        # Add absolute opportunity column for visualization
        df_filtered['Absolute Opportunity'] = df_filtered['Total Opportunity'].abs()

        # Format float values
        for col in df_filtered.select_dtypes(include=['float64']).columns:
            df_filtered[col] = df_filtered[col].round(2)

        return df_filtered
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        return None

def generate_insights_wwp(df):
    """Generate key insights from the processed data"""
    total_opportunity = df['Total Opportunity'].sum()
    avg_qty_projection = df['Qty/projection'].mean()
    
    # Use absolute values for top suppliers
    top_suppliers = df.groupby('Supplier Name')['Absolute Opportunity'].sum().sort_values(ascending=False).head(5)
    top_categories = df.groupby('Category Code')['Absolute Opportunity'].sum().sort_values(ascending=False).head(5)
    
    return {
        'total_opportunity': total_opportunity,
        'avg_qty_projection': avg_qty_projection,
        'top_suppliers': top_suppliers,
        'top_categories': top_categories
    }

def create_visualizations_wwp(df):
    template = {
        'layout': {
            'plot_bgcolor': '#1E1E1E',
            'paper_bgcolor': '#1E1E1E',
            'font': {'color': SCHNEIDER_COLORS['white']},
            'title': {'font': {'color': SCHNEIDER_COLORS['white']}},
            'xaxis': {'gridcolor': '#333333', 'linecolor': '#333333'},
            'yaxis': {'gridcolor': '#333333', 'linecolor': '#333333'}
        }
    }

    # Opportunity by Category
    category_data = df.groupby('Category Code')['Absolute Opportunity'].sum().reset_index()
    category_data = category_data.sort_values('Absolute Opportunity', ascending=True)
    
    fig1 = px.bar(
        category_data,
        x='Absolute Opportunity',
        y='Category Code',
        title='Savings Opportunity by Category (EUR)',
        orientation='h',
        color_discrete_sequence=[SCHNEIDER_COLORS['primary_green']]
    )
    fig1.update_layout(
        template=template,
        yaxis_title="Category Code",
        xaxis_title="Savings Opportunity (EUR)",
        height=500,
        showlegend=False,
        hovermode='closest',
        hoverlabel=dict(bgcolor=SCHNEIDER_COLORS['dark_green'])
    )

    # Opportunity by Supplier (Top 10)
    supplier_data = df.groupby('Supplier Name')['Absolute Opportunity'].sum().sort_values(ascending=False).head(10).reset_index()
    
    fig2 = px.pie(
        supplier_data,
        values='Absolute Opportunity',
        names='Supplier Name',
        title='Top 10 Suppliers by Savings Opportunity',
        color_discrete_sequence=SCHNEIDER_COLORS['chart_colors']
    )
    fig2.update_layout(
        template=template,
        hoverlabel=dict(bgcolor=SCHNEIDER_COLORS['dark_green'])
    )
    fig2.update_traces(textposition='inside', textinfo='percent+label')

    # Bar chart for top suppliers
    fig3 = px.bar(
        supplier_data,
        x='Supplier Name',
        y='Absolute Opportunity',
        title='Top 10 Suppliers by Savings Opportunity (EUR)',
        color_discrete_sequence=[SCHNEIDER_COLORS['primary_green']]
    )
    fig3.update_layout(
        template=template,
        xaxis_title="Supplier Name",
        yaxis_title="Savings Opportunity (EUR)",
        xaxis={'tickangle': 45},
        height=500,
        showlegend=False,
        hoverlabel=dict(bgcolor=SCHNEIDER_COLORS['dark_green'])
    )

    return [fig1, fig2, fig3]
def get_table_download_link_wwp(df):
    """Generate a styled download link for the processed data"""
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="processed_data.csv" class="download-link">📥 Download Processed Data</a>'
    return href

## Open PO Anaysis 
# Currency conversion rates (from OPO)
CONVERSION_RATES = {
    'USD': 0.93,
    'GBP': 1.2,
    'INR': 0.011,
    'JPY': 0.0061
}

def convert_to_euro(price, currency):
    """Converts a price to Euros based on the provided currency."""
    if currency in CONVERSION_RATES:
        return price * CONVERSION_RATES[currency]
    return price  # Return original price if currency not found instead of None

def process_data_opo(open_po_df, workbench_df):
    try:
        # Filter Open PO for LINE_TYPE = Inventory
        open_po_df = open_po_df[open_po_df['LINE_TYPE'] == 'Inventory']
        
        # Clean up column names before merge
        open_po_df.columns = open_po_df.columns.str.strip()
        workbench_df.columns = workbench_df.columns.str.strip()
        
        # Rename UNIT_PRICE columns before merge to avoid confusion
        open_po_df = open_po_df.rename(columns={'UNIT_PRICE': 'UNIT_PRICE_OPO'})
        workbench_df = workbench_df.rename(columns={'UNIT_PRICE': 'UNIT_PRICE_WB'})
        
        # Merge dataframes
        merged_df = pd.merge(
            workbench_df,
            open_po_df,
            left_on=['PART_NUMBER', 'VENDOR_NUM'],
            right_on=['ITEM', 'VENDOR_NUM'],
            how='inner'
        )
        
        # Drop redundant column
        merged_df = merged_df.drop('ITEM', axis=1)
        
        # Rename columns
        merged_df = merged_df.rename(columns={
            'DANDB': 'VENDOR_DUNS',
            'CURRENCY_CODE': 'CURRENCY_CODE_WB',
            'CURRNECY': 'CURRENCY_CODE_OPO'  # Fixed typo in CURRENCY
        })
        
        # Add IG/OG classification
        merged_df['IG/OG'] = merged_df['VENDOR_NAME'].apply(
            lambda x: 'IG' if 'SCHNEIDER' in str(x).upper() or 'WUXI' in str(x).upper() else 'OG'
        )
        
        # Add PO Year
        merged_df['PO Year'] = pd.to_datetime(merged_df['PO_SHIPMENT_CREATION_DATE']).dt.year
        
        # Convert prices to EUR
        merged_df['UNIT_PRICE_WB_EUR'] = merged_df.apply(
            lambda row: convert_to_euro(row['UNIT_PRICE_WB'], row['CURRENCY_CODE_WB']), axis=1
        )
        merged_df['UNIT_PRICE_OPO_EUR'] = merged_df.apply(
            lambda row: convert_to_euro(row['UNIT_PRICE_OPO'], row['CURRENCY_CODE_OPO']), axis=1
        )
        
        # Calculate metrics
        merged_df['Price_Delta'] = merged_df['UNIT_PRICE_OPO_EUR'] - merged_df['UNIT_PRICE_WB_EUR']
        merged_df['Impact in Euros'] = merged_df['Price_Delta'] * merged_df['QTY_ELIGIBLE_TO_SHIP']
        merged_df['Open PO Value'] = merged_df['QTY_ELIGIBLE_TO_SHIP'] * merged_df['UNIT_PRICE_OPO_EUR']
        
        # Sort by impact
        merged_df = merged_df.sort_values('Impact in Euros', ascending=False)
        
        return merged_df
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        return None

def generate_insights_opo(df):
    """Generate key insights from the processed data"""
    if df is None or df.empty:
        return None
        
    total_impact = df['Impact in Euros'].sum()
    total_po_value = df['Open PO Value'].sum()
    distinct_parts_count = df['PART_NUMBER'].nunique() 
    unique_vendors = df['VENDOR_NAME'].nunique()
    
    # Group by analyses
    impact_by_vendor = df.groupby('VENDOR_NAME')['Impact in Euros'].sum().sort_values(ascending=False).head(5)
    impact_by_category = df.groupby('STARS Category Code')['Impact in Euros'].sum().sort_values(ascending=False).head(5)
    
    return {
        'total_impact': total_impact,
        'total_po_value': total_po_value,
        'distinct_parts_count': distinct_parts_count,
        'unique_vendors': unique_vendors,
        'impact_by_vendor': impact_by_vendor,
        'impact_by_category': impact_by_category
    }

def create_visualizations_opo(df):
    """Create visualizations using Plotly"""
    if df is None or df.empty:
        return None
        
    # Impact by Category
    category_fig = px.bar(
        df.groupby('STARS Category Code')['Impact in Euros'].sum().sort_values(ascending=False).head(10).reset_index(),
        x='Impact in Euros',
        y='STARS Category Code',
        title='Price Impact by Category (EUR)',
        orientation='h'
    )
    category_fig.update_layout(height=500)
    
    # Impact by Vendor (Top 10)
    vendor_fig = px.pie(
        df.groupby('VENDOR_NAME')['Impact in Euros'].sum().sort_values(ascending=False).head(10).reset_index(),
        values='Impact in Euros',
        names='VENDOR_NAME',
        title='Top 10 Vendors by Price Impact'
    )
    
    # Impact by IG/OG
    ig_og_fig = px.bar(
        df.groupby('IG/OG')['Impact in Euros'].sum().reset_index(),
        x='IG/OG',
        y='Impact in Euros',
        title='Price Impact by IG/OG Classification',
        color='IG/OG'
    )
    
    # Timeline of PO Creation
    timeline_fig = px.line(
        df.groupby('PO_SHIPMENT_CREATION_DATE')['Impact in Euros'].sum().reset_index(),
        x='PO_SHIPMENT_CREATION_DATE',
        y='Impact in Euros',
        title='Price Impact Timeline'
    )
    
    return [category_fig, vendor_fig, ig_og_fig, timeline_fig]

def get_download_link_opo(df, filename="processed_data.csv"):
    """Generate a download link for the processed data"""
    if df is None or df.empty:
        return ""
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def main():
    st.set_page_config(
        page_title="Procurement Analysis",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    load_css()

    # Header with logo
    st.markdown("""
        <div class="header-container">
            <img src="https://www.se.com/ww/en/assets/wiztopic/615aeb0184d20b323d58575e/Schneider-Electric-logo-jpg-_original.jpg" 
                 style="width: 150px; margin-right: 20px;">
            <div>
                <h1 style="margin: 0;">Procurement Analysis</h1>
                <p style="color: #3DCD58; margin: 0;">India Effectiveness Team</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    analysis_type = st.sidebar.radio(
        "Select Analysis Type",
        ["Worldwide Price Analysis", "Open PO Analysis"]
    )
    if analysis_type == "Worldwide Price Analysis":
        st.markdown("<h2>Worldwide Price Analysis</h2>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Upload WWP Data", type=['xlsx', 'csv'])
        
        if uploaded_file is not None:
            try:
                with st.spinner('Processing your data...'):
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                    else:
                        df = pd.read_excel(uploaded_file)
                    time.sleep(0.5)  # Short delay for visual feedback
    
                st.success("✅ File uploaded and processed successfully!")
    
                # Process data
                df_processed = process_dataframe_wwp(df)
                
                if df_processed is not None and not df_processed.empty:
                    # Generate insights
                    insights = generate_insights_wwp(df_processed)
                    
                    # Display metrics with enhanced styling
                    st.markdown("<div class='metrics-container'>", unsafe_allow_html=True)
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric(
                            "💰 Total Savings Opportunity",
                            f"€{abs(insights['total_opportunity']):,.2f}"
                        )
                    with col2:
                        st.metric(
                            "📊 Avg Qty/Projection Ratio",
                            f"{insights['avg_qty_projection']:.2f}%"
                        )
                    with col3:
                        st.metric(
                            "🔢 Number of Parts",
                            f"{len(df_processed):,}"
                        )
                    with col4:
                        st.metric(
                            "🏢 Number of Suppliers",
                            f"{df_processed['Supplier Name'].nunique():,}"
                        )
                    st.markdown("</div>", unsafe_allow_html=True)
    
                    # Create tabs for different views
                    tab1, tab2, tab3 = st.tabs(["📈 Visualizations", "📋 Data Table", "🎯 Top Analysis"])
    
                    with tab1:
                        figures = create_visualizations_wwp(df_processed)
                        st.plotly_chart(figures[0], use_container_width=True)
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.plotly_chart(figures[1], use_container_width=True)
                        with col2:
                            st.plotly_chart(figures[2], use_container_width=True)
    
                    with tab2:
                        st.dataframe(df_processed, height=400)
                        st.markdown(get_table_download_link_wwp(df_processed), unsafe_allow_html=True)
    
                    with tab3:
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown(f"""
                            <div class='metric-card'>
                                <h3 style='color: {SCHNEIDER_COLORS["primary_green"]}'>
                                    🏆 Top Suppliers by Savings Opportunity
                                </h3>
                            """, unsafe_allow_html=True)
                            supplier_table = pd.DataFrame({
                                'Supplier': insights['top_suppliers'].index,
                                'Savings Opportunity (EUR)': insights['top_suppliers'].values.round(2)
                            })
                            st.table(supplier_table)
                            st.markdown("</div>", unsafe_allow_html=True)
                        
                        with col2:
                            st.markdown(f"""
                            <div class='metric-card'>
                                <h3 style='color: {SCHNEIDER_COLORS["primary_green"]}'>
                                    📊 Top Categories by Savings Opportunity
                                </h3>
                            """, unsafe_allow_html=True)
                            category_table = pd.DataFrame({
                                'Category': insights['top_categories'].index,
                                'Savings Opportunity (EUR)': insights['top_categories'].values.round(2)
                            })
                            st.table(category_table)
                            st.markdown("</div>", unsafe_allow_html=True)
    
                else:
                    st.warning("⚠️ No data matches the filtering criteria.")
    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.info("📝 Please ensure your file has the required columns and format.")
            
        else:
            st.markdown("<h2>Open PO Analysis</h2>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                open_po_file = st.file_uploader("Upload Open PO Report", type=['xlsx'])
            with col2:
                workbench_file = st.file_uploader("Upload Workbench Report", type=['xlsx'])
    
            if open_po_file is not None and workbench_file is not None:
                       try:
                            open_po_df = pd.read_excel(
                                open_po_file,
                                usecols=['     ORDER_TYPE', 'LINE_TYPE', 'ITEM', 'VENDOR_NUM', 'PO_NUM', 'RELEASE_NUM', 
                                        'LINE_NUM', 'SHIPMENT_NUM', 'AUTHORIZATION_STATUS', 'PO_SHIPMENT_CREATION_DATE',
                                        'QTY_ELIGIBLE_TO_SHIP', 'UNIT_PRICE', 'CURRNECY']
                            )
                            
                            workbench_df = pd.read_excel(
                                workbench_file,
                                usecols=['PART_NUMBER', 'DESCRIPTION', 'VENDOR_NUM', 'VENDOR_NAME', 'DANDB',
                                        'STARS Category Code', 'ASL_MPN', 'UNIT_PRICE', 'CURRENCY_CODE']
                            )
                
                            st.success("Files uploaded successfully!")
                
                            # Process data
                            processed_df = process_data_opo(open_po_df, workbench_df)
                            
                            if processed_df is not None and not processed_df.empty:
                                # Generate insights
                                insights = generate_insights_opo(processed_df)
                                
                                if insights:
                                    # Display metrics
                                    col1, col2, col3, col4 = st.columns(4)
                                    with col1:
                                        st.metric("Total Price Impact (EUR)", f"{insights['total_impact']:,.2f}")
                                    with col2:
                                        st.metric("Total Open PO Value (EUR)", f"{insights['total_po_value']:,.2f}")
                                    with col3:
                                        st.metric("Number of Parts", insights['distinct_parts_count'])
                                    with col4:
                                        st.metric("Number of Vendors", insights['unique_vendors'])
                
                                    # Create tabs
                                    tab1, tab2, tab3 = st.tabs(["Visualizations", "Data Table", "Top Impact Analysis"])
                
                                    with tab1:
                                        figures = create_visualizations_opo(processed_df)
                                        if figures:
                                            col1, col2 = st.columns(2)
                                            with col1:
                                                st.plotly_chart(figures[0], use_container_width=True)
                                                st.plotly_chart(figures[2], use_container_width=True)
                                            with col2:
                                                st.plotly_chart(figures[1], use_container_width=True)
                                                st.plotly_chart(figures[3], use_container_width=True)
                
                                    with tab2:
                                        st.dataframe(processed_df)
                                        st.markdown(get_download_link_opo(processed_df), unsafe_allow_html=True)
                
                                    with tab3:
                                        col1, col2 = st.columns(2)
                                        with col1:
                                            st.subheader("Top Vendors by Price Impact")
                                            st.table(pd.DataFrame({
                                                'Vendor': insights['impact_by_vendor'].index,
                                                'Impact (EUR)': insights['impact_by_vendor'].values.round(2)
                                            }))
                                        
                                        with col2:
                                            st.subheader("Top Categories by Price Impact")
                                            st.table(pd.DataFrame({
                                                'Category': insights['impact_by_category'].index,
                                                'Impact (EUR)': insights['impact_by_category'].values.round(2)
                                            }))
                
                            else:
                                st.warning("No data matches the analysis criteria.")
                
                       except Exception as e:
                           st.error(f"Error: {str(e)}")
                           st.info("Please ensure your files have the required columns and format.")
            
if __name__ == "__main__":
    main()
