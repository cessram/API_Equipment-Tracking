import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io

# Page Configuration
st.set_page_config(
    page_title="Equipment Tracker | Bird Construction",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS with Bird Construction Colors
st.markdown("""
    <style>
    /* Main color scheme: Blue, Green, White, Yellow, Orange */
    .stApp {
        background-color: #FFFFFF;
    }
    
    /* Headers */
    h1 {
        color: #1B4D89 !important;
        font-weight: 700;
    }
    
    h2, h3 {
        color: #2E7D32 !important;
    }
    
    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 28px;
        font-weight: 700;
    }
    
    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1B4D89 0%, #2E7D32 100%);
    }
    
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    /* Buttons */
    .stButton>button {
        background-color: #FF9800;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
        transition: all 0.3s;
    }
    
    .stButton>button:hover {
        background-color: #F57C00;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #F5F5F5;
        padding: 8px;
        border-radius: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: white;
        border-radius: 6px;
        color: #1B4D89;
        font-weight: 600;
        padding: 8px 16px;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #1B4D89 0%, #2E7D32 100%);
        color: white !important;
    }
    
    /* Cards */
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border-left: 4px solid #2E7D32;
        margin: 10px 0;
    }
    
    .alert-card {
        background: #FFF3E0;
        padding: 15px;
        border-radius: 8px;
        border-left: 4px solid #FF9800;
        margin: 10px 0;
    }
    
    /* Success message */
    .success-box {
        background: #E8F5E9;
        padding: 15px;
        border-radius: 8px;
        border-left: 4px solid #2E7D32;
        color: #1B5E20;
        font-weight: 600;
    }
    
    /* Download section */
    .download-section {
        background: linear-gradient(135deg, #FFEB3B 0%, #FF9800 100%);
        padding: 20px;
        border-radius: 12px;
        margin: 20px 0;
    }
    
    /* Status badges */
    .status-active {
        background-color: #2E7D32;
        color: white;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 12px;
    }
    
    .status-idle {
        background-color: #FF9800;
        color: white;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 12px;
    }
    
    .status-maintenance {
        background-color: #D32F2F;
        color: white;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 12px;
    }
    </style>
""", unsafe_allow_html=True)

# Load and process data
@st.cache_data
def load_data(file):
    """Load and process equipment data"""
    df = pd.read_excel(file) if file.name.endswith('.xlsx') else pd.read_csv(file)
    
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Convert date columns
    date_cols = ['Last Inspection Date', 'Next Inspection Due', 
                 'Mobilization Date', 'Planned Demob Date', 'Actual Demob Date']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Calculate days onsite
    if 'Mobilization Date' in df.columns:
        df['Days Onsite'] = (datetime.now() - df['Mobilization Date']).dt.days
        df['Days Onsite'] = df['Days Onsite'].clip(lower=0)
    
    # Calculate variance
    if 'Planned Duration (Days)' in df.columns and 'Days Onsite' in df.columns:
        df['Duration Variance'] = df['Days Onsite'] - df['Planned Duration (Days)']
    
    # Convert numeric columns
    if 'Unit Rate' in df.columns:
        df['Unit Rate'] = pd.to_numeric(df['Unit Rate'], errors='coerce')
    if 'Quantity' in df.columns:
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
    
    # Calculate total cost estimate
    if 'Unit Rate' in df.columns and 'Days Onsite' in df.columns and 'Billing Basis' in df.columns:
        df['Estimated Total Cost'] = 0
        for idx, row in df.iterrows():
            days = row['Days Onsite'] if pd.notnull(row['Days Onsite']) else 0
            rate = row['Unit Rate'] if pd.notnull(row['Unit Rate']) else 0
            
            if pd.notnull(row['Billing Basis']):
                if row['Billing Basis'] == 'Daily':
                    df.at[idx, 'Estimated Total Cost'] = days * rate
                elif row['Billing Basis'] == 'Weekly':
                    df.at[idx, 'Estimated Total Cost'] = (days / 7) * rate
                elif row['Billing Basis'] == 'Monthly':
                    df.at[idx, 'Estimated Total Cost'] = (days / 30) * rate
    
    return df

def calculate_kpis(df):
    """Calculate key performance indicators"""
    total_equipment = len(df)
    
    if 'Current Status' in df.columns:
        active = len(df[df['Current Status'].str.contains('Active', case=False, na=False)])
        idle = len(df[df['Current Status'].str.contains('Idle', case=False, na=False)])
        maintenance = len(df[df['Current Status'].str.contains('Maintenance', case=False, na=False)])
    else:
        active = idle = maintenance = 0
    
    total_cost = df['Estimated Total Cost'].sum() if 'Estimated Total Cost' in df.columns else 0
    
    # Count alerts
    alerts = 0
    if 'Next Inspection Due' in df.columns:
        overdue = df['Next Inspection Due'] < datetime.now()
        alerts += overdue.sum()
    
    if 'Duration Variance' in df.columns:
        over_duration = df['Duration Variance'] > 7
        alerts += over_duration.sum()
    
    return {
        'total': total_equipment,
        'active': active,
        'idle': idle,
        'maintenance': maintenance,
        'total_cost': total_cost,
        'alerts': alerts
    }

def create_status_chart(df):
    """Create equipment status distribution chart"""
    if 'Current Status' not in df.columns:
        return None
    
    status_counts = df['Current Status'].value_counts()
    
    colors = {
        'Active': '#2E7D32',
        'Idle': '#FF9800',
        'Under Maintenance': '#D32F2F',
        'Demobilized': '#757575'
    }
    
    fig = go.Figure(data=[go.Pie(
        labels=status_counts.index,
        values=status_counts.values,
        marker=dict(colors=[colors.get(status, '#1B4D89') for status in status_counts.index]),
        hole=0.4,
        textinfo='label+percent+value',
        textposition='auto'
    )])
    
    fig.update_layout(
        title='Equipment Status Distribution',
        showlegend=True,
        height=400,
        paper_bgcolor='white',
        plot_bgcolor='white'
    )
    
    return fig

def create_vendor_chart(df):
    """Create vendor analysis chart"""
    if 'Vendor' not in df.columns:
        return None
    
    vendor_data = df.groupby('Vendor').agg({
        'Equipment Description': 'count',
        'Estimated Total Cost': 'sum'
    }).reset_index()
    
    vendor_data.columns = ['Vendor', 'Equipment Count', 'Total Cost']
    vendor_data = vendor_data.sort_values('Total Cost', ascending=False).head(10)
    
    fig = go.Figure(data=[
        go.Bar(
            x=vendor_data['Vendor'],
            y=vendor_data['Total Cost'],
            marker_color='#1B4D89',
            text=vendor_data['Equipment Count'],
            texttemplate='%{text} items',
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Total Cost: $%{y:,.2f}<br>Equipment: %{text}<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title='Top 10 Vendors by Cost',
        xaxis_title='Vendor',
        yaxis_title='Total Cost ($)',
        showlegend=False,
        height=400,
        paper_bgcolor='white',
        plot_bgcolor='white',
        xaxis_tickangle=-45
    )
    
    return fig

def create_timeline_chart(df):
    """Create equipment timeline"""
    if 'Mobilization Date' not in df.columns or 'Equipment Description' not in df.columns:
        return None
    
    timeline_df = df[['Equipment Description', 'Vendor', 'Mobilization Date', 
                      'Planned Demob Date', 'Current Status']].copy()
    timeline_df = timeline_df.dropna(subset=['Mobilization Date'])
    
    # Use actual or planned demob date
    timeline_df['End Date'] = timeline_df['Planned Demob Date'].fillna(datetime.now() + timedelta(days=30))
    
    fig = px.timeline(
        timeline_df,
        x_start='Mobilization Date',
        x_end='End Date',
        y='Equipment Description',
        color='Current Status',
        hover_data=['Vendor'],
        color_discrete_map={
            'Active': '#2E7D32',
            'Idle': '#FF9800',
            'Under Maintenance': '#D32F2F',
            'Demobilized': '#757575'
        }
    )
    
    fig.update_layout(
        title='Equipment Timeline',
        height=600,
        xaxis_title='Date',
        yaxis_title='Equipment',
        showlegend=True,
        paper_bgcolor='white',
        plot_bgcolor='white'
    )
    
    fig.update_yaxes(categoryorder='total ascending')
    
    return fig

def create_category_chart(df):
    """Create category distribution"""
    if 'Category' not in df.columns:
        return None
    
    category_counts = df['Category'].value_counts().head(10)
    
    fig = go.Figure(data=[go.Bar(
        x=category_counts.values,
        y=category_counts.index,
        orientation='h',
        marker_color='#2E7D32',
        text=category_counts.values,
        textposition='outside'
    )])
    
    fig.update_layout(
        title='Equipment by Category',
        xaxis_title='Count',
        yaxis_title='Category',
        showlegend=False,
        height=400,
        paper_bgcolor='white',
        plot_bgcolor='white'
    )
    
    return fig

def generate_summary_report(df):
    """Generate summary report for download"""
    summary = {
        'Report Generated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'Total Equipment': len(df),
        'Active Equipment': len(df[df['Current Status'].str.contains('Active', case=False, na=False)]) if 'Current Status' in df.columns else 0,
        'Total Vendors': df['Vendor'].nunique() if 'Vendor' in df.columns else 0,
        'Total Estimated Cost': f"${df['Estimated Total Cost'].sum():,.2f}" if 'Estimated Total Cost' in df.columns else 0
    }
    
    summary_df = pd.DataFrame([summary])
    return summary_df

def main():
    # Header with Bird Construction branding
    col1, col2, col3 = st.columns([1, 3, 1])
    with col2:
        st.markdown("""
            <div style='text-align: center; padding: 20px;'>
                <h1 style='color: #1B4D89; font-size: 42px; margin-bottom: 5px;'>🏗️ Bird Construction</h1>
                <h2 style='color: #2E7D32; font-size: 24px; margin-top: 0;'>Equipment Tracking Dashboard</h2>
                <p style='color: #666; font-size: 16px;'>API Early Works Project - Real-Time Monitoring</p>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 📁 Upload Equipment Data")
        uploaded_file = st.file_uploader(
            "Upload Excel or CSV file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your equipment tracking spreadsheet"
        )
        
        if uploaded_file:
            st.markdown("<div class='success-box'>✅ File loaded successfully!</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 🔍 Filters")
    
    # Main content
    if uploaded_file is None:
        st.markdown("""
            <div style='text-align: center; padding: 60px 20px; background: linear-gradient(135deg, #E3F2FD 0%, #E8F5E9 100%); border-radius: 12px; margin: 40px 0;'>
                <h2 style='color: #1B4D89; margin-bottom: 20px;'>👈 Upload Your Equipment Data to Begin</h2>
                <p style='color: #666; font-size: 18px;'>Upload your Excel or CSV file using the sidebar</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Features overview
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown("""
                <div class='metric-card'>
                    <h3 style='color: #1B4D89;'>📊 Real-Time KPIs</h3>
                    <p style='color: #666;'>Track equipment status, costs, and utilization instantly</p>
                </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
                <div class='metric-card'>
                    <h3 style='color: #2E7D32;'>📈 Analytics</h3>
                    <p style='color: #666;'>Vendor performance and cost breakdowns</p>
                </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
                <div class='metric-card'>
                    <h3 style='color: #FF9800;'>⚠️ Alerts</h3>
                    <p style='color: #666;'>Inspection due dates and duration tracking</p>
                </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown("""
                <div class='metric-card'>
                    <h3 style='color: #FFEB3B;'>📥 Reports</h3>
                    <p style='color: #666;'>Download detailed analytics and summaries</p>
                </div>
            """, unsafe_allow_html=True)
        
        return
    
    # Load data
    try:
        df = load_data(uploaded_file)
        
        # Sidebar filters
        with st.sidebar:
            if 'Vendor' in df.columns:
                vendors = ['All'] + sorted(df['Vendor'].dropna().unique().tolist())
                selected_vendor = st.selectbox("🏢 Vendor", vendors)
            else:
                selected_vendor = 'All'
            
            if 'Current Status' in df.columns:
                statuses = ['All'] + sorted(df['Current Status'].dropna().unique().tolist())
                selected_status = st.selectbox("📊 Status", statuses)
            else:
                selected_status = 'All'
            
            if 'Category' in df.columns:
                categories = ['All'] + sorted(df['Category'].dropna().unique().tolist())
                selected_category = st.selectbox("🏷️ Category", categories)
            else:
                selected_category = 'All'
            
            if 'Payment Type' in df.columns:
                payment_types = ['All'] + sorted(df['Payment Type'].dropna().unique().tolist())
                selected_payment = st.selectbox("💰 Payment Type", payment_types)
            else:
                selected_payment = 'All'
        
        # Apply filters
        filtered_df = df.copy()
        if selected_vendor != 'All':
            filtered_df = filtered_df[filtered_df['Vendor'] == selected_vendor]
        if selected_status != 'All':
            filtered_df = filtered_df[filtered_df['Current Status'] == selected_status]
        if selected_category != 'All':
            filtered_df = filtered_df[filtered_df['Category'] == selected_category]
        if selected_payment != 'All':
            filtered_df = filtered_df[filtered_df['Payment Type'] == selected_payment]
        
        # Calculate KPIs
        kpis = calculate_kpis(filtered_df)
        
        # KPI Section
        st.markdown("### 📊 Key Performance Indicators")
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        with col1:
            st.metric("Total Equipment", f"{kpis['total']:,}")
        with col2:
            st.metric("🟢 Active", f"{kpis['active']:,}")
        with col3:
            st.metric("🟡 Idle", f"{kpis['idle']:,}")
        with col4:
            st.metric("🔴 Maintenance", f"{kpis['maintenance']:,}")
        with col5:
            st.metric("💰 Total Cost", f"${kpis['total_cost']:,.0f}")
        with col6:
            st.metric("⚠️ Alerts", f"{kpis['alerts']:,}")
        
        st.markdown("---")
        
        # Create tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Dashboard Overview",
            "📈 Vendor Analytics",
            "📅 Timeline",
            "📋 Equipment Details",
            "📥 Download Reports"
        ])
        
        # TAB 1: Dashboard Overview
        with tab1:
            col1, col2 = st.columns(2)
            
            with col1:
                status_fig = create_status_chart(filtered_df)
                if status_fig:
                    st.plotly_chart(status_fig, use_container_width=True)
            
            with col2:
                category_fig = create_category_chart(filtered_df)
                if category_fig:
                    st.plotly_chart(category_fig, use_container_width=True)
            
            # Payment Type Analysis
            if 'Payment Type' in filtered_df.columns and 'Billing Basis' in filtered_df.columns:
                st.markdown("### 💰 Payment Structure Analysis")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    payment_counts = filtered_df['Payment Type'].value_counts()
                    fig = go.Figure(data=[go.Pie(
                        labels=payment_counts.index,
                        values=payment_counts.values,
                        marker=dict(colors=['#1B4D89', '#2E7D32', '#FF9800', '#FFEB3B']),
                        hole=0.4
                    )])
                    fig.update_layout(title='Equipment by Payment Type', height=350)
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    billing_counts = filtered_df['Billing Basis'].value_counts()
                    fig = go.Figure(data=[go.Bar(
                        x=billing_counts.index,
                        y=billing_counts.values,
                        marker_color='#2E7D32',
                        text=billing_counts.values,
                        textposition='outside'
                    )])
                    fig.update_layout(
                        title='Equipment by Billing Basis',
                        xaxis_title='Billing Basis',
                        yaxis_title='Count',
                        height=350,
                        showlegend=False
                    )
                    st.plotly_chart(fig, use_container_width=True)
            
            # Alerts Section
            if kpis['alerts'] > 0:
                st.markdown("### ⚠️ Action Required")
                
                alert_items = []
                
                if 'Next Inspection Due' in filtered_df.columns:
                    overdue = filtered_df[filtered_df['Next Inspection Due'] < datetime.now()]
                    if len(overdue) > 0:
                        st.markdown(f"""
                            <div class='alert-card'>
                                <strong>🔴 {len(overdue)} equipment items have overdue inspections</strong>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        with st.expander(f"View {len(overdue)} Overdue Inspections"):
                            st.dataframe(
                                overdue[['Equipment Description', 'Vendor', 'Next Inspection Due']],
                                use_container_width=True
                            )
                
                if 'Duration Variance' in filtered_df.columns:
                    over_duration = filtered_df[filtered_df['Duration Variance'] > 7]
                    if len(over_duration) > 0:
                        st.markdown(f"""
                            <div class='alert-card'>
                                <strong>🟡 {len(over_duration)} equipment items exceed planned duration by >7 days</strong>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        with st.expander(f"View {len(over_duration)} Over Duration Items"):
                            st.dataframe(
                                over_duration[['Equipment Description', 'Vendor', 'Duration Variance']],
                                use_container_width=True
                            )
        
        # TAB 2: Vendor Analytics
        with tab2:
            st.markdown("### 🏢 Vendor Performance Analysis")
            
            vendor_fig = create_vendor_chart(filtered_df)
            if vendor_fig:
                st.plotly_chart(vendor_fig, use_container_width=True)
            
            # Vendor summary table
            if 'Vendor' in filtered_df.columns:
                st.markdown("### 📊 Detailed Vendor Metrics")
                
                vendor_summary = filtered_df.groupby('Vendor').agg({
                    'Equipment Description': 'count',
                    'Estimated Total Cost': 'sum',
                    'Days Onsite': 'mean'
                }).reset_index()
                
                vendor_summary.columns = ['Vendor', 'Equipment Count', 'Total Cost', 'Avg Days Onsite']
                vendor_summary = vendor_summary.sort_values('Total Cost', ascending=False)
                
                st.dataframe(
                    vendor_summary.style.format({
                        'Total Cost': '${:,.2f}',
                        'Avg Days Onsite': '{:.1f}'
                    }).background_gradient(subset=['Total Cost'], cmap='YlOrRd'),
                    use_container_width=True,
                    height=400
                )
        
        # TAB 3: Timeline
        with tab3:
            st.markdown("### 📅 Equipment Mobilization Timeline")
            
            timeline_fig = create_timeline_chart(filtered_df)
            if timeline_fig:
                st.plotly_chart(timeline_fig, use_container_width=True)
            else:
                st.info("Timeline requires Mobilization Date data")
        
        # TAB 4: Equipment Details
        with tab4:
            st.markdown("### 📋 Detailed Equipment Information")
            
            # Column selector
            all_columns = filtered_df.columns.tolist()
            default_columns = ['Equipment Description', 'Vendor', 'Current Status', 
                             'Mobilization Date', 'Unit Rate', 'Estimated Total Cost']
            available_defaults = [col for col in default_columns if col in all_columns]
            
            selected_columns = st.multiselect(
                "Select columns to display",
                options=all_columns,
                default=available_defaults
            )
            
            if selected_columns:
                # Search functionality
                search = st.text_input("🔍 Search equipment", "")
                
                display_df = filtered_df[selected_columns].copy()
                
                if search:
                    mask = display_df.astype(str).apply(
                        lambda x: x.str.contains(search, case=False, na=False)
                    ).any(axis=1)
                    display_df = display_df[mask]
                
                st.markdown(f"**Showing {len(display_df)} of {len(filtered_df)} items**")
                
                st.dataframe(
                    display_df,
                    use_container_width=True,
                    height=500
                )
        
        # TAB 5: Download Reports
        with tab5:
            st.markdown("""
                <div class='download-section'>
                    <h2 style='color: #1B4D89; text-align: center; margin-bottom: 10px;'>📥 Download Analytics & Reports</h2>
                    <p style='text-align: center; color: #666;'>Export your data and analysis for offline review</p>
                </div>
            """, unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            
            # Download 1: Filtered Data
            with col1:
                st.markdown("### 📊 Filtered Equipment Data")
                st.write(f"Current view: **{len(filtered_df)} items**")
                
                csv_data = filtered_df.to_csv(index=False)
                st.download_button(
                    label="📥 Download Filtered Data (CSV)",
                    data=csv_data,
                    file_name=f"equipment_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            # Download 2: Vendor Summary
            with col2:
                st.markdown("### 🏢 Vendor Summary Report")
                st.write("Performance metrics by vendor")
                
                if 'Vendor' in filtered_df.columns:
                    vendor_summary = filtered_df.groupby('Vendor').agg({
                        'Equipment Description': 'count',
                        'Estimated Total Cost': 'sum',
                        'Days Onsite': 'mean'
                    }).reset_index()
                    
                    vendor_summary.columns = ['Vendor', 'Equipment Count', 'Total Cost', 'Avg Days Onsite']
                    
                    csv_vendor = vendor_summary.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Vendor Report (CSV)",
                        data=csv_vendor,
                        file_name=f"vendor_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
            
            # Download 3: Summary Report
            with col3:
                st.markdown("### 📈 Executive Summary")
                st.write("High-level KPI report")
                
                summary_df = generate_summary_report(filtered_df)
                csv_summary = summary_df.to_csv(index=False)
                
                st.download_button(
                    label="📥 Download Summary (CSV)",
                    data=csv_summary,
                    file_name=f"executive_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            st.markdown("---")
            
            # Download 4: Full Dataset
            st.markdown("### 📚 Complete Dataset")
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.write(f"Export all **{len(df)} equipment records** with all columns")
            
            with col2:
                csv_full = df.to_csv(index=False)
                st.download_button(
                    label="📥 Download Complete Dataset",
                    data=csv_full,
                    file_name=f"equipment_complete_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            # Download 5: Excel Format
            st.markdown("### 📊 Excel Format Export")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, sheet_name='Equipment Data', index=False)
                
                if 'Vendor' in filtered_df.columns:
                    vendor_summary.to_excel(writer, sheet_name='Vendor Summary', index=False)
                
                summary_df.to_excel(writer, sheet_name='Executive Summary', index=False)
            
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Multi-Sheet Excel Report",
                data=excel_data,
                file_name=f"equipment_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.info("Please ensure your Excel file has the correct column structure")

if __name__ == "__main__":
    main()


