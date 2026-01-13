import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np

# Page configuration
st.set_page_config(
    page_title="Equipment Tracking Dashboard",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for dark theme
st.markdown("""
    <style>
    .metric-card {
        background-color: #1e1e1e;
        padding: 20px;
        border-radius: 10px;
        border-left: 4px solid #4CAF50;
    }
    .stMetric {
        background-color: #2d2d2d;
        padding: 15px;
        border-radius: 8px;
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data(file):
    """Load and preprocess equipment tracking data"""
    df = pd.read_excel(file)
    
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Convert date columns
    date_columns = ['Mobilization Date', 'LEM Remove Demobilization Date', 'Date (Jenny)']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Extract phase codes and categories
    if 'Phase Code (From Previous LEM)' in df.columns:
        df['Phase_Category'] = df['Phase Code (From Previous LEM)'].str.extract(r'TP\d+ - (.+?)(?:\s|$)')
        df['Phase_Code'] = df['Phase Code (From Previous LEM)'].str.extract(r'(50 \d+ \d+ \d+ \d+)')
    
    # Calculate duration
    if 'Mobilization Date' in df.columns and 'LEM Remove Demobilization Date' in df.columns:
        df['Duration_Days'] = (df['LEM Remove Demobilization Date'] - df['Mobilization Date']).dt.days
    
    # Convert monthly cost to numeric
    if 'VTC Monthly Cost' in df.columns:
        df['VTC Monthly Cost'] = pd.to_numeric(df['VTC Monthly Cost'], errors='coerce')
    
    # Calculate total cost based on duration
    if 'Duration_Days' in df.columns and 'VTC Monthly Cost' in df.columns:
        df['Total_Cost'] = (df['Duration_Days'] / 30) * df['VTC Monthly Cost']
    
    return df

def calculate_kpis(df):
    """Calculate key performance indicators"""
    total_equipment = len(df)
    active_equipment = df[df['LEM Remove Demobilization Date'].isna() | 
                         (df['LEM Remove Demobilization Date'] > datetime.now())].shape[0]
    total_cost = df['Total_Cost'].sum() if 'Total_Cost' in df.columns else 0
    avg_duration = df['Duration_Days'].mean() if 'Duration_Days' in df.columns else 0
    
    return {
        'total_equipment': total_equipment,
        'active_equipment': active_equipment,
        'total_cost': total_cost,
        'avg_duration': avg_duration
    }

def vendor_analysis(df):
    """Analyze equipment by vendor"""
    if 'Vendor' not in df.columns:
        return None
    
    vendor_stats = df.groupby('Vendor').agg({
        'Equipment Description': 'count',
        'Total_Cost': 'sum',
        'Duration_Days': 'mean',
        'VTC Monthly Cost': 'sum'
    }).reset_index()
    
    vendor_stats.columns = ['Vendor', 'Equipment_Count', 'Total_Cost', 'Avg_Duration', 'Monthly_Cost']
    return vendor_stats.sort_values('Total_Cost', ascending=False)

def create_timeline_chart(df):
    """Create Gantt-style timeline for equipment mobilization"""
    timeline_df = df[['Equipment Description', 'Vendor', 'Mobilization Date', 
                      'LEM Remove Demobilization Date']].dropna(subset=['Mobilization Date'])
    
    fig = px.timeline(
        timeline_df,
        x_start='Mobilization Date',
        x_end='LEM Remove Demobilization Date',
        y='Equipment Description',
        color='Vendor',
        title='Equipment Mobilization Timeline',
        hover_data=['Vendor']
    )
    
    fig.update_yaxes(categoryorder="total ascending")
    fig.update_layout(height=600, showlegend=True)
    return fig

def main():
    st.title("🏗️ Equipment Tracking & Analytics Dashboard")
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("📁 Data Upload")
        uploaded_file = st.file_uploader("Upload Equipment Tracking Excel", type=['xlsx', 'xls'])
        
        if uploaded_file:
            st.success("File uploaded successfully!")
        
        st.markdown("---")
        st.header("🔍 Filters")
    
    # Main content
    if uploaded_file is None:
        st.info("👈 Please upload your Equipment Tracking Log Excel file to begin analysis")
        st.markdown("""
        ### Dashboard Features:
        - **Real-time KPIs**: Track total equipment, active equipment, costs
        - **Vendor Analysis**: Compare vendor performance and costs
        - **Timeline Visualization**: Gantt chart of equipment mobilization
        - **Phase Code Analytics**: Equipment distribution by project phase
        - **Cost Breakdown**: Monthly and total cost analysis
        - **Export Capabilities**: Download filtered data and reports
        """)
        return
    
    # Load data
    try:
        df = load_data(uploaded_file)
        
        # Sidebar filters
        with st.sidebar:
            vendors = ['All'] + sorted(df['Vendor'].dropna().unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendors)
            
            if 'Phase_Category' in df.columns:
                phases = ['All'] + sorted(df['Phase_Category'].dropna().unique().tolist())
                selected_phase = st.selectbox("Select Phase", phases)
            else:
                selected_phase = 'All'
            
            date_range = st.date_input(
                "Date Range",
                value=(df['Mobilization Date'].min(), datetime.now()),
                key='date_range'
            )
        
        # Apply filters
        filtered_df = df.copy()
        if selected_vendor != 'All':
            filtered_df = filtered_df[filtered_df['Vendor'] == selected_vendor]
        if selected_phase != 'All' and 'Phase_Category' in df.columns:
            filtered_df = filtered_df[filtered_df['Phase_Category'] == selected_phase]
        
        # KPI Section
        st.header("📊 Key Performance Indicators")
        kpis = calculate_kpis(filtered_df)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Equipment", f"{kpis['total_equipment']:,}")
        with col2:
            st.metric("Active Equipment", f"{kpis['active_equipment']:,}")
        with col3:
            st.metric("Total Cost", f"${kpis['total_cost']:,.2f}")
        with col4:
            st.metric("Avg Duration", f"{kpis['avg_duration']:.0f} days")
        
        st.markdown("---")
        
        # Vendor Analysis Section
        st.header("🏢 Vendor Analysis")
        vendor_stats = vendor_analysis(filtered_df)
        
        if vendor_stats is not None:
            col1, col2 = st.columns(2)
            
            with col1:
                # Vendor cost breakdown
                fig_vendor_cost = px.bar(
                    vendor_stats,
                    x='Vendor',
                    y='Total_Cost',
                    title='Total Cost by Vendor',
                    color='Total_Cost',
                    color_continuous_scale='Viridis'
                )
                st.plotly_chart(fig_vendor_cost, use_container_width=True)
            
            with col2:
                # Equipment count by vendor
                fig_vendor_count = px.pie(
                    vendor_stats,
                    values='Equipment_Count',
                    names='Vendor',
                    title='Equipment Distribution by Vendor'
                )
                st.plotly_chart(fig_vendor_count, use_container_width=True)
            
            # Vendor statistics table
            st.subheader("Vendor Performance Metrics")
            st.dataframe(
                vendor_stats.style.format({
                    'Total_Cost': '${:,.2f}',
                    'Avg_Duration': '{:.1f}',
                    'Monthly_Cost': '${:,.2f}'
                }),
                use_container_width=True
            )
        
        st.markdown("---")
        
        # Timeline Section
        st.header("📅 Equipment Timeline")
        if 'Mobilization Date' in filtered_df.columns:
            timeline_fig = create_timeline_chart(filtered_df)
            st.plotly_chart(timeline_fig, use_container_width=True)
        
        st.markdown("---")
        
        # Phase Code Analysis
        if 'Phase_Category' in filtered_df.columns:
            st.header("🔧 Phase Code Analysis")
            col1, col2 = st.columns(2)
            
            with col1:
                phase_counts = filtered_df['Phase_Category'].value_counts()
                fig_phase = px.bar(
                    x=phase_counts.index,
                    y=phase_counts.values,
                    title='Equipment Count by Phase Category',
                    labels={'x': 'Phase Category', 'y': 'Count'}
                )
                st.plotly_chart(fig_phase, use_container_width=True)
            
            with col2:
                if 'Total_Cost' in filtered_df.columns:
                    phase_costs = filtered_df.groupby('Phase_Category')['Total_Cost'].sum()
                    fig_phase_cost = px.pie(
                        values=phase_costs.values,
                        names=phase_costs.index,
                        title='Cost Distribution by Phase'
                    )
                    st.plotly_chart(fig_phase_cost, use_container_width=True)
        
        st.markdown("---")
        
        # Detailed Data Table
        st.header("📋 Detailed Equipment Data")
        
        # Column selection
        display_columns = st.multiselect(
            "Select columns to display",
            options=df.columns.tolist(),
            default=['Equipment Description', 'Vendor', 'VTC Equipment #', 
                    'Mobilization Date', 'VTC Monthly Cost', 'Phase_Category']
        )
        
        if display_columns:
            st.dataframe(
                filtered_df[display_columns].style.format({
                    'VTC Monthly Cost': '${:,.2f}',
                    'Total_Cost': '${:,.2f}',
                    'Mobilization Date': lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) else '',
                    'LEM Remove Demobilization Date': lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) else ''
                }),
                use_container_width=True,
                height=400
            )
        
        # Export section
        st.markdown("---")
        st.header("📥 Export Data")
        col1, col2 = st.columns(2)
        
        with col1:
            # Export filtered data
            csv = filtered_df.to_csv(index=False)
            st.download_button(
                label="Download Filtered Data (CSV)",
                data=csv,
                file_name=f"equipment_tracking_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        with col2:
            # Export vendor summary
            if vendor_stats is not None:
                vendor_csv = vendor_stats.to_csv(index=False)
                st.download_button(
                    label="Download Vendor Summary (CSV)",
                    data=vendor_csv,
                    file_name=f"vendor_summary_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.info("Please ensure your Excel file has the correct format and column headers.")

if __name__ == "__main__":
    main()
