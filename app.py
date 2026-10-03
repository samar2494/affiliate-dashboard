import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# PAGE CONFIG
# ============================================================================
st.set_page_config(page_title="Affiliate Retention Command Center", layout="wide", initial_sidebar_state="expanded")

# ============================================================================
# STYLING
# ============================================================================
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .metric-value {
        font-size: 32px;
        font-weight: bold;
        margin: 10px 0;
    }
    .metric-label {
        font-size: 12px;
        opacity: 0.9;
    }
    .status-growing { color: #2ecc71; font-weight: bold; }
    .status-stable { color: #f39c12; font-weight: bold; }
    .status-watch { color: #e74c3c; font-weight: bold; }
    .status-red { color: #c0392b; font-weight: bold; }
    .status-new { color: #95a5a6; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def load_excel_data(file):
    """Load Excel file with auto-detection of sheet names"""
    try:
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        
        # Try common sheet names
        data_sheet = None
        for sheet in sheet_names:
            if 'affiliate' in sheet.lower() or 'data' in sheet.lower() or 'partners' in sheet.lower():
                data_sheet = sheet
                break
        
        if not data_sheet:
            data_sheet = sheet_names[0]
        
        df = pd.read_excel(file, sheet_name=data_sheet)
        return df, data_sheet
    except Exception as e:
        st.error(f"Error loading Excel: {e}")
        return None, None

def clean_dataframe(df):
    """Clean and standardize the dataframe"""
    # Standardize column names
    df.columns = df.columns.str.strip().str.lower()
    
    # Map common variations
    column_mapping = {
        'partner id': 'partner_id',
        'partner name': 'partner_name',
        'geo': 'geo',
        'am': 'am',
        'list': 'list',
        'traffic source': 'traffic_source',
        'ftd current month': 'ftd_current',
        'ftd previous month': 'ftd_previous',
        'ftd 3-month baseline': 'ftd_baseline_3m',
        'decline reason': 'decline_reason',
        'am feedback': 'am_feedback',
        'proposed action': 'proposed_action',
        'first added': 'first_added',
        'status': 'status'
    }
    
    df.columns = [column_mapping.get(col, col) for col in df.columns]
    
    # Fill blanks with 0 for FTD columns, NaN for text columns
    ftd_cols = ['ftd_current', 'ftd_previous', 'ftd_baseline_3m']
    for col in ftd_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    # Fill text columns with empty string
    text_cols = ['decline_reason', 'am_feedback', 'proposed_action', 'traffic_source']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna('')
    
    # Ensure required columns exist
    required_cols = ['partner_id', 'partner_name', 'geo', 'am', 'list', 'ftd_current', 'ftd_previous', 'ftd_baseline_3m']
    for col in required_cols:
        if col not in df.columns:
            df[col] = 0 if 'ftd' in col else ''
    
    return df

def calculate_performance_metrics(df):
    """Calculate all performance metrics"""
    df = df.copy()
    
    # FTD calculations
    df['ftd_mom_change'] = df['ftd_current'] - df['ftd_previous']
    df['ftd_mom_pct'] = np.where(
        df['ftd_previous'] > 0,
        ((df['ftd_current'] - df['ftd_previous']) / df['ftd_previous'] * 100),
        0
    )
    
    df['ftd_vs_baseline_pct'] = np.where(
        df['ftd_baseline_3m'] > 0,
        ((df['ftd_current'] - df['ftd_baseline_3m']) / df['ftd_baseline_3m'] * 100),
        0
    )
    
    # Classification logic
    def classify_partner(row):
        if row['ftd_baseline_3m'] == 0:
            return 'New/Insufficient Baseline'
        
        baseline_pct = row['ftd_vs_baseline_pct']
        
        if baseline_pct >= 15:
            return 'Growing'
        elif baseline_pct >= -15:
            return 'Stable'
        elif baseline_pct >= -30:
            return 'Watch'
        else:
            return 'Red'
    
    df['performance_status'] = df.apply(classify_partner, axis=1)
    
    return df

def get_performance_color(status):
    """Return color for performance status"""
    colors = {
        'Growing': '#2ecc71',
        'Stable': '#f39c12',
        'Watch': '#e74c3c',
        'Red': '#c0392b',
        'New/Insufficient Baseline': '#95a5a6'
    }
    return colors.get(status, '#95a5a6')

def detect_list_movements(current_df, previous_df):
    """Detect partner list movements between months"""
    movements = []
    
    for _, current_row in current_df.iterrows():
        partner_id = current_row['partner_id']
        current_list = current_row['list']
        
        # Check if first_added is this month (indicates new addition to this list)
        first_added = current_row.get('first_added', '')
        if first_added and str(first_added).strip():
            # This partner was added to this list this month
            previous_list = previous_df[previous_df['partner_id'] == partner_id]['list'].values
            previous_list = previous_list[0] if len(previous_list) > 0 else 'New Partner'
            
            if previous_list != current_list:
                movements.append({
                    'partner_id': partner_id,
                    'partner_name': current_row['partner_name'],
                    'from_list': previous_list,
                    'to_list': current_list,
                    'date_moved': first_added,
                    'geo': current_row['geo'],
                    'am': current_row['am']
                })
    
    return pd.DataFrame(movements) if movements else pd.DataFrame()

def identify_missing_data(df):
    """Identify data quality issues"""
    issues = []
    
    for idx, row in df.iterrows():
        if not row.get('traffic_source') or row.get('traffic_source') == '':
            issues.append({
                'type': 'Missing Traffic Source',
                'partner_id': row['partner_id'],
                'partner_name': row['partner_name'],
                'severity': 'Medium'
            })
        
        if not row.get('am_feedback') or row.get('am_feedback') == '':
            if row['performance_status'] in ['Watch', 'Red']:
                issues.append({
                    'type': 'Missing AM Feedback (High Priority)',
                    'partner_id': row['partner_id'],
                    'partner_name': row['partner_name'],
                    'severity': 'High'
                })
        
        if row['ftd_current'] == 0 and row['list'] == 'Main Tops':
            issues.append({
                'type': 'Zero FTD - Main Tops Partner',
                'partner_id': row['partner_id'],
                'partner_name': row['partner_name'],
                'severity': 'Critical'
            })
    
    return pd.DataFrame(issues) if issues else pd.DataFrame()

# ============================================================================
# MAIN APP
# ============================================================================

def main():
    st.title("🎯 Affiliate Retention Command Center")
    st.subheader("FTD Performance Analysis & Partner Management Dashboard")
    
    # ========================================================================
    # SIDEBAR - DATA IMPORT
    # ========================================================================
    with st.sidebar:
        st.header("📊 Data Import")
        
        uploaded_file = st.file_uploader("📁 Import Latest Excel", type=['xlsx', 'xls'], key='main_upload')
        
        if uploaded_file:
            df_current, sheet_name = load_excel_data(uploaded_file)
            
            if df_current is not None:
                df_current = clean_dataframe(df_current)
                df_current = calculate_performance_metrics(df_current)
                
                st.success(f"✅ Loaded {len(df_current)} partners from '{sheet_name}'")
                
                # Store in session state
                st.session_state['df_current'] = df_current
                st.session_state['upload_date'] = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        # Previous month upload for list movement tracking
        st.subheader("📅 Previous Month (Optional)")
        previous_file = st.file_uploader("📁 Upload Previous Month", type=['xlsx', 'xls'], key='prev_upload')
        
        if previous_file:
            df_previous, _ = load_excel_data(previous_file)
            if df_previous is not None:
                df_previous = clean_dataframe(df_previous)
                st.session_state['df_previous'] = df_previous
                st.success("✅ Previous month data loaded")
        
        st.divider()
        st.info("💡 Upload current month Excel to populate dashboard. Previous month is optional for list movement tracking.")
    
    # Check if data loaded
    if 'df_current' not in st.session_state:
        st.info("👈 Please upload Excel file in sidebar to begin")
        return
    
    df = st.session_state['df_current']
    
    # ========================================================================
    # PAGE NAVIGATION
    # ========================================================================
    pages = [
        "📊 Executive Dashboard",
        "🌍 GEO Performance",
        "📈 Growth & Decline",
        "🔄 List Movement",
        "⚡ Action Center",
        "🔍 Partner Tracker",
        "👥 AM Visibility",
        "📋 Data Quality"
    ]
    
    selected_page = st.sidebar.radio("Navigation", pages)
    
    # ========================================================================
    # PAGE 1: EXECUTIVE DASHBOARD
    # ========================================================================
    if selected_page == "📊 Executive Dashboard":
        st.header("Executive Dashboard")
        
        # Key metrics
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        total_ftd_current = df['ftd_current'].sum()
        total_ftd_previous = df['ftd_previous'].sum()
        total_ftd_baseline = df['ftd_baseline_3m'].sum()
        
        ftd_gained = df[df['ftd_current'] > df['ftd_previous']]['ftd_mom_change'].sum()
        ftd_lost = abs(df[df['ftd_current'] < df['ftd_previous']]['ftd_mom_change'].sum())
        ftd_net = total_ftd_current - total_ftd_previous
        
        mom_pct = ((total_ftd_current - total_ftd_previous) / total_ftd_previous * 100) if total_ftd_previous > 0 else 0
        baseline_pct = ((total_ftd_current - total_ftd_baseline) / total_ftd_baseline * 100) if total_ftd_baseline > 0 else 0
        
        with col1:
            st.metric("Current FTD", f"{total_ftd_current:,}", delta=f"{ftd_net:+,}")
        with col2:
            st.metric("Previous FTD", f"{total_ftd_previous:,}")
        with col3:
            st.metric("MoM Change %", f"{mom_pct:.1f}%", delta=f"{ftd_net:+,}" if ftd_net != 0 else "Flat")
        with col4:
            st.metric("vs 3M Baseline %", f"{baseline_pct:.1f}%")
        with col5:
            st.metric("FTD Gained", f"{ftd_gained:,}", delta_color="inverse")
        with col6:
            st.metric("FTD Lost", f"{ftd_lost:,}", delta_color="inverse")
        
        st.divider()
        
        # Performance distribution
        col1, col2 = st.columns(2)
        
        with col1:
            performance_counts = df['performance_status'].value_counts()
            fig = px.pie(values=performance_counts.values, names=performance_counts.index, 
                        title="Partner Performance Distribution",
                        color_discrete_map={
                            'Growing': '#2ecc71',
                            'Stable': '#f39c12',
                            'Watch': '#e74c3c',
                            'Red': '#c0392b',
                            'New/Insufficient Baseline': '#95a5a6'
                        })
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            list_counts = df['list'].value_counts()
            fig = px.bar(x=list_counts.index, y=list_counts.values, title="Partners by List",
                        labels={'x': 'List', 'y': 'Count'},
                        color=list_counts.index,
                        color_discrete_sequence=['#3498db', '#9b59b6', '#e67e22', '#95a5a6'])
            st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        
        # Management Summary
        st.subheader("📋 Management Summary")
        
        growing_count = len(df[df['performance_status'] == 'Growing'])
        watch_count = len(df[df['performance_status'] == 'Watch'])
        red_count = len(df[df['performance_status'] == 'Red'])
        main_tops_count = len(df[df['list'] == 'Main Tops'])
        potential_count = len(df[df['list'] == 'Potential Tops'])
        
        summary_text = f"""
        **Current Period Overview:**
        
        • **Total FTD Performance**: {total_ftd_current:,} FTDs this month ({mom_pct:+.1f}% MoM)
        • **Performance Health**: {growing_count} Growing partners, {watch_count} on Watch, {red_count} in Red zone
        • **Portfolio Structure**: {main_tops_count} Main Tops, {potential_count} Potential Tops
        • **Baseline Comparison**: {baseline_pct:+.1f}% vs 3-month average
        • **Net Change**: {ftd_net:+,} FTDs ({ftd_gained:,} gained, {ftd_lost:,} lost)
        """
        
        st.markdown(summary_text)
        
        st.divider()
        
        # Top 5 performers
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🚀 Top 5 Growing Partners")
            top_growing = df[df['performance_status'] == 'Growing'].nlargest(5, 'ftd_mom_change')[
                ['partner_name', 'ftd_current', 'ftd_mom_change', 'ftd_mom_pct', 'geo']
            ].reset_index(drop=True)
            top_growing['ftd_mom_pct'] = top_growing['ftd_mom_pct'].round(1).astype(str) + '%'
            top_growing.columns = ['Partner', 'Current FTD', 'MoM Change', 'MoM %', 'GEO']
            st.dataframe(top_growing, use_container_width=True, hide_index=True)
        
        with col2:
            st.subheader("📉 Top 5 Declining Partners")
            top_decline = df[df['performance_status'].isin(['Watch', 'Red'])].nsmallest(5, 'ftd_mom_change')[
                ['partner_name', 'ftd_current', 'ftd_mom_change', 'ftd_mom_pct', 'geo']
            ].reset_index(drop=True)
            top_decline['ftd_mom_pct'] = top_decline['ftd_mom_pct'].round(1).astype(str) + '%'
            top_decline.columns = ['Partner', 'Current FTD', 'MoM Change', 'MoM %', 'GEO']
            st.dataframe(top_decline, use_container_width=True, hide_index=True)
    
    # ========================================================================
    # PAGE 2: GEO PERFORMANCE
    # ========================================================================
    elif selected_page == "🌍 GEO Performance":
        st.header("Geographic Performance Analysis")
        
        # GEO summary
        geo_data = df.groupby('geo').agg({
            'partner_id': 'count',
            'ftd_current': 'sum',
            'ftd_previous': 'sum',
            'ftd_baseline_3m': 'sum'
        }).rename(columns={'partner_id': 'partner_count'})
        
        geo_data['ftd_mom_change'] = geo_data['ftd_current'] - geo_data['ftd_previous']
        geo_data['ftd_mom_pct'] = (geo_data['ftd_mom_change'] / geo_data['ftd_previous'] * 100).round(1)
        geo_data['ftd_vs_baseline_pct'] = ((geo_data['ftd_current'] - geo_data['ftd_baseline_3m']) / geo_data['ftd_baseline_3m'] * 100).round(1)
        
        st.dataframe(geo_data, use_container_width=True)
        
        st.divider()
        
        # GEO charts
        col1, col2 = st.columns(2)
        
        with col1:
            fig = px.bar(geo_data.reset_index(), x='geo', y='ftd_current', 
                        title="Current FTD by GEO",
                        labels={'geo': 'Geography', 'ftd_current': 'FTD Count'},
                        color='ftd_mom_pct', color_continuous_scale='RdYlGn')
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = px.bar(geo_data.reset_index(), x='geo', y='partner_count',
                        title="Partner Count by GEO",
                        labels={'geo': 'Geography', 'partner_count': 'Count'},
                        color='partner_count', color_continuous_scale='Viridis')
            st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        
        # Performance distribution by GEO
        st.subheader("Performance Distribution by GEO")
        
        for geo in df['geo'].unique():
            if pd.isna(geo):
                continue
            
            geo_df = df[df['geo'] == geo]
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(f"{geo} - FTD", f"{geo_df['ftd_current'].sum():,}", 
                         delta=f"{geo_df['ftd_mom_change'].sum():+,}")
            
            with col2:
                perf_counts = geo_df['performance_status'].value_counts()
                st.metric(f"{geo} - Growing", perf_counts.get('Growing', 0))
            
            with col3:
                st.metric(f"{geo} - Red Zone", perf_counts.get('Red', 0))
            
            # Drill-down table
            geo_partners = geo_df[['partner_name', 'list', 'ftd_current', 'ftd_mom_change', 
                                   'performance_status', 'am']].sort_values('ftd_current', ascending=False)
            st.dataframe(geo_partners, use_container_width=True, hide_index=True)
            st.divider()
    
    # ========================================================================
    # PAGE 3: GROWTH & DECLINE
    # ========================================================================
    elif selected_page == "📈 Growth & Decline":
        st.header("Growth & Decline Analysis")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🚀 Top Growers (by Absolute FTD Impact)")
            top_growers = df[df['ftd_mom_change'] > 0].nlargest(10, 'ftd_mom_change')[
                ['partner_name', 'list', 'ftd_current', 'ftd_mom_change', 'ftd_mom_pct', 
                 'traffic_source', 'am', 'am_feedback']
            ].reset_index(drop=True)
            top_growers['ftd_mom_pct'] = top_growers['ftd_mom_pct'].round(1).astype(str) + '%'
            top_growers.columns = ['Partner', 'List', 'Current FTD', 'MoM Change', 'MoM %', 'Source', 'AM', 'Feedback']
            st.dataframe(top_growers, use_container_width=True, hide_index=True)
        
        with col2:
            st.subheader("📉 Top Decliners (by Absolute FTD Loss)")
            top_decliners = df[df['ftd_mom_change'] < 0].nsmallest(10, 'ftd_mom_change')[
                ['partner_name', 'list', 'ftd_current', 'ftd_mom_change', 'ftd_mom_pct', 
                 'traffic_source', 'am', 'decline_reason']
            ].reset_index(drop=True)
            top_decliners['ftd_mom_pct'] = top_decliners['ftd_mom_pct'].round(1).astype(str) + '%'
            top_decliners.columns = ['Partner', 'List', 'Current FTD', 'MoM Change', 'MoM %', 'Source', 'AM', 'Reason']
            st.dataframe(top_decliners, use_container_width=True, hide_index=True)
        
        st.divider()
        
        # Performance status breakdown
        st.subheader("Partners by Status")
        
        for status in ['Growing', 'Stable', 'Watch', 'Red', 'New/Insufficient Baseline']:
            status_df = df[df['performance_status'] == status]
            
            if len(status_df) > 0:
                with st.expander(f"{status} ({len(status_df)} partners)"):
                    display_df = status_df[['partner_name', 'list', 'geo', 'ftd_current', 
                                           'ftd_vs_baseline_pct', 'am', 'traffic_source']].reset_index(drop=True)
                    display_df['ftd_vs_baseline_pct'] = display_df['ftd_vs_baseline_pct'].round(1).astype(str) + '%'
                    display_df.columns = ['Partner', 'List', 'GEO', 'Current FTD', 'vs Baseline %', 'AM', 'Source']
                    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # ========================================================================
    # PAGE 4: LIST MOVEMENT
    # ========================================================================
    elif selected_page == "🔄 List Movement":
        st.header("Partner List Movement Analysis")
        
        if 'df_previous' in st.session_state:
            df_previous = st.session_state['df_previous']
            movements = detect_list_movements(df, df_previous)
            
            if len(movements) > 0:
                st.subheader(f"Partners Moved This Month ({len(movements)} movements)")
                
                # Movement summary
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Movements", len(movements))
                with col2:
                    promoted = len(movements[movements['to_list'].isin(['Main Tops', 'Potential Tops'])])
                    st.metric("Promoted", promoted)
                with col3:
                    demoted = len(movements[movements['to_list'].isin(['Problematic', 'Dropped Out'])])
                    st.metric("Demoted", demoted)
                with col4:
                    st.metric("Stationary", len(df) - len(movements))
                
                st.divider()
                
                # Sankey preparation
                from_to_counts = movements.groupby(['from_list', 'to_list']).size().reset_index(name='count')
                
                st.subheader("Movement Flow (Sankey)")
                
                # Create nodes and links for Sankey
                all_lists = set(from_to_counts['from_list'].unique()) | set(from_to_counts['to_list'].unique())
                all_lists = sorted([l for l in all_lists if l != 'New Partner'])
                
                source_indices = [all_lists.index(x) for x in from_to_counts['from_list']]
                target_indices = [all_lists.index(x) for x in from_to_counts['to_list']]
                
                fig = go.Figure(data=[go.Sankey(
                    node=dict(
                        pad=15,
                        thickness=20,
                        line=dict(color='black', width=0.5),
                        label=all_lists,
                        color=['#3498db', '#9b59b6', '#e67e22', '#95a5a6']
                    ),
                    link=dict(
                        source=source_indices,
                        target=target_indices,
                        value=from_to_counts['count'].tolist()
                    )
                )])
                
                fig.update_layout(title="Partner List Movement Flow", font=dict(size=10), height=400)
                st.plotly_chart(fig, use_container_width=True)
                
                st.divider()
                
                # Movement details table
                st.subheader("Movement Details")
                movements_display = movements[['partner_name', 'from_list', 'to_list', 'geo', 'am']].copy()
                movements_display.columns = ['Partner', 'From List', 'To List', 'GEO', 'AM']
                st.dataframe(movements_display, use_container_width=True, hide_index=True)
            else:
                st.info("No partner list movements detected this month.")
        else:
            st.warning("⚠️ Upload previous month Excel to track list movements.")
        
        st.divider()
        
        # Current list distribution
        st.subheader("Current List Distribution")
        
        list_dist = df['list'].value_counts()
        fig = px.pie(values=list_dist.values, names=list_dist.index, title="Current Partner Distribution",
                    color_discrete_sequence=['#3498db', '#9b59b6', '#e67e22', '#95a5a6'])
        st.plotly_chart(fig, use_container_width=True)
    
    # ========================================================================
    # PAGE 5: ACTION CENTER
    # ========================================================================
    elif selected_page == "⚡ Action Center":
        st.header("Action Center - Priority Task Queue")
        
        tasks = []
        
        # Critical FTDs lost
        critical_losses = df[(df['ftd_current'] == 0) & (df['list'] == 'Main Tops')]
        for _, row in critical_losses.iterrows():
            tasks.append({
                'priority': 'CRITICAL',
                'task': '🔴 Zero FTD - Main Tops Partner',
                'partner': row['partner_name'],
                'details': f"Lost all FTD (was {row['ftd_previous']})",
                'am': row['am'],
                'action': 'Contact immediately'
            })
        
        # Major declines (Red zone)
        red_partners = df[(df['performance_status'] == 'Red') & (df['list'] == 'Main Tops')]
        for _, row in red_partners.nlargest(5, 'ftd_current').iterrows():
            tasks.append({
                'priority': 'HIGH',
                'task': '🔴 Red Zone - Main Tops',
                'partner': row['partner_name'],
                'details': f"Down {row['ftd_mom_change']:+,} ({row['ftd_mom_pct']:.1f}%)",
                'am': row['am'],
                'action': 'Review & intervention needed'
            })
        
        # Missing AM Feedback on declining partners
        missing_feedback = df[(df['performance_status'].isin(['Watch', 'Red'])) & 
                             ((df['am_feedback'] == '') | (df['am_feedback'].isna()))]
        for _, row in missing_feedback.nlargest(5, 'ftd_mom_change').iterrows():
            tasks.append({
                'priority': 'HIGH',
                'task': '⚠️ Missing AM Feedback',
                'partner': row['partner_name'],
                'details': f"Status: {row['performance_status']}, FTD: {row['ftd_current']}",
                'am': row['am'],
                'action': f"Request feedback from {row['am']}"
            })
        
        # Missing traffic source
        missing_source = df[(df['traffic_source'] == '') & (df['list'].isin(['Main Tops', 'Potential Tops']))]
        for _, row in missing_source.nlargest(3, 'ftd_current').iterrows():
            tasks.append({
                'priority': 'MEDIUM',
                'task': '📋 Missing Traffic Source',
                'partner': row['partner_name'],
                'details': f"List: {row['list']}, FTD: {row['ftd_current']}",
                'am': row['am'],
                'action': 'Verify traffic source'
            })
        
        # Sort by priority
        priority_order = {'CRITICAL': 0, 'HIGH': 1, 'MEDIUM': 2}
        tasks.sort(key=lambda x: priority_order.get(x['priority'], 3))
        
        st.metric(f"Total Action Items", len(tasks))
        
        st.divider()
        
        # Display tasks
        for i, task in enumerate(tasks[:20]):  # Show top 20
            color_map = {'CRITICAL': '🔴', 'HIGH': '🟠', 'MEDIUM': '🟡'}
            
            with st.expander(f"{color_map[task['priority']]} {task['priority']} | {task['task']} | {task['partner']}"):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.write(f"**Partner:** {task['partner']}")
                with col2:
                    st.write(f"**AM:** {task['am']}")
                with col3:
                    st.write(f"**Priority:** {task['priority']}")
                
                st.write(f"**Details:** {task['details']}")
                st.write(f"**Recommended Action:** {task['action']}")
        
        if len(tasks) == 0:
            st.success("✅ No urgent actions required.")
    
    # ========================================================================
    # PAGE 6: PARTNER TRACKER
    # ========================================================================
    elif selected_page == "🔍 Partner Tracker":
        st.header("Partner Tracker - Master Database")
        
        # Filters
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            selected_geo = st.multiselect("GEO", df['geo'].unique(), default=df['geo'].unique())
        with col2:
            selected_list = st.multiselect("List", df['list'].unique(), default=df['list'].unique())
        with col3:
            selected_am = st.multiselect("AM", df['am'].unique(), default=df['am'].unique())
        with col4:
            selected_status = st.multiselect("Status", df['performance_status'].unique(), 
                                            default=df['performance_status'].unique())
        with col5:
            search_term = st.text_input("Search Partner Name")
        
        # Apply filters
        filtered_df = df[
            (df['geo'].isin(selected_geo)) &
            (df['list'].isin(selected_list)) &
            (df['am'].isin(selected_am)) &
            (df['performance_status'].isin(selected_status))
        ]
        
        if search_term:
            filtered_df = filtered_df[filtered_df['partner_name'].str.contains(search_term, case=False, na=False)]
        
        st.metric(f"Showing {len(filtered_df)} of {len(df)} partners")
        
        # Display full partner table
        display_df = filtered_df[[
            'partner_id', 'partner_name', 'geo', 'am', 'list', 'performance_status',
            'ftd_current', 'ftd_previous', 'ftd_baseline_3m', 'ftd_mom_change', 'ftd_vs_baseline_pct',
            'traffic_source', 'am_feedback', 'proposed_action'
        ]].copy().reset_index(drop=True)
        
        display_df['ftd_vs_baseline_pct'] = display_df['ftd_vs_baseline_pct'].round(1).astype(str) + '%'
        display_df.columns = [
            'ID', 'Partner', 'GEO', 'AM', 'List', 'Status', 'Current FTD', 'Prev FTD', 
            '3M Base', 'MoM Δ', 'vs Base %', 'Source', 'Feedback', 'Action'
        ]
        
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        # Export functionality
        st.divider()
        col1, col2 = st.columns(2)
        
        with col1:
            csv = display_df.to_csv(index=False)
            st.download_button("📥 Download as CSV", csv, "partner_tracker.csv", "text/csv")
        
        with col2:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                display_df.to_excel(writer, index=False)
            excel_buffer.seek(0)
            st.download_button("📥 Download as Excel", excel_buffer, "partner_tracker.xlsx", 
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # ========================================================================
    # PAGE 7: AM VISIBILITY
    # ========================================================================
    elif selected_page == "👥 AM Visibility":
        st.header("Affiliate Manager Operational Metrics")
        
        # AM aggregates
        am_stats = df.groupby('am').agg({
            'partner_id': 'count',
            'ftd_current': 'sum',
            'ftd_previous': 'sum'
        }).rename(columns={'partner_id': 'partner_count'})
        
        am_stats['ftd_mom_change'] = am_stats['ftd_current'] - am_stats['ftd_previous']
        am_stats['ftd_mom_pct'] = (am_stats['ftd_mom_change'] / am_stats['ftd_previous'] * 100).round(1)
        
        # Add contact rates
        am_contact_rates = {}
        am_feedback_rates = {}
        
        for am in df['am'].unique():
            am_df = df[df['am'] == am]
            feedback_count = len(am_df[(am_df['am_feedback'] != '') & (am_df['am_feedback'].notna())])
            feedback_rate = (feedback_count / len(am_df) * 100) if len(am_df) > 0 else 0
            am_feedback_rates[am] = feedback_rate
        
        am_stats['feedback_rate_%'] = pd.Series(am_feedback_rates)
        
        st.dataframe(am_stats.reset_index(), use_container_width=True)
        
        st.divider()
        
        # AM drill-down
        st.subheader("AM Performance Details")
        
        for am in df['am'].unique():
            am_df = df[df['am'] == am]
            
            with st.expander(f"{am} ({len(am_df)} partners, {am_stats.loc[am, 'ftd_current']:,} FTD)"):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric(f"{am} - Current FTD", f"{am_df['ftd_current'].sum():,}")
                with col2:
                    st.metric(f"{am} - MoM Change", f"{am_df['ftd_mom_change'].sum():+,}")
                with col3:
                    feedback_pct = (len(am_df[(am_df['am_feedback'] != '') & (am_df['am_feedback'].notna())]) / len(am_df) * 100)
                    st.metric(f"{am} - Feedback Rate", f"{feedback_pct:.0f}%")
                
                # Partner table for this AM
                am_partners = am_df[[
                    'partner_name', 'list', 'ftd_current', 'ftd_mom_change', 'performance_status', 'am_feedback'
                ]].sort_values('ftd_current', ascending=False)
                
                am_partners.columns = ['Partner', 'List', 'Current FTD', 'MoM Δ', 'Status', 'Feedback']
                st.dataframe(am_partners, use_container_width=True, hide_index=True)
    
    # ========================================================================
    # PAGE 8: DATA QUALITY
    # ========================================================================
    elif selected_page == "📋 Data Quality":
        st.header("Data Quality Dashboard")
        
        issues = identify_missing_data(df)
        
        # Summary
        col1, col2, col3, col4 = st.columns(4)
        
        critical_count = len(issues[issues['severity'] == 'Critical'])
        high_count = len(issues[issues['severity'] == 'High'])
        medium_count = len(issues[issues['severity'] == 'Medium'])
        
        with col1:
            st.metric("🔴 Critical Issues", critical_count)
        with col2:
            st.metric("🟠 High Issues", high_count)
        with col3:
            st.metric("🟡 Medium Issues", medium_count)
        with col4:
            data_completeness = ((len(df) * len(df.columns) - len(df[df.isna().any(axis=1)])) / 
                                (len(df) * len(df.columns)) * 100)
            st.metric("Data Completeness", f"{data_completeness:.1f}%")
        
        st.divider()
        
        if len(issues) > 0:
            st.subheader("Identified Issues")
            
            for severity in ['Critical', 'High', 'Medium']:
                severity_issues = issues[issues['severity'] == severity]
                
                if len(severity_issues) > 0:
                    with st.expander(f"{severity} Severity ({len(severity_issues)} issues)"):
                        st.dataframe(severity_issues[['type', 'partner_name', 'partner_id']].reset_index(drop=True), 
                                   use_container_width=True, hide_index=True)
        else:
            st.success("✅ No data quality issues detected!")
        
        st.divider()
        
        # Data completeness by field
        st.subheader("Field Completeness")
        
        completeness = {}
        for col in df.columns:
            if col.startswith('ftd') or col in ['traffic_source', 'am_feedback', 'proposed_action', 'decline_reason']:
                filled = df[col].notna().sum() - (df[col] == '').sum()
                completeness[col] = (filled / len(df) * 100)
        
        if completeness:
            completeness_df = pd.DataFrame(list(completeness.items()), columns=['Field', 'Completeness %'])
            completeness_df = completeness_df.sort_values('Completeness %', ascending=False)
            
            fig = px.bar(completeness_df, x='Field', y='Completeness %', 
                        title="Data Completeness by Field",
                        color='Completeness %', color_continuous_scale='RdYlGn')
            st.plotly_chart(fig, use_container_width=True)
    
    # ========================================================================
    # FOOTER
    # ========================================================================
    st.divider()
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if 'upload_date' in st.session_state:
            st.caption(f"Data loaded: {st.session_state['upload_date']}")
    
    with col2:
        st.caption(f"Total partners: {len(df)} | Total FTD: {df['ftd_current'].sum():,}")
    
    with col3:
        export_buffer = BytesIO()
        with pd.ExcelWriter(export_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Full Data', index=False)
        export_buffer.seek(0)
        st.download_button("📥 Export Full Dataset", export_buffer, "full_dataset.xlsx", 
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
