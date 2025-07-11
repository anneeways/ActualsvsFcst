import streamlit as st
import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

# Optional imports for exports
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4, letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Configure Streamlit page
st.set_page_config(
    page_title="HR ROI Tracker - Actuals vs Plan",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 0.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        text-align: center;
    }
    .status-green { color: #28a745; }
    .status-yellow { color: #ffc107; }
    .status-red { color: #dc3545; }
    .variance-positive { background-color: #d4edda; padding: 0.5rem; border-radius: 0.25rem; }
    .variance-negative { background-color: #f8d7da; padding: 0.5rem; border-radius: 0.25rem; }
    .variance-neutral { background-color: #fff3cd; padding: 0.5rem; border-radius: 0.25rem; }
</style>
""", unsafe_allow_html=True)

# Sample plan templates for demo purposes
SAMPLE_PLANS = {
    'leadership_q1_2024': {
        'name': "Leadership Development Q1 2024",
        'initiative_type': "Leadership Development",
        'start_date': "2024-01-01",
        'duration_months': 6,
        'participants': 25,
        'planned_metrics': {
            'total_investment': 165000,
            'annual_roi': 285,
            'productivity_gain': 18,
            'retention_improvement': 30,
            'team_performance_gain': 15,
            'payback_months': 14
        },
        'kpis': [
            'productivity_improvement',
            'employee_retention_rate', 
            'engagement_scores',
            'promotion_rates',
            'program_completion_rate'
        ]
    },
    'recruiting_optimization_2024': {
        'name': "Recruiting Optimization 2024",
        'initiative_type': "Recruiting Process",
        'start_date': "2024-02-01", 
        'duration_months': 12,
        'participants': 50,
        'planned_metrics': {
            'total_investment': 50000,
            'annual_roi': 320,
            'time_to_hire_reduction': 35,
            'cost_per_hire_reduction': 25,
            'hire_quality_improvement': 20
        },
        'kpis': [
            'time_to_hire',
            'cost_per_hire',
            'hire_quality_score',
            'recruiter_productivity',
            'candidate_satisfaction'
        ]
    }
}

def format_currency(amount):
    """Format amount as currency"""
    return f"${amount:,.0f}"

def get_variance_status(variance_pct):
    """Get status color and icon based on variance percentage"""
    if variance_pct >= 10:
        return "🟢", "Exceeding Plan", "status-green"
    elif variance_pct >= -5:
        return "🟡", "On Track", "status-yellow"
    else:
        return "🔴", "Below Plan", "status-red"

def calculate_variance(actual, planned):
    """Calculate variance percentage"""
    if planned == 0:
        return 0
    return ((actual - planned) / planned) * 100

def generate_sample_actuals(plan_data, progress_pct):
    """Generate realistic sample actuals data for demo purposes"""
    import random
    random.seed(42)  # For consistent demo data
    
    # Simulate some variance in performance
    variance_factor = random.uniform(0.85, 1.15)  # +/- 15% variance
    completion_factor = progress_pct / 100
    
    planned_metrics = plan_data['planned_metrics']
    
    # Generate actuals with some realistic variance
    actuals = {
        'measurement_date': datetime.now().strftime('%Y-%m-%d'),
        'program_progress': progress_pct,
        'participants_completed': int(plan_data['participants'] * completion_factor * variance_factor),
        'actual_investment': planned_metrics['total_investment'] * completion_factor * random.uniform(0.9, 1.1),
        'productivity_improvement': planned_metrics.get('productivity_gain', 15) * variance_factor,
        'retention_improvement': planned_metrics.get('retention_improvement', 25) * variance_factor,
        'time_to_hire_current': 45 * (1 - (planned_metrics.get('time_to_hire_reduction', 30) / 100) * completion_factor * variance_factor),
        'cost_per_hire_current': 5000 * (1 - (planned_metrics.get('cost_per_hire_reduction', 20) / 100) * completion_factor * variance_factor),
        'engagement_score_change': random.uniform(0.5, 2.5) * completion_factor,
        'program_satisfaction': random.uniform(7.5, 9.5),
        'completion_rate': min(100, progress_pct * random.uniform(0.9, 1.1))
    }
    
    return actuals

def create_executive_dashboard(plan_data, actuals_data):
    """Create executive dashboard with key metrics"""
    
    # Calculate key variances
    planned_roi = plan_data['planned_metrics']['annual_roi']
    actual_investment = actuals_data['actual_investment']
    planned_investment = plan_data['planned_metrics']['total_investment']
    
    # Estimate actual ROI based on progress
    progress_factor = actuals_data['program_progress'] / 100
    estimated_annual_roi = planned_roi * (actuals_data.get('productivity_improvement', 15) / plan_data['planned_metrics'].get('productivity_gain', 15))
    
    roi_variance = calculate_variance(estimated_annual_roi, planned_roi)
    investment_variance = calculate_variance(actual_investment, planned_investment * progress_factor)
    
    return {
        'roi': {
            'planned': planned_roi,
            'actual': estimated_annual_roi,
            'variance_pct': roi_variance,
            'status': get_variance_status(roi_variance)
        },
        'investment': {
            'planned': planned_investment * progress_factor,
            'actual': actual_investment,
            'variance_pct': investment_variance,
            'status': get_variance_status(-investment_variance)  # Negative because lower cost is better
        },
        'progress': {
            'planned': 100,
            'actual': actuals_data['program_progress'],
            'status': get_variance_status(actuals_data['program_progress'] - 100)
        }
    }

def create_variance_waterfall_chart(plan_data, actuals_data):
    """Create waterfall chart showing variance drivers"""
    
    planned_roi = plan_data['planned_metrics']['annual_roi']
    
    # Calculate impact of each variance driver
    productivity_impact = (actuals_data.get('productivity_improvement', 15) - plan_data['planned_metrics'].get('productivity_gain', 15)) * 5
    retention_impact = (actuals_data.get('retention_improvement', 25) - plan_data['planned_metrics'].get('retention_improvement', 25)) * 3
    cost_impact = -(actuals_data['actual_investment'] - plan_data['planned_metrics']['total_investment']) / 1000
    
    actual_roi = planned_roi + productivity_impact + retention_impact + cost_impact
    
    fig = go.Figure(go.Waterfall(
        name="ROI Variance Analysis",
        orientation="v",
        measure=["absolute", "relative", "relative", "relative", "total"],
        x=["Planned ROI", "Productivity Impact", "Retention Impact", "Cost Impact", "Actual ROI"],
        y=[planned_roi, productivity_impact, retention_impact, cost_impact, actual_roi],
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        decreasing={"marker": {"color": "#ff6b6b"}},
        increasing={"marker": {"color": "#51cf66"}},
        totals={"marker": {"color": "#339af0"}}
    ))
    
    fig.update_layout(
        title="ROI Variance Waterfall Analysis",
        xaxis_title="Factors",
        yaxis_title="ROI Impact (%)",
        showlegend=False
    )
    
    return fig

def create_trend_analysis(plan_data, periods=6):
    """Create trend analysis showing plan vs actual over time"""
    
    # Generate sample trend data
    dates = pd.date_range(start=plan_data['start_date'], periods=periods, freq='M')
    
    planned_progress = np.linspace(0, 100, periods)
    actual_progress = planned_progress * np.random.uniform(0.9, 1.1, periods)
    
    planned_roi = np.linspace(0, plan_data['planned_metrics']['annual_roi'], periods)
    actual_roi = planned_roi * np.random.uniform(0.85, 1.15, periods)
    
    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=("Program Progress (%)", "Cumulative ROI (%)"),
        vertical_spacing=0.12
    )
    
    # Progress tracking
    fig.add_trace(
        go.Scatter(x=dates, y=planned_progress, name="Planned Progress", line=dict(color="#339af0", dash="dash")),
        row=1, col=1
    )
    fig.add_trace(
        go.Scatter(x=dates, y=actual_progress, name="Actual Progress", line=dict(color="#51cf66")),
        row=1, col=1
    )
    
    # ROI tracking
    fig.add_trace(
        go.Scatter(x=dates, y=planned_roi, name="Planned ROI", line=dict(color="#339af0", dash="dash")),
        row=2, col=1
    )
    fig.add_trace(
        go.Scatter(x=dates, y=actual_roi, name="Actual ROI", line=dict(color="#51cf66")),
        row=2, col=1
    )
    
    fig.update_layout(
        title="Plan vs Actual Trend Analysis",
        height=500,
        showlegend=True
    )
    
    return fig

def create_kpi_scorecard(plan_data, actuals_data):
    """Create KPI scorecard with traffic light indicators"""
    
    kpis = plan_data.get('kpis', [])
    scorecard_data = []
    
    # Sample KPI data
    kpi_values = {
        'productivity_improvement': {
            'planned': plan_data['planned_metrics'].get('productivity_gain', 15),
            'actual': actuals_data.get('productivity_improvement', 15),
            'unit': '%'
        },
        'employee_retention_rate': {
            'planned': plan_data['planned_metrics'].get('retention_improvement', 25),
            'actual': actuals_data.get('retention_improvement', 25),
            'unit': '%'
        },
        'engagement_scores': {
            'planned': 1.5,
            'actual': actuals_data.get('engagement_score_change', 1.2),
            'unit': 'pts'
        },
        'time_to_hire': {
            'planned': 30,
            'actual': actuals_data.get('time_to_hire_current', 32),
            'unit': 'days',
            'lower_is_better': True
        },
        'cost_per_hire': {
            'planned': 3750,
            'actual': actuals_data.get('cost_per_hire_current', 4000),
            'unit': '$',
            'lower_is_better': True
        },
        'program_completion_rate': {
            'planned': 95,
            'actual': actuals_data.get('completion_rate', 92),
            'unit': '%'
        }
    }
    
    for kpi in kpis:
        if kpi in kpi_values:
            kpi_data = kpi_values[kpi]
            planned = kpi_data['planned']
            actual = kpi_data['actual']
            
            if kpi_data.get('lower_is_better', False):
                variance_pct = calculate_variance(planned, actual)  # Flip for lower is better
            else:
                variance_pct = calculate_variance(actual, planned)
            
            icon, status, css_class = get_variance_status(variance_pct)
            
            scorecard_data.append({
                'KPI': kpi.replace('_', ' ').title(),
                'Planned': f"{planned}{kpi_data['unit']}",
                'Actual': f"{actual:.1f}{kpi_data['unit']}",
                'Variance': f"{variance_pct:+.1f}%",
                'Status': f"{icon} {status}"
            })
    
    return pd.DataFrame(scorecard_data)

def main():
    # Header
    st.markdown("""
    <div style='background: linear-gradient(135deg, #28a745 0%, #20c997 100%); 
                padding: 2rem; border-radius: 10px; margin-bottom: 2rem;'>
        <h1 style='color: white; margin: 0; font-size: 2.5rem;'>📈 HR ROI Tracker</h1>
        <p style='color: rgba(255,255,255,0.8); margin: 0.5rem 0 0 0; font-size: 1.2rem;'>
            Track Actual Performance vs Original Plan
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show export capabilities
    col1, col2, col3 = st.columns(3)
    with col1:
        if REPORTLAB_AVAILABLE:
            st.success("✅ PDF Export Available")
        else:
            st.warning("⚠️ PDF Export: Install reportlab")
    with col2:
        if PPTX_AVAILABLE:
            st.success("✅ PowerPoint Export Available")
        else:
            st.warning("⚠️ PowerPoint: Install python-pptx")
    with col3:
        st.success("✅ Data Export Available")
    
    # Initialize session state
    if 'selected_plan' not in st.session_state:
        st.session_state.selected_plan = None
    if 'actuals_data' not in st.session_state:
        st.session_state.actuals_data = {}
    
    # Sidebar for plan selection and setup
    with st.sidebar:
        st.header("🎯 Plan Selection")
        
        # Plan selection
        st.subheader("📋 Available Plans")
        plan_options = [''] + list(SAMPLE_PLANS.keys()) + ['Upload New Plan']
        selected_plan_key = st.selectbox("Select HR Initiative Plan", plan_options)
        
        if selected_plan_key and selected_plan_key != 'Upload New Plan':
            st.session_state.selected_plan = SAMPLE_PLANS[selected_plan_key]
            st.success(f"✅ Loaded: {st.session_state.selected_plan['name']}")
            
            # Show plan summary
            with st.expander("📊 Plan Summary"):
                plan = st.session_state.selected_plan
                st.write(f"**Initiative:** {plan['initiative_type']}")
                st.write(f"**Duration:** {plan['duration_months']} months")
                st.write(f"**Participants:** {plan['participants']}")
                st.write(f"**Planned ROI:** {plan['planned_metrics']['annual_roi']}%")
                st.write(f"**Investment:** {format_currency(plan['planned_metrics']['total_investment'])}")
        
        elif selected_plan_key == 'Upload New Plan':
            st.subheader("📤 Upload Plan")
            uploaded_file = st.file_uploader("Upload ROI Plan (JSON)", type=['json'])
            if uploaded_file:
                try:
                    plan_data = json.load(uploaded_file)
                    st.session_state.selected_plan = plan_data
                    st.success("✅ Plan uploaded successfully!")
                except Exception as e:
                    st.error(f"Error loading plan: {str(e)}")
        
        st.divider()
        
        # Progress tracking
        if st.session_state.selected_plan:
            st.subheader("📅 Progress Tracking")
            progress_pct = st.slider(
                "Program Progress (%)", 
                0, 100, 
                value=st.session_state.actuals_data.get('program_progress', 60),
                help="How far along is the program?"
            )
            
            if st.button("🔄 Generate Sample Actuals", help="Generate realistic sample data for demo"):
                st.session_state.actuals_data = generate_sample_actuals(st.session_state.selected_plan, progress_pct)
                st.success("✅ Sample actuals generated!")
    
    # Main content
    if not st.session_state.selected_plan:
        st.info("👈 Please select an HR initiative plan from the sidebar to begin tracking actual performance.")
        
        # Show overview of tracking capabilities
        st.subheader("🔍 ROI Tracking Capabilities")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            #### 📊 Performance Monitoring
            - Real-time variance analysis
            - KPI scorecards with traffic lights
            - Trend analysis and forecasting
            - Progress tracking dashboards
            """)
        
        with col2:
            st.markdown("""
            #### 🎯 Action Management  
            - Root cause analysis
            - Corrective action tracking
            - Risk identification
            - Performance optimization
            """)
        
        with col3:
            st.markdown("""
            #### 📈 Executive Reporting
            - Variance waterfall charts
            - Portfolio performance views
            - Stakeholder communications
            - ROI reforecast updates
            """)
        
        return
    
    plan = st.session_state.selected_plan
    
    # Check if we have actuals data
    if not st.session_state.actuals_data:
        st.warning("⚡ Click 'Generate Sample Actuals' in the sidebar to see tracking in action, or input your actual data below.")
        
        # Manual actuals input
        with st.expander("📝 Input Actual Performance Data", expanded=True):
            col1, col2 = st.columns(2)
            
            with col1:
                actual_progress = st.number_input("Program Progress (%)", 0, 100, 50)
                actual_investment = st.number_input("Actual Investment to Date ($)", 0, plan['planned_metrics']['total_investment'], plan['planned_metrics']['total_investment']//2)
                actual_productivity = st.number_input("Actual Productivity Gain (%)", 0, 50, plan['planned_metrics'].get('productivity_gain', 15))
            
            with col2:
                actual_retention = st.number_input("Actual Retention Improvement (%)", 0, 50, plan['planned_metrics'].get('retention_improvement', 25))
                engagement_change = st.number_input("Engagement Score Change", 0.0, 5.0, 1.5)
                completion_rate = st.number_input("Program Completion Rate (%)", 0, 100, 95)
            
            if st.button("💾 Save Actuals Data"):
                st.session_state.actuals_data = {
                    'measurement_date': datetime.now().strftime('%Y-%m-%d'),
                    'program_progress': actual_progress,
                    'actual_investment': actual_investment,
                    'productivity_improvement': actual_productivity,
                    'retention_improvement': actual_retention,
                    'engagement_score_change': engagement_change,
                    'completion_rate': completion_rate
                }
                st.success("✅ Actuals data saved!")
                st.rerun()
        
        return
    
    actuals = st.session_state.actuals_data
    
    # Executive Dashboard
    st.subheader("🎯 Executive Dashboard")
    
    dashboard_data = create_executive_dashboard(plan, actuals)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        roi_data = dashboard_data['roi']
        icon, status, css_class = roi_data['status']
        st.metric(
            "ROI Performance",
            f"{roi_data['actual']:.0f}%",
            delta=f"{roi_data['variance_pct']:+.1f}% vs Plan",
            help=f"Planned: {roi_data['planned']:.0f}%"
        )
        st.markdown(f"<div class='{css_class}'>{icon} {status}</div>", unsafe_allow_html=True)
    
    with col2:
        inv_data = dashboard_data['investment']
        icon, status, css_class = inv_data['status']
        st.metric(
            "Investment Track",
            format_currency(inv_data['actual']),
            delta=f"{inv_data['variance_pct']:+.1f}% vs Plan",
            help=f"Planned: {format_currency(inv_data['planned'])}"
        )
        st.markdown(f"<div class='{css_class}'>{icon} {status}</div>", unsafe_allow_html=True)
    
    with col3:
        progress_data = dashboard_data['progress']
        st.metric(
            "Program Progress",
            f"{progress_data['actual']:.0f}%",
            delta="On Schedule" if progress_data['actual'] >= 90 else "In Progress"
        )
        
        # Progress bar
        progress_color = "#28a745" if progress_data['actual'] >= 90 else "#ffc107" if progress_data['actual'] >= 70 else "#dc3545"
        st.markdown(f"""
        <div style="background-color: #e9ecef; border-radius: 0.25rem; padding: 0.25rem;">
            <div style="background-color: {progress_color}; width: {progress_data['actual']}%; height: 20px; border-radius: 0.25rem; text-align: center; color: white; font-weight: bold;">
                {progress_data['actual']:.0f}%
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.metric(
            "Last Updated",
            actuals['measurement_date'],
            delta="Data Fresh" if (datetime.now() - datetime.strptime(actuals['measurement_date'], '%Y-%m-%d')).days <= 7 else "Update Needed"
        )
        
        if st.button("🔄 Refresh Data"):
            st.rerun()
    
    st.divider()
    
    # Main analysis tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Variance Analysis", "📈 Trend Analysis", "🎯 KPI Scorecard", 
        "🚨 Action Management", "📋 Reports"
    ])
    
    with tab1:
        st.subheader("📊 Variance Analysis")
        
        # Variance waterfall chart
        waterfall_fig = create_variance_waterfall_chart(plan, actuals)
        st.plotly_chart(waterfall_fig, use_container_width=True)
        
        # Variance summary table
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Key Variances")
            
            variance_data = []
            
            # ROI Variance
            planned_roi = plan['planned_metrics']['annual_roi']
            actual_roi = planned_roi * (actuals.get('productivity_improvement', 15) / plan['planned_metrics'].get('productivity_gain', 15))
            roi_variance = calculate_variance(actual_roi, planned_roi)
            
            variance_data.append({
                'Metric': 'ROI (%)',
                'Planned': f"{planned_roi:.0f}%",
                'Actual': f"{actual_roi:.0f}%",
                'Variance': f"{roi_variance:+.1f}%",
                'Status': get_variance_status(roi_variance)[1]
            })
            
            # Investment Variance  
            planned_inv = plan['planned_metrics']['total_investment']
            actual_inv = actuals['actual_investment']
            inv_variance = calculate_variance(actual_inv, planned_inv)
            
            variance_data.append({
                'Metric': 'Investment',
                'Planned': format_currency(planned_inv),
                'Actual': format_currency(actual_inv),
                'Variance': f"{inv_variance:+.1f}%",
                'Status': get_variance_status(-inv_variance)[1]  # Lower is better
            })
            
            # Productivity Variance
            planned_prod = plan['planned_metrics'].get('productivity_gain', 15)
            actual_prod = actuals.get('productivity_improvement', 15)
            prod_variance = calculate_variance(actual_prod, planned_prod)
            
            variance_data.append({
                'Metric': 'Productivity Gain',
                'Planned': f"{planned_prod}%",
                'Actual': f"{actual_prod:.1f}%", 
                'Variance': f"{prod_variance:+.1f}%",
                'Status': get_variance_status(prod_variance)[1]
            })
            
            df_variance = pd.DataFrame(variance_data)
            st.dataframe(df_variance, use_container_width=True, hide_index=True)
        
        with col2:
            st.subheader("🎯 Impact Analysis")
            
            # Calculate business impact
            participant_impact = actuals.get('participants_completed', plan['participants']) - plan['participants']
            investment_impact = actual_inv - planned_inv
            roi_impact = actual_roi - planned_roi
            
            impact_text = f"""
            **Program Performance Impact:**
            
            📊 **ROI Impact:** {roi_impact:+.1f} percentage points
            - Current trajectory: {actual_roi:.0f}% vs {planned_roi:.0f}% planned
            
            💰 **Investment Impact:** {format_currency(investment_impact)}
            - {"Over budget" if investment_impact > 0 else "Under budget" if investment_impact < 0 else "On budget"}
            
            👥 **Participation Impact:** {participant_impact:+.0f} participants
            - Completion rate: {actuals.get('completion_rate', 95):.0f}%
            
            **📈 Projected Annual Value:**
            - Based on current performance: {format_currency(actual_roi * planned_inv / 100)}
            - Original projection: {format_currency(planned_roi * planned_inv / 100)}
            - Net impact: {format_currency((actual_roi - planned_roi) * planned_inv / 100)}
            """
            
            st.markdown(impact_text)
    
    with tab2:
        st.subheader("📈 Trend Analysis")
        
        # Trend chart
        trend_fig = create_trend_analysis(plan)
        st.plotly_chart(trend_fig, use_container_width=True)
        
        # Forecast section
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🔮 Performance Forecast")
            
            current_trend = actuals.get('productivity_improvement', 15) / plan['planned_metrics'].get('productivity_gain', 15)
            forecasted_roi = plan['planned_metrics']['annual_roi'] * current_trend
            
            forecast_text = f"""
            **Based on Current Performance Trend:**
            
            📊 **Forecasted Final ROI:** {forecasted_roi:.0f}%
            - vs Original Plan: {plan['planned_metrics']['annual_roi']:.0f}%
            - Confidence Level: {"High" if abs(current_trend - 1) < 0.1 else "Medium" if abs(current_trend - 1) < 0.2 else "Low"}
            
            ⏱️ **Timeline Assessment:**
            - Current Progress: {actuals['program_progress']:.0f}%
            - Expected Completion: {"On Time" if actuals['program_progress'] >= 80 else "Delayed" if actuals['program_progress'] < 60 else "Monitor Closely"}
            
            🎯 **Key Risk Factors:**
            - {"✅ Performance on track" if current_trend >= 0.95 else "⚠️ Performance below target"}
            - {"✅ Budget on track" if abs(calculate_variance(actual_inv, planned_inv)) < 10 else "⚠️ Budget variance detected"}
            """
            
            st.markdown(forecast_text)
        
        with col2:
            st.subheader("📊 Monthly Progress")
            
            # Monthly progress tracking
            months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
            planned_monthly = [15, 30, 45, 65, 80, 100]
            actual_monthly = [12, 28, 42, 58, 72, actuals['program_progress']]
            
            fig_monthly = go.Figure()
            fig_monthly.add_trace(go.Scatter(
                x=months, y=planned_monthly, 
                mode='lines+markers', name='Planned',
                line=dict(color='#339af0', dash='dash')
            ))
            fig_monthly.add_trace(go.Scatter(
                x=months, y=actual_monthly,
                mode='lines+markers', name='Actual',
                line=dict(color='#51cf66')
            ))
            
            fig_monthly.update_layout(
                title="Monthly Progress Tracking",
                xaxis_title="Month",
                yaxis_title="Progress (%)",
                yaxis=dict(range=[0, 100])
            )
            
            st.plotly_chart(fig_monthly, use_container_width=True)
    
    with tab3:
        st.subheader("🎯 KPI Scorecard")
        
        # Create KPI scorecard
        kpi_df = create_kpi_scorecard(plan, actuals)
        
        if not kpi_df.empty:
            # Display KPI table with styling
            st.dataframe(
                kpi_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Status": st.column_config.TextColumn("Status", width="medium"),
                    "Variance": st.column_config.TextColumn("Variance", width="small")
                }
            )
            
            # KPI Performance Chart
            col1, col2 = st.columns(2)
            
            with col1:
                # Extract numeric values for charting
                if len(kpi_df) > 0:
                    kpi_names = kpi_df['KPI'].tolist()
                    variances = []
                    
                    for _, row in kpi_df.iterrows():
                        variance_str = row['Variance'].replace('%', '').replace('+', '')
                        try:
                            variances.append(float(variance_str))
                        except:
                            variances.append(0)
                    
                    colors = ['#28a745' if v >= 0 else '#dc3545' for v in variances]
                    
                    fig_kpi = go.Figure(data=[
                        go.Bar(x=kpi_names, y=variances, marker_color=colors)
                    ])
                    
                    fig_kpi.update_layout(
                        title="KPI Variance from Plan (%)",
                        xaxis_title="KPIs",
                        yaxis_title="Variance (%)",
                        xaxis_tickangle=-45
                    )
                    
                    st.plotly_chart(fig_kpi, use_container_width=True)
            
            with col2:
                st.subheader("🏆 Performance Summary")
                
                # Calculate overall performance score
                on_target_kpis = len([v for v in variances if v >= -5])
                exceeding_kpis = len([v for v in variances if v >= 10])
                total_kpis = len(variances)
                
                performance_score = (on_target_kpis / total_kpis * 100) if total_kpis > 0 else 0
                
                st.metric("Overall KPI Score", f"{performance_score:.0f}%")
                st.metric("KPIs On Target", f"{on_target_kpis}/{total_kpis}")
                st.metric("KPIs Exceeding Plan", f"{exceeding_kpis}/{total_kpis}")
                
                if performance_score >= 80:
                    st.success("🟢 Strong Performance")
                elif performance_score >= 60:
                    st.warning("🟡 Monitor Performance")
                else:
                    st.error("🔴 Action Required")
        else:
            st.info("No KPI data available for this initiative.")
    
    with tab4:
        st.subheader("🚨 Action Management")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🔍 Risk Assessment")
            
            # Identify risks based on performance
            risks = []
            
            # ROI risk
            roi_variance = calculate_variance(actual_roi, planned_roi)
            if roi_variance < -10:
                risks.append({
                    'risk': 'ROI Below Target',
                    'severity': 'High',
                    'impact': f"{roi_variance:.1f}% below plan",
                    'action': 'Review program delivery and participant engagement'
                })
            
            # Budget risk
            inv_variance = calculate_variance(actuals['actual_investment'], plan['planned_metrics']['total_investment'])
            if inv_variance > 15:
                risks.append({
                    'risk': 'Budget Overrun',
                    'severity': 'Medium',
                    'impact': f"{inv_variance:.1f}% over budget",
                    'action': 'Implement cost controls and review vendor contracts'
                })
            
            # Completion risk
            if actuals.get('completion_rate', 95) < 85:
                risks.append({
                    'risk': 'Low Completion Rate',
                    'severity': 'High',
                    'impact': f"Only {actuals.get('completion_rate', 95):.0f}% completion",
                    'action': 'Enhance participant support and motivation'
                })
            
            if risks:
                for i, risk in enumerate(risks):
                    severity_color = '#dc3545' if risk['severity'] == 'High' else '#ffc107'
                    st.markdown(f"""
                    <div style="border-left: 4px solid {severity_color}; padding: 1rem; margin: 1rem 0; background-color: #f8f9fa;">
                        <strong>🚨 {risk['risk']}</strong><br>
                        <em>Severity: {risk['severity']}</em><br>
                        Impact: {risk['impact']}<br>
                        <strong>Recommended Action:</strong> {risk['action']}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("🟢 No major risks identified")
        
        with col2:
            st.subheader("📋 Action Items")
            
            # Action tracking
            if 'action_items' not in st.session_state:
                st.session_state.action_items = []
            
            # Add new action
            with st.expander("➕ Add Action Item"):
                action_desc = st.text_input("Action Description")
                action_owner = st.text_input("Owner")
                action_due = st.date_input("Due Date")
                action_priority = st.selectbox("Priority", ["Low", "Medium", "High"])
                
                if st.button("Add Action"):
                    st.session_state.action_items.append({
                        'description': action_desc,
                        'owner': action_owner,
                        'due_date': action_due.strftime('%Y-%m-%d'),
                        'priority': action_priority,
                        'status': 'Open',
                        'created_date': datetime.now().strftime('%Y-%m-%d')
                    })
                    st.success("Action item added!")
            
            # Display existing actions
            if st.session_state.action_items:
                for i, action in enumerate(st.session_state.action_items):
                    priority_color = '#dc3545' if action['priority'] == 'High' else '#ffc107' if action['priority'] == 'Medium' else '#28a745'
                    
                    st.markdown(f"""
                    <div style="border: 1px solid #dee2e6; border-radius: 0.25rem; padding: 1rem; margin: 0.5rem 0;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <strong>{action['description']}</strong>
                            <span style="background-color: {priority_color}; color: white; padding: 0.2rem 0.5rem; border-radius: 0.25rem; font-size: 0.8rem;">
                                {action['priority']}
                            </span>
                        </div>
                        <small>Owner: {action['owner']} | Due: {action['due_date']} | Status: {action['status']}</small>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No action items yet")
    
    with tab5:
        st.subheader("📋 Executive Reports")
        
        # Report generation
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📊 Report Options")
            
            report_type = st.selectbox(
                "Select Report Type",
                ["Executive Summary", "Detailed Variance Report", "KPI Performance Report", "Action Plan Report"]
            )
            
            report_period = st.selectbox(
                "Reporting Period",
                ["Current Month", "Quarter to Date", "Year to Date", "Full Program"]
            )
            
            include_charts = st.checkbox("Include Charts", value=True)
            include_actions = st.checkbox("Include Action Items", value=True)
        
        with col2:
            st.subheader("📈 Key Insights")
            
            insights = f"""
            **Performance Highlights:**
            
            🎯 **Overall Status:** {"On Track" if roi_variance >= -5 else "Needs Attention"}
            
            📊 **ROI Performance:** {actual_roi:.0f}% ({roi_variance:+.1f}% vs plan)
            
            💰 **Investment Efficiency:** {format_currency(actual_inv)} spent
            
            👥 **Participant Engagement:** {actuals.get('completion_rate', 95):.0f}% completion rate
            
            🔮 **Forecast:** {"Likely to meet targets" if roi_variance >= -10 else "May miss targets without intervention"}
            
            **Next Steps:**
            - {"Continue current approach" if roi_variance >= 0 else "Implement performance improvements"}
            - Monitor key risk factors
            - {"Prepare scale-up plan" if roi_variance >= 10 else "Focus on optimization"}
            """
            
            st.markdown(insights)
        
        # Export buttons
        st.divider()
        st.subheader("📤 Export Reports")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if st.button("📋 Text Report", type="primary"):
                report_text = create_variance_report(plan, actuals, dashboard_data)
                st.download_button(
                    label="📥 Download Text",
                    data=report_text,
                    file_name=f"roi_variance_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    mime="text/plain"
                )
        
        with col2:
            if REPORTLAB_AVAILABLE:
                if st.button("📄 PDF Report", type="primary"):
                    st.info("PDF generation feature coming soon!")
            else:
                st.button("📄 PDF Report", disabled=True, help="Install reportlab")
        
        with col3:
            if PPTX_AVAILABLE:
                if st.button("📊 PowerPoint", type="primary"):
                    st.info("PowerPoint generation feature coming soon!")
            else:
                st.button("📊 PowerPoint", disabled=True, help="Install python-pptx")
        
        with col4:
            if st.button("📊 Data Export", type="secondary"):
                export_data = {
                    'plan': plan,
                    'actuals': actuals,
                    'dashboard': dashboard_data,
                    'kpis': kpi_df.to_dict('records') if not kpi_df.empty else [],
                    'action_items': st.session_state.action_items,
                    'export_date': datetime.now().isoformat()
                }
                
                st.download_button(
                    label="📥 Download JSON",
                    data=json.dumps(export_data, indent=2, default=str),
                    file_name=f"roi_tracking_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json"
                )

def create_variance_report(plan, actuals, dashboard_data):
    """Create a text-based variance report"""
    
    report = f"""
HR ROI VARIANCE ANALYSIS REPORT
Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}

INITIATIVE: {plan['name']}
REPORTING PERIOD: {actuals['measurement_date']}

EXECUTIVE SUMMARY
================
Program Progress: {actuals['program_progress']:.0f}%
Overall Status: {dashboard_data['roi']['status'][1]}

KEY PERFORMANCE INDICATORS
==========================
ROI Performance:
  Planned: {dashboard_data['roi']['planned']:.0f}%
  Actual: {dashboard_data['roi']['actual']:.0f}%
  Variance: {dashboard_data['roi']['variance_pct']:+.1f}%
  Status: {dashboard_data['roi']['status'][1]}

Investment Performance:
  Planned: {format_currency(dashboard_data['investment']['planned'])}
  Actual: {format_currency(dashboard_data['investment']['actual'])}
  Variance: {dashboard_data['investment']['variance_pct']:+.1f}%
  Status: {dashboard_data['investment']['status'][1]}

DETAILED ANALYSIS
================
Productivity Improvement: {actuals.get('productivity_improvement', 15):.1f}%
Retention Improvement: {actuals.get('retention_improvement', 25):.1f}%
Program Completion Rate: {actuals.get('completion_rate', 95):.0f}%
Participant Satisfaction: {actuals.get('program_satisfaction', 8.5):.1f}/10

RECOMMENDATIONS
==============
{"✅ Continue current approach - program performing well" if dashboard_data['roi']['variance_pct'] >= 0 else "⚠️ Implement performance improvement actions" if dashboard_data['roi']['variance_pct'] >= -10 else "🚨 Urgent intervention required - significant underperformance"}

Next Review Date: {(datetime.now() + timedelta(days=30)).strftime('%B %d, %Y')}

Generated by HR ROI Tracker
"""
    
    return report

if __name__ == "__main__":
    main()
