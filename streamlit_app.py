# streamlit_trading_app.py - COMPLETE VERSION WITH ALL 20 COLUMNS
import streamlit as st
import pandas as pd
import asyncio
import json
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import time
from typing import Dict, List, Optional
import requests

# Page configuration
st.set_page_config(
    page_title="Option Trading System",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Import local modules
try:
    from main import OptimizedOptionTradingSystem as OptionTradingSystem
    SYSTEM_AVAILABLE = True
except ImportError:
    st.error("⚠️ Cannot import main.py - please ensure it exists in the same directory!")
    SYSTEM_AVAILABLE = False

try:
    from kite_autologin import AutomatedDailyLogin
    AUTOLOGIN_AVAILABLE = True
except ImportError:
    AUTOLOGIN_AVAILABLE = False

# Custom CSS
st.markdown("""
<style>
    .stMetric {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
    }
    .success-box {
        padding: 10px;
        border-radius: 5px;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .warning-box {
        padding: 10px;
        border-radius: 5px;
        background-color: #fff3cd;
        border: 1px solid #ffeeba;
        color: #856404;
    }
    .error-box {
        padding: 10px;
        border-radius: 5px;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
    }
    .info-box {
        padding: 10px;
        border-radius: 5px;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
def init_session_state():
    """Initialize all session state variables"""
    defaults = {
        'system_initialized': False,
        'system_obj': None,
        'dropdown_options': None,
        'option_data': None,
        'init_time': None,
        'fetch_time': None,
        'login_status': None,
        'api_mode': False,
        'api_url': "http://localhost:8000",
        'selected_symbol': None,
        'selected_option_expiry': None,
        'selected_future_expiry': None,
        'last_fetch_params': None,
        'last_fetch_key': None
    }
    
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value

init_session_state()

# Utility functions
def create_option_chain_chart(df: pd.DataFrame, underlying_price: float) -> go.Figure:
    """Create interactive option chain visualization"""
    fig = go.Figure()
    
    # Call LTP
    fig.add_trace(go.Scatter(
        x=df['strike'],
        y=df['call_ltp'],
        name='Call LTP',
        mode='lines+markers',
        line=dict(color='#00cc66', width=2),
        marker=dict(size=8)
    ))
    
    # Put LTP
    fig.add_trace(go.Scatter(
        x=df['strike'],
        y=df['put_ltp'],
        name='Put LTP',
        mode='lines+markers',
        line=dict(color='#ff6666', width=2),
        marker=dict(size=8)
    ))
    
    # Add vertical line at underlying price
    fig.add_vline(
        x=underlying_price,
        line_dash="dash",
        line_color="blue",
        line_width=2,
        annotation_text=f"Spot: ₹{underlying_price:,.0f}",
        annotation_position="top"
    )
    
    fig.update_layout(
        title="Option Chain - Premium Prices",
        xaxis_title="Strike Price",
        yaxis_title="Premium (₹)",
        hovermode='x unified',
        height=450,
        template='plotly_white',
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    return fig

def create_oi_chart(df: pd.DataFrame) -> go.Figure:
    """Create Open Interest comparison chart"""
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=df['strike'],
        y=df['call_oi'],
        name='Call OI',
        marker_color='#66ccff',
        text=df['call_oi'],
        textposition='outside',
        texttemplate='%{text:.2s}'
    ))
    
    fig.add_trace(go.Bar(
        x=df['strike'],
        y=df['put_oi'],
        name='Put OI',
        marker_color='#ff9999',
        text=df['put_oi'],
        textposition='outside',
        texttemplate='%{text:.2s}'
    ))
    
    fig.update_layout(
        title="Open Interest Comparison",
        xaxis_title="Strike Price",
        yaxis_title="Open Interest",
        barmode='group',
        height=450,
        template='plotly_white',
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    return fig

def format_number(num):
    """Format number with Indian numbering system"""
    if num >= 10000000:  # 1 crore
        return f"₹{num/10000000:.2f}Cr"
    elif num >= 100000:  # 1 lakh
        return f"₹{num/100000:.2f}L"
    elif num >= 1000:
        return f"₹{num/1000:.2f}K"
    else:
        return f"₹{num:,.0f}"

def check_api_health(api_url: str) -> bool:
    """Check if FastAPI server is running"""
    try:
        response = requests.get(f"{api_url}/health", timeout=2)
        return response.status_code == 200
    except:
        return False

# Header
st.title("📊 Option Trading System")
st.markdown("**Real-time Option Chain Analysis & Trading Dashboard**")

# Sidebar - System Control
with st.sidebar:
    st.header("⚙️ System Control")
    
    # Mode selection
    mode = st.radio(
        "Operation Mode",
        ["🖥️ Local System", "🌐 API Mode"],
        index=0 if not st.session_state.api_mode else 1,
        help="Local: Direct system integration | API: Connect to FastAPI server"
    )
    
    st.session_state.api_mode = (mode == "🌐 API Mode")
    
    if st.session_state.api_mode:
        # API Mode
        st.subheader("🌐 API Configuration")
        api_url = st.text_input("API URL", value=st.session_state.api_url)
        st.session_state.api_url = api_url
        
        api_healthy = check_api_health(api_url)
        if api_healthy:
            st.success("✅ API Connected")
        else:
            st.error("❌ API Not Reachable")
        
        if st.button("🔄 Test Connection", use_container_width=True):
            if check_api_health(api_url):
                st.success("API is healthy!")
            else:
                st.error("Cannot reach API server")
    
    else:
        # Local Mode
        st.subheader("🖥️ Local System")
        
        # System status
        if st.session_state.system_initialized:
            st.success("✅ System Ready")
            if st.session_state.init_time:
                st.info(f"⏱️ Init: {st.session_state.init_time:.2f}s")
            
            # Show extracted dates info
            if st.session_state.dropdown_options:
                with st.expander("📋 Available Options", expanded=False):
                    opts = st.session_state.dropdown_options
                    st.write(f"**Symbols:** {len(opts.get('symbols', []))}")
                    st.write(f"**Option Dates:** {len(opts.get('option_expiry', []))}")
                    st.write(f"**Future Dates:** {len(opts.get('future_expiry', []))}")
        else:
            st.warning("⏳ Not Initialized")
        
        # Initialize button
        if not st.session_state.system_initialized:
            if st.button("🚀 Initialize System", type="primary", use_container_width=True):
                with st.spinner("🔧 Initializing trading system..."):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        start_time = time.time()
                        
                        # Step 1: Initialize system
                        status_text.text("Step 1/2: Performing auto-login...")
                        progress_bar.progress(25)
                        
                        if st.session_state.system_obj is None:
                            st.session_state.system_obj = OptionTradingSystem()
                        
                        status_text.text("Step 2/2: Extracting dates from Excel...")
                        progress_bar.progress(50)
                        
                        loop = asyncio.new_event_loop()
                        asyncio.set_event_loop(loop)
                        success = loop.run_until_complete(
                            st.session_state.system_obj.initialize_system_fast()
                        )
                        loop.close()
                        
                        progress_bar.progress(100)
                        
                        if success:
                            options = st.session_state.system_obj.get_dropdown_options()
                            st.session_state.dropdown_options = options
                            st.session_state.system_initialized = True
                            st.session_state.init_time = time.time() - start_time
                            
                            status_text.empty()
                            progress_bar.empty()
                            
                            st.success(f"✅ System initialized in {st.session_state.init_time:.2f}s!")
                            st.balloons()
                            time.sleep(1)
                            st.rerun()
                        else:
                            status_text.empty()
                            progress_bar.empty()
                            st.error("❌ Initialization failed!")
                    except Exception as e:
                        status_text.empty()
                        progress_bar.empty()
                        st.error(f"Error: {str(e)}")
        else:
            # Refresh dates button
            if st.button("🔄 Refresh Dates", use_container_width=True):
                with st.spinner("Refreshing dates from Excel..."):
                    try:
                        from date_extractor import EnhancedDateExtractor
                        extractor = EnhancedDateExtractor()
                        new_dates = extractor.extract_all_dates()
                        
                        if new_dates:
                            st.session_state.dropdown_options = new_dates
                            st.success("✅ Dates refreshed!")
                            st.rerun()
                        else:
                            st.warning("⚠️ Could not refresh dates")
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
            
            # Reset button
            if st.button("🔄 Reset System", type="secondary", use_container_width=True):
                if st.session_state.system_obj:
                    try:
                        loop = asyncio.new_event_loop()
                        asyncio.set_event_loop(loop)
                        loop.run_until_complete(st.session_state.system_obj.cleanup())
                        loop.close()
                    except:
                        pass
                
                # Clear all session state
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                
                st.success("System reset!")
                time.sleep(1)
                st.rerun()
    
    st.divider()
    
    # Market status
    st.subheader("📈 Market Status")
    now = datetime.now()
    market_open = datetime.strptime("09:15", "%H:%M").time()
    market_close = datetime.strptime("15:30", "%H:%M").time()
    is_open = market_open <= now.time() <= market_close and now.weekday() < 5
    
    if is_open:
        st.success("🟢 Market OPEN")
    else:
        st.error("🔴 Market CLOSED")
    
    st.info(f"🕐 {now.strftime('%H:%M:%S')}")

# Main content area
tab1, tab2, tab3 = st.tabs(["📊 Trading", "📈 Analysis", "⚙️ Settings"])

# Tab 1: Trading
with tab1:
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.header("🎯 Trading Controls")
        
        can_trade = st.session_state.system_initialized and not st.session_state.api_mode
        
        if can_trade:
            # Get dropdown options
            options = st.session_state.dropdown_options
            
            if not options or not options.get('symbols'):
                st.error("❌ No dropdown options available. Please initialize the system.")
            else:
                # Symbol selection
                symbol = st.selectbox(
                    "Symbol",
                    options=options.get('symbols', []),
                    index=0,
                    key='symbol_select'
                )
                
                # Option Expiry selection
                option_expiry_dates = options.get('option_expiry', [])
                if option_expiry_dates:
                    option_expiry = st.selectbox(
                        "Option Expiry",
                        options=option_expiry_dates,
                        index=0,
                        key='option_expiry_select'
                    )
                else:
                    st.error("No option expiry dates available")
                    option_expiry = None
                
                # Future Expiry selection
                future_expiry_dates = options.get('future_expiry', [])
                if future_expiry_dates:
                    future_expiry = st.selectbox(
                        "Future Expiry",
                        options=future_expiry_dates,
                        index=0,
                        key='future_expiry_select'
                    )
                else:
                    st.error("No future expiry dates available")
                    future_expiry = None
                
                # Chain length
                chain_length = st.slider(
                    "Chain Length",
                    min_value=10,
                    max_value=50,
                    value=20,
                    step=5,
                    key='chain_length_select'
                )
                
                st.divider()
                
                # Fetch button
                if st.button("📡 Fetch Option Data", type="primary", use_container_width=True, key="fetch_button"):
                    if not option_expiry or not future_expiry:
                        st.error("Please select valid expiry dates")
                    else:
                        # Set a flag to prevent re-fetching
                        fetch_key = f"{symbol}_{option_expiry}_{future_expiry}_{chain_length}"
                        
                        # Only fetch if this is a new request
                        if st.session_state.get('last_fetch_key') != fetch_key:
                            with st.spinner("🔄 Fetching live data from Excel..."):
                                try:
                                    fetch_start = time.time()
                                    
                                    # Create new event loop for async operation
                                    loop = asyncio.new_event_loop()
                                    asyncio.set_event_loop(loop)
                                    
                                    # Fetch data
                                    data = loop.run_until_complete(
                                        st.session_state.system_obj.fetch_option_data_fast(
                                            symbol=symbol,
                                            option_expiry=option_expiry,
                                            future_expiry=future_expiry,
                                            chain_length=chain_length
                                        )
                                    )
                                    
                                    loop.close()
                                    
                                    fetch_time = time.time() - fetch_start
                                    
                                    # Store results
                                    st.session_state.option_data = data
                                    st.session_state.fetch_time = fetch_time
                                    st.session_state.last_fetch_key = fetch_key
                                    st.session_state.last_fetch_params = {
                                        'symbol': symbol,
                                        'option_expiry': option_expiry,
                                        'future_expiry': future_expiry,
                                        'chain_length': chain_length
                                    }
                                    
                                    # Check data source
                                    data_source = data.get('data_source', 'unknown')
                                    
                                    if data_source == 'excel_live':
                                        st.success(f"✅ Live data fetched in {fetch_time:.2f}s")
                                    elif data_source == 'excel_manual':
                                        st.success(f"✅ Data fetched (manual trigger) in {fetch_time:.2f}s")
                                        if 'warning' in data:
                                            st.warning(data['warning'])
                                    elif data_source == 'excel_partial':
                                        st.warning("⚠️ Partial data received - check Excel manually")
                                    elif data_source == 'fallback':
                                        st.error("❌ Using fallback data - Excel fetch failed")
                                        if 'fallback_reason' in data:
                                            st.error(f"Reason: {data['fallback_reason']}")
                                    else:
                                        st.warning(f"⚠️ Unknown data source: {data_source}")
                                    
                                except Exception as e:
                                    st.error(f"❌ Fetch failed: {str(e)}")
                                    import traceback
                                    with st.expander("Error Details"):
                                        st.code(traceback.format_exc())
                        else:
                            st.info("Data already fetched for these parameters. Change parameters and click again to fetch new data.")
                
                # Auto-refresh
                auto_refresh = st.checkbox("Auto Refresh", value=False)
                if auto_refresh:
                    refresh_interval = st.slider(
                        "Refresh Interval (seconds)",
                        min_value=30,
                        max_value=300,
                        value=60,
                        step=30
                    )
                    st.info(f"Will refresh every {refresh_interval}s")
        
        else:
            st.warning("⚠️ Please initialize the system first")
    
    with col2:
        st.header("📊 Option Chain Data")
        
        if st.session_state.option_data:
            data = st.session_state.option_data
            
            # Display metrics - NOW WITH 6 METRICS INCLUDING SIGNALS
            col_a, col_b, col_c, col_d, col_e, col_f = st.columns(6)
            
            market_data = data.get('market_data', {})
            
            with col_a:
                spot_ltp = market_data.get('spot_ltp', 0)
                spot_change = market_data.get('spot_ltp_change', 0)
                st.metric(
                    "Spot LTP",
                    f"₹{spot_ltp:,.2f}",
                    f"{spot_change:+,.2f}"
                )
            
            with col_b:
                spot_change_pct = market_data.get('spot_ltp_change_pct', 0)
                st.metric(
                    "Change %",
                    f"{spot_change_pct:+.2f}%"
                )
            
            with col_c:
                pcr = market_data.get('pcr', 0)
                st.metric("PCR", f"{pcr:.2f}")
            
            with col_d:
                vix = market_data.get('india_vix', 0)
                st.metric("India VIX", f"{vix:.2f}")
            
            # NEW: Display signals
            with col_e:
                intraday = data.get('signals', {}).get('intraday', 'N/A')
                st.metric("Intraday", intraday)
            
            with col_f:
                weekly = data.get('signals', {}).get('weekly', 'N/A')
                st.metric("Weekly", weekly)
            
            st.divider()
            
            # Display option chain WITH ALL 20 COLUMNS
            option_chain = data.get('option_chain', [])
            
            if option_chain and len(option_chain) > 0:
                # Convert to DataFrame with ALL 20 columns
                rows = []
                for opt in option_chain:
                    try:
                        rows.append({
                            # CALL columns (8 columns)
                            'call_interpretation': opt.get('call', {}).get('interpretation', ''),
                            'call_avg_price': float(opt.get('call', {}).get('avg_price', 0)),
                            'call_iv': float(opt.get('call', {}).get('iv', 0)),
                            'call_oi_change': float(opt.get('call', {}).get('oi_change', 0)),
                            'call_oi': float(opt.get('call', {}).get('oi', 0)),
                            'call_volume': float(opt.get('call', {}).get('volume', 0)),
                            'call_ltp_change': float(opt.get('call', {}).get('ltp_change', 0)),
                            'call_ltp': float(opt.get('call', {}).get('ltp', 0)),
                            
                            # STRIKE (1 column)
                            'strike': float(opt.get('strike', 0)),
                            
                            # PUT columns (8 columns)
                            'put_ltp': float(opt.get('put', {}).get('ltp', 0)),
                            'put_ltp_change': float(opt.get('put', {}).get('ltp_change', 0)),
                            'put_volume': float(opt.get('put', {}).get('volume', 0)),
                            'put_oi': float(opt.get('put', {}).get('oi', 0)),
                            'put_oi_change': float(opt.get('put', {}).get('oi_change', 0)),
                            'put_iv': float(opt.get('put', {}).get('iv', 0)),
                            'put_avg_price': float(opt.get('put', {}).get('avg_price', 0)),
                            'put_interpretation': opt.get('put', {}).get('interpretation', '')
                        })
                    except Exception as e:
                        st.error(f"Error processing option row: {e}")
                        continue
                
                df = pd.DataFrame(rows)
                
                if df.empty:
                    st.error("❌ No valid data rows extracted from option chain")
                    st.write("**Raw option chain data:**")
                    st.json(option_chain[:3])  # Show first 3 items for debugging
                else:
                    # Show column count
                    st.info(f"📊 Displaying {len(df.columns)} columns x {len(df)} rows")
                    
                    # Create tabs for different views
                    view_tab1, view_tab2, view_tab3 = st.tabs(["📈 Chart", "📋 Table", "📊 OI Analysis"])
                    
                    with view_tab1:
                        fig = create_option_chain_chart(df, spot_ltp)
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with view_tab2:
                        # Format the dataframe for display
                        display_df = df.copy()
                        display_df['strike'] = display_df['strike'].astype(int)
                        
                        st.dataframe(
                            display_df,
                            use_container_width=True,
                            height=400
                        )
                        
                        # Download button
                        csv = display_df.to_csv(index=False)
                        st.download_button(
                            "📥 Download CSV",
                            csv,
                            f"option_chain_{data['symbol']}_{data['option_expiry']}.csv",
                            "text/csv"
                        )
                    
                    with view_tab3:
                        fig_oi = create_oi_chart(df)
                        st.plotly_chart(fig_oi, use_container_width=True)
            else:
                st.warning("⚠️ No option chain data available in response")
            
            # Data info
            with st.expander("📋 Fetch Information"):
                st.write(f"**Symbol:** {data.get('symbol')}")
                st.write(f"**Option Expiry:** {data.get('option_expiry')}")
                st.write(f"**Future Expiry:** {data.get('future_expiry')}")
                st.write(f"**Chain Length:** {data.get('chain_length')}")
                st.write(f"**Data Source:** {data.get('data_source')}")
                st.write(f"**Timestamp:** {data.get('timestamp')}")
                if st.session_state.fetch_time:
                    st.write(f"**Fetch Time:** {st.session_state.fetch_time:.2f}s")
                
                # Show signals
                if data.get('signals'):
                    st.write(f"**Intraday Signal:** {data['signals'].get('intraday', 'N/A')}")
                    st.write(f"**Weekly Signal:** {data['signals'].get('weekly', 'N/A')}")
        
        else:
            st.info("👆 Select parameters and click 'Fetch Option Data' to view the option chain")

# Tab 2: Analysis
with tab2:
    st.header("📈 Advanced Analysis")
    
    if st.session_state.option_data:
        data = st.session_state.option_data
        option_chain = data.get('option_chain', [])
        
        if option_chain:
            # Calculate analytics
            total_call_oi = sum(opt['call']['oi'] for opt in option_chain)
            total_put_oi = sum(opt['put']['oi'] for opt in option_chain)
            pcr = total_put_oi / total_call_oi if total_call_oi > 0 else 0
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Call OI", format_number(total_call_oi))
            with col2:
                st.metric("Total Put OI", format_number(total_put_oi))
            with col3:
                st.metric("PCR Ratio", f"{pcr:.3f}")
            
            st.divider()
            
            # Find max pain
            strikes = [opt['strike'] for opt in option_chain]
            max_pain_losses = []
            
            for strike in strikes:
                total_loss = 0
                for opt in option_chain:
                    if opt['strike'] < strike:
                        # Calls are ITM
                        total_loss += opt['call']['oi'] * (strike - opt['strike'])
                    elif opt['strike'] > strike:
                        # Puts are ITM
                        total_loss += opt['put']['oi'] * (opt['strike'] - strike)
                
                max_pain_losses.append(total_loss)
            
            max_pain_strike = strikes[max_pain_losses.index(min(max_pain_losses))]
            
            st.subheader("🎯 Max Pain Analysis")
            st.metric("Max Pain Strike", f"₹{max_pain_strike:,.0f}")
            
            # Plot max pain
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=strikes,
                y=max_pain_losses,
                mode='lines+markers',
                name='Total Loss',
                line=dict(color='purple', width=2)
            ))
            
            fig.add_vline(
                x=max_pain_strike,
                line_dash="dash",
                line_color="red",
                annotation_text=f"Max Pain: {max_pain_strike}",
                annotation_position="top"
            )
            
            fig.update_layout(
                title="Max Pain Analysis",
                xaxis_title="Strike Price",
                yaxis_title="Total Loss at Expiry",
                height=400,
                template='plotly_white'
            )
            
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Fetch option data first to see analysis")

# Tab 3: Settings
with tab3:
    st.header("⚙️ System Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🗑️ Data Management")
        
        if st.button("🗑️ Clear Cache"):
            # Clear cached files
            import os
            cache_files = ['excel_dates_cache.json', 'token_cache.json']
            cleared = []
            for f in cache_files:
                if os.path.exists(f):
                    os.remove(f)
                    cleared.append(f)
            
            if cleared:
                st.success(f"Cleared: {', '.join(cleared)}")
            else:
                st.info("No cache files to clear")
        
        if st.button("💾 Export Current Data"):
            if st.session_state.option_data:
                import json
                json_data = json.dumps(st.session_state.option_data, indent=2)
                st.download_button(
                    "Download JSON",
                    json_data,
                    f"option_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    "application/json"
                )
            else:
                st.warning("No data to export")
    
    with col2:
        st.subheader("ℹ️ System Information")
        
        st.write("**System Status:**")
        st.write(f"- Initialized: {'✅' if st.session_state.system_initialized else '❌'}")
        st.write(f"- Mode: {'API' if st.session_state.api_mode else 'Local'}")
        
        if st.session_state.dropdown_options:
            opts = st.session_state.dropdown_options
            st.write(f"- Symbols: {len(opts.get('symbols', []))}")
            st.write(f"- Option Dates: {len(opts.get('option_expiry', []))}")
            st.write(f"- Future Dates: {len(opts.get('future_expiry', []))}")
        
        if st.session_state.last_fetch_params:
            st.write("\n**Last Fetch:**")
            params = st.session_state.last_fetch_params
            st.write(f"- Symbol: {params['symbol']}")
            st.write(f"- Expiry: {params['option_expiry']}")

# Footer
st.divider()
st.caption("📊 Option Trading System | Built with Streamlit")