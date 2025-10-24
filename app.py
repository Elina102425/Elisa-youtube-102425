import os
import json
import time
import yaml
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime
from tenacity import retry, wait_fixed, stop_after_attempt
from typing import Dict, List, Optional

# Import custom modules
try:
    from google_utils import (
        create_sheet, read_sheet_df, ensure_control_columns,
        upload_docx_as_gdoc, create_gdoc_from_text, copy_gdoc,
        replace_placeholders, export_pdf, ensure_folder, write_back
    )
    from llm_providers import run_agent
except ImportError as e:
    st.error(f"Import error: {e}. Ensure google_utils.py and llm_providers.py are present.")
    st.stop()

# Constants
TRIGGER_COL = "Generate"
STATUS_COL = "Status"
URL_COL = "Doc URL"
TS_COL = "Generated At"
PDF_URL_COL = "PDF URL"

# Theme definitions
THEMES = {
    "Cyber Neon": {
        "primary": "#00ffff",
        "secondary": "#ff00ff",
        "background": "#0a0e27",
        "surface": "#1a1f3a",
        "text": "#e0e0e0",
        "accent": "#00ff88"
    },
    "Sunset Glow": {
        "primary": "#ff6b6b",
        "secondary": "#ffd93d",
        "background": "#2d1b2e",
        "surface": "#3d2b3e",
        "text": "#f5f5f5",
        "accent": "#ff9a56"
    },
    "Ocean Depth": {
        "primary": "#0077be",
        "secondary": "#00b4d8",
        "background": "#011627",
        "surface": "#1a2332",
        "text": "#f1f1f1",
        "accent": "#48cae4"
    },
    "Forest Twilight": {
        "primary": "#2d6a4f",
        "secondary": "#52b788",
        "background": "#1b263b",
        "surface": "#2d3e50",
        "text": "#e8f5e9",
        "accent": "#95d5b2"
    },
    "Royal Purple": {
        "primary": "#7b2cbf",
        "secondary": "#c77dff",
        "background": "#10002b",
        "surface": "#240046",
        "text": "#f0e6ff",
        "accent": "#e0aaff"
    },
    "Desert Sunset": {
        "primary": "#d4a373",
        "secondary": "#ee9b00",
        "background": "#1a1410",
        "surface": "#2d2418",
        "text": "#fdf6e3",
        "accent": "#ca6702"
    },
    "Arctic Ice": {
        "primary": "#4cc9f0",
        "secondary": "#7209b7",
        "background": "#03045e",
        "surface": "#023e8a",
        "text": "#caf0f8",
        "accent": "#90e0ef"
    },
    "Volcanic Ash": {
        "primary": "#e63946",
        "secondary": "#f77f00",
        "background": "#1d1d1d",
        "surface": "#2d2d2d",
        "text": "#f1faee",
        "accent": "#ffb703"
    },
    "Mint Fresh": {
        "primary": "#06ffa5",
        "secondary": "#00d9ff",
        "background": "#0f1419",
        "surface": "#1a2027",
        "text": "#e8fff3",
        "accent": "#4fffb0"
    },
    "Lavender Dreams": {
        "primary": "#b8a7d4",
        "secondary": "#d4a5d8",
        "background": "#1a1423",
        "surface": "#2d243a",
        "text": "#f8f4ff",
        "accent": "#c4b5f3"
    },
    "Monochrome Pro": {
        "primary": "#ffffff",
        "secondary": "#b0b0b0",
        "background": "#000000",
        "surface": "#1a1a1a",
        "text": "#ffffff",
        "accent": "#808080"
    },
    "Coral Reef": {
        "primary": "#ff6f61",
        "secondary": "#ffb399",
        "background": "#1a0f14",
        "surface": "#2d1f24",
        "text": "#fff5f5",
        "accent": "#ff9a8a"
    },
    "Emerald Night": {
        "primary": "#50c878",
        "secondary": "#00a86b",
        "background": "#0a1612",
        "surface": "#142822",
        "text": "#e8fff2",
        "accent": "#7cf5a0"
    },
    "Golden Hour": {
        "primary": "#ffd700",
        "secondary": "#ffed4e",
        "background": "#1a1410",
        "surface": "#2d2418",
        "text": "#fffef0",
        "accent": "#ffe55c"
    },
    "Deep Space": {
        "primary": "#8b5cf6",
        "secondary": "#ec4899",
        "background": "#0c0a1d",
        "surface": "#1a1631",
        "text": "#f3e8ff",
        "accent": "#a78bfa"
    },
    "Crimson Edge": {
        "primary": "#dc143c",
        "secondary": "#ff6b9d",
        "background": "#1a0a0f",
        "surface": "#2d141f",
        "text": "#fff0f5",
        "accent": "#ff4d6d"
    },
    "Teal Fusion": {
        "primary": "#14b8a6",
        "secondary": "#06b6d4",
        "background": "#0f1419",
        "surface": "#1a2530",
        "text": "#e0f2fe",
        "accent": "#5eead4"
    },
    "Amber Glow": {
        "primary": "#f59e0b",
        "secondary": "#fbbf24",
        "background": "#1c1410",
        "surface": "#2d2318",
        "text": "#fffbeb",
        "accent": "#fcd34d"
    },
    "Indigo Wave": {
        "primary": "#6366f1",
        "secondary": "#818cf8",
        "background": "#0f0f1e",
        "surface": "#1e1e3f",
        "text": "#e0e7ff",
        "accent": "#a5b4fc"
    },
    "Rose Garden": {
        "primary": "#f43f5e",
        "secondary": "#fb7185",
        "background": "#1f0a13",
        "surface": "#2d1420",
        "text": "#fff1f2",
        "accent": "#fda4af"
    }
}

def apply_theme(theme_name: str):
    """Apply custom theme styling"""
    theme = THEMES.get(theme_name, THEMES["Cyber Neon"])
    
    st.markdown(f"""
    <style>
        :root {{
            --primary-color: {theme['primary']};
            --secondary-color: {theme['secondary']};
            --background-color: {theme['background']};
            --surface-color: {theme['surface']};
            --text-color: {theme['text']};
            --accent-color: {theme['accent']};
        }}
        
        .stApp {{
            background: linear-gradient(135deg, {theme['background']} 0%, {theme['surface']} 100%);
            color: {theme['text']};
        }}
        
        .stButton>button {{
            background: linear-gradient(90deg, {theme['primary']} 0%, {theme['secondary']} 100%);
            color: {theme['background']};
            border: none;
            border-radius: 12px;
            padding: 0.75rem 2rem;
            font-weight: 600;
            box-shadow: 0 4px 15px rgba(0,0,0,0.3);
            transition: all 0.3s ease;
        }}
        
        .stButton>button:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(0,0,0,0.4);
        }}
        
        .metric-card {{
            background: {theme['surface']};
            border-left: 4px solid {theme['primary']};
            padding: 1.5rem;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            margin: 0.5rem 0;
        }}
        
        .status-badge {{
            display: inline-block;
            padding: 0.25rem 0.75rem;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }}
        
        .status-success {{
            background: linear-gradient(90deg, #10b981 0%, #34d399 100%);
            color: white;
        }}
        
        .status-error {{
            background: linear-gradient(90deg, #ef4444 0%, #f87171 100%);
            color: white;
        }}
        
        .status-pending {{
            background: linear-gradient(90deg, #f59e0b 0%, #fbbf24 100%);
            color: white;
        }}
        
        .agent-card {{
            background: {theme['surface']};
            border: 2px solid {theme['primary']};
            border-radius: 15px;
            padding: 1.5rem;
            margin: 1rem 0;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        }}
        
        .pulse {{
            animation: pulse 2s ease-in-out infinite;
        }}
        
        @keyframes pulse {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.5; }}
        }}
        
        .stTabs [data-baseweb="tab-list"] {{
            gap: 8px;
            background: {theme['surface']};
            border-radius: 12px;
            padding: 0.5rem;
        }}
        
        .stTabs [data-baseweb="tab"] {{
            background: transparent;
            border-radius: 8px;
            color: {theme['text']};
            padding: 0.75rem 1.5rem;
        }}
        
        .stTabs [aria-selected="true"] {{
            background: linear-gradient(90deg, {theme['primary']} 0%, {theme['secondary']} 100%);
            color: {theme['background']};
        }}
        
        .dataframe {{
            background: {theme['surface']} !important;
            border-radius: 12px;
            overflow: hidden;
        }}
        
        .progress-ring {{
            stroke: {theme['primary']};
            fill: none;
            stroke-width: 8;
            stroke-linecap: round;
        }}
    </style>
    """, unsafe_allow_html=True)

def render_status_badge(status: str) -> str:
    """Render status badge with appropriate styling"""
    if pd.isna(status) or status == "":
        badge_class = "status-pending"
        badge_text = "Pending"
    elif str(status).lower().startswith("done"):
        badge_class = "status-success"
        badge_text = "✓ Done"
    elif str(status).lower().startswith("error"):
        badge_class = "status-error"
        badge_text = "✗ Error"
    else:
        badge_class = "status-pending"
        badge_text = status
    
    return f'<span class="status-badge {badge_class}">{badge_text}</span>'

def create_circular_progress(percentage: float, size: int = 120) -> str:
    """Create circular progress indicator"""
    radius = 50
    circumference = 2 * 3.14159 * radius
    offset = circumference - (percentage / 100 * circumference)
    
    return f"""
    <svg width="{size}" height="{size}" viewBox="0 0 120 120">
        <circle cx="60" cy="60" r="{radius}" class="progress-ring" 
                style="stroke: rgba(255,255,255,0.1); stroke-dasharray: {circumference}; stroke-dashoffset: 0;"/>
        <circle cx="60" cy="60" r="{radius}" class="progress-ring" 
                style="stroke-dasharray: {circumference}; stroke-dashoffset: {offset}; 
                       transform: rotate(-90deg); transform-origin: 60px 60px;"/>
        <text x="60" y="65" text-anchor="middle" style="fill: white; font-size: 20px; font-weight: bold;">
            {int(percentage)}%
        </text>
    </svg>
    """

# Page configuration
st.set_page_config(
    page_title="Agentic Data Studio Pro",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if "theme" not in st.session_state:
    st.session_state.theme = "Cyber Neon"
if "agents_yaml" not in st.session_state:
    st.session_state.agents_yaml = ""

# Apply selected theme
apply_theme(st.session_state.theme)

# Sidebar configuration
with st.sidebar:
    st.markdown("# ⚙️ Configuration")
    
    # Theme selector with preview
    st.markdown("### 🎨 Theme Selection")
    selected_theme = st.selectbox(
        "Choose your theme",
        options=list(THEMES.keys()),
        index=list(THEMES.keys()).index(st.session_state.theme)
    )
    if selected_theme != st.session_state.theme:
        st.session_state.theme = selected_theme
        st.rerun()
    
    st.markdown("---")
    
    # API Configuration
    st.markdown("### 🔑 API Keys")
    with st.expander("Configure API Keys", expanded=False):
        gsa_inline = st.text_area(
            "Google Service Account JSON",
            value=os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", ""),
            type="password",
            height=100,
            help="Paste your service account JSON here"
        )
        openai_key = st.text_input(
            "OpenAI API Key",
            value=os.environ.get("OPENAI_API_KEY", ""),
            type="password"
        )
        gemini_key = st.text_input(
            "Google AI API Key (Gemini)",
            value=os.environ.get("GOOGLE_API_KEY", ""),
            type="password"
        )
        grok_key = st.text_input(
            "xAI API Key (Grok)",
            value=os.environ.get("XAI_API_KEY", ""),
            type="password"
        )
        
        if st.button("💾 Apply Keys", use_container_width=True):
            if gsa_inline:
                os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"] = gsa_inline
            if openai_key:
                os.environ["OPENAI_API_KEY"] = openai_key
            if gemini_key:
                os.environ["GOOGLE_API_KEY"] = gemini_key
            if grok_key:
                os.environ["XAI_API_KEY"] = grok_key
            st.success("✓ Keys applied successfully!")
    
    st.markdown("---")
    
    # Agents configuration
    st.markdown("### 🤖 Agents Configuration")
    agents_file = st.file_uploader(
        "Upload agents.yaml",
        type=["yaml", "yml"],
        help="Upload a YAML file with agent definitions"
    )
    
    if agents_file:
        try:
            st.session_state.agents_yaml = agents_file.read().decode("utf-8")
            st.success("✓ Agents loaded!")
        except Exception as e:
            st.error(f"Error loading agents: {e}")
    
    if st.button("📥 Load Default Agents", use_container_width=True):
        # Will be loaded from default in main area
        st.info("Default agents will be loaded in the Agents Runner tab")
    
    st.markdown("---")
    st.caption("💡 Tip: Set API keys in Hugging Face Space secrets for production")

# Main header
st.markdown("""
    <div style='text-align: center; padding: 2rem 0;'>
        <h1 style='font-size: 3rem; font-weight: 800; margin: 0;'>
            🧠 Agentic Data Studio Pro
        </h1>
        <p style='font-size: 1.2rem; opacity: 0.8; margin-top: 0.5rem;'>
            Advanced Data Mining & Document Automation Platform
        </p>
    </div>
""", unsafe_allow_html=True)

# Main tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Dashboard",
    "📝 Sheet Setup",
    "📄 Template Manager",
    "🚀 Generate Docs",
    "🤖 Agents Runner"
])

with tab1:
    st.markdown("### 📊 Real-time Analytics Dashboard")
    
    # Check if we have data
    if "spreadsheet_id" in st.session_state and st.session_state.spreadsheet_id:
        try:
            df = read_sheet_df(st.session_state.spreadsheet_id)
            
            if not df.empty and STATUS_COL in df.columns:
                # Calculate metrics
                total_rows = len(df)
                done_count = df[STATUS_COL].astype(str).str.startswith("Done").sum()
                error_count = df[STATUS_COL].astype(str).str.startswith("Error").sum()
                pending_count = total_rows - done_count - error_count
                success_rate = (done_count / total_rows * 100) if total_rows > 0 else 0
                
                # Metrics row
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.markdown(f"""
                        <div class='metric-card'>
                            <h3 style='margin: 0; font-size: 2.5rem;'>{total_rows}</h3>
                            <p style='margin: 0; opacity: 0.8;'>Total Records</p>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                        <div class='metric-card'>
                            <h3 style='margin: 0; font-size: 2.5rem; color: #10b981;'>{done_count}</h3>
                            <p style='margin: 0; opacity: 0.8;'>Completed</p>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                        <div class='metric-card'>
                            <h3 style='margin: 0; font-size: 2.5rem; color: #ef4444;'>{error_count}</h3>
                            <p style='margin: 0; opacity: 0.8;'>Errors</p>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                        <div class='metric-card'>
                            <h3 style='margin: 0; font-size: 2.5rem; color: #f59e0b;'>{pending_count}</h3>
                            <p style='margin: 0; opacity: 0.8;'>Pending</p>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col5:
                    st.markdown(f"""
                        <div class='metric-card'>
                            <h3 style='margin: 0; font-size: 2.5rem;'>{success_rate:.1f}%</h3>
                            <p style='margin: 0; opacity: 0.8;'>Success Rate</p>
                        </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Charts row
                chart_col1, chart_col2 = st.columns(2)
                
                with chart_col1:
                    # Status distribution pie chart
                    status_data = pd.DataFrame({
                        'Status': ['Completed', 'Errors', 'Pending'],
                        'Count': [done_count, error_count, pending_count],
                        'Color': ['#10b981', '#ef4444', '#f59e0b']
                    })
                    
                    fig_pie = px.pie(
                        status_data,
                        values='Count',
                        names='Status',
                        title='Status Distribution',
                        color='Status',
                        color_discrete_map={
                            'Completed': '#10b981',
                            'Errors': '#ef4444',
                            'Pending': '#f59e0b'
                        }
                    )
                    fig_pie.update_layout(
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        font_color=THEMES[st.session_state.theme]['text']
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                with chart_col2:
                    # Circular progress
                    st.markdown("### Overall Progress")
                    progress_html = create_circular_progress(success_rate)
                    st.markdown(f"<div style='text-align: center;'>{progress_html}</div>", unsafe_allow_html=True)
                
                # Recent activity table
                st.markdown("### 📋 Recent Activity")
                display_df = df.copy()
                if STATUS_COL in display_df.columns:
                    display_df[STATUS_COL] = display_df[STATUS_COL].apply(
                        lambda x: render_status_badge(x)
                    )
                st.markdown(display_df.tail(10).to_html(escape=False, index=False), unsafe_allow_html=True)
                
            else:
                st.info("📊 No data available yet. Create some records to see analytics!")
        
        except Exception as e:
            st.error(f"Error loading dashboard data: {e}")
    else:
        st.info("👋 Welcome! Create or connect a spreadsheet to view your dashboard.")
        st.markdown("""
            <div style='text-align: center; padding: 3rem;'>
                <p style='font-size: 1.2rem; opacity: 0.7;'>
                    Get started by creating a new sheet in the <strong>Sheet Setup</strong> tab
                </p>
            </div>
        """, unsafe_allow_html=True)

with tab2:
    st.markdown("### 📝 Google Sheet Configuration")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        sheet_title = st.text_input(
            "📌 Sheet Title",
            value="Data Mining Workspace",
            help="Name for your new Google Sheet"
        )
        
        headers_input = st.text_input(
            "📋 Column Headers (comma-separated)",
            value="Company, Industry, Revenue, Employees, Location, Website, Contact",
            help="These will become columns in your sheet and placeholders in templates"
        )
    
    with col2:
        st.markdown("#### Destination Settings")
        folder_id = st.text_input(
            "📁 Folder ID (optional)",
            help="Google Drive folder ID for generated docs"
        )
        create_dest_folder = st.text_input(
            "🆕 Or create new folder",
            help="Name for a new destination folder"
        )
    
    if st.button("🚀 Create Sheet", use_container_width=True):
        try:
            with st.spinner("Creating your sheet..."):
                dest_folder_id = ensure_folder(folder_id, create_dest_folder) if (folder_id or create_dest_folder) else None
                
                headers = [h.strip() for h in headers_input.split(",") if h.strip()]
                control_cols = [TRIGGER_COL, STATUS_COL, URL_COL, TS_COL, PDF_URL_COL]
                full_headers = headers + control_cols
                
                info = create_sheet(sheet_title, full_headers)
                
                st.session_state.spreadsheet_id = info["spreadsheet_id"]
                st.session_state.spreadsheet_url = info["url"]
                st.session_state.dest_folder_id = dest_folder_id
                
                st.success("✅ Sheet created successfully!")
                st.markdown(f"""
                    <div class='metric-card'>
                        <p><strong>Sheet URL:</strong></p>
                        <a href="{info['url']}" target="_blank">{info['url']}</a>
                    </div>
                """, unsafe_allow_html=True)
                
                if dest_folder_id:
                    st.info(f"📁 Destination folder: {dest_folder_id}")
        
        except Exception as e:
            st.error(f"❌ Error creating sheet: {e}")

with tab3:
    st.markdown("### 📄 Template Management")
    
    st.markdown("""
        <div class='agent-card'>
            <p>Choose how to provide your document template with <code>{{placeholders}}</code></p>
        </div>
    """, unsafe_allow_html=True)
    
    template_mode = st.radio(
        "Template Source",
        ["Existing Google Doc", "Upload .docx", "Paste Text"],
        horizontal=True
    )
    
    if template_mode == "Existing Google Doc":
        template_id = st.text_input(
            "Google Doc ID",
            help="Extract from the doc URL: docs.google.com/document/d/{ID}/edit"
        )
        if st.button("📌 Use This Template"):
            if template_id.strip():
                st.session_state.template_id = template_id.strip()
                st.success(f"✅ Template set: {template_id.strip()}")
            else:
                st.warning("Please enter a template ID")
    
    elif template_mode == "Upload .docx":
        col1, col2 = st.columns([3, 1])
        with col1:
            tpl_upload = st.file_uploader("Upload .docx file", type=["docx"])
        with col2:
            tpl_title = st.text_input("Template name", value="Uploaded Template")
        
        if tpl_upload and st.button("📤 Convert & Save"):
            try:
                with st.spinner("Converting to Google Doc..."):
                    file_bytes = tpl_upload.read()
                    dest_folder_id = st.session_state.get("dest_folder_id")
                    new_id = upload_docx_as_gdoc(file_bytes, tpl_title, dest_folder_id)
                    st.session_state.template_id = new_id
                    st.success(f"✅ Converted successfully! ID: {new_id}")
            except Exception as e:
                st.error(f"❌ Conversion error: {e}")
    
    else:  # Paste Text
        col1, col2 = st.columns([3, 1])
        with col1:
            tpl_text = st.text_area(
                "Paste template text (use {{placeholders}})",
                height=300,
                placeholder="Example:\n\nDear {{Name}},\n\nWe analyzed {{Company}} in the {{Industry}} sector..."
            )
        with col2:
            tpl_title = st.text_input("Template name", value="Text Template")
        
        if tpl_text.strip() and st.button("💾 Create Template"):
            try:
                with st.spinner("Creating Google Doc..."):
                    dest_folder_id = st.session_state.get("dest_folder_id")
                    new_id = create_gdoc_from_text(tpl_title, tpl_text, dest_folder_id)
                    st.session_state.template_id = new_id
                    st.success(f"✅ Template created! ID: {new_id}")
            except Exception as e:
                st.error(f"❌ Creation error: {e}")

with tab4:
    st.markdown("### 🚀 Document Generation Engine")
    
    if not st.session_state.get("spreadsheet_id"):
        st.warning("⚠️ Please create or connect a spreadsheet first (Sheet Setup tab)")
    elif not st.session_state.get("template_id"):
        st.warning("⚠️ Please configure a template first (Template Manager tab)")
    else:
        try:
            spreadsheet_id = st.session_state.spreadsheet_id
            ensure_control_columns(spreadsheet_id, [TRIGGER_COL, STATUS_COL, URL_COL, TS_COL, PDF_URL_COL])
            df = read_sheet_df(spreadsheet_id)
            
            if df.empty:
                st.info("📭 No data rows found. Add some data to your sheet first.")
            else:
                # Data preview
                st.markdown("#### 📊 Data Preview")
                st.dataframe(df, use_container_width=True)
                
                # Generation options
                st.markdown("#### ⚙️ Generation Options")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    gen_mode = st.selectbox(
                        "Selection Mode",
                        ["Rows with Generate=TRUE", "All rows without Doc URL", "Manual selection"]
                    )
                
                with col2:
                    create_pdf = st.checkbox("📄 Export PDF", value=True)
                
                with col3:
                    custom_filename = st.text_input(
                        "Filename pattern",
                        value="{{Company}} - {{Industry}} Report"
                    )
                
                manual_sel = []
                if gen_mode == "Manual selection":
                    manual_sel = st.multiselect(
                        "Select rows (1-based, including header)",
                        options=list(range(2, len(df) + 2))
                    )
                
                # Generate button
                if st.button("🚀 Generate Documents Now", use_container_width=True, type="primary"):
                    success_count = 0
                    error_count = 0
                    
                    progress_bar = st.progress(0)
                    status_container = st.container()
                    
                    for idx, row in df.iterrows():
                        row_index_1 = idx + 2
                        
                        # Check selection criteria
                        if gen_mode == "Rows with Generate=TRUE":
                            trig_val = str(row.get(TRIGGER_COL, "")).strip().lower()
                            if trig_val not in ["true", "yes", "y", "1", "checked"]:
                                continue
                        elif gen_mode == "All rows without Doc URL":
                            if str(row.get(URL_COL, "")).strip():
                                continue
                        elif gen_mode == "Manual selection":
                            if row_index_1 not in manual_sel:
                                continue
                        
                        # Skip if already has URL (idempotency)
                        if str(row.get(URL_COL, "")).strip():
                            continue
                        
                        try:
                            # Build filename
                            filename = custom_filename
                            for col in df.columns:
                                filename = filename.replace(f"{{{{{col}}}}}", str(row.get(col, "")))
                            
                            # Create document
                            new_doc_id = copy_gdoc(
                                st.session_state.template_id,
                                filename,
                                st.session_state.get("dest_folder_id")
                            )
                            
                            # Replace placeholders
                            mapping = {col: row.get(col, "") for col in df.columns}
                            replace_placeholders(new_doc_id, mapping)
                            
                            doc_url = f"https://docs.google.com/document/d/{new_doc_id}/edit"
                            
                            updates = {
                                STATUS_COL: "Done",
                                URL_COL: doc_url,
                                TS_COL: datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            }
                            
                            if create_pdf:
                                pdf_url = f"https://docs.google.com/document/d/{new_doc_id}/export?format=pdf"
                                updates[PDF_URL_COL] = pdf_url
                            
                            write_back(spreadsheet_id, row_index_1, updates)
                            success_count += 1
                            
                            with status_container:
                                st.success(f"✅ Row {row_index_1}: {filename}")
                        
                        except Exception as e:
                            error_count += 1
                            write_back(spreadsheet_id, row_index_1, {STATUS_COL: f"Error: {str(e)[:100]}"})
                            with status_container:
                                st.error(f"❌ Row {row_index_1}: {str(e)[:100]}")
                        
                        progress_bar.progress(min(1.0, (idx + 1) / len(df)))
                    
                    st.markdown(f"""
                        <div class='metric-card' style='text-align: center;'>
                            <h3>Generation Complete</h3>
                            <p>✅ Success: {success_count} | ❌ Errors: {error_count}</p>
                        </div>
                    """, unsafe_allow_html=True)
        
        except Exception as e:
            st.error(f"Error loading data: {e}")

with tab5:
    st.markdown("### 🤖 AI Agents Runner")
    
    # Load agents configuration
    if not st.session_state.agents_yaml:
        # Load default comprehensive agents
        default_agents = """agents:
  - id: data_validator
    name: Data Quality Validator
    provider: openai
    model: gpt-4o-mini
    temperature: 0.1
    max_tokens: 512
    prompt: |
      You are a strict data quality validator. Analyze the data row and identify:
      1. Missing required fields
      2. Invalid data formats
      3. Suspicious or inconsistent values
      4. Data completeness score (0-100)
      
      Row data: {{row_json}}
      Required fields: {{required_fields_csv}}
      
      Return a JSON with: {{"score": 0-100, "issues": [], "recommendations": []}}

  - id: company_enricher
    name: Company Data Enricher
    provider: gemini
    model: gemini-2.0-flash-exp
    temperature: 0.4
    max_tokens: 800
    prompt: |
      Enrich this company data with intelligent insights:
      - Industry classification and sub-sector
      - Estimated market position
      - Growth indicators
      - Competitive landscape
      - Key business model characteristics
      
      Company data: {{row_json}}
      
      Provide a structured enrichment in 3-4 sentences.

  - id: sentiment_analyzer
    name: Content Sentiment Analyzer
    provider: openai
    model: gpt-4o-mini
    temperature: 0.2
    max_tokens: 400
    prompt: |
      Analyze the sentiment and tone of the content in this data row.
      Provide:
      - Overall sentiment (positive/negative/neutral)
      - Confidence score (0-100)
      - Key emotional indicators
      - Tone characteristics
      
      Data: {{row_json}}

  - id: pattern_detector
    name: Pattern & Anomaly Detector
    provider: grok
    model: grok-beta
    temperature: 0.3
    max_tokens: 600
    prompt: |
      Analyze this data point for patterns and anomalies:
      - Unusual patterns compared to typical data
      - Statistical outliers
      - Data quality concerns
      - Recommendations for further investigation
      
      Data: {{row_json}}
      
      Be specific and actionable.

  - id: summary_generator
    name: Executive Summary Generator
    provider: gemini
    model: gemini-2.0-flash-exp
    temperature: 0.5
    max_tokens: 700
    prompt: |
      Generate a concise executive summary for this data record.
      Focus on:
      - Key highlights
      - Critical metrics
      - Notable insights
      - Actionable recommendations
      
      Data: {{row_json}}
      
      Keep it under 150 words, professional tone.

  - id: category_classifier
    name: Smart Category Classifier
    provider: openai
    model: gpt-4o-mini
    temperature: 0.2
    max_tokens: 300
    prompt: |
      Classify this data into relevant categories:
      - Primary category
      - Secondary categories (if applicable)
      - Tags for searchability
      - Confidence level for each classification
      
      Data: {{row_json}}
      
      Return structured classification data.

  - id: risk_assessor
    name: Risk Assessment Agent
    provider: grok
    model: grok-beta
    temperature: 0.3
    max_tokens: 500
    prompt: |
      Perform a risk assessment on this data:
      - Identify potential risks or red flags
      - Assign risk level (low/medium/high)
      - Suggest mitigation strategies
      - Compliance considerations
      
      Data: {{row_json}}

  - id: insight_extractor
    name: Key Insights Extractor
    provider: gemini
    model: gemini-2.0-flash-exp
    temperature: 0.4
    max_tokens: 600
    prompt: |
      Extract and synthesize key insights from this data:
      - Most important findings
      - Hidden patterns or correlations
      - Strategic implications
      - Questions for further investigation
      
      Data: {{row_json}}
      
      Be analytical and thought-provoking.

  - id: completeness_scorer
    name: Data Completeness Scorer
    provider: openai
    model: gpt-4o-mini
    temperature: 0.1
    max_tokens: 400
    prompt: |
      Score the completeness and quality of this data record:
      - Completeness score (0-100)
      - Quality score (0-100)
      - Missing information analysis
      - Suggestions for improvement
      
      Data: {{row_json}}
      Required fields: {{required_fields_csv}}

  - id: trend_analyzer
    name: Trend Analysis Agent
    provider: grok
    model: grok-beta
    temperature: 0.3
    max_tokens: 500
    prompt: |
      Analyze trends and patterns in this data point:
      - Temporal trends (if dates present)
      - Value trends
      - Comparative analysis
      - Future projections or implications
      
      Data: {{row_json}}
"""
        st.session_state.agents_yaml = default_agents
    
    try:
        agents_config = yaml.safe_load(st.session_state.agents_yaml)
        agents = agents_config.get("agents", [])
        
        if not agents:
            st.warning("⚠️ No agents found. Upload agents.yaml in the sidebar.")
        else:
            st.markdown(f"<div class='metric-card'><p>📦 Loaded <strong>{len(agents)}</strong> agents</p></div>", unsafe_allow_html=True)
            
            # Agent configuration section
            st.markdown("#### 🛠️ Configure Agents")
            
            for i, agent in enumerate(agents):
                with st.expander(f"🤖 {agent.get('name', f'Agent {i+1}')} ({agent.get('id', '')})", expanded=False):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        provider_options = ["openai", "gemini", "grok"]
                        current_provider = agent.get("provider", "openai")
                        agent["provider"] = st.selectbox(
                            "Provider",
                            provider_options,
                            index=provider_options.index(current_provider) if current_provider in provider_options else 0,
                            key=f"provider_{i}"
                        )
                        
                        agent["model"] = st.text_input(
                            "Model",
                            value=agent.get("model", ""),
                            key=f"model_{i}"
                        )
                    
                    with col2:
                        agent["temperature"] = st.slider(
                            "Temperature",
                            0.0, 2.0,
                            float(agent.get("temperature", 0.3)),
                            0.1,
                            key=f"temp_{i}"
                        )
                        
                        agent["max_tokens"] = st.number_input(
                            "Max Tokens",
                            min_value=50,
                            max_value=4096,
                            value=int(agent.get("max_tokens", 512)),
                            key=f"tokens_{i}"
                        )
                    
                    agent["prompt"] = st.text_area(
                        "Prompt Template",
                        value=agent.get("prompt", ""),
                        height=200,
                        key=f"prompt_{i}"
                    )
            
            st.markdown("---")
            
            # Execution section
            st.markdown("#### 🎯 Execute Agents")
            
            col1, col2 = st.columns(2)
            
            with col1:
                spreadsheet_id = st.text_input(
                    "Spreadsheet ID",
                    value=st.session_state.get("spreadsheet_id", ""),
                    help="The sheet to analyze"
                )
            
            with col2:
                if spreadsheet_id:
                    try:
                        df = read_sheet_df(spreadsheet_id)
                        row_number = st.number_input(
                            "Row to analyze (2 = first data row)",
                            min_value=2,
                            max_value=len(df) + 1 if not df.empty else 2,
                            value=2
                        )
                    except:
                        row_number = 2
                        st.warning("Could not load sheet data")
                else:
                    row_number = 2
            
            col1, col2 = st.columns(2)
            
            with col1:
                required_fields = st.text_input(
                    "Required fields (comma-separated)",
                    value="Company, Industry, Revenue",
                    help="Fields that must be present in data"
                )
            
            with col2:
                num_agents = st.number_input(
                    "Number of agents to run",
                    min_value=1,
                    max_value=len(agents),
                    value=min(3, len(agents))
                )
            
            if st.button("🚀 Run Agents Pipeline", use_container_width=True, type="primary"):
                if not spreadsheet_id:
                    st.error("Please provide a spreadsheet ID")
                else:
                    try:
                        df = read_sheet_df(spreadsheet_id)
                        
                        if row_number - 2 >= len(df) or row_number - 2 < 0:
                            st.error(f"Row {row_number} is out of range")
                        else:
                            row_data = df.iloc[row_number - 2].to_dict()
                            context = {
                                "row_json": json.dumps(row_data, ensure_ascii=False, indent=2),
                                "required_fields_csv": required_fields
                            }
                            
                            st.markdown(f"""
                                <div class='agent-card'>
                                    <h4>📊 Analyzing Row {row_number}</h4>
                                    <pre>{json.dumps(row_data, indent=2)[:500]}...</pre>
                                </div>
                            """, unsafe_allow_html=True)
                            
                            results = []
                            progress_bar = st.progress(0)
                            
                            for i in range(int(num_agents)):
                                agent = agents[i]
                                agent_name = agent.get("name", agent.get("id", f"Agent {i+1}"))
                                
                                with st.status(f"🤖 Running: {agent_name}", expanded=True) as status:
                                    try:
                                        st.write(f"Provider: {agent.get('provider')} | Model: {agent.get('model')}")
                                        
                                        output = run_agent(agent, context)
                                        results.append({
                                            "agent": agent_name,
                                            "status": "success",
                                            "output": output
                                        })
                                        
                                        status.update(label=f"✅ {agent_name} - Complete", state="complete")
                                        time.sleep(0.5)
                                    
                                    except Exception as e:
                                        results.append({
                                            "agent": agent_name,
                                            "status": "error",
                                            "output": f"Error: {str(e)}"
                                        })
                                        status.update(label=f"❌ {agent_name} - Error", state="error")
                                
                                progress_bar.progress((i + 1) / num_agents)
                            
                            # Display results
                            st.markdown("---")
                            st.markdown("### 📊 Agent Results")
                            
                            for result in results:
                                if result["status"] == "success":
                                    st.markdown(f"""
                                        <div class='agent-card'>
                                            <h4>✅ {result['agent']}</h4>
                                        </div>
                                    """, unsafe_allow_html=True)
                                    st.markdown(result["output"])
                                else:
                                    st.markdown(f"""
                                        <div class='agent-card' style='border-color: #ef4444;'>
                                            <h4>❌ {result['agent']}</h4>
                                        </div>
                                    """, unsafe_allow_html=True)
                                    st.error(result["output"])
                                
                                st.markdown("<br>", unsafe_allow_html=True)
                    
                    except Exception as e:
                        st.error(f"Error running agents: {e}")
    
    except yaml.YAMLError as e:
        st.error(f"Error parsing agents.yaml: {e}")
    except Exception as e:
        st.error(f"Error in agents runner: {e}")

# Footer
st.markdown("---")
st.markdown("""
    <div style='text-align: center; padding: 2rem; opacity: 0.6;'>
        <p>🧠 Agentic Data Studio Pro | Powered by Claude, OpenAI, Google AI & xAI</p>
        <p style='font-size: 0.9rem;'>Advanced data mining and document automation platform</p>
    </div>
""", unsafe_allow_html=True)
