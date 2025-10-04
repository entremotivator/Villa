import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, time, timedelta
import pandas as pd
import json
import plotly.express as px
import plotly.graph_objects as go
from typing import Optional, Dict, Any, List
import hashlib

# Page configuration
st.set_page_config(
    page_title="Villa Booking Management System Pro",
    page_icon="🏖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced CSS with all styles
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');
    
    * { font-family: 'Poppins', sans-serif; }
    
    .main {
        padding: 2rem;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        padding: 0.75rem 1.5rem;
        border-radius: 12px;
        border: none;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
    }
    
    .stButton>button:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
    }
    
    .booking-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 3rem 2rem;
        border-radius: 20px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 40px rgba(102, 126, 234, 0.3);
        text-align: center;
    }
    
    .booking-header h1 {
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .status-new { 
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white; padding: 8px 16px; border-radius: 20px; 
        font-weight: 600; display: inline-block;
        box-shadow: 0 4px 15px rgba(17, 153, 142, 0.3);
    }
    
    .status-reserved { 
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        color: white; padding: 8px 16px; border-radius: 20px; 
        font-weight: 600; display: inline-block;
        box-shadow: 0 4px 15px rgba(79, 172, 254, 0.3);
    }
    
    .status-canceled { 
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        color: white; padding: 8px 16px; border-radius: 20px; 
        font-weight: 600; display: inline-block;
        box-shadow: 0 4px 15px rgba(245, 87, 108, 0.3);
    }
    
    .status-update { 
        background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
        color: white; padding: 8px 16px; border-radius: 20px; 
        font-weight: 600; display: inline-block;
        box-shadow: 0 4px 15px rgba(250, 112, 154, 0.3);
    }
    
    .metric-card {
        background: linear-gradient(135deg, #e0f7ff 0%, #d6f0ff 100%);
        padding: 1.5rem;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(79, 172, 254, 0.15);
        transition: all 0.3s ease;
        border: 1px solid rgba(79, 172, 254, 0.2);
        color: #333333;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 30px rgba(79, 172, 254, 0.25);
    }
    
    .sidebar-card {
        background: linear-gradient(135deg, #e8f4ff 0%, #dceeff 100%);
        padding: 1rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(79, 172, 254, 0.15);
        border: 1px solid rgba(79, 172, 254, 0.2);
        margin: 0.5rem 0;
        transition: all 0.3s ease;
    }
    
    .sidebar-card:hover {
        transform: translateX(3px);
        box-shadow: 0 6px 20px rgba(79, 172, 254, 0.25);
    }
    
    .sidebar-card h4 {
        color: #667eea;
        margin: 0 0 0.5rem 0;
        font-size: 14px;
        font-weight: 600;
    }
    
    .sidebar-card p {
        color: #333;
        margin: 0;
        font-size: 13px;
    }
    
    .booking-card {
        background: linear-gradient(135deg, #e8f4ff 0%, #dceeff 100%);
        padding: 1.5rem;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(79, 172, 254, 0.15);
        margin-bottom: 1rem;
        border-left: 5px solid #667eea;
        transition: all 0.3s ease;
        color: #333333;
    }
    
    .booking-card:hover {
        transform: translateX(5px);
        box-shadow: 0 8px 30px rgba(79, 172, 254, 0.25);
    }
    
    .editable-card {
        background: linear-gradient(135deg, #fff9e6 0%, #fff3cc 100%);
        padding: 1.5rem;
        border-radius: 15px;
        border-left: 5px solid #fee140;
        margin-bottom: 1rem;
    }
    
    .live-indicator {
        display: inline-block;
        width: 12px;
        height: 12px;
        background: #38ef7d;
        border-radius: 50%;
        animation: pulse 2s infinite;
        margin-right: 8px;
    }
    
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
    
    .log-entry {
        background: #f8f9fa;
        padding: 12px;
        border-left: 4px solid #667eea;
        border-radius: 8px;
        margin: 8px 0;
        font-size: 13px;
    }
    
    .log-entry-success { border-left-color: #38ef7d; }
    .log-entry-error { border-left-color: #f5576c; }
    .log-entry-info { border-left-color: #4facfe; }
    
    .url-box {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 10px;
        border: 2px solid #667eea;
        word-break: break-all;
        font-family: 'Courier New', monospace;
        font-size: 12px;
    }
    
    .success-message {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        font-weight: 600;
        animation: slideIn 0.5s ease;
    }
    
    .error-message {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        font-weight: 600;
        animation: slideIn 0.5s ease;
    }
    
    @keyframes slideIn {
        from { transform: translateY(-20px); opacity: 0; }
        to { transform: translateY(0); opacity: 1; }
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'client' not in st.session_state:
    st.session_state.client = None
if 'worksheet' not in st.session_state:
    st.session_state.worksheet = None
if 'credentials' not in st.session_state:
    st.session_state.credentials = None
if 'activity_log' not in st.session_state:
    st.session_state.activity_log = []
if 'current_spreadsheet_id' not in st.session_state:
    st.session_state.current_spreadsheet_id = None
if 'spreadsheet_url' not in st.session_state:
    st.session_state.spreadsheet_url = None
if 'email_notifications_enabled' not in st.session_state:
    st.session_state.email_notifications_enabled = True
if 'available_spreadsheets' not in st.session_state:
    st.session_state.available_spreadsheets = []
if 'auto_loaded_sheets' not in st.session_state:
    st.session_state.auto_loaded_sheets = False

# Constants
ORIGINAL_SPREADSHEET_ID = "1-3FLLEkUmiHzW7DGVAPI6PebdRc_24t3vM0OCBnDhco"
ORIGINAL_SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{ORIGINAL_SPREADSHEET_ID}/edit"
NOTIFICATION_EMAIL = "entremotivator@gmail.com"
SHEET_NAME = "Copy of Test Template Sheet Reservations"

def log_activity(message: str, log_type: str = "info"):
    """Log activity with timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {
        'timestamp': timestamp,
        'message': message,
        'type': log_type
    }
    st.session_state.activity_log.insert(0, log_entry)
    if len(st.session_state.activity_log) > 100:
        st.session_state.activity_log = st.session_state.activity_log[:100]

def send_email_notification(subject: str, body: str):
    """Log email notification (actual sending would require SMTP configuration)"""
    log_activity(f"📧 Email notification: {subject}", "info")
    # Note: To actually send emails, you would need to configure Gmail App Password
    # and use smtplib with proper authentication

def authenticate_google_sheets(creds_dict: Dict[str, Any]) -> Optional[gspread.Client]:
    """Authenticate with Google Sheets"""
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(creds)
        log_activity("Successfully authenticated with Google Sheets", "success")
        
        # Auto-load default spreadsheet
        try:
            default_worksheet = get_worksheet(client, ORIGINAL_SPREADSHEET_ID)
            if default_worksheet:
                st.session_state.worksheet = default_worksheet
                st.session_state.current_spreadsheet_id = ORIGINAL_SPREADSHEET_ID
                st.session_state.spreadsheet_url = ORIGINAL_SPREADSHEET_URL
                st.session_state.default_sheet_loaded = True
                log_activity("Default spreadsheet loaded successfully", "success")
        except Exception as e:
            log_activity(f"Could not auto-load default sheet: {str(e)}", "error")
        
        return client
    except Exception as e:
        log_activity(f"Authentication failed: {str(e)}", "error")
        st.error(f"Authentication failed: {str(e)}")
        return None

def add_to_drive(client: gspread.Client, spreadsheet_id: str) -> bool:
    """Add a spreadsheet to user's Drive by creating a copy"""
    try:
        source = client.open_by_key(spreadsheet_id)
        new_title = f"{source.title} - Added {datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # Create a copy which automatically adds to Drive
        new_spreadsheet = client.copy(spreadsheet_id, title=new_title, copy_permissions=False)
        
        # Share with notification email
        try:
            new_spreadsheet.share(NOTIFICATION_EMAIL, perm_type='user', role='writer')
            log_activity(f"Added to Drive and shared: {new_title}", "success")
        except Exception as e:
            log_activity(f"Added to Drive but failed to share: {str(e)}", "error")
        
        return True
    except Exception as e:
        log_activity(f"Failed to add to Drive: {str(e)}", "error")
        return False

def get_all_spreadsheets(client: gspread.Client) -> List[Dict[str, str]]:
    """Get all spreadsheets from Google Drive"""
    try:
        spreadsheets = client.openall()
        return [{'id': s.id, 'title': s.title, 'url': s.url} for s in spreadsheets]
    except Exception as e:
        log_activity(f"Failed to fetch spreadsheets: {str(e)}", "error")
        return []

def clone_spreadsheet(client: gspread.Client, source_id: str, new_title: str = None) -> Optional[Dict[str, str]]:
    """Clone a spreadsheet and share with notification email"""
    try:
        source = client.open_by_key(source_id)
        
        if new_title is None:
            new_title = f"Copy of {source.title} - {datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        new_spreadsheet = client.copy(source_id, title=new_title, copy_permissions=False)
        
        try:
            new_spreadsheet.share(NOTIFICATION_EMAIL, perm_type='user', role='writer')
            log_activity(f"Cloned and shared spreadsheet: {new_title}", "success")
            send_email_notification(
                f"New Spreadsheet Created: {new_title}",
                f"A new spreadsheet has been created and shared with you.\n\nView it here: {new_spreadsheet.url}"
            )
        except Exception as e:
            log_activity(f"Cloned but failed to share: {str(e)}", "error")
        
        return {
            'id': new_spreadsheet.id,
            'title': new_spreadsheet.title,
            'url': new_spreadsheet.url
        }
    except Exception as e:
        log_activity(f"Failed to clone spreadsheet: {str(e)}", "error")
        return None

def get_worksheet(client: gspread.Client, spreadsheet_id: str) -> Optional[gspread.Worksheet]:
    """Get the specific worksheet"""
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.get_worksheet(0)
        log_activity(f"Opened worksheet: {spreadsheet.title}", "info")
        return worksheet
    except Exception as e:
        log_activity(f"Error accessing worksheet: {str(e)}", "error")
        return None

def read_bookings(worksheet: gspread.Worksheet) -> pd.DataFrame:
    """Read all bookings from the sheet"""
    try:
        all_values = worksheet.get_all_values()
        
        if len(all_values) < 7:
            return pd.DataFrame()
        
        header_row_idx = None
        for idx, row in enumerate(all_values):
            if 'DATE:' in row:
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            return pd.DataFrame()
        
        headers = all_values[header_row_idx]
        data = all_values[header_row_idx + 1:]
        df = pd.DataFrame(data, columns=headers)
        
        if 'DATE:' in df.columns:
            df = df[df['DATE:'].notna() & (df['DATE:'] != '')]
        
        return df
    except Exception as e:
        log_activity(f"Error reading bookings: {str(e)}", "error")
        return pd.DataFrame()

def add_booking(worksheet: gspread.Worksheet, booking_data: Dict[str, Any]) -> bool:
    """Add a new booking to the sheet"""
    try:
        all_values = worksheet.get_all_values()
        header_row_idx = None
        
        for idx, row in enumerate(all_values):
            if 'DATE:' in row:
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            return False
        
        row_data = [
            booking_data['date'],
            booking_data['villa'],
            booking_data['type_clean'],
            booking_data['pax'],
            booking_data['start_time'],
            booking_data['end_time'],
            booking_data['reservation_status'],
            booking_data['laundry_services'],
            booking_data['comments']
        ]
        
        worksheet.append_row(row_data, value_input_option='USER_ENTERED')
        
        log_activity(f"New booking added: {booking_data['villa']} on {booking_data['date']}", "success")
        
        if st.session_state.email_notifications_enabled:
            send_email_notification(
                f"New Booking: {booking_data['villa']}",
                f"New booking created:\nVilla: {booking_data['villa']}\nDate: {booking_data['date']}\nTime: {booking_data['start_time']} - {booking_data['end_time']}\nStatus: {booking_data['reservation_status']}"
            )
        
        return True
    except Exception as e:
        log_activity(f"Error adding booking: {str(e)}", "error")
        return False

def update_booking(worksheet: gspread.Worksheet, row_index: int, booking_data: Dict[str, Any]) -> bool:
    """Update an existing booking"""
    try:
        row_data = [
            booking_data['date'],
            booking_data['villa'],
            booking_data['type_clean'],
            booking_data['pax'],
            booking_data['start_time'],
            booking_data['end_time'],
            booking_data['reservation_status'],
            booking_data['laundry_services'],
            booking_data['comments']
        ]
        
        actual_row = row_index + 7
        for col_idx, value in enumerate(row_data, start=1):
            worksheet.update_cell(actual_row, col_idx, value)
        
        log_activity(f"Booking updated: {booking_data['villa']}", "success")
        
        if st.session_state.email_notifications_enabled:
            send_email_notification(
                f"Booking Updated: {booking_data['villa']}",
                f"Booking updated:\nVilla: {booking_data['villa']}\nNew Status: {booking_data['reservation_status']}"
            )
        
        return True
    except Exception as e:
        log_activity(f"Error updating booking: {str(e)}", "error")
        return False

def delete_booking(worksheet: gspread.Worksheet, row_index: int, villa_name: str) -> bool:
    """Delete a booking"""
    try:
        worksheet.delete_rows(row_index + 7)
        log_activity(f"Booking deleted: {villa_name}", "success")
        
        if st.session_state.email_notifications_enabled:
            send_email_notification(
                f"Booking Deleted: {villa_name}",
                f"A booking for {villa_name} has been deleted."
            )
        
        return True
    except Exception as e:
        log_activity(f"Error deleting booking: {str(e)}", "error")
        return False

def get_booking_statistics(df: pd.DataFrame) -> Dict[str, int]:
    """Calculate booking statistics"""
    stats = {
        'total': len(df),
        'new': len(df[df['RESERVATION STATUS:'] == 'NEW!']),
        'reserved': len(df[df['RESERVATION STATUS:'] == 'RESERVED!']),
        'canceled': len(df[df['RESERVATION STATUS:'] == 'CANCELED!']),
        'update': len(df[df['RESERVATION STATUS:'] == 'UPDATE!'])
    }
    return stats

def create_status_chart(stats: Dict[str, int]) -> go.Figure:
    """Create a pie chart for booking status"""
    labels = ['New', 'Reserved', 'Canceled', 'Update']
    values = [stats['new'], stats['reserved'], stats['canceled'], stats['update']]
    colors = ['#38ef7d', '#00f2fe', '#f5576c', '#fee140']
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        marker=dict(colors=colors),
        hole=0.4,
        textinfo='label+percent',
        textfont=dict(size=14, color='white'),
        hoverinfo='label+value'
    )])
    
    fig.update_layout(
        title='Booking Status Distribution',
        height=400,
        showlegend=True,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    
    return fig

def render_login_sidebar():
    """Render authentication section"""
    st.sidebar.markdown("## 🔐 Authentication")
    
    if not st.session_state.authenticated:
        st.sidebar.markdown("""
        <div class="sidebar-card">
            <h4 style="text-align: center;">📁 Upload Credentials</h4>
            <p style="text-align: center; font-size: 12px;">Upload your Google Service Account JSON</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.sidebar.file_uploader("Choose JSON credentials file", type=['json'])
        
        if uploaded_file is not None:
            try:
                creds_dict = json.load(uploaded_file)
                
                with st.spinner("Authenticating..."):
                    client = authenticate_google_sheets(creds_dict)
                    
                    if client:
                        st.session_state.client = client
                        st.session_state.credentials = creds_dict
                        st.session_state.authenticated = True
                        st.sidebar.success("Authentication successful!")
                        st.rerun()
            except Exception as e:
                st.sidebar.error(f"Error: {str(e)}")
    else:
        st.sidebar.markdown("""
        <div class="sidebar-card">
            <h4>✅ Connected</h4>
            <p>Google Sheets API Active</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.credentials:
            email = st.session_state.credentials.get('client_email', 'N/A')
            st.sidebar.markdown(f"""
            <div class="sidebar-card">
                <h4>📧 Service Account</h4>
                <p style="font-size: 11px; word-break: break-all;">{email[:35]}...</p>
            </div>
            """, unsafe_allow_html=True)
        
        if st.sidebar.button("🚪 Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.client = None
            st.session_state.worksheet = None
            st.session_state.credentials = None
            st.session_state.current_spreadsheet_id = None
            st.session_state.available_spreadsheets = []
            st.session_state.auto_loaded_sheets = False
            st.rerun()

def render_spreadsheet_selector():
    """Render spreadsheet management with improved dropdown"""
    st.sidebar.markdown("---")
    st.sidebar.markdown("## 📊 Spreadsheet Manager")
    
    if st.session_state.authenticated and st.session_state.client:
        
        # Show default sheet status
        if st.session_state.default_sheet_loaded:
            st.sidebar.markdown("""
            <div class="sidebar-card" style="background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%); border-left: 4px solid #38ef7d;">
                <h4 style="color: #155724;">✅ Default Sheet Loaded</h4>
                <p style="color: #155724; font-size: 12px;">Original template is ready to use</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Auto-load spreadsheets on first authentication
        if not st.session_state.auto_loaded_sheets:
            with st.spinner("Loading your spreadsheets..."):
                spreadsheets = get_all_spreadsheets(st.session_state.client)
                st.session_state.available_spreadsheets = spreadsheets
                st.session_state.auto_loaded_sheets = True
        
        # Original template link with Add to Drive button
        st.sidebar.markdown("### 📄 Original Template")
        st.sidebar.markdown(f"""
        <div class="sidebar-card">
            <a href="{ORIGINAL_SPREADSHEET_URL}" target="_blank" style="color: #667eea; text-decoration: none; font-weight: 600;">
                🔗 View Original Template
            </a>
            <p style="margin-top: 0.5rem; font-size: 11px;">Base template for all bookings</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.sidebar.columns(2)
        with col1:
            if st.button("➕ Add to Drive", use_container_width=True, help="Copy to your Google Drive"):
                with st.spinner("Adding to Drive..."):
                    if add_to_drive(st.session_state.client, ORIGINAL_SPREADSHEET_ID):
                        st.sidebar.success("✅ Added to Drive!")
                        # Refresh spreadsheet list
                        spreadsheets = get_all_spreadsheets(st.session_state.client)
                        st.session_state.available_spreadsheets = spreadsheets
                        st.rerun()
                    else:
                        st.sidebar.error("Failed to add")
        
        with col2:
            if st.button("🔄 Use Default", use_container_width=True, help="Load original template"):
                with st.spinner("Loading..."):
                    worksheet = get_worksheet(st.session_state.client, ORIGINAL_SPREADSHEET_ID)
                    if worksheet:
                        st.session_state.worksheet = worksheet
                        st.session_state.current_spreadsheet_id = ORIGINAL_SPREADSHEET_ID
                        st.session_state.spreadsheet_url = ORIGINAL_SPREADSHEET_URL
                        st.session_state.default_sheet_loaded = True
                        log_activity("Default spreadsheet loaded", "success")
                        st.rerun()
        
        # Clone functionality with improved UI
        st.sidebar.markdown("### 📋 Clone Template")
        
        with st.sidebar.form("clone_form"):
            clone_name = st.text_input(
                "New spreadsheet name",
                placeholder="My Villa Bookings 2025",
                help="Leave empty for auto-generated name"
            )
            
            clone_btn = st.form_submit_button("🔄 Clone & Auto-Share", use_container_width=True)
            
            if clone_btn:
                with st.spinner("Cloning and sharing..."):
                    result = clone_spreadsheet(
                        st.session_state.client,
                        ORIGINAL_SPREADSHEET_ID,
                        clone_name if clone_name else None
                    )
                    
                    if result:
                        st.success(f"✅ Created: {result['title'][:30]}...")
                        # Refresh spreadsheet list
                        spreadsheets = get_all_spreadsheets(st.session_state.client)
                        st.session_state.available_spreadsheets = spreadsheets
                        st.session_state.current_spreadsheet_id = result['id']
                        st.session_state.spreadsheet_url = result['url']
                        
                        # Auto-select the new worksheet
                        worksheet = get_worksheet(st.session_state.client, result['id'])
                        if worksheet:
                            st.session_state.worksheet = worksheet
                        
                        st.rerun()
                    else:
                        st.error("Failed to clone")
        
        # Enhanced dropdown selector
        st.sidebar.markdown("### 📚 Your Spreadsheets")
        
        # Refresh button
        col1, col2 = st.sidebar.columns([3, 1])
        with col1:
            st.markdown('<div class="sidebar-card"><p>Select from all sheets in Drive</p></div>', unsafe_allow_html=True)
        with col2:
            if st.button("🔄", help="Refresh list"):
                with st.spinner("Refreshing..."):
                    spreadsheets = get_all_spreadsheets(st.session_state.client)
                    st.session_state.available_spreadsheets = spreadsheets
                    st.rerun()
        
        # Display count
        if st.session_state.available_spreadsheets:
            # Filter for cloned/copied spreadsheets
            all_sheets = st.session_state.available_spreadsheets
            cloned_sheets = [s for s in all_sheets if 'Copy of' in s['title'] or 'Villa' in s['title'] or 'Added' in s['title']]
            
            st.sidebar.markdown(f"""
            <div class="sidebar-card">
                <h4>📊 Sheet Statistics</h4>
                <p>Total sheets: <strong>{len(all_sheets)}</strong></p>
                <p>Cloned/Added: <strong>{len(cloned_sheets)}</strong></p>
            </div>
            """, unsafe_allow_html=True)
            
            # Tabs for filtering
            sheet_filter = st.sidebar.radio(
                "Filter by:",
                ["🌟 All Sheets", "📋 Cloned/Added Only", "🔍 Search"],
                label_visibility="collapsed"
            )
            
            # Determine which sheets to show
            if sheet_filter == "📋 Cloned/Added Only":
                display_sheets = cloned_sheets
            elif sheet_filter == "🔍 Search":
                search_term = st.sidebar.text_input("🔍 Search sheets", "")
                if search_term:
                    display_sheets = [s for s in all_sheets if search_term.lower() in s['title'].lower()]
                else:
                    display_sheets = all_sheets
            else:
                display_sheets = all_sheets
            
            if display_sheets:
                # Create dropdown options
                spreadsheet_options = {
                    f"📄 {s['title'][:50]}{'...' if len(s['title']) > 50 else ''}": s['id'] 
                    for s in display_sheets
                }
                
                # Find current selection
                current_selection = None
                if st.session_state.current_spreadsheet_id:
                    for title, sheet_id in spreadsheet_options.items():
                        if sheet_id == st.session_state.current_spreadsheet_id:
                            current_selection = title
                            break
                
                selected_title = st.sidebar.selectbox(
                    f"Select spreadsheet ({len(display_sheets)} available)",
                    options=list(spreadsheet_options.keys()),
                    index=list(spreadsheet_options.keys()).index(current_selection) if current_selection else 0
                )
                
                if st.sidebar.button("📂 Open Selected Sheet", use_container_width=True):
                    selected_id = spreadsheet_options[selected_title]
                    selected_sheet = next(s for s in display_sheets if s['id'] == selected_id)
                    
                    with st.spinner("Opening spreadsheet..."):
                        worksheet = get_worksheet(st.session_state.client, selected_id)
                        
                        if worksheet:
                            st.session_state.worksheet = worksheet
                            st.session_state.current_spreadsheet_id = selected_id
                            st.session_state.spreadsheet_url = selected_sheet['url']
                            st.sidebar.success(f"✅ Opened: {selected_sheet['title'][:30]}...")
                            st.rerun()
                        else:
                            st.sidebar.error("Failed to open sheet")
            else:
                st.sidebar.info("No sheets found matching criteria")
        else:
            st.sidebar.markdown("""
            <div class="sidebar-card">
                <h4>⚠️ No Sheets Loaded</h4>
                <p>Click the refresh button above to load your sheets</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Current spreadsheet info with enhanced styling
        if st.session_state.current_spreadsheet_id and st.session_state.spreadsheet_url:
            st.sidebar.markdown("---")
            st.sidebar.markdown("### <span class='live-indicator'></span> Active Sheet", unsafe_allow_html=True)
            
            # Get current sheet name
            current_sheet_name = "Unknown"
            if st.session_state.available_spreadsheets:
                for sheet in st.session_state.available_spreadsheets:
                    if sheet['id'] == st.session_state.current_spreadsheet_id:
                        current_sheet_name = sheet['title']
                        break
            
            # If it's the default sheet
            if st.session_state.current_spreadsheet_id == ORIGINAL_SPREADSHEET_ID:
                current_sheet_name = "Original Template (Default)"
            
            st.sidebar.markdown(f"""
            <div class="sidebar-card">
                <h4>📊 {current_sheet_name[:30]}{'...' if len(current_sheet_name) > 30 else ''}</h4>
                <a href="{st.session_state.spreadsheet_url}" target="_blank" style="color: #667eea; font-weight: 600; text-decoration: none;">
                    🔗 View Live Spreadsheet →
                </a>
                <p style="margin-top: 0.5rem; font-size: 11px; color: #666;">Opens in new tab</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Quick actions
            st.sidebar.markdown("""
            <div class="sidebar-card">
                <h4>⚡ Quick Actions</h4>
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.sidebar.columns(2)
            with col1:
                if st.button("📋 Copy URL", use_container_width=True):
                    st.sidebar.code(st.session_state.spreadsheet_url, language="text")
            with col2:
                if st.button("🔄 Reload", use_container_width=True):
                    st.rerun()

def render_activity_log_sidebar():
    """Render activity log with enhanced styling"""
    st.sidebar.markdown("---")
    st.sidebar.markdown("## 📋 Activity & Notifications")
    
    # Email notifications toggle with card styling
    st.sidebar.markdown("""
    <div class="sidebar-card">
        <h4>📧 Email Notifications</h4>
        <p style="font-size: 11px;">Auto-notify: entremotivator@gmail.com</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.session_state.email_notifications_enabled = st.sidebar.checkbox(
        "Enable Email Alerts",
        value=st.session_state.email_notifications_enabled,
        help="Send notifications for all booking activities"
    )
    
    # Activity log stats
    log_count_total = len(st.session_state.activity_log)
    log_success = len([l for l in st.session_state.activity_log if l['type'] == 'success'])
    log_errors = len([l for l in st.session_state.activity_log if l['type'] == 'error'])
    
    st.sidebar.markdown(f"""
    <div class="sidebar-card">
        <h4>📊 Activity Statistics</h4>
        <p>Total Events: <strong style="color: #667eea;">{log_count_total}</strong></p>
        <p>✅ Success: <strong style="color: #38ef7d;">{log_success}</strong></p>
        <p>❌ Errors: <strong style="color: #f5576c;">{log_errors}</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Log display controls
    col1, col2 = st.sidebar.columns(2)
    with col1:
        if st.button("🔄 Refresh", use_container_width=True, key="refresh_log"):
            st.rerun()
    with col2:
        if st.button("🗑️ Clear", use_container_width=True, key="clear_log"):
            st.session_state.activity_log = []
            st.rerun()
    
    log_count = st.sidebar.slider("Show entries", 5, 50, 15, key="log_slider")
    
    # Display logs with enhanced styling
    st.sidebar.markdown("### Recent Activity")
    
    if st.session_state.activity_log:
        for log in st.session_state.activity_log[:log_count]:
            log_class = f"log-entry log-entry-{log['type']}"
            
            # Icon based on type
            icon = "ℹ️"
            if log['type'] == 'success':
                icon = "✅"
            elif log['type'] == 'error':
                icon = "❌"
            elif log['type'] == 'info':
                icon = "📌"
            
            st.sidebar.markdown(f"""
            <div class="{log_class}">
                <strong style="font-size: 11px;">{icon} {log['timestamp']}</strong><br>
                <span style="font-size: 12px;">{log['message']}</span>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.sidebar.markdown("""
        <div class="sidebar-card">
            <p style="text-align: center; color: #666;">No activity yet</p>
        </div>
        """, unsafe_allow_html=True)

def main():
    render_login_sidebar()
    
    if not st.session_state.authenticated:
        st.markdown("""
            <div class="booking-header">
                <h1>🏖️ Villa Booking Management System Pro</h1>
                <p>Advanced booking management with real-time sync & notifications</p>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.info("""
            ### Welcome!
            
            Upload your Google Service Account credentials to get started.
            
            **Enhanced Features:**
            - 📝 Create & edit bookings
            - 📧 Email notifications to entremotivator@gmail.com
            - 📊 Real-time analytics
            - 🔄 Clone spreadsheets
            - 📚 Access all your sheets
            - 📋 Activity logging
            - 🔗 Direct sheet links
            """)
        return
    
    render_spreadsheet_selector()
    render_activity_log_sidebar()
    
    if not st.session_state.worksheet:
        st.warning("Please select or create a spreadsheet from the sidebar")
        return
    
    worksheet = st.session_state.worksheet
    
    st.markdown("""
        <div class="booking-header">
            <h1>🏖️ Villa Booking Management System Pro</h1>
            <p>Manage bookings with real-time synchronization</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Navigation
    st.sidebar.markdown("---")
    st.sidebar.title("📋 Navigation")
    page = st.sidebar.radio("", [
        "📊 Dashboard",
        "📝 New Booking",
        "📅 View & Edit Bookings",
        "📈 Analytics",
        "⚙️ Settings"
    ], label_visibility="collapsed")
    
    # DASHBOARD
    if page == "📊 Dashboard":
        st.header("📊 Dashboard")
        
        with st.spinner("Loading..."):
            df = read_bookings(worksheet)
        
        if not df.empty:
            stats = get_booking_statistics(df)
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            metrics = [
                ("📊 Total", stats['total'], "#667eea"),
                ("🆕 New", stats['new'], "#38ef7d"),
                ("✅ Reserved", stats['reserved'], "#00f2fe"),
                ("❌ Canceled", stats['canceled'], "#f5576c"),
                ("🔄 Update", stats['update'], "#fee140")
            ]
            
            for col, (label, value, color) in zip([col1, col2, col3, col4, col5], metrics):
                with col:
                    st.markdown(f"""
                    <div class="metric-card">
                        <h3 style="color: {color}; margin: 0;">{label}</h3>
                        <h1 style="margin: 10px 0; color: #333;">{value}</h1>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig = create_status_chart(stats)
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.subheader("📅 Recent Bookings")
                recent_df = df.head(5)
                for idx, row in recent_df.iterrows():
                    st.markdown(f"""
                    <div class="booking-card">
                        <strong>🏠 {row['VILLA:']}</strong><br>
                        <span>📅 {row['DATE:']} | ⏰ {row['START TIME:']} - {row['END TIME:']}</span><br>
                        <span class="status-{row['RESERVATION STATUS:'].lower().replace('!', '')}">{row['RESERVATION STATUS:']}</span>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("No bookings found. Create your first booking!")
    
    # NEW BOOKING
    elif page == "📝 New Booking":
        st.header("📝 Create New Booking")
        
        with st.form("booking_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📍 Basic Information")
                date = st.date_input("📅 Booking Date", datetime.now())
                villa = st.text_input("🏠 Villa Name", placeholder="Enter villa name")
                type_clean = st.selectbox("🧹 Type of Cleaning", [
                    "Standard Clean", "Deep Clean", "Check-out Clean", 
                    "Mid-stay Clean", "Move-in Clean", "Post-event Clean"
                ])
                pax = st.number_input("👥 PAX", min_value=1, max_value=50, value=2)
            
            with col2:
                st.subheader("⏰ Schedule & Services")
                start_time = st.time_input("⏰ Start Time", time(10, 0))
                end_time = st.time_input("⏰ End Time", time(14, 0))
                reservation_status = st.selectbox("📊 Status", ["NEW!", "RESERVED!", "CANCELED!", "UPDATE!"])
                laundry_services = st.selectbox("🧺 Laundry", ["Yes", "No", "Not Required", "Pending"])
            
            st.subheader("💬 Additional Information")
            comments = st.text_area("Comments", placeholder="Special notes...", height=100)
            
            col1, col2, col3 = st.columns([1, 1, 1])
            with col2:
                submit_button = st.form_submit_button("✅ Submit Booking", use_container_width=True)
            
            if submit_button:
                if not villa:
                    st.error("Please enter a villa name")
                elif start_time >= end_time:
                    st.error("End time must be after start time")
                else:
                    booking_data = {
                        'date': date.strftime('%m/%d/%Y'),
                        'villa': villa,
                        'type_clean': type_clean,
                        'pax': str(pax),
                        'start_time': start_time.strftime('%H:%M'),
                        'end_time': end_time.strftime('%H:%M'),
                        'reservation_status': reservation_status,
                        'laundry_services': laundry_services,
                        'comments': comments
                    }
                    
                    with st.spinner("Saving..."):
                        if add_booking(worksheet, booking_data):
                            st.markdown('<div class="success-message">✅ Booking added successfully!</div>', unsafe_allow_html=True)
                            st.balloons()
                        else:
                            st.markdown('<div class="error-message">❌ Failed to add booking</div>', unsafe_allow_html=True)
    
    # VIEW & EDIT BOOKINGS
    elif page == "📅 View & Edit Bookings":
        st.header("📅 View & Edit Bookings")
        
        with st.expander("🔍 Filters", expanded=True):
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                filter_status = st.selectbox("Status", ["All", "NEW!", "RESERVED!", "CANCELED!", "UPDATE!"])
            with col2:
                filter_date_from = st.date_input("From Date", value=None)
            with col3:
                filter_date_to = st.date_input("To Date", value=None)
            with col4:
                search_villa = st.text_input("🔍 Search Villa")
        
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("🔄 Refresh", use_container_width=True):
                st.rerun()
        
        with st.spinner("Loading bookings..."):
            df = read_bookings(worksheet)
        
        if not df.empty:
            filtered_df = df.copy()
            
            if filter_status != "All":
                filtered_df = filtered_df[filtered_df['RESERVATION STATUS:'] == filter_status]
            if filter_date_from:
                filtered_df = filtered_df[pd.to_datetime(filtered_df['DATE:']) >= pd.to_datetime(filter_date_from)]
            if filter_date_to:
                filtered_df = filtered_df[pd.to_datetime(filtered_df['DATE:']) <= pd.to_datetime(filter_date_to)]
            if search_villa:
                filtered_df = filtered_df[filtered_df['VILLA:'].str.contains(search_villa, case=False, na=False)]
            
            st.info(f"Showing {len(filtered_df)} of {len(df)} bookings")
            
            for idx, row in filtered_df.iterrows():
                with st.expander(f"🏠 {row['VILLA:']} - {row['DATE:']} ({row['RESERVATION STATUS:']})"):
                    edit_mode = st.checkbox(f"✏️ Edit Mode", key=f"edit_{idx}")
                    
                    if edit_mode:
                        st.markdown('<div class="editable-card">', unsafe_allow_html=True)
                        
                        with st.form(key=f"edit_form_{idx}"):
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                new_date = st.date_input("Date", value=datetime.strptime(row['DATE:'], '%m/%d/%Y'))
                                new_villa = st.text_input("Villa", value=row['VILLA:'])
                                new_type_clean = st.selectbox("Cleaning Type", [
                                    "Standard Clean", "Deep Clean", "Check-out Clean", 
                                    "Mid-stay Clean", "Move-in Clean", "Post-event Clean"
                                ], index=0)
                                new_pax = st.number_input("PAX", min_value=1, value=int(row['PAX:']) if row['PAX:'].isdigit() else 2)
                            
                            with col2:
                                new_start_time = st.time_input("Start Time", value=datetime.strptime(row['START TIME:'], '%H:%M').time() if ':' in row['START TIME:'] else time(10, 0))
                                new_end_time = st.time_input("End Time", value=datetime.strptime(row['END TIME:'], '%H:%M').time() if ':' in row['END TIME:'] else time(14, 0))
                                new_status = st.selectbox("Status", ["NEW!", "RESERVED!", "CANCELED!", "UPDATE!"], index=["NEW!", "RESERVED!", "CANCELED!", "UPDATE!"].index(row['RESERVATION STATUS:']) if row['RESERVATION STATUS:'] in ["NEW!", "RESERVED!", "CANCELED!", "UPDATE!"] else 0)
                                new_laundry = st.selectbox("Laundry", ["Yes", "No", "Not Required", "Pending"])
                            
                            new_comments = st.text_area("Comments", value=row['COMMENTS:'])
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                update_btn = st.form_submit_button("💾 Update", use_container_width=True)
                            with col3:
                                delete_btn = st.form_submit_button("🗑️ Delete", use_container_width=True)
                            
                            if update_btn:
                                booking_data = {
                                    'date': new_date.strftime('%m/%d/%Y'),
                                    'villa': new_villa,
                                    'type_clean': new_type_clean,
                                    'pax': str(new_pax),
                                    'start_time': new_start_time.strftime('%H:%M'),
                                    'end_time': new_end_time.strftime('%H:%M'),
                                    'reservation_status': new_status,
                                    'laundry_services': new_laundry,
                                    'comments': new_comments
                                }
                                
                                if update_booking(worksheet, idx, booking_data):
                                    st.success("Booking updated!")
                                    st.rerun()
                            
                            if delete_btn:
                                if delete_booking(worksheet, idx, row['VILLA:']):
                                    st.success("Booking deleted!")
                                    st.rerun()
                        
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.markdown("**📅 Date & Time**")
                            st.write(f"Date: {row['DATE:']}")
                            st.write(f"Start: {row['START TIME:']}")
                            st.write(f"End: {row['END TIME:']}")
                        
                        with col2:
                            st.markdown("**🏠 Details**")
                            st.write(f"Villa: {row['VILLA:']}")
                            st.write(f"Cleaning: {row['TYPE CLEAN:']}")
                            st.write(f"PAX: {row['PAX:']}")
                        
                        with col3:
                            st.markdown("**📊 Status**")
                            status_class = f"status-{row['RESERVATION STATUS:'].lower().replace('!', '')}"
                            st.markdown(f"<span class='{status_class}'>{row['RESERVATION STATUS:']}</span>", unsafe_allow_html=True)
                            st.write(f"Laundry: {row['LAUNDRY SERVICES WITH VIDEMI:']}")
                        
                        if row['COMMENTS:']:
                            st.markdown("**💬 Comments**")
                            st.info(row['COMMENTS:'])
        else:
            st.info("No bookings found")
    
    # ANALYTICS
    elif page == "📈 Analytics":
        st.header("📈 Analytics")
        
        with st.spinner("Analyzing..."):
            df = read_bookings(worksheet)
        
        if not df.empty:
            stats = get_booking_statistics(df)
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig_status = create_status_chart(stats)
                st.plotly_chart(fig_status, use_container_width=True)
            
            with col2:
                cleaning_counts = df['TYPE CLEAN:'].value_counts()
                fig_cleaning = px.bar(
                    x=cleaning_counts.index,
                    y=cleaning_counts.values,
                    labels={'x': 'Cleaning Type', 'y': 'Count'},
                    title='Bookings by Cleaning Type',
                    color=cleaning_counts.values,
                    color_continuous_scale='Viridis'
                )
                st.plotly_chart(fig_cleaning, use_container_width=True)
            
            st.subheader("🏆 Top Villas")
            villa_counts = df['VILLA:'].value_counts().head(10)
            fig_villas = px.bar(
                x=villa_counts.values,
                y=villa_counts.index,
                orientation='h',
                title='Top 10 Most Booked Villas',
                color=villa_counts.values,
                color_continuous_scale='Sunset'
            )
            st.plotly_chart(fig_villas, use_container_width=True)
            
            st.markdown("---")
            st.subheader("📥 Export Data")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "⬇️ Download CSV",
                    data=csv,
                    file_name=f"bookings_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            with col2:
                json_str = df.to_json(orient='records', indent=2)
                st.download_button(
                    "⬇️ Download JSON",
                    data=json_str,
                    file_name=f"bookings_{datetime.now().strftime('%Y%m%d')}.json",
                    mime="application/json",
                    use_container_width=True
                )
        else:
            st.info("No data available")
    
    # SETTINGS
    elif page == "⚙️ Settings":
        st.header("⚙️ Settings")
        
        st.subheader("📧 Notification Settings")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <h4>Current Email</h4>
                <p style="font-size: 18px; color: #667eea; font-weight: 600;">{NOTIFICATION_EMAIL}</p>
                <p style="font-size: 14px; color: #666;">All notifications are sent to this email</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <h4>Notification Status</h4>
                <p style="font-size: 18px; color: {'#38ef7d' if st.session_state.email_notifications_enabled else '#f5576c'}; font-weight: 600;">
                    {'✅ Enabled' if st.session_state.email_notifications_enabled else '❌ Disabled'}
                </p>
                <p style="font-size: 14px; color: #666;">Toggle in sidebar activity log</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.subheader("📊 System Information")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div class="metric-card">
                <h4>Original Template</h4>
                <code style="font-size: 11px; word-break: break-all;">1-3FLLEkUm...</code>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="metric-card">
                <h4>Status</h4>
                <p style="color: #38ef7d; font-weight: 600;">● Connected</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            log_count = len(st.session_state.activity_log)
            st.markdown(f"""
            <div class="metric-card">
                <h4>Activity Logs</h4>
                <h2 style="color: #667eea;">{log_count}</h2>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.subheader("⚠️ Danger Zone")
        
        with st.expander("🚨 Advanced Operations"):
            st.warning("These operations may affect your data")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🗑️ Clear Activity Log", use_container_width=True):
                    st.session_state.activity_log = []
                    log_activity("Activity log cleared", "info")
                    st.success("Log cleared")
            
            with col2:
                if st.button("🔄 Reset Connection", use_container_width=True):
                    st.session_state.worksheet = None
                    st.session_state.current_spreadsheet_id = None
                    st.rerun()
    
    # Footer
    st.sidebar.markdown("---")
    st.sidebar.caption("🏖️ Villa Booking System Pro v3.0")
    st.sidebar.caption("Built with Streamlit")
    st.sidebar.caption(f"© 2025 | Email: {NOTIFICATION_EMAIL}")

if __name__ == "__main__":
    main()
