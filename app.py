import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, timedelta
import json
from typing import List, Dict, Optional
import re
import logging
from io import StringIO
import calendar as cal

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Page Configuration
st.set_page_config(
    page_title="Professional Booking Management System",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for enhanced styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 2rem;
        padding: 1rem;
    }
    .section-header {
        font-size: 1.8rem;
        font-weight: bold;
        color: #2c3e50;
        margin-top: 2rem;
        margin-bottom: 1.5rem;
        border-bottom: 3px solid #667eea;
        padding-bottom: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .info-box {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #667eea;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .warning-box {
        background: linear-gradient(135deg, #fff9e6 0%, #ffe8b3 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #ffc107;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .success-box {
        background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #28a745;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .error-box {
        background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #f44336;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        text-align: center;
        border: 2px solid #e9ecef;
        transition: transform 0.3s ease;
    }
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
    }
    .sidebar-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        color: white;
        box-shadow: 0 6px 12px rgba(0,0,0,0.2);
    }
    .sidebar-card h3 {
        color: white;
        font-size: 1.3rem;
        margin-bottom: 0.8rem;
        font-weight: bold;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .sidebar-card p {
        color: rgba(255,255,255,0.9);
        font-size: 0.95rem;
        line-height: 1.6;
        margin: 0.5rem 0;
    }
    .sidebar-card .highlight {
        background: rgba(255,255,255,0.2);
        padding: 0.3rem 0.6rem;
        border-radius: 5px;
        font-weight: bold;
        display: inline-block;
        margin: 0.3rem 0;
    }
    .status-badge {
        display: inline-block;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        font-weight: bold;
        margin: 0.3rem;
        color: white;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .log-entry {
        background: #f8f9fa;
        padding: 0.8rem;
        border-left: 3px solid #6c757d;
        margin: 0.5rem 0;
        font-family: 'Courier New', monospace;
        font-size: 0.85rem;
        border-radius: 5px;
    }
    .log-info { border-left-color: #17a2b8; }
    .log-success { border-left-color: #28a745; }
    .log-warning { border-left-color: #ffc107; }
    .log-error { border-left-color: #dc3545; }
    .stButton>button {
        border-radius: 8px;
        font-weight: bold;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .property-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px solid #e9ecef;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    .property-card:hover {
        border-color: #667eea;
        box-shadow: 0 6px 12px rgba(102, 126, 234, 0.3);
    }
    .booking-cell {
        padding: 0.5rem;
        border-radius: 5px;
        text-align: center;
        font-weight: bold;
        margin: 0.2rem;
    }
    .calendar-grid {
        display: grid;
        gap: 0.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'gc' not in st.session_state:
    st.session_state.gc = None
if 'workbooks' not in st.session_state:
    st.session_state.workbooks = []
if 'current_workbook' not in st.session_state:
    st.session_state.current_workbook = None
if 'current_sheet' not in st.session_state:
    st.session_state.current_sheet = None
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'drive_folder_id' not in st.session_state:
    st.session_state.drive_folder_id = None
if 'service_account_email' not in st.session_state:
    st.session_state.service_account_email = None
if 'current_view' not in st.session_state:
    st.session_state.current_view = "Dashboard"
if 'selected_calendar' not in st.session_state:
    st.session_state.selected_calendar = None
if 'filter_status' not in st.session_state:
    st.session_state.filter_status = "All"
if 'filter_property' not in st.session_state:
    st.session_state.filter_property = "All"

# Status code definitions
STATUS_CODES = {
    'CI': {
        'name': 'Check-In',
        'description': 'Complete cleaning and preparation for incoming guests',
        'color': '#4CAF50'
    },
    'SO': {
        'name': 'Stay-over',
        'description': 'Mid-stay cleaning with linen and towel refresh',
        'color': '#2196F3'
    },
    'CO/CI': {
        'name': 'Check-out/Check-in',
        'description': 'Same-day turnover between guests',
        'color': '#FF9800'
    },
    'FU': {
        'name': 'Fresh-up',
        'description': 'Quick refresh for already clean property',
        'color': '#9C27B0'
    },
    'DC': {
        'name': 'Deep Cleaning',
        'description': 'Thorough deep clean of entire property',
        'color': '#F44336'
    },
    'COC': {
        'name': 'Construction Cleaning',
        'description': 'Post-renovation construction cleanup',
        'color': '#795548'
    },
    'CO': {
        'name': 'Check-Out',
        'description': 'Final cleaning after guest departure',
        'color': '#607D8B'
    }
}

def add_log(message: str, level: str = "INFO"):
    """Add a log entry with timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {
        'timestamp': timestamp,
        'level': level,
        'message': message
    }
    st.session_state.logs.append(log_entry)
    
    # Keep only last 100 logs
    if len(st.session_state.logs) > 100:
        st.session_state.logs = st.session_state.logs[-100:]
    
    # Also log to Python logger
    if level == "INFO":
        logger.info(message)
    elif level == "SUCCESS":
        logger.info(f"✓ {message}")
    elif level == "WARNING":
        logger.warning(message)
    elif level == "ERROR":
        logger.error(message)

class BookingManager:
    """Manages Google Sheets operations for booking system"""
    
    def __init__(self, credentials_dict: Dict):
        """Initialize with service account credentials"""
        try:
            add_log("Initializing Booking Manager...", "INFO")
            
            self.scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/drive.file',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
            
            self.creds = Credentials.from_service_account_info(
                credentials_dict, 
                scopes=self.scopes
            )
            
            add_log("Creating Google Sheets client...", "INFO")
            self.gc = gspread.authorize(self.creds)
            
            # Store service account email
            st.session_state.service_account_email = credentials_dict.get('client_email', 'Unknown')
            
            add_log(f"Successfully authenticated as: {st.session_state.service_account_email}", "SUCCESS")
            add_log("Booking Manager initialized successfully", "SUCCESS")
            
        except Exception as e:
            add_log(f"Failed to initialize Booking Manager: {str(e)}", "ERROR")
            raise
    
    def list_workbooks(self, folder_id: Optional[str] = None) -> List[Dict]:
        """List all spreadsheets in Google Drive or specific folder"""
        try:
            add_log(f"Listing workbooks from {'folder: ' + folder_id if folder_id else 'all accessible files'}...", "INFO")
            
            files = self.gc.openall()
            workbooks = []
            
            for file in files:
                try:
                    workbooks.append({
                        'id': file.id,
                        'name': file.title,
                        'url': file.url
                    })
                except Exception as e:
                    add_log(f"Skipped file due to error: {str(e)}", "WARNING")
                    continue
            
            add_log(f"Found {len(workbooks)} workbook(s)", "SUCCESS")
            
            for wb in workbooks:
                add_log(f"  • {wb['name']} (ID: {wb['id'][:20]}...)", "INFO")
            
            return workbooks
            
        except Exception as e:
            add_log(f"Error listing workbooks: {str(e)}", "ERROR")
            return []
    
    def open_workbook(self, workbook_id: str):
        """Open a specific workbook by ID"""
        try:
            add_log(f"Opening workbook ID: {workbook_id[:20]}...", "INFO")
            workbook = self.gc.open_by_key(workbook_id)
            add_log(f"Successfully opened: {workbook.title}", "SUCCESS")
            add_log(f"Workbook contains {len(workbook.worksheets())} sheet(s)", "INFO")
            return workbook
        except Exception as e:
            add_log(f"Error opening workbook: {str(e)}", "ERROR")
            return None
    
    def get_client_profile(self, workbook) -> Dict:
        """Extract client profile from first sheet"""
        try:
            add_log(f"Reading client profile from: {workbook.title}", "INFO")
            sheet = workbook.get_worksheet(0)
            add_log(f"Accessing sheet: {sheet.title}", "INFO")
            
            all_values = sheet.get_all_values()
            add_log(f"Retrieved {len(all_values)} rows from profile sheet", "INFO")
            
            profile = {
                'client_name': all_values[0][0] if len(all_values) > 0 else 'Unknown',
                'check_out_time': all_values[8][1] if len(all_values) > 8 else '',
                'check_in_time': all_values[9][1] if len(all_values) > 9 else '',
                'amenities': all_values[10][1] if len(all_values) > 10 else '',
                'laundry_services': all_values[11][1] if len(all_values) > 11 else '',
                'keys': all_values[12][1] if len(all_values) > 12 else '',
                'codes': all_values[13][1] if len(all_values) > 13 else '',
                'properties': []
            }
            
            add_log(f"Client: {profile['client_name']}", "INFO")
            
            # Extract properties (starting from row 18)
            property_count = 0
            for i in range(17, min(len(all_values), 50)):
                if len(all_values[i]) > 1 and all_values[i][0]:
                    property_count += 1
                    profile['properties'].append({
                        'name': all_values[i][0],
                        'address': all_values[i][1] if len(all_values[i]) > 1 else '',
                        'hours': all_values[i][2] if len(all_values[i]) > 2 else '',
                        'so_hours': all_values[i][3] if len(all_values[i]) > 3 else ''
                    })
            
            add_log(f"Found {property_count} propert(y/ies)", "SUCCESS")
            
            return profile
        except Exception as e:
            add_log(f"Error reading client profile: {str(e)}", "ERROR")
            return {}
    
    def get_calendar_sheets(self, workbook) -> List[Dict]:
        """Get all calendar sheets (excluding first sheet)"""
        try:
            add_log("Retrieving calendar sheets...", "INFO")
            worksheets = workbook.worksheets()
            calendars = []
            
            for i, sheet in enumerate(worksheets[1:], start=1):
                calendars.append({
                    'index': i,
                    'name': sheet.title,
                    'sheet': sheet
                })
                add_log(f"  • Calendar sheet found: {sheet.title}", "INFO")
            
            add_log(f"Total calendar sheets: {len(calendars)}", "SUCCESS")
            return calendars
            
        except Exception as e:
            add_log(f"Error listing calendar sheets: {str(e)}", "ERROR")
            return []
    
    def read_calendar(self, sheet, start_row: int = 13) -> pd.DataFrame:
        """Read booking calendar starting from specified row"""
        try:
            add_log(f"Reading calendar: {sheet.title} (starting from row {start_row})", "INFO")
            all_values = sheet.get_all_values()
            
            if len(all_values) < start_row:
                add_log(f"Sheet has insufficient rows (found {len(all_values)}, need {start_row})", "WARNING")
                return pd.DataFrame()
            
            # Get headers (row before start_row)
            headers = all_values[start_row - 2] if start_row > 1 else all_values[0]
            add_log(f"Headers: {', '.join([h for h in headers[:10] if h])}", "INFO")
            
            # Get data starting from start_row
            data = all_values[start_row - 1:]
            
            df = pd.DataFrame(data, columns=headers)
            
            # Clean empty rows
            initial_rows = len(df)
            df = df[df.apply(lambda row: row.astype(str).str.strip().any(), axis=1)]
            
            add_log(f"Loaded {len(df)} booking rows (removed {initial_rows - len(df)} empty rows)", "SUCCESS")
            
            return df
            
        except Exception as e:
            add_log(f"Error reading calendar: {str(e)}", "ERROR")
            return pd.DataFrame()
    
    def update_booking(self, sheet, row: int, col: int, value: str) -> bool:
        """Update a single booking cell"""
        try:
            add_log(f"Updating cell [{row}, {col}] in {sheet.title} to: {value}", "INFO")
            sheet.update_cell(row, col, value)
            add_log("Cell updated successfully", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error updating booking: {str(e)}", "ERROR")
            return False
    
    def add_booking_row(self, sheet, data: List) -> bool:
        """Add a new booking row"""
        try:
            add_log(f"Adding new booking row to {sheet.title}", "INFO")
            all_values = sheet.get_all_values()
            last_row = len(all_values) + 1
            
            sheet.insert_row(data, last_row)
            add_log(f"Booking added at row {last_row}", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error adding booking: {str(e)}", "ERROR")
            return False
    
    def get_booking_statistics(self, workbook) -> Dict:
        """Calculate booking statistics across all calendars"""
        try:
            add_log("Calculating booking statistics...", "INFO")
            calendars = self.get_calendar_sheets(workbook)
            
            total_bookings = 0
            status_counts = {code: 0 for code in STATUS_CODES.keys()}
            property_counts = {}
            
            for calendar in calendars:
                df = self.read_calendar(calendar['sheet'])
                
                if not df.empty:
                    total_bookings += len(df)
                    
                    # Count status codes
                    for col in df.columns:
                        for status in STATUS_CODES.keys():
                            status_counts[status] += df[col].astype(str).str.contains(status, case=False, na=False).sum()
                    
                    # Count properties (first column usually contains property names)
                    if len(df.columns) > 0:
                        first_col = df.iloc[:, 0]
                        for prop in first_col:
                            prop_str = str(prop).strip()
                            if prop_str:
                                property_counts[prop_str] = property_counts.get(prop_str, 0) + 1
            
            add_log(f"Statistics calculated: {total_bookings} total bookings", "SUCCESS")
            
            return {
                'total_bookings': total_bookings,
                'status_counts': status_counts,
                'property_counts': property_counts,
                'calendar_count': len(calendars)
            }
            
        except Exception as e:
            add_log(f"Error calculating statistics: {str(e)}", "ERROR")
            return {}

def render_sidebar_info_cards():
    """Render informative cards in sidebar"""
    
    # Connection Status Card
    if st.session_state.authenticated:
        st.markdown(f"""
        <div class="sidebar-card">
            <h3>🟢 Connection Active</h3>
            <p>You are successfully connected to Google Sheets API with full access to your booking workbooks.</p>
            <p class="highlight">Service Account</p>
            <p style="font-size: 0.8rem; word-break: break-all;">{st.session_state.service_account_email}</p>
            <p class="highlight">Total Workbooks: {len(st.session_state.workbooks)}</p>
            <p>All your client booking spreadsheets are loaded and ready for management.</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="sidebar-card">
            <h3>🔴 Not Connected</h3>
            <p>Connect to Google Sheets to access your booking management system.</p>
            <p class="highlight">Required Setup:</p>
            <p>• Google Cloud Service Account</p>
            <p>• Sheets API enabled</p>
            <p>• Drive API enabled</p>
            <p>• Workbooks shared with service account</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Current Workbook Card
    if st.session_state.current_workbook and st.session_state.authenticated:
        manager = st.session_state.gc
        try:
            workbook = manager.open_workbook(st.session_state.current_workbook)
            if workbook:
                profile = manager.get_client_profile(workbook)
                calendars = manager.get_calendar_sheets(workbook)
                
                st.markdown(f"""
                <div class="sidebar-card" style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);">
                    <h3>📊 Active Workbook</h3>
                    <p class="highlight">{workbook.title}</p>
                    <p><strong>Client:</strong> {profile.get('client_name', 'Unknown')}</p>
                    <p><strong>Properties:</strong> {len(profile.get('properties', []))} location(s)</p>
                    <p><strong>Calendars:</strong> {len(calendars)} month(s)</p>
                </div>
                """, unsafe_allow_html=True)
        except:
            pass
    
    # System Features Card
    st.markdown("""
    <div class="sidebar-card" style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);">
        <h3>⚡ System Features</h3>
        <p class="highlight">Multi-Workbook Support</p>
        <p>Switch between client workbooks seamlessly.</p>
        
        <p class="highlight">Smart Calendar System</p>
        <p>Bookings start from row 13 in each calendar sheet.</p>
        
        <p class="highlight">Real-Time Sync</p>
        <p>Changes reflect immediately in Google Sheets.</p>
        
        <p class="highlight">Advanced Analytics</p>
        <p>Track occupancy, cleaning schedules, and revenue.</p>
    </div>
    """, unsafe_allow_html=True)

def authenticate():
    """Authentication section with enhanced UI"""
    st.markdown('<div class="main-header">🏠 Professional Booking Management System</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div class="info-box">
            <h3 style="margin-top: 0;">Welcome to Your Booking Hub</h3>
            <p>Manage multiple client properties, booking calendars, and service schedules all in one place. 
            Connect to your Google Sheets workbooks to get started.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.header("🔐 Authentication Center")
        
        render_sidebar_info_cards()
        
        st.markdown("---")
        
        with st.expander("🔑 **Setup Instructions**", expanded=False):
            st.markdown("""
            ### Getting Your Credentials
            
            1. **Google Cloud Console**
               - Go to console.cloud.google.com
               - Create new project or select existing
            
            2. **Enable APIs**
               - Enable Google Sheets API
               - Enable Google Drive API
            
            3. **Create Service Account**
               - IAM & Admin → Service Accounts
               - Create new service account
               - Generate JSON key
            
            4. **Share Workbooks**
               - Copy service account email
               - Share each Google Sheet with that email
               - Grant Editor permissions
            
            5. **Authenticate Here**
               - Paste JSON or upload file
               - Click connect button
            """)
        
        st.markdown("---")
        
        auth_method = st.radio(
            "Authentication Method",
            ["📝 Paste JSON Credentials", "📁 Upload Credentials File"],
            help="Choose how to provide your service account credentials"
        )
        
        if auth_method == "📝 Paste JSON Credentials":
            st.info("📋 Copy your entire service account JSON and paste below")
            creds_text = st.text_area(
                "Service Account JSON",
                height=250,
                placeholder='{\n  "type": "service_account",\n  "project_id": "...",\n  ...\n}',
                help="Paste the complete JSON from your downloaded credentials file"
            )
            
            if st.button("🚀 Connect to Google Sheets", type="primary", use_container_width=True):
                if creds_text:
                    with st.spinner("Authenticating..."):
                        try:
                            add_log("Starting authentication process...", "INFO")
                            creds_dict = json.loads(creds_text)
                            add_log("JSON credentials parsed successfully", "SUCCESS")
                            
                            manager = BookingManager(creds_dict)
                            st.session_state.gc = manager
                            st.session_state.authenticated = True
                            
                            add_log("Authentication completed successfully!", "SUCCESS")
                            st.success("✅ Successfully connected to Google Sheets!")
                            st.balloons()
                            
                            # Auto-load workbooks
                            with st.spinner("Loading workbooks..."):
                                st.session_state.workbooks = manager.list_workbooks()
                            
                            st.rerun()
                            
                        except json.JSONDecodeError as e:
                            add_log(f"JSON parsing error: {str(e)}", "ERROR")
                            st.error("❌ Invalid JSON format. Please check your credentials.")
                        except Exception as e:
                            add_log(f"Authentication failed: {str(e)}", "ERROR")
                            st.error(f"❌ Authentication failed: {str(e)}")
                else:
                    st.warning("⚠️ Please provide credentials")
        
        else:
            st.info("📁 Select your credentials JSON file")
            uploaded_file = st.file_uploader(
                "Upload Service Account JSON",
                type=['json'],
                help="Upload the JSON file downloaded from Google Cloud Console"
            )
            
            if uploaded_file:
                st.success(f"✅ File loaded: {uploaded_file.name}")
            
            if st.button("🚀 Connect with File", type="primary", use_container_width=True):
                if uploaded_file:
                    with st.spinner("Authenticating..."):
                        try:
                            add_log(f"Reading credentials from file: {uploaded_file.name}", "INFO")
                            creds_dict = json.load(uploaded_file)
                            add_log("Credentials file loaded successfully", "SUCCESS")
                            
                            manager = BookingManager(creds_dict)
                            st.session_state.gc = manager
                            st.session_state.authenticated = True
                            
                            add_log("Authentication completed successfully!", "SUCCESS")
                            st.success("✅ Successfully connected to Google Sheets!")
                            st.balloons()
                            
                            # Auto-load workbooks
                            with st.spinner("Loading workbooks..."):
                                st.session_state.workbooks = manager.list_workbooks()
                            
                            st.rerun()
                            
                        except Exception as e:
                            add_log(f"Authentication failed: {str(e)}", "ERROR")
                            st.error(f"❌ Authentication failed: {str(e)}")
                else:
                    st.warning("⚠️ Please upload a credentials file")

def render_dashboard():
    """Render main dashboard view"""
    st.markdown('<div class="section-header">📊 Dashboard Overview</div>', unsafe_allow_html=True)
    
    manager = st.session_state.gc
    
    if not st.session_state.current_workbook:
        st.markdown("""
        <div class="warning-box">
            <h3>⚠️ No Workbook Selected</h3>
            <p>Please select a workbook from the sidebar to view dashboard metrics and manage bookings.</p>
        </div>
        """, unsafe_allow_html=True)
        return
    
    try:
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("Failed to open workbook")
            return
        
        # Get client profile
        profile = manager.get_client_profile(workbook)
        
        # Get statistics
        stats = manager.get_booking_statistics(workbook)
        
        # Display metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <h2 style="color: #667eea; margin: 0;">📅 {stats.get('calendar_count', 0)}</h2>
                <p style="color: #6c757d; margin: 0.5rem 0 0 0;">Calendar Months</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <h2 style="color: #28a745; margin: 0;">🏠 {len(profile.get('properties', []))}</h2>
                <p style="color: #6c757d; margin: 0.5rem 0 0 0;">Properties</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <h2 style="color: #ff9800; margin: 0;">📋 {stats.get('total_bookings', 0)}</h2>
                <p style="color: #6c757d; margin: 0.5rem 0 0 0;">Total Bookings</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            ci_count = stats.get('status_counts', {}).get('CI', 0)
            st.markdown(f"""
            <div class="metric-card">
                <h2 style="color: #4CAF50; margin: 0;">✓ {ci_count}</h2>
                <p style="color: #6c757d; margin: 0.5rem 0 0 0;">Check-Ins</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Client Profile Section
        st.markdown('<div class="section-header">👤 Client Profile</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"""
            <div class="info-box">
                <h3 style="margin-top: 0;">📝 Basic Information</h3>
                <p><strong>Client Name:</strong> {profile.get('client_name', 'N/A')}</p>
                <p><strong>Check-Out Time:</strong> {profile.get('check_out_time', 'N/A')}</p>
                <p><strong>Check-In Time:</strong> {profile.get('check_in_time', 'N/A')}</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="info-box">
                <h3 style="margin-top: 0;">🔑 Access Information</h3>
                <p><strong>Keys:</strong> {profile.get('keys', 'N/A')}</p>
                <p><strong>Codes:</strong> {profile.get('codes', 'N/A')}</p>
                <p><strong>Laundry:</strong> {profile.get('laundry_services', 'N/A')}</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Properties Section
        st.markdown('<div class="section-header">🏘️ Properties</div>', unsafe_allow_html=True)
        
        properties = profile.get('properties', [])
        if properties:
            for i, prop in enumerate(properties):
                with st.expander(f"🏠 {prop.get('name', f'Property {i+1}')}", expanded=False):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write(f"**Address:** {prop.get('address', 'N/A')}")
                        st.write(f"**Standard Hours:** {prop.get('hours', 'N/A')}")
                    with col2:
                        st.write(f"**Stay-Over Hours:** {prop.get('so_hours', 'N/A')}")
        else:
            st.info("No properties found in client profile")
        
        # Status Code Distribution
        st.markdown('<div class="section-header">📊 Booking Status Distribution</div>', unsafe_allow_html=True)
        
        status_counts = stats.get('status_counts', {})
        if any(status_counts.values()):
            cols = st.columns(len(STATUS_CODES))
            for i, (code, info) in enumerate(STATUS_CODES.items()):
                with cols[i]:
                    count = status_counts.get(code, 0)
                    st.markdown(f"""
                    <div style="background: {info['color']}; color: white; padding: 1rem; border-radius: 8px; text-align: center;">
                        <h3 style="margin: 0;">{count}</h3>
                        <p style="margin: 0.5rem 0 0 0; font-size: 0.9rem;">{code}</p>
                        <p style="margin: 0; font-size: 0.75rem; opacity: 0.9;">{info['name']}</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("No booking status data available")
        
    except Exception as e:
        add_log(f"Error rendering dashboard: {str(e)}", "ERROR")
        st.error(f"Error loading dashboard: {str(e)}")

def render_calendar_view():
    """Render calendar view with booking details"""
    st.markdown('<div class="section-header">📅 Calendar View</div>', unsafe_allow_html=True)
    
    manager = st.session_state.gc
    
    if not st.session_state.current_workbook:
        st.warning("Please select a workbook first")
        return
    
    try:
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("Failed to open workbook")
            return
        
        calendars = manager.get_calendar_sheets(workbook)
        
        if not calendars:
            st.info("No calendar sheets found in this workbook")
            return
        
        # Calendar selector
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col1:
            calendar_names = [cal['name'] for cal in calendars]
            selected_calendar_name = st.selectbox(
                "Select Calendar Month",
                calendar_names,
                key="calendar_selector"
            )
        
        with col2:
            st.session_state.filter_status = st.selectbox(
                "Filter by Status",
                ["All"] + list(STATUS_CODES.keys()),
                key="status_filter"
            )
        
        with col3:
            profile = manager.get_client_profile(workbook)
            property_names = ["All"] + [prop['name'] for prop in profile.get('properties', [])]
            st.session_state.filter_property = st.selectbox(
                "Filter by Property",
                property_names,
                key="property_filter"
            )
        
        # Find selected calendar
        selected_calendar = next((cal for cal in calendars if cal['name'] == selected_calendar_name), None)
        
        if selected_calendar:
            st.session_state.selected_calendar = selected_calendar
            
            # Read calendar data
            df = manager.read_calendar(selected_calendar['sheet'])
            
            if not df.empty:
                st.markdown(f"""
                <div class="success-box">
                    <h4 style="margin-top: 0;">📋 {selected_calendar_name}</h4>
                    <p>Showing {len(df)} booking entries. Bookings start from row 13 in the Google Sheet.</p>
                </div>
                """, unsafe_allow_html=True)
                
                # Apply filters
                filtered_df = df.copy()
                
                if st.session_state.filter_status != "All":
                    # Filter rows containing the status code
                    mask = filtered_df.apply(lambda row: any(st.session_state.filter_status in str(cell) for cell in row), axis=1)
                    filtered_df = filtered_df[mask]
                
                if st.session_state.filter_property != "All":
                    # Filter by property name in first column
                    if len(filtered_df.columns) > 0:
                        filtered_df = filtered_df[filtered_df.iloc[:, 0].astype(str).str.contains(st.session_state.filter_property, case=False, na=False)]
                
                # Display filtered results
                st.markdown(f"**Filtered Results:** {len(filtered_df)} entries")
                
                # Display dataframe with styling
                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    height=600
                )
                
                # Export options
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download as CSV",
                        data=csv,
                        file_name=f"{selected_calendar_name}_bookings.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                with col2:
                    if st.button("🔄 Refresh Calendar", use_container_width=True):
                        st.rerun()
                
                with col3:
                    if st.button("📊 View Statistics", use_container_width=True):
                        st.session_state.current_view = "Dashboard"
                        st.rerun()
                
            else:
                st.info("No booking data found in this calendar")
        
    except Exception as e:
        add_log(f"Error rendering calendar view: {str(e)}", "ERROR")
        st.error(f"Error loading calendar: {str(e)}")

def render_booking_manager():
    """Render booking creation and editing interface"""
    st.markdown('<div class="section-header">➕ Booking Manager</div>', unsafe_allow_html=True)
    
    manager = st.session_state.gc
    
    if not st.session_state.current_workbook:
        st.warning("Please select a workbook first")
        return
    
    try:
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("Failed to open workbook")
            return
        
        profile = manager.get_client_profile(workbook)
        calendars = manager.get_calendar_sheets(workbook)
        
        # Booking form
        st.markdown("""
        <div class="info-box">
            <h3 style="margin-top: 0;">📝 Create New Booking</h3>
            <p>Fill in the details below to add a new booking to your calendar.</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("new_booking_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                # Calendar selection
                calendar_names = [cal['name'] for cal in calendars]
                selected_calendar = st.selectbox("Select Calendar Month", calendar_names)
                
                # Property selection
                property_names = [prop['name'] for prop in profile.get('properties', [])]
                selected_property = st.selectbox("Select Property", property_names if property_names else ["No properties available"])
                
                # Date selection
                booking_date = st.date_input("Booking Date", datetime.now())
                
                # Status code
                status_code = st.selectbox("Status Code", list(STATUS_CODES.keys()))
            
            with col2:
                # Guest name
                guest_name = st.text_input("Guest Name (Optional)")
                
                # Notes
                notes = st.text_area("Notes (Optional)", height=100)
                
                # Time
                booking_time = st.time_input("Time", datetime.now().time())
            
            # Status code info
            if status_code:
                status_info = STATUS_CODES[status_code]
                st.markdown(f"""
                <div style="background: {status_info['color']}; color: white; padding: 1rem; border-radius: 8px; margin: 1rem 0;">
                    <strong>{status_info['name']}</strong>: {status_info['description']}
                </div>
                """, unsafe_allow_html=True)
            
            submitted = st.form_submit_button("➕ Add Booking", type="primary", use_container_width=True)
            
            if submitted:
                if not property_names:
                    st.error("No properties available. Please add properties to the client profile first.")
                else:
                    # Find the selected calendar sheet
                    calendar_sheet = next((cal['sheet'] for cal in calendars if cal['name'] == selected_calendar), None)
                    
                    if calendar_sheet:
                        # Prepare booking data
                        booking_data = [
                            selected_property,
                            booking_date.strftime("%Y-%m-%d"),
                            status_code,
                            guest_name,
                            booking_time.strftime("%H:%M"),
                            notes
                        ]
                        
                        # Add booking
                        success = manager.add_booking_row(calendar_sheet, booking_data)
                        
                        if success:
                            st.success(f"✅ Booking added successfully to {selected_calendar}!")
                            st.balloons()
                        else:
                            st.error("Failed to add booking. Check logs for details.")
        
        # Quick actions
        st.markdown('<div class="section-header">⚡ Quick Actions</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📅 View All Calendars", use_container_width=True):
                st.session_state.current_view = "Calendar View"
                st.rerun()
        
        with col2:
            if st.button("📊 View Dashboard", use_container_width=True):
                st.session_state.current_view = "Dashboard"
                st.rerun()
        
        with col3:
            if st.button("🔄 Refresh Data", use_container_width=True):
                st.rerun()
        
    except Exception as e:
        add_log(f"Error in booking manager: {str(e)}", "ERROR")
        st.error(f"Error: {str(e)}")

def render_analytics():
    """Render analytics and reporting view"""
    st.markdown('<div class="section-header">📈 Analytics & Reports</div>', unsafe_allow_html=True)
    
    manager = st.session_state.gc
    
    if not st.session_state.current_workbook:
        st.warning("Please select a workbook first")
        return
    
    try:
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("Failed to open workbook")
            return
        
        stats = manager.get_booking_statistics(workbook)
        profile = manager.get_client_profile(workbook)
        
        # Overview metrics
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total Bookings", stats.get('total_bookings', 0))
        
        with col2:
            st.metric("Active Properties", len(profile.get('properties', [])))
        
        with col3:
            st.metric("Calendar Months", stats.get('calendar_count', 0))
        
        # Status distribution chart
        st.markdown("### 📊 Booking Status Distribution")
        
        status_counts = stats.get('status_counts', {})
        if any(status_counts.values()):
            # Create a simple bar chart using columns
            max_count = max(status_counts.values()) if status_counts.values() else 1
            
            for code, count in status_counts.items():
                if count > 0:
                    info = STATUS_CODES.get(code, {})
                    percentage = (count / max_count) * 100
                    
                    st.markdown(f"""
                    <div style="margin: 1rem 0;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 0.5rem;">
                            <span><strong>{code}</strong> - {info.get('name', code)}</span>
                            <span><strong>{count}</strong> bookings</span>
                        </div>
                        <div style="background: #e9ecef; border-radius: 10px; overflow: hidden;">
                            <div style="background: {info.get('color', '#667eea')}; width: {percentage}%; height: 30px; border-radius: 10px;"></div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("No booking data available for analysis")
        
        # Property distribution
        st.markdown("### 🏠 Bookings by Property")
        
        property_counts = stats.get('property_counts', {})
        if property_counts:
            for prop, count in sorted(property_counts.items(), key=lambda x: x[1], reverse=True):
                if prop.strip():
                    st.markdown(f"""
                    <div class="property-card">
                        <h4 style="margin: 0 0 0.5rem 0;">🏠 {prop}</h4>
                        <p style="margin: 0; color: #667eea; font-size: 1.5rem; font-weight: bold;">{count} bookings</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("No property data available")
        
        # Export analytics
        st.markdown("### 📥 Export Reports")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📊 Generate Full Report", use_container_width=True):
                report_data = {
                    'Client': profile.get('client_name', 'Unknown'),
                    'Total Bookings': stats.get('total_bookings', 0),
                    'Properties': len(profile.get('properties', [])),
                    'Calendars': stats.get('calendar_count', 0),
                    'Status Breakdown': status_counts,
                    'Property Breakdown': property_counts
                }
                
                report_json = json.dumps(report_data, indent=2)
                st.download_button(
                    label="💾 Download JSON Report",
                    data=report_json,
                    file_name=f"booking_report_{datetime.now().strftime('%Y%m%d')}.json",
                    mime="application/json"
                )
        
        with col2:
            if st.button("📧 Email Report (Coming Soon)", use_container_width=True, disabled=True):
                st.info("Email reporting feature coming soon!")
        
    except Exception as e:
        add_log(f"Error rendering analytics: {str(e)}", "ERROR")
        st.error(f"Error loading analytics: {str(e)}")

def render_logs():
    """Render system logs"""
    st.markdown('<div class="section-header">📜 System Logs</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.markdown("""
        <div class="info-box">
            <p style="margin: 0;">View detailed system activity logs including authentication, data loading, and operations.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if st.button("🗑️ Clear Logs", use_container_width=True):
            st.session_state.logs = []
            add_log("Logs cleared by user", "INFO")
            st.rerun()
    
    # Display logs
    if st.session_state.logs:
        st.markdown(f"**Total Entries:** {len(st.session_state.logs)}")
        
        # Reverse to show newest first
        for log in reversed(st.session_state.logs[-50:]):  # Show last 50 logs
            level_class = f"log-{log['level'].lower()}"
            st.markdown(f"""
            <div class="log-entry {level_class}">
                <strong>[{log['timestamp']}]</strong> 
                <span style="color: {'#28a745' if log['level'] == 'SUCCESS' else '#17a2b8' if log['level'] == 'INFO' else '#ffc107' if log['level'] == 'WARNING' else '#dc3545'};">
                    {log['level']}
                </span>: {log['message']}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No logs available")

def main_app():
    """Main application interface"""
    manager = st.session_state.gc
    
    # Sidebar Navigation
    with st.sidebar:
        st.header("🗂️ Workbook Management")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Refresh", use_container_width=True):
                add_log("Refreshing workbook list...", "INFO")
                with st.spinner("Refreshing..."):
                    st.session_state.workbooks = manager.list_workbooks()
                st.success("✅ Refreshed!")
                st.rerun()
        
        with col2:
            if st.button("🚪 Logout", use_container_width=True):
                st.session_state.authenticated = False
                st.session_state.gc = None
                st.session_state.workbooks = []
                st.session_state.current_workbook = None
                add_log("User logged out", "INFO")
                st.rerun()
        
        st.markdown("---")
        
        # Workbook selector
        if st.session_state.workbooks:
            workbook_options = {wb['name']: wb['id'] for wb in st.session_state.workbooks}
            selected_workbook_name = st.selectbox(
                "📚 Select Workbook",
                options=list(workbook_options.keys()),
                help="Choose a client workbook to manage"
            )
            
            if selected_workbook_name:
                selected_id = workbook_options[selected_workbook_name]
                if st.session_state.current_workbook != selected_id:
                    st.session_state.current_workbook = selected_id
                    add_log(f"Switched to workbook: {selected_workbook_name}", "INFO")
                    st.rerun()
        else:
            st.warning("No workbooks found. Make sure your service account has access to Google Sheets.")
        
        st.markdown("---")
        
        # View selector
        st.header("📍 Navigation")
        
        view_options = [
            "Dashboard",
            "Calendar View",
            "Booking Manager",
            "Analytics",
            "System Logs"
        ]
        
        for view in view_options:
            icon = {
                "Dashboard": "📊",
                "Calendar View": "📅",
                "Booking Manager": "➕",
                "Analytics": "📈",
                "System Logs": "📜"
            }.get(view, "📄")
            
            if st.button(f"{icon} {view}", use_container_width=True, type="primary" if st.session_state.current_view == view else "secondary"):
                st.session_state.current_view = view
                add_log(f"Navigated to: {view}", "INFO")
                st.rerun()
        
        st.markdown("---")
        
        # Sidebar info cards
        render_sidebar_info_cards()
    
    # Main content area
    st.markdown('<div class="main-header">🏠 Booking Management System</div>', unsafe_allow_html=True)
    
    # Render selected view
    if st.session_state.current_view == "Dashboard":
        render_dashboard()
    elif st.session_state.current_view == "Calendar View":
        render_calendar_view()
    elif st.session_state.current_view == "Booking Manager":
        render_booking_manager()
    elif st.session_state.current_view == "Analytics":
        render_analytics()
    elif st.session_state.current_view == "System Logs":
        render_logs()

# Main application flow
def main():
    """Main application entry point"""
    if not st.session_state.authenticated:
        authenticate()
    else:
        main_app()

if __name__ == "__main__":
    main()
