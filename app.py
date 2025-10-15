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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

DRIVE_FOLDER_ID = "1Fk5dJGkm5dNMZkfsITe5Lt9x-yCsBiF2"

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
        font-size: 2.8rem;
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
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .drive-config-card {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        color: white;
        box-shadow: 0 6px 12px rgba(0,0,0,0.2);
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
    .sheet-selector {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border: 2px solid #dee2e6;
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
if 'service_account_email' not in st.session_state:
    st.session_state.service_account_email = None
if 'all_sheets' not in st.session_state:
    st.session_state.all_sheets = []
if 'selected_sheet_index' not in st.session_state:
    st.session_state.selected_sheet_index = 0

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
    
    def list_workbooks_from_folder(self, folder_id: str) -> List[Dict]:
        """List all spreadsheets from a specific Google Drive folder"""
        try:
            add_log(f"Listing workbooks from folder: {folder_id}", "INFO")
            
            try:
                from googleapiclient.discovery import build
                from googleapiclient.errors import HttpError
                add_log("Google API client imported successfully", "INFO")
            except ImportError as e:
                add_log(f"Failed to import Google API client: {str(e)}", "ERROR")
                add_log("Attempting to use gspread's built-in methods instead...", "WARNING")
                return self._list_workbooks_fallback(folder_id)
            
            try:
                # Build Drive service
                add_log("Building Drive API service...", "INFO")
                drive_service = build('drive', 'v3', credentials=self.creds)
                add_log("Drive API service created successfully", "SUCCESS")
                
                # Query for spreadsheets in the specific folder
                query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
                
                add_log(f"Executing Drive API query: {query}", "INFO")
                
                results = drive_service.files().list(
                    q=query,
                    pageSize=100,
                    fields="files(id, name, webViewLink, modifiedTime, owners)",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True
                ).execute()
                
                files = results.get('files', [])
                add_log(f"Drive API returned {len(files)} file(s)", "INFO")
                
                workbooks = []
                for file in files:
                    workbook_info = {
                        'id': file['id'],
                        'name': file['name'],
                        'url': file.get('webViewLink', f"https://docs.google.com/spreadsheets/d/{file['id']}"),
                        'modified': file.get('modifiedTime', 'Unknown')
                    }
                    workbooks.append(workbook_info)
                    add_log(f"  ✓ Found: {file['name']} (ID: {file['id']})", "SUCCESS")
                
                if len(workbooks) == 0:
                    add_log("No workbooks found in folder", "WARNING")
                    add_log(f"Make sure folder {folder_id} is shared with: {st.session_state.service_account_email}", "WARNING")
                else:
                    add_log(f"Successfully loaded {len(workbooks)} workbook(s) from folder", "SUCCESS")
                
                return workbooks
                
            except HttpError as e:
                add_log(f"Drive API HTTP Error: {str(e)}", "ERROR")
                add_log("Attempting fallback method...", "WARNING")
                return self._list_workbooks_fallback(folder_id)
            except Exception as e:
                add_log(f"Drive API Error: {str(e)}", "ERROR")
                add_log("Attempting fallback method...", "WARNING")
                return self._list_workbooks_fallback(folder_id)
                
        except Exception as e:
            add_log(f"Error listing workbooks from folder: {str(e)}", "ERROR")
            return []
    
    def _list_workbooks_fallback(self, folder_id: str) -> List[Dict]:
        """Fallback method to list workbooks using gspread"""
        try:
            add_log("Using gspread fallback method to list spreadsheets...", "INFO")
            
            all_spreadsheets = self.gc.openall()
            add_log(f"Found {len(all_spreadsheets)} accessible spreadsheet(s)", "INFO")
            
            workbooks = []
            for sheet in all_spreadsheets:
                workbook_info = {
                    'id': sheet.id,
                    'name': sheet.title,
                    'url': sheet.url,
                    'modified': 'Unknown'
                }
                workbooks.append(workbook_info)
                add_log(f"  ✓ {sheet.title} (ID: {sheet.id})", "INFO")
            
            if len(workbooks) > 0:
                add_log(f"Fallback method found {len(workbooks)} workbook(s)", "SUCCESS")
            else:
                add_log("No accessible spreadsheets found", "WARNING")
            
            return workbooks
            
        except Exception as e:
            add_log(f"Fallback method also failed: {str(e)}", "ERROR")
            return []
    
    def open_workbook(self, workbook_id: str):
        """Open a specific workbook by ID"""
        try:
            add_log(f"Opening workbook ID: {workbook_id}", "INFO")
            workbook = self.gc.open_by_key(workbook_id)
            add_log(f"Successfully opened: {workbook.title}", "SUCCESS")
            
            all_sheets = workbook.worksheets()
            st.session_state.all_sheets = [
                {
                    'index': i,
                    'name': sheet.title,
                    'id': sheet.id,
                    'sheet_obj': sheet
                }
                for i, sheet in enumerate(all_sheets)
            ]
            
            add_log(f"Workbook contains {len(all_sheets)} sheet(s)", "INFO")
            for i, sheet_info in enumerate(st.session_state.all_sheets):
                add_log(f"  Sheet {i}: {sheet_info['name']}", "INFO")
            
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
            for i in range(17, min(len(all_values), 30)):
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
                add_log(f"  Calendar sheet found: {sheet.title}", "INFO")
            
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
            add_log(f"Headers: {', '.join([h for h in headers if h])}", "INFO")
            
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
    
    def read_sheet_all_data(self, sheet) -> pd.DataFrame:
        """Read all data from any sheet"""
        try:
            add_log(f"Reading all data from sheet: {sheet.title}", "INFO")
            all_values = sheet.get_all_values()
            
            if not all_values:
                add_log("Sheet is empty", "WARNING")
                return pd.DataFrame()
            
            # Use first row as headers
            df = pd.DataFrame(all_values[1:], columns=all_values[0])
            
            add_log(f"Loaded {len(df)} rows with {len(df.columns)} columns", "SUCCESS")
            
            return df
            
        except Exception as e:
            add_log(f"Error reading sheet: {str(e)}", "ERROR")
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
    
    def add_booking_row(self, sheet, start_row: int, data: List) -> bool:
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

def authenticate():
    """Authentication section with JSON file upload only"""
    st.markdown('<div class="main-header">🏠 Professional Booking Management System</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div class="info-box">
            <h3 style="margin-top: 0;">Welcome to Your Booking Hub</h3>
            <p>Manage multiple client properties, booking calendars, and service schedules all in one place. 
            Upload your Google service account JSON file to get started.</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="drive-config-card">
            <h3 style="color: white; margin-top: 0;">📁 Google Drive Folder</h3>
            <p style="color: white;"><strong>Folder ID:</strong> {DRIVE_FOLDER_ID}</p>
            <p style="color: white; font-size: 0.9rem;">All workbooks will be loaded from this folder automatically.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.header("🔐 Authentication")
        
        st.markdown("""
        <div class="sidebar-card">
            <h3>📋 Setup Instructions</h3>
            <p class="highlight">Step 1: Google Cloud Console</p>
            <p>Create a service account with Sheets & Drive API access</p>
            
            <p class="highlight">Step 2: Download JSON Key</p>
            <p>Generate and download the credentials JSON file</p>
            
            <p class="highlight">Step 3: Share Folder</p>
            <p>Share the Google Drive folder with your service account email</p>
            
            <p class="highlight">Step 4: Upload & Connect</p>
            <p>Upload the JSON file below to authenticate</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.info("📁 Upload your service account credentials JSON file")
        uploaded_file = st.file_uploader(
            "Service Account JSON File",
            type=['json'],
            help="Upload the JSON file downloaded from Google Cloud Console"
        )
        
        if uploaded_file:
            st.success(f"✅ File loaded: {uploaded_file.name}")
        
        if st.button("🚀 Connect to Google Sheets", type="primary", use_container_width=True):
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
                        
                        with st.spinner(f"Loading workbooks from folder {DRIVE_FOLDER_ID}..."):
                            add_log("Starting workbook discovery...", "INFO")
                            st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
                        
                        if len(st.session_state.workbooks) > 0:
                            st.success(f"✅ Found {len(st.session_state.workbooks)} workbook(s)!")
                            st.balloons()
                        else:
                            st.warning("⚠️ No workbooks found in folder")
                            st.info(f"Make sure folder {DRIVE_FOLDER_ID} is shared with: {st.session_state.service_account_email}")
                        
                        st.rerun()
                        
                    except Exception as e:
                        add_log(f"Authentication failed: {str(e)}", "ERROR")
                        st.error(f"❌ Authentication failed: {str(e)}")
            else:
                st.warning("⚠️ Please upload a credentials file")

def main_app():
    """Main application interface"""
    manager = st.session_state.gc
    
    st.markdown('<div class="main-header">🏠 Booking Management Dashboard</div>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("🗂️ Workbook Selection")
        
        st.markdown(f"""
        <div class="drive-config-card">
            <h3 style="color: white; margin-top: 0;">📁 Active Folder</h3>
            <p style="color: white; font-size: 0.85rem; word-break: break-all;">{DRIVE_FOLDER_ID}</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Refresh", use_container_width=True):
                add_log("Refreshing workbook list...", "INFO")
                with st.spinner("Refreshing..."):
                    st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
                st.success(f"✅ Found {len(st.session_state.workbooks)} workbooks!")
                st.rerun()
        
        with col2:
            if st.button("🚪 Logout", use_container_width=True):
                st.session_state.authenticated = False
                st.session_state.gc = None
                st.session_state.workbooks = []
                st.session_state.current_workbook = None
                st.session_state.all_sheets = []
                add_log("Logged out successfully", "INFO")
                st.rerun()
        
        st.markdown("---")
        
        if st.session_state.workbooks:
            workbook_names = [wb['name'] for wb in st.session_state.workbooks]
            selected_workbook_name = st.selectbox(
                "Select Workbook",
                workbook_names,
                help="Choose a client workbook to manage"
            )
            
            selected_workbook = next(
                (wb for wb in st.session_state.workbooks if wb['name'] == selected_workbook_name),
                None
            )
            
            if selected_workbook:
                if st.session_state.current_workbook != selected_workbook['id']:
                    st.session_state.current_workbook = selected_workbook['id']
                    st.session_state.selected_sheet_index = 0
                    add_log(f"Switched to workbook: {selected_workbook['name']}", "INFO")
                    st.rerun()
                
                st.markdown(f"""
                <div class="info-box">
                    <p><strong>📊 Active Workbook:</strong><br>{selected_workbook['name']}</p>
                    <p><strong>🔗 ID:</strong><br><span style="font-size: 0.8rem; word-break: break-all;">{selected_workbook['id']}</span></p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.warning("⚠️ No workbooks found in the folder")
            st.info("Make sure the Google Drive folder is shared with your service account email")
        
        st.markdown("---")
        
        # View selector
        view_mode = st.radio(
            "📑 View Mode",
            ["Dashboard", "All Sheets Viewer", "Calendar View", "Booking Manager", "System Logs"],
            help="Choose what to display"
        )
        
        st.markdown("---")
        
        # Connection info
        if st.session_state.service_account_email:
            st.markdown(f"""
            <div class="sidebar-card">
                <h3>🟢 Connected</h3>
                <p class="highlight">Service Account</p>
                <p style="font-size: 0.75rem; word-break: break-all;">{st.session_state.service_account_email}</p>
                <p class="highlight">Workbooks: {len(st.session_state.workbooks)}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Main content area
    if not st.session_state.current_workbook:
        st.info("👈 Please select a workbook from the sidebar to begin")
        return
    
    # Load current workbook
    workbook = manager.open_workbook(st.session_state.current_workbook)
    if not workbook:
        st.error("❌ Failed to load workbook")
        return
    
    if view_mode == "Dashboard":
        render_dashboard(manager, workbook)
    elif view_mode == "All Sheets Viewer":
        render_all_sheets_viewer(manager, workbook)
    elif view_mode == "Calendar View":
        render_calendar_view(manager, workbook)
    elif view_mode == "Booking Manager":
        render_booking_manager(manager, workbook)
    elif view_mode == "System Logs":
        render_system_logs()

def render_all_sheets_viewer(manager, workbook):
    """View and select any sheet from the workbook"""
    st.markdown('<div class="section-header">📄 All Sheets Viewer</div>', unsafe_allow_html=True)
    
    if not st.session_state.all_sheets:
        st.warning("No sheets found in this workbook")
        return
    
    st.markdown('<div class="sheet-selector">', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([3, 1, 1])
    
    with col1:
        sheet_names = [sheet['name'] for sheet in st.session_state.all_sheets]
        selected_sheet_name = st.selectbox(
            "Select Sheet to View",
            sheet_names,
            index=st.session_state.selected_sheet_index,
            help="Choose any sheet from the workbook"
        )
        
        # Update selected index
        st.session_state.selected_sheet_index = sheet_names.index(selected_sheet_name)
    
    with col2:
        st.metric("Total Sheets", len(st.session_state.all_sheets))
    
    with col3:
        if st.button("🔄 Refresh"):
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Get selected sheet
    selected_sheet_info = st.session_state.all_sheets[st.session_state.selected_sheet_index]
    sheet_obj = selected_sheet_info['sheet_obj']
    
    st.markdown(f"""
    <div class="info-box">
        <h3 style="margin-top: 0;">📋 {selected_sheet_info['name']}</h3>
        <p><strong>Sheet Index:</strong> {selected_sheet_info['index']}</p>
        <p><strong>Sheet ID:</strong> {selected_sheet_info['id']}</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        view_type = st.radio(
            "View Type",
            ["Full Data", "Calendar Format (Row 13+)"],
            help="Choose how to display the sheet"
        )
    
    with col2:
        show_info = st.checkbox("Show Sheet Statistics", value=True)
    
    # Load and display data
    with st.spinner(f"Loading {selected_sheet_info['name']}..."):
        if view_type == "Full Data":
            df = manager.read_sheet_all_data(sheet_obj)
        else:
            df = manager.read_calendar(sheet_obj, start_row=13)
        
        if not df.empty:
            st.success(f"✅ Loaded {len(df)} rows with {len(df.columns)} columns")
            
            # Display dataframe
            st.dataframe(
                df,
                use_container_width=True,
                height=600
            )
            
            # Export option
            csv = df.to_csv(index=False)
            st.download_button(
                label="📥 Download as CSV",
                data=csv,
                file_name=f"{selected_sheet_info['name']}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
            # Show statistics
            if show_info:
                st.markdown("### 📊 Sheet Statistics")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Rows", len(df))
                with col2:
                    st.metric("Total Columns", len(df.columns))
                with col3:
                    non_empty = df.notna().sum().sum()
                    st.metric("Non-Empty Cells", non_empty)
                with col4:
                    empty = df.isna().sum().sum()
                    st.metric("Empty Cells", empty)
        else:
            st.warning("⚠️ No data found in this sheet")

def render_dashboard(manager, workbook):
    """Render dashboard view"""
    st.markdown('<div class="section-header">📊 Dashboard Overview</div>', unsafe_allow_html=True)
    
    # Get client profile
    profile = manager.get_client_profile(workbook)
    
    # Display client info
    st.markdown(f"""
    <div class="success-box">
        <h2 style="margin-top: 0;">👤 {profile.get('client_name', 'Unknown Client')}</h2>
        <p><strong>Check-out Time:</strong> {profile.get('check_out_time', 'N/A')}</p>
        <p><strong>Check-in Time:</strong> {profile.get('check_in_time', 'N/A')}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Properties
    if profile.get('properties'):
        st.markdown("### 🏘️ Properties")
        for prop in profile['properties']:
            st.markdown(f"""
            <div class="property-card">
                <h4>{prop['name']}</h4>
                <p><strong>Address:</strong> {prop['address']}</p>
                <p><strong>Hours:</strong> {prop['hours']} | <strong>SO Hours:</strong> {prop['so_hours']}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Calendar sheets
    calendars = manager.get_calendar_sheets(workbook)
    st.markdown(f"### 📅 Calendar Sheets ({len(calendars)} months)")
    
    for cal in calendars:
        with st.expander(f"📆 {cal['name']}"):
            df = manager.read_calendar(cal['sheet'])
            if not df.empty:
                st.dataframe(df, use_container_width=True)
            else:
                st.info("No bookings found")

def render_calendar_view(manager, workbook):
    """Render calendar view"""
    st.markdown('<div class="section-header">📅 Calendar View</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar = st.selectbox("Select Calendar Month", calendar_names)
    
    selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
    
    if selected_cal:
        df = manager.read_calendar(selected_cal['sheet'])
        
        if not df.empty:
            st.dataframe(df, use_container_width=True, height=600)
            
            csv = df.to_csv(index=False)
            st.download_button(
                "📥 Download Calendar",
                data=csv,
                file_name=f"{selected_calendar}.csv",
                mime="text/csv"
            )
        else:
            st.info("No bookings in this calendar")

def render_booking_manager(manager, workbook):
    """Render booking manager"""
    st.markdown('<div class="section-header">➕ Booking Manager</div>', unsafe_allow_html=True)
    
    st.info("Booking creation interface - Coming soon!")

def render_system_logs():
    """Render system logs"""
    st.markdown('<div class="section-header">📋 System Logs</div>', unsafe_allow_html=True)
    
    if st.button("🗑️ Clear Logs"):
        st.session_state.logs = []
        st.rerun()
    
    if st.session_state.logs:
        for log in reversed(st.session_state.logs):
            level_class = f"log-{log['level'].lower()}"
            st.markdown(f"""
            <div class="log-entry {level_class}">
                <strong>[{log['timestamp']}]</strong> [{log['level']}] {log['message']}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No logs yet")

# Main execution
if __name__ == "__main__":
    if not st.session_state.authenticated:
        authenticate()
    else:
        main_app()
