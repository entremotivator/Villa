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
EXAMPLE_SPREADSHEET_IDS = [
    "1ge6-Rzor5jbQ7zaaQk3B7I0Vx31Nv80QH6zW2NfBUz8",
    "1-I0lHMXrA16v07Qtc3BNqnjRux7PRf8fMxJhVlxAcME"
]

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
    /* Updated all info boxes to have black text */
    .info-box {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #667eea;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: #000000;
    }
    .info-box h2, .info-box h3, .info-box h4, .info-box p, .info-box strong {
        color: #000000 !important;
    }
    .warning-box {
        background: linear-gradient(135deg, #fff9e6 0%, #ffe8b3 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #ffc107;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: #000000;
    }
    .warning-box h2, .warning-box h3, .warning-box h4, .warning-box p, .warning-box strong {
        color: #000000 !important;
    }
    .success-box {
        background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #28a745;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: #000000;
    }
    .success-box h2, .success-box h3, .success-box h4, .success-box p, .success-box strong {
        color: #000000 !important;
    }
    .error-box {
        background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #f44336;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: #000000;
    }
    .error-box h2, .error-box h3, .error-box h4, .error-box p, .error-box strong {
        color: #000000 !important;
    }
    /* Updated metric cards to have black text */
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        text-align: center;
        border: 2px solid #e9ecef;
        transition: transform 0.3s ease;
        color: #000000;
    }
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
    }
    .metric-card h1, .metric-card h2, .metric-card h3, .metric-card h4, .metric-card p, .metric-card strong {
        color: #000000 !important;
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
    /* Updated property cards to have black text */
    .property-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px solid #e9ecef;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
        color: #000000;
    }
    .property-card:hover {
        border-color: #667eea;
        box-shadow: 0 6px 12px rgba(102, 126, 234, 0.3);
    }
    .property-card h1, .property-card h2, .property-card h3, .property-card h4, .property-card p, .property-card strong {
        color: #000000 !important;
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
                'https://www.googleapis.com/auth/drive.readonly',
                'https://www.googleapis.com/auth/drive.metadata.readonly'
            ]
            
            self.creds = Credentials.from_service_account_info(
                credentials_dict, 
                scopes=self.scopes
            )
            
            add_log("Creating Google Sheets client...", "INFO")
            self.gc = gspread.authorize(self.creds)
            
            try:
                from googleapiclient.discovery import build
                self.drive_service = build('drive', 'v3', credentials=self.creds)
                add_log("✅ Drive API service initialized", "SUCCESS")
            except ImportError:
                self.drive_service = None
                add_log("⚠️ google-api-python-client not installed - install with: pip install google-api-python-client", "WARNING")
            
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
            add_log("=" * 70, "INFO")
            add_log(f"🔍 SCANNING GOOGLE DRIVE FOLDER", "INFO")
            add_log(f"📁 Folder ID: {folder_id}", "INFO")
            add_log(f"👤 Service Account: {st.session_state.service_account_email}", "INFO")
            add_log("=" * 70, "INFO")
            
            workbooks = []
            
            if self.drive_service:
                add_log("🚀 Using Google Drive API v3...", "INFO")
                
                try:
                    query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
                    add_log(f"📋 Query: {query}", "INFO")
                    
                    page_token = None
                    total_files = 0
                    page_num = 0
                    
                    while True:
                        page_num += 1
                        add_log(f"📄 Fetching page {page_num}...", "INFO")
                        
                        results = self.drive_service.files().list(
                            q=query,
                            pageSize=100,
                            fields="nextPageToken, files(id, name, webViewLink, modifiedTime, owners)",
                            supportsAllDrives=True,
                            includeItemsFromAllDrives=True,
                            corpora='allDrives',
                            pageToken=page_token
                        ).execute()
                        
                        files = results.get('files', [])
                        add_log(f"   Found {len(files)} file(s) on page {page_num}", "INFO")
                        
                        for file in files:
                            total_files += 1
                            workbooks.append({
                                'id': file['id'],
                                'name': file['name'],
                                'url': file.get('webViewLink', f"https://docs.google.com/spreadsheets/d/{file['id']}"),
                                'modified': file.get('modifiedTime', 'Unknown')
                            })
                            add_log(f"   ✅ {total_files}. {file['name']}", "SUCCESS")
                            add_log(f"      ID: {file['id']}", "INFO")
                        
                        page_token = results.get('nextPageToken')
                        if not page_token:
                            break
                    
                    add_log("=" * 70, "SUCCESS")
                    add_log(f"✅ SCAN COMPLETE: Found {total_files} spreadsheet(s)", "SUCCESS")
                    add_log("=" * 70, "SUCCESS")
                    
                    if total_files == 0:
                        add_log("", "WARNING")
                        add_log("⚠️ NO SPREADSHEETS FOUND IN FOLDER", "WARNING")
                        add_log("", "WARNING")
                        add_log("TROUBLESHOOTING:", "WARNING")
                        add_log(f"1. Verify folder ID is correct: {folder_id}", "WARNING")
                        add_log(f"2. Share folder with: {st.session_state.service_account_email}", "WARNING")
                        add_log("3. Grant 'Viewer' or 'Editor' access", "WARNING")
                        add_log("4. Check that spreadsheets exist in the folder", "WARNING")
                        add_log("5. Wait a few minutes after sharing for permissions to propagate", "WARNING")
                    
                    return workbooks
                    
                except Exception as e:
                    add_log(f"❌ Drive API Error: {str(e)}", "ERROR")
                    add_log(f"Error type: {type(e).__name__}", "ERROR")
                    
                    if "403" in str(e):
                        add_log("", "ERROR")
                        add_log("🔒 PERMISSION DENIED", "ERROR")
                        add_log(f"The service account does not have access to this folder.", "ERROR")
                        add_log(f"", "ERROR")
                        add_log(f"TO FIX:", "WARNING")
                        add_log(f"1. Open: https://drive.google.com/drive/folders/{folder_id}", "WARNING")
                        add_log(f"2. Click 'Share' button", "WARNING")
                        add_log(f"3. Add: {st.session_state.service_account_email}", "WARNING")
                        add_log(f"4. Grant 'Viewer' or 'Editor' access", "WARNING")
                        add_log(f"5. Click 'Send' and wait 1-2 minutes", "WARNING")
                    elif "404" in str(e):
                        add_log("", "ERROR")
                        add_log("❌ FOLDER NOT FOUND", "ERROR")
                        add_log(f"Folder ID '{folder_id}' does not exist or is not accessible", "ERROR")
                    
                    return []
            else:
                add_log("❌ Drive API not available", "ERROR")
                add_log("Install google-api-python-client: pip install google-api-python-client", "ERROR")
                return []
                
        except Exception as e:
            add_log(f"❌ CRITICAL ERROR: {str(e)}", "ERROR")
            return []
    
    def open_workbook_by_id(self, workbook_id: str):
        """Open a specific workbook by ID (for manual entry)"""
        try:
            add_log(f"Opening workbook by ID: {workbook_id}", "INFO")
            workbook = self.gc.open_by_key(workbook_id)
            add_log(f"Successfully opened: {workbook.title}", "SUCCESS")
            
            # Add to workbooks list if not already there
            if not any(wb['id'] == workbook_id for wb in st.session_state.workbooks):
                st.session_state.workbooks.append({
                    'id': workbook.id,
                    'name': workbook.title,
                    'url': workbook.url,
                    'modified': 'Unknown'
                })
                add_log(f"Added {workbook.title} to workbooks list", "SUCCESS")
            
            return workbook
        except Exception as e:
            add_log(f"Error opening workbook by ID: {str(e)}", "ERROR")
            return None
    
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
            
            seen = {}
            unique_headers = []
            for header in headers:
                if header in seen:
                    seen[header] += 1
                    unique_headers.append(f"{header}_{seen[header]}")
                else:
                    seen[header] = 0
                    unique_headers.append(header)
            
            add_log(f"Headers: {', '.join([h for h in unique_headers if h])}", "INFO")
            
            # Get data starting from start_row
            data = all_values[start_row - 1:]
            
            df = pd.DataFrame(data, columns=unique_headers)
            
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
            
            headers = all_values[0]
            seen = {}
            unique_headers = []
            for header in headers:
                if header in seen:
                    seen[header] += 1
                    unique_headers.append(f"{header}_{seen[header]}")
                else:
                    seen[header] = 0
                    unique_headers.append(header)
            
            # Use unique headers
            df = pd.DataFrame(all_values[1:], columns=unique_headers)
            
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

    def copy_row(self, sheet, row_index: int) -> List:
        """Copy a row from the sheet"""
        try:
            add_log(f"Copying row {row_index} from {sheet.title}", "INFO")
            all_values = sheet.get_all_values()
            
            if row_index < len(all_values):
                row_data = all_values[row_index]
                add_log(f"Row copied successfully: {len(row_data)} cells", "SUCCESS")
                return row_data
            else:
                add_log(f"Row index {row_index} out of range", "ERROR")
                return []
        except Exception as e:
            add_log(f"Error copying row: {str(e)}", "ERROR")
            return []
    
    def append_row(self, sheet, data: List) -> bool:
        """Append a new row to the sheet"""
        try:
            add_log(f"Appending row to {sheet.title}", "INFO")
            sheet.append_row(data)
            add_log(f"Row appended successfully with {len(data)} cells", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error appending row: {str(e)}", "ERROR")
            return False
    
    def update_cell_live(self, sheet, row: int, col: int, value: str) -> bool:
        """Update a single cell in real-time"""
        try:
            add_log(f"Updating cell [{row}, {col}] in {sheet.title} to: {value}", "INFO")
            sheet.update_cell(row, col, value)
            add_log("Cell updated successfully", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error updating cell: {str(e)}", "ERROR")
            return False
    
    def batch_update_cells(self, sheet, cell_list: List[Dict]) -> bool:
        """Batch update multiple cells at once
        cell_list format: [{'row': 1, 'col': 1, 'value': 'text'}, ...]
        """
        try:
            add_log(f"Batch updating {len(cell_list)} cells in {sheet.title}", "INFO")
            
            for cell_data in cell_list:
                sheet.update_cell(cell_data['row'], cell_data['col'], cell_data['value'])
            
            add_log(f"Batch update completed: {len(cell_list)} cells updated", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error in batch update: {str(e)}", "ERROR")
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
            <p style="color: white;"><a href="https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}" target="_blank" style="color: white; text-decoration: underline;">🔗 Open Folder in Google Drive</a></p>
            <p style="color: white; font-size: 0.9rem; margin-top: 1rem;">⚠️ Make sure to share this folder with your service account email!</p>
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
        
        st.warning("📦 Required: `pip install google-api-python-client`")
        
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
                        
                        st.info(f"📧 Service Account: {st.session_state.service_account_email}")
                        st.info(f"📁 Make sure folder {DRIVE_FOLDER_ID} is shared with this email!")
                        
                        with st.spinner(f"Scanning folder for spreadsheets..."):
                            add_log("Starting folder scan...", "INFO")
                            st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
                        
                        if len(st.session_state.workbooks) > 0:
                            st.success(f"✅ Found {len(st.session_state.workbooks)} spreadsheet(s)!")
                            for wb in st.session_state.workbooks:
                                st.write(f"📊 {wb['name']}")
                            st.balloons()
                        else:
                            st.error("❌ No spreadsheets found in folder")
                            st.markdown(f"""
                            **To fix this:**
                            1. Open [this folder](https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID})
                            2. Click the 'Share' button
                            3. Add: `{st.session_state.service_account_email}`
                            4. Grant 'Viewer' or 'Editor' access
                            5. Click 'Send' and wait 1-2 minutes
                            6. Click 'Refresh' button in the sidebar
                            """)
                        
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
        
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🔄 Refresh", use_container_width=True):
                add_log("Refreshing workbook list...", "INFO")
                with st.spinner("Refreshing..."):
                    st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
                if st.session_state.workbooks:
                    st.success(f"✅ Found {len(st.session_state.workbooks)} workbooks!")
                else:
                    st.warning("⚠️ No workbooks found")
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
        
        with col3:
            st.markdown("[📖 Docs](https://docs.streamlit.io/library/api-reference/index)", unsafe_allow_html=True)
        
        st.markdown("---")
        
        with st.expander("➕ Add Spreadsheet by ID"):
            st.markdown("""
            <div class="info-box">
                <p style="margin: 0;">If automatic loading fails, paste a spreadsheet ID here:</p>
            </div>
            """, unsafe_allow_html=True)
            
            manual_id = st.text_input(
                "Spreadsheet ID",
                placeholder="1ge6-Rzor5jbQ7zaaQk3B7I0Vx31Nv80QH6zW2NfBUz8",
                help="Get the ID from the spreadsheet URL"
            )
            
            if st.button("📥 Add Spreadsheet", use_container_width=True):
                if manual_id:
                    with st.spinner("Opening spreadsheet..."):
                        wb = manager.open_workbook_by_id(manual_id.strip())
                        if wb:
                            st.success(f"✅ Added: {wb.title}")
                            st.rerun()
                        else:
                            st.error("❌ Failed to open spreadsheet")
                else:
                    st.warning("⚠️ Please enter a spreadsheet ID")
        
        st.markdown("---")
        
        if st.session_state.workbooks:
            st.success(f"📊 {len(st.session_state.workbooks)} workbook(s) available")
            
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
            st.warning("⚠️ No workbooks loaded")
            st.info("Click 'Refresh' or add a spreadsheet ID manually above")
        
        st.markdown("---")
        
        if st.session_state.service_account_email:
            st.markdown(f"""
            <div class="sidebar-card">
                <h3>🟢 Connected</h3>
                <p class="highlight">Service Account</p>
                <p style="font-size: 0.75rem; word-break: break-all;">{st.session_state.service_account_email}</p>
                <p class="highlight">Workbooks: {len(st.session_state.workbooks)}</p>
            </div>
            """, unsafe_allow_html=True)

    view_mode = st.radio(
        "📑 View Mode",
        ["All Clients", "Dashboard", "All Sheets Viewer", "Calendar View", "Booking Manager", "System Logs"],
        help="Choose what to display"
    )
    
    st.markdown("---")
    
    if view_mode == "All Clients":
        render_all_clients(manager)
    elif view_mode == "Dashboard":
        # Check if workbook is selected
        if not st.session_state.current_workbook:
            st.info("👈 Please select a workbook from the sidebar to view dashboard")
            return
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("❌ Failed to load workbook")
            return
        render_dashboard(manager, workbook)
    elif view_mode == "All Sheets Viewer":
        if not st.session_state.current_workbook:
            st.info("👈 Please select a workbook from the sidebar to view sheets")
            return
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("❌ Failed to load workbook")
            return
        render_all_sheets_viewer(manager, workbook)
    elif view_mode == "Calendar View":
        if not st.session_state.current_workbook:
            st.info("👈 Please select a workbook from the sidebar to view calendar")
            return
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("❌ Failed to load workbook")
            return
        render_calendar_view(manager, workbook)
    elif view_mode == "Booking Manager":
        if not st.session_state.current_workbook:
            st.info("👈 Please select a workbook from the sidebar to manage bookings")
            return
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("❌ Failed to load workbook")
            return
        render_booking_manager(manager, workbook)
    elif view_mode == "System Logs":
        render_system_logs()

def render_all_clients(manager):
    """Render comprehensive view of all clients/workbooks"""
    st.markdown('<div class="section-header">👥 All Clients Overview</div>', unsafe_allow_html=True)
    
    if not st.session_state.workbooks:
        st.warning("⚠️ No workbooks found. Click 'Refresh' in the sidebar to scan for spreadsheets.")
        
        st.markdown(f"""
        <div class="info-box">
            <h3>📁 Folder Configuration</h3>
            <p><strong>Folder ID:</strong> {DRIVE_FOLDER_ID}</p>
            <p><strong>Service Account:</strong> {st.session_state.service_account_email}</p>
            <p><a href="https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}" target="_blank">🔗 Open Folder in Google Drive</a></p>
        </div>
        """, unsafe_allow_html=True)
        return
    
    # Summary metrics at the top
    st.markdown("### 📊 Portfolio Summary")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #667eea; margin: 0;">📚 Total Clients</h3>
            <h1 style="margin: 0.5rem 0;">{len(st.session_state.workbooks)}</h1>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #11998e; margin: 0;">📁 Active Folder</h3>
            <p style="margin: 0.5rem 0; font-size: 0.9rem;">Connected</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #f093fb; margin: 0;">🔄 Last Sync</h3>
            <p style="margin: 0.5rem 0; font-size: 0.9rem;">{datetime.now().strftime("%H:%M")}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        if st.button("🔄 Refresh All", use_container_width=True, type="primary"):
            with st.spinner("Refreshing all clients..."):
                st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
            st.success(f"✅ Refreshed! Found {len(st.session_state.workbooks)} clients")
            st.rerun()
    
    st.markdown("---")
    
    # Display each client as an enhanced card
    st.markdown("### 🏢 Client Details")
    
    for idx, workbook_info in enumerate(st.session_state.workbooks):
        with st.spinner(f"Loading {workbook_info['name']}..."):
            try:
                # Open workbook to get details
                wb = manager.open_workbook(workbook_info['id'])
                if not wb:
                    continue
                
                # Get client profile
                profile = manager.get_client_profile(wb)
                
                # Get calendar sheets
                calendars = manager.get_calendar_sheets(wb)
                
                # Calculate total bookings across all calendars
                total_bookings = 0
                total_properties = len(profile.get('properties', []))
                
                for cal in calendars:
                    df = manager.read_calendar(cal['sheet'])
                    if not df.empty:
                        total_bookings += len(df)
                
                # Create expandable client card
                with st.expander(f"🏢 {profile.get('client_name', workbook_info['name'])} - {total_properties} Properties, {len(calendars)} Calendars", expanded=(idx == 0)):
                    
                    # Client header with key info
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4 style="color: #667eea; margin: 0;">🏘️ Properties</h4>
                            <h2 style="margin: 0.5rem 0;">{total_properties}</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4 style="color: #11998e; margin: 0;">📅 Calendars</h4>
                            <h2 style="margin: 0.5rem 0;">{len(calendars)}</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4 style="color: #f093fb; margin: 0;">📋 Bookings</h4>
                            <h2 style="margin: 0.5rem 0;">{total_bookings}</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        if st.button(f"📊 Open Dashboard", key=f"open_{workbook_info['id']}", use_container_width=True):
                            st.session_state.current_workbook = workbook_info['id']
                            st.rerun()
                    
                    st.markdown("---")
                    
                    # Client details section
                    col_left, col_right = st.columns([1, 1])
                    
                    with col_left:
                        st.markdown(f"""
                        <div class="info-box">
                            <h4 style="margin-top: 0;">⏰ Check-in/out Times</h4>
                            <p><strong>Check-out:</strong> {profile.get('check_out_time', 'N/A')}</p>
                            <p><strong>Check-in:</strong> {profile.get('check_in_time', 'N/A')}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.markdown(f"""
                        <div class="info-box">
                            <h4 style="margin-top: 0;">🔑 Access Information</h4>
                            <p><strong>Keys:</strong> {profile.get('keys', 'N/A')}</p>
                            <p><strong>Codes:</strong> {profile.get('codes', 'N/A')}</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col_right:
                        st.markdown(f"""
                        <div class="info-box">
                            <h4 style="margin-top: 0;">🧺 Services</h4>
                            <p><strong>Amenities:</strong> {profile.get('amenities', 'N/A')}</p>
                            <p><strong>Laundry:</strong> {profile.get('laundry_services', 'N/A')}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.markdown(f"""
                        <div class="info-box">
                            <h4 style="margin-top: 0;">🔗 Spreadsheet</h4>
                            <p><a href="{workbook_info['url']}" target="_blank">📊 Open in Google Sheets</a></p>
                            <p style="font-size: 0.8rem; word-break: break-all;"><strong>ID:</strong> {workbook_info['id']}</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Properties section
                    if profile.get('properties'):
                        st.markdown("#### 🏘️ Property Portfolio")
                        
                        # Display properties in a grid
                        prop_cols = st.columns(min(len(profile['properties']), 3))
                        
                        for prop_idx, prop in enumerate(profile['properties']):
                            with prop_cols[prop_idx % 3]:
                                st.markdown(f"""
                                <div class="property-card">
                                    <h4 style="margin-top: 0; color: #667eea;">🏠 {prop['name']}</h4>
                                    <p style="font-size: 0.9rem;"><strong>📍 Address:</strong><br>{prop['address']}</p>
                                    <p style="font-size: 0.9rem;"><strong>⏱️ Hours:</strong> {prop['hours']}</p>
                                    <p style="font-size: 0.9rem;"><strong>🔄 SO Hours:</strong> {prop['so_hours']}</p>
                                </div>
                                """, unsafe_allow_html=True)
                    
                    # Calendar sheets section
                    if calendars:
                        st.markdown("#### 📅 Calendar Sheets")
                        
                        cal_cols = st.columns(min(len(calendars), 4))
                        
                        for cal_idx, cal in enumerate(calendars):
                            with cal_cols[cal_idx % 4]:
                                df = manager.read_calendar(cal['sheet'])
                                booking_count = len(df) if not df.empty else 0
                                
                                st.markdown(f"""
                                <div class="metric-card">
                                    <h4 style="margin: 0; color: #11998e;">📆 {cal['name']}</h4>
                                    <h3 style="margin: 0.5rem 0;">{booking_count}</h3>
                                    <p style="margin: 0; font-size: 0.85rem;">bookings</p>
                                </div>
                                """, unsafe_allow_html=True)
                    
                    st.markdown("---")
                    
            except Exception as e:
                st.error(f"❌ Error loading {workbook_info['name']}: {str(e)}")
                add_log(f"Error loading workbook {workbook_info['name']}: {str(e)}", "ERROR")

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
    """Render booking manager with live editing capabilities"""
    st.markdown('<div class="section-header">✏️ Booking Manager - Live Edit</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar = st.selectbox("Select Calendar to Edit", calendar_names)
    
    selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
    
    if selected_cal:
        sheet = selected_cal['sheet']
        
        tab1, tab2, tab3 = st.tabs(["📝 Edit Data", "➕ Add Booking", "📋 Copy/Append"])
        
        with tab1:
            st.markdown("### Edit Existing Data")
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                st.dataframe(df, use_container_width=True)
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    edit_row = st.number_input("Row Number (from row 13)", min_value=13, value=13)
                with col2:
                    edit_col = st.number_input("Column Number", min_value=1, value=1)
                with col3:
                    new_value = st.text_input("New Value")
                
                if st.button("💾 Update Cell", type="primary"):
                    if manager.update_cell_live(sheet, edit_row, edit_col, new_value):
                        st.success("✅ Cell updated successfully!")
                        st.rerun()
                    else:
                        st.error("❌ Failed to update cell")
            else:
                st.info("No data to edit")
        
        with tab2:
            st.markdown("### Add New Booking")
            
            df = manager.read_calendar(sheet)
            if not df.empty:
                num_cols = len(df.columns)
                
                st.info(f"This sheet has {num_cols} columns. Fill in the values below:")
                
                new_row_data = []
                cols = st.columns(min(num_cols, 4))
                
                for i in range(num_cols):
                    with cols[i % 4]:
                        col_name = df.columns[i] if i < len(df.columns) else f"Column {i+1}"
                        value = st.text_input(f"{col_name}", key=f"new_col_{i}")
                        new_row_data.append(value)
                
                if st.button("➕ Add Booking Row", type="primary"):
                    if manager.append_row(sheet, new_row_data):
                        st.success("✅ Booking added successfully!")
                        st.rerun()
                    else:
                        st.error("❌ Failed to add booking")
            else:
                st.warning("Cannot add booking - sheet structure unknown")
        
        with tab3:
            st.markdown("### Copy and Append Rows")
            
            df = manager.read_calendar(sheet)
            if not df.empty:
                st.dataframe(df, use_container_width=True)
                
                row_to_copy = st.number_input(
                    "Row Number to Copy (from row 13)", 
                    min_value=13, 
                    value=13,
                    help="Enter the row number you want to copy"
                )
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if st.button("📋 Copy Row", type="secondary"):
                        copied_data = manager.copy_row(sheet, row_to_copy - 1)
                        if copied_data:
                            st.session_state['copied_row'] = copied_data
                            st.success(f"✅ Row {row_to_copy} copied! ({len(copied_data)} cells)")
                        else:
                            st.error("❌ Failed to copy row")
                
                with col2:
                    if st.button("➕ Append Copied Row", type="primary"):
                        if 'copied_row' in st.session_state:
                            if manager.append_row(sheet, st.session_state['copied_row']):
                                st.success("✅ Row appended successfully!")
                                st.rerun()
                            else:
                                st.error("❌ Failed to append row")
                        else:
                            st.warning("⚠️ No row copied yet. Copy a row first!")
                
                if 'copied_row' in st.session_state:
                    st.markdown("### 📋 Copied Row Preview")
                    st.json(st.session_state['copied_row'])
            else:
                st.info("No data available to copy")

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
