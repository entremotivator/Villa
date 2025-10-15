import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, timedelta
import json
from typing import List, Dict, Optional, Callable
import re
import logging
from io import StringIO
import time
from functools import wraps

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
    page_title="Professional Booking Management System - Enhanced",
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
    .progress-box {
        background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #2196F3;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: #000000;
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
if 'workbooks_cache' not in st.session_state:
    st.session_state.workbooks_cache = {}
if 'cache_timestamp' not in st.session_state:
    st.session_state.cache_timestamp = None
if 'loading_progress' not in st.session_state:
    st.session_state.loading_progress = {'current': 0, 'total': 0, 'status': ''}

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

def retry_with_backoff(max_retries: int = 3, initial_delay: float = 1.0):
    """Decorator for retrying functions with exponential backoff"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            delay = initial_delay
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise
                    add_log(f"Attempt {attempt + 1} failed: {str(e)}. Retrying in {delay}s...", "WARNING")
                    time.sleep(delay)
                    delay *= 2  # Exponential backoff
            return None
        return wrapper
    return decorator

class EnhancedBookingManager:
    """Enhanced Booking Manager with improved file loading and caching"""
    
    def __init__(self, credentials_dict: Dict):
        """Initialize with service account credentials"""
        try:
            add_log("Initializing Enhanced Booking Manager...", "INFO")
            
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
            
            # Extract service account email
            st.session_state.service_account_email = credentials_dict.get('client_email', 'Unknown')
            add_log(f"Service Account: {st.session_state.service_account_email}", "SUCCESS")
            
            # Cache settings
            self.cache_ttl = 300  # 5 minutes
            
            add_log("Enhanced Booking Manager initialized successfully", "SUCCESS")
            
        except Exception as e:
            add_log(f"Failed to initialize Enhanced Booking Manager: {str(e)}", "ERROR")
            raise
    
    def _is_cache_valid(self, cache_key: str) -> bool:
        """Check if cache is still valid"""
        if cache_key not in st.session_state.workbooks_cache:
            return False
        
        if st.session_state.cache_timestamp is None:
            return False
        
        elapsed = (datetime.now() - st.session_state.cache_timestamp).total_seconds()
        return elapsed < self.cache_ttl
    
    def _update_cache(self, cache_key: str, data: List[Dict]):
        """Update cache with new data"""
        st.session_state.workbooks_cache[cache_key] = data
        st.session_state.cache_timestamp = datetime.now()
        add_log(f"Cache updated for key: {cache_key}", "INFO")
    
    def _update_progress(self, current: int, total: int, status: str):
        """Update loading progress"""
        st.session_state.loading_progress = {
            'current': current,
            'total': total,
            'status': status
        }
    
    @retry_with_backoff(max_retries=3, initial_delay=2.0)
    def _fetch_drive_files_page(self, drive_service, query: str, page_token: Optional[str] = None):
        """Fetch a single page of files from Drive API with retry logic"""
        return drive_service.files().list(
            q=query,
            pageSize=1000,  # Maximum page size
            fields="nextPageToken, files(id, name, webViewLink, modifiedTime, createdTime, owners, size, mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageToken=page_token
        ).execute()
    
    def list_workbooks_from_folder_enhanced(
        self, 
        folder_id: str,
        use_cache: bool = True,
        force_refresh: bool = False,
        recursive: bool = False,
        max_depth: int = 3,
        progress_callback: Optional[Callable] = None
    ) -> List[Dict]:
        """Enhanced version with complete loading, caching, and progress tracking"""
        try:
            cache_key = f"folder_{folder_id}_recursive_{recursive}"
            
            # Check cache first
            if use_cache and not force_refresh and self._is_cache_valid(cache_key):
                add_log(f"Using cached data for folder: {folder_id}", "INFO")
                cached_data = st.session_state.workbooks_cache.get(cache_key, [])
                add_log(f"Loaded {len(cached_data)} workbook(s) from cache", "SUCCESS")
                return cached_data
            
            add_log(f"Listing workbooks from folder: {folder_id} (Enhanced Mode)", "INFO")
            add_log(f"Settings: Cache={use_cache}, ForceRefresh={force_refresh}, Recursive={recursive}", "INFO")
            
            workbooks = []
            
            # Method 1: Drive API with enhanced pagination
            try:
                from googleapiclient.discovery import build
                from googleapiclient.errors import HttpError
                add_log("Attempting Enhanced Drive API method...", "INFO")
                
                drive_service = build('drive', 'v3', credentials=self.creds)
                
                # Query for all spreadsheets in the folder
                query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
                
                page_token = None
                page_count = 0
                total_files = 0
                max_pages = 100  # Safety limit
                
                add_log("Starting pagination loop (max 1000 files per page)...", "INFO")
                
                while page_count < max_pages:
                    page_count += 1
                    add_log(f"Fetching page {page_count}...", "INFO")
                    
                    self._update_progress(total_files, -1, f"Loading page {page_count}...")
                    
                    # Fetch page with retry logic
                    results = self._fetch_drive_files_page(drive_service, query, page_token)
                    
                    files = results.get('files', [])
                    add_log(f"Page {page_count}: Found {len(files)} file(s)", "INFO")
                    
                    for file in files:
                        total_files += 1
                        workbooks.append({
                            'id': file['id'],
                            'name': file['name'],
                            'url': file.get('webViewLink', f"https://docs.google.com/spreadsheets/d/{file['id']}"),
                            'modified': file.get('modifiedTime', 'Unknown'),
                            'created': file.get('createdTime', 'Unknown'),
                            'size': file.get('size', 'Unknown'),
                            'owners': file.get('owners', []),
                            'mimeType': file.get('mimeType', 'Unknown')
                        })
                        
                        if total_files % 10 == 0:
                            add_log(f"  Loaded {total_files} files so far...", "INFO")
                            self._update_progress(total_files, -1, f"Loaded {total_files} files...")
                    
                    page_token = results.get('nextPageToken')
                    
                    if not page_token:
                        add_log(f"No more pages. Total pages processed: {page_count}", "SUCCESS")
                        break
                    
                    add_log(f"More pages available, continuing...", "INFO")
                    time.sleep(0.5)  # Rate limiting
                
                if page_count >= max_pages:
                    add_log(f"Reached maximum page limit ({max_pages}). There may be more files.", "WARNING")
                
                if workbooks:
                    add_log(f"Drive API method successful: {len(workbooks)} workbook(s) loaded across {page_count} page(s)", "SUCCESS")
                    self._update_cache(cache_key, workbooks)
                    self._update_progress(len(workbooks), len(workbooks), "Complete")
                    return workbooks
                    
            except ImportError:
                add_log("google-api-python-client not installed, trying alternative methods", "WARNING")
            except Exception as e:
                add_log(f"Drive API method failed: {str(e)}", "WARNING")
            
            # Method 2: Try example spreadsheet IDs
            add_log("Trying to open example spreadsheets from folder...", "INFO")
            for sheet_id in EXAMPLE_SPREADSHEET_IDS:
                try:
                    sheet = self.gc.open_by_key(sheet_id)
                    workbooks.append({
                        'id': sheet.id,
                        'name': sheet.title,
                        'url': sheet.url,
                        'modified': 'Unknown',
                        'created': 'Unknown',
                        'size': 'Unknown',
                        'owners': [],
                        'mimeType': 'application/vnd.google-apps.spreadsheet'
                    })
                    add_log(f"  ✓ Opened: {sheet.title}", "SUCCESS")
                except Exception as e:
                    add_log(f"  ✗ Could not open {sheet_id}: {str(e)}", "WARNING")
            
            if workbooks:
                add_log(f"Example spreadsheet method successful: {len(workbooks)} workbook(s)", "SUCCESS")
                self._update_cache(cache_key, workbooks)
                return workbooks
            
            # Method 3: List all accessible spreadsheets
            add_log("Trying to list all accessible spreadsheets...", "INFO")
            try:
                all_spreadsheets = self.gc.openall()
                add_log(f"Found {len(all_spreadsheets)} accessible spreadsheet(s)", "INFO")
                
                for i, sheet in enumerate(all_spreadsheets, 1):
                    workbooks.append({
                        'id': sheet.id,
                        'name': sheet.title,
                        'url': sheet.url,
                        'modified': 'Unknown',
                        'created': 'Unknown',
                        'size': 'Unknown',
                        'owners': [],
                        'mimeType': 'application/vnd.google-apps.spreadsheet'
                    })
                    add_log(f"  ✓ {sheet.title}", "INFO")
                    
                    if i % 10 == 0:
                        self._update_progress(i, len(all_spreadsheets), f"Loading {i}/{len(all_spreadsheets)}...")
                
                if workbooks:
                    add_log(f"List all method successful: {len(workbooks)} workbook(s)", "SUCCESS")
                    self._update_cache(cache_key, workbooks)
                    self._update_progress(len(workbooks), len(workbooks), "Complete")
                    return workbooks
                    
            except Exception as e:
                add_log(f"List all method failed: {str(e)}", "ERROR")
            
            # No workbooks found
            add_log("No workbooks found with any method", "ERROR")
            add_log(f"Please ensure folder {folder_id} is shared with: {st.session_state.service_account_email}", "WARNING")
            add_log("Or add spreadsheet IDs manually below", "INFO")
            
            return []
                
        except Exception as e:
            add_log(f"Error listing workbooks: {str(e)}", "ERROR")
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
                    'modified': 'Unknown',
                    'created': 'Unknown',
                    'size': 'Unknown',
                    'owners': [],
                    'mimeType': 'application/vnd.google-apps.spreadsheet'
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
                    'sheet': sheet,
                    'id': sheet.id
                }
                for i, sheet in enumerate(all_sheets)
            ]
            
            add_log(f"Found {len(all_sheets)} sheet(s) in workbook", "SUCCESS")
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
    
    def clear_cache(self):
        """Clear the workbooks cache"""
        st.session_state.workbooks_cache = {}
        st.session_state.cache_timestamp = None
        add_log("Cache cleared successfully", "SUCCESS")

def authenticate():
    """Authentication section with JSON file upload only"""
    st.markdown('<div class="main-header">🏠 Professional Booking Management System - Enhanced</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown("""
        <div class="info-box">
            <h3 style="margin-top: 0;">Welcome to Your Enhanced Booking Hub</h3>
            <p>Manage multiple client properties, booking calendars, and service schedules all in one place. 
            Upload your Google service account JSON file to get started.</p>
            <p><strong>✨ New Features:</strong> Complete file loading, smart caching, progress tracking, and retry logic!</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="drive-config-card">
            <h3 style="color: white; margin-top: 0;">📁 Google Drive Folder</h3>
            <p style="color: white;"><strong>Folder ID:</strong> {DRIVE_FOLDER_ID}</p>
            <p style="color: white; font-size: 0.9rem;">All workbooks will be loaded from this folder automatically with enhanced pagination.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-card">
            <h3>🚀 Quick Setup</h3>
            
            <p class="highlight">Step 1: Service Account</p>
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
    
    col1, col2 = st.columns(2)
    
    with col1:
        use_cache = st.checkbox("Use Cache (5 min)", value=True, help="Use cached results for faster loading")
    
    with col2:
        force_refresh = st.checkbox("Force Refresh", value=False, help="Ignore cache and reload all files")
    
    if st.button("🚀 Connect to Google Sheets", type="primary", use_container_width=True):
        if uploaded_file:
            with st.spinner("Authenticating..."):
                try:
                    add_log(f"Reading credentials from file: {uploaded_file.name}", "INFO")
                    creds_dict = json.load(uploaded_file)
                    add_log("Credentials file loaded successfully", "SUCCESS")
                    
                    manager = EnhancedBookingManager(creds_dict)
                    st.session_state.gc = manager
                    st.session_state.authenticated = True
                    
                    st.success("✅ Successfully connected to Google Sheets!")
                    
                    progress_placeholder = st.empty()
                    
                    with st.spinner(f"Loading workbooks from folder {DRIVE_FOLDER_ID}..."):
                        add_log("Starting enhanced workbook discovery...", "INFO")
                        
                        # Show initial progress
                        progress_placeholder.markdown("""
                        <div class="progress-box">
                            <p><strong>📊 Loading workbooks...</strong></p>
                            <p>Scanning folder with enhanced pagination (up to 1000 files per page)</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.session_state.workbooks = manager.list_workbooks_from_folder_enhanced(
                            DRIVE_FOLDER_ID,
                            use_cache=use_cache,
                            force_refresh=force_refresh
                        )
                        
                        progress = st.session_state.loading_progress
                        if progress['total'] > 0:
                            progress_placeholder.markdown(f"""
                            <div class="success-box">
                                <p><strong>✅ Loading complete!</strong></p>
                                <p>Loaded {progress['current']} of {progress['total']} files</p>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    if len(st.session_state.workbooks) > 0:
                        st.success(f"✅ Found {len(st.session_state.workbooks)} workbook(s)!")
                        st.balloons()
                    else:
                        st.warning("⚠️ No workbooks found in folder")
                        st.info(f"Make sure folder {DRIVE_FOLDER_ID} is shared with: {st.session_state.service_account_email}")
                    
                    st.rerun()
                    
                except Exception as e:
                    add_log(f"Authentication error: {str(e)}", "ERROR")
                    st.error(f"❌ Authentication failed: {str(e)}")
        else:
            st.warning("⚠️ Please upload a credentials file")

def main_app():
    """Main application interface"""
    st.markdown('<div class="main-header">🏠 Booking Management Dashboard - Enhanced</div>', unsafe_allow_html=True)
    
    manager = st.session_state.gc
    
    with st.sidebar:
        st.markdown("### 🎛️ Control Panel")
        
        st.markdown(f"""
        <div class="drive-config-card">
            <h3 style="color: white; margin-top: 0;">📁 Active Folder</h3>
            <p style="color: white; font-size: 0.85rem; word-break: break-all;">{DRIVE_FOLDER_ID}</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            force_refresh = st.checkbox("Force Refresh", value=False, key="main_force_refresh")
        
        with col2:
            if st.button("🗑️ Clear Cache"):
                manager.clear_cache()
                st.success("Cache cleared!")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Refresh", use_container_width=True):
                add_log("Refreshing workbook list...", "INFO")
                with st.spinner("Refreshing..."):
                    st.session_state.workbooks = manager.list_workbooks_from_folder_enhanced(
                        DRIVE_FOLDER_ID,
                        force_refresh=force_refresh
                    )
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
                manager.clear_cache()
                add_log("Logged out successfully", "INFO")
                st.rerun()
        
        st.markdown("---")
        
        with st.expander("➕ Add Spreadsheet Manually"):
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
        
        # Display cache status
        if st.session_state.cache_timestamp:
            elapsed = (datetime.now() - st.session_state.cache_timestamp).total_seconds()
            cache_age = f"{int(elapsed)}s ago"
            cache_valid = elapsed < manager.cache_ttl
            
            st.markdown(f"""
            <div class="{'success-box' if cache_valid else 'warning-box'}">
                <p style="margin: 0;"><strong>Cache Status:</strong></p>
                <p style="margin: 0;">{'✅ Valid' if cache_valid else '⚠️ Expired'}</p>
                <p style="margin: 0; font-size: 0.85rem;">Updated: {cache_age}</p>
            </div>
            """, unsafe_allow_html=True)
        
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
                    <p><strong>📅 Modified:</strong><br>{selected_workbook.get('modified', 'Unknown')}</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.warning("⚠️ No workbooks loaded")
            st.info("Click 'Refresh' or add a spreadsheet ID manually above")
        
        st.markdown("---")
        
        # View mode selector
        view_mode = st.radio(
            "📑 View Mode",
            ["Dashboard", "All Sheets Viewer", "Calendar View", "Booking Manager", "System Logs"],
            help="Select the view mode"
        )

    # Main content area
    if not st.session_state.current_workbook:
        st.info("👈 Please select a workbook from the sidebar to begin")
        
        # Show loading statistics
        if st.session_state.workbooks:
            st.markdown("### 📊 Loaded Workbooks Statistics")
            
            total_workbooks = len(st.session_state.workbooks)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Workbooks", total_workbooks)
            
            with col2:
                cached = "Yes" if st.session_state.cache_timestamp else "No"
                st.metric("Using Cache", cached)
            
            with col3:
                progress = st.session_state.loading_progress
                st.metric("Last Load Status", progress.get('status', 'N/A'))
            
            # Show workbook list
            if st.checkbox("Show All Workbooks"):
                df_workbooks = pd.DataFrame(st.session_state.workbooks)
                st.dataframe(df_workbooks, use_container_width=True)
        
        return
    
    # Load current workbook
    workbook = manager.open_workbook(st.session_state.current_workbook)
    if not workbook:
        st.error("❌ Failed to load workbook")
        return
    
    # Connection status in sidebar
    with st.sidebar:
        st.markdown(f"""
        <div class="sidebar-card">
            <h3>✅ Connected</h3>
            <p class="highlight">Service Account</p>
            <p style="font-size: 0.75rem; word-break: break-all;">{st.session_state.service_account_email}</p>
            <p class="highlight">Workbooks: {len(st.session_state.workbooks)}</p>
            <p class="highlight">Sheets: {len(st.session_state.all_sheets)}</p>
        </div>
        """, unsafe_allow_html=True)
    
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
    
    st.info(f"📊 This workbook contains {len(st.session_state.all_sheets)} sheet(s)")
    
    with st.container():
        sheet_names = [s['name'] for s in st.session_state.all_sheets]
        selected_sheet_name = st.selectbox(
            "Select Sheet to View",
            sheet_names,
            index=st.session_state.selected_sheet_index,
            help="Choose any sheet from the workbook"
        )
        
        # Update selected index
        selected_index = sheet_names.index(selected_sheet_name)
        if selected_index != st.session_state.selected_sheet_index:
            st.session_state.selected_sheet_index = selected_index
        
        selected_sheet_info = st.session_state.all_sheets[selected_index]
        selected_sheet = selected_sheet_info['sheet']
        
        st.markdown(f"""
        <div class="info-box">
            <p><strong>📄 Current Sheet:</strong> {selected_sheet_info['name']}</p>
            <p><strong>🔢 Sheet Index:</strong> {selected_sheet_info['index']}</p>
            <p><strong>🆔 Sheet ID:</strong> {selected_sheet_info['id']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Read data
        with st.spinner(f"Loading data from '{selected_sheet_info['name']}'..."):
            df = manager.read_sheet_all_data(selected_sheet)
        
        if not df.empty:
            st.success(f"✅ Loaded {len(df)} rows and {len(df.columns)} columns")
            
            # Display options
            col1, col2, col3 = st.columns(3)
            
            with col1:
                show_full = st.checkbox("Show Full Data", value=True)
            
            with col2:
                max_rows = st.number_input("Max Rows to Display", min_value=10, max_value=10000, value=100, step=10)
            
            with col3:
                # Export option
                csv = df.to_csv(index=False)
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv,
                    file_name=f"{selected_sheet_info['name']}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            # Display data
            st.markdown("### 📊 Sheet Data")
            
            if show_full:
                st.dataframe(df.head(max_rows), use_container_width=True)
                if len(df) > max_rows:
                    st.info(f"Showing first {max_rows} of {len(df)} rows. Increase 'Max Rows' to see more.")
            else:
                st.dataframe(df.head(20), use_container_width=True)
            
            # Column info
            with st.expander("📋 Column Information"):
                col_info = pd.DataFrame({
                    'Column Name': df.columns,
                    'Data Type': df.dtypes.astype(str),
                    'Non-Empty Count': df.count().values,
                    'Sample Value': [df[col].iloc[0] if len(df) > 0 else '' for col in df.columns]
                })
                st.dataframe(col_info, use_container_width=True)
        else:
            st.warning("⚠️ No data found in this sheet")

def render_dashboard(manager, workbook):
    """Render dashboard view"""
    st.markdown('<div class="section-header">📊 Dashboard Overview</div>', unsafe_allow_html=True)
    
    # Get client profile
    profile = manager.get_client_profile(workbook)
    
    # Display client info
    st.markdown(f"""
    <div class="info-box">
        <h3>{profile.get('client_name', 'Unknown Client')}</h3>
        <p><strong>Check-out Time:</strong> {profile.get('check_out_time', 'N/A')}</p>
        <p><strong>Check-in Time:</strong> {profile.get('check_in_time', 'N/A')}</p>
        <p><strong>Amenities:</strong> {profile.get('amenities', 'N/A')}</p>
        <p><strong>Laundry Services:</strong> {profile.get('laundry_services', 'N/A')}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Display properties
    if profile.get('properties'):
        st.markdown(f"### 🏘️ Properties ({len(profile['properties'])})")
        for prop in profile['properties']:
            st.markdown(f"""
            <div class="property-card">
                <h4>{prop['name']}</h4>
                <p><strong>Address:</strong> {prop['address']}</p>
                <p><strong>Hours:</strong> {prop['hours']}</p>
                <p><strong>SO Hours:</strong> {prop['so_hours']}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Calendar sheets
    calendars = manager.get_calendar_sheets(workbook)
    st.markdown(f"### 📅 Calendar Sheets ({len(calendars)} months)")
    
    for cal in calendars[:3]:  # Show first 3
        with st.expander(f"📅 {cal['name']}"):
            df = manager.read_calendar(cal['sheet'])
            if not df.empty:
                st.dataframe(df.head(10), use_container_width=True)
            else:
                st.info("No bookings found")

def render_calendar_view(manager, workbook):
    """Render calendar view"""
    st.markdown('<div class="section-header">📅 Calendar View</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    calendar_names = [c['name'] for c in calendars]
    selected_calendar = st.selectbox("Select Calendar Month", calendar_names)
    
    selected_cal = next((c for c in calendars if c['name'] == selected_calendar), None)
    
    if selected_cal:
        df = manager.read_calendar(selected_cal['sheet'])
        
        if not df.empty:
            st.dataframe(df, use_container_width=True)
            
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
    
    st.info("Select a calendar to manage bookings")
    
    calendar_names = [c['name'] for c in calendars]
    selected_calendar = st.selectbox("Select Calendar", calendar_names)
    
    selected_cal = next((c for c in calendars if c['name'] == selected_calendar), None)
    
    if selected_cal:
        df = manager.read_calendar(selected_cal['sheet'])
        
        if not df.empty:
            st.markdown("### Current Bookings")
            st.dataframe(df, use_container_width=True)
            
            st.markdown("### Quick Actions")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("➕ Add New Booking"):
                    st.info("Add booking functionality - implement as needed")
            
            with col2:
                if st.button("🔄 Refresh Data"):
                    st.rerun()
        else:
            st.info("No bookings found. Add your first booking!")

def render_system_logs():
    """Render system logs"""
    st.markdown('<div class="section-header">📋 System Logs</div>', unsafe_allow_html=True)
    
    if st.button("🗑️ Clear Logs"):
        st.session_state.logs = []
        st.rerun()
    
    if st.session_state.logs:
        st.info(f"Showing {len(st.session_state.logs)} log entries (max 100)")
        
        for log in reversed(st.session_state.logs):
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
        st.info("No logs yet")

# Main entry point
if __name__ == "__main__":
    if not st.session_state.authenticated:
        authenticate()
    else:
        main_app()

