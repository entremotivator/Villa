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

# Page Configuration
st.set_page_config(
    page_title="Professional Booking Management System - Fixed",
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
        color: #000000;
    }
    .property-card:hover {
        border-color: #667eea;
        box-shadow: 0 6px 12px rgba(102, 126, 234, 0.3);
    }
    .property-card h1, .property-card h2, .property-card h3, .property-card h4, .property-card p, .property-card strong {
        color: #000000 !important;
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
if 'discovered_sheets' not in st.session_state:
    st.session_state.discovered_sheets = []

def add_log(message: str, level: str = "INFO"):
    """Add a log entry with timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {
        'timestamp': timestamp,
        'level': level,
        'message': message
    }
    st.session_state.logs.append(log_entry)
    
    # Keep only last 200 logs
    if len(st.session_state.logs) > 200:
        st.session_state.logs = st.session_state.logs[-200:]
    
    # Also log to Python logger
    if level == "INFO":
        logger.info(message)
    elif level == "SUCCESS":
        logger.info(f"✓ {message}")
    elif level == "WARNING":
        logger.warning(message)
    elif level == "ERROR":
        logger.error(message)

def retry_with_backoff(max_retries: int = 5, initial_delay: float = 2.0):
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
                    add_log(f"Attempt {attempt + 1}/{max_retries} failed: {str(e)}. Retrying in {delay}s...", "WARNING")
                    time.sleep(delay)
                    delay *= 2  # Exponential backoff
            return None
        return wrapper
    return decorator

class EnhancedBookingManager:
    """Enhanced Booking Manager with improved file loading"""
    
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
        add_log(f"Cache updated for key: {cache_key} with {len(data)} items", "INFO")
    
    def _update_progress(self, current: int, total: int, status: str):
        """Update loading progress"""
        st.session_state.loading_progress = {
            'current': current,
            'total': total,
            'status': status
        }
    
    @retry_with_backoff(max_retries=5, initial_delay=2.0)
    def _fetch_drive_files_page(self, drive_service, query: str, page_token: Optional[str] = None):
        """Fetch a single page of files from Drive API with retry logic"""
        add_log(f"Fetching Drive API page (token: {page_token[:20] if page_token else 'None'}...)", "INFO")
        result = drive_service.files().list(
            q=query,
            pageSize=1000,  # Maximum page size
            fields="nextPageToken, files(id, name, webViewLink, modifiedTime, createdTime, owners, size, mimeType, parents)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageToken=page_token,
            corpora='allDrives'
        ).execute()
        add_log(f"Drive API returned {len(result.get('files', []))} files", "SUCCESS")
        return result
    
    def list_all_workbooks_from_folder(self, folder_id: str, force_refresh: bool = False) -> List[Dict]:
        """
        FIXED VERSION: List ALL spreadsheets from Google Drive folder
        This version ensures complete file loading with proper pagination
        """
        try:
            cache_key = f"folder_{folder_id}_all"
            
            # Check cache first
            if not force_refresh and self._is_cache_valid(cache_key):
                add_log(f"Using cached data for folder: {folder_id}", "INFO")
                cached_data = st.session_state.workbooks_cache.get(cache_key, [])
                add_log(f"Loaded {len(cached_data)} workbook(s) from cache", "SUCCESS")
                return cached_data
            
            add_log(f"=" * 60, "INFO")
            add_log(f"STARTING COMPLETE FOLDER SCAN", "INFO")
            add_log(f"Folder ID: {folder_id}", "INFO")
            add_log(f"Force Refresh: {force_refresh}", "INFO")
            add_log(f"=" * 60, "INFO")
            
            workbooks = []
            
            # Method 1: Drive API with COMPLETE pagination
            try:
                from googleapiclient.discovery import build
                from googleapiclient.errors import HttpError
                
                add_log("Building Drive API service...", "INFO")
                drive_service = build('drive', 'v3', credentials=self.creds)
                add_log("Drive API service built successfully", "SUCCESS")
                
                # Query for ALL spreadsheets in the folder
                query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
                add_log(f"Query: {query}", "INFO")
                
                page_token = None
                page_count = 0
                total_files = 0
                max_pages = 1000  # Increased safety limit
                
                add_log("Starting pagination loop...", "INFO")
                add_log(f"Max pages: {max_pages}, Max files per page: 1000", "INFO")
                
                while page_count < max_pages:
                    page_count += 1
                    add_log(f"", "INFO")
                    add_log(f">>> FETCHING PAGE {page_count} <<<", "INFO")
                    
                    self._update_progress(total_files, -1, f"Loading page {page_count}...")
                    
                    try:
                        # Fetch page with retry logic
                        results = self._fetch_drive_files_page(drive_service, query, page_token)
                        
                        files = results.get('files', [])
                        add_log(f"Page {page_count}: Received {len(files)} file(s) from API", "SUCCESS")
                        
                        if len(files) == 0:
                            add_log(f"Page {page_count} returned 0 files - stopping", "WARNING")
                            break
                        
                        # Process each file
                        for idx, file in enumerate(files, 1):
                            total_files += 1
                            
                            workbook_info = {
                                'id': file['id'],
                                'name': file['name'],
                                'url': file.get('webViewLink', f"https://docs.google.com/spreadsheets/d/{file['id']}"),
                                'modified': file.get('modifiedTime', 'Unknown'),
                                'created': file.get('createdTime', 'Unknown'),
                                'size': file.get('size', 'Unknown'),
                                'owners': file.get('owners', []),
                                'mimeType': file.get('mimeType', 'Unknown'),
                                'parents': file.get('parents', [])
                            }
                            
                            workbooks.append(workbook_info)
                            
                            # Log every file
                            add_log(f"  [{total_files}] {file['name']}", "SUCCESS")
                            
                            if total_files % 50 == 0:
                                add_log(f"  >>> Loaded {total_files} files so far...", "INFO")
                                self._update_progress(total_files, -1, f"Loaded {total_files} files...")
                        
                        # Check for next page
                        page_token = results.get('nextPageToken')
                        
                        if not page_token:
                            add_log(f"", "INFO")
                            add_log(f"No more pages found. Pagination complete.", "SUCCESS")
                            add_log(f"Total pages processed: {page_count}", "SUCCESS")
                            break
                        else:
                            add_log(f"Next page token found: {page_token[:30]}...", "INFO")
                            add_log(f"Continuing to next page...", "INFO")
                            time.sleep(0.5)  # Rate limiting
                    
                    except Exception as page_error:
                        add_log(f"Error on page {page_count}: {str(page_error)}", "ERROR")
                        add_log(f"Continuing with {total_files} files loaded so far...", "WARNING")
                        break
                
                if page_count >= max_pages:
                    add_log(f"Reached maximum page limit ({max_pages}). There may be more files.", "WARNING")
                
                add_log(f"", "INFO")
                add_log(f"=" * 60, "INFO")
                add_log(f"FOLDER SCAN COMPLETE", "SUCCESS")
                add_log(f"Total workbooks found: {len(workbooks)}", "SUCCESS")
                add_log(f"Total pages scanned: {page_count}", "SUCCESS")
                add_log(f"=" * 60, "INFO")
                
                if workbooks:
                    self._update_cache(cache_key, workbooks)
                    self._update_progress(len(workbooks), len(workbooks), "Complete")
                    return workbooks
                else:
                    add_log("No workbooks found via Drive API", "WARNING")
                    
            except ImportError as ie:
                add_log(f"google-api-python-client not installed: {str(ie)}", "ERROR")
                add_log("Install with: pip install google-api-python-client", "WARNING")
            except Exception as e:
                add_log(f"Drive API method failed: {str(e)}", "ERROR")
                import traceback
                add_log(f"Traceback: {traceback.format_exc()}", "ERROR")
            
            # Method 2: List all accessible spreadsheets (fallback)
            add_log("Trying fallback method: list all accessible spreadsheets...", "INFO")
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
                        'mimeType': 'application/vnd.google-apps.spreadsheet',
                        'parents': []
                    })
                    add_log(f"  [{i}] {sheet.title}", "INFO")
                    
                    if i % 10 == 0:
                        self._update_progress(i, len(all_spreadsheets), f"Loading {i}/{len(all_spreadsheets)}...")
                
                if workbooks:
                    add_log(f"Fallback method successful: {len(workbooks)} workbook(s)", "SUCCESS")
                    self._update_cache(cache_key, workbooks)
                    self._update_progress(len(workbooks), len(workbooks), "Complete")
                    return workbooks
                    
            except Exception as e:
                add_log(f"Fallback method failed: {str(e)}", "ERROR")
            
            # No workbooks found
            add_log("=" * 60, "ERROR")
            add_log("NO WORKBOOKS FOUND WITH ANY METHOD", "ERROR")
            add_log(f"Please verify:", "WARNING")
            add_log(f"1. Folder {folder_id} exists", "WARNING")
            add_log(f"2. Folder is shared with: {st.session_state.service_account_email}", "WARNING")
            add_log(f"3. Folder contains Google Sheets files", "WARNING")
            add_log(f"4. google-api-python-client is installed", "WARNING")
            add_log("=" * 60, "ERROR")
            
            return []
                
        except Exception as e:
            add_log(f"CRITICAL ERROR in list_all_workbooks_from_folder: {str(e)}", "ERROR")
            import traceback
            add_log(f"Full traceback: {traceback.format_exc()}", "ERROR")
            return []
    
    def open_workbook_by_id(self, workbook_id: str):
        """Open a specific workbook by ID"""
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
                    'mimeType': 'application/vnd.google-apps.spreadsheet',
                    'parents': []
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
    
    def clear_cache(self):
        """Clear the workbooks cache"""
        st.session_state.workbooks_cache = {}
        st.session_state.cache_timestamp = None
        add_log("Cache cleared successfully", "SUCCESS")

def authenticate():
    """Authentication section"""
    st.markdown('<div class="main-header">🏠 Booking Management System - FIXED VERSION</div>', unsafe_allow_html=True)
    
    with st.container():
        st.markdown("""
        <div class="success-box">
            <h3 style="margin-top: 0;">✅ FIXED: Complete Folder Scanning</h3>
            <p>This version will load <strong>ALL spreadsheets</strong> from your Google Drive folder with:</p>
            <ul>
                <li>✅ Complete pagination (1000 files per page)</li>
                <li>✅ Detailed logging of every file found</li>
                <li>✅ Enhanced error handling and retry logic</li>
                <li>✅ Progress tracking for large folders</li>
                <li>✅ Fallback methods if Drive API fails</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="drive-config-card">
            <h3 style="color: white; margin-top: 0;">📁 Google Drive Folder</h3>
            <p style="color: white;"><strong>Folder ID:</strong> {DRIVE_FOLDER_ID}</p>
            <p style="color: white; font-size: 0.9rem;">All workbooks will be loaded from this folder automatically.</p>
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
    
    force_refresh = st.checkbox("Force Refresh (ignore cache)", value=True, help="Always reload from Drive")
    
    if st.button("🚀 Connect & Load ALL Spreadsheets", type="primary", use_container_width=True):
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
                    
                    with st.spinner(f"Loading ALL workbooks from folder {DRIVE_FOLDER_ID}..."):
                        add_log("Starting COMPLETE workbook discovery...", "INFO")
                        
                        progress_placeholder.markdown("""
                        <div class="progress-box">
                            <p><strong>📊 Loading ALL workbooks...</strong></p>
                            <p>Scanning folder with complete pagination</p>
                            <p>Check System Logs for detailed progress</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.session_state.workbooks = manager.list_all_workbooks_from_folder(
                            DRIVE_FOLDER_ID,
                            force_refresh=force_refresh
                        )
                        
                        progress = st.session_state.loading_progress
                        progress_placeholder.markdown(f"""
                        <div class="success-box">
                            <p><strong>✅ Loading complete!</strong></p>
                            <p><strong>Total workbooks found: {len(st.session_state.workbooks)}</strong></p>
                            <p>Status: {progress.get('status', 'Complete')}</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    if len(st.session_state.workbooks) > 0:
                        st.success(f"✅ Found {len(st.session_state.workbooks)} workbook(s)!")
                        st.balloons()
                    else:
                        st.error("❌ No workbooks found in folder")
                        st.warning(f"Please check System Logs and verify folder {DRIVE_FOLDER_ID} is shared with: {st.session_state.service_account_email}")
                    
                    time.sleep(2)
                    st.rerun()
                    
                except Exception as e:
                    add_log(f"Authentication error: {str(e)}", "ERROR")
                    st.error(f"❌ Authentication failed: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
        else:
            st.warning("⚠️ Please upload a credentials file")

def main_app():
    """Main application interface"""
    st.markdown('<div class="main-header">🏠 Booking Management Dashboard</div>', unsafe_allow_html=True)
    
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
            if st.button("🔄 Reload ALL", use_container_width=True):
                add_log("Manual reload initiated...", "INFO")
                with st.spinner("Reloading ALL workbooks..."):
                    st.session_state.workbooks = manager.list_all_workbooks_from_folder(
                        DRIVE_FOLDER_ID,
                        force_refresh=True
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
        
        # Display workbook count
        if st.session_state.workbooks:
            st.success(f"📊 {len(st.session_state.workbooks)} workbook(s) loaded")
            
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
                    <p><strong>📅 Modified:</strong><br>{selected_workbook.get('modified', 'Unknown')[:19]}</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.warning("⚠️ No workbooks loaded")
            st.info("Click 'Reload ALL' to scan the folder")
        
        st.markdown("---")
        
        # View mode selector
        view_mode = st.radio(
            "📑 View Mode",
            ["Workbook List", "Dashboard", "All Sheets Viewer", "Calendar View", "System Logs"],
            help="Select the view mode"
        )

    # Main content area
    if view_mode == "Workbook List":
        st.markdown('<div class="section-header">📋 All Loaded Workbooks</div>', unsafe_allow_html=True)
        
        if st.session_state.workbooks:
            st.success(f"Total: {len(st.session_state.workbooks)} workbook(s)")
            
            # Create DataFrame
            df_workbooks = pd.DataFrame(st.session_state.workbooks)
            
            # Display
            st.dataframe(df_workbooks, use_container_width=True)
            
            # Download option
            csv = df_workbooks.to_csv(index=False)
            st.download_button(
                "📥 Download Workbook List",
                data=csv,
                file_name="workbooks_list.csv",
                mime="text/csv"
            )
        else:
            st.info("No workbooks loaded yet. Click 'Reload ALL' in the sidebar.")
    
    elif view_mode == "System Logs":
        st.markdown('<div class="section-header">📋 System Logs</div>', unsafe_allow_html=True)
        
        if st.button("🗑️ Clear Logs"):
            st.session_state.logs = []
            st.rerun()
        
        if st.session_state.logs:
            st.info(f"Showing {len(st.session_state.logs)} log entries")
            
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
    
    elif not st.session_state.current_workbook:
        st.info("👈 Please select a workbook from the sidebar to begin")
    
    else:
        # Load current workbook
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if not workbook:
            st.error("❌ Failed to load workbook")
            return
        
        if view_mode == "Dashboard":
            st.markdown('<div class="section-header">📊 Dashboard Overview</div>', unsafe_allow_html=True)
            profile = manager.get_client_profile(workbook)
            st.markdown(f"""
            <div class="info-box">
                <h3>{profile.get('client_name', 'Unknown Client')}</h3>
                <p><strong>Check-out Time:</strong> {profile.get('check_out_time', 'N/A')}</p>
                <p><strong>Check-in Time:</strong> {profile.get('check_in_time', 'N/A')}</p>
            </div>
            """, unsafe_allow_html=True)
        
        elif view_mode == "All Sheets Viewer":
            st.markdown('<div class="section-header">📄 All Sheets Viewer</div>', unsafe_allow_html=True)
            
            if st.session_state.all_sheets:
                sheet_names = [s['name'] for s in st.session_state.all_sheets]
                selected_sheet_name = st.selectbox("Select Sheet", sheet_names)
                
                selected_index = sheet_names.index(selected_sheet_name)
                selected_sheet_info = st.session_state.all_sheets[selected_index]
                selected_sheet = selected_sheet_info['sheet']
                
                df = manager.read_sheet_all_data(selected_sheet)
                
                if not df.empty:
                    st.success(f"✅ Loaded {len(df)} rows")
                    st.dataframe(df, use_container_width=True)
                else:
                    st.warning("No data in sheet")
        
        elif view_mode == "Calendar View":
            st.markdown('<div class="section-header">📅 Calendar View</div>', unsafe_allow_html=True)
            calendars = manager.get_calendar_sheets(workbook)
            
            if calendars:
                calendar_names = [c['name'] for c in calendars]
                selected_calendar = st.selectbox("Select Calendar", calendar_names)
                
                selected_cal = next((c for c in calendars if c['name'] == selected_calendar), None)
                if selected_cal:
                    df = manager.read_calendar(selected_cal['sheet'])
                    if not df.empty:
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.info("No bookings")
            else:
                st.warning("No calendar sheets found")

# Main entry point
if __name__ == "__main__":
    if not st.session_state.authenticated:
        authenticate()
    else:
        main_app()

