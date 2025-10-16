
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
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# --- Configuration --- #
DRIVE_FOLDER_ID = "1Fk5dJGkm5dNMZkfsITe5Lt9x-yCsBiF2" # Example Drive Folder ID
# EXAMPLE_SPREADSHEET_IDS = [
#     "1ge6-Rzor5jbQ7zaaQk3B7I0Vx31Nv80QH6zW2NfBUz8",
#     "1-I0lHMXrA16v07Qtc3BNqnjRux7PRf8fMxJhVlxAcME"
# ]

# Page Configuration
st.set_page_config(
    page_title="Professional Booking Management System",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for enhanced styling (keeping user's original styles and adding new ones)
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
        color: #000000;
        margin-top: 2rem;
        margin-bottom: 1.5rem;
        border-bottom: 3px solid #667eea;
        padding-bottom: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    /* Updated all boxes and cards to have black text */
    .info-box {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #667eea;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: #000000;
    }
    .info-box h1, .info-box h2, .info-box h3, .info-box h4, .info-box p, .info-box strong, .info-box span, .info-box div {
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
    .warning-box h1, .warning-box h2, .warning-box h3, .warning-box h4, .warning-box p, .warning-box strong, .warning-box span, .warning-box div {
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
    .success-box h1, .success-box h2, .success-box h3, .success-box h4, .success-box p, .success-box strong, .success-box span, .success-box div {
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
    .error-box h1, .error-box h2, .error-box h3, .error-box h4, .error-box p, .error-box strong, .error-box span, .error-box div {
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
    .metric-card h1, .metric-card h2, .metric-card h3, .metric-card h4, .metric-card p, .metric-card strong, .metric-card span, .metric-card div {
        color: #000000 !important;
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
    .property-card h1, .property-card h2, .property-card h3, .property-card h4, .property-card p, .property-card strong, .property-card span, .property-card div {
        color: #000000 !important;
    }
    /* Added new card styles for enhanced features */
    .feature-card {
        background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #2196F3;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        color: #000000;
    }
    .feature-card h1, .feature-card h2, .feature-card h3, .feature-card h4, .feature-card p, .feature-card strong, .feature-card span, .feature-card div {
        color: #000000 !important;
    }
    .analytics-card {
        background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #ff9800;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        color: #000000;
    }
    .analytics-card h1, .analytics-card h2, .analytics-card h3, .analytics-card h4, .analytics-card p, .analytics-card strong, .analytics-card span, .analytics-card div {
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
        color: #000000;
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
    .sheet-selector {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border: 2px solid #dee2e6;
        margin: 1rem 0;
    }
    /* Added search and filter card styles */
    .search-card {
        background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #9c27b0;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        color: #000000;
    }
    .search-card h1, .search-card h2, .search-card h3, .search-card h4, .search-card p, .search-card strong, .search-card span, .search-card div {
        color: #000000 !important;
    }
    /* New styles for profile page and booking example */
    .profile-card {
        background: linear-gradient(135deg, #e0f7fa 0%, #b2ebf2 100%);
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
        margin-bottom: 2rem;
        border: 2px solid #00bcd4;
        color: #000000;
    }
    .profile-card h2 {
        color: #006064 !important;
        font-size: 2rem;
        margin-bottom: 1rem;
    }
    .profile-card p {
        font-size: 1.1rem;
        line-height: 1.8;
        color: #263238 !important;
    }
    .profile-card .detail-label {
        font-weight: bold;
        color: #00838f;
    }
    .booking-form-card {
        background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
        margin-bottom: 2rem;
        border: 2px solid #388e3c;
        color: #000000;
    }
    .booking-form-card h2 {
        color: #1b5e20 !important;
        font-size: 2rem;
        margin-bottom: 1rem;
    }
    .booking-form-card label {
        color: #1b5e20 !important;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'gc' not in st.session_state:
    st.session_state.gc = None
if 'drive_service' not in st.session_state:
    st.session_state.drive_service = None
if 'workbooks' not in st.session_state:
    st.session_state.workbooks = []
if 'current_workbook_id' not in st.session_state:
    st.session_state.current_workbook_id = None
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
if 'search_term' not in st.session_state:
    st.session_state.search_term = ""
if 'filter_status' not in st.session_state:
    st.session_state.filter_status = "All"
if 'selected_properties' not in st.session_state:
    st.session_state.selected_properties = []
if 'uploaded_csv_data' not in st.session_state:
    st.session_state.uploaded_csv_data = None

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
    'NA': {
        'name': 'Not Applicable',
        'description': 'No specific status required',
        'color': '#607D8B'
    }
}

def add_log(message: str, level: str = "INFO"):
    """Add a log entry to the session state"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.logs.append({
        'timestamp': timestamp,
        'level': level,
        'message': message
    })
    logger.log(getattr(logging, level), message)

class GoogleSheetManager:
    """Manages interactions with Google Sheets and Google Drive."""
    def __init__(self):
        self.gc = None
        self.drive_service = None
        if st.session_state.authenticated:
            self.gc = st.session_state.gc
            self.drive_service = st.session_state.drive_service

    def authenticate_gspread(self, credentials_json: str):
        """Authenticates with Google Sheets API using service account credentials."""
        try:
            creds_dict = json.loads(credentials_json)
            scopes = [
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/drive.file',
            ]
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            self.gc = gspread.authorize(creds)
            self.drive_service = build('drive', 'v3', credentials=creds)
            st.session_state.gc = self.gc
            st.session_state.drive_service = self.drive_service
            st.session_state.authenticated = True
            st.session_state.service_account_email = creds_dict['client_email']
            add_log("Successfully authenticated with Google Sheets and Drive.", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Authentication failed: {e}", "ERROR")
            st.error(f"Authentication failed: {e}")
            return False

    def list_workbooks_from_folder(self, folder_id: str) -> List[Dict]:
        """Lists all Google Sheets workbooks within a specified Google Drive folder."""
        if not self.drive_service:
            add_log("Drive service not authenticated.", "ERROR")
            return []
        try:
            add_log(f"Listing workbooks in folder ID: {folder_id}", "INFO")
            query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
            results = self.drive_service.files().list(q=query, fields="files(id, name, modifiedTime, webViewLink)").execute()
            items = results.get('files', [])
            
            workbooks_info = []
            for item in items:
                workbooks_info.append({
                    'id': item['id'],
                    'name': item['name'],
                    'modified': item.get('modifiedTime', 'Unknown'),
                    'url': item.get('webViewLink', '#')
                })
            add_log(f"Found {len(workbooks_info)} workbooks in folder {folder_id}", "SUCCESS")
            st.session_state.workbooks = workbooks_info # Update session state
            return workbooks_info
        except Exception as e:
            add_log(f"Error listing workbooks from Drive folder: {e}", "ERROR")
            return []

    def open_workbook_by_id(self, workbook_id: str):
        """Opens a specific workbook by ID and stores its sheets in session state."""
        if not self.gc:
            add_log("GSpread client not authenticated.", "ERROR")
            return None
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
            st.session_state.current_workbook_id = workbook_id
            add_log(f"Workbook contains {len(all_sheets)} sheet(s)", "INFO")
            return workbook
        except Exception as e:
            add_log(f"Error opening workbook: {e}", "ERROR")
            return None

    def get_client_profile(self, workbook) -> Dict:
        """Extracts client profile from the first sheet of a workbook."""
        try:
            add_log(f"Reading client profile from: {workbook.title}", "INFO")
            sheet = workbook.get_worksheet(0)
            add_log(f"Accessing sheet: {sheet.title}", "INFO")
            
            all_values = sheet.get_all_values()
            profile = {
                'client_name': all_values[0][0] if len(all_values) > 0 and len(all_values[0]) > 0 else 'Unknown',
                'contact_person': all_values[1][1] if len(all_values) > 1 and len(all_values[1]) > 1 else '',
                'contact_email': all_values[2][1] if len(all_values) > 2 and len(all_values[2]) > 1 else '',
                'phone_number': all_values[3][1] if len(all_values) > 3 and len(all_values[3]) > 1 else '',
                'address': all_values[4][1] if len(all_values) > 4 and len(all_values[4]) > 1 else '',
                'check_out_time': all_values[8][1] if len(all_values) > 8 and len(all_values[8]) > 1 else '',
                'check_in_time': all_values[9][1] if len(all_values) > 9 and len(all_values[9]) > 1 else '',
                'amenities': all_values[10][1] if len(all_values) > 10 and len(all_values[10]) > 1 else '',
                'laundry_services': all_values[11][1] if len(all_values) > 11 and len(all_values[11]) > 1 else '',
                'keys': all_values[12][1] if len(all_values) > 12 and len(all_values[12]) > 1 else '',
                'codes': all_values[13][1] if len(all_values) > 13 and len(all_values[13]) > 1 else '',
                'properties': []
            }
            
            # Extract properties (starting from row 18, assuming a structured table)
            for i in range(17, len(all_values)):
                if len(all_values[i]) > 1 and all_values[i][0]: # Ensure there's data and a property name
                    profile['properties'].append({
                        'name': all_values[i][0],
                        'address': all_values[i][1] if len(all_values[i]) > 1 else '',
                        'hours': all_values[i][2] if len(all_values[i]) > 2 else '',
                        'so_hours': all_values[i][3] if len(all_values[i]) > 3 else ''
                    })
            
            add_log(f"Client: {profile['client_name']} with {len(profile['properties'])} properties.", "INFO")
            return profile
        except Exception as e:
            add_log(f"Error reading client profile: {e}", "ERROR")
            return {}

    def get_calendar_sheets(self, workbook) -> List[Dict]:
        """Retrieves all sheets in a workbook, excluding the first (profile) sheet."""
        try:
            add_log("Retrieving calendar sheets...", "INFO")
            worksheets = workbook.worksheets()
            calendars = []
            for i, sheet in enumerate(worksheets[1:], start=1): # Start from the second sheet
                calendars.append({
                    'index': i,
                    'name': sheet.title,
                    'sheet': sheet
                })
                add_log(f"  Calendar sheet found: {sheet.title}", "INFO")
            add_log(f"Total calendar sheets: {len(calendars)}", "SUCCESS")
            return calendars
        except Exception as e:
            add_log(f"Error listing calendar sheets: {e}", "ERROR")
            return []

    def read_calendar(self, sheet, start_row: int = 13) -> pd.DataFrame:
        """Reads booking calendar data from a sheet, starting from a specified row."""
        try:
            add_log(f"Reading calendar: {sheet.title} (starting from row {start_row})", "INFO")
            all_values = sheet.get_all_values()
            
            if len(all_values) < start_row:
                add_log(f"Sheet has insufficient rows (found {len(all_values)}, need {start_row})", "WARNING")
                return pd.DataFrame()
            
            headers = all_values[start_row - 2] if start_row > 1 else all_values[0]
            unique_headers = []
            seen = {}
            for header in headers:
                if header in seen:
                    seen[header] += 1
                    unique_headers.append(f"{header}_{seen[header]}")
                else:
                    seen[header] = 0
                    unique_headers.append(header)
            
            data = all_values[start_row - 1:]
            df = pd.DataFrame(data, columns=unique_headers)
            df = df[df.apply(lambda row: row.astype(str).str.strip().any(), axis=1)] # Clean empty rows
            
            add_log(f"Loaded {len(df)} booking rows from {sheet.title}", "SUCCESS")
            return df
        except Exception as e:
            add_log(f"Error reading calendar: {e}", "ERROR")
            return pd.DataFrame()

    def read_sheet_all_data(self, sheet) -> pd.DataFrame:
        """Reads all data from any given sheet."""
        try:
            add_log(f"Reading all data from sheet: {sheet.title}", "INFO")
            all_values = sheet.get_all_values()
            if not all_values:
                add_log("Sheet is empty", "WARNING")
                return pd.DataFrame()
            
            headers = all_values[0]
            unique_headers = []
            seen = {}
            for header in headers:
                if header in seen:
                    seen[header] += 1
                    unique_headers.append(f"{header}_{seen[header]}")
                else:
                    seen[header] = 0
                    unique_headers.append(header)
            
            df = pd.DataFrame(all_values[1:], columns=unique_headers)
            add_log(f"Loaded {len(df)} rows with {len(df.columns)} columns from {sheet.title}", "SUCCESS")
            return df
        except Exception as e:
            add_log(f"Error reading sheet: {e}", "ERROR")
            return pd.DataFrame()

    def update_cell_live(self, sheet, row: int, col: int, value: str) -> bool:
        """Updates a single cell in a sheet in real-time."""
        try:
            add_log(f"Updating cell [{row}, {col}] in {sheet.title} to: {value}", "INFO")
            sheet.update_cell(row, col, value)
            add_log("Cell updated successfully", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error updating cell: {e}", "ERROR")
            return False

    def append_row(self, sheet, data: List) -> bool:
        """Appends a new row of data to the sheet."""
        try:
            add_log(f"Appending row to {sheet.title}", "INFO")
            sheet.append_row(data)
            add_log(f"Row appended successfully with {len(data)} cells", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error appending row: {e}", "ERROR")
            return False

    def copy_row(self, sheet, row_index: int) -> List:
        """Copies a specific row from the sheet."""
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
            add_log(f"Error copying row: {e}", "ERROR")
            return []

    def batch_update_cells(self, sheet, cell_list: List[Dict]) -> bool:
        """Batch updates multiple cells in a sheet."""
        try:
            add_log(f"Batch updating {len(cell_list)} cells in {sheet.title}", "INFO")
            for cell_data in cell_list:
                sheet.update_cell(cell_data['row'], cell_data['col'], cell_data['value'])
            add_log(f"Batch update completed: {len(cell_list)} cells updated", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error during batch update: {e}", "ERROR")
            return False

    def create_new_spreadsheet(self, title: str, folder_id: Optional[str] = None) -> Optional[Dict]:
        """Creates a new Google Sheet spreadsheet."""
        if not self.gc:
            add_log("GSpread client not authenticated.", "ERROR")
            return None
        try:
            add_log(f"Creating new spreadsheet: {title}", "INFO")
            spreadsheet = self.gc.create(title)
            if folder_id and self.drive_service:
                # Move the created spreadsheet to the specified folder
                file_id = spreadsheet.id
                self.drive_service.files().update(
                    fileId=file_id,
                    addParents=folder_id,
                    removeParents='root' # Remove from root folder
                ).execute()
                add_log(f"Moved spreadsheet '{title}' to folder ID: {folder_id}", "SUCCESS")
            add_log(f"Spreadsheet '{title}' created successfully with ID: {spreadsheet.id}", "SUCCESS")
            return {'id': spreadsheet.id, 'name': title, 'url': spreadsheet.url, 'modified': datetime.now().isoformat()}
        except Exception as e:
            add_log(f"Error creating new spreadsheet: {e}", "ERROR")
            return None

    def upload_csv_to_sheet(self, file_content: str, sheet_name: str, workbook_id: str) -> bool:
        """Uploads CSV content to a new sheet within an existing workbook."""
        if not self.gc:
            add_log("GSpread client not authenticated.", "ERROR")
            return False
        try:
            workbook = self.gc.open_by_key(workbook_id)
            new_sheet = workbook.add_worksheet(title=sheet_name, rows="100", cols="20")
            csv_data = pd.read_csv(StringIO(file_content))
            new_sheet.update([csv_data.columns.values.tolist()] + csv_data.values.tolist())
            add_log(f"CSV data uploaded to new sheet '{sheet_name}' in workbook '{workbook.title}'", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error uploading CSV to sheet: {e}", "ERROR")
            return False

    def download_csv_from_url(self, csv_url: str) -> Optional[pd.DataFrame]:
        """Downloads CSV data from a given URL and returns a pandas DataFrame."""
        try:
            add_log(f"Attempting to download CSV from URL: {csv_url}", "INFO")
            df = pd.read_csv(csv_url)
            add_log(f"Successfully downloaded CSV with {len(df)} rows.", "SUCCESS")
            return df
        except Exception as e:
            add_log(f"Error downloading CSV from URL: {e}", "ERROR")
            st.error(f"Failed to download CSV from URL: {e}")
            return None

    def get_drive_file_details(self, file_id: str) -> Optional[Dict]:
        """Gets details of a file from Google Drive."""
        if not self.drive_service:
            add_log("Drive service not authenticated.", "ERROR")
            return None
        try:
            file = self.drive_service.files().get(fileId=file_id, fields="id, name, mimeType, webViewLink, parents, modifiedTime").execute()
            add_log(f"Retrieved details for file ID {file_id}: {file['name']}", "INFO")
            return file
        except Exception as e:
            add_log(f"Error getting Drive file details for {file_id}: {e}", "ERROR")
            return None

    def list_drive_files(self, folder_id: str) -> List[Dict]:
        """Lists all files and folders within a specified Google Drive folder."""
        if not self.drive_service:
            add_log("Drive service not authenticated.", "ERROR")
            return []
        try:
            add_log(f"Listing all files in Drive folder ID: {folder_id}", "INFO")
            query = f"'{folder_id}' in parents and trashed=false"
            results = self.drive_service.files().list(q=query, fields="files(id, name, mimeType, modifiedTime, webViewLink)").execute()
            items = results.get('files', [])
            add_log(f"Found {len(items)} files/folders in folder {folder_id}", "SUCCESS")
            return items
        except Exception as e:
            add_log(f"Error listing Drive files: {e}", "ERROR")
            return []

    def download_drive_file_as_csv(self, file_id: str) -> Optional[pd.DataFrame]:
        """Downloads a Google Sheet file as CSV from Drive and returns a pandas DataFrame."""
        if not self.drive_service:
            add_log("Drive service not authenticated.", "ERROR")
            return None
        try:
            add_log(f"Downloading file ID {file_id} as CSV from Drive.", "INFO")
            request = self.drive_service.files().export(fileId=file_id, mimeType='text/csv')
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                add_log(f"Download {int(status.progress() * 100)}%.", "INFO")
            fh.seek(0)
            df = pd.read_csv(fh)
            add_log(f"Successfully downloaded Drive file {file_id} as CSV with {len(df)} rows.", "SUCCESS")
            return df
        except Exception as e:
            add_log(f"Error downloading Drive file {file_id} as CSV: {e}", "ERROR")
            st.error(f"Failed to download Drive file as CSV: {e}")
            return None


# --- Streamlit UI Functions --- #

def authenticate():
    """Handles Google Sheets API authentication via Streamlit sidebar."""
    st.sidebar.header("🔑 Google API Authentication")
    credentials_json = st.sidebar.text_area(
        "Enter your Google Service Account Credentials JSON",
        help="Paste the content of your service account JSON key file here.",
        type="password"
    )
    
    if credentials_json:
        manager = GoogleSheetManager()
        if manager.authenticate_gspread(credentials_json):
            st.sidebar.success("✅ Authentication successful!")
            st.rerun()
        else:
            st.sidebar.error("❌ Authentication failed. Check your credentials.")
    else:
        st.sidebar.info("Please enter your service account credentials to authenticate.")

def render_sidebar(manager):
    """Renders the Streamlit sidebar with navigation and controls."""
    st.sidebar.title("🏠 Navigation")
    
    if st.session_state.authenticated:
        st.sidebar.markdown(f"""
        <div class="sidebar-card">
            <h3>Authenticated</h3>
            <p>Service Account: <span class="highlight">{st.session_state.service_account_email}</span></p>
        </div>
        """, unsafe_allow_html=True)
        
        # Main navigation
        view_mode = st.sidebar.radio(
            "Select View",
            ["All Clients", "Sheets Viewer", "Calendar View", "Booking Manager", "Drive Explorer", "Profile Page", "Booking Example", "System Logs"],
            key="main_view_mode"
        )
        
        st.sidebar.markdown("---")
        
        # Workbook selection
        st.sidebar.markdown("<div class='sidebar-card'><h3>Workbooks</h3></div>", unsafe_allow_html=True)
        
        if st.sidebar.button("🔄 Refresh Workbooks", use_container_width=True):
            with st.spinner("Refreshing workbooks from Drive..."):
                manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
            st.sidebar.success(f"Found {len(st.session_state.workbooks)} workbooks.")
            st.rerun()

        if st.session_state.workbooks:
            workbook_names = [wb['name'] for wb in st.session_state.workbooks]
            selected_workbook_name = st.sidebar.selectbox(
                "Select a Client Workbook",
                workbook_names,
                index=workbook_names.index(next((wb['name'] for wb in st.session_state.workbooks if wb['id'] == st.session_state.current_workbook_id), workbook_names[0])) if st.session_state.current_workbook_id and st.session_state.current_workbook_id in [wb['id'] for wb in st.session_state.workbooks] else 0,
                key="selected_workbook"
            )
            
            selected_workbook_info = next((wb for wb in st.session_state.workbooks if wb['name'] == selected_workbook_name), None)
            if selected_workbook_info and st.session_state.current_workbook_id != selected_workbook_info['id']:
                st.session_state.current_workbook_id = selected_workbook_info['id']
                st.session_state.all_sheets = [] # Clear sheets when workbook changes
                st.session_state.current_sheet = None
                st.rerun()
            
            if selected_workbook_info:
                st.sidebar.markdown(f"<p style='font-size:0.8em;'>ID: {selected_workbook_info['id']}</p>", unsafe_allow_html=True)
                st.sidebar.markdown(f"<p style='font-size:0.8em;'>Last Modified: {selected_workbook_info['modified'][:10]}</p>", unsafe_allow_html=True)
                st.sidebar.link_button("Open in Google Sheets", url=selected_workbook_info['url'], use_container_width=True)
        else:
            st.sidebar.info("No workbooks found. Ensure the Drive Folder ID is correct and service account has access.")
            st.session_state.current_workbook_id = None

        st.sidebar.markdown("---")
        st.sidebar.markdown("<div class='sidebar-card'><h3>New Spreadsheet</h3></div>", unsafe_allow_html=True)
        new_ss_title = st.sidebar.text_input("New Spreadsheet Title", key="new_ss_title")
        if st.sidebar.button("➕ Create New Spreadsheet", use_container_width=True):
            if new_ss_title:
                with st.spinner(f"Creating '{new_ss_title}'..."):
                    new_workbook_info = manager.create_new_spreadsheet(new_ss_title, DRIVE_FOLDER_ID)
                    if new_workbook_info:
                        st.session_state.workbooks.append(new_workbook_info)
                        st.session_state.current_workbook_id = new_workbook_info['id']
                        st.sidebar.success(f"✅ Spreadsheet '{new_ss_title}' created!")
                        st.rerun()
                    else:
                        st.sidebar.error("❌ Failed to create spreadsheet.")
            else:
                st.sidebar.warning("Please enter a title for the new spreadsheet.")

        st.sidebar.markdown("---")
        st.sidebar.markdown("<div class='sidebar-card'><h3>Upload CSV</h3></div>", unsafe_allow_html=True)
        uploaded_file = st.sidebar.file_uploader("Upload CSV to Current Workbook", type=["csv"])
        if uploaded_file and st.session_state.current_workbook_id:
            if st.sidebar.button("⬆️ Upload CSV", use_container_width=True):
                bytes_data = uploaded_file.getvalue()
                string_data = bytes_data.decode('utf-8')
                sheet_name = uploaded_file.name.replace('.csv', '')[:100] # Max sheet name length
                with st.spinner(f"Uploading '{uploaded_file.name}'..."):
                    if manager.upload_csv_to_sheet(string_data, sheet_name, st.session_state.current_workbook_id):
                        st.sidebar.success(f"✅ CSV '{uploaded_file.name}' uploaded as new sheet!")
                        st.session_state.all_sheets = [] # Force refresh sheets list
                        st.rerun()
                    else:
                        st.sidebar.error("❌ Failed to upload CSV.")
        elif uploaded_file and not st.session_state.current_workbook_id:
            st.sidebar.warning("Please select or create a workbook first to upload CSV.")

        st.sidebar.markdown("---")
        st.sidebar.markdown("<div class='sidebar-card'><h3>Live CSV from URL</h3></div>", unsafe_allow_html=True)
        csv_url_input = st.sidebar.text_input("Google Sheet CSV Export URL", key="csv_url_input", help="e.g., https://docs.google.com/spreadsheets/d/.../export?format=csv")
        if st.sidebar.button("🔗 Load Live CSV", use_container_width=True):
            if csv_url_input:
                with st.spinner("Loading live CSV..."):
                    df_live_csv = manager.download_csv_from_url(csv_url_input)
                    if df_live_csv is not None:
                        st.session_state.uploaded_csv_data = df_live_csv
                        st.sidebar.success("✅ Live CSV loaded successfully!")
                        st.rerun()
                    else:
                        st.session_state.uploaded_csv_data = None
                        st.sidebar.error("❌ Failed to load live CSV. Check URL and permissions.")
            else:
                st.sidebar.warning("Please enter a Google Sheet CSV export URL.")

        return view_mode
    else:
        st.sidebar.info("Please authenticate to access features.")
        return "Authentication"

def render_all_clients(manager):
    """Renders comprehensive view of all clients/workbooks."""
    st.markdown('<div class="section-header">👥 All Clients Overview</div>', unsafe_allow_html=True)
    
    if not st.session_state.workbooks:
        st.warning("⚠️ No workbooks found. Click 'Refresh Workbooks' in the sidebar to scan for spreadsheets.")
        st.markdown(f"""
        <div class="info-box">
            <h3>📁 Folder Configuration</h3>
            <p><strong>Drive Folder ID:</strong> {DRIVE_FOLDER_ID}</p>
            <p><strong>Service Account:</strong> {st.session_state.service_account_email}</p>
            <p><a href="https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}" target="_blank">🔗 Open Folder in Google Drive</a></p>
        </div>
        """, unsafe_allow_html=True)
        return
    
    st.markdown("### 🔍 Search & Filter")
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        search_term = st.text_input(
            "Search clients by name",
            value=st.session_state.search_term,
            placeholder="Type to search...",
            key="client_search"
        )
        st.session_state.search_term = search_term
    
    with col2:
        sort_by = st.selectbox(
            "Sort by",
            ["Name (A-Z)", "Name (Z-A)", "Properties (Most)", "Properties (Least)", "Last Modified (Newest)", "Last Modified (Oldest)"]
        )
    
    with col3:
        view_style = st.selectbox(
            "View Style",
            ["Detailed", "Compact"]
        )
    
    st.markdown("---")
    
    # Summary metrics at the top
    st.markdown("### 📊 Portfolio Summary")
    
    total_properties = 0
    total_calendars = 0
    total_bookings = 0
    
    # Pre-fetch profiles and calendar counts for sorting and metrics to avoid repeated API calls
    workbook_details = []
    for workbook_info in st.session_state.workbooks:
        try:
            wb = manager.open_workbook_by_id(workbook_info['id'])
            if wb:
                profile = manager.get_client_profile(wb)
                calendars = manager.get_calendar_sheets(wb)
                
                current_bookings = 0
                for cal in calendars:
                    df = manager.read_calendar(cal['sheet'])
                    if not df.empty:
                        current_bookings += len(df)
                
                workbook_details.append({
                    'info': workbook_info,
                    'profile': profile,
                    'num_properties': len(profile.get('properties', [])),
                    'num_calendars': len(calendars),
                    'num_bookings': current_bookings
                })
                total_properties += len(profile.get('properties', []))
                total_calendars += len(calendars)
                total_bookings += current_bookings
        except Exception as e:
            add_log(f"Error processing workbook {workbook_info['name']}: {e}", "WARNING")
            workbook_details.append({
                'info': workbook_info,
                'profile': {'client_name': workbook_info['name'], 'properties': []},
                'num_properties': 0,
                'num_calendars': 0,
                'num_bookings': 0
            })

    col1, col2, col3, col4, col5 = st.columns(5)
    
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
            <h3 style="color: #11998e; margin: 0;">🏘️ Properties</h3>
            <h1 style="margin: 0.5rem 0;">{total_properties}</h1>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #f093fb; margin: 0;">📅 Calendars</h3>
            <h1 style="margin: 0.5rem 0;">{total_calendars}</h1>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <h3 style="color: #ff9800; margin: 0;">📋 Bookings</h3>
            <h1 style="margin: 0.5rem 0;">{total_bookings}</h1>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        if st.button("🔄 Refresh All", use_container_width=True, type="primary"):
            with st.spinner("Refreshing all clients..."):
                st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
            st.success(f"✅ Refreshed! Found {len(st.session_state.workbooks)} clients")
            st.rerun()
    
    st.markdown("---")
    
    st.markdown("### 📥 Bulk Actions")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📊 Export All Client Data", use_container_width=True):
            with st.spinner("Generating export..."):
                all_export_data = []
                for detail in workbook_details:
                    profile = detail['profile']
                    all_export_data.append({
                        'Client': profile.get('client_name', 'Unknown'),
                        'Properties': detail['num_properties'],
                        'Calendars': detail['num_calendars'],
                        'Total Bookings': detail['num_bookings'],
                        'Check-out Time': profile.get('check_out_time', ''),
                        'Check-in Time': profile.get('check_in_time', ''),
                        'Spreadsheet ID': detail['info']['id'],
                        'Spreadsheet URL': detail['info']['url']
                    })
                
                if all_export_data:
                    df_export = pd.DataFrame(all_export_data)
                    csv = df_export.to_csv(index=False)
                    st.download_button(
                        "💾 Download CSV",
                        data=csv,
                        file_name=f"all_clients_summary_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                else:
                    st.warning("No data to export.")
    
    with col2:
        if st.button("📋 Generate Summary Report", use_container_width=True):
            st.info("Summary report feature - coming soon!")
    
    with col3:
        if st.button("📧 Email Report", use_container_width=True):
            st.info("Email report feature - coming soon!")
    
    st.markdown("---")
    
    filtered_workbook_details = workbook_details
    
    if search_term:
        filtered_workbook_details = [
            detail for detail in filtered_workbook_details 
            if search_term.lower() in detail['info']['name'].lower()
        ]
    
    # Sort workbooks
    if sort_by == "Name (A-Z)":
        filtered_workbook_details = sorted(filtered_workbook_details, key=lambda x: x['info']['name'])
    elif sort_by == "Name (Z-A)":
        filtered_workbook_details = sorted(filtered_workbook_details, key=lambda x: x['info']['name'], reverse=True)
    elif sort_by == "Properties (Most)":
        filtered_workbook_details = sorted(filtered_workbook_details, key=lambda x: x['num_properties'], reverse=True)
    elif sort_by == "Properties (Least)":
        filtered_workbook_details = sorted(filtered_workbook_details, key=lambda x: x['num_properties'])
    elif sort_by == "Last Modified (Newest)":
        filtered_workbook_details = sorted(filtered_workbook_details, key=lambda x: x['info'].get('modified', ''), reverse=True)
    elif sort_by == "Last Modified (Oldest)":
        filtered_workbook_details = sorted(filtered_workbook_details, key=lambda x: x['info'].get('modified', ''))

    st.markdown(f"### 🏢 Client Details ({len(filtered_workbook_details)} shown)")
    
    if not filtered_workbook_details:
        st.warning("No clients match your search criteria")
        return
    
    for idx, detail in enumerate(filtered_workbook_details):
        workbook_info = detail['info']
        profile = detail['profile']
        total_properties = detail['num_properties']
        total_bookings = detail['num_bookings']

        # Recalculate status counts for display, as it's not cached in detail
        status_counts = {code: 0 for code in STATUS_CODES.keys()}
        try:
            wb_obj = manager.open_workbook_by_id(workbook_info['id'])
            if wb_obj:
                calendars_for_status = manager.get_calendar_sheets(wb_obj)
                for cal in calendars_for_status:
                    df_cal = manager.read_calendar(cal['sheet'])
                    if not df_cal.empty:
                        for col in df_cal.columns:
                            if 'status' in col.lower() or 'code' in col.lower():
                                for status in STATUS_CODES.keys():
                                    status_counts[status] += df_cal[col].astype(str).str.contains(status, case=False, na=False).sum()
        except Exception as e:
            add_log(f"Error getting status counts for {workbook_info['name']}: {e}", "WARNING")

        if view_style == "Compact":
            col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
            with col1:
                st.markdown(f"**🏢 {profile.get('client_name', workbook_info['name'])}**")
            with col2:
                st.metric("Properties", total_properties)
            with col3:
                st.metric("Calendars", detail['num_calendars'])
            with col4:
                st.metric("Bookings", total_bookings)
            with col5:
                if st.button("View", key=f"view_client_{idx}", use_container_width=True):
                    st.session_state.current_workbook_id = workbook_info['id']
                    st.session_state.all_sheets = []
                    st.session_state.current_sheet = None
                    st.session_state.main_view_mode = "Sheets Viewer" # Navigate to Sheets Viewer
                    st.rerun()
        else: # Detailed view
            st.markdown(f"""
            <div class="property-card">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <h3 style="margin-top: 0; color: #667eea;">🏢 {profile.get('client_name', workbook_info['name'])}</h3>
                    <a href="{workbook_info['url']}" target="_blank" style="text-decoration: none;">
                        <button style="background-color: #667eea; color: white; border: none; padding: 8px 15px; border-radius: 5px; cursor: pointer;">Open Sheet</button>
                    </a>
                </div>
                <p><span class="detail-label">Spreadsheet ID:</span> {workbook_info['id']}</p>
                <p><span class="detail-label">Last Modified:</span> {workbook_info['modified'][:10]}</p>
                <p><span class="detail-label">Contact:</span> {profile.get('contact_person', 'N/A')} ({profile.get('contact_email', 'N/A')})</p>
                <p><span class="detail-label">Address:</span> {profile.get('address', 'N/A')}</p>
                
                <h4>Properties ({total_properties})</h4>
                <ul>
            """, unsafe_allow_html=True)
            for prop in profile.get('properties', [])[:3]: # Show up to 3 properties in detailed view
                st.markdown(f"<li><strong>{prop['name']}</strong>: {prop['address']} (Hours: {prop['hours']})</li>", unsafe_allow_html=True)
            if total_properties > 3:
                st.markdown(f"<li>... and {total_properties - 3} more.</li>", unsafe_allow_html=True)
            st.markdown("</ul>", unsafe_allow_html=True)

            st.markdown(f"<h4>Booking Summary</h4>", unsafe_allow_html=True)
            cols_status = st.columns(len(STATUS_CODES))
            for s_idx, (code, info) in enumerate(STATUS_CODES.items()):
                with cols_status[s_idx]:
                    count = status_counts.get(code, 0)
                    st.markdown(f"""
                    <div class="status-badge" style="background-color: {info['color']};">
                        {code}: {count}
                    </div>
                    """, unsafe_allow_html=True)
            
            if st.button("Manage Bookings", key=f"manage_bookings_{idx}", use_container_width=True):
                st.session_state.current_workbook_id = workbook_info['id']
                st.session_state.all_sheets = []
                st.session_state.current_sheet = None
                st.session_state.main_view_mode = "Booking Manager" # Navigate to Booking Manager
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("--- ")

def render_all_sheets_viewer(manager, workbook):
    """Renders a viewer for all sheets within a selected workbook."""
    st.markdown('<div class="section-header">📄 Sheets Viewer</div>', unsafe_allow_html=True)
    
    if not st.session_state.all_sheets:
        workbook_obj = manager.open_workbook_by_id(st.session_state.current_workbook_id)
        if not workbook_obj:
            st.error("❌ Could not load workbook sheets.")
            return

    sheet_names = [s['name'] for s in st.session_state.all_sheets]
    if not sheet_names:
        st.warning("No sheets found in this workbook.")
        return

    selected_sheet_name = st.selectbox(
        "Select a Sheet to View",
        sheet_names,
        index=st.session_state.selected_sheet_index if st.session_state.selected_sheet_index < len(sheet_names) else 0,
        key="sheet_viewer_select"
    )
    st.session_state.selected_sheet_index = sheet_names.index(selected_sheet_name)
    
    selected_sheet_info = next((s for s in st.session_state.all_sheets if s['name'] == selected_sheet_name), None)
    if not selected_sheet_info:
        st.error("Selected sheet not found.")
        return
    
    sheet_obj = selected_sheet_info['sheet_obj']
    st.session_state.current_sheet = sheet_obj

    st.markdown("--- ")
    st.markdown(f"### Viewing: {selected_sheet_name}")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        view_type = st.radio("Data View Type", ["Calendar Data (from row 13)", "Full Data"], key="sheet_data_view_type")
    with col2:
        show_info = st.checkbox("Show Sheet Statistics", value=True)
    
    with st.spinner(f"Loading {selected_sheet_info['name']}..."):
        if view_type == "Full Data":
            df = manager.read_sheet_all_data(sheet_obj)
        else:
            df = manager.read_calendar(sheet_obj, start_row=13)
        
        if not df.empty:
            st.success(f"✅ Loaded {len(df)} rows with {len(df.columns)} columns")
            st.dataframe(df, use_container_width=True, height=600)
            
            csv = df.to_csv(index=False)
            st.download_button(
                label="📥 Download as CSV",
                data=csv,
                file_name=f"{selected_sheet_info['name']}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
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
    """Renders calendar view for booking data."""
    st.markdown('<div class="section-header">📅 Calendar View</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    if not calendars:
        st.warning("No calendar sheets found in this workbook.")
        return
    
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar_name = st.selectbox("Select Calendar Month", calendar_names, key="calendar_view_select")
    
    selected_cal_info = next((cal for cal in calendars if cal['name'] == selected_calendar_name), None)
    
    if selected_cal_info:
        df = manager.read_calendar(selected_cal_info['sheet'])
        
        if not df.empty:
            st.dataframe(df, use_container_width=True, height=600)
            csv = df.to_csv(index=False)
            st.download_button(
                "📥 Download Calendar",
                data=csv,
                file_name=f"{selected_calendar_name}.csv",
                mime="text/csv"
            )
        else:
            st.info("No bookings in this calendar sheet.")

def render_booking_manager(manager, workbook):
    """Renders booking manager with live editing capabilities."""
    st.markdown('<div class="section-header">✏️ Booking Manager - Live Edit</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    if not calendars:
        st.warning("No calendar sheets found in this workbook.")
        return
    
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar_name = st.selectbox("Select Calendar to Edit", calendar_names, key="booking_manager_select")
    
    selected_cal_info = next((cal for cal in calendars if cal['name'] == selected_calendar_name), None)
    
    if selected_cal_info:
        sheet = selected_cal_info['sheet']
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Edit Data", "➕ Add Booking", "📋 Copy/Append", "🔍 Search & Filter", "📊 Analytics"])
        
        with tab1:
            st.markdown("### Edit Existing Data")
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                st.dataframe(df, use_container_width=True)
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    # Adjust row number for gspread (1-indexed, and data starts from start_row)
                    edit_row_display = st.number_input("Row Number (from calendar data)", min_value=1, value=1, key="edit_row_display")
                    actual_row_in_sheet = edit_row_display + 12 # Assuming calendar data starts at row 13
                with col2:
                    edit_col = st.number_input("Column Number", min_value=1, value=1, key="edit_col")
                with col3:
                    new_value = st.text_input("New Value", key="new_value_input")
                
                if st.button("💾 Update Cell", type="primary", key="update_cell_button"):
                    if manager.update_cell_live(sheet, actual_row_in_sheet, edit_col, new_value):
                        st.success("✅ Cell updated successfully!")
                        st.rerun()
                    else:
                        st.error("❌ Failed to update cell")
            else:
                st.info("No data to edit in this calendar sheet.")
        
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
                
                if st.button("➕ Add Booking Row", type="primary", key="add_booking_row_button"):
                    if manager.append_row(sheet, new_row_data):
                        st.success("✅ Booking added successfully!")
                        st.rerun()
                    else:
                        st.error("❌ Failed to add booking")
            else:
                st.warning("Cannot add booking - sheet structure unknown. Please ensure the sheet has headers.")
        
        with tab3:
            st.markdown("### Copy and Append Rows")
            df = manager.read_calendar(sheet)
            if not df.empty:
                st.dataframe(df, use_container_width=True)
                
                row_to_copy_display = st.number_input(
                    "Row Number to Copy (from calendar data)", 
                    min_value=1, 
                    value=1,
                    help="Enter the row number from the displayed calendar data you want to copy",
                    key="row_to_copy_display"
                )
                actual_row_to_copy_in_sheet = row_to_copy_display + 12 # Adjust for header rows

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("📋 Copy Row", type="secondary", key="copy_row_button"):
                        copied_data = manager.copy_row(sheet, actual_row_to_copy_in_sheet -1) # gspread is 0-indexed for get_all_values
                        if copied_data:
                            st.session_state['copied_row'] = copied_data
                            st.success(f"✅ Row {row_to_copy_display} copied! ({len(copied_data)} cells)")
                        else:
                            st.error("❌ Failed to copy row")
                
                with col2:
                    if st.button("➕ Append Copied Row", type="primary", key="append_copied_row_button"):
                        if 'copied_row' in st.session_state and st.session_state['copied_row']:
                            if manager.append_row(sheet, st.session_state['copied_row']):
                                st.success("✅ Row appended successfully!")
                                st.rerun()
                            else:
                                st.error("❌ Failed to append row")
                        else:
                            st.warning("⚠️ No row copied yet. Copy a row first!")
                
                if 'copied_row' in st.session_state and st.session_state['copied_row']:
                    st.markdown("### 📋 Copied Row Preview")
                    st.json(st.session_state['copied_row'])
            else:
                st.info("No data available to copy in this calendar sheet.")
        
        with tab4:
            st.markdown("### 🔍 Search & Filter Bookings")
            df = manager.read_calendar(sheet)
            if not df.empty:
                col1, col2 = st.columns(2)
                with col1:
                    search_booking = st.text_input("Search in bookings", placeholder="Type to search...", key="search_booking_input")
                with col2:
                    filter_column = st.selectbox("Filter by column", ["All"] + list(df.columns), key="filter_column_select")
                
                filtered_df = df.copy()
                if search_booking:
                    mask = filtered_df.astype(str).apply(lambda row: row.str.contains(search_booking, case=False, na=False).any(), axis=1)
                    filtered_df = filtered_df[mask]
                
                if filter_column != "All":
                    unique_values = filtered_df[filter_column].unique()
                    selected_value = st.selectbox(f"Select {filter_column}", ["All"] + list(unique_values), key="filter_value_select")
                    if selected_value != "All":
                        filtered_df = filtered_df[filtered_df[filter_column] == selected_value]
                
                st.markdown(f"**Showing {len(filtered_df)} of {len(df)} bookings**")
                st.dataframe(filtered_df, use_container_width=True)
                
                if len(filtered_df) > 0:
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(
                        "📥 Download Filtered Data",
                        data=csv,
                        file_name=f"filtered_bookings_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="download_filtered_csv"
                    )
            else:
                st.info("No data to search in this calendar sheet.")
        
        with tab5:
            st.markdown("### 📊 Booking Analytics")
            df = manager.read_calendar(sheet)
            if not df.empty:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f"""
                    <div class="analytics-card">
                        <h4 style="margin: 0;">📋 Total Bookings</h4>
                        <h2 style="margin: 0.5rem 0;">{len(df)}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    non_empty_rows = df.dropna(how='all')
                    st.markdown(f"""
                    <div class="analytics-card">
                        <h4 style="margin: 0;">✅ Active Bookings</h4>
                        <h2 style="margin: 0.5rem 0;">{len(non_empty_rows)}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="analytics-card">
                        <h4 style="margin: 0;">📊 Columns</h4>
                        <h2 style="margin: 0.5rem 0;">{len(df.columns)}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                st.markdown("#### 📈 Status Code Distribution")
                status_data = {code: 0 for code in STATUS_CODES.keys()}
                for col in df.columns:
                    if 'status' in col.lower() or 'code' in col.lower():
                        for status in STATUS_CODES.keys():
                            status_data[status] += df[col].astype(str).str.contains(status, case=False, na=False).sum()
                
                status_cols = st.columns(len(STATUS_CODES))
                for idx, (code, info) in enumerate(STATUS_CODES.items()):
                    with status_cols[idx]:
                        count = status_data.get(code, 0)
                        st.markdown(f"""
                        <div class="metric-card">
                            <h5 style="margin: 0; color: {info['color']};">{code}</h5>
                            <h3 style="margin: 0.3rem 0;">{count}</h3>
                            <p style="margin: 0; font-size: 0.75rem;">{info['name']}</p>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                st.markdown("#### 📊 Column Fill Rates")
                fill_rates = []
                for col in df.columns:
                    non_empty = df[col].notna().sum()
                    fill_rate = (non_empty / len(df)) * 100 if len(df) > 0 else 0
                    fill_rates.append({
                        'Column': col,
                        'Non-Empty': non_empty,
                        'Fill Rate': f"{fill_rate:.1f}%"
                    })
                df_fill = pd.DataFrame(fill_rates)
                st.dataframe(df_fill, use_container_width=True)
            else:
                st.info("No data available for analytics in this calendar sheet.")

def render_drive_explorer(manager):
    """Renders a Google Drive file explorer for the configured folder."""
    st.markdown('<div class="section-header">📁 Google Drive Explorer</div>', unsafe_allow_html=True)

    st.markdown(f"""
    <div class="info-box">
        <h3>Current Drive Folder</h3>
        <p><strong>Folder ID:</strong> {DRIVE_FOLDER_ID}</p>
        <p><a href="https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}" target="_blank">🔗 Open Folder in Google Drive</a></p>
    </div>
    """, unsafe_allow_html=True)

    if st.button("🔄 Refresh Drive Files", use_container_width=True, key="refresh_drive_files_button"):
        st.session_state.drive_files = manager.list_drive_files(DRIVE_FOLDER_ID)
        st.rerun()

    if 'drive_files' not in st.session_state or not st.session_state.drive_files:
        st.session_state.drive_files = manager.list_drive_files(DRIVE_FOLDER_ID)

    if st.session_state.drive_files:
        st.markdown("### Files and Folders")
        file_data = []
        for item in st.session_state.drive_files:
            file_data.append({
                'Name': item['name'],
                'Type': item['mimeType'].split('.')[-1].replace('spreadsheet', 'Google Sheet').replace('folder', 'Folder'),
                'Last Modified': item['modifiedTime'][:10],
                'Link': item.get('webViewLink', '#'),
                'ID': item['id']
            })
        df_files = pd.DataFrame(file_data)
        st.dataframe(df_files, use_container_width=True, hide_index=True)

        st.markdown("--- ")
        st.markdown("### Download Google Sheet as CSV")
        sheet_files = [f for f in st.session_state.drive_files if 'spreadsheet' in f['mimeType']]
        if sheet_files:
            sheet_names = [f['name'] for f in sheet_files]
            selected_sheet_to_download = st.selectbox("Select a Google Sheet to download as CSV", sheet_names, key="download_sheet_select")
            selected_sheet_file = next((f for f in sheet_files if f['name'] == selected_sheet_to_download), None)

            if selected_sheet_file:
                if st.button(f"⬇️ Download '{selected_sheet_file['name']}' as CSV", use_container_width=True, type="primary"):
                    with st.spinner(f"Downloading '{selected_sheet_file['name']}'..."):
                        df_downloaded = manager.download_drive_file_as_csv(selected_sheet_file['id'])
                        if df_downloaded is not None:
                            csv_output = df_downloaded.to_csv(index=False)
                            st.download_button(
                                label=f"💾 Download {selected_sheet_file['name']}.csv",
                                data=csv_output,
                                file_name=f"{selected_sheet_file['name']}.csv",
                                mime="text/csv",
                                use_container_width=True
                            )
                            st.success(f"✅ '{selected_sheet_file['name']}' downloaded as CSV.")
                        else:
                            st.error(f"❌ Failed to download '{selected_sheet_file['name']}' as CSV.")
        else:
            st.info("No Google Sheet files found in this Drive folder to download.")

    else:
        st.info("No files or folders found in the specified Drive folder. Ensure the ID is correct and permissions are set.")

def render_profile_page(manager):
    """Renders a dedicated profile page for the currently selected client/workbook."""
    st.markdown('<div class="section-header">👤 Client Profile Page</div>', unsafe_allow_html=True)

    if not st.session_state.current_workbook_id:
        st.info("👈 Please select a client workbook from the sidebar to view their profile.")
        return

    workbook = manager.open_workbook_by_id(st.session_state.current_workbook_id)
    if not workbook:
        st.error("❌ Failed to load client workbook for profile.")
        return

    profile = manager.get_client_profile(workbook)

    if profile:
        st.markdown(f"""
        <div class="profile-card">
            <h2>{profile.get('client_name', 'Unknown Client')}</h2>
            <p><span class="detail-label">Contact Person:</span> {profile.get('contact_person', 'N/A')}</p>
            <p><span class="detail-label">Email:</span> {profile.get('contact_email', 'N/A')}</p>
            <p><span class="detail-label">Phone:</span> {profile.get('phone_number', 'N/A')}</p>
            <p><span class="detail-label">Address:</span> {profile.get('address', 'N/A')}</p>
            <p><span class="detail-label">Check-in Time:</span> {profile.get('check_in_time', 'N/A')}</p>
            <p><span class="detail-label">Check-out Time:</span> {profile.get('check_out_time', 'N/A')}</p>
            <p><span class="detail-label">Amenities:</span> {profile.get('amenities', 'N/A')}</p>
            <p><span class="detail-label">Laundry Services:</span> {profile.get('laundry_services', 'N/A')}</p>
            <p><span class="detail-label">Keys Info:</span> {profile.get('keys', 'N/A')}</p>
            <p><span class="detail-label">Codes Info:</span> {profile.get('codes', 'N/A')}</p>

            <h3>Associated Properties ({len(profile.get('properties', []))})</h3>
            <ul>
        """, unsafe_allow_html=True)
        if profile.get('properties'):
            for prop in profile['properties']:
                st.markdown(f"<li><strong>{prop['name']}</strong>: {prop['address']} (Hours: {prop['hours']}, SO Hours: {prop['so_hours']})</li>", unsafe_allow_html=True)
        else:
            st.markdown("<li>No properties listed.</li>", unsafe_allow_html=True)
        st.markdown("</ul></div>", unsafe_allow_html=True)

        st.markdown("--- ")
        st.markdown("### Quick Actions")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("View Bookings", use_container_width=True, key="profile_view_bookings"):
                st.session_state.main_view_mode = "Booking Manager"
                st.rerun()
        with col2:
            if st.button("Edit Profile Data (Coming Soon)", use_container_width=True, key="profile_edit_data", disabled=True):
                pass # Placeholder for future edit functionality
    else:
        st.warning("Could not retrieve profile information for the selected client.")

def render_booking_example(manager):
    """Renders a simplified booking example interface."""
    st.markdown('<div class="section-header">🗓️ Booking Example</div>', unsafe_allow_html=True)

    st.info("This is an example interface to demonstrate how a booking could be made or managed. It interacts with the currently selected workbook.")

    if not st.session_state.current_workbook_id:
        st.info("👈 Please select a client workbook from the sidebar to use the booking example.")
        return

    workbook = manager.open_workbook_by_id(st.session_state.current_workbook_id)
    if not workbook:
        st.error("❌ Failed to load client workbook for booking example.")
        return

    calendars = manager.get_calendar_sheets(workbook)
    if not calendars:
        st.warning("No calendar sheets found in this workbook to make a booking.")
        return
    
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar_name = st.selectbox("Select Calendar for Booking", calendar_names, key="booking_example_calendar_select")
    selected_cal_info = next((cal for cal in calendars if cal['name'] == selected_calendar_name), None)

    if selected_cal_info:
        sheet = selected_cal_info['sheet']
        st.markdown(f"### Make a New Booking in '{selected_calendar_name}'")

        with st.form(key='new_booking_form', clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                guest_name = st.text_input("Guest Name", max_chars=100)
                check_in_date = st.date_input("Check-in Date", value=datetime.now())
                num_guests = st.number_input("Number of Guests", min_value=1, value=1)
            with col2:
                property_name = st.text_input("Property Name/Unit", max_chars=100)
                check_out_date = st.date_input("Check-out Date", value=datetime.now() + timedelta(days=1))
                booking_status = st.selectbox("Booking Status", list(STATUS_CODES.keys()), format_func=lambda x: STATUS_CODES[x]['name'])
            
            notes = st.text_area("Additional Notes", height=100)

            submit_button = st.form_submit_button(label='✨ Submit New Booking', type="primary")

            if submit_button:
                if guest_name and property_name and check_in_date and check_out_date:
                    # Attempt to get existing headers to match new booking data
                    df_current = manager.read_calendar(sheet)
                    if not df_current.empty:
                        headers = df_current.columns.tolist()
                    else:
                        # Fallback headers if sheet is empty or structure unknown
                        headers = ["Guest Name", "Property", "Check-in Date", "Check-out Date", "Guests", "Status", "Notes"]
                    
                    # Map form data to headers, filling empty for unmatched columns
                    new_booking_data_dict = {
                        "Guest Name": guest_name,
                        

                        "Property": property_name,
                        "Check-in Date": check_in_date.strftime("%Y-%m-%d"),
                        "Check-out Date": check_out_date.strftime("%Y-%m-%d"),
                        "Guests": str(num_guests),
                        "Status": booking_status,
                        "Notes": notes
                    }

                    # Create a list of values in the correct order of headers
                    new_row_values = [new_booking_data_dict.get(header, '') for header in headers]

                    if manager.append_row(sheet, new_row_values):
                        st.success("✅ New booking added successfully!")
                        st.rerun()
                    else:
                        st.error("❌ Failed to add new booking.")
                else:
                    st.warning("Please fill in all required booking details (Guest Name, Property, Check-in, Check-out).")

        st.markdown("--- ")
        st.markdown("### Existing Bookings Overview")
        df_bookings = manager.read_calendar(sheet)
        if not df_bookings.empty:
            st.dataframe(df_bookings, use_container_width=True)
        else:
            st.info("No bookings found in this calendar.")

def render_system_logs():
    """Render system logs"""
    st.markdown("<div class=\"section-header\">📋 System Logs</div>", unsafe_allow_html=True)
    
    if st.button("🗑️ Clear Logs"):
        st.session_state.logs = []
        st.rerun()
    
    if st.session_state.logs:
        for log in reversed(st.session_state.logs):
            level_class = f"log-{log["level"].lower()}"
            st.markdown(f"""
            <div class="log-entry {level_class}">
                <strong>[{log["timestamp"]}]</strong> [{log["level"]}] {log["message"]}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No logs yet")

# Main execution
if __name__ == "__main__":
    manager = GoogleSheetManager()
    
    view_mode = render_sidebar(manager)

    st.markdown("<h1 class=\"main-header\">Professional Booking Management System</h1>", unsafe_allow_html=True)

    if st.session_state.authenticated:
        if view_mode == "All Clients":
            render_all_clients(manager)
        elif view_mode == "Sheets Viewer":
            if st.session_state.current_workbook_id:
                workbook = manager.open_workbook_by_id(st.session_state.current_workbook_id)
                if workbook:
                    render_all_sheets_viewer(manager, workbook)
                else:
                    st.error("❌ Failed to load workbook for Sheets Viewer.")
            else:
                st.info("👈 Please select a workbook from the sidebar.")
        elif view_mode == "Calendar View":
            if st.session_state.current_workbook_id:
                workbook = manager.open_workbook_by_id(st.session_state.current_workbook_id)
                if workbook:
                    render_calendar_view(manager, workbook)
                else:
                    st.error("❌ Failed to load workbook for Calendar View.")
            else:
                st.info("👈 Please select a workbook from the sidebar.")
        elif view_mode == "Booking Manager":
            if st.session_state.current_workbook_id:
                workbook = manager.open_workbook_by_id(st.session_state.current_workbook_id)
                if workbook:
                    render_booking_manager(manager, workbook)
                else:
                    st.error("❌ Failed to load workbook for Booking Manager.")
            else:
                st.info("👈 Please select a workbook from the sidebar.")
        elif view_mode == "Drive Explorer":
            render_drive_explorer(manager)
        elif view_mode == "Profile Page":
            render_profile_page(manager)
        elif view_mode == "Booking Example":
            render_booking_example(manager)
        elif view_mode == "System Logs":
            render_system_logs()

        # Display live CSV data if available
        if st.session_state.uploaded_csv_data is not None:
            st.markdown("<div class=\"section-header\">📊 Live CSV Data from URL</div>", unsafe_allow_html=True)
            st.dataframe(st.session_state.uploaded_csv_data, use_container_width=True)
            csv_output = st.session_state.uploaded_csv_data.to_csv(index=False)
            st.download_button(
                label="📥 Download Live CSV Data",
                data=csv_output,
                file_name=f"live_csv_data_{datetime.now().strftime("%Y%m%d")}.csv",
                mime="text/csv",
                use_container_width=True
            )

    else:
        authenticate()


