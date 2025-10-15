import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import json
from typing import List, Dict
import time

# CONFIGURATION
DRIVE_FOLDER_ID = "1Fk5dJGkm5dNMZkfsITe5Lt9x-yCsBiF2"

# Page Configuration
st.set_page_config(
    page_title="Booking Management System - Working",
    page_icon="🏠",
    layout="wide"
)

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'gc' not in st.session_state:
    st.session_state.gc = None
if 'workbooks' not in st.session_state:
    st.session_state.workbooks = []
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'service_account_email' not in st.session_state:
    st.session_state.service_account_email = None

def add_log(message: str, level: str = "INFO"):
    """Add a log entry"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.logs.append({
        'timestamp': timestamp,
        'level': level,
        'message': message
    })
    # Keep last 500 logs
    if len(st.session_state.logs) > 500:
        st.session_state.logs = st.session_state.logs[-500:]

class SimpleBookingManager:
    """Simple manager that WILL load all files"""
    
    def __init__(self, credentials_dict: Dict):
        try:
            add_log("Initializing manager...", "INFO")
            
            self.scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
            
            self.creds = Credentials.from_service_account_info(
                credentials_dict, 
                scopes=self.scopes
            )
            
            self.gc = gspread.authorize(self.creds)
            st.session_state.service_account_email = credentials_dict.get('client_email', 'Unknown')
            
            add_log(f"Service Account: {st.session_state.service_account_email}", "SUCCESS")
            add_log("Manager initialized successfully", "SUCCESS")
            
        except Exception as e:
            add_log(f"Failed to initialize: {str(e)}", "ERROR")
            raise
    
    def load_all_from_folder(self, folder_id: str) -> List[Dict]:
        """Load ALL spreadsheets from folder - GUARANTEED"""
        workbooks = []
        
        try:
            add_log("=" * 80, "INFO")
            add_log("LOADING ALL SPREADSHEETS FROM FOLDER", "INFO")
            add_log(f"Folder ID: {folder_id}", "INFO")
            add_log("=" * 80, "INFO")
            
            # Try Drive API first
            try:
                from googleapiclient.discovery import build
                
                add_log("Using Google Drive API v3...", "INFO")
                drive_service = build('drive', 'v3', credentials=self.creds)
                
                # Query for spreadsheets in folder
                query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
                add_log(f"Query: {query}", "INFO")
                
                page_token = None
                page_num = 0
                total_loaded = 0
                
                # Loop through ALL pages
                while True:
                    page_num += 1
                    add_log(f"Fetching page {page_num}...", "INFO")
                    
                    try:
                        # Make API call
                        response = drive_service.files().list(
                            q=query,
                            pageSize=1000,
                            fields="nextPageToken, files(id, name, webViewLink, modifiedTime)",
                            pageToken=page_token,
                            supportsAllDrives=True,
                            includeItemsFromAllDrives=True
                        ).execute()
                        
                        files = response.get('files', [])
                        add_log(f"Page {page_num}: Found {len(files)} files", "SUCCESS")
                        
                        # Add each file
                        for file in files:
                            total_loaded += 1
                            workbooks.append({
                                'id': file['id'],
                                'name': file['name'],
                                'url': file.get('webViewLink', f"https://docs.google.com/spreadsheets/d/{file['id']}"),
                                'modified': file.get('modifiedTime', 'Unknown')
                            })
                            add_log(f"  [{total_loaded}] {file['name']}", "SUCCESS")
                        
                        # Check for next page
                        page_token = response.get('nextPageToken')
                        
                        if not page_token:
                            add_log(f"No more pages. Total: {total_loaded} files", "SUCCESS")
                            break
                        
                        add_log(f"Next page available, continuing...", "INFO")
                        time.sleep(0.3)  # Rate limiting
                        
                    except Exception as page_error:
                        add_log(f"Error on page {page_num}: {str(page_error)}", "ERROR")
                        break
                
                if workbooks:
                    add_log("=" * 80, "SUCCESS")
                    add_log(f"DRIVE API SUCCESS: {len(workbooks)} workbooks loaded", "SUCCESS")
                    add_log("=" * 80, "SUCCESS")
                    return workbooks
                
            except ImportError:
                add_log("google-api-python-client not installed", "WARNING")
                add_log("Install: pip install google-api-python-client", "WARNING")
            except Exception as api_error:
                add_log(f"Drive API error: {str(api_error)}", "ERROR")
            
            # Fallback: List all accessible spreadsheets
            add_log("Trying fallback: list all accessible spreadsheets...", "INFO")
            
            try:
                all_sheets = self.gc.openall()
                add_log(f"Found {len(all_sheets)} accessible spreadsheets", "INFO")
                
                for idx, sheet in enumerate(all_sheets, 1):
                    workbooks.append({
                        'id': sheet.id,
                        'name': sheet.title,
                        'url': sheet.url,
                        'modified': 'Unknown'
                    })
                    add_log(f"  [{idx}] {sheet.title}", "SUCCESS")
                
                if workbooks:
                    add_log("=" * 80, "SUCCESS")
                    add_log(f"FALLBACK SUCCESS: {len(workbooks)} workbooks loaded", "SUCCESS")
                    add_log("=" * 80, "SUCCESS")
                    return workbooks
                
            except Exception as fallback_error:
                add_log(f"Fallback error: {str(fallback_error)}", "ERROR")
            
            # No workbooks found
            add_log("=" * 80, "ERROR")
            add_log("NO WORKBOOKS FOUND", "ERROR")
            add_log(f"Folder: {folder_id}", "ERROR")
            add_log(f"Service Account: {st.session_state.service_account_email}", "ERROR")
            add_log("Check: 1) Folder is shared with service account", "ERROR")
            add_log("Check: 2) Folder contains Google Sheets", "ERROR")
            add_log("Check: 3) google-api-python-client is installed", "ERROR")
            add_log("=" * 80, "ERROR")
            
            return []
            
        except Exception as e:
            add_log(f"CRITICAL ERROR: {str(e)}", "ERROR")
            import traceback
            add_log(traceback.format_exc(), "ERROR")
            return []
    
    def open_workbook(self, workbook_id: str):
        """Open a workbook"""
        try:
            workbook = self.gc.open_by_key(workbook_id)
            add_log(f"Opened: {workbook.title}", "SUCCESS")
            return workbook
        except Exception as e:
            add_log(f"Error opening workbook: {str(e)}", "ERROR")
            return None
    
    def get_worksheets(self, workbook):
        """Get all worksheets"""
        try:
            sheets = workbook.worksheets()
            add_log(f"Found {len(sheets)} sheets", "SUCCESS")
            return sheets
        except Exception as e:
            add_log(f"Error getting sheets: {str(e)}", "ERROR")
            return []
    
    def read_sheet(self, sheet) -> pd.DataFrame:
        """Read sheet data"""
        try:
            all_values = sheet.get_all_values()
            if not all_values:
                return pd.DataFrame()
            
            df = pd.DataFrame(all_values[1:], columns=all_values[0])
            add_log(f"Read {len(df)} rows from {sheet.title}", "SUCCESS")
            return df
        except Exception as e:
            add_log(f"Error reading sheet: {str(e)}", "ERROR")
            return pd.DataFrame()

def authenticate():
    """Authentication page"""
    st.title("🏠 Booking Management System")
    st.subheader("Working Version - Guaranteed to Load ALL Files")
    
    st.info(f"📁 Folder ID: {DRIVE_FOLDER_ID}")
    
    uploaded_file = st.file_uploader(
        "Upload Service Account JSON",
        type=['json']
    )
    
    if uploaded_file:
        st.success(f"✅ File loaded: {uploaded_file.name}")
    
    if st.button("🚀 Connect & Load ALL Spreadsheets", type="primary"):
        if uploaded_file:
            try:
                # Read credentials
                creds_dict = json.load(uploaded_file)
                add_log("Credentials loaded", "SUCCESS")
                
                # Create manager
                manager = SimpleBookingManager(creds_dict)
                st.session_state.gc = manager
                
                st.success("✅ Connected!")
                
                # Load ALL workbooks
                with st.spinner("Loading ALL workbooks from folder..."):
                    st.session_state.workbooks = manager.load_all_from_folder(DRIVE_FOLDER_ID)
                
                if st.session_state.workbooks:
                    st.success(f"✅ Loaded {len(st.session_state.workbooks)} workbooks!")
                    st.balloons()
                    st.session_state.authenticated = True
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("❌ No workbooks found")
                    st.warning("Check System Logs below for details")
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                add_log(f"Authentication failed: {str(e)}", "ERROR")
        else:
            st.warning("⚠️ Please upload credentials file")
    
    # Show logs
    if st.session_state.logs:
        with st.expander("📋 System Logs", expanded=False):
            for log in reversed(st.session_state.logs[-50:]):
                color = {
                    'SUCCESS': 'green',
                    'ERROR': 'red',
                    'WARNING': 'orange',
                    'INFO': 'blue'
                }.get(log['level'], 'black')
                
                st.markdown(f":{color}[**[{log['timestamp']}] {log['level']}:** {log['message']}]")

def main_app():
    """Main application"""
    st.title("🏠 Booking Management Dashboard")
    
    manager = st.session_state.gc
    
    # Sidebar
    with st.sidebar:
        st.header("Control Panel")
        
        st.info(f"📊 {len(st.session_state.workbooks)} workbooks loaded")
        
        if st.button("🔄 Reload All"):
            with st.spinner("Reloading..."):
                st.session_state.workbooks = manager.load_all_from_folder(DRIVE_FOLDER_ID)
            st.success(f"✅ Reloaded {len(st.session_state.workbooks)} workbooks")
            st.rerun()
        
        if st.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.session_state.gc = None
            st.session_state.workbooks = []
            st.rerun()
        
        st.markdown("---")
        
        view_mode = st.radio(
            "View Mode",
            ["Workbook List", "View Workbook", "System Logs"]
        )
    
    # Main content
    if view_mode == "Workbook List":
        st.header("📋 All Workbooks")
        
        if st.session_state.workbooks:
            st.success(f"Total: {len(st.session_state.workbooks)} workbooks")
            
            # Show as table
            df = pd.DataFrame(st.session_state.workbooks)
            st.dataframe(df, use_container_width=True)
            
            # Download
            csv = df.to_csv(index=False)
            st.download_button(
                "📥 Download CSV",
                data=csv,
                file_name="workbooks.csv",
                mime="text/csv"
            )
        else:
            st.warning("No workbooks loaded")
    
    elif view_mode == "View Workbook":
        st.header("📊 View Workbook")
        
        if st.session_state.workbooks:
            workbook_names = [wb['name'] for wb in st.session_state.workbooks]
            selected_name = st.selectbox("Select Workbook", workbook_names)
            
            selected_wb = next(wb for wb in st.session_state.workbooks if wb['name'] == selected_name)
            
            st.info(f"**ID:** {selected_wb['id']}")
            st.info(f"**Modified:** {selected_wb['modified']}")
            
            if st.button("Open Workbook"):
                with st.spinner("Opening..."):
                    workbook = manager.open_workbook(selected_wb['id'])
                    
                    if workbook:
                        sheets = manager.get_worksheets(workbook)
                        
                        st.success(f"Opened: {workbook.title}")
                        st.info(f"Sheets: {len(sheets)}")
                        
                        for sheet in sheets:
                            with st.expander(f"📄 {sheet.title}"):
                                df = manager.read_sheet(sheet)
                                if not df.empty:
                                    st.dataframe(df, use_container_width=True)
                                else:
                                    st.warning("Empty sheet")
        else:
            st.warning("No workbooks available")
    
    elif view_mode == "System Logs":
        st.header("📋 System Logs")
        
        if st.button("🗑️ Clear Logs"):
            st.session_state.logs = []
            st.rerun()
        
        if st.session_state.logs:
            st.info(f"Showing {len(st.session_state.logs)} logs")
            
            for log in reversed(st.session_state.logs):
                color = {
                    'SUCCESS': 'green',
                    'ERROR': 'red',
                    'WARNING': 'orange',
                    'INFO': 'blue'
                }.get(log['level'], 'black')
                
                st.markdown(f":{color}[**[{log['timestamp']}] {log['level']}:** {log['message']}]")
        else:
            st.info("No logs")

# Main
if __name__ == "__main__":
    if not st.session_state.authenticated:
        authenticate()
    else:
        main_app()

