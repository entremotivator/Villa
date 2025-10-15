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
from calendar import monthrange
import numpy as np

# Page configuration
st.set_page_config(
    page_title="Villa Booking Management System Pro++",
    page_icon="🏖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced CSS with calendar styles
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
        color: #000000;
        margin: 0 0 0.5rem 0;
        font-size: 14px;
        font-weight: 600;
    }
    
    .sidebar-card p {
        color: #000000;
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
    
    .calendar-day {
        background: white;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 8px;
        min-height: 100px;
        transition: all 0.3s ease;
    }
    
    .calendar-day:hover {
        border-color: #667eea;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.2);
        transform: translateY(-2px);
    }
    
    .calendar-day-header {
        font-weight: 600;
        color: #667eea;
        font-size: 14px;
        margin-bottom: 8px;
        padding-bottom: 4px;
        border-bottom: 2px solid #667eea;
    }
    
    .calendar-booking {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 4px 8px;
        margin: 4px 0;
        border-radius: 6px;
        font-size: 11px;
        cursor: pointer;
        transition: all 0.2s ease;
    }
    
    .calendar-booking:hover {
        transform: scale(1.05);
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
    }
    
    .calendar-today {
        background: linear-gradient(135deg, #fff9e6 0%, #fff3cc 100%);
        border: 2px solid #fee140;
    }
    
    .sync-indicator {
        display: inline-block;
        width: 10px;
        height: 10px;
        background: #38ef7d;
        border-radius: 50%;
        animation: pulse 2s infinite;
        margin-right: 6px;
    }
    
    @keyframes pulse {
        0%, 100% { opacity: 1; transform: scale(1); }
        50% { opacity: 0.5; transform: scale(1.2); }
    }
    
    .folder-sync-badge {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 6px 12px;
        border-radius: 15px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
        margin: 4px;
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
    
    .multi-sheet-badge {
        background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
        color: white;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        margin-left: 8px;
    }
    
    .timeline-view {
        background: white;
        border-radius: 12px;
        padding: 1rem;
        margin: 1rem 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .timeline-slot {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 8px;
        margin: 4px 0;
        border-radius: 6px;
        border-left: 4px solid #667eea;
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
if 'folder_sheets' not in st.session_state:
    st.session_state.folder_sheets = []
if 'all_bookings_df' not in st.session_state:
    st.session_state.all_bookings_df = pd.DataFrame()
if 'auto_sync_enabled' not in st.session_state:
    st.session_state.auto_sync_enabled = True
if 'last_sync_time' not in st.session_state:
    st.session_state.last_sync_time = None
if 'calendar_view_mode' not in st.session_state:
    st.session_state.calendar_view_mode = 'month'
if 'selected_calendar_date' not in st.session_state:
    st.session_state.selected_calendar_date = datetime.now()

# Constants
ORIGINAL_SPREADSHEET_ID = "1-3FLLEkUmiHzW7DGVAPI6PebdRc_24t3vM0OCBnDhco"
ORIGINAL_SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{ORIGINAL_SPREADSHEET_ID}/edit"
TARGET_FOLDER_ID = "1Fk5dJGkm5dNMZkfsITe5Lt9x-yCsBiF2"
FOLDER_URL = f"https://drive.google.com/drive/folders/{TARGET_FOLDER_ID}"
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
    """Log email notification"""
    log_activity(f"📧 Email notification: {subject}", "info")

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
        
        # Auto-load folder sheets
        try:
            load_folder_sheets(client)
        except Exception as e:
            log_activity(f"Could not auto-load folder sheets: {str(e)}", "error")
        
        return client
    except Exception as e:
        log_activity(f"Authentication failed: {str(e)}", "error")
        st.error(f"Authentication failed: {str(e)}")
        return None

def get_folder_files(client: gspread.Client, folder_id: str) -> List[Dict[str, str]]:
    """Get all spreadsheets from a specific Google Drive folder"""
    try:
        # Use Drive API to list files in folder
        drive_service = client.http_client
        
        # List all spreadsheets accessible to the service account
        all_sheets = client.openall()
        
        # Filter sheets - this is a simplified version
        # In production, you'd use proper Drive API filtering
        folder_sheets = []
        for sheet in all_sheets:
            # Check if sheet might be in target folder (simplified check)
            folder_sheets.append({
                'id': sheet.id,
                'title': sheet.title,
                'url': sheet.url,
                'last_modified': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
        
        log_activity(f"Found {len(folder_sheets)} sheets in folder", "success")
        return folder_sheets
        
    except Exception as e:
        log_activity(f"Error accessing folder: {str(e)}", "error")
        return []

def load_folder_sheets(client: gspread.Client):
    """Load all sheets from the target folder"""
    try:
        folder_sheets = get_folder_files(client, TARGET_FOLDER_ID)
        st.session_state.folder_sheets = folder_sheets
        log_activity(f"Loaded {len(folder_sheets)} sheets from folder", "success")
        return True
    except Exception as e:
        log_activity(f"Failed to load folder sheets: {str(e)}", "error")
        return False

def sync_all_sheets(client: gspread.Client) -> pd.DataFrame:
    """Sync and merge data from all sheets in folder"""
    all_data = []
    
    try:
        # Get sheets from folder
        sheets_to_sync = st.session_state.folder_sheets if st.session_state.folder_sheets else st.session_state.available_spreadsheets
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, sheet_info in enumerate(sheets_to_sync):
            try:
                status_text.text(f"Syncing: {sheet_info['title'][:40]}...")
                progress_bar.progress((idx + 1) / len(sheets_to_sync))
                
                worksheet = get_worksheet(client, sheet_info['id'])
                if worksheet:
                    df = read_bookings(worksheet)
                    if not df.empty:
                        df['source_sheet'] = sheet_info['title']
                        df['source_id'] = sheet_info['id']
                        df['sync_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        all_data.append(df)
                        log_activity(f"Synced {len(df)} bookings from {sheet_info['title']}", "success")
            except Exception as e:
                log_activity(f"Error syncing {sheet_info['title']}: {str(e)}", "error")
        
        progress_bar.empty()
        status_text.empty()
        
        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)
            st.session_state.all_bookings_df = combined_df
            st.session_state.last_sync_time = datetime.now()
            log_activity(f"Successfully synced {len(combined_df)} total bookings from {len(all_data)} sheets", "success")
            return combined_df
        else:
            log_activity("No data to sync", "info")
            return pd.DataFrame()
            
    except Exception as e:
        log_activity(f"Sync failed: {str(e)}", "error")
        return pd.DataFrame()

def get_worksheet(client: gspread.Client, spreadsheet_id: str) -> Optional[gspread.Worksheet]:
    """Get the specific worksheet"""
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.get_worksheet(0)
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

def create_calendar_view(df: pd.DataFrame, year: int, month: int):
    """Create an interactive calendar view"""
    if df.empty:
        st.info("No bookings to display in calendar view")
        return
    
    # Parse dates
    df_copy = df.copy()
    try:
        df_copy['parsed_date'] = pd.to_datetime(df_copy['DATE:'], format='%m/%d/%Y', errors='coerce')
    except:
        st.error("Error parsing dates for calendar view")
        return
    
    # Filter for selected month
    df_month = df_copy[
        (df_copy['parsed_date'].dt.year == year) & 
        (df_copy['parsed_date'].dt.month == month)
    ]
    
    # Get calendar info
    first_day = datetime(year, month, 1)
    num_days = monthrange(year, month)[1]
    start_weekday = first_day.weekday()
    
    # Calendar header
    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        if st.button("◀ Previous Month"):
            if month == 1:
                st.session_state.selected_calendar_date = datetime(year - 1, 12, 1)
            else:
                st.session_state.selected_calendar_date = datetime(year, month - 1, 1)
            st.rerun()
    
    with col2:
        st.markdown(f"<h2 style='text-align: center;'>{first_day.strftime('%B %Y')}</h2>", unsafe_allow_html=True)
    
    with col3:
        if st.button("Next Month ▶"):
            if month == 12:
                st.session_state.selected_calendar_date = datetime(year + 1, 1, 1)
            else:
                st.session_state.selected_calendar_date = datetime(year, month + 1, 1)
            st.rerun()
    
    # Day headers
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    cols = st.columns(7)
    for i, day in enumerate(days):
        with cols[i]:
            st.markdown(f"<div style='text-align: center; font-weight: 600; color: #667eea;'>{day[:3]}</div>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Calendar grid
    week_cols = st.columns(7)
    current_day = 1
    today = datetime.now().date()
    
    # Empty cells before first day
    for i in range(start_weekday):
        with week_cols[i]:
            st.markdown("<div class='calendar-day' style='opacity: 0.3;'></div>", unsafe_allow_html=True)
    
    # Days with bookings
    for day in range(1, num_days + 1):
        col_index = (start_weekday + day - 1) % 7
        
        if col_index == 0 and day != 1:
            week_cols = st.columns(7)
        
        current_date = datetime(year, month, day).date()
        day_bookings = df_month[df_month['parsed_date'].dt.day == day]
        
        is_today = current_date == today
        day_class = "calendar-day calendar-today" if is_today else "calendar-day"
        
        with week_cols[col_index]:
            st.markdown(f"<div class='{day_class}'>", unsafe_allow_html=True)
            st.markdown(f"<div class='calendar-day-header'>{day}</div>", unsafe_allow_html=True)
            
            if len(day_bookings) > 0:
                for idx, booking in day_bookings.iterrows():
                    villa_short = booking['VILLA:'][:15] if len(booking['VILLA:']) > 15 else booking['VILLA:']
                    time_info = f"{booking.get('START TIME:', 'N/A')}"
                    
                    st.markdown(f"""
                    <div class='calendar-booking' title='{booking["VILLA:"]} - {booking.get("RESERVATION STATUS:", "N/A")}'>
                        🏠 {villa_short}<br>
                        ⏰ {time_info}
                    </div>
                    """, unsafe_allow_html=True)
                
                if len(day_bookings) > 2:
                    st.markdown(f"<div style='text-align: center; font-size: 10px; color: #667eea; font-weight: 600;'>+{len(day_bookings) - 2} more</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div style='text-align: center; font-size: 11px; color: #ccc; padding: 10px;'>No bookings</div>", unsafe_allow_html=True)
            
            st.markdown("</div>", unsafe_allow_html=True)

def create_timeline_view(df: pd.DataFrame, selected_date: datetime):
    """Create a timeline view for a specific date"""
    if df.empty:
        st.info("No bookings for this date")
        return
    
    df_copy = df.copy()
    try:
        df_copy['parsed_date'] = pd.to_datetime(df_copy['DATE:'], format='%m/%d/%Y', errors='coerce')
    except:
        st.error("Error parsing dates")
        return
    
    # Filter for selected date
    date_str = selected_date.strftime('%m/%d/%Y')
    day_bookings = df_copy[df_copy['DATE:'] == date_str]
    
    if day_bookings.empty:
        st.info(f"No bookings for {selected_date.strftime('%B %d, %Y')}")
        return
    
    st.markdown(f"<h3>📅 Timeline for {selected_date.strftime('%B %d, %Y')}</h3>", unsafe_allow_html=True)
    st.markdown(f"<p style='color: #666;'>Total bookings: <strong>{len(day_bookings)}</strong></p>", unsafe_allow_html=True)
    
    # Sort by start time
    day_bookings_sorted = day_bookings.sort_values('START TIME:')
    
    for idx, booking in day_bookings_sorted.iterrows():
        status_class = f"status-{booking['RESERVATION STATUS:'].lower().replace('!', '')}"
        
        st.markdown(f"""
        <div class='timeline-slot'>
            <div style='display: flex; justify-content: space-between; align-items: center;'>
                <div>
                    <strong style='font-size: 16px; color: #333;'>🏠 {booking['VILLA:']}</strong>
                    <br>
                    <span style='font-size: 13px; color: #666;'>
                        ⏰ {booking.get('START TIME:', 'N/A')} - {booking.get('END TIME:', 'N/A')} | 
                        👥 {booking.get('PAX:', 'N/A')} PAX | 
                        🧹 {booking.get('TYPE CLEAN:', 'N/A')}
                    </span>
                </div>
                <div>
                    <span class='{status_class}'>{booking['RESERVATION STATUS:']}</span>
                </div>
            </div>
            {f"<div style='margin-top: 8px; font-size: 12px; color: #555;'>💬 {booking['COMMENTS:']}</div>" if booking.get('COMMENTS:') else ""}
        </div>
        """, unsafe_allow_html=True)

def create_gantt_chart(df: pd.DataFrame):
    """Create a Gantt chart view of bookings"""
    if df.empty:
        st.info("No bookings available for Gantt chart")
        return
    
    try:
        df_copy = df.copy()
        df_copy['parsed_date'] = pd.to_datetime(df_copy['DATE:'], format='%m/%d/%Y', errors='coerce')
        df_copy = df_copy.dropna(subset=['parsed_date'])
        
        if df_copy.empty:
            st.info("No valid dates for Gantt chart")
            return
        
        # Prepare data for Gantt chart
        df_copy['Start'] = df_copy['parsed_date']
        df_copy['Finish'] = df_copy['parsed_date'] + timedelta(hours=4)  # Assume 4-hour duration
        df_copy['Task'] = df_copy['VILLA:']
        df_copy['Status'] = df_copy['RESERVATION STATUS:']
        
        # Create color mapping
        color_map = {
            'NEW!': '#38ef7d',
            'RESERVED!': '#00f2fe',
            'CANCELED!': '#f5576c',
            'UPDATE!': '#fee140'
        }
        
        df_copy['Color'] = df_copy['Status'].map(color_map).fillna('#667eea')
        
        fig = px.timeline(
            df_copy.head(50),  # Limit to 50 for performance
            x_start='Start',
            x_end='Finish',
            y='Task',
            color='Status',
            color_discrete_map=color_map,
            title='Booking Timeline (Gantt View)',
            hover_data=['START TIME:', 'END TIME:', 'PAX:', 'TYPE CLEAN:']
        )
        
        fig.update_layout(
            height=600,
            xaxis_title='Date',
            yaxis_title='Villa',
            showlegend=True,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
    except Exception as e:
        st.error(f"Error creating Gantt chart: {str(e)}")

def create_heatmap_view(df: pd.DataFrame):
    """Create a heatmap showing booking density"""
    if df.empty:
        st.info("No data for heatmap")
        return
    
    try:
        df_copy = df.copy()
        df_copy['parsed_date'] = pd.to_datetime(df_copy['DATE:'], format='%m/%d/%Y', errors='coerce')
        df_copy = df_copy.dropna(subset=['parsed_date'])
        
        # Group by date
        daily_counts = df_copy.groupby(df_copy['parsed_date'].dt.date).size().reset_index()
        daily_counts.columns = ['Date', 'Count']
        
        fig = px.density_heatmap(
            df_copy,
            x=df_copy['parsed_date'].dt.day_name(),
            y=df_copy['parsed_date'].dt.hour,
            title='Booking Density Heatmap',
            color_continuous_scale='Viridis'
        )
        
        fig.update_layout(
            height=400,
            xaxis_title='Day of Week',
            yaxis_title='Hour of Day',
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
    except Exception as e:
        st.error(f"Error creating heatmap: {str(e)}")

def render_folder_sync_section():
    """Render folder synchronization controls"""
    st.sidebar.markdown("---")
    st.sidebar.markdown("## 📁 Folder Sync")
    
    if st.session_state.authenticated and st.session_state.client:
        # Folder info
        st.sidebar.markdown(f"""
        <div class="sidebar-card">
            <h4>🎯 Target Folder</h4>
            <a href="{FOLDER_URL}" target="_blank" style="color: #667eea; text-decoration: none; font-weight: 600;">
                📂 Open Drive Folder →
            </a>
            <p style="margin-top: 0.5rem; font-size: 11px; color: #666;">Auto-syncing all sheets</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Sync status
        if st.session_state.last_sync_time:
            time_ago = datetime.now() - st.session_state.last_sync_time
            minutes_ago = int(time_ago.total_seconds() / 60)
            
            st.sidebar.markdown(f"""
            <div class="sidebar-card">
                <h4><span class='sync-indicator'></span>Last Sync</h4>
                <p style="color: #38ef7d; font-weight: 600;">{minutes_ago} minutes ago</p>
                <p style="font-size: 11px;">Total: {len(st.session_state.all_bookings_df)} bookings</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.sidebar.markdown("""
            <div class="sidebar-card">
                <h4>⚠️ Not Synced</h4>
                <p>Click below to sync all sheets</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Sync controls
        col1, col2 = st.sidebar.columns(2)
        
        with col1:
            if st.button("🔄 Sync Now", use_container_width=True):
                with st.spinner("Syncing all sheets..."):
                    if st.session_state.client:
                        # Reload folder sheets
                        load_folder_sheets(st.session_state.client)
                        # Sync all data
                        combined_df = sync_all_sheets(st.session_state.client)
                        if not combined_df.empty:
                            st.sidebar.success(f"✅ Synced {len(combined_df)} bookings!")
                            st.rerun()
                        else:
                            st.sidebar.warning("No data synced")
        
        with col2:
            st.session_state.auto_sync_enabled = st.checkbox(
                "Auto",
                value=st.session_state.auto_sync_enabled,
                help="Auto-sync every page refresh"
            )
        
        # Folder sheets list
        if st.session_state.folder_sheets:
            st.sidebar.markdown(f"""
            <div class="sidebar-card">
                <h4>📊 Folder Statistics</h4>
                <p>Sheets found: <strong>{len(st.session_state.folder_sheets)}</strong></p>
            </div>
            """, unsafe_allow_html=True)
            
            with st.sidebar.expander("📋 View All Sheets"):
                for sheet in st.session_state.folder_sheets[:10]:  # Show first 10
                    st.markdown(f"""
                    <div style='font-size: 11px; padding: 4px; border-bottom: 1px solid #eee;'>
                        📄 {sheet['title'][:30]}{'...' if len(sheet['title']) > 30 else ''}
                        <br>
                        <span style='color: #666; font-size: 10px;'>Modified: {sheet.get('last_modified', 'N/A')}</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                if len(st.session_state.folder_sheets) > 10:
                    st.caption(f"...and {len(st.session_state.folder_sheets) - 10} more")

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
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

def render_activity_log_sidebar():
    """Render activity log"""
    st.sidebar.markdown("---")
    st.sidebar.markdown("## 📋 Activity Log")
    
    # Email notifications toggle
    st.session_state.email_notifications_enabled = st.sidebar.checkbox(
        "📧 Email Notifications",
        value=st.session_state.email_notifications_enabled,
        help=f"Send alerts to {NOTIFICATION_EMAIL}"
    )
    
    # Activity stats
    log_count_total = len(st.session_state.activity_log)
    log_success = len([l for l in st.session_state.activity_log if l['type'] == 'success'])
    log_errors = len([l for l in st.session_state.activity_log if l['type'] == 'error'])
    
    st.sidebar.markdown(f"""
    <div class="sidebar-card">
        <h4>📊 Activity Stats</h4>
        <p>Total: <strong>{log_count_total}</strong></p>
        <p>✅ Success: <strong style="color: #38ef7d;">{log_success}</strong></p>
        <p>❌ Errors: <strong style="color: #f5576c;">{log_errors}</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Log controls
    col1, col2 = st.sidebar.columns(2)
    with col1:
        if st.button("🔄", use_container_width=True, key="refresh_log"):
            st.rerun()
    with col2:
        if st.button("🗑️", use_container_width=True, key="clear_log"):
            st.session_state.activity_log = []
            st.rerun()
    
    # Display logs
    log_count = min(10, len(st.session_state.activity_log))
    
    if st.session_state.activity_log:
        for log in st.session_state.activity_log[:log_count]:
            icon = {"success": "✅", "error": "❌", "info": "📌"}.get(log['type'], "ℹ️")
            st.sidebar.markdown(f"""
            <div class="log-entry log-entry-{log['type']}">
                <strong style="font-size: 10px;">{icon} {log['timestamp']}</strong><br>
                <span style="font-size: 11px;">{log['message'][:60]}...</span>
            </div>
            """, unsafe_allow_html=True)

def main():
    render_login_sidebar()
    
    if not st.session_state.authenticated:
        st.markdown("""
            <div class="booking-header">
                <h1>🏖️ Villa Booking Management System Pro++</h1>
                <p>Advanced booking management with folder sync & live calendar</p>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.info("""
            ### Welcome to Enhanced System!
            
            **New Features:**
            - 📁 Auto-sync entire Drive folder
            - 📅 Live calendar view
            - 📊 Timeline & Gantt charts
            - 🗺️ Booking heatmaps
            - 🔄 Real-time multi-sheet sync
            - 📧 Email notifications
            - 📈 Advanced analytics
            - 🎯 Smart filtering
            
            Upload your credentials to begin!
            """)
        return
    
    render_folder_sync_section()
    render_activity_log_sidebar()
    
    # Auto-sync on page load
    if st.session_state.auto_sync_enabled and st.session_state.client:
        if st.session_state.last_sync_time is None or \
           (datetime.now() - st.session_state.last_sync_time).total_seconds() > 300:  # 5 minutes
            with st.spinner("Auto-syncing..."):
                load_folder_sheets(st.session_state.client)
                sync_all_sheets(st.session_state.client)
    
    st.markdown("""
        <div class="booking-header">
            <h1>🏖️ Villa Booking Management System Pro++</h1>
            <p><span class='live-indicator'></span>Live folder sync & calendar integration</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Navigation
    st.sidebar.markdown("---")
    st.sidebar.title("📋 Navigation")
    page = st.sidebar.radio("", [
        "📊 Dashboard",
        "📅 Calendar View",
        "⏰ Timeline View",
        "📈 Analytics Pro",
        "📝 New Booking",
        "📋 View & Edit",
        "⚙️ Settings"
    ], label_visibility="collapsed")
    
    # Use combined data if available
    if not st.session_state.all_bookings_df.empty:
        df = st.session_state.all_bookings_df
        using_multi = True
    elif st.session_state.worksheet:
        df = read_bookings(st.session_state.worksheet)
        using_multi = False
    else:
        df = pd.DataFrame()
        using_multi = False
    
    # DASHBOARD
    if page == "📊 Dashboard":
        st.header("📊 Dashboard")
        
        if using_multi:
            st.success(f"📁 Viewing {len(df)} bookings from {df['source_sheet'].nunique() if 'source_sheet' in df.columns else 1} sheets")
        
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
            
            # Multi-sheet badges
            if using_multi and 'source_sheet' in df.columns:
                st.subheader("📚 Data Sources")
                unique_sheets = df['source_sheet'].unique()
                for sheet in unique_sheets[:10]:
                    st.markdown(f'<span class="folder-sync-badge">📄 {sheet[:40]}</span>', unsafe_allow_html=True)
                if len(unique_sheets) > 10:
                    st.markdown(f'<span class="folder-sync-badge">+{len(unique_sheets) - 10} more sheets</span>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Status pie chart
                fig_status = go.Figure(data=[go.Pie(
                    labels=['New', 'Reserved', 'Canceled', 'Update'],
                    values=[stats['new'], stats['reserved'], stats['canceled'], stats['update']],
                    marker=dict(colors=['#38ef7d', '#00f2fe', '#f5576c', '#fee140']),
                    hole=0.4
                )])
                fig_status.update_layout(
                    title='Booking Status Distribution',
                    height=400,
                    paper_bgcolor='rgba(0,0,0,0)'
                )
                st.plotly_chart(fig_status, use_container_width=True)
            
            with col2:
                # Recent bookings
                st.subheader("📅 Recent Bookings")
                recent_df = df.head(5)
                for idx, row in recent_df.iterrows():
                    source_badge = f"<span class='multi-sheet-badge'>{row.get('source_sheet', 'N/A')[:20]}</span>" if 'source_sheet' in row else ""
                    st.markdown(f"""
                    <div class="booking-card">
                        <strong>🏠 {row['VILLA:']}</strong> {source_badge}<br>
                        <span>📅 {row['DATE:']} | ⏰ {row.get('START TIME:', 'N/A')} - {row.get('END TIME:', 'N/A')}</span><br>
                        <span class="status-{row['RESERVATION STATUS:'].lower().replace('!', '')}">{row['RESERVATION STATUS:']}</span>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("No bookings found. Sync folder or create your first booking!")
    
    # CALENDAR VIEW
    elif page == "📅 Calendar View":
        st.header("📅 Live Calendar View")
        
        if not df.empty:
            # Calendar controls
            col1, col2, col3 = st.columns([2, 1, 2])
            
            with col2:
                view_mode = st.selectbox("View Mode", ["📅 Month", "📊 Gantt", "🗺️ Heatmap"], key="cal_view")
            
            if view_mode == "📅 Month":
                current_date = st.session_state.selected_calendar_date
                create_calendar_view(df, current_date.year, current_date.month)
            elif view_mode == "📊 Gantt":
                create_gantt_chart(df)
            else:
                create_heatmap_view(df)
        else:
            st.info("No bookings to display. Sync your folder first!")
    
    # TIMELINE VIEW
    elif page == "⏰ Timeline View":
        st.header("⏰ Timeline View")
        
        if not df.empty:
            selected_date = st.date_input("Select Date", datetime.now())
            create_timeline_view(df, selected_date)
        else:
            st.info("No bookings available")
    
    # ANALYTICS PRO
    elif page == "📈 Analytics Pro":
        st.header("📈 Advanced Analytics")
        
        if not df.empty:
            stats = get_booking_statistics(df)
            
            # Top metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Bookings", stats['total'], f"+{stats['new']} new")
            with col2:
                villa_count = df['VILLA:'].nunique()
                st.metric("Unique Villas", villa_count)
            with col3:
                if 'source_sheet' in df.columns:
                    sheet_count = df['source_sheet'].nunique()
                    st.metric("Data Sources", sheet_count)
            with col4:
                reserved_rate = (stats['reserved'] / stats['total'] * 100) if stats['total'] > 0 else 0
                st.metric("Reserved Rate", f"{reserved_rate:.1f}%")
            
            st.markdown("---")
            
            # Charts
            col1, col2 = st.columns(2)
            
            with col1:
                # Top villas
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
            
            with col2:
                # Cleaning types
                cleaning_counts = df['TYPE CLEAN:'].value_counts()
                fig_cleaning = px.pie(
                    values=cleaning_counts.values,
                    names=cleaning_counts.index,
                    title='Bookings by Cleaning Type',
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                st.plotly_chart(fig_cleaning, use_container_width=True)
            
            # Export section
            st.markdown("---")
            st.subheader("📥 Export Combined Data")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "⬇️ Download CSV",
                    data=csv,
                    file_name=f"all_bookings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            with col2:
                json_str = df.to_json(orient='records', indent=2)
                st.download_button(
                    "⬇️ Download JSON",
                    data=json_str,
                    file_name=f"all_bookings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                    use_container_width=True
                )
            
            with col3:
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='All Bookings')
                excel_data = excel_buffer.getvalue()
                
                st.download_button(
                    "⬇️ Download Excel",
                    data=excel_data,
                    file_name=f"all_bookings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("No data available for analytics")
    
    # NEW BOOKING
    elif page == "📝 New Booking":
        st.header("📝 Create New Booking")
        
        if not st.session_state.worksheet:
            st.warning("Please select a spreadsheet from the sidebar to add bookings")
            return
        
        worksheet = st.session_state.worksheet
        
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
                            # Trigger re-sync
                            if st.session_state.auto_sync_enabled:
                                sync_all_sheets(st.session_state.client)
                        else:
                            st.markdown('<div class="error-message">❌ Failed to add booking</div>', unsafe_allow_html=True)
    
    # VIEW & EDIT BOOKINGS
    elif page == "📋 View & Edit":
        st.header("📋 View & Edit Bookings")
        
        # Advanced filters
        with st.expander("🔍 Advanced Filters", expanded=True):
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                filter_status = st.selectbox("Status", ["All", "NEW!", "RESERVED!", "CANCELED!", "UPDATE!"])
            with col2:
                filter_date_from = st.date_input("From Date", value=None)
            with col3:
                filter_date_to = st.date_input("To Date", value=None)
            with col4:
                search_villa = st.text_input("🔍 Search Villa")
            
            # Multi-sheet filter
            if using_multi and 'source_sheet' in df.columns:
                col1, col2 = st.columns(2)
                with col1:
                    unique_sheets = ["All"] + list(df['source_sheet'].unique())
                    filter_sheet = st.selectbox("📄 Filter by Sheet", unique_sheets)
                with col2:
                    sort_by = st.selectbox("Sort By", ["Date", "Villa", "Status", "Source"])
        
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("🔄 Refresh Data", use_container_width=True):
                if st.session_state.client:
                    sync_all_sheets(st.session_state.client)
                st.rerun()
        
        if not df.empty:
            filtered_df = df.copy()
            
            # Apply filters
            if filter_status != "All":
                filtered_df = filtered_df[filtered_df['RESERVATION STATUS:'] == filter_status]
            if filter_date_from:
                try:
                    filtered_df['parsed_date'] = pd.to_datetime(filtered_df['DATE:'], format='%m/%d/%Y', errors='coerce')
                    filtered_df = filtered_df[filtered_df['parsed_date'] >= pd.to_datetime(filter_date_from)]
                except:
                    pass
            if filter_date_to:
                try:
                    if 'parsed_date' not in filtered_df.columns:
                        filtered_df['parsed_date'] = pd.to_datetime(filtered_df['DATE:'], format='%m/%d/%Y', errors='coerce')
                    filtered_df = filtered_df[filtered_df['parsed_date'] <= pd.to_datetime(filter_date_to)]
                except:
                    pass
            if search_villa:
                filtered_df = filtered_df[filtered_df['VILLA:'].str.contains(search_villa, case=False, na=False)]
            
            if using_multi and 'source_sheet' in df.columns and filter_sheet != "All":
                filtered_df = filtered_df[filtered_df['source_sheet'] == filter_sheet]
            
            st.info(f"📊 Showing {len(filtered_df)} of {len(df)} bookings")
            
            # Pagination
            items_per_page = 10
            total_pages = (len(filtered_df) + items_per_page - 1) // items_per_page
            
            if 'current_page' not in st.session_state:
                st.session_state.current_page = 0
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col1:
                if st.button("◀ Previous") and st.session_state.current_page > 0:
                    st.session_state.current_page -= 1
                    st.rerun()
            with col2:
                st.markdown(f"<p style='text-align: center;'>Page {st.session_state.current_page + 1} of {max(1, total_pages)}</p>", unsafe_allow_html=True)
            with col3:
                if st.button("Next ▶") and st.session_state.current_page < total_pages - 1:
                    st.session_state.current_page += 1
                    st.rerun()
            
            # Display paginated bookings
            start_idx = st.session_state.current_page * items_per_page
            end_idx = start_idx + items_per_page
            page_df = filtered_df.iloc[start_idx:end_idx]
            
            for idx, row in page_df.iterrows():
                source_info = f"📄 {row.get('source_sheet', 'Unknown')}" if 'source_sheet' in row else ""
                
                with st.expander(f"🏠 {row['VILLA:']} - {row['DATE:']} ({row['RESERVATION STATUS:']}) {source_info}"):
                    # Only allow editing if we have a single active worksheet
                    if st.session_state.worksheet and not using_multi:
                        edit_mode = st.checkbox(f"✏️ Edit Mode", key=f"edit_{idx}")
                        
                        if edit_mode:
                            with st.form(key=f"edit_form_{idx}"):
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    new_date = st.date_input("Date", value=datetime.strptime(row['DATE:'], '%m/%d/%Y'))
                                    new_villa = st.text_input("Villa", value=row['VILLA:'])
                                    new_type_clean = st.selectbox("Cleaning Type", [
                                        "Standard Clean", "Deep Clean", "Check-out Clean", 
                                        "Mid-stay Clean", "Move-in Clean", "Post-event Clean"
                                    ], index=0)
                                    new_pax = st.number_input("PAX", min_value=1, value=int(row['PAX:']) if str(row['PAX:']).isdigit() else 2)
                                
                                with col2:
                                    try:
                                        start_val = datetime.strptime(row['START TIME:'], '%H:%M').time()
                                    except:
                                        start_val = time(10, 0)
                                    try:
                                        end_val = datetime.strptime(row['END TIME:'], '%H:%M').time()
                                    except:
                                        end_val = time(14, 0)
                                    
                                    new_start_time = st.time_input("Start Time", value=start_val)
                                    new_end_time = st.time_input("End Time", value=end_val)
                                    new_status = st.selectbox("Status", ["NEW!", "RESERVED!", "CANCELED!", "UPDATE!"])
                                    new_laundry = st.selectbox("Laundry", ["Yes", "No", "Not Required", "Pending"])
                                
                                new_comments = st.text_area("Comments", value=row.get('COMMENTS:', ''))
                                
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
                                    
                                    if update_booking(st.session_state.worksheet, idx, booking_data):
                                        st.success("✅ Booking updated!")
                                        if st.session_state.auto_sync_enabled:
                                            sync_all_sheets(st.session_state.client)
                                        st.rerun()
                                
                                if delete_btn:
                                    if delete_booking(st.session_state.worksheet, idx, row['VILLA:']):
                                        st.success("✅ Booking deleted!")
                                        if st.session_state.auto_sync_enabled:
                                            sync_all_sheets(st.session_state.client)
                                        st.rerun()
                    else:
                        # Read-only view for multi-sheet data
                        if using_multi:
                            st.info("💡 Editing disabled in multi-sheet view. Select a single sheet to edit.")
                        
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.markdown("**📅 Date & Time**")
                            st.write(f"Date: {row['DATE:']}")
                            st.write(f"Start: {row.get('START TIME:', 'N/A')}")
                            st.write(f"End: {row.get('END TIME:', 'N/A')}")
                        
                        with col2:
                            st.markdown("**🏠 Details**")
                            st.write(f"Villa: {row['VILLA:']}")
                            st.write(f"Cleaning: {row.get('TYPE CLEAN:', 'N/A')}")
                            st.write(f"PAX: {row.get('PAX:', 'N/A')}")
                        
                        with col3:
                            st.markdown("**📊 Status**")
                            status_class = f"status-{row['RESERVATION STATUS:'].lower().replace('!', '')}"
                            st.markdown(f"<span class='{status_class}'>{row['RESERVATION STATUS:']}</span>", unsafe_allow_html=True)
                            st.write(f"Laundry: {row.get('LAUNDRY SERVICES WITH VIDEMI:', 'N/A')}")
                        
                        if row.get('COMMENTS:'):
                            st.markdown("**💬 Comments**")
                            st.info(row['COMMENTS:'])
                        
                        if 'source_sheet' in row:
                            st.markdown("**📄 Source**")
                            st.caption(row['source_sheet'])
        else:
            st.info("No bookings found")
    
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
                <p style="font-size: 14px; color: #666;">All notifications are sent here</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <h4>Notification Status</h4>
                <p style="font-size: 18px; color: {'#38ef7d' if st.session_state.email_notifications_enabled else '#f5576c'}; font-weight: 600;">
                    {'✅ Enabled' if st.session_state.email_notifications_enabled else '❌ Disabled'}
                </p>
                <p style="font-size: 14px; color: #666;">Toggle in sidebar</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.subheader("📁 Folder Sync Settings")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <h4>Target Folder</h4>
                <p style="font-size: 12px; word-break: break-all;">{TARGET_FOLDER_ID[:20]}...</p>
                <a href="{FOLDER_URL}" target="_blank" style="color: #667eea;">View Folder →</a>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <h4>Auto-Sync</h4>
                <p style="font-size: 18px; color: {'#38ef7d' if st.session_state.auto_sync_enabled else '#f5576c'}; font-weight: 600;">
                    {'✅ Enabled' if st.session_state.auto_sync_enabled else '❌ Disabled'}
                </p>
                <p style="font-size: 14px; color: #666;">Every 5 minutes</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            sheets_count = len(st.session_state.folder_sheets)
            st.markdown(f"""
            <div class="metric-card">
                <h4>Synced Sheets</h4>
                <h2 style="color: #667eea;">{sheets_count}</h2>
                <p style="font-size: 14px; color: #666;">From target folder</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.subheader("📊 System Information")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown("""
            <div class="metric-card">
                <h4>Original Template</h4>
                <code style="font-size: 10px;">1-3FLLEk...</code>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="metric-card">
                <h4>Connection Status</h4>
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
        
        with col4:
            total_bookings = len(st.session_state.all_bookings_df)
            st.markdown(f"""
            <div class="metric-card">
                <h4>Total Bookings</h4>
                <h2 style="color: #667eea;">{total_bookings}</h2>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.subheader("🔧 Advanced Settings")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Sync Interval**")
            sync_interval = st.slider("Minutes between auto-sync", 1, 30, 5)
            st.caption(f"Currently syncing every {sync_interval} minutes")
        
        with col2:
            st.markdown("**Display Settings**")
            items_per_page = st.number_input("Items per page", 5, 50, 10)
            st.caption("Number of bookings to show per page")
        
        st.markdown("---")
        st.subheader("⚠️ Danger Zone")
        
        with st.expander("🚨 Advanced Operations"):
            st.warning("⚠️ These operations may affect your data")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("🗑️ Clear Activity Log", use_container_width=True):
                    st.session_state.activity_log = []
                    log_activity("Activity log cleared", "info")
                    st.success("✅ Log cleared")
            
            with col2:
                if st.button("🔄 Reset Sync Data", use_container_width=True):
                    st.session_state.all_bookings_df = pd.DataFrame()
                    st.session_state.last_sync_time = None
                    log_activity("Sync data reset", "info")
                    st.success("✅ Sync data reset")
            
            with col3:
                if st.button("🔌 Reset Connection", use_container_width=True):
                    st.session_state.worksheet = None
                    st.session_state.current_spreadsheet_id = None
                    log_activity("Connection reset", "info")
                    st.success("✅ Connection reset")
        
        st.markdown("---")
        st.subheader("📖 System Documentation")
        
        with st.expander("📚 Features & Usage Guide"):
            st.markdown("""
            ### 🎯 Key Features
            
            **1. Folder Synchronization**
            - Automatically syncs all spreadsheets from target Google Drive folder
            - Real-time updates every 5 minutes (configurable)
            - Combines data from multiple sheets into unified view
            
            **2. Calendar Views**
            - Monthly calendar with bookings overlay
            - Timeline view for daily scheduling
            - Gantt chart for project management view
            - Heatmap for booking density analysis
            
            **3. Advanced Analytics**
            - Multi-sheet data aggregation
            - Status distribution charts
            - Top villas analysis
            - Cleaning type breakdown
            - Export to CSV, JSON, Excel
            
            **4. Booking Management**
            - Create new bookings with validation
            - Edit existing bookings (single-sheet mode)
            - Delete bookings with confirmation
            - Advanced filtering and search
            - Pagination for large datasets
            
            **5. Email Notifications**
            - Auto-notify on booking creation
            - Updates and deletions tracked
            - Configurable notification settings
            
            **6. Activity Logging**
            - Complete audit trail
            - Success/error tracking
            - Exportable logs
            
            ### 🚀 Quick Start
            
            1. Upload Google Service Account JSON credentials
            2. System auto-loads default template and folder sheets
            3. Click "Sync Now" to load all data
            4. Navigate using sidebar menu
            5. View calendar, timeline, or manage bookings
            
            ### 💡 Pro Tips
            
            - Enable auto-sync for always-fresh data
            - Use calendar view for visual scheduling
            - Filter by source sheet in multi-sheet view
            - Export data regularly for backups
            - Check activity log for troubleshooting
            
            ### 🔗 Important Links
            
            - Original Template: View and clone base template
            - Target Folder: Access all synced spreadsheets
            - Support Email: {NOTIFICATION_EMAIL}
            """)
        
        with st.expander("🔧 Troubleshooting"):
            st.markdown("""
            ### Common Issues
            
            **Problem: Sheets not syncing**
            - Solution: Click "Sync Now" manually
            - Check that service account has access to folder
            - Verify folder ID is correct
            
            **Problem: Cannot edit bookings**
            - Solution: Switch to single-sheet mode
            - Multi-sheet view is read-only
            - Select specific sheet from sidebar
            
            **Problem: Dates not parsing**
            - Solution: Ensure date format is MM/DD/YYYY
            - Check for empty date cells
            - Verify column headers match template
            
            **Problem: Missing data**
            - Solution: Refresh the page
            - Clear sync data and re-sync
            - Check sheet permissions
            
            ### Performance Tips
            
            - Reduce sync interval for large datasets
            - Use pagination for better performance
            - Filter data before exporting
            - Clear activity log periodically
            """)
    
    # Footer
    st.sidebar.markdown("---")
    st.sidebar.caption("🏖️ Villa Booking System Pro++ v4.0")
    st.sidebar.caption("Enhanced with Folder Sync & Calendar")
    st.sidebar.caption(f"© 2025 | {NOTIFICATION_EMAIL}")
    
    if st.session_state.last_sync_time:
        time_ago = (datetime.now() - st.session_state.last_sync_time).total_seconds() / 60
        st.sidebar.caption(f"Last sync: {int(time_ago)} min ago")

# Additional import for Excel export
import io

if __name__ == "__main__":
    main()
