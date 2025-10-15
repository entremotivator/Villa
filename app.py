import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import json

# Page configuration
st.set_page_config(
    page_title="Booking Manager",
    page_icon="📅",
    layout="wide"
)

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'gc' not in st.session_state:
    st.session_state.gc = None
if 'workbooks' not in st.session_state:
    st.session_state.workbooks = []
if 'selected_workbook' not in st.session_state:
    st.session_state.selected_workbook = None
if 'logs' not in st.session_state:
    st.session_state.logs = []

def log_activity(message):
    """Add activity to logs"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.logs.append(f"[{timestamp}] {message}")

def authenticate_google_sheets(credentials_dict, folder_id=None):
    """Authenticate with Google Sheets and Drive API using service account"""
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
        
        creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        gc = gspread.authorize(creds)
        
        # Get workbooks from folder or all accessible
        if folder_id:
            log_activity(f"Loading workbooks from folder: {folder_id}")
            workbooks = gc.list_spreadsheet_files(folder_id=folder_id)
        else:
            log_activity("Loading all accessible workbooks")
            workbooks = gc.openall()
            workbooks = [{'id': wb.id, 'name': wb.title} for wb in workbooks]
        
        log_activity(f"Successfully loaded {len(workbooks)} workbook(s)")
        return gc, workbooks
    except Exception as e:
        log_activity(f"Authentication error: {str(e)}")
        st.error(f"Authentication failed: {str(e)}")
        return None, []

def load_client_profile(workbook):
    """Load client profile from first sheet"""
    try:
        first_sheet = workbook.get_worksheet(0)
        data = first_sheet.get_all_values()
        
        profile = {}
        for row in data[:20]:  # Check first 20 rows for profile data
            if len(row) >= 2 and row[0]:
                key = row[0].strip()
                value = row[1].strip() if len(row) > 1 else ""
                if key and value:
                    profile[key] = value
        
        log_activity(f"Loaded client profile with {len(profile)} fields")
        return profile
    except Exception as e:
        log_activity(f"Error loading client profile: {str(e)}")
        return {}

def load_calendar_data(workbook, sheet_index):
    """Load booking calendar data from a specific sheet (starting from row 13)"""
    try:
        sheet = workbook.get_worksheet(sheet_index)
        sheet_name = sheet.title
        
        # Get all data
        all_data = sheet.get_all_values()
        
        if len(all_data) < 13:
            return pd.DataFrame(), sheet_name
        
        # Headers are in row 12 (index 11), data starts at row 13 (index 12)
        headers = all_data[11]
        data_rows = all_data[12:]
        
        # Create DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Remove empty rows
        df = df[df.iloc[:, 0].str.strip() != '']
        
        log_activity(f"Loaded {len(df)} bookings from sheet: {sheet_name}")
        return df, sheet_name
    except Exception as e:
        log_activity(f"Error loading calendar data: {str(e)}")
        return pd.DataFrame(), "Unknown"

def get_all_properties(profile):
    """Extract property names from client profile"""
    properties = []
    for key, value in profile.items():
        if 'property' in key.lower() or 'address' in key.lower():
            if value and value not in properties:
                properties.append(value)
    return properties

def get_status_color(status):
    """Return color for status badge"""
    colors = {
        'CI': '#4CAF50',    # Green - Check In
        'SO': '#2196F3',    # Blue - Stay Over
        'CO/CI': '#FF9800', # Orange - Check Out/Check In
        'FU': '#9C27B0',    # Purple - Follow Up
        'DC': '#F44336',    # Red - Deep Clean
        'COC': '#00BCD4',   # Cyan - Check Out Clean
        'CO': '#FF5722'     # Deep Orange - Check Out
    }
    return colors.get(status, '#757575')

# Main App
st.title("📅 Booking Management System")

# Sidebar
with st.sidebar:
    st.header("🔐 Authentication")
    
    # <CHANGE> Only JSON file upload authentication
    uploaded_file = st.file_uploader(
        "Upload Service Account JSON",
        type=['json'],
        help="Upload your Google Service Account credentials JSON file"
    )
    
    folder_id = st.text_input(
        "Google Drive Folder ID (Optional)",
        help="Enter folder ID to load workbooks from a specific folder. Leave empty to load all accessible workbooks."
    )
    
    if uploaded_file is not None:
        if st.button("🔓 Authenticate", type="primary"):
            try:
                credentials_dict = json.load(uploaded_file)
                gc, workbooks = authenticate_google_sheets(credentials_dict, folder_id if folder_id else None)
                
                if gc and workbooks:
                    st.session_state.authenticated = True
                    st.session_state.gc = gc
                    st.session_state.workbooks = workbooks
                    st.success("✅ Authentication successful!")
                    st.rerun()
                else:
                    st.error("❌ No workbooks found")
            except json.JSONDecodeError:
                st.error("❌ Invalid JSON file")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    
    if st.session_state.authenticated:
        st.success("✅ Authenticated")
        
        if st.button("🔄 Refresh Workbooks"):
            try:
                if folder_id:
                    workbooks = st.session_state.gc.list_spreadsheet_files(folder_id=folder_id)
                else:
                    workbooks = st.session_state.gc.openall()
                    workbooks = [{'id': wb.id, 'name': wb.title} for wb in workbooks]
                
                st.session_state.workbooks = workbooks
                log_activity(f"Refreshed workbooks: {len(workbooks)} found")
                st.rerun()
            except Exception as e:
                st.error(f"Error refreshing: {str(e)}")
        
        st.divider()
        
        # Workbook selector
        if st.session_state.workbooks:
            st.subheader("📚 Select Workbook")
            workbook_names = [wb['name'] for wb in st.session_state.workbooks]
            selected_name = st.selectbox("Client Workbook", workbook_names)
            
            if selected_name:
                selected_wb = next(wb for wb in st.session_state.workbooks if wb['name'] == selected_name)
                st.session_state.selected_workbook = selected_wb
                
                st.info(f"📖 **{selected_name}**")
        else:
            st.warning("No workbooks found")
        
        if st.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.session_state.gc = None
            st.session_state.workbooks = []
            st.session_state.selected_workbook = None
            st.rerun()

# Main content
if not st.session_state.authenticated:
    st.info("👈 Please upload your service account JSON file and authenticate to begin")
    
    with st.expander("ℹ️ How to get started"):
        st.markdown("""
        ### Setup Instructions:
        
        1. **Create a Google Cloud Service Account**
           - Go to [Google Cloud Console](https://console.cloud.google.com/)
           - Create a new project or select existing
           - Enable Google Sheets API and Google Drive API
           - Create a Service Account and download JSON key
        
        2. **Share your Google Sheets**
           - Open your Google Drive folder with booking workbooks
           - Share the folder with the service account email (found in JSON file)
           - Give "Viewer" or "Editor" permissions
        
        3. **Get Folder ID (Optional)**
           - Open your Google Drive folder
           - Copy the ID from URL: `drive.google.com/drive/folders/FOLDER_ID_HERE`
        
        4. **Upload JSON and Authenticate**
           - Upload the service account JSON file
           - Optionally enter folder ID
           - Click Authenticate
        """)

elif not st.session_state.selected_workbook:
    st.warning("📚 Please select a workbook from the sidebar")

else:
    # Load selected workbook
    try:
        workbook = st.session_state.gc.open_by_key(st.session_state.selected_workbook['id'])
        
        # Navigation tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Dashboard",
            "📅 Calendar View",
            "➕ Booking Manager",
            "📈 Analytics",
            "📋 System Logs"
        ])
        
        with tab1:
            st.header("📊 Dashboard Overview")
            
            # Load client profile
            profile = load_client_profile(workbook)
            
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader("👤 Client Profile")
                if profile:
                    for key, value in profile.items():
                        st.text(f"{key}: {value}")
                else:
                    st.info("No profile data found")
            
            with col2:
                st.subheader("📈 Quick Stats")
                
                # Count total bookings across all calendar sheets
                total_bookings = 0
                sheet_count = workbook.worksheet_count - 1  # Exclude profile sheet
                
                for i in range(1, workbook.worksheet_count):
                    df, _ = load_calendar_data(workbook, i)
                    total_bookings += len(df)
                
                st.metric("Total Calendar Sheets", sheet_count)
                st.metric("Total Bookings", total_bookings)
                
                # Properties
                properties = get_all_properties(profile)
                st.metric("Properties", len(properties))
            
            # Recent bookings preview
            st.subheader("📅 Recent Bookings Preview")
            if workbook.worksheet_count > 1:
                df, sheet_name = load_calendar_data(workbook, 1)
                if not df.empty:
                    st.caption(f"From: {sheet_name}")
                    st.dataframe(df.head(10), use_container_width=True)
                else:
                    st.info("No bookings found in first calendar sheet")
        
        with tab2:
            st.header("📅 Calendar View")
            
            # Sheet selector
            sheet_names = [workbook.get_worksheet(i).title for i in range(1, workbook.worksheet_count)]
            
            if sheet_names:
                selected_sheet_name = st.selectbox("Select Month", sheet_names)
                selected_sheet_index = sheet_names.index(selected_sheet_name) + 1
                
                # Load calendar data
                df, sheet_name = load_calendar_data(workbook, selected_sheet_index)
                
                if not df.empty:
                    # Filters
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Status filter
                        if 'Status' in df.columns or 'status' in df.columns:
                            status_col = 'Status' if 'Status' in df.columns else 'status'
                            statuses = ['All'] + sorted(df[status_col].unique().tolist())
                            selected_status = st.selectbox("Filter by Status", statuses)
                    
                    with col2:
                        # Property filter
                        properties = get_all_properties(load_client_profile(workbook))
                        if properties:
                            selected_property = st.selectbox("Filter by Property", ['All'] + properties)
                    
                    # Apply filters
                    filtered_df = df.copy()
                    
                    if selected_status != 'All':
                        status_col = 'Status' if 'Status' in df.columns else 'status'
                        filtered_df = filtered_df[filtered_df[status_col] == selected_status]
                    
                    st.subheader(f"📆 {sheet_name}")
                    st.caption(f"Showing {len(filtered_df)} of {len(df)} bookings")
                    
                    # Display with color coding
                    st.dataframe(filtered_df, use_container_width=True, height=500)
                    
                    # Status legend
                    st.caption("**Status Codes:**")
                    status_codes = {
                        'CI': 'Check In',
                        'SO': 'Stay Over',
                        'CO/CI': 'Check Out / Check In',
                        'FU': 'Follow Up',
                        'DC': 'Deep Clean',
                        'COC': 'Check Out Clean',
                        'CO': 'Check Out'
                    }
                    
                    cols = st.columns(7)
                    for idx, (code, description) in enumerate(status_codes.items()):
                        with cols[idx]:
                            st.markdown(f"**{code}** - {description}")
                else:
                    st.info("No bookings found in this calendar sheet")
            else:
                st.warning("No calendar sheets found")
        
        with tab3:
            st.header("➕ Booking Manager")
            
            st.info("📝 Create new bookings (Feature in development)")
            
            with st.form("new_booking_form"):
                col1, col2 = st.columns(2)
                
                with col1:
                    booking_date = st.date_input("Date")
                    guest_name = st.text_input("Guest Name")
                    
                    properties = get_all_properties(load_client_profile(workbook))
                    property_select = st.selectbox("Property", properties if properties else ["No properties found"])
                
                with col2:
                    status = st.selectbox("Status", ['CI', 'SO', 'CO/CI', 'FU', 'DC', 'COC', 'CO'])
                    check_in_time = st.time_input("Check-in Time")
                    check_out_time = st.time_input("Check-out Time")
                
                notes = st.text_area("Notes")
                
                submitted = st.form_submit_button("Create Booking", type="primary")
                
                if submitted:
                    st.success("✅ Booking creation feature coming soon!")
                    log_activity(f"Booking form submitted for {guest_name} on {booking_date}")
        
        with tab4:
            st.header("📈 Analytics & Reports")
            
            # Collect all booking data
            all_bookings = []
            for i in range(1, workbook.worksheet_count):
                df, sheet_name = load_calendar_data(workbook, i)
                if not df.empty:
                    df['Sheet'] = sheet_name
                    all_bookings.append(df)
            
            if all_bookings:
                combined_df = pd.concat(all_bookings, ignore_index=True)
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total Bookings", len(combined_df))
                
                with col2:
                    st.metric("Calendar Sheets", len(all_bookings))
                
                with col3:
                    if 'Status' in combined_df.columns or 'status' in combined_df.columns:
                        status_col = 'Status' if 'Status' in combined_df.columns else 'status'
                        unique_statuses = combined_df[status_col].nunique()
                        st.metric("Unique Statuses", unique_statuses)
                
                # Status distribution
                if 'Status' in combined_df.columns or 'status' in combined_df.columns:
                    st.subheader("📊 Status Distribution")
                    status_col = 'Status' if 'Status' in combined_df.columns else 'status'
                    status_counts = combined_df[status_col].value_counts()
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        st.bar_chart(status_counts)
                    
                    with col2:
                        for status, count in status_counts.items():
                            percentage = (count / len(combined_df)) * 100
                            st.markdown(f"**{status}**: {count} ({percentage:.1f}%)")
                
                # Export option
                st.subheader("💾 Export Data")
                csv = combined_df.to_csv(index=False)
                st.download_button(
                    label="📥 Download All Bookings as CSV",
                    data=csv,
                    file_name=f"bookings_export_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("No booking data available for analytics")
        
        with tab5:
            st.header("📋 System Activity Logs")
            
            if st.session_state.logs:
                st.text_area("Activity Log", "\n".join(reversed(st.session_state.logs[-50:])), height=400)
                
                if st.button("🗑️ Clear Logs"):
                    st.session_state.logs = []
                    st.rerun()
            else:
                st.info("No activity logged yet")
    
    except Exception as e:
        st.error(f"Error loading workbook: {str(e)}")
        log_activity(f"Error loading workbook: {str(e)}")
