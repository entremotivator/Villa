import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import json

# Configuration
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'
]

# <CHANGE> Hardcoded Google Drive folder ID
DRIVE_FOLDER_ID = "1Fk5dJGkm5dNMZkfsITe5Lt9x-yCsBiF2"

# Status code definitions
STATUS_CODES = {
    'CI': {'name': 'Check In', 'color': '#4CAF50'},
    'SO': {'name': 'Stay Over', 'color': '#2196F3'},
    'CO/CI': {'name': 'Check Out/Check In', 'color': '#FF9800'},
    'FU': {'name': 'Follow Up', 'color': '#9C27B0'},
    'DC': {'name': 'Deep Clean', 'color': '#F44336'},
    'COC': {'name': 'Check Out Clean', 'color': '#FF5722'},
    'CO': {'name': 'Check Out', 'color': '#795548'}
}

# Page config
st.set_page_config(
    page_title="Booking Management System",
    page_icon="📅",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: bold;
    }
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    .status-badge {
        padding: 0.25rem 0.75rem;
        border-radius: 15px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
        color: white;
    }
    .info-card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #1f77b4;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'client' not in st.session_state:
    st.session_state.client = None
if 'workbooks' not in st.session_state:
    st.session_state.workbooks = []
if 'selected_workbook' not in st.session_state:
    st.session_state.selected_workbook = None
if 'logs' not in st.session_state:
    st.session_state.logs = []

def add_log(message, level="INFO"):
    """Add a log entry"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.logs.append({
        'timestamp': timestamp,
        'level': level,
        'message': message
    })

def authenticate_google(credentials_dict):
    """Authenticate with Google Sheets and Drive API"""
    try:
        creds = Credentials.from_service_account_info(credentials_dict, scopes=SCOPES)
        client = gspread.authorize(creds)
        add_log("Successfully authenticated with Google APIs", "SUCCESS")
        return client
    except Exception as e:
        add_log(f"Authentication failed: {str(e)}", "ERROR")
        return None

def load_workbooks_from_folder(client):
    """Load all spreadsheets from the specified Google Drive folder"""
    try:
        # <CHANGE> Using hardcoded folder ID
        add_log(f"Loading workbooks from folder: {DRIVE_FOLDER_ID}", "INFO")
        
        # Query spreadsheets in the specific folder
        query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
        files = client.list_spreadsheet_files(query=query)
        
        workbooks = []
        for file in files:
            workbooks.append({
                'id': file['id'],
                'name': file['name']
            })
        
        add_log(f"Found {len(workbooks)} workbooks in folder", "SUCCESS")
        return workbooks
    except Exception as e:
        add_log(f"Error loading workbooks: {str(e)}", "ERROR")
        return []

def get_client_profile(workbook):
    """Extract client profile from first sheet"""
    try:
        first_sheet = workbook.get_worksheet(0)
        data = first_sheet.get_all_values()
        
        profile = {}
        properties = []
        
        # Parse client info (rows 1-10)
        for i in range(min(10, len(data))):
            if len(data[i]) >= 2 and data[i][0]:
                key = data[i][0].strip()
                value = data[i][1].strip() if len(data[i]) > 1 else ""
                if key and value:
                    profile[key] = value
        
        # Parse properties (starting from row 11)
        if len(data) > 10:
            for i in range(10, len(data)):
                if len(data[i]) >= 2 and data[i][0]:
                    prop_name = data[i][0].strip()
                    prop_address = data[i][1].strip() if len(data[i]) > 1 else ""
                    if prop_name:
                        properties.append({
                            'name': prop_name,
                            'address': prop_address
                        })
        
        add_log(f"Loaded client profile with {len(properties)} properties", "INFO")
        return profile, properties
    except Exception as e:
        add_log(f"Error loading client profile: {str(e)}", "ERROR")
        return {}, []

def get_calendar_sheets(workbook):
    """Get all calendar sheets (excluding first sheet)"""
    try:
        all_sheets = workbook.worksheets()
        calendar_sheets = all_sheets[1:]  # Skip first sheet (client profile)
        add_log(f"Found {len(calendar_sheets)} calendar sheets", "INFO")
        return calendar_sheets
    except Exception as e:
        add_log(f"Error loading calendar sheets: {str(e)}", "ERROR")
        return []

def parse_calendar_data(sheet):
    """Parse booking data from calendar sheet (starting at row 13)"""
    try:
        data = sheet.get_all_values()
        
        if len(data) < 13:
            return pd.DataFrame()
        
        # Get headers from row 12 (index 11)
        headers = data[11]
        
        # Get booking data from row 13 onwards (index 12+)
        booking_rows = data[12:]
        
        # Create DataFrame
        df = pd.DataFrame(booking_rows, columns=headers)
        
        # Clean up empty rows
        df = df[df.iloc[:, 0].str.strip() != '']
        
        add_log(f"Parsed {len(df)} bookings from {sheet.title}", "INFO")
        return df
    except Exception as e:
        add_log(f"Error parsing calendar data: {str(e)}", "ERROR")
        return pd.DataFrame()

def get_all_bookings(workbook):
    """Get all bookings from all calendar sheets"""
    all_bookings = []
    calendar_sheets = get_calendar_sheets(workbook)
    
    for sheet in calendar_sheets:
        df = parse_calendar_data(sheet)
        if not df.empty:
            df['Month'] = sheet.title
            all_bookings.append(df)
    
    if all_bookings:
        combined_df = pd.concat(all_bookings, ignore_index=True)
        add_log(f"Total bookings loaded: {len(combined_df)}", "SUCCESS")
        return combined_df
    
    return pd.DataFrame()

def create_booking(workbook, month_sheet_name, booking_data):
    """Create a new booking entry"""
    try:
        sheet = workbook.worksheet(month_sheet_name)
        data = sheet.get_all_values()
        
        # Find next empty row after row 13
        next_row = 13
        for i in range(12, len(data)):
            if not data[i][0].strip():
                next_row = i + 1
                break
        else:
            next_row = len(data) + 1
        
        # Prepare row data
        row_data = [
            booking_data.get('date', ''),
            booking_data.get('property', ''),
            booking_data.get('status', ''),
            booking_data.get('guest_name', ''),
            booking_data.get('notes', '')
        ]
        
        sheet.insert_row(row_data, next_row)
        add_log(f"Created booking on {booking_data.get('date')} for {booking_data.get('property')}", "SUCCESS")
        return True
    except Exception as e:
        add_log(f"Error creating booking: {str(e)}", "ERROR")
        return False

# Main App
st.markdown('<div class="main-header">📅 Booking Management System</div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("🔐 Authentication")
    
    # <CHANGE> Removed folder ID input field since it's now hardcoded
    st.info(f"📁 Using folder: ...{DRIVE_FOLDER_ID[-20:]}")
    
    if not st.session_state.authenticated:
        st.write("Upload your service account JSON file:")
        uploaded_file = st.file_uploader("Service Account JSON", type=['json'], label_visibility="collapsed")
        
        if uploaded_file is not None:
            try:
                credentials_dict = json.load(uploaded_file)
                
                if st.button("🔓 Authenticate", use_container_width=True):
                    with st.spinner("Authenticating..."):
                        client = authenticate_google(credentials_dict)
                        
                        if client:
                            st.session_state.client = client
                            st.session_state.authenticated = True
                            
                            # Load workbooks from folder
                            workbooks = load_workbooks_from_folder(client)
                            st.session_state.workbooks = workbooks
                            
                            st.success("✅ Authenticated successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Authentication failed. Check your credentials.")
            except Exception as e:
                st.error(f"❌ Error reading JSON file: {str(e)}")
    else:
        st.success("✅ Authenticated")
        
        if st.button("🔄 Refresh Workbooks", use_container_width=True):
            with st.spinner("Refreshing..."):
                workbooks = load_workbooks_from_folder(st.session_state.client)
                st.session_state.workbooks = workbooks
                st.rerun()
        
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.client = None
            st.session_state.workbooks = []
            st.session_state.selected_workbook = None
            st.rerun()
    
    st.divider()
    
    # Workbook selector
    if st.session_state.authenticated and st.session_state.workbooks:
        st.header("📚 Select Workbook")
        
        workbook_names = [wb['name'] for wb in st.session_state.workbooks]
        selected_name = st.selectbox(
            "Client Workbook",
            workbook_names,
            key="workbook_selector"
        )
        
        if selected_name:
            selected_wb = next(wb for wb in st.session_state.workbooks if wb['name'] == selected_name)
            
            if st.session_state.selected_workbook != selected_wb['id']:
                st.session_state.selected_workbook = selected_wb['id']
                st.rerun()
    
    st.divider()
    
    # Navigation
    st.header("🧭 Navigation")
    view = st.radio(
        "Select View",
        ["Dashboard", "Calendar View", "Booking Manager", "Analytics", "System Logs"],
        label_visibility="collapsed"
    )

# Main content area
if not st.session_state.authenticated:
    st.info("👈 Please authenticate using your service account JSON file in the sidebar.")
    
    with st.expander("ℹ️ How to get started"):
        st.markdown("""
        ### Setup Instructions:
        
        1. **Create a Google Cloud Service Account**
           - Go to Google Cloud Console
           - Create a new service account
           - Download the JSON credentials file
        
        2. **Enable APIs**
           - Enable Google Sheets API
           - Enable Google Drive API
        
        3. **Share Access**
           - Share your Google Drive folder with the service account email
           - Grant "Viewer" or "Editor" permissions
        
        4. **Upload Credentials**
           - Upload the JSON file in the sidebar
           - Click "Authenticate"
        
        5. **Start Managing**
           - Select a workbook from the dropdown
           - View and manage bookings
        """)

elif not st.session_state.workbooks:
    st.warning("⚠️ No workbooks found in the specified folder. Please check folder permissions.")

elif not st.session_state.selected_workbook:
    st.info("👈 Please select a workbook from the sidebar to get started.")

else:
    # Load selected workbook
    try:
        workbook = st.session_state.client.open_by_key(st.session_state.selected_workbook)
        profile, properties = get_client_profile(workbook)
        
        # Dashboard View
        if view == "Dashboard":
            st.header("📊 Dashboard")
            
            # Client info
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader("👤 Client Profile")
                for key, value in profile.items():
                    st.markdown(f"**{key}:** {value}")
            
            with col2:
                st.subheader("🏠 Properties")
                for prop in properties:
                    st.markdown(f"**{prop['name']}**")
                    st.caption(prop['address'])
                    st.divider()
            
            # Booking statistics
            st.subheader("📈 Booking Overview")
            
            all_bookings = get_all_bookings(workbook)
            
            if not all_bookings.empty:
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{len(all_bookings)}</div>
                        <div class="metric-label">Total Bookings</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    calendar_count = len(get_calendar_sheets(workbook))
                    st.markdown(f"""
                    <div class="metric-card" style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);">
                        <div class="metric-value">{calendar_count}</div>
                        <div class="metric-label">Calendar Months</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    property_count = len(properties)
                    st.markdown(f"""
                    <div class="metric-card" style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);">
                        <div class="metric-value">{property_count}</div>
                        <div class="metric-label">Properties</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    # Count unique status codes
                    if 'Status' in all_bookings.columns or len(all_bookings.columns) > 2:
                        status_col = 'Status' if 'Status' in all_bookings.columns else all_bookings.columns[2]
                        unique_statuses = all_bookings[status_col].nunique()
                    else:
                        unique_statuses = 0
                    
                    st.markdown(f"""
                    <div class="metric-card" style="background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);">
                        <div class="metric-value">{unique_statuses}</div>
                        <div class="metric-label">Status Types</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.divider()
                
                # Status distribution
                st.subheader("📊 Status Distribution")
                if len(all_bookings.columns) > 2:
                    status_col = 'Status' if 'Status' in all_bookings.columns else all_bookings.columns[2]
                    status_counts = all_bookings[status_col].value_counts()
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        st.bar_chart(status_counts)
                    
                    with col2:
                        for status, count in status_counts.items():
                            status_info = STATUS_CODES.get(status, {'name': status, 'color': '#999999'})
                            st.markdown(f"""
                            <div style="margin: 0.5rem 0;">
                                <span class="status-badge" style="background-color: {status_info['color']};">
                                    {status}
                                </span>
                                <span style="margin-left: 0.5rem; font-weight: 600;">{count} bookings</span>
                            </div>
                            """, unsafe_allow_html=True)
            else:
                st.info("No bookings found in this workbook.")
        
        # Calendar View
        elif view == "Calendar View":
            st.header("📅 Calendar View")
            
            calendar_sheets = get_calendar_sheets(workbook)
            
            if calendar_sheets:
                # Month selector
                month_names = [sheet.title for sheet in calendar_sheets]
                selected_month = st.selectbox("Select Month", month_names)
                
                if selected_month:
                    sheet = workbook.worksheet(selected_month)
                    df = parse_calendar_data(sheet)
                    
                    if not df.empty:
                        # Filters
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if len(df.columns) > 2:
                                status_col = 'Status' if 'Status' in df.columns else df.columns[2]
                                unique_statuses = ['All'] + sorted(df[status_col].unique().tolist())
                                filter_status = st.selectbox("Filter by Status", unique_statuses)
                        
                        with col2:
                            if len(df.columns) > 1:
                                property_col = 'Property' if 'Property' in df.columns else df.columns[1]
                                unique_properties = ['All'] + sorted(df[property_col].unique().tolist())
                                filter_property = st.selectbox("Filter by Property", unique_properties)
                        
                        # Apply filters
                        filtered_df = df.copy()
                        
                        if filter_status != 'All' and len(df.columns) > 2:
                            status_col = 'Status' if 'Status' in df.columns else df.columns[2]
                            filtered_df = filtered_df[filtered_df[status_col] == filter_status]
                        
                        if filter_property != 'All' and len(df.columns) > 1:
                            property_col = 'Property' if 'Property' in df.columns else df.columns[1]
                            filtered_df = filtered_df[filtered_df[property_col] == filter_property]
                        
                        st.info(f"Showing {len(filtered_df)} of {len(df)} bookings")
                        
                        # Display data
                        st.dataframe(filtered_df, use_container_width=True, height=600)
                    else:
                        st.info(f"No bookings found for {selected_month}")
            else:
                st.warning("No calendar sheets found in this workbook.")
        
        # Booking Manager
        elif view == "Booking Manager":
            st.header("➕ Create New Booking")
            
            calendar_sheets = get_calendar_sheets(workbook)
            
            if calendar_sheets and properties:
                with st.form("new_booking_form"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        month_names = [sheet.title for sheet in calendar_sheets]
                        selected_month = st.selectbox("Month", month_names)
                        
                        booking_date = st.date_input("Date")
                        
                        property_names = [prop['name'] for prop in properties]
                        selected_property = st.selectbox("Property", property_names)
                    
                    with col2:
                        status_options = list(STATUS_CODES.keys())
                        selected_status = st.selectbox("Status", status_options)
                        
                        # Show status info
                        status_info = STATUS_CODES[selected_status]
                        st.markdown(f"""
                        <div class="info-card">
                            <strong>{status_info['name']}</strong>
                            <span class="status-badge" style="background-color: {status_info['color']}; margin-left: 0.5rem;">
                                {selected_status}
                            </span>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        guest_name = st.text_input("Guest Name")
                    
                    notes = st.text_area("Notes")
                    
                    submitted = st.form_submit_button("Create Booking", use_container_width=True)
                    
                    if submitted:
                        booking_data = {
                            'date': booking_date.strftime("%Y-%m-%d"),
                            'property': selected_property,
                            'status': selected_status,
                            'guest_name': guest_name,
                            'notes': notes
                        }
                        
                        with st.spinner("Creating booking..."):
                            success = create_booking(workbook, selected_month, booking_data)
                            
                            if success:
                                st.success("✅ Booking created successfully!")
                                st.balloons()
                            else:
                                st.error("❌ Failed to create booking. Check system logs.")
            else:
                st.warning("Please ensure the workbook has calendar sheets and properties defined.")
        
        # Analytics
        elif view == "Analytics":
            st.header("📊 Analytics & Reports")
            
            all_bookings = get_all_bookings(workbook)
            
            if not all_bookings.empty:
                # Export option
                col1, col2, col3 = st.columns([2, 1, 1])
                
                with col3:
                    csv = all_bookings.to_csv(index=False)
                    st.download_button(
                        label="📥 Export to CSV",
                        data=csv,
                        file_name=f"bookings_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                st.divider()
                
                # Bookings by month
                st.subheader("📅 Bookings by Month")
                if 'Month' in all_bookings.columns:
                    month_counts = all_bookings['Month'].value_counts().sort_index()
                    st.bar_chart(month_counts)
                
                st.divider()
                
                # Bookings by property
                st.subheader("🏠 Bookings by Property")
                if len(all_bookings.columns) > 1:
                    property_col = 'Property' if 'Property' in all_bookings.columns else all_bookings.columns[1]
                    property_counts = all_bookings[property_col].value_counts()
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.bar_chart(property_counts)
                    
                    with col2:
                        for prop, count in property_counts.items():
                            percentage = (count / len(all_bookings)) * 100
                            st.markdown(f"""
                            <div class="info-card">
                                <strong>{prop}</strong><br>
                                <span style="font-size: 1.5rem; font-weight: bold;">{count}</span> bookings
                                <span style="color: #666; margin-left: 0.5rem;">({percentage:.1f}%)</span>
                            </div>
                            """, unsafe_allow_html=True)
                
                st.divider()
                
                # Status breakdown
                st.subheader("📋 Status Code Breakdown")
                if len(all_bookings.columns) > 2:
                    status_col = 'Status' if 'Status' in all_bookings.columns else all_bookings.columns[2]
                    
                    for status_code, info in STATUS_CODES.items():
                        count = len(all_bookings[all_bookings[status_col] == status_code])
                        if count > 0:
                            percentage = (count / len(all_bookings)) * 100
                            st.markdown(f"""
                            <div class="info-card" style="border-left-color: {info['color']};">
                                <span class="status-badge" style="background-color: {info['color']};">
                                    {status_code}
                                </span>
                                <strong style="margin-left: 0.5rem;">{info['name']}</strong>
                                <span style="float: right; font-size: 1.2rem; font-weight: bold;">
                                    {count} <span style="font-size: 0.9rem; color: #666;">({percentage:.1f}%)</span>
                                </span>
                            </div>
                            """, unsafe_allow_html=True)
            else:
                st.info("No booking data available for analytics.")
        
        # System Logs
        elif view == "System Logs":
            st.header("📋 System Logs")
            
            if st.session_state.logs:
                # Filter options
                col1, col2, col3 = st.columns([2, 1, 1])
                
                with col2:
                    log_levels = ['All'] + list(set(log['level'] for log in st.session_state.logs))
                    filter_level = st.selectbox("Filter by Level", log_levels)
                
                with col3:
                    if st.button("🗑️ Clear Logs", use_container_width=True):
                        st.session_state.logs = []
                        st.rerun()
                
                # Display logs
                filtered_logs = st.session_state.logs if filter_level == 'All' else [
                    log for log in st.session_state.logs if log['level'] == filter_level
                ]
                
                for log in reversed(filtered_logs):
                    level_colors = {
                        'INFO': '#2196F3',
                        'SUCCESS': '#4CAF50',
                        'ERROR': '#F44336',
                        'WARNING': '#FF9800'
                    }
                    color = level_colors.get(log['level'], '#999999')
                    
                    st.markdown(f"""
                    <div class="info-card" style="border-left-color: {color};">
                        <span style="color: {color}; font-weight: bold;">[{log['level']}]</span>
                        <span style="color: #666; margin-left: 0.5rem;">{log['timestamp']}</span><br>
                        <span style="margin-top: 0.25rem; display: block;">{log['message']}</span>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No logs available yet.")
    
    except Exception as e:
        st.error(f"❌ Error loading workbook: {str(e)}")
        add_log(f"Error loading workbook: {str(e)}", "ERROR")
