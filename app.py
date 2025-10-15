import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import pandas as pd
import json
from typing import List, Dict, Optional, Tuple
import plotly.express as px
import plotly.graph_objects as go

# Page configuration
st.set_page_config(
    page_title="Booking Management System",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .status-badge {
        padding: 0.25rem 0.75rem;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
    }
    .sidebar-card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #1f77b4;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# <CHANGE> Added Google Drive API scope and folder configuration
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'  # Added Drive API scope
]

# Status code definitions
STATUS_CODES = {
    'CI': {'name': 'Check In', 'color': '#28a745', 'description': 'Guest checking in'},
    'SO': {'name': 'Stay Over', 'color': '#17a2b8', 'description': 'Guest staying, no cleaning'},
    'CO/CI': {'name': 'Check Out/Check In', 'color': '#fd7e14', 'description': 'Same day turnover'},
    'FU': {'name': 'Follow Up', 'color': '#ffc107', 'description': 'Follow up required'},
    'DC': {'name': 'Deep Clean', 'color': '#6f42c1', 'description': 'Deep cleaning service'},
    'COC': {'name': 'Check Out Clean', 'color': '#e83e8c', 'description': 'Check out cleaning'},
    'CO': {'name': 'Check Out', 'color': '#dc3545', 'description': 'Guest checking out'}
}

class BookingManager:
    def __init__(self):
        self.client = None
        self.current_workbook = None
        self.workbooks = []
        # <CHANGE> Added folder_id attribute
        self.folder_id = None
        
    def authenticate(self, credentials_dict: dict, folder_id: str = None) -> bool:
        """Authenticate with Google Sheets and Drive API"""
        try:
            creds = Credentials.from_service_account_info(
                credentials_dict,
                scopes=SCOPES
            )
            self.client = gspread.authorize(creds)
            self.folder_id = folder_id  # <CHANGE> Store folder ID
            self.log_activity("Authentication successful")
            return True
        except Exception as e:
            st.error(f"Authentication failed: {str(e)}")
            self.log_activity(f"Authentication failed: {str(e)}", "ERROR")
            return False
    
    # <CHANGE> Completely rewrote list_workbooks to pull from Google Drive folder
    def list_workbooks(self) -> List[Dict]:
        """List all spreadsheets from Google Drive folder"""
        if not self.client:
            return []
        
        try:
            if self.folder_id:
                # Query spreadsheets from specific folder
                query = f"'{self.folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
                files = self.client.list_spreadsheet_files(query=query)
            else:
                # List all accessible spreadsheets
                files = self.client.list_spreadsheet_files()
            
            self.workbooks = []
            for file in files:
                try:
                    workbook = self.client.open_by_key(file['id'])
                    self.workbooks.append({
                        'id': file['id'],
                        'name': file['name'],
                        'workbook': workbook,
                        'url': f"https://docs.google.com/spreadsheets/d/{file['id']}"
                    })
                except Exception as e:
                    self.log_activity(f"Error loading workbook {file['name']}: {str(e)}", "WARNING")
                    continue
            
            self.log_activity(f"Loaded {len(self.workbooks)} workbooks from {'folder' if self.folder_id else 'Drive'}")
            return self.workbooks
            
        except Exception as e:
            st.error(f"Error listing workbooks: {str(e)}")
            self.log_activity(f"Error listing workbooks: {str(e)}", "ERROR")
            return []
    
    def load_workbook(self, workbook_id: str) -> bool:
        """Load a specific workbook"""
        try:
            workbook_data = next((wb for wb in self.workbooks if wb['id'] == workbook_id), None)
            if workbook_data:
                self.current_workbook = workbook_data['workbook']
                self.log_activity(f"Loaded workbook: {workbook_data['name']}")
                return True
            return False
        except Exception as e:
            st.error(f"Error loading workbook: {str(e)}")
            self.log_activity(f"Error loading workbook: {str(e)}", "ERROR")
            return False
    
    def get_client_profile(self) -> Dict:
        """Extract client profile from first sheet"""
        if not self.current_workbook:
            return {}
        
        try:
            profile_sheet = self.current_workbook.get_worksheet(0)
            all_values = profile_sheet.get_all_values()
            
            profile = {
                'client_name': '',
                'properties': [],
                'contact_info': {},
                'notes': ''
            }
            
            # Parse client profile data
            for i, row in enumerate(all_values):
                if not row or not any(row):
                    continue
                
                # Look for client name
                if 'client' in str(row[0]).lower() and len(row) > 1:
                    profile['client_name'] = row[1]
                
                # Look for property information
                if 'property' in str(row[0]).lower() and len(row) > 1:
                    if row[1]:
                        profile['properties'].append(row[1])
                
                # Look for contact info
                if any(term in str(row[0]).lower() for term in ['email', 'phone', 'address']):
                    if len(row) > 1 and row[1]:
                        profile['contact_info'][row[0]] = row[1]
            
            return profile
            
        except Exception as e:
            self.log_activity(f"Error reading client profile: {str(e)}", "ERROR")
            return {}
    
    def get_calendar_sheets(self) -> List[Dict]:
        """Get all calendar sheets (excluding first profile sheet)"""
        if not self.current_workbook:
            return []
        
        try:
            worksheets = self.current_workbook.worksheets()
            calendars = []
            
            for sheet in worksheets[1:]:  # Skip first sheet (profile)
                calendars.append({
                    'name': sheet.title,
                    'sheet': sheet,
                    'id': sheet.id
                })
            
            return calendars
            
        except Exception as e:
            self.log_activity(f"Error getting calendar sheets: {str(e)}", "ERROR")
            return []
    
    def get_bookings_from_sheet(self, sheet, property_filter: str = None, status_filter: str = None) -> pd.DataFrame:
        """Extract bookings from a calendar sheet starting at row 13"""
        try:
            all_values = sheet.get_all_values()
            
            if len(all_values) < 13:
                return pd.DataFrame()
            
            # Get headers from row 12 (index 11)
            headers = all_values[11] if len(all_values) > 11 else []
            
            # Get data starting from row 13 (index 12)
            data_rows = all_values[12:]
            
            if not headers or not data_rows:
                return pd.DataFrame()
            
            # Create DataFrame
            df = pd.DataFrame(data_rows, columns=headers)
            
            # Clean up empty rows
            df = df[df.apply(lambda row: any(row.astype(str).str.strip() != ''), axis=1)]
            
            # Apply filters
            if property_filter and 'Property' in df.columns:
                df = df[df['Property'].str.contains(property_filter, case=False, na=False)]
            
            if status_filter and 'Status' in df.columns:
                df = df[df['Status'] == status_filter]
            
            # Add sheet name as source
            df['Calendar'] = sheet.title
            
            return df
            
        except Exception as e:
            self.log_activity(f"Error reading bookings from {sheet.title}: {str(e)}", "ERROR")
            return pd.DataFrame()
    
    def get_all_bookings(self, property_filter: str = None, status_filter: str = None) -> pd.DataFrame:
        """Get all bookings from all calendar sheets"""
        all_bookings = []
        
        for calendar in self.get_calendar_sheets():
            bookings = self.get_bookings_from_sheet(
                calendar['sheet'],
                property_filter,
                status_filter
            )
            if not bookings.empty:
                all_bookings.append(bookings)
        
        if all_bookings:
            return pd.concat(all_bookings, ignore_index=True)
        return pd.DataFrame()
    
    def add_booking(self, sheet, booking_data: Dict) -> bool:
        """Add a new booking to a calendar sheet"""
        try:
            all_values = sheet.get_all_values()
            
            if len(all_values) < 13:
                st.error("Sheet structure is invalid (needs at least 13 rows)")
                return False
            
            headers = all_values[11]
            
            # Find next empty row
            next_row = len(all_values) + 1
            
            # Prepare row data matching headers
            row_data = []
            for header in headers:
                row_data.append(booking_data.get(header, ''))
            
            # Append the row
            sheet.append_row(row_data)
            
            self.log_activity(f"Added booking to {sheet.title}")
            return True
            
        except Exception as e:
            st.error(f"Error adding booking: {str(e)}")
            self.log_activity(f"Error adding booking: {str(e)}", "ERROR")
            return False
    
    def log_activity(self, message: str, level: str = "INFO"):
        """Log activity to session state"""
        if 'activity_log' not in st.session_state:
            st.session_state.activity_log = []
        
        log_entry = {
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'level': level,
            'message': message
        }
        st.session_state.activity_log.insert(0, log_entry)
        
        # Keep only last 100 entries
        st.session_state.activity_log = st.session_state.activity_log[:100]

def render_status_badge(status_code: str) -> str:
    """Render a colored status badge"""
    if status_code in STATUS_CODES:
        info = STATUS_CODES[status_code]
        return f'<span class="status-badge" style="background-color: {info["color"]}; color: white;">{status_code} - {info["name"]}</span>'
    return f'<span class="status-badge" style="background-color: #6c757d; color: white;">{status_code}</span>'

def main():
    st.markdown('<h1 class="main-header">📅 Booking Management System</h1>', unsafe_allow_html=True)
    
    # Initialize session state
    if 'manager' not in st.session_state:
        st.session_state.manager = BookingManager()
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'current_view' not in st.session_state:
        st.session_state.current_view = 'Dashboard'
    
    manager = st.session_state.manager
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 🔐 Authentication")
        
        # <CHANGE> Added folder ID input field
        st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
        st.markdown("**Google Drive Folder**")
        folder_id_input = st.text_input(
            "Folder ID (optional)",
            value=st.session_state.get('folder_id', ''),
            help="Enter the Google Drive folder ID to filter spreadsheets. Leave empty to show all accessible sheets.",
            placeholder="e.g., 1a2b3c4d5e6f7g8h9i0j"
        )
        
        if folder_id_input:
            st.session_state.folder_id = folder_id_input
            st.info("📁 Will load sheets from specified folder")
        else:
            st.session_state.folder_id = None
            st.info("📂 Will load all accessible sheets")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Credentials input
        credentials_input = st.text_area(
            "Service Account JSON",
            height=150,
            help="Paste your Google Service Account credentials JSON here"
        )
        
        if st.button("🔓 Authenticate", use_container_width=True):
            if credentials_input:
                try:
                    creds_dict = json.loads(credentials_input)
                    # <CHANGE> Pass folder_id to authenticate
                    if manager.authenticate(creds_dict, st.session_state.get('folder_id')):
                        st.session_state.authenticated = True
                        st.success("✅ Authenticated successfully!")
                        st.rerun()
                except json.JSONDecodeError:
                    st.error("Invalid JSON format")
            else:
                st.warning("Please enter credentials")
        
        if st.session_state.authenticated:
            st.success("✅ Authenticated")
            
            # <CHANGE> Added refresh button for workbooks
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Refresh", use_container_width=True):
                    with st.spinner("Loading workbooks..."):
                        manager.list_workbooks()
                    st.rerun()
            with col2:
                if st.button("🚪 Logout", use_container_width=True):
                    st.session_state.authenticated = False
                    st.session_state.manager = BookingManager()
                    st.rerun()
            
            st.markdown("---")
            
            # Workbook selection
            if not manager.workbooks:
                with st.spinner("Loading workbooks from Drive..."):
                    manager.list_workbooks()
            
            if manager.workbooks:
                st.markdown("### 📚 Select Workbook")
                
                # <CHANGE> Enhanced workbook display with count
                st.info(f"Found {len(manager.workbooks)} workbook(s)")
                
                workbook_names = [wb['name'] for wb in manager.workbooks]
                selected_workbook_name = st.selectbox(
                    "Client Workbook",
                    workbook_names,
                    key="workbook_selector"
                )
                
                if selected_workbook_name:
                    selected_wb = next(wb for wb in manager.workbooks if wb['name'] == selected_workbook_name)
                    
                    if st.session_state.get('current_workbook_id') != selected_wb['id']:
                        manager.load_workbook(selected_wb['id'])
                        st.session_state.current_workbook_id = selected_wb['id']
                        st.rerun()
                    
                    # Display workbook info
                    st.markdown('<div class="sidebar-card">', unsafe_allow_html=True)
                    st.markdown(f"**Current Workbook:**")
                    st.markdown(f"📄 {selected_wb['name']}")
                    st.markdown(f"[Open in Google Sheets]({selected_wb['url']})")
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.warning("No workbooks found. Check your folder ID or permissions.")
            
            st.markdown("---")
            
            # Navigation
            st.markdown("### 🧭 Navigation")
            views = ['Dashboard', 'Calendar View', 'Booking Manager', 'Analytics', 'System Logs']
            st.session_state.current_view = st.radio("Go to", views, key="nav_radio")
            
            # Status code legend
            st.markdown("---")
            st.markdown("### 📋 Status Codes")
            for code, info in STATUS_CODES.items():
                st.markdown(
                    f'<div style="margin: 0.5rem 0;">{render_status_badge(code)}</div>',
                    unsafe_allow_html=True
                )
    
    # Main content area
    if not st.session_state.authenticated:
        st.info("👈 Please authenticate using the sidebar to get started")
        
        # <CHANGE> Added instructions for folder ID
        st.markdown("### 📖 Getting Started")
        st.markdown("""
        1. **Get your Google Drive Folder ID** (optional):
           - Open your Google Drive folder in a browser
           - Copy the folder ID from the URL: `https://drive.google.com/drive/folders/YOUR_FOLDER_ID`
           - Paste it in the sidebar
           
        2. **Create a Service Account**:
           - Go to [Google Cloud Console](https://console.cloud.google.com/)
           - Create a new project or select existing
           - Enable Google Sheets API and Google Drive API
           - Create a Service Account and download JSON key
           
        3. **Share your folder/sheets**:
           - Share your Google Drive folder or individual sheets with the service account email
           - Grant "Viewer" or "Editor" permissions
           
        4. **Paste credentials** in the sidebar and click Authenticate
        """)
        
        st.markdown("### 📊 Expected Sheet Structure")
        st.markdown("""
        Each workbook should have:
        - **First sheet**: Client profile with property details
        - **Subsequent sheets**: Monthly booking calendars
        - **Booking data**: Starts at row 13
        - **Headers**: Located at row 12
        """)
        
        return
    
    if not manager.current_workbook:
        st.warning("Please select a workbook from the sidebar")
        return
    
    # Render selected view
    if st.session_state.current_view == 'Dashboard':
        render_dashboard(manager)
    elif st.session_state.current_view == 'Calendar View':
        render_calendar_view(manager)
    elif st.session_state.current_view == 'Booking Manager':
        render_booking_manager(manager)
    elif st.session_state.current_view == 'Analytics':
        render_analytics(manager)
    elif st.session_state.current_view == 'System Logs':
        render_system_logs()

def render_dashboard(manager: BookingManager):
    """Render the dashboard view"""
    st.markdown("## 📊 Dashboard")
    
    # Get client profile
    profile = manager.get_client_profile()
    
    # Display client info
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 👤 Client Profile")
        if profile.get('client_name'):
            st.markdown(f"**Client:** {profile['client_name']}")
        
        if profile.get('properties'):
            st.markdown("**Properties:**")
            for prop in profile['properties']:
                st.markdown(f"- 🏠 {prop}")
        
        if profile.get('contact_info'):
            st.markdown("**Contact Information:**")
            for key, value in profile['contact_info'].items():
                st.markdown(f"- **{key}:** {value}")
    
    with col2:
        calendars = manager.get_calendar_sheets()
        st.metric("📅 Calendar Sheets", len(calendars))
        st.metric("🏠 Properties", len(profile.get('properties', [])))
    
    st.markdown("---")
    
    # Get all bookings
    all_bookings = manager.get_all_bookings()
    
    if not all_bookings.empty:
        # Metrics
        st.markdown("### 📈 Booking Metrics")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Bookings", len(all_bookings))
        
        with col2:
            if 'Status' in all_bookings.columns:
                unique_statuses = all_bookings['Status'].nunique()
                st.metric("Status Types", unique_statuses)
        
        with col3:
            if 'Property' in all_bookings.columns:
                unique_properties = all_bookings['Property'].nunique()
                st.metric("Active Properties", unique_properties)
        
        with col4:
            if 'Calendar' in all_bookings.columns:
                unique_calendars = all_bookings['Calendar'].nunique()
                st.metric("Calendar Months", unique_calendars)
        
        # Status distribution
        if 'Status' in all_bookings.columns:
            st.markdown("### 📊 Status Distribution")
            status_counts = all_bookings['Status'].value_counts()
            
            fig = go.Figure(data=[go.Pie(
                labels=status_counts.index,
                values=status_counts.values,
                hole=0.4,
                marker=dict(colors=[STATUS_CODES.get(s, {}).get('color', '#6c757d') for s in status_counts.index])
            )])
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        
        # Recent bookings
        st.markdown("### 🕐 Recent Bookings")
        display_df = all_bookings.head(10)
        
        if 'Status' in display_df.columns:
            display_df['Status'] = display_df['Status'].apply(
                lambda x: render_status_badge(x) if x in STATUS_CODES else x
            )
        
        st.markdown(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)
    else:
        st.info("No bookings found in the current workbook")

def render_calendar_view(manager: BookingManager):
    """Render the calendar view"""
    st.markdown("## 📅 Calendar View")
    
    calendars = manager.get_calendar_sheets()
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    # Filters
    col1, col2, col3 = st.columns(3)
    
    with col1:
        selected_calendar = st.selectbox(
            "Select Calendar",
            ['All Calendars'] + [cal['name'] for cal in calendars]
        )
    
    with col2:
        profile = manager.get_client_profile()
        properties = profile.get('properties', [])
        property_filter = st.selectbox(
            "Filter by Property",
            ['All Properties'] + properties
        )
    
    with col3:
        status_filter = st.selectbox(
            "Filter by Status",
            ['All Statuses'] + list(STATUS_CODES.keys())
        )
    
    # Apply filters
    prop_filter = None if property_filter == 'All Properties' else property_filter
    stat_filter = None if status_filter == 'All Statuses' else status_filter
    
    if selected_calendar == 'All Calendars':
        bookings = manager.get_all_bookings(prop_filter, stat_filter)
    else:
        calendar = next(cal for cal in calendars if cal['name'] == selected_calendar)
        bookings = manager.get_bookings_from_sheet(calendar['sheet'], prop_filter, stat_filter)
    
    if not bookings.empty:
        st.markdown(f"### 📋 Showing {len(bookings)} booking(s)")
        
        # Display bookings
        if 'Status' in bookings.columns:
            bookings_display = bookings.copy()
            bookings_display['Status'] = bookings_display['Status'].apply(
                lambda x: render_status_badge(x) if x in STATUS_CODES else x
            )
            st.markdown(bookings_display.to_html(escape=False, index=False), unsafe_allow_html=True)
        else:
            st.dataframe(bookings, use_container_width=True)
        
        # Export option
        csv = bookings.to_csv(index=False)
        st.download_button(
            label="📥 Download as CSV",
            data=csv,
            file_name=f"bookings_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("No bookings match the selected filters")

def render_booking_manager(manager: BookingManager):
    """Render the booking manager"""
    st.markdown("## ➕ Booking Manager")
    
    calendars = manager.get_calendar_sheets()
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    st.markdown("### Create New Booking")
    
    with st.form("new_booking_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            calendar_select = st.selectbox(
                "Select Calendar *",
                [cal['name'] for cal in calendars]
            )
            
            profile = manager.get_client_profile()
            properties = profile.get('properties', [''])
            property_select = st.selectbox("Property *", properties)
            
            status_select = st.selectbox(
                "Status *",
                list(STATUS_CODES.keys()),
                format_func=lambda x: f"{x} - {STATUS_CODES[x]['name']}"
            )
        
        with col2:
            date_input = st.date_input("Date *")
            guest_name = st.text_input("Guest Name")
            notes = st.text_area("Notes")
        
        submitted = st.form_submit_button("➕ Add Booking", use_container_width=True)
        
        if submitted:
            calendar = next(cal for cal in calendars if cal['name'] == calendar_select)
            
            booking_data = {
                'Date': date_input.strftime('%Y-%m-%d'),
                'Property': property_select,
                'Status': status_select,
                'Guest': guest_name,
                'Notes': notes
            }
            
            if manager.add_booking(calendar['sheet'], booking_data):
                st.success("✅ Booking added successfully!")
                st.balloons()
            else:
                st.error("❌ Failed to add booking")

def render_analytics(manager: BookingManager):
    """Render analytics view"""
    st.markdown("## 📊 Analytics & Reports")
    
    all_bookings = manager.get_all_bookings()
    
    if all_bookings.empty:
        st.info("No booking data available for analysis")
        return
    
    # Status analysis
    if 'Status' in all_bookings.columns:
        st.markdown("### 📈 Status Analysis")
        
        status_counts = all_bookings['Status'].value_counts()
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig = px.bar(
                x=status_counts.index,
                y=status_counts.values,
                labels={'x': 'Status', 'y': 'Count'},
                title='Bookings by Status',
                color=status_counts.index,
                color_discrete_map={s: STATUS_CODES.get(s, {}).get('color', '#6c757d') for s in status_counts.index}
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = go.Figure(data=[go.Pie(
                labels=status_counts.index,
                values=status_counts.values,
                hole=0.4
            )])
            fig.update_layout(title='Status Distribution')
            st.plotly_chart(fig, use_container_width=True)
    
    # Property analysis
    if 'Property' in all_bookings.columns:
        st.markdown("### 🏠 Property Analysis")
        
        property_counts = all_bookings['Property'].value_counts()
        
        fig = px.bar(
            x=property_counts.index,
            y=property_counts.values,
            labels={'x': 'Property', 'y': 'Bookings'},
            title='Bookings by Property'
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Calendar analysis
    if 'Calendar' in all_bookings.columns:
        st.markdown("### 📅 Monthly Analysis")
        
        calendar_counts = all_bookings['Calendar'].value_counts()
        
        fig = px.line(
            x=calendar_counts.index,
            y=calendar_counts.values,
            labels={'x': 'Month', 'y': 'Bookings'},
            title='Bookings by Month',
            markers=True
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Export analytics
    st.markdown("### 📥 Export Data")
    csv = all_bookings.to_csv(index=False)
    st.download_button(
        label="Download Full Dataset",
        data=csv,
        file_name=f"booking_analytics_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

def render_system_logs():
    """Render system logs"""
    st.markdown("## 📋 System Logs")
    
    if 'activity_log' not in st.session_state or not st.session_state.activity_log:
        st.info("No activity logged yet")
        return
    
    # Filter options
    col1, col2 = st.columns([3, 1])
    
    with col1:
        search_term = st.text_input("🔍 Search logs", "")
    
    with col2:
        level_filter = st.selectbox("Level", ['All', 'INFO', 'WARNING', 'ERROR'])
    
    # Display logs
    logs = st.session_state.activity_log
    
    if level_filter != 'All':
        logs = [log for log in logs if log['level'] == level_filter]
    
    if search_term:
        logs = [log for log in logs if search_term.lower() in log['message'].lower()]
    
    st.markdown(f"### Showing {len(logs)} log entries")
    
    for log in logs:
        level_color = {
            'INFO': '#17a2b8',
            'WARNING': '#ffc107',
            'ERROR': '#dc3545'
        }.get(log['level'], '#6c757d')
        
        st.markdown(
            f"""
            <div style="padding: 0.75rem; margin: 0.5rem 0; border-left: 4px solid {level_color}; background: #f8f9fa; border-radius: 4px;">
                <strong style="color: {level_color};">[{log['level']}]</strong> 
                <span style="color: #6c757d;">{log['timestamp']}</span><br/>
                {log['message']}
            </div>
            """,
            unsafe_allow_html=True
        )
    
    # Clear logs button
    if st.button("🗑️ Clear All Logs"):
        st.session_state.activity_log = []
        st.rerun()

if __name__ == "__main__":
    main()
