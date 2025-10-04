import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, time, timedelta
import pandas as pd
import json
import plotly.express as px
import plotly.graph_objects as go
from typing import Optional, Dict, Any

# Page configuration
st.set_page_config(
    page_title="Villa Booking Management System",
    page_icon="🏖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for enhanced UI
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');
    
    * {
        font-family: 'Poppins', sans-serif;
    }
    
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
        font-size: 16px;
    }
    
    .stButton>button:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
    }
    
    .stButton>button:active {
        transform: translateY(-1px);
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
    
    .booking-header p {
        font-size: 1.2rem;
        opacity: 0.95;
        font-weight: 300;
    }
    
    .status-new { 
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white; 
        padding: 8px 16px; 
        border-radius: 20px; 
        font-weight: 600;
        display: inline-block;
        box-shadow: 0 4px 15px rgba(17, 153, 142, 0.3);
    }
    
    .status-reserved { 
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        color: white; 
        padding: 8px 16px; 
        border-radius: 20px; 
        font-weight: 600;
        display: inline-block;
        box-shadow: 0 4px 15px rgba(79, 172, 254, 0.3);
    }
    
    .status-canceled { 
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        color: white; 
        padding: 8px 16px; 
        border-radius: 20px; 
        font-weight: 600;
        display: inline-block;
        box-shadow: 0 4px 15px rgba(245, 87, 108, 0.3);
    }
    
    .status-update { 
        background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
        color: white; 
        padding: 8px 16px; 
        border-radius: 20px; 
        font-weight: 600;
        display: inline-block;
        box-shadow: 0 4px 15px rgba(250, 112, 154, 0.3);
    }
    
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 30px rgba(0,0,0,0.12);
    }
    
    .booking-card {
        background: white;
        padding: 1.5rem;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        margin-bottom: 1rem;
        border-left: 5px solid #667eea;
        transition: all 0.3s ease;
    }
    
    .booking-card:hover {
        transform: translateX(5px);
        box-shadow: 0 8px 30px rgba(0,0,0,0.12);
    }
    
    .login-container {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
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
        from {
            transform: translateY(-20px);
            opacity: 0;
        }
        to {
            transform: translateY(0);
            opacity: 1;
        }
    }
    
    .stTextInput>div>div>input, .stTextArea>div>div>textarea, .stSelectbox>div>div>select {
        border-radius: 10px;
        border: 2px solid #e0e0e0;
        padding: 0.75rem;
        transition: all 0.3s ease;
    }
    
    .stTextInput>div>div>input:focus, .stTextArea>div>div>textarea:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    .upload-section {
        background: #f8f9fa;
        padding: 2rem;
        border-radius: 15px;
        border: 2px dashed #667eea;
        text-align: center;
        margin: 1rem 0;
    }
    
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin: 2rem 0;
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

# Google Sheets Configuration
SPREADSHEET_ID = "1-3FLLEkUmiHzW7DGVAPI6PebdRc_24t3vM0OCBnDhco"
SHEET_NAME = "Copy of Test Template Sheet Reservations"

def authenticate_google_sheets(creds_dict: Dict[str, Any]) -> Optional[gspread.Client]:
    """Authenticate with Google Sheets using service account credentials"""
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ]
        
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(creds)
        
        # Test the connection
        _ = client.open_by_key(SPREADSHEET_ID)
        
        return client
    except Exception as e:
        st.error(f"❌ Authentication failed: {str(e)}")
        return None

def get_worksheet(client: gspread.Client) -> Optional[gspread.Worksheet]:
    """Get the specific worksheet"""
    try:
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet(SHEET_NAME)
        return worksheet
    except Exception as e:
        st.error(f"❌ Error accessing worksheet: {str(e)}")
        return None

def read_bookings(worksheet: gspread.Worksheet) -> pd.DataFrame:
    """Read all bookings from the sheet with enhanced error handling"""
    try:
        all_values = worksheet.get_all_values()
        
        if len(all_values) < 7:
            return pd.DataFrame()
        
        headers = all_values[5]
        data = all_values[6:]
        
        df = pd.DataFrame(data, columns=headers)
        df = df[df['DATE:'].notna() & (df['DATE:'] != '')]
        
        return df
    except Exception as e:
        st.error(f"❌ Error reading bookings: {str(e)}")
        return pd.DataFrame()

def add_booking(worksheet: gspread.Worksheet, booking_data: Dict[str, Any]) -> bool:
    """Add a new booking to the sheet"""
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
        
        worksheet.append_row(row_data, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"❌ Error adding booking: {str(e)}")
        return False

def update_client_info(worksheet: gspread.Worksheet, client_name: str, service_type: str) -> bool:
    """Update client name and service type"""
    try:
        worksheet.update('B2', client_name)
        worksheet.update('B3', service_type)
        return True
    except Exception as e:
        st.error(f"❌ Error updating client info: {str(e)}")
        return False

def delete_booking(worksheet: gspread.Worksheet, row_index: int) -> bool:
    """Delete a booking from the sheet"""
    try:
        worksheet.delete_rows(row_index + 7)  # +7 because data starts at row 7
        return True
    except Exception as e:
        st.error(f"❌ Error deleting booking: {str(e)}")
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

def create_status_chart(stats: Dict[str, int]) -> go.Figure:
    """Create a pie chart for booking status"""
    labels = ['New', 'Reserved', 'Canceled', 'Update']
    values = [stats['new'], stats['reserved'], stats['canceled'], stats['update']]
    colors = ['#38ef7d', '#00f2fe', '#f5576c', '#fee140']
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        marker=dict(colors=colors),
        hole=0.4,
        textinfo='label+percent',
        textfont=dict(size=14, color='white'),
        hoverinfo='label+value'
    )])
    
    fig.update_layout(
        title='Booking Status Distribution',
        height=400,
        showlegend=True,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    
    return fig

# Sidebar Login Section
def render_login_sidebar():
    """Render the login section in sidebar"""
    st.sidebar.markdown("## 🔐 Authentication")
    
    if not st.session_state.authenticated:
        st.sidebar.markdown("""
        <div class="login-container">
            <h3 style="text-align: center; color: #667eea;">📁 Upload Credentials</h3>
            <p style="text-align: center; font-size: 14px; color: #666;">
                Upload your Google Service Account JSON file to connect
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.sidebar.file_uploader(
            "Choose JSON credentials file",
            type=['json'],
            help="Upload your Google Service Account credentials JSON file"
        )
        
        if uploaded_file is not None:
            try:
                creds_dict = json.load(uploaded_file)
                
                # Validate required fields
                required_fields = ['type', 'project_id', 'private_key', 'client_email']
                missing_fields = [field for field in required_fields if field not in creds_dict]
                
                if missing_fields:
                    st.sidebar.error(f"❌ Missing required fields: {', '.join(missing_fields)}")
                    return
                
                with st.spinner("🔄 Authenticating..."):
                    client = authenticate_google_sheets(creds_dict)
                    
                    if client:
                        worksheet = get_worksheet(client)
                        
                        if worksheet:
                            st.session_state.authenticated = True
                            st.session_state.client = client
                            st.session_state.worksheet = worksheet
                            st.session_state.credentials = creds_dict
                            
                            st.sidebar.success("✅ Authentication successful!")
                            st.balloons()
                            st.rerun()
            
            except json.JSONDecodeError:
                st.sidebar.error("❌ Invalid JSON file. Please upload a valid credentials file.")
            except Exception as e:
                st.sidebar.error(f"❌ Error: {str(e)}")
        
        # Instructions
        with st.sidebar.expander("📖 Setup Instructions"):
            st.markdown("""
            ### How to get your credentials:
            
            1. **Go to Google Cloud Console**
               - Visit [console.cloud.google.com](https://console.cloud.google.com)
            
            2. **Create/Select Project**
               - Create a new project or select existing
            
            3. **Enable APIs**
               - Enable Google Sheets API
               - Enable Google Drive API
            
            4. **Create Service Account**
               - Go to IAM & Admin → Service Accounts
               - Click "Create Service Account"
               - Give it a name and create
            
            5. **Generate Key**
               - Click on service account
               - Go to Keys tab
               - Add Key → Create new key
               - Choose JSON format
               - Download the file
            
            6. **Share Your Sheet**
               - Open your Google Sheet
               - Click Share
               - Add service account email
               - Grant Editor access
            
            7. **Upload JSON**
               - Upload the downloaded JSON file above
            """)
    
    else:
        st.sidebar.success("✅ Connected to Google Sheets")
        
        # Display connection info
        if st.session_state.credentials:
            creds = st.session_state.credentials
            st.sidebar.markdown(f"""
            **📧 Service Account:**  
            `{creds.get('client_email', 'N/A')[:30]}...`
            
            **🆔 Project ID:**  
            `{creds.get('project_id', 'N/A')}`
            """)
        
        if st.sidebar.button("🚪 Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.client = None
            st.session_state.worksheet = None
            st.session_state.credentials = None
            st.rerun()

# Main App
def main():
    # Render login sidebar
    render_login_sidebar()
    
    # Check authentication
    if not st.session_state.authenticated:
        st.markdown("""
            <div class="booking-header">
                <h1>🏖️ Villa Booking Management System</h1>
                <p>Professional vacation rental booking platform with real-time Google Sheets integration</p>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            st.info("""
            ### 👋 Welcome!
            
            Please upload your Google Service Account credentials in the sidebar to get started.
            
            **Features:**
            - 📝 Create and manage bookings
            - 📊 Real-time analytics
            - 🔍 Advanced filtering
            - 📈 Visual reports
            - 🔄 Live sync with Google Sheets
            """)
        
        return
    
    worksheet = st.session_state.worksheet
    
    # Header
    st.markdown("""
        <div class="booking-header">
            <h1>🏖️ Villa Booking Management System</h1>
            <p>Manage your vacation rental bookings with real-time synchronization</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Navigation
    st.sidebar.markdown("---")
    st.sidebar.title("📋 Navigation")
    page = st.sidebar.radio(
        "Select Page",
        ["📊 Dashboard", "📝 New Booking", "📅 View Bookings", "📈 Analytics", "⚙️ Settings"],
        label_visibility="collapsed"
    )
    
    # DASHBOARD PAGE
    if page == "📊 Dashboard":
        st.header("📊 Booking Dashboard")
        
        # Load bookings
        with st.spinner("Loading data..."):
            df = read_bookings(worksheet)
        
        if not df.empty:
            stats = get_booking_statistics(df)
            
            # Display metrics
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.markdown("""
                <div class="metric-card">
                    <h3 style="color: #667eea; margin: 0;">📊 Total</h3>
                    <h1 style="margin: 10px 0;">{}</h1>
                </div>
                """.format(stats['total']), unsafe_allow_html=True)
            
            with col2:
                st.markdown("""
                <div class="metric-card">
                    <h3 style="color: #38ef7d; margin: 0;">🆕 New</h3>
                    <h1 style="margin: 10px 0;">{}</h1>
                </div>
                """.format(stats['new']), unsafe_allow_html=True)
            
            with col3:
                st.markdown("""
                <div class="metric-card">
                    <h3 style="color: #00f2fe; margin: 0;">✅ Reserved</h3>
                    <h1 style="margin: 10px 0;">{}</h1>
                </div>
                """.format(stats['reserved']), unsafe_allow_html=True)
            
            with col4:
                st.markdown("""
                <div class="metric-card">
                    <h3 style="color: #f5576c; margin: 0;">❌ Canceled</h3>
                    <h1 style="margin: 10px 0;">{}</h1>
                </div>
                """.format(stats['canceled']), unsafe_allow_html=True)
            
            with col5:
                st.markdown("""
                <div class="metric-card">
                    <h3 style="color: #fee140; margin: 0;">🔄 Update</h3>
                    <h1 style="margin: 10px 0;">{}</h1>
                </div>
                """.format(stats['update']), unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Charts
            col1, col2 = st.columns(2)
            
            with col1:
                fig = create_status_chart(stats)
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Recent bookings
                st.subheader("📅 Recent Bookings")
                recent_df = df.head(5)
                for idx, row in recent_df.iterrows():
                    st.markdown(f"""
                    <div class="booking-card">
                        <strong>🏠 {row['VILLA:']}</strong><br>
                        📅 {row['DATE:']} | ⏰ {row['START TIME:']} - {row['END TIME:']}<br>
                        <span class="status-{row['RESERVATION STATUS:'].lower().replace('!', '')}">{row['RESERVATION STATUS:']}</span>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("📭 No bookings found. Create your first booking!")
    
    # NEW BOOKING PAGE
    elif page == "📝 New Booking":
        st.header("📝 Create New Booking")
        
        with st.form("booking_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📍 Basic Information")
                date = st.date_input(
                    "📅 Booking Date",
                    datetime.now(),
                    help="Select the date for the booking"
                )
                
                villa = st.text_input(
                    "🏠 Villa Name",
                    placeholder="Enter villa name or number",
                    help="Specify the villa for this booking"
                )
                
                type_clean = st.selectbox(
                    "🧹 Type of Cleaning",
                    ["Standard Clean", "Deep Clean", "Check-out Clean", "Mid-stay Clean", "Move-in Clean", "Post-event Clean"],
                    help="Select the type of cleaning service required"
                )
                
                pax = st.number_input(
                    "👥 Number of Guests (PAX)",
                    min_value=1,
                    max_value=50,
                    value=2,
                    help="Enter the number of guests"
                )
            
            with col2:
                st.subheader("⏰ Schedule & Services")
                start_time = st.time_input(
                    "⏰ Start Time",
                    time(10, 0),
                    help="When should the service start?"
                )
                
                end_time = st.time_input(
                    "⏰ End Time",
                    time(14, 0),
                    help="Expected completion time"
                )
                
                reservation_status = st.selectbox(
                    "📊 Reservation Status",
                    ["NEW!", "RESERVED!", "CANCELED!", "UPDATE!"],
                    help="Current status of the booking"
                )
                
                laundry_services = st.selectbox(
                    "🧺 Laundry Services with Videmi",
                    ["Yes", "No", "Not Required", "Pending"],
                    help="Does this booking include laundry services?"
                )
            
            st.subheader("💬 Additional Information")
            comments = st.text_area(
                "Comments & Special Requirements",
                placeholder="Add any special notes, requirements, or instructions...",
                height=100,
                help="Include any special requests or important notes"
            )
            
            col1, col2, col3 = st.columns([1, 1, 1])
            with col2:
                submit_button = st.form_submit_button("✅ Submit Booking", use_container_width=True)
            
            if submit_button:
                if not villa:
                    st.error("❌ Please enter a villa name")
                elif start_time >= end_time:
                    st.error("❌ End time must be after start time")
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
                    
                    with st.spinner("💾 Saving booking..."):
                        if add_booking(worksheet, booking_data):
                            st.markdown("""
                            <div class="success-message">
                                ✅ Booking added successfully!
                            </div>
                            """, unsafe_allow_html=True)
                            st.balloons()
                        else:
                            st.markdown("""
                            <div class="error-message">
                                ❌ Failed to add booking. Please try again.
                            </div>
                            """, unsafe_allow_html=True)
    
    # VIEW BOOKINGS PAGE
    elif page == "📅 View Bookings":
        st.header("📅 All Bookings")
        
        # Advanced filters
        with st.expander("🔍 Advanced Filters", expanded=True):
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                filter_status = st.selectbox(
                    "Status",
                    ["All", "NEW!", "RESERVED!", "CANCELED!", "UPDATE!"]
                )
            
            with col2:
                filter_date_from = st.date_input("From Date", value=None)
            
            with col3:
                filter_date_to = st.date_input("To Date", value=None)
            
            with col4:
                search_villa = st.text_input("🔍 Search Villa")
        
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("🔄 Refresh Data", use_container_width=True):
                st.rerun()
        
        # Load and filter bookings
        with st.spinner("Loading bookings..."):
            df = read_bookings(worksheet)
        
        if not df.empty:
            filtered_df = df.copy()
            
            # Apply filters
            if filter_status != "All":
                filtered_df = filtered_df[filtered_df['RESERVATION STATUS:'] == filter_status]
            
            if filter_date_from:
                filtered_df = filtered_df[pd.to_datetime(filtered_df['DATE:']) >= pd.to_datetime(filter_date_from)]
            
            if filter_date_to:
                filtered_df = filtered_df[pd.to_datetime(filtered_df['DATE:']) <= pd.to_datetime(filter_date_to)]
            
            if search_villa:
                filtered_df = filtered_df[filtered_df['VILLA:'].str.contains(search_villa, case=False, na=False)]
            
            st.info(f"📈 Showing **{len(filtered_df)}** booking(s) out of **{len(df)}** total")
            
            # Display bookings
            for idx, row in filtered_df.iterrows():
                status = row['RESERVATION STATUS:']
                status_class = f"status-{status.lower().replace('!', '')}"
                
                with st.expander(f"🏠 {row['VILLA:']} - {row['DATE:']} ({status})"):
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown("**📅 Date & Time**")
                        st.write(f"Date: {row['DATE:']}")
                        st.write(f"Start: {row['START TIME:']}")
                        st.write(f"End: {row['END TIME:']}")
                    
                    with col2:
                        st.markdown("**🏠 Details**")
                        st.write(f"Villa: {row['VILLA:']}")
                        st.write(f"Cleaning: {row['TYPE CLEAN:']}")
                        st.write(f"PAX: {row['PAX:']}")
                    
                    with col3:
                        st.markdown("**📊 Status & Services**")
                        st.markdown(f"<span class='{status_class}'>{status}</span>", unsafe_allow_html=True)
                        st.write(f"Laundry: {row['LAUNDRY SERVICES WITH VIDEMI:']}")
                    
                    if row['COMMENTS:']:
                        st.markdown("**💬 Comments**")
                        st.info(row['COMMENTS:'])
        else:
            st.info("📭 No bookings found")
    
    # ANALYTICS PAGE
    elif page == "📈 Analytics":
        st.header("📈 Booking Analytics")
        
        with st.spinner("Analyzing data..."):
            df = read_bookings(worksheet)
        
        if not df.empty:
            stats = get_booking_statistics(df)
            
            # Status distribution
            col1, col2 = st.columns(2)
            
            with col1:
                fig_status = create_status_chart(stats)
                st.plotly_chart(fig_status, use_container_width=True)
            
            with col2:
                # Bookings by cleaning type
                cleaning_counts = df['TYPE CLEAN:'].value_counts()
                fig_cleaning = px.bar(
                    x=cleaning_counts.index,
                    y=cleaning_counts.values,
                    labels={'x': 'Cleaning Type', 'y': 'Count'},
                    title='Bookings by Cleaning Type',
                    color=cleaning_counts.values,
                    color_continuous_scale='Viridis'
                )
                fig_cleaning.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    showlegend=False
                )
                st.plotly_chart(fig_cleaning, use_container_width=True)
            
            # Timeline analysis
            st.subheader("📅 Booking Timeline")
            
            try:
                df['DATE_PARSED'] = pd.to_datetime(df['DATE:'], format='%m/%d/%Y', errors='coerce')
                df_sorted = df.sort_values('DATE_PARSED')
                
                timeline_data = df_sorted.groupby(['DATE_PARSED', 'RESERVATION STATUS:']).size().reset_index(name='count')
                
                fig_timeline = px.line(
                    timeline_data,
                    x='DATE_PARSED',
                    y='count',
                    color='RESERVATION STATUS:',
                    title='Booking Trends Over Time',
                    labels={'DATE_PARSED': 'Date', 'count': 'Number of Bookings'},
                    color_discrete_map={
                        'NEW!': '#38ef7d',
                        'RESERVED!': '#00f2fe',
                        'CANCELED!': '#f5576c',
                        'UPDATE!': '#fee140'
                    }
                )
                fig_timeline.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    hovermode='x unified'
                )
                st.plotly_chart(fig_timeline, use_container_width=True)
            except Exception as e:
                st.warning("Unable to generate timeline chart")
            
            # PAX distribution
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("👥 Guest Distribution (PAX)")
                pax_data = df['PAX:'].value_counts().sort_index()
                fig_pax = px.bar(
                    x=pax_data.index,
                    y=pax_data.values,
                    labels={'x': 'Number of Guests', 'y': 'Frequency'},
                    title='Guest Count Distribution',
                    color=pax_data.values,
                    color_continuous_scale='Blues'
                )
                fig_pax.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    showlegend=False
                )
                st.plotly_chart(fig_pax, use_container_width=True)
            
            with col2:
                st.subheader("🧺 Laundry Services")
                laundry_counts = df['LAUNDRY SERVICES WITH VIDEMI:'].value_counts()
                fig_laundry = go.Figure(data=[go.Pie(
                    labels=laundry_counts.index,
                    values=laundry_counts.values,
                    hole=0.4,
                    marker=dict(colors=['#667eea', '#764ba2', '#fa709a', '#fee140'])
                )])
                fig_laundry.update_layout(
                    title='Laundry Service Distribution',
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    height=400
                )
                st.plotly_chart(fig_laundry, use_container_width=True)
            
            # Top villas
            st.subheader("🏆 Most Booked Villas")
            villa_counts = df['VILLA:'].value_counts().head(10)
            
            fig_villas = px.bar(
                x=villa_counts.values,
                y=villa_counts.index,
                orientation='h',
                labels={'x': 'Number of Bookings', 'y': 'Villa'},
                title='Top 10 Villas by Booking Count',
                color=villa_counts.values,
                color_continuous_scale='Sunset'
            )
            fig_villas.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                showlegend=False,
                height=500
            )
            st.plotly_chart(fig_villas, use_container_width=True)
            
            # Download data option
            st.markdown("---")
            st.subheader("📥 Export Data")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="⬇️ Download as CSV",
                    data=csv,
                    file_name=f"bookings_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            with col2:
                json_str = df.to_json(orient='records', indent=2)
                st.download_button(
                    label="⬇️ Download as JSON",
                    data=json_str,
                    file_name=f"bookings_{datetime.now().strftime('%Y%m%d')}.json",
                    mime="application/json",
                    use_container_width=True
                )
            
            with col3:
                excel_buffer = pd.ExcelWriter('temp.xlsx', engine='openpyxl')
                df.to_excel(excel_buffer, index=False, sheet_name='Bookings')
                excel_buffer.close()
        
        else:
            st.info("📭 No data available for analytics")
    
    # SETTINGS PAGE
    elif page == "⚙️ Settings":
        st.header("⚙️ System Settings")
        
        # Client Information Section
        st.subheader("👤 Client Information")
        
        with st.form("client_settings_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                client_name = st.text_input(
                    "Client Name",
                    placeholder="Enter client or company name",
                    help="This will be displayed in the sheet"
                )
            
            with col2:
                service_type = st.selectbox(
                    "Type of Service",
                    [
                        "Vacation Rental",
                        "Property Management",
                        "Cleaning Service",
                        "Hospitality",
                        "Real Estate",
                        "Other"
                    ],
                    help="Select the primary service type"
                )
            
            col1, col2, col3 = st.columns([1, 1, 1])
            with col2:
                submit_settings = st.form_submit_button("💾 Save Settings", use_container_width=True)
            
            if submit_settings:
                if not client_name:
                    st.error("❌ Please enter a client name")
                else:
                    with st.spinner("Updating settings..."):
                        if update_client_info(worksheet, client_name, service_type):
                            st.markdown("""
                            <div class="success-message">
                                ✅ Settings updated successfully!
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.markdown("""
                            <div class="error-message">
                                ❌ Failed to update settings
                            </div>
                            """, unsafe_allow_html=True)
        
        # Display current settings
        st.markdown("---")
        st.subheader("📋 Current Configuration")
        
        try:
            current_client = worksheet.acell('B2').value
            current_service = worksheet.acell('B3').value
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #667eea;">👤 Current Client</h4>
                    <h2>{current_client or 'Not Set'}</h2>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #764ba2;">🏷️ Service Type</h4>
                    <h2>{current_service or 'Not Set'}</h2>
                </div>
                """, unsafe_allow_html=True)
        except Exception as e:
            st.warning("Unable to load current settings")
        
        # System Information
        st.markdown("---")
        st.subheader("ℹ️ System Information")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div class="metric-card">
                <h4 style="color: #667eea;">📊 Spreadsheet ID</h4>
                <code style="font-size: 12px; word-break: break-all;">{}</code>
            </div>
            """.format(SPREADSHEET_ID[:20] + "..."), unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="metric-card">
                <h4 style="color: #764ba2;">📄 Sheet Name</h4>
                <p style="font-size: 14px;">{}</p>
            </div>
            """.format(SHEET_NAME), unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div class="metric-card">
                <h4 style="color: #38ef7d;">🔗 Status</h4>
                <p style="font-size: 14px; color: #38ef7d;"><strong>● Connected</strong></p>
            </div>
            """, unsafe_allow_html=True)
        
        # Advanced Settings
        st.markdown("---")
        st.subheader("🔧 Advanced Settings")
        
        with st.expander("⚙️ Configuration Options"):
            st.info("""
            **Available Configuration Options:**
            
            - **Auto-refresh interval**: Set automatic data refresh
            - **Date format**: Customize date display format
            - **Time zone**: Configure timezone settings
            - **Notification settings**: Email alerts for new bookings
            - **Export preferences**: Default export format
            
            *These features are coming in future updates*
            """)
        
        # Danger Zone
        st.markdown("---")
        st.subheader("⚠️ Danger Zone")
        
        with st.expander("🚨 Advanced Operations", expanded=False):
            st.warning("""
            **Warning:** The following operations are irreversible. Please proceed with caution.
            """)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🗑️ Clear All Filters", use_container_width=True):
                    st.info("Filters cleared (feature in development)")
            
            with col2:
                if st.button("🔄 Reset Connection", use_container_width=True):
                    st.session_state.authenticated = False
                    st.rerun()
    
    # Footer
    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    <div style="text-align: center; padding: 1rem; background: white; border-radius: 10px; margin: 1rem 0;">
        <h4 style="color: #667eea; margin: 0;">📊 Quick Stats</h4>
    </div>
    """, unsafe_allow_html=True)
    
    try:
        df = read_bookings(worksheet)
        if not df.empty:
            stats = get_booking_statistics(df)
            
            st.sidebar.markdown(f"""
            <div style="background: white; padding: 1rem; border-radius: 10px;">
                <p style="margin: 5px 0;"><strong>Total Bookings:</strong> {stats['total']}</p>
                <p style="margin: 5px 0; color: #38ef7d;"><strong>New:</strong> {stats['new']}</p>
                <p style="margin: 5px 0; color: #00f2fe;"><strong>Reserved:</strong> {stats['reserved']}</p>
                <p style="margin: 5px 0; color: #f5576c;"><strong>Canceled:</strong> {stats['canceled']}</p>
            </div>
            """, unsafe_allow_html=True)
    except:
        pass
    
    st.sidebar.markdown("---")
    st.sidebar.caption("🏖️ Villa Booking System v2.0")
    st.sidebar.caption("Built with ❤️ using Streamlit")
    st.sidebar.caption("© 2025 - Real-time Google Sheets Integration")

if __name__ == "__main__":
    main()
