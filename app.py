import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, timedelta
import json
from typing import List, Dict, Optional
import re
import logging
from io import StringIO, BytesIO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import hashlib
import requests
import plotly.express as px
import plotly.graph_objects as go
import calendar as cal_module

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

DRIVE_FOLDER_ID = "1Fk5dJGkm5dNMZkfsITe5Lt9x-yCsBiF2"

# Page Configuration
st.set_page_config(
    page_title="Professional Booking Management System",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    .info-box, .warning-box, .success-box, .error-box, .metric-card, 
    .property-card, .feature-card, .analytics-card, .search-card {
        color: #000000;
    }
    .info-box h1, .info-box h2, .info-box h3, .info-box h4, .info-box p, 
    .info-box strong, .info-box span, .info-box div,
    .warning-box h1, .warning-box h2, .warning-box h3, .warning-box h4, 
    .warning-box p, .warning-box strong, .warning-box span, .warning-box div,
    .success-box h1, .success-box h2, .success-box h3, .success-box h4, 
    .success-box p, .success-box strong, .success-box span, .success-box div,
    .error-box h1, .error-box h2, .error-box h3, .error-box h4, 
    .error-box p, .error-box strong, .error-box span, .error-box div,
    .metric-card h1, .metric-card h2, .metric-card h3, .metric-card h4, 
    .metric-card p, .metric-card strong, .metric-card span, .metric-card div,
    .property-card h1, .property-card h2, .property-card h3, .property-card h4, 
    .property-card p, .property-card strong, .property-card span, .property-card div,
    .feature-card h1, .feature-card h2, .feature-card h3, .feature-card h4, 
    .feature-card p, .feature-card strong, .feature-card span, .feature-card div,
    .analytics-card h1, .analytics-card h2, .analytics-card h3, .analytics-card h4, 
    .analytics-card p, .analytics-card strong, .analytics-card span, .analytics-card div,
    .search-card h1, .search-card h2, .search-card h3, .search-card h4, 
    .search-card p, .search-card strong, .search-card span, .search-card div {
        color: #000000 !important;
    }
    .info-box {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #667eea;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .warning-box {
        background: linear-gradient(135deg, #fff9e6 0%, #ffe8b3 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #ffc107;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .success-box {
        background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #28a745;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .error-box {
        background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #f44336;
        margin: 1.5rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        text-align: center;
        border: 2px solid #e9ecef;
        transition: transform 0.3s ease;
    }
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.2);
    }
    .property-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px solid #e9ecef;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    .property-card:hover {
        border-color: #667eea;
        box-shadow: 0 6px 12px rgba(102, 126, 234, 0.3);
    }
    .feature-card {
        background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #2196F3;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .analytics-card {
        background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #ff9800;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .search-card {
        background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #9c27b0;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
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
        color: #000000;
    }
    .log-info { border-left-color: #17a2b8; }
    .log-success { border-left-color: #28a745; }
    .log-warning { border-left-color: #ffc107; }
    .log-error { border-left-color: #dc3545; }
    .log-email { border-left-color: #9c27b0; }
    .drive-config-card {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        color: white;
        box-shadow: 0 6px 12px rgba(0,0,0,0.2);
    }
    .email-card {
        background: linear-gradient(135deg, #ee9ca7 0%, #ffdde1 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #e91e63;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        color: #000000;
    }
    .email-card h1, .email-card h2, .email-card h3, .email-card h4, 
    .email-card p, .email-card strong, .email-card span, .email-card div {
        color: #000000 !important;
    }
    .edit-panel {
        background: linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #ff6b6b;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        color: #000000;
    }
    .edit-panel h1, .edit-panel h2, .edit-panel h3, .edit-panel h4, 
    .edit-panel p, .edit-panel strong, .edit-panel span, .edit-panel div {
        color: #000000 !important;
    }
    .notification-badge {
        background: #dc3545;
        color: white;
        padding: 0.2rem 0.5rem;
        border-radius: 10px;
        font-size: 0.75rem;
        font-weight: bold;
        margin-left: 0.5rem;
    }
    .bulk-edit-card {
        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #00d2ff;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        color: #000000;
    }
    .bulk-edit-card h1, .bulk-edit-card h2, .bulk-edit-card h3, .bulk-edit-card h4, 
    .bulk-edit-card p, .bulk-edit-card strong, .bulk-edit-card span, .bulk-edit-card div {
        color: #000000 !important;
    }
    .calendar-card {
        background: linear-gradient(135deg, #ffeaa7 0%, #fdcb6e 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #fdcb6e;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        color: #000000;
    }
    .template-card {
        background: linear-gradient(135deg, #a29bfe 0%, #6c5ce7 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        margin: 1rem 0;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    /* Specific styles for enhanced calendar view */
    .calendar-grid {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        gap: 5px;
        margin: 20px 0;
    }
    .calendar-day {
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 10px;
        min-height: 120px;
        background: white;
    }
    .calendar-day-header {
        font-weight: bold;
        color: #333;
        margin-bottom: 5px;
    }
    .calendar-day-empty {
        background: #f5f5f5;
    }
    .booking-item {
        background: #e3f2fd;
        border-left: 3px solid #2196f3;
        padding: 4px 6px;
        margin: 3px 0;
        font-size: 0.85em;
        border-radius: 3px;
        cursor: pointer;
    }
    .booking-item:hover {
        background: #bbdefb;
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
if 'logs' not in st.session_state:
    st.session_state.logs = []
if 'service_account_email' not in st.session_state:
    st.session_state.service_account_email = None
if 'all_sheets' not in st.session_state:
    st.session_state.all_sheets = []
if 'selected_sheet_index' not in st.session_state:
    st.session_state.selected_sheet_index = 0
if 'email_logs' not in st.session_state:
    st.session_state.email_logs = []
if 'email_config' not in st.session_state:
    st.session_state.email_config = {}
if 'edit_history' not in st.session_state:
    st.session_state.edit_history = []
if 'pending_changes' not in st.session_state:
    st.session_state.pending_changes = []
if 'notification_count' not in st.session_state:
    st.session_state.notification_count = 0
if 'templates' not in st.session_state:
    st.session_state.templates = []
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False
if 'auto_backup' not in st.session_state:
    st.session_state.auto_backup = True
if 'backup_history' not in st.session_state:
    st.session_state.backup_history = []

if 'selected_bookings' not in st.session_state:
    st.session_state.selected_bookings = []
if 'calendar_view_date' not in st.session_state:
    st.session_state.calendar_view_date = datetime.now()
if 'quick_filters' not in st.session_state:
    st.session_state.quick_filters = {}
if 'booking_form_data' not in st.session_state:
    st.session_state.booking_form_data = {}
if 'grid_edit_mode' not in st.session_state:
    st.session_state.grid_edit_mode = False
if 'visual_calendar_data' not in st.session_state:
    st.session_state.visual_calendar_data = {}
if 'selected_calendar' not in st.session_state: # For enhanced calendar view
    st.session_state.selected_calendar = None
if 'show_date_details' not in st.session_state: # For enhanced calendar view
    st.session_state.show_date_details = False
if 'selected_date' not in st.session_state: # For enhanced calendar view
    st.session_state.selected_date = None
if 'editing_booking' not in st.session_state: # For enhanced calendar view
    st.session_state.editing_booking = None
if 'editing_calendar' not in st.session_state: # For enhanced calendar view
    st.session_state.editing_calendar = None


CALENDAR_COLUMNS = {
    'DATE': 0,
    'VILLA': 1,
    'TYPE CLEAN': 2,
    'PAX': 3,
    'START TIME': 4,
    'RESERVATION STATUS': 5,
    'LAUNDRY': 6,
    'COMMENTS': 7
}

HEADER_ROW = 13
DATA_START_ROW = 14

STATUS_CODES = {
    'CI': {'name': 'Check-In', 'description': 'Complete cleaning and preparation for incoming guests', 'color': '#4CAF50', 'icon': '🏠'},
    'SO': {'name': 'Stay-over', 'description': 'Mid-stay cleaning with linen and towel refresh', 'color': '#2196F3', 'icon': '🔄'},
    'CO/CI': {'name': 'Check-out/Check-in', 'description': 'Same-day turnover between guests', 'color': '#FF9800', 'icon': '🔁'},
    'FU': {'name': 'Fresh-up', 'description': 'Quick refresh for already clean property', 'color': '#9C27B0', 'icon': '✨'},
    'DC': {'name': 'Deep Cleaning', 'description': 'Thorough deep clean of entire property', 'color': '#F44336', 'icon': '🧹'},
    'COC': {'name': 'Construction Cleaning', 'description': 'Post-renovation construction cleanup', 'color': '#795548', 'icon': '🏗️'},
    'CO': {'name': 'Check-Out', 'description': 'Final cleaning after guest departure', 'color': '#607D8B', 'icon': '🚪'}
}

# Data loading functions
@st.cache_data(ttl=300)
def load_client_data():
    """Load client information from CSV"""
    url = "https://hebbkx1anhila5yf.public.blob.vercel-storage.com/Reservations%20%20Casa%20Bohemian%20Curacao%20-%20Casa%20Bohemian%20Client%20Info%20-IePfFT3QCALdQtGNNBVbbLkVBW4eDR.csv"
    try:
        response = requests.get(url)
        response.raise_for_status()
        df = pd.read_csv(StringIO(response.text))
        return df
    except Exception as e:
        st.error(f"Error loading client data: {str(e)}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def load_reservations_data():
    """Load reservations from CSV"""
    url = "https://hebbkx1anhila5yf.public.blob.vercel-storage.com/Reservations%20%20Casa%20Bohemian%20Curacao%20-%20Nov%2725-uf4T55IrkhN7ak0nY1OqWgPhIaAWdA.csv"
    try:
        response = requests.get(url)
        response.raise_for_status()
        df = pd.read_csv(StringIO(response.text))
        return df
    except Exception as e:
        st.error(f"Error loading reservations: {str(e)}")
        return pd.DataFrame()

def add_log(message: str, level: str = "INFO"):
    """Add a log entry with timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {
        'timestamp': timestamp,
        'level': level,
        'message': message
    }
    st.session_state.logs.append(log_entry)
    
    if len(st.session_state.logs) > 500:
        st.session_state.logs = st.session_state.logs[-500:]
    
    if level == "INFO":
        logger.info(message)
    elif level == "SUCCESS":
        logger.info(f"✓ {message}")
    elif level == "WARNING":
        logger.warning(message)
    elif level == "ERROR":
        logger.error(message)
    elif level == "EMAIL":
        logger.info(f"📧 {message}")

def add_edit_history(action: str, details: Dict):
    """Track edit history for audit trail"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    edit_entry = {
        'timestamp': timestamp,
        'action': action,
        'details': details,
        'user': st.session_state.service_account_email
    }
    st.session_state.edit_history.append(edit_entry)
    
    if len(st.session_state.edit_history) > 1000:
        st.session_state.edit_history = st.session_state.edit_history[-1000:]
    
    add_log(f"Edit: {action} - {details}", "SUCCESS")

def create_backup(workbook, sheet_name: str) -> Dict:
    """Create a backup of current sheet data"""
    try:
        sheet = workbook.worksheet(sheet_name)
        all_values = sheet.get_all_values()
        
        backup = {
            'timestamp': datetime.now().isoformat(),
            'workbook': workbook.title,
            'sheet': sheet_name,
            'data': all_values,
            'row_count': len(all_values),
            'col_count': len(all_values[0]) if all_values else 0
        }
        
        st.session_state.backup_history.append(backup)
        
        # Keep only last 50 backups
        if len(st.session_state.backup_history) > 50:
            st.session_state.backup_history = st.session_state.backup_history[-50:]
        
        add_log(f"Backup created: {sheet_name} ({len(all_values)} rows)", "SUCCESS")
        return backup
    except Exception as e:
        add_log(f"Backup failed: {str(e)}", "ERROR")
        return None

def restore_backup(backup: Dict, manager) -> bool:
    """Restore data from a backup"""
    try:
        workbook = manager.gc.open(backup['workbook'])
        sheet = workbook.worksheet(backup['sheet'])
        
        # Clear existing data
        sheet.clear()
        
        # Restore backup data
        if backup['data']:
            sheet.update(backup['data'], 'A1')
        
        add_log(f"Backup restored: {backup['sheet']} from {backup['timestamp']}", "SUCCESS")
        return True
    except Exception as e:
        add_log(f"Restore failed: {str(e)}", "ERROR")
        return False

def save_template(name: str, description: str, data: Dict):
    """Save a booking template"""
    template = {
        'id': hashlib.md5(f"{name}{datetime.now().isoformat()}".encode()).hexdigest()[:8],
        'name': name,
        'description': description,
        'data': data,
        'created': datetime.now().isoformat(),
        'usage_count': 0
    }
    st.session_state.templates.append(template)
    add_log(f"Template saved: {name}", "SUCCESS")
    return template

def apply_template(template: Dict, sheet, row_index: int, manager) -> bool:
    """Apply a template to a specific row"""
    try:
        data = template['data']
        updates = []
        
        for col_idx, value in enumerate(data.values(), start=1):
            updates.append({
                'row': row_index,
                'col': col_idx,
                'value': value
            })
        
        if manager.batch_update_cells(sheet, updates):
            template['usage_count'] += 1
            add_log(f"Template applied: {template['name']} to row {row_index}", "SUCCESS")
            return True
        return False
    except Exception as e:
        add_log(f"Template application failed: {str(e)}", "ERROR")
        return False

class EmailManager:
    """Manages email sending functionality"""
    
    @staticmethod
    def configure_email(smtp_server: str, smtp_port: int, sender_email: str, sender_password: str):
        """Configure email settings"""
        st.session_state.email_config = {
            'smtp_server': smtp_server,
            'smtp_port': smtp_port,
            'sender_email': sender_email,
            'sender_password': sender_password,
            'configured': True
        }
        add_log(f"Email configured: {sender_email} via {smtp_server}:{smtp_port}", "EMAIL")
        return True
    
    @staticmethod
    def send_email(to_email: str, subject: str, body: str, attachment_data: Optional[bytes] = None, 
                   attachment_name: Optional[str] = None) -> bool:
        """Send email with optional attachment"""
        try:
            if not st.session_state.email_config.get('configured'):
                add_log("Email not configured. Please configure email settings first.", "ERROR")
                return False
            
            config = st.session_state.email_config
            
            add_log(f"Preparing email to {to_email}", "EMAIL")
            
            msg = MIMEMultipart()
            msg['From'] = config['sender_email']
            msg['To'] = to_email
            msg['Subject'] = subject
            msg['Date'] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z")
            
            msg.attach(MIMEText(body, 'html'))
            
            if attachment_data and attachment_name:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment_data)
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename= {attachment_name}')
                msg.attach(part)
                add_log(f"Attached file: {attachment_name}", "EMAIL")
            
            add_log(f"Connecting to SMTP server {config['smtp_server']}:{config['smtp_port']}", "EMAIL")
            
            server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
            server.starttls()
            server.login(config['sender_email'], config['sender_password'])
            
            text = msg.as_string()
            server.sendmail(config['sender_email'], to_email, text)
            server.quit()
            
            email_log = {
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'to': to_email,
                'subject': subject,
                'status': 'Sent',
                'has_attachment': attachment_name is not None
            }
            st.session_state.email_logs.append(email_log)
            
            add_log(f"✅ Email sent successfully to {to_email}", "SUCCESS")
            st.session_state.notification_count += 1
            return True
            
        except Exception as e:
            add_log(f"❌ Failed to send email: {str(e)}", "ERROR")
            email_log = {
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'to': to_email,
                'subject': subject,
                'status': f'Failed: {str(e)}',
                'has_attachment': False
            }
            st.session_state.email_logs.append(email_log)
            return False
    
    @staticmethod
    def send_booking_summary(to_email: str, workbook_name: str, summary_data: pd.DataFrame) -> bool:
        """Send booking summary as email"""
        subject = f"Booking Summary - {workbook_name}"
        
        body = f"""
        <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
                    th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                    th {{ background-color: #667eea; color: white; }}
                    tr:nth-child(even) {{ background-color: #f2f2f2; }}
                    h2 {{ color: #667eea; }}
                </style>
            </head>
            <body>
                <h2>📊 Booking Summary Report</h2>
                <p><strong>Client:</strong> {workbook_name}</p>
                <p><strong>Generated:</strong> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                <p><strong>Total Bookings:</strong> {len(summary_data)}</p>
                
                {summary_data.to_html(index=False, escape=False)}
                
                <br>
                <p style="color: #666; font-size: 0.9em;">
                    This is an automated report from the Professional Booking Management System.
                </p>
            </body>
        </html>
        """
        
        csv_data = summary_data.to_csv(index=False).encode()
        filename = f"booking_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        return EmailManager.send_email(to_email, subject, body, csv_data, filename)

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
            
            self.creds = Credentials.from_service_account_info(credentials_dict, scopes=self.scopes)
            
            add_log("Creating Google Sheets client...", "INFO")
            self.gc = gspread.authorize(self.creds)
            
            try:
                from googleapiclient.discovery import build
                self.drive_service = build('drive', 'v3', credentials=self.creds)
                add_log("✅ Drive API service initialized", "SUCCESS")
            except ImportError:
                self.drive_service = None
                add_log("⚠️ google-api-python-client not installed", "WARNING")
            
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
            add_log("=" * 70, "INFO")
            
            workbooks = []
            
            if self.drive_service:
                try:
                    query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
                    
                    page_token = None
                    total_files = 0
                    
                    while True:
                        results = self.drive_service.files().list(
                            q=query,
                            pageSize=100,
                            fields="nextPageToken, files(id, name, webViewLink, modifiedTime)",
                            supportsAllDrives=True,
                            includeItemsFromAllDrives=True,
                            corpora='allDrives',
                            pageToken=page_token
                        ).execute()
                        
                        files = results.get('files', [])
                        
                        for file in files:
                            total_files += 1
                            workbooks.append({
                                'id': file['id'],
                                'name': file['name'],
                                'url': file.get('webViewLink', f"https://docs.google.com/spreadsheets/d/{file['id']}"),
                                'modified': file.get('modifiedTime', 'Unknown')
                            })
                        
                        page_token = results.get('nextPageToken')
                        if not page_token:
                            break
                    
                    add_log(f"✅ Found {total_files} spreadsheet(s)", "SUCCESS")
                    return workbooks
                    
                except Exception as e:
                    add_log(f"❌ Drive API Error: {str(e)}", "ERROR")
                    return []
            else:
                add_log("❌ Drive API not available", "ERROR")
                return []
                
        except Exception as e:
            add_log(f"❌ CRITICAL ERROR: {str(e)}", "ERROR")
            return []
    
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
            
            return workbook
        except Exception as e:
            add_log(f"Error opening workbook: {str(e)}", "ERROR")
            return None
    
    def get_client_profile(self, workbook) -> Dict:
        """Extract client profile from first sheet"""
        try:
            sheet = workbook.get_worksheet(0)
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
            
            for i in range(17, min(len(all_values), 30)):
                if len(all_values[i]) > 1 and all_values[i][0]:
                    profile['properties'].append({
                        'name': all_values[i][0],
                        'address': all_values[i][1] if len(all_values[i]) > 1 else '',
                        'hours': all_values[i][2] if len(all_values[i]) > 2 else '',
                        'so_hours': all_values[i][3] if len(all_values[i]) > 3 else ''
                    })
            
            return profile
        except Exception as e:
            add_log(f"Error reading client profile: {str(e)}", "ERROR")
            return {}
    
    def get_calendar_sheets(self, workbook) -> List[Dict]:
        """Get all calendar sheets (excluding first sheet which is Client Info)"""
        try:
            worksheets = workbook.worksheets()
            calendars = []
            
            for i, sheet in enumerate(worksheets[1:], start=1):
                calendars.append({
                    'index': i,
                    'name': sheet.title,
                    'sheet': sheet,
                    'property_name': sheet.title  # Sheet name is property name
                })
            
            add_log(f"Found {len(calendars)} property calendar sheets", "INFO")
            return calendars
        except Exception as e:
            add_log(f"Error listing calendar sheets: {str(e)}", "ERROR")
            return []
    
    def read_calendar(self, sheet, start_row: int = DATA_START_ROW) -> pd.DataFrame:
        """Read booking calendar starting from row 14 (data), with headers from row 13"""
        try:
            all_values = sheet.get_all_values()
            
            if len(all_values) < HEADER_ROW:
                return pd.DataFrame()
            
            # Get headers from row 13 (index 12)
            headers = all_values[HEADER_ROW - 1]
            
            # Clean headers and handle duplicates
            seen = {}
            unique_headers = []
            for header in headers:
                clean_header = header.strip()
                if clean_header in seen:
                    seen[clean_header] += 1
                    unique_headers.append(f"{clean_header}_{seen[clean_header]}")
                else:
                    seen[clean_header] = 0
                    unique_headers.append(clean_header)
            
            # Get data starting from row 14 (index 13)
            data = all_values[DATA_START_ROW - 1:]
            
            # Create DataFrame
            df = pd.DataFrame(data, columns=unique_headers)
            
            # Remove completely empty rows
            df = df[df.apply(lambda row: row.astype(str).str.strip().any(), axis=1)]
            
            # Add actual row numbers for reference (starting from 14)
            df.insert(0, 'Row#', range(DATA_START_ROW, DATA_START_ROW + len(df)))
            
            add_log(f"Read {len(df)} bookings from {sheet.title}", "INFO")
            return df
        except Exception as e:
            add_log(f"Error reading calendar: {str(e)}", "ERROR")
            return pd.DataFrame()
    
    def create_booking(self, sheet, booking_data: Dict) -> bool:
        """Create a new booking in the calendar sheet"""
        try:
            new_row = [
                booking_data.get('date', ''),
                booking_data.get('villa', booking_data.get('property', '')),
                booking_data.get('type_clean', booking_data.get('service_type', '')),
                booking_data.get('pax', booking_data.get('guests', '')),
                booking_data.get('start_time', booking_data.get('time', '')),
                booking_data.get('reservation_status', booking_data.get('status', '')),
                booking_data.get('laundry', ''),
                booking_data.get('comments', booking_data.get('notes', ''))
            ]
            
            # Append to sheet
            sheet.append_row(new_row)
            
            add_edit_history("Create Booking", {
                'sheet': sheet.title,
                'booking_data': booking_data
            })
            
            add_log(f"✅ Booking created in {sheet.title}", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error creating booking: {str(e)}", "ERROR")
            return False
    
    def update_booking(self, sheet, row_num: int, booking_data: Dict) -> bool:
        """Update an existing booking"""
        try:
            updates = []
            
            if 'date' in booking_data:
                updates.append({'row': row_num, 'col': 1, 'value': booking_data['date']})
            if 'villa' in booking_data or 'property' in booking_data:
                updates.append({'row': row_num, 'col': 2, 'value': booking_data.get('villa', booking_data.get('property', ''))})
            if 'type_clean' in booking_data or 'service_type' in booking_data:
                updates.append({'row': row_num, 'col': 3, 'value': booking_data.get('type_clean', booking_data.get('service_type', ''))})
            if 'pax' in booking_data or 'guests' in booking_data:
                updates.append({'row': row_num, 'col': 4, 'value': booking_data.get('pax', booking_data.get('guests', ''))})
            if 'start_time' in booking_data or 'time' in booking_data:
                updates.append({'row': row_num, 'col': 5, 'value': booking_data.get('start_time', booking_data.get('time', ''))})
            if 'reservation_status' in booking_data or 'status' in booking_data:
                updates.append({'row': row_num, 'col': 6, 'value': booking_data.get('reservation_status', booking_data.get('status', ''))})
            if 'laundry' in booking_data:
                updates.append({'row': row_num, 'col': 7, 'value': booking_data['laundry']})
            if 'comments' in booking_data or 'notes' in booking_data:
                updates.append({'row': row_num, 'col': 8, 'value': booking_data.get('comments', booking_data.get('notes', ''))})
            
            return self.batch_update_cells(sheet, updates)
        except Exception as e:
            add_log(f"Error updating booking: {str(e)}", "ERROR")
            return False
    
    def get_calendar_data_for_month(self, sheet, year: int, month: int) -> Dict:
        """Get all bookings for a specific month organized by date"""
        try:
            df = self.read_calendar(sheet)
            
            if df.empty:
                return {}
            
            calendar_data = {}
            
            for _, row in df.iterrows():
                date_str = str(row.get('DATE', '')).strip()
                if not date_str:
                    continue
                
                try:
                    # Try multiple date formats
                    for fmt in ['%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                        try:
                            booking_date = datetime.strptime(date_str, fmt)
                            break
                        except ValueError:
                            continue
                    else:
                        continue
                    
                    # Check if booking is in the requested month
                    if booking_date.year == year and booking_date.month == month:
                        date_key = booking_date.strftime('%Y-%m-%d')
                        
                        if date_key not in calendar_data:
                            calendar_data[date_key] = []
                        
                        calendar_data[date_key].append({
                            'row': row.get('Row#', ''),
                            'date': date_str,
                            'villa': row.get('VILLA', ''),
                            'type_clean': row.get('TYPE CLEAN', ''),
                            'pax': row.get('PAX', ''),
                            'start_time': row.get('START TIME', ''),
                            'status': row.get('RESERVATION STATUS', ''),
                            'laundry': row.get('LAUNDRY', ''),
                            'comments': row.get('COMMENTS', '')
                        })
                except Exception as e:
                    continue
            
            return calendar_data
        except Exception as e:
            add_log(f"Error getting calendar data: {str(e)}", "ERROR")
            return {}
    
    # Rest of the existing code in BookingManager class ...
    def update_cell_live(self, sheet, row: int, col: int, value: str, log_edit: bool = True) -> bool:
        """Update a single cell in real-time with edit tracking"""
        try:
            old_value = sheet.cell(row, col).value
            sheet.update_cell(row, col, value)
            
            if log_edit:
                add_edit_history("Cell Update", {
                    'sheet': sheet.title,
                    'row': row,
                    'col': col,
                    'old_value': old_value,
                    'new_value': value
                })
            
            add_log(f"Cell [{row}, {col}] updated: '{old_value}' → '{value}'", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error updating cell: {str(e)}", "ERROR")
            return False
    
    def batch_update_cells(self, sheet, updates: List[Dict]) -> bool:
        """Batch update multiple cells at once"""
        try:
            add_log(f"Batch updating {len(updates)} cells in {sheet.title}", "INFO")
            
            # Group updates by row for efficiency
            for update in updates:
                sheet.update_cell(update['row'], update['col'], update['value'])
            
            add_edit_history("Batch Update", {
                'sheet': sheet.title,
                'cell_count': len(updates)
            })
            
            add_log(f"✅ Batch update completed: {len(updates)} cells", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error in batch update: {str(e)}", "ERROR")
            return False
    
    def bulk_text_update(self, sheet, start_row: int, start_col: int, text_data: str) -> bool:
        """Update multiple cells from pasted text data (tab/newline separated)"""
        try:
            lines = text_data.strip().split('\n')
            updates = []
            
            for row_offset, line in enumerate(lines):
                cells = line.split('\t')
                for col_offset, value in enumerate(cells):
                    updates.append({
                        'row': start_row + row_offset,
                        'col': start_col + col_offset,
                        'value': value.strip()
                    })
            
            return self.batch_update_cells(sheet, updates)
        except Exception as e:
            add_log(f"Error in bulk text update: {str(e)}", "ERROR")
            return False
    
    def append_row(self, sheet, data: List) -> bool:
        """Append a new row to the sheet"""
        try:
            sheet.append_row(data)
            
            add_edit_history("Row Appended", {
                'sheet': sheet.title,
                'row_data': data
            })
            
            add_log(f"Row appended successfully with {len(data)} cells", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error appending row: {str(e)}", "ERROR")
            return False
    
    def delete_row(self, sheet, row_index: int) -> bool:
        """Delete a row from the sheet"""
        try:
            sheet.delete_rows(row_index)
            
            add_edit_history("Row Deleted", {
                'sheet': sheet.title,
                'row_index': row_index
            })
            
            add_log(f"Row {row_index} deleted successfully", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error deleting row: {str(e)}", "ERROR")
            return False
    
    def copy_row(self, sheet, row_index: int) -> List:
        """Copy a row from the sheet"""
        try:
            all_values = sheet.get_all_values()
            
            if row_index < len(all_values):
                row_data = all_values[row_index]
                add_log(f"Row {row_index} copied: {len(row_data)} cells", "SUCCESS")
                return row_data
            else:
                add_log(f"Row index {row_index} out of range", "ERROR")
                return []
        except Exception as e:
            add_log(f"Error copying row: {str(e)}", "ERROR")
            return []
    
    def search_bookings(self, sheet, search_term: str, columns: List[str] = None) -> pd.DataFrame:
        """Search for bookings matching a term"""
        try:
            df = self.read_calendar(sheet)
            
            if df.empty:
                return pd.DataFrame()
            
            if columns:
                search_df = df[columns]
            else:
                search_df = df
            
            mask = search_df.astype(str).apply(
                lambda row: row.str.contains(search_term, case=False, na=False).any(), 
                axis=1
            )
            
            result = df[mask]
            add_log(f"Search '{search_term}' found {len(result)} results", "INFO")
            return result
            
        except Exception as e:
            add_log(f"Error searching bookings: {str(e)}", "ERROR")
            return pd.DataFrame()
    
    def get_bookings_by_date_range(self, sheet, start_date: datetime, end_date: datetime, 
                                   date_column: str = 'Date') -> pd.DataFrame:
        """Get bookings within a date range"""
        try:
            df = self.read_calendar(sheet)
            
            if df.empty or date_column not in df.columns:
                return df
            
            df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
            
            mask = (df[date_column] >= start_date) & (df[date_column] <= end_date)
            result = df[mask]
            
            add_log(f"Found {len(result)} bookings between {start_date.date()} and {end_date.date()}", "INFO")
            return result
            
        except Exception as e:
            add_log(f"Error filtering by date range: {str(e)}", "ERROR")
            return pd.DataFrame()

    # Added new BookingManager methods for new features
    def get_all_bookings_combined(self, workbook) -> pd.DataFrame:
        """Get all bookings from all calendar sheets combined"""
        try:
            calendars = self.get_calendar_sheets(workbook)
            all_bookings = []
            
            for cal in calendars:
                df = self.read_calendar(cal['sheet'])
                if not df.empty:
                    df['calendar_source'] = cal['name']
                    all_bookings.append(df)
            
            if all_bookings:
                combined = pd.concat(all_bookings, ignore_index=True)
                add_log(f"Combined {len(combined)} bookings from {len(calendars)} calendars", "SUCCESS")
                return combined
            return pd.DataFrame()
        except Exception as e:
            add_log(f"Error combining bookings: {str(e)}", "ERROR")
            return pd.DataFrame()
    
    def create_full_booking(self, sheet, booking_data: Dict) -> bool:
        """Create a complete booking with all fields"""
        try:
            add_log(f"Creating full booking in {sheet.title}", "INFO")
            
            # Get current headers to match data structure
            all_values = sheet.get_all_values()
            if len(all_values) < 12:
                add_log("Sheet structure invalid for booking creation", "ERROR")
                return False
            
            headers = all_values[11]  # Row 12 (index 11) typically has headers
            
            # Build row data matching headers
            row_data = []
            for header in headers:
                header_lower = header.lower()
                value = ''
                
                # Map booking_data to appropriate columns
                if 'date' in header_lower:
                    value = booking_data.get('date', '')
                elif 'property' in header_lower or 'location' in header_lower:
                    value = booking_data.get('property', '')
                elif 'guest' in header_lower or 'name' in header_lower:
                    value = booking_data.get('guest_name', '')
                elif 'status' in header_lower or 'code' in header_lower:
                    value = booking_data.get('status', '')
                elif 'time' in header_lower:
                    value = booking_data.get('time', '')
                elif 'notes' in header_lower or 'comment' in header_lower:
                    value = booking_data.get('notes', '')
                elif 'email' in header_lower:
                    value = booking_data.get('email', '')
                elif 'phone' in header_lower:
                    value = booking_data.get('phone', '')
                elif 'price' in header_lower or 'amount' in header_lower:
                    value = booking_data.get('price', '')
                elif 'duration' in header_lower or 'hours' in header_lower:
                    value = booking_data.get('duration', '')
                else:
                    value = booking_data.get(header, '')
                
                row_data.append(str(value))
            
            # Append the booking
            sheet.append_row(row_data)
            
            add_edit_history("Create Full Booking", {
                'sheet': sheet.title,
                'data': booking_data
            })
            
            add_log(f"✅ Full booking created successfully", "SUCCESS")
            return True
            
        except Exception as e:
            add_log(f"Error creating full booking: {str(e)}", "ERROR")
            return False
    
    def bulk_update_cells_optimized(self, sheet, updates: List[Dict]) -> bool:
        """Optimized bulk update using batch API"""
        try:
            add_log(f"Bulk updating {len(updates)} cells in {sheet.title}", "INFO")
            
            # Group updates by row for efficiency
            for update in updates:
                sheet.update_cell(update['row'], update['col'], update['value'])
            
            add_edit_history("Bulk Update", {
                'sheet': sheet.title,
                'cell_count': len(updates)
            })
            
            add_log(f"✅ Bulk update completed: {len(updates)} cells", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error in bulk update: {str(e)}", "ERROR")
            return False
    
    def duplicate_booking(self, sheet, row_index: int, new_date: str = None) -> bool:
        """Duplicate an existing booking with optional new date"""
        try:
            row_data = self.copy_row(sheet, row_index)
            if not row_data:
                return False
            
            # Update date if provided
            if new_date:
                # Find date column and update
                all_values = sheet.get_all_values()
                headers = all_values[11] if len(all_values) > 11 else []
                for i, header in enumerate(headers):
                    if 'date' in header.lower():
                        row_data[i] = new_date
                        break
            
            self.append_row(sheet, row_data)
            
            add_edit_history("Duplicate Booking", {
                'sheet': sheet.title,
                'source_row': row_index,
                'new_date': new_date
            })
            
            add_log(f"✅ Booking duplicated successfully", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error duplicating booking: {str(e)}", "ERROR")
            return False
    
    def delete_booking(self, sheet, row_index: int) -> bool:
        """Delete a booking row"""
        try:
            sheet.delete_rows(row_index)
            
            add_edit_history("Delete Booking", {
                'sheet': sheet.title,
                'row': row_index
            })
            
            add_log(f"✅ Booking deleted from row {row_index}", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error deleting booking: {str(e)}", "ERROR")
            return False
    
    def search_bookings(self, workbook, search_term: str, search_columns: List[str] = None) -> pd.DataFrame:
        """Search across all bookings"""
        try:
            all_bookings = self.get_all_bookings_combined(workbook)
            
            if all_bookings.empty:
                return pd.DataFrame()
            
            if search_columns:
                # Search specific columns
                mask = False
                for col in search_columns:
                    if col in all_bookings.columns:
                        mask |= all_bookings[col].astype(str).str.contains(search_term, case=False, na=False)
            else:
                # Search all columns
                mask = all_bookings.astype(str).apply(
                    lambda row: row.str.contains(search_term, case=False, na=False).any(), 
                    axis=1
                )
            
            results = all_bookings[mask]
            add_log(f"Search '{search_term}' found {len(results)} results", "SUCCESS")
            return results
            
        except Exception as e:
            add_log(f"Error searching bookings: {str(e)}", "ERROR")
            return pd.DataFrame()
    
    def get_bookings_by_date_range(self, workbook, start_date: datetime, end_date: datetime) -> pd.DataFrame:
        """Get bookings within a date range"""
        try:
            all_bookings = self.get_all_bookings_combined(workbook)
            
            if all_bookings.empty:
                return pd.DataFrame()
            
            # Find date columns
            date_cols = [col for col in all_bookings.columns if 'date' in col.lower()]
            
            if not date_cols:
                return all_bookings
            
            # Filter by date range
            date_col = date_cols[0]
            all_bookings[date_col] = pd.to_datetime(all_bookings[date_col], errors='coerce')
            
            mask = (all_bookings[date_col] >= start_date) & (all_bookings[date_col] <= end_date)
            results = all_bookings[mask]
            
            add_log(f"Found {len(results)} bookings between {start_date.date()} and {end_date.date()}", "SUCCESS")
            return results
            
        except Exception as e:
            add_log(f"Error filtering by date range: {str(e)}", "ERROR")
            return pd.DataFrame()

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
                        creds_dict = json.load(uploaded_file)
                        manager = BookingManager(creds_dict)
                        st.session_state.gc = manager
                        st.session_state.authenticated = True
                        
                        st.success("✅ Successfully connected to Google Sheets!")
                        st.info(f"📧 Service Account: {st.session_state.service_account_email}")
                        
                        with st.spinner(f"Scanning folder for spreadsheets..."):
                            st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
                        
                        if len(st.session_state.workbooks) > 0:
                            st.success(f"✅ Found {len(st.session_state.workbooks)} spreadsheet(s)!")
                            st.balloons()
                        else:
                            st.error("❌ No spreadsheets found in folder")
                        
                        st.rerun()
                        
                    except Exception as e:
                        add_log(f"Authentication failed: {str(e)}", "ERROR")
                        st.error(f"❌ Authentication failed: {str(e)}")
            else:
                st.warning("⚠️ Please upload a credentials file")

def render_dashboard(manager, workbook):
    """Render enhanced dashboard view with interactive cards"""
    st.markdown('<div class="section-header">📊 Interactive Dashboard</div>', unsafe_allow_html=True)
    
    profile = manager.get_client_profile(workbook)
    calendars = manager.get_calendar_sheets(workbook)
    
    # Client header with all info
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        st.markdown(f"""
        <div class="success-box">
            <h2 style="margin-top: 0;">👤 {profile.get('client_name', 'Unknown Client')}</h2>
            <p><strong>Check-out:</strong> {profile.get('check_out_time', 'N/A')} | 
               <strong>Check-in:</strong> {profile.get('check_in_time', 'N/A')}</p>
            <p><strong>Amenities:</strong> {profile.get('amenities', 'N/A')}</p>
            <p><strong>Laundry:</strong> {profile.get('laundry_services', 'N/A')}</p>
            <p><strong>Keys:</strong> {profile.get('keys', 'N/A')} | <strong>Codes:</strong> {profile.get('codes', 'N/A')}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.metric("🏘️ Properties", len(profile.get('properties', [])))
    
    with col3:
        st.metric("📅 Calendars", len(calendars))
    
    # Quick stats with enhanced metrics
    total_bookings = 0
    upcoming_bookings = 0
    today = datetime.now().date()
    status_counts = {}
    
    for cal in calendars:
        df = manager.read_calendar(cal['sheet'])
        if not df.empty:
            total_bookings += len(df)
            
            # Count status codes
            for col in df.columns:
                if 'status' in col.lower() or 'code' in col.lower():
                    for status in df[col]:
                        status_counts[status] = status_counts.get(status, 0) + 1
            
            for col in df.columns:
                if 'date' in col.lower():
                    dates = pd.to_datetime(df[col], errors='coerce')
                    upcoming_bookings += (dates >= pd.Timestamp(today)).sum()
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <h4 style="color: #667eea;">📋 Total Bookings</h4>
            <h2>{total_bookings}</h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <h4 style="color: #11998e;">📅 Upcoming</h4>
            <h2>{upcoming_bookings}</h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <h4 style="color: #f093fb;">✏️ Edits Today</h4>
            <h2>{len([e for e in st.session_state.edit_history if e['timestamp'].startswith(today.strftime('%Y-%m-%d'))])}</h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <h4 style="color: #ff9800;">📧 Emails Sent</h4>
            <h2>{len(st.session_state.email_logs)}</h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div class="metric-card">
            <h4 style="color: #e91e63;">💾 Backups</h4>
            <h2>{len(st.session_state.backup_history)}</h2>
        </div>
        """, unsafe_allow_html=True)
    
    # Status code distribution
    if status_counts:
        st.markdown("### 📊 Status Code Distribution")
        status_df = pd.DataFrame(list(status_counts.items()), columns=['Status', 'Count'])
        
        fig = px.pie(status_df, values='Count', names='Status', 
                     title='Booking Status Distribution',
                     color_discrete_sequence=px.colors.qualitative.Set3)
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("### 🏘️ Property Portfolio - Click to View Details")
    
    if profile.get('properties'):
        # Display properties in expandable cards
        for idx, prop in enumerate(profile['properties']):
            with st.expander(f"🏠 {prop['name']}", expanded=(idx == 0)):
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.markdown(f"""
                    <div class="property-card">
                        <h3 style="margin-top: 0; color: #667eea;">{prop['name']}</h3>
                        <p><strong>📍 Address:</strong> {prop['address']}</p>
                        <p><strong>⏱️ Standard Hours:</strong> {prop['hours']}</p>
                        <p><strong>🔄 Stay-over Hours:</strong> {prop['so_hours']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    # Quick actions for property
                    if st.button(f"📅 View Bookings", key=f"view_prop_{idx}"):
                        st.session_state.quick_filters['property'] = prop['name']
                        st.info(f"Filter applied: {prop['name']}")
                    
                    if st.button(f"➕ New Booking", key=f"new_prop_{idx}"):
                        st.session_state.booking_form_data['property'] = prop['name']
                        st.success(f"Property pre-filled: {prop['name']}")
    
    st.markdown("### 📅 Calendar Sheets - Click to View & Edit")
    
    cal_cols = st.columns(min(len(calendars), 4))
    
    for idx, cal in enumerate(calendars):
        with cal_cols[idx % 4]:
            df = manager.read_calendar(cal['sheet'])
            booking_count = len(df) if not df.empty else 0
            
            st.markdown(f"""
            <div class="calendar-card">
                <h4 style="margin: 0; color: #000000;">📆 {cal['name']}</h4>
                <h2 style="margin: 0.5rem 0; color: #000000;">{booking_count}</h2>
                <p style="margin: 0; color: #000000;">bookings</p>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button(f"📖 Open", key=f"open_cal_{idx}", use_container_width=True):
                st.session_state.selected_sheet_index = cal['index']
                st.success(f"Switched to: {cal['name']}")
                st.rerun()
            
            if st.button(f"✏️ Edit", key=f"edit_cal_{idx}", use_container_width=True):
                st.session_state.selected_sheet_index = cal['index']
                st.session_state.grid_edit_mode = True
                st.rerun()
    
    # Recent activity with enhanced display
    st.markdown("### 📊 Recent Activity")
    
    if st.session_state.edit_history:
        recent_edits = st.session_state.edit_history[-10:]
        
        for edit in reversed(recent_edits):
            st.markdown(f"""
            <div class="log-entry log-success">
                <strong>[{edit['timestamp']}]</strong> {edit['action']} by {edit['user'][:30]}...
                <br><small>{str(edit['details'])[:100]}...</small>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No recent activity")
    
    # Quick actions
    st.markdown("### ⚡ Quick Actions")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("📝 New Booking", use_container_width=True):
            st.session_state.quick_action = "new_booking"
    
    with col2:
        if st.button("📧 Send Report", use_container_width=True):
            st.session_state.quick_action = "send_report"
    
    with col3:
        if st.button("💾 Create Backup", use_container_width=True):
            if calendars:
                create_backup(workbook, calendars[0]['name'])
                st.success("✅ Backup created!")
    
    with col4:
        if st.button("📊 View Analytics", use_container_width=True):
            st.session_state.quick_action = "analytics"

def render_bulk_text_editor(manager, workbook):
    """Render bulk text editing interface for pasting from Excel/Sheets"""
    st.markdown('<div class="section-header">📝 Bulk Text Editor</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="bulk-edit-card">
        <h3 style="margin-top: 0;">✨ Paste & Update Multiple Cells</h3>
        <p>Copy cells from Excel or Google Sheets and paste them here to update multiple cells at once.</p>
        <p><strong>Format:</strong> Tab-separated columns, newline-separated rows (standard copy-paste format)</p>
    </div>
    """, unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar = st.selectbox("📅 Select Calendar", calendar_names)
    
    selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
    
    if selected_cal:
        sheet = selected_cal['sheet']
        
        # Show current data
        df = manager.read_calendar(sheet)
        
        if not df.empty:
            st.markdown("### Current Data")
            st.dataframe(df, use_container_width=True, height=300)
        
        st.markdown("---")
        st.markdown("### Paste Your Data")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            start_row = st.number_input("Starting Row", min_value=1, value=13, 
                                       help="Row number where pasted data will start")
        
        with col2:
            start_col = st.number_input("Starting Column", min_value=1, value=1,
                                       help="Column number where pasted data will start")
        
        # Large text area for pasting
        pasted_data = st.text_area(
            "Paste Data Here (Ctrl+V / Cmd+V)",
            height=300,
            placeholder="Copy cells from Excel/Sheets and paste here...\nExample:\nDate\tProperty\tStatus\n2025-01-15\tVilla A\tCI\n2025-01-16\tVilla B\tSO",
            help="Paste tab-separated data directly from spreadsheets"
        )
        
        if pasted_data:
            # Preview what will be updated
            st.markdown("### Preview")
            lines = pasted_data.strip().split('\n')
            preview_data = []
            
            for row_offset, line in enumerate(lines[:10]):  # Show first 10 rows
                cells = line.split('\t')
                preview_data.append({
                    'Row': start_row + row_offset,
                    'Data': ' | '.join(cells[:5])  # Show first 5 columns
                })
            
            st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
            
            if len(lines) > 10:
                st.info(f"... and {len(lines) - 10} more rows")
            
            col1, col2, col3 = st.columns([1, 1, 2])
            
            with col1:
                if st.button("💾 Update Cells", type="primary", use_container_width=True):
                    with st.spinner("Updating cells..."):
                        if st.session_state.auto_backup:
                            create_backup(workbook, sheet.title)
                        
                        if manager.bulk_text_update(sheet, start_row, start_col, pasted_data):
                            st.success(f"✅ Updated {len(lines)} rows successfully!")
                            st.balloons()
                            st.rerun()
                        else:
                            st.error("❌ Update failed")
            
            with col2:
                if st.button("🔄 Clear", use_container_width=True):
                    st.rerun()

def render_live_editor(manager, workbook):
    """Render advanced live editing interface"""
    st.markdown('<div class="section-header">✏️ Live Booking Editor</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar = st.selectbox("📅 Select Calendar", calendar_names)
    
    selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
    
    if selected_cal:
        sheet = selected_cal['sheet']
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Edit Cells", "➕ Add Booking", "🗑️ Delete Row", "📋 Bulk Operations", "🎨 Quick Status"])
        
        with tab1:
            st.markdown("### Single Cell Editor")
            
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                st.dataframe(df, use_container_width=True, height=400)
                
                col1, col2, col3, col4 = st.columns([2, 1, 2, 1])
                
                with col1:
                    edit_row = st.number_input("Row #", min_value=13, value=13, key="edit_row")
                
                with col2:
                    edit_col = st.number_input("Col #", min_value=1, value=1, key="edit_col")
                
                with col3:
                    new_value = st.text_input("New Value", key="new_val")
                
                with col4:
                    st.write("")
                    st.write("")
                    if st.button("💾 Update", type="primary", use_container_width=True):
                        if manager.update_cell_live(sheet, edit_row, edit_col, new_value):
                            st.success("✅ Updated!")
                            st.rerun()
                        else:
                            st.error("❌ Failed")
        
        with tab2:
            st.markdown("### Add New Booking")
            
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                st.info(f"Sheet has {len(df.columns)} columns")
                
                # Check if templates exist
                if st.session_state.templates:
                    use_template = st.checkbox("Use Template")
                    if use_template:
                        template_names = [t['name'] for t in st.session_state.templates]
                        selected_template_name = st.selectbox("Select Template", template_names)
                        selected_template = next((t for t in st.session_state.templates if t['name'] == selected_template_name), None)
                        
                        if selected_template and st.button("Apply Template"):
                            new_row = list(selected_template['data'].values())
                            if manager.append_row(sheet, new_row):
                                st.success("✅ Booking added from template!")
                                st.rerun()
                
                with st.form("add_booking_form"):
                    new_row = []
                    cols = st.columns(2)
                    
                    for i, col_name in enumerate(df.columns):
                        with cols[i % 2]:
                            value = st.text_input(f"{col_name}", key=f"add_{i}")
                            new_row.append(value)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        submitted = st.form_submit_button("➕ Add Booking", type="primary")
                    with col2:
                        save_as_template = st.form_submit_button("💾 Save as Template")
                    
                    if submitted:
                        if manager.append_row(sheet, new_row):
                            st.success("✅ Booking added!")
                            st.rerun()
                        else:
                            st.error("❌ Failed to add")
                    
                    if save_as_template:
                        template_name = st.text_input("Template Name")
                        if template_name:
                            template_data = {df.columns[i]: new_row[i] for i in range(len(new_row))}
                            save_template(template_name, "Booking template", template_data)
                            st.success("✅ Template saved!")
        
        with tab3:
            st.markdown("### Delete Row")
            
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                st.dataframe(df, use_container_width=True, height=300)
                
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    row_to_delete = st.number_input(
                        "Row Number to Delete (from row 13)", 
                        min_value=13, 
                        value=13,
                        help="WARNING: This action cannot be undone!"
                    )
                
                with col2:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ Delete", type="secondary", use_container_width=True):
                        confirm = st.checkbox("I confirm deletion")
                        if confirm:
                            if st.session_state.auto_backup:
                                create_backup(workbook, sheet.title)
                            
                            if manager.delete_row(sheet, row_to_delete):
                                st.success("✅ Row deleted!")
                                st.rerun()
                            else:
                                st.error("❌ Failed")
                        else:
                            st.warning("⚠️ Please confirm deletion")
        
        with tab4:
            st.markdown("### Bulk Operations")
            
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                st.markdown("#### 📋 Copy & Paste Rows")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    row_to_copy = st.number_input("Row to Copy", min_value=13, value=13)
                    
                    if st.button("📋 Copy Row", use_container_width=True):
                        copied_data = manager.copy_row(sheet, row_to_copy - 1)
                        if copied_data:
                            st.session_state['copied_row'] = copied_data
                            st.success(f"✅ Copied {len(copied_data)} cells")
                
                with col2:
                    if st.button("➕ Paste Row", use_container_width=True):
                        if 'copied_row' in st.session_state:
                            if manager.append_row(sheet, st.session_state['copied_row']):
                                st.success("✅ Pasted!")
                                st.rerun()
                        else:
                            st.warning("⚠️ No row copied")
                
                if 'copied_row' in st.session_state:
                    st.json(st.session_state['copied_row'])
                
                st.markdown("---")
                st.markdown("#### 🔄 Batch Update")
                
                st.text_area(
                    "Paste batch update commands (JSON format)",
                    height=150,
                    help='Format: [{"row": 13, "col": 1, "value": "text"}, ...]',
                    key="batch_json"
                )
                
                if st.button("🚀 Execute Batch Update", type="primary"):
                    try:
                        updates = json.loads(st.session_state.batch_json)
                        if st.session_state.auto_backup:
                            create_backup(workbook, sheet.title)
                        
                        if manager.batch_update_cells(sheet, updates):
                            st.success(f"✅ Updated {len(updates)} cells!")
                            st.rerun()
                    except json.JSONDecodeError:
                        st.error("❌ Invalid JSON format")
        
        with tab5:
            st.markdown("### Quick Status Update")
            st.markdown("Update status codes quickly for multiple bookings")
            
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                # Find status column
                status_col = None
                for col in df.columns:
                    if 'status' in col.lower() or 'code' in col.lower():
                        status_col = col
                        break
                
                if status_col:
                    st.info(f"Status column: {status_col}")
                    
                    # Display status codes
                    st.markdown("#### Available Status Codes")
                    cols = st.columns(len(STATUS_CODES))
                    
                    for idx, (code, info) in enumerate(STATUS_CODES.items()):
                        with cols[idx]:
                            st.markdown(f"""
                            <div style="background: {info['color']}; color: white; padding: 0.5rem; border-radius: 5px; text-align: center;">
                                <strong>{info['icon']} {code}</strong><br>
                                <small>{info['name']}</small>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    st.markdown("---")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        row_num = st.number_input("Row Number", min_value=13, value=13)
                    
                    with col2:
                        new_status = st.selectbox("New Status", list(STATUS_CODES.keys()))
                    
                    with col3:
                        st.write("")
                        st.write("")
                        if st.button("🎨 Update Status", type="primary", use_container_width=True):
                            # Find column index
                            col_idx = list(df.columns).index(status_col) + 1
                            if manager.update_cell_live(sheet, row_num, col_idx, new_status):
                                st.success(f"✅ Status updated to {new_status}!")
                                st.rerun()
                else:
                    st.warning("No status column found in this sheet")

# Render enhanced calendar view
def render_enhanced_calendar_view(manager, workbook):
    """Render enhanced calendar view with extended date range and interactive features"""
    st.markdown('<div class="section-header">📅 Extended Calendar View</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No property calendars found")
        return
    
    st.markdown("### 🏠 Select Property Calendar")
    
    cols = st.columns(min(len(calendars), 4))
    
    for idx, calendar_info in enumerate(calendars):
        with cols[idx % 4]:
            if st.button(
                f"📍 {calendar_info['property_name']}", 
                key=f"cal_btn_{idx}",
                use_container_width=True
            ):
                st.session_state.selected_calendar = calendar_info['name']
    
    # Get selected calendar
    if 'selected_calendar' not in st.session_state:
        st.session_state.selected_calendar = calendars[0]['name']
    
    selected_cal = next((cal for cal in calendars if cal['name'] == st.session_state.selected_calendar), calendars[0])
    
    st.markdown(f"### 📆 {selected_cal['property_name']} - Booking Calendar")
    
    col1, col2, col3 = st.columns([2, 2, 1])
    
    with col1:
        current_date = datetime.now()
        selected_month = st.selectbox(
            "Month",
            range(1, 13),
            index=current_date.month - 1,
            format_func=lambda x: cal_module.month_name[x]
        )
    
    with col2:
        selected_year = st.selectbox(
            "Year",
            range(current_date.year - 1, current_date.year + 3),
            index=1
        )
    
    with col3:
        view_mode = st.radio("View", ["Month", "Week"], horizontal=True)
    
    # Get calendar data
    calendar_data = manager.get_calendar_data_for_month(selected_cal['sheet'], selected_year, selected_month)
    
    if view_mode == "Month":
        render_month_calendar(calendar_data, selected_year, selected_month, selected_cal, manager)
    else:
        render_week_calendar(calendar_data, selected_year, selected_month, selected_cal, manager)
    
    st.markdown("---")
    st.markdown("### ➕ Quick Add Booking")
    
    with st.expander("Create New Booking", expanded=False):
        render_quick_booking_form(manager, selected_cal, selected_year, selected_month)

def render_month_calendar(calendar_data: Dict, year: int, month: int, calendar_info: Dict, manager):
    """Render full month calendar grid with bookings"""
    
    # Get calendar for the month
    cal = cal_module.monthcalendar(year, month)
    
    # Day headers
    st.markdown("""
    <style>
    .calendar-grid {
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        gap: 5px;
        margin: 20px 0;
    }
    .calendar-day {
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 10px;
        min-height: 120px;
        background: white;
    }
    .calendar-day-header {
        font-weight: bold;
        color: #333;
        margin-bottom: 5px;
    }
    .calendar-day-empty {
        background: #f5f5f5;
    }
    .booking-item {
        background: #e3f2fd;
        border-left: 3px solid #2196f3;
        padding: 4px 6px;
        margin: 3px 0;
        font-size: 0.85em;
        border-radius: 3px;
        cursor: pointer;
    }
    .booking-item:hover {
        background: #bbdefb;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Render day headers
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    cols = st.columns(7)
    for idx, day in enumerate(days):
        with cols[idx]:
            st.markdown(f"**{day}**")
    
    # Render calendar weeks
    for week in cal:
        cols = st.columns(7)
        for idx, day in enumerate(week):
            with cols[idx]:
                if day == 0:
                    st.markdown('<div class="calendar-day calendar-day-empty"></div>', unsafe_allow_html=True)
                else:
                    date_key = f"{year}-{month:02d}-{day:02d}"
                    bookings = calendar_data.get(date_key, [])
                    
                    # Day number
                    st.markdown(f"**{day}**")
                    
                    # Bookings for this day
                    if bookings:
                        for booking in bookings[:3]:  # Show max 3 bookings
                            status_icon = STATUS_CODES.get(booking['status'], {}).get('icon', '📅')
                            st.markdown(
                                f"<div class='booking-item'>{status_icon} {booking['start_time']} - {booking['villa']}</div>",
                                unsafe_allow_html=True
                            )
                        
                        if len(bookings) > 3:
                            st.caption(f"+{len(bookings) - 3} more")
                        
                        # Click to view details
                        if st.button(f"View", key=f"view_{date_key}", use_container_width=True):
                            st.session_state.selected_date = date_key
                            st.session_state.show_date_details = True
                            st.rerun() # Rerun to show details immediately
    
    # Show selected date details
    if st.session_state.get('show_date_details') and st.session_state.get('selected_date'):
        render_date_details(calendar_data, st.session_state.selected_date, calendar_info, manager)

def render_week_calendar(calendar_data: Dict, year: int, month: int, calendar_info: Dict, manager):
    """Render week view calendar"""
    
    # Get current week
    today = datetime(year, month, 1)
    week_start = today - timedelta(days=today.weekday())
    
    st.markdown("### 📅 Week View")
    
    # Week navigation
    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        if st.button("◀ Previous Week"):
            st.session_state.calendar_view_date -= timedelta(days=7) # Use session state for date
            st.rerun()
    with col3:
        if st.button("Next Week ▶"):
            st.session_state.calendar_view_date += timedelta(days=7) # Use session state for date
            st.rerun()
    
    # Render 7 days
    for i in range(7):
        current_day = st.session_state.calendar_view_date.replace(day=1) + timedelta(days=i) # Use session state for date
        date_key = current_day.strftime('%Y-%m-%d')
        bookings = calendar_data.get(date_key, [])
        
        with st.expander(f"{current_day.strftime('%A, %B %d, %Y')} ({len(bookings)} bookings)", expanded=True):
            if bookings:
                for booking in bookings:
                    render_booking_card(booking, calendar_info, manager)
            else:
                st.info("No bookings for this day")

def render_date_details(calendar_data: Dict, date_key: str, calendar_info: Dict, manager):
    """Render detailed view of bookings for a specific date"""
    
    st.markdown("---")
    st.markdown(f"### 📋 Bookings for {date_key}")
    
    bookings = calendar_data.get(date_key, [])
    
    if not bookings:
        st.info("No bookings for this date")
        return
    
    for booking in bookings:
        render_booking_card(booking, calendar_info, manager)
    
    if st.button("✖ Close Details"):
        st.session_state.show_date_details = False
        st.session_state.selected_date = None # Clear selected date
        st.rerun()

def render_booking_card(booking: Dict, calendar_info: Dict, manager):
    """Render an individual booking card with edit capabilities"""
    
    status_info = STATUS_CODES.get(booking['status'], {'icon': '📅', 'name': 'Unknown', 'color': '#999'})
    
    with st.container():
        st.markdown(f"""
        <div style="background: {status_info['color']}15; border-left: 4px solid {status_info['color']}; padding: 15px; border-radius: 8px; margin: 10px 0;">
            <h4 style="margin: 0;">{status_info['icon']} {booking.get('villa', 'N/A')} - {booking.get('type_clean', 'N/A')}</h4>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"**🕐 Time:** {booking.get('start_time', 'N/A')}")
            st.markdown(f"**👥 PAX:** {booking.get('pax', 'N/A')}")
        
        with col2:
            st.markdown(f"**📊 Status:** {booking.get('status', 'N/A')}")
            st.markdown(f"**🧺 Laundry:** {booking.get('laundry', 'N/A')}")
        
        with col3:
            st.markdown(f"**💬 Comments:** {booking.get('comments', 'N/A')}")
            st.markdown(f"**📍 Row:** {booking.get('row', 'N/A')}")
        
        # Edit button
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("✏️ Edit", key=f"edit_{booking['row']}"):
                st.session_state.editing_booking = booking
                st.session_state.editing_calendar = calendar_info
                st.session_state.show_date_details = False # Close details to show modal
                st.rerun()
        
        with col2:
            if st.button("🗑️ Delete", key=f"del_{booking['row']}"):
                if manager.delete_row(calendar_info['sheet'], int(booking['row'])):
                    st.success("✅ Booking deleted!")
                    st.rerun()

def render_quick_booking_form(manager, calendar_info: Dict, year: int, month: int):
    """Render quick booking creation form"""
    
    with st.form(f"quick_booking_{calendar_info['name']}"):
        col1, col2 = st.columns(2)
        
        with col1:
            booking_date = st.date_input("Date", value=datetime(year, month, 1))
            start_time = st.time_input("Start Time", value=datetime.strptime('10:00', '%H:%M').time())
            villa = st.text_input("Villa/Property", value=calendar_info['property_name'])
            type_clean = st.selectbox("Type/Service", ["Deep Clean", "Regular Clean", "Check-in Clean", "Check-out Clean", "Maintenance"])
        
        with col2:
            pax = st.number_input("PAX (Guests)", min_value=1, max_value=20, value=2)
            status = st.selectbox("Status", list(STATUS_CODES.keys()), format_func=lambda x: f"{STATUS_CODES[x]['icon']} {x} - {STATUS_CODES[x]['name']}")
            laundry = st.selectbox("Laundry", ["Yes", "No", "N/A"])
            comments = st.text_area("Comments", height=80)
        
        submitted = st.form_submit_button("✅ Create Booking", type="primary", use_container_width=True)
        
        if submitted:
            booking_data = {
                'date': booking_date.strftime('%d/%m/%Y'),
                'villa': villa,
                'type_clean': type_clean,
                'pax': str(pax),
                'start_time': start_time.strftime('%H:%M'),
                'reservation_status': status,
                'laundry': laundry,
                'comments': comments
            }
            
            if manager.create_booking(calendar_info['sheet'], booking_data):
                st.success("✅ Booking created successfully!")
                st.balloons()
                st.rerun()
            else:
                st.error("❌ Failed to create booking")


def render_calendar_view(manager, workbook):
    """Render calendar view of bookings"""
    st.markdown('<div class="section-header">📅 Calendar View</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="calendar-card">
        <h3 style="margin-top: 0;">📆 Visual Booking Calendar</h3>
        <p>View and manage bookings in a calendar format</p>
    </div>
    """, unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendars found")
        return
    
    selected_calendar = st.selectbox("Select Calendar", [cal['name'] for cal in calendars])
    selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
    
    if selected_cal:
        sheet = selected_cal['sheet']
        df = manager.read_calendar(sheet)
        
        if not df.empty:
            # Find date column
            date_col = None
            for col in df.columns:
                if 'date' in col.lower():
                    date_col = col
                    break
            
            if date_col:
                # Date range selector
                col1, col2 = st.columns(2)
                
                with col1:
                    start_date = st.date_input("Start Date", datetime.now().date())
                
                with col2:
                    end_date = st.date_input("End Date", datetime.now().date() + timedelta(days=30))
                
                # Filter bookings by date range
                filtered_df = manager.get_bookings_by_date_range(
                    sheet, 
                    datetime.combine(start_date, datetime.min.time()),
                    datetime.combine(end_date, datetime.max.time()),
                    date_col
                )
                
                if not filtered_df.empty:
                    st.markdown(f"### Bookings: {start_date} to {end_date}")
                    st.markdown(f"**Total:** {len(filtered_df)} bookings")
                    
                    # Create timeline visualization
                    if date_col in filtered_df.columns:
                        filtered_df[date_col] = pd.to_datetime(filtered_df[date_col], errors='coerce')
                        
                        # Group by date
                        daily_counts = filtered_df.groupby(filtered_df[date_col].dt.date).size().reset_index()
                        daily_counts.columns = ['Date', 'Bookings']
                        
                        fig = px.bar(daily_counts, x='Date', y='Bookings',
                                    title='Daily Booking Count',
                                    color='Bookings',
                                    color_continuous_scale='Blues')
                        st.plotly_chart(fig, use_container_width=True)
                    
                    # Display bookings
                    st.dataframe(filtered_df, use_container_width=True)
                else:
                    st.info("No bookings in selected date range")
            else:
                st.warning("No date column found in this calendar")
                st.dataframe(df, use_container_width=True)

# Modal for editing booking
def render_edit_booking_modal(manager):
    """Modal for editing an existing booking"""
    if st.session_state.editing_booking and st.session_state.editing_calendar:
        booking = st.session_state.editing_booking
        calendar_info = st.session_state.editing_calendar
        
        with st.dialog(f"Edit Booking - Row {booking['row']}"):
            st.markdown(f"**Villa:** {booking['villa']} | **Date:** {booking['date']} | **Time:** {booking['start_time']}")
            
            with st.form("edit_booking_form"):
                col1, col2 = st.columns(2)
                
                with col1:
                    edit_villa = st.text_input("Villa/Property", value=booking.get('villa', ''))
                    edit_date = st.date_input("Date", value=datetime.strptime(booking.get('date', '01/01/2000'), '%d/%m/%Y'))
                    edit_start_time = st.time_input("Start Time", value=datetime.strptime(booking.get('start_time', '10:00'), '%H:%M').time())
                    edit_pax = st.number_input("PAX (Guests)", min_value=1, value=int(booking.get('pax', 1)))
                    edit_laundry = st.selectbox("Laundry", ["Yes", "No", "N/A"], index=["Yes", "No", "N/A"].index(booking.get('laundry', 'N/A')))
                
                with col2:
                    edit_type_clean = st.selectbox("Type/Service", ["Deep Clean", "Regular Clean", "Check-in Clean", "Check-out Clean", "Maintenance"], index=["Deep Clean", "Regular Clean", "Check-in Clean", "Check-out Clean", "Maintenance"].index(booking.get('type_clean', 'Deep Clean')))
                    edit_status = st.selectbox("Status", list(STATUS_CODES.keys()), format_func=lambda x: f"{STATUS_CODES[x]['icon']} {x} - {STATUS_CODES[x]['name']}", index=list(STATUS_CODES.keys()).index(booking.get('status', 'CI')))
                    edit_comments = st.text_area("Comments", value=booking.get('comments', ''), height=100)
                
                submitted = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
                
                if submitted:
                    updated_data = {
                        'villa': edit_villa,
                        'date': edit_date.strftime('%d/%m/%Y'),
                        'start_time': edit_start_time.strftime('%H:%M'),
                        'pax': str(edit_pax),
                        'laundry': edit_laundry,
                        'type_clean': edit_type_clean,
                        'status': edit_status,
                        'comments': edit_comments
                    }
                    
                    if manager.update_booking(calendar_info['sheet'], int(booking['row']), updated_data):
                        st.success("✅ Booking updated successfully!")
                        st.session_state.editing_booking = None
                        st.session_state.editing_calendar = None
                        st.rerun()
                    else:
                        st.error("❌ Failed to update booking")
            
            if st.button("❌ Cancel", use_container_width=True):
                st.session_state.editing_booking = None
                st.session_state.editing_calendar = None
                st.rerun()


def render_email_center(manager):
    """Render email management center"""
    st.markdown('<div class="section-header">📧 Email Center</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["⚙️ Configuration", "📤 Send Email", "📊 Email Log"])
    
    with tab1:
        st.markdown("### Email Configuration")
        
        st.markdown("""
        <div class="email-card">
            <h4>📋 SMTP Setup</h4>
            <p>Configure your email server settings to send booking reports and notifications.</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("email_config_form"):
            smtp_server = st.text_input(
                "SMTP Server",
                value="smtp.gmail.com",
                help="e.g., smtp.gmail.com for Gmail"
            )
            
            smtp_port = st.number_input(
                "SMTP Port",
                value=587,
                help="Usually 587 for TLS or 465 for SSL"
            )
            
            sender_email = st.text_input(
                "Sender Email",
                help="Your email address"
            )
            
            sender_password = st.text_input(
                "Password / App Password",
                type="password",
                help="For Gmail, use an App Password"
            )
            
            submitted = st.form_submit_button("💾 Save Configuration", type="primary")
            
            if submitted:
                if EmailManager.configure_email(smtp_server, smtp_port, sender_email, sender_password):
                    st.success("✅ Email configured successfully!")
                else:
                    st.error("❌ Configuration failed")
        
        if st.session_state.email_config.get('configured'):
            st.markdown(f"""
            <div class="success-box">
                <h4>✅ Email Configured</h4>
                <p><strong>Server:</strong> {st.session_state.email_config['smtp_server']}</p>
                <p><strong>From:</strong> {st.session_state.email_config['sender_email']}</p>
            </div>
            """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown("### Send Email")
        
        if not st.session_state.email_config.get('configured'):
            st.warning("⚠️ Please configure email settings first")
            return
        
        email_type = st.radio(
            "Email Type",
            ["📊 Booking Summary", "✉️ Custom Email"]
        )
        
        if email_type == "📊 Booking Summary":
            if st.session_state.current_workbook:
                workbook = manager.open_workbook(st.session_state.current_workbook)
                if workbook:
                    profile = manager.get_client_profile(workbook)
                    calendars = manager.get_calendar_sheets(workbook)
                    
                    selected_calendar = st.selectbox(
                        "Select Calendar",
                        [cal['name'] for cal in calendars]
                    )
                    
                    recipient = st.text_input("Recipient Email")
                    
                    if st.button("📧 Send Booking Summary", type="primary"):
                        selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
                        if selected_cal:
                            df = manager.read_calendar(selected_cal['sheet'])
                            if EmailManager.send_booking_summary(
                                recipient, 
                                profile.get('client_name', 'Unknown'), 
                                df
                            ):
                                st.success("✅ Email sent successfully!")
                                st.balloons()
                            else:
                                st.error("❌ Failed to send email")
            else:
                st.info("👈 Select a workbook first")
        
        else:  # Custom Email
            with st.form("custom_email_form"):
                recipient = st.text_input("To")
                subject = st.text_input("Subject")
                body = st.text_area("Message", height=200)
                
                attachment = st.file_uploader("Attachment (optional)")
                
                sent = st.form_submit_button("📤 Send Email", type="primary")
                
                if sent:
                    attachment_data = None
                    attachment_name = None
                    
                    if attachment:
                        attachment_data = attachment.read()
                        attachment_name = attachment.name
                    
                    if EmailManager.send_email(recipient, subject, body, attachment_data, attachment_name):
                        st.success("✅ Email sent!")
                    else:
                        st.error("❌ Failed to send")
    
    with tab3:
        st.markdown("### Email Log")
        
        if st.session_state.email_logs:
            df_emails = pd.DataFrame(st.session_state.email_logs)
            st.dataframe(df_emails, use_container_width=True)
            
            csv = df_emails.to_csv(index=False)
            st.download_button(
                "📥 Download Log",
                data=csv,
                file_name=f"email_log_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        else:
            st.info("No emails sent yet")

def render_search_analytics(manager, workbook):
    """Render enhanced search and analytics interface"""
    st.markdown('<div class="section-header">🔍 Search & Analytics</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendars found")
        return
    
    tab1, tab2, tab3 = st.tabs(["🔍 Advanced Search", "📊 Analytics Dashboard", "📈 Trends"])
    
    with tab1:
        selected_calendar = st.selectbox("Calendar", [cal['name'] for cal in calendars])
        selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
        
        if selected_cal:
            sheet = selected_cal['sheet']
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    search_term = st.text_input("🔍 Search", placeholder="Enter search term...")
                
                with col2:
                    search_columns = st.multiselect("Search in columns", df.columns.tolist())
                
                if st.button("🔍 Search", type="primary"):
                    if search_term:
                        result = manager.search_bookings(
                            sheet,
                            search_term,
                            search_columns if search_columns else None
                        )
                        
                        if not result.empty:
                            st.success(f"✅ Found {len(result)} matching records")
                            st.dataframe(result, use_container_width=True)
                            
                            csv = result.to_csv(index=False)
                            st.download_button(
                                "📥 Download Results",
                                data=csv,
                                file_name=f"search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                mime="text/csv"
                            )
                        else:
                            st.warning("No results found")
                    else:
                        st.warning("Please enter a search term")
            else:
                st.info("No data available")
    
    with tab2:
        st.markdown("### 📊 Analytics Dashboard")
        
        selected_calendar = st.selectbox("Select Calendar for Analytics", [cal['name'] for cal in calendars], key="analytics_calendar")
        selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
        
        if selected_cal:
            sheet = selected_cal['sheet']
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("📋 Total Bookings", len(df))
                
                with col2:
                    property_cols = [col for col in df.columns if 'property' in col.lower() or 'location' in col.lower()]
                    if property_cols:
                        unique_properties = df[property_cols[0]].nunique()
                        st.metric("🏘️ Properties", unique_properties)
                    else:
                        st.metric("🏘️ Properties", "N/A")
                
                with col3:
                    status_cols = [col for col in df.columns if 'status' in col.lower() or 'code' in col.lower()]
                    if status_cols:
                        status_counts = df[status_cols[0]].value_counts()
                        st.metric("📊 Status Types", len(status_counts))
                    else:
                        st.metric("📊 Status Types", "N/A")
                
                with col4:
                    st.metric("📅 Columns", len(df.columns))
                
                st.markdown("---")
                
                # Status distribution chart
                if status_cols:
                    st.markdown("#### Status Code Distribution")
                    status_df = pd.DataFrame(status_counts).reset_index()
                    status_df.columns = ['Status', 'Count']
                    
                    fig = px.pie(status_df, values='Count', names='Status',
                                title='Booking Status Distribution',
                                color_discrete_sequence=px.colors.qualitative.Pastel)
                    st.plotly_chart(fig, use_container_width=True)
                
                # Property distribution
                if property_cols:
                    st.markdown("#### Property Distribution")
                    property_counts = df[property_cols[0]].value_counts()
                    property_df = pd.DataFrame(property_counts).reset_index()
                    property_df.columns = ['Property', 'Count']
                    
                    fig = px.bar(property_df, x='Property', y='Count',
                                title='Bookings by Property',
                                color='Count',
                                color_continuous_scale='Viridis')
                    st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("---")
                st.markdown("#### Export Data")
                col1, col2 = st.columns(2)
                
                with col1:
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "📥 Download CSV",
                        data=csv,
                        file_name=f"{selected_calendar}_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
                
                with col2:
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name=selected_calendar[:31])
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        "📥 Download Excel",
                        data=excel_data,
                        file_name=f"{selected_calendar}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("No data available")
    
    with tab3:
        st.markdown("### 📈 Booking Trends")
        st.info("Trend analysis coming soon - will show booking patterns over time")

def render_templates(manager):
    """Render template management interface"""
    st.markdown('<div class="section-header">🎨 Booking Templates</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="template-card">
        <h3 style="margin-top: 0; color: white;">✨ Template Library</h3>
        <p style="color: white;">Save and reuse common booking patterns to speed up data entry</p>
    </div>
    """, unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["📚 Template Library", "➕ Create Template"])
    
    with tab1:
        if st.session_state.templates:
            st.markdown(f"### {len(st.session_state.templates)} Templates Available")
            
            for template in st.session_state.templates:
                with st.expander(f"🎨 {template['name']} (Used {template['usage_count']} times)"):
                    st.markdown(f"**Description:** {template['description']}")
                    st.markdown(f"**Created:** {template['created']}")
                    st.markdown(f"**ID:** {template['id']}")
                    
                    st.json(template['data'])
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"🗑️ Delete", key=f"del_{template['id']}"):
                            st.session_state.templates = [t for t in st.session_state.templates if t['id'] != template['id']]
                            st.success("✅ Template deleted")
                            st.rerun()
                    
                    with col2:
                        st.info("Use this template in the 'Add Booking' tab")
        else:
            st.info("No templates saved yet. Create one in the 'Create Template' tab!")
    
    with tab2:
        st.markdown("### Create New Template")
        
        with st.form("create_template_form"):
            template_name = st.text_input("Template Name", placeholder="e.g., Standard Check-In")
            template_desc = st.text_area("Description", placeholder="Describe when to use this template")
            
            st.markdown("#### Template Fields")
            st.info("Add the fields and default values for this template")
            
            num_fields = st.number_input("Number of Fields", min_value=1, max_value=20, value=5)
            
            template_data = {}
            cols = st.columns(2)
            
            for i in range(num_fields):
                with cols[i % 2]:
                    field_name = st.text_input("Field Name", key=f"field_name_{i}")
                    field_value = st.text_input("Field Value", key=f"field_value_{i}")
                    if field_name:
                        template_data[field_name] = field_value
            
            submitted = st.form_submit_button("💾 Save Template", type="primary")
            
            if submitted:
                if template_name and template_data:
                    save_template(template_name, template_desc, template_data)
                    st.success(f"✅ Template '{template_name}' saved!")
                    st.balloons()
                else:
                    st.error("❌ Please provide template name and at least one field")

def render_backup_restore(manager, workbook):
    """Render backup and restore interface"""
    st.markdown('<div class="section-header">💾 Backup & Restore</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-box">
        <h3 style="margin-top: 0;">🛡️ Data Protection</h3>
        <p>Create backups before making major changes and restore previous versions if needed.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Auto-backup toggle
    st.session_state.auto_backup = st.toggle("🔄 Auto-backup before edits", value=st.session_state.auto_backup)
    
    if st.session_state.auto_backup:
        st.success("✅ Auto-backup is enabled - backups will be created automatically before destructive operations")
    else:
        st.warning("⚠️ Auto-backup is disabled - you'll need to create backups manually")
    
    st.markdown("---")
    
    tab1, tab2 = st.tabs(["💾 Create Backup", "↩️ Restore Backup"])
    
    with tab1:
        st.markdown("### Create Manual Backup")
        
        calendars = manager.get_calendar_sheets(workbook)
        
        if calendars:
            selected_calendar = st.selectbox("Select Sheet to Backup", [cal['name'] for cal in calendars])
            
            if st.button("💾 Create Backup Now", type="primary", use_container_width=True):
                with st.spinner("Creating backup..."):
                    backup = create_backup(workbook, selected_calendar)
                    if backup:
                        st.success(f"✅ Backup created successfully!")
                        st.json({
                            'timestamp': backup['timestamp'],
                            'sheet': backup['sheet'],
                            'rows': backup['row_count'],
                            'columns': backup['col_count']
                        })
                    else:
                        st.error("❌ Backup failed")
        else:
            st.warning("No sheets available to backup")
    
    with tab2:
        st.markdown("### Restore from Backup")
        
        if st.session_state.backup_history:
            st.markdown(f"**{len(st.session_state.backup_history)} backups available**")
            
            for idx, backup in enumerate(reversed(st.session_state.backup_history)):
                with st.expander(f"💾 Backup #{len(st.session_state.backup_history) - idx} - {backup['timestamp']}"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown(f"**Sheet:** {backup['sheet']}")
                        st.markdown(f"**Workbook:** {backup['workbook']}")
                    
                    with col2:
                        st.markdown(f"**Rows:** {backup['row_count']}")
                        st.markdown(f"**Columns:** {backup['col_count']}")
                    
                    st.warning("⚠️ Restoring will overwrite current data!")
                    
                    if st.button(f"↩️ Restore This Backup", key=f"restore_{idx}"):
                        confirm = st.checkbox(f"I confirm restoration", key=f"confirm_{idx}")
                        if confirm:
                            with st.spinner("Restoring backup..."):
                                if restore_backup(backup, manager):
                                    st.success("✅ Backup restored successfully!")
                                    st.balloons()
                                else:
                                    st.error("❌ Restore failed")
        else:
            st.info("No backups available yet. Create your first backup in the 'Create Backup' tab!")
        
        st.markdown("---")
        
        if st.session_state.backup_history:
            if st.button("🗑️ Clear All Backups", type="secondary"):
                confirm_clear = st.checkbox("I confirm clearing all backups")
                if confirm_clear:
                    st.session_state.backup_history = []
                    st.success("✅ All backups cleared")
                    st.rerun()

def render_edit_history():
    """Render edit history log"""
    st.markdown('<div class="section-header">📜 Edit History</div>', unsafe_allow_html=True)
    
    if not st.session_state.edit_history:
        st.info("No edit history available")
        return
    
    # Filter options
    col1, col2, col3 = st.columns(3)
    
    with col1:
        filter_action = st.multiselect(
            "Filter by Action",
            list(set([e['action'] for e in st.session_state.edit_history])),
            default=list(set([e['action'] for e in st.session_state.edit_history]))
        )
    
    with col2:
        date_filter = st.date_input("Filter by Date", datetime.now().date())
    
    with col3:
        if st.button("🗑️ Clear History", type="secondary"):
            st.session_state.edit_history = []
            st.rerun()
    
    # Display filtered history
    filtered_history = [
        e for e in st.session_state.edit_history 
        if e['action'] in filter_action and e['timestamp'].startswith(date_filter.strftime('%Y-%m-%d'))
    ]
    
    st.markdown(f"### Showing {len(filtered_history)} of {len(st.session_state.edit_history)} edits")
    
    for idx, edit in enumerate(reversed(filtered_history)):
        with st.expander(f"Edit #{len(filtered_history) - idx} - {edit['timestamp']} - {edit['action']}"):
            st.markdown(f"**User:** {edit['user']}")
            st.markdown(f"**Action:** {edit['action']}")
            st.markdown(f"**Timestamp:** {edit['timestamp']}")
            st.markdown("**Details:**")
            st.json(edit['details'])
    
    # Export history
    if st.button("📥 Export Edit History"):
        df_history = pd.DataFrame(st.session_state.edit_history)
        csv = df_history.to_csv(index=False)
        st.download_button(
            "Download CSV",
            data=csv,
            file_name=f"edit_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

def render_system_logs():
    """Render system logs"""
    st.markdown('<div class="section-header">📊 System Logs</div>', unsafe_allow_html=True)
    
    if not st.session_state.logs:
        st.info("No system logs available")
        return
    
    # Filter options
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        filter_level = st.multiselect(
            "Filter by Level",
            ["INFO", "SUCCESS", "WARNING", "ERROR", "EMAIL"],
            default=["INFO", "SUCCESS", "WARNING", "ERROR", "EMAIL"]
        )
    
    with col2:
        log_limit = st.number_input("Show Last N Logs", min_value=10, max_value=500, value=100)
    
    with col3:
        if st.button("🗑️ Clear Logs"):
            st.session_state.logs = []
            st.rerun()
    
    # Display filtered logs
    filtered_logs = [log for log in st.session_state.logs if log['level'] in filter_level]
    displayed_logs = filtered_logs[-log_limit:]
    
    st.markdown(f"### Showing {len(displayed_logs)} of {len(st.session_state.logs)} logs")
    
    for log in reversed(displayed_logs):
        level_class = f"log-{log['level'].lower()}"
        st.markdown(
            f'<div class="log-entry {level_class}">'
            f'<strong>[{log["timestamp"]}]</strong> '
            f'<span style="color: #666;">[{log["level"]}]</span> '
            f'{log["message"]}'
            f'</div>',
            unsafe_allow_html=True
        )
    
    # Export logs
    st.markdown("---")
    if st.button("📥 Export System Logs"):
        df_logs = pd.DataFrame(st.session_state.logs)
        csv = df_logs.to_csv(index=False)
        st.download_button(
            "Download CSV",
            data=csv,
            file_name=f"system_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

def render_full_booking_creator(manager, workbook):
    """Comprehensive booking creation form"""
    st.markdown('<div class="section-header">➕ Create Full Booking</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    profile = manager.get_client_profile(workbook)
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    st.markdown("""
    <div class="feature-card">
        <h3 style="margin-top: 0;">📝 Complete Booking Form</h3>
        <p>Fill in all details to create a comprehensive booking entry</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Select target calendar
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar = st.selectbox("📅 Target Calendar", calendar_names)
    selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
    
    if not selected_cal:
        return
    
    # Booking form in columns
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📋 Basic Information")
        
        booking_date = st.date_input(
            "Booking Date",
            value=st.session_state.booking_form_data.get('date', datetime.now())
        )
        
        booking_time = st.time_input(
            "Booking Time",
            value=datetime.strptime(st.session_state.booking_form_data.get('time', '10:00'), '%H:%M').time() 
            if st.session_state.booking_form_data.get('time') else datetime.strptime('10:00', '%H:%M').time()
        )
        
        property_options = [prop['name'] for prop in profile.get('properties', [])]
        if not property_options:
            property_options = ['Property 1', 'Property 2', 'Property 3']
        
        property_name = st.selectbox(
            "Property/Location",
            property_options,
            index=property_options.index(st.session_state.booking_form_data.get('property', property_options[0])) 
            if st.session_state.booking_form_data.get('property') in property_options else 0
        )
        
        status_options = list(STATUS_CODES.keys())
        status_code = st.selectbox(
            "Status Code",
            status_options,
            format_func=lambda x: f"{x} - {STATUS_CODES[x]['name']}"
        )
        
        duration = st.number_input("Duration (hours)", min_value=0.5, max_value=24.0, value=2.0, step=0.5)
    
    with col2:
        st.markdown("#### 👤 Guest Information")
        
        guest_name = st.text_input("Guest Name", value=st.session_state.booking_form_data.get('guest_name', ''))
        guest_email = st.text_input("Guest Email", value=st.session_state.booking_form_data.get('email', ''))
        guest_phone = st.text_input("Guest Phone", value=st.session_state.booking_form_data.get('phone', ''))
        
        st.markdown("#### 💰 Pricing")
        price = st.number_input("Price/Amount", min_value=0.0, value=0.0, step=10.0)
        
        st.markdown("#### 📝 Additional Notes")
        notes = st.text_area("Notes/Comments", value=st.session_state.booking_form_data.get('notes', ''), height=100)
    
    # Preview booking data
    st.markdown("---")
    st.markdown("### 👁️ Booking Preview")
    
    booking_data = {
        'date': booking_date.strftime('%Y-%m-%d'),
        'time': booking_time.strftime('%H:%M'),
        'property': property_name,
        'status': status_code,
        'duration': str(duration),
        'guest_name': guest_name,
        'email': guest_email,
        'phone': guest_phone,
        'price': str(price),
        'notes': notes
    }
    
    preview_df = pd.DataFrame([booking_data])
    st.dataframe(preview_df, use_container_width=True)
    
    # Action buttons
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("✅ Create Booking", type="primary", use_container_width=True):
            if manager.create_full_booking(selected_cal['sheet'], booking_data):
                st.success("✅ Booking created successfully!")
                st.balloons()
                
                # Clear form
                st.session_state.booking_form_data = {}
                st.rerun()
            else:
                st.error("❌ Failed to create booking")
    
    with col2:
        if st.button("💾 Save as Template", use_container_width=True):
            template_name = st.text_input("Template Name", key="template_name_input")
            if template_name:
                save_template(template_name, f"Booking template for {property_name}", booking_data)
                st.success(f"✅ Template '{template_name}' saved!")
    
    with col3:
        if st.button("🔄 Reset Form", use_container_width=True):
            st.session_state.booking_form_data = {}
            st.rerun()
    
    with col4:
        if st.button("📧 Create & Email", use_container_width=True):
            if manager.create_full_booking(selected_cal['sheet'], booking_data):
                if guest_email:
                    # Send confirmation email
                    subject = f"Booking Confirmation - {property_name}"
                    body = f"""
                    <html>
                        <body>
                            <h2>Booking Confirmation</h2>
                            <p>Dear {guest_name},</p>
                            <p>Your booking has been confirmed:</p>
                            <ul>
                                <li><strong>Date:</strong> {booking_date.strftime('%B %d, %Y')}</li>
                                <li><strong>Time:</strong> {booking_time.strftime('%I:%M %p')}</li>
                                <li><strong>Property:</strong> {property_name}</li>
                                <li><strong>Duration:</strong> {duration} hours</li>
                                <li><strong>Service:</strong> {STATUS_CODES[status_code]['name']}</li>
                            </ul>
                            <p>Thank you for your booking!</p>
                        </body>
                    </html>
                    """
                    
                    if EmailManager.send_email(guest_email, subject, body):
                        st.success("✅ Booking created and confirmation email sent!")
                    else:
                        st.warning("⚠️ Booking created but email failed to send")
                else:
                    st.warning("⚠️ No email address provided")

def render_interactive_calendar_grid(manager, workbook):
    """Visual calendar grid with clickable dates"""
    st.markdown('<div class="section-header">📅 Interactive Calendar Grid</div>', unsafe_allow_html=True)
    
    # Month/Year selector
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col1:
        if st.button("◀️ Previous Month"):
            st.session_state.calendar_view_date -= timedelta(days=30) # Approximates previous month
            st.rerun()
    
    with col2:
        current_date = st.session_state.calendar_view_date
        st.markdown(f"""
        <div class="calendar-card">
            <h2 style="text-align: center; margin: 0; color: #000000;">
                {current_date.strftime('%B %Y')}
            </h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        if st.button("▶️ Next Month"):
            st.session_state.calendar_view_date += timedelta(days=30) # Approximates next month
            st.rerun()
    
    # Get bookings for current month
    start_date = st.session_state.calendar_view_date.replace(day=1)
    
    # Calculate end of month correctly
    if start_date.month == 12:
        end_date = start_date.replace(year=start_date.year + 1, month=1, day=1) - timedelta(days=1)
    else:
        end_date = start_date.replace(month=start_date.month + 1, day=1) - timedelta(days=1)

    # Fetch bookings from all calendars for the range
    all_bookings_for_range = pd.DataFrame()
    calendars = manager.get_calendar_sheets(workbook)
    if calendars:
        for cal in calendars:
            df = manager.read_calendar(cal['sheet'])
            if not df.empty:
                # Ensure date column exists and is in datetime format
                date_cols = [col for col in df.columns if 'date' in col.lower()]
                if date_cols:
                    date_col = date_cols[0]
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                    
                    # Filter by date range
                    mask = (df[date_col] >= pd.Timestamp(start_date)) & (df[date_col] <= pd.Timestamp(end_date))
                    all_bookings_for_range = pd.concat([all_bookings_for_range, df[mask]])

    # Create calendar grid
    year = st.session_state.calendar_view_date.year
    month = st.session_state.calendar_view_date.month
    
    # Get calendar matrix
    cal_matrix = cal_module.monthcalendar(year, month)
    
    # Display calendar
    st.markdown("### 📆 Calendar View")
    
    # Day headers
    day_cols = st.columns(7)
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    for i, day in enumerate(days):
        with day_cols[i]:
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #667eea;'>{day}</div>", 
                       unsafe_allow_html=True)
    
    # Calendar days
    for week in cal_matrix:
        week_cols = st.columns(7)
        for i, day in enumerate(week):
            with week_cols[i]:
                if day == 0:
                    st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)
                else:
                    # Count bookings for this day
                    day_date = datetime(year, month, day).date()
                    
                    day_bookings_count = 0
                    if not all_bookings_for_range.empty:
                        date_cols = [col for col in all_bookings_for_range.columns if 'date' in col.lower()]
                        if date_cols:
                            day_bookings_count = len(all_bookings_for_range[all_bookings_for_range[date_cols[0]].dt.date == day_date])
                    
                    # Color based on booking count
                    if day_bookings_count == 0:
                        bg_color = "#f8f9fa"
                        text_color = "#000000"
                    elif day_bookings_count <= 2:
                        bg_color = "#d4edda"
                        text_color = "#000000"
                    elif day_bookings_count <= 5:
                        bg_color = "#fff3cd"
                        text_color = "#000000"
                    else:
                        bg_color = "#f8d7da"
                        text_color = "#000000"
                    
                    st.markdown(f"""
                    <div style='
                        background: {bg_color};
                        border: 1px solid #dee2e6;
                        border-radius: 5px;
                        padding: 10px;
                        height: 80px;
                        text-align: center;
                    '>
                        <div style='font-size: 1.2rem; font-weight: bold; color: {text_color};'>{day}</div>
                        <div style='font-size: 0.8rem; color: {text_color}; margin-top: 5px;'>
                            {day_bookings_count} booking{'s' if day_bookings_count != 1 else ''}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Button to view bookings for this day
                    if st.button("View", key=f"day_{day}", use_container_width=True):
                        st.session_state.quick_filters['date'] = day_date
                        st.info(f"Showing bookings for {day_date}")
                        st.rerun() # Rerun to filter view below

    # Show bookings for selected date if filtered
    if 'date' in st.session_state.quick_filters:
        st.markdown("---")
        st.markdown(f"### 📋 Bookings for {st.session_state.quick_filters['date']}")
        
        filtered_date = st.session_state.quick_filters['date']
        if not all_bookings_for_range.empty:
            date_cols = [col for col in all_bookings_for_range.columns if 'date' in col.lower()]
            if date_cols:
                day_bookings_df = all_bookings_for_range[all_bookings_for_range[date_cols[0]].dt.date == filtered_date]
                
                if not day_bookings_df.empty:
                    st.dataframe(day_bookings_df, use_container_width=True)
        
        if st.button("Clear Filter"):
            del st.session_state.quick_filters['date']
            st.rerun()

def render_advanced_bulk_editor(manager, workbook):
    """Advanced bulk editing with grid interface"""
    st.markdown('<div class="section-header">✏️ Advanced Bulk Editor</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendar sheets found")
        return
    
    st.markdown("""
    <div class="bulk-edit-card">
        <h3 style="margin-top: 0;">🔧 Bulk Operations</h3>
        <p>Select multiple cells and perform batch updates efficiently</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sheet selector
    calendar_names = [cal['name'] for cal in calendars]
    selected_calendar = st.selectbox("Select Calendar Sheet", calendar_names)
    selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
    
    if not selected_cal:
        return
    
    sheet = selected_cal['sheet']
    df = manager.read_calendar(sheet)
    
    if df.empty:
        st.info("No data in this sheet")
        return
    
    # Display editable dataframe
    st.markdown("### 📊 Data Grid")
    st.dataframe(df, use_container_width=True, height=400)
    
    # Bulk edit options
    tab1, tab2, tab3, tab4 = st.tabs(["🔄 Find & Replace", "➕ Batch Add", "🗑️ Batch Delete", "📋 Multi-Cell Edit"])
    
    with tab1:
        st.markdown("#### Find and Replace")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            find_text = st.text_input("Find", key="find_text")
        
        with col2:
            replace_text = st.text_input("Replace with", key="replace_text")
        
        with col3:
            target_column = st.selectbox("In Column", ['All Columns'] + list(df.columns))
        
        if st.button("🔍 Preview Changes", type="secondary"):
            if find_text:
                preview_df = df.copy()
                if target_column == 'All Columns':
                    preview_df = preview_df.replace(find_text, replace_text, regex=False)
                else:
                    preview_df[target_column] = preview_df[target_column].replace(find_text, replace_text, regex=False)
                
                st.markdown("**Preview:**")
                st.dataframe(preview_df, use_container_width=True)
        
        if st.button("✅ Apply Find & Replace", type="primary"):
            if find_text:
                updates = []
                
                for row_idx, row in df.iterrows():
                    for col_idx, col_name in enumerate(df.columns):
                        if target_column == 'All Columns' or col_name == target_column:
                            cell_value = str(row[col_name])
                            if find_text in cell_value:
                                new_value = cell_value.replace(find_text, replace_text)
                                updates.append({
                                    'row': row_idx + 13,  # Adjust for header rows
                                    'col': col_idx + 1,
                                    'value': new_value
                                })
                
                if updates:
                    if manager.bulk_update_cells_optimized(sheet, updates):
                        st.success(f"✅ Updated {len(updates)} cells!")
                        st.rerun()
                else:
                    st.info("No matches found")
    
    with tab2:
        st.markdown("#### Batch Add Rows")
        
        num_rows = st.number_input("Number of rows to add", min_value=1, max_value=50, value=1)
        
        if st.button("➕ Add Empty Rows", type="primary"):
            for i in range(num_rows):
                empty_row = [''] * len(df.columns)
                manager.append_row(sheet, empty_row)
            
            st.success(f"✅ Added {num_rows} empty row(s)!")
            st.rerun()
    
    with tab3:
        st.markdown("#### Batch Delete Rows")
        
        st.warning("⚠️ This action cannot be undone. Create a backup first!")
        
        delete_start = st.number_input("Start Row (from row 13)", min_value=13, value=13)
        delete_end = st.number_input("End Row", min_value=13, value=13)
        
        if st.button("🗑️ Delete Rows", type="primary"):
            if delete_end >= delete_start:
                try:
                    # Create backup first
                    create_backup(workbook, sheet.title)
                    
                    # Delete rows (from end to start to maintain indices)
                    for row in range(delete_end, delete_start - 1, -1):
                        sheet.delete_rows(row)
                    
                    st.success(f"✅ Deleted rows {delete_start} to {delete_end}!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
    
    with tab4:
        st.markdown("#### Multi-Cell Edit")
        
        st.info("Enter cell updates in format: ROW,COL,VALUE (one per line)")
        
        bulk_text = st.text_area(
            "Bulk Edit Input",
            height=200,
            placeholder="Example:\n13,1,2024-01-15\n13,2,Property A\n14,1,2024-01-16",
            help="Format: row_number,column_number,new_value"
        )
        
        if st.button("📝 Preview Bulk Edit", type="secondary"):
            if bulk_text:
                lines = bulk_text.strip().split('\n')
                updates = []
                
                for line in lines:
                    parts = line.split(',')
                    if len(parts) >= 3:
                        try:
                            row = int(parts[0])
                            col = int(parts[1])
                            value = ','.join(parts[2:])  # Handle values with commas
                            updates.append({'row': row, 'col': col, 'value': value})
                        except:
                            st.warning(f"Invalid line: {line}")
                
                if updates:
                    st.success(f"✅ Parsed {len(updates)} updates")
                    preview_df = pd.DataFrame(updates)
                    st.dataframe(preview_df, use_container_width=True)
        
        if st.button("✅ Apply Bulk Edit", type="primary"):
            if bulk_text:
                lines = bulk_text.strip().split('\n')
                updates = []
                
                for line in lines:
                    parts = line.split(',')
                    if len(parts) >= 3:
                        try:
                            row = int(parts[0])
                            col = int(parts[1])
                            value = ','.join(parts[2:])
                            updates.append({'row': row, 'col': col, 'value': value})
                        except:
                            pass
                
                if updates:
                    if manager.bulk_update_cells_optimized(sheet, updates):
                        st.success(f"✅ Applied {len(updates)} updates!")
                        st.balloons()
                        st.rerun()

def main_app():
    """Main application interface"""
    manager = st.session_state.gc
    
    st.markdown('<div class="main-header">🏠 Booking Management Dashboard</div>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("🗂️ Workbook Selection")
        
        if st.session_state.notification_count > 0:
            st.markdown(f'<span class="notification-badge">{st.session_state.notification_count} new</span>', 
                       unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Refresh", use_container_width=True):
                with st.spinner("Refreshing..."):
                    st.session_state.workbooks = manager.list_workbooks_from_folder(DRIVE_FOLDER_ID)
                st.rerun()
        
        with col2:
            if st.button("🚪 Logout", use_container_width=True):
                st.session_state.authenticated = False
                st.session_state.gc = None
                st.session_state.workbooks = []
                st.rerun()
        
        st.markdown("---")
        
        if st.session_state.workbooks:
            workbook_names = [wb['name'] for wb in st.session_state.workbooks]
            selected_workbook_name = st.selectbox("Select Workbook", workbook_names)
            
            selected_workbook = next(
                (wb for wb in st.session_state.workbooks if wb['name'] == selected_workbook_name),
                None
            )
            
            if selected_workbook:
                if st.session_state.current_workbook != selected_workbook['id']:
                    st.session_state.current_workbook = selected_workbook['id']
                    st.rerun()
                
                st.markdown(f"""
                <div class="info-box">
                    <p><strong>📊 Active:</strong><br>{selected_workbook['name']}</p>
                </div>
                """, unsafe_allow_html=True)

    menu_options = {
        "🏠 Dashboard": render_dashboard,
        # "👥 All Clients": render_all_clients, # Placeholder for future features
        # "📊 All Sheets Viewer": render_all_sheets_viewer, # Placeholder for future features
        "📅 Extended Calendar": render_enhanced_calendar_view,  # New enhanced calendar
        "✏️ Booking Manager": render_live_editor, # Renamed to Live Editor
        "➕ Create Full Booking": render_full_booking_creator,
        "📅 Calendar View": render_calendar_view, # Basic Calendar View
        "⚙️ Email Center": render_email_center,
        "🔍 Search & Analytics": render_search_analytics,
        "🎨 Templates": render_templates,
        "💾 Backup & Restore": render_backup_restore,
        "📜 Edit History": render_edit_history,
        "📊 System Logs": render_system_logs,
        "🔧 Advanced Editor": render_advanced_bulk_editor, # Renamed Advanced Bulk Editor
        # "⚙️ Settings": render_settings # Placeholder for future features
    }
    
    view_mode = st.radio(
        "Select View",
        list(menu_options.keys()),
        format_func=lambda x: x, # Display full text from dictionary keys
        horizontal=True,
        key="main_view_radio"
    )
    
    st.markdown("---")
    
    # Render selected view
    if st.session_state.current_workbook:
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if workbook:
            render_function = menu_options.get(view_mode)
            if render_function:
                render_function(manager, workbook)
            else:
                st.error("Invalid view selected.")
        else:
            st.error("Could not open selected workbook.")
    else:
        st.info("👈 Please select a workbook from the sidebar to continue.")
    
    # Render edit booking modal if active
    render_edit_booking_modal(manager)


# Main entry point
if __name__ == "__main__":
    if not st.session_state.authenticated:
        authenticate()
    else:
        if st.session_state.gc:
            main_app()
        else:
            st.error("❌ Google Sheets client not initialized")
            if st.button("Return to Login"):
                st.session_state.authenticated = False
                st.rerun()
