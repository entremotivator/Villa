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

# Status code definitions
STATUS_CODES = {
    'CI': {'name': 'Check-In', 'description': 'Complete cleaning and preparation for incoming guests', 'color': '#4CAF50'},
    'SO': {'name': 'Stay-over', 'description': 'Mid-stay cleaning with linen and towel refresh', 'color': '#2196F3'},
    'CO/CI': {'name': 'Check-out/Check-in', 'description': 'Same-day turnover between guests', 'color': '#FF9800'},
    'FU': {'name': 'Fresh-up', 'description': 'Quick refresh for already clean property', 'color': '#9C27B0'},
    'DC': {'name': 'Deep Cleaning', 'description': 'Thorough deep clean of entire property', 'color': '#F44336'},
    'COC': {'name': 'Construction Cleaning', 'description': 'Post-renovation construction cleanup', 'color': '#795548'},
    'CO': {'name': 'Check-Out', 'description': 'Final cleaning after guest departure', 'color': '#607D8B'}
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
        """Get all calendar sheets (excluding first sheet)"""
        try:
            worksheets = workbook.worksheets()
            calendars = []
            
            for i, sheet in enumerate(worksheets[1:], start=1):
                calendars.append({
                    'index': i,
                    'name': sheet.title,
                    'sheet': sheet
                })
            
            return calendars
        except Exception as e:
            add_log(f"Error listing calendar sheets: {str(e)}", "ERROR")
            return []
    
    def read_calendar(self, sheet, start_row: int = 13) -> pd.DataFrame:
        """Read booking calendar starting from specified row"""
        try:
            all_values = sheet.get_all_values()
            
            if len(all_values) < start_row:
                return pd.DataFrame()
            
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
            
            data = all_values[start_row - 1:]
            df = pd.DataFrame(data, columns=unique_headers)
            df = df[df.apply(lambda row: row.astype(str).str.strip().any(), axis=1)]
            
            return df
        except Exception as e:
            add_log(f"Error reading calendar: {str(e)}", "ERROR")
            return pd.DataFrame()
    
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
            
            for update in updates:
                sheet.update_cell(update['row'], update['col'], update['value'])
            
            add_edit_history("Batch Update", {
                'sheet': sheet.title,
                'updates_count': len(updates),
                'cells': updates
            })
            
            add_log(f"✅ Batch update completed: {len(updates)} cells", "SUCCESS")
            return True
        except Exception as e:
            add_log(f"Error in batch update: {str(e)}", "ERROR")
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
                return df
            
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

    view_mode = st.radio(
        "📑 View Mode",
        ["Dashboard", "Live Editor", "Email Center", "Search & Analytics", "Edit History", "System Logs"],
        help="Choose what to display"
    )
    
    st.markdown("---")
    
    if view_mode == "Dashboard":
        if not st.session_state.current_workbook:
            st.info("👈 Please select a workbook from the sidebar")
            return
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if workbook:
            render_dashboard(manager, workbook)
    
    elif view_mode == "Live Editor":
        if not st.session_state.current_workbook:
            st.info("👈 Please select a workbook from the sidebar")
            return
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if workbook:
            render_live_editor(manager, workbook)
    
    elif view_mode == "Email Center":
        render_email_center(manager)
    
    elif view_mode == "Search & Analytics":
        if not st.session_state.current_workbook:
            st.info("👈 Please select a workbook from the sidebar")
            return
        workbook = manager.open_workbook(st.session_state.current_workbook)
        if workbook:
            render_search_analytics(manager, workbook)
    
    elif view_mode == "Edit History":
        render_edit_history()
    
    elif view_mode == "System Logs":
        render_system_logs()

def render_dashboard(manager, workbook):
    """Render enhanced dashboard view"""
    st.markdown('<div class="section-header">📊 Dashboard Overview</div>', unsafe_allow_html=True)
    
    profile = manager.get_client_profile(workbook)
    calendars = manager.get_calendar_sheets(workbook)
    
    # Client header
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        st.markdown(f"""
        <div class="success-box">
            <h2 style="margin-top: 0;">👤 {profile.get('client_name', 'Unknown Client')}</h2>
            <p><strong>Check-out:</strong> {profile.get('check_out_time', 'N/A')} | 
               <strong>Check-in:</strong> {profile.get('check_in_time', 'N/A')}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.metric("🏘️ Properties", len(profile.get('properties', [])))
    
    with col3:
        st.metric("📅 Calendars", len(calendars))
    
    # Quick stats
    total_bookings = 0
    upcoming_bookings = 0
    today = datetime.now().date()
    
    for cal in calendars:
        df = manager.read_calendar(cal['sheet'])
        if not df.empty:
            total_bookings += len(df)
            
            for col in df.columns:
                if 'date' in col.lower():
                    dates = pd.to_datetime(df[col], errors='coerce')
                    upcoming_bookings += (dates >= pd.Timestamp(today)).sum()
    
    col1, col2, col3, col4 = st.columns(4)
    
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
    
    # Properties section
    if profile.get('properties'):
        st.markdown("### 🏘️ Properties")
        prop_cols = st.columns(min(len(profile['properties']), 3))
        
        for idx, prop in enumerate(profile['properties']):
            with prop_cols[idx % 3]:
                st.markdown(f"""
                <div class="property-card">
                    <h4 style="color: #667eea;">🏠 {prop['name']}</h4>
                    <p><strong>📍</strong> {prop['address']}</p>
                    <p><strong>⏱️</strong> {prop['hours']} | <strong>🔄</strong> {prop['so_hours']}</p>
                </div>
                """, unsafe_allow_html=True)
    
    # Recent activity
    st.markdown("### 📊 Recent Activity")
    
    if st.session_state.edit_history:
        recent_edits = st.session_state.edit_history[-10:]
        
        for edit in reversed(recent_edits):
            st.markdown(f"""
            <div class="log-entry log-success">
                <strong>[{edit['timestamp']}]</strong> {edit['action']} by {edit['user'][:30]}...
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No recent activity")

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
        
        tab1, tab2, tab3, tab4 = st.tabs(["📝 Edit Cells", "➕ Add Booking", "🗑️ Delete Row", "📋 Bulk Operations"])
        
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
                
                with st.form("add_booking_form"):
                    new_row = []
                    cols = st.columns(2)
                    
                    for i, col_name in enumerate(df.columns):
                        with cols[i % 2]:
                            value = st.text_input(f"{col_name}", key=f"add_{i}")
                            new_row.append(value)
                    
                    submitted = st.form_submit_button("➕ Add Booking", type="primary")
                    
                    if submitted:
                        if manager.append_row(sheet, new_row):
                            st.success("✅ Booking added!")
                            st.rerun()
                        else:
                            st.error("❌ Failed to add")
        
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
                        if manager.batch_update_cells(sheet, updates):
                            st.success(f"✅ Updated {len(updates)} cells!")
                            st.rerun()
                    except json.JSONDecodeError:
                        st.error("❌ Invalid JSON format")

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
    """Render search and analytics interface"""
    st.markdown('<div class="section-header">🔍 Search & Analytics</div>', unsafe_allow_html=True)
    
    calendars = manager.get_calendar_sheets(workbook)
    
    if not calendars:
        st.warning("No calendars found")
        return
    
    tab1, tab2 = st.tabs(["🔍 Advanced Search", "📊 Analytics"])
    
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
                        result = manager.search_bookings(                            df,
                            search_term,
                            search_columns
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
                            st.warning("No results found matching your search criteria.")
                    else:
                        st.warning("Please enter a search term.")
            else:
                st.info("No data available in this calendar.")
    
    with tab2:
        st.markdown('<div class="analytics-card"><h3>📊 Analytics Dashboard</h3></div>', unsafe_allow_html=True)
        
        selected_calendar = st.selectbox("Select Calendar for Analytics", [cal['name'] for cal in calendars], key="analytics_calendar")
        selected_cal = next((cal for cal in calendars if cal['name'] == selected_calendar), None)
        
        if selected_cal:
            sheet = selected_cal['sheet']
            df = manager.read_calendar(sheet)
            
            if not df.empty:
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown(f'<div class="metric-card"><h2>{len(df)}</h2><p>Total Bookings</p></div>', unsafe_allow_html=True)
                
                with col2:
                    property_cols = [col for col in df.columns if 'property' in col.lower() or 'location' in col.lower()]
                    if property_cols:
                        unique_properties = df[property_cols[0]].nunique()
                        st.markdown(f'<div class="metric-card"><h2>{unique_properties}</h2><p>Properties</p></div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="metric-card"><h2>N/A</h2><p>Properties</p></div>', unsafe_allow_html=True)
                
                with col3:
                    status_cols = [col for col in df.columns if 'status' in col.lower() or 'code' in col.lower()]
                    if status_cols:
                        status_counts = df[status_cols[0]].value_counts()
                        st.markdown(f'<div class="metric-card"><h2>{len(status_counts)}</h2><p>Status Types</p></div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="metric-card"><h2>N/A</h2><p>Status Types</p></div>', unsafe_allow_html=True)
                
                st.markdown("---")
                
                if status_cols:
                    st.markdown("#### Status Code Distribution")
                    status_df = pd.DataFrame(status_counts).reset_index()
                    status_df.columns = ['Status', 'Count']
                    st.dataframe(status_df, use_container_width=True)
                
                st.markdown("#### Recent Bookings")
                st.dataframe(df.head(10), use_container_width=True)
                
                st.markdown("---")
                st.markdown("#### Export Data")
                col1, col2 = st.columns(2)
                
                with col1:
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "📥 Download Full Calendar (CSV)",
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
                        "📥 Download Full Calendar (Excel)",
                        data=excel_data,
                        file_name=f"{selected_calendar}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("No data available for analytics.")

def render_settings():
    """Render settings and configuration interface"""
    st.markdown('<div class="section-header">⚙️ System Settings</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["📧 Email Configuration", "📊 System Logs", "📝 Edit History"])
    
    with tab1:
        st.markdown('<div class="email-card"><h3>Email Server Configuration</h3></div>', unsafe_allow_html=True)
        
        with st.form("email_config_form"):
            st.markdown("Configure your SMTP email server settings to enable email functionality.")
            
            col1, col2 = st.columns(2)
            
            with col1:
                email = st.text_input(
                    "Email Address",
                    value=st.session_state.email_config.get('email', ''),
                    placeholder="your-email@example.com"
                )
                server = st.text_input(
                    "SMTP Server",
                    value=st.session_state.email_config.get('server', ''),
                    placeholder="smtp.gmail.com"
                )
            
            with col2:
                password = st.text_input(
                    "Password",
                    type="password",
                    value=st.session_state.email_config.get('password', ''),
                    placeholder="Your email password"
                )
                port = st.number_input(
                    "SMTP Port",
                    min_value=1,
                    max_value=65535,
                    value=st.session_state.email_config.get('port', 587)
                )
            
            submitted = st.form_submit_button("💾 Save Configuration", type="primary")
            
            if submitted:
                EmailManager.configure(email, password, server, port)
                st.success("✅ Email configuration saved successfully!")
                st.balloons()
        
        st.markdown("---")
        st.markdown("#### Common SMTP Settings")
        st.markdown("""
        <div class="info-box">
        <p><strong>Gmail:</strong> smtp.gmail.com (Port 587)<br>
        <strong>Outlook:</strong> smtp-mail.outlook.com (Port 587)<br>
        <strong>Yahoo:</strong> smtp.mail.yahoo.com (Port 587)<br>
        <strong>Note:</strong> You may need to enable "Less secure app access" or use App Passwords for Gmail.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown("### System Activity Log")
        
        if st.session_state.logs:
            col1, col2 = st.columns([3, 1])
            
            with col1:
                filter_level = st.multiselect(
                    "Filter by level",
                    ["INFO", "SUCCESS", "WARNING", "ERROR"],
                    default=["INFO", "SUCCESS", "WARNING", "ERROR"]
                )
            
            with col2:
                if st.button("🗑️ Clear Logs"):
                    st.session_state.logs = []
                    st.rerun()
            
            filtered_logs = [log for log in st.session_state.logs if log['level'] in filter_level]
            
            for log in filtered_logs:
                level_class = f"log-{log['level'].lower()}"
                st.markdown(
                    f'<div class="log-entry {level_class}">'
                    f'<strong>[{log["timestamp"]}]</strong> '
                    f'<span style="color: #666;">[{log["level"]}]</span> '
                    f'{log["message"]}'
                    f'</div>',
                    unsafe_allow_html=True
                )
            
            st.markdown("---")
            df_logs = pd.DataFrame(st.session_state.logs)
            csv = df_logs.to_csv(index=False)
            st.download_button(
                "📥 Download System Logs",
                data=csv,
                file_name=f"system_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            st.info("No system logs available.")
    
    with tab3:
        st.markdown("### Edit History")
        
        if st.session_state.edit_history:
            col1, col2 = st.columns([3, 1])
            
            with col2:
                if st.button("🗑️ Clear History"):
                    st.session_state.edit_history = []
                    st.rerun()
            
            for edit in st.session_state.edit_history:
                st.markdown(
                    f'<div class="edit-panel">'
                    f'<strong>{edit["action"]}</strong> at {edit["timestamp"]}<br>'
                    f'<pre>{json.dumps(edit["details"], indent=2)}</pre>'
                    f'</div>',
                    unsafe_allow_html=True
                )
            
            st.markdown("---")
            df_edits = pd.DataFrame(st.session_state.edit_history)
            csv = df_edits.to_csv(index=False)
            st.download_button(
                "📥 Download Edit History",
                data=csv,
                file_name=f"edit_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            st.info("No edit history available.")

if __name__ == "__main__":
    if not st.session_state.authenticated:
        authenticate()
    else:
        if st.session_state.gc:
            main_app()
        else:
            st.error("❌ Google Sheets connection not established. Please re-authenticate.")
            if st.button("Return to Login"):
                st.session_state.authenticated = False
                st.rerun()
