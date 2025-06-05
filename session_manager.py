import os
import io
import pickle
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime

# Directory to store session data
SESSION_DIR = "sessions"

# Create the sessions directory if it doesn't exist
if not os.path.exists(SESSION_DIR):
    os.makedirs(SESSION_DIR)

def save_upload(department, file_data, sheet_name, schedule_col, amount_col):
    """
    Save a department's upload to the session storage.
    
    Args:
        department (str): "claims" or "finance"
        file_data (BytesIO): The uploaded Excel file
        sheet_name (str): The selected sheet name
        schedule_col (str): Column name for schedule numbers
        amount_col (str): Column name for amounts
    """
    # Create unique session ID based on current week
    session_id = datetime.datetime.now().strftime("%Y-W%U")  # Year and week number
    session_file = os.path.join(SESSION_DIR, f"session_{session_id}.pkl")
    
    # Initialize or load session data
    if os.path.exists(session_file):
        with open(session_file, 'rb') as f:
            session_data = pickle.load(f)
    else:
        session_data = {
            'claims': None,
            'finance': None
        }
    
    # Create a copy of the file data to avoid memory issues
    file_copy = io.BytesIO()
    file_data.seek(0)
    file_copy.write(file_data.read())
    file_copy.seek(0)
    
    # Save department data (preserving other department's data if it exists)
    session_data[department] = {
        'file_data': file_copy,
        'sheet_name': sheet_name,
        'schedule_col': schedule_col,
        'amount_col': amount_col,
        'timestamp': datetime.datetime.now().isoformat()
    }
    
    # Save session data
    with open(session_file, 'wb') as f:
        pickle.dump(session_data, f)
        
    # Send email notification
    send_notification_email(department)
    
    # Check if both departments have uploaded
    both_uploaded = session_data['claims'] is not None and session_data['finance'] is not None
    if both_uploaded:
        # Send a special notification when both departments have uploaded
        send_notification_email('both')
    
    return session_id

def get_session_data(session_id=None):
    """
    Retrieve the session data for the current week or a specific session ID.
    
    Args:
        session_id (str, optional): Specific session ID to load. 
                                   If None, uses current week.
    
    Returns:
        dict: Session data with claims and finance uploads
    """
    if session_id is None:
        # Use current week as session ID if not specified
        session_id = datetime.datetime.now().strftime("%Y-W%U")
    
    session_file = os.path.join(SESSION_DIR, f"session_{session_id}.pkl")
    
    if not os.path.exists(session_file):
        return {'claims': None, 'finance': None}
    
    with open(session_file, 'rb') as f:
        session_data = pickle.load(f)
    
    return session_data

def get_available_sessions():
    """
    Get a list of all available sessions.
    
    Returns:
        list: List of session IDs (week numbers)
    """
    sessions = []
    if os.path.exists(SESSION_DIR):
        for file in os.listdir(SESSION_DIR):
            if file.startswith("session_") and file.endswith(".pkl"):
                session_id = file.replace("session_", "").replace(".pkl", "")
                sessions.append(session_id)
    
    # Sort sessions by date (newest first)
    sessions.sort(reverse=True)
    return sessions

def send_notification_email(department):
    """
    Send an email notification when a department uploads their file.
    
    Args:
        department (str): The department that uploaded the file
    """
    sender_email = os.getenv("POWERBI_SENDER_EMAIL")
    recipient_email = "ifeoluwa.adeniyi@avonhealthcare.com"
    password = os.getenv("POWERBI_PASSWORD")
    
    # Check if credentials are available
    if not sender_email or not password:
        print("❌ PowerBI credentials not found in environment variables")
        print("Please set POWERBI_SENDER_EMAIL and POWERBI_PASSWORD in Secrets")
        return False
    
    # Create message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient_email
    
    if department == 'both':
        message["Subject"] = f"Claims Reconciliation Tool: ALL FILES READY FOR RECONCILIATION"
        body = f"""
        <html>
        <body>
            <p>This is an automated notification from the Claims Reconciliation Tool.</p>
            <p><b>IMPORTANT: Both Claims and Finance departments have uploaded their files!</b></p>
            <p>You can now proceed with the full reconciliation process by visiting the Claims Reconciliation Tool.</p>
            <p>This is an automated message, please do not reply.</p>
        </body>
        </html>
        """
    else:
        message["Subject"] = f"Claims Reconciliation Tool: {department.capitalize()} Department Upload"
        # Email body for individual department uploads
        body = f"""
        <html>
        <body>
            <p>This is an automated notification from the Claims Reconciliation Tool.</p>
            <p>The <b>{department.capitalize()} Department</b> has uploaded their file at {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.</p>
            <p>Please note:</p>
            <ul>
                <li>{'Both Claims and Finance uploads are now available.' if department.lower() != 'claims' else 'Waiting for Finance Department upload.'}</li>
                <li>{'Both Claims and Finance uploads are now available.' if department.lower() != 'finance' else 'Waiting for Claims Department upload.'}</li>
            </ul>
            <p>You can now proceed with the reconciliation process by visiting the Claims Reconciliation Tool.</p>
            <p>This is an automated message, please do not reply.</p>
        </body>
        </html>
        """
    
    message.attach(MIMEText(body, "html"))
    
    try:
        # Create SMTP session
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()  # Secure the connection
        server.login(sender_email, password)
        server.send_message(message)
        server.quit()
        print(f"Email notification sent for {department} upload")
        return True
    except Exception as e:
        print(f"Failed to send email: {str(e)}")
        return False