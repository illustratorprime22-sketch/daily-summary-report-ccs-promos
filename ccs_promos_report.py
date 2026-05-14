import os
import json
import smtplib
import ssl
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials
from jinja2 import Template
import base64
from datetime import datetime, timedelta

# Configuration
SPREADSHEET_ID = '1lHua-KgYRwS2FesanfOxRhdogbN8izZOvMSYe_clo-4'
SMTP_HOST = 'secure.emailsrvr.com'
SMTP_PORT = 465
SENDER_EMAIL = 'mayur.kambli@artworkservicesusa.com'

# Recipients
RECIPIENTS_TO = ['balaji.alagesan@artworkservicesusa.com']
RECIPIENTS_CC = ['rupesh.pardeshi@artworkservicesusa.com']
RECIPIENTS_BCC = ['mayur.kambli@artworkservicesusa.com']

def get_google_client():
    creds_raw = os.environ.get('GOOGLE_CREDENTIALS_JSON', '').strip()
    if not creds_raw:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set")
    
    creds_dict = None
    try:
        decoded = base64.b64decode(creds_raw).decode('utf-8')
        if decoded.strip().startswith('{'):
            creds_dict = json.loads(decoded)
    except Exception:
        pass

    if not creds_dict:
        try:
            cleaned_raw = creds_raw.replace('\r\n', '\\n').replace('\n', '\\n')
            creds_dict = json.loads(cleaned_raw)
        except Exception:
            try:
                creds_dict = json.loads(creds_raw)
            except json.JSONDecodeError as e:
                raise ValueError(f"CRITICAL: GOOGLE_CREDENTIALS_JSON is malformed. Error: {e}")

    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def get_target_date(sh):
    # Try to get from D.count!Z2 like Natural Trends
    try:
        ws_count = sh.worksheet("D.count")
        date_val = ws_count.acell('Z2').value
        if date_val:
            print(f"Target Date from D.count!Z2: {date_val}")
            return date_val
    except:
        pass
    
    # Fallback to previous business day
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        target_date = today - timedelta(days=3)
    elif today.weekday() == 6:  # Sunday
        target_date = today - timedelta(days=2)
    else:
        target_date = today - timedelta(days=1)
    
    return f"{target_date.day}-{target_date.strftime('%b-%y')}"

def fetch_data(sh, target_date):
    print(f"Target Date: {target_date}")
    
    today = datetime.now()
    today_str = f"{today.day}-{today.strftime('%b-%y')}"
    
    # Get Data from CCS-Promos
    try:
        ws = sh.worksheet("CCS-Promos")
    except gspread.exceptions.WorksheetNotFound:
        try:
            ws = sh.worksheet("CCS Promos")
        except gspread.exceptions.WorksheetNotFound:
            print(f"Available sheets: {[w.title for w in sh.worksheets()]}")
            ws = sh.worksheets()[0]
            print(f"Fell back to using sheet: {ws.title}")
            
    all_data = ws.get_all_values()
    if not all_data:
        return 0, 0, 0, 0, []

    headers = all_data[0]
    def find_col(names, default):
        for i, h in enumerate(headers):
            if any(n.lower() in h.lower() for n in names):
                return i
        return default

    idx_date = find_col(["date"], 0)
    idx_from = find_col(["from", "name"], 5)
    idx_subject = find_col(["subject"], 6)
    idx_count = find_col(["count", "total item", "total virtual"], 9)
    idx_done = find_col(["done date"], 15)

    def normalize_date(d):
        d = d.strip()
        if not d: return ""
        parts = d.split('-')
        if len(parts) == 3:
            day = parts[0].lstrip('0')
            return f"{day}-{parts[1]}-{parts[2]}"
        return d

    target_date_norm = normalize_date(target_date)
    rows = all_data[1:]
    
    emails_received = 0
    emails_completed = 0
    total_completed_items = 0
    pending_count = 0
    detailed_rows = []

    for row in rows:
        max_idx = max(idx_date, idx_from, idx_subject, idx_count, idx_done)
        if len(row) <= max_idx:
            row = row + [""] * (max_idx + 1 - len(row))
        
        row_date_raw = str(row[idx_date]).strip()
        row_date = normalize_date(row_date_raw)
        
        email_from = str(row[idx_from]).strip()
        email_subject = str(row[idx_subject]).strip()
        total_items = str(row[idx_count]).strip()
        
        done_date_raw = str(row[idx_done]).strip()
        done_date = normalize_date(done_date_raw)
        
        # Skip if no date or if it matches today
        if not row_date or row_date == today_str:
            continue
            
        # --- VALIDATION: Skip header rows and empty placeholder rows ---
        if not email_subject or email_subject.lower() == 'email subject' or email_from.lower() == 'email from':
            continue
            
        is_pending = not done_date or done_date.lower() == 'pending'
        
        # 1. Emails Received
        if row_date == target_date_norm:
            emails_received += 1
            
        # 2. Emails Completed
        if done_date == target_date_norm:
            emails_completed += 1
            try:
                val = "".join(filter(str.isdigit, total_items))
                total_completed_items += int(val) if val else 0
            except:
                pass
        
        # 3. Pending
        if is_pending:
            pending_count += 1
            
        # 4. Detailed Rows
        if (row_date == target_date_norm or done_date == target_date_norm):
            detailed_rows.append([
                row_date_raw,
                email_from,
                email_subject,
                row[idx_count] if not is_pending else "",
                done_date_raw if done_date_raw.strip() else "Pending"
            ])

    return emails_received, emails_completed, total_completed_items, pending_count, detailed_rows

def format_html(target_date, emails_received, emails_completed, total_completed_items, pending_count, detailed_rows):
    if emails_received == 0 and emails_completed == 0 and pending_count == 0:
        template_str = """
        <html>
        <head>
        <style>
            body { font-family: Calibri, sans-serif; font-size: 10pt; line-height: 1.2; }
            table { border-collapse: collapse; border: 1px solid #000000; margin-top: 10px; }
            td { border: 1px solid #000000; padding: 2px 6px; font-size: 10pt; }
            .header-cell { font-weight: bold; }
        </style>
        </head>
        <body>
            <p>Hi team,</p>
            <p>Please see below summary.</p>
            <table>
                <tr><td class="header-cell">CCS Promos</td><td></td></tr>
                <tr><td class="header-cell">Date</td><td class="header-cell">Emails received</td></tr>
                <tr><td>{{ target_date }}</td><td>No orders received</td></tr>
            </table>
            <br>
            <p>Thanks and Regards,<br>Mayur</p>
        </body>
        </html>
        """
        template = Template(template_str)
        return template.render(target_date=target_date)

    template_str = """
    <html>
    <head>
    <style>
        body { font-family: Calibri, sans-serif; font-size: 10pt; line-height: 1.2; }
        table { border-collapse: collapse; border: 1px solid #000000; margin-top: 10px; width: auto; min-width: 400px; }
        td { border: 1px solid #000000; padding: 2px 6px; font-size: 10pt; }
        .header-cell { font-weight: bold; }
        .detail-table { width: 100%; max-width: 800px; }
        .section-title { font-weight: bold; text-align: center; }
    </style>
    </head>
    <body>
        <p>Hi team,</p>
        <p>Please see below mentioned summary report for your reference.</p>
        
        <table>
            <tr><td colspan="4" class="header-cell">CCS Promos</td></tr>
            <tr class="header-cell">
                <td>Date</td><td>Emails received</td><td>Emails completed</td><td>Total virtual completed</td>
            </tr>
            <tr>
                <td>{{ target_date }}</td><td>{{ emails_received }}</td><td>{{ emails_completed }}</td><td>{{ total_completed_items }}</td>
            </tr>
            <tr><td>&nbsp;</td><td></td><td></td><td></td></tr>
            <tr><td class="header-cell">Pending</td><td>{{ pending_count }}</td><td></td><td></td></tr>
        </table>
        <br>
        <table class="detail-table">
            <tr><td colspan="5" class="section-title">CCS Promos</td></tr>
            <tr class="header-cell">
                <td>Date</td><td>Emails from</td><td>Email subject</td><td>Count</td><td>Done date</td>
            </tr>
            {% for row in detailed_rows %}
            <tr>
                <td>{{ row[0] }}</td><td>{{ row[1] }}</td><td>{{ row[2] }}</td><td>{{ row[3] }}</td><td>{{ row[4] }}</td>
            </tr>
            {% endfor %}
        </table>
        <br>
        <p>Thanks and Regards,<br>Mayur</p>
    </body>
    </html>
    """
    template = Template(template_str)
    return template.render(
        target_date=target_date,
        emails_received=emails_received,
        emails_completed=emails_completed,
        total_completed_items=total_completed_items,
        pending_count=pending_count,
        detailed_rows=detailed_rows
    )

def send_email(subject, html_content):
    password = os.environ.get('SMTP_PASSWORD')
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECIPIENTS_TO)
    msg['Cc'] = ", ".join(RECIPIENTS_CC)
    msg['Bcc'] = ", ".join(RECIPIENTS_BCC)
    msg.set_content("Please enable HTML to view this report.")
    msg.add_alternative(html_content, subtype='html')
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(SENDER_EMAIL, password)
        server.send_message(msg)

def main():
    try:
        gc = get_google_client()
        sh = gc.open_by_key(SPREADSHEET_ID)
        target_date = get_target_date(sh)
        emails_received, emails_completed, total_items, pending, detailed = fetch_data(sh, target_date)
        html_content = format_html(target_date, emails_received, emails_completed, total_items, pending, detailed)
        subject = f"CCS Promos Summary: {target_date}"
        print(f"Sending email: {subject}")
        send_email(subject, html_content)
        print("Success!")
    except Exception as e:
        print(f"Error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
