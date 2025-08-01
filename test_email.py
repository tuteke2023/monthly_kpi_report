import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# SMTP2GO Configuration
smtp_server = "mail.smtp2go.com"
smtp_port = 2525
smtp_username = "ttaccountancy.com.au"
smtp_password = "A90FFF85F5A343D288733CBE74F87141"
sender_email = "noreply@ttaccountancy.com.au"  # Update if needed

# Test recipients
test_recipients = ["teke@ttaccountancy.com.au", "sydelle@ttaccountancy.com.au"]

# Email content
subject = "Test Email - Monthly KPI App"
body = f"""Hello,

This is a test email to confirm that the email system is working correctly for the Monthly KPI Report application.

The email integration with SMTP2GO has been successfully configured and is ready to send automated performance reports.

Test Details:
- Server: {smtp_server}
- Port: {smtp_port}
- Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

If you receive this email, it means the email system is functioning properly.

Best regards,
Monthly KPI Report System
"""

def send_test_email():
    for recipient in test_recipients:
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient
            msg['Subject'] = subject
            
            # Attach body
            msg.attach(MIMEText(body, 'plain'))
            
            # Connect to SMTP2GO server and send
            print(f"Connecting to SMTP2GO server...")
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            print(f"Logging in with username: {smtp_username}")
            server.login(smtp_username, smtp_password)
            
            print(f"Sending email to {recipient}...")
            server.send_message(msg)
            server.quit()
            
            print(f"✓ Test email successfully sent to {recipient}")
            
        except Exception as e:
            print(f"✗ Failed to send email to {recipient}: {str(e)}")

if __name__ == "__main__":
    print("Starting email test...")
    send_test_email()
    print("\nEmail test completed.")