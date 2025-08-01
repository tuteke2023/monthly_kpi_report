import requests
import json
from datetime import datetime

# SMTP2GO API Configuration
api_key = "api-A90FFF85F5A343D288733CBE74F87141"
api_url = "https://api.smtp2go.com/v3/email/send"

# Test recipients
test_recipients = ["teke@ttaccountancy.com.au", "sydelle@ttaccountancy.com.au"]

# Email content
subject = "Test Email - Monthly KPI App"
html_body = f"""<html>
<body>
<p>Hello,</p>

<p>This is a test email to confirm that the email system is working correctly for the Monthly KPI Report application.</p>

<p>The email integration with SMTP2GO has been successfully configured and is ready to send automated performance reports.</p>

<p><strong>Test Details:</strong></p>
<ul>
<li>API Integration: SMTP2GO</li>
<li>Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</li>
</ul>

<p>If you receive this email, it means the email system is functioning properly.</p>

<p>Best regards,<br>
Monthly KPI Report System</p>
</body>
</html>"""

text_body = f"""Hello,

This is a test email to confirm that the email system is working correctly for the Monthly KPI Report application.

The email integration with SMTP2GO has been successfully configured and is ready to send automated performance reports.

Test Details:
- API Integration: SMTP2GO
- Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

If you receive this email, it means the email system is functioning properly.

Best regards,
Monthly KPI Report System
"""

def send_test_email_api():
    headers = {
        "accept": "application/json",
        "Content-Type": "application/json",
        "X-Smtp2go-Api-Key": api_key
    }
    
    for recipient in test_recipients:
        email_data = {
            "api_key": api_key,
            "to": [recipient],
            "sender": "Monthly KPI Report <noreply@ttaccountancy.com.au>",
            "subject": subject,
            "text_body": text_body,
            "html_body": html_body
        }
        
        try:
            print(f"Sending email to {recipient} via SMTP2GO API...")
            response = requests.post(api_url, headers=headers, data=json.dumps(email_data))
            
            if response.status_code == 200:
                result = response.json()
                if result.get("data", {}).get("succeeded", 0) > 0:
                    print(f"✓ Test email successfully sent to {recipient}")
                else:
                    print(f"✗ Failed to send email to {recipient}: {result}")
            else:
                print(f"✗ Failed to send email to {recipient}: HTTP {response.status_code} - {response.text}")
                
        except Exception as e:
            print(f"✗ Failed to send email to {recipient}: {str(e)}")

if __name__ == "__main__":
    print("Starting email test using SMTP2GO API...")
    send_test_email_api()
    print("\nEmail test completed.")