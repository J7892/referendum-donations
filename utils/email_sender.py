import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(subject, body_html, body_text):
    smtp_server = os.environ.get("SMTP_SERVER")
    smtp_port = os.environ.get("SMTP_PORT", "587")
    smtp_user = os.environ.get("SMTP_USERNAME")
    smtp_pass = os.environ.get("SMTP_PASSWORD")
    email_to = os.environ.get("EMAIL_TO")
    email_from = os.environ.get("EMAIL_FROM", smtp_user)

    if not all([smtp_server, smtp_user, smtp_pass, email_to]):
        print("\n=== SMTP CONFIGURATION MISSING OR INCOMPLETE ===")
        print("Set SMTP_SERVER, SMTP_PORT, SMTP_USERNAME, SMTP_PASSWORD, and EMAIL_TO as environment variables.")
        print(f"SMTP Server: {smtp_server}")
        print(f"SMTP User: {smtp_user}")
        print(f"Email To: {email_to}")
        print("\n--- Dry Run: Email Details ---")
        print(f"Subject: {subject}")
        print(f"From: {email_from}")
        print(f"To: {email_to}")
        print("\nPlain Text Body:")
        print(body_text)
        print("================================================\n")
        return False

    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = email_from
        msg['To'] = email_to

        msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
        msg.attach(MIMEText(body_html, 'html', 'utf-8'))

        port = int(smtp_port)
        if port == 465:
            # SSL Connection
            print(f"Connecting to SMTP Server {smtp_server}:{port} via SSL...")
            server = smtplib.SMTP_SSL(smtp_server, port, timeout=30)
        else:
            # TLS Connection (port 587 or 25)
            print(f"Connecting to SMTP Server {smtp_server}:{port} via TLS...")
            server = smtplib.SMTP(smtp_server, port, timeout=30)
            server.ehlo()
            server.starttls()
            server.ehlo()
        
        print("Logging in to SMTP server...")
        server.login(smtp_user, smtp_pass)
        print("Sending email...")
        server.sendmail(email_from, [email_to], msg.as_string())
        server.quit()
        print("Email sent successfully!")
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False
