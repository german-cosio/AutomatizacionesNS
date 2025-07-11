import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv

load_dotenv()

def send_email(recipients, subject, body, attachments=None):
    """
    Send an email with specified subject and body to the given recipients and optionally attach files.

    :param recipients: List of recipient email addresses.
    :param subject: Subject line of the email.
    :param body: Body text of the email.
    :param attachments: Optional list of file paths to attach to the email. Defaults to None.
    """
    email_address = os.getenv('EMAIL_ADDRESS')
    email_password = os.getenv('EMAIL_PASSWORD')
    smtp_server = os.getenv('SMTP_SERVER')
    smtp_port = int(os.getenv('SMTP_PORT'))

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = email_address
    message['To'] = ', '.join(recipients)
    message['Subject'] = subject

    # Attach the body with the msg instance
    message.attach(MIMEText(body, 'plain'))

    # Attach files to the email if any
    if attachments:
        for file_path in attachments:
            part = MIMEBase('application', "octet-stream")
            try:
                with open(file_path, 'rb') as file:
                    part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(file_path))
                message.attach(part)
            except IOError:
                print(f"Error: Could not find file {file_path} or read data.")

    # Create SMTP session for sending the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.connect(smtp_server, smtp_port)  # Explicitly connect, normally not necessary
        server.ehlo()  # Identify ourselves to the smtp server
        server.starttls()  # enable security
        server.ehlo()  # Re-identify ourselves as an encrypted connection
        server.login(email_address, email_password)  # login with mail_id and password
        text = message.as_string()
        server.sendmail(email_address, recipients, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")
