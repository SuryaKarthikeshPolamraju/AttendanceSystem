import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import schedule
import time

def send_email():
    # Email configuration
    sender_email = "sathviknandyala012@gmail.com"  # Replace with your email
    receiver_email = "pskarthikesh@gmail.com"  # Replace with recipient's email
    subject = "Scheduled Email with Attachment"
    body = "Hello, This is the body of the email."

    # Attach the file
    attachment_path = "data.xlsx"  # Replace with your file path
    attachment_name = "data.xlsx"

    # Create the email message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Attach the body
    message.attach(MIMEText(body, "plain"))

    # Attach the file
    with open(attachment_path, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name=attachment_name)
        part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
        message.attach(part)

    # Connect to the SMTP server (Gmail example)
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()

        # Use the App Password generated from your Google Account
        app_password = "zdrv yzup vmmz ufgy"  # Replace with your app password
        server.login(sender_email, app_password)

        server.sendmail(sender_email, receiver_email, message.as_string())

    print(f"Email sent at {time.strftime('%Y-%m-%d %H:%M:%S')}")

# Schedule the email to be sent every day at 14:30 (2:30 PM)
schedule.every().day.at("15:51").do(send_email)

# Run the scheduler
while True:
    schedule.run_pending()
    time.sleep(1)
