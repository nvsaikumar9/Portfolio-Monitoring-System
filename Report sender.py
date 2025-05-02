import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import schedule
import time
import pandas as pd

with open('Report.txt', 'r') as file:
    file_data = file.read()

for i in file_data.split('\n'):
    if i.startswith('Email'):
        email = i.split(':')[1].strip()
        break

print(email)

# Function to send email
def send_email():
    sender_email = 'n.v.saikumar9@gmail.com'
    sender_password = "upuuozxqebztquel"
    recipient_email = email

    subject = "Daily Report on Portfolio"
    body = file_data

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        # Connect to the SMTP server and send the email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Schedule the email to be sent every 2 minutes
schedule.every(2).minutes.do(send_email)

print("Scheduler is running...")

# Keep the script running
while True:
    schedule.run_pending()
    time.sleep(1)