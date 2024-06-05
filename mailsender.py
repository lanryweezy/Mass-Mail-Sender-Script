import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

# Email configuration
smtp_server = 'your_smtp_server'
smtp_port = 25
sender_email = 'your_email@example.com'
sender_password = None  # Update if a password is required
subject = 'Your Subject Here'

# Read recipients' email addresses and names from an Excel file
recipients_df = pd.read_excel('recipients.xlsx')
recipients = recipients_df.values.tolist()  # Convert DataFrame to list of lists

# Connect to the SMTP server
server = smtplib.SMTP(smtp_server, smtp_port)

# Iterate over recipients and send personalized emails
for recipient in recipients:
    recipient_email = recipient[0]
    recipient_name = recipient[1]

    # Create a personalized message for each recipient
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = subject
    personalized_message = f"Dear {recipient_name},\n\nYour personalized message here.\n\nBest regards,\n[Your Name]"

    # Attach the personalized message to the email
    message.attach(MIMEText(personalized_message, 'plain'))

    # Send the email
    server.sendmail(sender_email, recipient_email, message.as_string())

# Close the connection to the SMTP server
server.quit()
