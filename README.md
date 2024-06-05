# Mass-Mail-Sender-Script
This script sends personalized emails to a list of recipients. It reads recipient information from an Excel file and uses an SMTP server to send the emails.
Prerequisites
Before you begin, ensure you have met the following requirements:

You have Python 3.x installed on your machine.
You have the pandas library installed. You can install it using pip install pandas.
You have the smtplib and email libraries, which are included in the Python standard library.
You have an Excel file named recipients.xlsx in the same directory as the script. This file should contain two columns: the first column for email addresses and the second column for recipient names.
Installation
Clone this repository to your local machine:

bash
Copy code
git clone https://github.com/yourusername/mass-mail-sender.git
Navigate to the project directory:

bash
Copy code
cd mass-mail-sender
Install the required packages:

bash
Copy code
pip install pandas
Usage
Update the script with your SMTP server details and sender email:

python
Copy code
smtp_server = 'your_smtp_server'  # Update with your SMTP server address
smtp_port = 25  # Update with your SMTP server port
sender_email = 'your_email@example.com'  # Update with your sender email address
sender_password = None  # Update if a password is required
subject = 'Your Subject Here'  # Update the email subject if needed
Ensure your recipients.xlsx file is correctly formatted with email addresses in the first column and names in the second column.

Customize the personalized message in the script:

python
Copy code
personalized_message = f"Dear {recipient_name},\n\nYour personalized message here.\n\nBest regards,\n[Your Name]"
Run the script:

bash
Copy code
python send_mass_mail.py
Editing the Code
To make improvements or changes to the script:

Fork this repository.
Create a new branch:
bash
Copy code
git checkout -b feature/your-feature-name
Make your changes and commit them:
bash
Copy code
git commit -m 'Add some feature'
Push to the branch:
bash
Copy code
git push origin feature/your-feature-name
Open a pull request and describe your changes.
Explanation of the Code
The script performs the following steps:

Imports the necessary libraries:

python
Copy code
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
Configures the email settings:

python
Copy code
smtp_server = 'your_smtp_server'
smtp_port = 25
sender_email = 'your_email@example.com'
sender_password = None  # Update if a password is required
subject = 'Your Subject Here'
Reads the recipients' email addresses and names from an Excel file:

python
Copy code
recipients_df = pd.read_excel('recipients.xlsx')
recipients = recipients_df.values.tolist()  # Convert DataFrame to list of lists
Connects to the SMTP server:

python
Copy code
server = smtplib.SMTP(smtp_server, smtp_port)
Iterates over the recipients and sends personalized emails:

python
Copy code
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
Closes the connection to the SMTP server:

python
Copy code
server.quit()
Contributing
Contributions are welcome! Please follow the steps in the "Editing the Code" section to contribute to this project.

License
This project is open-source and available under the MIT License.

