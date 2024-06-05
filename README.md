# Mass Mail Sender

This script sends personalized emails to a list of recipients. It reads recipient information from an Excel file and uses an SMTP server to send the emails.

## Prerequisites

Before you begin, ensure you have met the following requirements:
- You have Python 3.x installed on your machine.
- You have the `pandas` library installed. You can install it using `pip install pandas`.
- You have the `smtplib` and `email` libraries, which are included in the Python standard library.
- You have an Excel file named `recipients.xlsx` in the same directory as the script. This file should contain two columns: the first column for email addresses and the second column for recipient names.

## Installation

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/lanryweezy/mass-mail-sender.git
