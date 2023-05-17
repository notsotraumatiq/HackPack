import pandas as pd
import imaplib
import email

# Script to get email addresses from Gmail
# Gmail IMAP settings
IMAP_SERVER = 'imap.gmail.com'
IMAP_PORT = 993

# Your Gmail credentials
USERNAME = 'EMAIL'
PASSWORD = 'PASSWORD'

# Connect to the Gmail IMAP server
mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(USERNAME, PASSWORD)

# Select the folder
mail.select('"[Gmail]/Sent Mail"')

# Search for all sent emails
_, message_numbers = mail.search(None, 'ALL')
message_numbers = message_numbers[0].split()

email_addresses = []

# Iterate over the sent emails
for num in message_numbers:
    _, msg_data = mail.fetch(num, '(RFC822)')
    msg = email.message_from_bytes(msg_data[0][1])

    # Extract the sender's email address
    sender_email = msg['To']
    email_addresses.append(sender_email)

# Close the connection to the IMAP server
mail.logout()

# Save the email addresses to an Excel sheet
df = pd.DataFrame({'Email Addresses': email_addresses})
df.to_excel('sent_emails.xlsx', index=False)

print('Email addresses extracted and saved to sent_emails.xlsx')
