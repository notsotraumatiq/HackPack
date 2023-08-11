import pandas as pd
import imaplib
import email

# Script to get email addresses from Gmail
# Gmail IMAP settings
IMAP_SERVER = 'imap.gmail.com'
IMAP_PORT = 993

# Your Gmail credentials
USERNAME = 'atiqpatel81@gmail.com'
PASSWORD = ''

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

# Clean the email addresses
lookup = set()

for email in email_addresses:
    # get the data between < and >
    # data = email[0].split('<')[1].split('>')[0]
    # add to set
    if email not in lookup:

        try:
            #  I want to clean the email addresses so that I only get the email address and not the name sample
            # Ryan Au-Yeung <ryan@expledge.com>
            # use regex to get the data between < and >
            if email is not None:
                print( type(email))
                email = email.strip()
                email = email.split('<')[1].split('>')[0]
            # email = email.split('<')[1].split('>')[0]

        except IndexError:
            pass
        lookup.add(email)


# convert back to list
email_addresses = list(lookup)
print(email_addresses)

df = pd.DataFrame({'Email Addresses': email_addresses})
df.to_excel('sent_emails.xlsx', index=False)

print('Email addresses extracted and saved to sent_emails.xlsx')
