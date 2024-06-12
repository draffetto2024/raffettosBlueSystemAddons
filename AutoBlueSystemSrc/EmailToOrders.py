#The purpose of this code will be entirely to extract the orders from the emails
import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup

# Email details
username = "gingoso2@gmail.com"
password = "soiz avjw bdtu hmtn" #utilizing an app password

# Connect to the server
mail = imaplib.IMAP4_SSL("imap.gmail.com")

# Login to your account
mail.login(username, password)

# Select the mailbox you want to use
mail.select("inbox")

# Search for all emails in the inbox
status, messages = mail.search(None, "ALL")

# Convert messages to a list of email IDs
messages = messages[0].split()

# Fetch the latest email
latest_email_id = messages[-1]

# Fetch the email by ID
status, msg_data = mail.fetch(latest_email_id, "(RFC822)")

# Parse the email content
for response_part in msg_data:
    if isinstance(response_part, tuple):
        msg = email.message_from_bytes(response_part[1])
        subject = decode_header(msg["Subject"])[0][0]
        if isinstance(subject, bytes):
            subject = subject.decode()

        print("Subject:", subject)

        # Email sender
        from_ = msg.get("From")
        print("From:", from_)

        # Email content
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                try:
                    body = part.get_payload(decode=True).decode()
                except:
                    pass

                if content_type == "text/plain" and "attachment" not in content_disposition:
                    print("Body:", body)
                elif "text/html" in content_type:
                    soup = BeautifulSoup(part.get_payload(decode=True).decode(), "html.parser")
                    print("Body (HTML):", soup.get_text())
        else:
            content_type = msg.get_content_type()
            body = msg.get_payload(decode=True).decode()
            if content_type == "text/plain":
                print("Body:", body)
            elif "text/html" in content_type:
                soup = BeautifulSoup(body, "html.parser")
                print("Body (HTML):", soup.get_text())

# Close the connection and logout
mail.close()
mail.logout()
