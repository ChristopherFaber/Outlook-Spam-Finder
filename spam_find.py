import imaplib
import email
from email.header import decode_header 

username = "cmcfaber.eng@outlook.com"
password = "Chris2faXT2_Diamond$"

mail = imaplib.IMAP4_SSL("imap.outlook.com")

mail.login(username, password)

mail.select("inbox")

status, messages = mail.search(None, "ALL")

email_ids = messages[0].split()

for email_id in email_ids:

    latest_email_id = email_id

    status, msg_data = mail.fetch(latest_email_id, "(RFC822)")

    raw_email = msg_data[0][1]

    msg = email.message_from_bytes(raw_email)

# Decode the email subject
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
    # If it's a bytes type, decode to str
        subject = subject.decode(encoding if encoding else "utf-8")

# Print email subject


    if msg.is_multipart():
    # If the email message is multipart
        for part in msg.walk():
        # Each part is a MIME part
        # Check if the content type is text/plain or text/html
            if part.get_content_type() == "text/plain":
                try:
                    body = part.get_payload(decode=True).decode()
                    if "unsubscribe" in body:
                        print("From:", msg.get("From"))
                except UnicodeDecodeError:
                    pass
    else:
    # If the email message isn't multipart
        try:
            body = part.get_payload(decode=True).decode()
            if "unsubscribe" in body:
                print("From:", msg.get("From"))
        except UnicodeDecodeError:
                pass

mail.close()
mail.logout()
