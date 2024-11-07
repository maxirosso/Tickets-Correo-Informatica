import imaplib
import email
from email.header import decode_header
from pymongo import MongoClient
from datetime import datetime
import time
import logging
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Retrieve sensitive information from environment variables
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")
IMAP_PORT = int(os.getenv("IMAP_PORT"))
MONGO_URI = os.getenv("MONGO_URI")

# MongoDB setup
client = MongoClient(MONGO_URI)
db = client["TicketsMail"]
tickets_collection = db["tickets"]

# Keywords to filter for queries in Spanish
KEYWORDS = ["consulta", "pregunta", "duda", "ayuda", "información", "asistencia", "requerimiento"]

def connect_to_email():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL, PASSWORD)
        logging.info("Connected to the email server successfully.")
        return mail
    except Exception as e:
        logging.error(f"Error connecting to email server: {e}")
        return None

def fetch_latest_email(mail):
    try:
        mail.select("inbox")
        status, messages = mail.search(None, 'UNSEEN')
        email_ids = messages[0].split()

        if email_ids:
            latest_email_id = email_ids[-1]  # Get the most recent unread email
            return latest_email_id
        else:
            logging.info("No new unread emails.")
            return None
    except Exception as e:
        logging.error(f"Error fetching emails: {e}")
        return None

def process_email(email_id, mail):
    try:
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])

                # Decode the email's subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

                # Get the email's sender
                sender = msg.get("From")

                # Get email content
                body = get_email_content(msg)

                # Check if any keyword is in the subject or body
                subject_lower = subject.lower()
                body_lower = body.lower()
                if any(keyword in subject_lower or keyword in body_lower for keyword in KEYWORDS):
                    ticket = {
                        "subject": subject,
                        "sender": sender,
                        "date": datetime.now(),
                        "content": body
                    }
                    save_ticket(ticket)
                else:
                    logging.info(f"No keywords found in email: {subject}")

    except Exception as e:
        logging.error(f"Error processing email {email_id}: {e}")

def get_email_content(msg):
    try:
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if content_type == "text/plain" and "attachment" not in content_disposition:
                    return part.get_payload(decode=True).decode('utf-8', errors='replace')
        else:
            return msg.get_payload(decode=True).decode('utf-8', errors='replace')
    except Exception as e:
        logging.error(f"Error decoding email content: {e}")
        return ""

def save_ticket(ticket):
    try:
        tickets_collection.insert_one(ticket)
        logging.info(f"Ticket saved: {ticket['subject']}")
    except Exception as e:
        logging.error(f"Error saving ticket: {e}")

def main():
    while True:
        mail = connect_to_email()
        if mail:
            email_id = fetch_latest_email(mail)
            if email_id:
                process_email(email_id, mail)
            mail.logout()
        else:
            logging.error("Skipping email check due to connection failure.")
        
        logging.info("Waiting 1 minute before checking again...")
        time.sleep(60)

if __name__ == "__main__":
    main()
