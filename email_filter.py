import email
import smtplib
import imaplib
import os
import re
from email.parser import BytesParser

class FetchEmail():
    connection = None
    error = None
    def __init__(self, mail_server, username, password):
        self.connection = imaplib.IMAP4_SSL(mail_server)
        self.connection.login(username, password)
        self.connection.select(readonly=False) # so we can mark mails as read

    def close_connection(self):
        """
        Close the connection to the IMAP server
        """
        self.connection.close()

    def save_attachment(self, msg, download_folder):
        """
        Given a message, save its attachments to the specified
        download folder (default is /tmp)
        return: file path to attachment
        """
        att_path = "No attachment found."
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            filename = part.get_filename()
            att_path = os.path.join(download_folder, filename)
            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        return att_path
    
    def fetch_unread_messages(self, allowed_emails):
        """
        Retrieve unread messages and filter by allowed_emails list
        """
        emails = []
        uemails = []
        (result, messages) = self.connection.search(None, 'UnSeen')
        if result == "OK":
            for message in messages[0].decode().split(' '):
                try: 
                    ret, data = self.connection.fetch(message,'(RFC822)')
                except:
                    print("No new emails to read.")
                    self.close_connection()
                    exit()
                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    # Extract email address from the 'From' field of the email
                    email_address = msg['From'].split()[-1].strip('<>')
                    # Check if the email address is allowed
                    if email_address in allowed_emails:
                        emails.append(msg)
                    else:
                        print(f"Unauthorized email - {email_address}.")
                        uemails.append(msg)
                        continue
                response, data = self.connection.store(message, '+FLAGS','\\Seen')
            z=[]
            for i in uemails:
                z.append(i['FROM'][(i['FROM'].find("<") + 1):(i['FROM'].find(">"))])
            if len(z)>0:
                self.send_notification(z)
            return emails
        self.error = "Failed to retrieve emails."
        return emails

    def parse_email(self, msg_data):
        """
        Parse email data and return an EmailMessage object
        """
        msg_bytes = msg_data[0][1]
        msg = BytesParser().parsebytes(msg_bytes)
        return msg

    def has_attachment(self, msg):
        """
        Check if an email has any attachments
        """
        if 'Content-Disposition' in msg and 'attachment' in msg['Content-Disposition']:
            return True
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is not None:
                return True
        return False

    def move_to_downloaded_folder(self, from_email):
        """
        Move unseen emails from the specified email address to a "Downloaded" folder
        and mark them as deleted in the inbox
        """
        pattern_uid = r'\(UID (?P<uid>\d+)\)'
        downloaded_folder = "Downloaded"  # Update folder name as needed
        try:
            # Check if the downloaded folder exists
            typ, data = self.connection.list()
            folders = [folder.decode().split(' "/" ')[1] for folder in data]
            if downloaded_folder not in folders:
                # Create the downloaded folder if it doesn't exist
                self.connection.create(downloaded_folder)

            # Search for unseen emails from the specified email address
            typ, data = self.connection.search(None, f'(UNSEEN FROM "{from_email}")')
            if typ == 'OK':
                msg_ids = data[0].split()
                for msg_id in msg_ids:
                    # Fetch the email details
                    typ, msg_data = self.connection.fetch(msg_id, '(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])')
                    if typ == 'OK':
                        msg = self.parse_email(msg_data)
                        if self.has_attachment(msg):
                            # Move the email to the downloaded folder
                            self.connection.uid('COPY', msg_id, downloaded_folder)
                            if self.connection.uid('COPY', msg_id, downloaded_folder)[0] == 'OK':
                                mov, data = self.connection.uid('STORE', msg_id, '+FLAGS', '\\Deleted')
                                self.connection.expunge()
                            print(f"Unseen email from '{from_email}' with an attachment moved to '{downloaded_folder}' folder.")
                        else:
                            print(f"Unseen email from '{from_email}' does not have an attachment.")
                    else:
                        print("Failed to fetch email details.")
            else:
                print(f"No unseen emails found from '{from_email}'.")
        except Exception as e:
            print("Failed to move email:", str(e))

    def send_notification(self, blacklist_mails):
        """
        Sends a notification email to the specified receiver email
        """
        blacklist_mails_str = [str(msg) for msg in blacklist_mails]
        subject = "Unauthorized Email Notification"
        body = "You have received an email from an unauthorized senders: \n" + '\n'.join(blacklist_mails_str) 
        message = f"Subject: {subject}\n\n{body}"
        try:
            with smtplib.SMTP('smtp.office365.com', 587) as server:
                server.starttls()
                server.login("nnani5961@outlook.com", "Dru@#1234")
                server.sendmail("nnani5961@outlook.com", "nnani5961@outlook.com", message)
                print("Notification email sent successfully.")
        except Exception as e:
            print("Failed to send notification email:", str(e))


# initialize the FetchEmail object
mail_server = "imap-mail.outlook.com"
username = "please provide a outlook mail"
password = "your password"
output = "C:/Users/druthik.goud/Desktop/mail_ret/output"
invoices = "C:/Users/druthik.goud/Desktop/mail_ret/Invoice"
fetch_email = FetchEmail(mail_server, username, password)
has_attachment = False


email_id = "l.e.druthikgoud@gmail.com"  # specify the email id from which you want to download attachment from
allowed_emails = ["nnani5961@gmail.com", "l.e.druthikgoud@gmail.com"] #to check any unread mails from these mails

# fetch all unread email messages from the server
fetch_email.move_to_downloaded_folder(email_id)
emails = fetch_email.fetch_unread_messages(allowed_emails)

# iterate through each email message
print('____________________')
print("ALL UNSEEN MAIL INFO")
print('--------------------')
for email in emails:
    for part in email.walk():
        if part.get_filename():
            has_attachment = True
    # check if the email id matches the specified email id
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
    print(f"{email['From']}  {email['Subject']}  {part.get_content_type()}")
    print('--------------------------------------------------------------')


for email in emails:
    #print(has_attachment)
    if email_id in email['From'] and has_attachment == True and email['Subject'] in ['p2p','P2P','a2p','A2P']:
        fetch_email.save_attachment(email, download_folder1)
    if email_id in email['From'] and has_attachment == True and email['Subject'] in ['invoice','Invoice','INVOICE']:
        fetch_email.save_attachment(email, download_folder2)


# close the connection to the IMAP server
fetch_email.close_connection()








