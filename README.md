This program for fetching emails from an IMAP server, saving attachments, and sending notification emails for unauthorized senders. Let's break down the main components of the code:

   Class: FetchEmail (This class is designed to handle email fetching operations. Here's an overview of its methods):
     
      a. __init__: Initializes the class by establishing an IMAP connection with SSL, logging in, and selecting the inbox.
      b. close_connection: Closes the connection to the IMAP server.
      c. save_attachment: Saves attachments from a given email message to a specified download folder.
      d. fetch_unread_messages: Fetches unread messages from the server, filters them based on a list of allowed email addresses, and marks them as read.
      e. parse_email: Parses email data and returns an EmailMessage object.
      f. has_attachment: Checks if an email has any attachments.
      g. move_to_downloaded_folder: Moves unseen emails from a specified email address to a "Downloaded" folder and marks them as deleted in the inbox.
      h. send_notification: Sends a notification email for unauthorized senders.

   Main Part of the Script:
     
      a. Initialization:
          Sets up the IMAP server details, username, and password.
          Specifies download folders for attachments.
          Creates an instance of the FetchEmail class.
      b. Email Processing:
          Moves unseen emails from a specified email address to a "Downloaded" folder and marks them as deleted in the inbox.
      c. Fetching Unread Emails:
          Retrieves unread messages from the server, filters them based on a list of allowed email addresses, and marks them as read.
          If unauthorized emails are found, sends a notification email.
      d. Processing Unread Emails:
            Iterates through each unread email.
            Checks for attachments and prints information about the sender, subject, and content type.
      e. Saving Attachments:
            Saves attachments from specific senders and with specific subjects to specified download folders.
      f. Closing Connection:
            Closes the connection to the IMAP server.

   Note:
   
      The script is designed to work with Outlook's IMAP server, so ensure your email provider supports IMAP.
      This script heavily relies on the email and smtplib modules for email handling and notification sending.

   Clone the repository:
   
       git clone https://github.com/your-username/your-repository.git
       cd your-repository

   Install dependencies & used tools:
   
        Python 3.x 
        Spyder
        pip install imaplib==2.58
        pip install email==0.6.2
 
   Update configuration:
   
        a. mail_server: IMAP server details
        b. username: Your Outlook email address
        c. password: Your Outlook email password
        d. download_folder1 and download_folder2: Paths to download folders for attachments
        e. email_id: Target email ID for downloading attachments
        f. allowed_emails: List of allowed email addresses
