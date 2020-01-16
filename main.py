# ===Python Script to Download Attachments From Mail Automatically ===

# Import the required packages (* mandatory input fields)
"""
1. **imaplib**
2. **base64**
3. **os**
4. **email**
5. **argparse**
"""
import imaplib
import base64
import os
import email
import argparse

# ==== Note: ====

"""
1. **-d**       - Path To Downloaded File *
2. **-m**       - User MailId *
3. **-p**       - Mail Password *
4. **-f**       - File Name 
5. **-imap**    - Imap Server Name *
6. **-port**    - Imap Server Port *
"""

# ==== main function ====
def main(args):
    #  Provide MailId and Password from which need to login
    email_user = args.mail_id

    email_passwd = args.password
    #  Imap Server and Port
    mail = imaplib.IMAP4_SSL(args.imap_server, args.imap_port)
    #  login with user details
    mail.login(email_user, email_passwd)
    """
    Select the Folder Name.

    For example I  used **Indox** here.

    The parameter  **readonly** is optional if it is assign to  **True** it make the **unread** unread else **read**.
    """
    mail.select("Inbox", readonly=True)
    """
    Search for particular **Subject** and **From Id** in the mail list .

    **Unseen** is an optional Keyword if we provide it  will also search in  ***unseen*** mail list.

    Here , I searched for **From Id** - Rupesh with **subject**  - SQL.
    """
    type, data = mail.search(
        None, '(FROM "Rupesh" SUBJECT "SQL")', 'UnSeen')
    """
    Reverse the mail list from latest to older .

    Incase you need all mail remove break statement.
    """
    # Fetch Mail details for last sent mail only
    for num in data[0].split()[::-1]:
        typ, data = mail.fetch(num, '(RFC822)')
        raw_email = data[0][1]
        break

    # Decode the Mail String to readable Data Format
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
    # Download the file from mail
    for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            #  Get the filename as same as mail attachment (Change if u Needed)
            fileName = part.get_filename()
            # In case of renaming the file provide the fileName (in arguments).
            if args.filename:
                fileName = args.filename
            if bool(fileName):
                # Specify the Download location Path
                filePath = os.path.join(
                    args.target_dir, fileName)
                #  Save the file in Specified Location
                if not os.path.isfile(filePath):
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Download Attachments From Mail')
    parser.add_argument('-d', '--target_dir', required=True, help='path_to_save')
    parser.add_argument('-m', '--mail_id', required=True, help='mailId')
    parser.add_argument('-p', '--password', required=True, help='mail password')
    parser.add_argument('-f', '--filename', help='FileName for file')
    parser.add_argument('-imap','--imap_server',default="imap.gmail.com", help='Imap Server Name')
    parser.add_argument('-port','--imap_port',default="993", help='Imap Server Port')
    args = parser.parse_args()
    # defined in ***line 29*** (top)
    main(args)
