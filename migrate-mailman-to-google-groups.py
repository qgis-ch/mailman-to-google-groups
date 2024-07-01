#!/usr/bin/env python3

# A quick hack on the "Google Apps Groups Migration API Quickstart" program
#
# Giving a Google Group, walk through the legacy Mailman mbox file
# and extract only the text/plain part
#
# Many scripts walk through the mbox and just forward to the google group
# however, that doesn't preserve the original time stamp
#
# multiple runs on the mbox are idempotent since Messages-ID is constant
#
# At some point I may loop back to figure out how to do multi-part preservation
# passes since some e-mails had PDFs, M$ docs, etc.

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

from io import StringIO
import mailbox
from email.mime.text import MIMEText

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/groupsmigration-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/apps.groups.migration'
CLIENT_SECRET_FILE = 'credentials.json'
APPLICATION_NAME = 'Google Apps Groups Migration API Python Quickstart'


def get_credentials():
    """Shows basic usage of the Docs API.

    Returns:
        Credentials, the obtained credential.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)
    
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds

def main():
    """Shows basic usage of the Google Admin-SDK Groups Migration API.

    Creates a Google Admin-SDK Groups Migration API service object and
    inserts a test email into a group.
    """
    credentials = get_credentials()

    try:
        service = build('groupsmigration', 'v1', credentials=credentials)

        groupId = input(
            'Enter the email address of a Google Group in your domain: ')

        mbox_in = input('Name of mbox to migrate: ')

        src_mbox = mailbox.mbox(mbox_in)
        msg_count = 1
        msg_total = len(src_mbox)

        for msg in src_mbox:
            if msg.is_multipart():
                found = 0
                for part in msg.walk():
                    if part.get_content_type() == 'text/plain':
                        found = 1
                        message = MIMEText(part.get_payload(decode=True), 'plain', 'utf-8')
                if not found:
                    print('Error! No text/plain part found!')
                    continue
            else:
                message = MIMEText(msg.get_payload(decode=True), 'plain', 'utf-8')

            # Reformat the sender field
            from_raw = msg.get_from().split(" ")
            from_formatted = "%s@%s" % (from_raw[0], from_raw[2])

            message['Message-ID'] = msg['Message-ID']
            message['Subject'] = msg['Subject']
            message['From'] = from_formatted
            message['To'] = groupId
            message['Date'] = msg['Date']

            stream = StringIO()
            stream.write(message.as_string())
            media = MediaIoBaseUpload(stream,
                                    mimetype='message/rfc822')

            result = service.archive().insert(groupId=groupId,
                                            media_body=media).execute()

            if result['responseCode'] != 'SUCCESS':
                print('Issue with Message # ', msg_count, ' Message-ID: ', msg['Message-ID'], 'Reponse Code: ', result['responseCode'])
            else:
                print('Working on msg: ', msg_count, ' out of ', msg_total)

            msg_count += 1
    except HttpError as err:
        print(err)
        return 1

    return 0

if __name__ == '__main__':
    main()
