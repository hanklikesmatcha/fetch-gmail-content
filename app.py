from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import re
from openpyxl import Workbook, load_workbook

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/gmail/v1/users/userId/messages']
wb = Workbook()
ws1 = wb.active
file_name = "gift-codes.xlsx"
ws1.title = 'Gift Codes'


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API
    sent_emails = service.users().messages().list(userId='me', q="to='hank@sharesies.co.nz'").execute()

    if len(sent_emails['messages']) < 1:
        print('No sender found')

    for index, email in enumerate(sent_emails['messages']):
        raw_contents = service.users().messages().get(userId='me', id=email['id']).execute()
        encoded_contents = raw_contents['payload']['parts'][0]['body']['data']
        decoded_contents = base64.urlsafe_b64decode(encoded_contents).decode('utf-8')
        result = re.search('your Gift Card Code is (.*)', decoded_contents).group(1)
        if len(result) != 17:
            print("no gift card code found")
        ws1.cell(row=index+1, column=1, value=result)
        wb.save(file_name)


if __name__ == '__main__':
    main()