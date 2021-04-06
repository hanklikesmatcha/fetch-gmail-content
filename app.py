from __future__ import print_function
import pickle
import os.path
import sys
import argparse
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import re
from datetime import date, timedelta, datetime
from openpyxl import Workbook, load_workbook

# If modifying these scopes, delete the file token.pickle.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
wb = Workbook()
ws = wb.active
file_name = "gift-codes"
ws.title = "Gift Codes"


def main(sender: str, starts_from: str):
    # 'hello@thegoodregistry.com'
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    service = build("gmail", "v1", credentials=creds)

    start = starts_from[:4] + "-" + starts_from[4:6] + "-" + starts_from[6:]
    timestamp = str(datetime.strptime(start, "%Y-%m-%d").date() - timedelta(1)).replace(
        "-", "/"
    )
    # Call the Gmail API
    mail_group = (
        service.users()
        .messages()
        .list(userId="me", q="from='{}' after: {}".format(sender, timestamp))
        .execute()
    )
    count = 0

    gift_codes = []

    if len(mail_group["messages"]) < 1:
        print("No sender found")

    next_page = True

    def generate_file(count: int, codes: []):
        while count != 0:
            if count != len(codes):
                print("Number doesn't match with the total orders")
                return
            for index, code in enumerate(gift_codes):
                ws.append([code])
                count -= 1
            file = wb.save(file_name + "{}.xlsx".format(date.today()))

            print("total orders {}".format(len(codes)))
        return count

    while next_page:
        if next_page is False:
            break

        for index, email in enumerate(mail_group["messages"]):
            raw_contents = (
                service.users().messages().get(userId="me", id=email["id"]).execute()
            )
            encoded_contents = raw_contents["payload"]["parts"][0]["body"]["data"]
            decoded_contents = base64.urlsafe_b64decode(encoded_contents).decode(
                "utf-8"
            )
            match = re.search("and your Gift Card Code is\s+", decoded_contents)

            if match is None:
                return

            matched_number = decoded_contents[match.end() :]
            if len(matched_number) != 16:
                print("no gift card code found")
            count += 1
            gift_codes.append(matched_number)

        if mail_group.get('nextPageToken'):
            mail_group = (
                service.users()
                .messages()
                .list(
                    userId="me",
                    q="from='{}' after: {}".format(sender, timestamp),
                    pageToken=mail_group["nextPageToken"],
                )
                .execute()
            )
        else:
            next_page = False 
            print("estimated:{}".format(count))
            generate_file(count=count, codes=gift_codes)


parser = argparse.ArgumentParser(
    description="you may search based on sender and timestamp"
)
parser.add_argument("sender", type=str, help="Sender email")
parser.add_argument("starts_from", type=str, help="Day after, format: yyyymmdd")
args = parser.parse_args()

if __name__ == "__main__":
    main(args.sender, args.starts_from)
