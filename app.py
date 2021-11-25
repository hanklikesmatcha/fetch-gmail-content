from __future__ import print_function
import argparse
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import base64
import re
from datetime import date, timedelta, datetime
from openpyxl import Workbook

# If modifying these scopes, delete the file token.pickle.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
wb = Workbook()
ws = wb.active
file_name = "gift-codes"
ws.title = "Gift Codes"


def generate_file(count: int, codes: []):
    while count != 0:
        if count != len(codes):
            print("Number doesn't match with the total orders")
            return
        for index, code in enumerate(codes):
            ws.append([code])
            count -= 1
        wb.save(file_name + "{}.xlsx".format(date.today()))
    print(count)
    return count


def extract_mails(threads: list, service: build):
    count = 0
    gift_codes: list = []
    for index, email in enumerate(threads):
        raw_contents = (
            service.users().messages().get(userId="me", id=email["id"]).execute()
        )["payload"]["parts"][0]["body"]["data"]
        decoded_contents = base64.urlsafe_b64decode(raw_contents).decode("utf-8")
        match = re.search("and your Gift Card Code is\s+", decoded_contents)

        if match is None:
            generate_file(count=count, codes=gift_codes)

        else:
            matched_number = decoded_contents[match.end() :]
            if len(matched_number) != 16:
                print("no gift card code found")
            count += 1
            gift_codes.append(matched_number)
    return count, gift_codes


def main(sender: str, starts_from: str):
    # 'hello@thegoodregistry.com'
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
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

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

    if len(mail_group["messages"]) < 1:
        print("No sender found")

    threads = (
        service.users()
        .messages()
        .list(
            userId="me",
            q="from='{}' after: {}".format(sender, timestamp),
            # pageToken=mail_group.get("nextPageToken"),
            pageToken="",
            # Make this blank to blank to get mails from another mail group
        )
        .execute()
    )["messages"]
    count, gift_codes = extract_mails(threads=threads, service=service)
    return generate_file(count=count, codes=gift_codes)


parser = argparse.ArgumentParser(
    description="you may search based on sender and timestamp"
)
parser.add_argument("sender", type=str, help="Sender email")
parser.add_argument("starts_from", type=str, help="Day after, format: yyyymmdd")
args = parser.parse_args()

if __name__ == "__main__":
    main(args.sender, args.starts_from)
