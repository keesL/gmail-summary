#!/usr/bin/python3
import os.path
import datetime
import re
import smtplib

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from settings import FROM
from settings import TO
from settings import STYLE
from settings import FOLLOWUP_LABEL

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

def main():
  """Based on Google Quickstart.py
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
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    # Call the Gmail API
    service = build("gmail", "v1", credentials=creds)

    # fetch the purgatory label
    results = service.users().labels().list(userId="me").execute()
    label_id = [label['id'] for label in results.get("labels", []) if label['name'] == FOLLOWUP_LABEL]
    if len(label_id) == 0:
        print(f"Unable to find label {FOLLOWUP_LABEL}.")
        return
    label_id = label_id[0]

    # Get all message in the Inbox and return their IDs
    results = service.users().messages().list(userId="me", labelIds = [label_id]).execute()
    message_ids = results.get("messages", [])
    if not message_ids:
        print("No messages found.")
        return
    messages = []

    """ 
    Helper function for batch processing
    """
    def add_message_to_batch(id, msg, err):
        # id is given because this will not be called in the same order
        if err:
            print(err)
        else:
            messages.append(msg)

    batch = service.new_batch_http_request()
    for msg in message_ids:
        batch.add(service.users().messages().get(userId='me', id=msg['id']), add_message_to_batch)
    batch.execute()

    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = 'Daily email summary'
    msg['From'] = FROM
    msg['To'] = TO

    html='<div>You have the following pending action items:</div><div><table>';
    txt='You have the following pending action items\n\n'
    # determine if message needs to be purgatoried
    for message in messages:
        headers=message['payload']['headers']
        hit = False
        to=subject=frm=date=None
        for h in headers:
            # extract relevant headers
            if h['name'] == 'To': to = h['value'].lower()
            elif h['name'] == 'Subject': subject = h['value']
            elif h['name'] == 'From': frm=h['value'].replace('"', '')
            elif h['name'] == 'Date': date=h['value']
            
        if not (to and frm and subject):
            print(f'Unable to parse message.')
        else:
            m=re.search('^(.+)<([a-zA-Z0-9_.-]+@[a-zA-Z0-9_.-]+)>', frm)
            if m:
                frm=m.group(1)[:22]
            else:
                frm=frm[:22]
            # Sat, 30 Mar 2024 07:34:55 -040
            m=re.search('^(.+ [+-][0-9]{4})', date)
            if m:
                date=m.group(1)
                dt_obj = datetime.datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
                date=dt_obj.ctime()
            else:
                date=date[:32]
            txt += f'{date} | {frm:22} | {subject}'
            html += f'''<tr>
    <td style="{STYLE}">{date}</td>
    <td style="{STYLE}">{frm}</td>
    <td style="{STYLE}">{subject}</td>
    </tr>\n
            '''
   
    html += '</table></div>'

  except HttpError as error:
    # TODO(developer) - Handle errors from gmail API.
    print(f"An error occurred: {error}")


  part1=MIMEText(txt, 'plain')
  part2=MIMEText(html, 'html')
  msg.attach(part1)
  msg.attach(part2)
  s=smtplib.SMTP('localhost')
  s.sendmail(FROM, TO, msg.as_string())
  s.quit()

if __name__ == "__main__":
  main()
