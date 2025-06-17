from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from bs4 import BeautifulSoup
from requests_html import HTMLSession

# If changing  scopes, delete the file token.json so it can resets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of my spreadsheet.
SPREADSHEET_ID = 'SECRET'
RANGE_NAME = 'Artists!A2:E'

def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, created on first run
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in through the browser
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        line = 1
        for row in values:
            url_bugs_release = "https://music.bugs.co.kr/artist/" + row[1] + "/albums?type=RELEASE"
            session = HTMLSession()
            res = session.get(url_bugs_release)
            soup = BeautifulSoup(res.html.html, "html.parser")
            all_release = soup.find("ul", {"class": "list tileView albumList"}).findAll("li")
            number_release = len(all_release)
            page = 1
            while number_release % 70 == 0:
                page += 1
                page_url = url_bugs_release + "&page=" + str(page)
                session = HTMLSession()
                res = session.get(page_url)
                soup = BeautifulSoup(res.html.html, "html.parser")
                all_release = soup.find("ul", {"class": "list tileView albumList"}).findAll("li")
                number_release += len(all_release)
            if row[2] != str(number_release):
                batch_update_values_request_body = {
                    "valueInputOption": "USER_ENTERED",
                    "data": [
                        {
                            'range': 'Artists!C' + str(line + 1),
                            'values': [[str(number_release)]]
                        }
                    ]
                }
                service.spreadsheets().values().batchUpdate(
                   spreadsheetId=SPREADSHEET_ID, 
                   body=batch_update_values_request_body
                ).execute()
                print("Updated number of releases for " + row[1])
            percentage = int((line / len(values)) * 100)
            batch_update_values_request_body = {
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {
                        'range': 'Artists!I1',
                        'values': [[str(percentage) + "%"]]
                     }
                  ]
            }
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=SPREADSHEET_ID, 
                body=batch_update_values_request_body
            ).execute()
            line += 1
        batch_update_values_request_body = {
            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    'range': 'Artists!I1',
                    'values': [["DONE"]]
                 }
              ]
        }
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, 
            body=batch_update_values_request_body
        ).execute()
        print("Finished updating!")
        
if __name__ == '__main__':
    main()
