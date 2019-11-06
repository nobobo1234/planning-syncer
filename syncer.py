import pickle
from httplib2 import Http
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime, timedelta

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/calendar']

PLANNER_SPREADSHEET_ID = '17U_Xle35zJVBx_qI27uLamevZnZwULLGMaw8aQWBZpA'

def main():
    creds = None

    week = int(input('Which week do you want to sync?'))
    ranges = list(map(lambda range: f'Week {week}!{range}', ['A9:B', 'C9:D', 'E9:F', 'G9:H', 'I9:J']))

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    calendar_service = build('calendar', 'v3', credentials=creds)

    sheet = service.spreadsheets()
    result = sheet.values().batchGet(spreadsheetId=PLANNER_SPREADSHEET_ID, ranges=ranges).execute()
    day_values = result.get('valueRanges', [])

    batch = calendar_service.new_batch_http_request()

    current_day = datetime.today() - timedelta(days=datetime.today().weekday())
    for day in day_values:
       for values in day['values']:
           date_time = values[0].split(' - ')
           start = date_time[0].split(':')
           end = date_time[1].split(':')
           event = {
               'summary': values[1],
               'start': {
                   'dateTime': current_day.replace(hour=int(start[0]), minute=int(start[1])).isoformat(),
                   'timeZone': 'Europe/Amsterdam'
               },
               'end': {
                   'dateTime': current_day.replace(hour=int(end[0]), minute=int(end[1])).isoformat(),
                   'timeZone': 'Europe/Amsterdam'
               },
           }
           print(current_day.isoformat())
           batch.add(calendar_service.events().insert(calendarId='primary', body=event))
       current_day = current_day + timedelta(days=1)
       current_day = current_day.replace(hour=1, minute=0)
    http = Http()
    batch.execute(http=http)

if __name__ == '__main__':
    main()


