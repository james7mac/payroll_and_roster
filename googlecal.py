from __future__ import print_function
import datetime
import pickle
import os.path
import pytz
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def get_creds():
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
    service = build('calendar', 'v3', credentials=creds)
    return service



def update_calander(data, service):
    # Call the Calendar API
    #print('Getting the upcoming 10 events')
    service.events().insert(calendarId='primary',body=data).execute()
    print(data)

def check_work_event(start_date, service):
    # Returns the event ID if it exists else False
    local = pytz.timezone('Australia/Melbourne')
    local_dt_start = local.localize(start_date, is_dst=None)
    local_dt_end = local.localize(start_date + datetime.timedelta(hours=23, minutes=59, seconds=59
                                                                  ), is_dst=None)
    utc_dt_start = local_dt_start.astimezone(pytz.utc).isoformat()
    utc_dt_end = local_dt_end.astimezone(pytz.utc).isoformat()

    events = service.events().list(calendarId='primary',timeMin=utc_dt_start, timeMax=utc_dt_end).execute()
    work_names = ['work', 'rest', 'ex']
    for i in events['items']:
        if i['summary'].lower() in work_names:
            return i
    return False

def delete_event(id, service):
    service.events().delete(calendarId='primary', eventId=id).execute()

if __name__ == '__main__':
    pass