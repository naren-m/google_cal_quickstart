
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'

    #  Must be an RFC3339 timestamp with mandatory time zone offset,
    # e.g., 2011-06-03T10:00:00-07:00, 2011-06-03T10:00:00Z.
    # 'Z' indicates UTC time
    prevMonth = get_prev_nth_month_date(3).isoformat() + 'Z'

    print('Getting the past n months events')
    eventsResult = service.events().list(
        calendarId='primary',
        timeMin=prevMonth,
        timeMax=now,
        singleEvents=True,
        orderBy='startTime',
        timeZone='PST').execute()

    events = eventsResult.get('items', [])
    meetings = list()

    if not events:
        print('No upcoming events found.')
    for event in events:
        meeting = dict()
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])
        meeting['organizer'] = event['organizer']
        if 'attendees' in event.keys():
            meeting['attendees'] = event['attendees']
        else:
            meeting['attendees'] = None
        if 'description' in event.keys():
            meeting['description'] = event['description']
        else:
            meeting['description'] = None

        meeting['creator'] = event['creator']
        meeting['status'] = event['status']
        meeting['summary'] = event['summary']

        meeting['start'] = event['start']
        meeting['end'] = event['end']
        meeting['created'] = event['created']
        meetings.append(meeting)

    # print(meetings)


def get_prev_nth_month_date(n):
    days = n * 30
    lastMonth = datetime.datetime.now() - datetime.timedelta(days=days)
    print(lastMonth)
    return lastMonth


if __name__ == '__main__':
    main()
