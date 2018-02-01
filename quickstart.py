
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime

from collections import Counter
from dateutil.relativedelta import relativedelta


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

    num_months = 3
    #  Must be an RFC3339 timestamp with mandatory time zone offset,
    # e.g., 2011-06-03T10:00:00-07:00, 2011-06-03T10:00:00Z.
    # 'Z' indicates UTC time
    to_date = datetime.datetime.utcnow().isoformat() + 'Z'
    from_date = get_prev_nth_month_date(num_months).isoformat() + 'Z'

    print('Getting the past n months events')
    meetings = get_meetings(from_date, to_date)
    # print(meetings)
    print("Time spent in meetings:", time_spent_in_meetings(meetings))   # 1
    print("Time spent in interviews:", time_spent_in_interviews(meetings))  # 5
    print("Top attendees:", get_top_n_attendees(meetings, 3))


def event_has_attendees(attendees):
    if attendees is None:
        return False
    if len(attendees) <= 1:
        return False

    return True


def get_top_n_attendees(meetings, n):
    attendees = Counter()
    for meeting in meetings:
        if not event_has_attendees(meeting['attendees']):
            continue
        for attendee in meeting['attendees']:
            if 'displayName' in attendee.keys():
                name = attendee['displayName']
            else:
                name = attendee['email']
            attendees[name] += 1

    return attendees.most_common(n)


def get_meetings(from_date, to_date):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    events_result = service.events().list(
        calendarId='primary',
        timeMax=to_date,
        timeMin=from_date,
        singleEvents=True,
        orderBy='startTime').execute()

    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')

    meetings = list()

    for event in events:
        meeting = dict()

        if 'attendees' in event.keys():
            meeting['attendees'] = event['attendees']
        else:
            meeting['attendees'] = None

        if not event_has_attendees(meeting['attendees']):
            continue

        meeting['organizer'] = event['organizer']

        if 'description' in event.keys():
            meeting['description'] = event['description']
        else:
            meeting['description'] = None

        meeting['creator'] = event['creator']
        meeting['status'] = event['status']
        meeting['summary'] = event['summary']

        start_time = parse_date(event['start'].get(
            'dateTime', event['start'].get('date')))
        end_time = parse_date(event['end'].get(
            'dateTime', event['end'].get('date')))

        meeting['start'] = start_time
        meeting['end'] = end_time
        meeting['created'] = event['created']

        meeting['duration'] = end_time - start_time
        meetings.append(meeting)

        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'], meeting['duration'])

    return meetings


def is_interview(string):
    if string is None:
        return False

    interview_keywords = ['interview', 'connect call', 'recruitment']

    if any(word.upper() in string.upper() for word in interview_keywords):
        return True

    return False


def time_spent_in_interviews(meetings):
    total_time = datetime.timedelta(0)
    for meeting in meetings:
        if is_interview(meeting['description']) or \
                is_interview(meeting['summary']):
            total_time += meeting['duration']
    return total_time


def time_spent_in_meetings(meetings):
    total_time = datetime.timedelta(0)
    for meeting in meetings:
        total_time += meeting['duration']
    return total_time


"""

1. Total time spent in meetings per month for the last 3 months - Done
    - grouping by month

2. Busiest week / relaxed week - Which month had highest number of meetings/least
number of meetings
    - group by week
3. Average number of meetings per week, average time spent every week in meetings.
    - group by week
4. Top 3 persons with whom you have meetings
    - attendees list count of top N

5. Time spent in Recruting/Conducting interviews - Done
    - description or summary containing  words 'interview' 'connect call'


{
    1 : // months
        { 1: // week
            {meeting dict}
            }
}

"""


def parse_date(date_string):
    return datetime.datetime.strptime(date_string, "%Y-%m-%dT%H:%M:%S-08:00")


def get_date(date_string):
    """
        Returns date from given string

        Sample date string 2017-11-05T11:00:00-08:00
        Returns 2017-11-05
    """
    return datetime.datetime.strptime(date_string.split('T')[0], "%Y-%m-%d")


def get_prev_nth_month_date(n):
    lastMonth = datetime.datetime.now() - relativedelta(months=n)
    return lastMonth


if __name__ == '__main__':
    main()
