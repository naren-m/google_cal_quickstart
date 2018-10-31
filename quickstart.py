from __future__ import print_function
import httplib2
import os
import sys
import datetime

from collections import Counter
from dateutil.relativedelta import relativedelta

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage


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


def get_metings_count_and_weeks(meetings, num_months):
    num_of_weeks = 0
    num_of_meetings = 0

    last_month = get_prev_nth_month_date(num_months)

    for i in range(num_months):
        last_month = get_prev_nth_month_date(i)
        weeklist = get_weeks_of_month(last_month.month, last_month.year)
        num_of_weeks += len(weeklist)

    num_of_meetings = len(meetings)

    return num_of_meetings, num_of_weeks


def get_busiest_week(meetings, num_months):
    """
        If multiple week have same number of meetings, we will get most recent
        week statistics.
    """
    max_meetings = tuple()
    min_meetings = tuple()

    max_meeting_count = -1
    min_meeting_count = sys.maxsize

    for i in range(num_months):
        last_month = datetime.datetime.now() - relativedelta(months=i)
        weeklist = get_weeks_of_month(last_month.month, last_month.year)
        for week in weeklist:
            meetings_of_week = get_meetings_between_dates(
                meetings, week[0], week[-1])

            if len(meetings_of_week) > max_meeting_count:
                max_meeting_count = len(meetings_of_week)
                max_meetings = (week[0], week[-1], max_meeting_count)

            if len(meetings_of_week) < min_meeting_count:
                min_meeting_count = len(meetings_of_week)
                min_meetings = (week[0], week[-1], min_meeting_count)

    return min_meetings, max_meetings


def event_has_attendees(attendees):
    # Question: Does organizer comes under attendees
    # if yes: for it to be a meeting len(attendees) >1 else >=1
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


def get_meetings_between_dates(meetings, from_date, to_date):
    subset = list()
    for meeting in meetings:
        if from_date <= meeting['start'] <= to_date:
            subset.append(meeting)

    return subset


def get_meetings_from_calendar_api(from_date, to_date):
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


def get_weeks_of_month(month, year):
    import calendar

    week_list = calendar.Calendar().monthdatescalendar(year, month)
    filtered_weeklist = []
    for w in week_list:
        week = list()
        for day in w:
            if day.month == month:
                d = datetime.datetime.combine(
                    day, datetime.datetime.min.time())
                week.append(d)
        filtered_weeklist.append(week)

    return filtered_weeklist


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
    last_month = datetime.datetime.now() - relativedelta(months=n)
    return last_month


def main():
    num_months = 3
    #  Must be an RFC3339 timestamp with mandatory time zone offset,
    # e.g., 2011-06-03T10:00:00-07:00, 2011-06-03T10:00:00Z.
    # 'Z' indicates UTC time
    to_date = datetime.datetime.utcnow().isoformat() + 'Z'
    from_date = get_prev_nth_month_date(num_months).isoformat() + 'Z'

    meetings = get_meetings_from_calendar_api(from_date, to_date)

    print("===================================================================")
    print("==============           Meeting Stats            =================")
    print("===================================================================")
    total_meetings_duration = time_spent_in_meetings(meetings)
    print("Time spent in meetings:", total_meetings_duration)   # 1
    print("Time spent in interviews:", time_spent_in_interviews(meetings))  # 5
    print("Top attendees:", get_top_n_attendees(meetings, 3))

    relaxing_week, busiest_week = get_busiest_week(meetings, num_months)
    print("Busiest/relaxing weeks", busiest_week, relaxing_week)
    num_of_meetings, num_of_weeks = get_metings_count_and_weeks(
        meetings, num_months)

    print("Average meetings per week", num_of_meetings / num_of_weeks)
    print("Average meeting duration per week",
          total_meetings_duration / num_of_weeks)


if __name__ == '__main__':
    main()
