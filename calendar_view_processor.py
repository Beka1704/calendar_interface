from datetime import datetime
from datetime import timedelta
import pytz
import dateutil.parser
import json

cet = pytz.timezone('Europe/Amsterdam')


def iso_to_locale_datetime(datestring, timezonestring):
    format = '%Y-%m-%dT%H:%M:%S%z'
    utc = pytz.utc
    utc_d = dateutil.parser.parse(datestring)  # python 2.7
    utc_d = utc.localize(utc_d)#utc_d.replace(tzinfo=utc)  # - d.utcoffset()
    pst = pytz.timezone(timezonestring)
    return utc_d.astimezone(pst)


def split_multiple_day_events(start, end, subject, all_day, status):
    events = []
    date = start
    while date < end:
        end_of_day = date.replace(hour=23, minute=59, second=59, microsecond=999)
        if end_of_day < end:
            events.append((date, end_of_day, subject, all_day, status))
        else:
            events.append((date, end, subject, all_day, status))
            break
        date = (date + timedelta(days=1)).replace(hour=0, minute=0, second=59, microsecond=0)
    return events


def create_event_list(doc):
    events = []
    for meeting in doc['value']:
        timezone = 'Europe/Amsterdam'
        subject = meeting['subject']
        # print(subject)
        all_day = meeting['isAllDay']
        status = meeting['showAs']
        start = iso_to_locale_datetime(meeting['start']['dateTime'], timezone)
        end = iso_to_locale_datetime(meeting['end']['dateTime'], timezone)
        events.extend(split_multiple_day_events(start, end, subject, all_day, status))
    return events


def outlook_json_to_returnformat(outlook_responds, date):
    date_until = cet.localize(datetime.strptime(date, '%Y-%m-%d'))

    doc = json.loads(outlook_responds)
    days = {}
    t_format = '%H:%M:%S'
    d_format = '%y-%m-%d'
    for e in create_event_list(doc):
        print(str(e[0])+"<"+str(date_until)+":"+str(e[0] < date_until))
        if e[0] < date_until+timedelta(days=1):
            if e[0].date().strftime(d_format) not in days:
                days[e[0].date().strftime(d_format)] = []
            event = {'start': e[0].strftime(t_format), 'end': e[1].strftime(t_format), 'subject': e[2], 'allDay': e[3],
                     'status': e[4]}
            days[e[0].date().strftime(d_format)].append(event)
    return days