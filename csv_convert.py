import csv
from sys import argv
from os import path
from datetime import datetime, timedelta
from icalendar import Calendar, Event, vDatetime


def csv_to_ics(in_csv, busy_status='OOF', out_ics=None, ):
    """
    in_csv is the path to the input csv with calendar information, can be relative.
    busy_status is the status for all calendar items. Options are 'FREE', 'TENTATIVE', 'BUSY', or 'OOF' (out of office)
    out_ics is the path to the output ics file. Defaults to same name and location as in_csv, with .ics extension
    """
    print(in_csv)
    if not out_ics:
        out_ics = path.splitext(in_csv)[0] + ".ics"
    cal = Calendar()
    cal.add('prodid', '-//Travis Ormsby//')
    cal.add('version', '2.0')

    with open(in_csv, newline='') as f:
        reader = csv.DictReader(f)
        
        for row in reader:
            datetime_start = datetime.strptime(row['Start Date'], '%m/%d/%Y')
            datetime_end = datetime_start + timedelta(days=1)
            e = Event()
            e.add('summary', row['Subject'])
            e.add('dtstart', vDatetime(datetime_start))
            e.add('dtend', vDatetime(datetime_end))
            e.add('X-MICROSOFT-CDO-BUSYSTATUS', busy_status)
            e.add('X-MICROSOFT-CDO-IMPORTANCE', '1')
            cal.add_component(e)

    with open(out_ics, 'wb') as f:
        f.write(cal.to_ical())

if __name__ == '__main__':
    csv_to_ics(*argv[1:])
