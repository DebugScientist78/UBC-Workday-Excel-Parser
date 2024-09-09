from datetime import datetime
import dateutil.tz
import icalendar
import pandas as pd

def parse_time_range(time_range):
    # Split the time range into start and end times
    start_time_str, end_time_str = time_range.split('-')
    
    # Define a function to parse a time string to a datetime object
    def parse_time(time_str):
        return datetime.strptime(time_str, '%I:%M%p')
    
    # Parse the start and end times
    start_time = parse_time(start_time_str)
    end_time = parse_time(end_time_str)
    
    # Calculate the duration
    duration = end_time - start_time
    
    return duration

c = icalendar.Calendar()
excel = pd.read_excel("test.xlsx")
excel = excel.iloc[2:, :]
excel.reset_index(drop=True)

courses = [] # each course has a: course code & number, type, days, time, location, start, and end

#print(excel.iloc[0,4])

days_dic = {
    "Mon": "MO",
    "Tue": "TU",
    "Wed": "WE",
    "Thu": "TH",
    "Fri": "FR"
}

start_date_modifer = {
    "Mon": -1,
    "Tue": 0,
    "Wed": 1,
    "Thu": 2,
    "Fri": 3
}


for x in range(0, excel.shape[0]):
    course_name = excel.iloc[x,4]
    delimiter_pos = course_name.find('-')
    course_name = course_name[: course_name.find('-')+4].replace('-', ' ')

    meeting_pattern = excel.iloc[x,7]
    info = meeting_pattern.split('|')
    start_date = excel.iloc[x,10]
    end_date = excel.iloc[x,11]

    days = info[1].split(' ')
    days.pop(0)
    days.pop(len(days) - 1)

    location_description = info[-1]

    times = info[2].replace(' ', '').replace('.', '')
    start_time = datetime.strptime(times.split('-')[0], '%I:%M%p')
    end_time = datetime.strptime(times.split('-')[1], '%I:%M%p')
    duration = parse_time_range(times)
    duration = duration.seconds / (60*60)

    e = icalendar.Event()

    e.add('summary', course_name)
    e.add('version', 2.0)
    e.add('proid', "myCalender")

    #e.name = course_name

    if (start_date.weekday() == 0):
        start_date = start_date.replace(day= (start_date.day + start_date_modifer[days[0]] + 1) )
    else:
        start_date = start_date.replace(day= (start_date.day + start_date_modifer[days[0]]) )

    tz = dateutil.tz.tzstr("Canada/Vancouver")
    e["X-DT-ZONEINFO"] = icalendar.vDatetime(datetime(start_date.year, start_date.month, start_date.day, start_time.hour, start_time.minute, tzinfo=tz))
    e.add("dtstart", icalendar.vDatetime(datetime(start_date.year, start_date.month, start_date.day, start_time.hour, start_time.minute, tzinfo=tz)))
    #e.add("dtend", datetime(date_obj.year, date_obj.month, date_obj.day, end_time.hour, end_time.minute))
    e.add("description", location_description)

    '''
    e.begin = datetime(date_obj.year, date_obj.month, date_obj.day, start_time.hour, start_time.minute)
    e.duration = {'hours': duration}
    e.description = location_description
    '''

    d = []
    for day in days:
        d.append(days_dic[day])
    
    recurrance = icalendar.vRecur({
        'freq': 'weekly',
        'byday': d,
        'until': end_date
    })

    e.add('rrule', recurrance)

    '''
    extra_string = "FREQ=WEEKLY;BYDAY="
    for day in days:
        extra_string += f'{days_dic[day]},'
    extra_string = extra_string[:len(extra_string) - 1]
    '''
    e.add('dtend', icalendar.vDatetime(datetime(start_date.year, start_date.month, start_date.day, end_time.hour, end_time.minute, tzinfo=tz)) )
    c.add_component(e)

with open('my.ics', 'wb') as f:
    f.write(c.to_ical())