from datetime import datetime
from ics import Calendar, Event
import pandas as pd

c = Calendar()
e = Event()

excel = pd.read_excel("test.xlsx")
excel = excel.iloc[2:, :]
excel.reset_index(drop=True)

courses = [] # each course has a: course code & number, type, days, time, location, start, and end

#print(excel.iloc[0,4])

for x in range(0, excel.shape[0]):
    course_name = excel.iloc[x,4]
    delimiter_pos = course_name.find('-')
    course_name = course_name[: course_name.find('-')+4].replace('-', ' ')

    meeting_pattern = excel.iloc[x,7]
    info = meeting_pattern.split('|')

    #print(course_name)
