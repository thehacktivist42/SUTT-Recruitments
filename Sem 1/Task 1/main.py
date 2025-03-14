import pandas as pd
import json
times = { # Dictionary for the hour number to 24-hour time conversion
        1 : 8,
        2 : 9,
        3 : 10,
        4 : 11,
        5 : 12,
        6 : 13,
        7 : 14,
        8 : 15,
        9 : 16
    }
days = { # Dictionary for the day abbreviation to full-form conversion
    'M' : 'Monday',
    'T' : 'Tuesday',
    'W' : 'Wednesday',
    'Th' : 'Thursday',
    'F' : 'Friday',
    'S' : 'Saturday'
}
venues = { # QoL Feature addition: Dictionary for the venue names
    1 : 'FD-I',
    2 : 'FD-II',
    3 : 'FD-III',
    5 : 'LTC',
    6 : 'NAB',
    7 : 'Central Workshop'
}

file = pd.ExcelFile("Timetable Workbook - SUTT Task 1.xlsx") # Generates an object from the Excel file

d = {} # A dictionary for the multiple sheets in the Excel file. The following two lines create key value pairs for the sheet number and the respective dataframe.
for sheet in file.sheet_names:
    d[f'{sheet}'] = pd.read_excel(file, skiprows = 1, sheet_name = sheet)

subjects = {} # Main dictionary to be converted to JSON

for i in d: # For every sheet containing a subject, the following code is executed
    d[i].drop(0)
    temp = {}
    for j in d[i]["COURSE NO."]:
        if j == j: # Check for non-NaN values
            temp["courseCode"] = j
            break
    subjects[i] = temp
    for j in d[i]["COURSE TITLE"]:
        if j == j and j != "Tutorial": # Checks for values that are non-NaN and not called "Tutorial". Only one such value exists and that is the course title.
            temp["courseTitle"] = j
            break
    credits = {"L" : "-", "P" : "-", "U" : "-"} # Initializes a dictionary for the credit structure. If an integer is found in the respective columns, the '-' is replaced with it.
    for j in d[i]["CREDIT"]:
        if j == j and type(j) is int:
            credits["L"] = j
    for j in d[i]["Unnamed: 4"]:
        if j == j and type(j) is int:
            credits["P"] = j
    for j in d[i]["Unnamed: 5"]:
        if j == j and type(j) is int:
            credits["U"] = j
    temp["credits"] = credits
    secList = [] # List of sections
    for sec in d[i]["SEC"]:
        if sec == sec:
            secList.append(sec)
    sections = {sec : {"instructors" : [], "room" : '', "time" : '' } for sec in secList} # Initializes a dictionary for the section structure, containing a list of instructors, room and time
    profList = []
    curSec = ''
    curTime = ''
    for sec, prof, room, time in zip(d[i]["SEC"], d[i]["INSTRUCTOR-IN-CHARGE / Instructor"], d[i]["ROOM"], d[i]["DAYS & HOURS"]):
        temp1 = {}
        if sec == sec:
            curSec = sec # If the column contains a non-NaN value, it is taken as the current section
            sections[curSec]['room'] = f"{int(room)} ({venues[int(str(int(room))[0])]})" # The room and the location of the section is added to the sections dictionary.
            time = time.split('  ') # The time is converted into a list containing the day and the hours
            for i in time:
                if not i[0].isdigit(): # If not digit, it is the day and not the hours.
                    try:
                        sections[curSec]['time'] += days[i] # Only possible when only single character, i.e., single day exists.
                        if len(time[time.index(i)+1]) == 1: # Only possible when only single digit, i.e., single hour exists.
                            sections[curSec]['time'] += f" {times[int(time[time.index(i)+1])]}:00 - {times[int(time[time.index(i)+1])]+1}:00  "
                        else: # Considered when the class lasts for more than one hour.
                            sections[curSec]['time'] += f" {times[int(time[time.index(i)+1][0])]}:00 - {times[int(time[time.index(i)+1][-1])]+1}:00  "
                    except: # Considered when the class is on multiple days.
                        for j in i.split():
                            sections[curSec]['time'] += days[j]
                            sections[curSec]['time'] += f" {times[int(time[time.index(i)+1])]}:00 - {times[int(time[time.index(i)+1])]+1}:00  "
            sections[curSec]['time'] = sections[curSec]['time'].rstrip() # Removes extra whitespace on the right side.
        if curSec != '':
            sections[curSec]['instructors'].append(prof) # Adds instructors to the current section until the next section is found.
    temp["sections"] = sections

print(json.dumps(subjects, indent = 4)) # Pretty-prints the final dictionary to the terminal.
with open('data.json', 'w', encoding = 'utf-8') as f:
    json.dump(subjects, f, indent = 4)  # Dumps the dictionary to data.json