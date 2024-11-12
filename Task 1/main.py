import pandas as pd
import json
times = {
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
days = {
    'M' : 'Monday',
    'T' : 'Tuesday',
    'W' : 'Wednesday',
    'Th' : 'Thursday',
    'F' : 'Friday',
    'S' : 'Saturday'
}
venues = {
    1 : 'FD-I',
    2 : 'FD-II',
    3 : 'FD-III',
    5 : 'LTC',
    6 : 'NAB',
    7 : 'Central Workshop'
}

file = pd.ExcelFile("Timetable Workbook - SUTT Task 1.xlsx")

d = {}
for sheet in file.sheet_names:
    d[f'{sheet}'] = pd.read_excel(file, skiprows = 1, sheet_name = sheet)

subjects = {}

for i in d:
    d[i].drop(0)
    temp = {}
    for j in d[i]["COURSE NO."]:
        if j == j:
            temp["courseCode"] = j
            break
    subjects[i] = temp
    for j in d[i]["COURSE TITLE"]:
        if j == j and j != "Tutorial":
            temp["courseTitle"] = j
            break
    credits = {"L" : "-", "P" : "-", "U" : "-"}
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
    secList = []
    for sec in d[i]["SEC"]:
        if sec == sec:
            secList.append(sec)
    sections = {sec : {"instructors" : [], "room" : '', "time" : '' } for sec in secList}
    profList = []
    curSec = ''
    curTime = ''
    for sec, prof, room, time in zip(d[i]["SEC"], d[i]["INSTRUCTOR-IN-CHARGE / Instructor"], d[i]["ROOM"], d[i]["DAYS & HOURS"]):
        temp1 = {}
        if sec == sec:
            curSec = sec
            sections[curSec]['room'] = f"{int(room)} ({venues[int(str(int(room))[0])]})"
            time = time.split('  ')
            for i in time:
                if not i[0].isdigit():
                    try:
                        sections[curSec]['time'] += days[i]
                        if len(time[time.index(i)+1]) == 1:
                            sections[curSec]['time'] += f" {times[int(time[time.index(i)+1])]}:00 - {times[int(time[time.index(i)+1])]+1}:00  "
                        else:
                            sections[curSec]['time'] += f" {times[int(time[time.index(i)+1][0])]}:00 - {times[int(time[time.index(i)+1][-1])]+1}:00  "
                    except:
                        for j in i.split():
                            sections[curSec]['time'] += days[j]
                            sections[curSec]['time'] += f" {times[int(time[time.index(i)+1])]}:00 - {times[int(time[time.index(i)+1])]+1}:00  "
            sections[curSec]['time'] = sections[curSec]['time'].rstrip()
        if curSec != '':
            sections[curSec]['instructors'].append(prof)
    temp["sections"] = sections

print(json.dumps(subjects, indent = 4)) 
with open('data.json', 'w', encoding = 'utf-8') as f:
    json.dump(subjects, f, indent = 4)   