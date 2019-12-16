import tika
from tika import parser

def dictMachine(item):
    split = item.split(":")
    key = split[0]
    value = split[1].strip()
    d = {}
    d[key] = value
    return d

def tupleMachine(row):
    for index in range(len(row)):
        for key in row[index]:
            if "ST-WK-TOTAL" in key:
                stTotal = row[index].get(key)
            if "WORK ORDER #" in key:
                workOrderNum = row[index].get(key)
            if "OT-WK-TOTAL" in key:
                otTotal = row[index].get(key)
            if "JOB #" in key:
                jobNumber = row[index].get(key)
    return stTotal, otTotal, workOrderNum, jobNumber

timecard = parser.from_file('Time Card Form.pdf')
timecardContent = timecard['content']
split = timecardContent.split('\n')
split.sort()

row1 = []
row2 = []
row3 = []
row4 = []
row5 = []
row6 = []
row7 = []

for item in split:
    item = item.strip('\t')
    if item:
        if item.startswith('EmployeeNumber:'):
            split = item.split(' ')
            employeeNumber = split[1]
        if item.startswith('WeekEnding:'):
            split = item.split(' ')
            weekEnding = split[1]
        if item.startswith('1-'):
            row1.append(dictMachine(item))
        if item.startswith('2-'):
            row2.append(dictMachine(item))
        if item.startswith('3-'):
            row3.append(dictMachine(item))
        if item.startswith('4-'):
            row4.append(dictMachine(item))
        if item.startswith('5-'):
            row5.append(dictMachine(item))
        if item.startswith('6-'):
            row6.append(dictMachine(item)) 
        if item.startswith('7-'):
            row7.append(dictMachine(item))

# STTOTAL, OTTOTAL, WORKORDERNUM, JOBNUM
row1Tuple = tupleMachine(row1)
row2Tuple = tupleMachine(row2)
row3Tuple = tupleMachine(row3)
row4Tuple = tupleMachine(row4)
row5Tuple = tupleMachine(row5)
row6Tuple = tupleMachine(row6)
row7Tuple = tupleMachine(row7)

