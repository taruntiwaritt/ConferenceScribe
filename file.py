import datetime
import win32com.client as wincl
import csv
import re
speak = wincl.Dispatch("SAPI.SpVoice")
with open('helloworld.txt') as fp:
    lines = fp.read().split(".")
l=list()
for line in lines:
    if 'am' in line or 'pm' in line or 'o\' clock' in line or 'time' in line or 'meeting' in line or 'location' in line:
        l.append(line)

with open('timeLists.csv', 'w') as output:
    writer = csv.writer(output, lineterminator='\n')
    for val in l:
        writer.writerow([val])

query="what is the current system time"
speak.Speak("the query is "+query)
now=datetime.datetime.now()
str=now.date()
type(str)
if "current" in query or "today" in query:
    if "time" in query:
        t=now.time()
        speak.Speak('Time right now is %s'%t.hour)
        speak.Speak(t.minute)
    if "date" in query: 
        d=now.date()
        speak.Speak('Todays Date is %s'%d.isoformat())
query2="what is the summary of meeting"
speak.Speak("the query is "+query2)
if "summary" in query2 and "meeting" in query2:
       with open("timeLists.csv",'r') as input:
           read = csv.reader(input)
           speak.Speak("the Short summary for the previous meeting is")
           for row in read:
               speak.Speak(row)
    
query3="When is the next meeting"
speak.Speak("the query is "+query3)
if ("time" and "meeting" in query3) or ("when" and "meeting" in query3):
    with open("timeLists.csv",'r') as input:
           read = csv.reader(input)
           for row in read:
               v=row[0]
               reg=re.findall(r'((0[0-1]|[1-59]\d)(:(0[0-1]|[1-12]\d)\s(AM|am|PM|pm))?)',v)
               day=re.findall(r'\b((mon|tues|wed(nes)?|thur(s)?|fri|sat(ur)?|sun)(day)?)\b',v)
               if reg:
                   reg.sort(reverse=True)
                   speak.Speak(reg[0][0])
                   if day:
                       day.sort(reverse=True)
                       speak.Speak(day[0][0])