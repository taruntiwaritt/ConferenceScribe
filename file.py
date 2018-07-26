import datetime
import win32com.client as wincl
import csv
import re
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
speak = wincl.Dispatch("SAPI.SpVoice")
stopWords = set(stopwords.words("english"))
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

query="can you tell me the  time today"
words = word_tokenize(query)
filtered_query = [w for w in words if not w in stopWords]
filtered_query = []
for w in words:
    if w not in stopWords:
        filtered_query.append(w)
print(filtered_query)

now=datetime.datetime.now()
str=now.date()
type(str)
if "current" in filtered_query or "today" in filtered_query:
    if "time" in query:
        t=now.time()
        speak.Speak('Time right now is %s'%t.hour)
        speak.Speak(t.minute)
    if "date" in query: 
        d=now.date()
        speak.Speak('Todays Date is %s'%d.isoformat())

query="can you tell me the summary of meeting"
words = word_tokenize(query)
filtered_query = [w for w in words if not w in stopWords]
filtered_query = []
for w in words:
    if w not in stopWords:
        filtered_query.append(w)
print(filtered_query)
if "summary" in filtered_query and "meeting" in filtered_query:
       with open("timeLists.csv",'r') as input:
           read = csv.reader(input)
           speak.Speak("the Short summary for the previous meeting is")
           for row in read:
               speak.Speak(row)
    
query="When is the next meeting"
words = word_tokenize(query)
filtered_query = [w for w in words if not w in stopWords]
filtered_query = []
for w in words:
    if w not in stopWords:
        filtered_query.append(w)
print(filtered_query)
if ("time" and "meeting" in filtered_query) or ("when" and "meeting" in filtered_query):
    with open("timeLists.csv",'r') as input:
           read = csv.reader(input)
           for row in read:
               v=row[0]
               reg=re.findall(r'((0[0-1]|[1-59]\d)(:(0[0-1]|[1-59]\d)\s(AM|am|PM|pm))?)',v)
               day=re.findall(r'\b((mon|tues|wed(nes)?|thur(s)?|fri|sat(ur)?|sun)(day)?)\b',v)
               if reg:
                   reg.sort(reverse=True)
                   speak.Speak(reg[0][0])
                   if day:
                       day.sort(reverse=True)
                       speak.Speak(day[0][0])

