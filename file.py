import datetime
import win32com.client as wincl
import csv
import re
from PyDictionary import PyDictionary
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from weather import Weather, Unit
speak = wincl.Dispatch("SAPI.SpVoice")
import speech_recognition as sr
r = sr.Recognizer()
m = sr.Microphone()
stopWords = set (stopwords.words ("english"))
with open ('helloworld.txt') as fp:
    lines = fp.read ().split (".")
l = list()
for line in lines:
    if 'am' in line or 'pm' in line or 'o\' clock' in line or 'time' in line or 'meeting' in line or 'location' in line:
        l.append (line)
qq="meaning of abrasive"
for st in l:
    st.replace("\n","")
with open ('timeLists.csv', 'w') as output:
    writer = csv.writer (output, lineterminator='\n')
    for val in l:
        writer.writerow ([val])
#i=input("Press Enter To  Start Application ")
speak.Speak ("Hello! How can I help you?")
try:
    with m as source:
        r.adjust_for_ambient_noise (source)
    while True:
        with m as source:
            audio = r.listen (source)
            try:
                value = r.recognize_google (audio)
                if str is bytes:
                    query = u"{}".format (value).encode ("utf-8")
                else:
                    query = "{}".format (value)

                if(query=="goodbye"):
                    break
                print(query)
                words = word_tokenize (query)
                filtered_query = [w for w in words if not w in stopWords]
                filtered_query = []
                for w in words:
                    if w not in stopWords:
                        filtered_query.append (w)

                now = datetime.datetime.now ()
                str = now.date ()
                type (str)
                if "current" in filtered_query or "today" in filtered_query:
                    if "time" in query:
                        t = now.time ()
                        speak.Speak ('Time right now is %s' % t.hour)
                        speak.Speak (t.minute)
                    if "date" in query:
                        d = now.date ()
                        speak.Speak ('Todays Date is %s' % d.isoformat ())

                elif "summary" in filtered_query and "meeting" in filtered_query:
                    with open ("timeLists.csv", 'r') as input:
                        read = csv.reader (input)
                        speak.Speak ("the Short summary for the previous meeting is")
                        for row in read:
                            row[0].replace("\n","")
                            speak.Speak (row)

                elif ("time" in filtered_query and "meeting" in filtered_query) or ("when" in filtered_query and "meeting" in filtered_query):
                    with open ("timeLists.csv", 'r') as input:
                        read = csv.reader (input)
                        for row in read:
                            v = row[0]
                            reg = re.findall(r'((0[0-1]|[1-59]\d)(:(0[0-1]|[1-59]\d)\s(AM|am|PM|pm))?)', v)
                            day = re.findall(r'\b((mon|tues|wed(nes)?|thur(s)?|fri|sat(ur)?|sun)(day)?)\b', v)
                            if reg:
                                reg.sort(reverse=True)
                                speak.Speak(reg[0][0])
                                if day:
                                    day.sort(reverse=True)
                                    speak.Speak(day[0][0])
                elif ("location" in filtered_query and "meeting" in filtered_query):
                    with open ("timeLists.csv", 'r') as input:
                        read = csv.reader (input)
                        for row in read:
                            v=row[0]
                            if("location" in v and "meeting" in v):
                                speak.Speak(v)
                                
                elif("weather" in filtered_query or "prediction" in filtered_query):
                    weather = Weather(unit=Unit.CELSIUS)
                    location = weather.lookup_by_location('chennai')
                    x=1
                    z=0
                    forecasts = location.forecast
                    for forecast in forecasts:
                        if z<x:
                            speak.Speak("the prediction for chennai is ")
                            speak.Speak(forecast.text+"For the day")
                            speak.Speak(forecast.date)
                            speak.Speak("Highest temperature today will be")
                            speak.Speak(forecast.high)
                            speak.Speak("and Lowest temperature today will be")
                            speak.Speak(forecast.low)
                            z=z+1
                       
                elif("meaning" in filtered_query):
                    dictionary = PyDictionary()
                    wordList=qq.split(' ')
                    print(wordList)
                    for word in wordList:
                        if(word!="meaning" and word!="Meaning"):
                            print(dictionary.meaning(word))
                            speak.Speak(word)
                            speak.Speak(dictionary.meaning(word))

            except sr.UnknownValueError:
                pass
            except sr.RequestError as e:
                pass
            except KeyboardInterrupt:
                pass
except KeyboardInterrupt:
    
    pass
speak.Speak ("Thank You. Hope you got your response.")