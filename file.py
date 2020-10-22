
import re
import nltk
import heapq
import random
from textblob import TextBlob
import pyaudio
import wave
import speech_recognition as sr
import win32com.client as wincl
import tkinter as tk
speak = wincl.Dispatch("SAPI.SpVoice")


def audioToText():
    speak.Speak("Recording started")
    FORMAT = pyaudio.paInt16
    CHANNELS = 2
    RATE = 44100
    CHUNK = 1024
    RECORD_SECONDS = 5
    WAVE_OUTPUT_FILENAME = "output.wav"   #filename different to prerecorded audio
    
    audio = pyaudio.PyAudio()
    
    # start Recording
    stream = audio.open(format=FORMAT, channels=CHANNELS,
                    rate=RATE, input=True,
                    frames_per_buffer=CHUNK)
    
    print("recording...")
    frames = []
    
    for i in range(0, int(RATE / CHUNK * RECORD_SECONDS)):
        data = stream.read(CHUNK)
        frames.append(data)
    
    speak.Speak("recording ended")
    
    
    stream.stop_stream()
    stream.close()
    audio.terminate()
    waveFile = wave.open(WAVE_OUTPUT_FILENAME, 'wb')
    waveFile.setnchannels(CHANNELS)
    waveFile.setsampwidth(audio.get_sample_size(FORMAT))
    waveFile.setframerate(RATE)
    waveFile.writeframes(b''.join(frames))
    waveFile.close()
    r=sr.Recognizer()
    with sr.WavFile("file_Audio.wav") as source:     
        audio = r.record(source) 
    speak.Speak("You said-")
    try:
        str=r.recognize_google(audio)
        print(str)
        speak.Speak(str)
    except:
        speak.Speak("Connection problem or Voice unclear")
        print("voice unclear")
    with open("inputFile.txt",'w') as ir:
        ir.write(str)
        
        
'''win = tk.Tk() 
win.title('Text Editor')
win.geometry('500x500')  
button = tk.Button(win, text="Login", bg="green")#, command=attempt_log_in)
button.grid(row=5, column=2)
button.bind('<Button-1>', audioToText)'''

audioToText()
    
#PREPROCESSING TEXT DATA
with open("input_File.txt") as fp:
    text = fp.read() 
#text=re.sub(r'\s+'," ",text)
text=text.lower()
ages=nltk.sent_tokenize(text)
lines=text.split('.')
#print(lines)
for i in range(len(lines)):
    lines[i]=lines[i].strip('\n')
print(lines)
stop_words=nltk.corpus.stopwords.words('english')
 
 
#COMBINING RELATIVE SENTENCES
l=len(lines)
conjuction=['and','since','therefore','while','nor','or','but','so','yet','this']
for i in range(1,l):
    continous=0
    #print(i,l,len(lines))
    if i>=len(lines):
        break
    lines[i].replace('\n','')
    for conjucts in conjuction:
        if(lines[i].startswith(conjucts)):
            continous=1
            #print (lines[i])
            break
    if continous==1:
        lines[i-1]+=' '+lines[i]
        del lines[i]


#CALCULATING COUNT FOR EVERY WORD
words_count={}
for line in lines:
    for word in nltk.word_tokenize(line):
        if word not in stop_words:
            if word not in words_count:
                words_count[word]=1
            else:
                words_count[word]=words_count[word]+1
                
for key in words_count.keys():
    words_count[key]=words_count[key]/max(words_count.values())
    
    
#RANKING SENTENCES BASED ON WORDCOUNT RATIOS
sentence_score={}
for sentence in lines:
    for word in nltk.word_tokenize(sentence):
        if(word in words_count.keys()):
            if sentence not in sentence_score:
                sentence_score[sentence]=words_count[word]
            else:
                sentence_score[sentence]+=words_count[word]
                

#CALCULATING TOP 5 SENTENCES           
summary=heapq.nlargest(5,sentence_score,key=sentence_score.get)
speak.Speak("Conversation Summary is")

#POST PROCESSING 
for i in range(len(summary)):
    #print(para)
    summary[i]=summary[i].replace('\n','')
speak.Speak(summary)
print("The summary derived is-")
print(summary)
blob=TextBlob(text)
summary_nouns=list()
for word,tag in blob.tags:
    if tag=='NN':
        summary_nouns.append(word.lemmatize())
        
print("this conversation was about ...")
for item in random.sample(summary_nouns,5):
    word=item
    print(word)
    