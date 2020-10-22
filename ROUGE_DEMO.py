from rouge import Rouge 
import nltk
from textblob import TextBlob
import pyaudio
import speech_recognition as sr
import win32com.client as wincl
import tkinter as tk
speak = wincl.Dispatch("SAPI.SpVoice")

with open('pollution_summary.txt', 'r') as file:
    rouge_summary = file.read().replace('\n', '')

with open('pollution_reference.txt', 'r') as file:
    rouge_reference = file.read().replace('\n', '')
rouge = Rouge()
scores = rouge.get_scores(rouge_summary, rouge_reference)

#print(scores)
speak.Speak("Comparative rouge scores for summary")
print("ROUGE-1 ----------------------------------------------------------------")
speak.Speak("ROUGE 1 SCORES")
print("f-value")
print(round(scores[0]['rouge-1']['f'],4))
print("p-value")
print(round(scores[0]['rouge-1']['p'],4))
print("r-value")
print(round(scores[0]['rouge-1']['r'],4))
speak.Speak("f-value")
speak.Speak(round(scores[0]['rouge-1']['f'],4))
speak.Speak("p-value")
speak.Speak(round(scores[0]['rouge-1']['p'],4))
speak.Speak("r-value")
speak.Speak(round(scores[0]['rouge-1']['r'],4))
#print(scores)

print()
print("ROUGE-2 ----------------------------------------------------------------")
print("f-value")
print(round(scores[0]['rouge-2']['f'],4))
print("p-value")
print(round(scores[0]['rouge-2']['p'],4))
print("r-value")
print(round(scores[0]['rouge-2']['r'],4))
speak.Speak("f-value")
speak.Speak(round(scores[0]['rouge-2']['f'],4))
speak.Speak("p-value")
speak.Speak(round(scores[0]['rouge-2']['p'],4))
speak.Speak("r-value")
speak.Speak(round(scores[0]['rouge-2']['r'],4))


print()
print("ROUGE-3 ----------------------------------------------------------------")
print("f-value")
print(round(scores[0]['rouge-3']['f'],4))
print("p-value")
print(round(scores[0]['rouge-3']['p'],4))
print("r-value")
print(round(scores[0]['rouge-3']['r'],4))

speak.Speak("f-value")
speak.Speak(round(scores[0]['rouge-3']['f'],4))
speak.Speak("p-value")
speak.Speak(round(scores[0]['rouge-3']['p'],4))
speak.Speak("r-value")
speak.Speak(round(scores[0]['rouge-3']['r'],4))