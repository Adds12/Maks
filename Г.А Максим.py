import os
import time
import fuzzywuzzy as fuzz
import speech_recognition as sr
import sys
import webbrowser
import time
import win32com.client as wincl

speak = wincl.Dispatch("SAPI.SpVoice")
speak.Rate = 1
speak.Voice = speak.GetVoices().Item(0)
speak.Speak("Приветствую вас, мой повелитель!")
print("Приветствую вас, мой повелитель!")
def command():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( "Вы можете говорить, я вас слушаю" )
        print("Вы можете говорить, я вас слушаю")
        r.pause_threshold=1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)

    try:
        zd = r.recognize_google(audio, language="ru-RU").lower()
        print("Вы сказали:" + zd)
    except sr.UnknownValueError:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( "Я вас не понял" )
        print("Я вас не понял")
        zd = command()

    return zd

def makeSomething(zd):
    if 'вк'.lower() in zd:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( "Я уже открываю" )
        url = 'https://vk.com/feed'
        webbrowser.open(url)
    elif 'Закрой'.lower() in zd:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( "Я уже закрываюсь" )
        sys.exit()
    elif 'youtube'.lower() in zd:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( "Я уже открываю" )
        url = 'https://youtube.com'
        webbrowser.open( url )

    elif 'анекдот 1'.lower() in zd:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( " Я узнала почему в мультике 'Маша и Медведь' не показывают Машиных маму и папу потому что — Они, наверное, уже в дурдоме! ахахаахаха" )
        print(( " Я узнала почему в мультике 'Маша и Медведь' не показывают Машиных маму и папу потому что — Они, наверное, уже в дурдоме! ахахаахаха" ))

    elif 'Анекдот 2'.lower() in zd:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak("Заходит медсестpа в палату. — Больной Иванов, пpоснитесь, ..., ну пpоснитесь же... Больной пpосыпается. — Что такое случилось ? — Я вам снотвоpное пpинесла...")
        print("Заходит медсестpа в палату. — Больной Иванов, пpоснитесь, ..., ну пpоснитесь же... Больной пpосыпается. — Что такое случилось ? — Я вам снотвоpное пpинесла...")
    elif "Как тебя зовут".lower() in zd:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( "Меня зовут Максим, но вы можете называть меня Макс" )

    elif "".lower() in zd:
        speak = wincl.Dispatch( "SAPI.SpVoice" )
        speak.Rate = 1
        speak.Voice = speak.GetVoices().Item( 0 )
        speak.Speak( "Точное время: " )

while True:
    makeSomething(command())