import speech_recognition as sr
import webbrowser
import sys
import pyautogui as pg
import pyttsx3
from datetime import datetime
import win32api


engine = pyttsx3.init()
engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0')
#def talk(words):
    #speak = wincl.Dispatch("SAPI.SpVoice")
    #speak.Speak(words)
#talk("Hello, what you want to open?")

def talk(words):
    engine.say(words)
    engine.runAndWait()
talk("Hello, how I can help?")
win32api.LoadKeyboardLayout('00000409',1)

def command():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("You can talk")
        r.pause_threshold=1
        r.adjust_for_ambient_noise(source, duration=1)
        audio=r.listen(source)
    
    try:
        zadanie=r.recognize_google(audio).lower()
        print("You say = " + zadanie)    
    except sr.UnknownValueError:
        talk("I don't understand you")
        zadanie=command()
    return zadanie

def makeSomething(zadanie):
    if 'open google' in zadanie:
        talk("Opening")
        url='https://www.google.com.ua'
        webbrowser.open(url)
    elif 'open telegram' in zadanie:
        talk("Opening")
        pg.hotkey("winleft", "s")
        pg.PAUSE = 1
        pg.typewrite("Telegram \n", 0.3)
        pg.hotkey("winleft", "up")
        pg.PAUSE = 1
        talk("Opened")
    elif 'telegram' in zadanie:
        talk("Opening")
        pg.hotkey("winleft", "s")
        pg.PAUSE = 1
        pg.typewrite("Telegram \n", 0.3)
        pg.hotkey("winleft", "up")
        pg.PAUSE = 1
        talk("Opened")
    elif 'open spotify' in zadanie:
        talk("Opening")
        pg.hotkey("winleft", "s")
        pg.PAUSE = 1
        pg.typewrite("spotify \n", 0.3)
        pg.hotkey("winleft", "up")
        pg.PAUSE = 1
        talk("Opened")
    elif 'open web spotify' in zadanie:
        talk("Opening")
        url='https://open.spotify.com/'
        webbrowser.open(url)
        pg.PAUSE = 8
        pg.hotkey("space")
        talk("Opened")
    elif 'open soundcloud' in zadanie:
        talk("Opening")
        url='https://soundcloud.com/owenzbs'
        webbrowser.open(url)
        pg.PAUSE = 3
        pg.leftClick()
        pg.hotkey("space")
        talk("Opened")
    elif 'soundcloud' in zadanie:
        talk("Opening")
        url='https://soundcloud.com/owenzbs'
        webbrowser.open(url)
        pg.PAUSE = 3
        pg.leftClick()
        pg.hotkey("space")
        talk("Opened")    
    elif 'time' in zadanie:
        talk("Time is")
        current_datetime = datetime.now()
        talk(current_datetime.hour)
        talk("Hour and")
        talk(current_datetime.minute)
        talk("Minuts")
        talk("Done")
    elif 'thank you' in zadanie:
        talk("Your welcome!")
        sys.exit()
    elif 'stop' in zadanie:
        talk("Yeah, of course")
        sys.exit()
    elif 'goodbye' in zadanie:
        talk("Bye bye")
        sys.exit()
        
        

            
while True:
    makeSomething(command())