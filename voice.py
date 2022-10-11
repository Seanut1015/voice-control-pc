import pyttsx3
import speech_recognition as sr
import webbrowser 
import configparser
import pystray
import pyaudio
from PIL import Image
from time import sleep 
from win32api import keybd_event, ShellExecute
from win32con import KEYEVENTF_KEYUP
from pyperclip import copy
from pytube import Search

image = Image.open('.\icon.png')
config = configparser.ConfigParser()
config.read('.\config.ini',encoding="utf-8-sig")

wake_up = config['wake_up']['wake_up']
wake_up = wake_up.split(',')

keyword = []
for i in config['keyword']:
    keyword.append(config['keyword'][i])

search = config['keyword']['search']
open = config['keyword']['open']
close = config['keyword']['close']
play = config['keyword']['play']
nothing = config['keyword']['nothing']

mic_id = int(config['setting']['mic'])
volume = float(config['setting']['volume']) 

sound = config['speak']['sound']
wish = config['speak']['wish']
response = config['speak']['response']

program = []
path = []
for i in config['user']:
    path.append(config['user'][i])
    program.append(i)

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices') #gets you the details of the current voice
engine.setProperty('voice', voices[1].id)  # 0-male voice , 1-female voice
rate = engine.getProperty('rate')           #設定語音速率
engine.setProperty('rate', rate-50)         #設定語音速率

engine.setProperty('volume',volume)         #設定音量
run = 0

r = sr.Recognizer()
r.energy_threshold = 4000
r.pause_threshold = 0.5
r.dynamic_energy_threshold = 1  

wake = 0

def funtion(query):
    global wake
    if wake == 1 or wake == 2:
        if any(i in query for i in keyword):
            wake = 0
        else:
            wake = 2

        if search in query:
            query=query[len(search):]
            webbrowser.get('windows-default').open_new('https://www.google.com/search?q='+query)

        elif play in query:
            query=query[len(play):]
            s = Search(query)
            hash= str(s.results[0])
            hash=hash[41:].rstrip('>')
            webbrowser.get('windows-default').open_new('https://www.youtube.com/watch?v='+hash)

        elif open in query:
            query=query[len(open):]
            keybd_event(0x5B,0,0,0)
            keybd_event(0x5B,0,KEYEVENTF_KEYUP,0)
            sleep(0.1)
            copy(query)
            keybd_event(0x11,0,0,0)
            keybd_event(0x56,0,0,0)
            keybd_event(0x56,0,KEYEVENTF_KEYUP,0)
            keybd_event(0x11,0,KEYEVENTF_KEYUP,0)

            sleep(0.1)
            keybd_event(0x0D,0,0,0)
            keybd_event(0x0D,0,KEYEVENTF_KEYUP,0)

        elif close in query:
            keybd_event(0x12,0,0,0)
            keybd_event(0x73,0,0,0)
            keybd_event(0x73,0,KEYEVENTF_KEYUP,0)
            keybd_event(0x12,0,KEYEVENTF_KEYUP,0)

        elif any(i in query for i in program):
            for i in range(len(program)):
                if query == program[i]:
                    print(path[i])
                    ShellExecute(0, 'open', str(path[i]), '','',0)
        
        elif nothing in query:
            pass 

        if wake == 0:
            speak(response)
        
            
    if any(i in query for i in wake_up) and wake==0:
        wake = 1
       
def on_clicked(icon, item):
    global run
    if str(item) == '結束程式':
        run = 0
    if str(item) == '結束程式':
        run = 0
    if str(item) == '結束程式':
        run = 0
        icon.stop()

def speak(audio):   
    engine.say(audio)    
    engine.runAndWait() 

def takeCommand():
    with sr.Microphone(device_index = mic_id) as source:     
        #print(r.energy_threshold)
        r.adjust_for_ambient_noise(source,0.5)
        #print("Listening...")
        if wake == 1:
            speak('Hi')
        audio = r.listen(source)
        try:
            #print("Recognizing...")   
            query = r.recognize_google(audio, language='zh') #Using google for voice recognition.
            print(f"User said: {query}\n")  #User query will be printed.    
        except Exception as e:
            #pass
            #print(e) # use only if you want to print the error!
            print("Say that again please...")   #Say that again will be printed in case of improper voice 
            return "None" #None string will be returned
        return query

def wishMe():
    speak(wish)    

def voice(icon):
    global wake
    icon.visible = True
    while run == 1:  
        query = takeCommand().lower()
        if query == 'none':
            if wake == 1:
                wake = 2
            continue
        else:
            funtion(query)

        if '結束' in query:
            icon.stop()
            break

wishMe() 
icon = pystray.Icon ('Sean', image,"AI Helper", menu=pystray.Menu(
    pystray.MenuItem("結束程式", on_clicked),
))

run = 1 
icon.run(voice)
