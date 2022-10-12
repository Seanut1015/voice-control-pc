import pyttsx3
import speech_recognition as sr
import webbrowser 
import configparser
import pystray
from PIL import Image
from time import sleep 
from win32api import keybd_event, ShellExecute
from win32con import KEYEVENTF_KEYUP
from pyperclip import copy
from pytube import Search
import sys
from PyQt5 import QtWidgets
from UI import Ui_Form

image = Image.open('.\icon.png')

config = configparser.ConfigParser()
config.read('.\config.ini',encoding="utf-8-sig")

wake_up = config['wake_up']['wake_up']
wake_up_l = wake_up.split(',')

keyword = []
for i in config['keyword']:
    keyword.append(config['keyword'][i])

search = config['keyword']['search']
open_t = config['keyword']['open']
close = config['keyword']['close']
play = config['keyword']['play']
nothing = config['keyword']['nothing']

mic_id = int(config['setting']['mic'])
volume = int(config['setting']['volume'])
volume_f = float(volume/10)

sound = config['speak']['sound']
wish = config['speak']['wish']
response = config['speak']['response']

program = []
path = []
for i in config['user']:
    path.append(config['user'][i])
    program.append(i)

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
rate = engine.getProperty('rate')           
engine.setProperty('rate', rate-50)        
engine.setProperty('volume',volume_f)    

run = 0

r = sr.Recognizer()
r.energy_threshold = 4000
r.pause_threshold = 0.5
r.dynamic_energy_threshold = 1  

wake = 0

class MyWindow(QtWidgets.QWidget,Ui_Form):
    x = 0
    def __init__(self,parent=None):
        super(MyWindow,self).__init__(parent)
        self.setupUi(self)
        self.lineEdit.setText(wake_up)
        self.lineEdit_1.setText(search)
        self.lineEdit_2.setText(open_t)
        self.lineEdit_3.setText(close)
        self.lineEdit_4.setText(play)
        self.lineEdit_5.setText(nothing)
        self.lineEdit_6.setText(wish)
        self.lineEdit_7.setText(response)
        self.spinBox.setValue(volume)
        self.spinBox.setRange(1,10)
        self.radioButton_4.setChecked(1)
        self.radioButton_4.setDisabled(1)
        self.groupBox.setDisabled(1)
        x = 1

    def change(self):
        config['wake_up']['wake_up'] = self.lineEdit.text()
        config['keyword']['search'] = self.lineEdit_1.text()
        config['keyword']['open'] = self.lineEdit_2.text()
        config['keyword']['close'] = self.lineEdit_3.text()
        config['keyword']['play'] = self.lineEdit_4.text()
        config['keyword']['nothing'] = self.lineEdit_5.text()
        config['speak']['wish'] = self.lineEdit_6.text()
        config['speak']['response'] = self.lineEdit_7.text()
        config['setting']['volume'] = self.spinBox.text()
        user = config['user']
        x=self.lineEdit_8.text()
        if self.label_12.text() != 'path' and x != ' ' and len(x)!=0 :
            user[x] = self.label_12.text()
        with open('.\config.ini', 'w',encoding="utf-8-sig") as configfile:
            config.write(configfile)
        print(x)

    def reset(self):
        self.lineEdit.setText(wake_up)
        self.lineEdit_1.setText(search)
        self.lineEdit_2.setText(open_t)
        self.lineEdit_3.setText(close)
        self.lineEdit_4.setText(play)
        self.lineEdit_5.setText(nothing)
        self.lineEdit_6.setText(wish)
        self.lineEdit_7.setText(response)
        

    def open_file(self):
        filePath , filterType = QtWidgets.QFileDialog.getOpenFileNames()
        filePath = ''.join(filePath)
        self.label_12.setText('"'+filePath+'"')

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

        elif open_t in query:
            query=query[len(open_t):]
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
        
            
    if any(i in query for i in wake_up_l) and wake==0:
        wake = 1
       
def on_clicked(icon, item):
    global run
    if str(item) == '設定':
        app = QtWidgets.QApplication(sys.argv)
        ui=MyWindow()
        ui.setWindowTitle("VCP")
        ui.show()
        sys.exit(app.exec_())

    if str(item) == '結束程式':
        run = 0
        icon.stop()

def speak(audio):   
    engine.say(audio)    
    engine.runAndWait() 

def takeCommand():
    with sr.Microphone(device_index = mic_id) as source:     
        r.adjust_for_ambient_noise(source,0.5)
        if wake == 1:
            speak('Hi')
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language='zh') 
            #print(f"User said: {query}\n")   
        except Exception as e:
            return "None" 
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

icon = pystray.Icon ('VCP', image,"VCP", menu=pystray.Menu(
    pystray.MenuItem("設定", on_clicked),
    pystray.MenuItem("結束程式", on_clicked),
))

if __name__ == '__main__':
    wishMe() 
    run = 1 
    icon.run(voice)
