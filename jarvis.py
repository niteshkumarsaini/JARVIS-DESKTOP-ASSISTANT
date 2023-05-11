import win32com.client
import speech_recognition as sp
import wikipedia 
import webbrowser
import os
import random
from datetime import datetime
def speak(audio):
    speaker=win32com.client.Dispatch('SAPI.SpVoice')
    speaker.Speak(audio)
def wish():
    hour=int(datetime.now().hour)
    if(hour>11) and (hour<16):
        speak("Good Afternoon Sir !")
    elif(hour>16) and (hour<24):
        speak("Good Evening Sir !")
    else:
        speak("Good Morning Sir !")
    speak("I am Jarvis, how May I help you")
def takeCommand():
    r=sp.Recognizer()
    with sp.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        r.pause_threshold=1
        print("Listening...")
        audio=r.listen(source,timeout=8,phrase_time_limit=8)
    try:
        print("Recognizing...")
        query=r.recognize_google(audio, language='en-in')
        print(f"User Said : {query}\n")
    except Exception as e:
        # print(e)
        print("Say that Again Please...")
        return "None"
    return query
wish()
while True:
        query=takeCommand().lower()
        if 'wikipedia' in query:
            speak("Searching Wikipedia...")
            query=query.replace('wikipedia','')
            result=wikipedia.summary(query,sentences=2)
            speak("According to Wikipedia")
            print(result)
            speak(result)
        elif 'open youtube' in query:
            webbrowser.open('youtube.com')
        elif 'open google' in query:
            webbrowser.open("google.com")
        elif 'open instagram' in query:
            webbrowser.open("instagram.com")
        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")
        elif 'play music' in query:
            music_directory='G:\\Music Series\\New folder'
            songs=os.listdir(music_directory)
            select_Song=random.randint(0,len(songs))
            os.startfile(os.path.join(music_directory,songs[select_Song]))
        elif 'the time' in query:
            strTime=datetime.now().strftime("%H:%M:%S")
            print(strTime)
            
            speak(f"Sir, the Time is {strTime}")
        elif 'open vs code' in query:
            vs_path="C:\\Users\\Dell\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(vs_path)
        elif 'play video' in query:
            video_directory='G:\\New Video Series'
            videos=os.listdir(video_directory)
            select_video=random.randint(1,len(videos))
            os.startfile(os.path.join(video_directory,videos[select_video]))
        elif 'open study point'in query:
            study_path="D:\\Study Point\\Study Series"
            os.startfile(study_path)
        elif 'close chrome' in query:
            browser="chrome.exe"
            os.system("taskkill /f /im "+browser) 
        elif 'close code' in query:
            os.system("taskkill /f /im "+"Code.exe") 
        elif 'close system' in query:
            os.system("shutdown /s /t 1")
        elif 'open office' in query:
            office_path="C:\\Program Files\\Microsoft Office 15\\root\\office15\\WINWORD.EXE"
            os.startfile(office_path)
        elif 'open excel' in query:
            excel_path="C:\\Program Files\\Microsoft Office 15\\root\\office15\\EXCEL.EXE"
            os.startfile(excel_path)
        elif 'close office' in query:
            os.system("taskkill /f /im "+"WINWORD.EXE")    
        elif 'close excel' in query:
            os.system("taskkill /f /im "+"EXCEL.EXE")
        elif 'exit' in query:
            speak("Thanks Sir for Giving your time.")
            speak("See you soon.")
            break





    




  


