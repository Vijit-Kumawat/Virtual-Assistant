import pyttsx3 
import speech_recognition as sr  
import datetime
import wolframalpha
import wikipedia 
import webbrowser
import os        
import time
import mouse
import pyautogui
import pydirectinput
import wmi
import ctypes
from urllib.request import urlopen
import subprocess
import win32com.client
import win32gui, win32con
import winshell
import random   
import keyboard
import pyperclip
import pyjokes 
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from playsound import playsound

count = 0
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
newVoiceRate = 145
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate',newVoiceRate)

def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

def wishMe():
    
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")

def tellDay():
    day = datetime.datetime.today().weekday() + 1
    Day_dict = {1: 'Monday', 2: 'Tuesday', 
                3: 'Wednesday', 4: 'Thursday', 
                5: 'Friday', 6: 'Saturday',
                7: 'Sunday'}
    today = datetime.date.today()  
    strTime = datetime.datetime.now().strftime("%I:%M %p")
    speak(today)
    if day in Day_dict.keys():
        day_of_the_week = Day_dict[day]
        speak(day_of_the_week)
    speak(f",'time is ',{strTime}")

def takeCommand():
    #     It takes microphone input form the user and returns string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.5
        # audio = r.listen(source)
        audio = r.listen(source,0,2)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-US')
        print(f"User said: {query}")

    except Exception as e:
        print(e)
        return "None"
    return query

if __name__ == '__main__':
    playsound("starting.mp3")
    wishMe()
    tellDay()
    while True:
        query = takeCommand().lower()
        strings_start = {'jarvis', 'jarvis', 'hi jarvis', "hey jarvis", 'buddy', 'hey buddy','hi buddy' ,'suno', 'listen', 'wake up'}
        while query in strings_start:
            playsound("activate.mp3")
            while True:
                query = takeCommand().lower()
                screenshot_count = 0
                strings_hello = {'jarvis', 'badi', 'buddy', 'hey buddy', 'hello jarvis', 'hi jarvis', "hey jarvis"}
                strings_what_doing = {'buddy what are you doing', 'what are you doing', 'badi kya kar rahe ho'}
                strings_shut_up = {'shut up', 'chup ho jao'}
                strings_stop_listening = {'stop listening', "don't listen", 'kuchh nahi', 'suno mat'}
                Strings_help = {'how can you help me', 'how you can help me' , 'how you can help', 'what you can do', 'what can you do'}
                strings_how_are_you = {"how are you", "jarvis how are you"}
                strings_im_fine = {"i'm fine", "i am fine"}

                # Actions
                if "wikipedia" in query:
                    while True:
                        try:
                            speak('Searching Wikipedia...')
                            query = query.replace("wikipedia", "")
                            results = wikipedia.summary(query, sentences=2)
                            speak("According to wikipedia")
                            print(results)
                            speak(results)
                            break
                        except Exception:
                            speak('Sorry no result, Please be clear')
                            break
                        
                elif 'shutdown' in query:
                    speak("Do you really want to shutdown say 'yes'")
                    while True:
                        query = takeCommand()
                        if query == 'yes':
                            speak("Executing Shutdown")
                            os.system("shutdown /s /t 1")
                        else:
                            break                

                elif "game controls" in query: 
                    speak("let's go sir i am here to assist")
                    while True:
                        query = takeCommand().lower()
                        count=0
                        if "reload" in query:
                            try:
                                keyboard.press('r')
                                speak("done with reload")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "bang bang" in query:
                            try:
                                time.sleep(0.1)
                                while count<=10:
                                    pydirectinput.press('`')
                                    pyautogui.write("toolup")
                                    pydirectinput.press('enter')
                                    count+=1
                                speak("Now rock and roll")  
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')  

                        elif "remove police" in query:
                            try:
                                time.sleep(0.1)
                                while count<=6:
                                    pydirectinput.press('`')
                                    pyautogui.write("lawyerup")
                                    pydirectinput.press('enter')
                                    count+=1
                                speak("all stars removed")   
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')  

                        elif "superman mod" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("turtle")
                                pydirectinput.press('enter')

                                pydirectinput.press('`')
                                pyautogui.write("painkiller")
                                pydirectinput.press('enter')

                                pydirectinput.press('`')
                                pyautogui.write("catchme")
                                pydirectinput.press('enter')

                                pydirectinput.press('`')
                                pyautogui.write("hoptoit")
                                pydirectinput.press('enter')

                                pydirectinput.press('`')
                                pyautogui.write("gotgills")
                                pydirectinput.press('enter')

                                speak("Super man mode activated")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA') 
                
                        elif "power up" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("powerup")
                                pydirectinput.press('enter')
                                speak("power recharged")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA') 
                        
                        elif "hulk hands" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("hothands")
                                pydirectinput.press('enter')
                                speak("hulk hands activated")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA') 
                        
                        elif "explosive bullets" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("highex")
                                pydirectinput.press('enter')
                                speak("explosive bullets activated")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA') 
                        
                        elif "fire bullets" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("incendiary")
                                pydirectinput.press('enter')
                                speak("explosive bullets activated")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')


                        elif "health" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("turtle")
                                pydirectinput.press('enter')
                                speak("health full")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "painkiller" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("painkiller")
                                pydirectinput.press('enter')
                                speak("5 minute invincibility")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')
                        
                        elif "run faster" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("catchme")
                                pydirectinput.press('enter')
                                speak("run faster activated")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "super jump" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("hoptoit")
                                pydirectinput.press('enter')
                                speak("hulk jump activated")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "i need a helicopter" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("buzzoff")
                                pydirectinput.press('enter')
                                speak("you helicopter is waiting")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "i need a car" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("comet")
                                pydirectinput.press('enter')
                                speak("you car is waiting")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "dirt bikes" in query:
                            try:
                                time.sleep(0.1)
                                query = query.replace("dirt bikes ","")
                                count=0
                                try:
                                    number = int(query)
                                    number = number-1
                                    while count<=number:
                                        pydirectinput.press('`')
                                        pyautogui.write("offroad")
                                        pydirectinput.press('enter')
                                        count+=1
                                    speak("Bikes are here sir")
                                except Exception:
                                    speak("wrong input")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "give me a drink" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("liquor")
                                pydirectinput.press('enter')
                                speak("i am too high")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "change weather" in query:
                            try:
                                time.sleep(0.1)
                                pydirectinput.press('`')
                                pyautogui.write("makeitrain")
                                pydirectinput.press('enter')
                                speak("weather changed")
                            except Exception as e:
                                print('PROBLEM SOLVED HAHA')

                        elif "stop" in query:
                            time.sleep(0.1)
                            speak("game controls off")
                            if True==True:
                                break

                elif "type" in query: 
                    query = query.replace("type ","")                            
                    pyautogui.typewrite(query,interval=0.01)                          
                    pyautogui.typewrite(" ")    
                        
                elif query in strings_stop_listening:
                    playsound("stop_listening.mp3")
                    break

                elif "open youtube" in query:
                    webbrowser.open('https://www.youtube.com/')
                    speak("okay sir, Opening Youtube!")             

                elif "on youtube" in query:
                    try:
                        query = query.replace(" on youtube", "")
                        driver = webdriver.Chrome(executable_path=r"C:/Users/vijit/Downloads/chromedriver")
                        driver.maximize_window()
                        driver.get("https://www.youtube.com/results?search_query=" + query)
                        firstvideo = driver.find_element_by_xpath('//*[@id="contents"]/ytd-video-renderer[1]')
                        firstvideo.click()
                        speak("here")
                    except Exception as e:
                        print(e)                   

                elif "close chrome" in query:
                    os.system('TASKKILL /F /IM chrome.exe')
                    speak("okay sir, closing chrome!")

                elif "open google" in query:
                    webbrowser.open_new('https://www.google.com/')
                    speak("okay sir, opening google!")

                elif "open stack overflow" in query:
                    webbrowser.open('https://www.stackoverflow.com/')

                elif "play music" in query:
                    music_dir = 'E:\\Videoder'
                    songs = os.listdir(music_dir)
                    print(songs)
                    a = random.randint(1, 19)
                    os.startfile(os.path.join(music_dir, songs[a]))
                elif "stop music" in query:
                    os.system('TASKKILL /F /IM Music.UI.exe')
                    speak("okay sir, closing music!")

                elif ("the time" in query) or ("time kya hua hai" in query):
                    strTime = datetime.datetime.now().strftime("%I:%M %p")
                    speak(f"sir, the time is {strTime}")

                elif "open subject notes" in query:
                    codePath = "D:\\Notes"
                    os.startfile(codePath)
                    speak("here!")

                elif "open whatsapp" in query:
                    subprocess.call('C:\\Users\\vijit\\AppData\\Local\\WhatsApp\\WhatsApp.exe')
                    speak("okay sir, opening WhatsApp")
                elif "close whatsapp" in query:
                    os.system('TASKKILL /F /IM WhatsApp.exe')
                    speak("okay sir, closing WhatsApp")

                elif "open vs code" in query:
                    os.startfile("C:\\Users\\vijit\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe")
                    speak("okay sir, opening code")
                elif "close vs code" in query:
                    os.system('TASKKILL /F /IM Code.exe')
                    speak("okay sir, closing code")

                elif "open pycharm" in query:
                    os.startfile("C:\\Program Files\\JetBrains\\PyCharm Community Edition 2020.3.2\\bin\\pycharm64.exe")
                    speak("okay sir, opening pai charm")
                elif "close pycharm" in query:
                    os.system('TASKKILL /F /IM pycharm64.exe')
                    speak("okay sir, closing pai charm")

                elif "open video" in query:
                    os.startfile("C:\\Program Files (x86)\\VideoLAN\\VLC\\vlc.exe")
                    speak("okay sir, opening VLC media player")
                elif "close video" in query:
                    os.system('TASKKILL /F /IM vlc.exe')
                    speak("okay sir, closing VLC")

                elif "open photoshop" in query:
                    psApp = win32com.client.Dispatch("Photoshop.Application")
                    speak("okay sir, opening photoshop")
                elif "close photoshop" in query:
                    os.system('TASKKILL /F /IM photoshop.exe')
                    speak("okay sir, closing Photoshop")

                elif "open discord" in query:
                    os.system("C:\\Users\\vijit\\AppData\\Local\\Discord\\Update.exe --processStart Discord.exe")
                    speak("okay sir, opening discord")
                elif "close discord" in query:
                    os.system('TASKKILL /F /IM Discord.exe')
                    speak("okay sir, closing discord")

                elif "open teams" in query:
                    os.system("C:\\Users\\vijit\\AppData\\Local\\Microsoft\\Teams\\Update.exe --processStart Teams.exe")
                    speak("okay sir, opening teams")
                elif "close teams" in query:
                    os.system('TASKKILL /F /IM teams.exe')
                    speak("okay sir, closing teams")

                elif "open spotify" in query:
                    os.system("spotify")
                    speak("okay sir, opening spotify")
                elif "close spotify" in query:
                    os.system('TASKKILL /F /IM spotify.exe')
                    speak("okay sir, closing spotify")

                elif "open notes" in query:
                    webbrowser.open('https://keep.google.com/#home')
                    speak("okay, opening Notes")

                elif "search on google" in query:
                    query = query.replace("search on google", "")
                    url = "http://google.com/search?q=" + query
                    webbrowser.get().open(url)
                    speak("Here is what i found" + query)

                elif "find location" in query:
                    try:
                        driver = webdriver.Chrome(executable_path='C:/Users/vijit/Downloads/chromedriver')
                        driver.maximize_window()
                        query = query.replace("find location ", "")
                        location = query
                        driver.get("https://www.google.nl/maps/place/" + location + "")
                        time.sleep(2)
                        search_button = driver.find_element_by_class_name('searchbox-searchbutton-container')
                        search_button.click()
                    except Exception as e:
                        print("Try again")

                elif "where is" in query:
                    try:
                        query = query.replace("where is ","")
                        driver = webdriver.Chrome(executable_path = "C:/Users/vijit/Downloads/chromedriver")
                        driver.maximize_window()
                        
                        location = query
                        driver.get("https://www.google.nl/maps/place/" + location + "")
                        time.sleep(2)
                        search_button = driver.find_element_by_id('searchbox-searchbutton')
                        search_button.click()
                        time.sleep(3)
                        try:
                            try:
                                print('entered in first-------')
                                info = driver.find_element_by_class_name('ZwbPmb-text').text
                                print('second code is working-------')
                                pyperclip.copy(info)
                                from_clipboard = pyperclip.paste()
                                speak(from_clipboard)
                            except Exception as e:               
                                major = driver.find_elements_by_class_name("TkWrSb-mXWxJe")
                                info = driver.find_element_by_xpath('.//*[@id="pane"]/div/div[1]/div/div/div[7]/div[1]/span/span[1]').text
                                print('yes working')
                                pyperclip.copy(info)
                                from_clipboard = pyperclip.paste()
                                speak(from_clipboard)
                        except Exception as e:
                            print('vijit')
                    except Exception as e:
                        print(e)

                elif "cut this" in query:
                    keyboard.press_and_release("ctrl+x")
                    speak("ok")

                elif "copy it" in query:
                    keyboard.press_and_release("ctrl+c")
                    speak("ok")
                
                elif "select all" in query:
                    keyboard.press_and_release("ctrl+a")
                    
                elif "paste it here" in query:
                    keyboard.press_and_release("ctrl+v")         

                elif "delete this" in query:
                    keyboard.press_and_release("delete")
                    speak("ok")

                elif "close this" in query:
                    keyboard.press_and_release("alt+f4")

                elif "open this" in query:
                    keyboard.press_and_release("enter")
                
                elif "save it" in query:
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("s")
                    pyautogui.keyUp("ctrl")
                    speak("saved")
                
                elif "save as" in query:
                    pyautogui.keyDown("ctrl")
                    pyautogui.keyDown("shift")
                    pyautogui.keyDown("s")

                elif "don't save" in query:
                    pyautogui.keyDown("alt")
                    pyautogui.press("f4")
                    pyautogui.keyUp("alt")
                    pyautogui.press('right')
                    pyautogui.press('enter')
                    speak("okay")
                
                elif "open my computer" in query:
                    keyboard.press_and_release("win+e")

                elif "make a new folder" in query:
                    query = query.replace("make a new folder ", "")
                    keyboard.press_and_release("ctrl+shift+n")
                    pyautogui.typewrite(query)

                elif "create folders" in query:
                    speak("how many?")
                    numbers = takeCommand()
                    numbers = numbers.replace("folders","")
                    a=1
                    try:
                        while a<=int(numbers):
                            speak("what name you want?")
                            query = takeCommand()
                            keyboard.press_and_release("ctrl+shift+n")
                            pyautogui.typewrite(query)
                            a+=1

                    except Exception:
                        speak("Wrong input")

                elif "read this" in query:
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("c")  
                    pyautogui.keyUp("ctrl")
                    from_clipboard = pyperclip.paste()
                    speak(from_clipboard)

                elif "take a screenshot" in query:
                    while True:
                        screenshot_count += 1
                        query = screenshot_count
                        a = os.path.exists('D:\\screenshot\\' + str(query) + '.png')
                        if a == False:
                            pyautogui.screenshot('D:\\screenshot\\' + str(query) + '.png')
                            break
                        else:
                            print("The file already exists")
                    speak("taken!")                

                elif "screenshot folder" in query and "show screenshot" in query:
                    os.startfile("D:\screenshot")
                    speak("okay!")
                
                elif "show desktop" in query:
                    keyboard.press_and_release("win+d")

                elif "repeat after me" in query:
                    speak("Okay, sir")
                    while True:  
                        query = takeCommand().lower()
                        speak(query)
                        if "stop" in query:
                            speak("okay, sir")
                            break
                        
                elif "w3schools" in query:
                    query = query.replace("w3schools ","")
                    query = query.replace(" ","")
                    url = "https://www.w3schools.com/"+ query + "/"
                    webbrowser.get().open(url)
                    speak("here we go!")

                elif "open website" in query:
                    query = query.replace("open website ", "")
                    query = query.replace(" ","")
                    query = query.replace("dot",".")
                    url = "www." + query
                    webbrowser.get().open(url)
                    speak("opening" + query)

                elif "show history" in query:
                    keyboard.press_and_release("ctrl+h")
                
                elif "close tab" in query or "closed tab" in query:
                    keyboard.press_and_release("ctrl+w")

                elif "next" in query:
                    pyautogui.press("right")
                
                elif "previous" in query:
                    pyautogui.press("left")
                
                elif "stop" in query or "play" in query:
                    pyautogui.press("playpause")
  
                elif "increase volume" in query:  
                    x=0
                    count = 10
                    times = int(count)
                    while x<=10:
                        pyautogui.press("volumeup")
                        x+=1
               
                elif "decrease volume" in query:  
                    x=0
                    count = 10
                    times = int(count)
                    while x<=10:
                        pyautogui.press("volumedown")
                        x+=1
                           
                elif "mute" in query:
                    pyautogui.press("volumemute")

                elif "take a print" in query:
                    keyboard.press_and_release("ctrl+p")
                    speak("here!")

                elif "page up" in query:
                    pyautogui.press("pageup")

                elif "page down" in query:
                    pyautogui.press("pagedown")

                elif "go to the top" in query or "go at the top" in query:
                    pyautogui.press("pageup",presses=40)
                elif "go to the bottom" in query or "go at the bottom" in query:
                    pyautogui.press("pagedown",presses=40)
                
                elif "home" in query:
                    pyautogui.press("home")

                elif "end" in query:
                    pyautogui.press("end")

                elif "reload" in query:
                    keyboard.press_and_release("ctrl+r")

                elif "minimise" in query:
                    hwnd = win32gui.GetForegroundWindow()
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                
                elif "maximize" in query:
                    hwnd = win32gui.GetForegroundWindow()
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

                elif "calculator" in query:
                    from subprocess import call
                    call(["calc.exe"])     

                elif "calculate" in query:
                    try:
                        app_id = "JHQGK6-AH87QKWX2Q"
                        client = wolframalpha.Client(app_id)
                        indx = query.lower().split().index('calculate')
                        query = query.split()[indx + 1:]
                        res = client.query(' '.join(query))
                        answer = next(res.results).text
                        print("The answer is " + answer)
                        speak("The answer is " + answer)
                    except Exception:
                        speak("wrong Input")
                    
                elif "find word" in query:
                    query = query.replace("find word ", "")
                    keyboard.press_and_release("ctrl+f")
                    pyautogui.typewrite(query)
                    speak("here")

                elif "full brightness" in query:
                    brightness = 100 # percentage [0-100] For changing thee screen 
                    c = wmi.WMI(namespace='wmi')
                    methods = c.WmiMonitorBrightnessMethods()[0]    
                    methods.WmiSetBrightness(brightness, 0)
                elif "half brightness" in query:
                    brightness = 50 # percentage [0-100] For changing thee screen 
                    c = wmi.WMI(namespace='wmi')
                    methods = c.WmiMonitorBrightnessMethods()[0]    
                    methods.WmiSetBrightness(brightness, 0)
                elif "low brightness" in query:
                    brightness = 0 # percentage [0-100] For changing thee screen 
                    c = wmi.WMI(namespace='wmi')
                    methods = c.WmiMonitorBrightnessMethods()[0]    
                    methods.WmiSetBrightness(brightness, 0)
                elif "brightness" in query:
                    # speak("Give the brightness value between 1 to 100:")
                    # query = takeCommand()
                    query = query.replace("brightness","")
                    try:
                        brightness = int(query) # percentage [0-100] For changing thee screen 
                        c = wmi.WMI(namespace='wmi')
                        methods = c.WmiMonitorBrightnessMethods()[0]    
                        methods.WmiSetBrightness(brightness, 0)
                    except Exception:
                        speak("wrong input")

                elif 'lock window' in query:
                       speak("locking the device")
                       ctypes.windll.user32.LockWorkStation()  

                elif 'empty recycle bin' in query:
                    if True:
                        try:
                            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                            speak("Recycle Bin Recycled")
                        except Exception:
                            speak("Already Empty")
                    

                elif "what is" in query or "who is" in query:
                    client = wolframalpha.Client("JHQGK6-AH87QKWX2Q")
                    res = client.query(query)
                    try:
                        print (next(res.results).text)
                        speak (next(res.results).text)
                    except StopIteration:
                        print ("No results")

                elif "write a note" in query:
                    speak("What should i write, sir")
                    note = takeCommand()
                    file = open('jarvis.txt', 'w')
                    speak("Sir, Should i include date and time")
                    snfm = takeCommand()
                    if 'yes' in snfm or 'sure' in snfm:
                        strTime = datetime.datetime.now().strftime("%H:%M")
                        file.write(strTime)
                        file.write(" :- ")
                        file.write(note)
                    else:
                        file.write(note)

                elif "show note" in query or "show notes" in query:
                    speak("Showing Notes")
                    file = open("jarvis.txt", "r")
                    print(file.read())
                    reading = str(file.read())
                    speak(reading)                

                elif "inspect menu" in query:
                    keyboard.press_and_release("ctrl+shift+i")
                    speak("here")

                elif "delete line" in query:
                    keyboard.press_and_release("ctrl+shift+k")

                elif "open designer" in query:
                    os.startfile("C:\Program Files (x86)\Qt Designer\designer.exe")
                    speak("Opening Designer")

                elif "show time table" in query:
                    os.startfile("D:\\Jarvis\\timetable.png")
                    speak("here")

                elif "search on amazon" in query:             
                    try:
                        speak("what you wanna search")
                        query = takeCommand().lower()
                        url="https://www.amazon.in/s?k="+query
                        webbrowser.open(url)
                    except Exception:
                        print("wrong input")
                        
                elif "start spamming" in query:
                    query = query.replace("start spamming ","")               
                    speak("Starting ,Automation engine")         
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("c")  
                    pyautogui.keyUp("ctrl")
                    a = pyperclip.paste()
                    try:
                        query = int(query)
                        while query > 0:
                            query-=1
                            pyautogui.typewrite(a)
                            pyautogui.press("enter")
                    except Exception as e:
                        print(e)
                        speak('wrong Input')

                elif "nightmare" in query:
                    speak("ok")
                    try:
                        codePath = "Different things\start_spamming.py"
                        os.startfile(codePath)
                    except Exception as e:
                        print(e)
                        speak('wrong Input')

                elif "tell me a joke" in query:
                    speak(pyjokes.get_joke())

                elif "search" in query:
                    try:
                        pyautogui.press("win")
                        query = query.replace("search","")                            
                        pyautogui.typewrite(query,interval=0.01)                          
                        pyautogui.typewrite(" ")
                        pyautogui.press("enter")
                    except Exception as e:
                        print(e)
                        speak("oops")
                        
                elif "close everything" in query:
                    count=0
                    try:
                        while count <= 10:
                            pyautogui.hotkey('ctrl' ,'alt', 'tab')
                            pyautogui.hotkey('delete')
                            count+=1
                        pyautogui.hotkey('esc')
                        
                    except Exception as e:
                        print(e)
                        speak("something went wrong")

                elif "properties" in query:
                    try:
                        pyautogui.hotkey('alt', 'enter')
                    except Exception as e:
                        print(e)
                        speak('oops')

                elif "time to sleep" in query:
                    speak("signing out Jarvis Program")
                    os.system("taskkill /f /im pyw.exe")

                elif "comment this" in query:
                    try:
                        pyautogui.hotkey('ctrl','/')
                    except Exception as e:
                        speak("oops")
                        print(e)

                elif "local disk c" in query:
                    os.startfile("c:")

                elif "local disk d" in query:
                    os.startfile("d:")

                elif "local disk e" in query:
                    os.startfile("e:")

                elif "open" in query:
                    try:
                        keyboard.press_and_release("ctrl+F")
                        query = query.replace("open ","")                            
                        pyautogui.typewrite(query,interval=0.01)                          
                        # pyautogui.typewrite(" ")
                        pyautogui.press("enter")
                    except Exception as e:
                        print(e)
                        speak("oops")
                        
                elif "download" in query:
                    os.startfile(r"C:/Users/vijit/Downloads")
                    hwnd = win32gui.GetForegroundWindow()
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    speak("here!") 

                elif "back back back" in query:
                    pyautogui.press("backspace")
                    pyautogui.press("backspace")
                    pyautogui.press("backspace")

                elif "back back" in query:
                    pyautogui.press("backspace")    
                    pyautogui.press("backspace")    

                elif "back" in query:
                    pyautogui.press("backspace")

                elif "task manager" in query:
                    try:
                        keyboard.press_and_release("ctrl+shift+esc")
                        speak("here")
                    except Exception as e:
                        speak("oops")
                        print(e)

                elif "press" in query:
                    query=query.replace('press ','')
                    print(query)
                    try:
                        keyboard.press(query)
                        speak(query)
                    except Exception as e:
                        print("oops")
                        # print(e)

                elif "click" in query:
                    try:
                        # pyautogui.typewrite(query,interval=0.01)                          
                        query = query.replace("click ","")                            
                        keyboard.press_and_release("ctrl+f")
                        pyautogui.typewrite(query,interval=0.01)                          
                        pyautogui.typewrite("")
                        pyautogui.press("esc")
                        pyautogui.press("enter")
                    except Exception as e:
                        print(e)
                        speak("oops")

                elif "switch" in query:
                    try:
                        pyautogui.hotkey('alt', 'tab')
                    except Exception as e:
                        print(e)
                        speak("oops")

                elif "scroll down" in query:
                    try:
                        mouse.click('middle')
                        pyautogui.moveRel(0, 70)
                        query=takeCommand().lower()
                        if query=='stop':
                            mouse.click('middle')
                            
                    except Exception as e:
                        print(e)
                        speak("oops")

                elif "scroll up" in query:
                    try:
                        mouse.click('middle')
                        pyautogui.moveRel(0, -70)
                        query=takeCommand().lower()
                        if query=='stop':
                            mouse.click('middle')                           
                    except Exception as e:
                        print(e)
                        speak("oops")

                elif query in strings_hello:
                    mylist = ["hello, sir",
                              "how are you sir?",
                              "yes sir",
                              "im listening, sir",
                              "Nice to see you again!"]
                    n = random.randint(0, 4)
                    speak(mylist[n])

                elif query in strings_shut_up:
                    mylist = ["Got it", "Sure", "ok, sir"]
                    n = random.randint(0, 3)
                    speak(mylist[n])

                elif query in Strings_help:
                    mylist = [
                        "you can say the word open before the Application you want to open.",
                        "if you want me to write someting for you, say 'type' before the word and phase you want me to type"]
                    n = random.randint(0, 1)
                    speak(mylist[n])

                elif query in strings_how_are_you:
                    speak("Absolutely perfect sir")
                    speak("what about you?")

                elif query in strings_im_fine:
                    mylist = [
                        "that's awesome"
                        ]
                    n = random.randint(0, 1)
                    speak(mylist[n])
                    