import speech_recognition as sr
import pyttsx3
import datetime
import wikipedia
import webbrowser
import os
import time
import subprocess
from ecapture import ecapture as ec
import wolframalpha
import json
import requests
import smtplib
import pyjokes 
import tkinter
from twilio.rest import Client
from clint.textui import progress
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import win32com.client as wincl
from urllib.request import urlopen
from pptx import Presentation



print('Loading your AI personal assistant - G One')

engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice','voices[0].id')


def speak(text):
    engine.say(text)
    engine.runAndWait()


def usrname(): 
    speak("What should i call you sir") 
    uname = takeCommand() 
    speak("Welcome Mister") 
    speak(uname) 
  
      
    speak("How can i Help you, Sir") 

def tellDay(): 
  
    day = datetime.datetime.today().weekday() + 1
  
    Day_dict = {1: 'Monday', 2: 'Tuesday',  
                3: 'Wednesday', 4: 'Thursday',  
                5: 'Friday', 6: 'Saturday', 
                7: 'Sunday'} 
      
    if day in Day_dict.keys(): 
        day_of_the_week = Day_dict[day] 
        print(day_of_the_week) 
        speak("The day is " + day_of_the_week) 
  

def wishMe():
    hour=datetime.datetime.now().hour
    if hour>=0 and hour<12:
        speak("Hello,Good Morning")
        print("Hello,Good Morning")
    elif hour>=12 and hour<18:
        speak("Hello,Good Afternoon")
        print("Hello,Good Afternoon")
    else:
        speak("Hello,Good Evening")
        print("Hello,Good Evening")

          


def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio=r.listen(source)

        try:
            statement=r.recognize_google(audio,language='en-in')
            print(f"user said:{statement}\n")

        except Exception as e:
            speak("Pardon me, please say that again")
            return "None"
        return statement

def sendEmail(to, content): 
    server = smtplib.SMTP('smtp.gmail.com', 587) 
    server.ehlo() 
    server.starttls() 
      
    # Enable low security in gmail 
    server.login('whitedevil9677@gmail.com', 'devil@8999')
    server.sendmail('whitedevil9677@gmail.com', to, content)
    server.close() 


speak("Loading your AI personal assistant G-One")
wishMe()
usrname()

if __name__=='__main__':
	#clear = lambda: os.system('cls')


     while True:
        #speak("Tell me how can I help you now?")
        statement = takeCommand().lower()
        if statement==0:
            continue

        if "good bye" in statement or "ok bye"in statement:
            speak('your personal assistant G-one is shutting down,Good bye')
            print('your personal assistant G-one is shutting down,Good bye')
            break



        if 'wikipedia' in statement:
            speak('Searching Wikipedia...')
            statement =statement.replace("wikipedia", "")
            results = wikipedia.summary(statement, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in statement:
            webbrowser.open_new_tab("https://www.youtube.com")
            speak("youtube is open now")
            time.sleep(5)

        elif 'open google' in statement:
            webbrowser.open_new_tab("https://www.google.com")
            speak("Google chrome is open now")
            time.sleep(5)


        elif 'open gmail' in statement:
            webbrowser.open_new_tab("gmail.com")
            speak("Google Mail open now")
            time.sleep(5)

        elif 'send a mail' in statement:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")
        
        elif 'launch' in statement:
         reg_ex = re.search('launch (.*)', statement)
         if reg_ex:
            appname = reg_ex.group(1)
            appname1 = appname+".app"
            subprocess.Popen(["open", "-n", "/Applications/" + appname1], stdout=subprocess.PIPE)
            speak('I have launched the desired application')


         elif 'play music' in query or "play song" in statement:
             speak("Here you go with music")
             # music_dir = "G:\\Song"
             music_dir = "C:\\Users\\GAURAV\\Music"
             songs = os.listdir(music_dir)
             print(songs)
             random = os.startfile(os.path.join(music_dir, songs[1]))



        elif 'tell me a joke' in statement: 
            speak(pyjokes.get_joke()) 

        elif "which day it is" in statement: 
            tellDay()  

        elif 'what is love' in statement: 
            speak("It is 7th sense that destroy all other senses") 

        elif 'lock window' in statement: 
                speak("locking the device") 
                ctypes.windll.user32.LockWorkStation()

        

        elif "calculate" or "what is" in statement: 
            question=takeCommand()
            app_id="YTGLWE-UKKQEJ2KL2"
            client = wolframalpha.Client(app_id)
            res = client.statement(question)
            answer = next(res.results).statement
            speak("The answer is " + answer)


  
        elif 'shutdown system' in statement: 
                speak("Hold On a Sec ! Your system is on its way to shut down") 
                subprocess.call('shutdown / p /f')


        elif "where is" in statement:
            statement = statement.replace("where is", "")
            location = statement
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif "change name" in statement: 
            speak("What would you like to call me, Sir ") 
            assname = takeCommand() 
            speak("Thanks for naming me")


        
        elif "weather" in statement:
            api_key="8ef61edcf1c576d65d836254e11ea420"
            base_url="https://api.openweathermap.org/data/2.5/weather?"
            speak("whats the city name")
            city_name=takeCommand()
            complete_url=base_url+"appid="+api_key+"&q="+city_name
            response = requests.get(complete_url)
            x=response.json()
            if x["cod"]!="404":
                y=x["main"]
                current_temperature = y["temp"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                speak(" Temperature in kelvin unit is " +
                      str(current_temperature) +
                      "\n humidity in percentage is " +
                      str(current_humidiy) +
                      "\n description  " +
                      str(weather_description))
                print(" Temperature in kelvin unit = " +
                      str(current_temperature) +
                      "\n humidity (in percentage) = " +
                      str(current_humidiy) +
                      "\n description = " +
                      str(weather_description))

            else:
                speak(" City Not Found ")


        


        elif 'what time it is' in statement:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"the time is {strTime}")


        elif "don't listen" in statement or "stop listening" in statement:
            speak("for how much time you want to stop G-one from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "write a note" in statement:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('G_one.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% Y")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in statement:
            speak("Showing Notes")
            file = open("G_one.txt", "r")
            print(file.read())
            speak(file.read(6))



        elif 'who are you' in statement or 'what can you do' in statement:
            speak('I am G-one version 1 point O your persoanl assistant. I am programmed to minor tasks like'
                  'opening youtube,google chrome,gmail and stackoverflow ,predict time,take a photo,search wikipedia,predict weather' 
                  'in different cities , get top headline news from times of india and you can ask me computational or geographical questions too!')


        elif "who made you" in statement or "who created you" in statement or "who discovered you" in statement:
            speak("I was built by ramanand")
            print("I was built by ramanand")


        elif 'news headlines' in statement:
            news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
            speak('Here are some headlines from the Times of India,Happy reading')
            time.sleep(6)
        elif "camera" in statement or "take a photo" in statement:
            ec.capture(0,"robo camera","img.jpg")

        elif 'search'  in statement:
            statement = statement.replace("search", "")
            webbrowser.open_new_tab(statement)
            time.sleep(5)
        
        elif 'open stack overflow' in statement:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        elif "word" in input:
            speak("Opening Microsoft Word")
            os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office \\Microsoft Word 2010.lnk')
        

        elif 'close microsoft word' in statement:
            speak("closing microsoft word")
            os.close('')
        

        elif "open microsoft excel" in statement: 
            speak("Opening Microsoft Excel") 
            os.startfile('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office \\Microsoft Excel 2010.lnk') 
        
        elif 'close microsoft excel' in statement:
            speak("closing microsoft excel")
            os.close('')


        elif 'open my presentation ' in statement:
            speak("opening presentation")
            codePath1 =r"C:\Users\Rama\Desktop\cpp project\PPT(AI personal voice assistant-using python)"
            os.startfile(codePath1) 

        elif 'close my presentation' in statement:
            speak("closing presentation")
            codePath1 = r"C:\Users\Rama\Desktop\cpp project\PPT(AI personal voice assistant-using python)"
            os.close(codePath1)   
       
        elif 'how are you' in statement:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'bro' in statement or 'dude' in statement:
        	speak("yes boss")
 
        elif 'fine' in statement or "good" in statement:
            speak("It's good to know that your fine")

        elif "what's your name" in statement or "What is your name" in statement:
            speak("My friends call me G-One")
            
            print("My friends call me G-One")
        
        elif "who i am" in statement:
            speak("If you talk then definately your human.")
 
        elif "why you came to world" in statement:
            speak("Thanks to ramanand. further It's a secret")

        

        elif 'reason to create you' in statement:
            speak("I was created as a Minor project by Mister ramanand ")

        elif "who i am" in statement:
            speak("If you talk then definately your human.")

        elif 'ask' in statement:
            speak('I can answer to computational and geographical questions and what question do you want to ask now')
            question=takeCommand()
            app_id="YTGLWE-UKKQEJ2KL2"
            client = wolframalpha.Client('app_id')
            res = client.statement(question)
            answer = next(res.results).text
            speak(answer)
            print(answer)
        
        elif "send message " in statement:
                # You need to create an account on Twilio to use this service
            account_sid = 'ACda6855846979463ce061f061c89f70e4'
            auth_token = '4abe233803ffe4fea17c9a9e7d0d998c'
            client = Client(account_sid, auth_token)
 
            message = client.messages \
                                .create(
                                    body = takeCommand(),
                                    from_='Sender No',
                                    to ='Receiver No'
                                )
 
            print(message.sid)
        
        elif 'change background' in statement:
            ctypes.windll.user32.SystemParametersInfoW(20, 
                                                       0, 
                                                       r"C:\Users\Rama\Desktop\wallpaper\ ",
                                                       0)
            speak("Background changed succesfully")

        elif "will you be my girlfriend" in statement or "will you be my boyfriend" in statement:   
            speak("I'm not sure about, may be you should give me some time")
        
        elif "log off" in statement or "sign out" in statement:
            speak("Ok , your pc will log off in 10 sec make sure you exit from all applications")
            subprocess.call(["shutdown", "/l"])

        elif "open Android studio"in statement:
            speak("opening Android studio")
            codePath4 = r"C:\Program Files\Android\Android Studio1\bin\studio64.exe"
            os.startfile(codePath4)


        elif 'open pycharm' in statement:
           codePath3 = r"C:\Program Files\JetBrains\PyCharm Community Edition 2020.3.3\bin\pycharm64.exe"
           os.startfile(codePath3) 

        elif 'close PyCharm' in statement:
            speak("closing PyCharm")
            os.closefile(codePath3)

        elif 'open sublime' in statement:
            codePath2 = r"C:\Program Files (x86)\Sublime Text 3\sublime_text.exe"
            os.startfile(codePath2)  
        
        elif 'close sublime' in statement:
            speak("closing sublime")
            os.closefile(codePath2)

        elif 'open spotify' in statement:
            codePath = r"C:\Users\Rama\AppData\Roaming\Spotify\Spotify.exe"
            os.startfile(codePath)
            speak("opening spotify")

        elif 'close spotify' in statement:
            speak("closing spotify")
            os.close(codePath)


time.sleep(3)












