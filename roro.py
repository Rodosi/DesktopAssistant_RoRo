import os
import sys
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import psutil
import pylint
import random
import smtplib
import subprocess
import time
import json
import operator
import requests
import pyjokes
from PIL import Image , ImageGrab


engine = pyttsx3.init('sapi5')
voices= engine.getProperty('voices') 
engine.setProperty('voice', voices[1].id)


def wish_me():
    hour = int(datetime.datetime.now().hour)
    if hour>5 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<4:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Roro. Let me know how may I help you. Thank you.")

def take_command():   
    #takes input in audio form from user and returns string output
    
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.6
        audio = r.listen(source)
    
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language = 'en-in')
        print(f'User said: {query}\n')
    except Exception as e:
        print("Sorry, can't get you. Say that again please...")
        return "None"
    return query

def get_memory() :
       pid = os.getpid()
       py = psutil.Process(pid)
       return py.memory_info()[0] / 2. ** 30

def take_screenshot():
       image=ImageGrab.grab()
       image.show()

def sendEmail(to, content):
       server = smtplib.SMTP('smtp.gmail.com', 587)
       server.ehlo()
       server.starttls()
       server.login('youremailid@gmail.com', 'yourpassword')  #you need to provide your email id and password for it to run without any errors
       server.sendmail('youremailid@gmail.com', to, content)
       server.close() 

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


if __name__=="__main__":
    wish_me()
    while True:
        query = (take_command()).lower()

        #if 'hello' or 'hey' or 'hi' or 'hola' in query:
            #speak("Hey there!")

        if 'how are you' in query:
            speak("I am good. How are you?")

        elif 'joke' in query:   #returns a joke using pyjokes module
            speak(pyjokes.get_joke())


        elif 'wikipedia' in query:      #searching wikipedia
            speak("Searching Wikipedia")
            query = query.replace("wikipedia","")
            res = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia ")
            print(res)
            speak(res)

        elif 'google' in query:      
            if 'open' in query:      #if query is to open google
                webbrowser.open("https://google.com")
            else:                #if query is to search something in google
                query = query.replace("google", "")
                query = query.replace("search", "")
                query = query.replace(" ", "+")
                webbrowser.open("https://google.com/search?q="+query)

            
        elif 'youtube' in query:     
            if 'open' in query:                         #if query is to open youtube
                webbrowser.open("https://youtube.com")
            else:                                       #if query is to search something in youtube
                query = query.replace("youtube", "")
                query = query.replace("search", "")
                query = query.replace(" ", "+")
                print("search: ", query)
                webbrowser.open("https://www.youtube.com/results?search_query="+query)
        
        elif 'open twitter' in query: #opens twitter
            speak("Opening Twitter for you  ")
            webbrowser.open("twitter.com")

        elif 'open facebook' in query:
            speak("Opening Facebook for you ")
            webbrowser.open("facebook.com")

        elif 'open instagram' in query:           #opens instagram
            speak("Opening Instagram for you  ")
            webbrowser.open("instagram.com")

        elif 'open flipkart' in query: #opens flipkart
            speak("Opening Flipkart for you  ")
            webbrowser.open("flipkart.com")

        elif 'open amazon' in query:  #opens amazon
            speak("Opening Amazon for you  ")
            webbrowser.open("amazon.com")

        elif 'open netflix' in query:      #opens netflix
            speak("Opening Netflix for you  ")
            webbrowser.open("netflix.com")

        elif 'open imdb' in query:   #opens imdb
            speak("Opening Imdb for you  ")
            webbrowser.open("imdb.com")

        elif 'open pinterest' in query:      #opens pinterest
            speak("Opening Pinterest for you  ")
            webbrowser.open("pinterest.com")

        elif 'open linkedin' in query: #opens linkedin
            speak("Opening LinkedIn for you  ")
            webbrowser.open("linkedin.com")

        elif 'open coursera' in query:     #opens coursera 
            speak("Opening Coursera for you  ")
            webbrowser.open("coursera.com")

        elif 'open quora' in query:     #opens quora
            speak("Opening Quora for you  ")
            webbrowser.open("quora.com")

        elif 'open github' in query:      #opens github
            speak("Opening Github for you  ")
            webbrowser.open("github.com")

        elif 'open geeks for geeks' in query:       #opens geeksforgeeeks
            speak("Opening GeeksforGeeks for you ")
            webbrowser.open("geeksforgeeks.com")

        elif 'open stackoverflow' in query:
            speak("Opening Stackoverflow fo you ")
            webbrowser.open("stackoverflow.com")

        elif 'open reddit' in query:        #opens reddit
            speak("Opening Reddit for you  ")
            webbrowser.open("reddit.com")

        elif 'take screenshot' in query:
            time.sleep(1)
            take_screenshot()

        
        elif 'open notepad' in query:  #opens notepad on windows pc
            speak("Opening Notepad for you  ")
            os.system('notepad')

        elif "write a note" in query:    #function to write a note in notepad
            speak("What should i write, sir")
            note = take_command()       #takes input for body of note
            file = open('RORO.txt', 'w') 
            file.write(note)      
                    
        elif "display note" in query:    #display the written note
            file = open("RORO.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif "get memory" in query:   #get memory usage by voice assistant
            print(get_memory())

        elif 'the time' in query:       #time of the day in hrs , mins , secs
            strTime = datetime.datetime.now().strftime("%H:%M:%S") 
            print(f"Sir, the time is {strTime}")
            speak(f"Sir, the time is {strTime}")

        elif 'open microsoft word' in query:      #update file location for your own computer
            speak("Opening Microsoft Word")
            os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Word 2013.Ink")

        elif 'open microsoft powerpoint' in query:        #update file location for your own computer
            speak("Opening Microsoft Powerpoint")
            os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\PowerPoint 2013.lnk")

        elif 'open microsoft excel' in query:  #update file location for your own computer
            speak('Opening Microsoft Excel')
            os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Excel 2013.lnk")

        elif 'open microsoft outlook' in query:  #update file location for your own computer
            speak('Opening Microsoft Outlook')
            os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Outlook 2013.lnk")

        elif 'open paint' in query:     #update file location for your own computer
            speak("Opening paint")
            os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Paint.lnk")

        elif 'date' in query:      #current date returned
            cmd='date'
            os.system(cmd)

        elif 'c drive' in query:        #opens the machine's C drive
            os.system('explorer C:\\"{}"'.format(query.replace('c drive','')))

        elif 'email to recipient' in query:        #email to a predecided gmail user
            try:
                speak("What should I say?")
                content = take_command()         #takes command for body of the email
                to = "youremailid@gmail.com"    #your less secure app access must be turned on
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry, I could not send the email")      #error while trying to send the email


        elif "restart" in query:          #restarts your windows pc
            subprocess.call(["shutdown", "/r"])

        elif 'shutdown system' in query:      #shuts down your windows pc
            speak("Hold On a Second ! Your system is on its way to shut down")
            subprocess.call(["shutdown", "/s", "/t", "60"])          #it will shut down after 60sec

        elif "where is" in query:       #if the query is to find the location of a place 
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif 'play friends' in query:         #music played from a playlist 
            video_dir ='C:\\Users\\Rodosi\\Videos\\F.R.I.E.N.D.S\\SEASON 1'     #must provide the folder location where playlist is located 
            videos = os.listdir(video_dir)
            print(videos)
            os.startfile(os.path.join(video_dir, videos[random.randint(0, len(videos)-1)]))       #random module used for shuffle function in playlist


        elif 'awesome' in query or 'wow' in query or 'amazing' in query or 'wonderful' in query:
            speak("Thank you!")

        
        elif 'quit' in query or 'exit' in query or 'close' in query or 'abort' in query:  
            speak("Happy to help you. Thank you. Have a good day!")
            exit()


