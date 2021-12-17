from pyaudio import paBadStreamPtr
import pyttsx3 
import speech_recognition as sr
import datetime
import webbrowser
import wikipedia
import os
import winsound
import smtplib
from plyer import notification
import requests
import time
import requests    
 
def NewsFromBBC():
     
    # BBC news api
    # following query parameters are used
    # source, sortBy and apiKey
    query_params = {
    }
    main_url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=3460b40b92f84f5d83939b14163b5446"
 
    # fetching data in json format
    res = requests.get(main_url, params=query_params)
    open_bbc_page = res.json()
 
    # getting all articles in a string article
    article = open_bbc_page["articles"]
 
    # empty list which will
    # contain all trending news
    results = []
     
    for ar in article:
        results.append(ar["title"])
         
    for i in range(len(results)):
         
        # printing all trending news
        print(i + 1, results[i])
 
    #to read the news out loud for us
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak("Latest news from the day")
    speak.Speak(results)   
    print(results)             
 
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
print(voices[1].id)
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine = pyttsx3.init()
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour >= 0 and hour<=10:
        speak("Good Morning")
    elif hour > 10 and hour <=17:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak("My name is Jarvis. I can help you by opening Wikipedia pages, Gmail, Youtube, Music, VS Code,telling time, setting reminder and sending emails ")
    speak("How may I help you?")
    print('''My name is Jarvis. I can help you by opening: 
    Wikipedia pages
    Gmail
    Telling latest headlines
    Youtube
    Music
    VS Code
    Telling time and 
    Sending emails ''')

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio)
        print(f"User said:{query}") 
        print("")
    except Exception as e:
        print(e)
        print("Say that again please...")
        print("")
        return "None" 
    return query
def send_email(user, pwd, recipient, subject, body):
    import smtplib
    # Prepare actual message
    message = """From: %s\nTo: %s\nSubject: %s\n\n%s
    """ % (user, ", ".join(recipient), subject, body)

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls
    server.ehlo()
    server.login("emailid", "password")
    server.sendmail("emailid", to, content)
    server.quit()


if __name__ == "__main__":
    wishMe()
    takeCommand()
    while True:
        query = takeCommand().lower()
        if 'wikipedia' in query or 'wiki' in query:
            speak ("Searching Wikipedia")
            query=query.replace("wikipedia","")
            results = wikipedia.summary(query, sentences = 2)
            speak("According to Wikipedia")
            speak(results)
            print("According to Wikipedia")
            print(results)
        elif 'hi' in query or 'hello' in query:
            speak("Hello")
        elif 'how are you' in query:
            speak("As good as you")
            print("As good as you. Thank you")
        elif 'headline' in query or 'headlines' in query:
            NewsFromBBC()
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
            result = webbrowser.summary(query,sentences=2)
        elif 'open gmail' in query:
            webbrowser.open ("gmail.com")
        elif 'open google' in query:
             webbrowser.open("google.com")
        elif 'play music' in query:
            webbrowser.open("spotify.com")
        #elif 'set timer' in query or 'set reminder' in query or 'timer' in query or 'reminder' in query:
        #   if __name__ == "__main__":
        #       speak("Reminder? ")
        #       user=int(input("Enter"))
        #       speak("Title and Body")
        #       titl = takeCommand()
        #       messag = takeCommand()
        #       while True:
        #           notification.notify(
        #           title = titl,
        #           message = messag,
        #           timeout = 2
        #           )
        #   time.sleep(user*1)

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is{strTime}")
            print ("The time is: ", strTime)
        elif 'bye' in query or 'good bye' in query:
            speak("Bye. See you soon")
            break
        elif 'get lost' in query:
            speak("Bye. Never hope to see you again")
            break
        elif 'open code' in query or 'open vs code' in query or 'open visual studio' in query:
            codePath = "C:\\Users\\shivansh arora\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        elif 'send an email' in query or 'email' in query:
            try:
                speak("What should I say")
                content = takeCommand()
                to=input("Enter the email id: ")
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login('emailid', 'password')
                server.sendmail('emailid', to, content)
                server.close()
                speak(f"Email has been sent succesfully to {to}")
                print("Email has been sent succesfully ", to)
            except Exception as e:
                print(e)
                speak("Sorry. Email could not be sent")
else:
    speak("Sorry")
    print("Sorry")