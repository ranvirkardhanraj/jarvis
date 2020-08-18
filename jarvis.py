import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import requests
import random
import pyautogui
import wmi
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am Jarvis Sir. Please tell me how may I help you")

def takeCommand():


    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('ranvirkardhanraj007@gmail.com', '9890802923')
    server.sendmail('ranvirkardhanraj007@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    wishMe()
    while True:

        query = takeCommand().lower()

        if 'wikipedia' in query:
            print("you say:"+query)
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:

            webbrowser.open("youtube.com")

        elif 'open google' in query:

            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:

            webbrowser.open("stackoverflow.com")


        elif 'play music' in query:

            music_dir = 'G:\\mom fevorite song'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query:

            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open project' in query:

            codePath = "C:\\Users\\Unique pc\\PycharmProjects\\ERP-project\\billing management.py"
            os.startfile(codePath)

        elif 'open image' in query:
            codePath = "C:\\Users\\Unique pc\\Pictures\\Screenshots"
            os.startfile(codePath)



        elif 'email to dada' in query:

            try:
                speak("What should I say?")
                content = takeCommand()
                to = "uday.iicmr@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry sir. I am not able to send this email")

        elif 'email to sahi dada' in query:

            try:
                speak("What should I say?")
                content = takeCommand()
                to = "ranvirkarsainath@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry sir. I am not able to send this email")

        elif 'email to baba' in query:

            try:
                speak("What should I say?")
                content = takeCommand()
                to = "ranvirkar.baburao@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry sir. I am not able to send this email")

        elif "location" in query:

            print("you say "+query)
            speak("wait a minite sir")
            res = requests.get("https://ipinfo.io/")
            data = res.json()
            position = data['loc'].split(',')
            latitude = position[0]
            longitude = position[1]
            city = data['city']
            region = data['region']
            speak("your in"+ city+"city in "+region+" region \n latitude"+latitude+"and longitude is"+longitude)

        elif "thank you" in query:

            anslist = ["You're welcome.","No problem.","My pleasure.","Anytime."]
            speak(random.choice(anslist))

        elif "take screenshot" in query:
            myscreenshot = pyautogui.screenshot()
            speak("what should be name of screenshot")
            ssname = takeCommand()
            myscreenshot.save('C:\\Users\\Unique pc\\PycharmProjects\\chatbot\\screenshot\\'+ssname+'.png')

        elif "set brightness low" in query:
            speak("brightness set to zero percent")
            wmi.WMI(namespace='wmi').WmiMonitorBrightnessMethods()[0].WmiSetBrightness(0, 0)

        elif "set brightness full" in query:
            speak("brightness set to hundred percent")
            wmi.WMI(namespace='wmi').WmiMonitorBrightnessMethods()[0].WmiSetBrightness(100, 0)

        elif 'set brightness' in query:

            try:
                speak("what will be the percentage of brightness")
                ans = takeCommand()
                ll = ans.split(' ')
                anss = ll[0]
                fans = int(anss)
                speak("brightness set to " +ans+ " percent")
                wmi.WMI(namespace='wmi').WmiMonitorBrightnessMethods()[0].WmiSetBrightness(fans, 0)

            except:

                speak("sorry sir i am not get proper answer.\nplease try again")



        elif query.__contains__("bye") or query.__contains__("see you") or query.__contains__("exit") or query.__contains__("quit"):
            speak("good bye sir, have a nice day!")
            exit(1)
