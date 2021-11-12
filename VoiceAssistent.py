from __future__ import print_function
import warnings
import pyttsx3
import speech_recognition as sr
from gtts import gTTS
import playsound
import os
import datetime
import calendar
import random
import wikipedia
import webbrowser
import ctypes
import winshell
import subprocess
import pyjokes
import smtplib
import requests
import json
import datetime
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from twilio.rest import Client
import wolframalpha
from time import sleep
import time

warnings.filterwarnings("ignore")

engine = pyttsx3.init()
voices = engine.getProperty('voices')       #getting details of current voice
# engine.setProperty('voice', voices[0].id)  #changing index, changes voices. o for male
engine.setProperty('voice', voices[1].id)   #changing index, changes voices. 1 for female

def talk(audio):
    engine.say(audio)
    engine.runAndWait()


def rec_audio():
    recog = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        audio = recog.listen(source)

    data = " "

    try:

        data = recog.recognize_google(audio)
        print("You said: " + data)

    except sr.UnknownValueError:
        print("Assistant could not understand the audio")

    except sr.RequestError as ex:
        print("Request Error from Google Speech Recognition" + ex)

    return data


def response(text):
    print(text)

    tts = gTTS(text=text, lang="en")

    audio = "Audio.mp3"
    tts.save(audio)

    playsound.playsound(audio)

    os.remove(audio)


def call(text):
    action_call = "assistant"

    text = text.lower()

    if action_call in text:
        return True

    return False


def today_date():
    now = datetime.datetime.now()
    date_now = datetime.datetime.today()
    week_now = calendar.day_name[date_now.weekday()]
    month_now = now.month
    day_now = now.day

    months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ]

    ordinals = [
        "1st",
        "2nd",
        "3rd",
        "4th",
        "5th",
        "6th",
        "7th",
        "8th",
        "9th",
        "10th",
        "11th",
        "12th",
        "13th",
        "14th",
        "15th",
        "16th",
        "17th",
        "18th",
        "19th",
        "20th",
        "21st",
        "22nd",
        "23rd",
        "24th",
        "25th",
        "26th",
        "27th",
        "28th",
        "29th",
        "30th",
        "31st",
    ]

    return "Today is " + week_now + ", " + months[month_now - 1] + " the " + ordinals[day_now - 1] + "."

def say_hello(text):
    greet = ["hi", "hey", "hola", "greetings", "wassup", "hello"]

    response = ["howdy", "whats good", "hello", "hey there"]

    for word in text.split():
        if word.lower() in greet:
            return random.choice(response) + "."

    return ""

def wiki_person(text):
    list_wiki = text.split()
    for i in range(0, len(list_wiki)):
        if i + 3 <= len(list_wiki) - 1 and list_wiki[i].lower() == "who" and list_wiki[i + 1].lower() == "is":
            return list_wiki[i + 2] + " " + list_wiki[i + 3]

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])
    # If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


# def google_calendar():
#     """Shows basic usage of the Google Calendar API.
#     Prints the start and name of the next 10 events on the user's calendar.
#     """
#     creds = None
#     # The file token.json stores the user's access and refresh tokens, and is
#     # created automatically when the authorization flow completes for the first
#     # time.
#     if os.path.exists('token.json'):
#         creds = Credentials.from_authorized_user_file('token.json', SCOPES)
#     # If there are no (valid) credentials available, let the user log in.
#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(
#                 'credentials.json', SCOPES)
#             creds = flow.run_local_server(port=0)
#         # Save the credentials for the next run
#         with open('token.json', 'w') as token:
#             token.write(creds.to_json())

#     service = build('calendar', 'v3', credentials=creds)
#     return service


# def calendar_events(num, service):
#     talk('Hey there! Good Day. Hope you are doing fine. These are the events to do today')
#     now = datetime.datetime.utcnow().isoformat() + 'Z'
#     print(f'Getting the upcoming {num} events')
#     events_result = service.events().list(calendarId='primary', timeMin=now,
#                                           maxResults=num, singleEvents=True,
#                                           orderBy='startTime').execute()
#     events = events_result.get('items', [])

#     if not events:
#         talk('No upcoming events found.')
#     for event in events:
#         start = event['start'].get('dateTime', event['start'].get('date'))
#         events_today = (event['summary'])
#         start_time = str(start.split("T")[1].split("-")[0])
#         if int(start_time.split(":")[0]) < 12:
#             start_time = start_time + "am"
#         else:
#             start_time = str(int(start_time.split(":")[0]) - 12)
#             start_time = start_time + "pm"
#         talk(f'{events_today} at {start_time}')


# try:
#     service = google_calendar()
#     calendar_events(10, service)
# except:
#     talk("Could not connect to the local wifi network. Please try again later.")
#     exit()

def send_email(to, content):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.ehlo()
    server.starttls()

    server.login("voicelyassistant@gmail.com", "voicely5678")
    server.sendmail("voicelyassistant@gmail.com", to, content)
    server.close()

while True:

    try:

        text = rec_audio()
        speak = ""

        if call(text):

            speak = speak + say_hello(text)

            if "date" in text or "day" in text or "month" in text:
                get_today = today_date()
                speak = speak + " " + get_today

            elif "time" in text:
                now = datetime.datetime.now()
                meridiem = ""
                if now.hour >= 12:
                    meridiem = "p.m"
                    hour = now.hour - 12
                else:
                    meridiem = "a.m"
                    hour = now.hour

                if now.minute < 10:
                    minute = "0" + str(now.minute)
                else:
                    minute = str(now.minute)
                speak = speak + " " + "It is " + str(hour) + ":" + minute + " " + meridiem + " ."

            elif "wikipedia" in text or "Wikipedia" in text:
                if "who is" in text:
                    person = wiki_person(text)
                    wiki = wikipedia.summary(person, sentences=2)
                    speak = speak + " " + wiki
            elif "who are you" in text or "define yourself" in text:
                speak = speak + """Hello, I am an Assistant. Your Assistant. I am here to make your life easier.  
                You can command me to perform various tasks such as solving mathematical questions or opening 
                applications etcetera."""

            elif "your name" in text:
                speak = speak + "My name is Assistant."

            elif "who am I" in text:
                speak = speak + "You must probably be a human."

            elif "why do you exist" in text or "why did you come" in text:
                speak = speak + "It is a secret."

            elif "how are you" in text:
                speak = speak + "I am fine, Thank you!"
                speak = speak + "\nHow are you?"

            elif "fine" in text or "good" in text:
                speak = speak + "It's good to know that you are fine"

            elif "open" in text.lower():

                if "chrome" in text.lower():
                    speak = speak + "Opening Google Chrome"
                    os.startfile(
                        r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                    )

                elif "word" in text.lower():
                    speak = speak + "Opening Microsoft Word"
                    os.startfile(
                        r"C:\Program Files\Microsoft Office\Office12\WINWORD.EXE"
                    )

                elif "excel" in text.lower():
                    speak = speak + "Opening Microsoft Excel"
                    os.startfile(
                        r"C:\Program Files\Microsoft Office\Office12\EXCEL.EXE"
                    )

                elif "vs code" in text.lower():
                    speak = speak + "Opening Visual Studio Code"
                    os.startfile(
                        r"C:\Program Files\Microsoft VS Code\Code.exe"
                    )

                elif "youtube" in text.lower():
                    speak = speak + "Opening Youtube\n"
                    webbrowser.open("https://youtube.com/")

                elif "google" in text.lower():
                    speak = speak + "Opening Google\n"
                    webbrowser.open("https://google.com/")

                elif "stackoverflow" in text.lower():
                    speak = speak + "Opening StackOverFlow"
                    webbrowser.open("https://stackoverflow.com/")


                else:
                    speak=speak + "Application not available"


            elif "youtube" in text.lower():
                ind = text.lower().split().index("youtube")
                search = text.split()[ind + 1:]
                webbrowser.open(
                    "http://www.youtube.com/results?search_query=" +
                    "+".join(search)
                )
                speak = speak + "Opening " + str(search) + " on youtube"

            elif "search" in text.lower():
                ind = text.lower().split().index("search")
                search = text.split()[ind + 1:]
                webbrowser.open(
                    "https://www.google.com/search?q=" + "+".join(search))
                speak = speak + "Searching " + str(search) + " on google"

            elif "google" in text.lower():
                ind = text.lower().split().index("google")
                search = text.split()[ind + 1:]
                webbrowser.open(
                    "https://www.google.com/search?q=" + "+".join(search))
                speak = speak + "Searching " + str(search) + " on google"

            elif "change background" in text or "change wallpaper" in text:
                img = r"E:\voicely wallpapers"
                list_img = os.listdir(img)
                imgChoice = random.choice(list_img)
                randomImg = os.path.join(img, imgChoice)
                ctypes.windll.user32.SystemParametersInfoW(20, 0, randomImg, 0)
                speak = speak + "Background changed successfully"

            elif "play music" in text or "play song" in text:
                talk("Here you go with music")
                music_dir =r"C:\Users\Smart\Desktop\Music"
                songs = os.listdir(music_dir)
                d = random.choice(songs)
                random = os.path.join(music_dir, d)
                playsound.playsound(random)

            elif "empty recycle bin" in text:
                winshell.recycle_bin().empty(
                    confirm=True, show_progress=False, sound=True
                )
                speak = speak + "Recycle Bin Emptied"

            elif "note" in text or "remember this" in text:
                talk("What would you like me to write down?")
                note_text = rec_audio()
                note(note_text)
                speak = speak + "I have made a note of that."

            elif "joke" in text:
                speak = speak + pyjokes.get_joke()

            elif "where is" in text:
                ind = text.lower().split().index("is")
                location = text.split()[ind + 1:]
                url = "https://www.google.com/maps/place/" + "".join(location)
                speak = speak + "This is where " + str(location) + " is."
                webbrowser.open(url)

            elif "email to computer" in text or "gmail to computer" in text:
                try:
                    talk("What should I say?")
                    content = rec_audio()
                    to = "daduabhijeet@gmail.com"
                    send_email(to, content)
                    speak = speak + "Email has been sent !"
                except Exception as e:
                    print(e)
                    talk("I am not able to send this email")

            elif "mail" in text or "email" in text or "gmail" in text:
                try:
                    talk("What should I say?")
                    content = rec_audio()
                    talk("whom should i send")
                    to = input("Enter To Address: ")
                    send_email(to, content)
                    speak = speak + "Email has been sent !"
                except Exception as e:
                    print(e)
                    speak = speak + "I am not able to send this email"

            elif "weather" in text:
                key = "b7cff4cc8daa504b247cc89760fa0936"
                weather_url = "http://api.openweathermap.org/data/2.5/weather?"
                ind = text.split().index("in")
                location = text.split()[ind + 1:]
                location = "".join(location)
                url = weather_url + "appid=" + key + "&q=" + location
                js = requests.get(url).json()
                if js["cod"] != "404":
                    weather = js["main"]
                    temperature = weather["temp"]
                    temperature = temperature - 273.15
                    humidity = weather["humidity"]
                    desc = js["weather"][0]["description"]
                    weather_response = " The temperature in Celcius is " + str(temperature) + " The humidity is " + str(
                        humidity) + " and The weather description is " + str(desc)
                    speak = speak + weather_response
                else:
                    speak = speak + "City Not Found"

            elif "news" in text:
                url = ('https://newsapi.org/v2/top-headlines?country=us&apiKey=f25a6c868166490f916b8db02df4c709')

                try:
                    response = requests.get(url)
                except:
                    talk("Please check your connection")

                news = json.loads(response.text)

                for new in news["articles"]:
                    print(str(new["title"]), "\n")
                    talk(str(new["title"]))
                    engine.runAndWait()

                    print(str(new["description"]), "\n")
                    talk(str(new["description"]))
                    engine.runAndWait()
                
            elif "send message" in text or "send a message" in text:
                account_sid ="AC40d0ee34f62d070957e853dd8b2843f2"
                auth_token ="57a97dd8991ec61f0a69b9e8cba0c116"
                client = Client(account_sid,auth_token)
                
                talk("What should i send")
                message= client.messages.create(
                    body=rec_audio(), from_ ="57575701",to="9755333739")

                print(message.sid)
                speak =speak+"Message sent successfully"
                    #time.sleep(2)

            elif "calculate" in text:
                app_id = "3X2KK4-76VW76L663"
                client = wolframalpha.Client(app_id)
                ind = text.lower().split().index("calculate")
                text = text.split()[ind + 1:]
                res = client.query(" ".join(text))
                answer = next(res.results).text
                speak = speak + "The answer is " + answer

            elif "what is" in text or "who is" in text:
                app_id = "3X2KK4-76VW76L663"
                client = wolframalpha.Client(app_id)
                ind = text.lower().split().index("is")
                text = text.split()[ind + 1:]
                res = client.query(" ".join(text))
                answer = next(res.results).text
                speak = speak + answer

            elif "don't listen" in text or "stop listening" in text or "do not listen" in text:
                talk("for how many seconds do you want me to sleep")
                a = int(rec_audio())
                time.sleep(a)
                speak = speak + str(a) + " seconds completed. Now you can ask me anything"
                
            elif "exit" in text or "quit" in text:
                exit()
            
            response(speak)

            
    except:
        talk("I don't know that")

       

