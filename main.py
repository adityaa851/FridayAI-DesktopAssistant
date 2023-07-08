import datetime
import requests
import win32com.client
import os.path
import pickle
import pyttsx3
import speech_recognition as sr
import webbrowser
from datetime import date
from config import apikey
from config import apikey_weather
from config import apikey_news
from newsapi import NewsApiClient
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import errors
import openai
import pyautogui
import os
import time

speaker = win32com.client.Dispatch("SAPI.SpVoice")
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]  # for gmail


def gmail_read():
    # authentications for gmail
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    today = (date.today())
    today_main = today.strftime('%Y/%m/%d')

    # Call the Gmail API
    results = service.users().messages().list(userId='me',
                                              labelIds=["INBOX", "UNREAD"],
                                              q="after:{0} and category:Primary".format(today_main)).execute()

    # The above code will get emails from primary inbox which are unread

    messages = results.get('messages', [])
    if not messages:
        # if no new emails
        print('No messages found.')
        say('No messages found.')
    else:
        m = ""
        # if email found
        say("{} new emails found".format(len(messages)))
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id'], format='metadata').execute()
            for add in msg['payload']['headers']:
                if add['name'] == "From":
                    # fetching sender's email name
                    a = str(add['value'].split("<")[0])
                    print(a)
                    say("email from" + a)
                    text = takeCommand()
                    if text == "read":
                        print(msg['snippet'])
                        # speak up the mail
                        say(msg['snippet'])
                    else:
                        say("email passed")
    return


def news_headlines():
    try:
        newsapi = NewsApiClient(api_key=apikey_news)
        top_headlines = newsapi.get_top_headlines(language='en', country='in')

        # todo: put up a try catch
        # Extract and print the titles of the articles
        for article in top_headlines['articles']:
            print(f"{article['description']}\n")
            say(article['description'])

    except Exception as e:
        print("There was error some error while retrieving the news sir")


def spotify(song_request):
    say("What song should i play Sir?")
    song_name = takeCommand()
    webbrowser.open(f"https://open.spotify.com/search/{song_name}")
    time.sleep(2)
    say(f"Playing {song_name}")
    for key in ['pagedown', 'tab', 'enter', 'tab', 'enter', 'tab', 'tab', 'tab', 'enter']:
        time.sleep(1)
        pyautogui.press(key)


def weather(city):
    link = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={apikey_weather}"
    r = requests.get(link)
    api_data = r.json()
    weather_desc = api_data['weather'][0]['description']
    temp_celsius = int(api_data['main']['temp'] - 273.15)
    print(f"Sir the temperature today in {city} is {temp_celsius}°C")
    print(f"General Weather is {weather_desc}")
    say(f"Sir the temperature today in {city} is {temp_celsius} degree celsius")
    say(f"General Weather is {weather_desc}")
    say("Would you like the full weather report sir?")
    reply = takeCommand()
    if reply == "yes":
        say("Presenting full weather report")
        feels_like = int(api_data['main']['feels_like'] - 273.15)
        humidity = api_data["main"]["humidity"]
        temp_min = api_data["main"]["temp_min"]
        temp_max = api_data["main"]["temp_max"]
        wind_speed = api_data["wind"]["speed"]
        say(f"feels like {feels_like} degree celsius")
        say(f"Minimum temperature is {temp_min} degree celsius")
        say(f"Maximum temperature is {temp_max} degree celsius")
        say(f"humidity is {humidity}")
        say(f"Wind speed is {wind_speed}")


chatStr = ""


def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Aditya: {query}\nFriday: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    say(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]


def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")
    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)


def say(text):
    engine = pyttsx3.init()
    engine.setProperty('volume', 1.0)
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    voice_rate = engine.getProperty('rate')
    engine.setProperty('rate', voice_rate - 20)
    engine.say(text)
    engine.runAndWait()


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-US")
            print(query)
            return query
        except:
            print("Error occurred unable to recognize voice")


if __name__ == "__main__":
    print("Friday AI")
    say("Hello, this is friday at your service")
    while True:
        print("Listening...")
        command = takeCommand()
        # To open a website through Friday
        sites = [["youtube", "https://youtube.com"], ["spotify", "https://open.spotify.com/"],
                 ["the PS Website", "https://ps1.bits-pilani.in/login/index.php"],
                 ["linked in", "https://www.linkedin.com/feed/"], ["monkey type", "https://monkeytype.com/"]]

        for site in sites:
            if f"Open {site[0]}".lower() in command.lower():
                say(f" Opening {site[0]} Sir...")
                webbrowser.open(site[1])

        if "the time" in command:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir the time is {strTime}")

        # To play music on spotify
        elif "spotify play".lower() in command.lower() or "can u play some song friday" in command.lower() or "play me a song" in command.lower():
            spotify(command)

        elif "using artificial intelligence" in command.lower():
            ai(prompt=command.lower())

        elif "weather in" in command.lower():
            str1 = command.lower()
            weather(''.join(str1.split('in')[1:]).strip())

        elif "news headlines" in command.lower():
            news_headlines()

        elif "read my mails" in command.lower():
            gmail_read()

        elif "Friday Quit".lower() in command.lower() or "GoodBye Friday".lower() in command.lower():
            say("Good Bye Sir")
            exit()
        else:
            print("Chatting...")
            chat(command)
