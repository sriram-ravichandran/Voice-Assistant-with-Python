from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import time
import pyttsx3
import speech_recognition as sr
import pytz
import subprocess
import webbrowser
from googlesearch import search
import pyjokes
from pyowm import OWM
import wikipedia
from newsapi import NewsApiClient
from win10toast import ToastNotifier
import wolframalpha
import oxforddictionaries
from time import strftime
import calendar
import win32com.client as wincl
import sys
import re
import requests
import json
from bs4 import BeautifulSoup
import random
import tkinter as tk

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ["january", "february", "march", "april", "may", "june","july", "august", "september","october","november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENTIONS = ["rd", "th", "st", "nd"]

def Guide():
    root = tk.Tk()
    w = tk.Label(root, text="""

Wake Word:
    Hello Jack
    
    
To Access Google Calendar Events :
	Example : Do I have plans on saturday
		      What do i have on May 4th

To Make a Note:
	Example : Make a note, Today is a Good day

To Say Today's date:
	Example : Tell me Today's date

To Open a Website:
	Example : Open Google

To Search in the Internet:
	Example : Search Artificial Intelligence

To Say a Joke:
	Example : Tell me a Joke

To Say the Current Weather:
	Example : Tell me about Weather

To Say the Current News:
	Example : Tell me news

To Suggest Some Books:
	Example : Suggest me Books

To Launch the Application:
	Example : Launch Chrome

To Do Arithmetic Calculation:
	Example : Calculate 4 + 3

To Suggest Some movies:
	Example : Suggest some Movies

To Close the Application:
	Example : Close
		  Exit
		  Shut down""")
    w.pack()
    root.mainloop()

def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def GreetUser():
    
    day_time = int(strftime('%H'))
    if day_time < 12:
        speak("Hello user, Good morning")
        print('Jack :' + ' Hello user, Good morning')
    elif 12 <= day_time < 18:
        speak("Hello user, Good afternoon")
        print('Jack :' + ' Hello user, Good afternoon')
    else:
        speak("Hello user, Good evening")
        print('Jack :' + ' Hello user, Good evening')

    speak(".. I am Jack, Your Personal Voice Assistant ..")
    print('Jack :' + ' , Your Personal Voice Assistant')

def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source,phrase_time_limit=7)
        said = ""

        try:
            said = r.recognize_google(audio)
            print(said)
        except Exception as e:
            print("Exception: " + str(e))

    return said.lower()


def authenticate_google():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def get_events(day, service):
    # Call the Calendar API
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("-")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12) + start_time.split(":")[1]
                start_time = start_time + "pm"

            speak(event["summary"] + " at " + start_time)


def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENTIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass

    # THE NEW PART STARTS HERE
    if month < today.month and month != -1:  # if the month mentioned is before the current month set the year to the next
        year = year+1

    # This is slighlty different from the video but the correct version
    if month == -1 and day != -1:  # if we didn't find a month, but we have a day
        if day < today.day:
            month = today.month + 1
        else:
            month = today.month

    # if we only found a dta of the week
    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7

        return today + datetime.timedelta(dif)

    if day != -1:  # FIXED FROM VIDEO
        return datetime.date(month=month, day=day, year=year)

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])


WAKE = "hello jack"
SERVICE = authenticate_google()
i = 0
Guide()

while True:
    if(i<1):
        text = WAKE
        GreetUser()
        i = 2
    else:
        print("Listening")
        text = get_audio()

    if text.count(WAKE) > 0:
        speak("How can I Assist you")
        text = get_audio()

        CALENDAR_STRS = ["what do i have", "do i have plan", "am i busy", "do we have", "do i have anything"]
        for phrase in CALENDAR_STRS:
            if phrase in text:
                date = get_date(text)
                if date:
                    get_events(date, SERVICE)
                else:
                    speak("I don't understand")

        NOTE_STRS = ["make a note", "write this down", "remember this"]
        for phrase in NOTE_STRS:
            if phrase in text:
                speak("What would you like me to write down?")
                note_text = get_audio()
                note(note_text)
                speak("I've made a note of that.")

        DATE_STRS = ["today's date", "Today date"]
        for phrase in DATE_STRS:
            if phrase in text:
                def todays_date(date):
                    day, month, year = (int(i) for i in date.split(' '))
                    dayNumber = calendar.weekday(year, month, day)
                    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                    return (days[dayNumber])

                day_object = datetime.date.today().strftime("%d %m %y")
                print('Today is ' + todays_date(day_object) + ' And the Date is ' + day_object)
                speak(" Today is " + todays_date(day_object) + "And the Date is " + day_object)

        WEBSITE_STRS = ["open"]
        for phrase in WEBSITE_STRS:
            if phrase in text:
                website_request = re.search('open (.+)', text)
                if website_request:
                    domain = website_request.group(1)
                    print('The Requested Website is  https://www.' + domain + '.com')
                    url = 'https://www.' + domain + '.com'
                    webbrowser.open(url)

                print('The website you have requested has been opened for you.')
                speak("The website you have requested has been opened for you.")

        SEARCH_STRS = ["search"]
        for phrase in SEARCH_STRS:
            if phrase in text:
                search_request = re.search('search (.+)', text)
                if search_request:
                    query = search_request.group(1)
                    for url in search(query, stop=1):
                        print(url)
                        webbrowser.open(url)

                print("Your Search Request is Completed")
                speak("Your Search Request is Completed")

        JOKE_STRS = ["tell me a joke"]
        for phrase in JOKE_STRS:
            if phrase in text:
                joke = pyjokes.get_joke()
                print(joke)
                speak(joke)

        WEATHER_STRS = ["weather"]
        for phrase in WEATHER_STRS:
            if phrase in text:
                owm = OWM(api_key='f45914522bed55753f71931e220de2f4') # PyOWM API_key
                mgr = owm.weather_manager()
                obs = mgr.weather_at_place('Hyderabad,India') # location of required weather updates
                w = obs.weather
                x = w.detailed_status
                y = w.humidity # returns humidity
                z = w.temperature(unit='celsius') # returns temperatutre

                print('The Atmosphere Has ' + str(x))
                speak('The Atmosphere Has ' + str(x))
                print('The Maximum Temperature For Today is %d degree celsius and the Minimum Temperature is %d degree celsius' % (z['temp_max'], z['temp_min']))
                speak('The Maximum Temperature is %d degree celsius and the Minimum Temperature is %d degree celsius' % (z['temp_max'], z['temp_min']))
                print('And The Humidity Outside is %d percentage.' % (y))
                speak('And The Humidity Outside is %d percentage.' % (y))

                if z['temp'] >= 30:
                    print('The Weather Outside is Very Hot')
                    speak('The Weather Outside is Very Hot')
                elif z['temp'] > 24 and z['temp'] < 30:
                    print('The Weather Outside is Moderately Hot')
                    speak('The Weather Outside is Moderately Hot')
                elif z['temp'] > 16 and z['temp'] <= 24:
                    print('The Weather Outside is Warm')
                    speak('The Weather Outside is Warm')
                elif z['temp'] < 16:
                    print('The Weather Outside is Cold')
                    speak('The Weather Outside is Cold')

        NEWS_STRS = ["news"]
        for phrase in NEWS_STRS:
            if phrase in text:
                newsapi = NewsApiClient(api_key='d444d79c864946ea9661ac0bd781aa47') # Enter Your API Key from NewsAPI
                top_headlines = newsapi.get_top_headlines(q='India', language='en', )
                for article in top_headlines['articles'][:5]:
                    print('Title : ' + article['title'])
                    speak('Title : ' + article['title'])
                    print('Description : ' + article['description'], '\n')
                    speak('Description : ' + article['description'])

        MEANING_STRS = ["Tell me the meaning of"]
        for phrase in MEANING_STRS:
            if phrase in text:
                reg_ex = re.search('Tell me the meaning of (.+)', text)
                if reg_ex:
                    word_id = reg_ex.group(1)

                    app_id = '4ea13c64' # app_id from oxforddictionaries
                    app_key = 'ee330a18c6d1b84caa707af849421635' #app_key from oxforddictionaries
                    language = 'en'
                    url = 'https://od-api.oxforddictionaries.com:443/api/v2/entries/' + language + '/' + word_id.lower()
                    urlFR = 'https://od-api.oxforddictionaries.com:443/api/v2/stats/frequency/word/' + language + '/?corpus=nmc&lemma=' + word_id.lower()
                    r = requests.get(url, headers={'app_id': app_id, 'app_key': app_key})
                    name_json = r.json()
                    name_list = []
                    for name in name_json['results']:
                        name_list.append(name['word'])
                    print("You searched for the word : " + word_id)
                    speak("You searched for the word : " + word_id)
                    url_mean = 'https://od-api.oxforddictionaries.com:443/api/v1/entries/' + language + '/' + word_id.lower()
                    mean_json = r.json()
                    mean_list = []
                    speak("I can Read it out for you")
                    for result in mean_json['results']:
                        for lexicalEntry in result['lexicalEntries']:
                            for entry in lexicalEntry['entries']:
                                for sense in entry['senses']:
                                    mean_list.append(sense['definitions'][0])
                                for i in mean_list:
                                    print(word_id + " : " + i)
                                    speak(word_id + " : " + i)

        BOOK_STRS = ["book","books"]
        for phrase in BOOK_STRS:
            if phrase in text:
                url = "https://www.goodreads.com/book/most_read" # Website URL address used for web-scraping
                req = requests.get(url)
                bsObj = BeautifulSoup(req.text, "html.parser")
                book_title = bsObj.find_all(class_="bookTitle")
                author_name = bsObj.find_all(class_="authorName")

                print('Here are the 5 Most Read Books This Week In India : \n')

                print(book_title[0].text.strip() + " by " + author_name[0].text.strip())
                speak(book_title[0].text.strip() + " by " + author_name[0].text.strip())
                print(book_title[1].text.strip() + " by " + author_name[1].text.strip())
                speak(book_title[1].text.strip() + " by " + author_name[1].text.strip())
                print(book_title[2].text.strip() + " by " + author_name[2].text.strip())
                speak(book_title[2].text.strip() + " by " + author_name[2].text.strip())
                print(book_title[3].text.strip() + " by " + author_name[3].text.strip())
                speak(book_title[3].text.strip() + " by " + author_name[3].text.strip())
                print(book_title[4].text.strip() + " by " + author_name[4].text.strip())
                speak(book_title[4].text.strip() + " by " + author_name[4].text.strip())

        launch_STRS = ["launch"]
        for phrase in launch_STRS:
            if phrase in text:
                launch_application = re.search('launch (.+)', text)
                application = launch_application.group(1)
                if "chrome" in application: # Launches Google Chrome
                    print("Openning Google Chrome")
                    speak("Openning Google Chrome")
                    os.startfile('C:\Program Files\Google\Chrome\Application\chrome.exe') # File location
                    
                elif "word" in application: # Launches Microsoft Word
                    print("Opening Microsoft Word")
                    speak("Opening Microsoft Word")
                    os.startfile('C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE') # File location
                    
                elif "excel" in application: # Launches Microsoft Word
                    print("Opening Microsoft Excel")
                    speak("Opening Microsoft Excel")
                    os.startfile('C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE') # File location
                    

                elif "power" or "powerpoint" in application: 
                    print("Opening PowerPoint")
                    speak("Opening PowerPoint")
                    os.startfile('C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE') # File location
                    
                else:
                    speak("Application not available")
                    

        CALCULATE_STRS = ["calculate"]
        for phrase in CALCULATE_STRS:
            if phrase in text:
                app_id = "5A773V-U2P37TE6VH" # wolframalpha app_id
                client = wolframalpha.Client(app_id)
                indx = text.lower().split().index('calculate')
                query = text.split()[indx + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                print("The answer is " + answer)
                speak("The answer is " + answer)
                

        MOVIE_STRS = ["movie"]
        for phrase in MOVIE_STRS:
            if phrase in text:
                url = 'http://www.imdb.com/chart/top'
                response = requests.get(url)
                soup = BeautifulSoup(response.text, 'lxml')

                movies = soup.select('td.titleColumn')
                links = [a.attrs.get('href') for a in soup.select('td.titleColumn a')]
                crew = [a.attrs.get('title') for a in soup.select('td.titleColumn a')]
                ratings = [b.attrs.get('data-value') for b in soup.select('td.posterColumn span[name=ir]')]
                votes = [b.attrs.get('data-value') for b in soup.select('td.ratingColumn strong')]

                imdb = []

                # Store each item into dictionary (data), then put those into a list (imdb)
                randomlist = []
                for i in range(0,10):
                    n = random.randint(1,200)
                    randomlist.append(n)
                for index in randomlist:
                    # Seperate movie into: 'place', 'title', 'year'
                    movie_string = movies[index].get_text()
                    movie = (' '.join(movie_string.split()).replace('.', ''))
                    movie_title = movie[len(str(index))+1:-7]
                    year = re.search('\((.*?)\)', movie_string).group(1)
                    place = movie[:len(str(index))-(len(movie))]
                    data = {"movie_title": movie_title,
                            "year": year,
                            "place": place,
                            "star_cast": crew[index],
                            "rating": ratings[index],
                            "vote": votes[index],
                            "link": links[index]}
                    imdb.append(data)

                for item in imdb:
                    d = (item['movie_title'],item['year'])
                    speak(d)
                    print(item['place'], '-', item['movie_title'], '('+item['year']+') -', 'Starring:', item['star_cast'])

        exit_STRS = ["close", "shut","exit"]
        for phrase in exit_STRS:
            if phrase in text:
                print('Bye Bye user, Have a good day')
                speak('Bye Bye user, Have a good day')
                sys.exit()




        

        
