import subprocess
import playsound
import pyttsx3
import speech_recognition as sr
import datetime
import calendar
import wikipedia
import webbrowser
import os
import pywhatkit
import ctypes
import winshell
import pyjokes
import time
import sys
import random

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
rate = engine.getProperty('rate')
engine.setProperty('rate', rate - 25)
engine.setProperty('voice', voices[1].id)
engine.runAndWait()


def speak(audio):
    print(audio)
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good morning miss beautiful!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon sweetheart!")

    else:
        speak("Good Evening princess!")

    speak("I am Siyaana . Please tell me, how may i help you dear.")


def takecommand():
    r = sr.Recognizer()
    r.energy_threshold = 3093.7084974832646

    with sr.Microphone() as source:
        print("listening...")
        r.pause_threshold = 1
        # r.adjust_for_ambient_noise(source, 1.2)
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"user said:{query}\n")

    except Exception as e:
        print(e)
        print("Say that again please...")
        return "None"
    return query


if __name__ == '__main__':
    # speak(" Miss Shraddha Singh is a intelligent, cute, beautiful , smart girl. She is very loving. I love you.")
    wishMe()
    while True:
        query = takecommand().lower()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("Wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results, 'utf-8')
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("www.youtube.com")

        elif 'open google' in query:
            webbrowser.open("www.google.com")

        elif 'open stack overflow' in query:
            webbrowser.open("www.stackoverflow.com")

        elif 'open gmail' in query:
            webbrowser.open("www.gmail.com")

        elif "play song" in query:
            music_dir = "C:\\Users\\shrad\\Music\\Playlists"
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, random.choice(songs)))

        #elif 'play music' in query:
         #   musics_dir = "C:\\Users\\shrad\\OneDrive\\Desktop\\music"
          #  songs = os.listdir(musics_dir)
           # d = random.choice(songs)
            #random = os.path.join(musics_dir, d)
            #playsound.playsound(random)

        elif 'play music' in query:
            musics_dir = "C:\\Users\\shrad\\OneDrive\\Desktop\\music"
            music = os.listdir(musics_dir)
            d = random.choice(music)
            os.startfile(os.path.join(musics_dir, d))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H: %M: %S")
            speak(f"Dear, the time is {strTime}")

        elif 'what\'s up' in query or 'how are you' in query:
            stMsgs = ['Just doing my thing!', 'I am perfectly fine!', 'Nice!', 'I am NIce and full of energy']
            speak(random.choice(stMsgs))
            speak('and you?')

        elif "fine" in query or "good" in query or "perfect" in query:
            speak("it's good to know that you are fine.")

        elif 'nothing' in query or 'abort' in query or 'stop' in query:
            speak('okay Dear')
            speak('Bye Dear,have a good day. ')
            sys.exit()

        elif 'hello' in query:
            speak('Hello Dear ')

        elif 'bye' in query:
            speak('Bye dear, Have a good day.')
            sys.exit()

        elif 'weather' in query:
            webbrowser.open(
                "https://www.msn.com//en-us//weather//today//weather-today//we-city?el=yJTV9sMueHI%2FJDSKwuUII9E6BR0Advwn3ralyft%2B1uc7yYGyfhNZjg%2BxkBvUfJOw%2BrfN7TYiAW5sHNBITk3zt4NIFyK7sXOeO6HdtYP5zZ%2FMkL6j7vDWxaTK5Z08rrdE&weadegreetype=C&ocid=winp1taskbar&pfr=1")

        elif 'play' in query:
            song = query.replace('play', '')
            pywhatkit.playonyt(song)

        elif 'who is' in query:
            person = query.replace('who is', '')
            info = wikipedia.summary(person, 1)
            print(info)
            speak(info)

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif ' day' in query:
            now = datetime.datetime.now()
            date_now = datetime.datetime.today()
            week_now = calendar.day_name[date_now.weekday()]
            month_now = now.month
            day_now = now.day

            month = [
                "January", "February", "March", "April", "May", "june", "July", "August", "September", "October",
                "November",
                "December"]
            ordinals = ["1st", '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th'
                                                                                       '11th', '12th', '13th', '14th',
                        '15th', '16th', '17th', '18th', '19th', '20th',
                        "21st", "22nd", "23rd", "24th", "25th", "26th", "27th", "28th", "29th", "30th", "31st"]
            speak(f'Today is {week_now},{month[month_now - 1]} the {ordinals[day_now - 2]}.')

        elif 'who are you' in query or 'define yourself' in query or 'tell me about yourself' in query:
            me = (
                "Hello, i am your friend. I am here to make your life easier. you can ask me to perform various tasks such as solving mathematical questions or opening applications etc..")
            speak(me)

        elif 'who am i' in query:
            you = (" you are Miss Shraddha Singh.")
            speak(you)

        elif "why do you exist" in query:
            speak('it is secret')

        elif "open" in query.lower():
            if "chrome" in query.lower():
                speak("opening google chrome")
                os.startfile(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")

            elif "word" in query.lower():
                speak("opening Microsoft word")
                os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")

            elif "excel" in query:
                speak("opening Microsoft Excel")
                os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")

            elif "vs code" in query:
                speak("opening visual studio code")
                os.startfile(r"C:\Users\shrad\AppData\Local\Programs\Microsoft VS Code\Code.exe")

            else:
                speak("application not available")

        elif "youtube" in query.lower():
            ind = query.lower().split().index("youtube")
            search = query.split()[ind + 1:]
            webbrowser.open(
                "https://www.youtube.com/results?search_query=" + "+".join(search)
            )
            speak("opening" + str(search) + "on youtube")

        elif "search" in query.lower():
            ind = query.lower().split().index("search")
            search = query.split()[ind + 1:]
            webbrowser.open(
                "https://www.google.com/search?q=" + "+".join(search)
            )
            speak("Searching" + str(search) + "on google")

        elif "google" in query.lower():
            ind = query.lower().split().index("google")
            search = query.split()[ind + 1:]
            webbrowser.open(
                "https://www.google.com/search?q=" + "+".join(search)
            )
            speak("Searching" + str(search) + "on google")

        elif "empty recycle bin" in query:
            winshell.recycle_bin().empty(
                confirm=True, show_progress=False, sound=True
            )
            speak("recycle bin emptied")

        elif "note" in query or "remember this" in query:
            speak("what would you like me to write down?")
            date = datetime.datetime.now()
            file_name = str(date).replace(":", "-") + "-note.query"
            note_query = takecommand()
            with open(file_name, "w") as f:
                f.write(note_query)

            subprocess.Popen(["notepad.exe", file_name])
            speak("I have made a note of that")

        elif "where is" in query:
            ind = query.lower().split().index("is")
            location = query.split()[ind + 1:]
            url = "https://www.google.com/maps/place/" + "".join(location)
            speak("This is where" + str(location) + "is.")
            webbrowser.open(url)

        elif " change background" in query or "change wallpaper" in query:
            img = r'C:\Users\shrad\OneDrive\Desktop\img'
            list_img = os.listdir(img)
            imgChoice = random.choice(list_img)
            randomImg = os.path.join(img, imgChoice)
            ctypes.windll.user32.SystemParametersInfoW(20, 0, randomImg, 0)
            speak("Background changed successfully.")


        elif "don't listen " in query or "stop listening" in query or "do not listen" in query:
            speak("for how many seconds do you want me to sleep")
            a = 3
            time.sleep(a)
            speak(str(a) + "seconds completed. Now you can ask me anything.")

        elif "exit" in query or "quit" in query:
            exit()
