# Import necessary libraries
import pyautogui
import pyttsx3
import speech_recognition as sr
import datetime
import winsound
import os
import cv2
import pygetwindow as gw
import random
from requests import get
import wikipedia
import webbrowser
import pywhatkit
import urllib.request
import numpy as np
import smtplib
import sys
import time
from geopy.geocoders import Nominatim
import requests
import subprocess
import pyjokes
from twilio.rest import Client
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import psutil  # Import psutil for battery status

# Initialize pyttsx3 engine
engine = pyttsx3.init('sapi5')

# Iterate through available voices and select a male voice
voices = engine.getProperty('voices')
for voice in voices:
    if "male" in voice.name.lower():
        engine.setProperty('voice', voice.id)
        break
else:
    # If no male voice is found, use the first available voice
    engine.setProperty('voice', voices[1].id)

# API Key for news service
NEWS_API_KEY = 'c8b5228c870b48b4badb8547544453ea'

# Function to make the assistant speak
def speak(audio, speed=200, pitch=200, volume=1.0):
    engine.setProperty('rate', speed)
    engine.setProperty('pitch', pitch)  # Adjust pitch (50 is neutral)
    engine.setProperty('volume', volume)  # Adjust volume (0.0 to 1.0)
    engine.say(audio)
    engine.runAndWait()

# Function to take voice commands from the user
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.pause_threshold = 1
        try:
            audio = r.listen(source, timeout=10, phrase_time_limit=20)
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query
        except sr.WaitTimeoutError:
            print("Listening timed out while waiting for phrase to start. Please try again.")
            speak("Sorry, I did not hear anything. Please say that again.")
            return "none"
        except sr.UnknownValueError:
            print("Sorry, I did not understand your request. Please repeat.")
            speak("Sorry, I did not understand your request. Please repeat.")
            return "none"
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            speak("Sorry, I am having trouble with the speech recognition service. Please try again later.")
            return "none"

# Function to greet the user based on the time of day
def wish():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good morning")
    elif hour >= 12 and hour < 18:
        speak("Good afternoon")
    else:
        speak("Good evening")
    speak("Welcome Back Home, Sir.")

# Function to fetch and read out the news headlines
def get_news():
    main_url = f'https://newsapi.org/v2/top-headlines?country=in&apiKey={NEWS_API_KEY}'
    try:
        news_response = requests.get(main_url)
        news_data = news_response.json()
        articles = news_data.get('articles', [])
        if not articles:
            speak("Sorry, I couldn't fetch the news at the moment.")
            return
        speak("Here are today's top headlines:")
        for i, article in enumerate(articles[:5], 1):
            title = article.get('title', 'N/A')
            speak(f"Headline {i}: {title}")
    except Exception as e:
        print(f"Error fetching news: {e}")
        speak("Sorry, I encountered an issue while fetching the news.")

# Function to set an alarm
def alarm(Timing):
    try:
        formatted_timing = datetime.datetime.strptime(Timing.replace(".", "").lower(), "%I:%M %p").strftime("%I:%M %p")
        altime = datetime.datetime.strptime(formatted_timing, "%I:%M %p")
        speak(f"Done, Alarm is set for {formatted_timing}")
        print(f"Done, Alarm is set for {formatted_timing}")
        while True:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            if current_time == formatted_timing:
                print("Alarm is running")
                winsound.PlaySound('abc', winsound.SND_LOOP)
                break
            elif current_time > formatted_timing:
                winsound.PlaySound(None, winsound.SND_PURGE)
                break
    except ValueError as e:
        print(f"Error: {e}")
        speak("Sorry, I couldn't understand the time. Please provide a valid time format.")

# Function to close the current window
def close_window():
    try:
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)
    except Exception as e:
        print(f"Error closing window: {e}")

# Function to set the volume to maximum
def set_max_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(1.0, None)

# Function to set the volume to minimum
def set_min_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(0.0, None)

# Function to play a video on YouTube
def play_on_youtube(query):
    search_url = f'https://www.youtube.com/results?search_query={query}'
    webbrowser.open(search_url)

# Function to open WhatsApp
def open_whatsapp():
    whatsapp_path = r"C:\Path\To\WhatsApp.exe"
    try:
        subprocess.Popen([whatsapp_path])
        print("WhatsApp opened successfully.")
    except Exception as e:
        print(f"Error: {e}")

# Function to get the current screen brightness
def get_brightness():
    cmd = "(Get-WmiObject -Namespace root/wmi -Class WmiMonitorBrightness).CurrentBrightness"
    result = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
    return int(result.stdout.strip())

# Function to set the screen brightness
def set_brightness(level):
    cmd = f"(Get-WmiObject -Namespace root/wmi -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{level})"
    subprocess.run(["powershell", "-Command", cmd])

# Function to increase the screen brightness
def increase_brightness(increment=10):
    current_brightness = get_brightness()
    new_brightness = min(current_brightness + increment, 100)
    set_brightness(new_brightness)

# Function to decrease the screen brightness
def decrease_brightness(decrement=10):
    current_brightness = get_brightness()
    new_brightness = max(current_brightness - decrement, 0)
    set_brightness(new_brightness)

# Function to get the battery status
def get_battery_status():
    battery = psutil.sensors_battery()
    plugged = battery.power_plugged
    percent = battery.percent

    if plugged:
        speak(f"The battery is currently {percent}% charged and the charger is plugged in.")
    else:
        speak(f"The battery is currently {percent}% charged and the charger is not plugged in.")






if __name__ == "__main__":
    wish()

    while True:
        query = takecommand().lower()

        if "play on youtube" in query:
            speak("Sure, what do you want to play on YouTube?")
            youtube_query = takecommand().lower()
            play_on_youtube(youtube_query)

        elif "set alarm" in query:
            speak("Sure, please tell me the time to set the alarm. For example, set alarm to 5:30 AM")
            tt = takecommand().replace("set alarm to ", "").replace(".", "").upper()
            alarm(tt)


        elif "set brightness to maximum" in query:
            speak("Setting brightness to maximum.")
            set_brightness(100)

        elif "set brightness to minimum" in query:
            speak("Setting brightness to minimum.")
            set_brightness(0)

        elif "increase brightness" in query:
            speak("Increasing brightness.")
            increase_brightness()

        elif "decrease brightness" in query:
            speak("Decreasing brightness.")
            decrease_brightness()

        elif "exit" in query or "quit" in query:
            speak("Exiting the program. Have a good day!")
            break

        elif "hello jarvis" in query:
            speak("Hello sir.")
            continue

        elif "fine jarvis" in query:
            speak("That's good to hear.")
            continue

        elif "what is the time now" in query or "tell me the time" in query:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The current time is {current_time}")

        elif "max volume" in query:
            set_max_volume()
            speak("Setting volume to maximum.")

        elif "min volume" in query:
            set_min_volume()
            speak("Setting volume to minimum.")

        elif "tell me today's news" in query or "what's happening today" in query:
            get_news()

        elif "close the window" in query:
            close_window()

        elif "open notepad" in query:
            speak("Ok sir, opening notepad")
            npath = "C:\\Windows\\system32\\notepad.exe"
            os.startfile(npath)

        elif "open vs code" in query:
            speak("Ok sir, opening Visual Studio Code")
            npath = "C:\\Users\\Akash\\AppData\\Local\\Programs\\Microsoft VS Code\\code.exe"
            os.startfile(npath)

        elif "open powerpoint" in query:
            speak("Ok sir, opening PowerPoint")
            npath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(npath)

        elif "open word" in query:
            speak("Ok sir, opening Word")
            npath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(npath)

        elif "open document" in query:
            speak("Ok sir, opening Documents")
            npath = "C:\\Users\\Akash\\OneDrive\\Documents"
            os.startfile(npath)

        elif "open photos" in query:
            speak("Ok sir, opening Photos")
            npath = "C:\\Users\\Akash\\OneDrive\\Pictures"
            os.startfile(npath)

        elif "open cmd" in query:
            speak("Ok sir, opening Command Prompt")
            os.system("start cmd")

        elif "open camera" in query or "jarvis open camera" in query or "camera" in query:
            speak("Ok sir, opening camera")
            cap = cv2.VideoCapture(0)
            while True:
                ret, img = cap.read()
                cv2.imshow('webcam', img)
                if cv2.waitKey(50) == 27:
                    break
            cap.release()
            cv2.destroyAllWindows()

        elif "play music" in query:
            speak("Ok sir, playing music")
            music_dir = "C:\\Users\\Akash\\Music"
            songs = os.listdir(music_dir)
            for song in songs:
                if song.endswith('.mp3'):
                    os.startfile(os.path.join(music_dir, song))

        elif "wikipedia" in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak(results)
            print(results)

        elif "open youtube" in query or "jarvis open youtube" in query:
            speak("Ok sir, opening YouTube")
            webbrowser.open("www.youtube.com")

        elif "open facebook" in query or "jarvis open facebook" in query or "facebook" in query:
            speak("Ok sir, opening Facebook")
            webbrowser.open("www.facebook.com")

        elif "open instagram" in query or "jarvis open instagram" in query or "instagram" in query:
            speak("Ok sir, opening Instagram")
            webbrowser.open("www.instagram.com")

        elif "open google" in query:
            speak("Sir, what should I search on Google?")
            cm = takecommand().lower()
            webbrowser.open(f"{cm}")

        elif "open whatsapp" in query:
            speak("Ok sir, opening WhatsApp")
            webbrowser.open("https://web.whatsapp.com/")

        elif "volume mute" in query:
            pyautogui.press("volumemute")

        elif "open calculator" in query:
            speak("Opening calculator")
            os.system("start calc.exe")

        elif "open browser" in query:
            speak("Opening web browser")
            webbrowser.open("https://www.google.com")

        elif "close notepad" in query:
            speak("Ok sir, closing notepad")
            os.system("taskkill /f /im notepad.exe")

        elif "close calculator" in query:
            speak("Ok sir, closing calculator")
            os.system("taskkill /f /im calc.exe")

        elif "close vs code" in query:
            speak("Ok sir, closing Visual Studio Code")
            os.system("taskkill /f /im code.exe")

        elif "close document" in query:
            speak("Ok sir, closing Documents")
            close_window()

        elif "close whatsapp" in query:
            speak("Ok sir, closing WhatsApp")
            close_window()

        elif "tell me a joke" in query:
            joke = pyjokes.get_joke()
            speak(joke)

        elif "shutdown the system" in query:
            os.system("shutdown /s /t 5")

        elif "restart the system" in query:
            os.system("shutdown /r /t 5")

        elif "sleep the system" in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

        elif "switch the window" in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")

        elif "close window" in query:
            pyautogui.keyDown("alt")
            pyautogui.press("f4")
            time.sleep(1)

        elif "you can sleep" in query:
            speak("See you later! Take care and have a great day.")
            sys.exit()

        elif "battery level" in query or "check battery status" in query:
            get_battery_status()

















