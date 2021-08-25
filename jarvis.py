import webbrowser
import pyttsx3
import speech_recognition as sr
import pyautogui
from datetime import datetime
import time
import os
import windowsapps
import random
import cv2
import speedtest

engine = pyttsx3.init('sapi5')
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)

exit_message = "Thanks for using me Sir, have a good day."
request_error = "It looks like your internet is very slow, so try again later."
completed_task_msg = "Ready for the next task."


def speak(text_to_say):
    print(f"Jarvis: {text_to_say}")
    engine.say(text_to_say)
    engine.runAndWait()


def take_command():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        print("\nJarvis: Litening...")
        r.pause_threshold = 1
        audio = r.listen(source, phrase_time_limit=5)

    query = ""

    try:
        print("Jarvis: Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"\nUser: {query}\n")

    except sr.UnknownValueError:
        pass

    except sr.RequestError:
        speak(request_error)
        exit()

    except KeyboardInterrupt:
        speak(exit_message)
        exit()

    except Exception as err:
        print(err)

    return query.lower()


def wish():
    hour = int(datetime.now().hour)
    current_time = time.strftime("%I:%M %p")

    if 0 <= hour <= 12:
        speak(f"Good Morning, Its {current_time}")

    elif 12 < hour < 18:
        speak(f"Good Afternoon, Its {current_time}")

    else:
        speak(f"Good Evening, Its {current_time}")

    speak("I am Jarvis, How may I help you Sir ?")


def isAppInstalled(appName):
    name = windowsapps.find_app(appName)
    return name


def openApp(appName):
    result = isAppInstalled(appName)
    if isinstance(result, tuple):
        speak(f"Opening {result[0]}...")
        windowsapps.open_app(result[0])

    else:
        speak(f"Sorry but cannot find {appName} in your system")


def closeApp(appName, taskmgrName):
    speak(f"Closing {appName}...")
    os.system(f'taskkill /f /im {taskmgrName}.exe')
    speak(completed_task_msg)


def openWebsite(url, name):
    speak(f"Opening {name}...")
    webbrowser.open_new_tab(url)
    speak(completed_task_msg)


def tasks():
    os.system('cls')

    wish()

    while 1:
        query = take_command()

        if "exit" in query or "quit" in query:
            speak(exit_message)
            os.system('cls')
            exit()

        elif "update" in query:
            speak("Updating to the latest version...")
            speak("It might take a couple of minutes to update.")

            time.sleep(10)
            speak("Removing the older version from the system...")

            time.sleep(10)

            speak("Installing the latest version...")
            time.sleep(10)

            speak("Restarting Jarvis...")
            time.sleep(2)

            os.system('cls')
            os.system('python jarvis.py')
            exit()

        elif "sleep" in query:
            speak(
                "You call me anytime with the command 'Wake Up' Jarvis, have a good day")
            break

        # Opening and Closing Apps

        elif "open notepad" in query:
            openApp('Notepad')
            speak(completed_task_msg)

        elif "close notepad" in query:
            closeApp("Notepad", "notepad")

        elif "open cmd" in query or "open commond prompt" in query or "open terminal" in query:
            openApp('Command Prompt')
            speak(completed_task_msg)

        elif "close cmd" in query or "close command prompt" in query or "close terminal" in query:
            closeApp("Commamd Prompt", "cmd")

        elif "open vs code" in query:
            openApp("Visual Studio Code - Insiders")
            speak(completed_task_msg)

        elif "close vs code" in query:
            speak("Sorry, sir but I cant close vs code")
            speak(completed_task_msg)

        elif "open task manager" in query:
            speak("Opening Task manager...")
            pyautogui.hotkey('ctrl', "shift", "esc")
            speak(completed_task_msg)

        elif "close task manager" in query:
            speak("Sorry sir but I dont have access to close task manager")
            speak(completed_task_msg)

        elif "open whatsapp" in query:
            openApp("Whatsapp")
            speak(completed_task_msg)

        elif "close whatsapp" in query:
            closeApp('Whatsapp', 'Whatsapp')

        elif "open pycharm" in query or "open py charm" in query:
            openApp("Pycharm")
            speak(completed_task_msg)

        elif "close pycharm" in query or "close py charm" in query:
            closeApp('PyCharm', 'pycharm64')

        elif "open google chrome" in query or "open chrome" in query:
            openApp('Chrome')

        elif "open microsoft edge" in query or "open edge" in query:
            openApp('Microsoft Edge')

        elif "open calculator" in query:
            openApp('Calculator')
            speak(completed_task_msg)

        elif "open camera" in query:
            openApp('Camera')
            speak(completed_task_msg)

        elif "open email" in query or "open mail" in query:
            openApp('Mail')
            speak(completed_task_msg)

        elif "open maps" in query or "open map" in query:
            openApp('Maps')
            speak(completed_task_msg)

        elif "open store" in query or "open microsoft store" in query:
            openApp('Microsoft Store')
            speak(completed_task_msg)

        elif "open control panel" in query:
            openApp('Control Panel')
            speak(completed_task_msg)

        elif "open whiteboard" in query or "open microsoft whiteboard" in query:
            openApp('Microsoft Whiteboard')
            speak(completed_task_msg)

        elif "open ms office" in query or "open office" in query:
            openApp('Office')
            speak(completed_task_msg)

        elif "open onedrive" in query or "open one drive" in query:
            openApp('Onedrive')
            speak(completed_task_msg)

        elif "open onenote" in query or "open onenote for windows 10" in query or "open 1 note" in query or "open 1 note for windows 10" in query:
            openApp('Onenote for Windows 10')
            speak(completed_task_msg)

        elif "open paint 3d" in query:
            openApp('Paint 3D')
            speak(completed_task_msg)

        elif "open paint" in query:
            openApp('Paint')
            speak(completed_task_msg)

        elif "open photos" in query or "open photo" in query:
            openApp('Photos')
            speak(completed_task_msg)

        elif "open settings" in query:
            openApp('Settings')
            speak(completed_task_msg)

        elif "open skype" in query:
            openApp('Skype')
            speak(completed_task_msg)

        elif "open sticky notes" in query or "sticky notes" in query:
            openApp('Sticky Notes')
            speak(completed_task_msg)

        elif 'open windows tip' in query or "open tip app" in query:
            openApp('Tips')
            speak(completed_task_msg)

        elif "open video editor" in query:
            openApp('Video Editor')
            speak(completed_task_msg)

        elif "open voice recorder" in query:
            openApp('Voice recorder')
            speak(completed_task_msg)

        elif "open weather" in query:
            openApp('Weather')
            speak(completed_task_msg)

        elif "open clock" in query or "open alarm" in query:
            openApp('Alarms & Clock')
            speak(completed_task_msg)

        elif "open powershell" in query or "open windows powershell" in query:
            openApp('Windows PowerShell')
            speak(completed_task_msg)

        elif "open git" in query:
            openApp('Git Bash')
            speak(completed_task_msg)

        elif "open groove music" in query:
            openApp('Groove Music')
            speak(completed_task_msg)

        elif "open windows terminal" in query:
            openApp('Windows Terminal')
            speak(completed_task_msg)

        # Opening Websites

        elif "open google maps" in query:
            openWebsite('https://www.maps.google.com', 'Google Maps')

        elif "open google play" in query or "open play store" in query:
            openWebsite('https://play.google.com', 'Google Play')

        elif "open google docs" in query:
            openWebsite('https://docs.google.com', 'Google Docs')

        elif "open google forms" in query:
            openWebsite('https://forms.google.com', 'Google Forms')

        elif "open google drive" in query:
            openWebsite('https://drive.google.com', 'Google Drive')

        elif "open google meet" in query or "open meet" in query:
            openWebsite('https://meet.google.com', 'Google Meet')

        elif "open google" in query:
            openWebsite('https://google.com', 'Google')

        elif "open youtube" in query:
            openWebsite('https://youtube.com', 'Youtube')

        elif "open gmail" in query:
            openWebsite('https://mail.google.com', 'Gmail')

        # Other Commands

        elif "shutdown" in query:
            speak("Do you want to shutdown the computer ?")

            while 1:
                reply = take_command()
                if "yes" in reply:
                    os.system("shutdown /s /t 5")
                    break
                elif "no" in reply:
                    speak("Ok Sir")
                    speak("Cancelled the shutdown.")
                    break
                else:
                    speak("Say that again please")
                    continue

        elif "restart" in query:
            speak("Do you want to restart the computer ?")

            while 1:
                reply = take_command()
                if "yes" in reply:
                    os.system("shutdown /r /t 5")
                    break
                elif "no" in reply:
                    speak("Ok Sir")
                    speak("Cancelled the restart.")
                    break
                else:
                    speak("Say that again please")
                    continue

        elif "volume up" in query:
            for i in range(3):
                pyautogui.press('volumeup')

        elif "volume down" in query:
            for i in range(3):
                pyautogui.press('volumedown')

        elif "mute" in query:
            pyautogui.press('volumemute')

        elif "internet speed" in query:

            speak("Performing an internet speed test.")
            speak("It might take some time so please wait.")

            wifi = speedtest.Speedtest()
            speak(
                f"Our Downloading Speed is {round(wifi.download() / 8000, 2)} kbps")
            speak(
                f"Our Uploading Speed is {round(wifi.upload() / 8000, 2)} kbps")

            speak(completed_task_msg)

        elif "essay" in query:
            speak("Do you want me to write an essay?")

            while 1:
                answer = take_command()
                if "yes" in answer:
                    speak("Ok sir, writing an essay")
                    essaysDirectory = os.path.join(
                        os.getcwd(), "./data/essays")
                    listOfEssays = os.listdir(essaysDirectory)
                    essayChoosed = random.choice(listOfEssays)

                    openApp('Notepad')
                    time.sleep(1)
                    speak("Starting to write an essay")

                    with open(os.path.join(essaysDirectory, f'./{essayChoosed}'), "r") as essay:
                        pyautogui.write(essay.read(), interval=0.02)

                    break

                elif "no" in answer:
                    speak("Ok sir.")
                    speak("Cancelled the essay")
                    break

                else:
                    speak("Sorry but say that again please")
                    continue

            speak(completed_task_msg)

        elif "type" in query:
            speak("Do you want me to type something ?")

            while 1:
                reply = take_command()

                if "yes" in reply:
                    speak("\"Auto Type\" has been started")
                    speak("What should i type ?")

                    whatToType = take_command()

                    speak('You have 15 seconds to open any kind of text editor.')
                    speak("Then I will start to type")
                    time.sleep(15)

                    pyautogui.write(whatToType, interval=0.025)
                    speak("Successfully wrote it.")
                    break

                elif "no" in reply:
                    speak("Ok sir")
                    speak("Auto Type has been cancelled")
                    break

                else:
                    speak("Say that again please")
                    continue

        elif "screenshot" in query or "screen shot" in query:
            speak("Sir do you want me to take a screenshot")
            while 1:
                answer = take_command()
                if "yes" in answer:
                    speak("Sir please type the file name.")
                    time.sleep(0.5)
                    file_name = input("File name: ")

                    speak("Sir screenshot will be taken in 3 seconds")

                    time.sleep(3)

                    pyautogui.screenshot(f"data/screenshots/{file_name}.png")

                    speak(
                        "Sir screenshot has been taken and saved in the folder Jarvis/data/screenshots folder")
                    break

                elif "no" in answer:
                    speak("Ok sir, screenshot has been cancelled")
                    break

                else:
                    speak('Say that again please')
                    continue

            speak(completed_task_msg)

        elif "switch window" in query:
            pyautogui.hotkey('alt', 'tab')
            speak(completed_task_msg)

        elif "snap" in query or "snapshot" in query or "snip" in query:
            pyautogui.hotkey("win", "shift", "s")
            speak(completed_task_msg)

        elif "music" in query:
            speak("Do you want to listen music ?")
            while 1:
                answer = take_command()
                if "yes" in answer:
                    speak("Now playing music")
                    musicDirectory = os.path.join(os.getcwd(), "./data/music")
                    listOfMusics = os.listdir(musicDirectory)
                    musicChoosed = random.choice(listOfMusics)

                    if musicChoosed.endswith('.mp3'):
                        os.startfile(os.path.join(
                            musicDirectory, musicChoosed))
                    break

                elif "no" in answer:
                    speak("Ok sir")
                    break

                else:
                    speak("Say that again please")

        elif "search" in query:
            speak("Do you want to search something ?")
            while 1:
                reply = take_command()

                if "yes" in reply:
                    speak("Starting \"Voice Search\"")
                    speak("What should i Search ?")
                    searchQuery = take_command()

                    webbrowser.open_new_tab('https://www.google.com')
                    time.sleep(3)

                    pyautogui.write(searchQuery, interval=0.025)
                    pyautogui.press('enter')
                    speak("Here are your results.")
                    break

                elif "no" in reply:
                    speak("Ok Sir")
                    speak("\"Voice Search\" has been cancelled")
                    break

                else:
                    speak("Sorry but say that again please")
                    continue

        elif "start menu" in query:
            pyautogui.press('win')

        elif "sketch" in query:
            speak(
                "Do you want to make a sketch of a random image in the pictures folder ?")
            while 1:
                reply = take_command()

                if "yes" in reply:
                    speak("\"Image to Sketch\" has been started")

                    sketchDirectory = os.path.join(
                        os.getcwd(), "./data/pictures")
                    listOfSketch = os.listdir(sketchDirectory)
                    randomSketch = random.choice(listOfSketch)
                    sketchChoosed = os.path.join(sketchDirectory, randomSketch)

                    speak(f"Creating a sketch of the image {randomSketch}")
                    time.sleep(3)

                    speak("Sketch has been created")

                    image = cv2.imread(sketchChoosed)
                    cv2.waitKey(0)

                    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                    cv2.waitKey(0)

                    inverted_image = 255 - gray_image
                    cv2.waitKey(0)

                    blurred = cv2.GaussianBlur(inverted_image, (21, 21), 0)

                    inverted_blurred = 255 - blurred
                    pencil_sketch = cv2.divide(
                        gray_image, inverted_blurred, scale=256.0)
                    cv2.waitKey(0)

                    cv2.imshow("Original Image", image)
                    cv2.imshow("Sketch", pencil_sketch)
                    cv2.waitKey(0)
                    break

                elif "no" in reply:
                    speak("Cancelled \"Image to Sketch\"")
                    break

                else:
                    speak("Say that again please")
                    continue

        elif "watercolor effect" in query or "water color effect" in query:
            speak(
                "Do you want to add watercolor effect to a random image in the pictures folder ?")
            while 1:
                reply = take_command()

                if "yes" in reply:
                    speak("\"Water Color Effect\" has been started")

                    picturesDir = os.path.join(
                        os.getcwd(), "./data/pictures")
                    listOfPictures = os.listdir(picturesDir)
                    randomPicture = random.choice(listOfPictures)
                    pictureChoosed = os.path.join(
                        picturesDir, randomPicture)

                    speak(
                        f"Adding water color effect to the image {randomPicture}")
                    time.sleep(3)

                    speak("Effect has been added to the image")

                    image = cv2.imread(pictureChoosed)
                    res = cv2.stylization(image, sigma_s=60, sigma_r=0.6)
                    cv2.waitKey(0)

                    cv2.imshow("Original", image)
                    cv2.imshow("Effect Added", res)
                    cv2.waitKey(0)

                    break

                elif "no" in reply:
                    speak("Cancelled \"Water color effect\"")
                    break

                else:
                    speak("Say that again please")
                    continue

        # User Interactions

        elif "thank you" in query or "thanks" in query:
            speak("Its my pleasure to help you out sir.")

        else:
            continue


tasks()
