<h1 align='center'> Jarvis </h1>

## Basic Infomation

Jarvis is like google assistant it basically helps you out to automate almost every thing of a computer.

### **Basic things Jarvis can do**

-   Listen and Talk to you
-   Open Apps
-   Close Apps
-   Search the Web
-   Converts Images to Sketch
-   Converts Images to Paintings
-   Open Websites
-   Take Screenshots
-   Take Screenshots of specific parts of your screen
-   Handle the System
-   Automatically type what you say
-   Write Pretty long essays
-   Much More...

---

<div style="margin-top: 25px;"></div>

## Aim of this Project

The Aim of my project is to automate almost anything of a computer.

Because the more you click the less productive you are and alot of your time will be spent just clicking the buttons and options and
if you use your keyboard properly, you will be much productive but its always not the case that you remember every keyboard shortcut
and keybindings so there comes the concept of voice commands which are pretty straight forward and easy to understand and remember
and mostly to use voice commands to perform tasks we install third party softwares and in some cases there will be maleware in those
softwares so why take risk when you can build your own assistant to help you out.

---

<div style="margin-top: 25px;"></div>

## Technologies it uses:

-   `Python` ( default packages )
-   autopep8
-   comtypes
-   MouseInfo
-   numpy
-   opencv-python
-   Pillow
-   PyAudio
-   PyAutoGUI
-   pycodestyle
-   pygame
-   PyGetWindow
-   PyMsgBox
-   pyperclip
-   pypiwin32
-   PyRect
-   PyScreeze
-   pyttsx3
-   PyTweening
-   pywin32
-   SpeechRecognition
-   speedtest-cli
-   toml
-   windowsapps

---

<div style="margin-top: 25px;"></div>

## How Jarvis works ?

> **Text to Speech**

It uses python package called [pyttsx3](https://pypi.org/project/pyttsx3/)

Which uses [Microsoft Speech API (SAPI5)](https://www.google.com/search?q=What+is+sapi5&oq=What+is+sapi5&aqs=edge..69i57j0i512j0i390l2.5448j0j4&sourceid=chrome&ie=UTF-8)
to get the different voices and other properties
then it uses the first voice `( David )` in the voices array which we got from the engine.

```python
import pyttsx3

engine = pyttsx3.init('sapi5') # function to get a reference to a pyttsx3.Engine instance
voices = engine.getProperty("voices") # Gets the current value of the engine property voices.
engine.setProperty("voice", voices[0].id) # Uses the first voice ( David ) from the whole array of voices
```

Then we create a `speak` method to make jarvis speak

```python
def speak(text_to_say):
    print(f"Jarvis: {text_to_say}")
    engine.say(text_to_say) # To make Jarvis speak
    engine.runAndWait()
```

Why create a separate function for the speak functionality ?

Ans: To follow to the DRY ( Do not repeat yourself ) rule

> **Speech to Text**

It uses a python package called [SpeechRecognition](https://pypi.org/project/SpeechRecognition/)

It uses [google's speech recognition](https://github.com/googleapis/google-api-python-client) to process the microphone input from the user and converts it to text.

```python
import speech_recognition as sr

r = sr.Recognizer()

with sr.Microphone() as source:
	print("\nJarvis: Litening...")
	r.pause_threshold = 1
	audio = r.listen(source, phrase_time_limit=5)

try:
	print("Jarvis: Recognizing...")
	query = r.recognize_google(audio, language='en-in')
	print(f"\nUser: {query}\n")

except Exception as err:
	print(err)
```

Then we will add the basic error handling

```python
import speech_recognition as sr

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

except sr.RequestError:
	speak('It looks like your internet is very slow, so try again later.')
	exit()

except KeyboardInterrupt:
	speak("Thanks for using me sir, have a good day")
	exit()

except Exception as err:
	print(err)
```

Then we will create a function for this to implement DRY ( Do not repeat yourself )

```python
import speech_recognition as sr

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

    except sr.RequestError:
        speak('It looks like your internet is very slow, so try again later.')
		exit()

    except KeyboardInterrupt:
        speak("Thanks for using me sir, have a good day")
        exit()

    except Exception as err:
        print(err)

	return query

```

This is all the logic for the take command function ( Speech to text )

> **Opening Apps**

It uses a python package called [windowsapps](https://pypi.org/project/windowsapps/)

Its pretty straight forward and simple

```python
import windowsapps

def isAppInstalled(appName): # To find if the app we are looking for is installed or not
    name = windowsapps.find_app(appName)
    return name # If the app is installed, then it returns the names of that app in a tuple form


def openApp(appName):
    result = isAppInstalled(appName)
    if isinstance(result, tuple):
        speak(f"Opening {result[0]}...")
        windowsapps.open_app(result[0])

    else:
        speak(f"Sorry but cannot find {appName} in your system")
```

> **Closing Apps**

It uses a default python package called [os](https://docs.python.org/3/library/os.html#module-os)

Its pretty straight forward and simple it uses the function `os.system()` to run commands in the terminal ( shell, cmd, etc... )

```python
import os

def closeApp(appName, taskmgrName):
    speak(f"Closing {appName}...")
    os.system(f'taskkill /f /im {taskmgrName}.exe') # CMD Command to close programs and Apps
    speak(completed_task_msg)
```

> **Open Websites**

It uses the default python package called [webbrowser](https://docs.python.org/3/library/webbrowser.html)

Its pretty straight forward and simple here in this case we are using `webbrowser.open_new_tab(url)`

```python
import webbrowser

def openWebsite(url, name):
	"""
	Takes 2 Parameters:

	1: url -> The proper url with http / https
	2: name -> The name which Jarvis will tell you that it is openinig it

	Example:

	openWebsite("https://www.google.com", 'Google')

	# Opens www.google.com
	# Says "Opening Google"

	---

	Why use webbrowser.open_new_tab(url) ? Why not webbrowser.open(url) ?

	Ans: Because webbrowser.open() will only open the website if the browser is running else it will not open the website
		 and webbrowser.open_new_tab(url) will open the website no matter if the browser is running or not
		 if the browser is not running then it will start a new browser window else it will use the current window and open it in the new tab.
	"""

    speak(f"Opening {name}")
    webbrowser.open_new_tab(url)
    speak(completed_task_msg)
```

---

<div style="margin-top: 50px;">
Name: AR Ayush Kumar Sahoo

School: Odisha Adarsha Vidyalaya [OAV](https://www.oav.edu.in) Chanchra Pada

Class: VII

</div>
