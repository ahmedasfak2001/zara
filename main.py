import operator
import random
import sys
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import PyPDF2
import pyjokes
import pyttsx3
import requests
import speech_recognition as sr
import datetime
import os
from pywikihow import search_wikihow
from requests import get
import wikipedia
import webbrowser
import pywhatkit
import smtplib
import pyautogui
from bs4 import BeautifulSoup
from twilio.rest import Client


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
print(voices)
engine.setProperty('voice', voices[2].id)  # for lady voice use 1 otherwise use 0
engine.setProperty('rate', 180)  # to control the speech rate

# text to speech
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

# to convert voice into text
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('listening...')
        r.pause_threshold = 1
        audio = r.listen(source, timeout=100, phrase_time_limit=100)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-ln')
        print(f"user said: {query}")
    except Exception as e:
        # speak("say that again please sir...")
        return "none"
    query = query.lower()
    return query

# to read pdf
def pdf_reader():
    speak("sir please enter the location of the file")
    location = input("Enter the location here: ")
    book = open(location, 'rb')
    pdfReader = PyPDF2.PdfFileReader(book)
    pages = pdfReader.numPages
    speak(f"Total number of pages of this book {pages}")
    speak("sir please enter the page number i have to read")
    pg = int(input("Please enter the page number: "))
    page = pdfReader.getPage(pg)
    text = page.extractText()
    speak(text)


# to wish
def wish():
    user_name = 'asfak'
    hour = int(datetime.datetime.now().hour)
    tt = time.strftime("%I:%M %p")

    if hour >= 6 and hour <= 12:
        speak(f"good morning {user_name}, its {tt}")
    elif hour >= 12 and hour <= 15:
        speak(f"good noon {user_name}, its {tt}")
    elif hour > 15 and hour <= 18:
        speak(f"good afternoon {user_name}, its {tt}")
    else:
        speak(f"good evening {user_name}, its {tt}")
    speak("i am zara, your personal assistant. Tell me how may i help you?")


# def sendEmail(to, content):
#    server = smtplib.SMTP('smtp.gmail.com', 587)
#    server.ehlo()
#    server.starttls()
#    server.login('yourmail@gmail.com', 'your_password')
#    server.sendmail('yourmail@gmail.com', to, content)
#    server.close()


# for news update
def news():
    main_url = 'https://newsapi.org/v2/top-headlines?country=us&category=business&apiKey=4dde4df8a17d4b94aac2e43ee7b719c2'
    main_page = requests.get(main_url).json()
    articles = main_page["articles"]
    head = []
    day = ["first", "second", "third", "fourth", "fifth"]
    for ar in articles:
        head.append(ar["title"])
    for i in range(len(day)):
        speak(f"today's {day[i]} news is: {head[i]}")


def TaskExecution():
    wish()
    while True:

        query = takecommand()

        # logic building for tasks

        # to open notepad
        if "open notepad" in query:
            speak('opening notepad...')
            npath = "C:\\WINDOWS\\system32\\notepad.exe"
            os.startfile(npath)

        # to close notepad
        elif "close notepad" in query:
            speak("closing notepad...")
            os.system("taskkill /f /im notepad.exe")

        # to open word
        elif "open word" in query:
            speak("opening ms word...")
            npath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(npath)

        # to close word
        elif "close word" in query:
            speak("closing ms word...")
            os.system("taskkill /f /im WINWORD.EXE")

        # to open excel
        elif "open excel" in query:
            speak("opening ms excel...")
            npath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(npath)

        # to close excel
        elif "close excel" in query:
            speak("closing ms excel...")
            os.system("taskkill /f /im EXCEL.EXE")

        # to open power point
        elif "open powerpoint" in query or "open power point" in query:
            speak("opening powerpoint...")
            npath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(npath)

        # to close power point
        elif "close powerpoint" in query or "close power point" in query:
            speak("closing powerpoint...")
            os.system("taskkill /f /im POWERPNT.EXE")

        # to open command prompt
        elif "open command prompt" in query:
            speak('opening command prompt...')
            os.system("start cmd")

        # to close command prompt
        elif "close command prompt" in query:
            speak("closing...")
            os.system("taskkill /f /im cmd.exe")

        # to play music from a specific directory
        elif "play music" in query:
            music_dir = "path_of_the_music_file"
            songs = os.listdir(music_dir)
            rd = random.choice(songs)
            os.startfile(os.path.join(music_dir, rd))

        # to know your ip address
        elif "ip address" in query:
            ip = get('https://api.ipify.org').text
            speak(f"your ip address is {ip}")

        # to search in wikipedia
        elif "wikipedia" in query:
            speak("searching wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("according to wikipedia")
            speak(results)
            # print(results)

        # to open youtube
        elif "open youtube" in query:
            speak('opening youtube...')
            webbrowser.open("www.youtube.com")

        # to close youtube
        elif "close youtube" in query:
            speak("closing...")
            os.system("taskkill /f /im chrome.exe")

        # to open facebook
        elif "open facebook" in query:
            speak('opening facebook...')
            webbrowser.open("www.facebook.com")

        # to close facebook
        elif "close facebook" in query:
            speak("closing...")
            os.system("taskkill /f /im chrome.exe")

        # to open instagram
        elif "open instagram" in query:
            speak('opening instagram...')
            webbrowser.open("www.instagram.com")

        # to close instagram
        elif "close instagram" in query:
            speak("closing...")
            os.system("taskkill /f /im chrome.exe")

        # to open twitter
        elif "open twitter" in query:
            speak('opening twitter...')
            webbrowser.open("www.twitter.com")

        # to close twitter
        elif "close twitter" in query:
            speak("closing...")
            os.system("taskkill /f /im chrome.exe")

        # to open google and search
        elif "open google" in query:
            speak("sir, what should i search on google")
            cm = takecommand().lower()
            webbrowser.open(f"{cm}")

        # to send message in whatsapp
        elif "send message" in query:
            speak('please enter the number: ')
            num = input("")
            speak('what should i write in the message?')
            msg = takecommand().lower()
            pywhatkit.sendwhatmsg_instantly(num, msg, wait_time=5)
            speak(f"message send to {num} successfully")

        # to play song on youtube
        elif "play a song on youtube" in query:
            speak('which song you want to listen?')
            n = takecommand().lower()
            speak("playing...")
            pywhatkit.playonyt(n)

        # to close
        elif "close song" in query:
            speak("closing...")
            os.system("taskkill /f /im chrome.exe")

        # to set an alarm
        elif "set alarm" in query:
            speak("sir please tell me when i set the alarm, for example, set alarm to 12.30PM")
            tt = takecommand()
            tt = tt.replace("set alarm to ", "")
            tt = tt.replace(".", "")
            tt = tt.upper()
            import MyAlarm
            MyAlarm.alarm(tt)

        # for joke
        elif "tell me a joke" in query:
            joke = pyjokes.get_joke()
            speak(joke)

        # put assistant in sleep
        elif "you can sleep now" in query or "sleep now" in query:
            speak("okay sir, i am going to sleep. You can call me anytime")
            break

        # to shutdown system
        elif "shutdown my system" in query:
            speak("shutting down...")
            os.system("shutdown /s /t 5")
            sys.exit()

        # to restart system
        elif "restart my system" in query:
            speak("restarting system...")
            os.system("shutdown /r /t 5")
            sys.exit()

        # to sleep system
        # elif "sleep my system" in query:
        #    speak("your system is now in sleep...")
        #    os.system("rund1132.exe powrprof.dll,SetSuspendState 0,1,0")

        # for volume up
        elif "volume up" in query:
            pyautogui.press("volumeup")

        # for volume down
        elif "volume down" in query:
            pyautogui.press("volumedown")

        # for mute
        elif "mute" in query or "volume mute" in query:
            pyautogui.press("volumemute")

        # to switch the window
        elif "switch the window" in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")

        # to know the news
        elif "tell me news" in query:
            speak("please wait sir, fetching the latest news... ")
            news()

        # send email
        elif "send email" in query:
            speak("sir, what should i say?")
            query = takecommand().lower()
            try:
                if "send file" in query:
                    email = 'yourmail@gmail.com'  # your email
                    password = 'yourpassword'  # your password
                    speak("Please enter the email id: ")
                    send_mail = input("Enter here: ")
                    send_mail_to = send_mail + "@gmail.com"  # whom you want to send
                    print(send_mail_to)
                    speak("okay sir, what is the subject for this email?")
                    query = takecommand().lower()
                    subject = query  # subject in the email
                    speak("and sir, what is the message for this email?")
                    query2 = takecommand().lower()
                    message = query2  # message in the email
                    speak("sir please enter the correct path of the file into the shell:")
                    file_location = input("Please enter the path here: ")  # the file location for email

                    speak("please wait, i am sending email now...")

                    msg = MIMEMultipart()
                    msg['From'] = email
                    msg['To'] = send_mail_to
                    msg['Subject'] = subject

                    msg.attach(MIMEText(message, 'plain'))

                    # setup the attachment
                    filename = os.path.basename(file_location)
                    attachment = open(file_location, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

                    # attach the attachment to the MIMEMultipart object
                    msg.attach(part)

                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(email, password)
                    text = msg.as_string()
                    server.sendmail(email, send_mail_to, text)
                    server.quit()
                    speak(f"email has been sent to {send_mail_to}")

                else:
                    email = 'yourmail@gmail.com'  # your email id
                    password = 'yourpassword'  # your password
                    speak("please enter the email id: ")
                    send_mail = input("Enter here: ")
                    send_mail_to = send_mail + "@gmail.com"  # whom you want to send
                    print(send_mail_to)
                    speak("okay sir, what is the subject for this email?")
                    query3 = takecommand().lower()
                    subject1 = query3  # the message in the email
                    speak("please wait, i am sending email now...")

                    msg1 = MIMEMultipart()
                    msg1['From'] = email
                    msg1['To'] = send_mail_to
                    msg1['Subject'] = subject1

                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(email, password)
                    server.sendmail(email, send_mail_to, query)
                    server.quit()
                    speak(f"email has been sent to {send_mail_to}")

            except Exception as e:
                print(e)
                speak(f"sorry sir, i am unable to send the email {send_mail_to}")

        # to know our location
        elif "where i am" in query or "where we are" in query:
            speak("wait sir, let me check")
            try:
                ipAdd = requests.get('https://api.ipify.org').text
                print(ipAdd)
                url = 'https://get.geojs.io/v1/ip/geo/' + ipAdd + '.json'
                geo_requests = requests.get(url)
                geo_data = geo_requests.json()
                # print(geo_data)
                city = geo_data['city']
                region = geo_data['region']
                country = geo_data['country']
                speak(
                    f"sir i am not sure, but i think we are in {city} city somewhere in {region} of {country} country")
            except Exception as e:
                speak("sorry sir, Due to network issue i am unable to find our location")
                pass

        # to know weather forecast
        elif "temperature" in query:
            search = "temperature in kolkata"
            url = f"https://www.google.com/search?q={search}"
            r = requests.get(url)
            data = BeautifulSoup(r.text, "html.parser")
            temp = data.find("div", class_="BNeawe").text
            speak(f"current {search} is {temp}")

        # to search anything on internet
        elif "activate how to do mod" in query:
            speak("How to do mode is activated.")
            while True:
                speak("please tell me what do you want to know")
                how = takecommand()
                try:
                    if "exit" in how or "close" in how:
                        speak("okay sir, how to do mode is deactivated.")
                        break
                    else:
                        max_results = 1
                        how_to = search_wikihow(how, max_results)
                        assert len(how_to) == 1
                        how_to[0].print()
                        speak(how_to[0].summary)
                except Exception as e:
                    speak("sorry sir, i am unable to find this.")

        # to check battery
        elif "how much power left" in query or "check battery" in query or "how much power we have" in query:
            import psutil
            battery = psutil.sensors_battery()
            percentage = battery.percent
            speak(f"sir our system have {percentage} percent battery")
            if percentage >= 75:
                speak("we have enough power to continue our work.")
            elif percentage >= 40 and percentage < 75:
                speak("we should connect our system to power supply to give power to our system")
            elif percentage <= 15 and percentage < 40:
                speak("we don't have much power to work, please connect to power supply")
            elif percentage < 15:
                speak("we have very low power, please connect to power supply, system will shutdown soon")

        # to take screenshot
        elif "take screenshot" in query or "take a screenshot" in query:
            speak("sir, please tell me the name for this screenshot file")
            file = takecommand().lower()
            speak("please sir hold screen for few seconds, i am taking screenshot")
            time.sleep(3)
            img = pyautogui.screenshot()
            img.save(f"{file}.png")
            speak("i am done sir, the screenshot is saved in my memory. now i am ready for the next task.")

        # to calculate
        elif "can you calculate" in query or "do some calculations" in query or "do some calculation" in query:
            r = sr.Recognizer()
            with sr.Microphone() as source:
                speak("say what do you want to calculate, example 5 plus 6")
                print("listening...")
                r.adjust_for_ambient_noise(source)
                audio = r.listen(source)
            my_string = r.recognize_google(audio)
            print(my_string)

            def get_operator_fn(op):
                return {
                    '+': operator.add,
                    '-': operator.sub,
                    'x': operator.mul,
                    'divided': operator.__truediv__,
                }[op]

            def eval_binary_expr(op1, oper, op2):
                op1, op2 = int(op1), int(op2)
                return get_operator_fn(oper)(op1, op2)

            speak("your result is")
            speak(eval_binary_expr(*(my_string.split())))

        # to check internet speed
        elif "internet speed" in query:
            import speedtest
            st = speedtest.Speedtest()
            dl = st.download()
            up = st.upload()
            speak(f"sir we have {dl} bit per second downloading speed and {up} bit per second upload speed")

        # to send normal message
        elif "send text" in query:
            speak("sir what should i say")
            msz = takecommand()

            account_sid = 'yoursid'
            auth_token = 'yourtoken'
            client = Client(account_sid, auth_token)

            message = client.messages \
                .create(
                body=msz,
                from_='the number got from twilio',
                to='your number'
            )

            print(message.sid)
            speak("message has been sent")

        # to call
        elif "make a call" in query:
            account_sid = 'sid'
            auth_token = 'token'
            client = Client(account_sid, auth_token)

            message = client.calls \
                .create(
                twiml='<Response><Say>This is the testing call from zara..</Say></Response>',
                from_='twilio number',
                to='signup number'
            )
            print(message.sid)

        # to open mobile camera
        elif "open mobile camera" in query:
            import urllib.request
            import cv2
            import numpy as np
            import time
            URL = "http://192.168.0.110:8080/shot.jpg"
            while True:
                img_arr = np.array(bytearray(urllib.request.urlopen(URL).read()), dtype=np.uint8)
                img = cv2.imdecode(img_arr, -1)
                cv2.imshow('IPWebcam', img)
                q = cv2.waitKey(0)
                if q == ord("q"):
                    break;

            cv2.destroyAllWindows()

        # to open webcam
        elif "open camera" in query or "open webcam" in query:
            speak("opening...")
            import cv2
            cap = cv2.VideoCapture(0)
            while True:
                ret, frame = cap.read()
                cv2.imshow('frame', frame)
                if cv2.waitKey(10) == ord('q'):
                    break
            cap.release()
            cv2.destroyAllWindows()

        # to set reminder
        elif "remember that" in query:
            rememberMsg = query.replace("zara", "")
            rememberMsg = rememberMsg.replace("remember that", "")
            speak(f"you tell me to remember that:" + rememberMsg)
            remember = open('reminder.txt', 'w')
            remember.write(rememberMsg)
            remember.close()

        # to listen what zara remember
        elif "what do you remember" in query:
            remember = open('reminder.txt', 'r')
            speak("you tell me to remember that" + remember.read())

        # read pdf
        elif "read pdf" in query:
            pdf_reader()

        # interaction with zara
        elif "hello" in query or "hi" in query or "what's up" in query:
            speak("hello sir, how may i help you")
        elif "how are you" in query:
            speak("i am fine. what about you sir?")
        elif "am good" in query or "i'm good" in query:
            speak("that's great, sir.")
        elif "thank you" in query or "thanks" in query:
            speak("it's my duty sir.")


if __name__ == "__main__":
    # startExecution()
    while True:
        permission = takecommand()
        if "wake up" in permission or "good morning" in permission:
            TaskExecution()
        elif "goodbye" in permission:
            speak("thanks for using me sir, have a good day")
            sys.exit()
