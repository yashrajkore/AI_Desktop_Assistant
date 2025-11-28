import os
import re
import webbrowser
import subprocess
import threading
import getpass
import datetime
import sys
import ssl
import urllib.parse
import smtplib
from email.message import EmailMessage

# pip packages
try:
    import speech_recognition as sr
    import pyttsx3
    import dateparser
    from apscheduler.schedulers.background import BackgroundScheduler
except Exception as e:
    print("Missing packages. Run:\n"
          "pip install SpeechRecognition pyttsx3 dateparser apscheduler pyaudio\n"
          "On Windows, you may need to install PyAudio wheel separately if pip fails.\n")
    raise

# ------------ CONFIG ------------
# Desktop path (works on Windows)
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")

# Email config (recommended: set environment variables EMAIL_ADDRESS and EMAIL_PASSWORD,
# or you'll be prompted when sending the first email).
EMAIL_ADDRESS = os.environ.get("USER_Email")
EMAIL_PASSWORD = os.environ.get("PASSWORD")  # app password for Gmail typically

# A very small local map for opening common apps (Windows example).
# You can expand this mapping to include full paths to executables on your machine.
APP_PATHS = {
    "notepad": r"C:\Users\Rushikesh\AppData\Local\Packages\Microsoft.Windows.Search_cw5n1h2txyewy\LocalState\AppIconCache\125\{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}_notepad_exe",
    "calculator": r"C:\xampp\phpMyAdmin\vendor\brick\math\src\Internal\Calculator.php",
}

# ------------ TTS & Recognizer setup ------------
engine = pyttsx3.init()
engine.setProperty("rate", 170)

recognizer = sr.Recognizer()
mic = sr.Microphone()

# Scheduler for reminders
scheduler = BackgroundScheduler()
scheduler.start()

def speak(text: str):
    print("Assistant:", text)
    engine.say(text)
    engine.runAndWait()

def listen(timeout=6, phrase_time_limit=10):
    """Listen once from the microphone and return recognized text (lowercased)."""
    with mic:
        recognizer.adjust_for_ambient_noise(mic, duration=0.6)
        try:
            audio = recognizer.listen(mic, timeout=timeout, phrase_time_limit=phrase_time_limit)
        except sr.WaitTimeoutError:
            return ""
    try:
        text = recognizer.recognize_google(audio)
        print("You said:", text)
        return text.lower()
    except sr.UnknownValueError:
        return ""
    except sr.RequestError as e:
        speak("Network error for speech recognition.")
        return ""

# ------------ COMMAND HANDLERS ------------
def handle_search(query: str):
    if not query:
        speak("What would you like me to search for?")
        return
    url = f"https://www.google.com/search?q={urllib.parse.quote(query)}"
    webbrowser.open(url)
    speak(f"Searching for {query} on the web.")

def handle_open(target: str):
    if not target:
        speak("What should I open?")
        return

    target = target.lower().strip()

    # Common websites mapping
    WEBSITE_SHORTCUTS = {
        "youtube": "https://www.youtube.com",
        "google": "https://www.google.com",
        "gmail": "https://mail.google.com",
        "facebook": "https://facebook.com",
        "instagram": "https://instagram.com",
    }

    # 1. If shortcut website name
    if target in WEBSITE_SHORTCUTS:
        webbrowser.open(WEBSITE_SHORTCUTS[target])
        speak(f"Opening {target}.")
        return

    # 2. If it looks like a URL
    if "." in target:
        if not target.startswith("http"):
            target = "http://" + target
        webbrowser.open(target)
        speak(f"Opening website {target}.")
        return

    # 3. Try to open installed apps
    key = target.lower()
    if key in APP_PATHS:
        try:
            subprocess.Popen([APP_PATHS[key]])
            speak(f"Opening {key}.")
            return
        except Exception as e:
            speak(f"Could not open {key}. Error: {e}")

    # 4. Try OS startfile
    try:
        os.startfile(key)
        speak(f"Opening {key}.")
        return
    except Exception:
        pass

    # 5. Fallback: search the web
    speak(f"Couldn't find {target}. Searching on the web.")
    handle_search(target)

def handle_create_folder(name: str):
    if not name:
        speak("Folder name not provided.")
        return
    path = os.path.join(DESKTOP, name)
    try:
        os.makedirs(path, exist_ok=False)
        speak(f"Folder {name} created on desktop.")
    except FileExistsError:
        speak(f"A folder named {name} already exists on your desktop.")
    except Exception as e:
        speak(f"Failed to create folder: {e}")

def handle_create_file(name: str):
    if not name:
        speak("File name not provided.")
        return
    # Ensure extension; if none provided, create .txt
    if not os.path.splitext(name)[1]:
        name += ".txt"
    path = os.path.join(DESKTOP, name)
    try:
        if os.path.exists(path):
            speak(f"A file named {name} already exists. Opening it.")
        else:
            open(path, "w", encoding="utf-8").close()
            speak(f"File {name} created on desktop.")
        # Optionally open it
        try:
            os.startfile(path)
        except Exception:
            pass
    except Exception as e:
        speak(f"Failed to create file: {e}")

def send_email(to_addr: str, subject: str, body: str = ""):
    global EMAIL_ADDRESS, EMAIL_PASSWORD
    # get credentials if not set
    if not EMAIL_ADDRESS:
        EMAIL_ADDRESS = input("Sender email address (Gmail recommended): ").strip()
    if not EMAIL_PASSWORD:
        EMAIL_PASSWORD = getpass.getpass("Email password / app password: ").strip()

    msg = EmailMessage()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_addr
    msg["Subject"] = subject or "(no subject)"
    msg.set_content(body or "Sent via AI Voice Assistant")

    context = ssl.create_default_context()
    try:
        # Gmail SMTP. If you use another provider, change host/port accordingly.
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)
        speak(f"Email sent to {to_addr} with subject {subject}.")
    except smtplib.SMTPAuthenticationError:
        speak("Authentication failed. Check your email/password or use an app password for Gmail.")
    except Exception as e:
        speak(f"Failed to send email: {e}")

def schedule_reminder(task: str, time_str: str):
    """Parse time_str with dateparser, schedule a job that speaks reminder."""
    if not task or not time_str:
        speak("Couldn't understand the task or time for reminder.")
        return

    dt = dateparser.parse(time_str, settings={'PREFER_DATES_FROM': 'future'})
    if not dt:
        speak("I couldn't parse the time. Try: 'tomorrow 8 am' or 'on 25 december 7 pm'.")
        return

    # If parsed time is naive, assume local timezone; apscheduler will use local time
    job_id = f"reminder_{int(dt.timestamp())}_{hash(task) % 10000}"

    def remind_job():
        speak(f"Reminder: {task}")

    scheduler.add_job(remind_job, 'date', run_date=dt, id=job_id, replace_existing=True)
    speak(f"Reminder set for {dt.strftime('%Y-%m-%d %H:%M')} to {task}.")

# ------------ COMMAND PARSING ------------
def parse_and_execute(command: str):
    command = command.strip().lower()
    if not command:
        return

    # SEARCH
    m = re.match(r"(search|google)\s+(for\s+)?(.+)", command)
    if m:
        query = m.group(3)
        handle_search(query)
        return

    # OPEN
    m = re.match(r"(open)\s+(.+)", command)
    if m:
        target = m.group(2).strip()
        handle_open(target)
        return

    # CREATE FOLDER
    m = re.match(r"(create folder|make folder|create a folder)\s+(.+)", command)
    if m:
        name = m.group(2).strip()
        handle_create_folder(name)
        return

    # CREATE FILE
    m = re.match(r"(create file|make file|create a file|make a file)\s+(.+)", command)
    if m:
        name = m.group(2).strip()
        handle_create_file(name)
        return

    # SEND MAIL
    # pattern: send mail to abc@gmail.com about leave (subject)
    m = re.match(r"send (an )?email|send mail", command)
    if "send mail to" in command or "send email to" in command:
        # robust parsing
        m2 = re.search(r"(?:send (?:an )?email|send mail)(?: to)?\s+([^\s]+)\s+(?:about|regarding)\s+(.+)", command)
        if m2:
            to_addr = m2.group(1).strip()
            subject = m2.group(2).strip()
            # optional: ask body
            speak("Do you want to dictate the body of the email? Say 'yes' to dictate or 'no' to send without body.")
            ans = listen(timeout=4, phrase_time_limit=4)
            body = ""
            if "yes" in ans:
                speak("Start dictating the email body after the beep.")
                body = listen(timeout=6, phrase_time_limit=60)
            # run send in background thread (so assistant stays responsive)
            threading.Thread(target=send_email, args=(to_addr, subject, body), daemon=True).start()
            return
        else:
            speak("Couldn't parse the email command. Say: Send mail to abc@example.com about subject.")
            return

    # REMINDERS
    # Example: remind me to buy milk tomorrow 8 o'clock
    m = re.match(r"remind me to (.+)", command)
    if m:
        rest = m.group(1).strip()
        # try to split task and time by looking for common time prepositions
        # simple heuristic: last occurrence of " at | on | tomorrow | today | next | in "
        time_preps = [" at ", " on ", " tomorrow", " today", " next ", " in ", " by ", " after "]
        split_idx = None
        for prep in reversed(time_preps):
            idx = rest.find(prep)
            if idx != -1:
                split_idx = idx
                break
        if split_idx is not None:
            task = rest[:split_idx].strip()
            time_str = rest[split_idx:].strip()
        else:
            # fallback: no explicit time -> ask when
            task = rest
            speak("When should I remind you? Please say a time like 'tomorrow 8 am' or 'on 25 december 7 pm'.")
            time_str = listen(timeout=8, phrase_time_limit=6)
        schedule_reminder(task, time_str)
        return

    # If nothing matched:
    speak("Sorry, I didn't understand that command. Try: Search, Open, Create folder, Create file, Send mail, or Remind me to ...")

# ------------ MAIN LOOP ------------
def main_loop():
    speak("Voice assistant active. Say a command after the beep.")
    try:
        while True:
            speak("Listening...")
            text = listen(timeout=8, phrase_time_limit=10)
            if not text:
                # optionally continue silently; here we prompt again
                continue
            # Allow quitting
            if any(k in text for k in ["exit", "quit", "stop assistant", "shutdown"]):
                speak("Shutting down. Bye!")
                break
            parse_and_execute(text)
    except KeyboardInterrupt:
        speak("Interrupted by user. Exiting.")
    finally:
        scheduler.shutdown()

from flask import Flask, render_template

app = Flask(__name__)

@app.get("/")
def home():
    return render_template("assistant_ui.html")

@app.post("/listen")
def trigger_listen():
    speak("Listening...")
    text = listen()
    parse_and_execute(text)
    return "OK"

def run_server():
    app.run(host="127.0.0.1", port=5000)


if __name__ == "__main__":
    server_thread = threading.Thread(target=run_server)
    server_thread.daemon = True
    server_thread.start()

    print("Flask UI connected. Open the HTML file in your browser.")
    main_loop()
