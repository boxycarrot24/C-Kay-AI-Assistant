from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import whisper
import torch
import os
import datetime
import webbrowser
import subprocess
import requests
from dotenv import load_dotenv
import io
import tempfile
import random
import time
import os.path
import win32com.client as win32
import pyautogui
import json
import logging
from datetime import datetime, timedelta
from pydub import AudioSegment
import asyncio
import edge_tts
from transformers import pipeline

app = Flask(__name__)
CORS(app)

# Load environment variables
load_dotenv()

# Load models
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

logger.info("Loading Whisper model...")
whisper_model = whisper.load_model("base")
logger.info("Whisper model loaded!")

logger.info("Loading NLP model...")
chat_model = pipeline("text-generation", model="microsoft/DialoGPT-medium")
logger.info("NLP model loaded!")

# Preload edge-tts voices
logger.info("Loading edge-tts voices...")
VOICES = {}
async def load_voices():
    global VOICES
    voices = await edge_tts.VoicesManager.create()
    for voice in voices.voices:
        lang = voice['Locale']
        if lang not in VOICES:
            VOICES[lang] = voice['ShortName']
        base_lang = lang.split('-')[0]
        if base_lang not in VOICES:
            VOICES[base_lang] = voice['ShortName']
asyncio.run(load_voices())
logger.info(f"Loaded {len(VOICES)} edge-tts voices")

# API Keys
WEATHER_API_KEY = os.getenv("WEATHER_API_KEY")
NEWS_API_KEY = os.getenv("NEWS_API_KEY")

# User settings storage (in-memory for this example)
user_settings = {}

# Utility functions
def get_current_time():
    return datetime.now().strftime("%I:%M %p")

def get_current_date():
    return datetime.now().strftime("%A, %B %d, %Y")

def get_weather(city="Accra"):
    try:
        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={WEATHER_API_KEY}&units=metric"
        response = requests.get(url)
        data = response.json()
        if response.status_code == 200:
            return {
                "temperature": f"{round(data['main']['temp'])}°C",
                "conditions": data['weather'][0]['description'].title(),
                "humidity": f"{data['main']['humidity']}%",
                "wind": f"{data['wind']['speed']} m/s",
                "city": data['name'],
                "status": "success"
            }
        return {"error": "Weather data unavailable", "status": "error"}
    except Exception as e:
        return {"error": str(e), "status": "error"}

def get_news(country="us"):
    try:
        url = f"https://newsapi.org/v2/top-headlines?country={country}&apiKey={NEWS_API_KEY}"
        response = requests.get(url)
        data = response.json()
        if response.status_code == 200 and data['articles']:
            return [{
                "title": article['title'],
                "source": article['source']['name'],
                "url": article['url']
            } for article in data['articles'][:3]], "success"
        return [{"error": "No news available"}], "error"
    except Exception as e:
        return [{"error": str(e)}], "error"

def play_spotify(song=None):
    try:
        if song:
            webbrowser.open(f"spotify:search:{song}")
        else:
            webbrowser.open("spotify:")
        return True
    except Exception as e:
        print(f"Spotify error: {e}")
        return False

def open_application(app_name):
    try:
        app_name = app_name.lower()
        # Microsoft Office Apps
        if "word" in app_name or "document" in app_name:
            os.startfile("winword.exe")
            return True, "Microsoft Word"
        elif "excel" in app_name or "spreadsheet" in app_name:
            os.startfile("excel.exe")
            return True, "Microsoft Excel"
        elif "powerpoint" in app_name or "presentation" in app_name or "ppt" in app_name:
            os.startfile("powerpnt.exe")
            return True, "Microsoft PowerPoint"
        elif "outlook" in app_name or "email" in app_name:
            os.startfile("outlook.exe")
            return True, "Microsoft Outlook"
        
        # System Apps
        elif "calculator" in app_name or "calc" in app_name:
            os.startfile("calc.exe")
            return True, "Calculator"
        elif "paint" in app_name:
            os.startfile("mspaint.exe")
            return True, "Paint"
        elif "notepad" in app_name or "text editor" in app_name:
            os.startfile("notepad.exe")
            return True, "Notepad"
        elif "file explorer" in app_name or "explorer" in app_name or "files" in app_name:
            os.startfile("explorer.exe")
            return True, "File Explorer"
        elif "command prompt" in app_name or "cmd" in app_name:
            os.system("start cmd")
            return True, "Command Prompt"
        elif "task manager" in app_name:
            os.system("taskmgr")
            return True, "Task Manager"
        elif "control panel" in app_name:
            os.system("control")
            return True, "Control Panel"
        elif "settings" in app_name:
            os.system("start ms-settings:")
            return True, "Windows Settings"
        
        # Browsers
        elif "chrome" in app_name or "browser" in app_name or "google" in app_name:
            webbrowser.open("https://www.google.com")
            return True, "Google Chrome"
        elif "edge" in app_name or "microsoft edge" in app_name:
            webbrowser.open("microsoft-edge:")
            return True, "Microsoft Edge"
        elif "firefox" in app_name or "mozilla" in app_name:
            webbrowser.open("firefox")
            return True, "Mozilla Firefox"
        
        # Media Apps
        elif "spotify" in app_name or "music" in app_name:
            webbrowser.open("spotify:")
            return True, "Spotify"
        elif "vlc" in app_name or "media player" in app_name:
            os.startfile("vlc.exe")
            return True, "VLC Media Player"
        elif "photos" in app_name or "pictures" in app_name:
            os.startfile("ms-photos:")
            return True, "Photos App"
        
        # Development Tools
        elif "vs code" in app_name or "code" in app_name or "visual studio code" in app_name:
            os.system("code")
            return True, "Visual Studio Code"
        elif "pycharm" in app_name:
            os.system("pycharm64.exe")
            return True, "PyCharm"
        elif "sublime" in app_name or "sublime text" in app_name:
            os.system("sublime_text")
            return True, "Sublime Text"
        
        # Communication Apps
        elif "zoom" in app_name:
            os.system("zoom")
            return True, "Zoom"
        elif "teams" in app_name or "microsoft teams" in app_name:
            os.system("msteams")
            return True, "Microsoft Teams"
        elif "discord" in app_name:
            os.system("discord")
            return True, "Discord"
        elif "whatsapp" in app_name:
            webbrowser.open("https://web.whatsapp.com")
            return True, "WhatsApp Web"
        
        # Utilities
        elif "calendar" in app_name:
            os.startfile("outlookcal:")
            return True, "Calendar"
        elif "clock" in app_name or "alarm" in app_name:
            os.startfile("ms-clock:")
            return True, "Clock"
        elif "weather" in app_name:
            os.startfile("msnweather:")
            return True, "Weather"
        elif "maps" in app_name or "google maps" in app_name:
            webbrowser.open("https://maps.google.com")
            return True, "Google Maps"
        
        else:
            return False, None
    except Exception as e:
        print(f"Error opening application: {e}")
        return False, None

def control_application(action, app_name):
    try:
        app_name = app_name.lower()
        action = action.lower()
        
        if "close" in action or "exit" in action or "quit" in action:
            if "word" in app_name:
                os.system("taskkill /f /im winword.exe")
            elif "excel" in app_name:
                os.system("taskkill /f /im excel.exe")
            elif "powerpoint" in app_name:
                os.system("taskkill /f /im powerpnt.exe")
            elif "chrome" in app_name:
                os.system("taskkill /f /im chrome.exe")
            elif "edge" in app_name:
                os.system("taskkill /f /im msedge.exe")
            elif "firefox" in app_name:
                os.system("taskkill /f /im firefox.exe")
            elif "spotify" in app_name:
                os.system("taskkill /f /im spotify.exe")
            elif "vlc" in app_name:
                os.system("taskkill /f /im vlc.exe")
            elif "zoom" in app_name:
                os.system("taskkill /f /im zoom.exe")
            elif "teams" in app_name:
                os.system("taskkill /f /im teams.exe")
            elif "discord" in app_name:
                os.system("taskkill /f /im discord.exe")
            elif "code" in app_name or "vs code" in app_name:
                os.system("taskkill /f /im code.exe")
            elif "pycharm" in app_name:
                os.system("taskkill /f /im pycharm64.exe")
            else:
                return False, f"Couldn't close {app_name}"
            return True, f"Closed {app_name}"
        
        elif "minimize" in action or "hide" in action:
            pyautogui.hotkey('win', 'down')
            return True, f"Minimized {app_name}"
        
        elif "maximize" in action or "fullscreen" in action:
            pyautogui.hotkey('win', 'up')
            return True, f"Maximized {app_name}"
        
        elif "restore" in action or "normal" in action:
            pyautogui.hotkey('win', 'down')
            pyautogui.hotkey('win', 'up')
            return True, f"Restored {app_name}"
        
        elif "switch" in action or "alt tab" in action:
            pyautogui.hotkey('alt', 'tab')
            return True, f"Switched windows"
        
        elif "take screenshot" in action or "capture screen" in action:
            screenshot = pyautogui.screenshot()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshot_{timestamp}.png"
            screenshot.save(filename)
            return True, f"Screenshot saved as {filename}"
        
        return False, None
    except Exception as e:
        print(f"Error controlling application: {e}")
        return False, None

def create_email(recipient, subject, body):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def open_document(file_path):
    try:
        if os.path.exists(file_path):
            os.startfile(file_path)
            return True, os.path.basename(file_path)
        return False, None
    except Exception as e:
        print(f"Error opening document: {e}")
        return False, None

def process_command(text):
    text = text.lower()
    response = {
        "command": "unknown",
        "text": f"I'm not sure how to help with: {text}",
        "action": None,
        "data": None
    }

    # Greetings
    if any(word in text for word in ["hello", "hi", "hey", "greetings"]):
        response.update({
            "command": "greeting",
            "text": random.choice([
                "Hello there! I'm C-KAY, your personal voice assistant. How can I help you today?",
                "Hi! I'm C-KAY. What can I do for you?",
                "Hey! I'm C-KAY, ready to assist you. What do you need?"
            ])
        })
    
    # How are you
    elif "how are you" in text or "how's it going" in text:
        response.update({
            "command": "how_are_you",
            "text": random.choice([
                "I'm doing great, thanks for asking! How about you?",
                "I'm just a computer program, but I'm functioning perfectly!",
                "I'm C-KAY, always ready to help! How can I assist you today?"
            ])
        })
    
    # Time and Date
    elif "what time" in text or "current time" in text:
        response.update({
            "command": "get_time",
            "text": f"The current time is {get_current_time()}",
            "data": {"time": get_current_time()}
        })
    elif "what day" in text or "what date" in text or "today's date" in text:
        response.update({
            "command": "get_date",
            "text": f"Today is {get_current_date()}",
            "data": {"date": get_current_date()}
        })
    
    # Weather
    elif "weather" in text or "temperature" in text or "forecast" in text:
        city = "Accra"  # default city
        if "in " in text:
            city = text.split("in ")[1].split(" ")[0]
        weather = get_weather(city)
        if weather.get("status") == "success":
            response.update({
                "command": "get_weather",
                "text": f"Weather in {weather['city']}: {weather['temperature']}, {weather['conditions']}. "
                       f"Humidity: {weather['humidity']}, Wind: {weather['wind']}",
                "data": weather
            })
        else:
            response.update({
                "command": "weather_error",
                "text": f"Couldn't fetch weather data. {weather.get('error', 'Unknown error')}"
            })
    
    # News
    elif "news" in text or "headlines" in text or "latest news" in text:
        country = "us"  # default country
        if "from " in text:
            country = text.split("from ")[1].split(" ")[0]
        news, status = get_news(country)
        if status == "success":
            news_text = "Here are some news headlines:\n"
            for i, item in enumerate(news, 1):
                news_text += f"{i}. {item['title']} ({item['source']})\n"
            response.update({
                "command": "get_news",
                "text": news_text,
                "data": news
            })
        else:
            response.update({
                "command": "news_error",
                "text": "Couldn't fetch news headlines. Please try again later."
            })
    
    # Music
    elif any(phrase in text for phrase in ["play music", "play song", "play artist", "play album"]):
        query = None
        if "play " in text:
            query = text.split("play ")[1]
        if play_spotify(query):
            response.update({
                "command": "play_music",
                "text": f"Playing {query if query else 'music'} on Spotify" if query else "Opening Spotify",
                "data": {"query": query}
            })
        else:
            response.update({
                "command": "music_error",
                "text": "Couldn't open Spotify. Please make sure it's installed."
            })
    
    # Open Applications
    elif any(phrase in text for phrase in ["open ", "launch ", "start "]):
        app_name = text.split("open ")[1] if "open " in text else \
                  text.split("launch ")[1] if "launch " in text else \
                  text.split("start ")[1]
        success, app = open_application(app_name)
        if success:
            response.update({
                "command": "open_app",
                "text": f"Opening {app}",
                "data": {"app": app}
            })
        else:
            response.update({
                "command": "app_error",
                "text": f"Couldn't open {app_name}. Please try another application."
            })
    
    # Application Control
    elif any(word in text for word in ["close", "minimize", "maximize", "restore", "switch", "screenshot"]):
        app_name = text.split(" ")[-1] if not ("switch" in text or "screenshot" in text) else None
        action = text.split(" ")[0] if not ("take screenshot" in text or "capture screen" in text) else "take screenshot"
        success, result = control_application(action, app_name)
        if success:
            response.update({
                "command": "control_app",
                "text": result,
                "data": {"action": action, "app": app_name}
            })
        else:
            response.update({
                "command": "control_error",
                "text": f"Couldn't {action} {app_name if app_name else 'window'}."
            })
    
    # Email
    elif any(phrase in text for phrase in ["send email", "compose email", "write email"]):
        recipient = "example@example.com"  # default
        subject = "No subject"
        body = "No content"
        
        # Simple parsing (in a real app you'd use more sophisticated NLP)
        if "to " in text:
            recipient = text.split("to ")[1].split(" ")[0]
        if "about " in text:
            subject = text.split("about ")[1].split(" and ")[0]
        if "say " in text:
            body = text.split("say ")[1]
        
        if create_email(recipient, subject, body):
            response.update({
                "command": "send_email",
                "text": f"Email sent to {recipient} with subject '{subject}'",
                "data": {"recipient": recipient, "subject": subject}
            })
        else:
            response.update({
                "command": "email_error",
                "text": "Couldn't send email. Please try again."
            })
    
    # System Commands
    elif "shutdown" in text or "turn off" in text:
        response.update({
            "command": "shutdown",
            "text": "I can't actually shut down your computer for security reasons, but I can guide you through the process."
        })
    elif "restart" in text or "reboot" in text:
        response.update({
            "command": "restart",
            "text": "I can't actually restart your computer for security reasons, but I can guide you through the process."
        })
    elif "sleep" in text or "hibernate" in text:
        response.update({
            "command": "sleep",
            "text": "I can't put your computer to sleep for security reasons, but I can guide you through the process."
        })
    
    # Search
    elif "search for " in text or "look up " in text or "google " in text:
        query = text.split("search for ")[1] if "search for " in text else \
               text.split("look up ")[1] if "look up " in text else \
               text.split("google ")[1]
        webbrowser.open(f"https://www.google.com/search?q={query}")
        response.update({
            "command": "search",
            "text": f"Searching for {query} on Google",
            "data": {"query": query}
        })
    
    # Math Calculations
    elif any(word in text for word in ["calculate", "what is", "how much is", "+", "-", "*", "/"]):
        try:
            # Simple calculation handling (for more complex math, consider using eval or a math library)
            if "calculate" in text:
                expr = text.split("calculate ")[1]
            elif "what is" in text:
                expr = text.split("what is ")[1].split(" ")[0]
            elif "how much is" in text:
                expr = text.split("how much is ")[1].split(" ")[0]
            else:
                expr = text
            
            # Basic safety check (in a real app, use a proper expression evaluator)
            if any(c in expr for c in ["import", "exec", "eval", "open", "os.", "sys."]):
                raise ValueError("Invalid expression")
            
            result = eval(expr)  # Note: Using eval is dangerous in production!
            response.update({
                "command": "calculate",
                "text": f"The result is {result}",
                "data": {"expression": expr, "result": result}
            })
        except Exception as e:
            response.update({
                "command": "calculation_error",
                "text": "I couldn't perform that calculation. Please try a simpler one."
            })
    
    # Jokes
    elif any(word in text for word in ["tell me a joke", "make me laugh", "joke"]):
        jokes = [
            "Why don't scientists trust atoms? Because they make up everything!",
            "Did you hear about the mathematician who's afraid of negative numbers? He'll stop at nothing to avoid them!",
            "Why don't skeletons fight each other? They don't have the guts!",
            "I'm reading a book about anti-gravity. It's impossible to put down!",
            "Did you hear about the claustrophobic astronaut? He just needed a little space."
        ]
        response.update({
            "command": "joke",
            "text": random.choice(jokes)
        })
    
    # Help
    elif "help" in text or "what can you do" in text or "commands" in text:
        help_text = """
        I can help you with many things! Here are some examples:
        - Open applications (Word, Excel, Chrome, etc.)
        - Control applications (minimize, maximize, close)
        - Check weather and news
        - Play music on Spotify
        - Send emails
        - Search the web
        - Perform calculations
        - Tell jokes
        - And much more!
        
        Try asking me to do something specific!
        """
        response.update({
            "command": "help",
            "text": help_text
        })
    
    return response

# API Endpoints
@app.route('/transcribe', methods=['POST'])
def transcribe_audio():
    try:
        if 'audio' not in request.files:
            return jsonify({"status": "error", "message": "No audio file provided"}), 400

        audio_file = request.files['audio']
        temp_path = "temp_audio.wav"
        audio_file.save(temp_path)

        result = whisper_model.transcribe(temp_path, fp16=torch.cuda.is_available())
        transcription = result["text"].strip()
        
        if os.path.exists(temp_path):
            os.remove(temp_path)

        return jsonify({"text": transcription})
    
    except Exception as e:
        logger.error(f"Transcription error: {str(e)}")
        return jsonify({"error": "Audio processing failed", "detail": str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate_response():
    try:
        data = request.json
        user_text = data['text']
        session_id = request.headers.get('X-Session-Id', 'default')
        
        command_response = process_command(user_text)
        if command_response["command"] != "unknown":
            return jsonify({"response": command_response["text"]})
        
        response = chat_model(
            user_text,
            max_length=200,
            num_return_sequences=1,
            pad_token_id=chat_model.tokenizer.eos_token_id,
            no_repeat_ngram_size=3,
            do_sample=True,
            top_k=100,
            top_p=0.95,
            temperature=0.8
        )[0]['generated_text']
        
        response = response.replace(user_text, "").strip()
        
        return jsonify({"response": response})
    
    except Exception as e:
        logger.error(f"Response generation error: {str(e)}")
        return jsonify({"error": "Response generation failed", "detail": str(e)}), 500

@app.route('/synthesize', methods=['POST'])
def synthesize_speech():
    try:
        data = request.json
        text = data['text']
        lang = data.get('lang', 'en')
        session_id = request.headers.get('X-Session-Id', 'default')
        
        if len(text) > 500:
            return jsonify({"error": "Text too long (max 500 characters)"}), 400
        
        settings = user_settings.get(session_id, {})
        voice_gender = settings.get('voiceGender', 'female')
        voice_style = settings.get('voiceStyle', 'friendly')
        
        voice_name = VOICES.get(lang, 'en-US-AriaNeural')
        
        if voice_gender == 'male':
            voice_name = 'en-US-GuyNeural'
        elif voice_gender == 'neutral':
            voice_name = 'en-US-DavisNeural'
        
        if voice_style == 'professional':
            voice_name = 'en-US-AriaNeural'
        elif voice_style == 'cheerful':
            voice_name = 'en-US-JennyNeural'
        elif voice_style == 'calm':
            voice_name = 'en-US-AnaNeural'
        
        communicate = edge_tts.Communicate(text, voice_name)
        
        audio_buffer = io.BytesIO()
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        async def write_audio():
            async for message in communicate.stream():
                if message["type"] == "audio":
                    audio_buffer.write(message["data"])
        
        loop.run_until_complete(write_audio())
        loop.close()
        audio_buffer.seek(0)
        
        response = send_file(audio_buffer, mimetype='audio/mpeg')
        response.headers.add('Access-Control-Allow-Origin', '*')
        return response
    
    except Exception as e:
        logger.error(f"Speech synthesis error: {str(e)}")
        return jsonify({"error": "Speech synthesis failed", "detail": str(e)}), 500

@app.route('/settings', methods=['POST'])
def update_settings():
    try:
        data = request.json
        session_id = request.headers.get('X-Session-Id', 'default')
        
        user_settings[session_id] = data
        
        logger.info(f"Updated settings for session {session_id}")
        return jsonify({"status": "success"})
    
    except Exception as e:
        logger.error(f"Settings update error: {str(e)}")
        return jsonify({"error": "Settings update failed", "detail": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)