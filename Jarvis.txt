import os
import win32com.client
import speech_recognition as sr
import webbrowser
import datetime
import psutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
from webdriver_manager.chrome import ChromeDriverManager
import google.generativeai as genai


genai.configure(api_key="AIzaSyDiPszjzAeLYh4MIiQy3PE5mlNXDwDK-Tc")


def take_command():

    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.pause_threshold = 1
        try:
            audio = recognizer.listen(source)
            query = recognizer.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            return "I couldn't understand your query."
        except sr.RequestError as e:
            print(f"Could not request results: {e}")
            return "Speech recognition service is unavailable."


def say(text):

    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)


def search_with_gemini(question):

    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        response = model.generate_content(question)
        return response.text.strip()
    except Exception as e:
        print(f"Error with Gemini API: {e}")
        return "I encountered an error with the AI service."


def open_site(query, sites):

    for keyword, url in sites.items():
        if f"open {keyword}" in query.lower():
            say(f"Opening {keyword}...")
            webbrowser.open(url)
            return True
    return False


def close_application(query, apps):

    for keyword, process_name in apps.items():
        if f"close {keyword}" in query.lower():
            found = False
            try:
                for proc in psutil.process_iter(['pid', 'name', 'cmdline']):

                    if keyword.lower() == "youtube":
                        if 'chrome' in proc.info['name'].lower() or 'firefox' in proc.info['name'].lower():
                            if any('youtube' in cmd.lower() for cmd in proc.info['cmdline']):
                                proc.terminate()
                                found = True
                                print(f"Closed {keyword} in browser (PID: {proc.info['pid']})")
                                break
                    else:

                        if process_name in proc.info['name'].lower() or (
                                proc.info['cmdline'] and any(
                            process_name in cmd.lower() for cmd in proc.info['cmdline'])
                        ):
                            proc.terminate()
                            found = True
                            print(f"Closed {process_name} (PID: {proc.info['pid']})")
                            break
            except psutil.NoSuchProcess:
                say(f"The process {process_name} was not found or is already closed.")
                print(f"{process_name} is not running.")
            except Exception as e:
                say(f"An error occurred while closing {keyword}: {str(e)}")
                print(f"Error closing {keyword}: {e}")

            if found:
                say(f"Closing {keyword}...")
            else:
                say(f"{keyword} is not running.")
            return True
    return False


def play_music(song_name):
    """Play music on YouTube by searching for the song."""
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)

        driver.get(f"https://www.youtube.com/results?search_query={song_name.replace(' ', '+')}")
        time.sleep(2)  # Wait for the page to load

        # Find the first video element on the search results page and click it
        video = driver.find_element(By.XPATH, "//a[@id='video-title']")
        video.click()

        say(f"Playing {song_name} on YouTube.")
    except Exception as e:
        say(f"An error occurred while playing the song: {str(e)}")
        print(f"Error while playing song: {e}")


def main():
    """Main function to execute the virtual assistant loop."""
    say("Hi Ironman, I am Jarvis. How can I assist you today?")
    sites = {
        "youtube": "https://youtube.com",
        "instagram": "https://instagram.com",
        "whatsapp": "https://web.whatsapp.com",
        "facebook": "https://facebook.com",
        "twitter": "https://twitter.com",
        "linkedin": "https://linkedin.com",
        "google": "https://google.com",
        "wikipedia": "https://wikipedia.org",
        "spotify": "https://spotify.com"
    }
    apps = {
        "chrome": "chrome",
        "spotify": "spotify",
        "youtube": "chrome"
    }

    running = True
    while running:
        query = take_command()

        if "stop" in query.lower():
            say("Stopping, Goodbye!")
            running = False
            break

        if query.lower().startswith("jarvis"):
            answer = search_with_gemini(query.replace("jarvis", "").strip())
            say(answer)
            print(f"Gemini AI Response: {answer}")
            continue

        if open_site(query, sites):
            continue

        if close_application(query, apps):
            continue

        if "the time" in query.lower():
            time_str = datetime.datetime.now().strftime("%I:%M %p")
            say(f"Sir, the time is {time_str}.")
            continue

        if "open chrome" in query.lower():
            chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            if os.path.exists(chrome_path):
                say("Opening Chrome...")
                os.startfile(chrome_path)
            else:
                say("Chrome is not installed in the expected location.")
            continue

        if "play" in query.lower() and "song" in query.lower():
            song_name = query.lower().replace("play", "").replace("song", "").strip()
            play_music(song_name)
            continue

        say("I couldn't process your request. Please try again.")


if __name__ == "__main__":
    main()
