import streamlit as st
import speech_recognition as sr
import threading
import keyboard
import pyautogui
import win32com.client
import subprocess
import webbrowser
import time
import pygetwindow as gw

ALLOWED_APPS = ["chrome", "notepad", "word"]

def open_application(application_name):
    try:
        if "chrome" in application_name.lower():
            st.success("Opening Chrome...")
            # Add your application-specific code here
        elif "notepad" in application_name.lower():
            st.success("Opening Notepad...")
            st.write("Listening for text to type in Notepad...")
            subprocess.Popen(["notepad.exe"])
        elif "word" in application_name.lower():
            st.success("Opening Word...")
            st.write("Listening for text to type in Word...")
            word_path = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"  # Adjust the path accordingly
            subprocess.Popen([word_path])
        else:
            st.error(f"Application '{application_name}' not recognized.")
    except Exception as e:
        st.error(f"Error opening application: {e}")


def type_in_application(application_name, text):
    if "notepad" in application_name.lower():
        notepad_window = gw.getWindowsWithTitle("Untitled - Notepad")
        if notepad_window:
            notepad_window[0].activate()
            pyautogui.write(text, interval=0.1)
        else:
            st.warning("Notepad window not found.")
    elif "word" in application_name.lower():
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Add()
        word.Visible = True
        pyautogui.write(text, interval=0.1)

def voice_recognition():
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        st.success("Listening for a command...")
        recognizer.adjust_for_ambient_noise(source)

        while True:
            try:
                audio = recognizer.listen(source, timeout=5)  # Set timeout to 5 seconds
                command = recognizer.recognize_google(audio).lower()
                st.write(f"Command: {command}")

                if command.startswith("open"):
                    application_name = command[5:]  # Extract the application name after "open"
                    open_application(application_name)
                    audio_text = recognizer.listen(source, timeout=5)  # Set timeout to 5 seconds
                    text_to_type = recognizer.recognize_google(audio_text).lower()
                    st.write(f"Typed text: {text_to_type}")
                    type_in_application(application_name, text_to_type)
                else:
                    st.write("Searching Google...")
                    search_google(command)

            except sr.UnknownValueError:
                st.warning("Sorry, I could not understand the audio.")
            except sr.WaitTimeoutError:
                st.warning("Timeout: No speech detected. Continuing to listen...")
            except sr.RequestError as e:
                st.error(f"Could not request results from Google Speech Recognition service; {e}")

            # Check if the user pressed the 'q' key to terminate the program
            if keyboard.is_pressed('q'):
                st.warning("Terminating program...")
                break

    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        st.success("Listening for a command...")
        recognizer.adjust_for_ambient_noise(source)

        while True:
            try:
                audio = recognizer.listen(source, timeout=5)  # Set timeout to 5 seconds
                command = recognizer.recognize_google(audio).lower()
                st.write(f"Command: {command}")

                if command.startswith("open"):
                    application_name = command[5:]  # Extract the application name after "open"
                    open_application(application_name)
                    audio_text = recognizer.listen(source, timeout=5)  # Set timeout to 5 seconds
                    text_to_type = recognizer.recognize_google(audio_text).lower()
                    st.write(f"Typed text: {text_to_type}")
                    type_in_application(application_name, text_to_type)
                else:
                    st.write("Searching Google...")
                    search_google(command)

            except sr.UnknownValueError:
                st.warning("Sorry, I could not understand the audio.")
            except sr.WaitTimeoutError:
                st.warning("Timeout: No speech detected. Continuing to listen...")
            except sr.RequestError as e:
                st.error(f"Could not request results from Google Speech Recognition service; {e}")

            # Check if the user pressed the 'q' key to terminate the program
            if keyboard.is_pressed('q'):
                st.warning("Terminating program...")
                break

def search_google(query):
    # Check if the query seems like a general question
    if "who" in query or "what" in query or "where" in query or "when" in query or "why" in query:
        webbrowser.open(f"https://www.google.com/search?q={query}")
    else:
        webbrowser.open(f"https://www.google.com/search?q={query}")

if __name__ == "__main__":
    st.title("Voice Command Program")
    st.subheader("Click the microphone to start listening...")

    if st.button("🎤"):
        thread = threading.Thread(target=voice_recognition)
        thread.start()
        # Do not join the thread here; let it run in the background
