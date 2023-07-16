from gtts import gTTS
import datetime as datetime
import win32com.client
import speech_recognition as sr
import webbrowser
from pydub import AudioSegment
from pydub.playback import play
import pygame

say = win32com.client.Dispatch("SAPI.SpVoice")


def parroting():
    stpQuery = "Ok thankyou for using me"
    parr = sr.Recognizer()
    with sr.Microphone() as source:
        while True:
            print("listing..")
            audio = parr.listen(source)
            try:
                print("Recognizing...")
                query = parr.recognize_google(audio, language="en-in")
                print(f"User said: {query}")

                if "stop" in query.lower():
                    tts = gTTS(text=stpQuery, lang="hi", slow=False)
                    tts.save("query.mp3")
                    play_audio("query.mp3")
                    return

                tts = gTTS(text=query, lang="en", slow=False)
                tts.save("query.mp3")
                play_audio("query.mp3")



            except Exception as e:
                return "Some Error Occurred. Sorry from robust"


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")

            return query

        except Exception as e:
            return "Some Error Occurred. Sorry from MY side"


def play_music(query):
    say.Speak("Sure, I can play some music for you.")
    say.Speak("Please choose a music track number to play displayed below:")
    for track in music:
        print(track)

    chosen_track = input("Enter the track number: ")

    # Find the chosen track in the music list
    selected_track = None
    for track in music:
        if track[0] == chosen_track:
            selected_track = track[1]
            break
    if selected_track:
        pygame.mixer.init()
        pygame.mixer.music.load(selected_track)
        pygame.mixer.music.play()
        # Add an optional message to confirm the playing of the chosen track
        say.Speak(f"Now playing: the track you selected")

        # Wait for user input to stop the music
        stop_music = input("Enter 'stop' to stop the music: ").lower()
        if stop_music == "stop":
            pygame.mixer.music.stop()
            say.Speak("Music stopped.")
    else:
        say.Speak("I'm sorry, the selected track number is not valid.")


def play_audio(filename):
    audio = AudioSegment.from_file(filename)
    play(audio)


while True:
    print("ROBUST started")
    say.Speak("Hey SickCoder, welcome! I am Your Personal voice assistant, robust!")
    while True:
        print("Listening...")
        query = takeCommand()

        music = [["1", "NCS Project/beyond-the-horizon-136339.mp3"], ["2", "NCS Project/downfall.mp3"],
                 ["3", "NCS Project/holding-on-136343.mp3"],
                 ["4", "NCS Project/indian-music-with-sitar-tanpura-and-sarangi-74577.mp3"],
                 ["5", "NCS Project/indian-percussion-ethnic-drums-156542.mp3"],
                 ["6", "NCS Project/indian-trap-132594.mp3"],
                 ["7", "NCS Project/jonathan-gaming-143999.mp3"],["8", "NCS Project/keep-moving-on-136341.mp3"],
                 ["9", "NCS Project/love_bgm_no_copyright_music-113843.mp3"],
                 ["10", "NCS Project/silicon-valley-123990.mp3"],["11", "NCS Project/speed-122837.mp3"],
                 ["12", "NCS Project/tibet-13636.mp3"],["13", "NCS Project/time-to-on-136337.mp3"],
                 ["14", "NCS Project/vinee-heights-126947.mp3"]]

        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"],["Netflix", "https://www.netflix.com/in/"],
                 ["Amazon", "https://www.amazon.com/"],["Flipkart", "https://www.flipkart.com/"],
                 ["Reddit", "https://www.reddit.com/"],["Facebook", "https://www.facebook.com/"],
                 ["Instagram", "https://www.instagram.com/"],["twitter", "https://twitter.com/i/flow/login?redirect_after_login=%2F "],]

        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say.Speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        if "music" in query.lower() or "rock" in query.lower() or "pop" in query.lower():
            play_music(query)

        if "mimic me" in query.lower() or "copy me" in query.lower():
            parroting()

        if "the time" in query:
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            print(current_time)
            say.Speak(f"Sir, the current time is {current_time}")

        if "get out" in query.lower() or "stop now" in query.lower() or "bye" in query.lower():
            say.Speak("Thankyou, sir. If you need me another time, please say 'Hey robust'.")
            exit()

        say.Speak("Anything else, sir?")
