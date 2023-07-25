import speech_recognition as sr
import os
import win32com.client
import webbrowser
import datetime
from config import apikey
import openai


def say(text):
    speaker=win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold=1
        audio=r.listen(source)
        try:
            query=r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occurred"

chatStr = ""

def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"You: {query}\n MyAI: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        say(response["choices"][0]["text"])
        chatStr += f"{response['choices'][0]['text']}\n"
        return response["choices"][0]["text"]
    except Exception as e:
        return "Some error occurred"

def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )


    try:
        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")

        with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
            f.write(text)
    except Exception as e:
        return "Some error occurred"


if __name__ == '__main__':
    print('PyCharm')

    say("Hello")
    while True:
        print("Listening")
        query=takeCommand()
        if "quit" in query.lower():
            say("Exiting Program")
            exit()

        say(query)

        sites=[["youtube","https://www.youtube.com/"],["google","https://www.google.com/"],["wikipedia",        "https://www.wikipedia.com/"],["geeksforgeeks","https://www.geeksforgeeks.org/"]]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]}")
                webbrowser.open(site[1])

        if "open music" in query:
            musicPath="C:/Users/shrut/Downloads/jk.mp3"
            os.startfile(musicPath)

        elif "the time" in query:
            hr=datetime.datetime.now().strftime("%H")
            min=datetime.datetime.now().strftime("%M")
            say(f"Currently its {hr}Hours and {min}minutes")

        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)
