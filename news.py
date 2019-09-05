from win32com.client import Dispatch
import requests
import json

def speak(str):
    speak=Dispatch("SAPI.SpVoice")
    speak.Speak(str)

if __name__ == '__main__':
    speak("Today's top bulletin as follows ")
    url="https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=0d2b7cd3e22c4ba59973493686cd3a8b"
    news=requests.get(url).text
    news1=json.loads(news)
    print(news1["articles"])
    a=news1["articles"]
    i=1
    for article in a:

        speak(article["title"])
        speak(f"Time for another news no {i}")
        i+=1
