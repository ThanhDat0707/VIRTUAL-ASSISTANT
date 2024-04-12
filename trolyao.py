import os
import playsound
import speech_recognition as sr
import time
import sys
import ctypes
import wikipedia
import datetime
import json
import re
import webbrowser as wb
import smtplib
import requests
import urllib
import urllib.request as urllib2
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from time import strftime
from gtts import gTTS
from youtube_search import YoutubeSearch


#nói
def speak(text):
    print("Bot: {}".format(text))
    tts = gTTS(text=text, lang='vi')
    tts.save("sound.mp3")
    playsound.playsound("sound.mp3")
    os.remove("sound.mp3")

#Đọc
def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Tôi: ", end='')
        audio = r.record(source, duration=5)
        try:
            text = r.recognize_google(audio, language="vi")
            print(text)
            return text
        except:
            print("...")
            return 0

#Dừng
def stop():
    speak("Hẹn gặp lại chủ nhân sau!")

#lặp lại
def get_text():
    for i in range(3):
        text = get_audio()
        if text:
            return text.lower()
        elif i < 2:
            speak("Tiểu Cúc không nghe rõ. Ngài nói lại được không!")
    time.sleep(2)
    stop()
    return 0

#Chào hỏi
def hello(name):
    day_time = int(strftime('%H'))
    if day_time < 12:
        speak("Chào buổi sáng chủ nhân {}. Chúc ngài một ngày tốt lành.".format(name))
    elif 12 <= day_time < 18:
        speak("Chào buổi chiều chủ nhân {}. Ngài đã dự định gì cho chiều nay chưa.".format(name))
    else:
        speak("Chào buổi tối chủ nhân {}. Ngài đã ăn tối chưa nhỉ.".format(name))

# Thời gian
def get_time(text):
    now = datetime.datetime.now()
    if "giờ" in text:
        speak('Bây giờ là %d giờ %d phút' % (now.hour, now.minute))
    elif "ngày" in text:
        speak("Hôm nay là ngày %d tháng %d năm %d" %
              (now.day, now.month, now.year))
    else:
        speak("Tiểu Cúc chưa hiểu ý của chủ nhân. Ngài nói lại được không?")
        
#Đổi hình nền        
def change_wallpaper():
    api_key = 'RF3LyUUIyogjCpQwlf-zjzCf1JdvRwb--SLV6iCzOxw'
    url = 'https://api.unsplash.com/photos/random?client_id=' + \
        api_key  # pic from unspalsh.com
    f = urllib2.urlopen(url)
    json_string = f.read()
    f.close()
    parsed_json = json.loads(json_string)
    photo = parsed_json['urls']['full']
    # Location where we download the image to.
    urllib2.urlretrieve(photo, "C:/Users/luuth/Downloads/a.png")
    image=os.path.join("C:/Users/luuth/Downloads/a.png")
    ctypes.windll.user32.SystemParametersInfoW(20,0,image,3)
    speak('Hình nền máy tính vừa được thay đổi')

#Mở ứng dụng
def open_application(text):
    if "word" in text:
        speak("Mở Microsoft Word")
        os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
    elif "excel" in text:
        speak("Mở Microsoft Excel")
        os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
    elif "facebook" in text:
        speak("Mở facebook")
        url = f"https://www.facebook.com/"
        wb.get().open(url)
    elif "truyện tranh" in text:
        speak("Mở nhật truyện")
        url = f"http://nhattruyenbring.com/"
        wb.get().open(url)
    else:
        speak("Ứng dụng chưa được cài đặt. Chủ nhân hãy thử lại!")

#Đóng app
def close_application(text):
    if "đóng word" in text:
        os.system("TASKKILL /F /IM WINWORD.exe")
    elif "đóng excel" in text:
        os.system("TASKKILL /F /IM EXCEL.exe")
    elif "đóng cốc cốc" in text:
        os.system("TASKKILL /F /IM browser.exe")
    else:
        speak("Ứng dụng chưa được cài đặt. Chủ nhân hãy thử lại!")

#Thời tiết
def current_weather():
    speak("Chủ nhân muốn xem thời tiết ở đâu ạ.")
    ow_url = "http://api.openweathermap.org/data/2.5/weather?"
    city = get_text()
    if city:
        api_key = "2bbaf16f717edcf51c9d9768a8b31b27"
        call_url = ow_url + "appid=" + api_key + "&q=" + city + "&units=metric"
        response = requests.get(call_url)
        data = response.json()
        if data["cod"] != "404":
            city_res = data["main"]
            current_temperature = city_res["temp"]
            current_pressure = city_res["pressure"]
            current_humidity = city_res["humidity"]
            suntime = data["sys"]
            sunrise = datetime.datetime.fromtimestamp(suntime["sunrise"])
            sunset = datetime.datetime.fromtimestamp(suntime["sunset"])
            wthr = data["weather"]
            weather_description = wthr[0]["description"]
            now = datetime.datetime.now()
            content = """
            Hôm nay là ngày {day} tháng {month} năm {year}
            Mặt trời mọc vào {hourrise} giờ {minrise} phút
            Mặt trời lặn vào {hourset} giờ {minset} phút
            Nhiệt độ trung bình là {temp} độ C
            Áp suất không khí là {pressure} héc tơ Pascal
            Độ ẩm là {humidity}%
            Trời hôm nay quang mây. Dự báo mưa rải rác ở một số nơi.""".format(day = now.day,month = now.month, year= now.year, hourrise = sunrise.hour, minrise = sunrise.minute,
                                                                            hourset = sunset.hour, minset = sunset.minute, 
                                                                            temp = current_temperature, pressure = current_pressure, humidity = current_humidity)
            speak(content)
            time.sleep(20)
    else:
        speak("Không tìm thấy địa chỉ của chủ nhân")
#Chạy
speak("Xin chào, chủ nhân tên là gì nhỉ?")
name = get_text()
if name:
    speak("Chào chủ nhân {}".format(name))
    speak("Chủ nhân cần Tiểu Cúc giúp gì ạ?")
    while True:
        text = get_text()
        if not text:
            break
        elif "tạm biệt" in text:
            stop()
            break
        elif "chào" in text:
            hello(name)
        elif "giờ" in text or "ngày" in text:
            get_time(text)
        elif "thời tiết" in text:
            current_weather()
        elif "hình nền" in text:
            change_wallpaper()
        elif "đóng" in text:
            close_application(text)
        elif "word" in text or "excel" in text or "facebook" in text or "truyện tranh" in text:
            open_application(text)
        elif "tìm kiếm" in text:
            speak("Chủ nhân muốn kiếm gì trên Google?")
            search = get_text()
            url = f"https://www.google.com/search?q={search}"
            wb.get().open(url)
            speak(f"Kết quả tìm kiếm Google là {search}")
        elif "video" in text:
            speak("Chủ nhân muốn kiếm gì trên Youtube?")
            search = get_text()
            url = f"https://www.youtube.com/search?q={search}"
            wb.get().open(url)
            speak(f"Kết quả tìm kiếm Youtube là {search}")
        elif "tạm dừng" in text:
            speak("Ấn bất cứ phím nào để đánh thức Tiểu Cúc")
            os.system("pause")
        elif text:  #Tìm kiếm thông tin
            wikipedia.set_lang('vi')
            robot = wikipedia.summary(text, sentences = 1)
            speak(robot)
        else:
            speak("Chủ nhân cần Tiểu Cúc giúp gì ạ?")