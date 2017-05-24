import wx
import os
os.environ["HTTPS_PROXY"] = "http://username:pass@192.168.1.107:3128"
import wikipedia
import wolframalpha
import time
import webbrowser
import winshell
import json
import requests
import ctypes
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import speech_recognition as sr

headers = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'}
speak = wincl.Dispatch("SAPI.SpVoice")

#GUI creation
class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None,
            pos=wx.DefaultPosition, size=wx.Size(450, 100),
            style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION |
             wx.CLOSE_BOX | wx.CLIP_CHILDREN,
            title="BRUNO")
        panel = wx.Panel(self)

        ico = wx.Icon('boy2.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)

        my_sizer = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(panel,
        label="Welcome Sir, I'm the Python Digital Assistant. How can I help you?")
        my_sizer.Add(lbl, 0, wx.ALL, 5)
        self.txt = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER,size=(400,30))
        self.txt.SetFocus()
        self.txt.Bind(wx.EVT_TEXT_ENTER, self.OnEnter)
        my_sizer.Add(self.txt, 0, wx.ALL, 5)
        panel.SetSizer(my_sizer)
        self.Show()
        speak.Speak("Welcome Sir, how can I help")



    def OnEnter(self, event):
        put = self.txt.GetValue()
        put = put.lower()
        link=put.split()
        if put=='':
            r = sr.Recognizer()
            with sr.Microphone() as src:
                audio = r.listen(src)
            try:
                put=r.recognize_google(audio)
                put= put.lower()
                link=put.split()
                self.txt.SetValue(put)
             
            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
            except sr.RequestError as e:
                print("Could not request results from Google STT; {0}".format(e))
            except:
                print("Unknown exception occurred!")
                
#Open a webpage                
        if put.startswith('open '):
            try:
                speak.Speak("opening "+link[1])
                webbrowser.open('http://www.'+link[1]+'.com')
            except:
                print('Sorry, No Internet Connection!')
#Play Song on Youtube
        if put.startswith('play '):
            try:
                link='+'.join(link[1:])
                url='https://www.youtube.com/results?search_query='+link
                source_code = requests.get(url, headers=headers, timeout=15)
                plain_text=source_code.text
                soup=BeautifulSoup(plain_text,"html.parser")
                songs=soup.findAll('div',{'class':'yt-lockup-video'})
                song=songs[0].contents[0].contents[0].contents[0]
                hit=song['href']
                speak.Speak("playing "+link)
                webbrowser.open('https://www.youtube.com'+hit)


            except:
                print('Sorry, No internet connection!')
#Google Search
        if put.startswith('search '):
            try:
                link='+'.join(link[1:])
                #print(link)
                speak.Speak("searching on google for "+link)
                webbrowser.open('https://www.google.co.in/search?q='+link)
            except:
                print('Sorry, No internet connection!')

#Empty Recycle bin
        if put.startswith('empty '):
            try:
                winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                print("Recycle Bin Empty!!")
            except:
                print("Unknown Error")

#News
        if put.startswith('science '):
            try:
                jsonObj=urlopen('https://newsapi.org/v1/articles?source=new-scientist&sortBy=top&apiKey=1bb9352ea6964f539e39b431a3bcbda6')
                data=json.load(jsonObj)
                i=1
                print('             ================NEW SCIENTIST===============')
                for item in data['articles']:
                    print(str(i)+'. '+item['title']+'\n')
                    print(item['description']+'\n')
                    i+=1
            except:
                print('Sorry, No internet connection')

        if put.startswith('headlines '):
            try:
                jsonObj=urlopen('https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=1bb9352ea6964f539e39b431a3bcbda6')
                data=json.load(jsonObj)
                i=1
                print('             ===============TIMES OF INDIA===============')
                for item in data['articles']:
                    print(str(i)+'. '+item['title']+'\n')
                    print(item['description']+'\n')
                    i+=1
            except Exception as e:
                print(e)

#Lock the device
        if put.startswith('lock '):
            try:
                speak.Speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
            except Exception as e:
                print(str(e))
                
#Trigger GUI               
if __name__=="__main__":
    app=wx.App(True)
    frame=MyFrame()
    app.MainLoop()
