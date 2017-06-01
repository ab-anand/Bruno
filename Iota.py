import wx
import pyttsx
import wikipedia
import wolframalpha
import os
# os.environ["HTTPS_PROXY"] = "http://user:pass@192.168.1.107:3128"


def speak(value):
    engine = pyttsx.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[4].id)
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate-30)
    engine.say(value)
    engine.runAndWait()


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, pos=wx.DefaultPosition,
                          size=wx.Size(450, 100),
                          style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION |
                          wx.CLOSE_BOX | wx.CLIP_CHILDREN, title='KIARA')
        panel = wx.Panel(self)
        my_sizer = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(panel, label='''Hello I\'m Kiara the Python
                            Digital Assistant. How can I help you?''')
        my_sizer.Add(lbl, 0, wx.ALL, 5)
        self.txt = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER,
                               size=(400, 30))
        self.txt.SetFocus()
        self.txt.Bind(wx.EVT_TEXT_ENTER, self.OnEnter)
        my_sizer.Add(self.txt, 0, wx.ALL, 5)
        panel.SetSizer(my_sizer)
        self.Show()
        speak('Welcome my friend. I am Kiara. How can I help ?')

    def OnEnter(self, event):
        input = self.txt.GetValue()
        # input = input.lower()
        app_id = '2V3684-LXELTTTJ9J'
        try:
            # wolframalpha
            client = wolframalpha.Client(app_id)
            res = client.query(input)
            ans = next(res.results).text
            print(ans)
            speak(ans)
        except:
            # wikipedia
            input = input.split(' ')
            input = ' '.join(input[2:])
            print(wikipedia.summary(input))
            speak('Searched wikipedia for '+input)


if __name__ == "__main__":
    app = wx.App(True)
    frame = MyFrame()
    app.MainLoop()
