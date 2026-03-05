import wx
import sys
from gui import MainFrame

def ignore_sip_voidptr(exctype, value, traceback):
    if "sip.voidptr" in str(value):
        return 
    sys.__excepthook__(exctype, value, traceback)

sys.excepthook = ignore_sip_voidptr

class MyApp(wx.App):
    def OnInit(self):
        self.frame = MainFrame()
        self.frame.Show()
        return True

if __name__ == "__main__":
    app = MyApp(redirect=False)
    app.MainLoop()
