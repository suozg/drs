# main.py
import wx
from config import frame_title
from main_frame import DrsMainFrame

class DocumentSearchApp(wx.App):
    def OnInit(self):
        frame = DrsMainFrame(None, title=frame_title)
        self.SetTopWindow(frame)
        return True

if __name__ == '__main__':
    app = DocumentSearchApp(0)
    app.MainLoop()
