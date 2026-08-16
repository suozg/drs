import argparse
import wx
from config import frame_title, db_path
from ui_main import DocumentSearchFrame

class DocumentSearchApp(wx.App):
    def OnInit(self):
        global db_path
        parser_file_db = argparse.ArgumentParser(description='Process some database file.')
        parser_file_db.add_argument('-c', type=str, default=db_path, help='Path to the database file')
        args_db = parser_file_db.parse_args()
        
        # Оновлюємо шлях у глобальному модулі
        import config
        config.db_path = args_db.c
        
        frame = DocumentSearchFrame(None, title=frame_title)
        self.SetTopWindow(frame)
        return True

if __name__ == '__main__':
    app = DocumentSearchApp(0)
    app.MainLoop()
