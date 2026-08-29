# main_frame.py
import wx
import os
import config
from ui_dialogs import PasswordDialog, ConfirmPasswordDialog
from settings_db import SETTINGS_DB_NAME, init_settings_db, add_database_to_settings

# Імпортуємо наші вкладки
from tabs.tab_search import SearchTab
from tabs.tab_import import ImportTab
from tabs.tab_sql import SqlTab
from tabs.tab_about import AboutTab
from tabs.tab_settings import SettingsTab

class DrsMainFrame(wx.Frame):
    def __init__(self, parent, title):
        super(DrsMainFrame, self).__init__(parent, title=title, size=(800, 640))

        self.db_info = {
            "records": 0,
            "last_modified": "Недоступно",
            "last_update_app_time": "Не оновлювалось"
        }
        self.is_scanning_active = False

        # Спочатку запитуємо пароль
        if not self.prompt_for_password():
            self.Close()
            return

        # Створюємо інтерфейс тільки після успішного введення пароля
        self.InitUI()
        self.Centre()
        self.Show() 
    
    def InitUI(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        self.notebook = wx.Notebook(panel)
        main_sizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 5)
        
        # Footer
        footer_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.file_label = wx.StaticText(panel, label="Готово")
        font = self.file_label.GetFont()
        font.SetPointSize(8) 
        self.file_label.SetFont(font)
        footer_sizer.Add(self.file_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.BOTTOM, 2)
        main_sizer.Add(footer_sizer, 0, wx.EXPAND | wx.LEFT, 5)
        
        panel.SetSizer(main_sizer)

        # --- Створюємо та додаємо вкладки (передаємо self як головне вікно) ---
        self.tab_search = SearchTab(self.notebook, self)
        self.tab_import = ImportTab(self.notebook, self)
        self.tab_sql = SqlTab(self.notebook, self)
        self.tab_about = AboutTab(self.notebook, self)
        self.tab_settings = SettingsTab(self.notebook, self)

        self.notebook.AddPage(self.tab_search, "Пошук")
        self.notebook.AddPage(self.tab_import, "Імпорт")
        self.notebook.AddPage(self.tab_sql, "SQL") 
        self.notebook.AddPage(self.tab_settings, "Налаштування")   
        self.notebook.AddPage(self.tab_about, "Про")

        # Прив'язка подій notebook
        self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGING, self.on_notebook_page_changing)

    def set_status(self, text):
        """Оновлює текст у нижньому рядку сповіщень"""
        if hasattr(self, 'file_label'):
            wx.CallAfter(self.file_label.SetLabel, text)

    def on_notebook_page_changing(self, event):
        if getattr(self, 'is_scanning_active', False):
            wx.MessageBox("Будь ласка, дочекайтеся завершення процесу.", "Обробка активна", wx.OK | wx.ICON_INFORMATION)
            event.Veto()
        else:
            event.Skip()

    def prompt_for_password(self):
        from settings_db import get_databases_list, add_database_to_settings
        import sqlite3
        
        while True:
            is_first_run = not os.path.exists(SETTINGS_DB_NAME)
            
            if is_first_run:
                dlg = ConfirmPasswordDialog(self, title="Створення пароля")
            else:
                dlg = PasswordDialog(self, "Введіть пароль", "Пароль:")

            if dlg.ShowModal() == wx.ID_OK:
                if hasattr(dlg, "GetPassword"):
                    input_password = dlg.GetPassword().strip()
                elif hasattr(dlg, "GetValue"):
                    input_password = dlg.GetValue().strip()
                else:
                    input_password = ""
                    
                dlg.Destroy()
                
                if not input_password:
                    wx.MessageBox("Пароль не може бути порожнім.", "Помилка", wx.OK | wx.ICON_ERROR)
                    continue
                
                if is_first_run:
                    try:
                        init_settings_db(input_password)
                        config.master_password = input_password
                        
                        choice_dlg = wx.MessageDialog(
                            self, 
                            "Базу даних не знайдено. Бажаєте створити нову базу даних чи підключити існуючу?", 
                            "Початкове налаштування", 
                            wx.YES_NO | wx.ICON_QUESTION
                        )
                        choice_dlg.SetYesNoLabels("Створити нову", "Підключити існуючу")
                        
                        if choice_dlg.ShowModal() == wx.ID_YES:
                            choice_dlg.Destroy()
                            
                            with wx.FileDialog(self, "Створити нову базу даних", wildcard="SQLChiper files (*.db)|*.db",
                                                style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fd:
                                if fd.ShowModal() == wx.ID_CANCEL:
                                    return False
                                
                                db_path = fd.GetPath()
                                if not db_path.endswith('.db'):
                                    db_path += '.db'
                                db_name = os.path.basename(db_path)

                            pwd_dlg = ConfirmPasswordDialog(self, title=f"Пароль для нової бази: {db_name}")
                            if pwd_dlg.ShowModal() == wx.ID_OK:
                                db_password = pwd_dlg.GetPassword().strip() if hasattr(pwd_dlg, "GetPassword") else pwd_dlg.GetValue().strip()
                                pwd_dlg.Destroy()
                                
                                if not db_password:
                                    wx.MessageBox("Пароль бази не може бути порожнім.", "Помилка", wx.OK | wx.ICON_ERROR)
                                    return False
                                
                                config.db_path = db_path
                                config.password = db_password
                                
                                from database import connect_to_database
                                conn = connect_to_database(db_password)
                                if conn:
                                    conn.close()
                                
                                add_database_to_settings(input_password, db_name, db_path, db_password)
                            else:
                                pwd_dlg.Destroy()
                                return False
                        else:
                            choice_dlg.Destroy()
                        
                        break
                    except Exception as e:
                        wx.MessageBox(f"Помилка ініціалізації: {e}", "Помилка", wx.OK | wx.ICON_ERROR)
                        continue
                else:
                    config.master_password = input_password
                    try:
                        # Отримуємо список підключених баз із settings.db 
                        databases = get_databases_list(input_password)
                        
                        if not databases:
                            # Якщо список баз у settings.db взагалі порожній
                            choice_dlg = wx.MessageDialog(
                                self, 
                                "У налаштуваннях немає підключених баз даних. Створити нову чи підключити існуючу?", 
                                "База даних не знайдена", 
                                wx.YES_NO | wx.ICON_QUESTION
                            )
                            choice_dlg.SetYesNoLabels("Створити нову", "Підключити існуючу")
                            
                            if choice_dlg.ShowModal() == wx.ID_YES:
                                choice_dlg.Destroy()
                                with wx.FileDialog(self, "Створити нову базу даних", wildcard="SQLChiper files (*.db)|*.db",
                                                    style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fd:
                                    if fd.ShowModal() == wx.ID_CANCEL:
                                        return False
                                    db_path = fd.GetPath()
                                    if db_path.endswith('.db'):
                                        db_path += '.db'
                                    db_name = os.path.basename(db_path)

                                pwd_dlg = ConfirmPasswordDialog(self, title=f"Пароль для нової бази: {db_name}")
                                if pwd_dlg.ShowModal() == wx.ID_OK:
                                    db_password = pwd_dlg.GetPassword().strip() if hasattr(pwd_dlg, "GetPassword") else pwd_dlg.GetValue().strip()
                                    pwd_dlg.Destroy()
                                    
                                    config.db_path = db_path
                                    config.password = db_password
                                    
                                    from database import connect_to_database
                                    conn = connect_to_database(db_password)
                                    if conn:
                                        conn.close()
                                        
                                    add_database_to_settings(input_password, db_name, db_path, db_password)
                                else:
                                    pwd_dlg.Destroy()
                                    return False
                            else:
                                choice_dlg.Destroy()
                        else:
                            # Якщо бази є у списку — беремо першу активну (або першу у списку) за замовчуванням
                            # Формат рядка в databases: (id, name, path, password, is_active)
                            db_id, db_name, db_path, db_password, is_active = databases[0]
                            config.db_path = db_path
                            config.password = db_password
                        
                        break  
                    except (sqlite3.DatabaseError, Exception) as e:
                        wx.MessageBox(f"Невірний пароль програми або пошкоджений файл налаштувань!\n{e}", "Помилка авторизації", wx.OK | wx.ICON_ERROR)
                        config.master_password = ""
                        continue  
            else:
                dlg.Destroy()
                return False
        return True
