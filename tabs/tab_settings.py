# tabs/tab_settings.py
import threading
import os
from sqlcipher3 import dbapi2 as sqlite3
import wx
import config
from settings_db import (
    get_databases_list, 
    remove_database_from_settings, 
    add_database_to_settings, 
    verify_database_password
)
from ui_dialogs import PasswordDialog, ConfirmPasswordDialog
from database import create_new_database

class SettingsTab(wx.Panel):
    def __init__(self, parent, main_frame):
        super().__init__(parent)
        self.main_frame = main_frame
        self.selected_db_path = None
        self.selected_db_name = None
        
        self.init_ui()

    def init_ui(self):
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # --- СЕКЦІЯ 1: Управління базами даних ---
        db_box = wx.StaticBox(self, label="Підключені бази даних (оберіть кліком для зміни пароля)")
        db_sizer = wx.StaticBoxSizer(db_box, wx.VERTICAL)

        self.db_list_ctrl = wx.ListCtrl(db_box, style=wx.LC_REPORT | wx.BORDER_SUNKEN | wx.LC_SINGLE_SEL, size=(-1, 140))
        self.db_list_ctrl.InsertColumn(0, "Назва", width=150)
        self.db_list_ctrl.InsertColumn(1, "Шлях до файлу", width=380)
        self.db_list_ctrl.InsertColumn(2, "Статус", width=90)
        self.db_list_ctrl.Bind(wx.EVT_LIST_ITEM_SELECTED, self.on_db_selected)
        db_sizer.Add(self.db_list_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        create_btn = wx.Button(db_box, label="Створити базу")
        create_btn.Bind(wx.EVT_BUTTON, self.on_create_database)
        btn_sizer.Add(create_btn, 0, wx.RIGHT, 5)
        
        add_btn = wx.Button(db_box, label="Додати базу")
        add_btn.Bind(wx.EVT_BUTTON, self.on_add_database)
        btn_sizer.Add(add_btn, 0, wx.RIGHT, 5)

        refresh_btn = wx.Button(db_box, label="Оновити список")
        refresh_btn.Bind(wx.EVT_BUTTON, lambda e: self.load_databases_into_ui())
        btn_sizer.Add(refresh_btn, 0, wx.RIGHT, 5)

        delete_btn = wx.Button(db_box, label="Видалити базу")
        delete_btn.Bind(wx.EVT_BUTTON, self.on_delete_database)
        btn_sizer.Add(delete_btn, 0, wx.RIGHT, 5)
        
        db_sizer.Add(btn_sizer, 0, wx.ALL, 5)
        main_sizer.Add(db_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # --- СЕКЦІЯ 2: Безпека та зміна паролів ---
        pass_box = wx.StaticBox(self, label="Керування паролями")
        pass_sizer = wx.StaticBoxSizer(pass_box, wx.VERTICAL)

        # Радіокнопки вибору цілі для зміни пароля
        radio_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.radio_master = wx.RadioButton(pass_box, label="Змінити Майстер-пароль програми", style=wx.RB_GROUP)
        self.radio_db = wx.RadioButton(pass_box, label="Змінити пароль обраної бази даних:")
        
        self.radio_master.Bind(wx.EVT_RADIOBUTTON, self.on_pass_target_changed)
        self.radio_db.Bind(wx.EVT_RADIOBUTTON, self.on_pass_target_changed)

        radio_sizer.Add(self.radio_master, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 20)
        radio_sizer.Add(self.radio_db, 0, wx.ALIGN_CENTER_VERTICAL, 5)
        
        # Динамічний напис з назвою обраної бази
        self.lbl_selected_db = wx.StaticText(pass_box, label="(не обрано)")
        font = self.lbl_selected_db.GetFont()
        font.SetWeight(wx.FONTWEIGHT_BOLD)
        self.lbl_selected_db.SetFont(font)
        radio_sizer.Add(self.lbl_selected_db, 1, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 5)

        pass_sizer.Add(radio_sizer, 0, wx.EXPAND | wx.ALL, 8)

        # Поля введення паролів
        change_pass_panel = wx.Panel(pass_box)
        change_pass_sizer = wx.FlexGridSizer(rows=3, cols=2, vgap=8, hgap=10)
        change_pass_panel.SetSizer(change_pass_sizer)

        change_pass_sizer.Add(wx.StaticText(change_pass_panel, label="Введіть старий пароль:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 3)
        self.old_pass_entry = wx.TextCtrl(change_pass_panel, style=wx.TE_PASSWORD, size=(200, -1))
        change_pass_sizer.Add(self.old_pass_entry, 1, wx.EXPAND | wx.ALL, 3)

        change_pass_sizer.Add(wx.StaticText(change_pass_panel, label="Введіть новий пароль:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 3)
        self.new_pass1_entry = wx.TextCtrl(change_pass_panel, style=wx.TE_PASSWORD, size=(200, -1))
        change_pass_sizer.Add(self.new_pass1_entry, 1, wx.EXPAND | wx.ALL, 3)

        change_pass_sizer.Add(wx.StaticText(change_pass_panel, label="Повторіть новий пароль:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 3)
        self.new_pass2_entry = wx.TextCtrl(change_pass_panel, style=wx.TE_PASSWORD, size=(200, -1))
        change_pass_sizer.Add(self.new_pass2_entry, 1, wx.EXPAND | wx.ALL, 3)

        button_container_sizer = wx.BoxSizer(wx.VERTICAL)
        self.change_pass_button = wx.Button(pass_box, label="Змінити пароль")
        self.change_pass_button.Bind(wx.EVT_BUTTON, self.on_change_password)
        button_container_sizer.Add(self.change_pass_button, 1, wx.EXPAND | wx.ALL, 5)

        top_section_sizer = wx.BoxSizer(wx.HORIZONTAL)
        top_section_sizer.Add(change_pass_panel, 1, wx.EXPAND | wx.ALL, 5)
        top_section_sizer.Add(button_container_sizer, 0, wx.ALL | wx.ALIGN_BOTTOM, 5)

        pass_sizer.Add(top_section_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Лог операцій
        self.password_progress_text = wx.TextCtrl(pass_box, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL, size=(-1, 90))
        pass_sizer.Add(self.password_progress_text, 1, wx.EXPAND | wx.ALL, 5)

        main_sizer.Add(pass_sizer, 1, wx.EXPAND | wx.ALL, 10)

        self.SetSizer(main_sizer)
        
        self.load_databases_into_ui()
        self.on_pass_target_changed(None)

    def load_databases_into_ui(self):
        self.db_list_ctrl.DeleteAllItems()
        
        if not hasattr(config, 'master_password') or not config.master_password:
            return

        databases = get_databases_list(config.master_password)
        unavailable_count = 0
        
        for row in databases:
            db_id, name, path, password, is_active = row
            
            # Перевіряємо фізичну доступність файлу бази на диску
            file_exists = os.path.exists(path)
            
            if not file_exists:
                status_text = "Недоступна"
                unavailable_count += 1
            else:
                status_text = "Активна" if is_active else "Вимкнена"

            index = self.db_list_ctrl.InsertItem(self.db_list_ctrl.GetItemCount(), name)
            self.db_list_ctrl.SetItem(index, 1, path)
            self.db_list_ctrl.SetItem(index, 2, status_text)
            
            if not file_exists:
                self.db_list_ctrl.SetItemTextColour(index, wx.Colour(200, 0, 0)) # червоний для недоступних
                
            self.db_list_ctrl.SetItemData(index, db_id)

        # Автоматично обираємо перший рядок, якщо він є
        if self.db_list_ctrl.GetItemCount() > 0:
            self.db_list_ctrl.Select(0)

        # Виводимо статус у головне вікно
        if hasattr(self.main_frame, 'set_status'):
            total_db = len(databases)
            if unavailable_count > 0:
                self.main_frame.set_status(f"Список оновлено: баз: {total_db}, недоступних файлів: {unavailable_count}")
            else:
                self.main_frame.set_status(f"Список оновлено: усі {total_db} баз доступні")

    def on_db_selected(self, event):
        index = event.GetIndex()
        self.selected_db_name = self.db_list_ctrl.GetItemText(index, 0)
        self.selected_db_path = self.db_list_ctrl.GetItemText(index, 1)
        self.lbl_selected_db.SetLabel(f"[{self.selected_db_name}]")

    def on_pass_target_changed(self, event):
        is_db_selected = self.radio_db.GetValue()
        self.lbl_selected_db.Enable(is_db_selected)

    def on_delete_database(self, event):
        selected_item = self.db_list_ctrl.GetFirstSelected()
        if selected_item == -1:
            wx.MessageBox("Виберіть базу даних зі списку.", "Попередження", wx.OK | wx.ICON_WARNING)
            return

        db_id = self.db_list_ctrl.GetItemData(selected_item)
        db_name = self.db_list_ctrl.GetItemText(selected_item, 0)

        confirm = wx.MessageBox(f"Видалити '{db_name}' зі списку?", "Підтвердження", wx.YES_NO | wx.ICON_QUESTION)
        if confirm == wx.YES:
            if hasattr(config, 'master_password'):
                remove_database_from_settings(config.master_password, db_id)
            if hasattr(self.main_frame, 'set_status'):
                self.main_frame.set_status(f"Базу даних '{db_name}' видалено зі списку") 
            self.refresh_all_views()

    def on_add_database(self, event):
        with wx.FileDialog(self, "Виберіть файл бази даних", wildcard="SQLCipher files (*.db)|*.db",
                            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fd:
            if fd.ShowModal() == wx.ID_CANCEL:
                return
            
            db_path = fd.GetPath()
            db_name = os.path.basename(db_path)

            dlg = PasswordDialog(self, f"Введіть пароль для {db_name}:", "Пароль бази даних")
            if dlg.ShowModal() == wx.ID_OK:
                db_password = dlg.GetValue().strip()
                dlg.Destroy()

                is_valid, error_message = verify_database_password(db_path, db_password)
                if not is_valid:
                    wx.MessageBox(f"Помилка перевірки пароля: {error_message}", "Помилка", wx.OK | wx.ICON_ERROR)
                    return

                try:
                    add_database_to_settings(config.master_password, db_name, db_path, db_password)
                    if hasattr(self.main_frame, 'set_status'):
                        self.main_frame.set_status(f"Базу даних '{db_name}' успішно додано") 
                    self.refresh_all_views()
                except Exception as e:
                    wx.MessageBox(f"Помилка: {e}", "Помилка", wx.OK | wx.ICON_ERROR)

    def on_create_database(self, event):
        with wx.FileDialog(self, "Створити новий файл бази даних", wildcard="SQLCipher files (*.db)|*.db",
                            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fd:
            if fd.ShowModal() == wx.ID_CANCEL:
                return
            
            db_path = fd.GetPath()
            if not db_path.lower().endswith('.db'):
                db_path += '.db'
            
            db_name = os.path.basename(db_path)

            # Використовуємо правильний клас ConfirmPasswordDialog з ui_dialogs.py
            dlg = ConfirmPasswordDialog(self, title=f"Пароль нової бази даних: {db_name}")
            if dlg.ShowModal() == wx.ID_OK:
                db_password = dlg.GetPassword().strip()
                dlg.Destroy()

                if not db_password:
                    wx.MessageBox("Пароль не може бути порожнім!", "Помилка", wx.OK | wx.ICON_ERROR)
                    return

                success, error_msg = create_new_database(db_path, db_password)
                if not success:
                    wx.MessageBox(f"Не вдалося створити базу даних: {error_msg}", "Помилка", wx.OK | wx.ICON_ERROR)
                    return

                try:
                    add_database_to_settings(config.master_password, db_name, db_path, db_password)
                    if hasattr(self.main_frame, 'set_status'):
                        self.main_frame.set_status(f"Створено та додано нову базу: {db_name}")
                    self.refresh_all_views()
                    wx.MessageBox(f"Базу даних '{db_name}' успішно створено та додано!", "Успіх", wx.OK | wx.ICON_INFORMATION)
                except Exception as e:
                    wx.MessageBox(f"Базу створено, але сталася помилка при додаванні до налаштувань: {e}", "Попередження", wx.OK | wx.ICON_WARNING)
            else:
                dlg.Destroy()

    def refresh_all_views(self):
        self.load_databases_into_ui()
        for tab_attr in ['tab_import', 'tab_sql', 'tab_search']:
            if hasattr(self.main_frame, tab_attr):
                tab = getattr(self.main_frame, tab_attr)
                if hasattr(tab, 'load_databases_to_choice'):
                    tab.load_databases_to_choice()

    def on_change_password(self, event):
        old_pass = self.old_pass_entry.GetValue()
        new_pass1 = self.new_pass1_entry.GetValue()
        new_pass2 = self.new_pass2_entry.GetValue()
        is_master = self.radio_master.GetValue()

        if not old_pass or not new_pass1 or not new_pass2:
            self.log_password_progress("Будь ласка, заповніть усі поля.")
            return

        if new_pass1 != new_pass2:
            self.log_password_progress("Новий пароль та підтвердження не збігаються.")
            return

        if new_pass1 == old_pass:
            self.log_password_progress("Новий пароль не може бути таким самим, як старий.")
            return

        self.log_password_progress("Зміна пароля...")
        
        if hasattr(self.main_frame, 'notebook'):
            self.main_frame.notebook.Enable(False)
            
        self.change_pass_button.Enable(False)
        
        if is_master:
            threading.Thread(target=self._perform_master_password_change, args=(old_pass, new_pass1)).start()
        else:
            if not self.selected_db_path:
                self.log_password_progress("Помилка: не обрано базу даних у таблиці зверху.")
                self._restore_ui()
                return
            threading.Thread(target=self._perform_db_password_change, args=(self.selected_db_path, old_pass, new_pass1)).start()

    def _perform_master_password_change(self, old_pass, new_pass):
        conn = None
        try:
            settings_db = getattr(self.main_frame, 'settings_db_path', config.SETTINGS_DB_PATH)
            
            conn = sqlite3.connect(settings_db)
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA key = '{old_pass}';")
            cursor.execute("PRAGMA cipher_compatibility = 3;")
            
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            cursor.fetchall()
            
            cursor.execute(f"PRAGMA rekey = '{new_pass}';")
            conn.commit()
            conn.close()
            conn = None

            if hasattr(self.main_frame, 'master_password'):
                self.main_frame.master_password = new_pass

            wx.CallAfter(self.log_password_progress, "Майстер-пароль успішно змінено!")
            self._clear_entries()

        except sqlite3.DatabaseError as e:
            if "file is encrypted or is not a database" in str(e) or "not an error" in str(e):
                wx.CallAfter(self.log_password_progress, "Помилка: невірний старий майстер-пароль.")
            else:
                wx.CallAfter(self.log_password_progress, f"Помилка бази даних при зміні майстер-пароля: {e}")
        except Exception as e:
            wx.CallAfter(self.log_password_progress, f"Помилка при зміні майстер-пароля: {e}")
        finally:
            if conn:
                conn.close()
            self._restore_ui()

    def _perform_db_password_change(self, db_path, old_pass, new_pass):
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA key = '{old_pass}';")
            cursor.execute("PRAGMA cipher_compatibility = 3;")
            cursor.execute(f"PRAGMA rekey = '{new_pass}';")
            conn.commit()
            conn.close()
            conn = None
            master_pass = getattr(self.main_frame, 'master_password', None)
            settings_db = getattr(self.main_frame, 'settings_db_path', config.SETTINGS_DB_PATH)
            
            if master_pass:
                s_conn = sqlite3.connect(settings_db)
                s_cursor = s_conn.cursor()
                s_cursor.execute(f"PRAGMA key = '{master_pass}';")
                s_cursor.execute("PRAGMA cipher_compatibility = 3;")
                
                s_cursor.execute("UPDATE registered_databases SET password = ? WHERE path = ?", (new_pass, db_path))
                s_conn.commit()
                s_conn.close()
            
            wx.CallAfter(self.log_password_progress, "Пароль бази даних та запис у налаштуваннях успішно оновлено!")
            self._clear_entries()
            wx.CallAfter(self.refresh_all_views)
            
        except sqlite3.DatabaseError as e:
            if "file is encrypted or is not a database" in str(e) or "not an error" in str(e):
                wx.CallAfter(self.log_password_progress, "Помилка: невірний старий пароль пароля бази даних.")
            else:
                wx.CallAfter(self.log_password_progress, f"Помилка бази даних при зміні пароля: {e}")
        except Exception as e:
            wx.CallAfter(self.log_password_progress, f"Несподівана помилка при зміні пароля БД: {e}")
        finally:
            if conn: 
                conn.close()
            self._restore_ui()

    def _clear_entries(self):
        wx.CallAfter(self.old_pass_entry.Clear)
        wx.CallAfter(self.new_pass1_entry.Clear)
        wx.CallAfter(self.new_pass2_entry.Clear)

    def _restore_ui(self):
        wx.CallAfter(self.change_pass_button.Enable, True)
        if hasattr(self.main_frame, 'notebook'):
            wx.CallAfter(self.main_frame.notebook.Enable, True)

    def log_password_progress(self, message):
        wx.CallAfter(self.password_progress_text.AppendText, f"{message}\n")
