#tabs/tab_import.py
import os
import threading
import time
import hashlib
from datetime import datetime
import re
import sqlite3
import wx
import config
from config import document_number_pattern, filename_date_pattern, path_date_pattern
from database import connect_to_database, populate_databases_choice, connect_to_specific_database, check_patch_db
from settings_db import get_databases_list
from docx import Document  # Для читання .docx файлів
from utils import extract_text_libreoffice, get_document_date, normalize_text

class ImportTab(wx.Panel):
    def __init__(self, parent, main_frame):
        super().__init__(parent)
        self.main_frame = main_frame
        self.stop_processing = False
        self.is_scanning_active = False
        
        self.setup_import_tab()
        self.load_databases_to_choice()

    def setup_import_tab(self):
        import_sizer = wx.BoxSizer(wx.VERTICAL)

        info_panel = wx.Panel(self)
        info_sizer = wx.BoxSizer(wx.VERTICAL)
        info_panel.SetSizer(info_sizer)

        info_sizer.Add(wx.StaticText(info_panel, 
                label="УВАГА! Дозволені тільки документи .doc, .docx, .rtf. Перевірте каталог на відсутність інших форматів.",
                style=wx.ALIGN_LEFT), 0, wx.ALL, 5)

        import_sizer.Add(info_panel, 0, wx.EXPAND | wx.ALL, 10)

        # Створюємо мітку та випадаючий список для вибору бази
        import_db_sizer = wx.BoxSizer(wx.HORIZONTAL)
        import_db_label = wx.StaticText(self, label="Цільова база даних для імпорту:")
        import_db_sizer.Add(import_db_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)
        
        self.import_db_choice = wx.Choice(self)
        import_db_sizer.Add(self.import_db_choice, 1, wx.EXPAND)        
       
        import_sizer.Add(import_db_sizer, 0, wx.EXPAND | wx.ALL, 5)

        buttons_panel = wx.Panel(self)
        buttons_sizer = wx.BoxSizer(wx.HORIZONTAL)
        buttons_panel.SetSizer(buttons_sizer)

        self.scan_button = wx.Button(buttons_panel, label=" Сканувати папку ")
        self.scan_button.Bind(wx.EVT_BUTTON, self.on_process_documents)
        buttons_sizer.Add(self.scan_button, 1, wx.EXPAND | wx.ALL, 10)

        self.stop_button = wx.Button(buttons_panel, label=" Зупинити ")
        self.stop_button.Bind(wx.EVT_BUTTON, self.on_stop_processing_action)
        buttons_sizer.Add(self.stop_button, 1, wx.EXPAND | wx.ALL, 10)
        self.stop_button.Disable()
  
        import_sizer.Add(buttons_panel, 0, wx.EXPAND | wx.ALL, 5)

        self.status_label = wx.StaticText(self, label="")
        import_sizer.Add(self.status_label, 0, wx.EXPAND | wx.ALL, 5)

        self.output_text = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL, size=(-1, 300))
        import_sizer.Add(self.output_text, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(import_sizer)

    def load_databases_to_choice(self):
        populate_databases_choice(self.import_db_choice)
    
    def on_stop_processing_action(self, event):
        self.stop_processing = True
        wx.MessageBox("Обробка файлів буде зупинена після завершення поточного файлу.", "Зупинка", wx.OK | wx.ICON_INFORMATION)

    def on_process_documents(self, event):
        selected_index = self.import_db_choice.GetSelection()
        if selected_index == wx.NOT_FOUND:
            wx.MessageBox("Будь ласка, виберіть цільову базу даних для імпорту.", "Попередження", wx.OK | wx.ICON_WARNING)
            return

        self.output_text.SetEditable(True)
        self.output_text.Clear()
        self.output_text.SetEditable(False)
        self.status_label.SetLabel("")
        self.stop_processing = False

        dlg = wx.DirDialog(self, "Виберіть папку з документами", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            doc_folder = dlg.GetPath()
            dlg.Destroy()

            if not doc_folder:
                self.status_label.SetLabel("Вибір папки скасовано")
                self.status_label.SetForegroundColour(wx.RED)
                return

            # Отримуємо параметри обраної бази з випадаючого списку
            db_path, db_password = self.import_db_choice.GetClientData(selected_index)

            self.is_scanning_active = True 
            self.scan_button.Disable()
            self.stop_button.Enable()
            threading.Thread(target=self.process_documents_thread, args=(doc_folder, db_path, db_password)).start()

        else:
            dlg.Destroy()
            self.status_label.SetLabel("Вибір папки скасовано")
            self.status_label.SetForegroundColour(wx.RED)

    def _reset_tab2_after_scan(self):
        self.stop_button.Disable()
        self.scan_button.Enable()
        self.is_scanning_active = False 

    def process_documents_thread(self, doc_folder, db_path, db_password):
        # Перевіряємо, чи є пароль. Якщо він порожній, виводимо зрозумілу помилку.
        if not db_password:
            wx.CallAfter(wx.MessageBox, 
                f"Для бази даних '{os.path.basename(db_path)}' не вказано пароль у налаштуваннях!", 
                "Помилка підключення", 
                wx.OK | wx.ICON_ERROR
            )
            wx.CallAfter(self.status_label.SetLabel, "Помилка: відсутній пароль бази даних.")
            wx.CallAfter(self.status_label.SetForegroundColour, wx.RED)
            wx.CallAfter(self._reset_tab2_after_scan)
            return

        conn = connect_to_specific_database(db_path, db_password)
        if not conn:
            wx.CallAfter(self.status_label.SetLabel, "Помилка підключення до обраної БД.")
            wx.CallAfter(self.status_label.SetForegroundColour, wx.RED)
            wx.CallAfter(self._reset_tab2_after_scan)
            return

        try:
            cursor = conn.cursor()
            cursor.execute("PRAGMA cache_size = 64000;")

            total_files = 0
            processed_files = 0
            skipped_files = 0
            new_records = 0

            files_to_process = []
            allowed_extensions = ['.doc', '.docx', '.rtf']
            temp_file_prefixes = ('.', '~$', '#', '~')
            temp_file_suffixes = ('~',)

            for root, _, files in os.walk(doc_folder):
                for file in files:
                    filename_lower = file.lower()
                    ext = os.path.splitext(filename_lower)[1]

                    if filename_lower.startswith(temp_file_prefixes) or filename_lower.endswith(temp_file_suffixes):
                        continue

                    if ext in allowed_extensions:
                        files_to_process.append(os.path.join(root, file))

            total_files = len(files_to_process)

            wx.CallAfter(self.status_label.SetLabel, f"Обробка файлів... (0/{total_files})")
            wx.CallAfter(self.status_label.SetForegroundColour, wx.GREEN)
            wx.CallAfter(self.output_text.Clear)

            conn.isolation_level = None
            cursor.execute("BEGIN;")

            for i, filepath in enumerate(files_to_process):
                if self.stop_processing:
                    wx.CallAfter(self.status_label.SetLabel, "Обробку зупинено користувачем.")
                    wx.CallAfter(self.status_label.SetForegroundColour, wx.RED)
                    break

                filename = os.path.basename(filepath)
                ext = os.path.splitext(filename.lower())[1]

                if ext not in allowed_extensions:
                    skipped_files += 1
                    wx.CallAfter(self.output_text.AppendText, f"Пропущено: '{filename}' (непідтримуване розширення).\n")
                    continue

                content = ""
                extracted_successfully = False

                if ext == '.docx':
                    try:
                        doc = Document(filepath)
                        content = "\n".join([p.text for p in doc.paragraphs])
                        extracted_successfully = True
                    except Exception as e:
                        wx.CallAfter(self.output_text.AppendText, f"Помилка при отриманні тексту з DOCX '{filename}': {e}\n")

                elif ext == '.doc' or ext == '.rtf':
                    try:
                        content = extract_text_libreoffice(filepath)
                        if content and content.strip():
                            extracted_successfully = True
                        else:
                            wx.CallAfter(self.output_text.AppendText, f"LibreOffice не зміг конвертувати або отримав порожній текст з '{filename}'.\n")
                            extracted_successfully = False
                    except Exception as e:
                        wx.CallAfter(self.output_text.AppendText, f"Помилка при конвертації '{filename}' за допомогою LibreOffice: {e}\n")
                        extracted_successfully = False

                if not extracted_successfully or not content.strip():
                    skipped_files += 1
                    continue

                filename = normalize_text(filename)
                content = normalize_text(content)

                doc_year, doc_month, doc_day = get_document_date(filename, os.path.dirname(filepath))
                
                document_number = None
                doc_num_match = document_number_pattern.search(filename)
                if doc_num_match:
                    try:
                        document_number = int(doc_num_match.group(1))
                    except ValueError:
                        document_number = None
                
                created_timestamp = os.path.getctime(filepath)
                created_datetime = datetime.fromtimestamp(created_timestamp)
                created_at_str = created_datetime.strftime('%Y-%m-%d %H:%M:%S')
                text_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
                
                cursor.execute("SELECT content_hash FROM documents WHERE filename = ?", (filename,))
                existing_record = cursor.fetchone()

                was_updated_or_added = False

                if existing_record:
                    existing_hash = existing_record[0]
                    if existing_hash != text_hash:
                        cursor.execute("""
                        UPDATE documents SET 
                            year = ?, month = ?, day = ?, content = ?, 
                            document_number = ?, created_at = ?, content_hash = ?
                        WHERE filename = ?
                        """, (doc_year, doc_month, doc_day, content, document_number, created_at_str, text_hash, filename))
                        
                        was_updated_or_added = True
                        wx.CallAfter(self.output_text.AppendText, f"Оновлено: {filename} (новий вміст)\n")
                    else:
                        wx.CallAfter(self.output_text.AppendText, f"Пропуск: {filename} (вже в базі без змін)\n")
                else:
                    cursor.execute("""
                    INSERT INTO documents (filename, year, month, day, content, document_number, created_at, content_hash)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (filename, doc_year, doc_month, doc_day, content, document_number, created_at_str, text_hash))
                    
                    was_updated_or_added = True
                    wx.CallAfter(self.output_text.AppendText, f"Додано: {filename}\n")

                if was_updated_or_added:
                    new_records += 1
                else:
                    skipped_files += 1

                processed_files += 1
                progress_percent = int((processed_files) / total_files * 100) if total_files > 0 else 0
                wx.CallAfter(self.status_label.SetLabel, f"Обробка файлів... ({processed_files}/{total_files}, {progress_percent}%)")

            conn.commit()
            wx.CallAfter(self.output_text.AppendText, "Всі документи оброблені. Запуск оптимізації FTS...\n")
            wx.CallAfter(self.status_label.SetLabel, "Оптимізація FTS...")

            conn.execute("INSERT INTO documents_fts(documents_fts) VALUES('optimize');")
            conn.commit()
            wx.CallAfter(self.output_text.AppendText, "Оптимізація FTS завершена.\n")

        except sqlite3.Error as e:
            wx.CallAfter(self.output_text.AppendText, f"Помилка бази даних: {e}\n")
            try:
                conn.rollback()
            except sqlite3.Error as rb_e:
                wx.CallAfter(self.output_text.AppendText, f"Помилка при відкаті транзакції: {rb_e}\n")
        except Exception as e:
            wx.CallAfter(self.output_text.AppendText, f"Несподівана помилка в обробці документів: {e}\n")
            try:
                conn.rollback()
            except sqlite3.Error as rb_e:
                wx.CallAfter(self.output_text.AppendText, f"Помилка при відкаті транзакції: {rb_e}\n")
        finally:
            if conn:
                conn.close()

            final_status = f"Обробку завершено. Додано {new_records} нових записів. Пропущено {skipped_files} файлів."
            if self.stop_processing:
                final_status = "Обробку зупинено користувачем."
            wx.CallAfter(self.status_label.SetLabel, final_status)
            wx.CallAfter(self.status_label.SetForegroundColour, wx.NullColour)
            wx.CallAfter(self._reset_tab2_after_scan)

