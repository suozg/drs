# tabs/tab_search.py
import os
import threading
import time
import hashlib
import re
from datetime import datetime
import wx
import wx.adv as wx_adv
from sqlcipher3 import dbapi2 as sqlite3 
import config
from config import password
from database import connect_to_database, check_patch_db
from settings_db import get_databases_list
from concurrent.futures import ThreadPoolExecutor, as_completed

class SearchTab(wx.Panel):
    def __init__(self, parent, main_frame):
        super().__init__(parent)
        self.main_frame = main_frame
        self.documents = {}
        self.matches = []
        self.match_index = -1
        
        # Подія для скасування пошуку
        self.search_cancel_event = threading.Event()
        
        self.setup_search_tab()

    def setup_search_tab(self):
        search_sizer = wx.BoxSizer(wx.VERTICAL)

        query_date_panel = wx.Panel(self)
        query_date_sizer = wx.BoxSizer(wx.HORIZONTAL)
        query_date_panel.SetSizer(query_date_sizer)

        query_date_sizer.Add(wx.StaticText(query_date_panel, label="Запит:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.search_entry = wx.TextCtrl(query_date_panel, size=(150, -1), style=wx.TE_PROCESS_ENTER)
        self.search_entry.SetHint("\"прізв* ім* бать*\"")
        query_date_sizer.Add(self.search_entry, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.search_entry.Bind(wx.EVT_TEXT_ENTER, self.on_search_documents)

        query_date_sizer.Add(wx.StaticText(query_date_panel, label="∨"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.show_all_but = wx.Button(query_date_panel, label=" ∞ ", size=(40, -1))
        self.show_all_but.Bind(wx.EVT_BUTTON, self.show_all_docs) 
        query_date_sizer.Add(self.show_all_but, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        query_date_sizer.Add(wx.StaticText(query_date_panel, label="Період"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.start_date_entry = wx_adv.DatePickerCtrl(query_date_panel, style=wx_adv.DP_DROPDOWN | wx_adv.DP_SHOWCENTURY)
        query_date_sizer.Add(self.start_date_entry, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        query_date_sizer.Add(wx.StaticText(query_date_panel, label="-"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.end_date_entry = wx_adv.DatePickerCtrl(query_date_panel, style=wx_adv.DP_DROPDOWN | wx_adv.DP_SHOWCENTURY)
        query_date_sizer.Add(self.end_date_entry, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.search_button = wx.Button(query_date_panel, label=" Пошук ")
        self.search_button.Bind(wx.EVT_BUTTON, self.on_search_documents)
        query_date_sizer.Add(self.search_button, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.search_progress_bar = wx.Gauge(query_date_panel, range=100, size=(100, -1), style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        query_date_sizer.Add(self.search_progress_bar, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        query_date_sizer.Hide(self.search_progress_bar)

        self.delete_button = wx.Button(query_date_panel, label=" Видалити ")
        self.delete_button.Bind(wx.EVT_BUTTON, self.on_delete_selected_file)
        query_date_sizer.Add(self.delete_button, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.delete_progress_bar = wx.Gauge(query_date_panel, range=100, size=(100, -1), style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        query_date_sizer.Add(self.delete_progress_bar, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        query_date_sizer.Hide(self.delete_progress_bar)

        search_sizer.Add(query_date_panel, 0, wx.EXPAND | wx.ALL, 5)

        self.default_text_count_label = "(введіть запит, натисніть ↲, кнопка [∞] просто покаже усі підряд документи)"
        self.count_label = wx.StaticText(self, label=self.default_text_count_label)
        default_font = self.count_label.GetFont()
        smaller_font_size = default_font.GetPointSize() - 2
        if smaller_font_size < 8:
            smaller_font_size = 8
        smaller_font = wx.Font(smaller_font_size, default_font.GetFamily(),
                default_font.GetStyle(), default_font.GetWeight(),
                default_font.GetUnderlined(), default_font.GetFaceName())
        self.count_label.SetFont(smaller_font)
        search_sizer.Add(self.count_label, 0, wx.LEFT | wx.RIGHT, 10) 

        content_panel = wx.Panel(self)
        content_sizer = wx.BoxSizer(wx.HORIZONTAL)
        content_panel.SetSizer(content_sizer)

        self.search_output_listbox = wx.ListBox(content_panel, style=wx.LB_SINGLE, size=(250, -1))
        self.search_output_listbox.Bind(wx.EVT_LISTBOX, self.on_display_document)
        content_sizer.Add(self.search_output_listbox, 0, wx.EXPAND | wx.ALL, 5)

        text_display_panel = wx.Panel(content_panel)
        text_display_sizer = wx.BoxSizer(wx.VERTICAL)
        text_display_panel.SetSizer(text_display_sizer)

        self.content_text = wx.TextCtrl(text_display_panel, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL, size=(600, -1))
        text_display_sizer.Add(self.content_text, 1, wx.EXPAND | wx.ALL, 5)

        self.view_filename_label = wx.StaticText(text_display_panel, label="")
        text_display_sizer.Add(self.view_filename_label, 0, wx.EXPAND | wx.ALL, 5)

        search_in_text_panel = wx.Panel(text_display_panel)
        search_in_text_sizer = wx.BoxSizer(wx.HORIZONTAL)
        search_in_text_panel.SetSizer(search_in_text_sizer)

        search_in_text_sizer.Add(wx.StaticText(search_in_text_panel, label=" Знайти в тексті: "), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.search_in_text_entry = wx.TextCtrl(search_in_text_panel, size=(250, -1), style=wx.TE_PROCESS_ENTER)
        self.search_in_text_entry.Bind(wx.EVT_TEXT_ENTER, self.on_search_in_text)
        search_in_text_sizer.Add(self.search_in_text_entry, 1, wx.EXPAND | wx.ALL, 5)

        search_in_text_button = wx.Button(search_in_text_panel, label=" Шукати ")
        search_in_text_button.Bind(wx.EVT_BUTTON, self.on_search_in_text)
        search_in_text_sizer.Add(search_in_text_button, 0, wx.ALL, 5)

        prev_button = wx.Button(search_in_text_panel, label=" ⬅ Назад ")
        prev_button.Bind(wx.EVT_BUTTON, self.on_prev_match)
        search_in_text_sizer.Add(prev_button, 0, wx.ALL, 5)

        next_button = wx.Button(search_in_text_panel, label=" Вперед ➡ ")
        next_button.Bind(wx.EVT_BUTTON, self.on_next_match)
        search_in_text_sizer.Add(next_button, 0, wx.ALL, 5)

        text_display_sizer.Add(search_in_text_panel, 0, wx.EXPAND | wx.ALL, 5)
        content_sizer.Add(text_display_panel, 1, wx.EXPAND | wx.ALL, 5)
        search_sizer.Add(content_panel, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(search_sizer)
        query_date_panel.Layout()
    
    def format_date_for_fts3(self, queries: str) -> str:
        queries = re.sub(r'(\d{2})\.(\d{2})\.(\d{2,4})', r'\1 NEAR/0 \2 NEAR/0 \3', queries)
        return queries

    def show_all_docs(self, event):
        wx.CallAfter(self.search_output_listbox.Clear)
        self.count_label.SetLabel(self.default_text_count_label)
        query = "" 
        formatted_query = self.format_date_for_fts3(query)
        self._start_search_thread(formatted_query, query)

    def on_search_documents(self, event):
        self.search_output_listbox.Clear()
        self.count_label.SetLabel(self.default_text_count_label)

        query = self.search_entry.GetValue().strip()
        if not query or query == "*":
            self.search_output_listbox.Append("Введіть запит для пошуку.")
            return

        formatted_query = self.format_date_for_fts3(query)
        self._start_search_thread(formatted_query, query)

    def _start_search_thread(self, formatted_query, query):
        # Оновлюємо стан доступності всіх баз даних перед початком пошуку
        if hasattr(self.main_frame, 'tab_settings'):
            self.main_frame.tab_settings.load_databases_into_ui()

        self.search_cancel_event.clear()

        # Змінюємо саму кнопку "Пошук" на "Стоп"
        self.search_button.SetLabel(" Стоп ")
        self.search_button.Unbind(wx.EVT_BUTTON)
        self.search_button.Bind(wx.EVT_BUTTON, self.on_stop_search)

        sizer = self.search_button.GetContainingSizer()
        sizer.Show(self.search_progress_bar, recursive=True)
        self.search_button.GetParent().Layout()

        stop_pulsing_search = threading.Event()
        pulse_thread_search = threading.Thread(
            target=self._pulse_gauge_loop,
            args=(self.search_progress_bar, stop_pulsing_search)
        )
        pulse_thread_search.start()

        threading.Thread(
            target=self.perform_search,
            args=(formatted_query, query, stop_pulsing_search, pulse_thread_search)
        ).start()

    def on_stop_search(self, event):
        self.search_cancel_event.set()
        self.search_button.SetLabel(" Зупинка... ")
        self.search_button.Enable(False)

    def _search_single_database(self, db_info, formatted_query, original_query, start_date_num, end_date_num, filename_pattern):
        if self.search_cancel_event.is_set():
            return [], None

        db_id, db_name, db_path, db_password, is_active = db_info
        if not is_active:
            return [], None
            
        if not os.path.exists(db_path):
            return [], f"База даних '{db_name}' недоступна (файл не знайдено)."

        from database import connect_to_specific_database
        conn = connect_to_specific_database(db_path, db_password)
        if not conn:
            return [], f"Не вдалося підключитися до бази '{db_name}'."
            
        db_results = []
        try:
            cursor = conn.cursor()
            if original_query == "":
                sql_query = """
                    SELECT filename, content, created_at, year, month, day 
                    FROM documents
                    WHERE year * 10000 + month * 100 + day BETWEEN ? AND ?
                """
                params = (start_date_num, end_date_num)
            else:
                if start_date_num == end_date_num:
                    sql_query = """
                    SELECT filename, content, created_at, year, month, day FROM (
                        SELECT DISTINCT d.filename, d.content, d.created_at, d.year, d.month, d.day
                        FROM documents AS d
                        JOIN documents_fts AS fts ON d.id = fts.docid
                        WHERE fts.content MATCH ?
                        UNION
                        SELECT filename, content, created_at, year, month, day
                        FROM documents
                        WHERE filename LIKE ?
                        AND year * 10000 + month * 100 + day = ?
                    )
                    """
                    params = (formatted_query, filename_pattern, start_date_num)
                else:
                    sql_query = """
                    SELECT filename, content, created_at, year, month, day FROM (
                        SELECT DISTINCT d.filename, d.content, d.created_at, d.year, d.month, d.day
                        FROM documents AS d
                        JOIN documents_fts AS fts ON d.id = fts.docid
                        WHERE fts.content MATCH ?
                        AND (d.year * 10000 + d.month * 100 + d.day BETWEEN ? AND ?)
                        UNION
                        SELECT filename, content, created_at, year, month, day
                        FROM documents
                        WHERE filename LIKE ?
                        AND (year * 10000 + month * 100 + day BETWEEN ? AND ?)
                    )
                    """
                    params = (formatted_query, start_date_num, end_date_num, filename_pattern, start_date_num, end_date_num)

            cursor.execute(sql_query, params)
            rows = cursor.fetchall()
            
            for row in rows:
                if self.search_cancel_event.is_set():
                    break
                display_name = f"[{db_name}] {row[0]}"
                db_results.append((display_name, row[1], row[2], row[3], row[4], row[5]))
                
        except Exception as e:
            print(f"Помилка пошуку в БД {db_name}: {e}")
        finally:
            conn.close()
            
        return db_results, None

    def perform_search(self, formatted_query, original_query, stop_pulsing_event, pulse_thread_ref):
        start_date_wx = self.start_date_entry.GetValue()
        end_date_wx = self.end_date_entry.GetValue()

        start_date_str = start_date_wx.FormatISODate()
        end_date_str = end_date_wx.FormatISODate()

        if not start_date_str or not end_date_str:
            wx.CallAfter(self.search_output_listbox.Clear)
            wx.CallAfter(self.search_output_listbox.Append, "Виберіть обидві дати.")
            self._finish_search_ui(stop_pulsing_event, pulse_thread_ref)
            return

        start_date_num = int(start_date_str.replace("-", ""))
        end_date_num = int(end_date_str.replace("-", ""))

        if end_date_num < start_date_num:
            wx.CallAfter(self.search_output_listbox.Clear)
            wx.CallAfter(self.search_output_listbox.Append, "Кінцева дата не може бути раніше початкової.")
            self._finish_search_ui(stop_pulsing_event, pulse_thread_ref)
            return

        if original_query == "":
            start_year = start_date_num // 10000
            end_year = end_date_num // 10000
            if start_date_num == end_date_num or ((end_year - start_year) > 10):
                end_date_str = str(end_date_num)
                end_date = datetime.strptime(end_date_str, "%Y%m%d")
                ten_years_ago = end_date.replace(year=end_date.year - 10)
                start_date_num = int(ten_years_ago.strftime("%Y%m%d"))

                wx.CallAfter(self.search_output_listbox.Clear)
                wx.CallAfter(self.search_output_listbox.Append, "Період обмежено 10 роками")
                wx.CallAfter(self.start_date_entry.SetValue, wx.DateTime.FromDMY(
                    ten_years_ago.day, ten_years_ago.month - 1, ten_years_ago.year
                ))

        databases = get_databases_list(config.master_password) if hasattr(config, 'master_password') else []
        all_results = []
        warnings = []
        filename_pattern = f"%{original_query}%"

        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {
                executor.submit(
                    self._search_single_database, 
                    db, formatted_query, original_query, start_date_num, end_date_num, filename_pattern
                ): db for db in databases
            }
            
            for future in as_completed(futures):
                if self.search_cancel_event.is_set():
                    continue
                try:
                    db_results, warning_msg = future.result()
                    if db_results:
                        all_results.extend(db_results)
                    if warning_msg:
                        warnings.append(warning_msg)
                except Exception as e:
                    print(f"Помилка в потоці виконання: {e}")

        if self.search_cancel_event.is_set():
            wx.CallAfter(self.search_output_listbox.Clear)
            wx.CallAfter(self.search_output_listbox.Append, "Пошук зупинено користувачем.")
            self._finish_search_ui(stop_pulsing_event, pulse_thread_ref)
            return

        if warnings:
            for w in warnings:
                all_results.append((w, "", "", 0, 0, 0))

        all_results.sort(key=lambda x: (x[3] if x[3] else 0, x[4] if x[4] else 0, x[5] if x[5] else 0), reverse=True)

        ui_results = [(res[0], res[1], res[2]) for res in all_results]

        wx.CallAfter(self.update_search_results_ui, ui_results, original_query)
        self._finish_search_ui(stop_pulsing_event, pulse_thread_ref)

    def _finish_search_ui(self, stop_pulsing_event, pulse_thread_ref):
        wx.CallAfter(stop_pulsing_event.set)
        wx.CallAfter(pulse_thread_ref.join)
        
        def restore_button():
            self.search_button.SetLabel(" Пошук ")
            self.search_button.Enable(True)
            self.search_button.Unbind(wx.EVT_BUTTON)
            self.search_button.Bind(wx.EVT_BUTTON, self.on_search_documents)
            
            sizer = self.search_button.GetContainingSizer()
            sizer.Hide(self.search_progress_bar, recursive=True)
            self.search_button.GetParent().Layout()

        wx.CallAfter(restore_button)

    def update_search_results_ui(self, results, original_query):
        self.search_output_listbox.Clear()
        self.documents.clear()
        self.content_text.Clear()
        self.view_filename_label.SetLabel("")
        self.search_in_text_entry.SetValue("")
        self.content_text.SetEditable(True)
        self.content_text.SetStyle(0, self.content_text.GetLastPosition(), wx.TextAttr(wx.NullColour, wx.NullColour, wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)))
        self.content_text.SetEditable(False)

        if results:
            for index, (filename, content, created_at) in enumerate(results):
                doc_key = filename
                self.documents[doc_key] = (content, created_at)
                self.search_output_listbox.Append(doc_key)
        else:
            if not self.search_cancel_event.is_set():
                self.search_output_listbox.Append("Нічого не знайдено.")

        self.update_count_label(self.search_output_listbox.GetCount())

        if original_query:
            cleaned_query = re.sub(r'["\'`‘’“”*]', '', original_query)
            first_word = cleaned_query.split()[0] if cleaned_query else ""
            self.search_in_text_entry.SetValue(first_word)

    def on_delete_selected_file(self, event):
        if not check_patch_db():
            return

        if hasattr(self.main_frame, 'notebook') and self.main_frame.notebook:
            self.main_frame.notebook.Enable(False)

        selected_index = self.search_output_listbox.GetSelection()
        if selected_index == wx.NOT_FOUND:
            wx.MessageBox("Файл не вибрано.", "Попередження", wx.OK | wx.ICON_WARNING)
            if hasattr(self.main_frame, 'notebook') and self.main_frame.notebook:
                self.main_frame.notebook.Enable(True)
            return

        filename = self.search_output_listbox.GetString(selected_index)

        confirm = wx.MessageBox(f"Ви дійсно хочете видалити {filename}?", "Підтвердження", wx.YES_NO | wx.ICON_QUESTION)
        if confirm == wx.NO:
            if hasattr(self.main_frame, 'notebook') and self.main_frame.notebook:
                self.main_frame.notebook.Enable(True)
            return

        self.delete_button.GetContainingSizer().Hide(self.delete_button, recursive=True)
        self.delete_button.GetContainingSizer().Show(self.delete_progress_bar, recursive=True)
        self.delete_button.GetParent().Layout()

        thread = threading.Thread(target=self._perform_delete_file, args=(filename, selected_index))
        thread.start()


    def _perform_delete_file(self, filename, selected_index):
        conn = None
        stop_pulsing = threading.Event()
        pulse_thread = threading.Thread(target=self._pulse_gauge_loop, args=(self.delete_progress_bar, stop_pulsing))
        pulse_thread.start()
        
        try:
            db_name_match = re.match(r'^\[(.*?)\]\s+(.*)$', filename)
            if not db_name_match:
                wx.CallAfter(lambda: wx.MessageBox("Не вдалося визначити базу даних для цього файлу.", "Помилка", wx.OK | wx.ICON_ERROR))
                return
                
            target_db_name = db_name_match.group(1)
            clean_filename = db_name_match.group(2)

            databases = get_databases_list(config.master_password) if hasattr(config, 'master_password') else []
            target_db_path = None
            target_db_password = None

            for db_id, db_name, db_path, db_password, is_active in databases:
                if db_name == target_db_name:
                    target_db_path = db_path
                    target_db_password = db_password
                    break

            if not target_db_path:
                wx.CallAfter(lambda: wx.MessageBox(f"Не знайдено шлях до бази даних '{target_db_name}'.", "Помилка", wx.OK | wx.ICON_ERROR))
                return

            from database import connect_to_specific_database
            conn = connect_to_specific_database(target_db_path, target_db_password)

            if not conn:
                wx.CallAfter(lambda: wx.MessageBox(f"Не вдалося підключитися до бази даних '{target_db_name}'.", "Помилка", wx.OK | wx.ICON_ERROR))
                return

            cursor = conn.cursor()
            cursor.execute("DELETE FROM documents WHERE filename = ?", (clean_filename,))
            conn.commit()

            if cursor.rowcount == 0:
                wx.CallAfter(lambda: wx.MessageBox(f"Файл {clean_filename} не знайдено в базі.", "Увага!", wx.OK | wx.ICON_WARNING))
            else:
                wx.CallAfter(self._update_gui_after_delete, filename, selected_index)
                wx.CallAfter(lambda: wx.MessageBox(f"Файл {clean_filename} видалено з бази.", "Увага!", wx.OK | wx.ICON_INFORMATION))

        except sqlite3.DatabaseError as e:
            wx.CallAfter(lambda: wx.MessageBox(f"Помилка бази даних: {e}", "Помилка видалення", wx.OK | wx.ICON_ERROR))
        finally:
            stop_pulsing.set()
            pulse_thread.join()

            if conn:
                conn.close()
            wx.CallAfter(self.delete_button.GetContainingSizer().Hide, self.delete_progress_bar, recursive=True)
            wx.CallAfter(self.delete_button.GetContainingSizer().Show, self.delete_button, recursive=True)
            wx.CallAfter(self.delete_button.GetParent().Layout)
            
            if hasattr(self.main_frame, 'notebook') and self.main_frame.notebook:
                wx.CallAfter(self.main_frame.notebook.Enable, True)


    def _pulse_gauge_loop(self, gauge, stop_event):
        while not stop_event.is_set():
            wx.CallAfter(gauge.Pulse)
            time.sleep(0.1)

    def _update_gui_after_delete(self, filename, selected_index):
        self.search_output_listbox.Delete(selected_index)
        if filename in self.documents:
            del self.documents[filename]
        self.content_text.Clear()
        self.view_filename_label.SetLabel("")
        self.search_in_text_entry.SetValue("")
        self.update_count_label(self.search_output_listbox.GetCount())

    def update_count_label(self, count):
        if count > 0:
            self.count_label.SetLabel(f"Знайдено записів: {count}")
            self.count_label.Show()
        else:
            self.count_label.SetLabel("")
            self.count_label.Hide()
        self.Layout()

    def on_display_document(self, event):
        selected_index = self.search_output_listbox.GetSelection()
        if selected_index != wx.NOT_FOUND:
            filename = self.search_output_listbox.GetString(selected_index)
            content, created_at = self.documents.get(filename, ("", ""))

            self.content_text.SetEditable(True)
            self.content_text.Clear()
            self.content_text.WriteText(content)
            self.content_text.SetEditable(False)
            self.content_text.SetDefaultStyle(wx.TextAttr())

            self.view_filename_label.SetLabel(f"{filename} (дані додано: {created_at})")
            self.on_search_in_text(None)

    def on_search_in_text(self, event):
        query = self.search_in_text_entry.GetValue()
        if not query:
            self.content_text.SetStyle(0, self.content_text.GetLastPosition(), wx.TextAttr(wx.NullColour))
            self.matches = []
            current_filename_label = self.view_filename_label.GetLabel().split(" - [ збігів:")[0]
            self.view_filename_label.SetLabel(current_filename_label)
            return

        self.matches = []
        full_text = self.content_text.GetValue()
        query_lower = query.lower()
        text_lower = full_text.lower()

        start_pos = 0
        while True:
            idx = text_lower.find(query_lower, start_pos)
            if idx == -1:
                break
            self.matches.append(idx)
            start_pos = idx + len(query)

        self.content_text.SetStyle(0, self.content_text.GetLastPosition(), wx.TextAttr(wx.NullColour, wx.NullColour))

        if self.matches:
            self.match_index = 0
            self.go_to_match(self.match_index)
        else:
            self.match_index = -1

        current_filename_label = self.view_filename_label.GetLabel().split(" (дані додано:")[0]
        self.view_filename_label.SetLabel(f"{current_filename_label} - [ збігів: {len(self.matches)} ]")

    def go_to_match(self, index):
        if not self.matches:
            return

        self.match_index = index % len(self.matches)
        start_idx = self.matches[self.match_index]
        end_idx = start_idx + len(self.search_in_text_entry.GetValue())

        self.content_text.SetStyle(0, self.content_text.GetLastPosition(), wx.TextAttr(wx.NullColour, wx.NullColour))
        light_blue_bg_color = wx.Colour(173, 216, 230)

        for match_start in self.matches:
            match_end = match_start + len(self.search_in_text_entry.GetValue())
            self.content_text.SetStyle(match_start, match_end, wx.TextAttr(wx.NullColour, light_blue_bg_color))

        self.content_text.SetStyle(start_idx, end_idx, wx.TextAttr(wx.NullColour, wx.YELLOW))
        self.content_text.ShowPosition(start_idx)

    def on_next_match(self, event):
        if self.matches:
            self.go_to_match(self.match_index + 1)

    def on_prev_match(self, event):
        if self.matches:
            self.go_to_match(self.match_index - 1)
