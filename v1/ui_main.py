import os
import threading
import time
import hashlib
import re
from datetime import datetime
import wx
import wx.adv as wx_adv
from docx import Document
from sqlcipher3 import dbapi2 as sqlite3 
import config
from config import password, frame_title, db_path, document_number_pattern
from utils import get_document_date, extract_text_libreoffice, get_full_db_path, _get_full_db_path
from database import connect_to_database, check_patch_db
from ui_dialogs import PasswordDialog
from sql_console import SqlConsolePanel

class DocumentSearchFrame(wx.Frame):
    def __init__(self, parent, title):
        super(DocumentSearchFrame, self).__init__(parent, title=title, size=(800, 640))

        self.db_info = {
            "records": 0,
            "last_modified": "Недоступно",
            "last_update_app_time": "Не оновлювалось"
        }
        self.stop_processing = False
        self.is_scanning_active = False
        self.documents = {}
        self.matches = []
        self.match_index = -1

        self.InitUI()
        self.Centre()
        self.Show()

        # Запит пароля при запуску
        self.prompt_for_password()
    
    def InitUI(self):
        panel = wx.Panel(self)
        
        # 1. Головний сайзер (вертикальний)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 2. Створюємо Notebook (він тепер буде першим, щоб займати основне місце)
        self.notebook = wx.Notebook(panel)
        main_sizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 5)
        
        # 3. Створюємо компактний "footer" для назви бази (внизу)
        footer_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.file_label = wx.StaticText(panel, label="Вибрано файл бази: " + config.db_path)
        # Робимо шрифт меншим для економії місця
        font = self.file_label.GetFont()
        font.SetPointSize(8) 
        self.file_label.SetFont(font)
        
        # Додаємо з мінімальними відступами
        footer_sizer.Add(self.file_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.BOTTOM, 2)
        
        # Додаємо footer_sizer до головного сайзера (він буде під Notebook)
        main_sizer.Add(footer_sizer, 0, wx.EXPAND | wx.LEFT, 5)
        
        panel.SetSizer(main_sizer)

        # --- Вкладки ---
        self.tab1 = wx.Panel(self.notebook)
        self.tab2 = wx.Panel(self.notebook)
        self.tab3 = wx.Panel(self.notebook)
        self.tab4 = wx.Panel(self.notebook)
        self.tabH = wx.Panel(self.notebook)  #
        
        self.notebook.AddPage(self.tab1, "Пошук")
        self.notebook.AddPage(self.tab2, "Імпорт")
        self.notebook.AddPage(self.tab3, "Пароль")
        self.notebook.AddPage(self.tab4, "SQL") 
        self.notebook.AddPage(self.tabH, "Про")

        # --- Далі ваш код налаштування ---
        self.setup_search_tab()
        self.setup_import_tab()
        self.setup_password_tab()
        self.help_tab()        

        self.ShowHowTo() 

        self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_page_changed)
        self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGING, self.on_notebook_page_changing)

    def on_page_changed(self, event):
        password_tab_index = 2 # 'Пошук' (0), 'Імпорт' (1), 'Пароль' (2)

        if event.GetSelection() == password_tab_index:
            # Якщо вибрано вкладку "Пароль", оновлюємо інформацію про БД в окремому потоці
            threading.Thread(target=self.update_db_info).start()
        event.Skip() # Важливо пропустити подію, щоб вона продовжувала оброблятися за замовчуванням

    def on_notebook_page_changing(self, event):
        """
        Обробник подій для запобігання зміні вкладки
        """
        # Перевіряємо прапор, який вказує, чи активне сканування
        if hasattr(self, 'is_scanning_active') and self.is_scanning_active:
            # Якщо сканування активне, забороняємо перехід
            wx.MessageBox("Будь ласка, дочекайтеся завершення або зупиніть процес обробки документів, перш ніж переходити на іншу вкладку.", 
                          "Обробка активна", wx.OK | wx.ICON_INFORMATION)
            event.Veto() # Скасувати зміну вкладки
        else:
            event.Skip() # Дозволити зміну вкладки, якщо сканування не активне
    
    def prompt_for_password(self):
        global password
        while True:
            dlg = PasswordDialog(self, "Введіть пароль", "Пароль")
            result = dlg.ShowModal()
            
            if result == wx.ID_OK:
                input_password = dlg.GetValue().strip()
                if not input_password:
                    wx.MessageBox("Пароль не може бути порожнім.", "Помилка", wx.OK | wx.ICON_ERROR)
                    dlg.Destroy()
                    continue  # Продовжуємо цикл, викликаємо діалог знову
                
                conn = connect_to_database(input_password)
                if conn:
                    config.password = input_password
                    
                    # --- ІНІЦІАЛІЗАЦІЯ вкладки SQL --
                    self.sql_console = SqlConsolePanel(self.tab4, fut_place=None)
                    sql_sizer = wx.BoxSizer(wx.VERTICAL)
                    sql_sizer.Add(self.sql_console, 1, wx.EXPAND | wx.ALL, 0)
                    self.tab4.SetSizer(sql_sizer)
                    self.tab4.Layout()
                    # -----------------------------------------------

                    conn.close()
                    dlg.Destroy()
                    self.file_label.SetLabel("Вибрано файл бази: " + config.db_path)
                    # Запускаємо оновлення інформації про БД в окремому потоці
                    threading.Thread(target=self.update_db_info).start()
                    break  # Успішно! Виходимо з циклу
                else:
                    wx.MessageBox("Невірний пароль або помилка бази даних.", "Помилка підключення", wx.OK | wx.ICON_ERROR)
                    dlg.Destroy()
                    # Цикл піде на наступну ітерацію і знову покаже діалог
            else:
                # Якщо натиснули "Скасувати" або закрили вікно — тоді закриваємо програму
                dlg.Destroy()
                self.Close()
                return

    # --- Налаштування вкладки 1 ---

    def setup_search_tab(self):
        search_sizer = wx.BoxSizer(wx.VERTICAL)

        query_date_panel = wx.Panel(self.tab1)
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
        self.count_label = wx.StaticText(self.tab1, label=self.default_text_count_label)
        default_font = self.count_label.GetFont()
        # Створюємо новий шрифт, зменшуючи його розмір на 2-4 пункти
        smaller_font_size = default_font.GetPointSize() - 2 # Зменшуємо на 2 пункти
        if smaller_font_size < 8: # Мінімальний розумний розмір шрифту
            smaller_font_size = 8
        smaller_font = wx.Font(smaller_font_size, default_font.GetFamily(),
                default_font.GetStyle(), default_font.GetWeight(),
                default_font.GetUnderlined(), default_font.GetFaceName())
        self.count_label.SetFont(smaller_font)
        search_sizer.Add(self.count_label, 0, wx.LEFT | wx.RIGHT, 10) 

        content_panel = wx.Panel(self.tab1)
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

        self.tab1.SetSizer(search_sizer)
        query_date_panel.Layout()


    def ShowHowTo(self):
        """Відображає інструкції щодо синтаксису пошуку."""
        help_search = (
            "1. Пошук одного слова\n"
            "\tслово\n"
            "Знайде всі записи, що містять вказане слово.\n\n"
            "2. Пошук кількох слів (AND)\n"
            "\tслово1 слово2\n"
            "Знайде записи, що містять обидва слова (слово1 та слово2).\n\n"
            "3. Пошук з OR\n"
            "\tслово1 OR слово2\n"
            "Знайде записи, які містять хоча б одне зі слів.\n\n"
            "4. Пошук фрази (використовуємо лапки)\n"
            "\t'точна фраза'\n"
            "Знайде точний збіг заданої фрази, зберігаючи порядок слів.\n\n"
            "5. Виключення слів (NOT, -)\n"
            "\tслово1 -слово2\n"
            "Знайде записи, які містять слово1, але не містять слово2.\n\n"
            "6. Пошук за близькістю слів (NEAR)\n"
            "\tслово1 NEAR/5 слово2\n"
            "Знайде записи, в яких слово1 та слово2 знаходяться не далі ніж за 5 слів одне від одного.\n\n"
            "7. Пошук за префіксом (частина слова)\n"
            "\tсло*\n"
            "Знайде всі слова, які починаються зі сло (наприклад, слово, слон)."
        )

        self.content_text.SetEditable(True)
        self.content_text.Clear()

        for line in help_search.split("\n"):
            if line.strip().isdigit() or line.endswith("пошуку") or "." in line[:3]:
                self.content_text.SetDefaultStyle(wx.TextAttr(wx.NullColour, wx.NullColour, wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)))
                self.content_text.AppendText(line + "\n")
            elif line.startswith("\t"):
                self.content_text.SetDefaultStyle(wx.TextAttr(wx.NullColour, wx.NullColour, wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)))
                self.content_text.AppendText(line.strip() + "\n")
            else:
                self.content_text.SetDefaultStyle(wx.TextAttr(wx.NullColour, wx.NullColour, wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)))
                self.content_text.AppendText(line + "\n")
        self.content_text.SetEditable(False)
        self.content_text.SetDefaultStyle(wx.TextAttr())

    def format_date_for_fts3(self, queries: str) -> str:
        queries = re.sub(r'(\d{2})\.(\d{2})\.(\d{2,4})', r'\1 NEAR/0 \2 NEAR/0 \3', queries)
        return queries

    def show_all_docs(self, event):
        wx.CallAfter(self.search_output_listbox.Clear)
        self.count_label.SetLabel(self.default_text_count_label)
        query = ""  # Пустой запрос для отключения текстового поиска

        formatted_query = self.format_date_for_fts3(query)

        self.search_button.GetContainingSizer().Hide(self.search_button, recursive=True)
        self.search_button.GetContainingSizer().Show(self.search_progress_bar, recursive=True)
        self.search_button.GetParent().Layout()

        stop_pulsing_search = threading.Event()
        pulse_thread_search = threading.Thread(
            target=self._pulse_gauge_loop,
            args=(self.search_progress_bar, stop_pulsing_search)
        )
        pulse_thread_search.start()

        threading.Thread(
            target=self.perform_search,
            args=(password, formatted_query, query, stop_pulsing_search, pulse_thread_search)
        ).start()

    def on_search_documents(self, event):
        self.search_output_listbox.Clear()
        self.count_label.SetLabel(self.default_text_count_label)

        query = self.search_entry.GetValue().strip()
        #if not query or query == "*" or len(query) < 4:
        if not query or query == "*":
            self.search_output_listbox.Append("Введіть запит для пошуку.")
            return

        formatted_query = self.format_date_for_fts3(query)

        self.search_button.GetContainingSizer().Hide(self.search_button, recursive=True)
        self.search_button.GetContainingSizer().Show(self.search_progress_bar, recursive=True)
        self.search_button.GetParent().Layout()

        stop_pulsing_search = threading.Event()
        pulse_thread_search = threading.Thread(target=self._pulse_gauge_loop, args=(self.search_progress_bar, stop_pulsing_search))
        pulse_thread_search.start()

        threading.Thread(target=self.perform_search, args=(password, formatted_query, query, stop_pulsing_search, pulse_thread_search)).start()

    def perform_search(self, password, formatted_query, original_query, stop_pulsing_event, pulse_thread_ref):
        if not check_patch_db():
            wx.CallAfter(stop_pulsing_event.set)
            wx.CallAfter(pulse_thread_ref.join)
            wx.CallAfter(self.search_button.GetContainingSizer().Hide, self.search_progress_bar, recursive=True)
            wx.CallAfter(self.search_button.GetContainingSizer().Show, self.search_button, recursive=True)
            wx.CallAfter(self.search_button.GetParent().Layout)
            return

        start_date_wx = self.start_date_entry.GetValue()
        end_date_wx = self.end_date_entry.GetValue()

        start_date_str = start_date_wx.FormatISODate()
        end_date_str = end_date_wx.FormatISODate()

        if not start_date_str or not end_date_str:
            wx.CallAfter(self.search_output_listbox.Clear)
            wx.CallAfter(self.search_output_listbox.Append, "Виберіть обидві дати.")
            wx.CallAfter(stop_pulsing_event.set)
            wx.CallAfter(pulse_thread_ref.join)
            wx.CallAfter(self.search_button.GetContainingSizer().Hide, self.search_progress_bar, recursive=True)
            wx.CallAfter(self.search_button.GetContainingSizer().Show, self.search_button, recursive=True)
            wx.CallAfter(self.search_button.GetParent().Layout)
            return

        start_date_num = int(start_date_str.replace("-", ""))
        end_date_num = int(end_date_str.replace("-", ""))

        if end_date_num < start_date_num:
            wx.CallAfter(self.search_output_listbox.Clear)
            wx.CallAfter(self.search_output_listbox.Append, "Кінцева дата не може бути раніше початкової.")
            wx.CallAfter(stop_pulsing_event.set)
            wx.CallAfter(pulse_thread_ref.join)
            wx.CallAfter(self.search_button.GetContainingSizer().Hide, self.search_progress_bar, recursive=True)
            wx.CallAfter(self.search_button.GetContainingSizer().Show, self.search_button, recursive=True)
            wx.CallAfter(self.search_button.GetParent().Layout)
            return

        # Если нет запроса и выбран весь период больше 10 лет => слишком много данных
        if original_query == "":
            start_year = start_date_num // 10000
            end_year = end_date_num // 10000
            if start_date_num == end_date_num or ((end_year - start_year) > 10):
                # Автоматически ограничиваем стартовую дату
                end_date_str = str(end_date_num)
                end_date = datetime.strptime(end_date_str, "%Y%m%d")
                ten_years_ago = end_date.replace(year=end_date.year - 10)
                start_date_num = int(ten_years_ago.strftime("%Y%m%d"))

                wx.CallAfter(self.search_output_listbox.Clear)
                wx.CallAfter(self.search_output_listbox.Append, "Період обмежено 10 роками")
                self.start_date_entry.SetValue(wx.DateTime.FromDMY(
                    ten_years_ago.day, ten_years_ago.month - 1, ten_years_ago.year
                ))

 
        conn = None
        results = []
        try:
            conn = connect_to_database(config.password)
            if not conn:
                wx.CallAfter(stop_pulsing_event.set)
                wx.CallAfter(pulse_thread_ref.join)
                wx.CallAfter(self.search_button.GetContainingSizer().Hide, self.search_progress_bar, recursive=True)
                wx.CallAfter(self.search_button.GetContainingSizer().Show, self.search_button, recursive=True)
                wx.CallAfter(self.search_button.GetParent().Layout)
                return

            cursor = conn.cursor()

            # Формуємо шаблон для LIKE пошуку по filename
            filename_pattern = f"%{original_query}%"

            if original_query == "":
                # Без пошуку (показати все за період)
                sql_query = """
                    SELECT filename, content, created_at
                    FROM documents
                    WHERE year * 10000 + month * 100 + day BETWEEN ? AND ?
                    ORDER BY year DESC, month DESC, day DESC
                """
                params = (start_date_num, end_date_num)
            else:
                if start_date_num == end_date_num:
                    sql_query = """
                    SELECT filename, content, created_at FROM (
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
                    ORDER BY year DESC, month DESC, day DESC
                    """
                    params = (formatted_query, filename_pattern, start_date_num)
                else:                                    
                    sql_query = """
                    SELECT filename, content, created_at FROM (
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
                    ORDER BY year DESC, month DESC, day DESC
                    """
                    params = (formatted_query, start_date_num, end_date_num, filename_pattern, start_date_num, end_date_num)

            cursor.execute(sql_query, params)
            results = cursor.fetchall()

        except sqlite3.DatabaseError as e:
            wx.CallAfter(wx.MessageBox, f"Помилка бази даних: {e}", "Помилка пошуку", wx.OK | wx.ICON_ERROR)
        finally:
            wx.CallAfter(stop_pulsing_event.set)
            wx.CallAfter(pulse_thread_ref.join)

            if conn:
                conn.close()

            wx.CallAfter(self.update_search_results_ui, results, original_query)
            wx.CallAfter(self.search_button.GetContainingSizer().Hide, self.search_progress_bar, recursive=True)
            wx.CallAfter(self.search_button.GetContainingSizer().Show, self.search_button, recursive=True)
            wx.CallAfter(self.search_button.GetParent().Layout)

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
            self.ShowHowTo()
            for index, (filename, content, created_at) in enumerate(results):
                doc_key = filename
                self.documents[doc_key] = (content, created_at)
                self.search_output_listbox.Append(doc_key)
                if index % 2 != 0:
                    pass
        else:
            self.search_output_listbox.Append("Нічого не знайдено.")
            self.ShowHowTo()

        self.update_count_label(self.search_output_listbox.GetCount())

        cleaned_query = re.sub(r'["\'`‘’“”*]', '', original_query)
        first_word = cleaned_query.split()[0] if cleaned_query else ""
        self.search_in_text_entry.SetValue(first_word)

    def on_delete_selected_file(self, event):
        if not check_patch_db():
            return

        self.notebook.Enable(False)

        selected_index = self.search_output_listbox.GetSelection()
        if selected_index == wx.NOT_FOUND:
            wx.MessageBox("Файл не вибрано.", "Попередження", wx.OK | wx.ICON_WARNING)
            return

        filename = self.search_output_listbox.GetString(selected_index)

        confirm = wx.MessageBox(f"Ви дійсно хочете видалити {filename}?", "Підтвердження", wx.YES_NO | wx.ICON_QUESTION)
        if confirm == wx.NO:
            return

        self.delete_button.GetContainingSizer().Hide(self.delete_button, recursive=True)
        self.delete_button.GetContainingSizer().Show(self.delete_progress_bar, recursive=True)
        self.delete_button.GetParent().Layout()

        thread = threading.Thread(target=self._perform_delete_file, args=(filename, selected_index, password))
        thread.start()

    def _perform_delete_file(self, filename, selected_index, password):
        conn = None
        try:
            conn = connect_to_database(config.password)

            if not conn:
                wx.CallAfter(lambda: wx.MessageBox("Не вдалося підключитися до бази даних.", "Помилка", wx.OK | wx.ICON_ERROR))
                return

            cursor = conn.cursor()
            stop_pulsing = threading.Event()
            pulse_thread = threading.Thread(target=self._pulse_gauge_loop, args=(self.delete_progress_bar, stop_pulsing))
            pulse_thread.start()
            cursor.execute("DELETE FROM documents WHERE filename = ?", (filename,))
            conn.commit()

            if cursor.rowcount == 0:
                wx.CallAfter(lambda: wx.MessageBox(f"Файл {filename} не знайдено в базі.", "Увага!", wx.OK | wx.ICON_WARNING))
            else:
                wx.CallAfter(self._update_gui_after_delete, filename, selected_index)
                wx.CallAfter(lambda: wx.MessageBox(f"Файл {filename} видалено з бази.", "Увага!", wx.OK | wx.ICON_INFORMATION))

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
            # Оновлюємо інформацію про БД після видалення в окремому потоці
            threading.Thread(target=self.update_db_info).start()
            self.notebook.Enable(True)

    def _pulse_gauge_loop(self, gauge, stop_event):
        """
        Допоміжна функція для циклічної пульсації wx.Gauge у фоновому потоці.
        """
        while not stop_event.is_set():
            wx.CallAfter(gauge.Pulse)
            time.sleep(0.1)

    def _update_gui_after_delete(self, filename, selected_index):
        """
        Оновлює елементи інтерфейсу після успішного видалення.
        Викликається з головного потоку через wx.CallAfter.
        """
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
        self.tab1.Layout()

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

    # --- Налаштування вкладки 2 ---
    def setup_import_tab(self):
        import_sizer = wx.BoxSizer(wx.VERTICAL)

        info_panel = wx.Panel(self.tab2)
        info_sizer = wx.BoxSizer(wx.VERTICAL)
        info_panel.SetSizer(info_sizer)

        info_sizer.Add(wx.StaticText(info_panel, 
                label="УВАГА! Дозволені тільки документи .doc, .docx, .rtf. Перевірте каталог на відсутність інших форматів.",
                style=wx.ALIGN_LEFT,), 0, wx.ALL, 5)

        import_sizer.Add(info_panel, 0, wx.EXPAND | wx.ALL, 10)

        buttons_panel = wx.Panel(self.tab2)
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

        self.status_label = wx.StaticText(self.tab2, label="")
        import_sizer.Add(self.status_label, 0, wx.EXPAND | wx.ALL, 5)

        self.output_text = wx.TextCtrl(self.tab2, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL, size=(-1, 300))
        import_sizer.Add(self.output_text, 1, wx.EXPAND | wx.ALL, 5)

        self.tab2.SetSizer(import_sizer)

    # --- Налаштування вкладки 3 ---
    def setup_password_tab(self):
        password_sizer = wx.BoxSizer(wx.VERTICAL)

        change_pass_panel = wx.Panel(self.tab3)
        change_pass_sizer = wx.FlexGridSizer(rows=3, cols=2, vgap=10, hgap=10)
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
        self.change_pass_button = wx.Button(self.tab3, label=" Змінити пароль ")
        self.change_pass_button.Bind(wx.EVT_BUTTON, self.on_change_password)
        
        button_container_sizer.Add(self.change_pass_button, 1, wx.EXPAND | wx.ALL, 5)

        top_section_sizer = wx.BoxSizer(wx.HORIZONTAL)
        top_section_sizer.Add(change_pass_panel, 1, wx.EXPAND | wx.ALL, 10)
        top_section_sizer.Add(button_container_sizer, 0, wx.ALL | wx.EXPAND, 10)

        password_sizer.Add(top_section_sizer, 0, wx.EXPAND | wx.ALL, 10)

        self.password_progress_text = wx.TextCtrl(self.tab3, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL, size=(-1, 200))
        password_sizer.Add(self.password_progress_text, 1, wx.EXPAND | wx.ALL, 10)

        self.tab3.SetSizer(password_sizer)
        self.tab3.Layout()

    # --- Налаштування вкладки HELP ---
    
    def help_tab(self):
        full_path, full_absolute_path = _get_full_db_path(db_path)
        
        long_help_text = (
            f"\n{frame_title}\n\n(c) 2025. Холодов О.В. Ліцензія GPL v.2\n\n"
            f"\n{frame_title} оброблює текстові файли (щоденні звіти, накази) у визначеному каталозі "
            "та виконує пошук текстів за параметрами. \n"
            "\nОчікується дата (ДД.ММ.РРРР) в назві файла, а також групування файлів в каталогах по-місячно та по роках.\n"
            "Наприклад, каталог /2025/05 (рік/місяць) з файлами 'НАКАЗ №1 від 01.05.2025.docx', '№2 від 09.05.2025.docx'\n\n"
            f"За замовчуванням файл бази даних ({db_path}) буде створений автоматично "
            "під час першого запуску, або при його відсутності в каталозі програми.\n"
            "\nШлях до бази даних можна також вказати при запуску з командного рядка з параметром -с\n"
            f"Наприклад: {full_absolute_path} -c {full_path} \n"
        )

        help_sizer = wx.BoxSizer(wx.VERTICAL)        
        help_label = wx.StaticText(self.tabH, label=long_help_text, style=wx.ALIGN_LEFT)

        # 2. Розміщення по центру в сайзері
        help_sizer.Add(help_label, 1, wx.ALL | wx.EXPAND, 10)

        self.tabH.SetSizer(help_sizer)
        self.tabH.Layout()

    # --- Обробники подій вкладки "Імпорт" ---

    def on_stop_processing_action(self, event):
        self.stop_processing = True
        wx.MessageBox("Обробка файлів буде зупинена після завершення поточного файлу.", "Зупинка", wx.OK | wx.ICON_INFORMATION)

    def on_process_documents(self, event):
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

            self.is_scanning_active = True 
            self.scan_button.Disable() # Деактивуємо кнопку "Сканувати папку"
            self.stop_button.Enable()
            threading.Thread(target=self.process_documents_thread, args=(doc_folder,)).start()

        else:
            dlg.Destroy()
            self.status_label.SetLabel("Вибір папки скасовано")
            self.status_label.SetForegroundColour(wx.RED)

    def _reset_tab2_after_scan(self):
        self.stop_button.Disable()
        self.scan_button.Enable()
        self.is_scanning_active = False 

    def process_documents_thread(self, doc_folder):
        # Перевірка з'єднання з базою даних
        conn = connect_to_database(config.password)
        if not conn:
            wx.CallAfter(self.status_label.SetLabel, "Помилка підключення до БД.")
            wx.CallAfter(self.status_label.SetForegroundColour, wx.RED)
            wx.CallAfter(self._reset_tab2_after_scan)
            return

        try:
            cursor = conn.cursor()
            # Настройка PRAGMA для оптимизации производительности
            cursor.execute("PRAGMA cache_size = 64000;")

            total_files = 0
            processed_files = 0
            skipped_files = 0
            new_records = 0

            # --- Збір списку файлів для обробки ---
            # Фільтруємо файли за розширенням та ігноруємо тимчасові
            files_to_process = []
            allowed_extensions = ['.doc', '.docx', '.rtf'] # Дозволені розширення
            temp_file_prefixes = ('.', '~$', '#', '~')
            temp_file_suffixes = ('~',)

            for root, _, files in os.walk(doc_folder):
                for file in files:
                    filename_lower = file.lower() # Для порівняння без урахування регістру
                    ext = os.path.splitext(filename_lower)[1]

                    # Пропускаємо тимчасові файли
                    if filename_lower.startswith(temp_file_prefixes) or filename_lower.endswith(temp_file_suffixes):
                        # Можна додати до лічильника skipped_files тут, якщо потрібно рахувати пропущені тимчасові
                        # skipped_files += 1 # Якщо потрібно враховувати тимчасові файли у пропущених
                        continue

                    # Обробляємо тільки дозволені розширення
                    if ext in allowed_extensions:
                        files_to_process.append(os.path.join(root, file))
                    # else:
                        # Можна додати до лічильника skipped_files тут, якщо потрібно рахувати пропущені через розширення
                        # skipped_files += 1

            total_files = len(files_to_process)

            wx.CallAfter(self.status_label.SetLabel, f"Обробка файлів... (0/{total_files})")
            wx.CallAfter(self.status_label.SetForegroundColour, wx.GREEN)
            wx.CallAfter(self.output_text.Clear)

            # Використання транзакції для пакетної вставки
            conn.isolation_level = None
            cursor.execute("BEGIN;")

            for i, filepath in enumerate(files_to_process):
                if self.stop_processing:
                    wx.CallAfter(self.status_label.SetLabel, "Обробку зупинено користувачем.")
                    wx.CallAfter(self.status_label.SetForegroundColour, wx.RED)
                    break

                filename = os.path.basename(filepath)
                ext = os.path.splitext(filename.lower())[1] # Переконаємося, що ext також нижнього регістру

                # Ці перевірки тепер, по суті, зайві тут, оскільки файли вже відфільтровані
                # на етапі збору files_to_process. Але подвійна перевірка не зашкодить
                # і явно демонструє, які файли обробляються.
                if ext not in allowed_extensions:
                    skipped_files += 1
                    wx.CallAfter(self.output_text.AppendText, f"Пропущено: '{filename}' (непідтримуване розширення).\n")
                    continue

                content = ""
                extracted_successfully = False

                # --- Вилучення тексту з файлів ---
                if ext == '.docx':
                    try:
                        # Припустимо, Document імпортовано
                        doc = Document(filepath)
                        content = "\n".join([p.text for p in doc.paragraphs])
                        extracted_successfully = True
                    except Exception as e:
                        wx.CallAfter(self.output_text.AppendText, f"Помилка при отриманні тексту з DOCX '{filename}': {e}\n")

                elif ext == '.doc' or ext == '.rtf':
                    try:
                        # Припустимо, extract_text_libreoffice імпортовано
                        content = extract_text_libreoffice(filepath)
                        if content and content.strip():
                            extracted_successfully = True
                        else:
                            wx.CallAfter(self.output_text.AppendText, f"LibreOffice не зміг конвертувати або отримав порожній текст з '{filename}'.\n")
                            extracted_successfully = False
                    except Exception as e:
                        wx.CallAfter(self.output_text.AppendText, f"Помилка при конвертації '{filename}' за допомогою LibreOffice: {e}\n")
                        extracted_successfully = False

                # Перевірка на успішність вилучення тексту
                if not extracted_successfully or not content.strip():
                    wx.CallAfter(self.output_text.AppendText, f"Помилка: не вдалося вилучити текст або він порожній з '{filename}'. Пропускаємо файл.\n")
                    skipped_files += 1
                    continue

                # --- Вилучення дати та номера документа ---
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
                # --- створюємо hash для файлу
                text_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
                
                # --- Перевірка наявності та вставка/оновлення (сумісно зі старими версіями) ---
                # 1. Шукаємо, чи є вже файл з такою назвою та отримуємо його хеш
                cursor.execute("SELECT content_hash FROM documents WHERE filename = ?", (filename,))
                existing_record = cursor.fetchone()

                was_updated_or_added = False

                if existing_record:
                    existing_hash = existing_record[0]
                    # 2. Якщо файл є, порівнюємо хеш вмісту
                    if existing_hash != text_hash:
                        # Якщо existing_hash це None, значить ми вперше бачимо цей файл у новій системі
                        status_msg = "Оновлено вміст" if existing_hash else "Оновлено хеш (старий запис)"                       # Хеш інший — значить файл оновився
                        cursor.execute("""
                        UPDATE documents SET 
                            year = ?, month = ?, day = ?, content = ?, 
                            document_number = ?, created_at = ?, content_hash = ?
                        WHERE filename = ?
                        """, (doc_year, doc_month, doc_day, content, document_number, created_at_str, text_hash, filename))
                        
                        was_updated_or_added = True
                        wx.CallAfter(self.output_text.AppendText, f"Оновлено: {filename} (новий вміст)\n")
                    else:
                        # Хеш однаковий — нічого не робимо
                        wx.CallAfter(self.output_text.AppendText, f"Пропуск: {filename} (вже в базі без змін)\n")
                else:
                    # 3. Файлу взагалі немає в базі — додаємо як новий
                    cursor.execute("""
                    INSERT INTO documents (filename, year, month, day, content, document_number, created_at, content_hash)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (filename, doc_year, doc_month, doc_day, content, document_number, created_at_str, text_hash))
                    
                    was_updated_or_added = True
                    wx.CallAfter(self.output_text.AppendText, f"Додано: {filename}\n")

                # Оновлюємо лічильники
                if was_updated_or_added:
                    new_records += 1
                else:
                    skipped_files += 1

                processed_files += 1

                # Оновлення прогресу
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
            threading.Thread(target=self.update_db_info).start()

            final_status = f"Обробку завершено. Додано {new_records} нових записів. Пропущено {skipped_files} файлів."
            if self.stop_processing:
                final_status = "Обробку зупинено користувачем."
            wx.CallAfter(self.status_label.SetLabel, final_status)
            wx.CallAfter(self.status_label.SetForegroundColour, wx.NullColour)
            wx.CallAfter(self._reset_tab2_after_scan)

    def update_db_info(self):
        """
        Оновлює інформацію про базу даних та виводить її в password_progress_text.
        Викликається в окремому потоці.
        """
        wx.CallAfter(self.password_progress_text.Clear)

        if password is None:
            wx.CallAfter(self.log_password_progress, "Інформація про БД: Пароль не встановлено (або невірний).")
            return

        conn = None
        records = 0
        last_modified = "Недоступно"
        last_update_app_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            if os.path.exists(db_path):
                last_modified_timestamp = os.path.getmtime(db_path)
                last_modified_datetime = datetime.fromtimestamp(last_modified_timestamp)
                last_modified = last_modified_datetime.strftime("%Y-%m-%d %H:%M:%S")
            else:
                last_modified = "Файл БД не знайдено"

            info_message = (
                f"Остання зміна файлу БД: {last_modified}\n"
                f"Час отримання інформації: {last_update_app_time}\n"
            )
            wx.CallAfter(self.password_progress_text.Clear)
            wx.CallAfter(self.log_password_progress, info_message)
        except Exception as e:
            wx.CallAfter(self.log_password_progress, f"Помилка при отриманні інформації про БД: {e}")
        finally:
            pass

    # --- Обробники подій вкладки "Пароль" ---

    def on_change_password(self, event):
        old_pass = self.old_pass_entry.GetValue()
        new_pass1 = self.new_pass1_entry.GetValue()
        new_pass2 = self.new_pass2_entry.GetValue()

        if not old_pass or not new_pass1 or not new_pass2:
            self.log_password_progress("Будь ласка, заповніть усі поля.")
            return

        if old_pass != password:
            self.log_password_progress("Старий пароль невірний.")
            return

        if new_pass1 != new_pass2:
            self.log_password_progress("Новий пароль та підтвердження не збігаються.")
            return

        if new_pass1 == old_pass:
            self.log_password_progress("Новий пароль не може бути таким самим, як старий.")
            return

        self.log_password_progress("Зміна пароля...")
        self.notebook.Enable(False)
        self.change_pass_button.Enable(False)
        threading.Thread(target=self._perform_password_change, args=(old_pass, new_pass1)).start()

    def _perform_password_change(self, old_pass, new_pass):
        global password
        conn = None
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA key = '{old_pass}';")
            cursor.execute("PRAGMA cipher_compatibility = 3;")
            cursor.execute(f"PRAGMA rekey = '{new_pass}';")
            conn.commit()
            conn.close()
            conn = None # Важливо обнулити conn після закриття, щоб finally не намагався закрити вже закрите
            password = new_pass
            wx.CallAfter(self.log_password_progress, "Пароль успішно змінено!")
        except sqlite3.DatabaseError as e:
            if "file is encrypted or is not a database" in str(e) or "not an error" in str(e):
                wx.CallAfter(self.log_password_progress, "Помилка: невірний старий пароль або БД пошкоджена.")
            else:
                wx.CallAfter(self.log_password_progress, f"Помилка бази даних при зміні пароля: {e}")
        except Exception as e:
            wx.CallAfter(self.log_password_progress, f"Несподівана помилка при зміні пароля: {e}")
        finally:
            if conn: 
                conn.close()
            wx.CallAfter(self.change_pass_button.Enable, True)
            # Запускаємо оновлення інформації про БД в окремому потоці
            threading.Thread(target=self.update_db_info).start()
            self.notebook.Enable(True)

    def log_password_progress(self, message):
        wx.CallAfter(self.password_progress_text.AppendText, f"{message}\n")

