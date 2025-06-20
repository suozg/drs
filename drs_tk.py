import os
import re
import pysqlcipher3.dbapi2 as sqlite
from docx import Document
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox, simpledialog
import threading
import time
from tkcalendar import DateEntry
from datetime import datetime   


def show_main_window():


    #########################################################################
    #########################################################################
    #########################################################################
    ####################        PARSING && IMPORT        ####################
    #########################################################################
    #########################################################################
    #########################################################################

    stop_processing = False

    # Глобальные переменные для информации о базе данных
    db_info = {
        "records": 0,
        "updated": "Не оновлювалась"
    }

    def stop_processing_action():
        global stop_processing
        stop_processing = True


    # функция поиска дубликатов файла/записи
    def doubl_file():
        global password  # Используем глобальную переменну

        # Проверка, существует ли база данных
        if not check_patch_db():
            return  # Прерываем выполнение, если база данных не найдена

        try:
            conn = sqlite.connect("db.db")
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA key = '{password}';")

            cursor.execute("""
                SELECT filename, COUNT(*) 
                FROM documents 
                GROUP BY filename 
                HAVING COUNT(*) > 1
            """)

            # Выводим дубликаты
            duplicates = cursor.fetchall()
            for filename, count in duplicates:
                output_text.insert(tk.END, f"Файл {filename} имеет {count} дубликатов.\n")
                output_text.see(tk.END)
                root.update()
 
            conn.close()

        except sqlite.DatabaseError as e:
            messagebox.showinfo("Увага!", f"Помилка бази даних: {e}")
            return



    # функция РУЧНОГО удаления файла/записи
    def delete_selected_file():
        global password  # Используем глобальную переменну

        # Проверка, существует ли база данных
        if not check_patch_db():
            return  # Прерываем выполнение, если база данных не найдена

        start_pass = messagebox.askokcancel("Підтвердження", "Ви дійсно бажаєте видалити запис?")
        if start_pass:
            pass
        else:
            return

        selected = search_output_listbox.curselection()

        if not selected:
            status_label.config(text="Файл не вибрано.", foreground="red")
            return

        filename = search_output_listbox.get(selected[0])

        try:
            conn = sqlite.connect("db.db")
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA key = '{password}';")

            cursor.execute("DELETE FROM documents WHERE filename = ?", (filename,))
            conn.commit()

            if cursor.rowcount == 0:  # Если ничего не удалилось, возможно, файл не найден в БД
                messagebox.showwarning("Увага!", f"Файл {filename} не знайдено в базі.")
            else:
                search_output_listbox.delete(selected[0])
                messagebox.showinfo("Увага!", f"Файл {filename} видалений з бази.")

            conn.close()
            root.update()  # Обновляем интерфейс


        except sqlite.DatabaseError as e:
            messagebox.showinfo("Увага!", f"Помилка бази даних: {e}")
            return



    # функция АВТОМАТИЧЕСКОГО добавления файлов / записей
    def process_documents():
        global password  # Используем глобальную переменную

        output_text.pack(pady=5, expand=True, fill="both")
        status_label.config(text="")

        # Проверка, существует ли база данных
        if not check_patch_db():
            return  # Прерываем выполнение, если база данных не найдена

        global stop_processing
        stop_processing = False
        
        doc_folder = filedialog.askdirectory()

        if not doc_folder:
            status_label.config(text="Вибір теки скасований", foreground="red")
            return

        conn = None
        cursor = None
        try:
            conn = sqlite.connect("db.db")  # Открываем соединение
            cursor = conn.cursor()

            # Попытка выполнить запрос с паролем
            try:
                cursor.execute(f"PRAGMA key = '{password}';")

                # Получаем количество записей
                cursor.execute("SELECT COUNT(*) FROM documents")
                db_info["records"] = cursor.fetchone()[0]  
            
                # Получаем время последнего обновления
                cursor.execute("PRAGMA user_version;")
                db_info["updated"] = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                              
            except sqlite.DatabaseError as e:
                # Если пароль неверный
                if "file is encrypted or is not a database" in str(e):
                    status_label.config(text="Невірний пароль.", foreground="red")
                    return  # Прерываем выполнение, если пароль неверный
                else:
                    status_label.config(text=f"Помилка бази даних: {e}", foreground="red")
                    return

        except Exception as e:
            status_label.config(text=f"Неочікувана помилка: {e}", foreground="red")
            return

        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, f"Починається обробка файлів в {doc_folder}...\n")
        output_text.see(tk.END)
        root.update()



        # Регулярка для поиска номера и даты в названии файла
        filename_pattern = re.compile(r'(\d+)[^\d]*(\d{2})\.(\d{2})\.(\d{4})')

        #  регулярка для поиска года и месяца в пути
        path_pattern = re.compile(r"(\d{4})/(\d{1,2})")

        def extract_text(filepath):
            if filepath.endswith((".docx", ".doc", ".rtf")):
                if filepath.endswith(".docx"):
                    try:
                        doc = Document(filepath)
                        return "\n".join([p.text for p in doc.paragraphs])
                    except Exception as e:
                        return f"Помилка при обробці .docx: {e}"
                elif filepath.endswith(".doc"):
                    try:
                        result = subprocess.run(["/usr/bin/antiword", filepath], capture_output=True, text=True)
                        return result.stdout.strip()
                    except Exception as e:
                        return f"Помилка при обробці .doc: {e}"
                elif filepath.endswith(".rtf"):
                    try:
                        # Обработка .rtf файлов с использованием unrtf
                        result = subprocess.run(["unrtf", "--text", filepath], capture_output=True, text=True)
                        return result.stdout.strip()
                    except Exception as e:
                        return f"Помилка при обробці .rtf: {e}"
            else:
                return ""  # Возвращаем пустую строку, если файл не поддерживаемого формата


        # doc_folder -  каталог, который нужно сканировать
        for root_dir, _, files in os.walk(doc_folder):
            if stop_processing:
                break
            match = path_pattern.search(root_dir)
            if not match:
                continue

            year_from_path, month_from_path = match.groups()

            for filename in files:
                if stop_processing:
                    break
                filepath = os.path.join(root_dir, filename)
                if filename.startswith(('.', '~$', '#', '~')) or filename.endswith('~'):
                    continue

                # Поиск даты и номера в названии файла
                filename_match = filename.lower()
                document_number = None  # Значение по умолчанию
                date_match = filename_pattern.search(filename_match)

                if date_match:
                    try:
                        document_number, day, month, year = date_match.groups()
                        year, month, day = int(year), int(month), int(day)
                    except ValueError:
                        year, month, day = int(year_from_path), int(month_from_path), None
                else:
                    year, month, day = int(year_from_path), int(month_from_path), None

                text = extract_text(filepath)
                if not text:
                    continue

                try:
                    # Проверка соединения перед вставкой данных
                    if conn:
                        # удаляем старий файл в базе   
                        cursor.execute("DELETE FROM documents WHERE filename = ?", (filename,))

                        if cursor.rowcount > 0:
                            output_text.insert(tk.END, f"Існуючий в базі дублікат {filename} видалено.\n")
  
                        # вставляем новий файл в базе  
                        cursor.execute("""
                            INSERT INTO documents (filename, year, month, day, content, document_number, created_at)
                            VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                        """, (filename, year, month, day, text, document_number))
                        conn.commit()

                        output_text.insert(tk.END, f"Файл {filename} успішно доданий.\n")                                  
                        output_text.see(tk.END)
                        root.update()

                    else:
                        output_text.insert(tk.END, f"Не вдалось з'єднатися з базою даних.\n")
                        output_text.see(tk.END)
                        root.update()
                except sqlite.DatabaseError as e:
                    output_text.insert(tk.END, f"Помилка при додаванні файлу {filename}: {e}\n")
                    output_text.see(tk.END)
                    root.update()

        status_label.config(text=f"Обробка завершена\n{db_info}", foreground="green")
        root.update()

        # Закрытие соединения в конце
        if conn:
            conn.close()

















    #########################################################################
    #########################################################################
    #########################################################################
    ############################ SEARCH #####################################
    #########################################################################
    #########################################################################
    #########################################################################
    def howto():

        help_search = f"\n" \
            "1. Пошук одного слова\n" \
            "\tслово\n" \
            "Знайде усі записи, що містять указане слово.\n\n" \
            "2. Пошук декілька слів (AND)\n" \
            "\tслово1 слово2\n" \
            "Знайде записи, що містять обидва слова (слово1 и слово2).\n\n" \
            "3. Пошук з OR\n" \
            "\tслово1 OR слово2\n" \
            "Знайде записи, які містять хоча б одне зі слів.\n\n" \
            "4. Пошук фрази (використовуємо лапки)\n" \
            "\tʼточная фразаʼ\n" \
            "Знайде точне співпадіння заданої фрази, зберігаючи порядок слів.\n\n" \
            "5. Виключення слів (NOT, -)\n" \
            "\tслово1 -слово2\n" \
            "Знайде записи, які містять слово1, але не містять слово2.\n\n" \
            "6. Пошук за близкістю слів (NEAR)\n" \
            "\tслово1 NEAR/5 слово2\n" \
            "Знайже записи, в яких слово1 и слово2 знаходяться не далі ніж в 5 словах одне від іншого.\n\n" \
            "7. Пошук за префіксом (частина слова)\n" \
            "\tсло*\n" \
            "Знайде усі слова, які починаються з сло (наприклад, слово, слот)."

        # Включаем возможность редактирования
        content_text.config(state=tk.NORMAL)
        content_text.delete("1.0", tk.END)  # Очищаем предыдущее содержимое

        # Разбиение текста и вставка с раскраской
        lines = help_search.split("\n")
        for line in lines:
            if line.strip().isdigit() or line.endswith("пошуку") or "." in line[:3]:  # Заголовки
                content_text.insert("end", line + "\n", "header")
            elif line.startswith("\t"):  # Ключевые слова (с отступом)
                content_text.insert("end", line.strip() + "\n", "keyword")        
            else:  # Описание
                content_text.insert("end", line + "\n", "description")

        # Создаём стили для разных частей текста
        content_text.tag_config("header", foreground="blue", font=("Sans", 9, "bold"))
        content_text.tag_config("keyword", foreground="red", font=("Sans", 8, "bold"))
        content_text.tag_config("description", foreground="black", font=("Sans", 8))

        # Отключаем редактирование обратно
        content_text.config(state=tk.DISABLED)


    def format_date_for_fts3(queries: str) -> str:
        # Преобразуем точки (например даты) в формат NEAR/0
        queries = re.sub(r'(\d{2})\.(\d{2})\.(\d{2,4})', r'\1 NEAR/0 \2 NEAR/0 \3', queries)
        return queries


    def search_documents(event=None):
        global password  # Используем глобальную переменную

        # Проверка, существует ли база данных
        if not check_patch_db():
            return  # Прерываем выполнение, если база данных не найдена
       
        query = search_entry.get()
        if not query:
            search_output_listbox.insert(tk.END, "Введіть запит для пошуку.")
            return    

        formatted_query = format_date_for_fts3(query)

        progress_bar.pack(padx=10, pady=10, fill="x")
        progress_bar["value"] = 0
        search_thread = threading.Thread(target=perform_search, args=(password, formatted_query, query))
        search_thread.start()

    def perform_search(password, formatted_query, query):
        start_date = start_date_entry.get()
        end_date = end_date_entry.get()
        # Проверка на наличие обеих дат
        if not start_date or not end_date:
            search_output_listbox.insert(tk.END, "Оберіть обидві дати.")
            return

        try:
            conn = sqlite.connect("db.db")
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA key = '{password}';")

            # Если даты введены, преобразуем их в формат YYYYMMDD
            if start_date and end_date:
                # Преобразуем даты в объект datetime
                start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
                end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")

                # Преобразуем в формат YYYYMMDD
                start_date_num = int(start_date_obj.strftime("%Y%m%d"))
                end_date_num = int(end_date_obj.strftime("%Y%m%d"))
           
                # Проверка, чтобы конечная дата не была раньше начальной
                if end_date_num < start_date_num:
                    search_output_listbox.insert(tk.END, "Кінцева дата не може бути раніше початкової.")
                    progress_bar.pack_forget()
                    return

                # Если даты одинаковые, ищем без учета дат
                if start_date_obj == end_date_obj:
                    sql_query = """
                    SELECT filename, content, created_at FROM documents
                    WHERE content MATCH ?
                    ORDER BY year DESC, month DESC, day DESC
                    """
                    params = (formatted_query,)
                else:
                    sql_query = """
                    SELECT filename, content, created_at FROM documents
                    WHERE content MATCH ?
                    AND (year * 10000 + month * 100 + day BETWEEN ? AND ?)
                    ORDER BY year DESC, month DESC, day DESC
                    """
                    params = (formatted_query, start_date_num, end_date_num)


            cursor.execute(sql_query, params)
            results = cursor.fetchall()
            conn.close()

        except sqlite.DatabaseError as e:
            search_output_listbox.insert(tk.END, f"Помилка бази даних: {e}")
            progress_bar.pack_forget()
            return
        
        search_output_listbox.delete(0, tk.END)
        documents.clear()
 
        if results:
            howto()
            for index, (filename, content, created_at) in enumerate(results):
                doc_key = f"{filename}"
                documents[doc_key] = (content, created_at)             
                search_output_listbox.insert(tk.END, doc_key)
                
                # Выделение нечётных строк светло-серым цветом
                if index % 2 != 0:  # нечётный индекс
                    search_output_listbox.itemconfig(index, {'bg': '#f0f0f0'})  # светло-серый цвет
        else:
            search_output_listbox.insert(tk.END, "Нічого не знайдено.")


        progress_bar.pack_forget()
        x=len(results)
        update_count_label(x)  # Обновление метки после поиска
        # Удаляем все виды кавычек и звездочки
        query = re.sub(r'["\'`‘’“”*]', '', query)
        first_word = query.split()[0]  # Разделяем строку и берем первое слово

        # вставляем текст запроса в поле поиска по тексту
        search_text_entry.delete(0, 'end')  # Очищаем поле
        search_text_entry.insert(0, first_word)  # Вставляем новый текст


    # Функция для обновления метки с количеством записей
    def update_count_label(count):
        if count>1:
            # Обновляем существующую метку
            count_label.config(text=f"Записів: {count}")
            count_label.pack(side="left", padx=5)
        else:
            count_label.pack_forget()

    def display_document(event):
        selected = search_output_listbox.curselection()
        if selected:
            # Получаем имя файла из выбранного элемента
            filename = search_output_listbox.get(selected[0])
            # Извлекаем содержимое и дату из словаря по имени файла
            content, created_at = documents.get(filename, ("", ""))  # По умолчанию пустое содержимое и дата         
            content_text.config(state=tk.NORMAL)
            content_text.delete("1.0", tk.END)
            content_text.insert(tk.END, content)
            content_text.config(state=tk.DISABLED)
            view_filename.config(text=f"{filename} (дані додані: {created_at})")            
            search_in_text()


    def search_in_text():
        global matches, match_index  # Объявляем глобальные переменные

        query = search_text_entry.get()
        if not query:
            return

        content_text.tag_remove("highlight", "1.0", tk.END)
        matches = []  # Очищаем список перед новым поиском

        start_idx = "1.0"

        while True:
            start_idx = content_text.search(query, start_idx, stopindex=tk.END, nocase=True)
            if not start_idx:
                break
            end_idx = f"{start_idx}+{len(query)}c"
            content_text.tag_add("highlight", start_idx, end_idx)
            matches.append(start_idx)  # Добавляем найденную позицию
            start_idx = end_idx

        content_text.tag_config("highlight", background="lightblue")

        if matches:
            match_index = 0
            go_to_match(match_index)  # Перейти к первому совпадению 

        # обновляем инфо-строку
        get_view_filename = view_filename.cget("text").split(" - [ збігів:")
        view_filename.config(text=f"{get_view_filename[0]} - [ збігів: {len(matches)} ]")


    def go_to_match(index):
        """Прокручивает текст к нужному совпадению."""
        global match_index
        if matches:
            match_index = index % len(matches)  # Зацикливание при достижении конца
            content_text.see(matches[match_index])
            content_text.tag_remove("current_highlight", "1.0", tk.END)
            content_text.tag_add("current_highlight", matches[match_index], f"{matches[match_index]}+{len(search_text_entry.get())}c")
            content_text.tag_config("current_highlight", background="yellow")


    def next_match():
        """Переход к следующему совпадению."""
        if matches:
            go_to_match(match_index + 1)


    def prev_match():
        """Переход к предыдущему совпадению."""
        if matches:
            go_to_match(match_index - 1)



    ################################ FORMS ##################################

    root = tk.Tk()
    root.title("DRS TK")
    # Задайте бажані розміри вікна (наприклад, 400x300 пікселів)
    width = 1050
    height = 650

    # Задайте бажане положення вікна на екрані (наприклад, x=100, y=150)
    x_position = 10
    y_position = 10

    # Створіть рядок geometry
    geometry_string = f"{width}x{height}+{x_position}+{y_position}"

    # Застосуйте geometry до головного вікна
    root.geometry(geometry_string)

    # Если DPI меньше 96, устанавливаем scaling_factor = 1.3, иначе оставляем стандартное значение
    # Количество пикселей в одном дюйме 
    dpi = root.winfo_fpixels('1i')
    if isinstance(dpi, (int, float)) and dpi < 97:
        scaling_factor = 1.3
        root.tk.call("tk", "scaling", scaling_factor)

    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill="both")

    tab1 = ttk.Frame(notebook)
    tab2 = ttk.Frame(notebook)
    tab3 = ttk.Frame(notebook)

    notebook.add(tab1, text="Пошук")
    notebook.add(tab2, text="Імпорт")
    notebook.add(tab3, text="Пароль")















    #########################  TAB 1 ########################################
    documents = {}


    # "Пошук"
    first_frame = ttk.Frame(tab1)
    first_frame.pack(side="top", fill='both', expand=True)

    search_frame = ttk.LabelFrame(first_frame, text=" Пошук наказів ")
    search_frame.pack(side="top", fill='both', expand=True, padx=10, anchor="n")

    ttk.Label(search_frame, text="Запит:").pack(side="left", padx=5)
    search_entry = ttk.Entry(search_frame, width=30)
    search_entry.pack(side="left", padx=5)

    # Создаем метку найдених записей
    count_label = ttk.Label(search_frame, text="")

    ttk.Label(search_frame, text="Початкова дата:").pack(side="left", padx=5)
    start_date_entry = DateEntry(search_frame, width=12, date_pattern="yyyy-MM-dd")
    start_date_entry.pack(side="left", padx=5)

    ttk.Label(search_frame, text="Кінцева дата:").pack(side="left", padx=5)
    end_date_entry = DateEntry(search_frame, width=12, date_pattern="yyyy-MM-dd")
    end_date_entry.pack(side="left", padx=5)

    search_button = ttk.Button(search_frame, text=" Пошук ", command=search_documents)
    search_button.pack(side="left", padx=5, pady=10)
    root.bind("<Return>", search_documents)

    delete_button = tk.Button(search_frame, text="Видалити файл", command=delete_selected_file)
    delete_button.pack(side="left", padx=5, pady=10)

    # Основнойфрейм с Listbox
    main_frame = ttk.Frame(tab1)
    main_frame.pack(expand=True, fill="both", padx=10, pady=5)

    listbox_frame = ttk.Frame(main_frame)
    listbox_frame.pack(side="left", fill="y")

    scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical")
    search_output_listbox = tk.Listbox(listbox_frame, height=20, width=30, yscrollcommand=scrollbar.set)
    search_output_listbox.pack(side="left", fill="y")
    scrollbar.config(command=search_output_listbox.yview)
    scrollbar.pack(side="right", fill="y")

    search_output_listbox.bind("<<ListboxSelect>>", display_document)

    # Фрейм для отображения текста
    text_frame = ttk.Frame(main_frame)
    text_frame.pack(side="right", expand=True, fill="both")

    content_text = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, state=tk.DISABLED)
    content_text.pack(expand=True, fill="both")

    howto()


    ############ пошук по окремому документу

    input_frame = ttk.Frame(text_frame)
    input_frame.pack(fill="x", expand=True, pady=5)
    view_filename = ttk.Label(input_frame, text="")
    view_filename.pack(side="top", padx=5, expand=True, fill="both")

    ttk.Label(input_frame, text=" Знайти в тексті: ").pack(side="left", padx=5)
    search_text_entry = ttk.Entry(input_frame, width=40)
    search_text_entry.pack(side="left", padx=5)

    ttk.Button(input_frame, text=" Шукати ", command=search_in_text).pack(side="left", padx=5)
    # Кнопки навигации
    prev_button = tk.Button(input_frame, text=" ⬅ Назад ", command=prev_match)
    prev_button.pack(side=tk.LEFT)

    next_button = tk.Button(input_frame, text=" Вперед ➡ ", command=next_match)
    next_button.pack(side=tk.RIGHT)


    progress_bar = ttk.Progressbar(tab1, orient="horizontal", length=300, mode="determinate")


















    #########################  TAB 2 ########################################
    #  фрейм для импорта
    import_frame = ttk.Frame(tab2)
    import_frame.pack(side="top", fill='both', expand=True)

    pass_frame = ttk.LabelFrame(import_frame, text=" Імпорт даних ")
    pass_frame.pack(side="top", fill='both', expand=True, padx=5, anchor="n")

    ttk.Label(pass_frame, text="УВАГА! Дозволені тільки документи .doc, .docx, .rtf. Перевірте каталог на відсутність інших форматів.", foreground="brown").pack(side="left", anchor="n", pady=5, padx=5)

    #  фрейм для buttons 
    buttons_frame = ttk.Frame(tab2)
    buttons_frame.pack(side="top", fill="both", expand=True)

    ttk.Button(buttons_frame, text=" Сканувати теку ", command=process_documents).pack(pady=10, padx=10, fill='x', side="left", anchor="n")
    ttk.Button(buttons_frame, text=" Зупинити ", command=stop_processing_action).pack(pady=10, padx=10, fill='x', side="right", anchor="n")
 
    status_label = ttk.Label(buttons_frame, text="")
    status_label.pack(pady=15, side="left", fill="x", expand=True, anchor="n")

    #  фрейм для text
    text_frame = ttk.Frame(tab2)
    text_frame.pack(side="bottom", fill="x", expand=True)

    output_text = scrolledtext.ScrolledText(text_frame)










    ############################## TAB 3 ###########################################
    # вкладка смени пароля

    # Создание фреймов для каждого блока
    frame_6_top = ttk.Frame(tab3)
    frame_6_top.pack(side="top", fill='both', expand=True)

    # Фрейм для отображения текста
    frame_progress_text = ttk.Frame(tab3)
    frame_progress_text.pack(expand=True, fill="both")

    # Поля ввода паролей и кнопка пуск
    group_6_1 = ttk.LabelFrame(frame_6_top, text=" Форма зміни пароля ")
    group_6_1.pack(side="top", fill='both', expand=True, padx=5, anchor="n")

    ttk.Label(group_6_1, text="ввести старий пароль").grid(row=1, column=1,  sticky=tk.W, padx=3)
    pass1 = ttk.Entry(group_6_1, show="*")
    pass1.grid(row=1, column=2, sticky="nsew", padx=3, pady=3)

    ttk.Label(group_6_1, text="ввести новий пароль").grid(row=2, column=1, sticky=tk.W, padx=3)
    pass2 = ttk.Entry(group_6_1, show="*")
    pass2.grid(row=2, column=2, sticky="nsew", padx=3, pady=3)

    ttk.Label(group_6_1, text="повторити новий пароль").grid(row=3, column=1, sticky=tk.W, padx=3)
    pass3 = ttk.Entry(group_6_1, show="*")
    pass3.grid(row=3, column=2, sticky="nsew", padx=3, pady=3)


    # Поле для индикатора прогресса
    progress_text = scrolledtext.ScrolledText(frame_progress_text, wrap=tk.WORD)

    def chekPassChange():
        oldPass = pass1.get()
        newPass1 = pass2.get()
        newPass2 = pass3.get()

        if len(newPass1) == 0 or len(newPass2) == 0:
            messagebox.showinfo("Увага!", "Порожні паролі неприйнятні.")
            return
        if newPass1 != newPass2:
            messagebox.showinfo("Увага!", "Нові паролі не збігаються.")
            return

        # Показываем индикатор прогресса
        progress_text.pack(expand=True, fill="both")
        progress_text.insert(tk.END, "\nУВАГА! НЕ ПЕРЕРИВАЙТЕ ПРОЦЕС! Розмір бази впливає на час зміни пароля (5-10 хв).")
        tab3.update_idletasks()  # Обновляем интерфейс

        db = None  # Инициализируем db как None
        try:
            db = sqlite.connect('db.db')
            cursor = db.cursor()

            # Подключаемся к базе с текущим паролем            
            cursor.execute(f"PRAGMA key = '{oldPass}';")
            progress_text.insert(tk.END, "\nПеревірка старого паролю...")
            tab3.update_idletasks()

            # Проверяем доступность базы
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()

            if not tables:  # Если список пуст, значит база не расшифровалась
                progress_text.insert(tk.END, "\nВведено невірний старий пароль.")
                tab3.update_idletasks()
                return

            # Временно блокируем кнопку, чтобы пользователь не нажал её несколько раз
            passChan.config(state=tk.DISABLED)

            # Меняем пароль базы
            progress_text.insert(tk.END, "\nОК\nШифруємо дані новим паролем... ")
            tab3.update_idletasks()

            cursor.execute(f"PRAGMA rekey = '{newPass2}';")
            db.commit()

            # Проверка успешности изменения пароля
            progress_text.insert(tk.END,  "\nПароль успішно змінено. Перезапустіть програму.")
            tab3.update_idletasks()


        except sqlite.Error as e:
            progress_text.insert(tk.END, f"\nПомилка бази даних: {e}")
            tab3.update_idletasks()

        finally:
            if db:
                db.close()  

            # разблокируем кнопку        
            passChan.config(state=tk.NORMAL)
            tab3.update_idletasks()
           

    # Инициализация кнопки после определения всех переменных
    passChan = ttk.Button(group_6_1, text=" Змінити пароль ", width=15, command=chekPassChange)
    passChan.grid(row=1, column=3, rowspan=3, sticky="nsew", padx=13, pady=5)

    # ----------------- END TAB 3 ----------------------














    # Закрытие приложения при закрытии окна
    root.protocol("WM_DELETE_WINDOW", root.quit)  # Позволяет завершить программу при закрытии окна
    root.mainloop()


def check_patch_db():
    # проверка на наличие базы данных
    if not os.path.exists("db.db"):
        # Показываем всплывающее окно с сообщением
        messagebox.showerror("Помилка", "База даних не знайдена.")
        return False
    return True  # База данных существует


def connect_to_database(password):

    # Проверка, существует ли база данных
    if not check_patch_db():
        try:
            # Создаем новую базу данных, если ее нет
            conn = sqlite.connect('db.db')  # Укажите путь к вашей базе данных
            cursor = conn.cursor()

            # Вводим ключ (пароль) для доступа к базе
            cursor.execute(f"PRAGMA key = '{password}';")

            # Создаем таблицу, если ее нет
            cursor.execute("""
            CREATE VIRTUAL TABLE IF NOT EXISTS documents USING fts3(
                filename TEXT,
                year INTEGER,
                month INTEGER,
                day INTEGER,
                content TEXT,
                document_number INTEGER,
                created_at TEXT,
                tokenize=unicode61
            );
            """)
            conn.commit()
            return conn  # Подключение успешно, база данных была создана
        except Exception as e:
            print(f"Ошибка при создании базы данных: {e}")
            return None
    else:
        try:
            """Функция для подключения к базе данных с использованием пароля"""
            # Создаем подключение к базе данных с SQLCipher (SQLite с зашифрованной базой)
            conn = sqlite.connect('db.db')  # Укажите путь к вашей базе данных
            cursor = conn.cursor()

            # Вводим ключ (пароль) для доступа к базе
            cursor.execute(f"PRAGMA key = '{password}';")
            
            # Пробуем выполнить запрос, чтобы убедиться, что пароль правильный
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            
            if tables:
                return conn  # Подключение успешно
            else:
                conn.close()
                return None  # Неверный пароль
        except Exception as e:
            return None


# Функция для запроса пароля 
password = None  # Глобальная переменная для пароля
def ask_password():
    global password  # Используем глобальную переменную

    # Модальное окно для ввода пароля
    password = simpledialog.askstring("Ввод пароля", "Введіть пароль доступу:", show="*")

    if password:

        # Попытка подключения к базе данных
        conn = connect_to_database(password)
        
        if conn:
            # Закрыть окно для пароля после успешного ввода
            #password_window.destroy()  # Закрыть окно
            show_main_window()  # Если пароль правильный, показать основное окно

        else:
            messagebox.showerror("Помилка", "Невірний пароль.")
            ask_password()  # Запросить пароль снова

# Запуск запроса пароля
ask_password()
