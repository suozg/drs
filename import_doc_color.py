import os
import re
import pysqlcipher3.dbapi2 as sqlite
from docx import Document
import subprocess
import getpass
from colorama import init, Fore

# Инициализация colorama
init(autoreset=True)

# Запрашиваем пароль у пользователя
password = getpass.getpass("Введите пароль для базы данных: ")

# Подключаемся к зашифрованной базе
conn = sqlite.connect("documents_encrypted.db")
cursor = conn.cursor()
cursor.execute(f"PRAGMA key = '{password}';")

# Создаем таблицу FTS3, если её нет
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

# Функция для извлечения текста из .docx
def extract_text_docx(filepath):
    try:
        doc = Document(filepath)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        print(f"{Fore.RED}Ошибка при извлечении текста из {filepath}: {e}")
        return ""

# Функция для извлечения текста из .doc
def extract_text_doc(filepath):
    try:
        result = subprocess.run(["/usr/bin/antiword", filepath], capture_output=True, text=True)
        return result.stdout.strip()
    except Exception as e:
        print(f"{Fore.RED}Ошибка при извлечении текста из {filepath}: {e}")
        return ""

def extract_text_libreoffice(filepath):
    try:
        output_dir = "/tmp"
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "txt:Text", "--outdir", output_dir, filepath],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise Exception(f"Ошибка при конвертации {filepath}: {result.stderr}")

        # Определяем имя файла после конвертации
        converted_filename = os.path.basename(filepath).rsplit(".", 1)[0] + ".txt"
        converted_filepath = os.path.join(output_dir, converted_filename)

        # Читаем содержимое сконвертированного файла
        with open(converted_filepath, "r", encoding="utf-8") as file:
            text = file.read().strip()

        # Удаляем временный файл после использования
        os.remove(converted_filepath)

        return text

    except Exception as e:
        print(f"{Fore.RED}Ошибка при извлечении текста из {filepath} с помощью LibreOffice: {e}")
        return ""


# Запрашиваем у пользователя начальный каталог
while True:
    doc_folder = input(f"{Fore.YELLOW}Введите путь к каталогу с документами: ").strip()
    if os.path.isdir(doc_folder):
        break
    print(f"{Fore.RED}Ошибка: указанный каталог не существует.")

# Регулярное выражение для извлечения даты (день, месяц, год) из имени файла
filename_pattern = re.compile(r'(\d+)[^\d]*(\d{2})\.(\d{2})\.(\d{4})')

# Регулярка для извлечения года и месяца из пути
path_pattern = re.compile(r"(\d{4})/(\d{1,2})")

# Обход файлов
for root, _, files in os.walk(doc_folder):
    match = path_pattern.search(root)
    if not match:
        continue

    year_from_path, month_from_path = match.groups()

    for filename in files:
        filepath = os.path.join(root, filename)

        # Пропускаем временные файлы (начинаются с . или с ~$, #, ~ или заканчиваются на ~)
        if filename.startswith(('.', '~$', '#', '~')) or filename.endswith('~'):
            continue

        filename_match = filename.lower()
        date_match = filename_pattern.search(filename_match)

        if date_match:
            try:
                document_number, day, month, year = date_match.groups()
                year, month, day = int(year), int(month), int(day)
            except ValueError:
                year, month, day = int(year_from_path), int(month_from_path), None
        else:
            year, month, day = int(year_from_path), int(month_from_path), None

        # Проверяем, есть ли уже такой файл в базе
        #cursor.execute("SELECT COUNT(*) FROM documents WHERE filename = ?", (filename,))
        #exists = cursor.fetchone()[0] > 0
        exists = None
        if exists:
            choice = input(f"{Fore.YELLOW}Файл {filename} уже есть в базе. Заменить? (y/n): ").strip().lower()
            if choice != "y":
                print(f"{Fore.GREEN}Пропускаем файл {filename}")
                continue

        # Извлекаем текст
        text = extract_text_docx(filepath) if filename.endswith(".docx") else extract_text_doc(filepath)

        # Если текст пустой, пробуем повторно конвертировать с помощью LibreOffice
        if not text:
            print(f"{Fore.YELLOW}Текст из файла {filename} не был извлечен. Попытка конвертировать с помощью LibreOffice.")
            text = extract_text_libreoffice(filepath)

        if not text:
            print(f"{Fore.RED}Ошибка: не удалось извлечь текст из {filename}. Пропускаем файл.")
            continue

        if exists:
            pass
            #cursor.execute("""
            #    UPDATE documents
            #    SET year = ?, month = ?, day = ?, content = ?, document_number = ?
            #    WHERE filename = ?
            #""", (year, month, day, text, document_number, filename))
        else:
            cursor.execute("""
                INSERT INTO documents (filename, year, month, day, content, document_number, created_at)
                VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """, (filename, year, month, day, text, document_number))
        
        conn.commit()
        print(f"{Fore.GREEN}Файл {filename} обработан и добавлен/обновлен в базе. Вес {len(text)} символов")

conn.close()
