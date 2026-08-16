import os
from sqlcipher3 import dbapi2 as sqlite3 
import wx
from utils import get_full_db_path
import config

def check_patch_db():
    """Перевіряє, чи існує файл бази даних."""
    if not os.path.exists(config.db_path):
        full_db_path, _ = get_full_db_path(config.db_path)
        wx.MessageBox(f"Під час роботи втрачено доступ до бази даних, перевірте файл {full_db_path}.", "База даних не знайдена", wx.OK | wx.ICON_ERROR)
        return False
    return True

def sqlite_lower(value_):
    """Функція для приведення рядка до нижнього регістру для SQL-запитів."""
    return str(value_).lower()

def connect_to_database(db_password):
    """
    Підключається до зашифрованої бази даних з паролем через sqlcipher3.
    """
    conn = None
    try:
        conn = sqlite3.connect(config.db_path)
        cursor = conn.cursor()
        
        # Встановлюємо ключ та обов'язкову сумісність для розшифрування існуючої бази
        cursor.execute(f"PRAGMA key = '{db_password}';")
        cursor.execute("PRAGMA cipher_compatibility = 3;")  
        
        # Оптимізація швидкодії
        cursor.execute("PRAGMA journal_mode=DELETE;")
        cursor.execute("PRAGMA synchronous=NORMAL;")
        cursor.execute("PRAGMA cache_size = 64000;")
        cursor.execute("PRAGMA temp_store = MEMORY;")

        # Перевірка наявності таблиці 'documents'
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='documents';")
        main_table_exists = cursor.fetchone()

        if not main_table_exists:
            full_db_path, _ = get_full_db_path(config.db_path)
            wx.MessageBox(f"Не знайдено файл бази даних. \nБуде створений новий за адресою {full_db_path}.", "Відсутня база даних", wx.OK | wx.ICON_ERROR)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT UNIQUE NOT NULL,
                year INTEGER,
                month INTEGER,
                day INTEGER,
                content TEXT,
                document_number TEXT,
                created_at TEXT,
                content_hash TEXT
            );
            """)

            cursor.execute("""
            CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts3(
                filename, content,
                tokenize=unicode61
            );
            """)

            cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS documents_ai AFTER INSERT ON documents BEGIN
                INSERT INTO documents_fts(docid, filename, content) VALUES (new.id, new.filename, new.content);
            END;
            """)
            cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS documents_ad AFTER DELETE ON documents BEGIN
                DELETE FROM documents_fts WHERE docid = old.id;
            END;
            """)
            cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS documents_au AFTER UPDATE ON documents BEGIN
                UPDATE documents_fts SET filename = new.filename, content = new.content WHERE docid = old.id;
            END;
            """)
            conn.commit()
        else:
            try:
                cursor.execute("ALTER TABLE documents ADD COLUMN content_hash TEXT;")
                conn.commit()
            except Exception:
                pass           

        conn.create_function("LOWER", 1, sqlite_lower)
        return conn

    except Exception as e:
        print(f"Помилка підключення до бази даних: {e}")
        if conn:
            conn.close()
        return None
