# settings_db.py
import os
from sqlcipher3 import dbapi2 as sqlite3

SETTINGS_DB_NAME = "settings.db"

def init_settings_db(master_password):
    """Створює або підключається до зашифрованого файлу налаштувань."""
    conn = sqlite3.connect(SETTINGS_DB_NAME)
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA key = '{master_password}';")
    cursor.execute("PRAGMA cipher_compatibility = 3;")
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS registered_databases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        path TEXT NOT NULL,
        password TEXT,
        is_active INTEGER DEFAULT 1
    );
    """)
    conn.commit()
    return conn

def verify_database_password(path, password):
    """Перевіряє, чи правильний пароль до зашифрованої SQLite (SQLCipher) бази."""
    if not os.path.exists(path):
        return False, "Файл бази даних не знайдено за вказаним шляхом."
    
    try:
        conn = sqlite3.connect(path)
        cursor = conn.cursor()
        cursor.execute(f"PRAGMA key = '{password}';")
        cursor.execute("PRAGMA cipher_compatibility = 3;")
        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        cursor.fetchall()
        
        conn.close()
        return True, "Успішно"
    except Exception as e:
        return False, str(e)

def get_databases_list(master_password):
    """Отримує список усіх зареєстрованих баз даних."""
    if not os.path.exists(SETTINGS_DB_NAME):
        return []
    
    conn = None
    try:
        conn = sqlite3.connect(SETTINGS_DB_NAME)
        cursor = conn.cursor()
        cursor.execute(f"PRAGMA key = '{master_password}';")
        cursor.execute("PRAGMA cipher_compatibility = 3;")
        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        cursor.fetchall()
        
        cursor.execute("SELECT id, name, path, password, is_active FROM registered_databases")
        rows = cursor.fetchall()
        return rows
    except Exception as e:
        raise sqlite3.DatabaseError("Невірний пароль або пошкоджений файл налаштувань.")
    finally:
        if conn:
            conn.close()

def add_database_to_settings(master_password, name, path, password):
    """Додає або оновлює базу даних у конфігурації."""
    conn = sqlite3.connect(SETTINGS_DB_NAME)
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA key = '{master_password}';")
    cursor.execute("PRAGMA cipher_compatibility = 3;")
    
    cursor.execute("""
        INSERT OR REPLACE INTO registered_databases (name, path, password, is_active)
        VALUES (?, ?, ?, 1)
    """, (name, path, password))

    conn.commit()
    conn.close()

def remove_database_from_settings(master_password, db_id):
    """Видаляє зареєстровану базу даних за її ID."""
    conn = sqlite3.connect(SETTINGS_DB_NAME)
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA key = '{master_password}';")
    cursor.execute("PRAGMA cipher_compatibility = 3;")
    
    cursor.execute("DELETE FROM registered_databases WHERE id = ?", (db_id,))
    conn.commit()
    conn.close()
