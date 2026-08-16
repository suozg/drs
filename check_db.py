import os
import sys
import argparse
from sqlcipher3 import dbapi2 as sqlite3
import config

def inspect_database():
    parser = argparse.ArgumentParser(description='Check database connection.')
    parser.add_argument('-c', type=str, default=config.db_path, help='Path to the database file')
    args = parser.parse_args()
    
    db_file = args.c
    print(f"Шлях до файлу БД: {os.path.abspath(db_file)}")
    print(f"Файл існує: {os.path.exists(db_file)}")
    
    if not os.path.exists(db_file):
        print("Помилка: Файл бази даних не знайдено за цим шляхом!")
        return

    password = input("Введіть пароль від бази даних: ").strip()

    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        cursor.execute(f"PRAGMA key = '{password}';")
        cursor.execute("PRAGMA cipher_compatibility = 3;")
        cursor.execute("PRAGMA kdf_iter = 64000;")        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        
        print("\n Успішно підключено! База розшифрована.")
        print("Знайдені таблиці в БД:", tables)
        
        cursor.execute("PRAGMA cipher_version;")
        print("Версія SQLCipher:", cursor.fetchone())
        
        conn.close()
    except Exception as e:
        print(f"\n Помилка підключення або невірний пароль: {e}")

if __name__ == "__main__":
    inspect_database()
