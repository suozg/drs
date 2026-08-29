# config.py
import os
from pathlib import Path
import re

# Глобальні змінні (можуть змінюватися під час виконання)
password = None
frame_title = "Document Retrieval System"
master_password = ""
db_path = ""

def get_config_dir(app_name="drs"):
    """Повертає шлях до стандартної папки конфігурації ОС ( ~/.config/drs для Linux, AppData/Roaming/drs для Windows )."""
    if os.name == "nt":  # Windows
        base_dir = os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")
    else:  # Linux / macOS
        base_dir = os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")
    
    config_dir = Path(base_dir) / app_name
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir

# Повний шлях до файлу settings.db у стандартному системному каталозі
SETTINGS_DB_PATH = str(get_config_dir("drs") / "settings.db")

# Регулярний вираз для вилучення дати з імені файла
filename_date_pattern = re.compile(r'(?:від\s?)?(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})')

# Регулярний вираз для вилучення номера документа
document_number_pattern = re.compile(r'(?:№|\s|^)(\d+)')

# Регулярка для вилучення року та місяця з шляху
path_date_pattern = re.compile(r"[/\\](\d{4})[/\\](\d{1,2})[/\\]?")
