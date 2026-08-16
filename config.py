import re

# Глобальні змінні (можуть змінюватися під час виконання)
password = None
frame_title = "Document Retrieval System"
db_path = "db.db"

# Регулярний вираз для вилучення дати з імені файла
filename_date_pattern = re.compile(r'(?:від\s?)?(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})')

# Регулярний вираз для вилучення номера документа
document_number_pattern = re.compile(r'(?:№|\s|^)(\d+)')

# Регулярка для вилучення року та місяця з шляху
path_date_pattern = re.compile(r"[/\\](\d{4})[/\\](\d{1,2})[/\\]?")
