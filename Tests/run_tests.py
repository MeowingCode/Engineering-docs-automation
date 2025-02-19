import subprocess
from pathlib import Path

# Путь к программе
script_path = Path('CreateDoc.py')

test_folder = Path('Tests/input')

# Удаление всех файлов в папке test_folder
for file in Path('Tests/output').iterdir():
    if file.is_file():
        file.unlink()  # Удаляем файл

# создание новых тестовых файлов
for file in test_folder.iterdir():
    if file.is_file():
        if "спецификация" in file.name: doc_type = "Спецификация"
        elif "расчет_надежности" in file.name: doc_type = "Расчет надежности"
        elif "перечень_элементов" in file.name: doc_type = "Перечень элементов"
        else: doc_type = " "

        print(file.name, doc_type)
        # Запуск скрипта с аргументами
        subprocess.run(['python', script_path, file, doc_type])