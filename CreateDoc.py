from Modules import Constructor as C
import argparse
from pathlib import Path

def main():

    # значение по умолчанию
    def_data_path = "table_test.xml"
    def_mode = None 

    handler = argparse.ArgumentParser(description="Создание конструкторской документации")
    handler.add_argument('data_path', type=str, help='Путь к входному файлу.', default = def_data_path, nargs='?')
    handler.add_argument('type', type=str, help='Тип документации',default=def_mode, nargs='?')
    
    args = handler.parse_args()
    
    # Получаем путь к текущему файлу
    current_file_path = Path(__file__).resolve().parent

    C.PATH_TO_SAMPLES = current_file_path / "Samples"
    C.PATH_TO_NEW_FILES = current_file_path / "Docs"
    C.CAN_REPLACE_FILES = False

    constructor = C.Constructor(args.data_path, args.type)
    constructor.construct_document()



if __name__ == "__main__":
    main()
