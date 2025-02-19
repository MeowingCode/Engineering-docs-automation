from pathlib import Path
import shutil

from Modules import Parser as pr
from Modules import Builder as bld

PATH_TO_SAMPLES = ""
PATH_TO_NEW_FILES = ""
CAN_REPLACE_FILES = False

# Управляет процессом создания документа (Директор по ШП Строитель)
class Constructor: 
    def __init__(self, xml_file_path, doc_type = None):
        
        self.type = doc_type
        # проверка файла-инструкции 
        self.instruction = Path(xml_file_path)
        if not self._check_file(self.instruction):
            raise ValueError('Не удается найти файл ' + str(self.instruction))
        if self.instruction!=None and self.instruction.suffix != ".xml": 
            raise ValueError("Неверное расширение файла: " + str(self.instruction) + "  Ожидается расширение .xml")        
                
        self.builder = None 
        self.parser = None 
        self.commands = []
        self.replace_dict = {}

    def _check_file(self, path, mode = 'r+'): #проверяет что файл существует и не открыт где-то еще 
        path = Path(path)
        # проверка существования файла
        if not path.is_file():
            return False
        # проверка доступа
        try:
            with open(path, mode):
                return True
        except PermissionError:
            raise ValueError('Файл ' + str(path) + ' открыт в другой программе или недоступен')
        return True
    
    def _unique_file_name(self, path, active = True): #если указанный файл существует, создает новое имя  
        path = Path(path)
        if not active: 
            return path
        name = path.stem
        count = 1

        # подбираем новое незанятое имя
        while self._check_file(path):
            path = path.with_stem(name + '(' + str(count) + ')')
            count += 1
        return path

    def _check_suffix(self, path, correct_suffix = '.docx'): #проверяет правильное ли расширение файла
        path = Path(path)
        
        if not path.suffix: 
            return path.with_suffix(correct_suffix)
        elif path.suffix != correct_suffix:
            raise ValueError("Неверное расширение файла: " + str(path) + "  Ожидается расширение " + correct_suffix)
        else:
            return path

    
    def create_document(self): #создает документ для работы
        # проверяем расширение документа
        if self.doc_path == None: 
            self.doc_path = "Новый файл.docx"
        else: self.doc_path = self._check_suffix(self.doc_path, ".docx")
        # проверяем, существует ли имя 
        if self._check_file(self.doc_path): 
            self.doc_path = self._unique_file_name(self.doc_path, (not CAN_REPLACE_FILES))

        # проверяем существование шаблона
        if self.sample_path!=None:
            # проверяем расширение шаблона
            self.sample_path = self._check_suffix(self.sample_path, ".docx")
            if self._check_file(self.sample_path):
                shutil.copy(self.sample_path, self.doc_path)
            else:
                raise ValueError("Не удалось найти шаблон " + str(self.sample_path))
        


    def _read_parsed_data(self): #читает что получил парсер 
        # чтение путей к файлам
        self.doc_path = self.parser.get_doc_path()
        self.sample_path = self.parser.get_sample_path()

        self.table_sample_path = self.parser.get_table_sample_path()

        if self.doc_path is not None: 
            self.doc_path = Path(self.doc_path)
            if self.doc_path.parent == Path('.'): self.doc_path = PATH_TO_NEW_FILES / self.doc_path
        
        if self.sample_path is not None: 
            self.sample_path = Path(self.sample_path)
            if self.sample_path.parent == Path('.'): self.sample_path = PATH_TO_SAMPLES / self.sample_path

        self.table_sample_path = Path(self.table_sample_path) if self.table_sample_path is not None else None

        if self.table_sample_path!= None: 
            self.table_sample_path = self._check_suffix(self.table_sample_path)
            if not self._check_file(self.table_sample_path, 'r'):
                raise ValueError("Не удалось найти шаблон " + str(self.table_sample_path))
        
        # чтение слов для замены
        self.replace_dict = self.parser.get_words_to_replace()  

        # чтение массива команд инструкции
        self.commands = self.parser.get_instructions()

      
    def construct_document(self):
        self.parser = pr.Parser(self.instruction, self.type)
        self._read_parsed_data()

        # создание документа
        self.create_document()

        builder = self.builder = bld.DocBuilder(Path(self.doc_path), self.table_sample_path)
        
        builder.read_samples()
        builder.replace_words(self.replace_dict)
        
        # добавление содержимого в файл
        for command, parametrs in self.commands: 
            match command:
                case pr.Tags.after:
                    builder.set_add_after(parametrs)
                case pr.Tags.paragraph:
                    builder.add_paragraph(*parametrs)
                case pr.Tags.textblock:
                    builder.add_textblock(*parametrs)
                case pr.Tags.table:
                    builder.add_table(*parametrs)


if __name__ == "__main__":
    current_directory = Path.cwd()
    xml_path = current_directory / "final\расчет_надежности.xml"
    constructor = Constructor(xml_path)
    constructor.construct_document()
