import xml.etree.ElementTree as ET
# from pathlib import Path
from enum import Enum

from Modules import Preparsing as pr
from Modules.WordObjects import Gluing


class Tags(Enum):
    after = 'insertafter'
    paragraph = 'p'
    textblock = 'textblock'
    table = 'table'

class Params(Enum):
    textblock_name = 'mark'
    table_mode = 'mode'
    table_mark = 'table_in_file'
    # table_sample_name = 'sample'
    headline_level = 'headline_level'
    alignment = 'alignment'
    font_size = 'font_size'
    font_name = 'font_name'
    table_footer = 'footer'
    record_sample = 'row_sample'
    vertical = 'vertical'
    replace_dict = 'replace_dict'
    text = 'text'
    table_title = 'title'
    table_data = 'row'
    


# Поведение: 
# достает данные (но не проверяет их достоверность)
# не проверяет реальные ли файлы
class Parser:
    def __init__(self, path, doc_type):
        # достаем дерево файла 
        self.doc_type = doc_type
        tree = ET.parse(path)
        self.root = tree.getroot()
        self.instructions = []
        
        self.paragraph_params = (Params.text, Params.headline_level, Params.alignment, Params.font_size, Params.font_name)
        self.textblock_params = (Params.textblock_name, Params.replace_dict)
        self.table_params = (Params.table_mode, Params.table_mark, Params.vertical, Params.table_title,Params.record_sample, Params.table_footer, Params.table_data)
        
        self.preparser = pr.Preprocessor(self.root, self.doc_type)
        # Shablon_nadejnost
        self.preparser.preprocess_data()


        self.parse()

    
    def parse(self):
        # записываем имя документа, который будем создавать
        self.doc_path = self.root.attrib.get('name').strip()
        if self.doc_path=="": self.doc_path = None

        # читаем заголовок
        header = self.root.find('header')
        self._read_header(header)

        # читаем тело файла
        body = self.root.find('body')
        self._read_body(body)


    def _read_header(self, header):
        # считываем название шаблона
        if header.find('sample_path') is not None and header.find('sample_path').text is not None:
            self.sample_path = header.find('sample_path').text.strip()
            if self.sample_path == "": self.sample_path = None 
        else:
            self.sample_path = None

        # считываем название шаблона таблиц
        if header.find('table_sample_path') is not None and header.find('table_sample_path').text is not None:
            self.table_sample_path = header.find('table_sample_path').text.strip()
            if self.table_sample_path == "": self.table_sample_path = None 
        else:
            self.table_sample_path = None

        # читаем слова для замены
        self.replace_dict = self._read_replace_dict(header)

    def _read_body(self, body):

        for element in body:
            tag = Tags(element.tag)
            match tag: 
                case Tags.after:
                    self.instructions.append((tag, element.attrib.get('mark')))
                    self._read_body(element)
                    params = None
                
                case Tags.paragraph: 
                    params = self._read_tag(element, self.paragraph_params)

                case Tags.textblock:
                    params = self._read_tag(element, self.textblock_params)

                case Tags.table:
                    params = self._read_tag(element, self.table_params)

            self.instructions.append((tag, params))

        
    def _read_tag(self, element, params_to_find):
        params = []
        for param in params_to_find: 
            if param == Params.replace_dict:
                params.append(self._read_replace_dict(element)) 
            elif param == Params.text:
                text = ET.tostring(element, encoding='unicode', method='xml').strip()
                text = str(text[text.find('>')+1:-4])
                params.append(text)
            
            elif param == Params.table_title:
                params.append(self._read_table_data(element, Params.table_title.value)) 

            elif param == Params.record_sample:
                params.append(self._read_table_data(element, Params.record_sample.value))  

            elif param == Params.table_footer:
                params.append(self._read_table_data(element, Params.table_footer.value))  

            elif param == Params.table_data:
                params.append(self._read_table_data(element)) 
           
            elif param == Params.vertical:
                val = element.attrib.get(param.value)
                val = True if val=='True' else False
                params.append(val)
            else:
                params.append(element.attrib.get(param.value))
        return tuple(params)


    def _read_table_data(self, element, rows_type = './*'):
        # находим все элементы с конкретным тегом (title или row) или вообще всех 
        if rows_type!= './*': 
            root = element.find(rows_type)
            if not root: return None
        else: root = element
        
        rows = root.findall('row')
        data = []
        for row in rows:
            data_row = [] 
            for cell in row:
                if cell.tag == 'cell': 
                    if cell.text:
                        data_row.append(cell.text)
                    else:
                        data_row.append(None)
                elif cell.tag == 'gluing':
                    direction = cell.attrib.get('direction')
                    data_row.append(Gluing(direction))
            data.append(data_row.copy())
        return data

    def _read_replace_dict(self, root): 
        replace_dict = {}
        for replace_tag in root.findall('replace'):
            mark = replace_tag.attrib.get('mark')
            if replace_tag.text:
                replace_dict[mark] = replace_tag.text
            else: replace_dict[mark] = ''
        return replace_dict if replace_dict else None



# -----------Возврат значений--------------------------------------------------------
    def get_doc_path(self):
        return self.doc_path
    def get_sample_path(self):
        return self.sample_path
    def get_table_sample_path(self):
        return self.table_sample_path
    
    def get_words_to_replace(self):
        return self.replace_dict    
    
    def get_instructions(self):
        return self.instructions

# if __name__ == "__main__":
    # current_directory = Path.cwd()
    # path = current_directory / "final\исходник.xml"
    # par = Parser(path)