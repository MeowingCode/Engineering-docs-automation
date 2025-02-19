from copy import deepcopy
from Modules.WordObjects import WordTextRow as r
import docx
from lxml import etree
from copy import deepcopy

# Счетчик с определением номера относительно заголовков 
# требует инициализации Counter = C.Counter(doc) перед началом работы с файлом doc

class Counter:

    _instance = None  # Статическая переменная для хранения единственного экземпляра

    def __new__(cls, doc=None):
        if cls._instance is None:
            cls._instance = super(Counter, cls).__new__(cls)
            cls._instance.doc = doc
            
            #коды заголовков
            cls._instance.code = []
            # индексы заголовков
            cls._instance.index = []
            
            # информация для каждого класса. Название класса - ключ списокв 
            # форматы вывода 
            cls._instance.format = {}
            #счетчики
            cls._instance.counters = {}
        return cls._instance
    
    def add_counter(self, counter_name, format = None):
        if not counter_name in self.counters.keys():
            self.counters[counter_name] = 0
            self.format[counter_name] = format
    
    # сбрасывает все указанные счетчики
    def reset_counters(self, counter_names = None):
        if not counter_names:
            counter_names = self._instance.counters.keys()
        for key in counter_names:
            self._instance.counters[key] = 0


    def update_file_map(self, doc = None):
        if doc:
            self.doc = doc
        self.code = []
        self.index = []
        for i,p in enumerate(self.doc.paragraphs):
            if p.style.name.startswith('Heading'):
                num = int(p.style.name.split(' ')[-1])
                if len(self.code) == 0:
                    self.code.append("1."*num)
                    self.index.append(i)
                else:
                    
                    last_code = (self.code[-1]).split(".")[:-1]
                    if len(last_code)>=num:
                        code = last_code[:num]
                        code[-1] = str(int(code[-1]) + 1)
                    else:
                        code = ['1' for i in range(num)]
                        code = last_code + code[len(last_code):]
                    code = '.'.join(code)
                    self.code.append(code + '.')
                    self.index.append(i)   
    
    def get_number(self, paragraph, counter_name = None, format=None):
        
        if not counter_name: counter_name = 'Num'
        
        # Инициализируем счётчик, если не существует
        if counter_name not in self.counters:
            # если создается новый счетчик
            self.counters[counter_name] = 1
            self.format[counter_name] = format
        else:
            # если используется ранее созданный счетчик 
            self.counters[counter_name] += 1

            #задаем новый формат или читаем ранее заданный 
            if format: self.format[counter_name] = format
            else: format = self.format[counter_name]

        N = self.counters[counter_name]
        # Если формат 'num', просто возвращаем счётчик
        if format == 'num':
            return str(N)
        elif format.count('num') == 1: 
            return format.replace('num', str(N))
        # Обновляем карту файла повторно и проверяем снова
        self.update_file_map()
        # если нет никаких заголовков, возвращаем счетчик
        if not self.code:
            return str(N)

        # Получаем индекс текущего параграфа
        index = self._get_paragraph_index(paragraph)
        i = 0
        while i+1<len(self.index) and index>=self.index[i+1]: i+=1

        # Обрабатываем форматирование
        N = str(N)
        code = self.code[i] if self.code else None
        if not code:
            return N
        
        # Если формат 'format' существует, обрабатываем его
        if format and 'format' in format:
            return format.replace('num', N).replace('format', code[:-1])
        elif not format:
            return f"{code}{N}"


        format_code = []
        code = code.split('.')
        format = format.split('.')
        for i in range(min(len(format[:-1]),len(code[:-1]))):
            if format[i] != '_':
                format_code.append(str(format[i]))
                format_code[-1] = format_code[-1].replace('num', str(code[i]))

        if format[-1] != '_':
            format_code.append(str(format[-1]))
            format_code[-1] = format_code[-1].replace('num', N)
        
        format_code = ".".join(format_code)
       
        return format_code
    
    # возвращает индекс переданного параграфа
    def _get_paragraph_index(self, paragraph):
        
        add_after(paragraph)
        index = 0
        for p in self.doc.paragraphs:
            if p.text == "!!Num!!":
                delete_paragraph(p)
                break
            index += 1
        return index-1

#удаляет параграф
def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

#вставляет параграф после p_before после p_after или после таблицы, в которой находится p_after
def add_after(p_before, p_after = "!!Num!!"): 
    
    p = p_before._element

    if isinstance(p_after,str):
        p_insert = deepcopy(p_before)
        p_insert.text = p_after
        p_insert = p_insert._p
    else:
        p_insert = deepcopy(p_after._p)
    
    # Проверяем, находится ли второй параграф в таблице

    if 'w:tc' in str(p.getparent()):  # Параграф находится в таблице
        p.getparent().getparent().getparent().addnext(p_insert)
    
    else:  # Обычный параграф
        p.addnext(p_insert)


"""заметки:
1) работает довольно медлено, посмотрим, будет ли это критично в общей массе (только ручное обновление карты?)
2) после перехода через заголовок данные не обновляются"""