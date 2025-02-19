import xml.etree.ElementTree as ET
from copy import deepcopy

DEFAULT_FILE_NAME = "Новый файл.docx"

class StandartPreprocessor:
    shablon_name = None 
    def __init__(self, root, doc_type):
        self.root = root
        self.type = doc_type 

    def _add_default_file_name(self): 
        if self.root.attrib.get('name') is None or self.root.attrib.get('name').strip() == "": 
            self.root.set("name",DEFAULT_FILE_NAME)
            
        self.doc_path = self.root.attrib.get('name').strip()
    
    def _add_shablon_name(self):
        if self.shablon_name is None: 
            return False 
        
        header = self.root.find('header')
        tag = header.find('sample_path')

        if tag is None: 
            new_tag = ET.Element("sample_path")
            new_tag.text = self.shablon_name
            header.append(new_tag)
        elif tag.text is None or tag.text.strip() == "": 
            tag.text = self.shablon_name
        else: tag.text = tag.text.strip()
        return True

    def _check_row(self,root, num_of_cell, all_rows = False):
        row_tag = './/row' if all_rows else 'row'
        
        gluing_tag = ET.Element("gluing")
        gluing_tag.set("direction","left")
        cell_tag = ET.Element("cell")
        cell_tag.text = " "

        for row in root.findall(row_tag):
            cells = len(row.findall('./*'))
            if cells < num_of_cell: 
                if cells == 1:
                    for _ in range(num_of_cell-1): row.append(gluing_tag)
                else: 
                    for _ in range(num_of_cell - cells): row.append(cell_tag)


    def _check_tables(self):

        body = self.root.find('body')
        for table in body.findall('.//table'):
            cells_in_row = 0
            if not table.find('row_sample'):
                # считаем количество ячеек в каждой строке
                for row in table.findall('.//row'):
                    cells = len(row.findall('./*'))
                    cells_in_row = cells if cells>cells_in_row else cells_in_row
                
                self._check_row(table, cells_in_row, True)
            else: 
                val_cells_in_row = len(table.find('row_sample').findall('.//cell'))
                cells_in_row = max([len(row.findall('./*')) for row in table.find('row_sample').findall('row')])
                if table.find('title'): self._check_row(table.find('title'), cells_in_row)
                if table.find('footer'): self._check_row(table.find('footer'), cells_in_row)
                self._check_row(table, val_cells_in_row)

 
            
    def unique_preprocess(self):
        pass

    def preprocess_data(self):
        self._add_default_file_name()
        self._add_shablon_name()
        self._check_tables()
        self.unique_preprocess()
        return self.root


class Preprocessor(StandartPreprocessor):
    def __new__(cls, root, doc_type):
        # Создаем объект нужного подкласса в зависимости от doc_type
        subclasses = {
            "Расчет надежности": ReliabilityCalculation,
            "Перечень элементов": ElementsList,
            "Карты рабочих режимов": OperatingModesMap,
            "Спецификация": Specification
        }
        subclass = subclasses.get(doc_type, StandartPreprocessor)
        instance = super().__new__(subclass)
        instance.__init__(root, doc_type)
        return instance


class ElementsList(StandartPreprocessor):
    # Перечень элементов
    shablon_name = "Перечень_элементов_sample"
    row_in_title = 22
    row_in_data = 28
    cell_in_row = 4

    def _check_tables(self):
        # дополняет таблицы до одинакового кол-ва ячеек в каждой строке
        cell_tag = ET.Element("cell")
        cell_tag.text = " "
        cells_in_row = 4 
        body = self.root.find('body')
        for table in body.findall('.//table'):
            # добавляем недостающие
            for row in table.findall('./*'):
                cells = len(row.findall('./*'))
                if cells < cells_in_row: 
                    if cells == 1:
                        row.insert(0, cell_tag)
                        for _ in range(cells_in_row-2): row.append(cell_tag)
                    else: 
                        for _ in range(cells_in_row - cells): row.append(cell_tag)

    def _split_table(self):
        row_in_title = self.row_in_title
        row_in_data = self.row_in_data
        body = self.root.find('body')
        table = body.find('table[@mode="fill"]')
        title_table = ET.Element("table")
        title_table.set("mode","fill")
        title_table.set("table_in_file","1")

        data_table = ET.Element("table")
        data_table.set("mode","fill")
        data_table.set("table_in_file","2")

        row_num = len(list(table))
        for i, row in enumerate(list(table)):
            if i < row_in_title: 
                title_table.append(row)
            else:
                data_table.append(row)

        row_num -= row_in_title
        empty_row = ET.Element("row")
        empty_cell = ET.Element("cell")
        empty_cell.text = " "
        for _ in range(self.cell_in_row): empty_row.append(empty_cell)
                

        if row_num <= 0: 
            data_table.set("mode","delete")
            for _ in range(abs(row_num)): title_table.append(empty_row)
        elif row_num%row_in_data:
            for _ in range(row_in_data - row_num%row_in_data): data_table.append(empty_row)

        body.remove(table)
        body.insert(0, data_table)
        body.insert(0, title_table)

     
    def unique_preprocess(self):
        self._split_table()

class Specification(ElementsList): 
    # Спецификация 
    shablon_name = "Спецификация_sample"
    row_in_title = 16
    row_in_data = 28
    cell_in_row = 7

    def _create_empty_row(self, n): 
        empty_row = ET.Element("row")
        empty_cell = ET.Element("cell")
        empty_cell.text = " "
        for _ in range(n): 
            empty_row.append(empty_cell)
        return empty_row
    
    def _check_tables(self):
        # дополняет таблицы до одинакового кол-ва ячеек в каждой строке и пустые строки 
        new_table = ET.Element("table") 
        new_table.set("mode","fill")
        new_table.set("table_in_file","1")

        empty_row = self._create_empty_row(7)
        empty_cell = ET.Element("cell")
        empty_cell.text = " "
        count = 0 
        body = self.root.find('body')
        for table in body.findall('.//table'):
            # добавляем недостающие
            for row in table.findall('./*'):
                count+=1
                new_row = ET.Element("row")
                cells = len(row.findall('./*'))
                new_cell = ET.Element("cell")
                if cells == 1:
                    new_table.append(empty_row)
                    new_table.append(empty_row)
                    for _ in range(4): new_row.append(empty_cell)
                    new_cell.text = row.find('cell').text
                    new_row.append(new_cell)
                    for _ in range(2): new_row.append(empty_cell)
                    new_table.append(new_row)
                    new_table.append(empty_row)

                elif cells == 3: 
                    new_cell.text = str(count)
                    row.insert(0, empty_cell)
                    row.insert(0, new_cell)
                    for _ in range(2): row.insert(0, empty_cell)
                    new_row = row
                    new_table.append(new_row)
        body.remove(table)
        body.append(new_table)



class ReliabilityCalculation(StandartPreprocessor):
    # Расчет надежности
    shablon_name = "Расчет_надежности_sample"

    def _insert_into_third_section(self): 
        body = self.root.find('body')
        insertafter_tag = body.find('insertafter[@mark="Полный назначенный ресурс"]')

        if insertafter_tag is None:
            insertafter_tag = ET.Element('insertafter', mark="Полный назначенный ресурс")
            # извлекаем и перемещаем содержимое body в insertafter
            for content in list(body):
                insertafter_tag.append(content)    

            body.clear()  
            body.append(insertafter_tag)

    def unique_preprocess(self):
        self._insert_into_third_section()

class OperatingModesMap(StandartPreprocessor):
    # Карты рабочих режимов 
    shablon_name = "КРР_sample"
    def unique_preprocess(self):
        pass