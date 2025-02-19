# модуль с объявлением интерфейсов и структур данных
from enum import Enum

class WordObject:
    def add_to_file(self, place_to_add, add_after = None):
        pass

#перечисления для обозначений объединения ячеек
class Gluing(Enum):
    UP = 'up'
    LEFT = 'left'
    DIAGONAL = 'diagonal'