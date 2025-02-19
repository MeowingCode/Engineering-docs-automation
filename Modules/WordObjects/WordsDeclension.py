import pymorphy3
import re

DELIMITER = '_' 

morph = pymorphy3.MorphAnalyzer()

"""проверяет, является ли слово кодом/абривиатурой, которые не нужно изменять"""
def word_is_code(word:str):
    # Проверяем, есть ли в строке символы, которые не являются буквами
    if not word.isalpha():
        return True
    
    # Считаем количество заглавных букв
    uppercase_count = sum(1 for char in word if char.isupper())
    
    # Если больше одной заглавной буквы, возвращаем True
    if uppercase_count > 1:
        return True
    
    return False

"""возвращает нормальную форму маркера, слова или словосочетания"""
def normal_form(word:str):
    word = word.replace(' ', DELIMITER)
    words = [word for word in word.split(DELIMITER) if word]

    norm_words = []
    for word in words: 
        if word_is_code(word):
            norm_words.append(word)
        else: 
            norm_form = morph.parse(word)[0].inflect({'nomn'}).word
            norm_words.append(norm_form)
    return DELIMITER.join(norm_words)


"""принимает ключ словаря и одно слово значения. Изменяет значение по форме (падежу) ключа"""
def change_dislension(sample:str, word_to_change:str):
    
    #разделяем ключ на слова 
    sample = sample.replace('_', " ")
    sample = [word for word in sample.split(' ') if word]

    # определяем форму ключа 
    sample_case = None
    for word in sample: 
        if normal_form(word) == word: 
            sample_case = 'nomn'
        else:
            word_info = morph.parse(word)[0]
            word_case = word_info.tag.case
            # если хотя бы одно слово в нормальной форме, считаем нормальной
            sample_case = word_case
        
        if sample_case=='nomn': 
            break 
            
    #если ключ и так в нормальной форме или форма неопределима - не изменяем значение 
    if sample_case is None or sample_case=='nomn': 
        return word_to_change

    words = word_to_change.split(" ")
    new_form = []
    for word in words: 
        if word_is_code(word):
            new_form.append(word)
        else: 
            word_info = morph.parse(word)[0]
            new_word = word_info.inflect({sample_case}).word
            new_word = word[0] + new_word[1:]
            new_form.append(new_word)
    return " ".join(new_form)


"""разбивает словарь на два словаря (just_replace, replace_with_word_form)"""
def sort_by_wordform_changing(word_to_replace:dict): 
    just_replace = {}
    replace_with_word_form = {}

    for mark, params in word_to_replace.items(): 
        if params['auto_change_dislension']: 
            replace_with_word_form[normal_form(mark)] = (params['text'], params['auto_change_register'])
        else:
            just_replace[mark] = (params['text'], params['auto_change_register'])
    return (just_replace, replace_with_word_form)


def remove_xml_tags(text):
    """
    Удаляет HTML-теги из строки.

    :param text: Исходный текст с HTML-тегами.
    :return: Текст без HTML-тегов.
    """
    # Используем регулярное выражение для удаления тегов
    clean_text = re.sub(r'<[^>]+>', '', text)
    return clean_text



