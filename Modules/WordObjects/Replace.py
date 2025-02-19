import re
import shutil
from pathlib import Path
from Modules.WordObjects import WordsDeclension as WD

DELIMITER = '!'


"""Замена расширения файла."""
def rename_document(path: str, new_suffix) -> Path:
    doc_name = Path(path).parent / f'{Path(path).name.removesuffix(Path(path).suffix)}{new_suffix}'
    Path(path).rename(doc_name)
    return doc_name

"""Распаковка архива."""
def extract_doc(path: Path) -> Path:
    path_unpack = path.parent / Path(path).name.removesuffix(Path(path).suffix)
    shutil.unpack_archive(path, path_unpack, "zip")
    return path_unpack

"""Упаковка директории в архив."""
def zip_document(doc: str, path: Path):
    shutil.make_archive(str(path.parent / doc), 'zip', path)

def find_closest_brace(text, start_index):
    match = re.search(r'[{}]', text[start_index:])
    if match:
        return start_index + match.start()
    return -1

def update_txt(txt_xml: str, just_replace:dict = None, replace_with_word_form:dict = None): 

    new_txt_xml = ''
    start_index = -1
    index = -1
    next_index = 0
    open = False 
    text = []

    while next_index >= 0:

        next_index = find_closest_brace(txt_xml, index+1)

        if txt_xml[index] == '{': 
            if open:
                text.append(txt_xml[start_index+1:next_index])
                start_index = next_index
                open = True 
        elif txt_xml[index] == '}': 

            if open: 
                text.append(txt_xml[start_index+1:index])
                text.append(txt_xml[index+1:next_index])
                start_index = next_index
            open = False 
        index = next_index

            
    # while index >=0:
    #     index = find_closest_brace(txt_xml, start_index+1)
    #     if txt_xml[index] == '{': 
    #         if not open or len(text)==0: 
    #             text.append(txt_xml[start_index+1:index])
    #         else: 
    #             text[-1] += txt_xml[start_index:index]
    #         open = True 
        
    #     else: 
    #         if open or len(text)==0: 
    #             text.append(txt_xml[start_index+1:index])
    #         else: 
    #             text[-1] += txt_xml[start_index:index] 
    #         open = False

    #     start_index = index
        


    # text = txt_xml.split(DELIMITER)
    # for i, block in enumerate(text): 
    #     print(i, WD.remove_xml_tags(block))
    #     if i%2==0: 
    #         new_txt_xml += block
    #     else: 
    #         word = WD.remove_xml_tags(block)
    #         print(i, block)
    # for i in range(1, len(text),2): 
    #     new_txt_xml += text[i-1]
    #     word = WD.remove_xml_tags(text[i])
    #     print(i,word)
    #     if word in just_replace.keys(): 
    #         new_txt_xml += text[i]
    #     else: 
    #         new_txt_xml += DELIMITER + text[i] + DELIMITER
    # print(new_txt_xml)
    # start_index = 0
    # mark = 'meow'
    # # print(txt_xml)
    # index = txt_xml.find(DELIMITER, 0)

    # while start_index >= 0: 
    #     start_index = index
    #     index = txt_xml.find(DELIMITER, start_index+1)
    #     if mark: 
    #         mark = ''
    #         new_txt_xml+= txt_xml[start_index:index]
    #     else: 
    #         mark = txt_xml[start_index:index+1]
    #         print(mark)
    # if not mark:
    #     new_txt_xml += txt_xml[index+1:]
    
    # print('\n\n\n')
    # print([txt_xml[index:]])
    # print(index, start_index)
    # print(mark)
    return txt_xml



"""Замена текста в документе."""
# def change_text(file: Path, just_replace:dict, replace_with_word_form:dict) -> None:

#     with open(file, "r", encoding="utf-8") as doc_xml:
#         txt_xml = doc_xml.read()
#         updated_xml_string = update_txt(txt_xml)
    
#     with open(file, "w", encoding="utf-8") as doc_xml:
#         doc_xml.write(updated_xml_string)

def change_text(file: Path, old_text: str, new_text: str) -> None:

    with open(file, "r", encoding="utf-8") as doc_xml:
        txt_xml = doc_xml.read()
        updated_xml_string = re.sub(old_text, new_text, txt_xml)
    
    with open(file, "w", encoding="utf-8") as doc_xml:
        doc_xml.write(updated_xml_string)

"""Замена текста в колонтитулах."""
def change_headers_and_footers(path: Path, old_text: str, new_text: str) -> None:

    for header in path.glob('word/header*.xml'):
        with open(header, "r", encoding="utf-8") as hf_xml:
            hf_txt_xml = hf_xml.read()
            updated_hf_xml_string = re.sub(old_text, new_text, hf_txt_xml)

        with open(header, "w", encoding="utf-8") as hf_xml:
            hf_xml.write(updated_hf_xml_string)


    for footer in path.glob('word/footer*.xml'):
        with open(footer, "r", encoding="utf-8") as hf_xml:
            hf_txt_xml = hf_xml.read()
            updated_hf_xml_string = re.sub(old_text, new_text, hf_txt_xml)

        with open(footer, "w", encoding="utf-8") as hf_xml:
            hf_xml.write(updated_hf_xml_string)




def replace_words(path_doc_file:str, words_to_replace:dict):
    if not words_to_replace: return 

    # path_doc_file = "C:/prog/work/Тест.docx"
    origin_suffix = Path(path_doc_file).suffix  # оригинальное расширение файла
    rename_doc = rename_document(path_doc_file, ".zip")  # имя файла после смены расширения
    path_doc = extract_doc(rename_doc)  # директория куда распакован файл
    Path(rename_doc).unlink()  # удаление оригинального файла
    try:
        # just_replace, replace_with_word_form = WD.sort_by_wordform_changing(words_to_replace)
        # change_text(path_doc / "word" / "document.xml", just_replace, replace_with_word_form)
        #заменяем слова
        for word in words_to_replace.keys():
            change_text(path_doc / "word" / "document.xml", word, words_to_replace[word]) #замена в основном файле
            change_headers_and_footers(path_doc, word, words_to_replace[word])  # Замена текста в колонтитулах
    except: 
        print('replace er')
    finally:
        zip_document(f'{Path(rename_doc).name.removesuffix(Path(rename_doc).suffix)}', Path(path_doc))  # упаковка файла в архив
        rename_document(str(rename_doc), origin_suffix)  # переименование архива в оригинальное имя
        shutil.rmtree(path_doc)  # удаление директории с распакованным документом


if __name__ == "__main__":
    # replace_words("C:/prog/work/Shablon.docx","@Децимальный_номер", "!!!")
    # replace_words("C:/prog/work/Shablon.docx","@Децимальный_номер", "Дн53434")
    # replace_words("C:/prog/work/Shablon_test.docx","Децимальный_номер", "мяумяумяу")
    # replace_words("C:/prog/work/Shablon.docx","РАЗРАБОТАЛ", "Шалумов А.С")
    word_dict = {"ДЕЦИМАЛЬНЫЙ_НОМЕР": {"text": "НКГД.441467.654-400", "auto_change_dislension": False, "auto_change_register": False},
             "ИМЯ_ИЗДЕЛИЯ": {"text": "Прожектор ССД600-Л4", "auto_change_dislension": True, "auto_change_register": True},
             "РАЗРАБОТАЛ": {"text": "Шалумов А.С", "auto_change_dislension": True, "auto_change_register": False}}  
    # replace_words("C:/prog/work/Shablon_test.docx", word_dict)

    text = "текст}текст{текст}еще текст {} и еще {{ текст"
    print(text)
    update_txt(text, word_dict)


# import zipfile

# # Открываем .docx как ZIP-архив
# with zipfile.ZipFile('Тест3.docx', 'r') as docx:
#     # Получаем список всех файлов внутри архива
#     file_list = docx.namelist()
#     print(file_list)
