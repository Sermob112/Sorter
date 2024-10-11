import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import re
import shutil
class Sorter():
    def __init__(self, folder_path):
        self.folder_path = folder_path

    def count_files(self):
        
        if not os.path.isdir(self.folder_path):
            print(f"Папка {self.folder_path} не существует")
            return 0
        file_count = sum([len(files) for _, _, files in os.walk(self.folder_path)])
        return file_count
    def get_file_names(self):
        # Проверка, существует ли папка
        if not os.path.isdir(self.folder_path):
            print(f"Папка {self.folder_path} не существует")
            return []

        # Массив для хранения имен файлов
        file_names = []

        # Проход по папкам и файлам
        for _, _, files in os.walk(self.folder_path):
            file_names.extend(files) 

        return file_names
    def get_file_names(self):
        # Проверка, существует ли папка
        if not os.path.isdir(self.folder_path):
            print(f"Папка {self.folder_path} не существует")
            return []

        # Массив для хранения имен файлов
        file_names = []

        # Проход по папкам и файлам
        for _, _, files in os.walk(self.folder_path):
            file_names.extend(files)  # Добавляем все файлы в массив

        return file_names

    def clean_file_name(self, file_name):
  
        name_without_extension, extension = os.path.splitext(file_name)     
        clean_name = re.sub(r'\s*\(\d+\)$', '', name_without_extension)

        
        return clean_name + extension

    def find_duplicates(self, file_names):
        self.duplicate_counter = 0
        duplicates = []
        seen = {}

        for file in file_names:
            clean_name = self.clean_file_name(file)

            
            if clean_name in seen:
                duplicates.append(file)
            else:
                seen[clean_name] = file

        return duplicates
    
    

    def export_to_xlsx(self, output_file):
        
        file_names = self.get_file_names()

        if not file_names:
            print("Нет файлов для экспорта.")
            return

        # Найти дубликаты
        duplicates = self.find_duplicates(file_names)
        self.duplicate_counter = len(duplicates)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "File List"

        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        ws.append(["File Name"])

        
        for file_name in file_names:
            cell = ws.append([file_name])
            if file_name in duplicates:
                ws.cell(row=ws.max_row, column=1).fill = red_fill  # Выделяем ячейку красным

        # Сохраняем файл
        wb.save(output_file)
        # print(f"Список файлов успешно выгружен в {output_file}, дубликаты выделены.")
        # print(f"Количество дубликатов {self.duplicate_counter}")

    def replace_cyrillic_to_latin(self, text):
        """
        Заменяет кириллические символы Н и В на латинские H и B.
        """
        translation_table = str.maketrans({'Н': 'H', 'В': 'B'})
        return text.translate(translation_table)

    def extract_prefix(self, file_name):
        """
        Извлекает префикс, который допускает как кириллические, так и латинские символы.
        Преобразует кириллицу в латиницу перед проверкой.
        """
        # Заменяем кириллические символы на латинские
        file_name = self.replace_cyrillic_to_latin(file_name)

        # Проверяем, что имя начинается с 'HB' (в латинице или кириллице)
        match = re.match(r'^(HB\w{3})([\s.-]?)(\w{3,7})', file_name)
        if match:
            # Используем найденный разделитель (если он есть)
            separator = match.group(2) if match.group(2) else '.'
            return f"{match.group(1)}{separator}{match.group(3)}"
        return None

    
    def move_files_to_folders(self, destination_folder):
        """
        Распределение файлов по папкам на основе префикса XXXXX.XXXXXX и расширения файла.
        Дубликаты удаляются.
        """
        # Получаем список файлов
        file_names = self.get_file_names()

        # Словарь для хранения встречающихся имен файлов (чтобы избежать дубликатов)
        seen = {}

        # Проход по всем файлам
        for file_name in file_names:
            # Очищаем имя файла
            clean_name = self.clean_file_name(file_name)

            # Извлекаем префикс
            prefix = self.extract_prefix(clean_name)

            # Определяем целевую папку на основе префикса
            if prefix:
                target_folder = os.path.join(destination_folder, prefix)
            else:
                target_folder = os.path.join(destination_folder, "Прочие документы")

            # Извлекаем расширение файла и создаем подпапку
            extension = os.path.splitext(file_name)[1][1:].lower()  # Получаем расширение файла без точки
            if extension:
                target_folder = os.path.join(target_folder, extension)  # Добавляем подпапку с расширением

            # Создаем папку, если ее нет
            if not os.path.exists(target_folder):
                os.makedirs(target_folder)

            # Полный путь к исходному файлу
            src_path = os.path.join(self.folder_path, file_name)

            # Полный путь к целевому файлу
            dest_path = os.path.join(target_folder, file_name)

            # Проверяем, не является ли файл дубликатом
            if clean_name not in seen:
                # Перемещаем файл в целевую папку
                shutil.copy(src_path, dest_path)
                seen[clean_name] = True  # Добавляем файл в список встреченных
            # else:
            #     # Если файл дубликат, удаляем его
            #     os.remove(src_path)

        print("Файлы успешно распределены по папкам.")
