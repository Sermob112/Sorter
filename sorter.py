import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import re
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
        """
        Очищает имя файла, удаляя расширение и номера в скобках в конце.
        """
        # Убираем расширение файла
        name_without_extension = os.path.splitext(file_name)[0]

        # Убираем текст в скобках (например, "(3)")
        clean_name = re.sub(r'\s*\(\d+\)$', '', name_without_extension)

        return clean_name

    def find_duplicates(self, file_names):
        """
        Поиск дубликатов с учетом продвинутого сравнения имен (без расширений и номеров в скобках).
        """
        duplicates = []
        seen = {}

        for file in file_names:
            clean_name = self.clean_file_name(file)

            # Если чистое имя уже встречалось, то это дубликат
            if clean_name in seen:
                duplicates.append(file)
            else:
                seen[clean_name] = file

        return duplicates
    def find_duplicates(self, file_names):
        """
        Поиск дубликатов с учетом продвинутого сравнения имен (без расширений и номеров в скобках).
        """
        duplicates = []
        seen = {}

        for file in file_names:
            clean_name = self.clean_file_name(file)

            # Если чистое имя уже встречалось, то это дубликат
            if clean_name in seen:
                duplicates.append(file)
            else:
                seen[clean_name] = file

        return duplicates

    def export_to_xlsx(self, output_file):
        # Получаем список файлов
        file_names = self.get_file_names()

        if not file_names:
            print("Нет файлов для экспорта.")
            return

        # Найти дубликаты
        duplicates = self.find_duplicates(file_names)

        # Создаем новый XLSX файл
        wb = Workbook()
        ws = wb.active
        ws.title = "File List"

        # Определяем заливку для дубликатов (красная заливка)
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        # Записываем заголовок
        ws.append(["File Name"])

        # Записываем имена файлов и выделяем дубликаты
        for file_name in file_names:
            cell = ws.append([file_name])
            if file_name in duplicates:
                ws.cell(row=ws.max_row, column=1).fill = red_fill  # Выделяем ячейку красным

        # Сохраняем файл
        wb.save(output_file)
        print(f"Список файлов успешно выгружен в {output_file}, дубликаты выделены.")