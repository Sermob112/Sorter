from openpyxl import Workbook,load_workbook
from openpyxl.styles import PatternFill
import re
import os
class ExcelGenerator():
    def __init__(self, folder_path):
        self.folder_path = folder_path

    def export_to_xlsx(self, output_file):
        """
        Экспорт списка файлов и дубликатов в xlsx с добавлением расширения и размера файла.
        """
        def human_readable_size(size):
            """Форматирует размер файла в КБ, МБ и т.д."""
            for unit in ['Б', 'КБ', 'МБ', 'ГБ', 'ТБ']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}"
                size /= 1024.0
            return f"{size:.1f} ТБ"

        file_names = self.get_file_names()

        if not file_names:
            print("Нет файлов для экспорта.")
            return

        duplicates = self.find_duplicates(file_names)
        self.duplicate_counter = len(duplicates)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "File List"

        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        # Заголовки колонок
        ws.append(["Название файла", "Расширение", "Размер файла"])

        # Добавляем информацию о каждом файле
        for file_name in file_names:
            # Извлечение расширения и размера файла
            extension = os.path.splitext(file_name)[1].lower()  # Расширение файла
            file_size = os.path.getsize(os.path.join(self.folder_path, file_name))  # Размер файла
            readable_size = human_readable_size(file_size)  # Читаемый размер файла
            
            # Добавляем информацию в строку
            ws.append([file_name, extension, readable_size])
            
            # Выделяем дубликаты красным
            if file_name in duplicates:
                for col in range(1, 4):  # Первая, вторая и третья колонки
                    ws.cell(row=ws.max_row, column=col).fill = red_fill

        # Сохраняем Excel файл
        wb.save(output_file)


        
        for file_name in file_names:
            cell = ws.append([file_name])
            if file_name in duplicates:
                ws.cell(row=ws.max_row, column=1).fill = red_fill  # Выделяем ячейку красным

       
        wb.save(output_file)


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
    
    def clean_file_name(self, file_name):
  
        name_without_extension, extension = os.path.splitext(file_name)     
        clean_name = re.sub(r'\s*\(\d+\)$', '', name_without_extension)

        
        return clean_name + extension
    
    def count_files(self):
        
        if not os.path.isdir(self.folder_path):
            print(f"Папка {self.folder_path} не существует")
            return 0
        file_count = sum([len(files) for _, _, files in os.walk(self.folder_path)])
        return file_count
    def get_file_names(self):
       
        if not os.path.isdir(self.folder_path):
            print(f"Папка {self.folder_path} не существует")
            return []
        file_names = []
        for _, _, files in os.walk(self.folder_path):
            file_names.extend(files) 

        return file_names