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
        Преобразует имя файла в формат HBXXX.XXXXXX.XXX, добавляя точки между частями.
        Преобразует кириллицу в латиницу перед проверкой.
        """
        # Заменяем кириллические символы на латинские
        file_name = self.replace_cyrillic_to_latin(file_name)

        # Ищем префикс, который начинается с 'HB', за которым идут 3 цифры, 6 цифр и 3 или более символов
        match = re.match(r'^(HB\d{3})[\s.-]?(\d{6})', file_name)

        if match:
            # Добавляем точки между частями
            part1 = match.group(1)  # HBXXX
            part2 = match.group(2)  # XXXXXX


            # Возвращаем префикс в формате HBXXX.XXXXXX.XXX
            return f"{part1}.{part2}"
        
        return None

    
    def move_files_to_folders(self, destination_folder):
        """
        Распределение файлов по папкам на основе префикса, системы ЕСКД РФ и обработки неправильных форматов.
        """
        file_names = self.get_file_names()
        seen = {}

        for file_name in file_names:
            # clean_name = self.clean_file_name(file_name)
            prefix = self.extract_prefix(file_name)
            lat_file_name = self.replace_cyrillic_to_latin(file_name)
            if prefix:
                project_code = prefix[:5]  # Например, HB600
                project_number = prefix[6:]  # Например, 360063
                # Проверка, соответствует ли имя файла формату HBXXX.XXXXXX
                if re.match(r'^HB\d{3}\.\d{6}', prefix):
                    # Разбиваем префикс по частям
                    # project_code = prefix[:5]  # Например, HB600
                    # project_number = prefix[5:]  # Например, 360063
                    
                    # Создаем вложенные папки по ЕСКД (36 -> 360 -> 3600 -> 36006 -> 360063)
                    target_folder = os.path.join(destination_folder, project_code)  # Папка проекта
                    for i in range(2, len(project_number) + 1):
                        target_folder = os.path.join(target_folder, project_number[:i])
                else:
                    # Если файл не соответствует формату, отправляем в папку "Прочие файлы"
                    target_folder = os.path.join(destination_folder, project_code, "Прочие файлы")
                
            elif re.match(r'^HB\d{3}-\d{3}', lat_file_name):
                project_code = lat_file_name[:5]  # Извлекаем проект HBXXX
                target_folder = os.path.join(destination_folder, project_code, "Письма проектировщика")  # Папка "Письмо проекта"

            elif re.match(r'^120-\d{3}', lat_file_name):
                # Все файлы 120-XXX перемещаются в папку "ОССЗ.Письма"
                target_folder = os.path.join(destination_folder, "ОССЗ", "Письма")
            else:
                target_folder = os.path.join(destination_folder, "Прочие документы")
            # Извлекаем расширение файла и создаем подпапку для расширения
            extension = os.path.splitext(file_name)[1][1:].lower()
            if extension:
                target_folder = os.path.join(target_folder, extension)

            # Создаем папки, если они не существуют
            if not os.path.exists(target_folder):
                os.makedirs(target_folder)

            # Полный путь к исходному файлу
            src_path = os.path.join(self.folder_path, file_name)

            # Полный путь к целевому файлу
            dest_path = os.path.join(target_folder, file_name)

            # Проверяем, не является ли файл дубликатом
            if file_name not in seen:
                shutil.copy(src_path, dest_path)  # Перемещаем файл
                seen[file_name] = True
    #     else:
    #         os.remove(src_path)  # Удаляем дубликат

    # logging.info("Файлы успешно распределены по папкам.")

        
