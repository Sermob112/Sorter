import os
from model import File
from PySide6.QtCore import QObject, Signal
import re
import shutil
class Sorter(QObject):
    file_moved = Signal(int)
    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def get_file_names(self):
        print("Запуск метода get_file_names") 
        if not os.path.isdir(self.folder_path):
            print(f"Папка {self.folder_path} не существует")
            return []
        file_names = []
        for _, _, files in os.walk(self.folder_path):
            file_names.extend(files) 

        return file_names
    

    def replace_cyrillic_to_latin(self, text):
        """
        Заменяет кириллические символы Н и В на латинские H и B.
        """
        translation_table = str.maketrans({'Н': 'H', 'В': 'B'})
        return text.translate(translation_table)

    def extract_prefix(self, file_name):
        """
        Преобразует имя файла в формат HBXXX.XXXXXX, добавляя точки между частями.
        Преобразует кириллицу в латиницу перед проверкой.
        """
        file_name = self.replace_cyrillic_to_latin(file_name)
        match = re.match(r'^(HB\d{3})[\s.-]?(\d{6})', file_name)

        if match:
            part1 = match.group(1)  # HBXXX
            part2 = match.group(2)  # XXXXXX
            return f"{part1}.{part2}"       
        return None

    def create_escd_dict(self):
        """
        Читает данные из базы данных и создает словарь.
        Ключом является номер (например, 36), значением — название.
        """
        print("Запуск метода create_escd_dict") 
        escd_dict = {}
    
        
        for file in File.select():
            escd_dict[file.num] = file.name  # {'36': 'Суда, судовое оборудование'}
        
        return escd_dict
    def count_files(self,statusStay):
        file_names = self.get_file_names()
        count = 0
        for file_name in file_names:
            lat_file_name = self.replace_cyrillic_to_latin(file_name)
            prefix = self.extract_prefix(file_name)
            try:
                if prefix:
                    count += 1
                elif re.match(r'^120-\d{3}', lat_file_name) and re.search(r'(HB\d{3})', lat_file_name):
                    count += 1
                elif re.match(r'^HB\d{3}-\d{3}', lat_file_name):
                    count += 1
                elif statusStay == False:
                    count += 1

                # target_folder = self.append_extension_folder(target_folder, file_name)
                # self.copy_file_to_folder(file_name, target_folder, seen, statusMove)
            finally:
                continue
        return count
    def move_files_to_folders(self, destination_folder,statusMove,statusStay):
      
        file_names = self.get_file_names()
        escd_dict = self.create_escd_dict()
        seen = {}

        for file_name in file_names:
            lat_file_name = self.replace_cyrillic_to_latin(file_name)
            prefix = self.extract_prefix(file_name)
            try:
                if prefix:
                    target_folder = self.handle_prefix_case(destination_folder, prefix, escd_dict)
                    self.moveable(target_folder,file_name,seen,statusMove)
                elif re.match(r'^120-\d{3}', lat_file_name) and re.search(r'(HB\d{3})', lat_file_name):
                    target_folder = self.handle_specific_format_case_120(destination_folder, lat_file_name, "ОССЗ - Письма")
                    self.moveable(target_folder,file_name,seen,statusMove)
                elif re.match(r'^HB\d{3}-\d{3}', lat_file_name):
                    target_folder = self.handle_specific_format_case(destination_folder, lat_file_name, "ОССЗ - Письма проектировщика")
                    self.moveable(target_folder,file_name,seen,statusMove)
                if statusStay == False:
                    target_folder = os.path.join(destination_folder, "Прочие документы")
                    self.moveable(target_folder,file_name,seen,statusMove)

                # target_folder = self.append_extension_folder(target_folder, file_name)
                # self.copy_file_to_folder(file_name, target_folder, seen, statusMove)
            finally:
                continue
    def moveable (self,target_folder, file_name,seen,statusMove):
        target_folder = self.append_extension_folder(target_folder, file_name)
        self.copy_file_to_folder(file_name, target_folder, seen, statusMove)

    def handle_prefix_case(self, destination_folder, prefix, escd_dict):
        project_code = prefix[:5]
        project_number = prefix[6:]
        target_folder = os.path.join(destination_folder, project_code)

        for i in range(2, len(project_number) + 1):
            sub_folder = project_number[:i]
            folder_description = escd_dict.get(sub_folder, "")
            if folder_description:
                folder_name = f"{project_code}.{sub_folder} - {folder_description}"
            else:
                folder_name = f"{project_code}.{sub_folder}"

            target_folder = os.path.join(target_folder, folder_name)
        return target_folder
    def handle_specific_format_case_120(self, destination_folder, file_name, specific_folder):
        project_code = re.search(r'(HB\d{3})', file_name).group(1)
        return os.path.join(destination_folder, project_code, project_code + "." + specific_folder)
    
    def handle_specific_format_case(self, destination_folder, file_name, specific_folder):
        project_code = file_name[:5]
        return os.path.join(destination_folder, project_code, project_code + "." + specific_folder)

    def append_extension_folder(self, target_folder, file_name):
        extension = os.path.splitext(file_name)[1][1:].lower()
        if extension:
            target_folder = os.path.join(target_folder, extension)
        return target_folder

    def copy_file_to_folder(self, file_name, target_folder, seen, status):
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        src_path = os.path.join(self.folder_path, file_name)
        dest_path = os.path.join(target_folder, file_name)
        if status == False:
            if file_name not in seen:
                shutil.copy(src_path, dest_path)
                seen[file_name] = True
                self.file_moved.emit(1)
                
        else:
            if file_name not in seen:
                shutil.move(src_path, dest_path)
                seen[file_name] = True
                self.file_moved.emit(1)
                
            

        


    # def move_files_to_folders(self, destination_folder):
    #     """
    #     Распределение файлов по папкам на основе префикса, системы ЕСКД РФ и обработки неправильных форматов.
    #     """
    #     file_names = self.get_file_names()
    #     seen = {}
    #     escd_dict = self.create_escd_dict("ESCD.xlsx")
    #     for file_name in file_names:
    #         prefix = self.extract_prefix(file_name)
    #         lat_file_name = self.replace_cyrillic_to_latin(file_name)

    #         if prefix:
    #             project_code = prefix[:5]  
    #             project_number = prefix[6:]  #

    #             if re.match(r'^HB\d{3}\.\d{6}', prefix):
    #                 target_folder = os.path.join(destination_folder, project_code)  

    
    #                 for i in range(2, len(project_number) + 1):
    #                     sub_folder = project_number[:i]  
                        
    #                     if sub_folder in escd_dict:
    #                         folder_name = f"{sub_folder} - {escd_dict[sub_folder]}" 
    #                     else:
    #                         folder_name = sub_folder 
    #                     target_folder = os.path.join(target_folder, folder_name)
    #             else:
    #                 target_folder = os.path.join(destination_folder, project_code, "Прочие файлы")
            
    #         elif re.match(r'^HB\d{3}-\d{3}', lat_file_name):
    #             project_code = lat_file_name[:5]  
    #             target_folder = os.path.join(destination_folder, project_code, "ОССЗ.Письма проектировщика")
            
    #         elif re.match(r'^120-\d{3}', lat_file_name):
    #             target_folder = os.path.join(destination_folder, "ОССЗ", "Письма")
            
    #         else:
    #             target_folder = os.path.join(destination_folder, "Прочие документы")
            
    #         extension = os.path.splitext(file_name)[1][1:].lower()
    #         if extension:
    #             target_folder = os.path.join(target_folder, extension)

    #         if not os.path.exists(target_folder):
    #             os.makedirs(target_folder)

    #         src_path = os.path.join(self.folder_path, file_name)
    #         dest_path = os.path.join(target_folder, file_name)

    #         if file_name not in seen:
    #             shutil.copy(src_path, dest_path)  
    #             seen[file_name] = True