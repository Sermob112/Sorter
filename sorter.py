import os
import re
import shutil
import traceback
from model import File
from PySide6.QtCore import QObject, Signal


class Sorter(QObject):
    file_moved = Signal(int)
    finished = Signal()

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def get_file_names(self):
        try:
            if not os.path.isdir(self.folder_path):
                print(f"[Ошибка] Папка {self.folder_path} не существует")
                return []
            file_names = []
            for _, _, files in os.walk(self.folder_path):
                for f in files:
                    file_names.append(f)  # сохраняем оригинальные имена без замены
            return file_names
        except Exception as e:
            print(f"[Ошибка] Ошибка в get_file_names: {e}")
            traceback.print_exc()
            return []

    def replace_cyrillic_to_latin(self, text):
        try:
            translation_table = str.maketrans({'Н': 'H', 'В': 'B'})
            return text.translate(translation_table)
        except Exception as e:
            print(f"[Ошибка] Ошибка при замене кириллицы: {e}")
            traceback.print_exc()
            return text

    def extract_prefix(self, file_name):
        try:
            file_name = self.replace_cyrillic_to_latin(file_name)
            match = re.match(r'^(HB\d{3})[\s.-]?(\d{6})', file_name)
            if match:
                return f"{match.group(1)}.{match.group(2)}"
        except Exception as e:
            print(f"[Ошибка] Ошибка при извлечении префикса: {e}")
            traceback.print_exc()
        return None

    def create_escd_dict(self):
        print("Запуск метода create_escd_dict")
        escd_dict = {}
        try:
            for file in File.select():
                escd_dict[file.num] = file.name
        except Exception as e:
            print(f"[Ошибка] Ошибка при чтении из базы данных: {e}")
            traceback.print_exc()
        return escd_dict

    def count_files(self, statusStay):
        count = 0
        try:
            file_names = self.get_file_names()
            for file_name in file_names:
                try:
                    lat_file_name = self.replace_cyrillic_to_latin(file_name)
                    prefix = self.extract_prefix(file_name)
                    if prefix or re.match(r'^120-\d{3}', lat_file_name) and re.search(r'(HB\d{3})', lat_file_name) or re.match(r'^HB\d{3}-\d{3}', lat_file_name) or not statusStay:
                        count += 1
                except Exception as e:
                    print(f"[Ошибка] Ошибка при анализе файла {file_name}: {e}")
                    traceback.print_exc()
        except Exception as e:
            print(f"[Ошибка] Ошибка при подсчете файлов: {e}")
            traceback.print_exc()
        return count

    def move_files_to_folders(self, destination_folder, statusMove, statusStay):
        try:
            file_names = self.get_file_names()
            escd_dict = self.create_escd_dict()
            seen = {}

            for file_name in file_names:
                try:
                    lat_file_name = self.replace_cyrillic_to_latin(file_name)
                    prefix = self.extract_prefix(file_name)

                    if prefix:
                        target_folder = self.handle_prefix_case(destination_folder, prefix, escd_dict)
                        self.moveable(target_folder, file_name, seen, statusMove)

                    elif re.match(r'^120-\d{3}', lat_file_name) and re.search(r'(HB\d{3})', lat_file_name):
                        target_folder = self.handle_specific_format_case_120(destination_folder, lat_file_name, "ОССЗ - Письма")
                        self.moveable(target_folder, file_name, seen, statusMove)

                    elif re.match(r'^HB\d{3}-\d{3}', lat_file_name):
                        target_folder = self.handle_specific_format_case(destination_folder, lat_file_name, "ОССЗ - Письма проектировщика")
                        self.moveable(target_folder, file_name, seen, statusMove)

                    elif not statusStay:
                        target_folder = os.path.join(destination_folder, "Прочие документы")
                        self.moveable(target_folder, file_name, seen, statusMove)

                except Exception as e:
                    print(f"[Ошибка] Ошибка при обработке файла {file_name}: {e}")
                    traceback.print_exc()

        except Exception as e:
            print(f"[Ошибка] Ошибка в move_files_to_folders: {e}")
            traceback.print_exc()
        finally:
            self.finished.emit()

    def extract_description(self, file_name):
        try:
            parts = file_name.split("_")
            if len(parts) > 1:
                return parts[1].replace(".dwg", "").replace(".xlsx", "")
        except Exception as e:
            print(f"[Ошибка] Ошибка при извлечении описания из имени файла {file_name}: {e}")
            traceback.print_exc()
        return ""

    def moveable(self, target_folder, file_name, seen, statusMove):
        try:
            self.copy_file_to_folder(file_name, target_folder, seen, statusMove)
        except Exception as e:
            print(f"[Ошибка] Ошибка в moveable для файла {file_name}: {e}")
            traceback.print_exc()

    def handle_prefix_case(self, destination_folder, prefix, escd_dict):
        try:
            project_code = prefix[:5]
            project_number = prefix[6:]
            target_folder = os.path.join(destination_folder, project_code)

            for i in range(2, len(project_number) + 1):
                sub_folder = project_number[:i]
                folder_description = escd_dict.get(sub_folder, "")
                folder_name = f"{project_code}.{sub_folder} - {folder_description}" if folder_description else f"{project_code}.{sub_folder}"
                target_folder = os.path.join(target_folder, folder_name)
            return target_folder
        except Exception as e:
            print(f"[Ошибка] Ошибка при формировании пути по префиксу: {e}")
            traceback.print_exc()
            return destination_folder

    def handle_specific_format_case_120(self, destination_folder, file_name, specific_folder):
        try:
            project_code = re.search(r'(HB\d{3})', file_name).group(1)
            return os.path.join(destination_folder, project_code, f"{project_code}.{specific_folder}")
        except Exception as e:
            print(f"[Ошибка] Ошибка при разборе 120-файла: {e}")
            traceback.print_exc()
            return destination_folder

    def handle_specific_format_case(self, destination_folder, file_name, specific_folder):
        try:
            project_code = file_name[:5]
            return os.path.join(destination_folder, project_code, f"{project_code}.{specific_folder}")
        except Exception as e:
            print(f"[Ошибка] Ошибка при разборе HB-файла: {e}")
            traceback.print_exc()
            return destination_folder

    def append_extension_folder(self, target_folder, file_name):
        try:
            extension = os.path.splitext(file_name)[1][1:].lower()
            return os.path.join(target_folder, extension) if extension else target_folder
        except Exception as e:
            print(f"[Ошибка] Ошибка при добавлении папки по расширению: {e}")
            traceback.print_exc()
            return target_folder

    def copy_file_to_folder(self, file_name, target_folder, seen, status):
        try:
            os.makedirs(target_folder, exist_ok=True)

            src_path = self.find_file_recursive(file_name)
            if not src_path:
                print(f"[Ошибка] Файл не найден: {file_name} в {self.folder_path}")
                return

            dest_path = os.path.join(target_folder, file_name)

            if file_name not in seen:
                if status:
                    shutil.move(src_path, dest_path)
                else:
                    shutil.copy(src_path, dest_path)
                seen[file_name] = True
                self.file_moved.emit(1)

        except Exception as e:
            print(f"[Ошибка] Не удалось {'переместить' if status else 'скопировать'} файл {file_name}: {e}")
            traceback.print_exc()



    def find_file_recursive(self, file_name):
        """
        Ищет файл с указанным именем во всех подкаталогах self.folder_path.
        Возвращает абсолютный путь, если найден, иначе None.
        """
        for root, _, files in os.walk(self.folder_path):
            for f in files:
                if f == file_name:
                    return os.path.join(root, f)
        return None


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