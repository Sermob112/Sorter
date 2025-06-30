from openpyxl import Workbook,load_workbook
from openpyxl.styles import PatternFill
from collections import defaultdict
import re
import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from collections import defaultdict
import re
import os
import pandas as pd
import traceback  # для вывода полной информации об ошибках


class ExcelGenerator:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.duplicate_counter = 0

    def export_to_xlsx(self, output_file):
        def human_readable_size(size):
            for unit in ['Б', 'КБ', 'МБ', 'ГБ', 'ТБ']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}"
                size /= 1024.0
            return f"{size:.1f} ТБ"

        try:
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
            ws.append(["Название файла", "Расширение", "Размер файла"])

            for full_path in file_names:
                try:
                    file_name = os.path.relpath(full_path, self.folder_path)
                    extension = os.path.splitext(file_name)[1].lower()
                    file_size = os.path.getsize(full_path)
                    readable_size = human_readable_size(file_size)

                    ws.append([file_name, extension, readable_size])

                    # Проверка на дубликат по имени файла (без пути)
                    if os.path.basename(file_name) in duplicates:
                        for col in range(1, 4):
                            ws.cell(row=ws.max_row, column=col).fill = red_fill
                except Exception as e:
                    print(f"[Ошибка] Не удалось обработать файл {full_path}: {e}")
                    traceback.print_exc()

            wb.save(output_file)

        except Exception as e:
            print(f"[Ошибка] Ошибка при экспорте в Excel: {e}")
            traceback.print_exc()

    def find_duplicates(self, file_paths):
        duplicates = []
        seen = {}

        for full_path in file_paths:
            file_name = os.path.basename(full_path)
            clean_name = self.clean_file_name(file_name)

            if clean_name in seen:
                duplicates.append(file_name)
            else:
                seen[clean_name] = file_name

        return duplicates


    def clean_file_name(self, file_name):
        try:
            name_without_extension, extension = os.path.splitext(file_name)
            clean_name = re.sub(r'\s*\(\d+\)$', '', name_without_extension)
            return clean_name + extension
        except Exception as e:
            print(f"[Ошибка] Ошибка при очистке имени файла {file_name}: {e}")
            traceback.print_exc()
            return file_name

    def count_files(self):
        try:
            if not os.path.isdir(self.folder_path):
                print(f"Папка {self.folder_path} не существует")
                return 0
            file_count = sum([len(files) for _, _, files in os.walk(self.folder_path)])
            return file_count
        except Exception as e:
            print(f"[Ошибка] Ошибка при подсчете файлов: {e}")
            traceback.print_exc()
            return 0

    def get_file_names(self):
        """
        Возвращает список файлов с полными путями во всех подкаталогах.
        """
        if not os.path.isdir(self.folder_path):
            print(f"[Ошибка] Папка {self.folder_path} не существует")
            return []

        file_paths = []
        for root, _, files in os.walk(self.folder_path):
            for file in files:
                full_path = os.path.join(root, file)
                file_paths.append(full_path)
        return file_paths


    def generate_hierarchy_report(self, root_folder, excel_path="Отчет по файлам.xlsx"):
        try:
            hierarchy_data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(int)))))

            for root, dirs, files in os.walk(root_folder):
                try:
                    rel_path = os.path.relpath(root, root_folder).split(os.sep)

                    if len(rel_path) >= 4:
                        main_class, first_class, second_class, third_class = rel_path[:4]
                        for file_name in files:
                            extension = os.path.splitext(file_name)[1][1:].lower()
                            hierarchy_data[main_class][first_class][second_class][third_class][extension] += 1

                    elif len(rel_path) >= 2:
                        main_class, potential_folder = rel_path[:2]
                        if "ОССЗ - Письма" in potential_folder or "ОССЗ - Письма проектировщика" in potential_folder:
                            second_class = potential_folder
                            for file_name in files:
                                extension = os.path.splitext(file_name)[1][1:].lower()
                                hierarchy_data[main_class][None][second_class][None][extension] += 1
                        elif "Прочие документы" in main_class:
                            for file_name in files:
                                extension = os.path.splitext(file_name)[1][1:].lower()
                                hierarchy_data[main_class][None][None][None][extension] += 1
                except Exception as e:
                    print(f"[Ошибка] Ошибка при обходе папки {root}: {e}")
                    traceback.print_exc()

            report_data = []

            for main_class, first_classes in hierarchy_data.items():
                for first_class, second_classes in first_classes.items():
                    for second_class, third_classes in second_classes.items():
                        for third_class, extensions in third_classes.items():
                            row = {
                                "Проект": main_class,
                                "Классификатор код 1": first_class or "",
                                "Классификатор код 2": second_class,
                                "Классификатор код 3": third_class or ""
                            }
                            for ext, count in extensions.items():
                                row[ext] = count
                            report_data.append(row)

            df = pd.DataFrame(report_data)
            df.fillna(0, inplace=True)

            with pd.ExcelWriter(excel_path) as writer:
                df.to_excel(writer, index=False, sheet_name="Отчет")

        except Exception as e:
            print(f"[Ошибка] Ошибка при генерации иерархического отчета: {e}")
            traceback.print_exc()
