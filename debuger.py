from sorter import Sorter
from model import File
from collections import defaultdict
import os
sort = Sorter("НВ600 документация ОССЗ")

print(sort.count_files(True))
# sort.export_to_xlsx("test.xlsx")
# sort.move_files_to_folders()
# print(sort.create_escd_dict())


# def create_escd_dict(self, file_path):
#         """
#         Читает xlsx файл и создает словарь из данных.
#         Ключом является номер (например, 36), значением — название.
#         """
#         escd_dict = {}
#         workbook = load_workbook(file_path)
#         sheet = workbook.active
        
#         for row in sheet.iter_rows(min_row=1, values_only=True):
#             if row[0] and row[1]:
#                 escd_dict[str(row[0])] = row[1]  #  {'36': 'Суда, судовое оборудование'}
        
#         return escd_dict

# import pandas as pd
# from openpyxl import Workbook
# def generate_hierarchy_report(root_folder, excel_path="Отчет по файлам.xlsx"):
#         """
#         Генерирует сводную таблицу в Excel для первых четырех уровней иерархии с информацией о расширениях файлов.
#         Также включает папки-исключения "ОССЗ - Письма" и "ОССЗ - Письма проектировщика" как второй класс с указанным главным классом.
#         """
       
#         hierarchy_data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(int)))))

#         for root, dirs, files in os.walk(root_folder):
#             rel_path = os.path.relpath(root, root_folder).split(os.sep)

#             if len(rel_path) >= 4:
#                 main_class, first_class, second_class, third_class = rel_path[:4]
               
#                 for file_name in files:
#                     extension = os.path.splitext(file_name)[1][1:].lower()
#                     hierarchy_data[main_class][first_class][second_class][third_class][extension] += 1
#                 # if "Прочие документы" in main_class:
#                 #     for file_name in files:
#                 #         extension = os.path.splitext(file_name)[1][1:].lower()
#                 #         hierarchy_data[main_class][None][None][None][extension] += 1

#             elif len(rel_path) >= 2:
#                 main_class, potential_folder = rel_path[:2]
#                 if "ОССЗ - Письма"in potential_folder or "ОССЗ - Письма проектировщика" in potential_folder:
#                     second_class = potential_folder
#                     for file_name in files:
#                         extension = os.path.splitext(file_name)[1][1:].lower()
#                         hierarchy_data[main_class][None][second_class][None][extension] += 1
#                 elif "Прочие документы" in main_class:
             
#                     for file_name in files:
#                         extension = os.path.splitext(file_name)[1][1:].lower()
#                         hierarchy_data[main_class][None][None][None][extension] += 1
     
#         report_data = []

#         for main_class, first_classes in hierarchy_data.items():
#             for first_class, second_classes in first_classes.items():
#                 for second_class, third_classes in second_classes.items():
#                     for third_class, extensions in third_classes.items():
#                         row = {
#                             "Проект": main_class,
#                             "Классификатор код 1": first_class or "",
#                             "Классификатор код 2": second_class,
#                             "Классификатор код 3": third_class or ""
#                         }
#                         for ext, count in extensions.items():
#                             row[ext] = count
#                         report_data.append(row)


#         df = pd.DataFrame(report_data)
#         df.fillna(0, inplace=True)  

#         with pd.ExcelWriter(excel_path) as writer:
#             df.to_excel(writer, index=False, sheet_name="Отчет")
# Пример вызова
# generate_hierarchy_report("тест 2")
            



# def create_escd_dict():
#     """
#     Читает данные из базы данных и создает словарь.
#     Ключом является номер (например, 36), значением — название.
#     """
#     escd_dict = {}
#     for file in File.select():
#         escd_dict[file.num] = file.name  # {'36': 'Суда, судовое оборудование'}
#     return escd_dict

# def export_to_excel( escd_dict, excel_path="output.xlsx"):
#     """
#     Экспортирует данные в Excel, где столбцы распределены по длине ключа (num).
#     """
#     # Подготовка словарей для каждого столбца (двузначные, трехзначные и т.д.)
#     data_by_length = {2: [], 3: [], 4: [], 5: [], 6: []}

#     # Распределение по длине ключа
#     for num, name in escd_dict.items():
#         length = len(num)
#         if length in data_by_length:
#             data_by_length[length].append((num, name))
    
#     # Создание DataFrame для записи в Excel
#     max_len = max(len(v) for v in data_by_length.values())  # максимальная длина среди групп
#     df = pd.DataFrame()

#     for length, data in data_by_length.items():
#         nums, names = zip(*data) if data else ([], [])
#         # Добавляем столбцы для каждой группы (с запасом строк для выравнивания)
#         df[f"{length}-значные номера"] = list(nums) + [""] * (max_len - len(nums))
#         df[f"{length}-значные описания"] = list(names) + [""] * (max_len - len(names))

#     # Сохранение в Excel
#     with pd.ExcelWriter(excel_path) as writer:
#         df.to_excel(writer, index=False, sheet_name="ESCD")

# # Использование функций
# escd_dict = create_escd_dict()
# export_to_excel(escd_dict)