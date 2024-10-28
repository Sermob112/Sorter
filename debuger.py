from sorter import Sorter

sort = Sorter("НВ600 документация ОССЗ")


# sort.export_to_xlsx("test.xlsx")
# sort.move_files_to_folders()
print(sort.create_escd_dict())


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