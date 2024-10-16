from sorter import Sorter

sort = Sorter("НВ600 документация ОССЗ")

print(sort.count_files())
# sort.export_to_xlsx("test.xlsx")
# sort.move_files_to_folders()
print(sort.create_escd_dict("ESKD.xlsx"))