from sorter import Sorter

sort = Sorter("C:/Users/Sergey/Desktop/Work/НВ600 документация ОССЗ")

print(sort.count_files())
sort.export_to_xlsx("test.xlsx")