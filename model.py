from peewee import *
import pandas as pd

db = SqliteDatabase("ESCD.db")

class File(Model):
    num = CharField()
    name = CharField()

    class Meta:
        database = db
        table_name = 'files'

class ProjectFile(Model):
    project = CharField()                 # Наименование проекта
    drawing_number = CharField()          # Чертежный номер
    document_name = CharField()           # Наименование документа
    sheets = CharField(null=True)         # Листов, ед.
    scale = CharField(null=True)          # Масштаб
    format = CharField(null=True)         # Формат
    approved_by = CharField(null=True)    # Утвердил
    date = DateField(null=True)           # Дата (утверждения или создания)
    company = CharField(null=True)        # Компания
    file_link = CharField(null=True)      # Ссылка на файл
    note = TextField(null=True)           # Примечание
    file_size = CharField(null=True)      # Размер файла
    file_date = DateField(null=True)      # Дата файла

    class Meta:
        database = db
        table_name = 'project_files'

# db.connect()
# db.create_tables([File])

# def load_excel_to_db(excel_path):
#    
#     data = pd.read_excel(excel_path, header=None, names=['num', 'name'])
#     for _, row in data.iterrows():
#         # Добавление записи в базу данных
#         File.create(num=str(row['num']), name=row['name'])

# load_excel_to_db("ESCD.xlsx")

# for file in File.select():
#     print(file.id, file.num, file.name)

# db.close()