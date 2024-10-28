from peewee import *
import pandas as pd

db = SqliteDatabase("ESCD.db")

class File(Model):
    num = CharField()
    name = CharField()

    class Meta:
        database = db
        table_name = 'files'

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