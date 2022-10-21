import pymongo
import pandas as pd
import numpy as np
from pyexcelerate import Workbook, Style,Format
from datetime import datetime


# формат для даты
formatt = '%Y-%m-%dT%H:%M:%S'

#Подключение к mongodb
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["admin"]


# заданные данные
data = {
  "id":[1, 2, 3, 4, 5, 6, 7],
  "Name": ["Alex","Justin","Set","Carlos","Gareth","John","Bob"],
  "Surname": ["Smur","Forman","Carey","Carey","Chapman","James","James"],
  "Age":[21, 25, 35, 40, 19, 27, 25],
  "Job": ["Python Developer","Java Developer","Project Manager","Enterprise architect","Python Developer","IOS Developer","Python Developer"],
  "Datetime":[datetime.strptime('2022-01-01T09:45:12',formatt), datetime.strptime('2022-01-01T11:50:25',formatt), datetime.strptime('2022-01-01T10:00:45',formatt), datetime.strptime('2022-01-01T09:07:36',formatt), datetime.strptime('2022-01-01T11:54:10',formatt), datetime.strptime('2022-01-01T09:56:40',formatt), datetime.strptime('2022-01-01T09:52:45',formatt)]
}

# Базывый датафрейм
bace_df = pd.DataFrame(data)


# Записовает dataframe в Excel
def vExcel(DataF,name:str):
    values = [DataF.columns] + list(DataF.values)
    wb = Workbook()
    ws = wb.new_sheet('sheet name', data=values)
    ws.set_col_style(6, Style(format=Format('dd/mm/yy hh:mm:ss')))
    wb.save(f'{name}.xlsx')

# Создает новую коллекцию
def ToMongoDBColl(DataF,name:str):
    df =DataF
    mycol = mydb[name]
    mycol.insert_many(df.to_dict('records'))

# Датафрейм для разработчиков
df_dev = bace_df

# 1 условие
TimeToEnter = []
ages = bace_df['Age']
i = 0
for j in df_dev['Job']:
 index = j.find('Developer')
 age = ages[i]
 if index>0 and age<=21:
   TimeToEnter.append('9:00')
 elif index>0 and age>21:
   TimeToEnter.append('9:15')
 else:
   TimeToEnter.append(None)
 i = i+1

# Добовляем новый столбец в датафрейм
df_dev['TimeToEnter'] = TimeToEnter

# Записываем датафрейм в Excel
vExcel(df_dev,'First')


# Создание коллекции
ToMongoDBColl(df_dev,"18MoreAnd21andLess")



# Второе условие
df_mang = bace_df
TimeToEnter2 = []
i = 0
for j in df_mang['Job']:
 index = j.find('Manager')
 age = ages[i]
 if index>0 and age>=35:
   TimeToEnter2.append('11:00')
 elif index>0 and age<35:
   TimeToEnter2.append('11:30')
 else:
   TimeToEnter2.append(None)
 i = i+1
df_mang['TimeToEnter'] = TimeToEnter2

# Записываем датафрейм в Excel 2
vExcel(df_mang,'Second')


# Создание коллекции 2
ToMongoDBColl(df_mang,"35AndMore")


# Третье условие
df_arch = bace_df
TimeToEnter3 = []
i = 0
for j in df_mang['Job']:
 index = j.find('architect')
 age = ages[i]
 if index>0:
   TimeToEnter3.append('10:30')
 else:
   TimeToEnter3.append(None)
 i = i+1
df_arch['TimeToEnter'] = TimeToEnter3

# Записываем датафрейм в Excel 3
vExcel(df_arch,'Third')


# Создание коллекции 3
ToMongoDBColl(df_arch,"ArchitectEnterTime")