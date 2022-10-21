import pymongo
import pandas as pd
import numpy as np
from pyexcelerate import Workbook, Style,Format
from datetime import datetime


# формат для даты
formatt = '%Y-%m-%dT%H:%M:%S'

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

# Датафрейм для разработчиков
df_dev = bace_df

# 1 условие
TimeToEnter = []
ages = bace_df['Age']
i = 0
for j in bace_df['Job']:
 index = j.find('Developer')
 age = ages[i]
 if index>0 and age<=21:
   TimeToEnter.append('9:00')
 elif index>0 and age>21:
   TimeToEnter.append('9:15')
 else:
   TimeToEnter.append('None')
 i = i+1

# Добовляем новый столбец в датафрейм
df_dev['TimeToEnter'] = TimeToEnter

# Зписываем датафрейм в Excel
values = [df_dev.columns] + list(df_dev.values)
wb = Workbook()
ws = wb.new_sheet('sheet name', data=values)
ws.set_col_style(6,Style(format=Format('dd/mm/yy hh:mm:ss')))
wb.save('First.xlsx')