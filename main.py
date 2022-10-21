import pymongo
import pandas as pd
import numpy as np
from pyexcelerate import Workbook, Style
from datetime import datetime

formatt = '%Y-%m-%dT%H:%M:%S'
data = {
  "id":[1, 2, 3, 4, 5, 6, 7],
  "Name": ["Alex","Justin","Set","Carlos","Gareth","John","Bob"],
  "Surname": ["Smur","Forman","Carey","Carey","Chapman","James","James"],
  "Age":[21, 25, 35, 40, 19, 27, 25],
  "Job": ["Python Developer","Java Developer","Project Manager","Enterprise architect","Python Developer","IOS Developer","Python Developer"],
  "Datetime":[datetime.strptime('2022-01-01T09:45:12',formatt), datetime.strptime('2022-01-01T11:50:25',formatt), datetime.strptime('2022-01-01T10:00:45',formatt), datetime.strptime('2022-01-01T09:07:36',formatt), datetime.strptime('2022-01-01T11:54:10',formatt), datetime.strptime('2022-01-01T09:56:40',formatt), datetime.strptime('2022-01-01T09:52:45',formatt)]
}

bace_df = pd.DataFrame(data)