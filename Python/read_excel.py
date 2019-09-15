import excel2json
import xlrd
import json
import pymongo
import pandas
from pandas import ExcelWriter
from pandas import ExcelFile
pandas.read_excel('./death-data.xlsx','Deaths 2004').to_json('jfile.json')
# y=df.to_json


# j=json.dumps(y)
# Write to file
# with open('data.json', 'w') as f:
#     f.write(y)
# print(y)

