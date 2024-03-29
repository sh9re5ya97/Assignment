
from collections import OrderedDict
import xlrd
import json
import pymongo
from pymongo import MongoClient

# pandas.read_excel('./death-data.xlsx','Deaths 2004').to_json('jfile.json')
wb=xlrd.open_workbook('./death-data.xlsx')
main=[]


sh = wb.sheet_by_index(0)
rowval=sh.row_values(6) 
print(rowval[7])
print(sh.cell(6,7).value)
coldata=7
rowdata=17

def colnumcause(row):
    rowcause=sh.row_values(row)
    col=coldata-1
    if(rowcause[1]!=''):
        while(rowcause[col]==''):
            col=col-1
    return col

for colnum in range(coldata,sh.ncols):
    data=OrderedDict()
    data['country']=rowval[colnum]
    data['rates']=[]
    for rownum in range(rowdata,sh.nrows-1):
        leveldata=OrderedDict()
        leveldata['cause']=sh.cell(rownum,colnumcause(rownum)).value
        leveldata['rate']=sh.cell(rownum,colnum).value
        data['rates'].append(leveldata)
        
    main.append(data)


j = json.dumps(main)
# Write to file
with open('deathdata.json', 'w') as f:
    f.write(j)
        

# x=pandas.read_json('jfile.json')

myclient = pymongo.MongoClient("mongodb://localhost/27017")
print(myclient)
mydb = myclient["kaam"]
print(mydb)
mycol = mydb["data"]
print(mycol)
with open('deathdata.json','r')as f:
    file_data=json.load(f)
mycol.insert_many(file_data)
myclient.close()





