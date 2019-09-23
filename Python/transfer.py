from collections import OrderedDict
import xlrd
import json

wb=xlrd.open_workbook('./death-data.xlsx')
main=[]

# print(wb.sheet_names())
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
        





