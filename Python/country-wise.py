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

def appenddata(array,rownum,colnum):
    leveldata=OrderedDict()
    leveldata['cause']=sh.cell(rownum,colnumcause(rownum)).value
    leveldata['rate']=sh.cell(rownum,colnum).value
    array.append(leveldata)

def sublevel(rownum,colnum,colchecklevel):
    data['subcause']=[]
    for rowsublevel in range(rownum,sh.nrows-1):
        checknextrowvalsub=sh.row_values(rowsublevel+1)
        if(checknextrowvalsub[colchecklevel]!='' and checknextrowvalsub[colchecklevel-1]==''):
            appenddata(data['subcause'],rownum,colnum)
    return rownum,data['subcause']

for colnum in range(coldata,sh.ncols):
    data=OrderedDict()
    data['country']=rowval[colnum]
    data['rates']=[]
    for rownum in range(rowdata,sh.nrows-1):
        for colchecklevel in range(2,coldata-1):
            checkrowval=sh.row_values(rownum)
            checknextrowval=sh.row_values(rownum+1)
            if(checkrowval[colchecklevel]!='' and checknextrowval[colchecklevel]==''):
                # if(checknextrowval[colnum]==''):
                appenddata(data['rates'],rownum,colnum)
                (rownum,data['subcause'])=sublevel(rownum+1,colnum+1,colchecklevel) 
                data['rates'].append(data['subcause'])
                break 
            else:
                appenddata(data['rates'],rownum,colnum)
    main.append(data)
    



 






j = json.dumps(main)
# Write to file
with open('data.json', 'w') as f:
    f.write(j)
        





